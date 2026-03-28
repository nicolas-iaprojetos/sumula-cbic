var fs = require("fs");
var path = require("path");
var PizZip = require("pizzip");
var Docxtemplater = require("docxtemplater");
 
var TEMPLATE_PATH = path.join(__dirname, "../templates/Sumula_Template_v3.docx");
 
function formatDate(d) {
  if (!d) return "";
  if (d.indexOf("/") >= 0) return d;
  var parts = d.split("-");
  if (parts.length === 3) return parts[2] + "/" + parts[1] + "/" + parts[0];
  return d;
}
 
function esc(s) {
  if (!s) return "";
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
 
function mkBullet(text) {
  return '<w:p><w:pPr><w:pStyle w:val="PargrafodaLista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="120" w:after="120" w:line="276" w:lineRule="auto"/><w:ind w:right="457"/><w:contextualSpacing w:val="0"/><w:jc w:val="both"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
}
 
function mkSubBullet(text) {
  return '<w:p><w:pPr><w:pStyle w:val="PargrafodaLista"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="15"/></w:numPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/><w:ind w:right="318" w:hanging="357"/><w:contextualSpacing w:val="0"/><w:jc w:val="both"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
}
 
/**
 * Convert text content to Word XML bullet paragraphs.
 */
function contentToXml(text) {
  if (!text) return "";
  var lines = text.split("\n");
  var out = "";
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    if (line.indexOf(">> ") === 0) {
      out += mkSubBullet(line.substring(3).trim());
    } else if (line.indexOf("- ") === 0) {
      out += mkBullet(line.substring(2).trim());
    } else {
      out += mkBullet(line);
    }
  }
  return out;
}
 
/**
 * Post-process: replace content paragraphs with proper Word XML bullets.
 *
 * The template v3 uses paraId="FIX00004" for all {conteudo_secao} paragraphs.
 * After docxtemplater renders, each section gets a <w:p> with this paraId
 * containing the plain text (with <w:br/> for line breaks).
 *
 * Strategy: find all <w:p> with paraId="FIX00004" in order, match them 1:1
 * with the secoes array, and replace each with bullet XML generated from
 * the original conteudo_secao text.
 */
function postProcess(xml, secoes) {
  if (!secoes || secoes.length === 0) return xml;
 
  var MARKER = 'paraId="FIX00004"';
 
  // Find all positions of the marker
  var positions = [];
  var searchPos = 0;
  while (true) {
    var idx = xml.indexOf(MARKER, searchPos);
    if (idx < 0) break;
    positions.push(idx);
    searchPos = idx + 1;
  }
 
  if (positions.length === 0) return xml;
 
  // Process in REVERSE order to preserve positions
  for (var i = positions.length - 1; i >= 0; i--) {
    // Get the corresponding section content
    var secIdx = i;
    if (secIdx >= secoes.length) continue;
    var conteudo = secoes[secIdx].conteudo_secao || "";
    if (!conteudo) continue;
 
    // Find the <w:p that contains this marker
    var markerPos = positions[i];
    var pStart = xml.lastIndexOf("<w:p ", markerPos);
    if (pStart < 0) continue;
 
    // Find the closing </w:p>
    var pEnd = xml.indexOf("</w:p>", markerPos);
    if (pEnd < 0) continue;
    pEnd += 6;
 
    // Generate bullet XML
    var bulletXml = contentToXml(conteudo);
    if (!bulletXml) continue;
 
    // Replace
    xml = xml.substring(0, pStart) + bulletXml + xml.substring(pEnd);
  }
 
  return xml;
}
 
/**
 * Remove the ANEXO section when there are no images.
 */
function removeEmptyAnexo(xml) {
  var anexoPos = xml.indexOf(">ANEXO<");
  if (anexoPos < 0) return xml;
 
  // Check if there are real images after ANEXO
  var afterAnexo = xml.substring(anexoPos);
  if (afterAnexo.indexOf("w:drawing") >= 0 || afterAnexo.indexOf("pic:pic") >= 0) {
    return xml;
  }
 
  // Find the paragraph containing ANEXO
  var pStart = xml.lastIndexOf("<w:p ", anexoPos);
  if (pStart < 0) pStart = xml.lastIndexOf("<w:p>", anexoPos);
  if (pStart < 0) return xml;
 
  // Also need to remove the empty sz=48 paragraph BEFORE ANEXO
  // Walk backwards to find it
  var beforeAnexo = xml.substring(0, pStart).trimEnd();
  var prevPEnd = beforeAnexo.lastIndexOf("</w:p>");
  if (prevPEnd >= 0) {
    prevPEnd += 6;
    var prevPStart = beforeAnexo.lastIndexOf("<w:p ", prevPEnd - 200);
    if (prevPStart >= 0) {
      var prevPara = xml.substring(prevPStart, prevPEnd);
      // Only remove if it's the empty spacer paragraph (sz=48, no text content)
      if (prevPara.indexOf('w:sz w:val="48"') >= 0 && prevPara.indexOf("<w:t") < 0) {
        pStart = prevPStart;
      }
    }
  }
 
  // Remove from pStart to just before <w:sectPr
  var sectPr = xml.indexOf("<w:sectPr", anexoPos);
  if (sectPr < 0) return xml;
 
  xml = xml.substring(0, pStart) + xml.substring(sectPr);
  return xml;
}
 
async function generateSumula(data) {
  if (!fs.existsSync(TEMPLATE_PATH)) {
    throw new Error("Template nao encontrado em: " + TEMPLATE_PATH);
  }
 
  var content = fs.readFileSync(TEMPLATE_PATH, "binary");
  var zip = new PizZip(content);
  var doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{", end: "}" }
  });
 
  var secoes = (data.secoes || []).map(function(s) {
    return {
      titulo_secao: s.titulo_secao || s.titulo || "",
      conteudo_secao: s.conteudo_secao || (s.bullets ? s.bullets.join("\n") : "")
    };
  });
 
  // FIX 2: Add encerramento as last section if present
  if (data.encerramento && data.encerramento.trim()) {
    secoes.push({
      titulo_secao: "PALAVRA ABERTA E ENCERRAMENTO",
      conteudo_secao: "- " + data.encerramento.trim()
    });
  }
 
  doc.render({
    data: formatDate(data.data) || "",
    horario_inicio: data.horario_inicio || "",
    horario_fim: data.horario_fim || "",
    local: data.local || "",
    objetivo: data.objetivo || "",
    numero_reuniao: data.numero_reuniao || "",
    titulo_reuniao: data.titulo_reuniao || "",
    ano: data.ano || String(new Date().getFullYear()),
    informes: data.informes || "",
    pauta: data.pauta || [],
    secoes: secoes,
    quorum: data.quorum || [],
    anexos: data.anexos || []
  });
 
  var outputZip = doc.getZip();
  var docXml = outputZip.file("word/document.xml").asText();
 
  // FIX 1: Replace content paragraphs (paraId=FIX00004) with proper bullets
  docXml = postProcess(docXml, secoes);
 
  // FIX 3: Remove empty ANEXO section
  if (!data.anexos || data.anexos.length === 0) {
    docXml = removeEmptyAnexo(docXml);
  }
 
  outputZip.file("word/document.xml", docXml);
  return outputZip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}
 
module.exports = { generateSumula: generateSumula };
 
