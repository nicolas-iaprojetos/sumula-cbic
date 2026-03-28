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
 * Handles: "- " for bullets, ">> " for sub-bullets, plain text as bullets.
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
 * Post-process rendered XML:
 * 
 * The docxtemplater with linebreaks:true renders multi-line content as a SINGLE
 * <w:p> with <w:br/> between lines:
 * 
 *   <w:p>
 *     <w:r><w:t>- Bullet 1</w:t></w:r>
 *     <w:r><w:br/></w:r>
 *     <w:r><w:t>- Bullet 2</w:t></w:r>
 *   </w:p>
 *
 * We need to find these paragraphs and replace them with proper bullet XML.
 * Strategy: search for the NormalWeb-styled paragraph that contains the first
 * line of each section's content, then replace the ENTIRE <w:p>...</w:p>.
 */
function postProcess(xml, secoes) {
  if (!secoes || secoes.length === 0) return xml;

  // Process in REVERSE order to preserve string positions
  for (var i = secoes.length - 1; i >= 0; i--) {
    var conteudo = secoes[i].conteudo_secao || "";
    if (!conteudo) continue;

    // Get first meaningful line of content to use as search key
    var lines = conteudo.split("\n");
    var firstLine = "";
    for (var k = 0; k < lines.length; k++) {
      var l = lines[k].trim();
      if (l) { firstLine = l; break; }
    }
    if (!firstLine) continue;

    // Strip bullet prefix for search
    if (firstLine.indexOf("- ") === 0) firstLine = firstLine.substring(2);
    if (firstLine.indexOf(">> ") === 0) firstLine = firstLine.substring(3);

    // Take first 50 chars as search key (escaped for XML)
    var searchKey = esc(firstLine.substring(0, Math.min(50, firstLine.length)));

    var keyPos = xml.indexOf(searchKey);
    if (keyPos < 0) {
      // Try shorter key
      searchKey = esc(firstLine.substring(0, Math.min(25, firstLine.length)));
      keyPos = xml.indexOf(searchKey);
    }
    if (keyPos < 0) continue;

    // Find the enclosing <w:p> — walk backwards looking for <w:p> or <w:p 
    var pStart = keyPos;
    while (pStart > 0) {
      pStart = xml.lastIndexOf("<w:p", pStart - 1);
      if (pStart < 0) break;
      var nextChar = xml.charAt(pStart + 4);
      if (nextChar === ">" || nextChar === " ") break;
    }
    if (pStart < 0) continue;

    // Find the closing </w:p> — but we need to find the RIGHT one.
    // Since <w:p> can't nest inside <w:p>, the first </w:p> after pStart is ours.
    // However, we need to make sure we go past the search key position.
    var pEnd = xml.indexOf("</w:p>", keyPos);
    if (pEnd < 0) continue;
    pEnd += 6;

    // Sanity checks
    var paragraphXml = xml.substring(pStart, pEnd);
    if (paragraphXml.indexOf(searchKey) < 0) continue;
    if (paragraphXml.indexOf("w:pBdr") >= 0) continue; // Don't replace headlines

    // Generate bullet XML from the original content text
    var bulletXml = contentToXml(conteudo);
    if (!bulletXml) continue;

    // Replace the entire paragraph with bullet paragraphs
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

  // Check if there are any actual images after ANEXO
  var afterAnexo = xml.substring(anexoPos);
  if (afterAnexo.indexOf("w:drawing") >= 0 || afterAnexo.indexOf("pic:pic") >= 0) {
    return xml; // Has images, keep it
  }

  // Find the paragraph containing ANEXO
  var pStart = xml.lastIndexOf("<w:p ", anexoPos);
  if (pStart < 0) pStart = xml.lastIndexOf("<w:p>", anexoPos);
  if (pStart < 0) return xml;

  // Remove everything from ANEXO paragraph to just before <w:sectPr
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

  // FIX 2: Add encerramento as the last section if present
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

  // FIX 1: Convert plain-text content to proper Word XML bullets
  docXml = postProcess(docXml, secoes);

  // FIX 3: Remove empty ANEXO section
  if (!data.anexos || data.anexos.length === 0) {
    docXml = removeEmptyAnexo(docXml);
  }

  outputZip.file("word/document.xml", docXml);
  return outputZip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

module.exports = { generateSumula: generateSumula };
