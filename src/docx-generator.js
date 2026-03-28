var fs = require("fs");
var path = require("path");
var PizZip = require("pizzip");
var Docxtemplater = require("docxtemplater");

var TEMPLATE_PATH = path.join(__dirname, "../templates/Sumula_Template_v2.docx");

function formatDate(d) {
  if (!d) return "";
  if (d.indexOf("/") >= 0) return d;
  var parts = d.split("-");
  if (parts.length === 3) return parts[2] + "/" + parts[1] + "/" + parts[0];
  return d;
}

function esc(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function mkHeadline(text) {
  return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr><w:spacing w:before="240" w:line="276" w:lineRule="auto"/><w:ind w:right="315"/><w:jc w:val="both"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
}

function mkBullet(text) {
  return '<w:p><w:pPr><w:pStyle w:val="PargrafodaLista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="120" w:after="120" w:line="276" w:lineRule="auto"/><w:ind w:right="457"/><w:contextualSpacing w:val="0"/><w:jc w:val="both"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
}

function mkSubBullet(text) {
  return '<w:p><w:pPr><w:pStyle w:val="PargrafodaLista"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="15"/></w:numPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/><w:ind w:right="318" w:hanging="357"/><w:contextualSpacing w:val="0"/><w:jc w:val="both"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
}

function contentToXml(text) {
  if (!text) return "";
  var lines = text.split("\n");
  var out = "";
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    if (line.indexOf(">> ") === 0) out += mkSubBullet(line.substring(3).trim());
    else if (line.indexOf("- ") === 0) out += mkBullet(line.substring(2).trim());
    else out += mkBullet(line);
  }
  return out;
}

// Build replacement XML for a full section (headline + bullets)
function sectionToXml(titulo, conteudo) {
  return mkHeadline(titulo) + contentToXml(conteudo);
}

// Extract text from a <w:p> XML string
function extractText(pXml) {
  var result = "";
  var re = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  var m;
  while ((m = re.exec(pXml)) !== null) result += m[1];
  return result.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"');
}

function postProcess(xml, secoes) {
  if (!secoes || secoes.length === 0) return xml;

  // Process sections in REVERSE order to preserve offsets
  for (var i = secoes.length - 1; i >= 0; i--) {
    var titulo = secoes[i].titulo_secao || "";
    var conteudo = secoes[i].conteudo_secao || "";
    if (!titulo) continue;

    // Find the title text in the XML
    var searchText = esc(titulo);
    var titleTextPos = xml.indexOf(searchText);
    if (titleTextPos < 0) {
      // Try first 20 chars
      searchText = esc(titulo.substring(0, Math.min(20, titulo.length)));
      titleTextPos = xml.indexOf(searchText);
    }
    if (titleTextPos < 0) continue;

    // Find the enclosing <w:p> of the title
    var titlePStart = xml.lastIndexOf("<w:p", titleTextPos);
    var titlePEnd = xml.indexOf("</w:p>", titleTextPos) + 6;
    if (titlePStart < 0 || titlePEnd < 6) continue;

    // Find the NEXT <w:p> (the content paragraph)
    var contentPStart = xml.indexOf("<w:p", titlePEnd);
    if (contentPStart < 0) continue;
    var contentPEnd = xml.indexOf("</w:p>", contentPStart) + 6;
    if (contentPEnd < 6) continue;

    // Check: is the content paragraph another headline? If so, there's no content to replace
    var contentPar = xml.substring(contentPStart, contentPEnd);
    var isNextHeadline = contentPar.indexOf("w:pBdr") >= 0 && contentPar.indexOf("<w:b/>") >= 0;

    // Build the replacement
    var newXml;
    if (i === 0) {
      // First section: the template already has the headline with correct formatting
      // Just replace the content paragraph
      if (isNextHeadline) {
        // No content paragraph to replace, insert after headline
        newXml = contentToXml(conteudo);
        xml = xml.substring(0, titlePEnd) + newXml + xml.substring(titlePEnd);
      } else {
        newXml = contentToXml(conteudo);
        xml = xml.substring(0, contentPStart) + newXml + xml.substring(contentPEnd);
      }
    } else {
      // Subsequent sections: replace BOTH title paragraph + content paragraph
      // with our properly formatted headline + bullets
      var replaceEnd = isNextHeadline ? titlePEnd : contentPEnd;
      newXml = sectionToXml(titulo, conteudo);
      xml = xml.substring(0, titlePStart) + newXml + xml.substring(replaceEnd);
    }
  }

  return xml;
}

// Fix pauta table: vertical centering of text within cells
function fixPauta(xml) {
  var pautaStart = xml.indexOf("PAUTA DA");
  if (pautaStart < 0) return xml;
  
  var firstHeadline = xml.indexOf("w:pBdr", pautaStart);
  if (firstHeadline < 0) firstHeadline = xml.length;
  
  var before = xml.substring(0, pautaStart);
  var pautaBlock = xml.substring(pautaStart, firstHeadline);
  var after = xml.substring(firstHeadline);
  
  // In pauta rows: ensure trHeight exists and vAlign=center in cells
  pautaBlock = pautaBlock.replace(/<w:tr\b[^>]*>(?:(?!<\/w:tr>).)*<\/w:tr>/gs, function(match) {
    var hasTime = /\d{2}h\d{2}/.test(match);
    var hasApresentador = /Apresenta/.test(match);
    if (!hasTime && !hasApresentador) return match;
    
    // Add row height for vertical centering to work
    if (match.indexOf("w:trHeight") < 0) {
      if (match.indexOf("<w:trPr>") >= 0) {
        match = match.replace("<w:trPr>", '<w:trPr><w:trHeight w:val="454" w:hRule="atLeast"/>');
      } else {
        match = match.replace(/(<w:tr\b[^>]*>)/, '$1<w:trPr><w:trHeight w:val="454" w:hRule="atLeast"/></w:trPr>');
      }
    }
    
    // Ensure vAlign=center in cell properties (vertical centering within cell)
    // Simple approach: if vAlign not present, add it before </w:tcPr>
    if (match.indexOf("w:vAlign") < 0) {
      match = match.split("</w:tcPr>").join('<w:vAlign w:val="center"/></w:tcPr>');
    }
    
    // Do NOT change horizontal alignment - keep text left-aligned
    
    return match;
  });
  
  return before + pautaBlock + after;
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
  
  // Post-processing temporarily simplified for stability
  // TODO: re-enable after fixing XML corruption
  // docXml = postProcess(docXml, secoes);
  // docXml = fixPauta(docXml);
  
  outputZip.file("word/document.xml", docXml);
  return outputZip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

module.exports = { generateSumula: generateSumula };
