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

function escapeXml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// Headline paragraph XML (bold + border bottom, matching template style)
function headlineXml(text) {
  return '<w:p><w:pPr>' +
    '<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr>' +
    '<w:spacing w:before="240" w:line="276" w:lineRule="auto"/>' +
    '<w:ind w:right="315"/><w:jc w:val="both"/>' +
    '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>' +
    '<w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>' +
    '</w:pPr><w:r><w:rPr>' +
    '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>' +
    '<w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/>' +
    '</w:rPr><w:t xml:space="preserve">' + escapeXml(text) + '</w:t></w:r></w:p>';
}

// Bullet paragraph XML (filled circle marker)
function bulletXml(text) {
  return '<w:p><w:pPr>' +
    '<w:pStyle w:val="PargrafodaLista"/>' +
    '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' +
    '<w:spacing w:before="120" w:after="120" w:line="276" w:lineRule="auto"/>' +
    '<w:ind w:right="457"/><w:contextualSpacing w:val="0"/><w:jc w:val="both"/>' +
    '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>' +
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>' +
    '</w:pPr><w:r><w:rPr>' +
    '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>' +
    '<w:sz w:val="22"/><w:szCs w:val="22"/>' +
    '</w:rPr><w:t xml:space="preserve">' + escapeXml(text) + '</w:t></w:r></w:p>';
}

// Sub-bullet paragraph XML (open circle marker)
function subBulletXml(text) {
  return '<w:p><w:pPr>' +
    '<w:pStyle w:val="PargrafodaLista"/>' +
    '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="15"/></w:numPr>' +
    '<w:spacing w:after="120" w:line="276" w:lineRule="auto"/>' +
    '<w:ind w:right="318" w:hanging="357"/><w:contextualSpacing w:val="0"/><w:jc w:val="both"/>' +
    '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>' +
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>' +
    '</w:pPr><w:r><w:rPr>' +
    '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/>' +
    '<w:sz w:val="22"/><w:szCs w:val="22"/>' +
    '</w:rPr><w:t xml:space="preserve">' + escapeXml(text) + '</w:t></w:r></w:p>';
}

// Convert section content text to proper XML paragraphs
function contentToXml(text) {
  if (!text) return "";
  var lines = text.split("\n");
  var result = "";
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    if (line.indexOf(">> ") === 0) {
      result += subBulletXml(line.substring(3).trim());
    } else if (line.indexOf("- ") === 0) {
      result += bulletXml(line.substring(2).trim());
    } else {
      result += bulletXml(line);
    }
  }
  return result;
}

function postProcessDocXml(xml, secoes) {
  if (!secoes || secoes.length === 0) return xml;

  // Strategy: find each {secoes} output block and replace the content paragraphs
  // with properly formatted headline + bullet paragraphs.
  //
  // The docxtemplater with linebreaks:true puts \n as <w:br/> inside the same <w:p>,
  // which means the title paragraph is correct (bold+border) but the content paragraph
  // has all the text crammed into one paragraph with line breaks.
  //
  // We need to find each content paragraph (the one after a bold+border headline)
  // and split it into individual bullet paragraphs.

  for (var i = 0; i < secoes.length; i++) {
    var s = secoes[i];
    var titulo = s.titulo_secao || s.titulo || "";
    var conteudo = s.conteudo_secao || "";
    
    if (!titulo || !conteudo) continue;
    
    // Find the headline paragraph in the rendered XML
    var tituloEsc = escapeXml(titulo);
    var titlePos = xml.indexOf(tituloEsc);
    if (titlePos < 0) {
      // Try without accents
      tituloEsc = titulo.replace(/[áàãâ]/g, "a").replace(/[éèê]/g, "e").replace(/[íìî]/g, "i").replace(/[óòõô]/g, "o").replace(/[úùû]/g, "u");
      tituloEsc = escapeXml(tituloEsc);
      titlePos = xml.indexOf(tituloEsc);
    }
    if (titlePos < 0) continue;
    
    // Find the paragraph containing the title
    var titlePStart = xml.lastIndexOf("<w:p", titlePos);
    var titlePEnd = xml.indexOf("</w:p>", titlePos) + 6;
    
    // The NEXT paragraph should be the content paragraph
    var contentPStart = xml.indexOf("<w:p", titlePEnd);
    if (contentPStart < 0) continue;
    var contentPEnd = xml.indexOf("</w:p>", contentPStart) + 6;
    
    // Verify this is actually a content paragraph (not another headline)
    var contentPar = xml.substring(contentPStart, contentPEnd);
    if (contentPar.indexOf("w:pBdr") >= 0) continue; // skip if it's another headline
    
    // Replace the content paragraph with properly formatted bullet paragraphs
    var newContent = contentToXml(conteudo);
    if (newContent) {
      xml = xml.substring(0, contentPStart) + newContent + xml.substring(contentPEnd);
    }
  }
  
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

  // Post-process XML
  var outputZip = doc.getZip();
  var docXml = outputZip.file("word/document.xml").asText();
  docXml = postProcessDocXml(docXml, secoes);
  outputZip.file("word/document.xml", docXml);

  return outputZip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

module.exports = { generateSumula: generateSumula };
