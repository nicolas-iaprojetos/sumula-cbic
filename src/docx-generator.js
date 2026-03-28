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

var BULLET_XML = '<w:p><w:pPr><w:pStyle w:val="PargrafodaLista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="120" w:after="120" w:line="276" w:lineRule="auto"/><w:ind w:right="457"/><w:contextualSpacing w:val="0"/><w:jc w:val="both"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">TEXT_HERE</w:t></w:r></w:p>';

var SUB_BULLET_XML = '<w:p><w:pPr><w:pStyle w:val="PargrafodaLista"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="15"/></w:numPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/><w:ind w:right="318" w:hanging="357"/><w:contextualSpacing w:val="0"/><w:jc w:val="both"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">TEXT_HERE</w:t></w:r></w:p>';

function textToBulletXml(text) {
  if (!text) return "";
  var lines = text.split("\n");
  var result = "";
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    if (line.indexOf(">> ") === 0) {
      result += SUB_BULLET_XML.replace("TEXT_HERE", escapeXml(line.substring(3).trim()));
    } else if (line.indexOf("- ") === 0) {
      result += BULLET_XML.replace("TEXT_HERE", escapeXml(line.substring(2).trim()));
    } else {
      result += BULLET_XML.replace("TEXT_HERE", escapeXml(line));
    }
  }
  return result;
}

function postProcessDocXml(xml) {
  // Find <w:p> elements that contain bullet text markers and expand them
  var result = xml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>).)*<\/w:p>/gs, function(match) {
    // Extract all text from this paragraph
    var texts = [];
    var re = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    var m;
    while ((m = re.exec(match)) !== null) {
      texts.push(m[1]);
    }
    var fullText = texts.join("");
    
    // Unescape XML entities
    var raw = fullText
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"');
    
    // Check if it has bullets
    var bulletCount = 0;
    var pos = 0;
    while ((pos = raw.indexOf("- ", pos)) !== -1) { bulletCount++; pos += 2; }
    var hasSubBullets = raw.indexOf(">> ") >= 0;
    
    if (bulletCount < 2 && !hasSubBullets) return match;
    
    return textToBulletXml(raw);
  });
  
  return result;
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

  // Post-process XML for bullets
  var outputZip = doc.getZip();
  var docXml = outputZip.file("word/document.xml").asText();
  docXml = postProcessDocXml(docXml);
  outputZip.file("word/document.xml", docXml);

  return outputZip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

module.exports = { generateSumula: generateSumula };
