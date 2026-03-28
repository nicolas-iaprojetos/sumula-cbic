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
 * Convert text content (with - bullets and >> sub-bullets) to Word XML paragraphs.
 * Input format:
 *   "- First bullet point"
 *   "- Second bullet point"
 *   ">> Sub-bullet under second"
 *   "- Third bullet point"
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
      // Plain text line — treat as bullet
      out += mkBullet(line);
    }
  }
  return out;
}

/**
 * Post-process the rendered XML to convert plain-text section content
 * into proper Word XML bullets.
 *
 * Template v3 structure after docxtemplater renders:
 *   <w:p>...<w:t>{titulo_secao text}</w:t>...</w:p>     ← headline (bold + border)
 *   <w:p>...<w:t>{conteudo_secao text}</w:t>...</w:p>   ← plain text content → needs conversion
 *
 * Strategy: find each content paragraph that follows a headline paragraph,
 * extract its text, convert to bullet XML, and replace the paragraph.
 */
function postProcess(xml, secoes) {
  if (!secoes || secoes.length === 0) return xml;

  for (var i = secoes.length - 1; i >= 0; i--) {
    var conteudo = secoes[i].conteudo_secao || "";
    if (!conteudo) continue;

    // The content was rendered as plain text inside a <w:p> with NormalWeb style.
    // Find it by looking for the escaped content text.
    // Take first 40 chars of the first line as search key
    var firstLine = conteudo.split("\n")[0].trim();
    if (firstLine.indexOf("- ") === 0) firstLine = firstLine.substring(2);
    if (firstLine.indexOf(">> ") === 0) firstLine = firstLine.substring(3);
    var searchKey = esc(firstLine.substring(0, Math.min(40, firstLine.length)));

    var keyPos = xml.indexOf(searchKey);
    if (keyPos < 0) continue;

    // Find the enclosing <w:p> of this content
    var pStart = xml.lastIndexOf("<w:p", keyPos);
    var pEnd = xml.indexOf("</w:p>", keyPos);
    if (pStart < 0 || pEnd < 0) continue;
    pEnd += 6; // include </w:p>

    // Generate bullet XML from the content
    var bulletXml = contentToXml(conteudo);
    if (!bulletXml) continue;

    // Replace the plain-text paragraph with bullet paragraphs
    xml = xml.substring(0, pStart) + bulletXml + xml.substring(pEnd);
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

  var outputZip = doc.getZip();
  var docXml = outputZip.file("word/document.xml").asText();

  // Post-process: convert plain-text section content to proper Word XML bullets
  docXml = postProcess(docXml, secoes);

  outputZip.file("word/document.xml", docXml);
  return outputZip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

module.exports = { generateSumula: generateSumula };
