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
  return '<w:p><w:pPr><w:pStyle w:val="PargrafodaLista"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:right="457"/><w:jc w:val="both"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
}

function mkBulletSpacer() {
  return '<w:p><w:pPr><w:pStyle w:val="PargrafodaLista"/><w:ind w:right="457"/><w:jc w:val="both"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr></w:p>';
}

function mkSubBullet(text) {
  return '<w:p><w:pPr><w:pStyle w:val="PargrafodaLista"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="15"/></w:numPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/><w:ind w:right="318" w:hanging="357"/><w:contextualSpacing w:val="0"/><w:jc w:val="both"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorHAnsi"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
}

function contentToXml(text) {
  if (!text) return "";
  var lines = text.split("\n");
  var out = "";
  var prevWasBullet = false;
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    if (line.indexOf(">> ") === 0) {
      out += mkSubBullet(line.substring(3).trim());
      prevWasBullet = false;
    } else {
      if (prevWasBullet) out += mkBulletSpacer();
      if (line.indexOf("- ") === 0) {
        out += mkBullet(line.substring(2).trim());
      } else {
        out += mkBullet(line);
      }
      prevWasBullet = true;
    }
  }
  return out;
}

function postProcess(xml, secoes) {
  if (!secoes || secoes.length === 0) return xml;
  var MARKER = 'paraId="FIX00004"';
  var positions = [];
  var searchPos = 0;
  while (true) {
    var idx = xml.indexOf(MARKER, searchPos);
    if (idx < 0) break;
    positions.push(idx);
    searchPos = idx + MARKER.length;
  }
  console.log("[postProcess] Found " + positions.length + " FIX00004 for " + secoes.length + " secoes");
  if (positions.length === 0) return xml;
  var firstSectionStart = -1;
  for (var i = positions.length - 1; i >= 0; i--) {
    if (i >= secoes.length) continue;
    var conteudo = secoes[i].conteudo_secao || "";
    if (!conteudo) continue;
    var markerPos = positions[i];
    var pStart = xml.lastIndexOf("<w:p ", markerPos);
    if (pStart < 0) continue;
    var pEnd = xml.indexOf("</w:p>", markerPos);
    if (pEnd < 0) continue;
    pEnd += 6;
    var bulletXml = contentToXml(conteudo);
    if (!bulletXml) continue;
    console.log("[postProcess] Replacing secao " + i);
    xml = xml.substring(0, pStart) + bulletXml + xml.substring(pEnd);
    if (i === 0) firstSectionStart = pStart;
  }
  // Add spacing before the first section heading
  if (firstSectionStart >= 0) {
    var headingEnd = xml.lastIndexOf("</w:p>", firstSectionStart);
    if (headingEnd >= 0) {
      var headingStart = xml.lastIndexOf("<w:p ", headingEnd);
      if (headingStart >= 0) {
        var spacer = '<w:p><w:pPr><w:spacing w:before="480"/></w:pPr></w:p>';
        xml = xml.substring(0, headingStart) + spacer + xml.substring(headingStart);
      }
    }
  }
  return xml;
}

function removeEmptyAnexo(xml) {
  var anexoPos = xml.indexOf(">ANEXO<");
  if (anexoPos < 0) return xml;
  var afterAnexo = xml.substring(anexoPos);
  if (afterAnexo.indexOf("w:drawing") >= 0) return xml;
  var pAnexo = xml.lastIndexOf("<w:p ", anexoPos);
  if (pAnexo < 0) return xml;
  var beforeAnexo = xml.substring(0, pAnexo).trimEnd();
  var prevPEnd = beforeAnexo.lastIndexOf("</w:p>");
  if (prevPEnd >= 0) {
    prevPEnd += 6;
    var checkArea = xml.substring(Math.max(0, prevPEnd - 300), prevPEnd);
    if (checkArea.indexOf('w:sz w:val="48"') >= 0) {
      var prevPStart = xml.lastIndexOf("<w:p ", prevPEnd - 300);
      if (prevPStart >= 0) pAnexo = prevPStart;
    }
  }
  var sectPr = xml.indexOf("<w:sectPr", anexoPos);
  if (sectPr < 0) return xml;
  xml = xml.substring(0, pAnexo) + xml.substring(sectPr);
  return xml;
}

/**
 * Insert images into the ANEXO section.
 * Each image is added as a relationship in the .docx zip and referenced via w:drawing.
 * 
 * @param {PizZip} zip - The docx zip object
 * @param {string} xml - The document.xml content
 * @param {Array} images - Array of {data: Buffer, width: number, height: number, ext: 'png'|'jpeg'}
 * @returns {string} Modified XML with images inserted
 */
function insertAnexoImages(zip, xml, images) {
  if (!images || images.length === 0) return xml;

  // Find the {#anexos}...{/anexos} block or the ANEXO section
  // After docxtemplater renders with empty anexos, the block may be gone
  // So we look for the ANEXO heading and insert after "Slides Apresentados"
  var slidesPos = xml.indexOf("Slides Apresentados");
  if (slidesPos < 0) {
    console.log("[insertAnexoImages] 'Slides Apresentados' not found");
    return xml;
  }

  // Find the </w:p> after "Slides Apresentados"
  var insertAfter = xml.indexOf("</w:p>", slidesPos);
  if (insertAfter < 0) return xml;
  insertAfter += 6;

  // Read existing relationships to determine next rId
  var relsPath = "word/_rels/document.xml.rels";
  var relsXml = zip.file(relsPath) ? zip.file(relsPath).asText() : "";
  var maxRid = 0;
  var ridMatch;
  var ridRe = /Id="rId(\d+)"/g;
  while ((ridMatch = ridRe.exec(relsXml)) !== null) {
    var n = parseInt(ridMatch[1]);
    if (n > maxRid) maxRid = n;
  }

  // Ensure content types include png/jpeg
  var ctPath = "[Content_Types].xml";
  var ctXml = zip.file(ctPath) ? zip.file(ctPath).asText() : "";
  if (ctXml.indexOf('Extension="png"') < 0) {
    ctXml = ctXml.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>');
  }
  if (ctXml.indexOf('Extension="jpeg"') < 0) {
    ctXml = ctXml.replace("</Types>", '<Default Extension="jpeg" ContentType="image/jpeg"/></Types>');
  }
  zip.file(ctPath, ctXml);

  var imagesXml = "";
  var newRels = "";

  for (var i = 0; i < images.length; i++) {
    var img = images[i];
    var rid = maxRid + i + 1;
    var rIdStr = "rId" + rid;
    var ext = img.ext || "png";
    var mediaName = "image_anexo_" + (i + 1) + "." + ext;
    var mediaPath = "word/media/" + mediaName;

    // Add image file to zip
    zip.file(mediaPath, img.data);

    // Add relationship
    newRels += '<Relationship Id="' + rIdStr + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + mediaName + '"/>';

    // Calculate EMU dimensions (1 inch = 914400 EMU)
    // Target: fit within page width (~16cm = ~5760000 EMU) maintaining aspect ratio
    var maxWidthEmu = 5760000;
    var maxHeightEmu = 7680000;
    var wEmu = (img.width || 800) * 9525; // pixels to EMU (96 DPI)
    var hEmu = (img.height || 600) * 9525;

    // Scale to fit
    if (wEmu > maxWidthEmu) {
      var scale = maxWidthEmu / wEmu;
      wEmu = Math.round(wEmu * scale);
      hEmu = Math.round(hEmu * scale);
    }
    if (hEmu > maxHeightEmu) {
      var scale2 = maxHeightEmu / hEmu;
      wEmu = Math.round(wEmu * scale2);
      hEmu = Math.round(hEmu * scale2);
    }

    // Build image paragraph XML
    imagesXml += '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
      '<w:r><w:rPr><w:noProof/></w:rPr>' +
      '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="' + wEmu + '" cy="' + hEmu + '"/>' +
      '<wp:docPr id="' + (100 + i) + '" name="Anexo ' + (i + 1) + '"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:nvPicPr><pic:cNvPr id="' + (100 + i) + '" name="' + mediaName + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + rIdStr + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + wEmu + '" cy="' + hEmu + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic>' +
      '</wp:inline></w:drawing></w:r></w:p>';

    // Add page break between images (not after the last one)
    if (i < images.length - 1) {
      imagesXml += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    }
  }

  // Insert relationships
  if (newRels) {
    relsXml = relsXml.replace("</Relationships>", newRels + "</Relationships>");
    zip.file(relsPath, relsXml);
  }

  // Insert images after "Slides Apresentados"
  xml = xml.substring(0, insertAfter) + imagesXml + xml.substring(insertAfter);

  console.log("[insertAnexoImages] Inserted " + images.length + " images");
  return xml;
}

async function generateSumula(data) {
  if (!fs.existsSync(TEMPLATE_PATH)) {
    throw new Error("Template nao encontrado em: " + TEMPLATE_PATH);
  }
  console.log("[generateSumula] Starting...");
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

  if (data.encerramento && data.encerramento.trim()) {
    secoes.push({
      titulo_secao: "PALAVRA ABERTA E ENCERRAMENTO",
      conteudo_secao: "- " + data.encerramento.trim()
    });
  }

  console.log("[generateSumula] " + secoes.length + " secoes");

  var hasImages = data.anexo_images && data.anexo_images.length > 0;

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

  // FIX 1: Bullets
  docXml = postProcess(docXml, secoes);

  // FIX 3: Handle ANEXO — insert images or remove empty section
  if (hasImages) {
    docXml = insertAnexoImages(outputZip, docXml, data.anexo_images);
  } else {
    docXml = removeEmptyAnexo(docXml);
  }

  outputZip.file("word/document.xml", docXml);
  return outputZip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

module.exports = { generateSumula: generateSumula };
