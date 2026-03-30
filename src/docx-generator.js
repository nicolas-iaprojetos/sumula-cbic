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

  // Find all <w:p>...</w:p> blocks
  var paragraphs = [];
  var re = /<w:p[\s>]/g;
  var match;
  while ((match = re.exec(xml)) !== null) {
    var pStart = match.index;
    var pEnd = xml.indexOf("</w:p>", pStart);
    if (pEnd < 0) continue;
    pEnd += 6;
    paragraphs.push({ start: pStart, end: pEnd, text: xml.substring(pStart, pEnd) });
  }

  // Identify headlines: paragraphs with BOTH w:pBdr AND w:b in rPr
  var allHeadlines = [];
  for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i].text;
    if (p.indexOf("w:pBdr") >= 0 && p.indexOf("<w:b/>") >= 0) {
      allHeadlines.push(i);
    }
  }

  // Skip the first headline (INFORMES) — its content is filled by docxtemplater, not postProcess
  var headlines = allHeadlines.slice(1);

  console.log("[postProcess] Found " + allHeadlines.length + " headlines total, processing " + headlines.length + " (skipped INFORMES) for " + secoes.length + " secoes");

  var spacer = '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>';

  // Replace ALL paragraphs between consecutive headlines with bullet XML (iterate backwards)
  for (var h = headlines.length - 1; h >= 0; h--) {
    if (h >= secoes.length) continue;

    var contentStart = headlines[h] + 1;
    if (contentStart >= paragraphs.length) continue;

    // Find where this section's content ends: right before the next headline (or next allHeadlines entry)
    var nextBoundary;
    var allIdx = allHeadlines.indexOf(headlines[h]);
    if (allIdx >= 0 && allIdx + 1 < allHeadlines.length) {
      nextBoundary = allHeadlines[allIdx + 1]; // paragraph index of next headline
    } else {
      // Last headline — find the end boundary (e.g., sectPr or ANEXO or QUORUM)
      nextBoundary = contentStart + 1; // fallback: just replace the next paragraph
      for (var k = contentStart; k < paragraphs.length; k++) {
        var pText = paragraphs[k].text;
        if (pText.indexOf("w:sectPr") >= 0 || pText.indexOf(">ANEXO<") >= 0 ||
            pText.indexOf(">QUÓRUM<") >= 0 || pText.indexOf(">QUORUM<") >= 0 ||
            (pText.indexOf("w:pBdr") >= 0 && pText.indexOf("<w:b/>") >= 0)) {
          nextBoundary = k;
          break;
        }
        nextBoundary = k + 1;
      }
    }

    if (contentStart >= nextBoundary) continue;

    var conteudo = secoes[h].conteudo_secao || "";
    if (!conteudo) continue;

    var bulletXml = contentToXml(conteudo);
    if (!bulletXml) continue;

    console.log("[postProcess] Replacing secao " + h + " (headline at paragraph " + headlines[h] + ", content paragraphs " + contentStart + "-" + (nextBoundary - 1) + ")");

    var headlinePara = paragraphs[headlines[h]];
    var replaceEnd = paragraphs[nextBoundary - 1].end;
    var insertBefore = (h > 0) ? spacer : '';

    xml = xml.substring(0, headlinePara.start) + insertBefore + headlinePara.text + spacer + bulletXml + xml.substring(replaceEnd);

    // Re-parse paragraphs since offsets changed
    paragraphs = [];
    re.lastIndex = 0;
    while ((match = re.exec(xml)) !== null) {
      var pS = match.index;
      var pE = xml.indexOf("</w:p>", pS);
      if (pE < 0) continue;
      pE += 6;
      paragraphs.push({ start: pS, end: pE, text: xml.substring(pS, pE) });
    }

    // Re-find all headlines with updated paragraph indices
    allHeadlines = [];
    for (var j = 0; j < paragraphs.length; j++) {
      var pt = paragraphs[j].text;
      if (pt.indexOf("w:pBdr") >= 0 && pt.indexOf("<w:b/>") >= 0) {
        allHeadlines.push(j);
      }
    }
    headlines = allHeadlines.slice(1);
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

  function mkImageDrawing(rIdStr, wEmu, hEmu, docPrId, mediaName) {
    return '<w:r><w:rPr><w:noProof/></w:rPr>' +
      '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="' + wEmu + '" cy="' + hEmu + '"/>' +
      '<wp:docPr id="' + docPrId + '" name="Anexo ' + docPrId + '"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:nvPicPr><pic:cNvPr id="' + docPrId + '" name="' + mediaName + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + rIdStr + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + wEmu + '" cy="' + hEmu + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic>' +
      '</wp:inline></w:drawing></w:r>';
  }

  var imagesXml = "";
  var newRels = "";
  var maxWidthEmu = 5760000;  // 16cm full width
  var maxHeightEmu = 4800000; // ~13cm max height to fit ~2 per page

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

    // Calculate EMU dimensions (1 inch = 914400 EMU, 1 pixel at 96 DPI = 9525 EMU)
    var wEmu = (img.width || 800) * 9525;
    var hEmu = (img.height || 600) * 9525;

    // Scale to fit page width, then cap height so ~2 fit per page
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

    // Each image in a centered paragraph with small spacing — Word breaks pages automatically
    imagesXml += '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="120"/></w:pPr>' +
      mkImageDrawing(rIdStr, wEmu, hEmu, 100 + i, mediaName) + '</w:p>';
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

  // Extract INFORMES from secoes if informes field is empty
  var informesText = data.informes || "";
  if (!informesText) {
    var informesIdx = secoes.findIndex(function(s) { return /^informes$/i.test(s.titulo_secao); });
    if (informesIdx >= 0) {
      informesText = secoes[informesIdx].conteudo_secao || "";
      // Clean bullet prefixes for paragraph format
      informesText = informesText.replace(/^- /gm, "").replace(/\n/g, " ").trim();
      secoes.splice(informesIdx, 1);
    }
  }

  console.log("[generateSumula] " + secoes.length + " secoes (informes extracted: " + (informesText ? "yes" : "no") + ")");
  console.log("[generateSumula] pauta items: " + (data.pauta ? data.pauta.length : 0), JSON.stringify(data.pauta || []).substring(0, 500));

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
    informes: informesText,
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
