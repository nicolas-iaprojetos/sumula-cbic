var fs = require("fs");
var path = require("path");
var PizZip = require("pizzip");
var Docxtemplater = require("docxtemplater");

var TEMPLATE_PATH = path.join(__dirname, "../templates/Sumula_Template_v2.docx");

function formatDate(d) {
  // Convert 2026-03-25 to 25/03/2026
  if (!d) return "";
  if (d.indexOf("/") >= 0) return d;
  var parts = d.split("-");
  if (parts.length === 3) return parts[2] + "/" + parts[1] + "/" + parts[0];
  return d;
}

async function generateSumula(data) {
  if (!fs.existsSync(TEMPLATE_PATH)) {
    throw new Error("Template nao encontrado em: " + TEMPLATE_PATH + " - Envie o arquivo Sumula_Template_v2.docx para a pasta templates/");
  }

  var content = fs.readFileSync(TEMPLATE_PATH, "binary");
  var zip = new PizZip(content);

  var doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{", end: "}" }
  });

  // Format sections: ensure conteudo_secao exists
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

  return doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });
}

module.exports = { generateSumula: generateSumula };
