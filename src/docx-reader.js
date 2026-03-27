var mammoth = require("mammoth");
var fs = require("fs");
var path = require("path");

async function extractDocxText(filePath, originalName) {
  var ext = path.extname(originalName || filePath).toLowerCase();

  if (ext === ".docx" || ext === ".doc") {
    var result = await mammoth.extractRawText({ path: filePath });
    return result.value || "";
  }

  if (ext === ".txt") {
    return fs.readFileSync(filePath, "utf-8");
  }

  throw new Error("Formato nao suportado: " + ext);
}

module.exports = { extractDocxText: extractDocxText };
