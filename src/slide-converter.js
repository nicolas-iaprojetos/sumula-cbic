var fs = require("fs");
var path = require("path");
var { execSync } = require("child_process");

var UPLOAD_DIR = path.join(__dirname, "../data/uploads");

/**
 * Convert a PPT/PPTX/PDF file to an array of PNG images.
 * Uses LibreOffice for PPT→PDF, then pdftoppm for PDF→PNG.
 *
 * @param {string} filePath - Path to the uploaded file
 * @param {string} originalName - Original filename (for extension detection)
 * @returns {Array<{data: Buffer, width: number, height: number, ext: string}>}
 */
function convertToImages(filePath, originalName) {
  var ext = path.extname(originalName || filePath).toLowerCase();
  var baseName = "slides_" + Date.now();
  var workDir = path.join(UPLOAD_DIR, baseName);

  fs.mkdirSync(workDir, { recursive: true });

  var pdfPath;

  try {
    if (ext === ".pdf") {
      // Already PDF
      pdfPath = path.join(workDir, "input.pdf");
      fs.copyFileSync(filePath, pdfPath);
    } else if ([".ppt", ".pptx", ".odp"].indexOf(ext) >= 0) {
      // Convert PPT to PDF using LibreOffice
      console.log("[convertToImages] Converting " + ext + " to PDF...");
      try {
        execSync(
          'libreoffice --headless --convert-to pdf --outdir "' + workDir + '" "' + filePath + '"',
          { timeout: 120000, stdio: "pipe" }
        );
      } catch (e) {
        // Try soffice as fallback
        execSync(
          'soffice --headless --convert-to pdf --outdir "' + workDir + '" "' + filePath + '"',
          { timeout: 120000, stdio: "pipe" }
        );
      }
      // Find the generated PDF
      var files = fs.readdirSync(workDir).filter(function(f) { return f.endsWith(".pdf"); });
      if (files.length === 0) {
        throw new Error("LibreOffice não gerou o PDF. Verifique se está instalado no servidor.");
      }
      pdfPath = path.join(workDir, files[0]);
      console.log("[convertToImages] PDF generated: " + files[0]);
    } else {
      throw new Error("Formato não suportado: " + ext + ". Use .pptx, .ppt, ou .pdf");
    }

    // Convert PDF pages to PNG images using pdftoppm
    console.log("[convertToImages] Converting PDF pages to PNG...");
    var outputPrefix = path.join(workDir, "page");
    try {
      execSync(
        'pdftoppm -png -r 200 "' + pdfPath + '" "' + outputPrefix + '"',
        { timeout: 120000, stdio: "pipe" }
      );
    } catch (e) {
      throw new Error("pdftoppm não está instalado. Execute: apt install poppler-utils");
    }

    // Read generated PNG files
    var pngFiles = fs.readdirSync(workDir)
      .filter(function(f) { return f.startsWith("page-") && f.endsWith(".png"); })
      .sort();

    console.log("[convertToImages] Generated " + pngFiles.length + " page images");

    var images = [];
    for (var i = 0; i < pngFiles.length; i++) {
      var imgPath = path.join(workDir, pngFiles[i]);
      var imgData = fs.readFileSync(imgPath);

      // Get image dimensions from PNG header (bytes 16-23)
      var width = imgData.readUInt32BE(16);
      var height = imgData.readUInt32BE(20);

      images.push({
        data: imgData,
        width: width,
        height: height,
        ext: "png"
      });
    }

    return images;

  } finally {
    // Cleanup
    try {
      var cleanFiles = fs.readdirSync(workDir);
      for (var j = 0; j < cleanFiles.length; j++) {
        fs.unlinkSync(path.join(workDir, cleanFiles[j]));
      }
      fs.rmdirSync(workDir);
    } catch (e) {
      // ignore cleanup errors
    }
  }
}

module.exports = { convertToImages: convertToImages };
