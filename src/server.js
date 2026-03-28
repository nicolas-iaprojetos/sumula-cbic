const express = require("express");
const cors = require("cors");
const multer = require("multer");
const path = require("path");
const fs = require("fs");

const db = require("./database");
const { generateSumula } = require("./docx-generator");
const { extractDocxText } = require("./docx-reader");
const { matchAttendance } = require("./attendance-matcher");

// Global error handlers (prevent silent crashes)
process.on("uncaughtException", (err) => {
  console.error("UNCAUGHT EXCEPTION:", err);
});
process.on("unhandledRejection", (err) => {
  console.error("UNHANDLED REJECTION:", err);
});

const app = express();
const PORT = process.env.PORT || 3000;

// Ensure upload directory exists
const uploadDir = path.join(__dirname, "../data/uploads");
fs.mkdirSync(uploadDir, { recursive: true });

const upload = multer({
  dest: uploadDir,
  limits: { fileSize: 10 * 1024 * 1024 },
});

app.use(cors());
app.use(express.json({ limit: "5mb" }));
app.use(express.static(path.join(__dirname, "../public")));

// Healthcheck
app.get("/health", (req, res) => res.json({ status: "ok", uptime: process.uptime() }));

// ── GTs ──
app.get("/api/gts", (req, res) => {
  res.json(db.getGTs());
});

app.post("/api/gts", (req, res) => {
  res.json(db.createGT(req.body.nome, req.body.descricao));
});

app.put("/api/gts/:id", (req, res) => {
  db.updateGT(req.params.id, req.body.nome, req.body.descricao);
  res.json({ ok: true });
});

app.delete("/api/gts/:id", (req, res) => {
  db.deleteGT(req.params.id);
  res.json({ ok: true });
});

// ── Membros ──
app.get("/api/gts/:gtId/membros", (req, res) => {
  res.json(db.getMembros(req.params.gtId));
});

app.post("/api/gts/:gtId/membros", (req, res) => {
  res.json(db.addMembro(req.params.gtId, req.body));
});

app.put("/api/membros/:id", (req, res) => {
  db.updateMembro(req.params.id, req.body);
  res.json({ ok: true });
});

app.delete("/api/membros/:id", (req, res) => {
  db.deleteMembro(req.params.id);
  res.json({ ok: true });
});

// ── Reunioes & Quorum ──
app.get("/api/gts/:gtId/reunioes", (req, res) => {
  res.json(db.getReunioes(req.params.gtId));
});

app.post("/api/gts/:gtId/reunioes", (req, res) => {
  res.json(db.createReuniao(req.params.gtId, req.body));
});

app.get("/api/gts/:gtId/quorum", (req, res) => {
  var ano = req.query.ano || new Date().getFullYear();
  res.json(db.getQuorum(req.params.gtId, ano));
});

app.post("/api/reunioes/:reuniaoId/presencas", (req, res) => {
  db.registerPresencas(req.params.reuniaoId, req.body.presencas);
  res.json({ ok: true });
});

// ── Match Attendance ──
app.post("/api/gts/:gtId/match-attendance", upload.single("file"), async (req, res) => {
  try {
    var membros = db.getMembros(req.params.gtId);
    var result = await matchAttendance(req.file.path, membros);
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.json(result);
  } catch (err) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ error: err.message });
  }
});

// ── Extract Text ──
app.post("/api/extract-text", upload.single("file"), async (req, res) => {
  try {
    var text = await extractDocxText(req.file.path, req.file.originalname);
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.json({ text: text });
  } catch (err) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ error: err.message });
  }
});

// ── Generate Sumula DOCX ──
app.post("/api/generate-sumula", async (req, res) => {
  try {
    var data = req.body.data;
    var gtId = req.body.gtId;

    if (gtId) {
      var ano = data.ano || String(new Date().getFullYear());
      var quorum = db.getQuorum(gtId, parseInt(ano));
      if (quorum.length > 0) data.quorum = quorum;
      data.ano = ano;
    }

    var buffer = await generateSumula(data);

    var filename = "Sumula_" + (data.numero_reuniao || "reuniao") + "_" + (data.data || "doc").replace(/\//g, "-") + ".docx";
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", "attachment; filename=\"" + filename + "\"");
    res.send(buffer);
  } catch (err) {
    console.error("Erro gerando sumula:", err);
    res.status(500).json({ error: err.message });
  }
});

// ── Settings ──
app.get("/api/settings", (req, res) => {
  res.json(db.getSettings());
});

app.put("/api/settings", (req, res) => {
  db.saveSettings(req.body);
  res.json({ ok: true });
});
// ── AI Proxy (evita CORS do browser) ──
app.post("/api/ai/chat", async (req, res) => {
  try {
    var settings = db.getSettings();
    var apiKey = req.body.apiKey || settings.apiKey || "";
    var provider = req.body.provider || settings.provider || "anthropic";
    var model = req.body.model || settings.model || "claude-sonnet-4-20250514";
    var messages = req.body.messages || [];

    if (provider === "anthropic") {
      var response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01"
        },
        body: JSON.stringify({
          model: model,
          max_tokens: 4096,
          messages: messages
        })
      });
      var data = await response.json();
      res.json(data);
    } else if (provider === "openai") {
      var response2 = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": "Bearer " + apiKey
        },
        body: JSON.stringify({
          model: model,
          max_tokens: 4096,
          messages: messages
        })
      });
      var data2 = await response2.json();
      res.json(data2);
    } else {
      res.status(400).json({ error: "Provedor nao suportado: " + provider });
    }
  } catch (err) {
    console.error("AI Proxy error:", err);
    res.status(500).json({ error: err.message });
  }
});
// SPA fallback
app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "../public/index.html"));
});

app.listen(PORT, "0.0.0.0", function () {
  console.log("");
  console.log("  Sumulas COIC rodando em http://localhost:" + PORT);
  console.log("");
});
