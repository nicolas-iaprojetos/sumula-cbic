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
app.post("/api/gts/:gtId/match-attendance", upl
