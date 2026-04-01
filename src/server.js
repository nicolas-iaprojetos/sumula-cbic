const express = require("express");
const cors = require("cors");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const { execSync } = require("child_process");

const db = require("./database");
const { generateSumula } = require("./docx-generator");
const { extractDocxText } = require("./docx-reader");
const { matchAttendance } = require("./attendance-matcher");
const { convertToImages } = require("./slide-converter");

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
  limits: { fileSize: 50 * 1024 * 1024 },
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

app.delete("/api/reunioes/:reuniaoId", (req, res) => {
  db.deleteReuniao(req.params.reuniaoId);
  res.json({ ok: true });
});

app.get("/api/reunioes/:reuniaoId/presencas", (req, res) => {
  res.json(db.getPresencasReuniao(req.params.reuniaoId));
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

// ── Upload Slides ──
var slidesDir = path.join(__dirname, "../data/slides");
fs.mkdirSync(slidesDir, { recursive: true });
app.use("/data/slides", express.static(slidesDir));

app.post("/api/upload-slides", upload.single("slides"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "Nenhum arquivo enviado" });
    // Check dependencies
    try { execSync("which pdftoppm", { stdio: "pipe" }); } catch (e) {
      return res.status(500).json({ error: "pdftoppm não instalado no servidor. Execute: apt install poppler-utils" });
    }
    var images = convertToImages(req.file.path, req.file.originalname);
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);

    var batch = "slides_" + Date.now();
    var batchDir = path.join(slidesDir, batch);
    fs.mkdirSync(batchDir, { recursive: true });

    var result = [];
    for (var i = 0; i < images.length; i++) {
      var fname = "slide_" + (i + 1) + "." + (images[i].ext || "png");
      var fpath = path.join(batchDir, fname);
      fs.writeFileSync(fpath, images[i].data);
      result.push({ path: "data/slides/" + batch + "/" + fname, width: images[i].width, height: images[i].height });
    }
    res.json({ images: result });
  } catch (err) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    console.error("[upload-slides]", err);
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

    // Convert anexo_images paths to Buffer objects for docx-generator
    if (data.anexo_images && data.anexo_images.length > 0) {
      var resolvedImages = [];
      for (var ai = 0; ai < data.anexo_images.length; ai++) {
        var imgPath = path.join(__dirname, "..", data.anexo_images[ai]);
        if (fs.existsSync(imgPath)) {
          var imgData = fs.readFileSync(imgPath);
          var w = imgData.length > 24 ? imgData.readUInt32BE(16) : 800;
          var h = imgData.length > 24 ? imgData.readUInt32BE(20) : 600;
          resolvedImages.push({ data: imgData, width: w, height: h, ext: "png" });
        }
      }
      data.anexo_images = resolvedImages;
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
// ── Direcionamentos (Radar Estratégico) ──
app.get("/api/gts/:gtId/direcionamentos", (req, res) => {
  var filters = {};
  if (req.query.status) filters.status = req.query.status;
  if (req.query.tipo) filters.tipo = req.query.tipo;
  res.json(db.getDirecionamentos(req.params.gtId, filters));
});

app.post("/api/direcionamentos/extrair", async (req, res) => {
  try {
    var settings = db.getSettings();
    var apiKey = req.body.apiKey || settings.apiKey || "";
    var provider = settings.provider || "anthropic";
    var model = settings.model || "claude-sonnet-4-20250514";
    var texto = req.body.texto || "";
    var gtId = req.body.gtId || "";
    var reuniaoOrigem = req.body.reuniao_origem || "";
    var reuniaoData = req.body.reuniao_data || "";

    if (!texto || !gtId) {
      return res.status(400).json({ error: "texto e gtId sao obrigatorios" });
    }

    var promptExtracao = 'Analise o texto abaixo (sumula/ata de reuniao) e extraia TODOS os direcionamentos encontrados.\n\nClassifique cada item em um dos tipos:\n- deliberacao: decisoes tomadas pelo grupo\n- encaminhamento: acoes delegadas a alguem\n- compromisso: compromissos assumidos com prazo\n- risco: riscos ou alertas identificados\n\nPara cada item extraido, retorne:\n- tipo (deliberacao|encaminhamento|compromisso|risco)\n- titulo (frase curta e objetiva)\n- descricao (detalhamento do item)\n- responsavel (nome da pessoa ou entidade responsavel, se mencionado)\n- prazo (data ou periodo, se mencionado)\n- temas (palavras-chave separadas por virgula)\n\nRetorne APENAS um JSON valido no formato:\n{"direcionamentos":[{"tipo":"...","titulo":"...","descricao":"...","responsavel":"...","prazo":"...","temas":"..."}]}\n\nNao inclua explicacoes, apenas o JSON.\n\nTexto da sumula:\n' + texto;

    var aiMessages = [{ role: "user", content: promptExtracao }];
    var aiResult;

    if (provider === "anthropic") {
      var response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01"
        },
        body: JSON.stringify({ model: model, max_tokens: 4096, messages: aiMessages })
      });
      var data = await response.json();
      if (data.content && data.content[0]) {
        aiResult = data.content[0].text;
      } else {
        return res.status(500).json({ error: "Resposta inesperada da API", details: data });
      }
    } else if (provider === "openai") {
      var response2 = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": "Bearer " + apiKey
        },
        body: JSON.stringify({ model: model, max_tokens: 4096, messages: aiMessages })
      });
      var data2 = await response2.json();
      if (data2.choices && data2.choices[0]) {
        aiResult = data2.choices[0].message.content;
      } else {
        return res.status(500).json({ error: "Resposta inesperada da API", details: data2 });
      }
    } else {
      return res.status(400).json({ error: "Provedor nao suportado: " + provider });
    }

    // Parse JSON from AI response
    var jsonMatch = aiResult.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      return res.status(500).json({ error: "AI nao retornou JSON valido", raw: aiResult });
    }
    var parsed = JSON.parse(jsonMatch[0]);
    var direcionamentos = parsed.direcionamentos || [];

    // Save to database
    var items = [];
    for (var i = 0; i < direcionamentos.length; i++) {
      var d = direcionamentos[i];
      items.push({
        gt_id: gtId,
        reuniao_origem: reuniaoOrigem,
        reuniao_data: reuniaoData,
        tipo: d.tipo,
        titulo: d.titulo,
        descricao: d.descricao || "",
        responsavel: d.responsavel || "",
        prazo: d.prazo || "",
        status: "pendente",
        temas: d.temas || ""
      });
    }
    var saved = db.createDirecionamentosBatch(items);
    res.json({ extracted: direcionamentos.length, items: saved });
  } catch (err) {
    console.error("Erro extraindo direcionamentos:", err);
    res.status(500).json({ error: err.message });
  }
});

app.put("/api/direcionamentos/:id", (req, res) => {
  db.updateDirecionamento(req.params.id, req.body);
  res.json({ ok: true });
});

app.delete("/api/direcionamentos/:id", (req, res) => {
  db.deleteDirecionamento(req.params.id);
  res.json({ ok: true });
});

app.post("/api/direcionamentos/sugerir-pauta", async (req, res) => {
  try {
    var settings = db.getSettings();
    var apiKey = req.body.apiKey || settings.apiKey || "";
    var provider = settings.provider || "anthropic";
    var model = settings.model || "claude-sonnet-4-20250514";
    var gtId = req.body.gtId || "";

    if (!gtId) {
      return res.status(400).json({ error: "gtId e obrigatorio" });
    }

    // Get pending items
    var pendentes = db.getDirecionamentos(gtId, { status: "pendente" });
    var emAndamento = db.getDirecionamentos(gtId, { status: "em_andamento" });
    var todos = pendentes.concat(emAndamento);

    if (todos.length === 0) {
      return res.json({ sugestao: "Nao ha direcionamentos pendentes ou em andamento para sugerir pauta." });
    }

    var listaItens = "";
    for (var i = 0; i < todos.length; i++) {
      var item = todos[i];
      listaItens += "- [" + item.tipo.toUpperCase() + "] " + item.titulo + " (Status: " + item.status + ", Responsavel: " + (item.responsavel || "N/A") + ", Prazo: " + (item.prazo || "N/A") + ")\n";
    }

    var promptPauta = 'Com base nos direcionamentos pendentes e em andamento listados abaixo, sugira uma pauta estruturada para a proxima reuniao do grupo.\n\nA pauta deve:\n1. Priorizar itens com prazo proximo ou vencido\n2. Agrupar itens por tema quando possivel\n3. Incluir tempo estimado para cada item\n4. Sugerir ordem logica de discussao\n\nDirecionamentos ativos:\n' + listaItens + '\n\nRetorne a sugestao em formato texto estruturado, com horarios sugeridos e responsaveis por apresentar cada ponto.';

    var aiMessages = [{ role: "user", content: promptPauta }];
    var aiResult;

    if (provider === "anthropic") {
      var response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01"
        },
        body: JSON.stringify({ model: model, max_tokens: 4096, messages: aiMessages })
      });
      var data = await response.json();
      aiResult = (data.content && data.content[0]) ? data.content[0].text : "Erro ao gerar sugestao";
    } else if (provider === "openai") {
      var response2 = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": "Bearer " + apiKey
        },
        body: JSON.stringify({ model: model, max_tokens: 4096, messages: aiMessages })
      });
      var data2 = await response2.json();
      aiResult = (data2.choices && data2.choices[0]) ? data2.choices[0].message.content : "Erro ao gerar sugestao";
    } else {
      return res.status(400).json({ error: "Provedor nao suportado: " + provider });
    }

    res.json({ sugestao: aiResult, total_itens: todos.length });
  } catch (err) {
    console.error("Erro sugerindo pauta:", err);
    res.status(500).json({ error: err.message });
  }
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
