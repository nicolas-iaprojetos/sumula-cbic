var Database = require("better-sqlite3");
var path = require("path");
var crypto = require("crypto");
var fs = require("fs");

var DB_PATH = process.env.DB_PATH || path.join(__dirname, "../data/sumulas.db");
fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });

var db = new Database(DB_PATH);
db.pragma("journal_mode = WAL");
db.pragma("synchronous = NORMAL");
db.pragma("cache_size = 10000");
db.pragma("foreign_keys = ON");

db.exec(
  "CREATE TABLE IF NOT EXISTS gts (id TEXT PRIMARY KEY, nome TEXT NOT NULL, descricao TEXT DEFAULT '', created_at DATETIME DEFAULT CURRENT_TIMESTAMP);" +
  "CREATE TABLE IF NOT EXISTS membros (id TEXT PRIMARY KEY, gt_id TEXT NOT NULL REFERENCES gts(id) ON DELETE CASCADE, nome TEXT NOT NULL, nome_completo TEXT DEFAULT '', email TEXT DEFAULT '', telefone TEXT DEFAULT '', empresa TEXT DEFAULT '', funcao TEXT DEFAULT '', grupo TEXT DEFAULT '', ativo INTEGER DEFAULT 1, created_at DATETIME DEFAULT CURRENT_TIMESTAMP);" +
  "CREATE TABLE IF NOT EXISTS reunioes (id TEXT PRIMARY KEY, gt_id TEXT NOT NULL REFERENCES gts(id) ON DELETE CASCADE, nome TEXT NOT NULL, edicao TEXT DEFAULT '', data TEXT NOT NULL, horario TEXT DEFAULT '', local_ TEXT DEFAULT '', created_at DATETIME DEFAULT CURRENT_TIMESTAMP);" +
  "CREATE TABLE IF NOT EXISTS presencas (id INTEGER PRIMARY KEY AUTOINCREMENT, reuniao_id TEXT NOT NULL REFERENCES reunioes(id) ON DELETE CASCADE, membro_id TEXT NOT NULL REFERENCES membros(id) ON DELETE CASCADE, presente INTEGER NOT NULL DEFAULT 0, UNIQUE(reuniao_id, membro_id));" +
  "CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);"
);

function uid() {
  return crypto.randomBytes(8).toString("hex");
}

var SEED_MEMBROS = [
  { n: "Ilso Jose de Oliveira", c: "Ilso Jose de Oliveira", e: "ilso@retcinfraestrutura.com.br", p: "COIC", f: "Presidente", g: "A - CBIC" },
  { n: "Hugo Franca", c: "Hugo Franca Cavalcanti de Lima", e: "", p: "CBIC", f: "Gestor COINFRA", g: "A - CBIC" },
  { n: "Claudio Freitas", c: "Claudio Antonio Brito Freitas", e: "claudio.freitas@qualidados.com.br", p: "Qualidados", f: "Socio Diretor", g: "B - Engenharia" },
  { n: "Ricardo Fabel", c: "Ricardo Fabel Braga", e: "Ricardo.Fabel@k2mais.com.br", p: "Tractebel", f: "", g: "B - Engenharia" },
  { n: "Thomas Diepenbruck", c: "Thomas Martin Diepenbruck", e: "Thomas.Diepenbruck@htb.eng.br", p: "HTB", f: "", g: "B - Engenharia" },
  { n: "Thereza Cavalcanti", c: "Thereza Christina Coelho Cavalcanti", e: "", p: "Civil Eng", f: "", g: "B - Engenharia" },
  { n: "Eduardo Aragon", c: "Eduardo Antonio Villela De Aragon", e: "", p: "Brainmarket", f: "", g: "C - Consultoria" },
  { n: "Marcelo Figueiredo", c: "Marcelo Figueiredo", e: "", p: "MF Eng", f: "", g: "C - Consultoria" },
  { n: "Patricia Boson", c: "Patricia Helena Gambogi Boson", e: "", p: "Conciliare", f: "", g: "C - Consultoria" },
  { n: "Eduardo Silvino", c: "Eduardo Silvino", e: "", p: "Retc", f: "", g: "C - Consultoria" },
  { n: "Gustavo Bado", c: "Gustavo Almeida Bado", e: "gustavo.bado@estruturalrs.com.br", p: "Estrutural", f: "", g: "D - Manutencao" },
  { n: "Ricardo Abrahao Netto", c: "Ricardo Antonio Abrahao Netto", e: "ricardo@fortes.ind.br", p: "Fortes Eng", f: "", g: "E - Construcao Industrial" },
  { n: "Edmilson Pires", c: "Edmilson de Araujo Pires", e: "", p: "Marka", f: "", g: "E - Construcao Industrial" },
  { n: "Geraldo Menezes", c: "Geraldo Celso Cunha de Menezes", e: "", p: "IBPC", f: "", g: "E - Construcao Industrial" },
  { n: "Elissandra Silva", c: "Elissandra Candido Alves Silva", e: "", p: "KVG", f: "", g: "E - Construcao Industrial" },
  { n: "Celso Pimentel", c: "Celso Pimentel Fraga Filho", e: "celso.pimentel@mip.com.br", p: "MIP Eng", f: "", g: "F - Montagem" },
  { n: "Cezar Mortari", c: "Cezar Valmor Mortari", e: "", p: "Irontec", f: "", g: "F - Montagem" },
  { n: "Fernando Lima", c: "Fernando Lima", e: "fernando.lima@montisol.com.br", p: "Montisol", f: "", g: "F - Montagem" },
  { n: "Leonardo Scarpelli", c: "Leonardo Scarpelli", e: "", p: "FDC", f: "", g: "G - Instituicoes" },
  { n: "Rodrigo Koerich", c: "Rodrigo Broering Koerich", e: "", p: "BIM Forum Brasil", f: "", g: "G - Instituicoes" },
  { n: "Oscar Simonsen", c: "Oscar Simonsen", e: "", p: "ABEMI", f: "", g: "G - Instituicoes" }
];

if (db.prepare("SELECT COUNT(*) as c FROM gts").get().c === 0) {
  var gtId = "cie-coic";
  db.prepare("INSERT INTO gts (id, nome, descricao) VALUES (?, ?, ?)").run(gtId, "CIE-COIC", "Comite de Inteligencia Estrategica da COIC");
  var ins = db.prepare("INSERT INTO membros (id, gt_id, nome, nome_completo, email, telefone, empresa, funcao, grupo) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)");
  for (var i = 0; i < SEED_MEMBROS.length; i++) {
    var m = SEED_MEMBROS[i];
    ins.run(uid(), gtId, m.n, m.c, m.e, "", m.p, m.f, m.g);
  }
  console.log("CIE-COIC criado com 21 membros");
}

module.exports = {
  getGTs: function () {
    var gts = db.prepare("SELECT * FROM gts ORDER BY created_at").all();
    for (var i = 0; i < gts.length; i++) {
      gts[i].membros_count = db.prepare("SELECT COUNT(*) as c FROM membros WHERE gt_id = ? AND ativo = 1").get(gts[i].id).c;
    }
    return gts;
  },

  createGT: function (nome, descricao) {
    var id = uid();
    db.prepare("INSERT INTO gts (id, nome, descricao) VALUES (?, ?, ?)").run(id, nome, descricao || "");
    return { id: id, nome: nome, descricao: descricao };
  },

  updateGT: function (id, nome, descricao) {
    db.prepare("UPDATE gts SET nome = ?, descricao = ? WHERE id = ?").run(nome, descricao || "", id);
  },

  deleteGT: function (id) {
    db.prepare("DELETE FROM gts WHERE id = ?").run(id);
  },

  getMembros: function (gtId) {
    return db.prepare("SELECT * FROM membros WHERE gt_id = ? AND ativo = 1 ORDER BY grupo, nome").all(gtId);
  },

  addMembro: function (gtId, d) {
    var id = uid();
    db.prepare("INSERT INTO membros (id, gt_id, nome, nome_completo, email, telefone, empresa, funcao, grupo) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)").run(id, gtId, d.nome, d.nome_completo || d.nome, d.email || "", d.telefone || "", d.empresa || "", d.funcao || "", d.grupo || "");
    return { id: id, nome: d.nome, empresa: d.empresa };
  },

  updateMembro: function (id, d) {
    db.prepare("UPDATE membros SET nome = ?, nome_completo = ?, email = ?, telefone = ?, empresa = ?, funcao = ?, grupo = ? WHERE id = ?").run(d.nome, d.nome_completo || d.nome, d.email || "", d.telefone || "", d.empresa || "", d.funcao || "", d.grupo || "", id);
  },

  deleteMembro: function (id) {
    db.prepare("UPDATE membros SET ativo = 0 WHERE id = ?").run(id);
  },

  getReunioes: function (gtId) {
    return db.prepare("SELECT * FROM reunioes WHERE gt_id = ? ORDER BY data DESC").all(gtId);
  },

  createReuniao: function (gtId, d) {
    var id = uid();
    db.prepare("INSERT INTO reunioes (id, gt_id, nome, edicao, data, horario, local_) VALUES (?, ?, ?, ?, ?, ?, ?)").run(id, gtId, d.nome, d.edicao || "", d.data, d.horario || "", d.local || "");
    return { id: id, gt_id: gtId };
  },

  registerPresencas: function (reuniaoId, presencas) {
    var stmt = db.prepare("INSERT INTO presencas (reuniao_id, membro_id, presente) VALUES (?, ?, ?) ON CONFLICT(reuniao_id, membro_id) DO UPDATE SET presente = ?");
    var tx = db.transaction(function () {
      var keys = Object.keys(presencas);
      for (var i = 0; i < keys.length; i++) {
        var val = presencas[keys[i]] ? 1 : 0;
        stmt.run(reuniaoId, keys[i], val, val);
      }
    });
    tx();
  },

  getQuorum: function (gtId, ano) {
    var membros = db.prepare("SELECT * FROM membros WHERE gt_id = ? AND ativo = 1 ORDER BY grupo, nome").all(gtId);
    var totalReq = db.prepare("SELECT COUNT(*) as c FROM reunioes WHERE gt_id = ? AND data LIKE ?").get(gtId, "%" + ano + "%");
    var totalReunioes = (totalReq && totalReq.c) || 0;

    var result = [];
    for (var i = 0; i < membros.length; i++) {
      var m = membros[i];
      var presReq = db.prepare("SELECT COUNT(*) as c FROM presencas p JOIN reunioes r ON p.reuniao_id = r.id WHERE p.membro_id = ? AND p.presente = 1 AND r.data LIKE ?").get(m.id, "%" + ano + "%");
      var pres = (presReq && presReq.c) || 0;
      var pct = totalReunioes > 0 ? Math.round((pres / totalReunioes) * 100) : 0;
      result.push({
        nome: m.nome_completo || m.nome,
        entidade: m.empresa,
        total_reunioes: String(Math.max(totalReunioes, 1)),
        presencas: String(pres),
        percentual: pct + "%"
      });
    }
    return result;
  },

  getSettings: function () {
    var rows = db.prepare("SELECT * FROM settings").all();
    var obj = {};
    for (var i = 0; i < rows.length; i++) {
      obj[rows[i].key] = rows[i].value;
    }
    return obj;
  },

  saveSettings: function (settings) {
    var stmt = db.prepare("INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = ?");
    var tx = db.transaction(function () {
      var keys = Object.keys(settings);
      for (var i = 0; i < keys.length; i++) {
        stmt.run(keys[i], settings[keys[i]], settings[keys[i]]);
      }
    });
    tx();
  }
};
