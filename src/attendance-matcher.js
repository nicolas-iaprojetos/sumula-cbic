var fs = require("fs");
var path = require("path");

function norm(s) {
  return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
}

function cleanName(s) {
  return s
    .replace(/\s*\([^)]*\)\s*/g, " ")
    .replace(/^[A-Z]{2,10}\s*-\s*/i, "")
    .replace(/\s*\|[^|]*$/g, "")
    .replace(/\s+-\s+[A-Z][A-Z\-]+\s*$/g, "")
    .replace(/\s+(SME|SINDUSCON[\w-]*)\s*$/gi, "")
    .trim();
}

function tokens(s) {
  return norm(s).split(/[\s\-\.]+/).filter(function (t) {
    return t.length > 1 && ["de", "da", "do", "dos", "das", "e"].indexOf(t) === -1;
  });
}

function overlap(a, b) {
  var h = 0;
  for (var i = 0; i < a.length; i++) {
    for (var j = 0; j < b.length; j++) {
      if (a[i] === b[j] || (a[i].length > 3 && b[j].length > 3 && (a[i].indexOf(b[j]) >= 0 || b[j].indexOf(a[i]) >= 0))) {
        h++;
        break;
      }
    }
  }
  return h;
}

function parseCSV(text) {
  var lines = text.split(/\r?\n/);
  var rows = [];
  var inParticipants = false;
  var headerFound = false;
  var emailCol = -1;

  for (var i = 0; i < lines.length; i++) {
    var cols = lines[i].split("\t");
    if (!inParticipants) {
      if (/participantes/i.test(cols[0])) { inParticipants = true; continue; }
      if (/^nome$/i.test((cols[0] || "").trim())) { inParticipants = true; headerFound = true; emailCol = cols.findIndex(function (c) { return /email/i.test(c); }); continue; }
    } else if (!headerFound) {
      if (/^nome$/i.test((cols[0] || "").trim())) { headerFound = true; emailCol = cols.findIndex(function (c) { return /email/i.test(c); }); continue; }
    } else {
      var name = (cols[0] || "").trim();
      if (name && name.length > 1) {
        rows.push({ name: name, email: emailCol >= 0 ? (cols[emailCol] || "").trim() : "" });
      }
    }
  }

  if (rows.length === 0) {
    for (var j = 0; j < lines.length; j++) {
      var c = lines[j].split(/[,;\t]/);
      var n = (c[0] || "").replace(/"/g, "").trim();
      if (n && n.length > 2 && !/^(nome|name)/i.test(n)) {
        rows.push({ name: n, email: "" });
      }
    }
  }

  var deduped = new Map();
  for (var k = 0; k < rows.length; k++) {
    var key = norm(rows[k].name);
    if (!deduped.has(key)) deduped.set(key, rows[k]);
    else if (rows[k].email && !deduped.get(key).email) deduped.get(key).email = rows[k].email;
  }
  return Array.from(deduped.values());
}

async function matchAttendance(filePath, membros) {
  var buf = fs.readFileSync(filePath);
  var raw = (buf[0] === 0xFF && buf[1] === 0xFE) ? buf.toString("utf16le") : buf.toString("utf-8");

  var rows = parseCSV(raw);
  var presencas = {};
  var matched = new Set();
  membros.forEach(function (m) { presencas[m.id] = false; });

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r.email) continue;
    var eml = r.email.toLowerCase().trim();
    for (var j = 0; j < membros.length; j++) {
      var m = membros[j];
      if (matched.has(m.id)) continue;
      if (m.email && m.email.toLowerCase().trim() === eml) {
        presencas[m.id] = true;
        matched.add(m.id);
        r._matched = true;
        break;
      }
    }
  }

  for (var i2 = 0; i2 < rows.length; i2++) {
    var r2 = rows[i2];
    if (r2._matched) continue;
    var cleaned = cleanName(r2.name);
    var rowTokens = tokens(cleaned);
    if (!rowTokens.length) continue;

    var bestMatch = null;
    var bestScore = 0;
    for (var j2 = 0; j2 < membros.length; j2++) {
      var m2 = membros[j2];
      if (matched.has(m2.id)) continue;
      var names = [m2.nome, m2.nome_completo].filter(Boolean);
      for (var k2 = 0; k2 < names.length; k2++) {
        var mTokens = tokens(names[k2]);
        if (!mTokens.length) continue;
        var olap = overlap(rowTokens, mTokens);
        var score = olap / Math.min(rowTokens.length, mTokens.length);
        if (olap >= 2 || (score >= 0.8 && olap >= 1)) {
          if (score > bestScore) { bestScore = score; bestMatch = m2; }
        }
      }
    }

    if (bestMatch) {
      presencas[bestMatch.id] = true;
      matched.add(bestMatch.id);
      r2._matched = true;
    }
  }

  var bots = /fireflies|read\.ai|notetaker|bot/i;
  var seen = new Set();
  var unmatched = [];
  for (var i3 = 0; i3 < rows.length; i3++) {
    var r3 = rows[i3];
    if (r3._matched) continue;
    var cl = cleanName(r3.name);
    if (!cl || cl.length < 2 || bots.test(r3.name)) continue;
    var nk = norm(cl);
    if (seen.has(nk)) continue;
    seen.add(nk);
    unmatched.push({ name: cl, email: r3.email || "" });
  }

  return {
    presencas: presencas,
    unmatched: unmatched,
    totalParticipants: rows.length,
    matchedCount: matched.size
  };
}

module.exports = { matchAttendance: matchAttendance };
