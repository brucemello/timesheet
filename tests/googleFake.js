// Simulação mínima dos serviços do Google Apps Script para os testes em Node.
const fs = require('fs'), vm = require('vm'), path = require('path');
const pad = n => String(n).padStart(2, '0');
const fmtBR = d => `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
const exibir = v => v instanceof Date ? fmtBR(v) : (v == null ? '' : String(v));

class Planilha {
  constructor(nome, linhas) { this.nome = nome; this.dados = (linhas || []).map(r => r.slice()); }
  getName() { return this.nome; }
  getSheetId() { return this.nome.length; }
  getLastRow() { let n = this.dados.length; while (n && this.dados[n - 1].every(v => v === '' || v == null)) n--; return n; }
  getLastColumn() { return Math.max(0, ...this.dados.map(r => { let n = r.length; while (n && (r[n - 1] === '' || r[n - 1] == null)) n--; return n; })); }
  celula(r, c) { return (this.dados[r] || [])[c]; }
  set(r, c, v) { while (this.dados.length <= r) this.dados.push([]); const l = this.dados[r]; while (l.length <= c) l.push(''); l[c] = v; }
  getRange(r, c, nr, nc) {
    nr = nr || 1; nc = nc || 1; const sh = this;
    const ler = f => { const o = []; for (let i = 0; i < nr; i++) { const l = []; for (let j = 0; j < nc; j++) { const v = sh.celula(r - 1 + i, c - 1 + j); l.push(f(v == null ? '' : v)); } o.push(l); } return o; };
    const rg = {
      getValues: () => ler(v => v), getDisplayValues: () => ler(exibir), getValue: () => ler(v => v)[0][0],
      setValues: m => { m.forEach((l, i) => l.forEach((v, j) => sh.set(r - 1 + i, c - 1 + j, v))); return rg; },
      setValue: v => { sh.set(r - 1, c - 1, v); return rg; },
      clearContent: () => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (sh.dados[r - 1 + i]) sh.set(r - 1 + i, c - 1 + j, ''); return rg; }
    };
    ['setBackground', 'setFontColor', 'setFontWeight', 'setNumberFormat', 'setBackgrounds', 'setFontColors'].forEach(k => rg[k] = () => rg);
    return rg;
  }
  appendRow(l) { this.dados.splice(this.getLastRow(), 0, l.slice()); }
  clear() { this.dados = []; }
  autoResizeColumns() {}
}
class Pasta {
  constructor(abas, id) { this.abas = abas; this.id = id || 'ID_PLANILHA'; }
  getSheets() { return this.abas; }
  getSheetByName(n) { return this.abas.find(a => a.getName() === n) || null; }
  insertSheet(n) { const a = new Planilha(n, []); this.abas.push(a); return a; }
  getId() { return this.id; } getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id; } getName() { return 'Planilha ' + this.id; }
}

function criarContexto(opts) {
  const props = Object.assign({}, opts.props || {});
  const cache = {};
  const emails = [];
  const gatilhos = [];
  const ctx = {
    console, emails, gatilhos, props,
    Logger: { log: () => {} },
    Session: { getScriptTimeZone: () => 'America/Sao_Paulo', getActiveUser: () => ({ getEmail: () => ctx.usuarioAtivo }), getEffectiveUser: () => ({ getEmail: () => 'bruce.mello@totvs.com.br' }) },
    usuarioAtivo: 'bruce.mello@totvs.com.br',
    Utilities: { formatDate: (d, tz, f) => f === 'yyyy-MM-dd' ? `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`
      : f === 'dd/MM/yyyy HH:mm' ? `${fmtBR(d)} ${pad(d.getHours())}:${pad(d.getMinutes())}` : fmtBR(d) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] == null ? null : props[k], setProperty: (k, v) => { props[k] = String(v); }, deleteProperty: k => { delete props[k]; } }) },
    CacheService: { getScriptCache: () => ({ get: k => cache[k] == null ? null : cache[k], putAll: m => Object.assign(cache, m), getAll: ks => { const o = {}; ks.forEach(k => { if (k in cache) o[k] = cache[k]; }); return o; }, remove: k => { delete cache[k]; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    MailApp: { sendEmail: m => emails.push(m) },
    ScriptApp: {
      getProjectTriggers: () => gatilhos.slice(), deleteTrigger: t => gatilhos.splice(gatilhos.indexOf(t), 1),
      newTrigger: f => { const b = { timeBased: () => b, onWeekDay: () => b, atHour: () => b, create: () => { gatilhos.push({ getHandlerFunction: () => f }); } }; return b; },
      getScriptId: () => 'SCRIPT123',
      WeekDay: { MONDAY: 1, TUESDAY: 2, WEDNESDAY: 3, THURSDAY: 4, FRIDAY: 5 }, getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/X/exec' })
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => opts.ativa,
      openById: id => opts.externas[id] || (() => { throw new Error('sem acesso ' + id); })(),
      openByUrl: u => { const k = Object.keys(opts.externas).find(k => u.indexOf(k) >= 0); if (!k) throw new Error('sem acesso ' + u); return opts.externas[k]; },
      getUi: () => { throw new Error('sem UI'); }
    },
    HtmlService: { createHtmlOutputFromFile: () => { const o = { setTitle: () => o, addMetaTag: () => o, setXFrameOptionsMode: () => o, setWidth: () => o, setHeight: () => o }; return o; }, XFrameOptionsMode: { ALLOWALL: 1 } },
    UrlFetchApp: { fetch: () => { throw new Error('rede desligada nos testes'); } }
  };
  vm.createContext(ctx);
  const SRC = path.join(__dirname, '..', 'src');
  fs.readdirSync(SRC).filter(f => f.endsWith('.js')).sort().forEach(f => vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f }));
  return ctx;
}
module.exports = { Planilha, Pasta, criarContexto };
