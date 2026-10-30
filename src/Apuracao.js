/********************************************************************************
 * APURAÇÃO DO PERÍODO — fonte única para painel, abas da planilha e e-mails
 *
 * Junta: apontamentos do Tempo + DePara_Usuarios + ausências/férias/folgas
 * (Ausencias.js) + feriados, e calcula para cada dev e cada dia útil:
 *   horas apontadas, meta do dia, status e ausência considerada.
 ********************************************************************************/

const CHAVE_ID_PLANILHA = 'CONFIG_SPREADSHEET_ID';
const ABA_LOG_COBRANCAS = 'Log_Cobrancas';
const ABA_DISPENSAS = 'Dias_Dispensados';
const CABECALHO_DISPENSAS = ['E-mail do colaborador', 'Data dispensada', 'Motivo', 'Dispensado por', 'Quando'];

/** Dias que o líder avaliou e decidiu não cobrar: { 'email|yyyy-MM-dd': motivo }. */
function lerDispensas_(ss) {
  const mapa = {};
  const aba = ss.getSheetByName(ABA_DISPENSAS);
  if (!aba || aba.getLastRow() < 2) return mapa;
  const n = aba.getLastRow() - 1;
  const val = aba.getRange(2, 1, n, 3).getValues(), exib = aba.getRange(2, 1, n, 3).getDisplayValues();
  for (let i = 0; i < n; i++) {
    const email = String(exib[i][0] || '').trim().toLowerCase();
    const d = lerDataAus_(val[i][1] instanceof Date ? val[i][1] : exib[i][1]);
    if (email && d) mapa[email + '|' + isoAus_(d)] = String(exib[i][2] || '').trim();
  }
  return mapa;
}

/** Grava dias dispensados (sem repetir os que já existem). */
function gravarDispensas_(ss, email, dias, motivo) {
  let aba = ss.getSheetByName(ABA_DISPENSAS);
  if (!aba) {
    aba = ss.insertSheet(ABA_DISPENSAS);
    aba.getRange(1, 1, 1, CABECALHO_DISPENSAS.length).setValues([CABECALHO_DISPENSAS]);
    estilizarCabecalho_(aba, 1, CABECALHO_DISPENSAS.length, '#0f766e');
  }
  const ja = lerDispensas_(ss);
  let usuario = '';
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) { usuario = ''; }
  const quando = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const novos = dias.filter(iso => !((email + '|' + iso) in ja))
    .map(iso => [email, isoParaDataAus_(iso), motivo || '', usuario, quando]);
  if (novos.length) {
    aba.getRange(aba.getLastRow() + 1, 1, novos.length, 5).setValues(novos);
    aba.getRange(2, 2, aba.getLastRow() - 1, 1).setNumberFormat('dd/MM/yyyy');
  }
  return novos.length;
}

/** Planilha do timesheet: a ativa (menu/gatilho) ou a memorizada (Web App). */
function obterPlanilha_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (ss) {
    try { if (props.getProperty(CHAVE_ID_PLANILHA) !== ss.getId()) props.setProperty(CHAVE_ID_PLANILHA, ss.getId()); } catch (e) { /* ok */ }
    return ss;
  }
  const id = props.getProperty(CHAVE_ID_PLANILHA);
  if (!id) throw new Error('Abra a planilha uma vez (menu ⚙️ Painel Tempo API) para vincular o painel Web App.');
  return SpreadsheetApp.openById(id);
}

/** Identifica o dev do worklog usando a DePara_Usuarios. */
function resolverDevDoWorklog_(wl, mapaUsuarios) {
  const info = extrairIdentificadoresColaborador_(wl);
  const chaveUsername = String(info.username || '').trim().toLowerCase();
  const chaveAccountId = String(info.accountId || '').trim().toLowerCase();
  const chaveEmail = String(info.email || '').trim().toLowerCase();
  const chaveNome = String(info.nome || '').trim().toLowerCase();

  const dp = mapaUsuarios[chaveUsername] || mapaUsuarios[chaveAccountId] || mapaUsuarios[chaveEmail] ||
             mapaUsuarios[chaveEmail.split('@')[0]] || mapaUsuarios[chaveNome];

  let nome = (dp && dp.nome) ? dp.nome : (info.nome || info.email);
  let email = (dp && dp.email) ? dp.email : info.email;
  if (!email) email = chaveUsername || chaveAccountId || 'usuario.desconhecido';
  if (email && email.indexOf('@') < 0) email += '@totvs.com.br';
  email = email.toLowerCase();
  if (!nome || nome.indexOf('@') >= 0) nome = (dp && dp.nome) ? dp.nome : email;
  return { email: email, nome: nome };
}

/**
 * Apura o período.
 * @param {{mesDe?:string, mesAte?:string, usarCache?:boolean}} opcoes
 */
function apurarPeriodo_(opcoes) {
  opcoes = opcoes || {};
  const cfg = obterConfiguracoesAtuais();
  if (!cfg.jiraToken) throw new Error('Token da API do Tempo não configurado (Painel > Configurações).');

  let mesDe = opcoes.mesDe || cfg.mesDe, mesAte = opcoes.mesAte || cfg.mesAte;
  if (mesAte < mesDe) { const t = mesDe; mesDe = mesAte; mesAte = t; }

  const versao = PropertiesService.getScriptProperties().getProperty('CACHE_VERSAO') || '0';
  const chaveCache = `APUR_${versao}_${mesDe}_${mesAte}_${cfg.teamId}_${cfg.horaMinima}`;
  if (opcoes.usarCache) {
    const c = lerCacheGrande_(chaveCache);
    if (c) return c;
  }

  const ss = obterPlanilha_();
  const mapaUsuarios = carregarMapaUsuarios_(ss);
  const feriados = carregarFeriados_(ss);
  const periodo = calcularDatasDoPeriodo_(mesDe, mesAte);
  const horaMinima = parseFloat(cfg.horaMinima) || 6;
  const hojeIso = formatarDataISO_(new Date());

  // Devs: todos da DePara + quem apontou no período
  const devsInfo = {};
  Object.keys(mapaUsuarios).forEach(k => {
    const i = mapaUsuarios[k];
    if (i.email && i.nome) devsInfo[i.email] = { nome: i.nome, cobrar: i.cobrar !== false, idJira: i.idJira || '' };
  });

  const horasPorDevDia = {};
  const worklogs = buscarWorklogsTempoAPI_(periodo.dataDeIso, periodo.dataAteIso, cfg) || [];
  worklogs.forEach(wl => {
    const dev = resolverDevDoWorklog_(wl, mapaUsuarios);
    const dataWl = String(wl.started || wl.dateStarted || '').substring(0, 10);
    if (!dataWl || !dev.email) return;
    const k = dev.email + '|' + dataWl;
    horasPorDevDia[k] = (horasPorDevDia[k] || 0) + (wl.timeSpentSeconds || 0) / 3600;
    if (!devsInfo[dev.email]) devsInfo[dev.email] = { nome: dev.nome, cobrar: true, foraDePara: true };
  });

  // Dias de semana do período (feriados marcados)
  const dias = [];
  for (let d = new Date(periodo.inicio); d <= periodo.fim; d.setDate(d.getDate() + 1)) {
    if (d.getDay() === 0 || d.getDay() === 6) continue;
    const iso = formatarDataISO_(d);
    dias.push({ iso: iso, br: formatarDataBR_(d), dow: d.getDay(), feriado: feriados.has(iso) ? feriados.get(iso) : '' });
  }

  const emails = Object.keys(devsInfo).sort((a, b) => devsInfo[a].nome.localeCompare(devsInfo[b].nome));
  const ausencias = carregarAusencias_(ss, emails.map(e => ({ email: e, nome: devsInfo[e].nome })));
  const dispensas = lerDispensas_(ss);

  const devs = emails.map(email => {
    const info = devsInfo[email];
    const t = { esperadas: 0, apontadas: 0, saldo: 0, diasPendentes: 0, horasFaltantes: 0, diasAusencia: 0, horasAusencia: 0, diasUteis: 0, diasOk: 0, diasDispensados: 0 };
    const diasDev = dias.map(dia => {
      const horas = Math.round((horasPorDevDia[email + '|' + dia.iso] || 0) * 100) / 100;
      if (dia.feriado) return { h: horas, m: 0, s: 'FER', a: dia.feriado };
      const chaveDisp = email + '|' + dia.iso;
      if (chaveDisp in dispensas) {
        if (dia.iso < hojeIso) t.diasDispensados++;
        return { h: horas, m: 0, s: 'DISP', a: 'Dispensado' + (dispensas[chaveDisp] ? ': ' + dispensas[chaveDisp] : '') };
      }
      const av = avaliarStatusEAusenciaDev_(email, info.nome, lerDataAus_(dia.iso), ausencias, horaMinima);
      const desc = descreverAusenciaDia_(av);
      const ehHoje = dia.iso >= hojeIso;
      let s;
      if (av.eDiaInteiro) s = 'AUS';
      else if (ehHoje) s = horas >= av.cargaEsperadaNoDia ? 'OK' : 'HOJE';
      else s = horas >= av.cargaEsperadaNoDia ? 'OK' : 'PEND';

      t.apontadas += horas;
      if (!ehHoje) {
        t.diasUteis++;
        if (av.eDiaInteiro) t.diasAusencia++;
        else {
          t.esperadas += av.cargaEsperadaNoDia;
          t.horasAusencia += av.horasAbatidas;
          if (s === 'OK') t.diasOk++;
          if (s === 'PEND') { t.diasPendentes++; t.horasFaltantes += av.cargaEsperadaNoDia - horas; }
        }
      }
      return { h: horas, m: av.cargaEsperadaNoDia, s: s, a: desc, p: av.horasAbatidas > 0 };
    });
    t.saldo = t.apontadas - t.esperadas;
    Object.keys(t).forEach(k => t[k] = Math.round(t[k] * 100) / 100);
    return {
      email: email, nome: info.nome, cobrar: info.cobrar, foraDePara: !!info.foraDePara,
      dias: diasDev, totais: t, status: t.diasPendentes > 0 ? 'PENDENTE' : 'OK'
    };
  });

  const nomePorEmail = {};
  devs.forEach(d => nomePorEmail[d.email] = d.nome);

  const resultado = {
    periodo: { mesDe: mesDe, mesAte: mesAte, de: periodo.dataDeIso, ate: periodo.dataAteIso, hoje: hojeIso },
    horaMinima: horaMinima,
    dias: dias,
    devs: devs,
    ausencias: {
      registros: ausencias.registros.map(r => ({
        email: r.email, nome: nomePorEmail[r.email] || r.email, inicio: r.inicio, fim: r.fim, tipo: r.tipo,
        horas: r.horas, horario: r.horarioTexto, origem: r.origem, titulo: r.titulo
      })),
      naoVinculados: ausencias.naoVinculados,
      avisos: ausencias.avisos
    },
    totalWorklogs: worklogs.length,
    geradoEm: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  };
  gravarCacheGrande_(chaveCache, resultado, 1800);
  return resultado;
}

/** Pendências por dev a partir da apuração (somente dias úteis já encerrados). */
/**
 * @param {Object=} diasPorEmail opcional { email: ['yyyy-MM-dd', ...] } — cobra só esses dias
 *   (email ausente no mapa = todos os dias pendentes)
 */
function pendenciasDaApuracao_(apur, filtroEmails, diasPorEmail) {
  const filtro = filtroEmails ? filtroEmails.map(e => String(e).toLowerCase()) : null;
  const dias = {};
  Object.keys(diasPorEmail || {}).forEach(e => { if (Array.isArray(diasPorEmail[e])) dias[String(e).toLowerCase()] = diasPorEmail[e]; });
  return apur.devs
    .filter(d => !filtro || filtro.indexOf(d.email) >= 0)
    .map(d => ({
      nome: d.nome, email: d.email, cobrar: d.cobrar,
      horasFaltantes: d.totais.horasFaltantes,
      pendencias: d.dias.map((x, i) => ({ x: x, dia: apur.dias[i] }))
        .filter(o => o.x.s === 'PEND' && (!dias[d.email] || dias[d.email].indexOf(o.dia.iso) >= 0))
        .map(o => ({ data: o.dia.br, iso: o.dia.iso, horas: o.x.h, meta: o.x.m, ausencia: o.x.a }))
    }))
    .filter(d => d.pendencias.length > 0);
}

/* =============================================================================
 * ENVIO DE COBRANÇA (individual, em lote, semanal) + LOG
 * ===========================================================================*/
function enviarCobrancaDev_(dev, cfg, periodo, tipo, somenteLider) {
  const prefixo = somenteLider ? `[TESTE - Ref: ${dev.nome}] ` : '';
  MailApp.sendEmail({
    to: somenteLider ? cfg.emailLider : dev.email,
    cc: somenteLider ? '' : cfg.emailLider,
    subject: `${prefixo}Lembrete de Apontamento de Horas - Pendências Identificadas`,
    htmlBody: montarEmailIndividualDev_(dev, cfg.horaMinima, periodo)
  });
  registrarCobranca_(dev, periodo, somenteLider ? 'Teste (só líder)' : tipo);
}

function registrarCobranca_(dev, periodo, tipo) {
  try {
    const ss = obterPlanilha_();
    let aba = ss.getSheetByName(ABA_LOG_COBRANCAS);
    if (!aba) {
      aba = ss.insertSheet(ABA_LOG_COBRANCAS);
      aba.getRange(1, 1, 1, 8).setValues([['Data/Hora', 'Dev / Analista', 'E-mail', 'Período', 'Dias Cobrados', 'Horas Faltantes', 'Tipo', 'Enviado por']]);
      estilizarCabecalho_(aba, 1, 8, '#16213e');
    }
    const horas = dev.pendencias.reduce((s, p) => s + Math.max(0, (p.meta || 0) - p.horas), 0);
    let usuario = '';
    try { usuario = Session.getActiveUser().getEmail(); } catch (e) { usuario = ''; }
    aba.appendRow([
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      dev.nome, dev.email, `${periodo.de} a ${periodo.ate}`,
      dev.pendencias.map(p => p.data.substring(0, 5)).join(', '),
      Math.round(horas * 100) / 100, tipo, usuario || 'automático'
    ]);
  } catch (e) {
    Logger.log('Não foi possível registrar a cobrança: ' + e);
  }
}

function lerLogCobrancas_(limite) {
  const ss = obterPlanilha_();
  const aba = ss.getSheetByName(ABA_LOG_COBRANCAS);
  if (!aba || aba.getLastRow() < 2) return [];
  const n = aba.getLastRow() - 1;
  const ini = Math.max(2, aba.getLastRow() - (limite || 200) + 1);
  return aba.getRange(ini, 1, aba.getLastRow() - ini + 1, 8).getDisplayValues().reverse().map(r => ({
    quando: r[0], nome: r[1], email: String(r[2]).toLowerCase(), periodo: r[3], dias: r[4], horas: r[5], tipo: r[6], por: r[7]
  }));
}

/* =============================================================================
 * CACHE (CacheService aceita até 100KB por chave: divide em pedaços)
 * ===========================================================================*/
function gravarCacheGrande_(chave, obj, segundos) {
  try {
    const cache = CacheService.getScriptCache();
    const txt = JSON.stringify(obj);
    const tam = 90000, partes = Math.ceil(txt.length / tam);
    if (partes > 20) return;
    const mapa = { [chave]: String(partes) };
    for (let i = 0; i < partes; i++) mapa[chave + '_' + i] = txt.substr(i * tam, tam);
    cache.putAll(mapa, segundos || 1800);
  } catch (e) { Logger.log('Cache: ' + e); }
}

function lerCacheGrande_(chave) {
  try {
    const cache = CacheService.getScriptCache();
    const n = parseInt(cache.get(chave), 10);
    if (!n) return null;
    const chaves = []; for (let i = 0; i < n; i++) chaves.push(chave + '_' + i);
    const vals = cache.getAll(chaves);
    if (chaves.some(k => vals[k] == null)) return null;
    return JSON.parse(chaves.map(k => vals[k]).join(''));
  } catch (e) { return null; }
}

/** Invalida todas as apurações em cache (muda a versão usada na chave). */
function limparCacheApuracao_() {
  try { PropertiesService.getScriptProperties().setProperty('CACHE_VERSAO', Date.now().toString(36) + Math.random().toString(36).slice(2, 8)); } catch (e) { /* ok */ }
}
