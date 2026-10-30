/********************************************************************************
 * PAINEL DE CONTROLE (Web App + modal na planilha)
 * Tudo que o Painel.html chama via google.script.run fica aqui.
 * Funções do painel devolvem { sucesso, mensagem, ... } e nunca objetos Date.
 ********************************************************************************/

/** Web App */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Painel')
    .setTitle('Timesheet · Painel de Controle')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Modal dentro da planilha (menu) */
function abrirPainelLateral() {
  obterPlanilha_(); // memoriza o ID para o Web App
  const html = HtmlService.createHtmlOutputFromFile('Painel').setWidth(1400).setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, 'Timesheet · Painel de Controle');
}

function respostaPainel_(fn) {
  try {
    const r = fn();
    return Object.assign({ sucesso: true }, r || {});
  } catch (e) {
    Logger.log('Painel: ' + (e && e.stack || e));
    return { sucesso: false, mensagem: String(e && e.message || e) };
  }
}

/* =============================================================================
 * DADOS DO DASHBOARD
 * ===========================================================================*/
function obterDadosPainel(params) {
  return respostaPainel_(() => {
    params = params || {};
    const apur = apurarPeriodo_({ mesDe: params.mesDe, mesAte: params.mesAte, usarCache: !params.forcar });
    let log = [];
    try { log = lerLogCobrancas_(300); } catch (e) { log = []; }
    const ultima = {};
    log.forEach(l => { if (!ultima[l.email] && !/teste/i.test(l.tipo)) ultima[l.email] = l.quando; });
    apur.devs.forEach(d => d.ultimaCobranca = ultima[d.email] || '');

    const cfg = obterConfiguracoesAtuais();
    const admin = ehAdmin_();
    let urlPlanilha = '';
    if (admin) { try { urlPlanilha = obterPlanilha_().getUrl(); } catch (e) { urlPlanilha = ''; } }
    return {
      apuracao: apur,
      logCobrancas: log.slice(0, 60),
      contexto: {
        emailLider: cfg.emailLider, periodoAuto: cfg.periodoAuto, mesPadraoDe: cfg.mesDe, mesPadraoAte: cfg.mesAte,
        urlPlanilha: urlPlanilha, urlWebApp: obterUrlWebApp_(), admin: admin
      }
    };
  });
}

function obterUrlWebApp_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

/* =============================================================================
 * ÁREA DO ADMINISTRADOR
 * Admin = e-mails da propriedade CONFIG_ADMINS (separados por vírgula) ou,
 * se ela não existir, o dono da implantação (quem publicou o Web App).
 * Os links só saem do servidor para admin.
 * ===========================================================================*/
const CHAVE_ADMINS = 'CONFIG_ADMINS';
const REPO_GITHUB = 'https://github.com/brucemello/timesheet';

function usuarioAtual_() {
  try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) { return ''; }
}

function listaAdmins_() {
  const prop = PropertiesService.getScriptProperties().getProperty(CHAVE_ADMINS);
  if (prop) return prop.split(/[,;\s]+/).map(x => x.trim().toLowerCase()).filter(Boolean);
  try { const dono = String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); return dono ? [dono] : []; } catch (e) { return []; }
}

function ehAdmin_() {
  const u = usuarioAtual_();
  return !!u && listaAdmins_().indexOf(u) >= 0;
}

function exigirAdmin_() {
  if (!ehAdmin_()) throw new Error('Acesso restrito ao administrador.');
}

function obterAreaAdminPainel() {
  return respostaPainel_(() => {
    exigirAdmin_();
    const ss = obterPlanilha_();
    const scriptId = ScriptApp.getScriptId();
    const proj = `https://script.google.com/home/projects/${scriptId}`;
    const aba = nome => { const a = ss.getSheetByName(nome); return a ? `${ss.getUrl()}#gid=${a.getSheetId()}` : ''; };
    return {
      usuario: usuarioAtual_(),
      admins: listaAdmins_(),
      grupos: [
        { titulo: 'Planilha', links: [
          { icone: '📗', nome: 'Planilha do Timesheet', desc: ss.getName(), url: ss.getUrl() },
          { icone: '📧', nome: 'Log de cobranças', desc: 'Aba Log_Cobrancas', url: aba(ABA_LOG_COBRANCAS) },
          { icone: '👥', nome: 'Colaboradores (DePara)', desc: 'Aba DePara_Usuarios', url: aba('DePara_Usuarios') },
          { icone: '⚠️', nome: 'Horas pendentes', desc: 'Aba Horas_Pendentes', url: aba('Horas_Pendentes') }
        ].filter(l => l.url) },
        { titulo: 'Apps Script', links: [
          { icone: '🧩', nome: 'Editor de scripts', desc: 'Código do projeto', url: `${proj}/edit` },
          { icone: '📜', nome: 'Execuções', desc: 'Logs e erros das rotinas', url: `${proj}/executions` },
          { icone: '⏰', nome: 'Acionadores', desc: 'Gatilhos agendados', url: `${proj}/triggers` },
          { icone: '🚀', nome: 'Implantações', desc: 'Versões do Web App', url: `${proj}/deployments` }
        ] },
        { titulo: 'GitHub', links: [
          { icone: '🐙', nome: 'Repositório', desc: 'brucemello/timesheet', url: REPO_GITHUB },
          { icone: '⚙️', nome: 'Deploys (Actions)', desc: 'Histórico de publicações', url: `${REPO_GITHUB}/actions` }
        ] }
      ]
    };
  });
}

/* =============================================================================
 * COBRANÇA INDIVIDUAL / EM LOTE
 * ===========================================================================*/
function previewCobrancaPainel(params) {
  return respostaPainel_(() => {
    const cfg = obterConfiguracoesAtuais();
    const apur = apurarPeriodo_({ mesDe: params.mesDe, mesAte: params.mesAte, usarCache: true });
    const filtroDias = Array.isArray(params.dias) ? { [params.email]: params.dias } : null;
    const dev = pendenciasDaApuracao_(apur, [params.email], filtroDias)[0];
    if (!dev) return { html: '', mensagem: filtroDias ? 'Nenhum dos dias selecionados está pendente.' : 'Este colaborador não tem pendências no período.' };
    return {
      html: montarEmailIndividualDev_(dev, cfg.horaMinima, apur.periodo),
      para: dev.email, cc: cfg.emailLider,
      assunto: 'Lembrete de Apontamento de Horas - Pendências Identificadas',
      dias: dev.pendencias.length
    };
  });
}

/**
 * params: { emails: [], mesDe, mesAte, teste: bool }
 * Envio explícito pelo painel ignora a coluna "Cobrar" (quem escolheu foi o líder).
 */
function enviarCobrancasPainel(params) {
  return respostaPainel_(() => {
    const emails = (params && params.emails) || [];
    if (!emails.length) throw new Error('Selecione pelo menos um colaborador.');
    const cfg = obterConfiguracoesAtuais();
    if (!cfg.emailLider) throw new Error('Configure o e-mail do líder antes de enviar.');
    const apur = apurarPeriodo_({ mesDe: params.mesDe, mesAte: params.mesAte, usarCache: true });
    const devs = pendenciasDaApuracao_(apur, emails, params.dias || null);
    const semPendencia = emails.filter(e => !devs.some(d => d.email === String(e).toLowerCase()));

    const enviados = [], falhas = [];
    devs.forEach(dev => {
      try { enviarCobrancaDev_(dev, cfg, apur.periodo, (emails.length > 1 ? 'Painel (lote)' : 'Painel (individual)') + (params.dias && params.dias[dev.email] ? ' · dias escolhidos' : ''), !!params.teste); enviados.push(dev.nome); }
      catch (e) { falhas.push(`${dev.nome}: ${e.message}`); }
    });

    let msg = enviados.length
      ? (params.teste ? `Teste enviado para ${cfg.emailLider} (ref.: ${enviados.join(', ')}).` : `Cobrança enviada para ${enviados.join(', ')} com cópia para ${cfg.emailLider}.`)
      : 'Nenhum e-mail enviado.';
    if (semPendencia.length) msg += ` Sem pendência no período: ${semPendencia.length}.`;
    if (falhas.length) msg += ` Falhas: ${falhas.join(' | ')}`;
    return { mensagem: msg, enviados: enviados.length, falhas: falhas, log: lerLogCobrancas_(60) };
  });
}

/**
 * Dispensa dias da cobrança de forma permanente (painel, abas e rotinas automáticas).
 * params: { email, dias: ['yyyy-MM-dd'], motivo }
 */
function dispensarDiasPainel(params) {
  return respostaPainel_(() => {
    const email = String(params && params.email || '').trim().toLowerCase();
    const dias = ((params && params.dias) || []).filter(d => /^\d{4}-\d{2}-\d{2}$/.test(d));
    if (!email || !dias.length) throw new Error('Escolha pelo menos um dia para dispensar.');
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let n;
    try { n = gravarDispensas_(obterPlanilha_(), email, dias, String(params.motivo || '').trim()); }
    finally { lock.releaseLock(); }
    limparCacheApuracao_();
    return { mensagem: n ? `${n} dia(s) dispensado(s): não serão mais cobrados, nem pelas rotinas automáticas.` : 'Esses dias já estavam dispensados.' };
  });
}

/* =============================================================================
 * AÇÕES / ROTINAS
 * ===========================================================================*/
function executarAcaoPainel(acao) {
  return respostaPainel_(() => {
    const mapa = {
      extrair: () => extrairEGravarNaPlanilha().mensagem,
      relatorioLider: () => executarRelatorioDiario17h(),
      alertaSemanal: () => executarVerificacaoEAlertaSemanais(false),
      testeLider: () => executarVerificacaoEAlertaSemanais(true),
      ativarRotinas: () => { configurarGatilhosAutomaticos(); return 'Rotinas automáticas ativadas.'; },
      desativarRotinas: () => { removerGatilhosRotina_(); return 'Rotinas automáticas desativadas.'; },
      limparCache: () => { limparCacheApuracao_(); return 'Cache limpo: a próxima consulta busca tudo de novo no Tempo.'; }
    };
    if (!mapa[acao]) throw new Error(`Ação desconhecida: ${acao}`);
    const msg = mapa[acao]();
    return { mensagem: typeof msg === 'string' ? msg : 'Concluído.', rotinas: listarRotinas_() };
  });
}

const FUNCOES_ROTINA = ['rodarExtracaoMatutina7h', 'enviarRelatorioMatutino9h', 'executarDisparoSemanalOficial', 'executarRelatorioDiario17h', 'executarVerificacaoEAlertaSemanais'];

function removerGatilhosRotina_() {
  ScriptApp.getProjectTriggers().forEach(t => { if (FUNCOES_ROTINA.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
}

function listarRotinas_() {
  const nomes = {
    rodarExtracaoMatutina7h: 'Extração dos dados (07h, seg–sex)',
    enviarRelatorioMatutino9h: 'Relatório para o líder (09h, seg–sex)',
    executarDisparoSemanalOficial: 'Cobrança semanal dos devs (sexta, 14h)',
    executarRelatorioDiario17h: 'Relatório diário 17h (antigo)',
    executarVerificacaoEAlertaSemanais: 'Alerta semanal (antigo)'
  };
  const cont = {};
  try {
    ScriptApp.getProjectTriggers().forEach(t => { const f = t.getHandlerFunction(); cont[f] = (cont[f] || 0) + 1; });
  } catch (e) { /* sem permissão */ }
  return Object.keys(nomes).filter(f => cont[f] || FUNCOES_ROTINA.indexOf(f) < 3)
    .map(f => ({ funcao: f, nome: nomes[f], gatilhos: cont[f] || 0 }));
}

/* =============================================================================
 * CONFIGURAÇÕES
 * ===========================================================================*/
function obterConfigPainel() {
  return respostaPainel_(() => {
    const cfg = obterConfiguracoesAtuais();
    const props = PropertiesService.getScriptProperties();
    const tk = cfg.jiraToken || '';
    return {
      config: {
        jiraBaseUrl: cfg.jiraBaseUrl, teamId: cfg.teamId, horaMinima: cfg.horaMinima, emailLider: cfg.emailLider,
        periodoAuto: cfg.periodoAuto,
        mesDe: props.getProperty(CHAVES_CONFIG.MES_DE) || cfg.mesDe, mesAte: props.getProperty(CHAVES_CONFIG.MES_ATE) || cfg.mesAte,
        urlFerias: cfg.urlFerias, urlAgenda: cfg.urlAgenda,
        tokenConfigurado: !!tk, tokenFinal: tk ? tk.slice(-4) : ''
      },
      rotinas: listarRotinas_(),
      feriadosAuto: (() => { const a = new Date().getFullYear(); return feriadosDoAno_(a).concat(feriadosDoAno_(a + 1)).sort((x, y) => x.data.localeCompare(y.data)); })()
    };
  });
}

function salvarConfigPainel(c) {
  return respostaPainel_(() => {
    c = c || {};
    const props = PropertiesService.getScriptProperties();
    const url = String(c.jiraBaseUrl || '').trim().replace(/\/$/, '');
    if (!url || !/^https?:\/\//i.test(url)) throw new Error('Informe a URL do Jira (https://...).');
    if (!String(c.teamId || '').trim()) throw new Error('Informe o Team ID do Tempo.');
    const hm = parseFloat(String(c.horaMinima).replace(',', '.'));
    if (!(hm > 0 && hm <= 24)) throw new Error('Hora mínima diária inválida.');
    if (c.emailLider && !/@/.test(c.emailLider)) throw new Error('E-mail do líder inválido.');
    if (!c.periodoAuto && (!/^\d{4}-\d{2}$/.test(c.mesDe || '') || !/^\d{4}-\d{2}$/.test(c.mesAte || ''))) throw new Error('Informe os meses do período fixo.');

    props.setProperty(CHAVES_CONFIG.JIRA_BASE_URL, url);
    props.setProperty(CHAVES_CONFIG.TEAM_ID, String(c.teamId).trim());
    if (String(c.jiraToken || '').trim()) props.setProperty(CHAVES_CONFIG.JIRA_TOKEN, String(c.jiraToken).trim());
    props.setProperty(CHAVES_CONFIG.HORA_MINIMA, String(hm));
    props.setProperty(CHAVES_CONFIG.EMAIL_LIDER, String(c.emailLider || '').trim().toLowerCase());
    props.setProperty(CHAVES_CONFIG.PERIODO_AUTO, c.periodoAuto ? 'SIM' : 'NAO');
    if (c.mesDe) props.setProperty(CHAVES_CONFIG.MES_DE, c.mesDe);
    if (c.mesAte) props.setProperty(CHAVES_CONFIG.MES_ATE, c.mesAte);
    ['urlFerias', 'urlAgenda'].forEach(k => {
      const chave = k === 'urlFerias' ? CHAVES_CONFIG.URL_FERIAS : CHAVES_CONFIG.URL_AGENDA;
      const v = String(c[k] || '').trim();
      if (v) props.setProperty(chave, v); else props.deleteProperty(chave);
    });
    limparCacheApuracao_();
    return { mensagem: 'Configurações salvas.' };
  });
}

/** Testa o acesso ao Tempo e às planilhas de férias/agenda. */
function testarConexoesPainel() {
  return respostaPainel_(() => {
    const cfg = obterConfiguracoesAtuais();
    const r = [];
    try {
      const hoje = formatarDataISO_(new Date());
      const wl = buscarWorklogsTempoAPI_(hoje.substring(0, 8) + '01', hoje, cfg);
      r.push(`✅ Tempo: ${wl.length} apontamento(s) no mês`);
    } catch (e) { r.push('❌ Tempo: ' + e.message); }
    try { r.push('✅ Férias: ' + abrirPlanilhaRef_(cfg.urlFerias).getName()); } catch (e) { r.push('❌ Planilha de férias: ' + e.message); }
    try { r.push('✅ Agenda: ' + abrirPlanilhaRef_(cfg.urlAgenda).getName()); } catch (e) { r.push('❌ Agenda de ausências: ' + e.message); }
    return { mensagem: r.join(' · '), ok: r.every(x => x.indexOf('✅') === 0) };
  });
}

/* =============================================================================
 * CADASTROS (abas da planilha editadas pelo painel)
 * ===========================================================================*/
const CADASTROS_PAINEL = {
  colaboradores: {
    aba: 'DePara_Usuarios', titulo: 'Colaboradores',
    colunas: [
      { key: 'idJira', label: 'ID Jira (Username / AccountId)' },
      { key: 'nome', label: 'Nome Real do Colaborador', obrig: true },
      { key: 'email', label: 'E-mail Corporativo', obrig: true },
      { key: 'apelidos', label: 'Apelidos (agenda)' },
      { key: 'cobrar', label: 'Cobrar (SIM/NÃO)', opcoes: ['SIM', 'NÃO'] }
    ]
  },
  feriados: {
    aba: 'Feriados', titulo: 'Feriados',
    colunas: [{ key: 'data', label: 'Data', data: true, obrig: true }, { key: 'descricao', label: 'Descrição' }]
  },
  dispensas: {
    aba: ABA_DISPENSAS, titulo: 'Dias dispensados',
    colunas: [
      { key: 'email', label: 'E-mail do colaborador', obrig: true },
      { key: 'data', label: 'Data dispensada', data: true, obrig: true },
      { key: 'motivo', label: 'Motivo' },
      { key: 'por', label: 'Dispensado por' },
      { key: 'quando', label: 'Quando' }
    ]
  },
  ausencias: {
    aba: 'Ausencias', titulo: 'Ausências manuais',
    colunas: [
      { key: 'identificador', label: 'Identificador (nome ou e-mail)', obrig: true },
      { key: 'inicio', label: 'Início', data: true, obrig: true },
      { key: 'fim', label: 'Fim', data: true },
      { key: 'horas', label: 'Horas (parcial, opcional)', num: true },
      { key: 'motivo', label: 'Motivo' }
    ]
  }
};

function obterCadastroPainel(nome) {
  return respostaPainel_(() => {
    const def = CADASTROS_PAINEL[nome];
    if (!def) throw new Error('Cadastro desconhecido: ' + nome);
    const ss = obterPlanilha_();
    if (nome === 'colaboradores') carregarMapaUsuarios_(ss); // garante a aba e as colunas novas
    const aba = ss.getSheetByName(def.aba);
    const linhas = [];
    if (aba && aba.getLastRow() > 1) {
      const n = aba.getLastRow() - 1, nc = def.colunas.length;
      const valores = aba.getRange(2, 1, n, nc).getValues();
      const exibidos = aba.getRange(2, 1, n, nc).getDisplayValues();
      for (let i = 0; i < n; i++) {
        const o = {}; let vazio = true;
        def.colunas.forEach((c, j) => {
          let v;
          if (c.data) { const d = lerDataAus_(valores[i][j] instanceof Date ? valores[i][j] : exibidos[i][j]); v = d ? isoAus_(d) : ''; }
          else v = String(exibidos[i][j] == null ? '' : exibidos[i][j]).trim();
          if (v) vazio = false;
          o[c.key] = v;
        });
        if (!vazio) linhas.push(o);
      }
    }
    if (nome === 'colaboradores') linhas.forEach(l => { if (!l.cobrar) l.cobrar = 'SIM'; });
    return { colunas: def.colunas, linhas: linhas, titulo: def.titulo };
  });
}

function salvarCadastroPainel(nome, linhas) {
  return respostaPainel_(() => {
    const def = CADASTROS_PAINEL[nome];
    if (!def) throw new Error('Cadastro desconhecido: ' + nome);
    linhas = (linhas || []).filter(l => def.colunas.some(c => String(l[c.key] || '').trim()));
    linhas.forEach((l, i) => def.colunas.forEach(c => {
      if (c.obrig && !String(l[c.key] || '').trim()) throw new Error(`Linha ${i + 1}: preencha "${c.label}".`);
      if (c.data && l[c.key] && !/^\d{4}-\d{2}-\d{2}$/.test(l[c.key])) throw new Error(`Linha ${i + 1}: data inválida em "${c.label}".`);
    }));
    if (nome === 'colaboradores') {
      const vistos = {};
      linhas.forEach(l => {
        const e = String(l.email).trim().toLowerCase();
        if (!/@/.test(e)) throw new Error(`E-mail inválido: ${l.email}`);
        if (vistos[e]) throw new Error(`E-mail repetido: ${e}`);
        vistos[e] = true;
      });
    }
    if (nome === 'ausencias') linhas.forEach((l, i) => { if (l.fim && l.fim < l.inicio) throw new Error(`Linha ${i + 1}: fim antes do início.`); });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const ss = obterPlanilha_();
      const aba = ss.getSheetByName(def.aba) || ss.insertSheet(def.aba);
      const nc = def.colunas.length;
      aba.getRange(1, 1, 1, nc).setValues([def.colunas.map(c => c.label)]);
      estilizarCabecalho_(aba, 1, nc, '#16213e');
      if (aba.getLastRow() > 1) aba.getRange(2, 1, aba.getLastRow() - 1, Math.max(nc, aba.getLastColumn())).clearContent();
      if (linhas.length) {
        const matriz = linhas.map(l => def.colunas.map(c => {
          const v = String(l[c.key] == null ? '' : l[c.key]).trim();
          if (c.data) return v ? isoParaDataAus_(v) : '';
          if (c.num) return v === '' ? '' : parseFloat(v.replace(',', '.')) || '';
          if (c.key === 'email') return v.toLowerCase();
          return v;
        }));
        aba.getRange(2, 1, matriz.length, nc).setValues(matriz);
        def.colunas.forEach((c, j) => { if (c.data) aba.getRange(2, j + 1, matriz.length, 1).setNumberFormat('dd/MM/yyyy'); });
      }
    } finally {
      lock.releaseLock();
    }
    limparCacheApuracao_();
    return { mensagem: `${def.titulo}: ${linhas.length} linha(s) salvas.` };
  });
}
