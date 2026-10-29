/********************************************************************************
 * PROJETO: Automação TimeSheet + Extração com De-Para, Férias, Ausências e Alertas (API Tempo)
 * DASHBOARD EXECUTIVO ESTILO TOTVS + CÁLCULO DE META PROPORCIONAL POR AUSÊNCIA
 ********************************************************************************/

const URL_PLANILHA_FERIAS = 'https://docs.google.com/spreadsheets/d/1ox7bG8Z1NqSJVINDF2EjAq00XpjnbvSyD4SapgXvjD8/edit';
const URL_PLANILHA_AGENDA_AUSENCIAS = 'https://docs.google.com/spreadsheets/d/1oZTQ3XhiYgTmQFtDKcrXKHTKRWf8sagUfCsJi-V-t2A/edit';

const CHAVES_CONFIG = {
  JIRA_BASE_URL: 'CONFIG_JIRA_BASE_URL',
  TEAM_ID: 'CONFIG_TEAM_ID',
  JIRA_TOKEN: 'CONFIG_JIRA_TOKEN',
  HORA_MINIMA: 'CONFIG_HORA_MINIMA',
  EMAIL_LIDER: 'CONFIG_EMAIL_LIDER',
  MES_DE: 'CONFIG_MES_DE',
  MES_ATE: 'CONFIG_MES_ATE'
};

/********************************************************************************
 * GATILHO SIMPLES: EXECUTADO AUTOMATICAMENTE AO ABRIR A PLANILHA
 ********************************************************************************/
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('⚙️ Painel Tempo API')
    .addItem('🚀 Abrir Dashboard Executivo', 'abrirPainelLateral')
    .addSeparator()
    .addItem('📊 Extrair e Salvar na Planilha', 'extrairEGravarNaPlanilha')
    .addItem('📧 Disparar Relatório Diário 17h (Manual)', 'executarRelatorioDiario17h')
    .addItem('🚀 Executar Alertas Semanal Devs (Manual)', 'executarDisparoSemanalOficial')
    .addItem('👤 Gerar/Verificar Aba DePara_Usuarios', 'criarAbaDeParaManualmente')
    .addItem('🏖️ Conferir Ausências e Férias Lidas', 'testarIntegracaoAusenciasEAgenda')
    .addItem('⏰ Reconfigurar Gatilhos Automáticos', 'configurarGatilhosAutomaticos')
    .addToUi();

  abrirPainelLateral();
}

/**
 * Abre o Dashboard Executivo Centralizado (Janela Modal)
 */
function abrirPainelLateral() {
  const html = HtmlService.createHtmlOutputFromFile('Interface')
      .setTitle('Dashboard Agente Jira IA - Controle Executivo')
      .setWidth(1020)
      .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Agente Jira IA - Controle Executivo');
}

/********************************************************************************
 * GESTÃO DE CONFIGURAÇÕES E INTEGRAÇÃO COM INTERFACE
 ********************************************************************************/
function obterConfiguracoesAtuais() {
  const props = PropertiesService.getScriptProperties();
  const hoje = new Date();
  const anoAtual = hoje.getFullYear();
  const mesAtual = String(hoje.getMonth() + 1).padStart(2, '0');
  const mesPadrao = `${anoAtual}-${mesAtual}`;

  return {
    jiraBaseUrl: props.getProperty(CHAVES_CONFIG.JIRA_BASE_URL) || 'https://jira.totvs.com.br',
    teamId: props.getProperty(CHAVES_CONFIG.TEAM_ID) || '644',
    jiraToken: props.getProperty(CHAVES_CONFIG.JIRA_TOKEN) || '',
    horaMinima: props.getProperty(CHAVES_CONFIG.HORA_MINIMA) || '6',
    emailLider: props.getProperty(CHAVES_CONFIG.EMAIL_LIDER) || Session.getActiveUser().getEmail(),
    mesDe: props.getProperty(CHAVES_CONFIG.MES_DE) || mesPadrao,
    mesAte: props.getProperty(CHAVES_CONFIG.MES_ATE) || mesPadrao
  };
}

function salvarConfiguracoesPelaInterface(config) {
  try {
    if (!config.jiraBaseUrl || !config.teamId || !config.jiraToken) {
      throw new Error('Preencha os campos obrigatórios: URL do Jira, Team ID e Token API.');
    }

    const props = PropertiesService.getScriptProperties();
    props.setProperty(CHAVES_CONFIG.JIRA_BASE_URL, config.jiraBaseUrl.trim().replace(/\/$/, ''));
    props.setProperty(CHAVES_CONFIG.TEAM_ID, String(config.teamId).trim());
    props.setProperty(CHAVES_CONFIG.JIRA_TOKEN, config.jiraToken.trim());
    props.setProperty(CHAVES_CONFIG.HORA_MINIMA, String(config.horaMinima || 6).trim());
    props.setProperty(CHAVES_CONFIG.EMAIL_LIDER, config.emailLider.trim().toLowerCase());
    props.setProperty(CHAVES_CONFIG.MES_DE, config.mesDe);
    props.setProperty(CHAVES_CONFIG.MES_ATE, config.mesAte);

    configurarGatilhosAutomaticos();

    return { sucesso: true, mensagem: 'Configurações e gatilhos automáticos salvos com sucesso!' };
  } catch (erro) {
    return { sucesso: false, mensagem: 'Erro ao salvar: ' + erro.message };
  }
}

function executarFuncaoPeloPainel(nomeFuncao) {
  try {
    if (nomeFuncao === 'extrairEGravarNaPlanilha') {
      const res = extrairEGravarNaPlanilha();
      const textoMsg = (typeof res === 'object' && res.mensagem) ? res.mensagem : res;
      return { sucesso: true, mensagem: textoMsg };
    } 
    else if (nomeFuncao === 'executarVerificacaoEAlertaSemanais' || nomeFuncao === 'executarDisparoSemanalOficial') {
      const res = executarDisparoSemanalOficial();
      return { sucesso: true, mensagem: res };
    } 
    else if (nomeFuncao === 'executarRelatorioDiario17h') {
      const res = executarRelatorioDiario17h();
      return { sucesso: true, mensagem: res };
    }
    else if (nomeFuncao === 'testarEnvioApenasParaLider') {
      const res = executarVerificacaoEAlertaSemanais(true);
      return { sucesso: true, mensagem: res };
    }
    
    throw new Error(`A função '${nomeFuncao}' não está mapeada.`);
  } catch (erro) {
    return { sucesso: false, mensagem: erro.toString() };
  }
}

/********************************************************************************
 * MAPA DE-PARA DE USUÁRIOS
 ********************************************************************************/
function criarAbaDeParaManualmente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  carregarMapaUsuarios_(ss);
  SpreadsheetApp.getUi().alert('Aba "DePara_Usuarios" criada/verificada com sucesso!');
}

function carregarMapaUsuarios_(ss) {
  const mapa = {};
  let aba = ss.getSheetByName('DePara_Usuarios');

  if (!aba) {
    aba = ss.insertSheet('DePara_Usuarios');
    const dadosModelo = [
      ['ID Jira (Username / AccountId)', 'Nome Real do Colaborador', 'E-mail Corporativo'],
      ['jirauser102776', 'Yasmim Lira Castro', 'yasmim.castro@totvs.com.br'],
      ['jirauser103914', 'Eduardo Nunes Cirqueira', 'eduardo.cirqueira@totvs.com.br'],
      ['jirauser103915', 'Keilla Silva de Freitas', 'keilla.freitas@totvs.com.br'],
      ['jirauser104037', 'William Palma Alves', 'william.alves@totvs.com.br'],
      ['jirauser104319', 'Caique Henrique Carlos', 'caique.carlos@totvs.com.br'],
      ['jirauser82702', 'Evandro Italo Silva Peixoto', 'evandro.peixoto@totvs.com.br']
    ];

    aba.getRange(1, 1, dadosModelo.length, dadosModelo[0].length).setValues(dadosModelo);
    aba.getRange(1, 1, 1, 3)
      .setBackground('#16213e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    aba.autoResizeColumns(1, 3);
  }

  if (aba.getLastRow() <= 1) return mapa;

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 3).getValues();
  dados.forEach(linha => {
    const idJira = String(linha[0] || '').trim().toLowerCase();
    const nomeReal = String(linha[1] || '').trim();
    const emailReal = String(linha[2] || '').trim().toLowerCase();

    const infoObj = { nome: nomeReal, email: emailReal };

    if (idJira) mapa[idJira] = infoObj;
    if (emailReal) {
      mapa[emailReal] = infoObj;
      const prefixoEmail = emailReal.split('@')[0];
      mapa[prefixoEmail] = infoObj;
    }
  });

  return mapa;
}

/* Ausências, férias, folgas e feriados: ver Ausencias.js */

/********************************************************************************
 * EXTRAÇÃO E SALVAMENTO COM TRATAMENTO DE AUSÊNCIAS (FÉRIAS + AUDITORIA DE AGENDA)
 ********************************************************************************/
function extrairEGravarNaPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapaUsuarios = carregarMapaUsuarios_(ss);
  const feriados = carregarFeriados_(ss);

  const cfg = obterConfiguracoesAtuais();
  if (!cfg.jiraToken) throw new Error('Token da API não configurado no Painel.');

  const periodo = calcularDatasDoPeriodo_(cfg.mesDe, cfg.mesAte);
  const worklogs = buscarWorklogsTempoAPI_(periodo.dataDeIso, periodo.dataAteIso, cfg);

  const horaMinimaNum = parseFloat(cfg.horaMinima) || 6;
  const horasPorDevDia = {};
  const infoDevs = {};

  Object.keys(mapaUsuarios).forEach(chave => {
    const item = mapaUsuarios[chave];
    if (item.email && item.nome) {
      infoDevs[item.email.toLowerCase()] = item.nome;
    }
  });

  if (worklogs && worklogs.length > 0) {
    worklogs.forEach(wl => {
      const info = extrairIdentificadoresColaborador_(wl);
      const chaveUsername = String(info.username || '').trim().toLowerCase();
      const chaveAccountId = String(info.accountId || '').trim().toLowerCase();
      const chaveEmail = String(info.email || '').trim().toLowerCase();
      const chaveNome = String(info.nome || '').trim().toLowerCase();

      const deParaEncontrado = mapaUsuarios[chaveUsername] || 
                               mapaUsuarios[chaveAccountId] || 
                               mapaUsuarios[chaveEmail] || 
                               mapaUsuarios[chaveEmail.split('@')[0]] ||
                               mapaUsuarios[chaveNome];

      let nomeFinal = (deParaEncontrado && deParaEncontrado.nome) ? deParaEncontrado.nome : (info.nome || info.email);
      let emailFinal = (deParaEncontrado && deParaEncontrado.email) ? deParaEncontrado.email : info.email;

      if (!emailFinal) emailFinal = chaveUsername || chaveAccountId || 'usuario.desconhecido';
      if (emailFinal && !emailFinal.includes('@')) emailFinal += '@totvs.com.br';
      emailFinal = emailFinal.toLowerCase();

      if (!nomeFinal || nomeFinal.includes('@')) {
        nomeFinal = (deParaEncontrado && deParaEncontrado.nome) ? deParaEncontrado.nome : emailFinal;
      }

      const dataWl = (wl.started || wl.dateStarted || '').substring(0, 10);
      const horas = (wl.timeSpentSeconds || 0) / 3600;

      if (dataWl && emailFinal) {
        const chaveAgrupamento = `${emailFinal}|${dataWl}`;
        horasPorDevDia[chaveAgrupamento] = (horasPorDevDia[chaveAgrupamento] || 0) + horas;
        infoDevs[emailFinal] = nomeFinal;
      }
    });
  }

  const hojeIso = formatarDataISO_(new Date());
  const diasUteis = [];
  let cursor = new Date(periodo.inicio);
  while (cursor <= periodo.fim) {
    const iso = formatarDataISO_(cursor);
    if (cursor.getDay() !== 0 && cursor.getDay() !== 6 && !feriados.has(iso)) {
      diasUteis.push(iso);
    }
    cursor.setDate(cursor.getDate() + 1);
  }

  const listaEmailsOrdenados = Object.keys(infoDevs).sort((a, b) => {
    const nomeA = infoDevs[a] || a;
    const nomeB = infoDevs[b] || b;
    return nomeA.localeCompare(nomeB);
  });

  const ausencias = carregarAusencias_(ss, listaEmailsOrdenados.map(e => ({ email: e, nome: infoDevs[e] || e })));

  const registrosProcessados = [];
  const resumoConsolidadoDevs = [];
  const pendenciasDetalhadas = [];
  const matrizAuditoria = [
    ['Data', 'Dev / Analista', 'E-mail', 'Tipo Ausência', 'Título / Motivo Evento', 'Horário Registrado', 'Abatimento / Horas', 'Origem']
  ];

  const matrizExtracao = [
    ['Data', 'Dev / Analista', 'E-mail', 'Total Horas no Dia', `Status (Mín. ${horaMinimaNum}h)`, 'Meta do Dia', 'Ausência Considerada']
  ];

  listaEmailsOrdenados.forEach(email => {
    let nome = infoDevs[email] || email;
    if (mapaUsuarios[email] && mapaUsuarios[email].nome) {
      nome = mapaUsuarios[email].nome;
    }

    let totalHorasApontadasDev = 0;
    let diasFeriasDev = 0;
    let totalHorasAusenciaAbatidas = 0;
    let horasEsperadasDevTotal = 0;

    diasUteis.forEach(dataIso => {
      const dataObj = lerDataAus_(dataIso);
      const avaliacao = avaliarStatusEAusenciaDev_(email, nome, dataObj, ausencias, horaMinimaNum);
      const ehHoje = dataIso >= hojeIso; // dia em andamento: mostra, mas não cobra

      const chaveAgrupamento = `${email}|${dataIso}`;
      const totalHorasDia = horasPorDevDia[chaveAgrupamento] || 0;
      totalHorasApontadasDev += totalHorasDia;
      if (!ehHoje) totalHorasAusenciaAbatidas += avaliacao.horasAbatidas;

      // Alimenta dados para a aba de Auditoria
      if (avaliacao.eventosAuditoria && avaliacao.eventosAuditoria.length > 0) {
        avaliacao.eventosAuditoria.forEach(evt => {
          matrizAuditoria.push([
            formatarDataBR_(dataObj),
            nome,
            email,
            evt.tipo,
            evt.titulo,
            evt.horarioTexto,
            evt.eDiaInteiro ? 'Dia Inteiro' : Number(evt.horasAbatimento.toFixed(2)),
            evt.origem || ''
          ]);
        });
      }

      let status = 'PENDENTE';

      if (avaliacao.eDiaInteiro) {
        status = 'FÉRIAS / AUSÊNCIA';
        if (!ehHoje) diasFeriasDev++;
      } else if (ehHoje) {
        status = totalHorasDia >= avaliacao.cargaEsperadaNoDia ? 'OK' : 'EM ANDAMENTO';
      } else {
        horasEsperadasDevTotal += avaliacao.cargaEsperadaNoDia;
        if (totalHorasDia >= avaliacao.cargaEsperadaNoDia) {
          status = 'OK';
        }
      }

      matrizExtracao.push([dataIso, nome, email, Number(totalHorasDia.toFixed(2)), status,
        avaliacao.cargaEsperadaNoDia, descreverAusenciaDia_(avaliacao)]);
      registrosProcessados.push({ 
        dataIso: dataIso, 
        dataBr: formatarDataBR_(dataObj), 
        nome: nome, 
        email: email, 
        horas: totalHorasDia, 
        status: status 
      });

      if (status === 'PENDENTE') {
        pendenciasDetalhadas.push([dataIso, nome, email, Number(totalHorasDia.toFixed(2)), 'PENDENTE', avaliacao.cargaEsperadaNoDia, descreverAusenciaDia_(avaliacao)]);
      }
    });

    const saldoHorasDev = totalHorasApontadasDev - horasEsperadasDevTotal;

    resumoConsolidadoDevs.push({
      nome: nome,
      email: email,
      diasUteisMeta: diasUteis.filter(d => d < hojeIso).length,
      diasFerias: diasFeriasDev,
      horasAusencia: Number(totalHorasAusenciaAbatidas.toFixed(2)),
      horasEsperadas: Number(horasEsperadasDevTotal.toFixed(2)),
      horasApontadas: Number(totalHorasApontadasDev.toFixed(2)),
      saldoHoras: Number(saldoHorasDev.toFixed(2)),
      statusGeral: saldoHorasDev >= 0 ? 'OK' : 'PENDENTE'
    });
  });

  // ABA 1: Resumo_Mensal_Devs
  let abaResumo = ss.getSheetByName('Resumo_Mensal_Devs');
  if (!abaResumo) {
    abaResumo = ss.insertSheet('Resumo_Mensal_Devs');
  } else {
    abaResumo.clear();
  }
  
  const matrizResumo = [['Dev / Analista', 'E-mail', 'Dias Úteis (Até Ontem)', 'Dias Férias/Ausência', 'Horas Ausência (Abatidas)', 'Horas Esperadas', 'Horas Apontadas', 'Saldo (Horas)', 'Status Geral']];
  resumoConsolidadoDevs.forEach(d => {
    matrizResumo.push([d.nome, d.email, d.diasUteisMeta, d.diasFerias, d.horasAusencia, d.horasEsperadas, d.horasApontadas, d.saldoHoras, d.statusGeral]);
  });

  abaResumo.getRange(1, 1, matrizResumo.length, matrizResumo[0].length).setValues(matrizResumo);
  estilizarCabecalho_(abaResumo, 1, matrizResumo[0].length, '#16213e');
  abaResumo.getRange(2, 5, matrizResumo.length - 1, 4).setNumberFormat('0.00"h"');
  aplicarCoresStatusGeral_(abaResumo, matrizResumo.length);
  abaResumo.autoResizeColumns(1, matrizResumo[0].length);

  // ABA 2: Horas_Pendentes
  let abaPendentes = ss.getSheetByName('Horas_Pendentes');
  if (!abaPendentes) {
    abaPendentes = ss.insertSheet('Horas_Pendentes');
  } else {
    abaPendentes.clear();
  }

  const matrizPendentes = [['Data', 'Dev / Analista', 'E-mail', 'Total Horas no Dia', 'Status', 'Meta do Dia', 'Ausência Considerada']];
  pendenciasDetalhadas.forEach(p => matrizPendentes.push(p));

  abaPendentes.getRange(1, 1, matrizPendentes.length, matrizPendentes[0].length).setValues(matrizPendentes);
  estilizarCabecalho_(abaPendentes, 1, matrizPendentes[0].length, '#c5221f');
  if (matrizPendentes.length > 1) {
    abaPendentes.getRange(2, 4, matrizPendentes.length - 1, 1).setNumberFormat('0.00"h"');
    abaPendentes.getRange(2, 5, matrizPendentes.length - 1, 1).setBackground('#fce8e6').setFontColor('#c5221f').setFontWeight('bold');
  }
  abaPendentes.autoResizeColumns(1, matrizPendentes[0].length);

  // ABA 3: Auditoria_Ausencias
  let abaAuditoria = ss.getSheetByName('Auditoria_Ausencias');
  if (!abaAuditoria) {
    abaAuditoria = ss.insertSheet('Auditoria_Ausencias');
  } else {
    abaAuditoria.clear();
  }

  abaAuditoria.getRange(1, 1, matrizAuditoria.length, matrizAuditoria[0].length).setValues(matrizAuditoria);
  estilizarCabecalho_(abaAuditoria, 1, matrizAuditoria[0].length, '#e65100');
  if (matrizAuditoria.length > 1) {
    abaAuditoria.getRange(2, 7, matrizAuditoria.length - 1, 1).setNumberFormat('0.00"h"');
  }
  abaAuditoria.autoResizeColumns(1, matrizAuditoria[0].length);

  // ABA 4: Extracao_Tempo (Bruto)
  let abaExtracao = ss.getSheetByName('Extracao_Tempo');
  if (!abaExtracao) {
    abaExtracao = ss.insertSheet('Extracao_Tempo');
  } else {
    abaExtracao.clear();
  }

  abaExtracao.getRange(1, 1, matrizExtracao.length, matrizExtracao[0].length).setValues(matrizExtracao);
  estilizarCabecalho_(abaExtracao, 1, matrizExtracao[0].length, '#1a73e8');
  abaExtracao.getRange(2, 4, matrizExtracao.length - 1, 1).setNumberFormat('0.00"h"');
  aplicarCoresStatusDiario_(abaExtracao, matrizExtracao.length);
  abaExtracao.autoResizeColumns(1, matrizExtracao[0].length);

  return {
    mensagem: `Sucesso! Abas 'Resumo_Mensal_Devs', 'Horas_Pendentes', 'Auditoria_Ausencias' e 'Extracao_Tempo' geradas com sucesso. ` +
      `${ausencias.registros.length} ausência(s)/férias consideradas.` +
      (ausencias.naoVinculados.length ? ` Sem colaborador identificado (cadastre o apelido na coluna D da DePara_Usuarios): ${ausencias.naoVinculados.join(', ')}.` : '') +
      (ausencias.avisos.length ? ` Avisos: ${ausencias.avisos.join(' | ')}` : ''),
    registros: registrosProcessados,
    resumoDevs: resumoConsolidadoDevs,
    periodoIso: `${periodo.dataDeIso} a ${periodo.dataAteIso}`
  };
}

function estilizarCabecalho_(aba, numLinhas, numColunas, corFundo) {
  aba.getRange(1, 1, numLinhas, numColunas)
    .setBackground(corFundo)
    .setFontColor('#ffffff')
    .setFontWeight('bold');
}

function aplicarCoresStatusGeral_(aba, totalLinhas) {
  if (totalLinhas <= 1) return;
  const rangeStatus = aba.getRange(2, 9, totalLinhas - 1, 1);
  const valores = rangeStatus.getValues();
  const coresFundo = [];
  const coresTexto = [];

  valores.forEach(linha => {
    if (linha[0] === 'OK') {
      coresFundo.push(['#e6f4ea']);
      coresTexto.push(['#137333']);
    } else {
      coresFundo.push(['#fce8e6']);
      coresTexto.push(['#c5221f']);
    }
  });

  rangeStatus.setBackgrounds(coresFundo).setFontColors(coresTexto).setFontWeight('bold');
}

function aplicarCoresStatusDiario_(aba, totalLinhas) {
  if (totalLinhas <= 1) return;
  const rangeStatus = aba.getRange(2, 5, totalLinhas - 1, 1);
  const valores = rangeStatus.getValues();
  const coresFundo = [];
  const coresTexto = [];

  valores.forEach(linha => {
    if (linha[0] === 'OK') {
      coresFundo.push(['#e6f4ea']);
      coresTexto.push(['#137333']);
    } else if (linha[0] === 'EM ANDAMENTO') {
      coresFundo.push(['#fef7e0']);
      coresTexto.push(['#b06000']);
    } else if (linha[0].includes('FÉRIAS') || linha[0].includes('AUSÊNCIA')) {
      coresFundo.push(['#e8f0fe']);
      coresTexto.push(['#1a73e8']);
    } else {
      coresFundo.push(['#fce8e6']);
      coresTexto.push(['#c5221f']);
    }
  });

  rangeStatus.setBackgrounds(coresFundo).setFontColors(coresTexto).setFontWeight('bold');
}

/********************************************************************************
 * RELATÓRIO DIÁRIO 17H COM BALANÇO + AGRUPAMENTO DE DIAS PENDENTES
 ********************************************************************************/
function executarRelatorioDiario17h() {
  const cfg = obterConfiguracoesAtuais();
  if (!cfg.jiraToken) throw new Error('Token da API não configurado.');

  const resultado = extrairEGravarNaPlanilha();
  const resumoDevs = resultado.resumoDevs || [];
  const registrosDiarios = resultado.registros || [];

  if (resumoDevs.length === 0) {
    return 'Nenhum dado encontrado para gerar o relatório diário.';
  }

  let totalEsperadoTime = 0;
  let totalApontadoTime = 0;
  let devsPendentes = 0;

  resumoDevs.forEach(d => {
    totalEsperadoTime += d.horasEsperadas;
    totalApontadoTime += d.horasApontadas;
    if (d.statusGeral === 'PENDENTE') devsPendentes++;
  });

  const htmlBody = montarTemplateHTMLDiario_(
    resumoDevs, 
    registrosDiarios,
    resultado.periodoIso, 
    cfg.horaMinima, 
    totalEsperadoTime, 
    totalApontadoTime, 
    devsPendentes
  );

  const assunto = devsPendentes > 0 
    ? `⚠️ [Relatório TimeSheet] ${devsPendentes} Dev(s) com Pendências - ${formatarDataBR_(new Date())}`
    : `✅ [Relatório TimeSheet] Meta de Horas Cumprida pelo Time - ${formatarDataBR_(new Date())}`;

  MailApp.sendEmail({
    to: cfg.emailLider,
    subject: assunto,
    htmlBody: htmlBody
  });

  return `Relatório Diário das 17h enviado com sucesso para ${cfg.emailLider}!`;
}

function montarTemplateHTMLDiario_(resumoDevs, registrosDiarios, periodoStr, horaMinima, totalEsperado, totalApontado, devsPendentes) {
  let linhasTabelaBalanco = '';
  let totalHorasAusenciaTime = 0;
  let totalDiasUteisTime = 0;

  resumoDevs.forEach(d => {
    const eOk = d.statusGeral === 'OK';
    const bgStatus = eOk ? '#e6f4ea' : '#fce8e6';
    const corStatus = eOk ? '#137333' : '#c5221f';
    const corSaldo = d.saldoHoras >= 0 ? '#137333' : '#c5221f';

    totalHorasAusenciaTime += d.horasAusencia + (d.diasFerias * 6);
    totalDiasUteisTime += d.diasUteisMeta;

    linhasTabelaBalanco += `
      <tr style="border-bottom: 1px solid #e0e0e0;">
        <td style="padding: 10px; font-size: 13px; font-weight: bold; color: #202124;">${d.nome}</td>
        <td style="padding: 10px; font-size: 13px; text-align: center;">${d.diasFerias > 0 ? d.diasFerias + 'd' : '-'}</td>
        <td style="padding: 10px; font-size: 13px; text-align: center; color: #e67c73;">${d.horasAusencia > 0 ? d.horasAusencia.toFixed(2).replace('.', ',') + 'h' : '-'}</td>
        <td style="padding: 10px; font-size: 13px; text-align: center; font-weight: bold; color: #5f6368;">${d.horasEsperadas.toFixed(2).replace('.', ',')}h</td>
        <td style="padding: 10px; font-size: 13px; text-align: center; font-weight: bold; color: #1a73e8;">${d.horasApontadas.toFixed(2).replace('.', ',')}h</td>
        <td style="padding: 10px; font-size: 13px; text-align: center; font-weight: bold; color: ${corSaldo};">${d.saldoHoras > 0 ? '+' : ''}${d.saldoHoras.toFixed(2).replace('.', ',')}h</td>
        <td style="padding: 10px; text-align: center;">
          <span style="background-color: ${bgStatus}; color: ${corStatus}; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: bold; display: inline-block;">
            ${d.statusGeral}
          </span>
        </td>
      </tr>
    `;
  });

  const metaBrutaTotal = totalDiasUteisTime * 6;
  const percNaoProdutividade = metaBrutaTotal > 0 ? (totalHorasAusenciaTime / metaBrutaTotal) * 100 : 0;
  const percAtingido = totalEsperado > 0 ? Math.min(100, (totalApontado / totalEsperado) * 100) : 0;

  const pendenciasPorDev = {};
  registrosDiarios.forEach(r => {
    if (r.status === 'PENDENTE') {
      if (!pendenciasPorDev[r.nome]) pendenciasPorDev[r.nome] = [];
      pendenciasPorDev[r.nome].push({ dataBr: r.dataBr, horas: r.horas });
    }
  });

  let blocoPendenciasHTML = '';
  const nomesComPendencia = Object.keys(pendenciasPorDev);

  if (nomesComPendencia.length > 0) {
    let cardsDevs = '';
    nomesComPendencia.forEach(nomeDev => {
      const listaDias = pendenciasPorDev[nomeDev];
      const badgesDias = listaDias.map(p => `
        <span style="display: inline-block; background-color: #fce8e6; border: 1px solid #fad2cf; color: #c5221f; padding: 4px 8px; border-radius: 4px; font-size: 12px; margin: 3px 4px 3px 0;">
          <b>${p.dataBr}</b> (${p.horas.toFixed(2).replace('.', ',')}h)
        </span>
      `).join('');

      cardsDevs += `
        <div style="background-color: #fafafa; border: 1px solid #e0e0e0; border-left: 4px solid #c5221f; border-radius: 6px; padding: 12px; margin-bottom: 10px;">
          <div style="font-weight: bold; font-size: 13px; color: #202124; margin-bottom: 6px;">
            👤 ${nomeDev} <span style="font-weight: normal; font-size: 11px; color: #70757a;">(${listaDias.length} dia(s) pendente(s))</span>
          </div>
          <div>${badgesDias}</div>
        </div>
      `;
    });

    blocoPendenciasHTML = `
      <div style="padding: 0 20px 20px 20px;">
        <h4 style="margin: 20px 0 12px 0; font-size: 14px; color: #c5221f; text-transform: uppercase;">
          ⚠️ Detalhamento das Datas Pendentes por Colaborador
        </h4>
        ${cardsDevs}
      </div>
    `;
  } else {
    blocoPendenciasHTML = `
      <div style="padding: 20px; text-align: center; background-color: #e6f4ea; margin: 20px; border-radius: 8px;">
        <span style="font-size: 14px; font-weight: bold; color: #137333;">🎉 Nenhuma pendência diária encontrada para o time!</span>
      </div>
    `;
  }

  return `
    <!DOCTYPE html>
    <html>
      <head><meta charset="utf-8"></head>
      <body style="font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f6f9; margin: 0; padding: 20px; color: #333;">
        <div style="max-width: 760px; margin: 0 auto; background: #ffffff; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); overflow: hidden;">
          
          <div style="background: linear-gradient(135deg, #16213e 0%, #1a73e8 100%); padding: 25px; text-align: center; color: #ffffff;">
            <h2 style="margin: 0; font-size: 22px; font-weight: 600;">📊 Consolidado Mensal de Horas</h2>
            <p style="margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;">Horas Esperadas vs. Apontadas | Período: ${periodoStr}</p>
          </div>

          <div style="display: table; width: 100%; border-bottom: 1px solid #e0e0e0; background-color: #fafafa;">
            <div style="display: table-cell; width: 25%; padding: 15px; text-align: center; border-right: 1px solid #e0e0e0;">
              <span style="font-size: 10px; color: #5f6368; font-weight: bold; text-transform: uppercase;">Meta do Time</span>
              <div style="font-size: 16px; font-weight: bold; color: #5f6368; margin-top: 4px;">${totalEsperado.toFixed(1)}h</div>
            </div>
            <div style="display: table-cell; width: 25%; padding: 15px; text-align: center; border-right: 1px solid #e0e0e0;">
              <span style="font-size: 10px; color: #5f6368; font-weight: bold; text-transform: uppercase;">Total Apontado</span>
              <div style="font-size: 16px; font-weight: bold; color: #1a73e8; margin-top: 4px;">${totalApontado.toFixed(1)}h</div>
            </div>
            <div style="display: table-cell; width: 25%; padding: 15px; text-align: center; border-right: 1px solid #e0e0e0;">
              <span style="font-size: 10px; color: #5f6368; font-weight: bold; text-transform: uppercase;">Progresso Meta</span>
              <div style="font-size: 16px; font-weight: bold; color: ${percAtingido >= 100 ? '#137333' : '#c5221f'}; margin-top: 4px;">${percAtingido.toFixed(1)}%</div>
            </div>
            <div style="display: table-cell; width: 25%; padding: 15px; text-align: center;">
              <span style="font-size: 10px; color: #e65100; font-weight: bold; text-transform: uppercase;">% Não Produtividade</span>
              <div style="font-size: 16px; font-weight: bold; color: #e65100; margin-top: 4px;">${percNaoProdutividade.toFixed(1)}%</div>
            </div>
          </div>

          <div style="padding: 20px 20px 0 20px;">
            <h4 style="margin: 0 0 15px 0; font-size: 14px; color: #16213e; text-transform: uppercase;">Balanço por Colaborador</h4>
            <table style="width: 100%; border-collapse: collapse; text-align: left;">
              <thead>
                <tr style="background-color: #f1f3f4; color: #3c4043; font-size: 11px; text-transform: uppercase;">
                  <th style="padding: 10px;">Colaborador</th>
                  <th style="padding: 10px; text-align: center;">Férias</th>
                  <th style="padding: 10px; text-align: center;">Ausência</th>
                  <th style="padding: 10px; text-align: center;">Esperado</th>
                  <th style="padding: 10px; text-align: center;">Apontado</th>
                  <th style="padding: 10px; text-align: center;">Saldo</th>
                  <th style="padding: 10px; text-align: center;">Status</th>
                </tr>
              </thead>
              <tbody>${linhasTabelaBalanco}</tbody>
            </table>
          </div>

          ${blocoPendenciasHTML}

          <div style="background-color: #f8f9fa; padding: 15px; text-align: center; font-size: 12px; color: #70757a; border-top: 1px solid #e0e0e0;">
            Carga diária de referência: <b>${horaMinima}h/dia útil</b> | Gerado automaticamente pelo <b>Painel Tempo API</b>.
          </div>

        </div>
      </body>
    </html>
  `;
}

/********************************************************************************
 * AUTOMAÇÃO SEMANAL (DISPARO OFICIAL PARA OS DEVS)
 ********************************************************************************/
function executarDisparoSemanalOficial() {
  return executarVerificacaoEAlertaSemanais(false);
}

function executarVerificacaoEAlertaSemanais(somenteLider = false) {
  const cfg = obterConfiguracoesAtuais();
  if (!cfg.jiraToken) throw new Error('Token da API não configurado.');

  const periodo = calcularDatasDoPeriodo_(cfg.mesDe, cfg.mesAte);
  const worklogs = buscarWorklogsTempoAPI_(periodo.dataDeIso, periodo.dataAteIso, cfg);
  const pendenciasPorDev = calcularPendenciasDevs_(worklogs, periodo.inicio, periodo.fim, cfg);

  if (pendenciasPorDev.length === 0) return 'Excelente! Nenhuma pendência de horas encontrada para o time.';

  pendenciasPorDev.forEach(dev => {
    const destinoEmail = somenteLider ? cfg.emailLider : dev.email;
    const ccEmail = somenteLider ? '' : cfg.emailLider;
    const prefixoAssunto = somenteLider ? `[TESTE - Ref: ${dev.nome}] ` : '';

    MailApp.sendEmail({
      to: destinoEmail,
      cc: ccEmail,
      subject: `${prefixoAssunto}Lembrete de Apontamento de Horas - Pendências Identificadas`,
      htmlBody: montarEmailIndividualDev_(dev, cfg.horaMinima)
    });
  });

  MailApp.sendEmail({
    to: cfg.emailLider,
    subject: `${somenteLider ? '[TESTE] ' : ''}Consolidado Semanal de Pendências (TimeSheet)`,
    htmlBody: montarEmailConsolidadoLider_(pendenciasPorDev, periodo.dataDeIso, periodo.dataAteIso, cfg.horaMinima)
  });

  return somenteLider 
    ? `Teste concluído! E-mails de teste enviados para ${cfg.emailLider}.`
    : `Alertas oficiais enviados com sucesso para ${pendenciasPorDev.length} dev(s) com cópia para o líder!`;
}

function buscarWorklogsTempoAPI_(dataDe, dataAte, cfg) {
  let url = `${cfg.jiraBaseUrl}/rest/tempo-timesheets/4/worklogs/search`;
  let payload = JSON.stringify({ from: dataDe, to: dataAte, teamId: [parseInt(cfg.teamId, 10)] });

  let options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': `Bearer ${cfg.jiraToken}` },
    payload: payload,
    muteHttpExceptions: true
  };

  let response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() === 404 || response.getResponseCode() === 0) {
    url = `https://api.tempo.io/v3/worklogs/team/${cfg.teamId}?from=${dataDe}&to=${dataAte}`;
    options = {
      method: 'get',
      headers: { 'Authorization': `Bearer ${cfg.jiraToken}` },
      muteHttpExceptions: true
    };
    response = UrlFetchApp.fetch(url, options);
  }

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error(`Erro na API do Tempo (${response.getResponseCode()}): ${response.getContentText()}`);
  }

  const json = JSON.parse(response.getContentText());
  return Array.isArray(json) ? json : (json.results || []);
}

/********************************************************************************
 * CÁLCULO DE PENDÊNCIAS PARA DISPARO DE ALERTAS SEMANAIS
 ********************************************************************************/
function calcularPendenciasDevs_(worklogs, dataInicio, dataFim, cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapaUsuarios = carregarMapaUsuarios_(ss);
  const feriados = carregarFeriados_(ss);
  const horaMinimaNum = parseFloat(cfg.horaMinima) || 6;

  const horasPorDevDia = {};
  const nomesDevs = {};

  worklogs.forEach(wl => {
    const info = extrairIdentificadoresColaborador_(wl);
    const chaveUsername = String(info.username || '').trim().toLowerCase();
    const chaveAccountId = String(info.accountId || '').trim().toLowerCase();
    const chaveEmail = String(info.email || '').trim().toLowerCase();
    const chaveNome = String(info.nome || '').trim().toLowerCase();

    const deParaEncontrado = mapaUsuarios[chaveUsername] || 
                             mapaUsuarios[chaveAccountId] || 
                             mapaUsuarios[chaveEmail] || 
                             mapaUsuarios[chaveEmail.split('@')[0]] ||
                             mapaUsuarios[chaveNome];

    let nomeFinal = (deParaEncontrado && deParaEncontrado.nome) ? deParaEncontrado.nome : (info.nome || info.email);
    let emailFinal = (deParaEncontrado && deParaEncontrado.email) ? deParaEncontrado.email : info.email;

    if (!emailFinal) emailFinal = chaveUsername || chaveAccountId || 'usuario.desconhecido';
    if (emailFinal && !emailFinal.includes('@')) emailFinal += '@totvs.com.br';
    emailFinal = emailFinal.toLowerCase();

    const dataWl = (wl.started || wl.dateStarted || '').substring(0, 10);
    const horas = (wl.timeSpentSeconds || 0) / 3600;

    if (emailFinal && dataWl) {
      const chave = `${emailFinal}|${dataWl}`;
      horasPorDevDia[chave] = (horasPorDevDia[chave] || 0) + horas;
      nomesDevs[emailFinal] = nomeFinal;
    }
  });

  // Devs da DePara entram mesmo sem nenhum apontamento no período
  Object.keys(mapaUsuarios).forEach(chave => {
    const item = mapaUsuarios[chave];
    if (item.email && item.nome && !nomesDevs[item.email]) nomesDevs[item.email] = item.nome;
  });
  const ausencias = carregarAusencias_(ss, Object.keys(nomesDevs).map(e => ({ email: e, nome: nomesDevs[e] })));
  const hojeIso = formatarDataISO_(new Date());

  const diasUteis = [];
  let cursor = new Date(dataInicio);
  while (cursor <= dataFim) {
    if (cursor.getDay() !== 0 && cursor.getDay() !== 6) {
      diasUteis.push(new Date(cursor));
    }
    cursor.setDate(cursor.getDate() + 1);
  }

  const pendenciasPorDev = [];

  Object.keys(nomesDevs).forEach(emailDev => {
    const nomeDev = nomesDevs[emailDev];
    const pendencias = [];

    diasUteis.forEach(dataObj => {
      const dataChave = formatarDataISO_(dataObj);
      if (feriados.has(dataChave) || dataChave >= hojeIso) return; // feriado ou dia ainda em andamento

      const avaliacao = avaliarStatusEAusenciaDev_(emailDev, nomeDev, dataObj, ausencias, horaMinimaNum);

      if (avaliacao.eDiaInteiro) return;

      const chave = `${emailDev}|${dataChave}`;
      const horas = horasPorDevDia[chave] || 0;

      if (horas < avaliacao.cargaEsperadaNoDia) {
        pendencias.push({ data: formatarDataBR_(dataObj), horas: horas, meta: avaliacao.cargaEsperadaNoDia, ausencia: descreverAusenciaDia_(avaliacao) });
      }
    });

    if (pendencias.length > 0) {
      pendenciasPorDev.push({ nome: nomeDev, email: emailDev, pendencias: pendencias });
    }
  });

  return pendenciasPorDev;
}

function extrairIdentificadoresColaborador_(wl) {
  let nome = '', email = '', username = '', accountId = '';
  const autor = wl.author || wl.worker || wl.user || {};

  if (typeof autor === 'object' && autor !== null) {
    nome = autor.displayName || autor.name || autor.fullName || '';
    email = autor.emailAddress || autor.email || '';
    username = autor.name || autor.username || autor.key || '';
    accountId = autor.accountId || autor.key || '';
  } else if (typeof autor === 'string') {
    username = autor;
  }

  if (!email && wl.authorEmail) email = wl.authorEmail;
  if (!nome && wl.authorDisplayName) nome = wl.authorDisplayName;
  if (!username && wl.authorName) username = wl.authorName;

  return { nome: nome, email: email, username: username, accountId: accountId };
}

function montarEmailIndividualDev_(item, horaMinima) {
  const linhas = item.pendencias.map(p => `
    <tr>
      <td style="padding:8px;border:1px solid #ddd;">${p.data}</td>
      <td style="padding:8px;border:1px solid #ddd;text-align:center;font-weight:bold;${p.horas === 0 ? 'background:#f8d7da;color:#721c24;' : 'background:#fff3cd;color:#856404;'}">
        ${p.horas.toFixed(2).replace('.', ',')}h
      </td>
      <td style="padding:8px;border:1px solid #ddd;text-align:center;">${formatarHorasEmail_(p.meta, horaMinima)}</td>
      <td style="padding:8px;border:1px solid #ddd;font-size:12px;color:#555;">${p.ausencia || '-'}</td>
    </tr>
  `).join('');

  return `
    <div style="font-family:Arial,sans-serif;color:#333;">
      <p>Olá, <b>${item.nome}</b>.</p>
      <p>Lembrete automático sobre o apontamento de horas no Jira/Tempo.</p>
      <p>Identificamos dias úteis concluídos com carga abaixo da meta estipulada:</p>
      <table style="border-collapse:collapse;width:100%;margin-top:10px;">
        <thead>
          <tr style="background:#f3f3f3;">
            <th style="padding:8px;border:1px solid #ddd;text-align:left;">Data</th>
            <th style="padding:8px;border:1px solid #ddd;text-align:center;">Horas Apontadas</th>
            <th style="padding:8px;border:1px solid #ddd;text-align:center;">Meta do Dia</th>
            <th style="padding:8px;border:1px solid #ddd;text-align:left;">Ausência considerada</th>
          </tr>
        </thead>
        <tbody>${linhas}</tbody>
      </table>
      <p style="margin-top:15px;font-size:12px;color:#666;">
        <i>Férias, folgas, feriados e ausências da agenda do time já foram descontados da meta. Se faltou registrar alguma ausência na agenda, avise o seu líder.</i>
      </p>
    </div>
  `;
}

function formatarHorasEmail_(meta, horaMinima) {
  const v = meta == null ? parseFloat(horaMinima) || 0 : meta;
  return v.toFixed(2).replace('.', ',') + 'h';
}

function montarEmailConsolidadoLider_(devs, dataDe, dataAte, horaMinima) {
  const blocos = devs.map(item => `
    <h3>${item.nome} (${item.email})</h3>
    <table style="border-collapse:collapse;width:100%;margin-bottom:15px;">
      <thead>
        <tr style="background:#f3f3f3;">
          <th style="padding:6px;border:1px solid #ddd;text-align:left;">Data</th>
          <th style="padding:6px;border:1px solid #ddd;text-align:center;">Horas</th>
          <th style="padding:6px;border:1px solid #ddd;text-align:center;">Meta</th>
          <th style="padding:6px;border:1px solid #ddd;text-align:left;">Ausência</th>
        </tr>
      </thead>
      <tbody>
        ${item.pendencias.map(p => `<tr><td style="padding:6px;border:1px solid #ddd;">${p.data}</td><td style="padding:6px;border:1px solid #ddd;text-align:center;">${p.horas.toFixed(2).replace('.', ',')}h</td><td style="padding:6px;border:1px solid #ddd;text-align:center;">${formatarHorasEmail_(p.meta, horaMinima)}</td><td style="padding:6px;border:1px solid #ddd;font-size:12px;">${p.ausencia || '-'}</td></tr>`).join('')}
      </tbody>
    </table>
  `).join('');

  return `<div style="font-family:Arial,sans-serif;color:#333;"><h2>Consolidado Semanal de Pendências (Tempo API)</h2><p>Período: <b>${dataDe}</b> até <b>${dataAte}</b></p>${blocos}</div>`;
}

function calcularDatasDoPeriodo_(mesDe, mesAte) {
  const pDe = mesDe.split('-');
  const pAte = mesAte.split('-');
  const inicio = new Date(parseInt(pDe[0]), parseInt(pDe[1]) - 1, 1);
  const fim = new Date(parseInt(pAte[0]), parseInt(pAte[1]), 0);
  const hoje = zerarHora_(new Date());

  // Considera até o dia de HOJE para capturar os agendamentos e apontamentos em andamento
  const dataFimEfetiva = fim > hoje ? hoje : fim;

  return {
    inicio: inicio,
    fim: dataFimEfetiva,
    dataDeIso: formatarDataISO_(inicio),
    dataAteIso: formatarDataISO_(dataFimEfetiva)
  };
}

function zerarHora_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function formatarDataISO_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function formatarDataBR_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
function removerAcentos_(texto) { return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); }
function converterParaData_(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'string') {
    const p = val.split('/');
    if (p.length === 3) return new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
  }
  return null;
}

/********************************************************************************
 * FUNÇÃO DE TESTE MANUAL DE INTEGRAÇÃO
 ********************************************************************************/
function testarIntegracaoAusenciasEAgenda() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapaUsuarios = carregarMapaUsuarios_(ss);
  const devs = {};
  Object.keys(mapaUsuarios).forEach(k => { const i = mapaUsuarios[k]; if (i.email && i.nome) devs[i.email] = i.nome; });
  const aus = carregarAusencias_(ss, Object.keys(devs).map(e => ({ email: e, nome: devs[e] })));

  const porDev = {};
  aus.registros.forEach(r => { porDev[r.email] = (porDev[r.email] || 0) + 1; });
  const linhas = Object.keys(devs).sort().map(e => `• ${devs[e]}: ${porDev[e] || 0} registro(s)`);
  const msg = `Ausências/férias vinculadas: ${aus.registros.length}\n\n${linhas.join('\n')}` +
    (aus.naoVinculados.length ? `\n\nSem colaborador identificado (cadastre o apelido na coluna D da DePara_Usuarios):\n${aus.naoVinculados.join(', ')}` : '') +
    (aus.avisos.length ? `\n\nAvisos: ${aus.avisos.join(' | ')}` : '');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* sem UI */ }
  return msg;
}


/********************************************************************************
 * ETAPA 1 (07:00) - EXTRAÇÃO E PROCESSAMENTO DOS DADOS NA PLANILHA
 ********************************************************************************/
function rodarExtracaoMatutina7h() {
  try {
    Logger.log('⏰ Iniciando extração matutina das 07:00...');
    const resultado = extrairEGravarNaPlanilha();
    
    // Salva o resultado formatado temporariamente para o envio das 09h
    const props = PropertiesService.getScriptProperties();
    props.setProperty('CACHE_RELATORIO_MATUTINO', JSON.stringify(resultado));
    
    Logger.log('✅ Extração das 07:00 concluída e dados salvos no cache.');
    return 'Extração das 07:00 executada com sucesso!';
  } catch (erro) {
    Logger.log('❌ Erro na extração das 07:00: ' + erro.toString());
    throw erro;
  }
}

/********************************************************************************
 * ETAPA 2 (09:00) - DISPARO DO E-MAIL PARA O LÍDER
 ********************************************************************************/
function enviarRelatorioMatutino9h() {
  const cfg = obterConfiguracoesAtuais();
  if (!cfg.jiraToken) throw new Error('Token da API não configurado.');

  const props = PropertiesService.getScriptProperties();
  const cacheRaw = props.getProperty('CACHE_RELATORIO_MATUTINO');
  
  let resultado;
  if (cacheRaw) {
    resultado = JSON.parse(cacheRaw);
  } else {
    // Caso a extração das 07h não tenha rodado, faz a extração na hora
    resultado = extrairEGravarNaPlanilha();
  }

  const resumoDevs = resultado.resumoDevs || [];
  const registrosDiarios = resultado.registros || [];

  if (resumoDevs.length === 0) {
    return 'Nenhum dado encontrado para enviar no relatório das 09h.';
  }

  let totalEsperadoTime = 0;
  let totalApontadoTime = 0;
  let devsPendentes = 0;

  resumoDevs.forEach(d => {
    totalEsperadoTime += d.horasEsperadas;
    totalApontadoTime += d.horasApontadas;
    if (d.statusGeral === 'PENDENTE') devsPendentes++;
  });

  const htmlBody = montarTemplateHTMLDiario_(
    resumoDevs, 
    registrosDiarios,
    resultado.periodoIso, 
    cfg.horaMinima, 
    totalEsperadoTime, 
    totalApontadoTime, 
    devsPendentes
  );

  const assunto = devsPendentes > 0 
    ? `⚠️ [Relatório Matutino 09h] ${devsPendentes} Dev(s) com Pendências - ${formatarDataBR_(new Date())}`
    : `✅ [Relatório Matutino 09h] Meta de Horas Cumprida pelo Time - ${formatarDataBR_(new Date())}`;

  MailApp.sendEmail({
    to: cfg.emailLider,
    subject: assunto,
    htmlBody: htmlBody
  });

  return `Relatório das 09:00 enviado com sucesso para ${cfg.emailLider}!`;
}

/********************************************************************************
 * RECONFIGURAÇÃO DOS GATILHOS TEMPORAIS (TRIGGERS)
 ********************************************************************************/
function configurarGatilhosAutomaticos() {
  const gatilhos = ScriptApp.getProjectTriggers();
  
  // Apaga gatilhos antigos das funções de relatórios
  gatilhos.forEach(t => {
    const nomeFn = t.getHandlerFunction();
    if (nomeFn === 'executarRelatorioDiario17h' || 
        nomeFn === 'rodarExtracaoMatutina7h' || 
        nomeFn === 'enviarRelatorioMatutino9h' || 
        nomeFn === 'executarVerificacaoEAlertaSemanais' || 
        nomeFn === 'executarDisparoSemanalOficial') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Dias Úteis da semana (Segunda a Sexta)
  const diasUteis = [
    ScriptApp.WeekDay.MONDAY,
    ScriptApp.WeekDay.TUESDAY,
    ScriptApp.WeekDay.WEDNESDAY,
    ScriptApp.WeekDay.THURSDAY,
    ScriptApp.WeekDay.FRIDAY
  ];

  // 1. Extração matutina diária às 07:00
  diasUteis.forEach(dia => {
    ScriptApp.newTrigger('rodarExtracaoMatutina7h')
      .timeBased()
      .onWeekDay(dia)
      .atHour(7)
      .create();
  });

  // 2. Envio do relatório diário ao Líder às 09:00
  diasUteis.forEach(dia => {
    ScriptApp.newTrigger('enviarRelatorioMatutino9h')
      .timeBased()
      .onWeekDay(dia)
      .atHour(9)
      .create();
  });

  // 3. Disparo Semanal de Cobrança aos Devs (Sextas-feiras às 14:00)
  ScriptApp.newTrigger('executarDisparoSemanalOficial')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(14)
    .create();

  Logger.log('✅ Gatilhos automáticos (07h Extração / 09h E-mail Líder / 14h Sexta Devs) reconfigurados com sucesso!');
}