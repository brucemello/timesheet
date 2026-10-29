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
  MES_ATE: 'CONFIG_MES_ATE',
  PERIODO_AUTO: 'CONFIG_PERIODO_AUTO',
  URL_FERIAS: 'CONFIG_URL_FERIAS',
  URL_AGENDA: 'CONFIG_URL_AGENDA'
};

/********************************************************************************
 * GATILHO SIMPLES: EXECUTADO AUTOMATICAMENTE AO ABRIR A PLANILHA
 ********************************************************************************/
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('⚙️ Painel Tempo API')
    .addItem('🚀 Abrir Painel de Controle', 'abrirPainelLateral')
    .addSeparator()
    .addItem('📊 Extrair e Salvar na Planilha', 'extrairEGravarNaPlanilha')
    .addItem('📧 Disparar Relatório Diário (Manual)', 'executarRelatorioDiario17h')
    .addItem('🚀 Executar Alertas Semanal Devs (Manual)', 'executarDisparoSemanalOficial')
    .addItem('👤 Gerar/Verificar Aba DePara_Usuarios', 'criarAbaDeParaManualmente')
    .addItem('🏖️ Conferir Ausências e Férias Lidas', 'testarIntegracaoAusenciasEAgenda')
    .addItem('⏰ Reconfigurar Gatilhos Automáticos', 'configurarGatilhosAutomaticos')
    .addToUi();

  try { obterPlanilha_(); } catch (err) { /* memoriza o ID para o Web App */ }
}

/********************************************************************************
 * CONFIGURAÇÕES (Script Properties)
 ********************************************************************************/
function obterConfiguracoesAtuais() {
  const props = PropertiesService.getScriptProperties();
  const hoje = new Date();
  const mesAtual = `${hoje.getFullYear()}-${String(hoje.getMonth() + 1).padStart(2, '0')}`;
  // Período automático (padrão): sempre o mês corrente. Desligado: usa os meses fixos salvos.
  const periodoAuto = props.getProperty(CHAVES_CONFIG.PERIODO_AUTO) !== 'NAO';
  const mesDe = periodoAuto ? mesAtual : (props.getProperty(CHAVES_CONFIG.MES_DE) || mesAtual);
  const mesAte = periodoAuto ? mesAtual : (props.getProperty(CHAVES_CONFIG.MES_ATE) || mesAtual);
  let emailPadrao = '';
  try { emailPadrao = Session.getActiveUser().getEmail(); } catch (e) { emailPadrao = ''; }

  return {
    jiraBaseUrl: props.getProperty(CHAVES_CONFIG.JIRA_BASE_URL) || 'https://jira.totvs.com.br',
    teamId: props.getProperty(CHAVES_CONFIG.TEAM_ID) || '644',
    jiraToken: props.getProperty(CHAVES_CONFIG.JIRA_TOKEN) || '',
    horaMinima: props.getProperty(CHAVES_CONFIG.HORA_MINIMA) || '6',
    emailLider: props.getProperty(CHAVES_CONFIG.EMAIL_LIDER) || emailPadrao,
    periodoAuto: periodoAuto,
    mesDe: mesDe <= mesAte ? mesDe : mesAte,
    mesAte: mesDe <= mesAte ? mesAte : mesDe,
    urlFerias: props.getProperty(CHAVES_CONFIG.URL_FERIAS) || URL_PLANILHA_FERIAS,
    urlAgenda: props.getProperty(CHAVES_CONFIG.URL_AGENDA) || URL_PLANILHA_AGENDA_AUSENCIAS
  };
}

/********************************************************************************
 * MAPA DE-PARA DE USUÁRIOS
 ********************************************************************************/
function criarAbaDeParaManualmente() {
  carregarMapaUsuarios_(obterPlanilha_());
  SpreadsheetApp.getUi().alert('Aba "DePara_Usuarios" criada/verificada com sucesso!');
}

const CABECALHO_DEPARA = ['ID Jira (Username / AccountId)', 'Nome Real do Colaborador', 'E-mail Corporativo', 'Apelidos (agenda)', 'Cobrar (SIM/NÃO)'];

function carregarMapaUsuarios_(ss) {
  const mapa = {};
  let aba = ss.getSheetByName('DePara_Usuarios');

  if (!aba) {
    aba = ss.insertSheet('DePara_Usuarios');
    aba.getRange(1, 1, 1, CABECALHO_DEPARA.length).setValues([CABECALHO_DEPARA]);
  }
  // Garante as colunas novas (Apelidos / Cobrar) sem mexer nos dados
  try {
    const cab = aba.getRange(1, 1, 1, CABECALHO_DEPARA.length).getValues()[0];
    CABECALHO_DEPARA.forEach((t, i) => { if (!String(cab[i] || '').trim()) aba.getRange(1, i + 1).setValue(t); });
    aba.getRange(1, 1, 1, CABECALHO_DEPARA.length).setBackground('#16213e').setFontColor('#ffffff').setFontWeight('bold');
  } catch (e) { /* sem permissão de edição */ }

  if (aba.getLastRow() <= 1) return mapa;

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, CABECALHO_DEPARA.length).getValues();
  dados.forEach(linha => {
    const idJira = String(linha[0] || '').trim().toLowerCase();
    const nomeReal = String(linha[1] || '').trim();
    const emailReal = String(linha[2] || '').trim().toLowerCase();
    const cobrar = String(linha[4] || '').trim().toUpperCase().replace('Ã', 'A') !== 'NAO';

    const infoObj = { nome: nomeReal, email: emailReal, cobrar: cobrar, idJira: idJira };

    if (idJira) mapa[idJira] = infoObj;
    if (emailReal) {
      mapa[emailReal] = infoObj;
      mapa[emailReal.split('@')[0]] = infoObj;
    }
  });

  return mapa;
}

/* Ausências, férias, folgas e feriados: ver Ausencias.js | Cálculo do período: ver Apuracao.js */

/********************************************************************************
 * EXTRAÇÃO E SALVAMENTO COM TRATAMENTO DE AUSÊNCIAS (FÉRIAS + AUDITORIA DE AGENDA)
 ********************************************************************************/
function extrairEGravarNaPlanilha(opcoes) {
  const apur = apurarPeriodo_(opcoes || {});
  const ss = obterPlanilha_();
  const horaMinimaNum = apur.horaMinima;

  const rotuloStatus = { OK: 'OK', PEND: 'PENDENTE', AUS: 'FÉRIAS / AUSÊNCIA', FER: 'FERIADO', HOJE: 'EM ANDAMENTO' };
  const registrosProcessados = [];
  const resumoConsolidadoDevs = [];
  const pendenciasDetalhadas = [];
  const matrizExtracao = [['Data', 'Dev / Analista', 'E-mail', 'Total Horas no Dia', `Status (Mín. ${horaMinimaNum}h)`, 'Meta do Dia', 'Ausência Considerada']];
  const matrizAuditoria = [['Data', 'Dev / Analista', 'E-mail', 'Tipo Ausência', 'Título / Motivo Evento', 'Horário Registrado', 'Abatimento / Horas', 'Origem']];

  apur.devs.forEach(dev => {
    dev.dias.forEach((x, i) => {
      const dia = apur.dias[i];
      const status = rotuloStatus[x.s] || x.s;
      matrizExtracao.push([dia.iso, dev.nome, dev.email, x.h, status, x.m, x.a || '']);
      registrosProcessados.push({ dataIso: dia.iso, dataBr: dia.br, nome: dev.nome, email: dev.email, horas: x.h, status: status });
      if (x.s === 'PEND') pendenciasDetalhadas.push([dia.iso, dev.nome, dev.email, x.h, 'PENDENTE', x.m, x.a || '']);
    });
    const t = dev.totais;
    resumoConsolidadoDevs.push({
      nome: dev.nome, email: dev.email, diasUteisMeta: t.diasUteis, diasFerias: t.diasAusencia,
      horasAusencia: t.horasAusencia, horasEsperadas: t.esperadas, horasApontadas: t.apontadas,
      saldoHoras: t.saldo, statusGeral: t.saldo >= 0 && t.diasPendentes === 0 ? 'OK' : 'PENDENTE'
    });
  });

  apur.ausencias.registros.forEach(r => {
    matrizAuditoria.push([r.inicio === r.fim ? r.inicio : `${r.inicio} a ${r.fim}`, r.nome, r.email, r.tipo, r.titulo, r.horario,
      r.horas == null ? 'Dia Inteiro' : r.horas, r.origem]);
  });

  // ABA 1: Resumo_Mensal_Devs
  const matrizResumo = [['Dev / Analista', 'E-mail', 'Dias Úteis (Até Ontem)', 'Dias Férias/Ausência', 'Horas Ausência (Abatidas)', 'Horas Esperadas', 'Horas Apontadas', 'Saldo (Horas)', 'Status Geral']];
  resumoConsolidadoDevs.forEach(d => matrizResumo.push([d.nome, d.email, d.diasUteisMeta, d.diasFerias, d.horasAusencia, d.horasEsperadas, d.horasApontadas, d.saldoHoras, d.statusGeral]));
  const abaResumo = prepararAba_(ss, 'Resumo_Mensal_Devs');
  abaResumo.getRange(1, 1, matrizResumo.length, matrizResumo[0].length).setValues(matrizResumo);
  estilizarCabecalho_(abaResumo, 1, matrizResumo[0].length, '#16213e');
  if (matrizResumo.length > 1) abaResumo.getRange(2, 5, matrizResumo.length - 1, 4).setNumberFormat('0.00"h"');
  aplicarCoresStatusGeral_(abaResumo, matrizResumo.length);
  abaResumo.autoResizeColumns(1, matrizResumo[0].length);

  // ABA 2: Horas_Pendentes
  const matrizPendentes = [['Data', 'Dev / Analista', 'E-mail', 'Total Horas no Dia', 'Status', 'Meta do Dia', 'Ausência Considerada']].concat(pendenciasDetalhadas);
  const abaPendentes = prepararAba_(ss, 'Horas_Pendentes');
  abaPendentes.getRange(1, 1, matrizPendentes.length, matrizPendentes[0].length).setValues(matrizPendentes);
  estilizarCabecalho_(abaPendentes, 1, matrizPendentes[0].length, '#c5221f');
  if (matrizPendentes.length > 1) {
    abaPendentes.getRange(2, 4, matrizPendentes.length - 1, 1).setNumberFormat('0.00"h"');
    abaPendentes.getRange(2, 5, matrizPendentes.length - 1, 1).setBackground('#fce8e6').setFontColor('#c5221f').setFontWeight('bold');
  }
  abaPendentes.autoResizeColumns(1, matrizPendentes[0].length);

  // ABA 3: Auditoria_Ausencias
  const abaAuditoria = prepararAba_(ss, 'Auditoria_Ausencias');
  abaAuditoria.getRange(1, 1, matrizAuditoria.length, matrizAuditoria[0].length).setValues(matrizAuditoria);
  estilizarCabecalho_(abaAuditoria, 1, matrizAuditoria[0].length, '#e65100');
  abaAuditoria.autoResizeColumns(1, matrizAuditoria[0].length);

  // ABA 4: Extracao_Tempo
  const abaExtracao = prepararAba_(ss, 'Extracao_Tempo');
  abaExtracao.getRange(1, 1, matrizExtracao.length, matrizExtracao[0].length).setValues(matrizExtracao);
  estilizarCabecalho_(abaExtracao, 1, matrizExtracao[0].length, '#1a73e8');
  if (matrizExtracao.length > 1) abaExtracao.getRange(2, 4, matrizExtracao.length - 1, 1).setNumberFormat('0.00"h"');
  aplicarCoresStatusDiario_(abaExtracao, matrizExtracao.length);
  abaExtracao.autoResizeColumns(1, matrizExtracao[0].length);

  const aus = apur.ausencias;
  return {
    mensagem: `Sucesso! Abas 'Resumo_Mensal_Devs', 'Horas_Pendentes', 'Auditoria_Ausencias' e 'Extracao_Tempo' geradas com sucesso. ` +
      `${aus.registros.length} ausência(s)/férias consideradas.` +
      (aus.naoVinculados.length ? ` Sem colaborador identificado (cadastre o apelido no painel > Cadastros > Colaboradores): ${aus.naoVinculados.join(', ')}.` : '') +
      (aus.avisos.length ? ` Avisos: ${aus.avisos.join(' | ')}` : ''),
    registros: registrosProcessados,
    resumoDevs: resumoConsolidadoDevs,
    periodoIso: `${apur.periodo.de} a ${apur.periodo.ate}`
  };
}

function prepararAba_(ss, nome) {
  const aba = ss.getSheetByName(nome) || ss.insertSheet(nome);
  aba.clear();
  return aba;
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
  const apur = apurarPeriodo_({});
  const pendenciasPorDev = pendenciasDaApuracao_(apur).filter(d => d.cobrar);

  if (pendenciasPorDev.length === 0) return 'Excelente! Nenhuma pendência de horas encontrada para o time.';

  pendenciasPorDev.forEach(dev => enviarCobrancaDev_(dev, cfg, apur.periodo, 'Semanal', somenteLider));

  MailApp.sendEmail({
    to: cfg.emailLider,
    subject: `${somenteLider ? '[TESTE] ' : ''}Consolidado Semanal de Pendências (TimeSheet)`,
    htmlBody: montarEmailConsolidadoLider_(pendenciasPorDev, apur.periodo.de, apur.periodo.ate, cfg.horaMinima)
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

/**
 * E-mail de cobrança individual — layout em tabelas com estilos inline
 * (compatível com Gmail, Outlook e celular).
 */
function montarEmailIndividualDev_(item, horaMinima, periodo) {
  const esc = t => String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const DIAS = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb'];
  const br = iso => iso.split('-').reverse().join('/');
  const primeiroNome = String(item.nome || '').trim().split(/\s+/)[0] || '';
  const nomeFmt = primeiroNome.charAt(0).toUpperCase() + primeiroNome.slice(1).toLowerCase();
  const pend = item.pendencias || [];
  const faltaTotal = pend.reduce((s, p) => s + Math.max(0, (p.meta == null ? parseFloat(horaMinima) || 0 : p.meta) - p.horas), 0);
  const zerados = pend.filter(p => !p.horas).length;

  let urlTempo = '';
  try { urlTempo = montarUrlTimesheetTempo_(obterConfiguracoesAtuais(), periodo); } catch (e) { urlTempo = ''; }

  const linhas = pend.map((p, i) => {
    const iso = p.iso || p.data.split('/').reverse().join('-');
    const d = new Date(+iso.substr(0, 4), +iso.substr(5, 2) - 1, +iso.substr(8, 2));
    const meta = p.meta == null ? parseFloat(horaMinima) || 0 : p.meta;
    const falta = Math.max(0, meta - p.horas);
    const zero = !p.horas;
    const fundo = i % 2 ? '#f8fafc' : '#ffffff';
    const badge = zero
      ? '<span style="display:inline-block;padding:3px 10px;border-radius:999px;background:#fee2e2;color:#b91c1c;font-size:11px;font-weight:700;white-space:nowrap;">Sem registro</span>'
      : '<span style="display:inline-block;padding:3px 10px;border-radius:999px;background:#fef3c7;color:#92400e;font-size:11px;font-weight:700;">Parcial</span>';
    const td = 'padding:10px 8px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#1e293b;';
    return `<tr style="background:${fundo};">
      <td style="${td}white-space:nowrap;"><b>${esc((p.data || br(iso)).substring(0, 5))}</b><div style="font-size:11px;color:#64748b;text-transform:capitalize;">${DIAS[d.getDay()]}</div></td>
      <td style="${td}text-align:center;">${formatarHorasLegivel_(p.horas)}</td>
      <td style="${td}text-align:center;color:#64748b;">${formatarHorasLegivel_(meta)}</td>
      <td style="${td}text-align:center;font-weight:700;color:#b91c1c;">${formatarHorasLegivel_(falta)}</td>
      <td style="${td}text-align:center;">${badge}${p.ausencia ? `<div style="margin-top:4px;font-size:11px;color:#6d28d9;">${esc(p.ausencia)} já descontada</div>` : ''}</td>
    </tr>`;
  }).join('');

  const th = 'padding:10px 8px;background:#f1f5f9;border-bottom:1px solid #e2e8f0;font-size:10.5px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.3px;';
  const card = (valor, rotulo, cor) => `<td width="33%" style="padding:0 4px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e2e8f0;border-radius:10px;border-top:3px solid ${cor};">
        <tr><td style="padding:12px 6px;text-align:center;">
          <div style="font-size:24px;font-weight:800;color:#0f172a;line-height:1.1;">${valor}</div>
          <div style="font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.5px;margin-top:4px;">${rotulo}</div>
        </td></tr></table></td>`;
  const periodoTxt = periodo ? `${br(periodo.de)} a ${br(periodo.ate)}` : '';

  return `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f1f5f9;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f1f5f9;">
<tr><td align="center" style="padding:20px 6px;">
  <table role="presentation" width="640" cellpadding="0" cellspacing="0" style="width:100%;max-width:640px;background:#ffffff;border-radius:14px;overflow:hidden;font-family:'Segoe UI',Arial,Helvetica,sans-serif;box-shadow:0 2px 10px rgba(15,23,42,.06);">

    <tr><td bgcolor="#0b1f3a" style="background:#0b1f3a;background-image:linear-gradient(120deg,#0b1f3a 0%,#0f5b78 100%);padding:26px 30px;">
      <div style="font-size:11px;letter-spacing:2px;text-transform:uppercase;color:#7dd3fc;font-weight:700;">Timesheet · Apontamento de horas</div>
      <div style="font-size:22px;font-weight:800;color:#ffffff;margin-top:6px;">Pendências de apontamento</div>
      ${periodoTxt ? `<div style="font-size:13px;color:#cbd5e1;margin-top:4px;">Período de ${periodoTxt}</div>` : ''}
    </td></tr>

    <tr><td style="padding:28px 30px 8px;">
      <p style="margin:0 0 12px;font-size:15px;color:#0f172a;">Olá, <b>${esc(nomeFmt)}</b>!</p>
      <p style="margin:0;font-size:14px;line-height:1.6;color:#334155;">Identificamos dias úteis já encerrados em que as horas lançadas no Jira/Tempo ficaram abaixo da meta diária. Pedimos, por gentileza, que regularize os apontamentos abaixo.</p>
    </td></tr>

    <tr><td style="padding:18px 16px 6px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
        ${card(pend.length, pend.length === 1 ? 'dia pendente' : 'dias pend.', '#ef4444')}
        ${card(formatarHorasLegivel_(faltaTotal), 'a apontar', '#f97316')}
        ${card(zerados, 'sem registro', '#64748b')}
      </tr></table>
    </td></tr>

    <tr><td style="padding:18px 12px 6px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e2e8f0;border-radius:10px;border-collapse:separate;overflow:hidden;">
        <tr>
          <th align="left" style="${th}">Data</th>
          <th style="${th}text-align:center;">Apontado</th>
          <th style="${th}text-align:center;">Meta</th>
          <th style="${th}text-align:center;">Falta</th>
          <th style="${th}text-align:center;">Situação</th>
        </tr>
        ${linhas}
      </table>
    </td></tr>

    ${urlTempo ? `<tr><td align="center" style="padding:22px 30px 6px;">
      <table role="presentation" cellpadding="0" cellspacing="0"><tr><td bgcolor="#0f5b78" style="border-radius:8px;background:#0f5b78;">
        <a href="${esc(urlTempo)}" target="_blank" style="display:inline-block;padding:12px 26px;font-size:14px;font-weight:700;color:#ffffff;text-decoration:none;border-radius:8px;">Abrir o Timesheet no Tempo →</a>
      </td></tr></table>
    </td></tr>` : ''}

    <tr><td style="padding:20px 30px 26px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f0f9ff;border:1px solid #bae6fd;border-radius:10px;"><tr><td style="padding:12px 16px;font-size:12.5px;line-height:1.55;color:#0c4a6e;">
        <b>Já descontamos da meta:</b> férias, folgas, feriados e ausências registradas na agenda do time. Se você esteve ausente em algum desses dias e o registro não está na agenda, responda este e-mail ou avise seu líder para ajustarmos.
      </td></tr></table>
    </td></tr>

    <tr><td style="padding:16px 30px;background:#f8fafc;border-top:1px solid #e2e8f0;font-size:11.5px;color:#94a3b8;line-height:1.5;">
      Mensagem automática do controle de Timesheet · meta diária de ${formatarHorasLegivel_(parseFloat(horaMinima) || 0)} por dia útil.
    </td></tr>

  </table>
</td></tr></table></body></html>`;
}

/** Link da visão de timesheet do time no Tempo, já no período cobrado. */
function montarUrlTimesheetTempo_(cfg, periodo) {
  const base = String(cfg.jiraBaseUrl || '').replace(/\/$/, '');
  const team = encodeURIComponent(String(cfg.teamId || '').trim());
  if (!base || !team) return '';
  const hoje = formatarDataISO_(new Date());
  const de = periodo && periodo.de ? periodo.de : hoje.substring(0, 8) + '01';
  const ate = periodo && periodo.ate ? periodo.ate : hoje;
  const params = [
    'columns=WORKED_COLUMN', 'columns=PORTION_OF_REQUIRED_COLUMN', 'dateDisplayType=days',
    'from=' + de, 'groupBy=worker', 'order=ASCENDING', 'periodType=FIXED', 'sortBy=TITLE_COLUMN',
    'subPeriodType=MONTH', 'teamId=' + team, 'to=' + ate, 'viewType=TIMESHEET'
  ];
  return `${base}/secure/Tempo.jspa#/teams/team/${team}/timesheet?${params.join('&')}`;
}

/** 4 -> "4h", 0.33 -> "20min", 4.5 -> "4h30" */
function formatarHorasLegivel_(h) {
  const min = Math.round((Number(h) || 0) * 60);
  if (min <= 0) return '0h';
  const hh = Math.floor(min / 60), mm = min % 60;
  if (!hh) return mm + 'min';
  return hh + 'h' + (mm ? String(mm).padStart(2, '0') : '');
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
  const ss = obterPlanilha_();
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
    gravarCacheGrande_('RELATORIO_MATUTINO', resultado, 21600);
    
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

  let resultado = lerCacheGrande_('RELATORIO_MATUTINO');
  if (!resultado) {
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