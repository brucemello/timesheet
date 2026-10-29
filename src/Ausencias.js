/********************************************************************************
 * AUSÊNCIAS, FÉRIAS, FOLGAS E FERIADOS — base para cobrar o apontamento certo
 * (lógica trazida do projeto Capacidade Não Produtiva)
 *
 * Fontes:
 *   1. Planilha de Férias (URL_PLANILHA_FERIAS)
 *      - abas "Controle Consolidado*": férias Marcadas / Em gozo / Finalizadas / Agendadas
 *      - abas "Folgas e Emendas*": GIFTCARD, RESETDAY, MVT... (conta dias úteis)
 *      - aba "Contatos" (Nome, Email): ajuda a identificar o colaborador
 *   2. Planilha "Agendamento Ausências" (URL_PLANILHA_AGENDA_AUSENCIAS), espelho da agenda
 *      - "Dia Inteiro" -> dia sem cobrança (Data Término é exclusiva, como na agenda)
 *      - com horário   -> abate as horas do evento (sobreposições contam uma vez só)
 *   3. Aba local "Ausencias" (lançamento manual): Identificador | Início | Fim | Horas (opcional)
 *   4. Feriados: nacionais + SP calculados automaticamente + aba local "Feriados"
 *
 * Identificação do colaborador (pelo título primeiro, depois pelo criador):
 *   "Ausência: JOSE RIQUELMO GOMES DA SILVA [Marcadas]", "Keilla - Consulta",
 *   "Consulta - Pâmela", "Edu - Exame", "Fisioterapia Ombro | Fabio Mendonça"...
 *   Apelidos que não lembram o nome (ex.: "Carlão") vão na coluna D
 *   "Apelidos (agenda)" da aba DePara_Usuarios, separados por vírgula.
 ********************************************************************************/

const AUS_HORAS_DIA = 8;
const AUS_TIPO_INFORMATIVO = 'Informativo';
const AUS_ORIGEM_FERIAS = 'Planilha de Férias';
const AUS_ORIGEM_AGENDA = 'Agenda de Ausências';
const AUS_ORIGEM_MANUAL = 'Aba Ausencias';

/**
 * Lê todas as fontes e devolve:
 *   { registros: [...], mapa: { 'email|yyyy-MM-dd': {...} }, naoVinculados: [...], avisos: [...] }
 * @param {Spreadsheet} ss planilha do timesheet
 * @param {Array<{email:string,nome:string}>} devs colaboradores considerados no cálculo
 */
function carregarAusencias_(ss, devs) {
  const avisos = [];
  const lista = (devs || []).filter(d => d && d.email);
  const apelidos = lerApelidosDePara_(ss);

  let planilhaFerias = null;
  try { planilhaFerias = SpreadsheetApp.openByUrl(URL_PLANILHA_FERIAS); }
  catch (e) { avisos.push('Planilha de férias inacessível: ' + e.message); }

  let contatos = {};
  if (planilhaFerias) {
    const abaContatos = planilhaFerias.getSheetByName('Contatos');
    if (abaContatos && abaContatos.getLastRow() > 1) {
      contatos = interpretarContatos_(abaContatos.getRange(1, 1, abaContatos.getLastRow(), 2).getDisplayValues());
    }
  }

  const matcher = criarMatcherColaborador_(lista, contatos, apelidos);
  const registros = [];
  const naoVinculados = {};
  const coletar = r => {
    r.registros.forEach(x => registros.push(x));
    r.naoVinculados.forEach(n => { if (n) naoVinculados[n] = true; });
  };

  // 1. Planilha de Férias
  if (planilhaFerias) {
    planilhaFerias.getSheets().forEach(sh => {
      const nome = sh.getName();
      const ehConsolidado = /^controle consolidado/i.test(nome);
      const ehFolgas = /^folgas/i.test(nome);
      if ((!ehConsolidado && !ehFolgas) || sh.getLastRow() < 2) return;
      const dados = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
      coletar(ehConsolidado ? interpretarControleConsolidado_(dados, matcher) : interpretarFolgas_(dados, matcher));
    });
  }

  // 2. Agendamento de Ausências (espelho da agenda do time)
  try {
    const planilhaAgenda = SpreadsheetApp.openByUrl(URL_PLANILHA_AGENDA_AUSENCIAS);
    const aba = planilhaAgenda.getSheetByName('Página1') || planilhaAgenda.getSheets()[0];
    if (aba && aba.getLastRow() > 1) {
      coletar(interpretarAgendamento_(aba.getRange(2, 1, aba.getLastRow() - 1, 7).getDisplayValues(), matcher));
    }
  } catch (e) {
    avisos.push('Planilha de agendamento de ausências inacessível: ' + e.message);
  }

  // 3. Aba local "Ausencias"
  const abaLocal = ss.getSheetByName('Ausencias');
  if (abaLocal && abaLocal.getLastRow() > 1) {
    coletar(interpretarAusenciasLocais_(abaLocal.getRange(2, 1, abaLocal.getLastRow() - 1, 4).getDisplayValues(), matcher));
  }

  const nv = Object.keys(naoVinculados).sort();
  avisos.forEach(a => Logger.log('⚠️ ' + a));
  Logger.log(`✅ ${registros.length} ausência(s) vinculada(s). Não vinculadas: ${nv.length ? nv.join(', ') : 'nenhuma'}`);

  return { registros: registros, mapa: montarMapaAusencias_(registros), naoVinculados: nv, avisos: avisos };
}

/* =============================================================================
 * INTERPRETAÇÃO DAS FONTES (funções puras: recebem matrizes de texto)
 * ===========================================================================*/

/** "Controle Consolidado": cabeçalho na linha 1, colunas localizadas pelo nome. */
function interpretarControleConsolidado_(dados, matcher) {
  const out = { registros: [], naoVinculados: [] };
  const h = dados[0].map(x => normalizarTextoAus_(x));
  const col = nome => h.findIndex(x => x.indexOf(nome) === 0);
  const iNome = col('nome'), iSit = col('situacao'), iAg = col('agendamento');
  const iIni = col('data inicio programacao'), iFim = col('data fim programacao');
  if ([iNome, iIni, iFim].some(i => i < 0)) return out;

  for (let r = 1; r < dados.length; r++) {
    const nome = String(dados[r][iNome] || '').trim();
    const sitTxt = iSit >= 0 ? String(dados[r][iSit] || '') : '';
    const agTxt = iAg >= 0 ? String(dados[r][iAg] || '') : '';
    const sit = normalizarTextoAus_(sitTxt), ag = normalizarTextoAus_(agTxt);
    const ini = lerDataAus_(dados[r][iIni]), fim = lerDataAus_(dados[r][iFim]);
    if (!nome || !ini || !fim || fim < ini) continue;
    const confirmada = /marcad|em gozo|finalizad/.test(sit) || ag === 'agendados' || ag === 'agendado';
    if (!confirmada) continue;
    const email = matcher(nome, '', isoAus_(ini));
    if (!email) { out.naoVinculados.push(nome); continue; }
    out.registros.push({
      email: email, inicio: isoAus_(ini), fim: isoAus_(fim), tipo: 'Férias', horas: null,
      origem: AUS_ORIGEM_FERIAS, titulo: ('Férias · ' + [sitTxt, agTxt].filter(x => x && x !== '-').join(' · ')).trim(),
      horarioTexto: 'Dia Inteiro'
    });
  }
  return out;
}

/** "Folgas e Emendas": Colaborador | Benefício | Data da Folga | Quantidade de Dias | ... */
function interpretarFolgas_(dados, matcher) {
  const out = { registros: [], naoVinculados: [] };
  const h = dados[0].map(x => normalizarTextoAus_(x));
  const col = nome => h.findIndex(x => x.indexOf(nome) === 0);
  const iNome = col('colaborador'), iBen = col('beneficio'), iIni = col('data da folga'), iQtd = col('quantidade');
  if (iNome < 0 || iIni < 0) return out;

  for (let r = 1; r < dados.length; r++) {
    const nome = String(dados[r][iNome] || '').trim();
    const ini = lerDataAus_(dados[r][iIni]);
    if (!nome || !ini) continue;
    const qtd = Math.max(1, parseInt(iQtd >= 0 ? dados[r][iQtd] : 1, 10) || 1);
    const fim = ultimoDiaUtilApos_(ini, qtd);
    const email = matcher(nome, '', isoAus_(ini));
    if (!email) { out.naoVinculados.push(nome); continue; }
    out.registros.push({
      email: email, inicio: isoAus_(ini), fim: isoAus_(fim), tipo: 'Folga', horas: null,
      origem: AUS_ORIGEM_FERIAS, titulo: 'Folga · ' + (String(dados[r][iBen] || '').trim() || 'Folga'),
      horarioTexto: 'Dia Inteiro'
    });
  }
  return out;
}

/**
 * "Agendamento Ausências" (sem cabeçalho):
 * Título | Criador/Autor | Data Início | Hora Início | Data Término | Hora Término | ID do Evento
 */
function interpretarAgendamento_(linhas, matcher) {
  const out = { registros: [], naoVinculados: [] };
  const vistos = {};
  linhas.forEach(r => {
    const titulo = String(r[0] || '').trim();
    const autor = String(r[1] || '').trim();
    const dIni = lerDataAus_(r[2]);
    let dFim = lerDataAus_(r[4]) || dIni;
    const hIni = String(r[3] || '').trim(), hFim = String(r[5] || '').trim();
    if (!dIni || (!titulo && !autor)) return;

    const diaInteiro = /dia inteiro/i.test(hIni) || /dia inteiro/i.test(hFim);
    const inicioIso = isoAus_(dIni);
    let intervalo = null, horas = null;

    if (diaInteiro) {
      if (dFim > dIni) dFim = new Date(dFim.getFullYear(), dFim.getMonth(), dFim.getDate() - 1); // término exclusivo
    } else if (isoAus_(dFim) === inicioIso) {
      const a = minutosAus_(hIni), b = minutosAus_(hFim);
      if (a < 0 || b <= a) return;
      intervalo = [a, b];
      horas = Math.round((b - a) / 60 * 100) / 100;
    } // com horário em vários dias -> trata como dias inteiros

    const email = matcher(titulo, autor, inicioIso);
    if (!email) { out.naoVinculados.push(titulo || autor); return; }

    // mesma ausência repetida (eventos duplicados na agenda / linhas repetidas)
    const chave = [email, inicioIso, isoAus_(dFim), intervalo ? intervalo.join('-') : 'integral'].join('|');
    if (vistos[chave]) return;
    vistos[chave] = true;

    out.registros.push({
      email: email, inicio: inicioIso, fim: isoAus_(dFim),
      tipo: inferirTipoAusencia_(titulo), horas: horas, intervalo: intervalo,
      origem: AUS_ORIGEM_AGENDA, titulo: titulo || autor,
      horarioTexto: intervalo ? `${hIni.substring(0, 5)} às ${hFim.substring(0, 5)}` : 'Dia Inteiro'
    });
  });
  return out;
}

/** Aba local "Ausencias": Identificador (nome/e-mail) | Início | Fim | Horas (opcional: parcial por dia) */
function interpretarAusenciasLocais_(linhas, matcher) {
  const out = { registros: [], naoVinculados: [] };
  linhas.forEach(r => {
    const ident = String(r[0] || '').trim();
    const ini = lerDataAus_(r[1]);
    const fim = lerDataAus_(r[2]) || ini;
    if (!ident || !ini || fim < ini) return;
    const horasNum = parseFloat(String(r[3] || '').replace(',', '.'));
    const parcial = horasNum > 0 && horasNum < AUS_HORAS_DIA;
    const email = matcher(ident, ident, isoAus_(ini));
    if (!email) { out.naoVinculados.push(ident); return; }
    out.registros.push({
      email: email, inicio: isoAus_(ini), fim: isoAus_(fim), tipo: 'Ausência', horas: parcial ? horasNum : null,
      origem: AUS_ORIGEM_MANUAL, titulo: 'Ausência registrada localmente',
      horarioTexto: parcial ? horasNum + 'h' : 'Dia Inteiro'
    });
  });
  return out;
}

function interpretarContatos_(dados) {
  const mapa = {};
  dados.slice(1).forEach(r => { if (r[0] && r[1]) mapa[String(r[1]).trim().toLowerCase()] = String(r[0]).trim(); });
  return mapa;
}

/** Coluna D "Apelidos (agenda)" da DePara_Usuarios: email -> ['carlao', ...]. Cria o cabeçalho se faltar. */
function lerApelidosDePara_(ss) {
  const mapa = {};
  const aba = ss.getSheetByName('DePara_Usuarios');
  if (!aba) return mapa;
  try {
    if (!String(aba.getRange(1, 4).getValue() || '').trim()) {
      aba.getRange(1, 4).setValue('Apelidos (agenda)').setBackground('#16213e').setFontColor('#ffffff').setFontWeight('bold');
    }
  } catch (e) { /* sem permissão de edição: segue sem o cabeçalho */ }
  if (aba.getLastRow() < 2) return mapa;
  aba.getRange(2, 1, aba.getLastRow() - 1, 4).getDisplayValues().forEach(r => {
    const email = String(r[2] || '').trim().toLowerCase();
    const lista = String(r[3] || '').split(/[,;]/).map(x => normalizarTextoAus_(x)).filter(Boolean);
    if (email && lista.length) mapa[email] = (mapa[email] || []).concat(lista);
  });
  return mapa;
}

/* =============================================================================
 * IDENTIFICAÇÃO DO COLABORADOR -> e-mail do dev
 * ===========================================================================*/
function criarMatcherColaborador_(devs, contatos, apelidos) {
  const ignorar = ['dos', 'das', 'de', 'da', 'do', 'e'];
  const lista = devs.map(d => ({
    email: d.email.toLowerCase(),
    prefixo: d.email.toLowerCase().split('@')[0],
    nomeNorm: normalizarTextoAus_(d.nome),
    tokens: normalizarTextoAus_(d.nome).split(/[^a-z]+/).filter(t => t.length > 2 && ignorar.indexOf(t) < 0),
    apelidos: (apelidos && apelidos[d.email.toLowerCase()]) || []
  }));

  function porTexto(texto) {
    const limpo = normalizarTextoAus_(String(texto || '').replace(/\[.*?\]/g, ' '));
    // apelido cadastrado (pode ter espaço, ex.: "ze riquelmo")
    const porApelido = lista.filter(c => c.apelidos.some(a => new RegExp('(^|[^a-z])' + a + '([^a-z]|$)').test(limpo)));
    if (porApelido.length === 1) return porApelido[0];

    const tk = limpo.split(/[^a-z]+/).filter(t => t.length >= 3);
    if (!tk.length) return null;
    let melhor = [], melhorScore = 0;
    lista.forEach(c => {
      let score = 0;
      tk.forEach(t => {
        if (c.tokens[0] === t) score += 3;                         // primeiro nome pesa mais
        else if (c.tokens.indexOf(t) >= 0) score += 2;
        else if (c.tokens[0] && c.tokens[0].indexOf(t) === 0) score += 1; // "edu" -> eduardo
      });
      if (score > melhorScore) { melhorScore = score; melhor = [c]; }
      else if (score === melhorScore && score > 0) melhor.push(c);
    });
    return melhorScore && melhor.length === 1 ? melhor[0] : null; // empate = ambíguo
  }

  function porEmail(texto) {
    const email = String(texto || '').toLowerCase().trim();
    if (email.indexOf('@') < 0) return null;
    const prefixo = email.split('@')[0];
    const direto = lista.filter(c => c.email === email || c.prefixo === prefixo);
    if (direto.length === 1) return direto[0];
    if (contatos && contatos[email]) return porTexto(contatos[email]);
    return null;
  }

  return function (titulo, criador) {
    const c = porEmail(titulo) || porTexto(titulo) || porEmail(criador) || porTexto(criador);
    return c ? c.email : null;
  };
}

function inferirTipoAusencia_(titulo) {
  const t = normalizarTextoAus_(titulo);
  const status = (t.match(/\[(.*)\]/) || [])[1] || '';
  if (/marcad|ferias|em gozo|agendad|sugest/.test(status)) return 'Férias';
  if (/resetday|giftcard|folga|mvt/.test(status)) return 'Folga';
  if (/ferias/.test(t)) return 'Férias';
  if (/resetday|giftcard|folga|emenda|mvt|day ?off/.test(t)) return 'Folga';
  if (/banco de horas/.test(t)) return 'Banco de horas';
  if (/atestado/.test(t)) return 'Atestado';
  if (/licenca/.test(t)) return 'Licença';
  if (/curso|treinamento|letramento|workshop|palestra/.test(t)) return 'Treinamento';
  if (/consulta|medic|exame|dentista|fisio|nutri|terapia|cardio|hospital|cirurgia|vacina|periodico|ocupacional/.test(t)) return 'Consulta/Exame';
  return 'Ausência';
}

/* =============================================================================
 * MAPA DIÁRIO E AVALIAÇÃO DO DIA
 * ===========================================================================*/
function montarMapaAusencias_(registros) {
  const mapa = {};
  registros.forEach(a => {
    if (!a.email || !a.inicio || a.tipo === AUS_TIPO_INFORMATIVO) return;
    const ini = isoParaDataAus_(a.inicio), fim = isoParaDataAus_(a.fim || a.inicio);
    for (let d = new Date(ini); d <= fim; d.setDate(d.getDate() + 1)) {
      const k = a.email + '|' + isoAus_(d);
      const m = mapa[k] || (mapa[k] = { integral: false, intervalos: [], horasAvulsas: 0, eventos: [] });
      if (a.horas == null) m.integral = true;
      else if (a.intervalo) m.intervalos.push(a.intervalo);
      else m.horasAvulsas += a.horas;
      const rotulo = a.tipo + '|' + a.titulo + '|' + a.horarioTexto;
      if (!m.eventos.some(e => e.rotulo === rotulo)) {
        m.eventos.push({
          rotulo: rotulo, tipo: a.tipo, titulo: a.titulo, origem: a.origem, horarioTexto: a.horarioTexto,
          eDiaInteiro: a.horas == null, horasAbatimento: a.horas == null ? 0 : a.horas
        });
      }
    }
  });
  return mapa;
}

/** Soma dos intervalos [ini,fim] em minutos, contando sobreposições uma vez só. */
function horasDosIntervalos_(intervalos) {
  const ord = intervalos.slice().sort((a, b) => a[0] - b[0]);
  let total = 0, ini = -1, fim = -1;
  ord.forEach(([a, b]) => {
    if (a > fim) { if (fim > ini) total += fim - ini; ini = a; fim = b; }
    else fim = Math.max(fim, b);
  });
  if (fim > ini) total += fim - ini;
  return total / 60;
}

/**
 * Meta do dia do dev considerando ausências.
 * Dia inteiro (férias, folga, ausência integral) -> meta 0.
 * Ausência parcial -> meta = hora mínima - horas ausentes (nunca negativa).
 */
function avaliarStatusEAusenciaDev_(emailDev, nomeDev, dataObj, ausencias, horaMinimaPadrao) {
  const mapa = (ausencias && ausencias.mapa) || {};
  const m = mapa[String(emailDev).toLowerCase() + '|' + isoAus_(dataObj)];
  if (!m) return { eDiaInteiro: false, horasAbatidas: 0, cargaEsperadaNoDia: horaMinimaPadrao, eventosAuditoria: [] };

  let horas = m.integral ? 0 : Math.min(horasDosIntervalos_(m.intervalos) + m.horasAvulsas, AUS_HORAS_DIA);
  const eDiaInteiro = m.integral || horas >= AUS_HORAS_DIA;
  horas = Math.round(horas * 100) / 100;
  return {
    eDiaInteiro: eDiaInteiro,
    horasAbatidas: eDiaInteiro ? 0 : horas,
    cargaEsperadaNoDia: eDiaInteiro ? 0 : Math.round(Math.max(0, horaMinimaPadrao - horas) * 100) / 100,
    eventosAuditoria: m.eventos
  };
}

/** Texto curto das ausências do dia (para e-mails). */
function descreverAusenciaDia_(avaliacao) {
  return (avaliacao.eventosAuditoria || []).map(e => e.tipo + (e.eDiaInteiro ? '' : ' ' + e.horarioTexto)).join(' + ');
}

/* =============================================================================
 * FERIADOS (nacionais + SP automáticos + aba "Feriados")
 * ===========================================================================*/
const FERIADOS_FIXOS_AUS = [
  [1, 1, 'Confraternização Universal'], [4, 21, 'Tiradentes'], [5, 1, 'Dia do Trabalho'],
  [7, 9, 'Revolução Constitucionalista (SP)'], [9, 7, 'Independência do Brasil'],
  [10, 12, 'Nossa Senhora Aparecida'], [11, 2, 'Finados'], [11, 15, 'Proclamação da República'],
  [11, 20, 'Dia Nacional de Zumbi e da Consciência Negra'], [12, 25, 'Natal']
];
const FERIADOS_MOVEIS_AUS = [[-48, 'Carnaval (segunda-feira)'], [-47, 'Carnaval (terça-feira)'], [-2, 'Sexta-feira Santa'], [60, 'Corpus Christi']];

function calcularPascoa_(ano) {
  const a = ano % 19, b = Math.floor(ano / 100), c = ano % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mes = Math.floor((h + l - 7 * m + 114) / 31);
  const dia = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(ano, mes - 1, dia);
}

function feriadosDoAno_(ano) {
  const lista = FERIADOS_FIXOS_AUS.map(f => ({ data: isoAus_(new Date(ano, f[0] - 1, f[1])), descricao: f[2] }));
  const p = calcularPascoa_(ano);
  FERIADOS_MOVEIS_AUS.forEach(f => lista.push({ data: isoAus_(new Date(p.getFullYear(), p.getMonth(), p.getDate() + f[0])), descricao: f[1] }));
  return lista;
}

/** Map 'yyyy-MM-dd' -> descrição. Aba "Feriados" (Data | Descrição) soma aos automáticos. */
function carregarFeriados_(ss) {
  const feriados = new Map();
  const ano = new Date().getFullYear();
  [ano - 1, ano, ano + 1].forEach(a => feriadosDoAno_(a).forEach(f => feriados.set(f.data, f.descricao)));
  const aba = ss && ss.getSheetByName('Feriados');
  if (aba && aba.getLastRow() > 1) {
    aba.getRange(2, 1, aba.getLastRow() - 1, Math.min(2, aba.getLastColumn())).getValues().forEach(r => {
      const dt = lerDataAus_(r[0]);
      if (dt) feriados.set(isoAus_(dt), String(r[1] || 'Feriado'));
    });
  }
  return feriados;
}

/* =============================================================================
 * UTILITÁRIOS
 * ===========================================================================*/
function normalizarTextoAus_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
}
function isoAus_(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}
function isoParaDataAus_(iso) {
  const p = String(iso).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}
/** Aceita Date, "dd/MM/yyyy" e "yyyy-MM-dd". */
function lerDataAus_(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}
function minutosAus_(h) {
  const m = String(h || '').match(/(\d{1,2}):(\d{2})/);
  return m ? +m[1] * 60 + +m[2] : -1;
}
function ultimoDiaUtilApos_(ini, qtdDias) {
  const d = new Date(ini);
  let contados = 1;
  while (contados < qtdDias) {
    d.setDate(d.getDate() + 1);
    if (d.getDay() !== 0 && d.getDay() !== 6) contados++;
  }
  return d;
}
