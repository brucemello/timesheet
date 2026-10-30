// Teste de ponta a ponta do painel: apuração, cobrança individual, cadastros e configurações.
const assert = require('assert');
const vm = require('vm');
const { Planilha, Pasta, criarContexto } = require('./googleFake');

const FERIAS = '1ox7bG8Z1NqSJVINDF2EjAq00XpjnbvSyD4SapgXvjD8', AGENDA = '1oZTQ3XhiYgTmQFtDKcrXKHTKRWf8sagUfCsJi-V-t2A';
const ferias = new Pasta([
  new Planilha('Controle Consolidado - 2026', [
    ['Período', 'Nome / Colaborador', 'Tempo', 'Função', 'Situação Férias', 'a', 'b', 'c', 'Agendamento', 'Data Início Programação', 'Data Fim Programação'],
    ['', 'EVANDRO ITALO SILVA PEIXOTO', '', '', 'Marcadas', '', '', '', 'Agendados', '26/10/2026', '09/11/2026']]),
  new Planilha('Folgas e Emendas - 2026', [['Colaborador', 'Benefício Utilizado', 'Data da Folga (Início)', 'Quantidade de Dias'], ['FABIO SANTOS DE MENDONCA', 'GIFTCARD', '2026-10-02', '1']]),
  new Planilha('Contatos', [['Nome', 'Email']])
], FERIAS);
const agenda = new Pasta([new Planilha('Página1', [
  ['Título', 'Criador', 'Data Início', 'Hora Início', 'Data Término', 'Hora Término', 'ID'],
  ['Fisioterapia Ombro | Fabio Mendonça', 'FABIO MENDONCA', '07/10/2026', '08:30:00', '07/10/2026', '10:30:00', 'a'],
  ['Carlão - Consulta Médica', 'Carlão', '08/10/2026', '08:00:00', '08/10/2026', '12:00:00', 'b']
])], AGENDA);
const ts = new Pasta([new Planilha('DePara_Usuarios', [
  ['ID Jira', 'Nome', 'E-mail'],
  ['u5', 'Evandro Italo Silva Peixoto', 'evandro.peixoto@totvs.com.br'],
  ['u6', 'Fabio Santos de Mendonca', 'fabio.mendonca@totvs.com.br'],
  ['u8', 'Carlos Eduardo Nonato da Silva', 'silva.carlos@totvs.com.br'],
  ['lider', 'Bruce Mello', 'bruce.mello@totvs.com.br']
])], 'TS');

const ctx = criarContexto({ ativa: ts, externas: { [FERIAS]: ferias, [AGENDA]: agenda, TS: ts },
  props: { CONFIG_JIRA_TOKEN: 'tok-1234', CONFIG_EMAIL_LIDER: 'bruce.mello@totvs.com.br', CONFIG_PERIODO_AUTO: 'NAO', CONFIG_MES_DE: '2026-10', CONFIG_MES_ATE: '2026-10', CONFIG_HORA_MINIMA: '6' } });
const R = c => vm.runInContext(c, ctx);

// Apontamentos fake do Tempo: Fabio 6h todo dia útil até 09/10 (4h no dia da fisio), Carlos só 2 dias, Bruce nada
const wls = [];
for (let d = 1; d <= 9; d++) {
  const iso = `2026-10-${String(d).padStart(2, '0')}`, dow = new Date(2026, 9, d).getDay();
  if (dow === 0 || dow === 6) continue;
  wls.push({ author: { name: 'u6' }, started: iso, timeSpentSeconds: (d === 7 ? 4 : 6) * 3600 });
}
wls.push({ author: { name: 'u8' }, started: '2026-10-01', timeSpentSeconds: 6 * 3600 });
wls.push({ author: { name: 'u8' }, started: '2026-10-08', timeSpentSeconds: 2 * 3600 });
wls.push({ author: { name: 'u99', displayName: 'Fulano Externo', emailAddress: 'fulano@totvs.com.br' }, started: '2026-10-05', timeSpentSeconds: 3600 });
ctx.buscarWorklogsTempoAPI_ = () => wls;

// "hoje" fixo em 10/10/2026 (sábado) para o período terminar em 09/10
const RealDate = Date;
ctx.Date = class extends RealDate { constructor(...a) { if (!a.length) return new RealDate(2026, 9, 10, 10, 0); super(...a); } static now() { return new RealDate(2026, 9, 10).getTime(); } };
vm.runInContext('Date = this.Date', ctx);

// 1) apelido "Carlão" ainda não cadastrado -> não vinculado
let r = R('obterDadosPainel({})');
assert.ok(r.sucesso, r.mensagem);
let ap = r.apuracao;
assert.strictEqual(JSON.stringify(ap.ausencias.naoVinculados), JSON.stringify(['Carlão - Consulta Médica']));
assert.ok(ts.getSheetByName('DePara_Usuarios').dados[0][4] === 'Cobrar (SIM/NÃO)', 'cabeçalho Cobrar criado');
const dev = e => ap.devs.find(d => d.email === e);
// período: 01..09/10 (hoje=10). dias úteis 1,2,5,6,7,8,9 = 7
assert.strictEqual(ap.dias.length, 7);
// Fabio: folga 02/10, fisio 07/10 (meta 4) -> em dia
assert.strictEqual(dev('fabio.mendonca@totvs.com.br').status, 'OK', JSON.stringify(dev('fabio.mendonca@totvs.com.br').totais));
assert.strictEqual(dev('fabio.mendonca@totvs.com.br').totais.diasAusencia, 1);
// Evandro: férias só a partir de 26/10 -> 7 dias pendentes
assert.strictEqual(dev('evandro.peixoto@totvs.com.br').totais.diasPendentes, 7);
// Carlos: 08/10 com 2h e meta 6 (apelido ainda não cadastrado)
const iCarlos08 = ap.dias.findIndex(d => d.iso === '2026-10-08');
assert.strictEqual(dev('silva.carlos@totvs.com.br').dias[iCarlos08].m, 6);
// fora do cadastro aparece marcado
assert.ok(dev('fulano@totvs.com.br').foraDePara);

// 2) cadastra o apelido e Cobrar = NÃO para o líder pelo painel
r = R('obterCadastroPainel("colaboradores")');
assert.ok(r.sucesso, r.mensagem);
const linhas = r.linhas.map(l => Object.assign({}, l));
linhas.find(l => l.email === 'silva.carlos@totvs.com.br').apelidos = 'Carlão';
linhas.find(l => l.email === 'bruce.mello@totvs.com.br').cobrar = 'NÃO';
ctx.__linhas = linhas;
r = R('salvarCadastroPainel("colaboradores", __linhas)');
assert.ok(r.sucesso, r.mensagem);
// e-mail repetido é recusado
ctx.__dup = linhas.concat([{ idJira: 'x', nome: 'X', email: 'bruce.mello@totvs.com.br', apelidos: '', cobrar: 'SIM' }]);
assert.strictEqual(R('salvarCadastroPainel("colaboradores", __dup)').sucesso, false);

r = R('obterDadosPainel({})'); ap = r.apuracao;
assert.strictEqual(JSON.stringify(ap.ausencias.naoVinculados), JSON.stringify([]));
assert.strictEqual(dev('silva.carlos@totvs.com.br').dias[iCarlos08].m, 2, 'consulta de 4h abatida');
assert.strictEqual(dev('silva.carlos@totvs.com.br').dias[iCarlos08].s, 'OK');
assert.strictEqual(dev('bruce.mello@totvs.com.br').cobrar, false);

// 3) ausência manual + feriado manual pelo painel
ctx.__aus = [{ identificador: 'evandro.peixoto@totvs.com.br', inicio: '2026-10-05', fim: '2026-10-06', horas: '', motivo: 'Atestado médico' }];
assert.ok(R('salvarCadastroPainel("ausencias", __aus)').sucesso);
ctx.__fer = [{ data: '2026-10-09', descricao: 'Ponto facultativo' }];
assert.ok(R('salvarCadastroPainel("feriados", __fer)').sucesso);
r = R('obterDadosPainel({})'); ap = r.apuracao;
assert.strictEqual(dev('evandro.peixoto@totvs.com.br').totais.diasPendentes, 4, '7 - 2 de atestado - 1 feriado');
assert.strictEqual(ap.dias.find(d => d.iso === '2026-10-09').feriado, 'Ponto facultativo');
assert.ok(ap.ausencias.registros.some(x => x.tipo === 'Atestado'));
// cadastro inválido
ctx.__ruim = [{ identificador: 'x', inicio: '2026-10-09', fim: '2026-10-01' }];
assert.strictEqual(R('salvarCadastroPainel("ausencias", __ruim)').sucesso, false);

// 4) prévia e cobrança individual
r = R('previewCobrancaPainel({ email: "evandro.peixoto@totvs.com.br" })');
assert.ok(r.sucesso && /Evandro/.test(r.html) && /01\/10\/2026/.test(r.html), 'prévia com nome e período');
assert.ok(!/<b>05\/10<\/b>/.test(r.html) && /<b>01\/10<\/b>/.test(r.html), 'dia de atestado não aparece na cobrança');
assert.ok(/Pendências de apontamento/.test(r.html) && /Tempo\.jspa/.test(r.html), 'layout novo com botão para o Tempo');
assert.ok(r.html.indexOf('/secure/Tempo.jspa#/teams/team/644/timesheet?columns=WORKED_COLUMN&amp;columns=PORTION_OF_REQUIRED_COLUMN&amp;dateDisplayType=days&amp;from=2026-10-01&amp;groupBy=worker') > 0, 'link do timesheet do time');
assert.ok(/teamId=644&amp;to=2026-10-\d\d&amp;viewType=TIMESHEET/.test(r.html), 'link termina no fim do período');
assert.strictEqual(R('formatarHorasLegivel_(0.33)'), '20min');
assert.strictEqual(R('formatarHorasLegivel_(4.5)'), '4h30');
r = R('previewCobrancaPainel({ email: "fabio.mendonca@totvs.com.br" })');
assert.ok(r.sucesso && !r.html, 'quem está em dia não tem prévia');

r = R('enviarCobrancasPainel({ emails: ["evandro.peixoto@totvs.com.br"], teste: true })');
assert.ok(r.sucesso, r.mensagem);
assert.strictEqual(ctx.emails.length, 1);
assert.strictEqual(ctx.emails[0].to, 'bruce.mello@totvs.com.br', 'teste vai só para o líder');
assert.ok(/^\[TESTE/.test(ctx.emails[0].subject));

r = R('enviarCobrancasPainel({ emails: ["evandro.peixoto@totvs.com.br", "fabio.mendonca@totvs.com.br"] })');
assert.ok(r.sucesso, r.mensagem);
assert.strictEqual(ctx.emails.length, 2, 'só quem tem pendência recebe');
assert.strictEqual(ctx.emails[1].to, 'evandro.peixoto@totvs.com.br');
assert.strictEqual(ctx.emails[1].cc, 'bruce.mello@totvs.com.br');
assert.ok(/Sem pendência no período: 1/.test(r.mensagem));
const log = ts.getSheetByName('Log_Cobrancas').dados;
assert.strictEqual(log.length, 3, 'cabeçalho + teste + envio');
assert.strictEqual(log[2][6], 'Painel (lote)');
r = R('obterDadosPainel({})');
assert.ok(r.apuracao.devs.find(d => d.email === 'evandro.peixoto@totvs.com.br').ultimaCobranca, 'última cobrança aparece');
assert.strictEqual(R('enviarCobrancasPainel({ emails: [] })').sucesso, false);

// 4b) cobrança só de dias escolhidos
r = R('previewCobrancaPainel({ email: "evandro.peixoto@totvs.com.br", dias: ["2026-10-01", "2026-10-07"] })');
assert.ok(r.sucesso && r.dias === 2 && /<b>01\/10<\/b>/.test(r.html) && !/<b>02\/10<\/b>/.test(r.html), 'prévia só com os dias escolhidos');
assert.ok(!R('previewCobrancaPainel({ email: "evandro.peixoto@totvs.com.br", dias: ["2026-10-05"] })').html, 'dia não pendente não entra');
const antes = ctx.emails.length;
r = R('enviarCobrancasPainel({ emails: ["evandro.peixoto@totvs.com.br"], dias: { "evandro.peixoto@totvs.com.br": ["2026-10-08"] } })');
assert.ok(r.sucesso, r.mensagem);
const ultimo = ctx.emails[ctx.emails.length - 1];
assert.strictEqual(ctx.emails.length, antes + 1);
assert.ok(/<b>08\/10<\/b>/.test(ultimo.htmlBody) && !/<b>01\/10<\/b>/.test(ultimo.htmlBody), 'e-mail só com o dia escolhido');
const logTs = ts.getSheetByName('Log_Cobrancas').dados;
assert.strictEqual(logTs[logTs.length - 1][4], '08/10');
assert.ok(/dias escolhidos/.test(logTs[logTs.length - 1][6]));
assert.ok(/<b>01\/10<\/b>/.test(R('previewCobrancaPainel({ email: "evandro.peixoto@totvs.com.br" })').html), 'sem filtro continua todos');

// 4c) dias dispensados: não são mais cobrados, nem na rotina semanal
r = R('dispensarDiasPainel({ email: "evandro.peixoto@totvs.com.br", dias: ["2026-10-01", "2026-10-02"], motivo: "Avaliado com o líder" })');
assert.ok(r.sucesso && /2 dia/.test(r.mensagem), r.mensagem);
assert.ok(/já estavam/.test(R('dispensarDiasPainel({ email: "evandro.peixoto@totvs.com.br", dias: ["2026-10-01"] })').mensagem), 'não duplica');
r = R('obterDadosPainel({})'); ap = r.apuracao;
const evd = dev('evandro.peixoto@totvs.com.br');
const i01 = ap.dias.findIndex(d => d.iso === '2026-10-01');
assert.strictEqual(evd.dias[i01].s, 'DISP'); assert.ok(/Avaliado com o líder/.test(evd.dias[i01].a));
assert.strictEqual(evd.totais.diasPendentes, 2, '4 pendentes - 2 dispensados'); assert.strictEqual(evd.totais.diasDispensados, 2);
assert.ok(!/<b>01\/10<\/b>/.test(R('previewCobrancaPainel({ email: "evandro.peixoto@totvs.com.br" })').html), 'dispensado fora do e-mail');
r = R('obterCadastroPainel("dispensas")');
assert.ok(r.sucesso && r.linhas.length === 2 && r.linhas[0].data === '2026-10-01', 'aparece no cadastro');
ctx.__disp = r.linhas.filter(l => l.data === '2026-10-02');
assert.ok(R('salvarCadastroPainel("dispensas", __disp)').sucesso, 'desfazer pelo cadastro');
r = R('obterDadosPainel({})'); ap = r.apuracao;
assert.strictEqual(dev('evandro.peixoto@totvs.com.br').totais.diasPendentes, 3, 'voltou a cobrar 01/10');

// 5) cobrança semanal respeita Cobrar = NÃO e quem está em dia
ctx.emails.length = 0;
R('executarVerificacaoEAlertaSemanais(false)');
const destinos = ctx.emails.map(m => m.to);
assert.ok(!destinos.includes('fabio.mendonca@totvs.com.br'));
const semanalEvandro = ctx.emails.find(m => m.to === 'evandro.peixoto@totvs.com.br');
assert.ok(semanalEvandro && !/<b>02\/10<\/b>/.test(semanalEvandro.htmlBody), 'semanal não cobra dia dispensado');
assert.ok(destinos.includes('evandro.peixoto@totvs.com.br'));
assert.strictEqual(destinos.filter(d => d === 'bruce.mello@totvs.com.br').length, 1, 'líder (Cobrar=NÃO) só recebe o consolidado');

// 6) configurações: token não volta para a tela; token em branco mantém o atual
r = R('obterConfigPainel()');
assert.ok(r.sucesso && r.config.tokenConfigurado && r.config.tokenFinal === '1234' && !('jiraToken' in r.config));
r = R('salvarConfigPainel({ jiraBaseUrl: "https://jiraproducao.totvs.com.br/", teamId: "644", jiraToken: "", horaMinima: "7", emailLider: "Bruce.Mello@totvs.com.br", periodoAuto: true, urlFerias: "", urlAgenda: "" })');
assert.ok(r.sucesso, r.mensagem);
assert.strictEqual(ctx.props.CONFIG_JIRA_TOKEN, 'tok-1234');
assert.strictEqual(ctx.props.CONFIG_JIRA_BASE_URL, 'https://jiraproducao.totvs.com.br');
assert.strictEqual(R('obterConfiguracoesAtuais().mesDe'), '2026-10');
assert.strictEqual(R('salvarConfigPainel({ jiraBaseUrl: "x", teamId: "1", horaMinima: "6" })').sucesso, false);

// 7) rotinas e abas da planilha
r = R('executarAcaoPainel("ativarRotinas")');
assert.ok(r.sucesso && r.rotinas.filter(x => x.gatilhos).length === 3);
assert.ok(R('executarAcaoPainel("desativarRotinas")').rotinas.every(x => !x.gatilhos));
r = R('executarAcaoPainel("extrair")');
assert.ok(r.sucesso, r.mensagem);
['Resumo_Mensal_Devs', 'Horas_Pendentes', 'Auditoria_Ausencias', 'Extracao_Tempo'].forEach(a => assert.ok(ts.getSheetByName(a).getLastRow() > 1, a));
assert.strictEqual(R('executarAcaoPainel("inexistente")').sucesso, false);

// 8) Web App: sem planilha ativa usa o ID memorizado
ctx.SpreadsheetApp.getActiveSpreadsheet = () => null;
assert.ok(R('obterDadosPainel({})').sucesso);

// 9) Área do administrador: só o admin recebe os links
r = R('obterAreaAdminPainel()');
assert.ok(r.sucesso && /script\.google\.com\/home\/projects\/SCRIPT123\/edit/.test(JSON.stringify(r.grupos)), 'admin vê o editor');
assert.ok(R('obterDadosPainel({})').contexto.admin);
ctx.usuarioAtivo = 'william.palves@totvs.com.br';
assert.strictEqual(R('obterAreaAdminPainel()').sucesso, false, 'não admin é barrado');
r = R('obterDadosPainel({})');
assert.ok(!r.contexto.admin && !r.contexto.urlPlanilha, 'não admin não recebe link da planilha');
ctx.props.CONFIG_ADMINS = 'william.palves@totvs.com.br, bruce.mello@totvs.com.br';
assert.ok(R('obterAreaAdminPainel()').sucesso, 'lista de admins na propriedade');
delete ctx.props.CONFIG_ADMINS; ctx.usuarioAtivo = 'bruce.mello@totvs.com.br';

// Dados de exemplo para o teste visual do painel
if (process.env.EXPORTAR_DADOS) require('fs').writeFileSync(process.env.EXPORTAR_DADOS, JSON.stringify({ painel: R('obterDadosPainel({})'), config: R('obterConfigPainel()'), cad: R('obterCadastroPainel("colaboradores")'), prev: R('previewCobrancaPainel({ email: "evandro.peixoto@totvs.com.br" })'), admin: R('obterAreaAdminPainel()') }));
console.log('TODOS OS TESTES DO PAINEL PASSARAM');
