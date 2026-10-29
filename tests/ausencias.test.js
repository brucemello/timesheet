const fs=require('fs'), vm=require('vm'), assert=require('assert');
const SRC=require('path').join(__dirname,'..','src')+'/';
// ---- fake sheets ----
function sheet(name, rows){ return { getName:()=>name, getLastRow:()=>rows.length, getLastColumn:()=>Math.max(...rows.map(r=>r.length)),
  getRange:(r,c,nr,nc)=>{ const sub=()=>rows.slice(r-1,r-1+(nr||1)).map(x=>{const y=x.slice(c-1,c-1+(nc||1)); while(y.length<(nc||1)) y.push(''); return y;});
    return { getDisplayValues:sub, getValues:sub, getValue:()=>sub()[0][0], setValue(){return this}, setBackground(){return this}, setFontColor(){return this}, setFontWeight(){return this} }; } }; }
function book(sheets){ return { getSheets:()=>sheets, getSheetByName:n=>sheets.find(s=>s.getName()===n)||null }; }
const consolidado=[["Período","Nome / Colaborador","Tempo","Função","Situação Férias","Ini Aq","Fim Aq","Limite","Agendamento","Data Início Programação","Data Fim Programação","Saldo"],
 ["","YASMIM LIRA CASTRO","","","Finalizadas","","","","Não Agendados","08/06/2026","22/06/2026","0"],
 ["","EVANDRO ITALO SILVA PEIXOTO","","","Marcadas","","","","Agendados","26/10/2026","09/11/2026","0"],
 ["","CARLA BIANCA SANTOS LATORRE","","","-","","","","Não Agendados","-","-","30"],
 ["","EDUARDO NUNES CIRQUEIRA","","","Sugestão do Dev (Padrão)","","","","Sugestão","05/10/2026","10/10/2026","6"]];
const folgas=[["Colaborador","Benefício Utilizado","Data da Folga (Início)","Quantidade de Dias","Data de Retorno"],
 ["FABIO SANTOS DE MENDONCA","GIFTCARD","2026-07-10","1","2026-07-13"],["EVANDRO ITALO SILVA PEIXOTO","RESETDAY","2026-07-17","1","2026-07-20"]];
const contatos=[["Nome","Email"],["CARLOS EDUARDO NONATO DA SILVA","silva.carlos@totvs.com.br"]];
const agenda=[["Título","Criador","Data Início","Hora Início","Data Término","Hora Término","ID"],
 ["Ausência: EDUARDO NUNES CIRQUEIRA [Férias]","EDUARDO NUNES CIRQUEIRA","05/10/2026","Dia Inteiro","12/10/2026","Dia Inteiro","a"],
 ["Fisioterapia Ombro | Fabio Mendonça","FABIO MENDONCA","14/10/2026","08:30:00","14/10/2026","10:30:00","b"],
 ["Keilla - Exame Medico ","Keilla","30/09/2026","13:00:00","30/09/2026","15:30:00","c"],
 ["Keilla - Exame Medico ","Keilla","30/09/2026","14:30:00","30/09/2026","17:30:00","c"],
 ["Consulta - Pâmela","Consulta","05/10/2026","11:00:00","05/10/2026","12:00:00","d"],
 ["Edu - Exame Periódico","Edu","16/10/2026","08:00:00","16/10/2026","10:00:00","e"],
 ["Carlão - Consulta Médica","Carlão","22/09/2026","08:00:00","22/09/2026","12:00:00","f"],
 ["","SILVA CARLOS","07/10/2026","13:30:00","07/10/2026","16:30:00","g"],
 ["Ausência: CAIQUE HENRIQUE CARLOS [Marcadas]","CAIQUE HENRIQUE CARLOS","03/11/2026","Dia Inteiro","23/11/2026","Dia Inteiro","h"],
 ["Ausência: CAIQUE HENRIQUE CARLOS [Marcadas]","CAIQUE HENRIQUE CARLOS","03/11/2026","Dia Inteiro","23/11/2026","Dia Inteiro","i"],
 ["Keilla - Atualização de Documentos pessoais.","Keilla","03/09/2026","11:30:00","03/09/2026","17:30:00","j"],
 ["teste","BRUCE MELLO","08/09/2026","08:00:00","08/09/2026","08:45:00","k"]];
const depara=[["ID Jira","Nome","E-mail","Apelidos (agenda)"],
 ["u1","Yasmim Lira Castro","yasmim.castro@totvs.com.br",""],["u2","Eduardo Nunes Cirqueira","eduardo.cirqueira@totvs.com.br",""],
 ["u3","Keilla Silva de Freitas","keilla.freitas@totvs.com.br",""],["u4","Caique Henrique Carlos","caique.carlos@totvs.com.br",""],
 ["u5","Evandro Italo Silva Peixoto","evandro.peixoto@totvs.com.br",""],["u6","Fabio Santos de Mendonca","fabio.mendonca@totvs.com.br",""],
 ["u7","Pamela Bernardo de Souza","pamela.bsouza@totvs.com.br",""],["u8","Carlos Eduardo Nonato da Silva","silva.carlos@totvs.com.br","Carlão"],
 ["u9","Carla Bianca Santos Latorre","carla.latorre@totvs.com.br",""]];
const ferias=book([sheet('Controle Consolidado - 2026',consolidado),sheet('Contatos',contatos),sheet('Folgas e Emendas - 2026',folgas)]);
const agendaBook=book([sheet('Página1',agenda)]);
const ts=book([sheet('DePara_Usuarios',depara)]);
const pad=n=>String(n).padStart(2,'0');
const ctx={ console, Logger:{log:()=>{}},
  Session:{getScriptTimeZone:()=>'America/Sao_Paulo', getActiveUser:()=>({getEmail:()=>'x'})},
  Utilities:{formatDate:(d,tz,f)=> f==='yyyy-MM-dd'? `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}` : `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()}`},
  SpreadsheetApp:{ openByUrl:u=> u.includes('1ox7') ? ferias : agendaBook, getActiveSpreadsheet:()=>ts } };
vm.createContext(ctx);
for (const f of ['Código.js','Ausencias.js']) vm.runInContext(fs.readFileSync(SRC+f,'utf8'), ctx, {filename:f});
const R=c=>vm.runInContext(c,ctx);
ctx.devs=depara.slice(1).map(r=>({email:r[2],nome:r[1]}));
const aus=R('carregarAusencias_(SpreadsheetApp.getActiveSpreadsheet(), devs)');
ctx.aus=aus;
const av=(e,d)=>R(`avaliarStatusEAusenciaDev_('${e}','x',lerDataAus_('${d}'),aus,6)`);
console.log('não vinculados:', aus.naoVinculados);
// Férias finalizadas contam (antes eram ignoradas)
assert.equal(av('yasmim.castro@totvs.com.br','2026-06-10').cargaEsperadaNoDia,0);
// Término exclusivo: Eduardo férias até 11/10 (dom) -> 12/10 normal (sem ausência)
assert.equal(av('eduardo.cirqueira@totvs.com.br','2026-10-09').eDiaInteiro,true);
assert.equal(av('eduardo.cirqueira@totvs.com.br','2026-10-12').eventosAuditoria.length,0);
// Caique: 03/11..22/11 (não 23)
assert.equal(av('caique.carlos@totvs.com.br','2026-11-20').eDiaInteiro,true);
assert.equal(av('caique.carlos@totvs.com.br','2026-11-23').eDiaInteiro,false);
// Folgas e Emendas
assert.equal(av('fabio.mendonca@totvs.com.br','2026-07-10').eDiaInteiro,true);
assert.equal(av('evandro.peixoto@totvs.com.br','2026-07-17').eDiaInteiro,true);
// parcial: fisio 2h -> meta 4
assert.equal(av('fabio.mendonca@totvs.com.br','2026-10-14').cargaEsperadaNoDia,4);
// sobreposição Keilla 13:00-15:30 + 14:30-17:30 = 4.5h (não 5.5)
assert.equal(av('keilla.freitas@totvs.com.br','2026-09-30').horasAbatidas,4.5);
// "Consulta - Pâmela" (criador "Consulta") -> Pamela 1h
assert.equal(av('pamela.bsouza@totvs.com.br','2026-10-05').horasAbatidas,1);
// "Edu" -> Eduardo Cirqueira, não Carlos Eduardo
assert.equal(av('eduardo.cirqueira@totvs.com.br','2026-10-16').horasAbatidas,2);
assert.equal(av('silva.carlos@totvs.com.br','2026-10-16').horasAbatidas,0);
// Apelido Carlão + título vazio com criador SILVA CARLOS
assert.equal(av('silva.carlos@totvs.com.br','2026-09-22').horasAbatidas,4);
assert.equal(av('silva.carlos@totvs.com.br','2026-10-07').horasAbatidas,3);
// Keilla 6h de ausência -> meta 0 mas não dia inteiro
assert.equal(av('keilla.freitas@totvs.com.br','2026-09-03').cargaEsperadaNoDia,0);
// Carla sem férias ("-") não vira ausência
assert.equal(aus.registros.filter(r=>r.email==='carla.latorre@totvs.com.br').length,0);
// Feriados automáticos
const fer=R('carregarFeriados_(null)');
for (const d of ['2026-11-02','2026-11-20','2026-04-03','2026-06-04','2026-02-16']) assert.ok(fer.has(d), d);
// Sem dev no Contatos com email diferente: Evandro (evandro.italo no Contatos) casa por nome
assert.equal(av('evandro.peixoto@totvs.com.br','2026-10-28').eDiaInteiro,true);
console.log('TODOS OS TESTES PASSARAM');
