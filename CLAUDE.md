# Timesheet — instruções para o Claude

Projeto Google Apps Script (planilha + Web App) de controle de apontamento de horas (Tempo API), alertas diários/semanais por e-mail.

## Padrão de trabalho (autonomia)
- Commitar e dar push **direto na `main`** — o workflow *Deploy Apps Script* publica sozinho (clasp push + atualiza a implantação `DEPLOYMENT_ID`).
- Antes de cada commit, testar sempre:
  1. `node --check` em cada `src/*.js` (copiar para um `.js` temporário se o nome tiver acento).
  2. `jq empty src/appsscript.json`.
  3. `node tests/ausencias.test.js` (roda também no deploy; adicione casos novos em `tests/*.test.js`).
  4. Quando a lógica for pura (datas, cálculos de horas, pendências), testar com um script Node isolado simulando as funções.
- Depois do push, acompanhar o run via `gh api repos/brucemello/timesheet/actions/runs` e confirmar `success` (o deploy inclui verificação de que o código publicado ficou igual ao repositório).
- Se alguém editou no editor do Google, rodar o workflow *Importar do Apps Script* antes de mexer.
- Mensagens de commit em português.

## Estrutura
- `src/Código.js` — lógica principal
- `src/Ausencias.js` — férias, folgas, agenda de ausências e feriados (meta do dia de cada dev); lógica trazida do Capacidade Não Produtiva
- `src/Painel.html` + `src/PainelServidor.js` — painel de controle (Web App `doGet` e modal pelo menu): dashboard, cobranças individuais, ausências, cadastros e configurações
- `src/Apuracao.js` — apuração única do período (Tempo + ausências + feriados), cache, envio e log de cobranças
- Área do administrador (Configurações): links de planilha, Apps Script e GitHub; admin = propriedade `CONFIG_ADMINS` ou o dono da implantação (checado no servidor)
- `tests/` — testes Node com dados reais de exemplo (Google simulado)
- `src/bkp_gs.html`, `src/bkp_html.html` — backups antigos (não usados)
