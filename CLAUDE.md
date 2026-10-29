# Timesheet — instruções para o Claude

Projeto Google Apps Script (planilha + Web App) de controle de apontamento de horas (Tempo API), alertas diários/semanais por e-mail.

## Padrão de trabalho (autonomia)
- Commitar e dar push **direto na `main`** — o workflow *Deploy Apps Script* publica sozinho (clasp push + atualiza a implantação `DEPLOYMENT_ID`).
- Antes de cada commit, testar sempre:
  1. `node --check` em cada `src/*.js` (copiar para um `.js` temporário se o nome tiver acento).
  2. `jq empty src/appsscript.json`.
  3. Quando a lógica for pura (datas, cálculos de horas, pendências), testar com um script Node isolado simulando as funções.
- Depois do push, acompanhar o run via `gh api repos/brucemello/timesheet/actions/runs` e confirmar `success` (o deploy inclui verificação de que o código publicado ficou igual ao repositório).
- Se alguém editou no editor do Google, rodar o workflow *Importar do Apps Script* antes de mexer.
- Mensagens de commit em português.

## Estrutura
- `src/Código.js` — lógica principal
- `src/Interface.html` — painel lateral
- `src/bkp_gs.html`, `src/bkp_html.html` — backups antigos (não usados)
