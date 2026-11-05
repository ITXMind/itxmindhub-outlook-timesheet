# Hub – Horas do Calendário (Outlook add-in)

Lê automaticamente o calendário do utilizador, apanha só as marcações cujo **assunto** contém um ID de
oportunidade (`OPP-AA-NNNNNNN`, ex.: `OPP-26-0000123`), faz match com projetos do ITXMind Hub e regista-as
para aparecerem como cards na tab **Calendário** do projeto (título da reunião, hora, colaborador e duração).
**Não cria registos de tempo em tarefas.**

As marcações chegam ao Hub **sem ninguém abrir o add-in**: uma sincronização no servidor (de 15 em 15 minutos)
lê os calendários de todos os utilizadores do Hub. O add-in é opcional e serve para ver a lista das OPPs
detetadas (OPP, projeto, nº de marcações e duração total) e forçar uma leitura imediata. Não há pesquisa nem
escolha de datas manual.

## Sincronização automática (servidor)

- **Função** `sync-outlook-calendar` (repo `itxmind-hub`), agendada por `pg_cron` de 15 em 15 minutos.
- Usa o Microsoft Graph com permissão de **aplicação** `Calendars.Read` (client credentials) e lê o calendário de
  cada utilizador do Hub (pelo email do perfil), do mês anterior até 3 meses à frente. Pede só as marcações com
  `OPP-` no assunto (`contains(subject,'OPP-')`).
- Aplica as mesmas regras do add-in: ID `OPP-\d{2}-\d{7}` no assunto, match com o projeto (orçamento do projeto
  ou itens importados), registo só onde o utilizador é membro do projeto (ou admin), sem duplicados
  (`iCalUId|início`).
- **Limpeza:** marcações apagadas, canceladas, sem OPP no assunto ou com a OPP trocada de projeto são removidas do
  Hub no ciclo seguinte (só dentro da janela e só quando a leitura desse utilizador correu toda bem).
- Utilizadores sem caixa de correio Exchange Online ou bloqueados por política de acesso são ignorados.
- Teste manual de um utilizador (resposta com o detalhe):
  `curl -X POST <SUPABASE_URL>/functions/v1/sync-outlook-calendar -H "Authorization: Bearer <SERVICE_ROLE_KEY>" -d '{"email":"nome@itxmind.com"}'`

## Como funciona

- **Login:** dialog Office → `hub.itxmind.com/auth/outlook` → `taskpane/auth-dialog.html` devolve a sessão
  Supabase. A função `outlook-addin-session` entrega ao add-in uma sessão **independente** da do site, por isso
  entrar no add-in não termina a sessão no Hub. Os tokens ficam em RoamingSettings/localStorage e são
  atualizados a cada refresh, pelo que a sessão persiste ao fechar e reabrir o Outlook.
- **Utilizador:** o email no topo abre um menu com **Terminar sessão** (termina só a sessão do add-in).
- **Leitura do calendário:** Microsoft Graph `/me/calendarView` (`Calendars.Read`) via Nested App Auth (MSAL).
  O pedido leva o filtro `contains(subject,'OPP-')`, por isso só são lidas as marcações com OPP no assunto,
  não o calendário todo.
- **Reconhecimento do ID:** apenas `OPP-\d{2}-\d{7}`, só no assunto. Variantes como `OPP-TESTE` ou IDs com
  outro número de dígitos são ignoradas.
- **Match OPP → projeto:** RPC `match_opps_to_projects` (via `quotes.opp_number` do projeto ou
  `project_items.opp_number`), com RLS: só projetos a que o utilizador tem acesso.
- **Registo:** upsert em `project_calendar_events` (chave `iCalUId|início`, sem duplicados).
- **Atualização:** ao abrir lê o mês atual e regista automaticamente (`AUTO_REGISTER` em `taskpane/index.html`;
  `false` pede confirmação). Repete de 5 em 5 minutos e ao reabrir o painel.
- **Calendário no Hub:** cada marcação mostra o título limpo (sem o ID nem a pontuação que sobra, ex.:
  `Batatas (OPP-26-1111111)` → `Batatas`), a hora, o primeiro nome do colaborador e a duração. A limpeza é feita
  na apresentação (`cleanEventTitle`); o assunto original fica guardado.

## Antes de usar

1. **Hub (repo `itxmind-hub`):**
   - aplicar as migrações `20261105120000_project_calendar_events.sql` e
     `20261105140000_sync_outlook_calendar_cron.sql` (`supabase db push`);
   - publicar as funções: `supabase functions deploy outlook-addin-session sync-outlook-calendar`;
   - definir os segredos da sincronização (ver passo 2):
     `supabase secrets set OUTLOOK_SYNC_TENANT_ID=… OUTLOOK_SYNC_CLIENT_ID=… OUTLOOK_SYNC_CLIENT_SECRET=…`;
   - fazer deploy do frontend (`CalendarView.tsx`, `useProjectCalendarEvents.ts`, `OutlookAuth.tsx`,
     `AuthContext.tsx`, `MicrosoftCallback.tsx`, `LoginForm.tsx`). A página `OutlookAuth.tsx` só entrega a sessão
     a callbacks da lista `ALLOWED_ADDIN_CALLBACKS`; o URL deste add-in tem de lá estar.
2. **Entra ID:** registar uma app SPA com permissão delegada `Calendars.Read` e **consentimento de administrador**
   concedido (sem isso aparece "Necessita de aprovação do administrador"), redirect URI
   `brk-multihub://itxmind.github.io`, e pré-autorizar o Outlook. Colocar o Client ID em `GRAPH_CLIENT_ID`
   (`taskpane/index.html`) e no `<WebApplicationInfo>` do `manifest.xml`.
   **Sincronização automática:** na mesma app (ou numa nova) adicionar a permissão de **Aplicação**
   `Calendars.Read` (Microsoft Graph) com consentimento de administrador, e criar um *client secret*
   (Certificates & secrets). O tenant, o client ID e o secret vão para os segredos `OUTLOOK_SYNC_*` acima.
   Esta permissão dá acesso a todas as caixas de correio: restringir com uma *Application Access Policy* do
   Exchange (`New-ApplicationAccessPolicy`) a um grupo de segurança com os utilizadores do Hub.
3. **Alojamento:** publicar esta pasta em HTTPS (GitHub Pages) e ajustar o URL base
   (`itxmind.github.io/itxmindhub-outlook-timesheet`) no `manifest.xml` e em `HUB_LOGIN_URL` (`taskpane/index.html`).
4. Instalar o `manifest.xml` no Outlook (Aplicações → Adicionar suplementos personalizados).

## Atualizar uma versão já instalada

O Outlook guarda em cache o URL do painel. Depois de publicar, espera a cache do Pages (~10 min) e reinicia o
Outlook. Se continuar a aparecer a versão antiga, sobe o `?v=N` dos URLs do taskpane/commands no `manifest.xml`
e reinstala o manifest.
