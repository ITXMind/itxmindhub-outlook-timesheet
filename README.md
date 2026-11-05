# Hub – Horas do Calendário (Outlook add-in)

Lê automaticamente o calendário do utilizador, apanha só as marcações cujo **assunto** contém um ID de
oportunidade (`OPP-AA-NNNNNNN`, ex.: `OPP-26-0000123`), faz match com projetos do ITXMind Hub e regista-as
para aparecerem como cards na tab **Calendário** do projeto (título da reunião, hora, colaborador e duração).
**Não cria registos de tempo em tarefas.**

O painel do add-in mostra apenas a lista das OPPs detetadas e adicionadas ao calendário do Hub
(OPP, projeto, nº de marcações e duração total). Não há pesquisa nem escolha de datas manual.

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
- **Match OPP → projeto:** RPC `match_opps_to_projects` (via `quotes.opp_number` do projeto,
  `project_items.opp_number`, ou nome do projeto igual ao ID / começado por "ID "), com RLS: só projetos a que
  o utilizador tem acesso.
- **Registo:** upsert em `project_calendar_events` (chave `iCalUId|início`, sem duplicados).
- **Atualização:** ao abrir lê o mês atual e regista automaticamente (`AUTO_REGISTER` em `taskpane/index.html`;
  `false` pede confirmação). Repete de 5 em 5 minutos e ao reabrir o painel.
- **Calendário no Hub:** cada marcação mostra o título limpo (sem o ID nem a pontuação que sobra, ex.:
  `Batatas (OPP-26-1111111)` → `Batatas`), a hora, o primeiro nome do colaborador e a duração. A limpeza é feita
  na apresentação (`cleanEventTitle`); o assunto original fica guardado.

## Antes de usar

1. **Hub (repo `itxmind-hub`):**
   - aplicar as migrações `20261105120000_project_calendar_events.sql` e
     `20261105130000_match_opps_by_project_name.sql` (`supabase db push`);
   - publicar a função: `supabase functions deploy outlook-addin-session`;
   - fazer deploy do frontend (`CalendarView.tsx`, `useProjectCalendarEvents.ts`, `OutlookAuth.tsx`,
     `AuthContext.tsx`, `MicrosoftCallback.tsx`, `LoginForm.tsx`). A página `OutlookAuth.tsx` só entrega a sessão
     a callbacks da lista `ALLOWED_ADDIN_CALLBACKS`; o URL deste add-in tem de lá estar.
2. **Entra ID:** registar uma app SPA com permissão delegada `Calendars.Read` e **consentimento de administrador**
   concedido (sem isso aparece "Necessita de aprovação do administrador"), redirect URI
   `brk-multihub://itxmind.github.io`, e pré-autorizar o Outlook. Colocar o Client ID em `GRAPH_CLIENT_ID`
   (`taskpane/index.html`) e no `<WebApplicationInfo>` do `manifest.xml`.
3. **Alojamento:** publicar esta pasta em HTTPS (GitHub Pages) e ajustar o URL base
   (`itxmind.github.io/itxmindhub-outlook-timesheet`) no `manifest.xml` e em `HUB_LOGIN_URL` (`taskpane/index.html`).
4. Instalar o `manifest.xml` no Outlook (Aplicações → Adicionar suplementos personalizados).

## Atualizar uma versão já instalada

O Outlook guarda em cache o URL do painel. Depois de publicar, espera a cache do Pages (~10 min) e reinicia o
Outlook. Se continuar a aparecer a versão antiga, sobe o `?v=N` dos URLs do taskpane/commands no `manifest.xml`
e reinstala o manifest.
