# Hub – Horas do Calendário (Outlook add-in)

Lê as marcações do calendário do utilizador, filtra as que contêm um ID de oportunidade
(`OPP-YY-NNNNNN`), faz match com projetos do ITXMind Hub e regista-as para aparecerem como
cards (utilizador + tempo) na tab **Calendário** do projeto. Não cria registos de tempo em tarefas.

## Como funciona

- **Login:** igual ao add-in de anexos — dialog Office → `hub.itxmind.com/auth/outlook` →
  `taskpane/auth-dialog.html` devolve a sessão Supabase; guardada em RoamingSettings/localStorage.
- **Calendário:** Microsoft Graph `/me/calendarView` (`Calendars.Read`) via Nested App Auth (MSAL).
- **Match OPP → projeto:** RPC `match_opps_to_projects` (via `quotes.opp_number` do projeto ou
  `project_items.opp_number`), com RLS: só projetos a que o utilizador tem acesso.
- **Registo:** upsert em `project_calendar_events` (chave `iCalUId|início`, sem duplicados).
- Ao abrir lê o mês atual e regista automaticamente (`AUTO_REGISTER` em `taskpane/index.html`; `false` pede confirmação). A leitura é sempre automática (mês atual, de 5 em 5 minutos e ao reabrir o painel), sem pesquisa manual.

## Antes de usar

1. **Hub:** aplicar a migração `supabase/migrations/20261105120000_project_calendar_events.sql`
   (repo `itxmind-hub`). A tab Calendário já mostra os dados (`CalendarView.tsx`).
2. **Entra ID:** registar uma app SPA com permissão delegada `Calendars.Read`, redirect URI
   `brk-multihub://itxmind.github.io`, e pré-autorizar o Outlook. Colocar o Client ID em
   `GRAPH_CLIENT_ID` (`taskpane/index.html`) e no `<WebApplicationInfo>` do `manifest.xml`.
3. **Alojamento:** publicar esta pasta em HTTPS e ajustar o URL base (`itxmind.github.io/itxmindhub-outlook-timesheet`)
   no `manifest.xml`, `HUB_LOGIN_URL` em `taskpane/index.html`.
4. Instalar o `manifest.xml` no Outlook (Aplicações → Adicionar suplementos personalizados).
