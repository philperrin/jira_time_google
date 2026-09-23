# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A standalone Google Apps Script web app (not a Node/npm project) that connects to the Jira REST API to streamline weekly time logging. It pulls active Jira issues, imports matching Google Calendar events, and posts time entries as Jira worklogs, all from a browser-based interface served by Apps Script.

## Development workflow

There is no build system, package manager, linter, or test runner — this is plain server-side `.js`/`.gs`-style Apps Script plus HTML templates, edited and run directly in the Apps Script editor.

- **No `clasp` CLI** — per company policy, deployment is manual: copy each `.js` file into an Apps Script script file and each `.html` file into an Apps Script HTML file via the online editor (script.google.com), then create/update a deployment there. See README.md for the full step-by-step.
- **Automated tests live in `Tests.js`, run manually.** Before creating or updating a deployment, run `runAllTests` from the Apps Script editor's function picker and confirm no `FAIL:` lines in the execution log, then walk the manual smoke-test checklist in `docs/superpowers/specs/SPEC-2026-09-23-test-suite.md` (§5). `Debug.js` contains separate developer-only diagnostic helpers (e.g. `validateImportCalendarEvents`), also run manually from the function picker, not from a CLI.
- **Manifest (`appsscript.json`)** must be kept in sync between this repo and the Apps Script project's manifest editor — it defines the timezone, OAuth scopes, and web app execution settings.
- When editing, changes only take effect in the live app after they're pasted into the Apps Script editor and redeployed; editing files in this repo alone does not affect any running deployment.

## Architecture

### Server / client split

- **`Code.js`** holds all server-side logic — every Jira REST API call, Google Calendar interaction, and Apps Script user-properties (per-user storage) access. Functions suffixed with `_` (e.g. `getAuthHeader_`, `getWorklogTotals_`, `parseTimeSpentHours_`) are private helpers, not called from client code.
- **`JavaScript.html`** is client-side JS shared across all tabs, wrapped in `<script>` and pulled into the page via Apps Script's HTML templating `include()` mechanism (see `Code.js`'s `include()` and `doGet()`). Client code calls server functions via `google.script.run`.
- **`Index.html`** is the app shell: tab navigation and the template that assembles every other `.html` file at request time.
- **`Stylesheet.html`** is shared CSS, similarly included into `Index.html`.
- Each **`Tab*.html`** file (`TabConfig`, `TabAssignments`, `TabTimecard`, `TabCreateIssue`, `TabWorklog`, `TabChangelog`) is the markup for one tab in the UI; `TabChangelog.html` pulls its content from `ChangelogData.html`, which defines the `CHANGELOG_ITEMS` data used to render the in-app changelog (the project's backlog/changelog lives here, not in a separate file).

### Data flow

1. **Config tab** stores the Jira base URL, Atlassian API key, and calendar-color → project allocation mapping in Apps Script user properties, scoped per Google account (`getConfig`/`saveJiraUrl`/`saveApiKey`/`getAllocation`/`saveAllocation` in `Code.js`). No config is shared across users of the same deployment.
2. **Assignments tab** calls `getJiraIssues` to fetch active issues from Jira (paginated internally), grouped by project.
3. **Timecard tab** calls `importCalendarEvents` to pull Google Calendar events in a date range and match them to projects by allocation color, lets the user assign Jira issues to entries, then calls `sendTimeEntries` to post worklogs to Jira.
4. **Create Issue tab** calls `makeJira` to create new Jira issues directly.
5. **Worklog tab** calls `getWorklogs`/`getWorklogTotals_` to pull posted worklogs for a selected year and render a month × project pivot with CSV export.
6. Calendar scheduling (`scheduleCalendarEvents`, `nextScheduleWorkDay_`, `scheduleDayCapacityMs_`) enforces a 9am–5pm workday capped at 8h/day, skipping weekends, and rolls overflow work to the next work day.

### Hardcoded business rules (in `Code.js`)

These are intentional and should not be "fixed" without confirming with the user first:

- `getJiraIssues`: only fetches `issuetype IN (Story, Task, Sub-Task)`; excludes Epics/Bugs/others. `Done` issues are included only if `updated >= -7d`. Projects whose name contains `'Archive'` or `'Managed Services Internal'` are excluded.
- `getWorklogs`: operates on a single calendar year (`Jan 1`–`Dec 31`); the UI only offers the current year or 3 prior years.
- All Jira API pagination uses `maxResults=100` per page, handled internally.
- Worklog fetches use `UrlFetchApp.fetchAll()` to parallelize requests across issues.
- Script timezone is `America/Denver` (set in `appsscript.json`); all times are converted to UTC before being posted to Jira.
