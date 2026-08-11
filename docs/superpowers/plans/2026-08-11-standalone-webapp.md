# Standalone Google Apps Script Web App Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Port the Google Sheets–bound Apps Script into a standalone web app (no spreadsheet) that performs the same Jira time-tracking workflow entirely through a browser UI.

**Architecture:** A single Google Apps Script standalone project serves an SPA via `doGet()`. Server-side `Code.js` functions handle all Jira REST API calls, Google Calendar access, and `PropertiesService` persistence; client-side HTML/CSS/JS renders tabbed UI and calls server functions via `google.script.run`. The Allocation configuration (calendar color → Jira project mapping) that previously lived in a sheet is stored as JSON in `PropertiesService`.

**Tech Stack:** Google Apps Script (server), HtmlService SPA with vanilla JS (client), Jira REST API v3, Google Calendar API (via `CalendarApp`), `PropertiesService` for persistence.

## Global Constraints

- Apps Script runtime — no Node.js APIs, no npm, no `fetch` (use `UrlFetchApp` server-side only)
- All HTML files included via `HtmlService.createHtmlOutputFromFile()` / `<?!= include('Stylesheet') ?>` pattern
- `PropertiesService.getUserProperties()` max value size ~9 KB — sufficient for allocation config
- Jira REST API base URL and API key stored per-user via `getUserProperties()`
- Calendar color numbers match `CalendarApp.EventColor` enum values (strings "1"–"11")
- No `ai()` formula — Jira task assignment per calendar event is a manual dropdown
- Fix pre-existing bug: `historyRows` referenced but undeclared in `sendTime`
- `clasp` project root is the repo root; `appsscript.json` lives at root

---

## File Map

| File | Role |
|---|---|
| `appsscript.json` | Manifest — webapp entrypoint, OAuth scopes |
| `Code.js` | All server-side functions (replaces existing `Code.js`) |
| `Index.html` | SPA shell — tab nav + `<?!= include() ?>` partials |
| `Stylesheet.html` | All CSS (included into Index) |
| `JavaScript.html` | All client-side JS (included into Index) |
| `TabAssignments.html` | Assignments tab markup |
| `TabTimecard.html` | Timecard tab markup (date range, event list, submit) |
| `TabCreateIssue.html` | Create Jira issue form (replaces `CreateNewJira.html`) |
| `TabWorklog.html` | Pull worklog tab |
| `TabConfig.html` | Config tab (API key, Jira URL, Allocation table) |

Existing files to retire (do not delete until end): `CreateNewJira.html`, old `Code.js`.

---

## Task 1: Project Scaffold + Manifest

**Files:**
- Create: `appsscript.json`
- Create: `Index.html`
- Create: `Stylesheet.html`
- Create: `JavaScript.html`
- Create: `TabAssignments.html`
- Create: `TabTimecard.html`
- Create: `TabCreateIssue.html`
- Create: `TabWorklog.html`
- Create: `TabConfig.html`

**Interfaces:**
- Produces: `doGet()` in `Code.js` returning `HtmlService` output; `include(filename)` helper

- [ ] **Step 1: Create `appsscript.json`**

```json
{
  "timeZone": "America/Chicago",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "webapp": {
    "executeAs": "USER_ACCESSING",
    "access": "MYSELF"
  },
  "oauthScopes": [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/script.storage"
  ]
}
```

- [ ] **Step 2: Create `Index.html` shell**

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <title>Jira Time Tracker</title>
  <?!= include('Stylesheet'); ?>
</head>
<body>
  <header><h1>Jira Time Tracker</h1></header>
  <nav>
    <button class="tab-btn active" data-tab="assignments">Assignments</button>
    <button class="tab-btn" data-tab="timecard">Timecard</button>
    <button class="tab-btn" data-tab="create-issue">Create Issue</button>
    <button class="tab-btn" data-tab="worklog">Worklog</button>
    <button class="tab-btn" data-tab="config">Config</button>
  </nav>
  <main>
    <section id="tab-assignments" class="tab-section active">
      <?!= include('TabAssignments'); ?>
    </section>
    <section id="tab-timecard" class="tab-section">
      <?!= include('TabTimecard'); ?>
    </section>
    <section id="tab-create-issue" class="tab-section">
      <?!= include('TabCreateIssue'); ?>
    </section>
    <section id="tab-worklog" class="tab-section">
      <?!= include('TabWorklog'); ?>
    </section>
    <section id="tab-config" class="tab-section">
      <?!= include('TabConfig'); ?>
    </section>
  </main>
  <?!= include('JavaScript'); ?>
</body>
</html>
```

- [ ] **Step 3: Create stub HTML files for each tab** (content filled in later tasks)

Each file should be a single comment for now:
- `TabAssignments.html` → `<!-- Assignments tab -->`
- `TabTimecard.html` → `<!-- Timecard tab -->`
- `TabCreateIssue.html` → `<!-- Create Issue tab -->`
- `TabWorklog.html` → `<!-- Worklog tab -->`
- `TabConfig.html` → `<!-- Config tab -->`

- [ ] **Step 4: Create `Stylesheet.html` with base styles**

```html
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Google Sans', Arial, sans-serif; font-size: 14px; color: #202124; background: #f8f9fa; }
  header { background: #1a73e8; color: white; padding: 12px 20px; }
  header h1 { font-size: 18px; font-weight: 500; }
  nav { display: flex; gap: 4px; padding: 8px 20px; background: white; border-bottom: 1px solid #dadce0; }
  .tab-btn { padding: 8px 16px; border: none; border-radius: 4px; background: transparent; cursor: pointer; font-size: 14px; color: #5f6368; }
  .tab-btn.active { background: #e8f0fe; color: #1a73e8; font-weight: 500; }
  .tab-section { display: none; padding: 20px; }
  .tab-section.active { display: block; }
  table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
  th { background: #f1f3f4; padding: 10px 12px; text-align: left; font-weight: 500; font-size: 12px; color: #5f6368; text-transform: uppercase; letter-spacing: 0.5px; }
  td { padding: 10px 12px; border-top: 1px solid #f1f3f4; vertical-align: middle; }
  tr:hover td { background: #f8f9fa; }
  .btn { padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500; }
  .btn-primary { background: #1a73e8; color: white; }
  .btn-primary:hover { background: #1557b0; }
  .btn-secondary { background: white; color: #1a73e8; border: 1px solid #dadce0; }
  .btn-secondary:hover { background: #e8f0fe; }
  .btn:disabled { opacity: 0.5; cursor: not-allowed; }
  .toolbar { display: flex; gap: 8px; margin-bottom: 16px; align-items: center; }
  input[type="text"], input[type="date"], input[type="number"], select, textarea {
    border: 1px solid #dadce0; border-radius: 4px; padding: 8px 10px; font-size: 14px; width: 100%;
  }
  input[type="number"] { width: 80px; }
  .status-msg { margin-top: 12px; padding: 10px 14px; border-radius: 4px; font-size: 13px; display: none; }
  .status-msg.success { background: #e6f4ea; color: #137333; display: block; }
  .status-msg.error { background: #fce8e6; color: #c5221f; display: block; }
  .status-msg.info { background: #e8f0fe; color: #1a73e8; display: block; }
  label { font-size: 13px; font-weight: 500; color: #5f6368; display: block; margin-bottom: 4px; }
  .form-group { margin-bottom: 16px; }
  .section-title { font-size: 16px; font-weight: 500; margin-bottom: 16px; color: #202124; }
  a { color: #1a73e8; text-decoration: none; }
  a:hover { text-decoration: underline; }
</style>
```

- [ ] **Step 5: Create `JavaScript.html` with tab switching**

```html
<script>
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-section').forEach(s => s.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
    });
  });

  function showStatus(elementId, message, type) {
    const el = document.getElementById(elementId);
    if (!el) return;
    el.textContent = message;
    el.className = 'status-msg ' + type;
  }

  function setLoading(btnId, loading, label) {
    const btn = document.getElementById(btnId);
    if (!btn) return;
    btn.disabled = loading;
    btn.textContent = loading ? 'Loading…' : label;
  }
</script>
```

- [ ] **Step 6: Add `doGet()` and `include()` to `Code.js`** (top of file, before all else)

```javascript
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Jira Time Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
```

- [ ] **Step 7: Deploy as web app and verify it loads**

In Apps Script editor → Deploy → New deployment → Web app → Execute as: Me → Who has access: Only myself → Deploy. Open the URL and confirm the tab nav renders.

- [ ] **Step 8: Commit**

```bash
git add appsscript.json Index.html Stylesheet.html JavaScript.html TabAssignments.html TabTimecard.html TabCreateIssue.html TabWorklog.html TabConfig.html Code.js
git commit -m "feat: scaffold standalone web app with tab navigation"
```

---

## Task 2: Server — Config Functions

**Files:**
- Modify: `Code.js` — add config functions, keep helpers

**Interfaces:**
- Produces:
  - `getConfig()` → `{ jiraUrl: string, hasApiKey: boolean }`
  - `saveJiraUrl(url: string)` → `void`
  - `saveApiKey(key: string)` → `void`
  - `getAllocation()` → `Array<{ colorNum: string, projectKey: string, colorLabel: string, projectName: string, hoursPerWeek: number, colorEnumName: string }>`
  - `saveAllocation(rows: Array<{ colorNum: string, projectKey: string, colorLabel: string, projectName: string, hoursPerWeek: number, colorEnumName: string }>)` → `void`
  - `getUserEmail()` → `string`
  - `getAuthHeader_()` → `string` (private helper, already exists)
  - `getUserProperties()` → PropertiesService store (private helper, already exists)

- [ ] **Step 1: Add config server functions to `Code.js`**

```javascript
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

function getConfig() {
  const props = getUserProperties();
  return {
    jiraUrl: props.getProperty('JIRA_BASE_URL') || '',
    hasApiKey: !!props.getProperty('JIRA_API_KEY')
  };
}

function saveJiraUrl(url) {
  getUserProperties().setProperty('JIRA_BASE_URL', url.trim().replace(/\/$/, ''));
}

function saveApiKey(key) {
  getUserProperties().setProperty('JIRA_API_KEY', key.trim());
}

function getAllocation() {
  const raw = getUserProperties().getProperty('ALLOCATION');
  return raw ? JSON.parse(raw) : [];
}

function saveAllocation(rows) {
  getUserProperties().setProperty('ALLOCATION', JSON.stringify(rows));
}
```

- [ ] **Step 2: Verify helpers are present in `Code.js`**

Confirm these already exist (they do — ported from original):
```javascript
function getUserProperties() {
  return PropertiesService.getUserProperties();
}

function getAuthHeader_() {
  const email = Session.getActiveUser().getEmail();
  const apiKey = getUserProperties().getProperty('JIRA_API_KEY');
  return 'Basic ' + Utilities.base64Encode(`${email}:${apiKey}`);
}
```

- [ ] **Step 3: Commit**

```bash
git add Code.js
git commit -m "feat: add server-side config functions (getConfig, saveJiraUrl, saveApiKey, getAllocation, saveAllocation)"
```

---

## Task 3: Config Tab UI

**Files:**
- Modify: `TabConfig.html`
- Modify: `JavaScript.html` — add config tab JS

**Interfaces:**
- Consumes: `getConfig()`, `saveJiraUrl()`, `saveApiKey()`, `getAllocation()`, `saveAllocation()`

- [ ] **Step 1: Write `TabConfig.html`**

```html
<h2 class="section-title">Configuration</h2>

<div style="max-width:600px;">
  <div class="form-group">
    <label>Jira Base URL</label>
    <input type="text" id="cfg-jira-url" placeholder="https://your-company.atlassian.net">
  </div>
  <div class="form-group">
    <label>Jira API Key</label>
    <input type="text" id="cfg-api-key" placeholder="Paste your API token here">
    <div style="font-size:12px;color:#5f6368;margin-top:4px;" id="cfg-api-key-status"></div>
  </div>
  <div class="toolbar">
    <button class="btn btn-primary" id="cfg-save-btn" onclick="saveConfig()">Save</button>
  </div>
  <div id="cfg-status" class="status-msg"></div>
</div>

<hr style="margin:24px 0;border:none;border-top:1px solid #dadce0;">

<h3 class="section-title">Allocation — Calendar Color → Jira Project Mapping</h3>
<p style="font-size:13px;color:#5f6368;margin-bottom:16px;">
  Maps Google Calendar event colors to Jira project keys. Only events with a matching color are imported into the timecard.
</p>

<div class="toolbar">
  <button class="btn btn-secondary" onclick="addAllocationRow()">+ Add Row</button>
  <button class="btn btn-primary" onclick="saveAllocationTable()">Save Allocation</button>
</div>

<table id="allocation-table">
  <thead>
    <tr>
      <th>Calendar Color #</th>
      <th>Jira Project Key</th>
      <th>Color Label</th>
      <th>Project Name</th>
      <th>Hours/Wk</th>
      <th>Color Enum Name</th>
      <th></th>
    </tr>
  </thead>
  <tbody id="allocation-body"></tbody>
</table>
<div id="allocation-status" class="status-msg"></div>
```

- [ ] **Step 2: Add config tab JS to `JavaScript.html`**

```javascript
// ── Config Tab ──────────────────────────────────────────────
function loadConfig() {
  google.script.run
    .withSuccessHandler(cfg => {
      document.getElementById('cfg-jira-url').value = cfg.jiraUrl;
      document.getElementById('cfg-api-key-status').textContent = cfg.hasApiKey ? '✓ API key saved' : 'No API key saved';
    })
    .getConfig();
  google.script.run
    .withSuccessHandler(renderAllocationTable)
    .getAllocation();
}

function saveConfig() {
  const url = document.getElementById('cfg-jira-url').value.trim();
  const key = document.getElementById('cfg-api-key').value.trim();
  setLoading('cfg-save-btn', true, 'Save');
  let pending = 0;
  function done() { if (--pending === 0) { setLoading('cfg-save-btn', false, 'Save'); showStatus('cfg-status', 'Saved!', 'success'); } }
  if (url) { pending++; google.script.run.withSuccessHandler(done).saveJiraUrl(url); }
  if (key) { pending++; google.script.run.withSuccessHandler(done).saveApiKey(key); }
  if (!pending) showStatus('cfg-status', 'Nothing to save.', 'info');
}

function renderAllocationTable(rows) {
  const tbody = document.getElementById('allocation-body');
  tbody.innerHTML = '';
  (rows || []).forEach(row => addAllocationRow(row));
}

function addAllocationRow(data) {
  const tbody = document.getElementById('allocation-body');
  const tr = document.createElement('tr');
  const fields = ['colorNum','projectKey','colorLabel','projectName','hoursPerWeek','colorEnumName'];
  const placeholders = ['1–11','CEOT','Orange','CFA Project','20','ORANGE'];
  tr.innerHTML = fields.map((f,i) => `<td><input type="text" style="width:100%" name="${f}" value="${(data && data[f]) || ''}" placeholder="${placeholders[i]}"></td>`).join('') +
    `<td><button class="btn" style="padding:4px 8px;" onclick="this.closest('tr').remove()">✕</button></td>`;
  tbody.appendChild(tr);
}

function saveAllocationTable() {
  const rows = Array.from(document.querySelectorAll('#allocation-body tr')).map(tr => {
    const inputs = tr.querySelectorAll('input');
    return {
      colorNum: inputs[0].value.trim(),
      projectKey: inputs[1].value.trim(),
      colorLabel: inputs[2].value.trim(),
      projectName: inputs[3].value.trim(),
      hoursPerWeek: parseFloat(inputs[4].value) || 0,
      colorEnumName: inputs[5].value.trim().toUpperCase()
    };
  }).filter(r => r.projectKey);
  google.script.run
    .withSuccessHandler(() => showStatus('allocation-status', 'Allocation saved.', 'success'))
    .withFailureHandler(e => showStatus('allocation-status', 'Error: ' + e.message, 'error'))
    .saveAllocation(rows);
}

// Load config when tab is clicked
document.querySelector('[data-tab="config"]').addEventListener('click', loadConfig);
```

- [ ] **Step 3: Reload the web app, click Config tab, verify the form loads and saving works**

Open the web app URL, go to Config tab, enter a dummy URL and API key, click Save, reload — confirm values persist.

- [ ] **Step 4: Commit**

```bash
git add TabConfig.html JavaScript.html
git commit -m "feat: config tab UI — Jira URL, API key, and allocation table"
```

---

## Task 4: Server — Assignments (Port `retrieveJiraIssues`)

**Files:**
- Modify: `Code.js` — add `getJiraIssues()`, `getWorklogTotals_()`, `getIssueWorklogTotal_()`, `parseTimeSpentHours_()`

**Interfaces:**
- Produces:
  - `getJiraIssues()` → `Array<{ dropdownValue: string, key: string, link: string, projectKey: string, name: string, status: string, project: string, timeLogged: number }>`

Note: `getWorklogTotals_`, `getIssueWorklogTotal_`, and `parseTimeSpentHours_` already exist in the original `Code.js` — keep them unchanged.

- [ ] **Step 1: Add `getJiraIssues()` to `Code.js`**

```javascript
function getJiraIssues() {
  const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
  const USER_EMAIL = Session.getActiveUser().getEmail();
  const authHeader = getAuthHeader_();
  const BASE_ENDPOINT = `${JIRA_URL}/rest/api/3/search/jql?jql=(assignee=currentUser()+OR+watcher=currentUser())+AND+issuetype+IN+(Story,Task,Sub-Task)+AND+(status!=Done+OR+(status=Done+AND+updated%3E=-7d))+ORDER+BY+key+ASC&fields=key,summary,status,project&maxResults=100`;
  const options = { headers: { Authorization: authHeader }, method: 'get', muteHttpExceptions: true };

  let allIssues = [];
  let nextPageToken = null;
  do {
    const endpoint = nextPageToken ? `${BASE_ENDPOINT}&nextPageToken=${nextPageToken}` : BASE_ENDPOINT;
    const data = JSON.parse(UrlFetchApp.fetch(endpoint, options).getContentText());
    if (data.issues) allIssues = allIssues.concat(data.issues);
    nextPageToken = data.nextPageToken || null;
  } while (nextPageToken);

  const filtered = allIssues.filter(issue => {
    const name = issue.fields.project.name;
    return !name.includes('Archive') && !name.includes('Managed Services Internal');
  });

  const worklogTotals = getWorklogTotals_(filtered.map(i => i.key), authHeader, JIRA_URL, USER_EMAIL);

  return filtered.map((issue, idx) => ({
    dropdownValue: `${issue.key} (${issue.fields.summary})`,
    key: issue.key,
    link: `${JIRA_URL}/browse/${issue.key}`,
    projectKey: issue.key.split('-')[0],
    name: issue.fields.summary,
    status: issue.fields.status.name,
    project: issue.fields.project.name,
    timeLogged: Math.round((worklogTotals[idx] / 3600) * 100) / 100
  }));
}
```

- [ ] **Step 2: Confirm `getWorklogTotals_` and `parseTimeSpentHours_` are present** — copy unchanged from original `Code.js` if not already there.

- [ ] **Step 3: Commit**

```bash
git add Code.js
git commit -m "feat: server getJiraIssues — port retrieveJiraIssues for web app"
```

---

## Task 5: Assignments Tab UI

**Files:**
- Modify: `TabAssignments.html`
- Modify: `JavaScript.html` — add assignments JS

**Interfaces:**
- Consumes: `getJiraIssues()` → array from Task 4
- Produces: `window.loadedIssues` — in-memory array used by Task 6 (scheduler)

- [ ] **Step 1: Write `TabAssignments.html`**

```html
<h2 class="section-title">Assignments</h2>
<div class="toolbar">
  <button class="btn btn-primary" id="assignments-load-btn" onclick="loadAssignments()">Load Assignments</button>
  <span id="assignments-count" style="color:#5f6368;font-size:13px;"></span>
</div>
<div id="assignments-status" class="status-msg"></div>

<table id="assignments-table" style="display:none;">
  <thead>
    <tr>
      <th>Key</th>
      <th>Name</th>
      <th>Project</th>
      <th>Status</th>
      <th>Time Logged (h)</th>
      <th>Schedule (h)</th>
    </tr>
  </thead>
  <tbody id="assignments-body"></tbody>
</table>

<div id="schedule-toolbar" style="display:none;margin-top:16px;" class="toolbar">
  <button class="btn btn-secondary" onclick="scheduleEvents()">⏱ Schedule Calendar Events</button>
</div>
<div id="schedule-status" class="status-msg"></div>
```

- [ ] **Step 2: Add assignments JS to `JavaScript.html`**

```javascript
// ── Assignments Tab ──────────────────────────────────────────
window.loadedIssues = [];

function loadAssignments() {
  setLoading('assignments-load-btn', true, 'Load Assignments');
  showStatus('assignments-status', 'Fetching from Jira…', 'info');
  google.script.run
    .withSuccessHandler(issues => {
      window.loadedIssues = issues;
      renderAssignments(issues);
      setLoading('assignments-load-btn', false, 'Load Assignments');
      document.getElementById('assignments-count').textContent = `${issues.length} issues`;
      showStatus('assignments-status', '', '');
    })
    .withFailureHandler(e => {
      setLoading('assignments-load-btn', false, 'Load Assignments');
      showStatus('assignments-status', 'Error: ' + e.message, 'error');
    })
    .getJiraIssues();
}

function renderAssignments(issues) {
  const tbody = document.getElementById('assignments-body');
  tbody.innerHTML = '';
  issues.forEach((issue, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><a href="${issue.link}" target="_blank">${issue.key}</a></td>
      <td>${issue.name}</td>
      <td>${issue.project}</td>
      <td>${issue.status}</td>
      <td>${issue.timeLogged}</td>
      <td><input type="number" min="0" step="0.25" style="width:80px;" data-idx="${idx}" class="schedule-input" placeholder="0"></td>`;
    tbody.appendChild(tr);
  });
  document.getElementById('assignments-table').style.display = '';
  document.getElementById('schedule-toolbar').style.display = '';
}
```

- [ ] **Step 3: Reload web app, click Assignments tab, click Load Assignments, verify table renders**

- [ ] **Step 4: Commit**

```bash
git add TabAssignments.html JavaScript.html
git commit -m "feat: assignments tab — load and display Jira issues with schedule input"
```

---

## Task 6: Server — Schedule Calendar Events

**Files:**
- Modify: `Code.js` — add `scheduleCalendarEvents(toSchedule)`

**Interfaces:**
- Consumes:
  ```javascript
  toSchedule: Array<{ key: string, jiraProject: string, dropdownValue: string, hours: number }>
  allocation: Array<{ colorNum, projectKey, colorLabel, projectName, hoursPerWeek, colorEnumName }> // from getAllocation()
  ```
- Produces: `scheduleCalendarEvents(toSchedule)` → `{ created: number, startTime: string }`

- [ ] **Step 1: Add `scheduleCalendarEvents` to `Code.js`**

```javascript
function scheduleCalendarEvents(toSchedule) {
  const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
  const allocation = getAllocation();

  const COLOR_ENUM_MAP = {
    'PALE_BLUE': CalendarApp.EventColor.PALE_BLUE,
    'PALE_GREEN': CalendarApp.EventColor.PALE_GREEN,
    'MAUVE': CalendarApp.EventColor.MAUVE,
    'PALE_RED': CalendarApp.EventColor.PALE_RED,
    'YELLOW': CalendarApp.EventColor.YELLOW,
    'ORANGE': CalendarApp.EventColor.ORANGE,
    'CYAN': CalendarApp.EventColor.CYAN,
    'GRAY': CalendarApp.EventColor.GRAY,
    'GREY': CalendarApp.EventColor.GRAY,
    'BLUE': CalendarApp.EventColor.BLUE,
    'GREEN': CalendarApp.EventColor.GREEN,
    'RED': CalendarApp.EventColor.RED
  };

  const projectColorMap = Object.fromEntries(
    allocation
      .filter(r => r.projectKey)
      .map(r => [r.projectKey, COLOR_ENUM_MAP[r.colorEnumName] || null])
  );

  const now = new Date();
  const startTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours() + 1, 0, 0, 0);
  const calendar = CalendarApp.getDefaultCalendar();
  let cursor = new Date(startTime);

  toSchedule.forEach(entry => {
    const durationMs = entry.hours * 60 * 60 * 1000;
    const endTime = new Date(cursor.getTime() + durationMs);
    const event = calendar.createEvent(
      entry.jiraProject,
      cursor,
      endTime,
      { description: `${entry.dropdownValue}\n${JIRA_URL}/browse/${entry.key}` }
    );
    const color = projectColorMap[entry.jiraProject];
    if (color) event.setColor(color);
    cursor = endTime;
  });

  return { created: toSchedule.length, startTime: startTime.toLocaleTimeString() };
}
```

- [ ] **Step 2: Add `scheduleEvents()` to `JavaScript.html`**

```javascript
function scheduleEvents() {
  const toSchedule = Array.from(document.querySelectorAll('.schedule-input'))
    .map(input => {
      const hours = parseFloat(input.value);
      if (!hours || hours <= 0) return null;
      const issue = window.loadedIssues[parseInt(input.dataset.idx)];
      return { key: issue.key, jiraProject: issue.projectKey, dropdownValue: issue.dropdownValue, hours };
    })
    .filter(Boolean);

  if (!toSchedule.length) {
    showStatus('schedule-status', 'Enter hours in the Schedule column first.', 'info');
    return;
  }
  showStatus('schedule-status', 'Creating calendar events…', 'info');
  google.script.run
    .withSuccessHandler(result => {
      showStatus('schedule-status', `Created ${result.created} event(s) starting at ${result.startTime}.`, 'success');
      document.querySelectorAll('.schedule-input').forEach(i => i.value = '');
    })
    .withFailureHandler(e => showStatus('schedule-status', 'Error: ' + e.message, 'error'))
    .scheduleCalendarEvents(toSchedule);
}
```

- [ ] **Step 3: Test end-to-end: load assignments, enter hours for one issue, click Schedule, verify event appears in Google Calendar**

- [ ] **Step 4: Commit**

```bash
git add Code.js JavaScript.html
git commit -m "feat: schedule calendar events from assignments tab"
```

---

## Task 7: Server — Import Calendar Events

**Files:**
- Modify: `Code.js` — add `importCalendarEvents(startDateStr, endDateStr)`

**Interfaces:**
- Produces:
  ```javascript
  importCalendarEvents(startDateStr: string, endDateStr: string)
  → Array<{
      title: string, date: string, start: string, end: string,
      description: string, status: string, projectKey: string,
      issueKey: string, duration: number
    }>
  ```
  Note: `issueKey` is empty string — user assigns it in the UI (replaces the `ai()` formula).

- [ ] **Step 1: Add `importCalendarEvents` to `Code.js`**

```javascript
function importCalendarEvents(startDateStr, endDateStr) {
  const allocation = getAllocation();
  const colorMap = Object.fromEntries(
    allocation.filter(r => r.projectKey && r.colorLabel).map(r => [r.colorNum, r.projectKey])
  );
  const validProjectKeys = new Set(
    allocation.filter(r => r.projectKey && r.colorLabel).map(r => r.projectKey)
  );

  const CALENDAR_ID = Session.getActiveUser().getEmail();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const tz = Session.getScriptTimeZone();

  const startDate = new Date(startDateStr + 'T00:00:00');
  const endDate = new Date(endDateStr + 'T23:59:59');

  const events = calendar.getEvents(startDate, endDate).map(event => {
    const colorNum = event.getColor();
    const projectKey = colorMap[colorNum] || '';
    if (!validProjectKeys.has(projectKey)) return null;
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    const durationHours = (endTime - startTime) / 3600000;
    const description = event.getDescription();
    const match = description ? description.match(/^[^\n_]+/) : null;
    return {
      title: event.getTitle(),
      date: Utilities.formatDate(startTime, tz, 'yyyy-MM-dd'),
      start: Utilities.formatDate(startTime, tz, 'hh:mm a'),
      end: Utilities.formatDate(endTime, tz, 'hh:mm a'),
      description: match ? match[0].trim() : '',
      status: event.getMyStatus(),
      projectKey,
      issueKey: '',
      duration: Math.round(durationHours * 4) / 4
    };
  }).filter(Boolean);

  return events;
}
```

- [ ] **Step 2: Commit**

```bash
git add Code.js
git commit -m "feat: server importCalendarEvents — port calendar import without ai() formula"
```

---

## Task 8: Server — Send Time to Jira (Port `sendTime`, fix bug)

**Files:**
- Modify: `Code.js` — add `sendTimeEntries(entries)`

**Interfaces:**
- Consumes:
  ```javascript
  entries: Array<{ date: string, issueKey: string, startTime: string, durationHours: number }>
  ```
- Produces: `sendTimeEntries(entries)` → `{ succeeded: number, failed: number, errors: string[] }`

Note: Fixes the pre-existing `historyRows` bug — the history append was commented out but `historyRows.push` was not. The new function returns results instead of writing to a sheet.

- [ ] **Step 1: Add `sendTimeEntries` to `Code.js`**

```javascript
function sendTimeEntries(entries) {
  const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
  const authHeader = getAuthHeader_();
  const UTC_FORMAT = "yyyy-MM-dd'T'HH:mm:ss'.000+0000'";
  let succeeded = 0, failed = 0;
  const errors = [];

  entries.forEach(entry => {
    if (!entry.issueKey) return;
    const combined = new Date(`${entry.date}T${entry.startTime}`);
    const utcString = Utilities.formatDate(combined, 'Etc/GMT', UTC_FORMAT);
    const durationMinutes = Math.round(entry.durationHours * 60);
    const options = {
      method: 'post',
      headers: { Authorization: authHeader, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ started: utcString, timeSpent: `${durationMinutes}m` }),
      muteHttpExceptions: true
    };
    try {
      const resp = UrlFetchApp.fetch(`${JIRA_URL}/rest/api/3/issue/${entry.issueKey}/worklog`, options);
      if (resp.getResponseCode() === 201) {
        succeeded++;
      } else {
        failed++;
        errors.push(`${entry.issueKey}: HTTP ${resp.getResponseCode()}`);
      }
    } catch (e) {
      failed++;
      errors.push(`${entry.issueKey}: ${e.message}`);
    }
  });

  return { succeeded, failed, errors };
}
```

- [ ] **Step 2: Commit**

```bash
git add Code.js
git commit -m "feat: server sendTimeEntries — port sendTime, fix historyRows bug, return results"
```

---

## Task 9: Timecard Tab UI

**Files:**
- Modify: `TabTimecard.html`
- Modify: `JavaScript.html` — add timecard JS

**Interfaces:**
- Consumes: `importCalendarEvents(startDateStr, endDateStr)`, `getJiraIssues()` (for dropdown options), `sendTimeEntries(entries)`

- [ ] **Step 1: Write `TabTimecard.html`**

```html
<h2 class="section-title">Timecard</h2>

<div class="toolbar" style="flex-wrap:wrap;gap:12px;">
  <div>
    <label>Start Date</label>
    <input type="date" id="tc-start-date" style="width:160px;">
  </div>
  <div>
    <label>End Date</label>
    <input type="date" id="tc-end-date" style="width:160px;">
  </div>
  <div style="align-self:flex-end;">
    <button class="btn btn-secondary" id="tc-load-btn" onclick="loadTimecardEvents()">Import Events</button>
  </div>
</div>
<div id="tc-status" class="status-msg"></div>

<table id="tc-table" style="display:none;margin-top:16px;">
  <thead>
    <tr>
      <th>Date</th>
      <th>Event</th>
      <th>Start</th>
      <th>End</th>
      <th>Project</th>
      <th>Jira Issue</th>
      <th>Duration (h)</th>
    </tr>
  </thead>
  <tbody id="tc-body"></tbody>
</table>

<div id="tc-submit-toolbar" style="display:none;margin-top:16px;" class="toolbar">
  <button class="btn btn-primary" id="tc-submit-btn" onclick="submitTimecard()">📥 Send to Jira</button>
</div>
<div id="tc-submit-status" class="status-msg"></div>
```

- [ ] **Step 2: Add timecard JS to `JavaScript.html`**

```javascript
// ── Timecard Tab ─────────────────────────────────────────────
(function initTimecardDates() {
  const today = new Date();
  const day = today.getDay();
  const monday = new Date(today);
  monday.setDate(today.getDate() - (day === 0 ? 6 : day - 1));
  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4);
  const fmt = d => d.toISOString().slice(0, 10);
  document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('tc-start-date').value = fmt(monday);
    document.getElementById('tc-end-date').value = fmt(friday);
  });
})();

window.timecardEvents = [];
window.issueOptions = [];

function loadTimecardEvents() {
  const start = document.getElementById('tc-start-date').value;
  const end = document.getElementById('tc-end-date').value;
  if (!start || !end) { showStatus('tc-status', 'Select a date range.', 'info'); return; }
  setLoading('tc-load-btn', true, 'Import Events');
  showStatus('tc-status', 'Importing calendar events…', 'info');

  // Load issues for dropdown in parallel
  google.script.run.withSuccessHandler(issues => { window.issueOptions = issues; }).getJiraIssues();

  google.script.run
    .withSuccessHandler(events => {
      window.timecardEvents = events;
      renderTimecardTable(events);
      setLoading('tc-load-btn', false, 'Import Events');
      showStatus('tc-status', `${events.length} event(s) imported.`, 'success');
    })
    .withFailureHandler(e => {
      setLoading('tc-load-btn', false, 'Import Events');
      showStatus('tc-status', 'Error: ' + e.message, 'error');
    })
    .importCalendarEvents(start, end);
}

function renderTimecardTable(events) {
  const tbody = document.getElementById('tc-body');
  tbody.innerHTML = '';
  events.forEach((ev, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${ev.date}</td>
      <td title="${ev.description}">${ev.title}</td>
      <td>${ev.start}</td>
      <td>${ev.end}</td>
      <td>${ev.projectKey}</td>
      <td>
        <select data-idx="${idx}" class="tc-issue-select" style="min-width:200px;">
          <option value="">— select —</option>
        </select>
      </td>
      <td>${ev.duration}</td>`;
    tbody.appendChild(tr);
  });
  document.getElementById('tc-table').style.display = '';
  document.getElementById('tc-submit-toolbar').style.display = '';

  // Populate dropdowns once issues are loaded (poll briefly)
  let attempts = 0;
  const populate = setInterval(() => {
    if (window.issueOptions.length || ++attempts > 20) {
      clearInterval(populate);
      document.querySelectorAll('.tc-issue-select').forEach(sel => {
        const ev = window.timecardEvents[parseInt(sel.dataset.idx)];
        window.issueOptions.forEach(issue => {
          const opt = document.createElement('option');
          opt.value = issue.key;
          opt.textContent = issue.dropdownValue;
          if (issue.projectKey === ev.projectKey) sel.appendChild(opt);
        });
      });
    }
  }, 300);
}

function submitTimecard() {
  const entries = Array.from(document.querySelectorAll('.tc-issue-select')).map(sel => {
    const ev = window.timecardEvents[parseInt(sel.dataset.idx)];
    return { date: ev.date, issueKey: sel.value, startTime: ev.start, durationHours: ev.duration };
  }).filter(e => e.issueKey);

  if (!entries.length) { showStatus('tc-submit-status', 'No issues selected.', 'info'); return; }
  setLoading('tc-submit-btn', true, '📥 Send to Jira');
  google.script.run
    .withSuccessHandler(result => {
      setLoading('tc-submit-btn', false, '📥 Send to Jira');
      const msg = `Sent ${result.succeeded} worklog(s).` + (result.failed ? ` ${result.failed} failed: ${result.errors.join(', ')}` : '');
      showStatus('tc-submit-status', msg, result.failed ? 'error' : 'success');
    })
    .withFailureHandler(e => {
      setLoading('tc-submit-btn', false, '📥 Send to Jira');
      showStatus('tc-submit-status', 'Error: ' + e.message, 'error');
    })
    .sendTimeEntries(entries);
}
```

- [ ] **Step 3: Test: import events for the current week, verify events appear, select a Jira issue, submit — confirm worklog appears in Jira**

- [ ] **Step 4: Commit**

```bash
git add TabTimecard.html JavaScript.html
git commit -m "feat: timecard tab — import calendar events, assign Jira issues, submit worklogs"
```

---

## Task 10: Create Jira Issue Tab

**Files:**
- Modify: `TabCreateIssue.html` — port `CreateNewJira.html`
- Modify: `Code.js` — port `makeJira`, `getDropdownValues` → `getProjectKeys`
- Modify: `JavaScript.html` — add create issue JS

**Interfaces:**
- Consumes: `getAllocation()` for project key dropdown (replaces `getDropdownValues` which read from a sheet)
- Produces:
  - `getProjectKeys()` → `string[]`
  - `makeJira(formData)` → `string` (status message)

Note: `getProjectKeys` derives project keys from `getAllocation()` instead of the Allocation sheet.

- [ ] **Step 1: Add `getProjectKeys` to `Code.js`**

```javascript
function getProjectKeys() {
  return getAllocation().map(r => r.projectKey).filter(Boolean);
}
```

- [ ] **Step 2: Keep `makeJira` in `Code.js`** — it has no sheet dependencies; copy it unchanged from the original.

- [ ] **Step 3: Write `TabCreateIssue.html`**

```html
<h2 class="section-title">Create Jira Issue</h2>
<div style="max-width:500px;">
  <div class="form-group">
    <label>Project</label>
    <select id="ci-project"></select>
  </div>
  <div class="form-group">
    <label>Technology</label>
    <select id="ci-technology">
      <option value=""></option>
      <option>AIMS</option><option>ArcGIS</option><option>Atlan</option>
      <option>AWS Glue</option><option>BI Support</option><option>dbt</option>
      <option>DQLabs</option><option>HeyWM</option><option>HVR</option>
      <option>Informatica</option><option>InfoSphere</option><option>Matillion</option>
      <option>Newton Insights</option><option>Power Apps</option><option>Power Automate</option>
      <option>Power BI</option><option>Qlik</option><option>Sigma</option>
      <option>Snowflake</option><option>Spotfire</option><option>Tableau</option>
    </select>
  </div>
  <div class="form-group">
    <label>MS ElasticOps Work Type</label>
    <select id="ci-worktype">
      <option value=""></option>
      <option>AdvOps</option>
      <option>Core</option>
    </select>
  </div>
  <div class="form-group">
    <label>Issue Type</label>
    <select id="ci-issuetype">
      <option value="Task">Task</option>
      <option value="Support">Incident</option>
    </select>
  </div>
  <div class="form-group">
    <label>Priority</label>
    <select id="ci-priority">
      <option value="10002">P-3 Low</option>
      <option value="10001" selected>P-2 Medium</option>
      <option value="10000">P-1 High</option>
    </select>
  </div>
  <div class="form-group">
    <label>Summary</label>
    <input type="text" id="ci-summary">
  </div>
  <div class="form-group">
    <label>Description</label>
    <textarea id="ci-description" rows="4"></textarea>
  </div>
  <button class="btn btn-primary" id="ci-submit-btn" onclick="submitNewIssue()">Create Issue</button>
  <div id="ci-status" class="status-msg"></div>
</div>
```

- [ ] **Step 4: Add create issue JS to `JavaScript.html`**

```javascript
// ── Create Issue Tab ─────────────────────────────────────────
document.querySelector('[data-tab="create-issue"]').addEventListener('click', () => {
  google.script.run.withSuccessHandler(keys => {
    const sel = document.getElementById('ci-project');
    sel.innerHTML = keys.map(k => `<option value="${k}">${k}</option>`).join('');
  }).getProjectKeys();
});

function submitNewIssue() {
  const formData = {
    input1: document.getElementById('ci-project').value,
    input2: document.getElementById('ci-technology').value,
    input3: document.getElementById('ci-worktype').value,
    input4: document.getElementById('ci-summary').value,
    input6: document.getElementById('ci-issuetype').value,
    input7: document.getElementById('ci-priority').value,
    notes: document.getElementById('ci-description').value
  };
  setLoading('ci-submit-btn', true, 'Create Issue');
  google.script.run
    .withSuccessHandler(msg => {
      setLoading('ci-submit-btn', false, 'Create Issue');
      showStatus('ci-status', msg, msg.startsWith('Error') ? 'error' : 'success');
    })
    .withFailureHandler(e => {
      setLoading('ci-submit-btn', false, 'Create Issue');
      showStatus('ci-status', 'Error: ' + e.message, 'error');
    })
    .makeJira(formData);
}
```

- [ ] **Step 5: Test: create a new issue, confirm it appears in Jira**

- [ ] **Step 6: Commit**

```bash
git add TabCreateIssue.html Code.js JavaScript.html
git commit -m "feat: create Jira issue tab — port CreateNewJira dialog into web app tab"
```

---

## Task 11: Worklog Tab

**Files:**
- Modify: `TabWorklog.html`
- Modify: `Code.js` — port `pullWorklog` → `getWorklogs()`
- Modify: `JavaScript.html` — add worklog JS

**Interfaces:**
- Produces:
  - `getWorklogs()` → `Array<{ projectKey, issueKey, summary, timeSpent, started, hours, month }>`

- [ ] **Step 1: Add `getWorklogs` to `Code.js`**

```javascript
function getWorklogs() {
  const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
  const authHeader = getAuthHeader_();
  const userEmail = Session.getActiveUser().getEmail();
  const currentYear = new Date().getFullYear();
  const fromDate = `${currentYear}-01-01`;
  const fromTimestamp = new Date(currentYear, 0, 1).getTime();
  const JQL = encodeURIComponent(`worklogAuthor=currentUser() AND worklogDate >= ${fromDate}`);
  const BASE_ENDPOINT = `${JIRA_URL}/rest/api/3/search/jql?fields=key,summary,worklog,project&jql=${JQL}&maxResults=100`;
  const fetchOpts = { headers: { Authorization: authHeader }, method: 'get', muteHttpExceptions: true };

  let allIssues = [];
  let nextPageToken = null;
  do {
    const endpoint = nextPageToken ? `${BASE_ENDPOINT}&nextPageToken=${nextPageToken}` : BASE_ENDPOINT;
    const data = JSON.parse(UrlFetchApp.fetch(endpoint, fetchOpts).getContentText());
    if (data.issues) allIssues = allIssues.concat(data.issues);
    nextPageToken = data.nextPageToken || null;
  } while (nextPageToken);

  const extraFetches = [];
  allIssues.forEach((issue, idx) => {
    const wl = issue.fields.worklog;
    if (!wl) return;
    for (let s = wl.worklogs ? wl.worklogs.length : 0; s < (wl.total || 0); s += 100) {
      extraFetches.push({ issueIdx: idx, startAt: s });
    }
  });
  if (extraFetches.length > 0) {
    UrlFetchApp.fetchAll(extraFetches.map(f => ({
      url: `${JIRA_URL}/rest/api/3/issue/${allIssues[f.issueIdx].key}/worklog?startAt=${f.startAt}&maxResults=100`,
      headers: { Authorization: authHeader }, method: 'get', muteHttpExceptions: true
    }))).forEach((resp, i) => {
      const data = JSON.parse(resp.getContentText());
      if (!data.worklogs) return;
      const wl = allIssues[extraFetches[i].issueIdx].fields.worklog;
      wl.worklogs = (wl.worklogs || []).concat(data.worklogs);
    });
  }

  const rows = [];
  allIssues.forEach(issue => {
    const projectKey = issue.fields.project.key;
    const issueKey = issue.key;
    const summary = issue.fields.summary;
    (issue.fields.worklog && issue.fields.worklog.worklogs || []).forEach(log => {
      if (!log.author || log.author.emailAddress !== userEmail) return;
      const startedDate = new Date(log.started || '');
      if (!log.started || startedDate.getTime() < fromTimestamp) return;
      rows.push({
        projectKey, issueKey, summary,
        timeSpent: log.timeSpent || '',
        started: log.started,
        hours: parseTimeSpentHours_(log.timeSpent || ''),
        month: `${startedDate.getFullYear()}-${String(startedDate.getMonth() + 1).padStart(2, '0')}`
      });
    });
  });
  return rows;
}
```

- [ ] **Step 2: Write `TabWorklog.html`**

```html
<h2 class="section-title">Worklog (YTD)</h2>
<div class="toolbar">
  <button class="btn btn-primary" id="wl-load-btn" onclick="loadWorklog()">Pull Worklog</button>
  <span id="wl-count" style="color:#5f6368;font-size:13px;"></span>
</div>
<div id="wl-status" class="status-msg"></div>
<table id="wl-table" style="display:none;margin-top:16px;">
  <thead>
    <tr>
      <th>Month</th><th>Project</th><th>Issue</th><th>Summary</th><th>Time Spent</th><th>Hours</th><th>Started</th>
    </tr>
  </thead>
  <tbody id="wl-body"></tbody>
</table>
```

- [ ] **Step 3: Add worklog JS to `JavaScript.html`**

```javascript
// ── Worklog Tab ──────────────────────────────────────────────
function loadWorklog() {
  setLoading('wl-load-btn', true, 'Pull Worklog');
  showStatus('wl-status', 'Fetching worklogs for this year…', 'info');
  google.script.run
    .withSuccessHandler(rows => {
      setLoading('wl-load-btn', false, 'Pull Worklog');
      const tbody = document.getElementById('wl-body');
      tbody.innerHTML = rows.map(r =>
        `<tr><td>${r.month}</td><td>${r.projectKey}</td><td>${r.issueKey}</td><td>${r.summary}</td><td>${r.timeSpent}</td><td>${r.hours.toFixed(2)}</td><td>${r.started.slice(0,10)}</td></tr>`
      ).join('');
      document.getElementById('wl-table').style.display = '';
      document.getElementById('wl-count').textContent = `${rows.length} entries`;
      showStatus('wl-status', '', '');
    })
    .withFailureHandler(e => {
      setLoading('wl-load-btn', false, 'Pull Worklog');
      showStatus('wl-status', 'Error: ' + e.message, 'error');
    })
    .getWorklogs();
}
```

- [ ] **Step 4: Test: pull worklog, verify entries appear matching Jira**

- [ ] **Step 5: Commit**

```bash
git add TabWorklog.html Code.js JavaScript.html
git commit -m "feat: worklog tab — pull YTD worklog entries from Jira"
```

---

## Task 12: Cleanup and Final Deployment

**Files:**
- Delete (or archive): old `CreateNewJira.html`
- Modify: `Code.js` — remove `onOpen`, `collectConfig`, `collectJiraUrl`, `retrieveJiraIssues`, `importCalendarEventsToSheet`, `scheduleCalendarEvents` (old sheet-based versions), `createJira`, `getDropdownValues`, `sendTime`, `pullWorklog`

- [ ] **Step 1: Remove all sheet-dependent functions from `Code.js`**

Functions to remove: `onOpen`, `collectConfig`, `collectJiraUrl`, `retrieveJiraIssues`, `importCalendarEventsToSheet`, `scheduleCalendarEvents` (old), `createJira`, `getDropdownValues`, `sendTime`, `pullWorklog`.

Functions to keep: `doGet`, `include`, `getConfig`, `saveJiraUrl`, `saveApiKey`, `getAllocation`, `saveAllocation`, `getUserEmail`, `getJiraIssues`, `importCalendarEvents`, `scheduleCalendarEvents` (new), `sendTimeEntries`, `makeJira`, `getProjectKeys`, `getWorklogs`, `getUserProperties`, `getAuthHeader_`, `getWorklogTotals_`, `getIssueWorklogTotal_`, `parseTimeSpentHours_`.

- [ ] **Step 2: Delete `CreateNewJira.html`**

- [ ] **Step 3: Update `README.md` — deployment instructions**

Add a section: **Deployment** — how to install clasp (`npm i -g @google/clasp`), authenticate (`clasp login`), create or clone a project (`clasp create --type webapp` or `clasp clone <scriptId>`), push (`clasp push`), and deploy (`clasp deploy`).

- [ ] **Step 4: Full end-to-end smoke test**

1. Open web app URL
2. Config tab: enter Jira URL + API key, set allocation rows, save
3. Assignments tab: load assignments, verify issues appear
4. Enter hours for one issue, click Schedule — verify calendar event created
5. Timecard tab: set date range for current week, import events — verify events appear with correct project
6. Assign a Jira issue to one event, submit — verify worklog in Jira
7. Create Issue tab: create a test issue — verify it appears in Jira
8. Worklog tab: pull YTD — verify entries appear

- [ ] **Step 5: Final commit**

```bash
git add -A
git commit -m "feat: complete standalone web app — remove sheet dependencies, cleanup old code"
```

---

## Known Gaps vs. Original

| Feature | Status |
|---|---|
| AI-suggested Jira task per calendar event (`=ai(...)`) | **Dropped** — replaced with manual dropdown |
| History sheet (log of submitted entries) | **Dropped** — was commented out in original |
| `pullWorklog` sheet output (LogPull tab) | **Replaced** — worklog now displayed in-page |
| Named ranges per project (used by Calendar dropdowns) | **Not needed** — replaced by client-side filtering |
| "Next Monday" auto-populate after submit | **Dropped** — date picker replaces this |
| `showUserProperties` debug function | **Dropped** — not needed in web app |
