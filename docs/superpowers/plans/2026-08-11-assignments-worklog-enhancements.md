# Assignments & Worklog UI Enhancements Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Comfortable/Compact density toggle and per-project table grouping to the Assignments tab, and add a month × project pivot summary table to the Worklog tab.

**Architecture:** All changes are client-side HTML, CSS, and JavaScript. No server-side changes. Assignments renders grouped `<table>` elements per projectKey with a CSS-class density toggle. Worklog builds a pivot table from existing `getWorklogs()` return data.

**Tech Stack:** Google Apps Script HtmlService, vanilla JavaScript, CSS class toggling, localStorage

## Global Constraints

- No changes to `Code.js` or `appsscript.json`
- XSS protection: all Jira/Calendar-sourced values rendered into innerHTML must pass through `esc()`
- `data-idx` on schedule inputs must continue referencing `window.loadedIssues` array indices correctly
- `scheduleEvents()` function and its `toSchedule` shape must remain unchanged
- Density preference persisted to `localStorage` key `jtt-density`, values `'comfortable'` or `'compact'`
- Default density is `'comfortable'`
- Pivot table rows: all 12 months of the current calendar year, labeled January–December
- Pivot columns: all unique `projectKey` values in the returned worklog data, sorted alphabetically
- Pivot values: sum of `hours` (decimal); cells with no data show `0.00`
- A `Total` column is the rightmost column; a `Total` footer row is the bottom row

---

### Task 1: CSS density rules in Stylesheet.html

**Files:**
- Modify: `Stylesheet.html`

**Interfaces:**
- Produces: `.density-compact td`, `.density-comfortable td` CSS rules used by Task 2's toggle and Task 3's container

- [ ] **Step 1: Open Stylesheet.html and locate the existing `td` rule**

The existing rule is on line 13:
```css
td { padding: 10px 12px; border-top: 1px solid #f1f3f4; vertical-align: middle; }
```
This is the "comfortable" default. The density rules will override padding and font-size when applied to a scoped container.

- [ ] **Step 2: Add density CSS rules after the existing `td` rule**

Insert immediately after the `td { ... }` line:
```css
.density-comfortable td { padding: 10px 12px; font-size: 14px; }
.density-compact td { padding: 4px 6px; font-size: 12px; }
```

- [ ] **Step 3: Add project heading style**

Insert after the density rules:
```css
.project-heading { font-size: 14px; font-weight: 500; color: #5f6368; margin: 20px 0 8px; text-transform: uppercase; letter-spacing: 0.5px; }
.project-heading:first-child { margin-top: 0; }
```

- [ ] **Step 4: Add pivot table Total column/row highlight style**

Insert after the project heading rules:
```css
.pivot-total { font-weight: 600; background: #f1f3f4; }
```

- [ ] **Step 5: Verify the file looks correct (no syntax errors, rules in logical order)**

Read back `Stylesheet.html` and confirm all four new rules are present and correctly placed within the `<style>` block.

- [ ] **Step 6: Commit**

```bash
git add Stylesheet.html
git commit -m "feat: add density toggle and pivot table CSS rules"
```

---

### Task 2: Assignments tab HTML — container and density toggle buttons

**Files:**
- Modify: `TabAssignments.html`

**Interfaces:**
- Consumes: CSS classes `.density-comfortable`, `.density-compact` from Task 1; `.project-heading` from Task 1
- Produces: `#assignments-container` div (replaces `#assignments-table`); `#density-comfortable-btn` and `#density-compact-btn` buttons; `#schedule-toolbar` and `#schedule-btn` unchanged

- [ ] **Step 1: Replace the entire contents of TabAssignments.html**

The new content replaces the static `<table id="assignments-table">` with a div container and adds density toggle buttons in the toolbar. The schedule toolbar and its button remain identical.

```html
<h2 class="section-title">Assignments</h2>
<div class="toolbar">
  <button class="btn btn-primary" id="assignments-load-btn" onclick="loadAssignments()">Load Assignments</button>
  <span id="assignments-count" style="color:#5f6368;font-size:13px;"></span>
  <span style="flex:1;"></span>
  <button class="btn btn-secondary" id="density-comfortable-btn" onclick="setDensity('comfortable')">Comfortable</button>
  <button class="btn btn-secondary" id="density-compact-btn" onclick="setDensity('compact')">Compact</button>
</div>
<div id="assignments-status" class="status-msg"></div>

<div id="assignments-container" style="display:none;"></div>

<div id="schedule-toolbar" style="display:none;margin-top:16px;" class="toolbar">
  <button class="btn btn-secondary" id="schedule-btn" onclick="scheduleEvents()">⏱ Schedule Calendar Events</button>
</div>
<div id="schedule-status" class="status-msg"></div>
```

- [ ] **Step 2: Verify the file is correct**

Read back `TabAssignments.html` and confirm:
- `#assignments-container` div is present (no static table)
- Density buttons have correct `onclick="setDensity('comfortable')"` and `onclick="setDensity('compact')"`
- `#schedule-btn` and `#schedule-toolbar` are unchanged

- [ ] **Step 3: Commit**

```bash
git add TabAssignments.html
git commit -m "feat: replace static assignments table with dynamic container and density toggle buttons"
```

---

### Task 3: JavaScript — density toggle and renderAssignments grouping

**Files:**
- Modify: `JavaScript.html` (the Assignments Tab section, lines 198–269)

**Interfaces:**
- Consumes: `#assignments-container` (from Task 2); CSS classes from Task 1; `window.loadedIssues` global array; `esc()` helper (line 2)
- Produces: updated `renderAssignments(issues)` that groups by projectKey; new `setDensity(mode)` function; `scheduleEvents()` unchanged

The `loadAssignments()` function changes only in that it now targets `#assignments-container` instead of `#assignments-table` for `style.display`, and calls `renderAssignments` as before.

- [ ] **Step 1: Add `setDensity` function and localStorage init after the existing `esc()` helper**

Locate the block at the top of `JavaScript.html` just after the `esc()` function definition (around line 3). Insert a new function and an IIFE that reads from localStorage:

```javascript
function setDensity(mode) {
  const container = document.getElementById('assignments-container');
  if (container) {
    container.classList.remove('density-comfortable', 'density-compact');
    container.classList.add('density-' + mode);
  }
  const comfortBtn = document.getElementById('density-comfortable-btn');
  const compactBtn = document.getElementById('density-compact-btn');
  if (comfortBtn && compactBtn) {
    if (mode === 'compact') {
      comfortBtn.className = 'btn btn-secondary';
      compactBtn.className = 'btn btn-primary';
    } else {
      comfortBtn.className = 'btn btn-primary';
      compactBtn.className = 'btn btn-secondary';
    }
  }
  localStorage.setItem('jtt-density', mode);
}

(function initDensity() {
  const saved = localStorage.getItem('jtt-density') || 'comfortable';
  document.addEventListener('DOMContentLoaded', () => setDensity(saved));
})();
```

- [ ] **Step 2: Replace `renderAssignments` with the grouped version**

Find and replace the entire `renderAssignments` function (lines 219–235 in current file):

Old function:
```javascript
function renderAssignments(issues) {
  const tbody = document.getElementById('assignments-body');
  tbody.innerHTML = '';
  issues.forEach((issue, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><a href="${esc(issue.link)}" target="_blank">${esc(issue.key)}</a></td>
      <td>${esc(issue.name)}</td>
      <td>${esc(issue.project)}</td>
      <td>${esc(issue.status)}</td>
      <td>${esc(issue.timeLogged)}</td>
      <td><input type="number" min="0" step="0.25" style="width:80px;" data-idx="${idx}" class="schedule-input" placeholder="0"></td>`;
    tbody.appendChild(tr);
  });
  document.getElementById('assignments-table').style.display = '';
  document.getElementById('schedule-toolbar').style.display = '';
}
```

New function:
```javascript
function renderAssignments(issues) {
  const container = document.getElementById('assignments-container');
  container.innerHTML = '';

  const groups = [];
  const groupMap = {};
  issues.forEach((issue, idx) => {
    const key = issue.projectKey || 'Unknown';
    if (!groupMap[key]) {
      groupMap[key] = [];
      groups.push(key);
    }
    groupMap[key].push({ issue, idx });
  });

  groups.forEach(projectKey => {
    const heading = document.createElement('h3');
    heading.className = 'project-heading';
    heading.textContent = projectKey;
    container.appendChild(heading);

    const table = document.createElement('table');
    table.innerHTML = `<thead><tr>
      <th>Key</th><th>Name</th><th>Status</th><th>Time Logged (h)</th><th>Schedule (h)</th>
    </tr></thead>`;
    const tbody = document.createElement('tbody');
    groupMap[projectKey].forEach(({ issue, idx }) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><a href="${esc(issue.link)}" target="_blank">${esc(issue.key)}</a></td>
        <td>${esc(issue.name)}</td>
        <td>${esc(issue.status)}</td>
        <td>${esc(issue.timeLogged)}</td>
        <td><input type="number" min="0" step="0.25" style="width:80px;" data-idx="${idx}" class="schedule-input" placeholder="0"></td>`;
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    container.appendChild(table);
  });

  container.style.display = '';
  document.getElementById('schedule-toolbar').style.display = '';
  setDensity(localStorage.getItem('jtt-density') || 'comfortable');
}
```

- [ ] **Step 3: Update `loadAssignments` to not reference `#assignments-table`**

In `loadAssignments()`, the count update line references no element that changed, but confirm there is no `document.getElementById('assignments-table')` call remaining. The function currently shows `#assignments-count` and calls `renderAssignments` — both still correct. No changes needed to `loadAssignments` itself.

- [ ] **Step 4: Confirm `scheduleEvents()` is unchanged**

Read the `scheduleEvents` function and verify it still queries `.schedule-input` by class and reads `data-idx` — this works regardless of how many tables contain the inputs. No changes needed.

- [ ] **Step 5: Verify the full Assignments section of JavaScript.html**

Read back JavaScript.html and confirm:
- `setDensity` function is present
- `initDensity` IIFE is present
- `renderAssignments` groups by `issue.projectKey` and no longer references `#assignments-body` or `#assignments-table`
- No reference to `#assignments-table` remains in this section

- [ ] **Step 6: Commit**

```bash
git add JavaScript.html
git commit -m "feat: add density toggle logic and per-project grouping to assignments render"
```

---

### Task 4: Worklog tab HTML — add pivot summary section

**Files:**
- Modify: `TabWorklog.html`

**Interfaces:**
- Produces: `#wl-summary` section (div wrapping an h3 and `<table id="wl-pivot">`) inserted between `#wl-status` and `#wl-table`; hidden by default via `style="display:none;"`

- [ ] **Step 1: Replace the entire contents of TabWorklog.html**

```html
<h2 class="section-title">Worklog (YTD)</h2>
<div class="toolbar">
  <button class="btn btn-primary" id="wl-load-btn" onclick="loadWorklog()">Pull Worklog</button>
  <span id="wl-count" style="color:#5f6368;font-size:13px;"></span>
</div>
<div id="wl-status" class="status-msg"></div>

<div id="wl-summary" style="display:none;margin-top:16px;">
  <h3 class="section-title" style="font-size:14px;">Monthly Hours by Project</h3>
  <table id="wl-pivot">
    <thead id="wl-pivot-head"></thead>
    <tbody id="wl-pivot-body"></tbody>
    <tfoot id="wl-pivot-foot"></tfoot>
  </table>
</div>

<table id="wl-table" style="display:none;margin-top:16px;">
  <thead>
    <tr>
      <th>Month</th><th>Project</th><th>Issue</th><th>Summary</th><th>Time Spent</th><th>Hours</th><th>Started</th>
    </tr>
  </thead>
  <tbody id="wl-body"></tbody>
</table>
```

- [ ] **Step 2: Verify the file**

Read back `TabWorklog.html` and confirm:
- `#wl-summary` div is present with `display:none`
- `#wl-pivot-head`, `#wl-pivot-body`, `#wl-pivot-foot` are present inside `#wl-pivot`
- `#wl-table` and `#wl-body` remain unchanged

- [ ] **Step 3: Commit**

```bash
git add TabWorklog.html
git commit -m "feat: add worklog pivot summary section to TabWorklog.html"
```

---

### Task 5: JavaScript — worklog pivot table rendering

**Files:**
- Modify: `JavaScript.html` (the Worklog Tab section, around line 303)

**Interfaces:**
- Consumes: `rows` array from `getWorklogs()` — each row has shape `{ month: "YYYY-MM", projectKey: string, issueKey: string, summary: string, timeSpent: string, hours: number, started: string }` (month is "YYYY-MM" format)
- Consumes: `#wl-pivot-head`, `#wl-pivot-body`, `#wl-pivot-foot` from Task 4; `#wl-summary` div from Task 4
- Produces: rendered pivot table shown above detail table; `#wl-summary` made visible when data loads

- [ ] **Step 1: Add a `buildPivot(rows)` helper function just before `loadWorklog()`**

Locate the comment `// ── Worklog Tab ──────────────────────────────────────────────` (around line 302) and insert the helper before `function loadWorklog()`:

```javascript
function buildPivot(rows) {
  const year = new Date().getFullYear();
  const MONTHS = ['January','February','March','April','May','June',
                  'July','August','September','October','November','December'];

  const projects = [...new Set(rows.map(r => r.projectKey))].sort();
  const data = {};
  rows.forEach(r => {
    if (!data[r.month]) data[r.month] = {};
    data[r.month][r.projectKey] = (data[r.month][r.projectKey] || 0) + r.hours;
  });

  const head = document.getElementById('wl-pivot-head');
  const body = document.getElementById('wl-pivot-body');
  const foot = document.getElementById('wl-pivot-foot');

  head.innerHTML = '<tr><th>Month</th>' +
    projects.map(p => `<th>${esc(p)}</th>`).join('') +
    '<th class="pivot-total">Total</th></tr>';

  const colTotals = new Array(projects.length).fill(0);
  body.innerHTML = MONTHS.map((name, mi) => {
    const key = `${year}-${String(mi + 1).padStart(2, '0')}`;
    const rowData = data[key] || {};
    let rowTotal = 0;
    const cells = projects.map((p, pi) => {
      const v = rowData[p] || 0;
      colTotals[pi] += v;
      rowTotal += v;
      return `<td>${v.toFixed(2)}</td>`;
    });
    return `<tr><td>${name}</td>${cells.join('')}<td class="pivot-total">${rowTotal.toFixed(2)}</td></tr>`;
  }).join('');

  const grandTotal = colTotals.reduce((s, v) => s + v, 0);
  foot.innerHTML = '<tr class="pivot-total"><td>Total</td>' +
    colTotals.map(v => `<td>${v.toFixed(2)}</td>`).join('') +
    `<td class="pivot-total">${grandTotal.toFixed(2)}</td></tr>`;

  document.getElementById('wl-summary').style.display = '';
}
```

- [ ] **Step 2: Update `loadWorklog()` to call `buildPivot` after rendering detail rows**

Find the current `loadWorklog` success handler (around line 307). The existing handler ends with:
```javascript
document.getElementById('wl-table').style.display = '';
document.getElementById('wl-count').textContent = `${rows.length} entries`;
showStatus('wl-status', '', '');
```

Add `buildPivot(rows);` immediately before `document.getElementById('wl-table').style.display = '';`:

```javascript
.withSuccessHandler(rows => {
  setLoading('wl-load-btn', false, 'Pull Worklog');
  const tbody = document.getElementById('wl-body');
  tbody.innerHTML = rows.map(r =>
    `<tr><td>${esc(r.month)}</td><td>${esc(r.projectKey)}</td><td>${esc(r.issueKey)}</td><td>${esc(r.summary)}</td><td>${esc(r.timeSpent)}</td><td>${esc(r.hours.toFixed(2))}</td><td>${esc(r.started.slice(0,10))}</td></tr>`
  ).join('');
  buildPivot(rows);
  document.getElementById('wl-table').style.display = '';
  document.getElementById('wl-count').textContent = `${rows.length} entries`;
  showStatus('wl-status', '', '');
})
```

- [ ] **Step 3: Verify the Worklog section of JavaScript.html**

Read back the file and confirm:
- `buildPivot(rows)` is defined before `loadWorklog()`
- `buildPivot(rows)` is called inside the success handler
- The `esc()` helper is applied to all `projectKey` values in the pivot header
- `#wl-summary` `style.display` is set to `''` inside `buildPivot`
- No reference to the old `wl-summary` is missing

- [ ] **Step 4: Commit**

```bash
git add JavaScript.html
git commit -m "feat: add worklog pivot table (month x project hours) to worklog tab"
```

---

### Task 6: Manual smoke test (document-only)

This task is a manual verification checklist for deploying and testing in the Google Apps Script web app. No code changes.

**Deploy steps:**
1. In the Google Apps Script editor, click **Deploy → Manage Deployments → Edit (pencil icon) → New version → Deploy**
2. Open the web app URL

**Assignments tab:**
- [ ] Click **Load Assignments** — issues load
- [ ] Verify issues are grouped under `h3` project headings (e.g. CEOT, WMEPO)
- [ ] Verify the **Project** column is gone from each table
- [ ] Click **Compact** — rows become tighter; **Comfortable** button turns blue; **Compact** button turns white
- [ ] Click **Comfortable** — rows return to normal spacing
- [ ] Reload the page — the last density choice is restored from localStorage
- [ ] Enter hours in a schedule input and click **Schedule Calendar Events** — events are created (correct issue keys in the request)

**Worklog tab:**
- [ ] Click **Pull Worklog** — worklogs load
- [ ] Verify **Monthly Hours by Project** table appears above detail table
- [ ] Verify rows are all 12 months (January–December) of the current year
- [ ] Verify columns are projectKey values found in the data, sorted alphabetically, plus **Total**
- [ ] Verify a **Total** footer row is present
- [ ] Verify values are decimal hours summed per project per month (spot-check one project against detail rows)
- [ ] Verify months with no data show `0.00`
