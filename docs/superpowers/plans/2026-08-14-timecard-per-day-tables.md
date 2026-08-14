# Timecard Per-Day Tables Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Redesign the Timecard tab to show events grouped by day in separate tables, add a Description column, and add a live-updating project × date summary table below.

**Architecture:** Client-side only. `TabTimecard.html` provides the new HTML scaffolding; `JavaScript.html` contains all rendering and update logic. `submitTimecard()` and `loadTimecardEvents()` are untouched.

**Tech Stack:** Vanilla JavaScript, Google Apps Script HtmlService, DOM manipulation

## Global Constraints

- No changes to `Code.js`, `Stylesheet.html`, or `appsscript.json`
- `submitTimecard()` must remain byte-for-byte identical — it reads `window.timecardEvents[parseInt(sel.dataset.idx)]` for date/issue/start/duration, never from table DOM
- `data-idx` on each `.tc-issue-select` must reference the flat `window.timecardEvents` array index (not a per-group index)
- XSS: all calendar/Jira-sourced strings rendered into innerHTML must pass through `esc()`
- Date heading format: full weekday + month + day + year, e.g. `"Friday, August 14, 2026"`, derived via `new Date(ev.date + 'T00:00:00').toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })`
- Summary table column headers: short date format `"Aug 14"` via `new Date(ev.date + 'T00:00:00').toLocaleDateString('en-US', { month: 'short', day: 'numeric' })`
- Description truncation: first 150 chars of `ev.description`, followed by `…` if longer; empty string renders as empty cell
- Summary table cell value: sum of `ev.duration` where projectKey and date match AND the `.tc-issue-select` for that event has a non-empty value; display `—` (em dash) when sum is zero

---

### Task 1: Replace static table with dynamic container in TabTimecard.html

**Files:**
- Modify: `TabTimecard.html`

**Interfaces:**
- Produces: `<div id="tc-days-container" style="display:none;">` (replaces `<table id="tc-table">` and `<tbody id="tc-body">`)
- Produces: `<div id="tc-summary" style="display:none;margin-top:24px;">` containing `<table id="tc-summary-table">` with `<thead id="tc-summary-head">`, `<tbody id="tc-summary-body">`, `<tfoot id="tc-summary-foot">`
- Preserves unchanged: `#tc-submit-toolbar`, `#tc-submit-btn`, `#tc-submit-status`, `#tc-status`, date range inputs, `#tc-load-btn`

- [ ] **Step 1: Replace the entire contents of TabTimecard.html**

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

<div id="tc-days-container" style="display:none;margin-top:16px;"></div>

<div id="tc-summary" style="display:none;margin-top:24px;">
  <h3 class="section-title" style="font-size:14px;">Daily Summary by Project</h3>
  <table id="tc-summary-table">
    <thead id="tc-summary-head"></thead>
    <tbody id="tc-summary-body"></tbody>
    <tfoot id="tc-summary-foot"></tfoot>
  </table>
</div>

<div id="tc-submit-toolbar" style="display:none;margin-top:16px;" class="toolbar">
  <button class="btn btn-primary" id="tc-submit-btn" onclick="submitTimecard()">📥 Send to Jira</button>
</div>
<div id="tc-submit-status" class="status-msg"></div>
```

- [ ] **Step 2: Verify the file**

Read back `TabTimecard.html` and confirm:
- `<table id="tc-table">` and `<tbody id="tc-body">` are GONE
- `<div id="tc-days-container" style="display:none;margin-top:16px;">` is present
- `<div id="tc-summary">` is present with `display:none` and contains `#tc-summary-head`, `#tc-summary-body`, `#tc-summary-foot`
- `#tc-submit-toolbar`, `#tc-submit-btn`, `#tc-status`, date inputs, `#tc-load-btn` are unchanged

- [ ] **Step 3: Commit**

```bash
git add TabTimecard.html
git commit -m "feat: replace static timecard table with per-day container and summary section"
```

---

### Task 2: Rewrite renderTimecardTable and add updateSummaryTable in JavaScript.html

**Files:**
- Modify: `JavaScript.html` (the Timecard Tab section, currently lines 120–225)

**Interfaces:**
- Consumes: `#tc-days-container`, `#tc-summary`, `#tc-summary-head`, `#tc-summary-body`, `#tc-summary-foot` from Task 1
- Consumes: `window.timecardEvents` (array of event objects with fields: `date`, `title`, `description`, `start`, `end`, `projectKey`, `duration`, `idx` implicit via array position)
- Consumes: `esc()` helper (line 2 of JavaScript.html)
- Produces: updated `renderTimecardTable(events)` — groups by date, renders per-day h3+table, attaches change listeners, calls `updateSummaryTable()`
- Produces: new `updateSummaryTable()` function — recalculates and re-renders summary tbody and tfoot
- Unchanged: `submitTimecard()`, `loadTimecardEvents()`, `initTimecardDates` IIFE

The existing `renderTimecardTable` shows `#tc-table` and `#tc-submit-toolbar`. The new version must show `#tc-days-container`, `#tc-summary`, and `#tc-submit-toolbar` instead.

- [ ] **Step 1: Add `updateSummaryTable()` function just before `submitTimecard()`**

Locate the `function submitTimecard()` definition and insert this function immediately before it:

```javascript
function updateSummaryTable() {
  const events = window.timecardEvents;
  if (!events || !events.length) return;

  const dates = [...new Set(events.map(e => e.date))].sort();
  const projects = [...new Set(events.map(e => e.projectKey))].sort();

  // Build a map: { projectKey: { date: totalHours } }
  const sums = {};
  projects.forEach(p => { sums[p] = {}; dates.forEach(d => { sums[p][d] = 0; }); });

  document.querySelectorAll('.tc-issue-select').forEach(sel => {
    if (!sel.value) return;
    const ev = events[parseInt(sel.dataset.idx)];
    if (!ev) return;
    sums[ev.projectKey][ev.date] = (sums[ev.projectKey][ev.date] || 0) + (ev.duration || 0);
  });

  const colTotals = {};
  dates.forEach(d => { colTotals[d] = 0; });

  document.getElementById('tc-summary-body').innerHTML = projects.map(p => {
    let rowTotal = 0;
    const cells = dates.map(d => {
      const v = sums[p][d] || 0;
      colTotals[d] += v;
      rowTotal += v;
      return `<td>${v > 0 ? v.toFixed(2) : '—'}</td>`;
    });
    return `<tr><td>${esc(p)}</td>${cells.join('')}<td class="pivot-total">${rowTotal > 0 ? rowTotal.toFixed(2) : '—'}</td></tr>`;
  }).join('');

  const grandTotal = dates.reduce((s, d) => s + colTotals[d], 0);
  document.getElementById('tc-summary-foot').innerHTML =
    '<tr class="pivot-total"><td>Total</td>' +
    dates.map(d => `<td>${colTotals[d] > 0 ? colTotals[d].toFixed(2) : '—'}</td>`).join('') +
    `<td class="pivot-total">${grandTotal > 0 ? grandTotal.toFixed(2) : '—'}</td></tr>`;
}
```

- [ ] **Step 2: Replace `renderTimecardTable(events)` with the new grouped version**

Find and replace the entire `renderTimecardTable` function (currently lines 165–203). The old function references `#tc-body` and `#tc-table` — both are gone.

New function:

```javascript
function renderTimecardTable(events) {
  const container = document.getElementById('tc-days-container');
  container.innerHTML = '';

  const dates = [...new Set(events.map(e => e.date))].sort();

  const fmtFull = dateStr => new Date(dateStr + 'T00:00:00').toLocaleDateString('en-US', {
    weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
  });
  const fmtShort = dateStr => new Date(dateStr + 'T00:00:00').toLocaleDateString('en-US', {
    month: 'short', day: 'numeric'
  });

  dates.forEach(date => {
    const dayEvents = events.filter(e => e.date === date);

    const heading = document.createElement('h3');
    heading.className = 'project-heading';
    heading.textContent = fmtFull(date);
    container.appendChild(heading);

    const table = document.createElement('table');
    table.innerHTML = `<thead><tr>
      <th>Event</th><th>Description</th><th>Start</th><th>End</th><th>Project</th><th>Jira Issue</th><th>Duration (h)</th>
    </tr></thead>`;
    const tbody = document.createElement('tbody');

    dayEvents.forEach(ev => {
      const idx = events.indexOf(ev);
      const desc = ev.description || '';
      const truncated = desc.length > 150 ? desc.slice(0, 150) + '…' : desc;
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td title="${esc(ev.description)}">${esc(ev.title)}</td>
        <td>${esc(truncated)}</td>
        <td>${esc(ev.start)}</td>
        <td>${esc(ev.end)}</td>
        <td>${esc(ev.projectKey)}</td>
        <td>
          <select data-idx="${idx}" class="tc-issue-select" style="min-width:200px;">
            <option value="">— select —</option>
          </select>
        </td>
        <td>${esc(String(ev.duration))}</td>`;
      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    container.appendChild(table);
  });

  // Build summary table headers (static — built once)
  const shortHeaders = dates.map(fmtShort);
  document.getElementById('tc-summary-head').innerHTML =
    '<tr><th>Project</th>' +
    shortHeaders.map(h => `<th>${esc(h)}</th>`).join('') +
    '<th class="pivot-total">Total</th></tr>';

  container.style.display = '';
  document.getElementById('tc-summary').style.display = '';
  document.getElementById('tc-submit-toolbar').style.display = '';

  // Populate issue dropdowns once issues are loaded (poll briefly)
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
        sel.addEventListener('change', updateSummaryTable);
      });
      updateSummaryTable();
    }
  }, 300);
}
```

Note: `updateSummaryTable()` is called inside the polling interval's final callback, after dropdowns are populated and change listeners are attached. This ensures the summary initialises correctly even if issue options load before or after calendar events.

- [ ] **Step 3: Verify no references to #tc-table or #tc-body remain**

Read back the Timecard section of `JavaScript.html` and confirm:
- `document.getElementById('tc-table')` does not appear
- `document.getElementById('tc-body')` does not appear
- `renderTimecardTable` groups by date, creates `h3.project-heading` per date
- `updateSummaryTable` is defined and called at the end of the polling interval
- Each `.tc-issue-select` gets `addEventListener('change', updateSummaryTable)` inside the polling callback
- `submitTimecard()` is byte-for-byte unchanged

- [ ] **Step 4: Commit**

```bash
git add JavaScript.html
git commit -m "feat: add per-day tables, description column, and live summary to timecard tab"
```
