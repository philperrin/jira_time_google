# Multi-Tab UX Improvements — Design

## Overview

Six UX changes across four tabs:

1. **Assignments:** Project Name lookup in headings; status-count summary under each heading; collapsible groups
2. **Timecard:** Per-day hour total below each day table; "Daily Summary by Project" heading style fix
3. **Worklog:** Remove detail table; add CSV export button
4. **Config:** Collapsible "Jira Configuration" section (collapsed by default)

No server-side changes. All modifications are client-side in HTML and JavaScript.html.

---

## 1. Assignments Tab

### 1a. Project Name lookup in group headings

`renderAssignments(issues)` currently uses `issue.projectKey` as the h3 text. It now receives a second argument `nameMap` — a `{ projectKey → projectName }` plain object derived from the allocation config — and uses it to look up the display name for each group. Falls back to the project key if not found.

`loadAssignments()` fires two parallel `google.script.run` calls: `getJiraIssues()` (already present) and `getAllocation()` (new). A coordination closure renders only when both have returned. The nameMap is built as:

```javascript
const nameMap = Object.fromEntries(
  (allocationRows || [])
    .filter(r => r.projectKey && r.projectName)
    .map(r => [r.projectKey, r.projectName])
);
```

Signature change: `renderAssignments(issues)` → `renderAssignments(issues, nameMap)`

### 1b. Status count summary under each heading

Immediately after the group heading (and before the table), render a multi-line status summary for the group. Computed from `groupMap[projectKey]` (the issues array for that project).

**Status order** (only statuses present in the group are shown):
1. To Do
2. In Progress
3. Done
4. TIME TRACKING ONLY
5. Any other statuses (alphabetical)

Each status appears on its own line. Format per line: `To Do: 2`

Rendered as a `<p class="project-status-summary">` element inside the collapsible body (see 1c), with `white-space: pre-line` or `<br>` between lines so each status appears stacked.

### 1c. Collapsible groups

Each project group is wrapped in a `<div class="project-group">`. Inside it:

```
<div class="project-group">
  <h3 class="project-heading" onclick="toggleGroup(this)">
    <span class="group-toggle">▼</span>
    <span class="group-label">CFA Project</span>
  </h3>
  <div class="project-group-body">
    <p class="project-status-summary">To Do: 2 · In Progress: 3</p>
    <table>...</table>
  </div>
</div>
```

`toggleGroup(h3El)` — defined in JavaScript.html:
- Toggles CSS class `.collapsed` on the parent `.project-group`
- When `.collapsed`: sets the `.group-toggle` span text to `▶`
- When not `.collapsed`: sets the `.group-toggle` span text to `▼`

When `Load Assignments` runs, all groups start expanded (no `.collapsed` class). The user's collapse state does NOT persist across loads.

**New CSS rules in Stylesheet.html:**
```css
.project-group-body { margin-bottom: 20px; }
.project-group.collapsed .project-group-body { display: none; }
.group-toggle { display: inline-block; margin-right: 6px; cursor: pointer; }
.project-heading { cursor: pointer; user-select: none; }
.project-status-summary {
  font-size: 12px; color: #5f6368; margin-bottom: 8px; letter-spacing: 0.3px;
}
```

---

## 2. Timecard Tab

### 2a. Per-day hour total below each day table

After each day's `<table>`, render a `<div class="tc-day-total" data-date="YYYY-MM-DD">` element showing the sum of hours for assigned events on that day.

Format: `Day total: 2.50 h (3 of 5 events assigned)` — or `Day total: — (0 of 5 events assigned)` when nothing is selected.

This div is rendered by `renderTimecardTable()` after each `<table>` is appended to the container. Its content is then kept live by `updateSummaryTable()` (renamed concern: it updates both the summary table AND the per-day totals). No new function needed — extend the existing loop.

Inside `updateSummaryTable()`, after building `sums`, add a per-date pass:

```javascript
dates.forEach(date => {
  const dayEvents = events.filter(e => e.date === date);
  const assigned = dayEvents.filter(e => {
    const sel = document.querySelector(`.tc-issue-select[data-idx="${events.indexOf(e)}"]`);
    return sel && sel.value;
  });
  const total = assigned.reduce((s, e) => s + (e.duration || 0), 0);
  const el = document.querySelector(`.tc-day-total[data-date="${date}"]`);
  if (el) el.textContent =
    'Day total: ' + (total > 0 ? total.toFixed(2) + ' h' : '—') +
    ' (' + assigned.length + ' of ' + dayEvents.length + ' events assigned)';
});
```

**New CSS rule in Stylesheet.html:**
```css
.tc-day-total {
  font-size: 12px; color: #5f6368; text-align: right;
  margin: 4px 0 16px; padding-right: 4px;
}
```

### 2b. "Daily Summary by Project" heading style

In `TabTimecard.html`, change:
```html
<h3 class="section-title" style="font-size:14px;">Daily Summary by Project</h3>
```
to:
```html
<h3 class="project-heading">Daily Summary by Project</h3>
```

The `project-heading` class already provides uppercase, bold, grey styling matching the day table headings. No new CSS needed.

---

## 3. Worklog Tab

### 3a. Remove detail table

Remove the entire `<table id="wl-table">` block (including its `<thead>` and `<tbody id="wl-body">`) from `TabWorklog.html`. It will not appear on the page at all.

In `loadWorklog()` (JavaScript.html):
- Remove `const tbody = document.getElementById('wl-body'); tbody.innerHTML = rows.map(...).join('');`
- Remove `document.getElementById('wl-table').style.display = '';`
- Add `window.worklogRows = rows;` to retain the data for CSV export

### 3b. CSV export button

In `TabWorklog.html`, add an export button to the toolbar (hidden initially):
```html
<div class="toolbar">
  <button class="btn btn-primary" id="wl-load-btn" onclick="loadWorklog()">Pull Worklog</button>
  <button class="btn btn-secondary" id="wl-export-btn" style="display:none;" onclick="exportWorklogCsv()">Export CSV</button>
  <span id="wl-count" style="color:#5f6368;font-size:13px;"></span>
</div>
```

In `loadWorklog()`, after storing rows: `document.getElementById('wl-export-btn').style.display = '';`

`exportWorklogCsv()` function in JavaScript.html:
```javascript
function exportWorklogCsv() {
  const rows = window.worklogRows || [];
  if (!rows.length) return;
  const headers = ['Month', 'Project', 'Issue', 'Summary', 'Time Spent', 'Hours', 'Started'];
  const escape = v => '"' + String(v || '').replace(/"/g, '""') + '"';
  const lines = [headers.map(escape).join(',')].concat(
    rows.map(r => [r.month, r.projectKey, r.issueKey, r.summary, r.timeSpent, r.hours.toFixed(2), r.started.slice(0,10)].map(escape).join(','))
  );
  const blob = new Blob([lines.join('\r\n')], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'jira-worklog-' + new Date().getFullYear() + '.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
```

`window.worklogRows` is initialized to `[]` alongside `window.timecardEvents` and `window.issueOptions` at module level.

---

## 4. Config Tab

### 4a. Collapsible Jira Configuration section

In `TabConfig.html`, wrap the Jira config form in a collapsible container and make the heading clickable. The section is **collapsed by default**.

New structure:
```html
<div class="project-group collapsed" id="cfg-jira-group">
  <h2 class="section-title cfg-collapsible-heading" onclick="toggleGroup(this)">
    <span class="group-toggle">▶</span>
    <span class="group-label">Jira Configuration</span>
  </h2>
  <div class="project-group-body">
    <!-- existing form content unchanged -->
  </div>
</div>
```

`toggleGroup()` already handles the `.collapsed` toggle and `▶`/`▼` swap (added in Assignments changes). No additional function needed.

**One additional CSS rule** to make `section-title` headings also respond to the cursor:
```css
.cfg-collapsible-heading { cursor: pointer; user-select: none; margin-bottom: 0; }
```

When collapsed, the form is hidden via the existing `.project-group.collapsed .project-group-body { display: none; }` rule. The Allocation section below the `<hr>` is NOT collapsible and is unaffected.

---

## Data Flow Summary

```
loadAssignments()
  → getJiraIssues() [parallel]       → window.loadedIssues, passed to renderAssignments
  → getAllocation() [parallel]       → nameMap, passed to renderAssignments
  → renderAssignments(issues, nameMap)
      → builds nameMap lookup
      → per project: renders .project-group > h3 (with toggle) + .project-group-body > p.status-summary + table
      → all groups start expanded

user changes a .tc-issue-select
  → updateSummaryTable()
      → recalculates project × date sums (existing)
      → recalculates per-date totals → updates .tc-day-total[data-date] text (new)

loadWorklog()
  → window.worklogRows = rows (new)
  → buildPivot(rows) (existing)
  → shows wl-export-btn (new)

exportWorklogCsv()
  → reads window.worklogRows
  → builds CSV string
  → triggers browser download via Blob URL
```

---

## Files Changed

| File | Change |
|---|---|
| `Stylesheet.html` | Add 5 new CSS rules (project-group-body, collapsed, group-toggle, project-heading cursor, project-status-summary, tc-day-total, cfg-collapsible-heading) |
| `TabAssignments.html` | No change (groups are rendered by JavaScript) |
| `TabTimecard.html` | Change summary heading class from `section-title` to `project-heading`; remove `style="font-size:14px;"` |
| `TabWorklog.html` | Remove `<table id="wl-table">` block; add `#wl-export-btn` to toolbar |
| `TabConfig.html` | Wrap Jira config form in collapsible `.project-group.collapsed` div |
| `JavaScript.html` | Update `loadAssignments()` (parallel allocation fetch + coordination); update `renderAssignments()` (nameMap arg, group wrapper, status summary, toggle); add `toggleGroup()`; update `renderTimecardTable()` (add `.tc-day-total` div per day); update `updateSummaryTable()` (per-day total update); update `loadWorklog()` (store rows, show export btn); add `exportWorklogCsv()` |
