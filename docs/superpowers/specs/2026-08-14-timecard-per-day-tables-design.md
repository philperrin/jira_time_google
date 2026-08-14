# Timecard Tab Redesign — Design

## Overview

Three visual enhancements to the Timecard tab:

1. **Per-day tables** — events grouped by date, each under a formatted date heading, with the Date column removed
2. **Description column** — first 150 characters of the calendar event's description, after the Event column
3. **Live summary table** — project × date pivot below the day tables, updating in real time as Jira Issue dropdowns are changed

No server-side changes. All modifications are client-side in `TabTimecard.html` and `JavaScript.html`. `submitTimecard()` is unchanged — it reads date and issue data from `window.timecardEvents` via `data-idx`, never from the table DOM.

---

## 1. Per-Day Tables

### Current structure

One flat `<table id="tc-table">` with columns: Date | Event | Start | End | Project | Jira Issue | Duration (h)

### New structure

A `<div id="tc-days-container">` replaces the table. For each unique date in the imported events (sorted ascending by date string):

```
<h3 class="day-heading">Friday, August 14, 2026</h3>
<table class="tc-day-table">
  <thead>
    <tr>
      <th>Event</th>
      <th>Description</th>
      <th>Start</th>
      <th>End</th>
      <th>Project</th>
      <th>Jira Issue</th>
      <th>Duration (h)</th>
    </tr>
  </thead>
  <tbody> ... events for that date ... </tbody>
</table>
```

- **Date column removed** — the h3 heading carries the date for the group
- **Date heading format**: Full weekday + month name + day + year — e.g. `"Friday, August 14, 2026"`. Derived client-side from the `ev.date` string (`"YYYY-MM-DD"`) using `toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })` on a local-time Date object (`new Date(ev.date + 'T00:00:00')` to avoid UTC-shift).
- **data-idx** on each `.tc-issue-select` continues to reference the flat `window.timecardEvents` array index — grouping is display-only, not structural

### Confirming submitTimecard() correctness

`submitTimecard()` builds entries by querying `.tc-issue-select` elements and reading `window.timecardEvents[parseInt(sel.dataset.idx)]` for each. It does not read from table cells. Removing the Date column and splitting into multiple tables does not affect this function.

---

## 2. Description Column

### Placement

Second column, after Event, before Start: **Event | Description | Start | End | Project | Jira Issue | Duration (h)**

### Content

`ev.description` is already extracted server-side from the first line of the Google Calendar event description (via the regex `match(/^[^\n_]+/)`).

Client-side truncation: if the string is longer than 150 characters, display the first 150 characters followed by `…`. Otherwise display as-is. Empty strings display as an empty cell.

```javascript
const desc = ev.description || '';
const truncated = desc.length > 150 ? desc.slice(0, 150) + '…' : desc;
```

All output goes through `esc()` before being inserted into innerHTML.

---

## 3. Live Summary Table

### Placement

Below `#tc-days-container`, above `#tc-submit-toolbar`. Hidden (`display:none`) until events load, then shown alongside the day tables.

### Structure

```
<div id="tc-summary" style="display:none; margin-top:24px;">
  <h3 class="section-title" style="font-size:14px;">Daily Summary by Project</h3>
  <table id="tc-summary-table">
    <thead id="tc-summary-head"></thead>
    <tbody id="tc-summary-body"></tbody>
    <tfoot id="tc-summary-foot"></tfoot>
  </table>
</div>
```

### Rows and columns

- **Rows**: one per unique `projectKey` found across all events in `window.timecardEvents`, sorted alphabetically
- **Columns**: one per unique date, sorted ascending (same dates as the day tables). Column headers use short date format: `"Aug 14"` (keeps columns narrow)
- **Final column**: **Total** — sum of selected hours for that project across all dates
- **Footer row**: column totals (sum of selected hours across all projects for each date) + grand total

### Values

A cell value is the sum of `ev.duration` for all events where:
- `ev.projectKey` matches the row's project
- `ev.date` matches the column's date
- The `.tc-issue-select` for that event currently has a non-empty selected value

Cells with a sum of zero (no events selected for that project/date combination) display `—` rather than `0.00`.

### Live update

`updateSummaryTable()` is called:
1. After `renderTimecardTable()` completes (initial render, all cells show `—`)
2. On `change` event of every `.tc-issue-select` element

The function re-reads all dropdown states from the DOM each time it runs. It only updates the tbody and tfoot (headers are static once built). No server calls.

Change listener is attached per dropdown inside `renderTimecardTable()`:
```javascript
select.addEventListener('change', updateSummaryTable);
```

---

## Data Flow Summary

```
loadTimecardEvents()
  → importCalendarEvents() [server]       → window.timecardEvents
  → getJiraIssues() [server, parallel]    → window.issueOptions
  → renderTimecardTable(events)
      → groups events by date
      → renders per-day h3 + table per date
      → attaches change listener on each .tc-issue-select
      → calls updateSummaryTable() [shows all dashes]
      → shows #tc-days-container, #tc-summary, #tc-submit-toolbar

user changes a .tc-issue-select
  → updateSummaryTable()
      → reads all .tc-issue-select values from DOM
      → recalculates project × date sums
      → re-renders #tc-summary-body and #tc-summary-foot

user clicks "Send to Jira"
  → submitTimecard()  [UNCHANGED]
      → reads window.timecardEvents via data-idx
      → filters where issueKey && durationHours > 0
      → calls sendTimeEntries(entries) [server]
```

---

## Files Changed

| File | Change |
|---|---|
| `TabTimecard.html` | Replace `<table id="tc-table">` with `<div id="tc-days-container">`; add `<div id="tc-summary">` section |
| `JavaScript.html` | Rewrite `renderTimecardTable()`; add `updateSummaryTable()`; `submitTimecard()` and `loadTimecardEvents()` unchanged |

No changes to `Code.js`, `Stylesheet.html`, or `appsscript.json`.
