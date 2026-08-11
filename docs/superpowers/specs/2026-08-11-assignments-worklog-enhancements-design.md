# Assignments & Worklog Enhancements — Design

## Overview

Two UI improvements to the standalone Google Apps Script web app:

1. **Assignments tab:** density toggle (Comfortable / Compact) and per-project table grouping
2. **Worklog tab:** month × project pivot summary table above the detail table

No server-side changes required. All logic is client-side, operating on data already returned by existing server functions.

---

## 1. Assignments Tab

### 1a. Per-Project Table Grouping

**Current behaviour:** A single table renders all issues in one flat list with a "Project" column.

**New behaviour:** Issues are grouped by `projectKey` (e.g. `CEOT`, `WMEPO`). Each group renders as:

```
<h3 class="project-heading">CEOT</h3>
<table>
  <thead> ... </thead>
  <tbody> ... (issues for CEOT) ... </tbody>
</table>
```

Groups are rendered in the order projectKeys first appear in the `window.loadedIssues` array (which is already ordered by Jira key ASC, so projects appear in alphabetical key order naturally). The "Project" column is removed from each table since the heading makes it redundant.

The container `#assignments-table` becomes a generic `<div id="assignments-container">` that holds all the per-project sections. The schedule toolbar and `scheduleEvents()` function are unchanged — `data-idx` on each schedule input still maps to the global `window.loadedIssues` array index.

Issues whose `projectKey` does not appear in any allocation config entry are grouped under their own `projectKey` heading normally — no special handling needed.

**Files changed:** `TabAssignments.html`, `JavaScript.html` (`renderAssignments`)

### 1b. Density Toggle

**Placement:** Two small buttons in the Assignments toolbar, right-aligned: `[ Comfortable ]  [ Compact ]`. The active mode button appears visually selected (uses `.btn-primary` style; inactive uses `.btn-secondary`).

**Comfortable mode:** Current row padding — `td { padding: 10px 12px }`, font-size 14px.

**Compact mode:** Reduced row padding — `td { padding: 4px 6px }`, font-size 12px.

**Implementation:** Clicking a density button adds the class `density-compact` or `density-comfortable` to `#assignments-container`. CSS rules scoped to that class override the global `td` padding and font-size. No JavaScript style manipulation — CSS class toggle only.

**Persistence:** The chosen density is saved to `localStorage` under the key `jtt-density` (`'compact'` or `'comfortable'`). On page load, the stored value is applied. Default is `'comfortable'`.

**Files changed:** `TabAssignments.html` (toggle buttons), `Stylesheet.html` (density CSS rules), `JavaScript.html` (toggle logic, localStorage read/write)

---

## 2. Worklog Tab — Summary Pivot Table

### Placement

A summary section renders between the "Pull Worklog" toolbar and the detail table. It is hidden until worklog data loads (same lifecycle as the detail table).

### Structure

```
Monthly Hours by Project

[ pivot table ]
```

**Rows:** All 12 months of the current year, labeled by full month name (January … December).

**Columns:** All unique `projectKey` values present in the returned worklog data, sorted alphabetically. A final **Total** column sums all projects for each month row.

**Values:** Sum of `hours` (decimal) for that project × month combination. Cells with no data show `0.00`. The Total column shows the row sum.

**Footer row:** A **Total** row at the bottom sums each project column across all months, and the grand total in the bottom-right cell.

### Data flow

Built entirely client-side in the `loadWorklog()` success handler, from the `rows` array (`Array<{ projectKey, month, hours, ... }>`) that `getWorklogs()` already returns. No new server function.

```
month field format: "YYYY-MM"  (e.g. "2026-08")
```

Steps:
1. Collect unique `projectKey` values from `rows`, sort alphabetically → column headers
2. Build a map: `{ "2026-08": { "CEOT": 12.5, "WMEPO": 4.0, ... }, ... }`
3. Render 12 rows for the current year (Jan = `YYYY-01` … Dec = `YYYY-12`), looking up values from the map (default 0)

### Styling

The pivot table uses the same global `<table>` styles as all other tables (white background, border-radius, box-shadow, striped hover). The **Total** column header and footer row use a slightly bolder style (matching `<th>` background) to visually separate them.

**Files changed:** `TabWorklog.html` (add summary section), `JavaScript.html` (`loadWorklog` — add pivot build + render)

---

## Summary of File Changes

| File | Change |
|---|---|
| `TabAssignments.html` | Replace single table with `#assignments-container` div + density toggle buttons |
| `TabWorklog.html` | Add `#wl-summary` section above `#wl-table` |
| `Stylesheet.html` | Add `.density-compact` and `.density-comfortable` CSS rules |
| `JavaScript.html` | Update `renderAssignments()` for grouping + toggle; update `loadWorklog()` for pivot |

No changes to `Code.js` or `appsscript.json`.
