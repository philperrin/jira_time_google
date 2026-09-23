# Spec: New "Allocation" Tab

Adds a tab where the user manually enters a monthly hours-allocation number per project (year selectable), rendered as a month × project pivot table of number inputs, persisted per Google account across sessions.

---

## 1. Tab placement

**Files:** `Index.html`

- New nav button `<button class="tab-btn" data-tab="allocation">Allocation</button>` inserted **after** the "Worklog" button and **before** the "Config" button.
- New `<section id="tab-allocation" class="tab-section"><?!= include('TabAllocation'); ?></section>` inserted after the `tab-worklog` section and before `tab-config`, matching the existing pattern for every other tab.

## 2. New file: `TabAllocation.html`

Modeled directly on `TabWorklog.html`'s toolbar + pivot table structure:

```html
<div class="toolbar">
  <select id="alloc-year" style="width:auto;"></select>
  <button class="btn btn-primary" id="alloc-load-btn" onclick="loadAllocation()">Set/Review Allocations</button>
</div>
<div id="alloc-status" class="status-msg"></div>

<div id="alloc-summary" style="display:none;margin-top:16px;">
  <h3 class="section-title" style="font-size:14px;">Monthly Allocation by Project</h3>
  <table id="alloc-pivot">
    <thead id="alloc-pivot-head"></thead>
    <tbody id="alloc-pivot-body"></tbody>
    <tfoot id="alloc-pivot-foot"></tfoot>
  </table>
  <div class="toolbar" style="margin-top:12px;">
    <button class="btn btn-primary" id="alloc-save-btn" onclick="saveAllocationGrid()">Save Allocations</button>
  </div>
</div>
```

- Year `<select>` (`#alloc-year`): three options — current year, previous year, next year (`currentYear - 1`, `currentYear`, `currentYear + 1`), default-selected to the current year. This differs from the Worklog tab's "current + 3 prior years" range per the user's explicit request (current/previous/next).
- "Set/Review Allocations" button populates the table (see section 4) — it does not, by itself, imply anything is saved yet.
- A dedicated **"Save Allocations"** button below the table persists the currently-entered grid values (see section 5). Matches the rest of the app's config-saving UX (`TabConfig.html`'s "Save"/"Save Allocation" buttons) — an explicit save action rather than autosave-on-blur. **Confirmed.**

## 3. Server: determine eligible project columns (`Code.js`)

**New function** `getAllocationTabData(year)`:

- Reuses the existing worklog-fetching logic in `getWorklogs(year)` to determine which Jira project keys have at least one worklog entry for the current user in the selected year. (Implementation detail: either call `getWorklogs(year)` directly and derive `[...new Set(rows.map(r => r.projectKey))]`, or factor the shared fetch into a private helper if avoiding the double JSON round-trip through the worklog rows is preferred — functionally equivalent either way.)
- Excludes any project key where the matching row in `getAllocation()` (the Config tab's allocation table) has `ignore: true`. A project with no matching allocation row at all (never configured on the Config tab) is **not** excluded — only an explicit `ignore` checkbox hides it.
- Sorts the resulting project key list alphabetically (matching `buildPivot()`'s `.sort()` behavior on the Worklog tab).
- Loads any previously saved allocation numbers for that year via `getAllocationValues_(year)` (section 5).
- Returns:
  ```js
  {
    projects: ['CFA', 'Experian', 'MGIC', ...],   // sorted project keys with ≥1 worklog this year, minus ignored
    values: { 'CFA': { '1': 38, '2': 38, ... }, 'MGIC': { '2': 60, ... } }  // month numbers 1-12, sparse — only cells the user has filled in are present
  }
  ```

**Column identity note:** the Worklog tab's pivot already keys and labels columns by `projectKey` as returned from Jira (see `buildPivot()` in `JavaScript.html` and `getWorklogs`'s `projectKey: issue.fields.project.key`), not by the Config tab's separate `projectName` field. The Allocation tab follows the same convention for consistency — column headers are the Jira project key string, not the friendlier "Project Name" from Config. **Confirmed.**

## 4. Client: rendering the grid (`JavaScript.html`)

**New functions**, alongside the existing Worklog-tab block:

- `initAllocationYearSelect()` — IIFE populating `#alloc-year` with `[currentYear - 1, currentYear, currentYear + 1]`, default-selected `currentYear`. Mirrors `initWorklogYearSelect()`'s DOMContentLoaded-guard pattern.
- `loadAllocation()` — reads `#alloc-year`, calls `getAllocationTabData(year)` via `google.script.run`, on success calls `buildAllocationGrid(data, year)` and shows `#alloc-summary`; on failure shows an error via `showStatus('alloc-status', ...)`. Mirrors `loadWorklog()`.
- `buildAllocationGrid(data, year)` — builds `#alloc-pivot-head`/`-body`/`-foot`, structured like `buildPivot()`:
  - Header row: `Month` + one `<th>` per `data.projects[i]` + a `Total` column.
  - One body row per month (`January`…`December`), each project column rendered as `<td><input type="number" step="any" class="alloc-input" data-project="${esc(project)}" data-month="${monthIndex+1}" value="${data.values[project]?.[monthIndex+1] ?? ''}"></td>`. `step="any"` (rather than `step="0.25"` used for scheduling durations elsewhere) since the screenshot shows values like `17.3` that aren't quarter-hour-aligned.
  - A row-total `<td>` per month, computed client-side as the live sum of that row's inputs.
  - A footer row with a column-total per project plus a grand total in the bottom-right cell, computed the same way.
  - Row/column/grand totals **recompute live on every input change** (`input` event listener attached once via delegation on `#alloc-pivot-body`, calling a `recomputeAllocationTotals()` helper) — so the user sees updated sums as they type, without needing to click Save first.
- `recomputeAllocationTotals()` — reads every `.alloc-input` currently in the DOM, recalculates row totals (per `<tr>`), column totals (per project, summed down `#alloc-pivot-body`), and the grand total, writing them into the pre-existing total cells created in `buildAllocationGrid`. Empty/non-numeric inputs count as `0` toward totals (matching the Worklog pivot's `v.toFixed(2)` / `|| 0` pattern), but an empty input's own cell renders as blank, not `0.00` — the input keeps whatever the user actually typed (or nothing).

## 5. Server: persisting entered values (`Code.js`)

**New user-properties key:** `ALLOCATION_VALUES`, alongside the existing `ALLOCATION` (Config tab's color/project mapping) and `JIRA_BASE_URL`/`JIRA_API_KEY` keys, scoped per-user the same way (`getUserProperties()`).

Stored shape (one JSON blob covering all years, matching the existing single-property pattern used for `ALLOCATION`):
```js
{
  "2026": { "CFA": { "1": 38, "2": 38, "5": 68 }, "MGIC": { "2": 60 } },
  "2025": { ... }
}
```

**New functions:**
- `getAllocationValues_(year)` (private helper, called from `getAllocationTabData`) — reads `ALLOCATION_VALUES`, returns `parsed[year] || {}`.
- `saveAllocationGrid(year, values)` (public, called from the client's Save button) — reads the existing `ALLOCATION_VALUES` blob, replaces only the `values[year]` key with the newly-submitted grid (so other years' saved data is untouched), writes the blob back. `values` is exactly the `{ projectKey: { month: number } }` shape built client-side from the current grid's non-empty inputs.

**Client `saveAllocationGrid()`** (in `JavaScript.html`, named the same as the server function per this codebase's existing convention of same-named client wrapper functions, e.g. `saveAllocation()`/`saveAllocationTable()` in Config):
- Walks all `.alloc-input` elements, builds `{ projectKey: { month: number } }` skipping blank/non-numeric cells entirely (not storing `0` for a cell the user never touched, vs. a cell they explicitly typed `0` into — an explicit `0` **is** saved).
- Calls `google.script.run...saveAllocationGrid(year, values)`, showing success/error via `showStatus('alloc-status', ...)` matching other Save buttons' feedback pattern.

## Edge cases

- **No projects with worklogs in the selected year:** `getAllocationTabData` returns `{ projects: [], values: {} }`. The client renders the table with just a `Month` column and a `Total` column (all zero) rather than showing an error — matches how `buildPivot()` on the Worklog tab already tolerates an empty `projects` array.
- **A project had worklogs (and saved allocation values) in a prior year but is now `ignore`d in Config:** its column simply won't appear when that year is loaded going forward; previously saved values for it remain in `ALLOCATION_VALUES` untouched (not deleted), so un-ignoring it later restores the historical numbers.
- **Switching years without saving:** unsaved in-progress edits in the currently-displayed grid are discarded (no unsaved-changes warning) when the user picks a different year and clicks "Set/Review Allocations" again — matches the Worklog tab's behavior of `loadWorklog()` fully replacing the pivot on each pull.
- **Non-numeric / negative input:** the browser's native `<input type="number">` handling is relied on for basic validation (as elsewhere in the app, e.g. the Assignments tab's duration inputs); no additional server-side validation is added beyond `parseFloat`-with-fallback when reading each cell client-side before save.
- **Decimal values:** explicitly supported (`step="any"`), per the screenshot showing `17.3`.

## Testing

- Load the Allocation tab: confirm it appears between Worklog and Config in the nav, and the year dropdown defaults to the current year with exactly three options (previous, current, next).
- Click "Set/Review Allocations" for a year with existing worklogs across several projects: confirm one column per project with ≥1 worklog that year, alphabetically ordered, excluding any project marked `ignore` on the Config tab.
- Type values into several cells: confirm row totals, column totals, and the grand total update live without needing to click Save.
- Click "Save Allocations", reload the page, switch back to the Allocation tab, select the same year, click "Set/Review Allocations" again: confirm previously entered values repopulate the correct cells.
- Select a different year, enter different values, save, then switch back to the first year: confirm each year's saved values are independent (no cross-year overwrite).
- For a year with zero worklogs: confirm the table still renders (empty of project columns) rather than erroring.

---

## Summary of decisions confirmed with user

- Persist via an explicit "Save Allocations" button (not autosave-on-blur), matching the Config tab's existing save conventions.
- Column headers use the raw Jira project key, matching the existing Worklog tab pivot, not the Config tab's friendlier `projectName`.
