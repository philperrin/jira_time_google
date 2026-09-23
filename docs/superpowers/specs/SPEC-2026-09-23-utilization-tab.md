# Spec: "Utilization" Tab (fills the existing "Pay & Profit" stub)

Adds a tab that reports, for a selected year: a monthly utilization percentage per project (worklog hours ÷ allocated hours), and revenue/cost/profit metrics through the most recently completed pay period.

---

## 0. Decisions confirmed with user

- The uncommitted `Index.html` diff already added a `Pay & Profit` nav button with no matching section/file. This spec **fills that stub** rather than adding a separate tab — the button's label changes to `Utilization`, and its section is wired to a new `TabUtilization.html`.
- Worklog data for a year is **cached persistently per-year** in user properties once collected. "Calculate Utilization" reuses the cache silently on later clicks/sessions and never auto-refetches — a `Refresh` action is the only way to update it (see §4).
- The new **Hourly Rate** and **Overhead Rate** inputs persist via the Config tab's user properties (like Pay Period End Date), but are displayed/edited as inputs on the Utilization tab itself, not added to the Config tab's UI.
- All pay-period math (periods to date, most recent end date, hours/revenue "thru" it) is scoped to the **year selected on the Utilization tab**, not the real-world current year.

## 1. Cache granularity

The request describes caching a "worklog summary (project by month)". Revenue needs hours **thru a specific day** (the most recent pay period end date), which can fall mid-month — a month-granularity cache can't answer that precisely.

**Confirmed with user:** the per-year cache stores the same row-level shape `getWorklogs(year)` already returns (`{projectKey, issueKey, summary, timeSpent, started, hours, month}`), not a pre-aggregated month total. Both the Utilization %-table (aggregated by month) and the Revenue section's day-precise cutoff are derived from these cached rows at render/calc time.

## 2. Tab wiring

**Files:** `Index.html`

- Change `<button class="tab-btn" data-tab="profit">Pay & Profit</button>` → `<button class="tab-btn" data-tab="profit">Utilization</button>` (keep `data-tab="profit"` as-is to avoid touching unrelated code, unless the user prefers renaming the id too — noted as low-risk either way).
- Add `<section id="tab-profit" class="tab-section"><?!= include('TabUtilization'); ?></section>` after the `tab-config` section (matching current button order: Worklog, Allocation, Config, Changelog, **Utilization**).

## 3. Config tab additions (`Code.js`, `TabConfig.html`, `JavaScript.html`)

**New user-properties keys**, alongside `JIRA_BASE_URL`/`JIRA_API_KEY`/`ALLOCATION`/`ALLOCATION_VALUES`:
- `PAY_PERIOD_END_DATE` — ISO date string (`YYYY-MM-DD`) for one known pay-period end date. Every other pay period end date (past or future, any year) is derived from this anchor at a fixed 14-day cadence.
- `UTIL_HOURLY_RATE` — number, gross hourly rate of pay.
- `UTIL_OVERHEAD_RATE` — number, overhead rate as a decimal fraction (e.g. `0.40` for 40%).

**New `Code.js` functions**, next to `saveJiraUrl`/`saveApiKey`:
```js
function getPayPeriodEndDate() { return getUserProperties().getProperty('PAY_PERIOD_END_DATE') || ''; }
function savePayPeriodEndDate(dateStr) { getUserProperties().setProperty('PAY_PERIOD_END_DATE', dateStr); }
function getUtilRates() {
  const props = getUserProperties();
  return {
    hourlyRate: parseFloat(props.getProperty('UTIL_HOURLY_RATE')) || 0,
    overheadRate: parseFloat(props.getProperty('UTIL_OVERHEAD_RATE')) || 0
  };
}
function saveUtilRates(hourlyRate, overheadRate) {
  getUserProperties().setProperty('UTIL_HOURLY_RATE', String(hourlyRate));
  getUserProperties().setProperty('UTIL_OVERHEAD_RATE', String(overheadRate));
}
```

**`TabConfig.html`:** add a labeled input for "Pay Period End Date" (`<input type="date" id="cfg-pay-period-end">`) in the Jira Configuration group (or its own small group) — the Hourly Rate / Overhead Rate inputs live on the Utilization tab per the decision above, not here.

**`JavaScript.html`:**
- `loadConfig()` also calls `getPayPeriodEndDate()` and populates `#cfg-pay-period-end`.
- `saveConfig()` also saves `#cfg-pay-period-end`'s value via `savePayPeriodEndDate(...)` when non-empty, following the existing `pending`/`done`/`fail` pattern.

## 4. Server: worklog cache (`Code.js`)

**New user-properties key:** `WORKLOG_CACHE` — one JSON blob keyed by year, same single-property pattern as `ALLOCATION_VALUES`:
```js
{ "2026": [ {projectKey, issueKey, summary, timeSpent, started, hours, month}, ... ], "2025": [...] }
```

**New functions:**
```js
/** Returns cached worklog rows for a year, or null if never collected. */
function getCachedWorklogs_(year) {
  const raw = getUserProperties().getProperty('WORKLOG_CACHE');
  const parsed = raw ? JSON.parse(raw) : {};
  return parsed[year] || null;
}

/** Stores worklog rows for a year, replacing only that year's entry. */
function setCachedWorklogs_(year, rows) {
  const raw = getUserProperties().getProperty('WORKLOG_CACHE');
  const parsed = raw ? JSON.parse(raw) : {};
  parsed[year] = rows;
  getUserProperties().setProperty('WORKLOG_CACHE', JSON.stringify(parsed));
}

/**
 * Returns worklog rows for a year, collecting from Jira via getWorklogs(year)
 * and caching them if not already cached. forceRefresh bypasses the cache.
 */
function getOrCollectWorklogs(year, forceRefresh) {
  if (!forceRefresh) {
    const cached = getCachedWorklogs_(year);
    if (cached) return cached;
  }
  const rows = getWorklogs(year);
  setCachedWorklogs_(year, rows);
  return rows;
}
```

- `getAllocationTabData(year)` (existing) keeps calling `getWorklogs(year)` directly — it's unrelated to this cache and out of scope for this change.

## 5. Server: pay-period math (`Code.js`)

```js
const PAY_PERIOD_DAYS = 14;

/** Returns all pay-period end dates (as Date objects, ascending) that fall within [Jan 1, Dec 31] of `year`, derived from the PAY_PERIOD_END_DATE anchor at a fixed 14-day cadence. */
function getPayPeriodEndDates_(year) {
  const anchorStr = getPayPeriodEndDate();
  if (!anchorStr) return [];
  const anchor = new Date(anchorStr + 'T00:00:00');
  const yearStart = new Date(year, 0, 1);
  const yearEnd = new Date(year, 11, 31, 23, 59, 59, 999);
  const msPerPeriod = PAY_PERIOD_DAYS * 24 * 60 * 60 * 1000;

  // Number of whole periods between the anchor and yearStart determines
  // which occurrence is the first to land on/after yearStart.
  const periodsFromAnchorToYearStart = Math.ceil((yearStart - anchor) / msPerPeriod);
  const dates = [];
  let n = periodsFromAnchorToYearStart;
  let d = new Date(anchor.getTime() + n * msPerPeriod);
  while (d <= yearEnd) {
    if (d >= yearStart) dates.push(new Date(d));
    n++;
    d = new Date(anchor.getTime() + n * msPerPeriod);
  }
  return dates;
}

/**
 * Returns {payPeriodsToDate, firstPayPeriodEnd, mostRecentPayPeriodEnd} for a
 * year, where "most recent" is the latest period end date <= today (for the
 * current year) — for a past year this is simply that year's last period end
 * date, since all of them are already in the past.
 */
function getPayPeriodSummary_(year) {
  const allDates = getPayPeriodEndDates_(year);
  const today = new Date();
  const eligible = allDates.filter(d => d <= today);
  if (eligible.length === 0) return { payPeriodsToDate: 0, firstPayPeriodEnd: null, mostRecentPayPeriodEnd: null };
  return {
    payPeriodsToDate: eligible.length,
    firstPayPeriodEnd: allDates[0],
    mostRecentPayPeriodEnd: eligible[eligible.length - 1]
  };
}
```

- If `PAY_PERIOD_END_DATE` isn't configured yet, all pay-period-derived metrics come back as zero/blank (see §7 edge cases) rather than throwing.

## 6. Server: utilization + revenue/cost/profit calc (`Code.js`)

**New function**, the one entry point the "Calculate Utilization" button calls:
```js
/**
 * Returns everything the Utilization tab needs to render for a year:
 * - projects: sorted project keys present in the year's worklog rows.
 * - utilization: { [projectKey]: { [month 1-12]: pct } }, plus rowTotals ({ [month]: pct }),
 *   colTotals ({ [projectKey]: pct }), and grandTotal (pct) — see totals note below.
 *   pct = round((hoursThatMonth / allocatedHoursThatMonth) * 100, 1), skipped (null) when
 *   the project has no allocation value for that month (see edge cases).
 * - payPeriods: { firstEnd, mostRecentEnd, payPeriodsToDate, totalCostHours } (ISO date strings for the two *End fields).
 * - revenue: { hoursThruMostRecent, grossRevenue, unratedProjects } (unratedProjects: project
 *   keys excluded from both figures because they have no rate configured — see §8).
 * - cost: { hourlyRate, overheadRate, hourlyCost, grossCost }.
 * - profit: { grossProfit, personalProfitMargin }.
 */
function calculateUtilization(year, forceRefresh) {
  const rows = getOrCollectWorklogs(year, forceRefresh);
  const allocationRows = getAllocation();
  const rateByProject = Object.fromEntries(allocationRows.filter(r => r.projectKey).map(r => [r.projectKey, r.rate]));
  const allocValues = getAllocationValues_(year); // { projectKey: { month: hours } }

  const projects = [...new Set(rows.map(r => r.projectKey))].sort();

  // Monthly hours per project, from cached rows.
  const hoursByProjectMonth = {};
  rows.forEach(r => {
    const m = parseInt(r.month.split('-')[1], 10);
    hoursByProjectMonth[r.projectKey] = hoursByProjectMonth[r.projectKey] || {};
    hoursByProjectMonth[r.projectKey][m] = (hoursByProjectMonth[r.projectKey][m] || 0) + r.hours;
  });

  // Per-cell percentages, plus running sums (hours/allocation, not percentages)
  // for the row/column/grand totals — see §7's totals note.
  const utilization = {};
  const rowHourSum = {}, rowAllocSum = {};       // keyed by month
  const colHourSum = {}, colAllocSum = {};       // keyed by projectKey
  let grandHourSum = 0, grandAllocSum = 0;
  projects.forEach(p => {
    utilization[p] = {};
    for (let m = 1; m <= 12; m++) {
      const allocated = allocValues[p] && allocValues[p][m];
      const worked = (hoursByProjectMonth[p] && hoursByProjectMonth[p][m]) || 0;
      if (allocated == null || allocated === 0) {
        utilization[p][m] = null;
        continue;
      }
      utilization[p][m] = Math.round((worked / allocated) * 1000) / 10;
      rowHourSum[m] = (rowHourSum[m] || 0) + worked;
      rowAllocSum[m] = (rowAllocSum[m] || 0) + allocated;
      colHourSum[p] = (colHourSum[p] || 0) + worked;
      colAllocSum[p] = (colAllocSum[p] || 0) + allocated;
      grandHourSum += worked;
      grandAllocSum += allocated;
    }
  });
  const pct = (h, a) => (a === 0 || a == null) ? null : Math.round((h / a) * 1000) / 10;
  const rowTotals = {};
  for (let m = 1; m <= 12; m++) rowTotals[m] = pct(rowHourSum[m] || 0, rowAllocSum[m] || 0);
  const colTotals = {};
  projects.forEach(p => colTotals[p] = pct(colHourSum[p] || 0, colAllocSum[p] || 0));
  const grandTotal = pct(grandHourSum, grandAllocSum);

  const payPeriodInfo = getPayPeriodSummary_(year);
  const cutoff = payPeriodInfo.mostRecentPayPeriodEnd; // Date or null

  let hoursThruMostRecent = 0, grossRevenue = 0;
  const unratedProjects = new Set();
  if (cutoff) {
    rows.forEach(r => {
      const rate = rateByProject[r.projectKey];
      if (rate == null) { unratedProjects.add(r.projectKey); return; } // excluded entirely, not treated as $0
      const started = new Date(r.started);
      if (started <= cutoff) {
        hoursThruMostRecent += r.hours;
        grossRevenue += r.hours * rate;
      }
    });
  }

  const { hourlyRate, overheadRate } = getUtilRates();
  const hourlyCost = hourlyRate * (1 + overheadRate);
  const totalCostHours = payPeriodInfo.payPeriodsToDate * 80;
  const grossCost = hourlyCost * totalCostHours;
  const grossProfit = grossRevenue - grossCost;
  const personalProfitMargin = grossRevenue === 0 ? null : Math.round((grossProfit / grossRevenue) * 1000) / 10;

  return {
    projects,
    utilization: { cells: utilization, rowTotals, colTotals, grandTotal },
    payPeriods: {
      firstEnd: payPeriodInfo.firstPayPeriodEnd ? payPeriodInfo.firstPayPeriodEnd.toISOString().slice(0, 10) : null,
      mostRecentEnd: cutoff ? cutoff.toISOString().slice(0, 10) : null,
      payPeriodsToDate: payPeriodInfo.payPeriodsToDate,
      totalCostHours
    },
    revenue: { hoursThruMostRecent, grossRevenue, unratedProjects: [...unratedProjects] },
    cost: { hourlyRate, overheadRate, hourlyCost, grossCost },
    profit: { grossProfit, personalProfitMargin }
  };
}
```

- Rounding: percentages round to one decimal (`toFixed(1)`-equivalent) per the request; dollar amounts are not rounded server-side, formatting happens client-side.
- A project with `rate == null` in Config is **excluded entirely** from `hoursThruMostRecent` and `grossRevenue` — its hours are not counted at all, not treated as `$0` (confirmed with user). Its key is still returned in `revenue.unratedProjects` so the UI can surface which projects were skipped (see §7/§8).

## 7. Client: `TabUtilization.html` + `JavaScript.html`

**`TabUtilization.html`**, modeled on `TabAllocation.html`'s toolbar/pivot structure plus new metric blocks matching the screenshot's three-section layout:

```html
<div class="toolbar">
  <select id="util-year" style="width:auto;"></select>
  <button class="btn btn-primary" id="util-calc-btn" onclick="calculateUtilization()">Calculate Utilization</button>
  <button class="btn btn-secondary" id="util-refresh-btn" onclick="calculateUtilization(true)">Refresh Worklog Data</button>
</div>
<div id="util-status" class="status-msg"></div>

<div id="util-results" style="display:none;margin-top:16px;">
  <h3 class="section-title" style="font-size:14px;">Monthly Utilization by Project</h3>
  <table id="util-pivot">
    <thead id="util-pivot-head"></thead>
    <tbody id="util-pivot-body"></tbody>
    <tfoot id="util-pivot-foot"></tfoot>
  </table>

  <h3 class="section-title" style="font-size:14px;margin-top:24px;">Revenue</h3>
  <table class="metrics-table">
    <tr><td>Hours Billed Thru <span id="util-cutoff-label"></span>:</td><td id="util-hours-thru"></td></tr>
    <tr><td>Gross Revenue:</td><td id="util-gross-revenue"></td></tr>
  </table>

  <h3 class="section-title" style="font-size:14px;margin-top:24px;">Cost</h3>
  <table class="metrics-table">
    <tr><td>Hourly Rate of Pay:</td><td><input type="number" step="any" id="util-hourly-rate"></td></tr>
    <tr><td>Overhead Rate (Est., %):</td><td><input type="number" step="any" id="util-overhead-rate"></td></tr>
    <tr><td>Hourly Cost:</td><td id="util-hourly-cost"></td></tr>
    <tr><td>Pay Periods To Date:</td><td id="util-periods-to-date"></td></tr>
    <tr><td>Total Cost Hours:</td><td id="util-total-cost-hours"></td></tr>
    <tr><td>Gross Cost:</td><td id="util-gross-cost"></td></tr>
  </table>
  <div class="toolbar" style="margin-top:8px;">
    <button class="btn btn-secondary" id="util-save-rates-btn" onclick="saveUtilRates()">Save Rates</button>
  </div>

  <h3 class="section-title" style="font-size:14px;margin-top:24px;">Profit</h3>
  <table class="metrics-table">
    <tr><td>Gross Profit:</td><td id="util-gross-profit"></td></tr>
    <tr><td>Personal Profit Margin:</td><td id="util-profit-margin"></td></tr>
  </table>
</div>
```

- Year `<select>` (`#util-year`): current year + 3 prior years, matching the Worklog tab's convention (this tab consumes the same worklog data).
- "Calculate Utilization" calls the server with `forceRefresh=false`; "Refresh Worklog Data" calls it with `forceRefresh=true` — this is the only UI path that re-hits Jira for a year already cached (per §0's decision).
- Overhead Rate is entered/displayed as a whole-number percentage (e.g. `40`) in the input but stored/sent as a decimal fraction (`0.40`) to match the server's `overheadRate` — client divides by 100 before sending, multiplies by 100 for display.

**`JavaScript.html`** new functions, alongside the Allocation tab block:
- `initUtilYearSelect()` — same IIFE pattern as `initWorklogYearSelect()`.
- `calculateUtilization(forceRefresh)` — reads `#util-year` and the two rate inputs (sending current values so a not-yet-saved edit is still used for this calculation), calls `calculateUtilization(year, !!forceRefresh)` via `google.script.run`, on success calls `renderUtilizationPivot(data)` and `renderUtilizationMetrics(data)`, shows `#util-results`; on failure shows an error via `showStatus('util-status', ...)`.
- `renderUtilizationPivot(data)` — builds `#util-pivot-head/-body/-foot` exactly like `buildPivot()` in structure, but each cell is `data.utilization.cells[project][month]` formatted as `${pct.toFixed(1)}%` (or `—` when `null`). The row-total column uses `data.utilization.rowTotals[month]`, the column-total footer row uses `data.utilization.colTotals[project]`, and the bottom-right cell uses `data.utilization.grandTotal` — all three already computed server-side as `sum(hours) / sum(allocation)` over the cells that went into them (not an average of percentages — confirmed with user), so the client only formats them.
- `renderUtilizationMetrics(data)` — writes `data.payPeriods`, `data.revenue`, `data.cost`, `data.profit` into the corresponding elements, formatting currency as `$` + `toLocaleString()` and percentages as `toFixed(1) + '%'`. If `data.revenue.unratedProjects` is non-empty, shows a note near the Revenue section (e.g. via `showStatus('util-status', ...)`, `info` type) listing which project keys were excluded for having no configured rate. Also populates `#util-hourly-rate`/`#util-overhead-rate` from `data.cost` on first load (via a separate `getUtilRates()` call on tab click, mirroring `loadConfig()`) so the fields aren't blank before the user clicks Calculate.
- `saveUtilRates()` — reads the two rate inputs, converts overhead from percentage to fraction, calls `saveUtilRates(hourlyRate, overheadRate)` via `google.script.run`, shows success/error via `showStatus('util-status', ...)`.
- Load rates when the tab is clicked: `document.querySelector('[data-tab="profit"]').addEventListener('click', () => google.script.run.withSuccessHandler(r => { document.getElementById('util-hourly-rate').value = r.hourlyRate; document.getElementById('util-overhead-rate').value = r.overheadRate * 100; }).getUtilRates());`

## 8. Edge cases

- **`PAY_PERIOD_END_DATE` not yet configured:** `getPayPeriodSummary_` returns all nulls/zero; the Revenue and Cost sections render `—`/`0` for pay-period-derived fields instead of erroring, with a note (`showStatus`) prompting the user to set it on the Config tab.
- **Project has worklogs but no allocation value for a given month:** that cell renders `—` (not `0%` or `N/A`-as-error) — allocation of `0` (explicitly entered) is treated differently from "never entered" (`null`), matching the distinction `getAllocationTabData` already respects for saved vs. unset cells.
- **Project has worklogs but no `rate` configured in Config (`rate: null`):** excluded entirely from both `hoursThruMostRecent` and `grossRevenue` — no calculation is performed for it, it is not treated as `$0` (confirmed with user). Its key is surfaced in `revenue.unratedProjects` and the UI shows a note listing the excluded project keys, so a missing rate reads as "excluded, here's why" rather than silently skewing the totals.
- **No worklogs collected yet for the selected year:** `getOrCollectWorklogs` calls `getWorklogs(year)` (same as today's Worklog tab), which can return `[]` if the user genuinely has none — renders an empty pivot and zeroed metrics, not an error.
- **Selected year has no pay-period end dates falling within it** (e.g. `PAY_PERIOD_END_DATE` anchor is many years in the future/past of an edge case) — `getPayPeriodEndDates_` returns `[]`; treated the same as "not configured" above.
- **Switching years:** clicking "Calculate Utilization" for a different year always reflects that year's own cache/state; it does not carry over the previous year's displayed numbers.

## 9. Testing

- Set a Pay Period End Date on Config, save, reload, confirm it round-trips.
- Enter Hourly Rate / Overhead Rate on the Utilization tab, click "Save Rates", reload, click the Utilization tab: confirm the inputs repopulate.
- Click "Calculate Utilization" for a year with no cache yet: confirm it fetches from Jira (slower) and subsequent clicks in the same/later sessions are fast (served from `WORKLOG_CACHE`) until "Refresh Worklog Data" is used.
- Verify the monthly pivot's percentages match hand-computed `hours / allocation * 100` for a couple of project/month cells, including a cell with allocation `0` explicitly entered (shows `0.0%`, not `—`) vs. a cell never entered (shows `—`).
- Verify Pay Periods To Date, Total Cost Hours, Gross Cost, Gross Revenue, Gross Profit, and Personal Profit Margin against a hand-computed example (e.g. reproduce the screenshot's numbers with matching inputs).
- Verify a project with `rate: null` in Config is fully excluded from both `hoursThruMostRecent` and `grossRevenue` (not counted as `$0`), and that its key appears in `revenue.unratedProjects` with a corresponding UI note.
- Verify selecting a past year computes "most recent pay period" as that year's last period end date (not gated on "today"), and a currently in-progress year only counts periods up to today.
- Verify the pivot's row total, column total, and grand-total cells equal `sum(hours) / sum(allocation)` over the contributing cells — not an average of the individual percentages — by hand-computing a case where a simple average would give a visibly different number (e.g. one month with allocation 10/hours 10 = 100% and another with allocation 100/hours 50 = 50%; the sum/sum total should be ~54.5%, not the 75% a naive average would give).
