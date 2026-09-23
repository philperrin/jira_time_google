# Spec: Pre-Deployment Test Suite

Adds an automated, GAS-native test suite for `Code.js`'s server-side business logic, plus a manual smoke-test checklist, to be run before every deployment. No build tooling, no npm, no CI — consistent with this project's existing manual, editor-based development workflow.

---

## 1. Context and constraints

This is a plain Apps Script project (per `CLAUDE.md`): no `clasp`, no package manager, no build system, no existing test runner. `Debug.js` already establishes the relevant precedent — a diagnostic function (`validateImportCalendarEvents`) meant to be run manually from the Apps Script editor's function picker, with output read from the execution log.

Most of the logic worth protecting against regressions lives in `Code.js`, interleaved with calls to GAS's built-in services (`UrlFetchApp`, `CalendarApp`, `PropertiesService`, `Session`), which only exist inside the Apps Script runtime. That rules out a local Node/Jest suite without either duplicating logic outside of what's actually deployed, or contradicting `CLAUDE.md`'s explicit "not a Node project" framing.

**Decision:** build the suite as a new `Tests.js` file, run manually from the Apps Script editor, using hand-rolled assertions and temporary global-service overrides to mock `UrlFetchApp`/`CalendarApp`/`PropertiesService`/`Session` without touching real Jira or Calendar data. Scope is `Code.js` only — client-side logic in `JavaScript.html` is out of scope for this pass.

## 2. Test runner architecture

**New file:** `Tests.js`.

A minimal runner, no external framework:

```js
function runAllTests() {
  const results = [];
  TEST_CASES.forEach(({ name, fn }) => {
    try {
      fn();
      results.push({ name, pass: true });
    } catch (e) {
      results.push({ name, pass: false, error: e.message });
    }
  });
  const passed = results.filter(r => r.pass).length;
  Logger.log(`${passed}/${results.length} passed`);
  results.filter(r => !r.pass).forEach(r => Logger.log(`FAIL: ${r.name} — ${r.error}`));
}

function test(name, fn) { TEST_CASES.push({ name, fn }); }
const TEST_CASES = [];

function assertEquals(actual, expected, msg) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`${msg || ''} expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  }
}
function assertTrue(cond, msg) { if (!cond) throw new Error(msg || 'expected true'); }
function assertThrows(fn, msg) {
  try { fn(); throw new Error(msg || 'expected a throw, none occurred'); }
  catch (e) { if (e.message === (msg || 'expected a throw, none occurred')) throw e; }
}
```

Test cases are registered at file-load time via `test(name, fn)` calls, grouped in `Tests.js` under the same numbered section comments `Code.js` uses (config, issues, scheduling, calendar import, worklogs, allocation, utilization, create issue). Running it is identical to the existing `Debug.js` workflow: open the Apps Script editor, select `runAllTests` in the function picker, run, read the summary and any `FAIL:` lines from the execution log.

## 3. Mocking approach

A small set of fake GAS services defined in `Tests.js`, installed and torn down per test via global reassignment (no refactor of `Code.js`):

```js
function mockResponse_(code, body) {
  return { getResponseCode: () => code, getContentText: () => body };
}

function withMockUrlFetch(handlers, fn) {
  const real = UrlFetchApp;
  UrlFetchApp = {
    fetch: handlers.fetch || (() => mockResponse_(200, '{}')),
    fetchAll: handlers.fetchAll || (requests => requests.map(() => mockResponse_(200, '{}')))
  };
  try { fn(); } finally { UrlFetchApp = real; }
}

function withMockCalendar(fakeCalendar, fn) {
  const real = CalendarApp;
  CalendarApp = {
    getDefaultCalendar: () => fakeCalendar,
    getCalendarById: () => fakeCalendar,
    EventColor: real.EventColor,
    GuestStatus: real.GuestStatus
  };
  try { fn(); } finally { CalendarApp = real; }
}

function withMockProperties(initial, fn) {
  const store = Object.assign({}, initial);
  const real = PropertiesService;
  PropertiesService = {
    getUserProperties: () => ({
      getProperty: k => (k in store ? store[k] : null),
      setProperty: (k, v) => { store[k] = v; }
    })
  };
  try { fn(); } finally { PropertiesService = real; }
}

function withMockSession(email, tz, fn) {
  const real = Session;
  Session = {
    getActiveUser: () => ({ getEmail: () => email }),
    getScriptTimeZone: () => tz || real.getScriptTimeZone()
  };
  try { fn(); } finally { Session = real; }
}
```

Each `withMock*` guarantees restoration of the real service even if the wrapped test throws, so a failing test can never leak a mock into a later test or a real run. Tests exercising pure logic (no GAS service dependency — e.g. `parseTimeSpentHours_`, `nextScheduleWorkDay_`) call the function directly with no mocking.

## 4. Test case inventory

Organized by `Code.js`'s existing numbered section comments.

**Config (section 00)**
- `saveJiraUrl` trims whitespace and strips a trailing slash; `getConfig` reflects the saved URL and reports `hasApiKey` as `true`/`false` correctly.
- `getAllocation`/`saveAllocation` round-trip through mocked `PropertiesService` (empty → `[]`; saved rows come back identical).
- `testJiraConnection`: returns `{ok:false, error}` when URL or key is blank (no fetch attempted); returns `{ok:true, displayName}` on a mocked 200 response; returns `{ok:false, error}` on a mocked 4xx/5xx response; never throws even if the mocked `fetch` itself throws.

**Jira issues**
- `getJiraIssues` throws when URL/API key are unset.
- Filters out issues whose project name contains `'Archive'` or `'Managed Services Internal'` (the hardcoded business rule `CLAUDE.md` calls out — must not regress silently).
- Paginates via `nextPageToken`: a mock returning two pages results in both being concatenated.
- `timeLogged` is computed correctly from mocked worklog-total seconds → hours, rounded to 2 decimals.
- `markIssueDone`: finds and posts the transition whose `to.name` is case-insensitively `"done"`; throws a clear error when no such transition exists.

**Scheduling**
- `nextScheduleWorkDay_`: from a Friday, returns the following Monday at 9:00 AM (skips Sat/Sun); from a Tuesday, returns Wednesday 9:00 AM.
- `scheduleDayCapacityMs_`: returns correct remaining ms before 5:00 PM; returns 0 when the cursor is already past 5:00 PM.
- `scheduleCalendarEvents`: an entry under 8h creates one event on the given day; an entry over 8h splits across multiple work days, correctly skipping a weekend in between; event title matches `"Client Task Time: <Project Name> - <Jira Key>"`; `colorHex` → `CalendarApp.EventColor` mapping is applied, verified via the fake calendar's `createEvent` call arguments.

**Calendar import**
- `importCalendarEvents`: events with `GuestStatus.NO` are excluded.
- Events whose color doesn't map to any active (non-ignored) allocation row are excluded.
- Events matching an active row's color get the correct `projectKey`; `duration` is computed in quarter-hour increments.
- Rows with `ignore: true` never match, even if their color matches an event.

**Worklogs / sendTimeEntries**
- `parseTimeSpentHours_`: `"1h 30m"` → 1.5, `"45m"` → 0.75, `"2h"` → 2, `""`/`undefined` → 0.
- `getWorklogs`: includes only entries whose `author.emailAddress` matches the current user; excludes entries outside `[Jan 1, Dec 31]` of the requested year; paginates extra worklog pages when an issue has more than 100 entries.
- `sendTimeEntries`: correctly parses `"hh:mm a"` start times, including AM/PM edge cases (12:00 AM → hour 0, 12:00 PM → hour 12), and formats them to the UTC string Jira expects; skips entries with no `issueKey`; correctly counts `succeeded`/`failed` against mocked 201 vs. non-201 responses; catches a thrown fetch error for one entry and records it in `errors` without aborting the rest of the batch.

**Allocation tab**
- `getAllocationValues_`: returns `{}` for a year never saved; returns the saved sparse grid for a saved year.
- `getAllocationTabData`: project list is the union of that year's worklog project keys, alphabetically sorted, minus any project marked `ignore: true` in `getAllocation()`.
- `saveAllocationGrid`: saving one year's grid doesn't overwrite another year's previously saved values.

**Utilization tab**
- `getPayPeriodEndDates_`: given an anchor date, returns correctly 14-day-spaced period-end dates within a target year, including one period whose end falls just outside the year boundary (excluded) and one just inside (included).
- `getPayPeriodSummary_`: `payPeriodsToDate` counts only periods `<= today`; returns a zeroed result when no anchor is set.
- `calculateUtilization`: a cell is `null` (not a divide-by-zero) when a project has no allocation value for that month; row/col/grand totals are computed as `sum(hours)/sum(allocation)` across contributing cells, not an average of per-cell percentages (verified with an unequal-weight fixture); `unratedProjects` correctly lists projects with no configured rate and excludes their hours from `revenue`/`grossRevenue`.
- `getOrCollectWorklogs`: returns cached rows without calling `getWorklogs` when a cache entry exists and `forceRefresh` is falsy; bypasses the cache and re-fetches when `forceRefresh` is true.

**Create Issue**
- `makeJira`: builds the correct payload shape (`fields.project.key`, `issuetype.name`, `priority.id`, custom fields) from `formData`; includes `assignee.accountId` only when the mocked user-search resolves a match; returns a friendly error string (not a throw) on a non-201 response or a failed user search.

## 5. Manual smoke-test checklist

Some behavior can't be exercised from a GAS-editor-run test file — real OAuth/API-key flow, real Calendar UI, actual Jira connectivity. Run this checklist against the deployed web app before/after each release, alongside `runAllTests`:

1. **Config tab** — save a Jira URL + API key, reload the page, confirm they persist; click "Test Connection" against real Jira and confirm success.
2. **Assignments tab** — load real issues; confirm the Story/Task/Sub-Task-only filter and the Archive/Managed-Services-Internal exclusion hold against live data.
3. **Timecard tab** — import a real week of calendar events; confirm color→project matching works against your actual calendar color scheme; assign an issue and post one real worklog to a test/sandbox issue; confirm it appears in Jira.
4. **Create Issue tab** — create one real issue in a test project; confirm it appears in Jira with the correct fields and assignee.
5. **Worklog tab** — load the current year; confirm the month×project pivot and CSV export match what's actually in Jira.
6. **Allocation / Utilization tabs** — save an allocation grid; confirm the Utilization tab's percentages/revenue figures look sane against it.

## 6. Workflow integration

Add one line to `CLAUDE.md`'s Development workflow section: run `runAllTests` from the Apps Script editor's function picker and walk the manual checklist (section 5 above) before creating or updating a deployment. No CI is introduced — this remains a manual, pre-deploy gate, consistent with the rest of this project's workflow.
