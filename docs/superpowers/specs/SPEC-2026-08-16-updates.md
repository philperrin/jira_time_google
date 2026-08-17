# Spec: Worklog Year Selector, Schedule Start-Time + Daily Cap, Manual Timecard Entry, Assignments Expand/Collapse

Covers TODO.md items **#2**, **#3**, **#17**, plus a new small addition (Assignments Expand All / Collapse All).

---

## 1. Worklog year selector (TODO #2)

**Goal:** Let the user view worklogs for a prior year instead of always the current year.

**Files:** `TabWorklog.html`, `JavaScript.html`, `Code.js`

**UI:**
- Add a `<select id="wl-year">` to the Worklog toolbar, before the "Pull Worklog" button.
- Populated client-side with the current year down to 3 years prior (4 options total), current year selected by default. No server call needed to populate it.

**Server (`Code.js`):**
- `getWorklogs(year)` takes an explicit `year` argument instead of hardcoding `new Date().getFullYear()`.
- `fromDate` = `${year}-01-01`; add a `toDate` = `${year}-12-31` and extend the JQL to `worklogDate >= fromDate AND worklogDate <= toDate` so a past year doesn't pull in later worklogs.
- No change to pagination/fetch-all logic.

**Client (`JavaScript.html`):**
- `loadWorklog()` reads `#wl-year` value and passes it to `getWorklogs(year)`.
- CSV filename uses the selected year instead of `new Date().getFullYear()`.

**Edge cases:**
- Selecting a future-looking year isn't possible (list only goes current → current-3).
- Empty result set for a year with no worklogs — existing "0 entries" handling already covers this.

**Testing:** Manually pull worklogs for current year and one prior year; confirm CSV export filename matches the selected year and pivot totals differ appropriately.

---

## 2. Schedule Calendar Events: start-time control + daily cap (TODO #3)

**Goal:** Let the user choose when scheduled work begins, and prevent long totals from spilling through nights/weekends.

**Files:** `Code.js`, `TabAssignments.html`, `JavaScript.html`

**UI (`TabAssignments.html`):**
- Add `date` and `time` inputs to `#schedule-toolbar`, before the "Schedule Calendar Events" button: `#schedule-date` (default: tomorrow) and `#schedule-time` (default: `09:00`).

**Client (`JavaScript.html` → `scheduleEvents()`):**
- Read `#schedule-date` + `#schedule-time`, combine into an ISO datetime string, pass to `scheduleCalendarEvents(toSchedule, startDateTimeIso)`.
- Initialize the date/time inputs on `DOMContentLoaded` similar to the existing Timecard Monday/Friday init.

**Server (`Code.js` → `scheduleCalendarEvents`):**
- Replace `now + 1 hour` with the passed-in start datetime as the initial cursor.
- **Daily cap logic (fixed workday window, confirmed):** the work window is 9:00 AM–5:00 PM (8h/day).
  - Day 1 is seeded at the user-chosen start time, even if outside 9–5 (e.g. a 7:00 PM start is honored for whatever fits before midnight isn't a concern — cap is against the 5:00 PM boundary only if the start is before it; if the chosen start is already at/after 5:00 PM, treat the *remaining capacity for that calendar day as 0* and roll everything to the next work day at 9:00 AM).
  - While scheduling a project's block, if the block would cross 5:00 PM, split it: fill to 5:00 PM, carry the remainder to 9:00 AM the next work day.
  - "Next work day" skips Saturday/Sunday, rolling to Monday.
  - This means one entry's hours can span multiple events/days if it doesn't fit in the remaining capacity of a day — each split segment is created as its own calendar event so the visible blocks never cross the 5 PM boundary.
- Returns `{ created, startTime }` as today, where `created` counts total events actually created (may be more than `toSchedule.length` if entries were split across days).

**Edge cases:**
- An entry longer than 8h alone spans multiple full days automatically via the same splitting logic.
- Chosen start date in the past: not blocked client-side (no explicit requirement), but the date input has no `min` — acceptable since this mirrors how Google Calendar itself allows past events.

**Testing:** Schedule ~10h total starting at 2:00 PM on a Friday; confirm it fills to 5:00 PM Friday, then resumes 9:00 AM Monday (skipping the weekend) for the remainder.

---

## 3. Manual Timecard entry (TODO #17)

**Goal:** Let the user add a time entry to a specific day that didn't come from a calendar import.

**Files:** `TabTimecard.html` (no structural change needed — day tables are built dynamically), `JavaScript.html`

**UI / Client (`JavaScript.html`):**
- In `renderTimecardTable()`, below each day's `<table>`, add a `<button class="btn btn-secondary" onclick="addManualEntry('${date}')">+ Add Entry</button>`.
- `addManualEntry(date)` appends a row to that day's `<tbody>` with:
  - Start time input (`<input type="time">`)
  - End time input (`<input type="time">`)
  - Project `<select>` populated from `getAllocation()` rows (projectKey/projectName, same source as the Config allocation table)
  - Jira Issue `<select>` — empty until a Project is chosen; on Project change, populate from `window.issueOptions` filtered by the selected `projectKey` (same filtering already used for imported rows)
  - Duration (h) — **read-only, auto-calculated** from Start/End (confirmed): recompute on every Start/End `change` event as `(end - start) / 3600000` hours, floored at 0, and re-run `updateSummaryTable()`
  - Remove button (✕) — confirmed: deletes the row and re-runs `updateSummaryTable()`
- Manually-added rows use a synthetic entry pushed into `window.timecardEvents` (title: `'Manual entry'`, `description: ''`) so they flow through the exact same `.tc-issue-select` / `data-idx` wiring, `updateSummaryTable()`, and `submitTimecard()` logic as imported rows — no special-casing needed in submit.
- New rows get an `idx` = `window.timecardEvents.length` at creation time (appended, not spliced), so existing indices for other rows stay valid; on removal, only that row's DOM element and its `timecardEvents` entry are cleared (set to `null` in place, not spliced) so other rows' indices remain stable.

**Validation:**
- Duration must be > 0 (End after Start) before the row counts toward the day total or submits — same `durationHours > 0` filter `submitTimecard()` already applies.
- If End < Start, duration shows `0` (row visually shows "—" like unassigned rows) rather than a negative number.

**Edge cases:**
- A day with zero imported events (e.g., user wants to log an entry on a day nothing was imported for): out of scope for this pass — "Add Entry" only appears below days that already have a table, i.e. days present in the imported date range. If needed later, this can be revisited.

**Testing:** Import a week, add a manual entry to one day, pick project + issue, confirm it appears in the day total, the pivot summary, and gets submitted in `submitTimecard()` alongside imported rows.

---

## 4. Assignments: Expand All / Collapse All buttons (new)

**Goal:** Quickly expand or collapse every project group at once, instead of clicking each `▶`/`▼` heading individually.

**Files:** `TabAssignments.html`, `JavaScript.html`

**UI (`TabAssignments.html`):**
- Add two buttons to the existing toolbar (`#assignments-load-btn`'s toolbar), placed next to the density buttons: `<button class="btn btn-secondary" id="expand-all-btn" onclick="setAllGroupsCollapsed(false)" style="display:none;">Expand All</button>` and `<button class="btn btn-secondary" id="collapse-all-btn" onclick="setAllGroupsCollapsed(true)" style="display:none;">Collapse All</button>`.
- Hidden (`display:none`) until assignments are loaded — same visibility pattern as `#density-comfortable-btn`/`#density-compact-btn`, shown alongside `#assignments-container` in `renderAssignments()`.

**Client (`JavaScript.html`):**
- `setAllGroupsCollapsed(collapsed)`: iterates `document.querySelectorAll('.project-group')`, toggling the `collapsed` class and each heading's `.group-toggle` glyph (▶/▼) to match — reusing the same class/glyph convention as `toggleGroup()`.
- `renderAssignments()` sets `#expand-all-btn` and `#collapse-all-btn` to `style.display = ''` alongside the existing `container.style.display = ''` / `schedule-toolbar` reveal.

**Interaction with TODO #13 (persist collapsed state):** not implemented here — out of scope for this pass. If #13 is picked up later, `setAllGroupsCollapsed` should also update whatever collapsed-state store that item introduces.

**Testing:** Load assignments, click "Collapse All" (all groups collapse, arrows flip to ▶), click "Expand All" (all reopen, arrows flip to ▼). Buttons stay hidden before Load Assignments is clicked.

---

## Summary of decisions confirmed with user

- Schedule daily cap uses a **fixed 9 AM–5 PM workday window**, splitting any block that would cross 5 PM.
- Day rollover **skips weekends**, resuming Monday 9 AM.
- Manual Timecard rows are **removable** and **auto-calculate duration from Start/End** (not manually typed).
