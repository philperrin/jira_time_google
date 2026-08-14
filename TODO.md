# Jira Time Tracker — Improvement Backlog

Opportunities identified from a full-app UX and efficiency review.

---

## High Impact

### 1. Remove or defer `getWorklogTotals_` from Assignments load
**File:** `Code.js` → `getJiraIssues()`

Every "Load Assignments" click triggers a batch of worklog API calls — one per active issue — just to populate the "Time Logged (h)" column. On 30+ issues this is the most expensive operation in the app and makes the Assignments tab feel slow.

Options:
- **Remove the column** from Assignments entirely (the same data is visible in the Worklog tab)
- **Defer it**: show assignments immediately with dashes, then populate time-logged asynchronously via a secondary button or lazy fetch

---

### 2. Add a year selector to the Worklog tab
**File:** `Code.js` → `getWorklogs()`, `TabWorklog.html`, `JavaScript.html`

`getWorklogs()` hardcodes the start date to January 1 of the current year. There is no way to view a previous year or a custom range.

Suggested fix: add a `<select>` to the Worklog toolbar populated with the current year and the 2–3 prior years. Pass the selected year to the server function instead of always using `new Date().getFullYear()`.

---

### 3. Add a start-time control to Schedule Calendar Events
**File:** `Code.js` → `scheduleCalendarEvents()`, `TabAssignments.html`, `JavaScript.html`

Events are always scheduled starting at the next full hour from the current time. There is no way to specify a date or starting time, which makes the feature impractical for planning work in advance.

Suggested fix: add a date + time input to the schedule toolbar (defaulting to tomorrow 9:00 AM). Pass the chosen start datetime to `scheduleCalendarEvents()` instead of computing `now + 1 hour`.

---

## Medium Impact

### 4. Explain empty Jira Issue dropdowns on the Timecard tab
**File:** `JavaScript.html` → `renderTimecardTable()`

When a calendar event's color doesn't match any allocation entry, or the matched project has no active issues, the Jira Issue dropdown renders with only "— select —" and no explanation. Users have no way to know whether the config is wrong or the project simply has no open issues.

Suggested fix: after populating each dropdown, check whether any options were added. If none, insert a disabled option like `No issues found for [projectKey]` or add a warning indicator to the row.

---

### 5. Replace the Timecard polling loop with callback coordination
**File:** `JavaScript.html` → `loadTimecardEvents()` / `renderTimecardTable()`

The `setInterval` (300 ms × up to 20 tries) that waits for issue options to arrive is fragile. If both the calendar and issues calls are slow, the events table appears before the dropdowns are populated. A callback-coordination pattern — accumulate both results, then populate — would be more reliable.

Suggested fix: use a shared state object; whichever of the two `google.script.run` calls returns second triggers dropdown population.

---

### 6. Improve the Allocation config table usability
**File:** `TabConfig.html`

The table has two columns that require prior knowledge:
- **Calendar Color #** — requires knowing that e.g. Sage = 2, Flamingo = 4, Tangerine = 6, etc.
- **Color Enum Name** — requires knowing the exact GAS constant name (PALE_BLUE, ORANGE, etc.)
- **Hours/Week** — not currently used by any app feature

Suggested fix: add a small reference table or tooltip below the allocation table mapping color numbers to their names and enum values. Consider removing "Hours/Week" until it drives a feature.

---

## Low Impact / Polish

### 7. Remove the redundant "Time Spent" column from the Worklog detail table
**File:** `TabWorklog.html`, `JavaScript.html` → `loadWorklog()`

The "Time Spent" column (e.g., "1h 30m") and the "Hours" column (1.50) express the same value. Keeping only "Hours" reduces horizontal scroll and table width.

---

### 8. Add a "Clear" button to the Timecard submit toolbar
**File:** `TabTimecard.html`, `JavaScript.html`

After a successful "Send to Jira," the dropdowns reset but the event table stays visible. There is no way to dismiss the table without clicking "Import Events" again (which re-fetches from the calendar).

Suggested fix: add a "Clear" button to `#tc-submit-toolbar` that hides `#tc-table` and `#tc-submit-toolbar` and resets `window.timecardEvents`.

---

### 9. Consider reordering tabs to match daily workflow
**File:** `Index.html`

Current order: Assignments → Timecard → Create Issue → Worklog → Config

The daily workflow is typically: import calendar events (Timecard) → review assignments (Assignments) → review hours (Worklog). "Create Issue" and "Config" are infrequent.

Suggested order: **Timecard → Assignments → Worklog → Create Issue → Config**

---

## Summary

| # | Area | File(s) | Effort |
|---|---|---|---|
| 1 | Remove/defer worklog totals from Assignments load | `Code.js` | Medium |
| 2 | Worklog year selector | `Code.js`, `TabWorklog.html`, `JavaScript.html` | Small |
| 3 | Schedule start-time picker | `Code.js`, `TabAssignments.html`, `JavaScript.html` | Small |
| 4 | Explain empty Timecard dropdowns | `JavaScript.html` | Small |
| 5 | Replace Timecard polling loop | `JavaScript.html` | Medium |
| 6 | Allocation table usability | `TabConfig.html` | Small |
| 7 | Remove redundant "Time Spent" column | `TabWorklog.html`, `JavaScript.html` | Trivial |
| 8 | "Clear" button on Timecard | `TabTimecard.html`, `JavaScript.html` | Trivial |
| 9 | Tab order | `Index.html` | Trivial |
