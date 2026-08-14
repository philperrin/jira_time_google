# Jira Time Tracker — Improvement Backlog

Opportunities identified from a full-app UX and efficiency review.

---

## High Impact

### 1. Remove or defer `getWorklogTotals_` from Assignments load
**File:** `Code.js` → `getJiraIssues()`

Every "Load Assignments" click triggers a batch of worklog API calls — one per active issue — just to populate the "Time Logged (h)" column. On 30+ issues this is the most expensive operation in the app and makes the Assignments tab feel slow.

It's also wasted a second time: `loadTimecardEvents()` in `JavaScript.html` calls `getJiraIssues()` purely to populate the issue dropdowns (`window.issueOptions`), but the Timecard tab never displays `timeLogged`. Every "Import Events" click pays for the full worklog-totals batch fetch for no benefit.

Options:
- **Remove the column** from Assignments entirely (the same data is visible in the Worklog tab)
- **Defer it**: show assignments immediately with dashes, then populate time-logged asynchronously via a secondary button or lazy fetch
- **Split the function**: give Timecard a lightweight issue-list fetch (key, summary, project) that skips `getWorklogTotals_` entirely, since it doesn't need time-logged data

---

### 2. Add a year selector to the Worklog tab
**File:** `Code.js` → `getWorklogs()`, `TabWorklog.html`, `JavaScript.html`

`getWorklogs()` hardcodes the start date to January 1 of the current year. There is no way to view a previous year or a custom range.

Suggested fix: add a `<select>` to the Worklog toolbar populated with the current year and the 2–3 prior years. Pass the selected year to the server function instead of always using `new Date().getFullYear()`. Consider defaulting to a narrower range (e.g. current month/quarter) if full-year fetches prove slow for high-volume users.

---

### 3. Add a start-time control to Schedule Calendar Events
**File:** `Code.js` → `scheduleCalendarEvents()`, `TabAssignments.html`, `JavaScript.html`

Events are always scheduled starting at the next full hour from the current time. There is no way to specify a date or starting time, which makes the feature impractical for planning work in advance.

Suggested fix: add a date + time input to the schedule toolbar (defaulting to tomorrow 9:00 AM). Pass the chosen start datetime to `scheduleCalendarEvents()` instead of computing `now + 1 hour`.

Also address the multi-day overflow case: events are currently chained back-to-back with no cap, so a large total (e.g. 40 hours across several issues) creates one continuous block running through nights and weekends. The fix should cap hours scheduled per day (e.g. 8h) and roll remaining time to the next work day.

---

### 4. Confirm before irreversible actions
**Files:** `JavaScript.html` → `submitTimecard()`, `scheduleEvents()`

"Send to Jira" (posts real worklogs) and "Schedule Calendar Events" (creates real calendar events) both fire immediately on click with no review step. A misclick or a wrong hours entry creates real data in Jira/Calendar that must be cleaned up manually — there's no undo.

Suggested fix: show a brief confirmation summary ("Submit 6 worklog(s) totaling 24h to Jira?" / "Create 4 calendar event(s) starting Tue 9:00 AM?") before firing the `google.script.run` call.

---

## Medium Impact

### 5. Explain empty Jira Issue dropdowns on the Timecard tab
**File:** `JavaScript.html` → `renderTimecardTable()`

When a calendar event's color doesn't match any allocation entry, or the matched project has no active issues, the Jira Issue dropdown renders with only "— select —" and no explanation. Users have no way to know whether the config is wrong or the project simply has no open issues.

Suggested fix: after populating each dropdown, check whether any options were added. If none, insert a disabled option like `No issues found for [projectKey]` or add a warning indicator to the row.

---

### 6. Replace the Timecard polling loop with callback coordination
**File:** `JavaScript.html` → `loadTimecardEvents()` / `renderTimecardTable()`

The `setInterval` (300 ms × up to 20 tries) that waits for issue options to arrive is fragile. If both the calendar and issues calls are slow, the events table appears before the dropdowns are populated. Worse, if issues still haven't arrived after 20 attempts (6s), the loop gives up **silently** — dropdowns stay empty with no error shown to the user.

Suggested fix: use a shared state object; whichever of the two `google.script.run` calls returns second triggers dropdown population. Surface an error state if issue loading fails or times out.

---

### 7. Improve the Allocation config table usability
**File:** `TabConfig.html`

The table has two columns that require prior knowledge:
- **Calendar Color #** — requires knowing that e.g. Sage = 2, Flamingo = 4, Tangerine = 6, etc.
- **Color Enum Name** — requires knowing the exact GAS constant name (PALE_BLUE, ORANGE, etc.)
- **Hours/Week** — not currently used by any app feature

This is enough of a pain point that a developer-only debug tool (`Debug.js` → `validateImportCalendarEvents()`) exists specifically to trace color/project mismatches — a strong signal the mapping UI itself needs to be clearer for end users, not just diagnosable after the fact.

Suggested fix: add a small reference table or tooltip below the allocation table mapping color numbers to their names and enum values. Consider removing "Hours/Week" until it drives a feature.

---

### 8. Prevent silent data loss when saving allocation rows
**File:** `JavaScript.html` → `saveAllocationTable()`

Rows missing a `projectKey` are silently filtered out before saving (`.filter(r => r.projectKey)`). If a user fills in a color and project name but forgets the key, that row disappears on save with no warning.

Suggested fix: validate before saving and show a status message listing which rows were dropped and why, or block save until required fields are filled.

---

### 9. Add a way to review individual worklog entries on-screen
**File:** `TabWorklog.html`, `JavaScript.html`

The on-screen worklog detail table was removed in favor of the monthly pivot view. The only way to see individual worklog line items now is the CSV export — there's no way to spot-check data without downloading a file.

Suggested fix: add an optional expandable detail view (e.g. click a pivot cell to see the underlying entries for that project/month) or a toggle to show the raw list alongside the pivot.

---

### 10. Add a "Test Connection" action to the Config tab
**File:** `TabConfig.html`, `JavaScript.html`, `Code.js`

There is no way to validate the Jira URL and API key until the user visits another tab and gets an error. Typos in the URL or an expired/invalid API key aren't caught at the point of entry.

Suggested fix: add a "Test Connection" button that calls a lightweight Jira endpoint (e.g. `/rest/api/3/myself`) and reports success/failure inline.

---

## Low Impact / Polish

### 11. Add a "Clear" button to the Timecard submit toolbar
**File:** `TabTimecard.html`, `JavaScript.html`

After a successful "Send to Jira," the dropdowns reset but the event table stays visible. There is no way to dismiss the table without clicking "Import Events" again (which re-fetches from the calendar).

Suggested fix: add a "Clear" button to `#tc-submit-toolbar` that hides `#tc-table` and `#tc-submit-toolbar` and resets `window.timecardEvents`.

---

### 12. Consider reordering tabs to match daily workflow
**File:** `Index.html`

Current order: Assignments → Timecard → Create Issue → Worklog → Config

The daily workflow is typically: import calendar events (Timecard) → review assignments (Assignments) → review hours (Worklog). "Create Issue" and "Config" are infrequent.

Suggested order: **Timecard → Assignments → Worklog → Create Issue → Config**

---

### 13. Persist collapsed group state on Assignments
**File:** `JavaScript.html` → `renderAssignments()`

Collapsing project groups is reset every time "Load Assignments" is clicked, since the container is rebuilt from scratch. Users with many projects who prefer a curated view have to re-collapse groups on every reload.

Suggested fix: track collapsed project keys in a `Set` (or `localStorage`, similar to the density preference) and reapply after render.

---

### 14. Show a summary of skipped events on Timecard submit
**File:** `JavaScript.html` → `submitTimecard()`

Events left without an assigned issue are silently excluded from submission (`.filter(e => e.issueKey && e.durationHours > 0)`) with no indication of how many were skipped.

Suggested fix: include a skipped count in the result message, e.g. "Sent 12 worklog(s). 3 event(s) skipped (no issue selected)."

---

### 15. Apply the density toggle consistently across tabs
**File:** `Stylesheet.html`, `JavaScript.html` → `setDensity()`

The Comfortable/Compact density toggle only affects the Assignments tab, even though Timecard and Worklog tables have similarly dense layouts.

Suggested fix: apply the `density-*` class at a shared container level (e.g. `body` or `main`) so the preference applies app-wide.

---

### 16. Add accessible labels to icon-only buttons
**Files:** `JavaScript.html`, `TabTimecard.html`, `TabAssignments.html`

Buttons using emoji-only or icon-only content (`✕` remove-row, `⏱ Schedule Calendar Events`, `📥 Send to Jira`) lack `aria-label`/`title` attributes for screen readers.

Suggested fix: add `aria-label` or `title` attributes describing the action in plain text.

---

## Resolved / Superseded

### ~~Remove the redundant "Time Spent" column from the Worklog detail table~~
The on-screen worklog detail table (with both "Time Spent" and "Hours" columns) was removed entirely in favor of the pivot-only view (commit `126e13f`). The redundancy now only exists in the CSV export, where it's low-cost — no action needed unless the export format is revisited.

---

## Summary

| # | Area | File(s) | Effort |
|---|---|---|---|
| 1 | Remove/defer worklog totals from Assignments + Timecard loads | `Code.js` | Medium |
| 2 | Worklog year selector | `Code.js`, `TabWorklog.html`, `JavaScript.html` | Small |
| 3 | Schedule start-time picker + multi-day cap | `Code.js`, `TabAssignments.html`, `JavaScript.html` | Medium |
| 4 | Confirm before irreversible actions | `JavaScript.html` | Small |
| 5 | Explain empty Timecard dropdowns | `JavaScript.html` | Small |
| 6 | Replace Timecard polling loop | `JavaScript.html` | Medium |
| 7 | Allocation table usability | `TabConfig.html` | Small |
| 8 | Prevent silent allocation row loss | `JavaScript.html` | Small |
| 9 | On-screen worklog detail/expand view | `TabWorklog.html`, `JavaScript.html` | Medium |
| 10 | Config "Test Connection" action | `TabConfig.html`, `JavaScript.html`, `Code.js` | Small |
| 11 | "Clear" button on Timecard | `TabTimecard.html`, `JavaScript.html` | Trivial |
| 12 | Tab order | `Index.html` | Trivial |
| 13 | Persist collapsed group state | `JavaScript.html` | Small |
| 14 | Skipped-event summary on Timecard submit | `JavaScript.html` | Trivial |
| 15 | App-wide density toggle | `Stylesheet.html`, `JavaScript.html` | Small |
| 16 | Accessible labels for icon-only buttons | `JavaScript.html`, `TabTimecard.html`, `TabAssignments.html` | Trivial |
