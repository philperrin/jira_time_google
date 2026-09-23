# Spec: Per-Row Calendar Scheduling on the Assignments Tab

Replaces the bulk "enter hours for any number of issues, then pick one shared date/time and click one Schedule Calendar Events button" flow with an independent date/time/duration/Schedule control on every issue row.

---

## 1. Deactivate the current bulk scheduling UI

**Files:** `TabAssignments.html`, `JavaScript.html`

- Remove the bottom `#schedule-toolbar` block (the shared `#schedule-date`, `#schedule-time` inputs and `#schedule-btn` button) and `#schedule-status` from `TabAssignments.html`.
- Remove the following from `JavaScript.html`:
  - `initScheduleDate()` (seeds the old shared date field)
  - `scheduleEvents()` (reads the shared inputs + all `.schedule-input` values, builds `toSchedule[]`, calls `scheduleCalendarEvents`)
- The `.schedule-input` (hours) column in `renderAssignments()`'s per-row `<tr>` template is removed — see section 2 for its replacement.

## 2. New per-row scheduling controls

**Goal:** Each issue row schedules its own calendar event(s) independently, with no shared bulk state.

**Files:** `JavaScript.html` (`renderAssignments()`)

**UI — per row, replacing the old single "Schedule" `<input type="number">` cell:**
- **Date** input (`type="date"`) — default: today.
- **Time** input (`type="time"`) — default: the next full hour from now (e.g. 2:37 PM → 3:00 PM; 11:58 PM → 12:00 AM next... no day-rollover needed here, just the `:00` boundary of the current or following hour).
- **Duration** input (`type="number" step="0.25" min="0"`) — no default value (starts empty).
- **Schedule** button — scoped to that row only.
- An inline text span next to the button for feedback (empty by default).

Each row's three inputs and button carry `data-idx="${idx}"` (same `idx` used elsewhere to index into `window.loadedIssues`) so the click handler can locate them via `closest('tr')` without any shared/global scheduling state.

**Column layout:** the existing table header row (`Key`, `Name`, `Status`, `Time Logged`, `Schedule`) keeps its 5 columns; the "Schedule" header's cell now contains the date/time/duration inputs + button + feedback span stacked or inlined (implementation's choice for spacing — functionally these are 4 controls in place of the old single hours input).

## 3. Row Schedule button behavior

**Files:** `JavaScript.html`

Replaces `scheduleEvents()` with a per-row handler, e.g. `scheduleRow(button)`:

- Reads that row's date, time, and duration inputs only (via `closest('tr')`).
- **Validation:** duration must be a number > 0 and date must be non-empty. If invalid, show the inline feedback span as "Error!" and stop — no server call.
- Looks up the row's issue from `window.loadedIssues[idx]` (same lookup `scheduleEvents()` used).
- Disables the button and shows a loading label (matching `setLoading()`'s pattern, scoped to this button rather than a global id) while the call is pending.
- Calls the existing, **unchanged** server function:
  ```js
  scheduleCalendarEvents(
    [{ key: issue.key, jiraProject: issue.projectKey, dropdownValue: issue.dropdownValue, hours: duration }],
    `${date}T${time}:00`
  )
  ```
- **On success:** re-enables the button, sets the inline feedback span to "Success!", and clears the duration input back to empty. Date and time inputs keep their current values (so the same row can be scheduled again for a different chunk of hours without re-entering date/time).
- **On failure:** re-enables the button, sets the inline feedback span to "Error!" (the underlying `e.message` may be attached as a `title` tooltip on the span for debugging, but no separate status area is used).

## 4. Server-side (`Code.js`)

**No changes.** `scheduleCalendarEvents(toSchedule, startDateTimeIso)` already:
- Accepts an array of `{ key, jiraProject, dropdownValue, hours }` entries — a single-entry array works unmodified.
- Applies the existing 9 AM–5 PM / 8h-per-day cap and next-work-day rollover (skipping weekends) to each entry independently, so a single row's duration that exceeds the remaining capacity for its chosen day still splits correctly across subsequent work days.
- Returns `{ created, startTime }`, used as-is for the per-row success feedback (though the UI only shows "Success!", not the count, per the design decision below).

---

## Edge cases

- Duration left empty or ≤ 0: blocked client-side before any server call, with inline "Error!".
- Date left empty: blocked client-side, with inline "Error!".
- Two different rows scheduled back-to-back with independent dates/times: fully independent server calls, no shared cursor or ordering between rows (unlike the old bulk flow, which scheduled all entered rows sequentially from one shared start time).
- A single row's duration long enough to spill into a following day (or across a weekend): still handled by the existing `scheduleCalendarEvents` splitting logic, unchanged from today.

## Testing

- Load Assignments, confirm the old bottom toolbar and shared hours column are gone, and each row shows date (defaulted to today) / time (defaulted to next full hour) / duration (empty) / Schedule button / feedback span.
- Enter 2.5h on one row's duration, leave date/time at defaults, click Schedule: confirm a single calendar event is created, the row shows "Success!", and its duration input clears while date/time remain.
- Enter 10h on a row starting at 2:00 PM: confirm events split at 5:00 PM and resume 9:00 AM the next work day (weekend-skipping), matching current `scheduleCalendarEvents` behavior.
- Click Schedule with duration empty or 0: confirm inline "Error!" and no server call / no calendar event created.
- Schedule two different rows in sequence with different chosen dates: confirm each creates events independently at its own date/time, with no interaction between the two.

---

## Summary of decisions confirmed with user

- Per-row feedback is a simple inline "Success!" / "Error!" text next to that row's Schedule button — no shared status bar.
- On success, the duration input clears to empty; date and time inputs retain their values for quick re-scheduling of the same issue.
- The existing 8h/day cap + next-work-day rollover logic in `scheduleCalendarEvents` is reused unchanged — this per-row change only affects how the client invokes it (one entry at a time, per row, instead of a shared bulk list).
