# Jira Time Tracking — Google Apps Script

A Google Sheets add-on that connects to the Jira REST API to streamline weekly time logging. It pulls active issues, imports matching Google Calendar events, and posts time entries as Jira worklogs — all from within the spreadsheet.

---

## Setup

1. Open the Google Sheet and go to **Extensions > Apps Script**.
2. Paste in the contents of `Code.js` into the `Code.gs` file and create a new HTML file in the project named `CreateNewJira` and paste in the contents of `CreateNewJira.html` file. Create a new Script file named `tests` and paste in the contents of `tests.js` file.
3. Open the **Log Jira Time > Config > Add/Update Jira Base URL** menu item and enter your Atlassian instance URL (e.g. `https://your-company.atlassian.net`).
4. Open the **Log Jira Time > Config > Add/Update Jira API** menu item and enter your Atlassian API key.
   - Generate a key at: `https://id.atlassian.com/manage-profile/security/api-tokens`
5. Reload the sheet. The **Log Jira Time** menu will appear in the toolbar.
6. You may want to familiarize yourself with the Google Apps Script Project Settings to validate the appsscript.json settings, etc.

---

## Sheets

| Sheet | Purpose |
|---|---|
| **Time Card** | Weekly time entries. Column A holds dates; columns B–D hold issue, start time, and duration. |
| **Assignments** | Populated by "Populate assignments". Lists active Jira issues with project, status, and hours logged. |
| **Calendar** | Populated by "Import calendar events". Lists Google Calendar events for the Time Card date range. |
| **Allocation** | Maps Google Calendar event colors to Jira project keys. Used to auto-populate issue dropdowns in Calendar. |
| **History** | Archive of all submitted time entries (date, issue key, duration). |

---

## Menu: Log Jira Time

### Populate assignments
Fetches all active Jira issues where you are the assignee or watcher (Story, Task, Sub-task). Filters out archived and internal projects, then writes them to the Assignments sheet with hours logged. Also creates named ranges per project used by Calendar dropdowns.

### Import calendar events
Reads the date range from the Time Card sheet and pulls matching events from your Google Calendar. Skips non-billable Personal and Internal events. Populates the Calendar sheet and sets per-row issue dropdowns in column H based on each event's calendar color (mapped via the Allocation sheet).

After setting the dropdowns, each row in column H receives a Sheets AI formula (`=ai(...)`) that automatically suggests the best matching Jira issue for that event. The formula passes the event name (column A) and description (column E) as context, along with the full list of Jira issues available for that event's project (resolved via `UNIQUE(INDIRECT(...))` from the named range in column G). Google Sheets evaluates this formula using its built-in AI feature to return the most relevant issue from the dropdown list. The suggestion can be accepted as-is or overwritten manually before sending time to Jira.

### Send time to Jira
Posts each Time Card entry as a Jira worklog via the REST API. Converts local time to UTC before posting. Successful entries are written to the History sheet in one batch. After sending, clears the Time Card, resets Calendar dropdowns, and advances cell A2 to the next Monday.

### Config > Add/Update Jira Base URL
Prompts for your Atlassian instance base URL (e.g. `https://your-company.atlassian.net`) and saves it to Apps Script user properties (scoped to your account only).

### Config > Add/Update Jira API
Prompts for your Atlassian API key and saves it to Apps Script user properties (scoped to your account only).

### Config > Create new Jira issue
Opens a modal dialog to create a new Jira Task. On submit, calls the Jira REST API and refreshes the Assignments sheet automatically.

---

## Files

| File | Purpose |
|---|---|
| `Code.js` | All Apps Script server-side logic |
| `CreateNewJira.html` | HTML/JS for the Create New Jira modal dialog |
| `appsscript.json` | Apps Script manifest (timezone, permissions, runtime) |
| `tests.js` | Unit/integration tests, runnable via Config > Run tests |

---

## Business Logic

The following values are hardcoded in `Code.js` and reflect intentional business rules. If your setup changes, these are the places to update.

| Location | Value | Purpose |
|---|---|---|
| `retrieveJiraIssues` | `issuetype IN (Story, Task, Sub-Task)` | Only these issue types are fetched from Jira. Epics, Bugs, and other types are excluded. |
| `retrieveJiraIssues` | `status=Done AND updated>=-7d` | Done issues are included only if updated within the last 7 days, so recently closed work stays available for logging. |
| `retrieveJiraIssues` | `maxResults=100` | Maximum issues returned per API page. Pagination is handled automatically, but each page is capped at 100 (the Jira API maximum). |
| `retrieveJiraIssues` | `'Archive'`, `'Managed Services Internal'` | Projects whose names contain either string are excluded from the Assignments sheet. Update these strings if project naming conventions change. |
| `scheduleCalendarEvents` | `ROWS_PER_DAY = 21` | The Time Card sheet allocates 21 rows per day. This drives where each day's query formula is placed after time is sent. There are actually only 20 rows of time card data per day - one extra row is used for spacing and has a sum of the hours for that day. |
| `scheduleCalendarEvents` | `DAYS_PER_WEEK = 5` | Five days (Mon–Fri) are supported per week. Changing this would also require adjusting the Time Card sheet layout. |
| `getWorklogTotals_` | `maxResults=100` | Maximum worklogs fetched per page per issue. Pagination is handled automatically. |

---

## Notes

- Worklog fetches use `UrlFetchApp.fetchAll()` to parallelize requests across all issues, which significantly reduces load time when many issues are present.
- The script timezone is set to `America/Denver` in `appsscript.json`. All times are converted to UTC before being posted to Jira.
