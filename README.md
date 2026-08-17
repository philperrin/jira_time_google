# Jira Time Tracker — Standalone Web App

A standalone Google Apps Script web app that connects to the Jira REST API to streamline weekly time logging. It pulls active issues, imports matching Google Calendar events, and posts time entries as Jira worklogs — all from a browser-based interface.

---

## Deployment

Deployment is done manually through the Apps Script editor (no `clasp` CLI, per company policy).

### Create a new standalone project

1. Go to [script.google.com](https://script.google.com) and click **New project**.
2. Rename the project (e.g. "Jira Time Tracker").

### Copy in the project files

For each `.js` and `.html` file in this repo, create a matching file in the Apps Script editor (**File > New > Script** or **HTML**) and paste in the contents:

- `Code.js`, `Debug.js` as script files
- `Index.html`, `JavaScript.html`, `Stylesheet.html`, `ChangelogData.html`, and every `Tab*.html` file as HTML files

Update the project's manifest (**Project Settings > Show "appsscript.json" manifest file in editor**, then edit `appsscript.json`) to match the `appsscript.json` in this repo (timezone, OAuth scopes, web app config).

### Deploy as a web app

1. Click **Deploy > New deployment**
2. Select type: **Web app**
3. Set **Execute as**: Me
4. Set **Who has access**: Only myself (or your org as needed)
5. Click **Deploy** and copy the web app URL

To push future changes, edit the files directly in the Apps Script editor and create a new deployment (or use **Manage deployments > Edit** to update an existing one).

---

## First-Time Setup

After deploying, open the web app URL and complete setup in this order:

1. **Config tab** — Enter your Jira base URL (e.g. `https://your-company.atlassian.net`) and your Atlassian API key, then set your project allocation rows (calendar color, project key/name) and save.
   - Generate an API key at: `https://id.atlassian.com/manage-profile/security/api-tokens`
2. **Assignments tab** — Load your active Jira issues, grouped by project. Groups can be expanded/collapsed individually or all at once, and rows include a schedule-hours input for creating calendar events.
3. **Timecard tab** — Set a date range, import calendar events (matched to projects via allocation color), assign Jira issues per entry, add manual entries for untracked time, and submit worklogs.
4. **Create Issue tab** — Create new Jira tasks directly from the app.
5. **Worklog tab** — Pull and review worklogs for a selected year (current year or the 3 prior), view the month × project pivot, and export to CSV.
6. **Changelog tab** — Browse known future enhancements, unselected/deferred ideas, and already-completed improvements.

---

## Files

| File | Purpose |
|---|---|
| `Code.js` | All Apps Script server-side logic |
| `Debug.js` | Developer-only logging/validation helpers, run manually from the Apps Script editor |
| `Index.html` | Web app shell (tab navigation, loads all other templates) |
| `JavaScript.html` | Client-side JS (shared across tabs) |
| `Stylesheet.html` | Shared CSS |
| `ChangelogData.html` | Changelog tab content data (`CHANGELOG_ITEMS`), included as its own template |
| `TabConfig.html` | Config tab UI |
| `TabAssignments.html` | Assignments tab UI |
| `TabTimecard.html` | Timecard tab UI |
| `TabCreateIssue.html` | Create Issue tab UI |
| `TabWorklog.html` | Worklog tab UI |
| `TabChangelog.html` | Changelog tab UI |
| `appsscript.json` | Apps Script manifest (timezone, permissions, runtime) |

---

## Business Logic

The following values are hardcoded in `Code.js` and reflect intentional business rules.

| Location | Value | Purpose |
|---|---|---|
| `getJiraIssues` | `issuetype IN (Story, Task, Sub-Task)` | Only these issue types are fetched. Epics, Bugs, and other types are excluded. |
| `getJiraIssues` | `status=Done AND updated>=-7d` | Done issues are included only if updated within the last 7 days. |
| `getJiraIssues` | `maxResults=100` | Maximum issues per API page; pagination is handled automatically. |
| `getJiraIssues` | `'Archive'`, `'Managed Services Internal'` | Projects whose names contain either string are excluded. |
| `getWorklogTotals_` | `maxResults=100` | Maximum worklogs fetched per page per issue; pagination is handled automatically. |
| `getWorklogs` | `year` (defaults to current year) | Worklogs are fetched for a single selected calendar year (`Jan 1`–`Dec 31`); the Worklog tab lets the user pick the current year or one of the 3 prior years. |
| `scheduleCalendarEvents` | 9:00 AM–5:00 PM workday, skips Sat/Sun | Scheduled calendar events are capped at 8h/day; work that doesn't fit is split and rolled to the next work day. |

---

## Notes

- All configuration (Jira URL, API key, allocation) is stored in Apps Script user properties scoped to your Google account — no data is shared with other users of the same deployment.
- Worklog fetches use `UrlFetchApp.fetchAll()` to parallelize requests, significantly reducing load time when many issues are present.
- The script timezone is set in `appsscript.json` (`America/Denver`). All times are converted to UTC before being posted to Jira.
- Known bugs, deferred ideas, and completed improvements are tracked in-app on the **Changelog** tab (`ChangelogData.html`) rather than in a separate backlog file.
