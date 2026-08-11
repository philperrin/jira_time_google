# Jira Time Tracker — Standalone Web App

A standalone Google Apps Script web app that connects to the Jira REST API to streamline weekly time logging. It pulls active issues, imports matching Google Calendar events, and posts time entries as Jira worklogs — all from a browser-based interface.

---

## Deployment

### Prerequisites

Install the `clasp` CLI (requires Node.js):

```bash
npm install -g @google/clasp
```

Authenticate with your Google account:

```bash
clasp login
```

### Create a new standalone project

From the repo root, create a new Apps Script project:

```bash
clasp create --type standalone
```

This writes a `.clasp.json` file linking the local files to the new script project.

### Push files to Apps Script

```bash
clasp push
```

### Deploy as a web app

#### Option A — via clasp

```bash
clasp deploy
```

#### Option B — via the Apps Script editor

1. Open the script project: `clasp open`
2. Click **Deploy > New deployment**
3. Select type: **Web app**
4. Set **Execute as**: Me
5. Set **Who has access**: Only myself (or your org as needed)
6. Click **Deploy** and copy the web app URL

---

## First-Time Setup

After deploying, open the web app URL and complete setup in this order:

1. **Config tab** — Enter your Jira base URL (e.g. `https://your-company.atlassian.net`) and your Atlassian API key, then set your project allocation rows and save.
   - Generate an API key at: `https://id.atlassian.com/manage-profile/security/api-tokens`
2. **Assignments tab** — Load your active Jira issues.
3. **Timecard tab** — Set a date range, import calendar events, assign Jira issues, and submit worklogs.
4. **Create Issue tab** — Create new Jira tasks directly from the app.
5. **Worklog tab** — Pull and review your year-to-date worklogs.

---

## Files

| File | Purpose |
|---|---|
| `Code.js` | All Apps Script server-side logic |
| `Index.html` | Web app shell (loads tabs) |
| `JavaScript.html` | Client-side JS (shared across tabs) |
| `Stylesheet.html` | Shared CSS |
| `TabConfig.html` | Config tab UI |
| `TabAssignments.html` | Assignments tab UI |
| `TabTimecard.html` | Timecard tab UI |
| `TabCreateIssue.html` | Create Issue tab UI |
| `TabWorklog.html` | Worklog tab UI |
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

---

## Notes

- All configuration (Jira URL, API key, allocation) is stored in Apps Script user properties scoped to your Google account — no data is shared with other users of the same deployment.
- Worklog fetches use `UrlFetchApp.fetchAll()` to parallelize requests, significantly reducing load time when many issues are present.
- The script timezone is set to `America/Denver` in `appsscript.json`. All times are converted to UTC before being posted to Jira.
