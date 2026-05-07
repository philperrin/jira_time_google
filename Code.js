/*
-----------------------------------------------
00: Adds UI menu options to the sheet.
-----------------------------------------------
*/
/** Adds the "Log Jira Time" menu and its items to the spreadsheet UI on open. */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const subMenu = ui.createMenu('Configuration')
        .addItem('Create new Jira task', 'createJira')
        .addItem('Add/Update Jira API', 'collectConfig')
        .addItem('Add/Update Jira Base URL', 'collectJiraUrl')
        .addItem('Run tests', 'runAllTests');
    ui.createMenu('Log Jira Time')
        .addItem('Populate assignments', 'retrieveJiraIssues')
        .addItem('Schedule calendar events', 'scheduleCalendarEvents')
        .addItem('Import calendar events', 'importCalendarEventsToSheet')
        .addItem('Send time to Jira', 'sendTime')
        .addSubMenu(subMenu)
        .addToUi();
}

/*
-----------------------------------------------
01: Prompts for the user to enter their Jira API key and the base URL.
-----------------------------------------------
*/
/** Prompts the user for their Jira API key and saves it to user properties. */
function collectConfig() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
        'Jira API Key',
        'Paste your API key here:',
        ui.ButtonSet.OK_CANCEL
    );
    const button = response.getSelectedButton();
    const key = response.getResponseText();
    if (button === ui.Button.OK) {
        getUserProperties().setProperty('JIRA_API_KEY', key);
        ui.alert('API key saved successfully!');
    } else {
        ui.alert('Input cancelled.');
    }
}

/** Prompts the user for their Jira base URL and saves it to user properties. */
function collectJiraUrl() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
        'Jira Base URL',
        'Paste your Jira base URL here (e.g. https://your-company.atlassian.net):',
        ui.ButtonSet.OK_CANCEL
    );
    const button = response.getSelectedButton();
    const url = response.getResponseText().trim().replace(/\/$/, '');
    if (button === ui.Button.OK) {
        getUserProperties().setProperty('JIRA_BASE_URL', url);
        ui.alert('Jira base URL saved successfully!');
    } else {
        ui.alert('Input cancelled.');
    }
}

/*
-----------------------------------------------
02: Fetch active Jira issues assigned to or watched by the current user
and populate the Assignments sheet with key, link, project, name, status,
and total time logged by the current user.
-----------------------------------------------
*/
/**
 * Fetches active Jira issues assigned to or watched by the current user and
 * populates the Assignments sheet. Excludes archived and internal projects.
 * Creates named ranges per project for use in Calendar sheet dropdowns.
 */
function retrieveJiraIssues() {
    const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
    const USER_EMAIL = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Assignments');
    const authHeader = getAuthHeader_();
    const BASE_ENDPOINT = `${JIRA_URL}/rest/api/3/search/jql?jql=(assignee=currentUser()+OR+watcher=currentUser())+AND+issuetype+IN+(Story,Task,Sub-Task)+AND+(status!=Done+OR+(status=Done+AND+updated%3E=-7d))+ORDER+BY+key+ASC&fields=key,summary,status,project&maxResults=100`;
    const options = {
        headers: { Authorization: authHeader },
        method: 'get',
        muteHttpExceptions: true
    };
    try {
        let allIssues = [];
        let nextPageToken = null;
        do {
            const endpoint = nextPageToken ? `${BASE_ENDPOINT}&nextPageToken=${nextPageToken}` : BASE_ENDPOINT;
            const jiraData = JSON.parse(UrlFetchApp.fetch(endpoint, options).getContentText());
            if (jiraData.issues && jiraData.issues.length > 0) {
                allIssues = allIssues.concat(jiraData.issues);
            }
            nextPageToken = jiraData.nextPageToken || null;
        } while (nextPageToken);

        sheet.clear();

        if (allIssues.length > 0) {
            const headers = ['Dropdown Value', 'Key', 'Link', 'Jira Project', 'Name', 'Status', 'Project', 'Time Logged (h)', 'Schedule Time'];
            sheet.appendRow(headers);
            const filteredIssues = allIssues.filter(issue => {
                const projName = issue.fields.project.name;
                return !projName.includes('Archive') && !projName.includes('Managed Services Internal');
            });
            const worklogTotals = getWorklogTotals_(filteredIssues.map(i => i.key), authHeader, JIRA_URL, USER_EMAIL);
            const rows = filteredIssues.map((issue, idx) => {
                const issueKey = issue.key;
                const issueName = issue.fields.summary;
                return [
                    `${issueKey} (${issueName})`,
                    issueKey,
                    `${JIRA_URL}/browse/${issueKey}`,
                    issueKey.split('-')[0],
                    issueName,
                    issue.fields.status.name,
                    issue.fields.project.name,
                    Math.round((worklogTotals[idx] / 3600) * 100) / 100,
                    ''
                ];
            });
            if (rows.length > 0) {
                sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
                // Overwrite the Link column with clickable HYPERLINK formulas.
                sheet.getRange(2, 3, rows.length, 1).setFormulas(
                    rows.map(row => [`=HYPERLINK("${row[2]}","${row[1]}")`])
                );
                sheet.getNamedRanges().forEach(nr => nr.remove());
                let prevProj = null;
                let startRow = 2;
                for (let i = 0; i < rows.length; i++) {
                    const currProj = rows[i][3];
                    if (currProj !== prevProj && prevProj !== null) {
                        ss.setNamedRange(prevProj, sheet.getRange(startRow, 1, i - (startRow - 2), 1));
                        startRow = i + 2;
                    }
                    prevProj = currProj;
                }
                // Always register the final group, regardless of its size.
                if (prevProj !== null) {
                    ss.setNamedRange(prevProj, sheet.getRange(startRow, 1, rows.length - (startRow - 2), 1));
                }
            }
            SpreadsheetApp.getUi().alert(`Successfully retrieved ${filteredIssues.length} issues.`);
        }
    } catch (e) {
        Logger.log(`retrieveJiraIssues error: ${e}`);
        SpreadsheetApp.getUi().alert('Error: Problem with getting your Jira tasks.');
    }
}

/*
-----------------------------------------------
03: Import calendar events for the dates on the Time Card sheet into the
Calendar sheet. Clears existing data first, then repopulates. Events
categorized as Personal or Internal are excluded as they are not billable.
-----------------------------------------------
*/
/**
 * Imports calendar events for the date range in the Time Card sheet into the
 * Calendar sheet. Skips non-billable Personal and Internal events, then
 * rebuilds issue dropdowns in column H based on each event's project color.
 */
function importCalendarEventsToSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const CALENDAR_ID = Session.getActiveUser().getEmail();
    const dateValues = ss.getSheetByName('Time Card').getRange('A2:A').getValues()
        .flat().filter(cell => cell instanceof Date && !isNaN(cell));
    if (dateValues.length === 0) {
        Logger.log('No valid dates found in Time Card.');
        return;
    }
    const START_DATE = new Date(Math.min(...dateValues.map(d => d.getTime())));
    const END_DATE = new Date(Math.max(...dateValues.map(d => d.getTime())));
    END_DATE.setDate(END_DATE.getDate() + 1);

    const sheet = ss.getSheetByName('Calendar');
    if (!sheet) {
        Logger.log('Sheet not found: Calendar');
        return;
    }
    sheet.getRange('A2:H').clear();
    try {
        const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
        if (!calendar) {
            Logger.log('Calendar not found or accessible.');
            return;
        }
        const allocationSheet = ss.getSheetByName('Allocation');
        const allocationLastRow = allocationSheet.getLastRow();
        const allocationData = allocationLastRow >= 2
            ? allocationSheet.getRange(2, 1, allocationLastRow - 1, 3).getValues()
            : [];
        const colorMap = Object.fromEntries(
            allocationData.filter(row => row[1] !== "").map(row => [row[0], row[1]])
        );
        // Only include events whose color maps to a Column B project key that has a non-empty Column C name.
        const validProjectKeys = new Set(
            allocationData.filter(row => row[1] !== "" && row[2] !== "").map(row => row[1])
        );
        const tz = Session.getScriptTimeZone();
        const data = calendar.getEvents(START_DATE, END_DATE).map(event => {
            const startTime = event.getStartTime();
            const description = event.getDescription();
            const match = description.match(/^[^________________________________________________________________________________]+/); // Strips Teams call-in details.
            return [
                event.getTitle(),
                Utilities.formatDate(startTime, tz, "yyyy-MM-dd"),
                Utilities.formatDate(startTime, tz, "hh:mm a"),
                Utilities.formatDate(event.getEndTime(), tz, "hh:mm a"),
                match ? match[0] : "",
                event.getMyStatus(),
                colorMap[event.getColor()] || "Unknown"
            ];
        });
        const filteredData = data.filter(row => validProjectKeys.has(row[6]));
        if (filteredData.length > 0) {
            sheet.getRange(2, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
        }
    } catch (error) {
        Logger.log(`importCalendarEventsToSheet error: ${error.toString()}`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const numRows = lastRow - 1;
    const gValues = sheet.getRange(2, 7, numRows, 1).getValues();
    const hRange = sheet.getRange(2, 8, numRows, 1);
    const hValidations = hRange.getDataValidations();

    for (let i = 0; i < gValues.length; i++) {
        const namedRangeName = gValues[i][0];
        if (!namedRangeName) continue;
        try {
            const nr = ss.getRangeByName(namedRangeName);
            if (!nr) continue;
            hValidations[i][0] = SpreadsheetApp.newDataValidation()
                .requireValueInRange(nr, true)
                .setAllowInvalid(false)
                .build();
        } catch (e) {
            Logger.log(`Skipping row ${i + 2} due to error: ${e}`);
        }
    }
    hRange.setDataValidations(hValidations);

    const hFormulas = Array.from({ length: numRows }, (_, i) => {
        const row = i + 2;
        if (!gValues[i][0]) return [''];
        return [`=ai("Based on the event name ("&A${row}&") and description ("&E${row}&"), what is the best Jira task for this event from the list of tasks: "&unique(indirect(G${row})))`];
    });
    hRange.setFormulas(hFormulas);
}

/*
-----------------------------------------------
04: Read "Schedule Time" values (column I) from the Assignments sheet and
create sequential Google Calendar events starting at the next full hour.
Each event's title is the Jira Project (column D) and its description is
the Dropdown Value (column A). Events are stacked back-to-back.
-----------------------------------------------
*/
/**
 * Creates Google Calendar events for each row in the Assignments sheet that
 * has a numeric value in the "Schedule Time" column (I). Events are scheduled
 * sequentially starting at the next full hour from now. After creating events,
 * the Schedule Time column is cleared.
 */
function scheduleCalendarEvents() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Assignments');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
        SpreadsheetApp.getUi().alert('No data found in the Assignments sheet.');
        return;
    }

    // Build project key → EventColor from Allocation columns B (col 2) and G (col 7).
    const allocationSheet = ss.getSheetByName('Allocation');
    const allocationLastRow = allocationSheet.getLastRow();
    const allocationData = allocationLastRow >= 2
        ? allocationSheet.getRange(2, 2, allocationLastRow - 1, 6).getValues()
        : [];
    const COLOR_ENUM_MAP = {
        'pale blue': CalendarApp.EventColor.PALE_BLUE,
        'pale green': CalendarApp.EventColor.PALE_GREEN,
        'mauve': CalendarApp.EventColor.MAUVE,
        'pale red': CalendarApp.EventColor.PALE_RED,
        'yellow': CalendarApp.EventColor.YELLOW,
        'orange': CalendarApp.EventColor.ORANGE,
        'cyan': CalendarApp.EventColor.CYAN,
        'gray': CalendarApp.EventColor.GRAY,
        'grey': CalendarApp.EventColor.GRAY,
        'blue': CalendarApp.EventColor.BLUE,
        'green': CalendarApp.EventColor.GREEN,
        'red': CalendarApp.EventColor.RED
    };
    // allocationData columns (0-indexed): 0=B (project key), 5=G (color name)
    const projectColorMap = Object.fromEntries(
        allocationData
            .filter(row => row[0] !== '')
            .map(row => [row[0], COLOR_ENUM_MAP[String(row[5]).trim().toLowerCase()] || null])
    );

    const numRows = lastRow - 1;
    const data = sheet.getRange(2, 1, numRows, 9).getValues();

    // Collect rows that have a positive numeric Schedule Time value.
    const toSchedule = data.reduce((acc, row, i) => {
        const scheduleTime = row[8]; // column I (0-indexed: 8)
        if (typeof scheduleTime === 'number' && scheduleTime > 0) {
            acc.push({ rowIndex: i + 2, dropdownValue: row[0], jiraProject: row[3], hours: scheduleTime });
        }
        return acc;
    }, []);

    if (toSchedule.length === 0) {
        SpreadsheetApp.getUi().alert('No Schedule Time values found. Enter a numeric value in column I for the rows you want to schedule.');
        return;
    }

    // Start time = next full hour from now.
    const now = new Date();
    const startTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours() + 1, 0, 0, 0);

    const calendar = CalendarApp.getDefaultCalendar();
    let cursor = new Date(startTime);
    const created = [];

    toSchedule.forEach(entry => {
        const durationMs = entry.hours * 60 * 60 * 1000;
        const endTime = new Date(cursor.getTime() + durationMs);
        const event = calendar.createEvent(entry.jiraProject, cursor, endTime, { description: entry.dropdownValue });
        const eventColor = projectColorMap[entry.jiraProject];
        if (eventColor) event.setColor(eventColor);
        created.push(entry.rowIndex);
        cursor = endTime;
    });

    // Clear Schedule Time values for rows that were processed.
    created.forEach(rowIndex => {
        sheet.getRange(rowIndex, 9).clearContent();
    });

    SpreadsheetApp.getUi().alert(`Created ${created.length} calendar event(s) starting at ${startTime.toLocaleTimeString()}.`);
}

/*
-----------------------------------------------
05: Open a modal dialog for creating a new Jira issue. On submit, the Jira
REST API is called to create the issue and the Assignments sheet is refreshed.
-----------------------------------------------
*/
/** Opens the Create New Jira modal dialog. */
function createJira() {
    const html = HtmlService.createHtmlOutputFromFile('CreateNewJira')
        .setTitle('newJira')
        .setWidth(500)
        .setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Create a new Jira issue');
}

/** Returns the list of Jira project keys from the Allocation sheet for use in the dialog dropdown. */
function getDropdownValues() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Allocation');
    return sheet.getRange('B2:B' + sheet.getLastRow()).getValues().map(row => row[0]).filter(String);
}

/**
 * Called by the Create New Jira dialog on submit. Posts the new issue to Jira
 * via REST API, then refreshes the Assignments sheet. Returns a status string
 * displayed in the dialog.
 */
function makeJira(formData) {
    const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
    const authHeader = getAuthHeader_();
    const payload = {
        fields: {
            project: { key: formData.input1 },
            summary: formData.input4,
            description: {
                type: 'doc',
                version: 1,
                content: [{ type: 'paragraph', content: [{ type: 'text', text: formData.input5 || '' }] }]
            },
            issuetype: { name: 'Task' }
        }
    };
    const options = {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: authHeader },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };
    try {
        const response = UrlFetchApp.fetch(`${JIRA_URL}/rest/api/3/issue`, options);
        const responseCode = response.getResponseCode();
        const responseData = JSON.parse(response.getContentText());
        if (responseCode === 201) {
            retrieveJiraIssues();
            return `Successfully created ${responseData.key}.`;
        } else {
            return `Error (${responseCode}): ${JSON.stringify(responseData.errors || responseData)}`;
        }
    } catch (e) {
        return `Error: ${e.message}`;
    }
}

/*
-----------------------------------------------
05: Read time entries from the Time Card sheet and post each as a worklog
to its Jira issue. Time is converted to UTC before posting. Successful
entries are archived to the History sheet. After all entries are sent,
the Time Card is cleared, Calendar dropdowns are reset, and the next
Monday's date is populated in A2.
-----------------------------------------------
*/
/**
 * Reads time entries from the Time Card sheet and posts each as a Jira worklog.
 * Successful entries are written to the History sheet in a single batch.
 * After sending, clears the Time Card, resets Calendar dropdowns, and sets
 * cell A2 to the next Monday.
 */
function sendTime() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Time Card');
    const UTC_FORMAT = "yyyy-MM-dd'T'HH:mm:ss'.000+0000'";
    const lastDataRow = sheet.getLastRow();
    const values = lastDataRow >= 2 ? sheet.getRange(2, 1, lastDataRow - 1, 4).getValues() : [];
    const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
    const authHeader = getAuthHeader_();
    const historySheet = ss.getSheetByName('History');
    const historyRows = [];

    for (let i = 0; i < values.length; i++) {
        const row = values[i];
        if (!row[1] || row[1].length === 0) continue;
        const dateValue = row[0];
        const issueID = row[1];
        const timeValue = row[2];
        const timeDuration = row[3];
        if (!(dateValue instanceof Date && timeValue instanceof Date)) {
            Logger.log("Error: Cells must contain valid date/time values.");
            continue;
        }
        const combinedDateTime = new Date(dateValue);
        combinedDateTime.setHours(timeValue.getHours(), timeValue.getMinutes(), timeValue.getSeconds());
        const utcString = Utilities.formatDate(combinedDateTime, 'Etc/GMT', UTC_FORMAT);
        const durationMinutes = Math.round((timeDuration * 24) * 60);
        const timeSpent = durationMinutes + "m";
        const options = {
            method: 'post',
            headers: { 'Authorization': authHeader, 'Content-Type': 'application/json' },
            payload: JSON.stringify({ started: utcString, timeSpent }),
            muteHttpExceptions: true
        };
        try {
            const response = UrlFetchApp.fetch(`${JIRA_URL}/rest/api/3/issue/${issueID}/worklog`, options);
            Logger.log(`Row ${i}: ${issueID} — ${response.getContentText()}`);
            historyRows.push([dateValue, issueID, timeSpent]);
        } catch (error) {
            Logger.log('Error during API request: ' + error.message);
        }
    }
    if (historyRows.length > 0) {
        historySheet.getRange(historySheet.getLastRow() + 1, 1, historyRows.length, 3).setValues(historyRows);
    }

    sheet.getRange(`B2:D${lastDataRow}`).clearContent();
    ss.getSheetByName('Assignments').getNamedRanges().forEach(nr => nr.remove());
    const calSheet = ss.getSheetByName('Calendar');
    calSheet.getRange('H2:H' + calSheet.getLastRow()).clearDataValidations();

    const today = new Date();
    const dayOfWeek = today.getDay();
    const daysUntilMonday = dayOfWeek === 1 ? 7 : (8 - dayOfWeek) % 7;
    sheet.getRange("A2").setValue(
        new Date(today.getFullYear(), today.getMonth(), today.getDate() + daysUntilMonday)
    );

    const queryTemplate = "=iferror(query(Calendar!B:K,\"SELECT Col9,Col2,Col10,Col6 WHERE Col1=Date '\"&text(A{r},\"YYYY-MM-DD\")&\"' AND Col7 IS NOT NULL ORDER BY Col2 ASC\",0),\"\")";
    const ROWS_PER_DAY = 21;
    const DAYS_PER_WEEK = 5;
    Array.from({ length: DAYS_PER_WEEK }, (_, i) => 2 + i * ROWS_PER_DAY).forEach(r => {
        sheet.getRange(`B${r}`).setFormula(queryTemplate.replace('{r}', r));
    });
}

/*
-----------------------------------------------
Helper Functions
-----------------------------------------------
*/
/** Returns the PropertiesService store scoped to the current user. */
function getUserProperties() {
    return PropertiesService.getUserProperties();
}

/** Builds the Basic auth header for Jira API requests. */
function getAuthHeader_() {
    const email = Session.getActiveUser().getEmail();
    const apiKey = getUserProperties().getProperty('JIRA_API_KEY');
    return 'Basic ' + Utilities.base64Encode(`${email}:${apiKey}`);
}

/**
 * Fetches worklog totals for all given issue keys in parallel using fetchAll.
 * Handles pagination — if an issue has more than 100 worklogs, subsequent pages
 * are also batched. Returns an array of totals (in seconds) in the same order
 * as issueKeys.
 */
function getWorklogTotals_(issueKeys, authHeader, jiraUrl, userEmail) {
    const totals = new Array(issueKeys.length).fill(0);
    if (issueKeys.length === 0) return totals;
    const makeRequest = (key, startAt) => ({
        url: `${jiraUrl}/rest/api/3/issue/${key}/worklog?startAt=${startAt}&maxResults=100`,
        headers: { Authorization: authHeader },
        method: 'get',
        muteHttpExceptions: true
    });
    let pending = issueKeys.map((_, i) => ({ keyIndex: i, startAt: 0 }));
    while (pending.length > 0) {
        const responses = UrlFetchApp.fetchAll(pending.map(p => makeRequest(issueKeys[p.keyIndex], p.startAt)));
        const nextPending = [];
        responses.forEach((response, i) => {
            const { keyIndex, startAt } = pending[i];
            const data = JSON.parse(response.getContentText());
            if (!data.worklogs) return;
            data.worklogs.forEach(log => {
                if (log.author && log.author.emailAddress === userEmail) {
                    totals[keyIndex] += log.timeSpentSeconds || 0;
                }
            });
            const nextStart = startAt + data.worklogs.length;
            if (nextStart < data.total) {
                nextPending.push({ keyIndex, startAt: nextStart });
            }
        });
        pending = nextPending;
    }
    return totals;
}

/** Returns the total seconds logged by userEmail on a single issue. */
function getIssueWorklogTotal_(issueKey, authHeader, jiraUrl, userEmail) {
    return getWorklogTotals_([issueKey], authHeader, jiraUrl, userEmail)[0];
}

/** Logs all stored user properties to the Apps Script console. Useful for debugging. */
function showUserProperties() {
    Logger.log('All User Properties: ' + JSON.stringify(getUserProperties().getProperties()));
}
