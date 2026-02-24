/*
-----------------------------------------------
00: This just adds some UI menu options on your sheet. Each option will run one of 
the subsequent functions.
-----------------------------------------------
*/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Log Jira Time')
    .addItem('Add/Update Jira API', 'collectConfig')
    .addItem('Populate Jira tasks', 'retrieveJiraIssues')
    .addItem('Import calendar events', 'importCalendarEventsToSheet')
    .addItem('Create new Jira issue', 'createJira')
    .addItem('Send time to Jira', 'sendTime')
    .addToUi();
}

/*
-----------------------------------------------
01: Store Jira API key and store project codes.
Maybe don't need to collection project codes. Currently developing this.
-----------------------------------------------
*/
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

/*
-----------------------------------------------
02: Get a list of Jira tasks to bill time against.
This Jira API looks for all tasks you are watching or are assigned.
-----------------------------------------------
*/
function retrieveJiraIssues() {
  const JIRA_URL = 'https://phdata.atlassian.net';
  const userProperties = getUserProperties();
  const API_KEY = userProperties.getProperty('JIRA_API_KEY');
  const API_ENDPOINT = `${JIRA_URL}/rest/api/3/search/jql?jql=(assignee=currentUser()+OR+watcher=currentUser())+AND+issuetype+IN+(Story,Task,Subtask)+AND+(status!=Done+OR+(status=Done+AND+updated%3E=-7d))+ORDER+BY+key+ASC&fields=key,summary,status,created,updated,customfield_10201,project&expand=%3Cstring%3C&maxResults=100`;
  const USER_EMAIL = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assignments');
  const credentials = `${USER_EMAIL}:${API_KEY}`;
  const authHeader = 'Basic ' + Utilities.base64Encode(credentials);
  const options = {
    headers: {
      Authorization: authHeader
    },
    method: 'get',
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(API_ENDPOINT, options);
    const jiraData = JSON.parse(response.getContentText());
    sheet.clear();
    if (jiraData.issues && jiraData.issues.length > 0) {
      const headers = ['Dropdown Value', 'Key', 'Jira Project', 'Name', 'Status', 'Created', 'Updated', 'Project', 'MS Billing Ref', 'Next Page Token'];
      sheet.appendRow(headers);
      const filteredIssues = jiraData.issues.filter(issue => {
        const issueProj = issue.fields.project.name;
        return !issueProj.includes('Archive') && !issueProj.includes('Managed Services Internal');
      });
      const rows = filteredIssues.map(issue => {
        const issueKey = issue.key;
        const issueKeySplit = issueKey.split('-');
        const issueProjName = issueKeySplit[0];
        const issueName = issue.fields.summary;
        const issueKeyName = `${issueKey} (${issueName})`;
        const issueProj = issue.fields.project.name;
        const issueStatus = issue.fields.status.name;
        const issueCreated = issue.fields.created;
        const issueUpdated = issue.fields.updated;
        const issueMSBilling = issue.fields.customfield_10201;
        const nextPageToken = jiraData.nextPageToken;
        return [issueKeyName, issueKey, issueProjName, issueName, issueStatus, issueCreated, issueUpdated, issueProj, issueMSBilling, nextPageToken];
      });
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
        const namedRanges = sheet.getNamedRanges();
        namedRanges.forEach(nr => nr.remove());
        let prevProj = null;
        let startRow = 2;
        for (let i = 0; i < rows.length; i++) {
          const currProj = rows[i][2];
          if (currProj !== prevProj && prevProj !== null) {
            const range = sheet.getRange(startRow, 1, i - (startRow - 2), 1);
            ss.setNamedRange(prevProj, range);
            startRow = i + 2;
          }
          prevProj = currProj;
        }
        if (rows.length > startRow) {
          const range = sheet.getRange(startRow, 1, rows.length - (startRow - 2), 1);
          ss.setNamedRange(prevProj, range);
        }
      }
      Browser.msgBox(`Successfully retrieved ${filteredIssues.length} issues.`);
    }
  } catch (e) {
    Browser.msgBox('Error: Problem with getting your Jira tasks.');
  }
}

/*
-----------------------------------------------
03: Import calendar events to the sheet. This function will clear the sheet first 
and then repopulate it with the events from your calendar - so if you don't see the item you're 
looking for, you can double check your calendar to make sure it is in your calendar, and if it 
is not, then you can create the event. Then just run this function again. 
-----------------------------------------------
*/
function importCalendarEventsToSheet() {
  const CALENDAR_ID = Session.getActiveUser().getEmail();;
  const SHEET_NAME = 'Calendar';
  const CALDATES = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Time Card').getRange('A2:A').getValues();
  const dateValues = CALDATES.flat().filter(cell => cell instanceof Date && !isNaN(cell));
  if (dateValues.length === 0) {
    Logger.log('No valid dates found in Time Card.');
    return;
  }
  const minTime = Math.min(...dateValues.map(date => date.getTime()));
  const maxTime = Math.max(...dateValues.map(date => date.getTime()));
  const START_DATE = new Date(minTime);
  const LAST_DATE = new Date(maxTime);
  const END_DATE = new Date();
  END_DATE.setDate(LAST_DATE.getDate() + 1);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`Sheet not found: ${SHEET_NAME}`);
    return;
  }
  sheet.getRange('A2:G').clear();
  try {
    const headers = ['Name', 'Start Date', 'Start Time', 'End Time', 'Details', 'Attending Status', 'Project'];
    sheet.appendRow(headers);
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) {
      Logger.log('Calendar not found or accessible.');
      return;
    }
    const events = calendar.getEvents(START_DATE, END_DATE);
    const allocationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Allocation');
    const allocationData = allocationSheet.getRange(1, 1, allocationSheet.getLastRow(), 2).getValues();
    const colorMap = Object.fromEntries(
      allocationData
        .filter(row => row[0] !== "")
        .map(row => [row[0], row[3]])
    );
    const data = events.map(event => {
      const title = event.getTitle();
      const tz = Session.getScriptTimeZone();
      const startTime_in = event.getStartTime();
      const startTime = Utilities.formatDate(startTime_in, tz, "hh:mm a");
      const endTime_in = event.getEndTime();
      const endTime = Utilities.formatDate(endTime_in, tz, "hh:mm a");
      const startDate = Utilities.formatDate(startTime, tz, "yyyy-MM-dd");
      const description = event.getDescription();
      const match = description.match(/^[^________________________________________________________________________________]+/); // Cuts off all the call-in info from Teams calls.
      const descriptionClean = match ? match[0] : "";
      const eventColor = event.getColor();
      const eventAttending = event.getMyStatus();
      const mappedValue = colorMap[eventColor] || "Unknown";
      return [title, startDate, startTime, endTime, descriptionClean, eventAttending, mappedValue];
    });
    const filteredData = data.filter(row => row[5] != 8 && row[5] != 11);
    if (filteredData.length > 0) {
      sheet.getRange(2, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
      Logger.log("Calendar data has been fetched and updated in the sheet.");
    } else {
      Logger.log("No events found in the specified date range.");
    }
  } catch (error) {
    Logger.log(`Error: ${error.toString()}`);
  }

  const START_ROW = 2;
  const COL_G = 7;
  const COL_H = 8;

  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) return;

  const numRows = lastRow - START_ROW + 1;
  const gRange = sheet.getRange(START_ROW, COL_G, numRows, 1);
  const hRange = sheet.getRange(START_ROW, COL_H, numRows, 1);
  const gValues = gRange.getValues();
  const hValidations = hRange.getDataValidations();

  for (let i = 0; i < gValues.length; i++) {
    const namedRangeName = gValues[i][0];
    if (!namedRangeName) {
      continue;
    }
    try {
      const nr = spreadsheet.getRangeByName(namedRangeName);
      if (!nr) {
        continue;
      }
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(nr, true)
        .setAllowInvalid(false)
        .build();

      hValidations[i][0] = rule;
    } catch (e) {
      Logger.log(`Skipping row ${START_ROW + i} due to error: ${e}`);
      continue;
    }
  }
  hRange.setDataValidations(hValidations);
}

/*
-----------------------------------------------
04: Create new Jira issue.
NOT YET BUILT.
-----------------------------------------------
*/
function createJira() {
  const html = HtmlService.createHtmlOutputFromFile('CreateNewJira')
    .setTitle('newJira')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Create a new Jira issue');
}

function getDropdownValues() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Allocation');
  const createJiraProj = sheet.getRange('C2:C' + sheet.getLastRow()).getValues();
  const flatList = createJiraProj.map(row => row[0]).filter(String);
  return flatList;
}

/*
-----------------------------------------------
05: The main event - send time entries to Jira. This takes the data from whatever 
sheet you are on, iterates through each line and adds the time to the Jira issue referenced on 
that line. This function also uses data from the "Config" sheet - as well as the "Assignments" 
sheet. Jira needs the 'Started' value to be UTC. So we have to convert whatever timezone you're 
in to UTC. It also calculates the duration per issue and returns the value in minutes - but 
don't worry you can still log things like 200 minutes. 
After the time gets logged to Jira, the values from your time card are also saved on the 
History tab. Once they are stored on that tab, the values on the Time Card tab are cleared out, 
preserving the data validation.
Finally, once all the data has been cleaned up, the date of the next Monday is populated in 
cell A2 on the Time Card sheet, ready for you to log your time then. But until then, go enjoy 
the weekend.
-----------------------------------------------
*/

function sendTime() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const UTC_FORMAT = "yyyy-MM-dd'T'HH:mm:ss'.000+0000'";
  const timeCardRange = sheet.getRange(2, 1, 105, 4);
  const values = timeCardRange.getValues();
  const CURR_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (!row[1] || row[1].length === 0) continue;
    const dateValue = row[0];
    const issueID = row[1];
    const timeValue = row[2];
    const timeValueEnd = row[3];
    if (!(dateValue instanceof Date && timeValue instanceof Date)) {
      Logger.log("Error: Cells must contain valid date/time values.");
      continue;
    }
    const combinedDateTime = new Date(dateValue);
    combinedDateTime.setHours(timeValue.getHours());
    combinedDateTime.setMinutes(timeValue.getMinutes());
    combinedDateTime.setSeconds(timeValue.getSeconds());
    const utcString = Utilities.formatDate(combinedDateTime, 'Etc/GMT', UTC_FORMAT);
    const durationMs = timeValueEnd.getTime() - timeValue.getTime();
    let durationMinutes = durationMs / (1000 * 60);
    durationMinutes += "m";
    const JIRA_URL = 'https://phdata.atlassian.net';
    const lookupSheet = CURR_SPREADSHEET.getSheetByName("Assignments");
    const dataRange = lookupSheet.getRange("A:G");
    const dataValues = dataRange.getValues();
    let resultValue = "Not Found";
    let projValue = "Not Found";
    let billRefValue = "Not Found";
    for (let j = 0; j < dataValues.length; j++) {
      if (dataValues[j][0] === issueID) {
        resultValue = dataValues[j][1];
        projValue = dataValues[j][5];
        billRefValue = dataValues[j][6];
        break;
      }
    }
    const API_ENDPOINT = `${JIRA_URL}/rest/api/2/issue/${resultValue}/worklog`;
    const USER_EMAIL2 = Session.getActiveUser().getEmail();
    const userProperties = getUserProperties();
    const API_TOKEN = userProperties.getProperty('JIRA_API_KEY');
    const payload = {
      started: utcString,
      timeSpent: durationMinutes
    };
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: 'Basic ' + Utilities.base64Encode(`${USER_EMAIL2}:${API_TOKEN}`)
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    try {
      const response = UrlFetchApp.fetch(API_ENDPOINT, options);
      Logger.log("Response Content: " + response.getContentText());
    } catch (error) {
      Logger.log('Error during API request: ' + error.message);
    }
    const rowData = [dateValue, resultValue, durationMinutes, projValue, billRefValue];
    CURR_SPREADSHEET.getSheetByName('History').appendRow(rowData);
  }
  const TIMECARDSHEET = CURR_SPREADSHEET.getSheetByName('Time Card');
  TIMECARDSHEET.getRange("B2:D105").clearContent();
  const today = new Date();
  const todayDayOfWeek = today.getDay();
  const targetDayOfWeek = 1;
  const daysUntilMonday = todayDayOfWeek === targetDayOfWeek ? 7 : (targetDayOfWeek + 7 - todayDayOfWeek) % 7;
  const nextMonday = new Date(today);
  const mondayOnly = new Date(nextMonday.getFullYear(), nextMonday.getMonth(), nextMonday.getDate());
  mondayOnly.setDate(today.getDate() + daysUntilMonday);
  TIMECARDSHEET.getRange("A2").setValue(mondayOnly);
}

/*
-----------------------------------------------
Helper Functions
-----------------------------------------------
*/
function getUserProperties() {
  return PropertiesService.getUserProperties();
}

function showUserProperties() {
  const allProperties = getUserProperties().getProperties();
  Logger.log('All User Properties: ' + JSON.stringify(allProperties));
}
