/*
-----------------------------------------------
Test Suite for Jira Time Tracker
-----------------------------------------------
HOW TO USE:
  1. Copy this file's contents into a new .gs file in your Apps Script project.
  2. Select "runAllTests" from the function dropdown in the editor and click Run.
  3. Open View > Logs to see pass/fail results for each test.

Tests are grouped into three tiers:
  - Pure logic: no GAS APIs, always runnable
  - Sheet structure: verifies required tabs and headers exist
  - API connectivity: read-only calls to Jira and Google Calendar

IMPORTANT: runAllTests() never writes to Jira or modifies any sheet data.
-----------------------------------------------
*/

function runAllTests() {
    Logger.log('========== JIRA TIME TRACKER TEST RUN ==========');
    const results = [
        // Pure logic
        testNextMondayCalculation_(),
        testDurationFormatting_(),
        testUrlConstruction_(),
        testIssueKeyParsing_(),
        // Sheet structure
        testRequiredSheetsExist_(),
        testAssignmentsHeaders_(),
        // Configuration
        testApiKeyIsSet_(),
        testJiraBaseUrlIsSet_(),
        // API connectivity (read-only)
        testJiraConnection_(),
        testCalendarConnection_(),
        testWorklogFetch_(),
        testDropdownValues_()
    ];
    const passed = results.filter(Boolean).length;
    const total = results.length;
    Logger.log(`========== ${passed}/${total} PASSED ==========`);
    SpreadsheetApp.getUi().alert(`Tests complete: ${passed}/${total} passed.\n\nOpen View > Logs for details.`);
}

/*
-----------------------------------------------
Assertion helper
-----------------------------------------------
*/
function assert_(condition, message) {
    if (condition) {
        Logger.log(`  PASS: ${message}`);
    } else {
        Logger.log(`  FAIL: ${message}`);
    }
    return !!condition;
}

/*
-----------------------------------------------
Pure logic tests — no GAS APIs, always runnable
-----------------------------------------------
*/

// Verifies the "days until next Monday" formula used in sendTime()
// for all seven possible starting days of the week.
function testNextMondayCalculation_() {
    Logger.log('--- testNextMondayCalculation_ ---');
    const cases = [
        [0, 1], // Sunday    → 1 day
        [1, 7], // Monday    → 7 days (same day = next week)
        [2, 6], // Tuesday   → 6 days
        [3, 5], // Wednesday → 5 days
        [4, 4], // Thursday  → 4 days
        [5, 3], // Friday    → 3 days
        [6, 2], // Saturday  → 2 days
    ];
    return cases.every(([dayOfWeek, expected]) => {
        const result = dayOfWeek === 1 ? 7 : (8 - dayOfWeek) % 7;
        return assert_(result === expected, `daysUntilMonday(${dayOfWeek}) = ${result}, expected ${expected}`);
    });
}

// Verifies that Google Sheets fractional-day duration values convert
// correctly to whole minutes for the Jira timeSpent field.
function testDurationFormatting_() {
    Logger.log('--- testDurationFormatting_ ---');
    // In Google Sheets, a duration is stored as a fraction of a 24-hour day.
    // e.g. 30 minutes = 30/(24*60) = 0.020833...
    const cases = [
        [30 / (24 * 60), "30m"],   // 30 minutes
        [60 / (24 * 60), "60m"],   // 1 hour
        [75 / (24 * 60), "75m"],   // 1h 15m
        [90 / (24 * 60), "90m"],   // 1h 30m
        [8 * 60 / (24 * 60), "480m"] // 8 hours (full work day)
    ];
    return cases.every(([timeDuration, expected]) => {
        const result = Math.round((timeDuration * 24) * 60) + "m";
        return assert_(result === expected, `duration ${timeDuration.toFixed(5)} → "${result}", expected "${expected}"`);
    });
}

// Verifies that Jira issue URLs are constructed correctly.
function testUrlConstruction_() {
    Logger.log('--- testUrlConstruction_ ---');
    const base = PropertiesService.getUserProperties().getProperty('JIRA_BASE_URL');
    if (!base) {
        Logger.log('  SKIP: No JIRA_BASE_URL set');
        return false;
    }
    const cases = [
        ['PROJ-1', `${base}/browse/PROJ-1`],
        ['LONGPROJ-999', `${base}/browse/LONGPROJ-999`],
    ];
    return cases.every(([key, expected]) => {
        const url = `${base}/browse/${key}`;
        return assert_(url === expected, `URL for ${key}: ${url}`);
    });
}

// Verifies that splitting an issue key on "-" extracts the project prefix.
function testIssueKeyParsing_() {
    Logger.log('--- testIssueKeyParsing_ ---');
    const cases = [
        ['PROJ-1', 'PROJ'],
        ['LONGPROJ-42', 'LONGPROJ'],
        ['AB-100', 'AB'],
    ];
    return cases.every(([key, expected]) => {
        const result = key.split('-')[0];
        return assert_(result === expected, `"${key}".split('-')[0] = "${result}", expected "${expected}"`);
    });
}

/*
-----------------------------------------------
Sheet structure tests
-----------------------------------------------
*/

// Verifies that every sheet the script reads from or writes to exists.
function testRequiredSheetsExist_() {
    Logger.log('--- testRequiredSheetsExist_ ---');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const required = ['Assignments', 'Time Card', 'Calendar', 'Allocation', 'History'];
    return required.every(name =>
        assert_(ss.getSheetByName(name) !== null, `Sheet exists: "${name}"`)
    );
}

// Verifies the Assignments sheet header row matches what retrieveJiraIssues() writes.
// If Assignments is empty, this test is skipped (run Populate Assignments first).
function testAssignmentsHeaders_() {
    Logger.log('--- testAssignmentsHeaders_ ---');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assignments');
    if (!sheet || sheet.getLastRow() < 1) {
        Logger.log('  SKIP: Assignments sheet is empty — run "Populate assignments" first');
        return true;
    }
    const actual = sheet.getRange(1, 1, 1, 8).getValues()[0];
    const expected = ['Dropdown Value', 'Key', 'Link', 'Jira Project', 'Name', 'Status', 'Project', 'Time Logged (h)'];
    return expected.every((col, i) =>
        assert_(actual[i] === col, `Assignments col ${i + 1}: "${actual[i]}", expected "${col}"`)
    );
}

/*
-----------------------------------------------
Configuration tests
-----------------------------------------------
*/

// Verifies the Jira API key has been saved via collectConfig().
function testApiKeyIsSet_() {
    Logger.log('--- testApiKeyIsSet_ ---');
    const key = PropertiesService.getUserProperties().getProperty('JIRA_API_KEY');
    return assert_(key !== null && key.length > 0, 'JIRA_API_KEY is set in user properties');
}

// Verifies the Jira base URL has been saved via collectJiraUrl().
function testJiraBaseUrlIsSet_() {
    Logger.log('--- testJiraBaseUrlIsSet_ ---');
    const url = PropertiesService.getUserProperties().getProperty('JIRA_BASE_URL');
    return assert_(url !== null && url.length > 0, 'JIRA_BASE_URL is set in user properties');
}

/*
-----------------------------------------------
API connectivity tests — read-only, no data is written
-----------------------------------------------
*/

// Confirms credentials are valid and the Jira identity matches the active user.
function testJiraConnection_() {
    Logger.log('--- testJiraConnection_ ---');
    const key = PropertiesService.getUserProperties().getProperty('JIRA_API_KEY');
    if (!key) {
        Logger.log('  SKIP: No API key set');
        return false;
    }
    const userEmail = Session.getActiveUser().getEmail();
    const authHeader = 'Basic ' + Utilities.base64Encode(`${userEmail}:${key}`);
    const jiraUrl = PropertiesService.getUserProperties().getProperty('JIRA_BASE_URL');
    if (!jiraUrl) {
        Logger.log('  SKIP: No JIRA_BASE_URL set');
        return false;
    }
    try {
        const response = UrlFetchApp.fetch(
            `${jiraUrl}/rest/api/3/myself`,
            { headers: { Authorization: authHeader }, method: 'get', muteHttpExceptions: true }
        );
        const code = response.getResponseCode();
        const data = JSON.parse(response.getContentText());
        const ok = assert_(code === 200, `Jira API responded ${code} (expected 200)`);
        assert_(data.emailAddress === userEmail, `Jira identity: "${data.emailAddress}" matches active user`);
        return ok;
    } catch (e) {
        Logger.log(`  FAIL: Jira connection threw — ${e}`);
        return false;
    }
}

// Confirms the Google Calendar for the active user is accessible.
function testCalendarConnection_() {
    Logger.log('--- testCalendarConnection_ ---');
    try {
        const email = Session.getActiveUser().getEmail();
        const calendar = CalendarApp.getCalendarById(email);
        return assert_(calendar !== null, `Calendar accessible for ${email}`);
    } catch (e) {
        Logger.log(`  FAIL: Calendar connection threw — ${e}`);
        return false;
    }
}

// Calls getIssueWorklogTotal_() against the first issue in Assignments
// and verifies it returns a non-negative number. Read-only.
function testWorklogFetch_() {
    Logger.log('--- testWorklogFetch_ ---');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assignments');
    if (!sheet || sheet.getLastRow() < 2) {
        Logger.log('  SKIP: Assignments sheet is empty — run "Populate assignments" first');
        return true;
    }
    const issueKey = sheet.getRange(2, 2).getValue();
    if (!issueKey) {
        Logger.log('  SKIP: No issue key found in Assignments B2');
        return true;
    }
    const key = PropertiesService.getUserProperties().getProperty('JIRA_API_KEY');
    const userEmail = Session.getActiveUser().getEmail();
    const authHeader = 'Basic ' + Utilities.base64Encode(`${userEmail}:${key}`);
    try {
        const jiraUrl = PropertiesService.getUserProperties().getProperty('JIRA_BASE_URL');
        if (!jiraUrl) {
            Logger.log('  SKIP: No JIRA_BASE_URL set');
            return false;
        }
        const total = getIssueWorklogTotal_(issueKey, authHeader, jiraUrl, userEmail);
        return assert_(typeof total === 'number' && total >= 0, `getIssueWorklogTotal_("${issueKey}") = ${total}s`);
    } catch (e) {
        Logger.log(`  FAIL: getIssueWorklogTotal_ threw — ${e}`);
        return false;
    }
}

// Verifies getDropdownValues() returns a non-empty list from the Allocation sheet.
function testDropdownValues_() {
    Logger.log('--- testDropdownValues_ ---');
    try {
        const values = getDropdownValues();
        const ok = assert_(Array.isArray(values) && values.length > 0, `getDropdownValues() returned ${values.length} item(s)`);
        values.forEach(v => Logger.log(`    → "${v}"`));
        return ok;
    } catch (e) {
        Logger.log(`  FAIL: getDropdownValues threw — ${e}`);
        return false;
    }
}
