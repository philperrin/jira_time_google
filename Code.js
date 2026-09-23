function doGet() {
  const FILE_ID = '1FJo2Y9M0tnVX9hkbzd9vrbvdTrPeb_jK';
  const faviconUrl = `https://drive.google.com/uc?id=${FILE_ID}&export=download&format=png`;


  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Jira Time Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl(faviconUrl);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/*
-----------------------------------------------
00: Server-side config functions for the standalone web app.
-----------------------------------------------
*/
/** Returns the email address of the active user. */
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

/** Returns the current Jira configuration (URL and whether an API key is set). */
function getConfig() {
  const props = getUserProperties();
  return {
    jiraUrl: props.getProperty('JIRA_BASE_URL') || '',
    hasApiKey: !!props.getProperty('JIRA_API_KEY')
  };
}

/** Saves the Jira base URL, trimmed and without trailing slash. */
function saveJiraUrl(url) {
  getUserProperties().setProperty('JIRA_BASE_URL', url.trim().replace(/\/$/, ''));
}

/** Saves the Jira API key. */
function saveApiKey(key) {
  getUserProperties().setProperty('JIRA_API_KEY', key.trim());
}

/** Returns the saved pay-period-end-date anchor (ISO date string), or '' if unset. */
function getPayPeriodEndDate() {
  return getUserProperties().getProperty('PAY_PERIOD_END_DATE') || '';
}

/** Saves the pay-period-end-date anchor. Every other pay period (any year, past or future) is derived from this date at a fixed 14-day cadence. */
function savePayPeriodEndDate(dateStr) {
  getUserProperties().setProperty('PAY_PERIOD_END_DATE', dateStr);
}

/** Returns the saved Utilization-tab cost inputs: {hourlyRate, overheadRate (decimal fraction)}. */
function getUtilRates() {
  const props = getUserProperties();
  return {
    hourlyRate: parseFloat(props.getProperty('UTIL_HOURLY_RATE')) || 0,
    overheadRate: parseFloat(props.getProperty('UTIL_OVERHEAD_RATE')) || 0
  };
}

/** Saves the Utilization-tab cost inputs. overheadRate is a decimal fraction (e.g. 0.40 for 40%). */
function saveUtilRates(hourlyRate, overheadRate) {
  getUserProperties().setProperty('UTIL_HOURLY_RATE', String(hourlyRate));
  getUserProperties().setProperty('UTIL_OVERHEAD_RATE', String(overheadRate));
}

/**
 * Tests connectivity to Jira using the given URL/API key, falling back to the
 * saved config for whichever of the two is blank. Returns {ok, displayName}
 * on success or {ok: false, error} on failure — never throws.
 */
function testJiraConnection(url, apiKey) {
  const props = getUserProperties();
  const jiraUrl = (url && url.trim()) ? url.trim().replace(/\/$/, '') : (props.getProperty('JIRA_BASE_URL') || '');
  const jiraApiKey = (apiKey && apiKey.trim()) ? apiKey.trim() : (props.getProperty('JIRA_API_KEY') || '');
  if (!jiraUrl || !jiraApiKey) return { ok: false, error: 'Jira URL and API key are required.' };

  const email = Session.getActiveUser().getEmail();
  const authHeader = 'Basic ' + Utilities.base64Encode(`${email}:${jiraApiKey}`);
  const options = { headers: { Authorization: authHeader }, method: 'get', muteHttpExceptions: true };

  try {
    const response = UrlFetchApp.fetch(`${jiraUrl}/rest/api/3/myself`, options);
    if (response.getResponseCode() >= 400) {
      return { ok: false, error: `Jira API error (${response.getResponseCode()}): ${response.getContentText()}` };
    }
    const data = JSON.parse(response.getContentText());
    return { ok: true, displayName: data.displayName || email };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/** Returns the allocation data (project-to-hours mapping) from user properties. */
function getAllocation() {
  const raw = getUserProperties().getProperty('ALLOCATION');
  return raw ? JSON.parse(raw) : [];
}

/** Saves the allocation data as a JSON string in user properties. */
function saveAllocation(rows) {
  getUserProperties().setProperty('ALLOCATION', JSON.stringify(rows));
}

/*
-----------------------------------------------
00: Fetch active Jira issues for the standalone web app.
-----------------------------------------------
*/
/**
 * Fetches active Jira issues assigned to or watched by the current user.
 * Returns an array of issues with dropdownValue, key, link, projectKey, name,
 * status, project, and timeLogged (in hours).
 */
function getJiraIssues() {
  const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
  const JIRA_API_KEY = getUserProperties().getProperty('JIRA_API_KEY');
  if (!JIRA_URL || !JIRA_API_KEY) throw new Error('Jira URL and API key must be configured in the Config tab.');
  const USER_EMAIL = Session.getActiveUser().getEmail();
  const authHeader = getAuthHeader_();
  const BASE_ENDPOINT = `${JIRA_URL}/rest/api/3/search/jql?jql=(assignee=currentUser()+OR+watcher=currentUser())+AND+issuetype+IN+(Story,Task,Sub-Task)+AND+(status!=Done+OR+(status=Done+AND+updated%3E=-7d))+ORDER+BY+key+ASC&fields=key,summary,status,project&maxResults=100`;
  const options = { headers: { Authorization: authHeader }, method: 'get', muteHttpExceptions: true };

  let allIssues = [];
  let nextPageToken = null;
  do {
    const endpoint = nextPageToken ? `${BASE_ENDPOINT}&nextPageToken=${nextPageToken}` : BASE_ENDPOINT;
    const response = UrlFetchApp.fetch(endpoint, options);
    if (response.getResponseCode() >= 400) throw new Error(`Jira API error (${response.getResponseCode()}): ${response.getContentText()}`);
    const data = JSON.parse(response.getContentText());
    if (data.issues) allIssues = allIssues.concat(data.issues);
    nextPageToken = data.nextPageToken || null;
  } while (nextPageToken);

  const filtered = allIssues.filter(issue => {
    const name = issue.fields.project.name;
    return !name.includes('Archive') && !name.includes('Managed Services Internal');
  });

  const worklogTotals = getWorklogTotals_(filtered.map(i => i.key), authHeader, JIRA_URL, USER_EMAIL);

  return filtered.map((issue, idx) => ({
    dropdownValue: `${issue.key} (${issue.fields.summary})`,
    key: issue.key,
    link: `${JIRA_URL}/browse/${issue.key}`,
    projectKey: issue.key.split('-')[0],
    name: issue.fields.summary,
    status: issue.fields.status.name,
    project: issue.fields.project.name,
    timeLogged: Math.round((worklogTotals[idx] / 3600) * 100) / 100
  }));
}

/**
 * Transitions the given issue to the "Done" status in Jira. Looks up the
 * issue's available transitions and executes whichever one targets a status
 * named "Done" (case-insensitive), since the transition ID varies by workflow.
 * Throws if no such transition is available (e.g. the workflow requires
 * additional fields, or has no status literally named "Done").
 */
function markIssueDone(issueKey) {
  const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
  const JIRA_API_KEY = getUserProperties().getProperty('JIRA_API_KEY');
  if (!JIRA_URL || !JIRA_API_KEY) throw new Error('Jira URL and API key must be configured in the Config tab.');
  const authHeader = getAuthHeader_();

  const getOptions = { headers: { Authorization: authHeader }, method: 'get', muteHttpExceptions: true };
  const getResponse = UrlFetchApp.fetch(`${JIRA_URL}/rest/api/3/issue/${issueKey}/transitions`, getOptions);
  if (getResponse.getResponseCode() >= 400) throw new Error(`Jira API error (${getResponse.getResponseCode()}): ${getResponse.getContentText()}`);
  const transitions = JSON.parse(getResponse.getContentText()).transitions || [];
  const doneTransition = transitions.find(t => t.to && t.to.name && t.to.name.toLowerCase() === 'done');
  if (!doneTransition) throw new Error(`No "Done" transition is available for ${issueKey}.`);

  const postOptions = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: authHeader },
    payload: JSON.stringify({ transition: { id: doneTransition.id } }),
    muteHttpExceptions: true
  };
  const postResponse = UrlFetchApp.fetch(`${JIRA_URL}/rest/api/3/issue/${issueKey}/transitions`, postOptions);
  if (postResponse.getResponseCode() !== 204) {
    throw new Error(`Jira API error (${postResponse.getResponseCode()}): ${postResponse.getContentText()}`);
  }
}

/** Workday window: 9:00 AM - 5:00 PM. */
const SCHEDULE_WORKDAY_START_HOUR = 9;
const SCHEDULE_WORKDAY_END_HOUR = 17;

/** Returns 9:00 AM on the next work day (skipping Saturday/Sunday) after the given date. */
function nextScheduleWorkDay_(date) {
  const next = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1, SCHEDULE_WORKDAY_START_HOUR, 0, 0, 0);
  while (next.getDay() === 0 || next.getDay() === 6) {
    next.setDate(next.getDate() + 1);
  }
  return next;
}

/** Returns remaining milliseconds before 5:00 PM on the cursor's calendar day (0 if already past). */
function scheduleDayCapacityMs_(cursor) {
  const dayEnd = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate(), SCHEDULE_WORKDAY_END_HOUR, 0, 0, 0);
  return Math.max(0, dayEnd.getTime() - cursor.getTime());
}

/**
 * Creates Google Calendar events from the provided toSchedule array.
 * Each event's title is "Client Task Time: <Project Name> - <Jira Key>"; the description
 * has the Jira summary, the JIRA link, and a note that the block can be scheduled over.
 * Events are scheduled sequentially starting at startDateTimeIso, capped at 8h/day
 * (9:00 AM-5:00 PM); any remainder rolls to the next work day (weekends skipped).
 * Returns { created: number, startTime: string }
 */
function scheduleCalendarEvents(toSchedule, startDateTimeIso) {
  const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
  const allocation = getAllocation();

  const HEX_TO_EVENT_COLOR = {
    '#828bc2': CalendarApp.EventColor.PALE_BLUE,
    '#55b080': CalendarApp.EventColor.PALE_GREEN,
    '#a75aba': CalendarApp.EventColor.MAUVE,
    '#d6837a': CalendarApp.EventColor.PALE_RED,
    '#e7ba51': CalendarApp.EventColor.YELLOW,
    '#e3683e': CalendarApp.EventColor.ORANGE,
    '#4b99d2': CalendarApp.EventColor.CYAN,
    '#7c7c7c': CalendarApp.EventColor.GRAY,
    '#6e72c3': CalendarApp.EventColor.BLUE,
    '#489160': CalendarApp.EventColor.GREEN,
    '#da5234': CalendarApp.EventColor.RED
  };

  const projectColorMap = Object.fromEntries(
    allocation
      .filter(r => r.projectKey && r.colorHex)
      .map(r => [r.projectKey, HEX_TO_EVENT_COLOR[r.colorHex] || null])
  );

  const projectNameMap = Object.fromEntries(
    allocation
      .filter(r => r.projectKey && r.projectName)
      .map(r => [r.projectKey, r.projectName])
  );

  const startTime = new Date(startDateTimeIso);
  const calendar = CalendarApp.getDefaultCalendar();
  let cursor = new Date(startTime);
  let created = 0;

  toSchedule.forEach(entry => {
    let remainingMs = entry.hours * 60 * 60 * 1000;

    while (remainingMs > 0) {
      const capacityMs = scheduleDayCapacityMs_(cursor);
      if (capacityMs <= 0) {
        cursor = nextScheduleWorkDay_(cursor);
        continue;
      }

      const segmentMs = Math.min(remainingMs, capacityMs);
      const endTime = new Date(cursor.getTime() + segmentMs);
      const projectName = projectNameMap[entry.jiraProject] || entry.jiraProject;
      const event = calendar.createEvent(
        `Client Task Time: ${projectName} - ${entry.key}`,
        cursor,
        endTime,
        { description: `${entry.summary}\n${JIRA_URL}/browse/${entry.key}\nTime blocking for client allocation - schedule over if needed.` }
      );
      const color = projectColorMap[entry.jiraProject];
      if (color) event.setColor(color);
      created++;

      remainingMs -= segmentMs;
      cursor = remainingMs > 0 ? nextScheduleWorkDay_(cursor) : endTime;
    }
  });

  return { created, startTime: startTime.toLocaleTimeString() };
}

/**
 * Imports calendar events for the given date range and returns an array of event objects.
 * Filters events based on their color mapping to valid project keys.
 *
 * @param {string} startDateStr - Start date in "yyyy-MM-dd" format
 * @param {string} endDateStr - End date in "yyyy-MM-dd" format
 * @returns {Array} Array of event objects with title, date, start, end, description, status, projectKey, issueKey, duration
 */
function importCalendarEvents(startDateStr, endDateStr) {
  const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
  const JIRA_API_KEY = getUserProperties().getProperty('JIRA_API_KEY');
  if (!JIRA_URL || !JIRA_API_KEY) throw new Error('Jira URL and API key must be configured in the Config tab.');
  const allocation = getAllocation();
  const HEX_TO_COLOR_ID = {
    '#828bc2': '1', '#55b080': '2', '#a75aba': '3', '#d6837a': '4',
    '#e7ba51': '5', '#e3683e': '6', '#4b99d2': '7', '#7c7c7c': '8',
    '#6e72c3': '9', '#489160': '10', '#da5234': '11'
  };
  const activeRows = allocation.filter(r => r.projectKey && r.colorHex && !r.ignore);
  const colorMap = Object.fromEntries(
    activeRows.map(r => [HEX_TO_COLOR_ID[r.colorHex], r.projectKey])
  );
  const validProjectKeys = new Set(activeRows.map(r => r.projectKey));

  const CALENDAR_ID = Session.getActiveUser().getEmail();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const tz = Session.getScriptTimeZone();

  const startDate = new Date(startDateStr + 'T00:00:00');
  const endDate = new Date(endDateStr + 'T23:59:59');

  const events = calendar.getEvents(startDate, endDate).map(event => {
    if (event.getMyStatus() === CalendarApp.GuestStatus.NO) return null;
    const colorNum = event.getColor();
    const projectKey = colorMap[colorNum] || '';
    if (!validProjectKeys.has(projectKey)) return null;
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    const durationHours = (endTime - startTime) / 3600000;
    const description = event.getDescription();
    const match = description ? description.match(/^[^\n_]+/) : null;
    return {
      title: event.getTitle(),
      date: Utilities.formatDate(startTime, tz, 'yyyy-MM-dd'),
      start: Utilities.formatDate(startTime, tz, 'hh:mm a'),
      end: Utilities.formatDate(endTime, tz, 'hh:mm a'),
      description: match ? match[0].trim() : '',
      status: event.getMyStatus(),
      projectKey,
      issueKey: '',
      duration: Math.round(durationHours * 4) / 4
    };
  }).filter(Boolean);

  return events;
}

/**
 * Sends time entries to Jira as worklogs.
 * @param {Array<{date: string, issueKey: string, startTime: string, durationHours: number}>} entries
 * @returns {{succeeded: number, failed: number, errors: Array<string>}}
 */
function sendTimeEntries(entries) {
  const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
  const authHeader = getAuthHeader_();
  const UTC_FORMAT = "yyyy-MM-dd'T'HH:mm:ss'.000+0000'";
  let succeeded = 0, failed = 0;
  const errors = [];

  entries.forEach(entry => {
    // Skip entries with no issue key
    if (!entry.issueKey) return;

    // Manually parse startTime (hh:mm a format) since V8 doesn't correctly parse combined datetime strings
    // Extract AM/PM and time portion
    const timeParts = entry.startTime.trim().split(' ');
    const ampm = timeParts[timeParts.length - 1].toUpperCase();
    const timePortion = timeParts.slice(0, -1).join(' '); // In case time has spaces
    const [hoursStr, minutesStr] = timePortion.split(':');
    let hours = parseInt(hoursStr, 10);
    const minutes = parseInt(minutesStr, 10);

    // Adjust for AM/PM
    if (ampm === 'PM' && hours !== 12) {
      hours += 12;
    } else if (ampm === 'AM' && hours === 12) {
      hours = 0;
    }

    // Create a Date object from the date string and set the time
    const combined = new Date(entry.date + 'T00:00:00');
    combined.setHours(hours, minutes, 0, 0);

    // Format to UTC
    const utcString = Utilities.formatDate(combined, 'Etc/GMT', UTC_FORMAT);
    const durationMinutes = Math.round(entry.durationHours * 60);
    const options = {
      method: 'post',
      headers: { Authorization: authHeader, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ started: utcString, timeSpent: `${durationMinutes}m` }),
      muteHttpExceptions: true
    };
    try {
      const resp = UrlFetchApp.fetch(`${JIRA_URL}/rest/api/3/issue/${entry.issueKey}/worklog`, options);
      if (resp.getResponseCode() === 201) {
        succeeded++;
      } else {
        failed++;
        errors.push(`${entry.issueKey}: HTTP ${resp.getResponseCode()}`);
      }
    } catch (e) {
      failed++;
      errors.push(`${entry.issueKey}: ${e.message}`);
    }
  });

  return { succeeded, failed, errors };
}

/**
 * Returns the list of project keys from the allocation data.
 * Derives from getAllocation() which contains the current user's project allocation.
 */
function getProjectKeys() {
  return [...new Set(getAllocation().map(r => r.projectKey).filter(Boolean))];
}

/*
-----------------------------------------------
04: Fetch worklogs logged by the current user since Jan 1 of the current year.
Returns an array of objects with projectKey, issueKey, summary, timeSpent,
started (ISO), hours, and month.
-----------------------------------------------
*/
/**
 * Fetches worklogs for the current user from Jan 1 of the current year.
 * Returns an array of {projectKey, issueKey, summary, timeSpent, started, hours, month}.
 * Handles paginated issue results (nextPageToken) and paginated worklog fields (fetchAll).
 */
function getWorklogs(year) {
  const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
  const JIRA_API_KEY = getUserProperties().getProperty('JIRA_API_KEY');
  if (!JIRA_URL || !JIRA_API_KEY) throw new Error('Jira URL and API key must be configured in the Config tab.');
  const authHeader = getAuthHeader_();
  const userEmail = Session.getActiveUser().getEmail();
  const targetYear = year || new Date().getFullYear();
  const fromDate = `${targetYear}-01-01`;
  const toDate = `${targetYear}-12-31`;
  const fromTimestamp = new Date(targetYear, 0, 1).getTime();
  const toTimestamp = new Date(targetYear, 11, 31, 23, 59, 59, 999).getTime();
  const JQL = encodeURIComponent(`worklogAuthor=currentUser() AND worklogDate >= ${fromDate} AND worklogDate <= ${toDate}`);
  const BASE_ENDPOINT = `${JIRA_URL}/rest/api/3/search/jql?fields=key,summary,worklog,project&jql=${JQL}&maxResults=100`;
  const fetchOpts = { headers: { Authorization: authHeader }, method: 'get', muteHttpExceptions: true };

  let allIssues = [];
  let nextPageToken = null;
  do {
    const endpoint = nextPageToken ? `${BASE_ENDPOINT}&nextPageToken=${nextPageToken}` : BASE_ENDPOINT;
    const response = UrlFetchApp.fetch(endpoint, fetchOpts);
    if (response.getResponseCode() >= 400) throw new Error(`Jira API error (${response.getResponseCode()}): ${response.getContentText()}`);
    const data = JSON.parse(response.getContentText());
    if (data.issues) allIssues = allIssues.concat(data.issues);
    nextPageToken = data.nextPageToken || null;
  } while (nextPageToken);

  const extraFetches = [];
  allIssues.forEach((issue, idx) => {
    const wl = issue.fields.worklog;
    if (!wl) return;
    for (let s = wl.worklogs ? wl.worklogs.length : 0; s < (wl.total || 0); s += 100) {
      extraFetches.push({ issueIdx: idx, startAt: s });
    }
  });
  // Fetch extra worklog pages sequentially — fetchAll throws on network-level
  // errors (e.g. rate limiting) and can't be caught per-request.
  for (const f of extraFetches) {
    try {
      const resp = UrlFetchApp.fetch(
        `${JIRA_URL}/rest/api/3/issue/${allIssues[f.issueIdx].key}/worklog?startAt=${f.startAt}&maxResults=100`,
        { headers: { Authorization: authHeader }, method: 'get', muteHttpExceptions: true }
      );
      if (resp.getResponseCode() >= 400) continue;
      const data = JSON.parse(resp.getContentText());
      if (!data.worklogs) continue;
      const wl = allIssues[f.issueIdx].fields.worklog;
      wl.worklogs = (wl.worklogs || []).concat(data.worklogs);
    } catch (e) {
      Logger.log(`getWorklogs: skipping extra page for ${allIssues[f.issueIdx].key} startAt=${f.startAt}: ${e.message}`);
    }
  }

  const rows = [];
  allIssues.forEach(issue => {
    const projectKey = issue.fields.project.key;
    const issueKey = issue.key;
    const summary = issue.fields.summary;
    (issue.fields.worklog && issue.fields.worklog.worklogs || []).forEach(log => {
      if (!log.author || log.author.emailAddress !== userEmail) return;
      const startedDate = new Date(log.started || '');
      if (!log.started || startedDate.getTime() < fromTimestamp || startedDate.getTime() > toTimestamp) return;
      rows.push({
        projectKey, issueKey, summary,
        timeSpent: log.timeSpent || '',
        started: log.started,
        hours: parseTimeSpentHours_(log.timeSpent || ''),
        month: `${startedDate.getFullYear()}-${String(startedDate.getMonth() + 1).padStart(2, '0')}`
      });
    });
  });
  return rows;
}

/*
-----------------------------------------------
05: Allocation tab — manual monthly hours-allocation entry per project, per year.
-----------------------------------------------
*/
/**
 * Returns the saved allocation-grid values for a given year:
 * { [projectKey]: { [month 1-12]: number } }. Sparse — only cells the user
 * has previously saved are present.
 */
function getAllocationValues_(year) {
  const raw = getUserProperties().getProperty('ALLOCATION_VALUES');
  const parsed = raw ? JSON.parse(raw) : {};
  return parsed[year] || {};
}

/**
 * Returns the data needed to render the Allocation tab's grid for a year:
 * - projects: sorted Jira project keys with at least one worklog that year,
 *   excluding any project marked "Ignore" in the Config tab's allocation table.
 * - values: previously saved allocation numbers for that year, see getAllocationValues_.
 */
function getAllocationTabData(year) {
  const rows = getWorklogs(year);
  const ignoredKeys = new Set(getAllocation().filter(r => r.ignore).map(r => r.projectKey));
  const projects = [...new Set(rows.map(r => r.projectKey))]
    .filter(key => !ignoredKeys.has(key))
    .sort();
  return { projects, values: getAllocationValues_(year) };
}

/**
 * Saves the Allocation tab's grid for a given year, replacing only that
 * year's entry in the ALLOCATION_VALUES user property (other years untouched).
 * values: { [projectKey]: { [month 1-12]: number } }.
 */
function saveAllocationGrid(year, values) {
  const raw = getUserProperties().getProperty('ALLOCATION_VALUES');
  const parsed = raw ? JSON.parse(raw) : {};
  parsed[year] = values;
  getUserProperties().setProperty('ALLOCATION_VALUES', JSON.stringify(parsed));
}

/*
-----------------------------------------------
06: Utilization tab — worklog caching, pay-period math, and the
    utilization/revenue/cost/profit calculation.
-----------------------------------------------
*/
/** Returns cached worklog rows for a year (same shape as getWorklogs), or null if never collected. */
function getCachedWorklogs_(year) {
  const raw = getUserProperties().getProperty('WORKLOG_CACHE');
  const parsed = raw ? JSON.parse(raw) : {};
  return parsed[year] || null;
}

/** Stores worklog rows for a year, replacing only that year's entry (other years untouched). */
function setCachedWorklogs_(year, rows) {
  const raw = getUserProperties().getProperty('WORKLOG_CACHE');
  const parsed = raw ? JSON.parse(raw) : {};
  parsed[year] = rows;
  getUserProperties().setProperty('WORKLOG_CACHE', JSON.stringify(parsed));
}

/**
 * Returns worklog rows for a year, serving from WORKLOG_CACHE when already
 * collected. forceRefresh bypasses the cache and re-fetches from Jira via
 * getWorklogs(year), overwriting the cached entry for that year.
 */
function getOrCollectWorklogs(year, forceRefresh) {
  if (!forceRefresh) {
    const cached = getCachedWorklogs_(year);
    if (cached) return cached;
  }
  const rows = getWorklogs(year);
  setCachedWorklogs_(year, rows);
  return rows;
}

const PAY_PERIOD_DAYS_ = 14;

/**
 * Returns all pay-period end dates (as Date objects, ascending) that fall
 * within [Jan 1, Dec 31] of `year`, derived from the PAY_PERIOD_END_DATE
 * anchor at a fixed 14-day cadence. Returns [] if the anchor isn't set.
 */
function getPayPeriodEndDates_(year) {
  const anchorStr = getPayPeriodEndDate();
  if (!anchorStr) return [];
  const anchor = new Date(anchorStr + 'T00:00:00');
  const yearStart = new Date(year, 0, 1);
  const yearEnd = new Date(year, 11, 31, 23, 59, 59, 999);
  const msPerPeriod = PAY_PERIOD_DAYS_ * 24 * 60 * 60 * 1000;

  const periodsFromAnchorToYearStart = Math.ceil((yearStart - anchor) / msPerPeriod);
  const dates = [];
  let n = periodsFromAnchorToYearStart;
  let d = new Date(anchor.getTime() + n * msPerPeriod);
  while (d <= yearEnd) {
    if (d >= yearStart) dates.push(new Date(d));
    n++;
    d = new Date(anchor.getTime() + n * msPerPeriod);
  }
  return dates;
}

/**
 * Returns {payPeriodsToDate, firstPayPeriodEnd, mostRecentPayPeriodEnd} for a
 * year, where "most recent" is the latest period end date <= today. For a
 * past year this is simply that year's last period end date, since all of
 * them are already in the past.
 */
function getPayPeriodSummary_(year) {
  const allDates = getPayPeriodEndDates_(year);
  const today = new Date();
  const eligible = allDates.filter(d => d <= today);
  if (eligible.length === 0) return { payPeriodsToDate: 0, firstPayPeriodEnd: null, mostRecentPayPeriodEnd: null };
  return {
    payPeriodsToDate: eligible.length,
    firstPayPeriodEnd: allDates[0],
    mostRecentPayPeriodEnd: eligible[eligible.length - 1]
  };
}

/**
 * Returns everything the Utilization tab needs to render for a year:
 * - projects: sorted project keys present in the year's worklog rows.
 * - utilization: { cells: { [projectKey]: { [month 1-12]: pct|null } }, rowTotals: { [month]: pct|null },
 *   colTotals: { [projectKey]: pct|null }, grandTotal: pct|null }. A cell is null when the project has
 *   no allocation value for that month; row/col/grand totals are sum(hours)/sum(allocation) over the
 *   contributing (non-null) cells, not an average of percentages.
 * - payPeriods: { firstEnd, mostRecentEnd (ISO date strings or null), payPeriodsToDate, totalCostHours }.
 * - revenue: { hoursThruMostRecent, grossRevenue, unratedProjects } — unratedProjects lists project keys
 *   excluded entirely from both figures because they have no rate configured in the Config tab.
 * - cost: { hourlyRate, overheadRate, hourlyCost, grossCost }.
 * - profit: { grossProfit, personalProfitMargin }.
 */
function calculateUtilization(year, forceRefresh) {
  const rows = getOrCollectWorklogs(year, forceRefresh);
  const allocationRows = getAllocation();
  const rateByProject = Object.fromEntries(allocationRows.filter(r => r.projectKey).map(r => [r.projectKey, r.rate]));
  const allocValues = getAllocationValues_(year); // { projectKey: { month: hours } }

  const projects = [...new Set(rows.map(r => r.projectKey))].sort();

  const hoursByProjectMonth = {};
  rows.forEach(r => {
    const m = parseInt(r.month.split('-')[1], 10);
    hoursByProjectMonth[r.projectKey] = hoursByProjectMonth[r.projectKey] || {};
    hoursByProjectMonth[r.projectKey][m] = (hoursByProjectMonth[r.projectKey][m] || 0) + r.hours;
  });

  const cells = {};
  const rowHourSum = {}, rowAllocSum = {};
  const colHourSum = {}, colAllocSum = {};
  let grandHourSum = 0, grandAllocSum = 0;
  projects.forEach(p => {
    cells[p] = {};
    for (let m = 1; m <= 12; m++) {
      const allocated = allocValues[p] && allocValues[p][m];
      const worked = (hoursByProjectMonth[p] && hoursByProjectMonth[p][m]) || 0;
      if (allocated == null || allocated === 0) {
        cells[p][m] = null;
        continue;
      }
      cells[p][m] = Math.round((worked / allocated) * 1000) / 10;
      rowHourSum[m] = (rowHourSum[m] || 0) + worked;
      rowAllocSum[m] = (rowAllocSum[m] || 0) + allocated;
      colHourSum[p] = (colHourSum[p] || 0) + worked;
      colAllocSum[p] = (colAllocSum[p] || 0) + allocated;
      grandHourSum += worked;
      grandAllocSum += allocated;
    }
  });
  const pct_ = (h, a) => (a === 0 || a == null) ? null : Math.round((h / a) * 1000) / 10;
  const rowTotals = {};
  for (let m = 1; m <= 12; m++) rowTotals[m] = pct_(rowHourSum[m] || 0, rowAllocSum[m] || 0);
  const colTotals = {};
  projects.forEach(p => colTotals[p] = pct_(colHourSum[p] || 0, colAllocSum[p] || 0));
  const grandTotal = pct_(grandHourSum, grandAllocSum);

  const payPeriodInfo = getPayPeriodSummary_(year);
  const cutoff = payPeriodInfo.mostRecentPayPeriodEnd; // Date or null

  let hoursThruMostRecent = 0, grossRevenue = 0;
  const unratedProjects = new Set();
  if (cutoff) {
    rows.forEach(r => {
      const rate = rateByProject[r.projectKey];
      if (rate == null) { unratedProjects.add(r.projectKey); return; }
      const started = new Date(r.started);
      if (started <= cutoff) {
        hoursThruMostRecent += r.hours;
        grossRevenue += r.hours * rate;
      }
    });
  }

  const { hourlyRate, overheadRate } = getUtilRates();
  const hourlyCost = hourlyRate * (1 + overheadRate);
  const totalCostHours = payPeriodInfo.payPeriodsToDate * 80;
  const grossCost = hourlyCost * totalCostHours;
  const grossProfit = grossRevenue - grossCost;
  const personalProfitMargin = grossRevenue === 0 ? null : Math.round((grossProfit / grossRevenue) * 1000) / 10;

  return {
    projects,
    utilization: { cells, rowTotals, colTotals, grandTotal },
    payPeriods: {
      firstEnd: payPeriodInfo.firstPayPeriodEnd ? payPeriodInfo.firstPayPeriodEnd.toISOString().slice(0, 10) : null,
      mostRecentEnd: cutoff ? cutoff.toISOString().slice(0, 10) : null,
      payPeriodsToDate: payPeriodInfo.payPeriodsToDate,
      totalCostHours
    },
    revenue: { hoursThruMostRecent, grossRevenue, unratedProjects: [...unratedProjects] },
    cost: { hourlyRate, overheadRate, hourlyCost, grossCost },
    profit: { grossProfit, personalProfitMargin }
  };
}

/**
 * Called by the Create Issue tab on submit. Posts the new issue to Jira
 * via REST API. Returns a status string displayed in the tab.
 */
function makeJira(formData) {
    const JIRA_URL = getUserProperties().getProperty('JIRA_BASE_URL');
    const authHeader = getAuthHeader_();
    const email = Session.getActiveUser().getEmail();

    // Resolve email to Jira accountId
    let accountId = null;
    try {
        const userSearch = UrlFetchApp.fetch(
            `${JIRA_URL}/rest/api/3/user/search?query=${encodeURIComponent(email)}`,
            { headers: { Authorization: authHeader }, muteHttpExceptions: true }
        );
        const userResults = JSON.parse(userSearch.getContentText());
        if (Array.isArray(userResults) && userResults.length > 0) {
            accountId = userResults[0].accountId;
        }
    } catch (e) {
        Logger.log('Could not resolve Jira accountId: ' + e.message);
    }

    const payload = {
        fields: {
            project: { key: formData.input1 },
            summary: formData.input4,
            description: {
                type: 'doc',
                version: 1,
                content: [{ type: 'paragraph', content: [{ type: 'text', text: formData.notes || '' }] }]
            },
            issuetype: { name: formData.input6 },
            priority: { id: formData.input7 },
            customfield_10201: formData.input2,
            customfield_10878: { value: formData.input3 },
            ...(accountId && { assignee: { accountId } })
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
            if (response.getResponseCode() >= 400) {
                Logger.log(`getWorklogTotals_: HTTP ${response.getResponseCode()} for ${issueKeys[keyIndex]} — skipping`);
                return;
            }
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

/**
 * Converts a Jira timeSpent string (e.g. "1h 30m", "45m", "2h") to decimal hours.
 * Minutes are divided by 60 and added to whole hours (e.g. 25m → 0.4167).
 */
function parseTimeSpentHours_(timeSpent) {
    if (!timeSpent) return 0;
    const h = timeSpent.match(/(\d+)h/);
    const m = timeSpent.match(/(\d+)m/);
    return (h ? parseInt(h[1]) : 0) + (m ? parseInt(m[1]) / 60 : 0);
}