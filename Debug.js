/*
  Debug.gs — Logging and validation helpers for Jira Time Tracker.
  Run these functions directly from the Apps Script editor.
  Results appear in View → Logs (or Execution Log).
*/

/**
 * Validates the importCalendarEvents import pipeline for a given date range.
 * Logs every calendar event found and, for any event that would be filtered,
 * explains exactly why.
 *
 * Run from the Apps Script editor with no arguments to inspect the current week.
 * Or call with explicit dates: validateImportCalendarEvents('2026-08-11', '2026-08-15')
 */
function validateImportCalendarEvents(startDateStr, endDateStr) {
  // Default to current Mon–Fri if no args provided
  if (!startDateStr || !endDateStr) {
    const today = new Date();
    const day = today.getDay();
    const monday = new Date(today);
    monday.setDate(today.getDate() - (day === 0 ? 6 : day - 1));
    const friday = new Date(monday);
    friday.setDate(monday.getDate() + 4);
    const fmt = d => Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    startDateStr = fmt(monday);
    endDateStr = fmt(friday);
  }

  const tz = Session.getScriptTimeZone();
  const log = lines => lines.forEach(l => Logger.log(l));

  Logger.log('='.repeat(70));
  Logger.log('validateImportCalendarEvents');
  Logger.log('Date range: %s → %s', startDateStr, endDateStr);
  Logger.log('='.repeat(70));

  // ── 1. Allocation config ──────────────────────────────────────────────────
  const allocation = getAllocation();
  Logger.log('\n── Allocation config (%s rows) ──', allocation.length);
  if (allocation.length === 0) {
    Logger.log('  (none — open the Config tab and save your allocation)');
  } else {
    allocation.forEach((r, i) => {
      Logger.log('  [%s] colorHex=%s  projectKey=%s  ignore=%s',
        i, r.colorHex || '(missing)', r.projectKey || '(missing)', r.ignore ? 'YES' : 'no');
    });
  }

  // ── 2. Derived colorMap ───────────────────────────────────────────────────
  const HEX_TO_COLOR_ID = {
    '#828bc2': '1',  // Lavender
    '#55b080': '2',  // Sage
    '#a75aba': '3',  // Grape
    '#d6837a': '4',  // Flamingo
    '#e7ba51': '5',  // Banana
    '#e3683e': '6',  // Tangerine
    '#4b99d2': '7',  // Peacock
    '#7c7c7c': '8',  // Graphite
    '#6e72c3': '9',  // Blueberry
    '#489160': '10', // Basil
    '#da5234': '11'  // Tomato
  };
  const COLOR_ID_TO_NAME = {
    '': '(default calendar color)',
    '1': 'Lavender', '2': 'Sage', '3': 'Grape', '4': 'Flamingo',
    '5': 'Banana', '6': 'Tangerine', '7': 'Peacock', '8': 'Graphite',
    '9': 'Blueberry', '10': 'Basil', '11': 'Tomato'
  };

  const skippedRows = { missingColorHex: [], missingProjectKey: [], ignored: [] };
  allocation.forEach(r => {
    if (!r.projectKey) { skippedRows.missingProjectKey.push(r); return; }
    if (!r.colorHex)   { skippedRows.missingColorHex.push(r); return; }
    if (r.ignore)      { skippedRows.ignored.push(r); }
  });

  const activeRows = allocation.filter(r => r.projectKey && r.colorHex && !r.ignore);
  const colorMap = Object.fromEntries(activeRows.map(r => [HEX_TO_COLOR_ID[r.colorHex], r.projectKey]));
  const validProjectKeys = new Set(activeRows.map(r => r.projectKey));

  Logger.log('\n── Active colorMap (color ID → project key) ──');
  if (Object.keys(colorMap).length === 0) {
    Logger.log('  (empty — no active, non-ignored rows with colorHex set)');
  } else {
    Object.entries(colorMap).forEach(([id, pk]) => {
      Logger.log('  Color ID %s (%s) → %s', id, COLOR_ID_TO_NAME[id] || '?', pk);
    });
  }

  if (skippedRows.ignored.length) {
    Logger.log('\n  Ignored allocation rows (events with these colors will be excluded):');
    skippedRows.ignored.forEach(r => Logger.log('    colorHex=%s  projectKey=%s', r.colorHex, r.projectKey));
  }
  if (skippedRows.missingColorHex.length) {
    Logger.log('\n  WARNING: allocation rows missing colorHex (will not match any event):');
    skippedRows.missingColorHex.forEach(r => Logger.log('    projectKey=%s', r.projectKey));
  }
  if (skippedRows.missingProjectKey.length) {
    Logger.log('\n  WARNING: allocation rows missing projectKey (skipped):');
    skippedRows.missingProjectKey.forEach(r => Logger.log('    colorHex=%s', r.colorHex));
  }

  // ── 3. Raw calendar events ────────────────────────────────────────────────
  const CALENDAR_ID = Session.getActiveUser().getEmail();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const startDate = new Date(startDateStr + 'T00:00:00');
  const endDate   = new Date(endDateStr   + 'T23:59:59');
  const rawEvents = calendar.getEvents(startDate, endDate);

  Logger.log('\n── Raw calendar events found: %s ──', rawEvents.length);

  // ── 4. Per-event filter trace ─────────────────────────────────────────────
  const results = { passed: [], declined: [], noColorMatch: [] };

  rawEvents.forEach((event, i) => {
    const title    = event.getTitle();
    const colorNum = event.getColor();
    const colorName = COLOR_ID_TO_NAME[colorNum] !== undefined
      ? COLOR_ID_TO_NAME[colorNum]
      : 'unknown (' + colorNum + ')';
    const status   = event.getMyStatus();
    const dateStr  = Utilities.formatDate(event.getStartTime(), tz, 'yyyy-MM-dd');

    const prefix = '[' + (i + 1) + '] ' + dateStr + '  "' + title + '"';
    const colorInfo = '  color=' + colorNum + ' (' + colorName + ')  status=' + status;

    if (status === CalendarApp.GuestStatus.NO) {
      Logger.log('\nFILTERED — declined (GuestStatus.NO)');
      Logger.log('  ' + prefix);
      Logger.log(colorInfo);
      results.declined.push(title);
      return;
    }

    const matchedProject = colorMap[colorNum] || '';
    if (!validProjectKeys.has(matchedProject)) {
      const reason = colorNum === ''
        ? 'event has default calendar color (no color set)'
        : 'color ID ' + colorNum + ' (' + colorName + ') is not mapped to any active allocation row';
      Logger.log('\nFILTERED — color not matched: ' + reason);
      Logger.log('  ' + prefix);
      Logger.log(colorInfo);
      results.noColorMatch.push({ title, colorNum, colorName });
      return;
    }

    Logger.log('\nPASSED → project=' + matchedProject);
    Logger.log('  ' + prefix);
    Logger.log(colorInfo);
    results.passed.push({ title, projectKey: matchedProject });
  });

  // ── 5. Summary ────────────────────────────────────────────────────────────
  Logger.log('\n' + '='.repeat(70));
  Logger.log('SUMMARY');
  Logger.log('  Total events in range : %s', rawEvents.length);
  Logger.log('  Passed (imported)     : %s', results.passed.length);
  Logger.log('  Filtered — declined   : %s', results.declined.length);
  Logger.log('  Filtered — no color   : %s', results.noColorMatch.length);

  if (results.noColorMatch.length) {
    Logger.log('\n  Tip: events filtered for color mismatch — their color IDs and what to do:');
    const seen = {};
    results.noColorMatch.forEach(e => {
      if (seen[e.colorNum]) return;
      seen[e.colorNum] = true;
      if (e.colorNum === '') {
        Logger.log('    Color ID "" (default) — change the event color in Google Calendar, or add an allocation row with a matching hex.');
      } else {
        Logger.log('    Color ID %s (%s) — add an allocation row in the Config tab using that color.', e.colorNum, e.colorName);
      }
    });
  }
  Logger.log('='.repeat(70));
}