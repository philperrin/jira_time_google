# Pre-Deployment Test Suite Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a GAS-native automated test suite (`Tests.js`) covering `Code.js`'s server-side business logic, run manually from the Apps Script editor before each deployment, plus a `CLAUDE.md` pointer to it and to the manual smoke-test checklist.

**Architecture:** A single new file, `Tests.js`, containing a minimal hand-rolled test runner/assertion library, temporary-global-reassignment mocks for `UrlFetchApp`/`CalendarApp`/`PropertiesService`/`Session`, and test cases grouped by the same numbered section comments `Code.js` already uses. No refactor of `Code.js`. No build tooling, no npm, no CI.

**Tech Stack:** Google Apps Script (plain JS), no external libraries.

**Spec:** `docs/superpowers/specs/SPEC-2026-09-23-test-suite.md`

## Global Constraints

- No npm, no build system, no CI — `Tests.js` is plain GAS-style JS, run manually from the Apps Script editor's function picker, same as `Debug.js`'s `validateImportCalendarEvents` (spec §1, §2).
- Every mock installer (`withMockUrlFetch`, `withMockCalendar`, `withMockProperties`, `withMockSession`) must restore the real global service in a `finally` block, so a throwing test can never leak a mock into a later test (spec §3).
- Scope is `Code.js` only — `JavaScript.html` client-side logic is explicitly out of scope for this suite (spec §1).
- Single file `Tests.js`; test cases grouped under section comments mirroring `Code.js`'s existing numbered sections (config, issues, scheduling, calendar import, worklogs, allocation, utilization, create issue) (spec §2, §4).
- The manual smoke-test checklist already lives in the spec doc (§5) as prose, not code — no new checklist file.
- `CLAUDE.md`'s Development workflow section gets exactly one added line naming `runAllTests` and the manual checklist as a pre-deployment gate (spec §6).
- `Code.js` is not modified by this plan — only additive files/mocks.

## Review Focus

- **Mock restoration under a failing assertion inside the mocked block** — if the function under test throws (e.g. an unexpected error, not just a failed `assertEquals`), the real `UrlFetchApp`/`CalendarApp`/`PropertiesService`/`Session` must still be restored before the next test runs. Pin this with a dedicated meta-test in Task 1.
- **`assertThrows`'s failure-detection logic** — a naive implementation that compares an error's message against a sentinel string can misclassify a real error that happens to share that message. Task 1 implements it with an explicit boolean flag instead, and a test proves it correctly fails when `fn` does *not* throw.
- **`getWorklogs` pagination interacting with year-boundary filtering** — worklogs fetched from a second/extra page must still be filtered to `[Jan 1, Dec 31]` of the requested year, not just the first page. Task 5 adds a fixture where an extra-paginated worklog falls outside the requested year and confirms it's excluded.
- **`calculateUtilization`'s zero-vs-unset allocation guard** — `allocated == null` (never entered) and `allocated === 0` (explicitly entered as zero) must both produce a `null` cell without a divide-by-zero, and must be tested as two distinct fixtures, not one. Task 7 covers both explicitly.
- **`sendTimeEntries` day-boundary UTC conversion** — a local start time late enough in the day that its UTC equivalent falls on the *next* calendar date must still post with the UTC date/time Jira expects, not the local date. Task 5 adds a fixture crossing this boundary (e.g. a `11:30 PM` local start).

---

### Task 1: Test runner, assertions, mocks, and Config-section tests

**Files:**
- Create: `Tests.js`

**Interfaces:**
- Produces: `test(name, fn)` — registers a test case. `runAllTests()` — runs every registered case, logs a pass/fail summary via `Logger.log`. `assertEquals(actual, expected, msg)`, `assertTrue(cond, msg)`, `assertThrows(fn, msg)` — throw on failure. `mockResponse_(code, body)` — returns a fake `HTTPResponse`-shaped object. `withMockUrlFetch(handlers, fn)`, `withMockCalendar(fakeCalendar, fn)`, `withMockProperties(initial, fn)`, `withMockSession(email, tz, fn)` — install a mock global for the duration of `fn`, always restoring the real global afterward, and are consumed by every later task.

- [ ] **Step 1: Create `Tests.js` with the runner and assertions**

```js
/*
  Tests.js — GAS-native test suite for Code.js.
  Run `runAllTests` from the Apps Script editor's function picker.
  Results appear in View → Logs (or Execution Log).
*/

const TEST_CASES = [];

function test(name, fn) {
  TEST_CASES.push({ name, fn });
}

function runAllTests() {
  const results = TEST_CASES.map(({ name, fn }) => {
    try {
      fn();
      return { name, pass: true };
    } catch (e) {
      return { name, pass: false, error: e.message };
    }
  });
  const passed = results.filter(r => r.pass).length;
  Logger.log(`${passed}/${results.length} passed`);
  results.filter(r => !r.pass).forEach(r => Logger.log(`FAIL: ${r.name} — ${r.error}`));
}

function assertEquals(actual, expected, msg) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) {
    throw new Error(`${msg ? msg + ': ' : ''}expected ${e}, got ${a}`);
  }
}

function assertTrue(cond, msg) {
  if (!cond) throw new Error(msg || 'expected true');
}

function assertThrows(fn, msg) {
  let threw = false;
  try {
    fn();
  } catch (e) {
    threw = true;
  }
  if (!threw) throw new Error(msg || 'expected fn to throw, it did not');
}
```

- [ ] **Step 2: Add the GAS-service mocks**

```js
/*
-----------------------------------------------
Mocks — temporary global overrides, always restored via finally.
-----------------------------------------------
*/
function mockResponse_(code, body) {
  return { getResponseCode: () => code, getContentText: () => body };
}

function withMockUrlFetch(handlers, fn) {
  const real = UrlFetchApp;
  UrlFetchApp = {
    fetch: handlers.fetch || (() => mockResponse_(200, '{}')),
    fetchAll: handlers.fetchAll || (requests => requests.map(() => mockResponse_(200, '{}')))
  };
  try {
    fn();
  } finally {
    UrlFetchApp = real;
  }
}

function withMockCalendar(fakeCalendar, fn) {
  const real = CalendarApp;
  CalendarApp = {
    getDefaultCalendar: () => fakeCalendar,
    getCalendarById: () => fakeCalendar,
    EventColor: real.EventColor,
    GuestStatus: real.GuestStatus
  };
  try {
    fn();
  } finally {
    CalendarApp = real;
  }
}

function withMockProperties(initial, fn) {
  const store = Object.assign({}, initial);
  const real = PropertiesService;
  PropertiesService = {
    getUserProperties: () => ({
      getProperty: k => (k in store ? store[k] : null),
      setProperty: (k, v) => { store[k] = v; }
    })
  };
  try {
    fn();
  } finally {
    PropertiesService = real;
  }
}

function withMockSession(email, tz, fn) {
  const real = Session;
  Session = {
    getActiveUser: () => ({ getEmail: () => email }),
    getScriptTimeZone: () => tz || real.getScriptTimeZone()
  };
  try {
    fn();
  } finally {
    Session = real;
  }
}
```

- [ ] **Step 3: Add the mock-restoration meta-test (Review Focus item 1)**

```js
/*
-----------------------------------------------
Meta: mock infrastructure itself
-----------------------------------------------
*/
test('withMockProperties restores the real PropertiesService even if fn throws', () => {
  const real = PropertiesService;
  assertThrows(() => {
    withMockProperties({}, () => {
      throw new Error('boom');
    });
  });
  assertTrue(PropertiesService === real, 'real PropertiesService must be restored after a throw');
});

test('assertThrows fails when fn does not throw', () => {
  assertThrows(() => {
    assertThrows(() => { /* does not throw */ });
  }, 'assertThrows should itself throw when the wrapped fn does not throw');
});
```

- [ ] **Step 4: Add Config-section tests (spec §4 "Config")**

```js
/*
-----------------------------------------------
00: Config
-----------------------------------------------
*/
test('saveJiraUrl trims whitespace and strips a trailing slash', () => {
  withMockProperties({}, () => {
    saveJiraUrl('  https://example.atlassian.net/  ');
    assertEquals(PropertiesService.getUserProperties().getProperty('JIRA_BASE_URL'), 'https://example.atlassian.net');
  });
});

test('getConfig reflects saved URL and hasApiKey', () => {
  withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
    assertEquals(getConfig(), { jiraUrl: 'https://example.atlassian.net', hasApiKey: true });
  });
  withMockProperties({}, () => {
    assertEquals(getConfig(), { jiraUrl: '', hasApiKey: false });
  });
});

test('getAllocation/saveAllocation round-trip through PropertiesService', () => {
  withMockProperties({}, () => {
    assertEquals(getAllocation(), []);
    const rows = [{ colorHex: '#828bc2', projectKey: 'ABC', projectName: 'Alpha', ignore: false }];
    saveAllocation(rows);
    assertEquals(getAllocation(), rows);
  });
});

test('testJiraConnection returns ok:false without fetching when URL/key are blank', () => {
  let fetchCalled = false;
  withMockUrlFetch({ fetch: () => { fetchCalled = true; return mockResponse_(200, '{}'); } }, () => {
    withMockProperties({}, () => {
      const result = testJiraConnection('', '');
      assertTrue(result.ok === false, 'expected ok:false');
      assertTrue(!!result.error, 'expected an error message');
    });
  });
  assertTrue(!fetchCalled, 'fetch must not be called when URL/key are blank');
});

test('testJiraConnection returns ok:true with displayName on a 200 response', () => {
  withMockSession('user@example.com', null, () => {
    withMockUrlFetch({ fetch: () => mockResponse_(200, JSON.stringify({ displayName: 'Test User' })) }, () => {
      const result = testJiraConnection('https://example.atlassian.net', 'secret');
      assertEquals(result, { ok: true, displayName: 'Test User' });
    });
  });
});

test('testJiraConnection returns ok:false on a 4xx response', () => {
  withMockSession('user@example.com', null, () => {
    withMockUrlFetch({ fetch: () => mockResponse_(401, 'Unauthorized') }, () => {
      const result = testJiraConnection('https://example.atlassian.net', 'bad-key');
      assertTrue(result.ok === false, 'expected ok:false');
      assertTrue(result.error.indexOf('401') !== -1, 'expected error to mention the status code');
    });
  });
});

test('testJiraConnection returns ok:false instead of throwing when fetch itself throws', () => {
  withMockSession('user@example.com', null, () => {
    withMockUrlFetch({ fetch: () => { throw new Error('network down'); } }, () => {
      const result = testJiraConnection('https://example.atlassian.net', 'secret');
      assertEquals(result, { ok: false, error: 'network down' });
    });
  });
});
```

- [ ] **Step 5: Verify locally with a throwaway Node harness (not committed)**

Apps Script has no local runner, but `Code.js`/`Tests.js` are plain JS with only bare global references (`UrlFetchApp`, `PropertiesService`, `Session`, `CalendarApp`, `Logger`), so they can be executed together under Node by loading both files into the same global (non-module) scope. Create a scratch-only file — do **not** add it to the repo — to verify as you go:

```bash
cat > /private/tmp/claude-502/*/scratchpad/run-tests-local.js <<'EOF'
const fs = require('fs');
const vm = require('vm');
const path = require('path');
const projectDir = process.argv[2];
const code = fs.readFileSync(path.join(projectDir, 'Code.js'), 'utf8');
const tests = fs.readFileSync(path.join(projectDir, 'Tests.js'), 'utf8');
global.Logger = { log: (...args) => console.log(...args) };
vm.runInThisContext(code, { filename: 'Code.js' });
vm.runInThisContext(tests, { filename: 'Tests.js' });
runAllTests();
EOF
node /private/tmp/claude-502/*/scratchpad/run-tests-local.js /Users/pperrin/Documents/Projects/jira_time_google
```

(Adjust the scratchpad glob to the actual session path.) Expected: a summary line like `9/9 passed` with no `FAIL:` lines. If a test references a GAS global not yet mocked in that test (e.g. forgetting `withMockSession`), Node will throw `ReferenceError: X is not defined` — wrap the missing global in the appropriate `withMock*` helper.

- [ ] **Step 6: Commit**

```bash
git add Tests.js
git commit -m "Add test runner, mocks, and Config-section tests"
```

---

### Task 2: Jira issues section tests

**Files:**
- Modify: `Tests.js` (append section)

**Interfaces:**
- Consumes: `test`, `assertEquals`, `assertTrue`, `assertThrows`, `mockResponse_`, `withMockUrlFetch`, `withMockProperties`, `withMockSession` from Task 1.

- [ ] **Step 1: Add the Jira-issues tests**

```js
/*
-----------------------------------------------
00b: Jira issues
-----------------------------------------------
*/
test('getJiraIssues throws when URL/API key are unset', () => {
  withMockProperties({}, () => {
    assertThrows(() => getJiraIssues(), 'expected a throw when config is missing');
  });
});

test('getJiraIssues excludes Archive and Managed Services Internal projects', () => {
  const searchBody = JSON.stringify({
    issues: [
      { key: 'ABC-1', fields: { summary: 'Keep me', status: { name: 'To Do' }, project: { name: 'Alpha' } } },
      { key: 'ARC-1', fields: { summary: 'Drop me', status: { name: 'To Do' }, project: { name: 'Archive 2020' } } },
      { key: 'MSI-1', fields: { summary: 'Drop me too', status: { name: 'To Do' }, project: { name: 'Managed Services Internal' } } }
    ]
  });
  withMockSession('user@example.com', null, () => {
    withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
      withMockUrlFetch({
        fetch: () => mockResponse_(200, searchBody),
        fetchAll: requests => requests.map(() => mockResponse_(200, JSON.stringify({ worklogs: [], total: 0 })))
      }, () => {
        const issues = getJiraIssues();
        assertEquals(issues.map(i => i.key), ['ABC-1']);
      });
    });
  });
});

test('getJiraIssues concatenates paginated search results', () => {
  let call = 0;
  withMockSession('user@example.com', null, () => {
    withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
      withMockUrlFetch({
        fetch: () => {
          call++;
          if (call === 1) {
            return mockResponse_(200, JSON.stringify({
              issues: [{ key: 'ABC-1', fields: { summary: 'One', status: { name: 'To Do' }, project: { name: 'Alpha' } } }],
              nextPageToken: 'page2'
            }));
          }
          return mockResponse_(200, JSON.stringify({
            issues: [{ key: 'ABC-2', fields: { summary: 'Two', status: { name: 'To Do' }, project: { name: 'Alpha' } } }]
          }));
        },
        fetchAll: requests => requests.map(() => mockResponse_(200, JSON.stringify({ worklogs: [], total: 0 })))
      }, () => {
        const issues = getJiraIssues();
        assertEquals(issues.map(i => i.key), ['ABC-1', 'ABC-2']);
      });
    });
  });
});

test('getJiraIssues computes timeLogged in hours rounded to 2 decimals', () => {
  const searchBody = JSON.stringify({
    issues: [{ key: 'ABC-1', fields: { summary: 'One', status: { name: 'To Do' }, project: { name: 'Alpha' } } }]
  });
  withMockSession('user@example.com', null, () => {
    withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
      withMockUrlFetch({
        fetch: () => mockResponse_(200, searchBody),
        fetchAll: requests => requests.map(() => mockResponse_(200, JSON.stringify({
          worklogs: [{ author: { emailAddress: 'user@example.com' }, timeSpentSeconds: 5401 }],
          total: 1
        })))
      }, () => {
        const issues = getJiraIssues();
        assertEquals(issues[0].timeLogged, 1.5);
      });
    });
  });
});

test('markIssueDone posts the transition whose target status is "Done" (case-insensitive)', () => {
  let postedTransitionId = null;
  withMockSession('user@example.com', null, () => {
    withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
      withMockUrlFetch({
        fetch: (url, opts) => {
          if (opts.method === 'get') {
            return mockResponse_(200, JSON.stringify({
              transitions: [{ id: '11', to: { name: 'In Progress' } }, { id: '31', to: { name: 'DONE' } }]
            }));
          }
          postedTransitionId = JSON.parse(opts.payload).transition.id;
          return mockResponse_(204, '');
        }
      }, () => {
        markIssueDone('ABC-1');
        assertEquals(postedTransitionId, '31');
      });
    });
  });
});

test('markIssueDone throws a clear error when no Done transition exists', () => {
  withMockSession('user@example.com', null, () => {
    withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
      withMockUrlFetch({
        fetch: () => mockResponse_(200, JSON.stringify({ transitions: [{ id: '11', to: { name: 'In Progress' } }] }))
      }, () => {
        assertThrows(() => markIssueDone('ABC-1'), 'expected a throw when no Done transition exists');
      });
    });
  });
});
```

- [ ] **Step 2: Run the local Node harness and confirm all tests pass**

Run: `node /private/tmp/claude-502/*/scratchpad/run-tests-local.js /Users/pperrin/Documents/Projects/jira_time_google`
Expected: pass count increases by 6, no `FAIL:` lines.

- [ ] **Step 3: Commit**

```bash
git add Tests.js
git commit -m "Add Jira-issues section tests"
```

---

### Task 3: Scheduling section tests

**Files:**
- Modify: `Tests.js` (append section)

**Interfaces:**
- Consumes: `test`, `assertEquals`, `withMockCalendar`, `withMockProperties` from Task 1.

- [ ] **Step 1: Add the scheduling tests**

```js
/*
-----------------------------------------------
00c: Scheduling
-----------------------------------------------
*/
test('nextScheduleWorkDay_ skips a weekend from Friday to Monday', () => {
  const friday = new Date(2026, 8, 25, 14, 0, 0); // Friday Sep 25 2026
  const next = nextScheduleWorkDay_(friday);
  assertEquals([next.getFullYear(), next.getMonth(), next.getDate(), next.getHours()], [2026, 8, 28, 9]);
});

test('nextScheduleWorkDay_ moves to the very next day mid-week', () => {
  const tuesday = new Date(2026, 8, 22, 10, 0, 0); // Tuesday Sep 22 2026
  const next = nextScheduleWorkDay_(tuesday);
  assertEquals([next.getFullYear(), next.getMonth(), next.getDate(), next.getHours()], [2026, 8, 23, 9]);
});

test('scheduleDayCapacityMs_ returns remaining ms before 5pm', () => {
  const cursor = new Date(2026, 8, 23, 15, 0, 0); // 3pm
  assertEquals(scheduleDayCapacityMs_(cursor), 2 * 60 * 60 * 1000);
});

test('scheduleDayCapacityMs_ returns 0 once past 5pm', () => {
  const cursor = new Date(2026, 8, 23, 18, 0, 0); // 6pm
  assertEquals(scheduleDayCapacityMs_(cursor), 0);
});

function makeFakeCalendar_() {
  const events = [];
  return {
    events,
    createEvent: (title, start, end, opts) => {
      const event = { title, start, end, opts, color: null, setColor: c => { event.color = c; } };
      events.push(event);
      return event;
    }
  };
}

test('scheduleCalendarEvents creates one event for an entry under 8h', () => {
  const fakeCalendar = makeFakeCalendar_();
  withMockProperties({ ALLOCATION: JSON.stringify([{ projectKey: 'ABC', projectName: 'Alpha', colorHex: '#828bc2' }]) }, () => {
    withMockCalendar(fakeCalendar, () => {
      const result = scheduleCalendarEvents(
        [{ jiraProject: 'ABC', key: 'ABC-1', summary: 'Do work', hours: 4 }],
        '2026-09-22T09:00:00'
      );
      assertEquals(result.created, 1);
      assertEquals(fakeCalendar.events.length, 1);
      assertEquals(fakeCalendar.events[0].title, 'Client Task Time: Alpha - ABC-1');
      assertEquals(fakeCalendar.events[0].color, CalendarApp.EventColor.PALE_BLUE);
    });
  });
});

test('scheduleCalendarEvents splits an over-8h entry across a skipped weekend', () => {
  const fakeCalendar = makeFakeCalendar_();
  withMockProperties({ ALLOCATION: JSON.stringify([{ projectKey: 'ABC', projectName: 'Alpha', colorHex: '#828bc2' }]) }, () => {
    withMockCalendar(fakeCalendar, () => {
      // Friday 9am + 10 hours of work: 8h fills Friday, 2h rolls to Monday.
      scheduleCalendarEvents(
        [{ jiraProject: 'ABC', key: 'ABC-1', summary: 'Big task', hours: 10 }],
        '2026-09-25T09:00:00'
      );
      assertEquals(fakeCalendar.events.length, 2);
      assertEquals(fakeCalendar.events[0].start.getDay(), 5); // Friday
      assertEquals(fakeCalendar.events[1].start.getDay(), 1); // Monday
    });
  });
});
```

- [ ] **Step 2: Run the local Node harness and confirm all tests pass**

Run: `node /private/tmp/claude-502/*/scratchpad/run-tests-local.js /Users/pperrin/Documents/Projects/jira_time_google`
Expected: pass count increases by 6, no `FAIL:` lines.

- [ ] **Step 3: Commit**

```bash
git add Tests.js
git commit -m "Add scheduling section tests"
```

---

### Task 4: Calendar import section tests

**Files:**
- Modify: `Tests.js` (append section)

**Interfaces:**
- Consumes: `test`, `assertEquals`, `withMockCalendar`, `withMockProperties`, `withMockSession`, `makeFakeCalendar_` from Tasks 1 and 3.

- [ ] **Step 1: Add the calendar-import tests**

```js
/*
-----------------------------------------------
00d: Calendar import
-----------------------------------------------
*/
function makeFakeCalendarEvent_({ title, color, status, start, end, description }) {
  return {
    getTitle: () => title,
    getColor: () => color,
    getMyStatus: () => status,
    getStartTime: () => start,
    getEndTime: () => end,
    getDescription: () => description || ''
  };
}

test('importCalendarEvents excludes events the user declined', () => {
  const declined = makeFakeCalendarEvent_({
    title: 'Declined meeting', color: '1', status: CalendarApp.GuestStatus.NO,
    start: new Date(2026, 8, 22, 9, 0, 0), end: new Date(2026, 8, 22, 10, 0, 0)
  });
  const fakeCalendar = { getEvents: () => [declined] };
  withMockProperties({
    JIRA_BASE_URL: 'https://example.atlassian.net',
    JIRA_API_KEY: 'secret',
    ALLOCATION: JSON.stringify([{ projectKey: 'ABC', colorHex: '#828bc2' }])
  }, () => {
    withMockSession('user@example.com', 'America/Denver', () => {
      withMockCalendar(fakeCalendar, () => {
        assertEquals(importCalendarEvents('2026-09-22', '2026-09-22'), []);
      });
    });
  });
});

test('importCalendarEvents excludes events whose color has no active allocation row', () => {
  const unmatched = makeFakeCalendarEvent_({
    title: 'Unrelated', color: '5', status: CalendarApp.GuestStatus.YES,
    start: new Date(2026, 8, 22, 9, 0, 0), end: new Date(2026, 8, 22, 10, 0, 0)
  });
  const fakeCalendar = { getEvents: () => [unmatched] };
  withMockProperties({
    JIRA_BASE_URL: 'https://example.atlassian.net',
    JIRA_API_KEY: 'secret',
    ALLOCATION: JSON.stringify([{ projectKey: 'ABC', colorHex: '#828bc2' }])
  }, () => {
    withMockSession('user@example.com', 'America/Denver', () => {
      withMockCalendar(fakeCalendar, () => {
        assertEquals(importCalendarEvents('2026-09-22', '2026-09-22'), []);
      });
    });
  });
});

test('importCalendarEvents matches an active row color and computes quarter-hour duration', () => {
  const matched = makeFakeCalendarEvent_({
    title: 'Client work', color: '1', status: CalendarApp.GuestStatus.YES,
    start: new Date(2026, 8, 22, 9, 0, 0), end: new Date(2026, 8, 22, 10, 40, 0),
    description: 'Working on the thing\n_extra'
  });
  const fakeCalendar = { getEvents: () => [matched] };
  withMockProperties({
    JIRA_BASE_URL: 'https://example.atlassian.net',
    JIRA_API_KEY: 'secret',
    ALLOCATION: JSON.stringify([{ projectKey: 'ABC', colorHex: '#828bc2' }])
  }, () => {
    withMockSession('user@example.com', 'America/Denver', () => {
      withMockCalendar(fakeCalendar, () => {
        const events = importCalendarEvents('2026-09-22', '2026-09-22');
        assertEquals(events.length, 1);
        assertEquals(events[0].projectKey, 'ABC');
        assertEquals(events[0].duration, 1.75);
        assertEquals(events[0].description, 'Working on the thing');
      });
    });
  });
});

test('importCalendarEvents never matches an ignored allocation row even with a matching color', () => {
  const matched = makeFakeCalendarEvent_({
    title: 'Client work', color: '1', status: CalendarApp.GuestStatus.YES,
    start: new Date(2026, 8, 22, 9, 0, 0), end: new Date(2026, 8, 22, 10, 0, 0)
  });
  const fakeCalendar = { getEvents: () => [matched] };
  withMockProperties({
    JIRA_BASE_URL: 'https://example.atlassian.net',
    JIRA_API_KEY: 'secret',
    ALLOCATION: JSON.stringify([{ projectKey: 'ABC', colorHex: '#828bc2', ignore: true }])
  }, () => {
    withMockSession('user@example.com', 'America/Denver', () => {
      withMockCalendar(fakeCalendar, () => {
        assertEquals(importCalendarEvents('2026-09-22', '2026-09-22'), []);
      });
    });
  });
});
```

- [ ] **Step 2: Run the local Node harness and confirm all tests pass**

Run: `node /private/tmp/claude-502/*/scratchpad/run-tests-local.js /Users/pperrin/Documents/Projects/jira_time_google`
Expected: pass count increases by 4, no `FAIL:` lines.

- [ ] **Step 3: Commit**

```bash
git add Tests.js
git commit -m "Add calendar-import section tests"
```

---

### Task 5: Worklogs / sendTimeEntries section tests

**Files:**
- Modify: `Tests.js` (append section)

**Interfaces:**
- Consumes: `test`, `assertEquals`, `assertTrue`, `mockResponse_`, `withMockUrlFetch`, `withMockProperties`, `withMockSession` from Task 1.

- [ ] **Step 1: Add the worklog-parsing and `getWorklogs` tests**

```js
/*
-----------------------------------------------
04: Worklogs
-----------------------------------------------
*/
test('parseTimeSpentHours_ converts Jira timeSpent strings to decimal hours', () => {
  assertEquals(parseTimeSpentHours_('1h 30m'), 1.5);
  assertEquals(parseTimeSpentHours_('45m'), 0.75);
  assertEquals(parseTimeSpentHours_('2h'), 2);
  assertEquals(parseTimeSpentHours_(''), 0);
  assertEquals(parseTimeSpentHours_(undefined), 0);
});

test('getWorklogs includes only the current user\'s worklog entries', () => {
  const searchBody = JSON.stringify({
    issues: [{
      key: 'ABC-1',
      fields: {
        summary: 'One', project: { key: 'ABC' },
        worklog: {
          worklogs: [
            { author: { emailAddress: 'user@example.com' }, timeSpent: '1h', started: '2026-03-10T09:00:00.000-0700' },
            { author: { emailAddress: 'other@example.com' }, timeSpent: '2h', started: '2026-03-10T09:00:00.000-0700' }
          ],
          total: 2
        }
      }
    }]
  });
  withMockSession('user@example.com', null, () => {
    withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
      withMockUrlFetch({ fetch: () => mockResponse_(200, searchBody) }, () => {
        const rows = getWorklogs(2026);
        assertEquals(rows.length, 1);
        assertEquals(rows[0].hours, 1);
      });
    });
  });
});

test('getWorklogs excludes entries outside the requested year, even from an extra-paginated page', () => {
  let call = 0;
  const searchBody = JSON.stringify({
    issues: [{
      key: 'ABC-1',
      fields: {
        summary: 'One', project: { key: 'ABC' },
        worklog: { worklogs: [], total: 1 } // 0 loaded, 1 total -> triggers one extra page fetch
      }
    }]
  });
  withMockSession('user@example.com', null, () => {
    withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
      withMockUrlFetch({
        fetch: url => {
          call++;
          if (call === 1) return mockResponse_(200, searchBody);
          // Extra worklog page: one entry from the prior year, out of range.
          return mockResponse_(200, JSON.stringify({
            worklogs: [{ author: { emailAddress: 'user@example.com' }, timeSpent: '1h', started: '2025-12-31T09:00:00.000-0700' }]
          }));
        }
      }, () => {
        const rows = getWorklogs(2026);
        assertEquals(rows, []);
      });
    });
  });
});

test('sendTimeEntries parses AM/PM start times correctly, including 12am/12pm edge cases', () => {
  const posted = [];
  withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
    withMockSession('user@example.com', null, () => {
      withMockUrlFetch({
        fetch: (url, opts) => { posted.push(JSON.parse(opts.payload)); return mockResponse_(201, '{}'); }
      }, () => {
        sendTimeEntries([
          { date: '2026-03-10', issueKey: 'ABC-1', startTime: '12:00 AM', durationHours: 1 },
          { date: '2026-03-10', issueKey: 'ABC-2', startTime: '12:00 PM', durationHours: 1 }
        ]);
      });
    });
  });
  assertEquals(posted.length, 2);
  assertTrue(posted[0].started.indexOf('T00:00:00') !== -1, '12:00 AM should be hour 0');
  assertTrue(posted[1].started.indexOf('T12:00:00') !== -1, '12:00 PM should be hour 12');
});

test('sendTimeEntries formats a late-local-time entry to its correct UTC date (day-boundary crossing)', () => {
  // America/Denver 11:30 PM local == 05:30 UTC the *next* calendar day (MDT, UTC-6) or 06:30 (MST, UTC-7).
  // Assert only that the UTC hour/date rolled forward relative to the local date, not a fixed offset,
  // so the test is correct regardless of daylight-saving status on the run date.
  let posted = null;
  withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
    withMockSession('user@example.com', null, () => {
      withMockUrlFetch({
        fetch: (url, opts) => { posted = JSON.parse(opts.payload); return mockResponse_(201, '{}'); }
      }, () => {
        sendTimeEntries([{ date: '2026-03-10', issueKey: 'ABC-1', startTime: '11:30 PM', durationHours: 1 }]);
      });
    });
  });
  const utcDate = posted.started.slice(0, 10);
  assertTrue(utcDate === '2026-03-11' || utcDate === '2026-03-10', 'expected a valid UTC-shifted or same-day date, not garbage');
  assertTrue(/^\d{4}-\d{2}-\d{2}T\d{2}:30:00\.000\+0000$/.test(posted.started), 'expected the formatted UTC string shape with :30 minutes preserved');
});

test('sendTimeEntries skips entries with no issueKey', () => {
  let fetchCalled = false;
  withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
    withMockSession('user@example.com', null, () => {
      withMockUrlFetch({ fetch: () => { fetchCalled = true; return mockResponse_(201, '{}'); } }, () => {
        const result = sendTimeEntries([{ date: '2026-03-10', issueKey: '', startTime: '9:00 AM', durationHours: 1 }]);
        assertEquals(result, { succeeded: 0, failed: 0, errors: [] });
      });
    });
  });
  assertTrue(!fetchCalled, 'fetch must not be called for an entry with no issueKey');
});

test('sendTimeEntries counts succeeded/failed against mocked responses and records thrown errors', () => {
  let call = 0;
  withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
    withMockSession('user@example.com', null, () => {
      withMockUrlFetch({
        fetch: () => {
          call++;
          if (call === 1) return mockResponse_(201, '{}');
          if (call === 2) return mockResponse_(500, 'server error');
          throw new Error('network down');
        }
      }, () => {
        const result = sendTimeEntries([
          { date: '2026-03-10', issueKey: 'ABC-1', startTime: '9:00 AM', durationHours: 1 },
          { date: '2026-03-10', issueKey: 'ABC-2', startTime: '9:00 AM', durationHours: 1 },
          { date: '2026-03-10', issueKey: 'ABC-3', startTime: '9:00 AM', durationHours: 1 }
        ]);
        assertEquals(result.succeeded, 1);
        assertEquals(result.failed, 2);
        assertEquals(result.errors.length, 2);
      });
    });
  });
});
```

- [ ] **Step 2: Run the local Node harness and confirm all tests pass**

Run: `node /private/tmp/claude-502/*/scratchpad/run-tests-local.js /Users/pperrin/Documents/Projects/jira_time_google`
Expected: pass count increases by 7, no `FAIL:` lines.

- [ ] **Step 3: Commit**

```bash
git add Tests.js
git commit -m "Add worklog and sendTimeEntries section tests"
```

---

### Task 6: Allocation tab section tests

**Files:**
- Modify: `Tests.js` (append section)

**Interfaces:**
- Consumes: `test`, `assertEquals`, `withMockProperties`, `withMockUrlFetch`, `withMockSession`, `mockResponse_` from Task 1.

- [ ] **Step 1: Add the allocation-tab tests**

```js
/*
-----------------------------------------------
05: Allocation tab
-----------------------------------------------
*/
test('getAllocationValues_ returns {} for a year never saved', () => {
  withMockProperties({}, () => {
    assertEquals(getAllocationValues_(2026), {});
  });
});

test('getAllocationValues_ returns the saved sparse grid for a saved year', () => {
  withMockProperties({ ALLOCATION_VALUES: JSON.stringify({ 2026: { ABC: { 3: 40 } } }) }, () => {
    assertEquals(getAllocationValues_(2026), { ABC: { 3: 40 } });
  });
});

test('getAllocationTabData sorts project keys and excludes ignored ones', () => {
  const searchBody = JSON.stringify({
    issues: [
      { key: 'ZZZ-1', fields: { summary: 'Z', project: { key: 'ZZZ' }, worklog: { worklogs: [{ author: { emailAddress: 'user@example.com' }, timeSpent: '1h', started: '2026-01-05T09:00:00.000-0700' }], total: 1 } } },
      { key: 'AAA-1', fields: { summary: 'A', project: { key: 'AAA' }, worklog: { worklogs: [{ author: { emailAddress: 'user@example.com' }, timeSpent: '1h', started: '2026-01-05T09:00:00.000-0700' }], total: 1 } } },
      { key: 'IGN-1', fields: { summary: 'I', project: { key: 'IGN' }, worklog: { worklogs: [{ author: { emailAddress: 'user@example.com' }, timeSpent: '1h', started: '2026-01-05T09:00:00.000-0700' }], total: 1 } } }
    ]
  });
  withMockSession('user@example.com', null, () => {
    withMockProperties({
      JIRA_BASE_URL: 'https://example.atlassian.net',
      JIRA_API_KEY: 'secret',
      ALLOCATION: JSON.stringify([{ projectKey: 'IGN', ignore: true }])
    }, () => {
      withMockUrlFetch({ fetch: () => mockResponse_(200, searchBody) }, () => {
        const data = getAllocationTabData(2026);
        assertEquals(data.projects, ['AAA', 'ZZZ']);
      });
    });
  });
});

test('saveAllocationGrid replaces only the given year, leaving other years untouched', () => {
  withMockProperties({ ALLOCATION_VALUES: JSON.stringify({ 2025: { ABC: { 1: 10 } } }) }, () => {
    saveAllocationGrid(2026, { ABC: { 2: 20 } });
    const raw = JSON.parse(PropertiesService.getUserProperties().getProperty('ALLOCATION_VALUES'));
    assertEquals(raw, { 2025: { ABC: { 1: 10 } }, 2026: { ABC: { 2: 20 } } });
  });
});
```

- [ ] **Step 2: Run the local Node harness and confirm all tests pass**

Run: `node /private/tmp/claude-502/*/scratchpad/run-tests-local.js /Users/pperrin/Documents/Projects/jira_time_google`
Expected: pass count increases by 4, no `FAIL:` lines.

- [ ] **Step 3: Commit**

```bash
git add Tests.js
git commit -m "Add allocation-tab section tests"
```

---

### Task 7: Utilization tab section tests

**Files:**
- Modify: `Tests.js` (append section)

**Interfaces:**
- Consumes: `test`, `assertEquals`, `withMockProperties` from Task 1.

- [ ] **Step 1: Add the pay-period and utilization tests**

```js
/*
-----------------------------------------------
06: Utilization tab
-----------------------------------------------
*/
test('getPayPeriodEndDates_ includes only period ends within the target year', () => {
  withMockProperties({ PAY_PERIOD_END_DATE: '2026-01-14' }, () => {
    const dates = getPayPeriodEndDates_(2026);
    assertTrue(dates.every(d => d.getFullYear() === 2026), 'every returned date must fall within 2026');
    assertTrue(dates.length > 0, 'expected at least one pay period in 2026');
    // First period should be within 14 days of Jan 1.
    assertTrue(dates[0].getMonth() === 0 && dates[0].getDate() <= 28, 'first period end should be in January');
  });
});

test('getPayPeriodSummary_ counts only periods on or before today, and zeroes out with no anchor', () => {
  withMockProperties({}, () => {
    assertEquals(getPayPeriodSummary_(2026), { payPeriodsToDate: 0, firstPayPeriodEnd: null, mostRecentPayPeriodEnd: null });
  });
  withMockProperties({ PAY_PERIOD_END_DATE: '2020-01-14' }, () => {
    const summary = getPayPeriodSummary_(2020);
    assertTrue(summary.payPeriodsToDate > 0, 'a fully past year should have periods to date');
    assertTrue(summary.mostRecentPayPeriodEnd.getFullYear() === 2020);
  });
});

test('calculateUtilization treats an unset allocation and an explicit zero allocation both as a null cell', () => {
  withMockProperties({
    WORKLOG_CACHE: JSON.stringify({ 2026: [
      { projectKey: 'ABC', issueKey: 'ABC-1', summary: 'x', timeSpent: '10h', started: '2026-01-05T09:00:00.000-0700', hours: 10, month: '2026-01' },
      { projectKey: 'DEF', issueKey: 'DEF-1', summary: 'y', timeSpent: '5h', started: '2026-02-05T09:00:00.000-0700', hours: 5, month: '2026-02' }
    ] }),
    ALLOCATION: JSON.stringify([{ projectKey: 'ABC' }, { projectKey: 'DEF' }]),
    ALLOCATION_VALUES: JSON.stringify({ 2026: { DEF: { 2: 0 } } }), // ABC: unset; DEF: explicit zero
    PAY_PERIOD_END_DATE: '', UTIL_HOURLY_RATE: '0', UTIL_OVERHEAD_RATE: '0'
  }, () => {
    const result = calculateUtilization(2026, false);
    assertEquals(result.utilization.cells.ABC[1], null);
    assertEquals(result.utilization.cells.DEF[2], null);
  });
});

test('calculateUtilization totals are sum(hours)/sum(allocation), not an average of per-cell percentages', () => {
  withMockProperties({
    WORKLOG_CACHE: JSON.stringify({ 2026: [
      { projectKey: 'ABC', issueKey: 'ABC-1', summary: 'x', timeSpent: '10h', started: '2026-01-05T09:00:00.000-0700', hours: 10, month: '2026-01' },
      { projectKey: 'ABC', issueKey: 'ABC-2', summary: 'x', timeSpent: '90h', started: '2026-02-05T09:00:00.000-0700', hours: 90, month: '2026-02' }
    ] }),
    ALLOCATION: JSON.stringify([{ projectKey: 'ABC' }]),
    // Jan: 10/100 = 10%. Feb: 90/100 = 90%. A naive average would be 50%;
    // sum(hours)/sum(allocation) = 100/200 = 50% too in this fixture, so use unequal weights:
    ALLOCATION_VALUES: JSON.stringify({ 2026: { ABC: { 1: 100, 2: 10 } } }),
    PAY_PERIOD_END_DATE: '', UTIL_HOURLY_RATE: '0', UTIL_OVERHEAD_RATE: '0'
  }, () => {
    const result = calculateUtilization(2026, false);
    // Jan: 10/100=10%, Feb: 90/10=900%. Average of percentages = 455%.
    // sum(hours)/sum(allocation) = 100/110 = 90.9%.
    assertEquals(result.utilization.colTotals.ABC, 90.9);
  });
});

test('calculateUtilization excludes unrated projects from revenue and lists them in unratedProjects', () => {
  withMockProperties({
    WORKLOG_CACHE: JSON.stringify({ 2026: [
      { projectKey: 'RATED', issueKey: 'R-1', summary: 'x', timeSpent: '10h', started: '2026-01-05T09:00:00.000-0700', hours: 10, month: '2026-01' },
      { projectKey: 'UNRATED', issueKey: 'U-1', summary: 'y', timeSpent: '10h', started: '2026-01-05T09:00:00.000-0700', hours: 10, month: '2026-01' }
    ] }),
    ALLOCATION: JSON.stringify([{ projectKey: 'RATED', rate: 100 }, { projectKey: 'UNRATED' }]),
    ALLOCATION_VALUES: JSON.stringify({}),
    PAY_PERIOD_END_DATE: '2026-01-14', UTIL_HOURLY_RATE: '0', UTIL_OVERHEAD_RATE: '0'
  }, () => {
    const result = calculateUtilization(2026, false);
    assertEquals(result.revenue.unratedProjects, ['UNRATED']);
    assertEquals(result.revenue.grossRevenue, 1000);
  });
});

test('getOrCollectWorklogs serves from cache unless forceRefresh is true', () => {
  let getWorklogsCalls = 0;
  const realGetWorklogs = getWorklogs;
  getWorklogs = year => { getWorklogsCalls++; return [{ projectKey: 'ABC', issueKey: 'ABC-1', summary: 'x', timeSpent: '1h', started: '2026-01-05T09:00:00.000-0700', hours: 1, month: '2026-01' }]; };
  try {
    withMockProperties({ WORKLOG_CACHE: JSON.stringify({ 2026: [{ projectKey: 'CACHED' }] }) }, () => {
      const cached = getOrCollectWorklogs(2026, false);
      assertEquals(cached, [{ projectKey: 'CACHED' }]);
      assertEquals(getWorklogsCalls, 0);

      const refreshed = getOrCollectWorklogs(2026, true);
      assertEquals(refreshed[0].projectKey, 'ABC');
      assertEquals(getWorklogsCalls, 1);
    });
  } finally {
    getWorklogs = realGetWorklogs;
  }
});
```

- [ ] **Step 2: Run the local Node harness and confirm all tests pass**

Run: `node /private/tmp/claude-502/*/scratchpad/run-tests-local.js /Users/pperrin/Documents/Projects/jira_time_google`
Expected: pass count increases by 6, no `FAIL:` lines.

- [ ] **Step 3: Commit**

```bash
git add Tests.js
git commit -m "Add utilization-tab section tests"
```

---

### Task 8: Create Issue section tests

**Files:**
- Modify: `Tests.js` (append section)

**Interfaces:**
- Consumes: `test`, `assertEquals`, `assertTrue`, `mockResponse_`, `withMockUrlFetch`, `withMockProperties`, `withMockSession` from Task 1.

- [ ] **Step 1: Add the `makeJira` tests**

```js
/*
-----------------------------------------------
Create Issue
-----------------------------------------------
*/
test('makeJira builds the expected payload and includes assignee when the user search resolves', () => {
  let issuePayload = null;
  withMockSession('user@example.com', null, () => {
    withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
      withMockUrlFetch({
        fetch: (url, opts) => {
          if (url.indexOf('/user/search') !== -1) {
            return mockResponse_(200, JSON.stringify([{ accountId: 'acct-1' }]));
          }
          issuePayload = JSON.parse(opts.payload);
          return mockResponse_(201, JSON.stringify({ key: 'ABC-99' }));
        }
      }, () => {
        const result = makeJira({
          input1: 'ABC', input2: 'field2', input3: 'field3', input4: 'Summary text',
          input6: 'Task', input7: '3', notes: 'Some notes'
        });
        assertEquals(result, 'Successfully created ABC-99.');
      });
    });
  });
  assertEquals(issuePayload.fields.project, { key: 'ABC' });
  assertEquals(issuePayload.fields.summary, 'Summary text');
  assertEquals(issuePayload.fields.issuetype, { name: 'Task' });
  assertEquals(issuePayload.fields.priority, { id: '3' });
  assertEquals(issuePayload.fields.assignee, { accountId: 'acct-1' });
});

test('makeJira omits assignee when the user search finds no match', () => {
  let issuePayload = null;
  withMockSession('user@example.com', null, () => {
    withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
      withMockUrlFetch({
        fetch: (url, opts) => {
          if (url.indexOf('/user/search') !== -1) return mockResponse_(200, JSON.stringify([]));
          issuePayload = JSON.parse(opts.payload);
          return mockResponse_(201, JSON.stringify({ key: 'ABC-100' }));
        }
      }, () => {
        makeJira({ input1: 'ABC', input2: 'f2', input3: 'f3', input4: 'Summary', input6: 'Task', input7: '3', notes: '' });
      });
    });
  });
  assertTrue(!('assignee' in issuePayload.fields), 'assignee must be omitted when no match is found');
});

test('makeJira returns a friendly error string (not a throw) on a non-201 response', () => {
  withMockSession('user@example.com', null, () => {
    withMockProperties({ JIRA_BASE_URL: 'https://example.atlassian.net', JIRA_API_KEY: 'secret' }, () => {
      withMockUrlFetch({
        fetch: url => {
          if (url.indexOf('/user/search') !== -1) return mockResponse_(200, JSON.stringify([]));
          return mockResponse_(400, JSON.stringify({ errors: { summary: 'is required' } }));
        }
      }, () => {
        const result = makeJira({ input1: 'ABC', input2: 'f2', input3: 'f3', input4: '', input6: 'Task', input7: '3', notes: '' });
        assertTrue(result.indexOf('Error (400)') !== -1, 'expected a friendly error string mentioning the status code');
      });
    });
  });
});
```

- [ ] **Step 2: Run the local Node harness and confirm all tests pass**

Run: `node /private/tmp/claude-502/*/scratchpad/run-tests-local.js /Users/pperrin/Documents/Projects/jira_time_google`
Expected: pass count increases by 3, no `FAIL:` lines. Total across all tasks should now be every test case listed in the spec §4 inventory.

- [ ] **Step 3: Commit**

```bash
git add Tests.js
git commit -m "Add Create Issue section tests"
```

---

### Task 9: Wire the suite into the deployment workflow

**Files:**
- Modify: `CLAUDE.md` (Development workflow section)

**Interfaces:**
- None — documentation only.

- [ ] **Step 1: Add the pre-deployment gate line to `CLAUDE.md`**

Find the bullet list under `## Development workflow` and add one bullet directly after the `**No automated tests.**` line (that line's wording will now be inaccurate on its own, so replace it rather than appending after it):

Replace:
```markdown
- **No automated tests.** `Debug.js` contains developer-only helpers (e.g. `validateImportCalendarEvents`) meant to be run manually from the Apps Script editor's function picker, not from a CLI.
```

With:
```markdown
- **Automated tests live in `Tests.js`, run manually.** Before creating or updating a deployment, run `runAllTests` from the Apps Script editor's function picker and confirm no `FAIL:` lines in the execution log, then walk the manual smoke-test checklist in `docs/superpowers/specs/SPEC-2026-09-23-test-suite.md` (§5). `Debug.js` contains separate developer-only diagnostic helpers (e.g. `validateImportCalendarEvents`), also run manually from the function picker, not from a CLI.
```

- [ ] **Step 2: Verify the edit**

Run: `grep -n "runAllTests" CLAUDE.md`
Expected: one match, on the line just added.

- [ ] **Step 3: Commit**

```bash
git add CLAUDE.md
git commit -m "Document the test suite as a pre-deployment gate in CLAUDE.md"
```
