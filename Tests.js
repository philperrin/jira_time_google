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
  withMockProperties({}, () => {
    withMockSession('user@example.com', null, () => {
      withMockUrlFetch({ fetch: () => mockResponse_(200, JSON.stringify({ displayName: 'Test User' })) }, () => {
        const result = testJiraConnection('https://example.atlassian.net', 'secret');
        assertEquals(result, { ok: true, displayName: 'Test User' });
      });
    });
  });
});

test('testJiraConnection returns ok:false on a 4xx response', () => {
  withMockProperties({}, () => {
    withMockSession('user@example.com', null, () => {
      withMockUrlFetch({ fetch: () => mockResponse_(401, 'Unauthorized') }, () => {
        const result = testJiraConnection('https://example.atlassian.net', 'bad-key');
        assertTrue(result.ok === false, 'expected ok:false');
        assertTrue(result.error.indexOf('401') !== -1, 'expected error to mention the status code');
      });
    });
  });
});

test('testJiraConnection returns ok:false instead of throwing when fetch itself throws', () => {
  withMockProperties({}, () => {
    withMockSession('user@example.com', null, () => {
      withMockUrlFetch({ fetch: () => { throw new Error('network down'); } }, () => {
        const result = testJiraConnection('https://example.atlassian.net', 'secret');
        assertEquals(result, { ok: false, error: 'network down' });
      });
    });
  });
});

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
