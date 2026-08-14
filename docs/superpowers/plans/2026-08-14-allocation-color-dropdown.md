# Allocation Color Dropdown & Event Filtering Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the allocation table's "Calendar Color #", "Color Label", and "Color Enum Name" fields with a single color dropdown and an "Ignore" checkbox, and update Code.js to filter imported calendar events by declined status and ignore flag.

**Architecture:** Client-side changes to TabConfig.html (table header) and JavaScript.html (row rendering + save logic); server-side changes to Code.js (two functions). No new files. All color mappings are hardcoded constants in Code.js.

**Tech Stack:** Google Apps Script, Vanilla JavaScript, DOM manipulation

## Global Constraints

- XSS: all user/external-sourced strings rendered into innerHTML must pass through `esc()`
- `saveAllocationTable()` must filter rows with no `projectKey` before passing to server
- `ignore` flag only affects `importCalendarEvents()` — NOT `scheduleCalendarEvents()`
- `GuestStatus.NO` filter applies to all events regardless of color/allocation
- Allocation row field names after this change: `colorHex`, `projectKey`, `projectName`, `hoursPerWeek`, `ignore`
- Old field names (`colorNum`, `colorLabel`, `colorEnumName`) must be completely removed from client and server code
- Color dropdown must list all 11 options in format "Label (#hexvalue)" e.g. "Lavender (#828bc2)"
- `HEX_TO_COLOR_ID` and `HEX_TO_EVENT_COLOR` maps must use the exact hex values and mappings from the spec

---

### Task 1: Update TabConfig.html and JavaScript.html

**Files:**
- Modify: `TabConfig.html` (allocation table thead)
- Modify: `JavaScript.html` (`addAllocationRow()` and `saveAllocationTable()`)

**Interfaces:**
- `addAllocationRow(data)` — accepts an object with optional fields `{ colorHex, projectKey, projectName, hoursPerWeek, ignore }`. Called by `renderAllocationTable()` for each saved row, and by the "Add Row" button with no args.
- `saveAllocationTable()` — reads the DOM table, calls `google.script.run.saveAllocation(rows)` where `rows` is an array of `{ colorHex, projectKey, projectName, hoursPerWeek, ignore }` objects (filtered to rows where `projectKey` is non-empty).

- [ ] **Step 1: Update the allocation table thead in TabConfig.html**

Find the allocation table `<thead>` and replace it. Old columns were: Calendar Color # | Jira Project Key | Color Label | Project Name | Hours/Wk | Color Enum Name | [delete button column]

Replace with:
```html
<tr>
  <th>Calendar Color</th>
  <th>Jira Project Key</th>
  <th>Project Name</th>
  <th>Hours/Wk</th>
  <th>Ignore</th>
  <th></th>
</tr>
```

- [ ] **Step 2: Verify TabConfig.html change**

Read back TabConfig.html and confirm:
- The old `<th>Calendar Color #</th>`, `<th>Color Label</th>`, and `<th>Color Enum Name</th>` headers are GONE
- The new `<th>Calendar Color</th>` and `<th>Ignore</th>` headers are present
- Column order is: Calendar Color | Jira Project Key | Project Name | Hours/Wk | Ignore | [empty delete column]

- [ ] **Step 3: Replace `addAllocationRow(data)` in JavaScript.html**

Find the entire `addAllocationRow` function and replace it:

```javascript
function addAllocationRow(data) {
  const tbody = document.getElementById('allocation-body');
  const tr = document.createElement('tr');

  const COLOR_OPTIONS = [
    ['#828bc2', 'Lavender (#828bc2)'],
    ['#55b080', 'Sage (#55b080)'],
    ['#a75aba', 'Grape (#a75aba)'],
    ['#d6837a', 'Flamingo (#d6837a)'],
    ['#e7ba51', 'Banana (#e7ba51)'],
    ['#e3683e', 'Tangerine (#e3683e)'],
    ['#4b99d2', 'Peacock (#4b99d2)'],
    ['#7c7c7c', 'Graphite (#7c7c7c)'],
    ['#6e72c3', 'Blueberry (#6e72c3)'],
    ['#489160', 'Basil (#489160)'],
    ['#da5234', 'Tomato (#da5234)'],
  ];

  const selectedHex = (data && data.colorHex) || '';
  const optionsHtml = COLOR_OPTIONS.map(([val, label]) =>
    '<option value="' + val + '"' + (val === selectedHex ? ' selected' : '') + '>' + label + '</option>'
  ).join('');

  const ignored = data && data.ignore ? 'checked' : '';

  tr.innerHTML =
    '<td><select name="colorHex" style="width:100%">' + optionsHtml + '</select></td>' +
    '<td><input type="text" style="width:100%" name="projectKey" value="' + esc((data && data.projectKey) || '') + '" placeholder="CEOT"></td>' +
    '<td><input type="text" style="width:100%" name="projectName" value="' + esc((data && data.projectName) || '') + '" placeholder="CFA Project"></td>' +
    '<td><input type="text" style="width:100%" name="hoursPerWeek" value="' + esc(String((data && data.hoursPerWeek) || '')) + '" placeholder="20"></td>' +
    '<td style="text-align:center"><input type="checkbox" name="ignore" ' + ignored + '></td>' +
    '<td><button class="btn" style="padding:4px 8px;" onclick="this.closest(\'tr\').remove()">✕</button></td>';

  tbody.appendChild(tr);
}
```

- [ ] **Step 4: Replace `saveAllocationTable()` in JavaScript.html**

Find the entire `saveAllocationTable` function and replace it:

```javascript
function saveAllocationTable() {
  const rows = Array.from(document.querySelectorAll('#allocation-body tr')).map(tr => {
    const cells = tr.querySelectorAll('td');
    return {
      colorHex: cells[0].querySelector('select').value,
      projectKey: cells[1].querySelector('input').value.trim(),
      projectName: cells[2].querySelector('input').value.trim(),
      hoursPerWeek: parseFloat(cells[3].querySelector('input').value) || 0,
      ignore: cells[4].querySelector('input').checked
    };
  }).filter(r => r.projectKey);
  google.script.run
    .withSuccessHandler(() => showStatus('allocation-status', 'Allocation saved.', 'success'))
    .withFailureHandler(e => showStatus('allocation-status', 'Error: ' + e.message, 'error'))
    .saveAllocation(rows);
}
```

- [ ] **Step 5: Verify JavaScript.html changes**

Read back the allocation section of JavaScript.html and confirm:
- `addAllocationRow` no longer references `colorNum`, `colorLabel`, or `colorEnumName`
- `addAllocationRow` builds a `<select>` with 11 `<option>` elements for the color dropdown
- `addAllocationRow` renders a checkbox `<input type="checkbox">` for the ignore field
- `saveAllocationTable` reads `cells[0].querySelector('select').value` for `colorHex`
- `saveAllocationTable` reads `cells[4].querySelector('input').checked` for `ignore`
- No text inputs for colorNum, colorLabel, or colorEnumName remain in either function

- [ ] **Step 6: Commit**

```bash
git add TabConfig.html JavaScript.html
git commit -m "feat: replace allocation color fields with color dropdown and ignore checkbox"
```

---

### Task 2: Update Code.js server functions

**Files:**
- Modify: `Code.js` (`importCalendarEvents()` and `scheduleCalendarEvents()`)

**Interfaces:**
- `importCalendarEvents(startDateStr, endDateStr)` — unchanged signature; internally uses new `HEX_TO_COLOR_ID` map and new `ignore` filter
- `scheduleCalendarEvents(toSchedule)` — unchanged signature; internally uses new `HEX_TO_EVENT_COLOR` map

- [ ] **Step 1: Read the current Code.js**

Read Code.js to find the exact current implementation of `importCalendarEvents()` and `scheduleCalendarEvents()`. Note the exact lines where the colorMap is constructed and where event filtering happens.

- [ ] **Step 2: Update `importCalendarEvents()` in Code.js**

Find the section in `importCalendarEvents()` where:
1. Allocation data is parsed to build a colorMap
2. Events are filtered by color

Replace the colorMap construction. Old pattern (approximately):
```javascript
const colorMap = Object.fromEntries(
  allocation.filter(r => r.projectKey && r.colorLabel).map(r => [r.colorNum, r.projectKey])
);
```

New pattern — add the `HEX_TO_COLOR_ID` constant and rebuild the map:
```javascript
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
```

Then update the per-event filtering to also exclude declined events. The filter must now exclude an event if:
- (a) `event.getMyStatus() === CalendarApp.GuestStatus.NO`, OR
- (b) The event's `colorNum` is not in `colorMap` (i.e., not matched to any active, non-ignored allocation row)

The existing color-match filter should already handle (b) — just add (a) as an additional early-exit check.

- [ ] **Step 3: Update `scheduleCalendarEvents()` in Code.js**

Find the section in `scheduleCalendarEvents()` where:
1. `COLOR_ENUM_MAP` is defined (mapping string names like `'PALE_BLUE'` to `CalendarApp.EventColor.*`)
2. `projectColorMap` is built using `r.colorEnumName`

Replace both with:
```javascript
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
  allocation.filter(r => r.projectKey && r.colorHex)
    .map(r => [r.projectKey, HEX_TO_EVENT_COLOR[r.colorHex] || null])
);
```

Note: `ignore` rows are NOT excluded here — ignore only affects import.

- [ ] **Step 4: Verify Code.js changes**

Read back the modified sections of Code.js and confirm:
- `COLOR_ENUM_MAP` is GONE from `scheduleCalendarEvents()`
- `colorEnumName` is not referenced anywhere in either function
- `colorNum` is not referenced in allocation filtering (it's still fine to use `event.getColor()` which returns the color index — that's different from the old `r.colorNum` field)
- `colorLabel` is not referenced in allocation filtering
- `r.ignore` filter is applied in `importCalendarEvents()` activeRows filter
- `CalendarApp.GuestStatus.NO` check is present in `importCalendarEvents()`
- `HEX_TO_COLOR_ID` is present in `importCalendarEvents()`
- `HEX_TO_EVENT_COLOR` is present in `scheduleCalendarEvents()`

- [ ] **Step 5: Commit**

```bash
git add Code.js
git commit -m "feat: update importCalendarEvents and scheduleCalendarEvents to use colorHex and ignore flag"
```
