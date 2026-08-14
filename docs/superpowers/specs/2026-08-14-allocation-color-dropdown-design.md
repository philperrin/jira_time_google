# Allocation Color Dropdown & Event Filtering — Design

## Overview

Three related changes:

1. **Config tab:** Replace "Calendar Color #", "Color Label", and "Color Enum Name" allocation columns with a single "Calendar Color" dropdown and an "Ignore" checkbox
2. **Code.js `importCalendarEvents`:** Filter out declined events (`GuestStatus.NO`) and events whose color maps to an ignored allocation row
3. **Code.js `scheduleCalendarEvents`:** Drive event color from `colorHex` instead of `colorEnumName`

---

## Allocation Row Schema Changes

### Removed fields
- `colorNum` — replaced by `colorHex` (hex → color ID lookup is now hardcoded in Code.js)
- `colorLabel` — redundant; dropdown label carries the color name
- `colorEnumName` — replaced by `colorHex` (hex → EventColor enum lookup is now hardcoded in Code.js)

### New fields
- `colorHex` — string, one of the 11 known hex values (e.g. `"#828bc2"`)
- `ignore` — boolean, when `true` events with this color are excluded from import

### Kept fields
- `projectKey`, `projectName`, `hoursPerWeek`

### New row shape (what `saveAllocationTable` returns and `getAllocation` stores)
```json
{
  "colorHex": "#828bc2",
  "projectKey": "CEOT",
  "projectName": "CFA Project",
  "hoursPerWeek": 20,
  "ignore": false
}
```

### Migration note
Existing saved allocation data uses the old field names (`colorNum`, `colorLabel`, `colorEnumName`). After deployment, the Config tab will load each row but the color dropdown will default to the first option (Lavender) since `colorHex` won't be in the stored data. **The user must open the Config tab, re-select each project's color, and click Save Allocation once** to migrate to the new schema. No automated migration is needed.

---

## Hardcoded Color Mappings (Code.js)

Both server functions share the same mapping logic. Defined once at the top of each function (or as a module-level constant):

```javascript
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
```

`HEX_TO_COLOR_ID` is used by `importCalendarEvents`. `HEX_TO_EVENT_COLOR` is used by `scheduleCalendarEvents`.

---

## 1. TabConfig.html Changes

### Allocation table thead

Old columns (7): Calendar Color # | Jira Project Key | Color Label | Project Name | Hours/Wk | Color Enum Name | [delete]

New columns (6): **Calendar Color** | **Jira Project Key** | **Project Name** | **Hours/Wk** | **Ignore** | [delete]

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

---

## 2. JavaScript.html Changes

### `addAllocationRow(data)`

The color field becomes a `<select>` with 11 options. The ignore field becomes a checkbox. The three removed text inputs are gone.

New row structure (6 cells):

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
    `<option value="${val}"${val === selectedHex ? ' selected' : ''}>${label}</option>`
  ).join('');

  const ignored = data && data.ignore ? 'checked' : '';

  tr.innerHTML =
    `<td><select name="colorHex" style="width:100%">${optionsHtml}</select></td>` +
    `<td><input type="text" style="width:100%" name="projectKey" value="${esc((data && data.projectKey) || '')}" placeholder="CEOT"></td>` +
    `<td><input type="text" style="width:100%" name="projectName" value="${esc((data && data.projectName) || '')}" placeholder="CFA Project"></td>` +
    `<td><input type="text" style="width:100%" name="hoursPerWeek" value="${esc(String((data && data.hoursPerWeek) || ''))}" placeholder="20"></td>` +
    `<td style="text-align:center"><input type="checkbox" name="ignore" ${ignored}></td>` +
    `<td><button class="btn" style="padding:4px 8px;" onclick="this.closest('tr').remove()">✕</button></td>`;

  tbody.appendChild(tr);
}
```

### `saveAllocationTable()`

Reads the new field set from each row. The select is the first child of the first cell; the checkbox is in the fifth cell.

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

---

## 3. Code.js Changes

### `importCalendarEvents()`

Replace the color-map construction and add two new filters:

**Old color map:**
```javascript
const colorMap = Object.fromEntries(
  allocation.filter(r => r.projectKey && r.colorLabel).map(r => [r.colorNum, r.projectKey])
);
const validProjectKeys = new Set(
  allocation.filter(r => r.projectKey && r.colorLabel).map(r => r.projectKey)
);
```

**New color map (uses `HEX_TO_COLOR_ID` + excludes ignored rows):**
```javascript
const HEX_TO_COLOR_ID = { /* 11-entry map as above */ };
const activeRows = allocation.filter(r => r.projectKey && r.colorHex && !r.ignore);
const colorMap = Object.fromEntries(
  activeRows.map(r => [HEX_TO_COLOR_ID[r.colorHex], r.projectKey])
);
const validProjectKeys = new Set(activeRows.map(r => r.projectKey));
```

**New event filter (add after existing color check, before return):**
```javascript
const colorNum = event.getColor();
const projectKey = colorMap[colorNum] || '';
if (!validProjectKeys.has(projectKey)) return null;
if (event.getMyStatus() === CalendarApp.GuestStatus.NO) return null;
```

Note: `GuestStatus.NO` covers events the user explicitly declined. Events with status `INVITED` (no response yet), `YES` (accepted), `MAYBE`, or `OWNER` are not filtered.

### `scheduleCalendarEvents()`

Replace the `COLOR_ENUM_MAP` + `colorEnumName` lookup:

**Old:**
```javascript
const COLOR_ENUM_MAP = { 'PALE_BLUE': CalendarApp.EventColor.PALE_BLUE, ... };
const projectColorMap = Object.fromEntries(
  allocation.filter(r => r.projectKey)
    .map(r => [r.projectKey, COLOR_ENUM_MAP[r.colorEnumName] || null])
);
```

**New:**
```javascript
const HEX_TO_EVENT_COLOR = { /* 11-entry map as above */ };
const projectColorMap = Object.fromEntries(
  allocation.filter(r => r.projectKey && r.colorHex)
    .map(r => [r.projectKey, HEX_TO_EVENT_COLOR[r.colorHex] || null])
);
```

Ignored projects are NOT excluded from scheduling (Ignore only affects event import).

---

## Files Changed

| File | Change |
|---|---|
| `TabConfig.html` | Update allocation table thead (6 columns instead of 7) |
| `JavaScript.html` | Rewrite `addAllocationRow()` and `saveAllocationTable()` |
| `Code.js` | Update `importCalendarEvents()` (new colorMap + two filters); update `scheduleCalendarEvents()` (new colorMap) |
