// TaskTime — server.
// Sheets: Tasks, Labels, TimeEntries (auto-created on first load).
// Active timer persisted via UserProperties (per-user, per-script).

const SHEETS = {
  Tasks:        ['id', 'name', 'labelId', 'createdAt'],
  Labels:       ['id', 'name', 'color',   'createdAt'],
  TimeEntries:  ['id', 'taskId', 'startTime', 'endTime', 'durationMinutes', 'source', 'notes', 'createdAt']
};

const PALETTE = ['#4285F4','#EA4335','#FBBC04','#34A853','#A142F4','#24C1E0','#FF6D00','#7B1FA2'];

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('SPREADSHEET_ID not set. Run setSpreadsheetId("your-sheet-id") from the script editor once, or set it in Project Settings → Script Properties.');
  return SpreadsheetApp.openById(id);
}

function setSpreadsheetId(id) {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', String(id).trim());
  return id;
}

function ensureSchema_() {
  const ss = ss_();
  Object.entries(SHEETS).forEach(([name, headers]) => {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    if (s.getLastRow() === 0) {
      s.getRange(1, 1, 1, headers.length).setValues([headers]);
      s.setFrozenRows(1);
    }
  });
}

function doGet() {
  ensureSchema_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('TaskTime')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function uid_() { return Utilities.getUuid(); }
function iso_() { return new Date().toISOString(); }
function toIso_(v) { return v instanceof Date ? v.toISOString() : (v || ''); }

function readRows_(name) {
  const s = ss_().getSheetByName(name);
  if (!s || s.getLastRow() <= 1) return [];
  const values = s.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    obj.createdAt = toIso_(obj.createdAt);
    obj.startTime = toIso_(obj.startTime);
    obj.endTime   = toIso_(obj.endTime);
    obj.durationMinutes = Number(obj.durationMinutes) || 0;
    return obj;
  });
}

function appendRow_(name, obj) {
  const s = ss_().getSheetByName(name);
  const headers = s.getDataRange().getValues()[0];
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  s.appendRow(row);
}

function findRowNum_(name, id) {
  const s = ss_().getSheetByName(name);
  if (!s || s.getLastRow() <= 1) return -1;
  const values = s.getDataRange().getValues();
  const idCol = values[0].indexOf('id');
  for (let i = 1; i < values.length; i++) {
    if (values[i][idCol] === id) return i + 1;
  }
  return -1;
}

function updateRow_(name, id, patch) {
  const s = ss_().getSheetByName(name);
  const rowNum = findRowNum_(name, id);
  if (rowNum === -1) throw new Error(`${name} row not found: ${id}`);
  const headers = s.getDataRange().getValues()[0];
  const row = s.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  Object.keys(patch).forEach(k => {
    const col = headers.indexOf(k);
    if (col >= 0) row[col] = patch[k];
  });
  s.getRange(rowNum, 1, 1, headers.length).setValues([row]);
}

function deleteRow_(name, id) {
  const s = ss_().getSheetByName(name);
  const rowNum = findRowNum_(name, id);
  if (rowNum > 0) s.deleteRow(rowNum);
}

// --- Bootstrap ---
function getInitialData() {
  return {
    tasks:  readRows_('Tasks'),
    labels: readRows_('Labels'),
    entries: readRows_('TimeEntries'),
    activeTimer: getActiveTimer()
  };
}

// --- Tasks ---
function getTasks() { return readRows_('Tasks'); }
function createTask(name, labelId) {
  const t = { id: uid_(), name: String(name).trim(), labelId: labelId || '', createdAt: iso_() };
  appendRow_('Tasks', t);
  return getTasks();
}
function updateTask(id, patch) {
  if ('name' in patch) patch.name = String(patch.name).trim();
  if ('labelId' in patch) patch.labelId = patch.labelId || '';
  updateRow_('Tasks', id, patch);
  return getTasks();
}
function deleteTask(id) {
  deleteRow_('Tasks', id);
  return getTasks();
}

// --- Labels ---
function getLabels() { return readRows_('Labels'); }
function nextColor_() {
  const labels = getLabels();
  return PALETTE[labels.length % PALETTE.length];
}
function createLabel(name) {
  const l = { id: uid_(), name: String(name).trim(), color: nextColor_(), createdAt: iso_() };
  appendRow_('Labels', l);
  return getLabels();
}
function updateLabel(id, patch) {
  if ('name' in patch) patch.name = String(patch.name).trim();
  updateRow_('Labels', id, patch);
  return getLabels();
}
function deleteLabel(id) {
  const labels = getLabels();
  const dying = labels.find(l => l.id === id);
  if (dying) {
    const tasks = getTasks().filter(t => t.labelId === id);
    tasks.forEach(t => updateRow_('Tasks', t.id, { labelId: '' }));
  }
  deleteRow_('Labels', id);
  return getLabels();
}

// --- Time entries ---
function getEntries() { return readRows_('TimeEntries'); }
function createEntry(payload) {
  const start = new Date(payload.startTime);
  const end   = new Date(payload.endTime);
  if (!payload.taskId) throw new Error('taskId required');
  if (isNaN(start) || isNaN(end)) throw new Error('invalid start/end time');
  if (end <= start) throw new Error('end must be after start');
  const durationMinutes = Math.max(1, Math.round((end - start) / 60000));
  appendRow_('TimeEntries', {
    id: uid_(),
    taskId: payload.taskId,
    startTime: start,
    endTime: end,
    durationMinutes,
    source: payload.source || 'manual',
    notes: payload.notes || '',
    createdAt: new Date()
  });
  return getEntries();
}
function deleteEntry(id) {
  deleteRow_('TimeEntries', id);
  return getEntries();
}

// ponytail: one-shot rebaser for entries stored under the buggy ISO parser — run from the editor with your local UTC offset (PDT = +7, UTC = 0)
function rebaseEntries(offsetHours) {
  const s = ss_().getSheetByName('TimeEntries');
  if (!s || s.getLastRow() <= 1) return 'No entries';
  const offsetMs = offsetHours * 3600000;
  const data = s.getDataRange().getValues();
  const headers = data[0];
  const startI = headers.indexOf('startTime');
  const endI   = headers.indexOf('endTime');
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    let changed = false;
    if (data[i][startI] instanceof Date) {
      data[i][startI] = new Date(data[i][startI].getTime() + offsetMs).toISOString();
      changed = true;
    }
    if (data[i][endI] instanceof Date) {
      data[i][endI] = new Date(data[i][endI].getTime() + offsetMs).toISOString();
      changed = true;
    }
    if (changed) count++;
  }
  s.getRange(1, 1, data.length, headers.length).setValues(data);
  return `Rebased ${count} entr${count === 1 ? 'y' : 'ies'} by ${offsetHours}h`;
}

// --- Timer (per-user, per-script via UserProperties) ---
function getActiveTimer() {
  const raw = PropertiesService.getUserProperties().getProperty('activeTimer');
  return raw ? JSON.parse(raw) : null;
}

// ponytail: diagnostics — run from the editor to inspect the sheet's raw state
function dumpTimeEntries() {
  const s = ss_().getSheetByName('TimeEntries');
  if (!s) return 'No TimeEntries sheet';
  const data = s.getDataRange().getValues();
  return JSON.stringify({
    sheetName: 'TimeEntries',
    lastRow: s.getLastRow(),
    headerRow: data[0],
    sample: data.slice(1, 4),
    types: data[1] ? data[1].map((v, i) => `[${i}] ${data[0][i]}: ${v === null ? 'null' : typeof v} (${v instanceof Date ? 'Date' : String(v).slice(0, 40)})`) : []
  }, null, 2);
}

function resetTimeEntries() {
  const s = ss_().getSheetByName('TimeEntries');
  if (!s) return 'No TimeEntries sheet';
  if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
  return 'Cleared all entries';
}
function startTimer(taskId) {
  const t = { taskId, startTime: iso_() };
  PropertiesService.getUserProperties().setProperty('activeTimer', JSON.stringify(t));
  return t;
}
function stopTimer() {
  const raw = PropertiesService.getUserProperties().getProperty('activeTimer');
  if (!raw) return null;
  const timer = JSON.parse(raw);
  PropertiesService.getUserProperties().deleteProperty('activeTimer');
  createEntry({
    taskId: timer.taskId,
    startTime: timer.startTime,
    endTime: iso_(),
    source: 'timer'
  });
  return getEntries();
}
