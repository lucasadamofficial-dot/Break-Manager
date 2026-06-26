/*  ============================================================
    BREAK MANAGEMENT — Google Apps Script backend (Google Sheets)
    ============================================================

    This is bound to your sheet:
    https://docs.google.com/spreadsheets/d/1gtVboWcy9xGs1vXhDZZVEGGyTqSQCFR8OpKd1Mwg7r8/edit

    SETUP (one time, ~2 min):
    1. Open that sheet → Extensions → Apps Script.
    2. Delete any code, paste THIS whole file, click Save (disk icon).
    3. Select the function `setup` in the toolbar dropdown → Run.
       Approve permissions when prompted. This creates the tabs
       (Employees, BreakTypes, BreakRecords, Settings, AuditLog) and
       seeds default break types + the logins.
    4. Deploy → New deployment → gear icon → Web app.
         - Execute as: Me
         - Who has access: Anyone
       Click Deploy, copy the Web app URL (ends with /exec).
    5. In index.html, near the top, set:
            const STORAGE_MODE   = "apps_script";   // already set
            const APPS_SCRIPT_URL = "PASTE_/exec_URL_HERE";
    6. Push to GitHub → redeploy on Vercel. Every device now shares live data.

    Default logins after setup:
       ADMIN / 1234        (Admin)
       LEAD-OPS / 1111     (Team Lead)
       EMP-101 / 0001      (Employee)
    To re-seed from scratch, run `resetAll` (wipes all data).
    ============================================================ */

// bound spreadsheet (works whether this script is bound to the sheet or standalone)
const SHEET_ID = "1gtVboWcy9xGs1vXhDZZVEGGyTqSQCFR8OpKd1Mwg7r8";

const SHEETS = {
  employees:  { name: 'Employees',    headers: ['id','empId','name','department','designation','teamLead','email','phone','role','pin','status'] },
  breakTypes: { name: 'BreakTypes',   headers: ['id','name','maxMinutes','color','active'] },
  records:    { name: 'BreakRecords', headers: ['id','employeeId','employeeName','department','teamLead','breakType','breakTypeColor','maxMinutes','startTime','endTime','durationMin','note','date','status'] },
  settings:   { name: 'Settings',     headers: ['key','value'] },
  audit:      { name: 'AuditLog',     headers: ['id','timestamp','employeeName','action','detail'] },
};

const NUMERIC = { maxMinutes: true, durationMin: true };

/* ---------- web entry points ---------- */
function doGet() {
  return json(getAll());
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    handleAction(body.action, body.payload || {});
    return json(getAll());
  } catch (err) {
    return json({ error: String(err) });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- sheet utilities ---------- */
function ss() {
  if (SHEET_ID && SHEET_ID.indexOf("PASTE") === -1) {
    try { return SpreadsheetApp.openById(SHEET_ID); } catch (e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheetFor(key) {
  const def = SHEETS[key];
  let sh = ss().getSheetByName(def.name);
  if (!sh) {
    sh = ss().insertSheet(def.name);
    sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
    sh.getRange(1, 1, 1, def.headers.length).setFontWeight('bold');
    // force text format so empId / pin / ids never lose leading zeros
    sh.getRange(1, 1, sh.getMaxRows(), def.headers.length).setNumberFormat('@');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll(key) {
  const def = SHEETS[key];
  const sh = sheetFor(key);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, def.headers.length).getValues();
  return values
    .filter(r => String(r[0]).trim() !== '')
    .map(r => {
      const o = {};
      def.headers.forEach((h, i) => {
        let v = r[i];
        if (NUMERIC[h]) v = (v === '' || v === null) ? null : Number(v);
        else if (h === 'active') v = (v === true || String(v).toLowerCase() === 'true');
        else v = (v === null) ? '' : String(v);
        o[h] = v;
      });
      return o;
    });
}

function writeAll(key, rows) {
  const def = SHEETS[key];
  const sh = sheetFor(key);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, def.headers.length).clearContent();
  if (!rows.length) return;
  sh.getRange(2, 1, rows.length, def.headers.length).setNumberFormat('@');
  const out = rows.map(o => def.headers.map(h => {
    const v = o[h];
    if (v === null || v === undefined) return '';
    if (h === 'active') return v ? 'TRUE' : 'FALSE';
    return v;
  }));
  sh.getRange(2, 1, out.length, def.headers.length).setValues(out);
}

/* ---------- read everything ---------- */
function getAll() {
  const audit = readAll('audit').sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  const settingsRows = readAll('settings');
  const settings = {};
  settingsRows.forEach(r => settings[r.key] = r.value);
  return {
    employees: readAll('employees'),
    breakTypes: readAll('breakTypes'),
    records: readAll('records'),
    audit: audit.slice(0, 500),
    settings: settings,
  };
}

/* ---------- mutations ---------- */
function uid(p) { return p + '_' + Date.now().toString(36) + '_' + Math.floor(Math.random() * 10000); }
function nowIso() { return new Date().toISOString(); }
function todayStr() {
  const d = new Date();
  const p = n => ('' + n).padStart(2, '0');
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}
function minutesBetween(a, b) { return Math.max(0, (new Date(b).getTime() - new Date(a).getTime()) / 60000); }

function handleAction(action, p) {
  switch (action) {

    case 'startBreak': {
      const emps = readAll('employees');
      const types = readAll('breakTypes');
      const emp = emps.find(e => e.empId === p.employeeId);
      const bt = types.find(b => b.id === p.breakTypeId);
      if (!emp || !bt) return;
      const recs = readAll('records');
      const now = nowIso();
      // close any open break for this employee
      recs.forEach(r => {
        if (r.employeeId === emp.empId && r.status === 'On Break') {
          r.endTime = now;
          r.durationMin = Math.round(minutesBetween(r.startTime, now));
          r.status = 'Returned';
        }
      });
      recs.push({
        id: uid('rec'), employeeId: emp.empId, employeeName: emp.name, department: emp.department, teamLead: emp.teamLead,
        breakType: bt.name, breakTypeColor: bt.color, maxMinutes: Number(bt.maxMinutes),
        startTime: now, endTime: '', durationMin: '', note: p.note || '', date: todayStr(), status: 'On Break',
      });
      writeAll('records', recs);
      break;
    }

    case 'endBreak': {
      const recs = readAll('records');
      const r = recs.find(x => x.id === p.recordId);
      if (r && r.status === 'On Break') {
        const now = nowIso();
        r.endTime = now;
        r.durationMin = Math.round(minutesBetween(r.startTime, now));
        r.status = 'Returned';
        writeAll('records', recs);
      }
      break;
    }

    case 'addEmployee': {
      const emps = readAll('employees');
      emps.push(Object.assign({ id: uid('emp'), status: 'Active' }, p));
      writeAll('employees', emps);
      break;
    }
    case 'updateEmployee': {
      const emps = readAll('employees');
      const i = emps.findIndex(e => e.id === p.id);
      if (i >= 0) { emps[i] = Object.assign(emps[i], p); writeAll('employees', emps); }
      break;
    }
    case 'deleteEmployee': {
      writeAll('employees', readAll('employees').filter(e => e.id !== p.id));
      break;
    }

    case 'addBreakType': {
      const types = readAll('breakTypes');
      types.push(Object.assign({ id: uid('bt'), active: true }, p));
      writeAll('breakTypes', types);
      break;
    }
    case 'updateBreakType': {
      const types = readAll('breakTypes');
      const i = types.findIndex(b => b.id === p.id);
      if (i >= 0) { types[i] = Object.assign(types[i], p); writeAll('breakTypes', types); }
      break;
    }
    case 'deleteBreakType': {
      writeAll('breakTypes', readAll('breakTypes').filter(b => b.id !== p.id));
      break;
    }

    case 'updateSettings': {
      const cur = {};
      readAll('settings').forEach(r => cur[r.key] = r.value);
      Object.keys(p).forEach(k => cur[k] = p[k]);
      writeAll('settings', Object.keys(cur).map(k => ({ key: k, value: cur[k] })));
      break;
    }

    case 'addAudit': {
      const sh = sheetFor('audit');
      const row = [uid('log'), nowIso(), p.employeeName || '—', p.action || '', p.detail || ''];
      sh.appendRow(row);
      sh.getRange(sh.getLastRow(), 1, 1, row.length).setNumberFormat('@').setValues([row]);
      break;
    }

    case 'resetDemo':
    case 'resetAll': {
      resetAll();
      break;
    }
  }
}

/* ---------- setup / seed ---------- */
function setup() {
  Object.keys(SHEETS).forEach(k => sheetFor(k));
  if (readAll('breakTypes').length === 0) seedDefaults();
  // remove the default empty "Sheet1" if present
  const junk = ss().getSheetByName('Sheet1');
  if (junk && ss().getSheets().length > 1) ss().deleteSheet(junk);
}

function resetAll() {
  Object.keys(SHEETS).forEach(k => {
    const sh = sheetFor(k);
    const last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, SHEETS[k].headers.length).clearContent();
  });
  seedDefaults();
}

function seedDefaults() {
  const types = [
    ['Tea Break', 15, '#10B981'], ['Lunch Break', 45, '#6366F1'], ['Prayer (Namaz)', 20, '#22D3EE'],
    ['Phone Call', 10, '#F59E0B'], ['Washroom', 8, '#A78BFA'], ['Meeting', 30, '#0EA5E9'],
    ['Personal Break', 15, '#F472B6'], ['Emergency', 30, '#F43F5E'], ['Other', 15, '#94A3B8'],
  ].map(t => ({ id: uid('bt'), name: t[0], maxMinutes: t[1], color: t[2], active: true }));
  writeAll('breakTypes', types);

  const emps = [
    ['ADMIN', 'System Admin', 'Management', 'Administrator', '', 'Admin', '1234'],
    ['LEAD-OPS', 'Omar Farooq', 'Operations', 'Team Lead', '', 'TeamLead', '1111'],
    ['LEAD-SUP', 'Hina Raza', 'Support', 'Team Lead', '', 'TeamLead', '2222'],
    ['EMP-101', 'Ayesha Khan', 'Operations', 'Executive', 'LEAD-OPS', 'Employee', '0001'],
    ['EMP-102', 'Bilal Ahmed', 'Operations', 'Executive', 'LEAD-OPS', 'Employee', '0002'],
    ['EMP-103', 'Sana Tariq', 'Support', 'Agent', 'LEAD-SUP', 'Employee', '0003'],
    ['EMP-104', 'Hamza Sheikh', 'Support', 'Agent', 'LEAD-SUP', 'Employee', '0004'],
    ['EMP-105', 'Zara Malik', 'Operations', 'Executive', 'LEAD-OPS', 'Employee', '0005'],
  ].map(e => ({
    id: uid('emp'), empId: e[0], name: e[1], department: e[2], designation: e[3],
    teamLead: e[4], email: '', phone: '', role: e[5], pin: e[6], status: 'Active',
  }));
  writeAll('employees', emps);

  writeAll('settings', [{ key: 'company', value: 'Your Company' }]);
}
