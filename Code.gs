/**
 * Apps Script backend for the To-Do page.
 * Paste into Extensions > Apps Script of a Google Sheet, then Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). Connect a browser once by opening the page with #api=<the /exec URL>.
 *
 * @OnlyCurrentDoc  Limits the script's access to this one spreadsheet.
 */
const SHEET_NAME = 'Tasks';
const HEADERS = ['id', 'text', 'done', 'created'];
const MAX_TEXT = 500;
const MAX_TASKS = 2000;
const ID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const DATE_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z$/;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readTasks_(sh) {
  const rows = sh.getDataRange().getValues().slice(1);
  return rows
    .filter(r => r[0])
    .map(r => ({ id: String(r[0]), text: String(r[1]), done: r[2] === true, created: String(r[3]) }));
}

function findRow_(sh, id) {
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 1;
  return -1;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Thrown for bad input; its message is safe to send back to the client
function BadRequest_(msg) { this.message = msg; }

function checkId_(id) {
  if (typeof id !== 'string' || !ID_RE.test(id)) throw new BadRequest_('Invalid id');
  return id;
}

function checkText_(text) {
  if (typeof text !== 'string') throw new BadRequest_('Invalid text');
  text = text.trim();
  if (!text || text.length > MAX_TEXT) throw new BadRequest_('Text must be 1-' + MAX_TEXT + ' characters');
  // Leading apostrophe stores the value as plain text, so "=..." is never run as a formula
  return "'" + text;
}

function checkDone_(done) {
  if (typeof done !== 'boolean') throw new BadRequest_('Invalid done');
  return done;
}

function checkCreated_(created) {
  if (typeof created !== 'string' || !DATE_RE.test(created)) throw new BadRequest_('Invalid created');
  return created;
}

function doGet() {
  return json_({ ok: true, tasks: readTasks_(getSheet_()).reverse() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let req;
    try { req = JSON.parse(e.postData.contents); } catch (_) { throw new BadRequest_('Invalid JSON'); }
    if (!req || typeof req !== 'object') throw new BadRequest_('Invalid request');
    const sh = getSheet_();

    switch (req.action) {
      case 'list':
        // newest first, matching the page
        return json_({ ok: true, tasks: readTasks_(sh).reverse() });

      case 'add': {
        const t = req.task || {};
        const row = [checkId_(t.id), checkText_(t.text), checkDone_(t.done), checkCreated_(t.created)];
        if (sh.getLastRow() - 1 >= MAX_TASKS) throw new BadRequest_('Task limit reached');
        if (findRow_(sh, t.id) > 0) throw new BadRequest_('Duplicate id');
        sh.appendRow(row);
        sh.getRange(sh.getLastRow(), 3).insertCheckboxes();
        return json_({ ok: true });
      }

      case 'update': {
        const t = req.task || {};
        const values = [[checkText_(t.text), checkDone_(t.done)]];
        const row = findRow_(sh, checkId_(t.id));
        if (row < 0) throw new BadRequest_('Task not found');
        sh.getRange(row, 2, 1, 2).setValues(values);
        return json_({ ok: true });
      }

      case 'delete': {
        const row = findRow_(sh, checkId_(req.id));
        if (row > 0) sh.deleteRow(row);
        return json_({ ok: true });
      }

      case 'deleteMany': {
        if (!Array.isArray(req.ids) || req.ids.length > MAX_TASKS) throw new BadRequest_('Invalid ids');
        // delete bottom-up so row numbers stay valid
        const rows = req.ids.map(id => findRow_(sh, checkId_(id))).filter(r => r > 0).sort((a, b) => b - a);
        rows.forEach(r => sh.deleteRow(r));
        return json_({ ok: true });
      }

      default:
        throw new BadRequest_('Unknown action');
    }
  } catch (err) {
    if (err instanceof BadRequest_) return json_({ ok: false, error: err.message });
    console.error(err);  // details stay in the Apps Script execution log
    return json_({ ok: false, error: 'Server error' });
  } finally {
    lock.releaseLock();
  }
}
