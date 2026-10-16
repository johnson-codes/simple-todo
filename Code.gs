/**
 * Apps Script backend for the To-Do page.
 * Paste into Extensions > Apps Script of a Google Sheet, then Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). Use the /exec URL in the page's "Google Sheets sync" box.
 */
const SHEET_NAME = 'Tasks';
const HEADERS = ['id', 'text', 'done', 'created'];

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

function doGet() {
  return json_({ ok: true, tasks: readTasks_(getSheet_()).reverse() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = JSON.parse(e.postData.contents);
    const sh = getSheet_();

    switch (req.action) {
      case 'list':
        // newest first, matching the page
        return json_({ ok: true, tasks: readTasks_(sh).reverse() });

      case 'add': {
        const t = req.task;
        sh.appendRow([t.id, t.text, !!t.done, t.created]);
        sh.getRange(sh.getLastRow(), 3).insertCheckboxes();
        return json_({ ok: true });
      }

      case 'update': {
        const t = req.task;
        const row = findRow_(sh, t.id);
        if (row < 0) return json_({ ok: false, error: 'Task not found' });
        sh.getRange(row, 2, 1, 2).setValues([[t.text, !!t.done]]);
        return json_({ ok: true });
      }

      case 'delete': {
        const row = findRow_(sh, req.id);
        if (row > 0) sh.deleteRow(row);
        return json_({ ok: true });
      }

      case 'deleteMany': {
        // delete bottom-up so row numbers stay valid
        const rows = req.ids.map(id => findRow_(sh, id)).filter(r => r > 0).sort((a, b) => b - a);
        rows.forEach(r => sh.deleteRow(r));
        return json_({ ok: true });
      }

      default:
        return json_({ ok: false, error: 'Unknown action: ' + req.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
