// ═══════════════════════════════════════════════════════════════
//  IED150S Bot — Google Apps Script Backend
//  Paste this entire file into the Apps Script editor.
//  See deployment instructions in the chat.
// ═══════════════════════════════════════════════════════════════

// ── WRITE: called by student_bot.html after each message ──
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (payload.action === 'saveAnalytics') {
      upsert(getSheet(ss, 'Analytics'), payload.studentNumber, JSON.stringify(payload.data));
      return ok({ success: true });
    }

    if (payload.action === 'saveChat') {
      upsert(getSheet(ss, 'Chats'), payload.studentNumber, JSON.stringify(payload.messages));
      return ok({ success: true });
    }

    return ok({ success: false, error: 'Unknown action: ' + payload.action });
  } catch (err) {
    return ok({ success: false, error: err.toString() });
  }
}

// ── READ: called by lecturer_dashboard.html ──
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = e.parameter.action;

  // Get all student analytics
  if (action === 'analytics') {
    const sheet = ss.getSheetByName('Analytics');
    if (!sheet || sheet.getLastRow() < 2) return ok({ data: [] });
    const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    const data = vals
      .map(row => { try { return JSON.parse(row[1]); } catch { return null; } })
      .filter(Boolean);
    return ok({ data });
  }

  // Get one student's chat history
  if (action === 'chat') {
    const sn = e.parameter.sn;
    const sheet = ss.getSheetByName('Chats');
    if (!sheet || sheet.getLastRow() < 2) return ok({ messages: [] });
    const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    const row = vals.find(r => r[0] === sn);
    const messages = row ? JSON.parse(row[1] || '[]') : [];
    return ok({ messages });
  }

  return ok({ error: 'Unknown action' });
}

// ── HELPERS ──

function getSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = name === 'Analytics'
      ? ['Student Number', 'Analytics JSON', 'Last Updated']
      : ['Student Number', 'Chat JSON', 'Last Updated'];
    const hRow = sheet.getRange(1, 1, 1, 3);
    hRow.setValues([headers]);
    hRow.setFontWeight('bold').setBackground('#0d1b3e').setFontColor('#f0a500');
    sheet.setColumnWidth(2, 600);
  }
  return sheet;
}

function upsert(sheet, key, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const idx = keys.indexOf(String(key));
    if (idx >= 0) {
      sheet.getRange(idx + 2, 2).setValue(value);
      sheet.getRange(idx + 2, 3).setValue(new Date().toISOString());
      return;
    }
  }
  sheet.appendRow([String(key), value, new Date().toISOString()]);
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
