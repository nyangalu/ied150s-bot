// ═══════════════════════════════════════════════════════════════
//  IED150S Bot — Google Apps Script Backend
//  1. Paste this into Extensions → Apps Script in your Google Sheet
//  2. Deploy → New deployment → Web app
//     Execute as: Me | Who has access: Anyone
//  3. Copy the Web app URL into SCRIPT_URL in index.html and 2_Dashboard.html
// ═══════════════════════════════════════════════════════════════

// Handle POST (student saves analytics + chat)
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (payload.action === 'saveAnalytics') {
      upsert(getSheet(ss, 'Analytics'), payload.studentNumber, JSON.stringify(payload.data));
      return jsonResponse({ success: true });
    }

    if (payload.action === 'saveChat') {
      upsert(getSheet(ss, 'Chats'), payload.studentNumber, JSON.stringify(payload.messages));
      return jsonResponse({ success: true });
    }

    return jsonResponse({ success: false, error: 'Unknown action: ' + payload.action });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// Handle GET (lecturer dashboard reads analytics + chat)
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = e.parameter.action;

    if (action === 'analytics') {
      const sheet = ss.getSheetByName('Analytics');
      if (!sheet || sheet.getLastRow() < 2) return jsonResponse({ data: [] });
      const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      const data = vals
        .map(row => { try { return JSON.parse(row[1]); } catch { return null; } })
        .filter(Boolean);
      return jsonResponse({ data });
    }

    if (action === 'chat') {
      const sn = e.parameter.sn;
      const sheet = ss.getSheetByName('Chats');
      if (!sheet || sheet.getLastRow() < 2) return jsonResponse({ messages: [] });
      const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      const row = vals.find(r => String(r[0]) === String(sn));
      const messages = row ? JSON.parse(row[1] || '[]') : [];
      return jsonResponse({ messages });
    }

    // Health check — open the URL in browser to test connectivity
    if (action === 'ping') {
      return jsonResponse({ status: 'ok', time: new Date().toISOString() });
    }

    return jsonResponse({ error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
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
    hRow.setFontWeight('bold').setBackground('#185FA5').setFontColor('#ffffff');
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 700);
    sheet.setColumnWidth(3, 180);
  }
  return sheet;
}

function upsert(sheet, key, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
    const idx = keys.indexOf(String(key));
    if (idx >= 0) {
      sheet.getRange(idx + 2, 2).setValue(value);
      sheet.getRange(idx + 2, 3).setValue(new Date().toISOString());
      return;
    }
  }
  sheet.appendRow([String(key), value, new Date().toISOString()]);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
