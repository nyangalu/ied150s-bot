// ═══════════════════════════════════════════════════════════════
//  IED150S Bot — Google Apps Script Backend  v3
//  1. Paste into Extensions → Apps Script in your Google Sheet
//  2. Deploy → New deployment → Web app
//     Execute as: Me  |  Who has access: Anyone
//  3. Copy the Web app URL into SCRIPT_URL in index.html + 2_Dashboard.html
// ═══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (payload.action === 'saveAnalytics') {
      upsert(getSheet(ss,'Analytics'), payload.studentNumber, JSON.stringify(payload.data));
      return ok({ success:true });
    }
    if (payload.action === 'saveChat') {
      upsert(getSheet(ss,'Chats'), payload.studentNumber, JSON.stringify(payload.messages));
      return ok({ success:true });
    }
    if (payload.action === 'saveRegistryOverride') {
      upsert(getSheet(ss,'RegistryOverrides'), payload.studentNumber, JSON.stringify(payload.entry));
      return ok({ success:true });
    }
    // Knowledge Base: add or update a document
    if (payload.action === 'saveKbDoc') {
      const sheet = getSheet(ss,'KnowledgeBase');
      const doc = payload.doc;
      // Store: id | metadata JSON (name,ext,size,active,addedAt) | text content
      const meta = JSON.stringify({ id:doc.id, name:doc.name, ext:doc.ext, size:doc.size, active:doc.active, addedAt:doc.addedAt });
      upsertKb(sheet, doc.id, meta, doc.text || '');
      return ok({ success:true });
    }
    // Knowledge Base: delete a document
    if (payload.action === 'deleteKbDoc') {
      const sheet = ss.getSheetByName('KnowledgeBase');
      if (sheet && sheet.getLastRow() >= 2) {
        const ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat().map(String);
        const idx = ids.indexOf(String(payload.id));
        if (idx >= 0) sheet.deleteRow(idx + 2);
      }
      return ok({ success:true });
    }

    return ok({ success:false, error:'Unknown action: '+payload.action });
  } catch (err) {
    return ok({ success:false, error:err.toString() });
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = e.parameter.action;

    if (action === 'analytics') {
      const sheet = ss.getSheetByName('Analytics');
      if (!sheet || sheet.getLastRow() < 2) return ok({ data:[] });
      const vals = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();
      return ok({ data: vals.map(r=>{ try{return JSON.parse(r[1]);}catch{return null;} }).filter(Boolean) });
    }
    if (action === 'chat') {
      const sn = e.parameter.sn;
      const sheet = ss.getSheetByName('Chats');
      if (!sheet || sheet.getLastRow() < 2) return ok({ messages:[] });
      const vals = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();
      const row = vals.find(r=>String(r[0])===String(sn));
      return ok({ messages: row ? JSON.parse(row[1]||'[]') : [] });
    }
    if (action === 'registryOverrides') {
      const sheet = ss.getSheetByName('RegistryOverrides');
      if (!sheet || sheet.getLastRow() < 2) return ok({ overrides:{} });
      const vals = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();
      const overrides = {};
      vals.forEach(r=>{ try{overrides[String(r[0])]=JSON.parse(r[1]);}catch{} });
      return ok({ overrides });
    }
    // Knowledge Base: return active documents for the student bot
    if (action === 'knowledgeBase') {
      const sheet = ss.getSheetByName('KnowledgeBase');
      if (!sheet || sheet.getLastRow() < 2) return ok({ docs:[] });
      const vals = sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues();
      const docs = [];
      vals.forEach(r=>{
        try {
          const meta = JSON.parse(r[1]);
          if (meta.active) docs.push({ id:meta.id, name:meta.name, text:String(r[2]) });
        } catch {}
      });
      return ok({ docs });
    }
    if (action === 'ping') return ok({ status:'ok', time:new Date().toISOString() });

    return ok({ error:'Unknown action: '+action });
  } catch (err) {
    return ok({ error:err.toString() });
  }
}

// ── HELPERS ──
function getSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const hdrs = {
      'Analytics':         ['Student Number','Analytics JSON','Last Updated'],
      'Chats':             ['Student Number','Chat JSON','Last Updated'],
      'RegistryOverrides': ['Student Number','Entry JSON','Last Updated'],
      'KnowledgeBase':     ['Doc ID','Metadata JSON','Text Content']
    }[name] || ['Key','Value','Updated'];
    const hr = sheet.getRange(1,1,1,3);
    hr.setValues([hdrs]).setFontWeight('bold').setBackground('#185FA5').setFontColor('#ffffff');
    sheet.setColumnWidth(1,140); sheet.setColumnWidth(2,300); sheet.setColumnWidth(3,700);
  }
  return sheet;
}

function upsert(sheet, key, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const keys = sheet.getRange(2,1,lastRow-1,1).getValues().flat().map(String);
    const idx = keys.indexOf(String(key));
    if (idx >= 0) {
      sheet.getRange(idx+2,2).setValue(value);
      sheet.getRange(idx+2,3).setValue(new Date().toISOString());
      return;
    }
  }
  sheet.appendRow([String(key), value, new Date().toISOString()]);
}

function upsertKb(sheet, id, meta, text) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const ids = sheet.getRange(2,1,lastRow-1,1).getValues().flat().map(String);
    const idx = ids.indexOf(String(id));
    if (idx >= 0) {
      sheet.getRange(idx+2,2).setValue(meta);
      sheet.getRange(idx+2,3).setValue(text);
      return;
    }
  }
  sheet.appendRow([String(id), meta, text]);
}

function ok(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
