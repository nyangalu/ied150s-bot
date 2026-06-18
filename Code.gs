// ═══════════════════════════════════════════════════════════════
//  IED150S Bot — Google Apps Script Backend  v5
//
//  SETUP:
//  1. Paste this into Extensions → Apps Script in your Google Sheet
//  2. Save, then Deploy → New deployment → Web app
//     Execute as: Me  |  Who has access: Anyone
//  3. Copy the Web app URL into SCRIPT_URL in index.html + 2_Dashboard.html
//
//  KNOWLEDGE BASE: Stored directly in Google Sheets (no Drive needed).
//  Each document's text is stored in the KnowledgeBase sheet.
// ═══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const p  = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Analytics ──
    if (p.action === 'saveAnalytics') {
      upsert(getSheet(ss, 'Analytics'), p.studentNumber, JSON.stringify(p.data));
      return ok({ success: true });
    }

    // ── Chat history ──
    if (p.action === 'saveChat') {
      upsert(getSheet(ss, 'Chats'), p.studentNumber, JSON.stringify(p.messages));
      return ok({ success: true });
    }

    // ── Student registry override ──
    if (p.action === 'saveRegistryOverride') {
      upsert(getSheet(ss, 'RegistryOverrides'), p.studentNumber, JSON.stringify(p.entry));
      return ok({ success: true });
    }

    // ── OneDrive / resource folder link ──
    if (p.action === 'saveOneDriveLink') {
      upsert(getSheet(ss, 'Settings'), 'oneDriveLink', p.url || '');
      return ok({ success: true });
    }

    // ── Knowledge Base: save a document (text stored directly in Sheets) ──
    if (p.action === 'saveKbDoc') {
      const doc  = p.doc;
      const meta = JSON.stringify({
        id:      doc.id,
        name:    doc.name,
        ext:     doc.ext,
        size:    doc.size,
        active:  doc.active,
        topics:  doc.topics || [],
        addedAt: doc.addedAt
      });
      const sheet = getSheet(ss, 'KnowledgeBase');
      upsertKb(sheet, doc.id, meta, doc.text || '');
      return ok({ success: true });
    }

    // ── Knowledge Base: toggle active/inactive ──
    if (p.action === 'toggleKbDoc') {
      const sheet = ss.getSheetByName('KnowledgeBase');
      if (sheet && sheet.getLastRow() >= 2) {
        const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
        const idx = ids.indexOf(String(p.id));
        if (idx >= 0) {
          const cell = sheet.getRange(idx + 2, 2);
          try {
            const meta = JSON.parse(cell.getValue());
            meta.active = p.active;
            cell.setValue(JSON.stringify(meta));
          } catch {}
        }
      }
      return ok({ success: true });
    }

    // ── Knowledge Base: delete a document ──
    if (p.action === 'deleteKbDoc') {
      const sheet = ss.getSheetByName('KnowledgeBase');
      if (sheet && sheet.getLastRow() >= 2) {
        const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
        const idx = ids.indexOf(String(p.id));
        if (idx >= 0) sheet.deleteRow(idx + 2);
      }
      return ok({ success: true });
    }

    // ── Knowledge Base: fetch & index a document from a public URL ──
    if (p.action === 'fetchKbUrl') {
      let url = convertOneDriveUrl(p.url);
      let text = '';
      let ext  = 'url';

      try {
        const response = UrlFetchApp.fetch(url, {
          followRedirects: true,
          muteHttpExceptions: true,
          headers: { 'User-Agent': 'Mozilla/5.0' }
        });
        const headers = response.getHeaders();
        const contentType = (headers['Content-Type'] || headers['content-type'] || '').toLowerCase();

        if (contentType.indexOf('pdf') >= 0) {
          return ok({ success: false, error: 'PDF links are not supported yet — please download the PDF and use "Upload File" instead.' });
        }

        text = response.getContentText();
        if (contentType.indexOf('html') >= 0) {
          text = text.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
                     .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
                     .replace(/<[^>]+>/g, ' ')
                     .replace(/&nbsp;/g, ' ')
                     .replace(/\s{2,}/g, ' ').trim();
        }
        ext = 'txt';
      } catch (fetchErr) {
        return ok({ success: false, error: 'Could not fetch URL: ' + fetchErr.message });
      }

      if (!text || text.trim().length < 5) {
        return ok({ success: false, error: 'No readable text found. Make sure the link is set to "Anyone with the link can view" and points directly to a text file.' });
      }

      const id = p.id || Date.now().toString();
      const capped = text.trim().slice(0, 30000);
      const meta = JSON.stringify({
        id, name: p.name, ext, size: capped.length,
        active: true, topics: p.topics || [], addedAt: new Date().toISOString(),
        sourceUrl: p.url
      });
      upsertKb(getSheet(ss, 'KnowledgeBase'), id, meta, capped);

      return ok({ success: true, id, size: capped.length, ext });
    }

    return ok({ success: false, error: 'Unknown action: ' + p.action });
  } catch (err) {
    return ok({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const action = e.parameter.action;

    // ── Analytics ──
    if (action === 'analytics') {
      const sheet = ss.getSheetByName('Analytics');
      if (!sheet || sheet.getLastRow() < 2) return ok({ data: [] });
      const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      return ok({ data: vals.map(r => { try { return JSON.parse(r[1]); } catch { return null; } }).filter(Boolean) });
    }

    // ── Chat history ──
    if (action === 'chat') {
      const sn    = e.parameter.sn;
      const sheet = ss.getSheetByName('Chats');
      if (!sheet || sheet.getLastRow() < 2) return ok({ messages: [] });
      const vals  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      const row   = vals.find(r => String(r[0]) === String(sn));
      return ok({ messages: row ? JSON.parse(row[1] || '[]') : [] });
    }

    // ── Registry overrides ──
    if (action === 'registryOverrides') {
      const sheet = ss.getSheetByName('RegistryOverrides');
      if (!sheet || sheet.getLastRow() < 2) return ok({ overrides: {} });
      const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      const out  = {};
      vals.forEach(r => { try { out[String(r[0])] = JSON.parse(r[1]); } catch {} });
      return ok({ overrides: out });
    }

    // ── Knowledge Base: full admin list (all docs, metadata + short preview, no full text) ──
    // This is what the dashboard calls — Sheets is the single source of truth,
    // so documents always appear correctly even after reload / cache clear / different device.
    if (action === 'kbAdminList') {
      const sheet = ss.getSheetByName('KnowledgeBase');
      const docs  = [];

      if (sheet && sheet.getLastRow() >= 2) {
        const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
        vals.forEach(row => {
          try {
            const meta = JSON.parse(row[1]);
            const text = String(row[2] || '');
            docs.push({
              id: meta.id, name: meta.name, ext: meta.ext,
              size: meta.size || text.length, active: meta.active !== false,
              topics: meta.topics || [], addedAt: meta.addedAt,
              sourceUrl: meta.sourceUrl || '',
              preview: text.slice(0, 180)
            });
          } catch {}
        });
        // Most recently added first
        docs.sort((a, b) => (b.addedAt || '').localeCompare(a.addedAt || ''));
      }

      let oneDriveLink = '';
      try {
        const s = ss.getSheetByName('Settings');
        if (s && s.getLastRow() >= 2) {
          const rows = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
          const r    = rows.find(r => String(r[0]) === 'oneDriveLink');
          if (r) oneDriveLink = r[1] || '';
        }
      } catch {}

      return ok({ docs, oneDriveLink });
    }

    // ── Knowledge Base: return active docs + OneDrive link ──
    if (action === 'knowledgeBase') {
      const sheet = ss.getSheetByName('KnowledgeBase');
      const docs  = [];

      if (sheet && sheet.getLastRow() >= 2) {
        const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
        vals.forEach(row => {
          try {
            const meta = JSON.parse(row[1]);
            if (!meta.active) return;
            const text = String(row[2] || '');
            if (text) docs.push({ id: meta.id, name: meta.name, topics: meta.topics || [], text });
          } catch {}
        });
      }

      // Get OneDrive folder link
      let oneDriveLink = '';
      try {
        const s = ss.getSheetByName('Settings');
        if (s && s.getLastRow() >= 2) {
          const rows = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
          const r    = rows.find(r => String(r[0]) === 'oneDriveLink');
          if (r) oneDriveLink = r[1] || '';
        }
      } catch {}

      return ok({ docs, oneDriveLink });
    }

    // ── Ping / health check ──
    if (action === 'ping') return ok({ status: 'ok', time: new Date().toISOString() });

    return ok({ error: 'Unknown action: ' + action });
  } catch (err) {
    return ok({ error: err.toString() });
  }
}

// ── Convert OneDrive/SharePoint share links to a direct-download URL ──
function convertOneDriveUrl(url) {
  if (!url) return url;
  if (url.indexOf('1drv.ms') >= 0 || url.indexOf('onedrive.live.com') >= 0 || url.indexOf('sharepoint.com') >= 0) {
    try {
      const encoded = Utilities.base64Encode(url).replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
      return 'https://api.onedrive.com/v1.0/shares/u!' + encoded + '/root/content';
    } catch (e) { return url; }
  }
  const gdMatch = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (gdMatch) return 'https://drive.google.com/uc?export=download&id=' + gdMatch[1];
  return url;
}

// ── Helpers ──
function getSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = {
      'Analytics':         ['Student Number', 'Analytics JSON',          'Last Updated'],
      'Chats':             ['Student Number', 'Chat JSON',               'Last Updated'],
      'RegistryOverrides': ['Student Number', 'Entry JSON',              'Last Updated'],
      'KnowledgeBase':     ['Doc ID',         'Metadata JSON',           'Text Content'],
      'Settings':          ['Key',            'Value',                   'Last Updated']
    }[name] || ['Key', 'Value', 'Updated'];
    const hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setValues([headers]).setFontWeight('bold').setBackground('#185FA5').setFontColor('#ffffff');
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 400);
    if (headers.length >= 3) sheet.setColumnWidth(3, name === 'KnowledgeBase' ? 600 : 180);
  }
  return sheet;
}

function upsert(sheet, key, value) {
  const last = sheet.getLastRow();
  if (last >= 2) {
    const keys = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(String);
    const idx  = keys.indexOf(String(key));
    if (idx >= 0) {
      sheet.getRange(idx + 2, 2).setValue(value);
      sheet.getRange(idx + 2, 3).setValue(new Date().toISOString());
      return;
    }
  }
  sheet.appendRow([String(key), value, new Date().toISOString()]);
}

function upsertKb(sheet, id, meta, text) {
  const last = sheet.getLastRow();
  if (last >= 2) {
    const ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat().map(String);
    const idx = ids.indexOf(String(id));
    if (idx >= 0) {
      sheet.getRange(idx + 2, 2).setValue(meta);
      sheet.getRange(idx + 2, 3).setValue(text);
      return;
    }
  }
  sheet.appendRow([String(id), meta, text]);
}

function ok(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
