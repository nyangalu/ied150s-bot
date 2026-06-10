// ═══════════════════════════════════════════════════════════════
//  IED150S Bot — Google Apps Script Backend  v4
//
//  SETUP:
//  1. Paste this into Extensions → Apps Script in your Google Sheet
//  2. Deploy → New deployment → Web app
//     Execute as: Me  |  Who has access: Anyone
//  3. Copy the Web app URL into SCRIPT_URL in index.html + 2_Dashboard.html
//
//  KNOWLEDGE BASE STORAGE:
//  Files are stored in Google Drive (unlimited space).
//  A folder called "IED150S Knowledge Base" is auto-created in your Drive.
//  The KnowledgeBase sheet stores metadata only — content stays in Drive.
// ═══════════════════════════════════════════════════════════════

// ── Get or create the KB folder in Google Drive ──
function getKbFolder() {
  const folderName = 'IED150S Knowledge Base';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

// ── POST handler ──
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Analytics ──
    if (payload.action === 'saveAnalytics') {
      upsert(getSheet(ss, 'Analytics'), payload.studentNumber, JSON.stringify(payload.data));
      return ok({ success: true });
    }

    // ── Chat history ──
    if (payload.action === 'saveChat') {
      upsert(getSheet(ss, 'Chats'), payload.studentNumber, JSON.stringify(payload.messages));
      return ok({ success: true });
    }

    // ── Student registry override ──
    if (payload.action === 'saveRegistryOverride') {
      upsert(getSheet(ss, 'RegistryOverrides'), payload.studentNumber, JSON.stringify(payload.entry));
      return ok({ success: true });
    }

    // ── OneDrive folder link (shown to students as "View Resources") ──
    if (payload.action === 'saveOneDriveLink') {
      upsert(getSheet(ss, 'Settings'), 'oneDriveLink', payload.url || '');
      return ok({ success: true });
    }

    // ── KB: Upload large file in chunks ──
    // Each chunk is appended to a Drive file. When all chunks arrive, finalise.
    if (payload.action === 'saveKbChunk') {
      const folder = getKbFolder();
      const chunkId = payload.id + '_chunk' + payload.chunkIndex;

      if (payload.chunkIndex === 0) {
        // First chunk — create or overwrite the main file
        const existing = folder.getFilesByName(payload.id + '.txt');
        while (existing.hasNext()) existing.next().setTrashed(true);
        folder.createFile(payload.id + '.txt', payload.chunk, MimeType.PLAIN_TEXT);
      } else {
        // Append subsequent chunks by reading existing + appending
        const files = folder.getFilesByName(payload.id + '.txt');
        if (files.hasNext()) {
          const file = files.next();
          const existing = file.getBlob().getDataAsString();
          file.setContent(existing + payload.chunk);
        }
      }

      // On last chunk, save metadata to sheet
      if (payload.chunkIndex === payload.totalChunks - 1) {
        const files2 = folder.getFilesByName(payload.id + '.txt');
        const driveId = files2.hasNext() ? files2.next().getId() : '';
        const meta = JSON.stringify({
          id: payload.id,
          name: payload.name,
          ext: payload.ext,
          size: payload.totalSize,
          active: true,
          addedAt: new Date().toISOString(),
          driveId: driveId
        });
        upsert(getSheet(ss, 'KnowledgeBase'), payload.id, meta);
      }

      return ok({ success: true });
    }

    // ── KB: Fetch a document from a public URL ──
    if (payload.action === 'fetchKbUrl') {
      let url = payload.url;
      let text = '';
      let ext  = 'url';

      // Convert OneDrive share links to direct download
      url = convertOneDriveUrl(url);

      try {
        const response = UrlFetchApp.fetch(url, {
          followRedirects: true,
          muteHttpExceptions: true,
          headers: { 'User-Agent': 'Mozilla/5.0' }
        });

        const contentType = response.getHeaders()['Content-Type'] || '';

        if (contentType.includes('pdf')) {
          // Save the raw PDF to Drive and extract text via Drive OCR
          const blob = response.getBlob().setName(payload.id + '.pdf');
          const folder = getKbFolder();
          const driveFile = folder.createFile(blob);
          // Use Drive to convert PDF → Google Doc → extract text
          const resource = { title: payload.id, mimeType: MimeType.GOOGLE_DOCS };
          const docFile  = Drive.Files.copy(resource, driveFile.getId());
          const doc      = DocumentApp.openById(docFile.id);
          text = doc.getBody().getText();
          // Clean up the temp Google Doc
          DriveApp.getFileById(docFile.id).setTrashed(true);
          driveFile.setTrashed(true);
          ext = 'pdf';
        } else {
          text = response.getContentText();
          // Strip HTML tags if it looks like HTML
          if (contentType.includes('html')) {
            text = text.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
                       .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
                       .replace(/<[^>]+>/g, ' ')
                       .replace(/\s{2,}/g, ' ').trim();
          }
          ext = 'txt';
        }
      } catch (fetchErr) {
        return ok({ success: false, error: 'Could not fetch URL: ' + fetchErr.message });
      }

      if (!text || text.length < 10) {
        return ok({ success: false, error: 'No readable text found at that URL. Make sure the link is public and points to a text or PDF file.' });
      }

      // Save to Drive
      const folder = getKbFolder();
      const driveFile = folder.createFile(payload.id + '.txt', text, MimeType.PLAIN_TEXT);
      const meta = JSON.stringify({
        id: payload.id,
        name: payload.name,
        ext: ext,
        size: text.length,
        active: true,
        addedAt: new Date().toISOString(),
        driveId: driveFile.getId(),
        sourceUrl: payload.url
      });
      upsert(getSheet(ss, 'KnowledgeBase'), payload.id, meta);

      return ok({ success: true, id: payload.id, size: text.length, ext });
    }

    // ── KB: Toggle active/inactive ──
    if (payload.action === 'toggleKbDoc') {
      const sheet = ss.getSheetByName('KnowledgeBase');
      if (sheet && sheet.getLastRow() >= 2) {
        const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
        const idx = ids.indexOf(String(payload.id));
        if (idx >= 0) {
          const metaCell = sheet.getRange(idx + 2, 2);
          try {
            const meta = JSON.parse(metaCell.getValue());
            meta.active = payload.active;
            metaCell.setValue(JSON.stringify(meta));
          } catch {}
        }
      }
      return ok({ success: true });
    }

    // ── KB: Delete a document ──
    if (payload.action === 'deleteKbDoc') {
      const sheet = ss.getSheetByName('KnowledgeBase');
      if (sheet && sheet.getLastRow() >= 2) {
        const ids  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String);
        const idx  = ids.indexOf(String(payload.id));
        if (idx >= 0) {
          // Also delete the Drive file
          try {
            const meta = JSON.parse(sheet.getRange(idx + 2, 2).getValue());
            if (meta.driveId) DriveApp.getFileById(meta.driveId).setTrashed(true);
          } catch {}
          sheet.deleteRow(idx + 2);
        }
      }
      return ok({ success: true });
    }

    return ok({ success: false, error: 'Unknown action: ' + payload.action });
  } catch (err) {
    return ok({ success: false, error: err.toString() });
  }
}

// ── GET handler ──
function doGet(e) {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const action = e.parameter.action;

    if (action === 'analytics') {
      const sheet = ss.getSheetByName('Analytics');
      if (!sheet || sheet.getLastRow() < 2) return ok({ data: [] });
      const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      return ok({ data: vals.map(r => { try { return JSON.parse(r[1]); } catch { return null; } }).filter(Boolean) });
    }

    if (action === 'chat') {
      const sn    = e.parameter.sn;
      const sheet = ss.getSheetByName('Chats');
      if (!sheet || sheet.getLastRow() < 2) return ok({ messages: [] });
      const vals  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      const row   = vals.find(r => String(r[0]) === String(sn));
      return ok({ messages: row ? JSON.parse(row[1] || '[]') : [] });
    }

    if (action === 'registryOverrides') {
      const sheet = ss.getSheetByName('RegistryOverrides');
      if (!sheet || sheet.getLastRow() < 2) return ok({ overrides: {} });
      const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      const overrides = {};
      vals.forEach(r => { try { overrides[String(r[0])] = JSON.parse(r[1]); } catch {} });
      return ok({ overrides });
    }

    // ── KB: Return active document contents + OneDrive folder link ──
    if (action === 'knowledgeBase') {
      const BOT_CTX_CHARS = 80000; // 80 KB per doc max
      const sheet = ss.getSheetByName('KnowledgeBase');
      const docs = [];

      if (sheet && sheet.getLastRow() >= 2) {
        const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
        vals.forEach(row => {
          try {
            const meta = JSON.parse(row[1]);
            if (!meta.active) return;
            let text = '';
            if (meta.driveId) {
              text = DriveApp.getFileById(meta.driveId).getBlob().getDataAsString().slice(0, BOT_CTX_CHARS);
            }
            if (text) docs.push({ id: meta.id, name: meta.name, topics: meta.topics || [], text });
          } catch {}
        });
      }

      // Get OneDrive folder link from Settings sheet
      let oneDriveLink = '';
      try {
        const settingsSheet = ss.getSheetByName('Settings');
        if (settingsSheet && settingsSheet.getLastRow() >= 2) {
          const rows = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 2).getValues();
          const odRow = rows.find(r => String(r[0]) === 'oneDriveLink');
          if (odRow) oneDriveLink = odRow[1] || '';
        }
      } catch {}

      return ok({ docs, oneDriveLink });
    }

    // ── KB: Get single document full text (for future preview feature) ──
    if (action === 'getKbDoc') {
      const sheet = ss.getSheetByName('KnowledgeBase');
      if (!sheet || sheet.getLastRow() < 2) return ok({ text: '' });
      const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      const row  = vals.find(r => String(r[0]) === String(e.parameter.id));
      if (!row) return ok({ text: '' });
      try {
        const meta = JSON.parse(row[1]);
        const file = DriveApp.getFileById(meta.driveId);
        return ok({ text: file.getBlob().getDataAsString() });
      } catch { return ok({ text: '' }); }
    }

    if (action === 'ping') return ok({ status: 'ok', time: new Date().toISOString() });

    return ok({ error: 'Unknown action: ' + action });
  } catch (err) {
    return ok({ error: err.toString() });
  }
}

// ── Convert OneDrive share links to direct download URLs ──
function convertOneDriveUrl(url) {
  // Already a direct link
  if (url.includes('download') && !url.includes('1drv.ms') && !url.includes('onedrive.live.com')) return url;

  // OneDrive short link: 1drv.ms/...
  // We use the sharing API: encode URL as base64url, prefix with "u!"
  if (url.includes('1drv.ms') || url.includes('onedrive.live.com') || url.includes('sharepoint.com')) {
    try {
      const encoded = Utilities.base64Encode(url)
        .replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
      return 'https://api.onedrive.com/v1.0/shares/u!' + encoded + '/root/content';
    } catch { return url; }
  }

  // Google Drive view link → direct download
  const gdMatch = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (gdMatch) {
    return 'https://drive.google.com/uc?export=download&id=' + gdMatch[1];
  }

  return url;
}

// ── HELPERS ──
function getSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const hdrs = {
      'Analytics':         ['Student Number', 'Analytics JSON', 'Last Updated'],
      'Chats':             ['Student Number', 'Chat JSON', 'Last Updated'],
      'RegistryOverrides': ['Student Number', 'Entry JSON', 'Last Updated'],
      'KnowledgeBase':     ['Doc ID', 'Metadata JSON', 'Last Updated'],
      'Settings':          ['Key', 'Value', 'Last Updated']
    }[name] || ['Key', 'Value', 'Updated'];
    const hr = sheet.getRange(1, 1, 1, 3);
    hr.setValues([hdrs]).setFontWeight('bold').setBackground('#185FA5').setFontColor('#ffffff');
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 500);
    sheet.setColumnWidth(3, 180);
  }
  return sheet;
}

function upsert(sheet, key, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String);
    const idx  = keys.indexOf(String(key));
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
