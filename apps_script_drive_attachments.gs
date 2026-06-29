/**
 * SSC Expense System - Google Sheets + Drive Attachments API
 *
 * วิธีใช้:
 * 1. เปิด Google Apps Script
 * 2. วางโค้ดนี้แทน Code.gs
 * 3. Deploy > New deployment > Web app
 * 4. Execute as: Me
 * 5. Who has access: Anyone with the link
 * 6. Copy Web app URL มาใส่หน้า "ตั้งค่า" ในเว็บ
 */

const API_USER = 'boy_ssc';
const API_PASS = 'SSC2569_Boy!@#';
const DATA_SHEET_NAME = 'Expenses';
const ROOT_FOLDER_NAME = 'SSC Expense System';
const RECEIPT_FOLDER_NAME = 'Receipts';
const PROP_SPREADSHEET_ID = 'SSC_EXPENSE_SPREADSHEET_ID';
const HEADERS = ['savedAt', 'docNo', 'id', 'issueDate', 'customer', 'total', 'attachmentCount', 'json'];

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (!isAuthorized_(p.user, p.pass)) return json_({ success: false, error: 'Unauthorized' });

    const action = p.action || 'list';
    if (action === 'ping') {
      return json_({ success: true, time: new Date().toISOString() });
    }

    if (action === 'list') {
      const items = listItems_();
      return json_({ success: true, count: items.length, items });
    }

    return json_({ success: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ success: false, error: String(err && err.message ? err.message : err) });
  }
}

function doPost(e) {
  try {
    const payload = parsePayload_(e);
    if (!isAuthorized_(payload.user, payload.pass)) return json_({ success: false, error: 'Unauthorized' });

    const action = payload.action || 'save_with_files';

    if (action === 'clear') {
      clearItems_();
      return json_({ success: true, count: 0 });
    }

    if (action === 'bulk_create' || action === 'bulk_create_with_files') {
      clearItems_();
      const items = (payload.items || []).map(item => prepareRecord_(item));
      appendItems_(items);
      return json_({ success: true, count: items.length, items });
    }

    if (action === 'save_with_files') {
      const rows = payload.rows || [];
      const docNo = payload.docNo || (rows[0] && rows[0].docNo) || '';
      if (docNo) deleteDoc_(docNo);
      const items = rows.map(item => prepareRecord_(item));
      appendItems_(items);
      return json_({ success: true, count: items.length, items });
    }

    return json_({ success: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ success: false, error: String(err && err.message ? err.message : err) });
  }
}

function parsePayload_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  return JSON.parse(raw);
}

function isAuthorized_(user, pass) {
  return String(user || '') === API_USER && String(pass || '') === API_PASS;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROP_SPREADSHEET_ID);
  let ss;

  if (id) {
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (err) {
      id = '';
    }
  }

  if (!id) {
    ss = SpreadsheetApp.create('SSC Expense System Data');
    props.setProperty(PROP_SPREADSHEET_ID, ss.getId());
  }

  ensureSheet_((ss || SpreadsheetApp.openById(id)).getSheetByName(DATA_SHEET_NAME) || (ss || SpreadsheetApp.openById(id)).insertSheet(DATA_SHEET_NAME));
  return ss || SpreadsheetApp.openById(id);
}

function getSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(DATA_SHEET_NAME);
  ensureSheet_(sheet);
  return sheet;
}

function ensureSheet_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    return;
  }

  const current = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needsHeader = HEADERS.some((h, i) => current[i] !== h);
  if (needsHeader) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function getRootFolder_() {
  const folders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
}

function getReceiptFolder_() {
  const root = getRootFolder_();
  const folders = root.getFoldersByName(RECEIPT_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : root.createFolder(RECEIPT_FOLDER_NAME);
}

function prepareRecord_(item) {
  const record = Object.assign({}, item || {});
  const docNo = sanitizeFileName_(record.docNo || 'NO-DOC');
  const rowId = sanitizeFileName_(record.id || Utilities.getUuid());
  const folder = getReceiptFolder_();

  const attachments = Array.isArray(record.attachments) ? record.attachments : [];
  record.attachments = attachments.map((att, index) => saveAttachment_(folder, docNo, rowId, att, index + 1));
  record.attachmentCount = record.attachments.length;
  record.savedAt = record.savedAt || new Date().toISOString();
  return record;
}

function saveAttachment_(folder, docNo, rowId, att, index) {
  const out = Object.assign({}, att || {});

  if (!out.data || !String(out.data).startsWith('data:')) {
    const src = out.data || out.thumbnailUrl || out.driveThumbnailUrl || out.url || out.driveUrl || '';
    out.data = src;
    out.url = out.url || src;
    return out;
  }

  const parsed = parseDataUrl_(out.data);
  const ext = mimeToExt_(parsed.mimeType);
  const fileName = sanitizeFileName_(`${docNo}_${rowId}_${String(index).padStart(2, '0')}.${ext}`);
  const blob = Utilities.newBlob(parsed.bytes, parsed.mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const thumb = `https://drive.google.com/thumbnail?id=${fileId}&sz=w1600`;

  out.driveFileId = fileId;
  out.driveUrl = file.getUrl();
  out.thumbnailUrl = thumb;
  out.url = thumb;
  out.data = thumb;
  out.mimeType = parsed.mimeType;
  out.name = out.name || fileName;
  return out;
}

function parseDataUrl_(dataUrl) {
  const match = String(dataUrl).match(/^data:([^;]+);base64,(.+)$/);
  if (!match) throw new Error('Invalid attachment data URL');
  return {
    mimeType: match[1],
    bytes: Utilities.base64Decode(match[2])
  };
}

function mimeToExt_(mimeType) {
  const map = {
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/png': 'png',
    'image/webp': 'webp',
    'image/gif': 'gif'
  };
  return map[String(mimeType || '').toLowerCase()] || 'jpg';
}

function sanitizeFileName_(value) {
  return String(value || '')
    .replace(/[\\/:*?"<>|#%{}~&]/g, '-')
    .replace(/\s+/g, '_')
    .slice(0, 140);
}

function appendItems_(items) {
  if (!items.length) return;
  const sheet = getSheet_();
  const values = items.map(item => [
    item.savedAt || new Date().toISOString(),
    item.docNo || '',
    item.id || '',
    item.issueDate || item.date || '',
    item.customer || '',
    Number(item.total || 0),
    Number(item.attachmentCount || 0),
    JSON.stringify(item)
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, HEADERS.length).setValues(values);
}

function listItems_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return values
    .map(row => {
      try {
        return JSON.parse(row[7] || '{}');
      } catch (err) {
        return null;
      }
    })
    .filter(Boolean);
}

function clearItems_() {
  const sheet = getSheet_();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).clearContent();
  }
}

function deleteDoc_(docNo) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][1]) === String(docNo)) {
      sheet.deleteRow(i + 2);
    }
  }
}
