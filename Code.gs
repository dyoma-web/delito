/**
 * Apps Script backend for the DOCX Comments Viewer.
 *
 * SETUP (una sola vez):
 *  1. Crea una Google Sheet vacía y copia su ID (parte entre /d/ y /edit en la URL).
 *  2. Ve a https://script.google.com y crea un proyecto nuevo.
 *  3. Pega este archivo completo en Code.gs.
 *  4. Project Settings (⚙ izquierda) → Script properties → Add script property:
 *       Property: SHEET_ID
 *       Value:    <el ID de tu Sheet>
 *  5. Ejecuta setup() una vez (Run > setup) y autoriza permisos.
 *  6. Deploy > New deployment > Type: Web app
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     Copia la URL del Web app.
 *  7. Pégala en el frontend (botón "Config" → URL de Apps Script).
 */

function getSheetId_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Falta Script Property "SHEET_ID". Configúrala en Project Settings → Script properties.');
  return id;
}

const SCHEMA = {
  documents:    ['doc_id', 'filename', 'color', 'uploaded_at', 'visible'],
  paragraphs:   ['doc_id', 'paragraph_index', 'text'],
  comments:     ['comment_id', 'doc_id', 'paragraph_index', 'comment_text', 'author', 'created_at', 'parent_comment_id', 'resolved'],
  tags:         ['tag_id', 'name', 'color'],
  comment_tags: ['comment_id', 'tag_id']
};

function setup() {
  const ss = SpreadsheetApp.openById(getSheetId_());
  for (const [name, headers] of Object.entries(SCHEMA)) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const empty = firstRow.every(v => v === '');
    if (empty) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  return 'ok';
}

function doGet(e) {
  return jsonOut({ ok: true, data: loadAll() });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const handlers = {
      loadAll, saveDocument, deleteDocument, updateDocument,
      saveTag, updateTag, deleteTag, setCommentTags
    };
    const fn = handlers[body.action];
    if (!fn) throw new Error('Unknown action: ' + body.action);
    return jsonOut({ ok: true, data: fn(body.payload || {}) });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sh(name) {
  return SpreadsheetApp.openById(getSheetId_()).getSheetByName(name);
}

function readAll(name) {
  const sheet = sh(name);
  const range = sheet.getDataRange().getValues();
  if (range.length <= 1) return [];
  const headers = range[0];
  return range.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function appendRows(name, rows) {
  if (!rows.length) return;
  const sheet = sh(name);
  const headers = SCHEMA[name];
  const values = rows.map(r => headers.map(h => r[h] !== undefined ? r[h] : ''));
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
}

function deleteRowsWhere(name, predicate) {
  const sheet = sh(name);
  const range = sheet.getDataRange().getValues();
  const headers = range[0];
  for (let i = range.length - 1; i >= 1; i--) {
    const obj = {};
    headers.forEach((h, j) => obj[h] = range[i][j]);
    if (predicate(obj)) sheet.deleteRow(i + 1);
  }
}

function updateRowsWhere(name, predicate, patch) {
  const sheet = sh(name);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, j) => obj[h] = data[i][j]);
    if (predicate(obj)) {
      Object.keys(patch).forEach(k => {
        const col = headers.indexOf(k);
        if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(patch[k]);
      });
    }
  }
}

function loadAll() {
  return {
    documents:    readAll('documents'),
    paragraphs:   readAll('paragraphs'),
    comments:     readAll('comments'),
    tags:         readAll('tags'),
    comment_tags: readAll('comment_tags')
  };
}

function saveDocument(payload) {
  const { document: doc, paragraphs, comments } = payload;
  const existing = readAll('documents').find(d => d.doc_id === doc.doc_id);
  if (existing) {
    deleteRowsWhere('paragraphs', r => r.doc_id === doc.doc_id);
    deleteRowsWhere('comments',   r => r.doc_id === doc.doc_id);
    updateRowsWhere('documents', r => r.doc_id === doc.doc_id, doc);
  } else {
    appendRows('documents', [doc]);
  }
  appendRows('paragraphs', paragraphs);
  appendRows('comments', comments);
  return { saved: true };
}

function deleteDocument(payload) {
  const { doc_id } = payload;
  const commentIds = readAll('comments')
    .filter(c => c.doc_id === doc_id)
    .map(c => String(c.comment_id));
  deleteRowsWhere('comment_tags', r => commentIds.includes(String(r.comment_id)));
  deleteRowsWhere('comments',     r => r.doc_id === doc_id);
  deleteRowsWhere('paragraphs',   r => r.doc_id === doc_id);
  deleteRowsWhere('documents',    r => r.doc_id === doc_id);
  return { deleted: true };
}

function updateDocument(payload) {
  const { doc_id, patch } = payload;
  updateRowsWhere('documents', r => r.doc_id === doc_id, patch);
  return { updated: true };
}

function saveTag(payload) {
  appendRows('tags', [payload]);
  return { saved: true };
}

function updateTag(payload) {
  const { tag_id, patch } = payload;
  updateRowsWhere('tags', r => String(r.tag_id) === String(tag_id), patch);
  return { updated: true };
}

function deleteTag(payload) {
  const { tag_id } = payload;
  deleteRowsWhere('comment_tags', r => String(r.tag_id) === String(tag_id));
  deleteRowsWhere('tags',         r => String(r.tag_id) === String(tag_id));
  return { deleted: true };
}

function setCommentTags(payload) {
  const { comment_id, tag_ids } = payload;
  deleteRowsWhere('comment_tags', r => String(r.comment_id) === String(comment_id));
  const rows = tag_ids.map(tid => ({ comment_id, tag_id: tid }));
  appendRows('comment_tags', rows);
  return { updated: true };
}
