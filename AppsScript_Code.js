/**
 * Disapprovals Tracker – Apps Script Web App
 *
 * Handles two modes:
 *   POST with "rows"   → Creates a new sheet from data, then moves it
 *   POST with "fileId"  → Moves an existing sheet (legacy/manual)
 *   GET  with "fileId"  → Moves an existing sheet (legacy/manual)
 *
 * To update:
 *   1. Replace the code in your existing Apps Script project with this file
 *   2. Deploy > Manage deployments > Edit > New version > Deploy
 */

/***** CONFIG *****/
const DISAPPROVALS_ROOT_FOLDER_ID = 'PUT_DISAPPROVALS_ROOT_FOLDER_ID_HERE';
const DRIVE_BASE = 'https://www.googleapis.com/drive/v3';

/***** Web app entrypoints *****/

function doGet(e) {
  try {
    var fileId = (e && e.parameter && (e.parameter.fileId || e.parameter.id)) || '';
    var url    = (e && e.parameter && e.parameter.url) || '';
    var id = fileId || extractFileIdFromUrl(url);
    if (!id) return htmlOut('Missing fileId');

    moveDisapprovalToClientFolder(id);
    return htmlOut('OK');
  } catch (err) {
    return htmlOut('ERR: ' + String(err));
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('No payload');
    var body = JSON.parse(e.postData.contents);

    // Mode 1: Create sheet from rows, then move it
    if (body.rows && body.rows.length) {
      var title      = body.title || 'Untitled';
      var rows       = body.rows;
      var clientName = body.clientName || '';
      var clientId   = body.clientId   || '';

      var ss = SpreadsheetApp.create(title);
      var sheet = ss.getActiveSheet();
      sheet.setName('Disapprovals');

      sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
      sheet.setFrozenRows(1);

      try {
        for (var c = 1; c <= rows[0].length; c++) {
          sheet.autoResizeColumn(c);
        }
      } catch (_) {}

      var createdId = ss.getId();
      var createdUrl = ss.getUrl();
      var moved = false;
      var warning = '';

      if (clientName && DISAPPROVALS_ROOT_FOLDER_ID &&
          DISAPPROVALS_ROOT_FOLDER_ID !== 'PUT_DISAPPROVALS_ROOT_FOLDER_ID_HERE') {
        try {
          var folderName = clientName + (clientId ? '_' + clientId : '');
          var targetFolderId = ensureChildFolder(DISAPPROVALS_ROOT_FOLDER_ID, folderName);
          var file = driveGet('/files/' + encodeURIComponent(createdId), {
            fields: 'id,parents',
            supportsAllDrives: 'true',
          });
          var currentParents = (file.parents || []).join(',') || '';
          drivePatch('/files/' + encodeURIComponent(createdId), {
            addParents: targetFolderId,
            removeParents: currentParents,
            supportsAllDrives: 'true',
            fields: 'id,parents',
          });
          moved = true;
        } catch (moveErr) {
          warning = 'Sheet created but move failed: ' + moveErr.message;
        }
      }

      return jsonOut({
        success: true,
        spreadsheetId: createdId,
        url: createdUrl,
        moved: moved,
        warning: warning || undefined,
      });
    }

    // Mode 2: Move existing file by fileId/url
    var id = body.fileId || extractFileIdFromUrl(body.url || '');
    if (!id) return jsonOut({ success: false, error: 'Missing fileId or rows' });

    moveDisapprovalToClientFolder(id);
    return jsonOut({ success: true, moved: true });
  } catch (err) {
    return jsonOut({ success: false, error: String(err) });
  }
}

/***** Core logic *****/
function moveDisapprovalToClientFolder(fileId) {
  var file = driveGet('/files/' + encodeURIComponent(fileId), {
    fields: 'id,name,parents',
    supportsAllDrives: 'true',
  });

  var name = file.name || '';
  var info = parseClientFromDisapprovalName(name);
  if (!info) throw new Error('Filename not recognized: ' + name);

  var folderName = info.clientName + '_' + info.clientId;
  var targetFolderId = ensureChildFolder(DISAPPROVALS_ROOT_FOLDER_ID, folderName);

  var currentParents = (file.parents || []).join(',') || '';
  if ((file.parents || []).indexOf(targetFolderId) !== -1) return;

  drivePatch('/files/' + encodeURIComponent(fileId), {
    addParents: targetFolderId,
    removeParents: currentParents,
    supportsAllDrives: 'true',
    fields: 'id,parents',
  });
}

/***** Helpers *****/

// Matches titles like:
//   "Disapprovals – 411Locals – 2026-02-26"  (created by Node script)
//   "Disapprovals_411Locals_1377"             (legacy)
function parseClientFromDisapprovalName(name) {
  // Try the new format first: "Disapprovals – ClientName – YYYY-MM-DD"
  var m1 = /Disapprovals\s*[\u2013\u2014-]\s*(.+?)\s*[\u2013\u2014-]\s*\d{4}-\d{2}-\d{2}/.exec(name || '');
  if (m1) {
    // clientId isn't in the title in the new format, but the mover
    // still needs it for the folder name. We'll use the name only.
    return { clientName: m1[1].trim(), clientId: '' };
  }

  // Legacy format: "Disapprovals_ClientName_1377"
  var m2 = /^\s*Disapprovals[\s_-]+(.+?)[\s_-]+(\d+)\s*$/i.exec(name || '');
  if (m2) return { clientName: m2[1], clientId: m2[2] };

  return null;
}

function ensureChildFolder(parentId, folderName) {
  var q = [
    "'" + parentId + "' in parents",
    "mimeType='application/vnd.google-apps.folder'",
    'trashed=false',
    "name='" + folderName.replace(/'/g, "\\'") + "'",
  ].join(' and ');

  var res = driveGet('/files', {
    q: q,
    corpora: 'allDrives',
    includeItemsFromAllDrives: 'true',
    supportsAllDrives: 'true',
    pageSize: '10',
    fields: 'files(id,name)',
  });

  if (res.files && res.files.length) return res.files[0].id;

  var created = drivePost('/files', {
    name: folderName,
    mimeType: 'application/vnd.google-apps.folder',
    parents: [parentId],
  }, { supportsAllDrives: 'true', fields: 'id,name' });

  return created.id;
}

function extractFileIdFromUrl(u) {
  if (!u) return '';
  var m = /\/spreadsheets\/d\/([A-Za-z0-9_-]+)/.exec(u);
  return m ? m[1] : '';
}

/***** Low-level Drive HTTP *****/
function driveGet(path, query) {
  return fetchJson(addQuery(DRIVE_BASE + path, query || {}), { method: 'get' });
}
function drivePost(path, resource, query) {
  return fetchJson(addQuery(DRIVE_BASE + path, query || {}), {
    method: 'post',
    payload: JSON.stringify(resource),
  });
}
function drivePatch(path, query) {
  return fetchJson(addQuery(DRIVE_BASE + path, query || {}), { method: 'patch' });
}

function fetchJson(url, opts) {
  var res = UrlFetchApp.fetch(url, {
    method: (opts.method || 'get').toUpperCase(),
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: opts.payload || null,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
  });
  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('HTTP ' + code + ' ' + url + '\n' + text);
  }
  return text ? JSON.parse(text) : {};
}

function addQuery(base, params) {
  var keys = Object.keys(params || {});
  var q = keys
    .filter(function(k) { return params[k] !== undefined && params[k] !== null && params[k] !== ''; })
    .map(function(k) { return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]); })
    .join('&');
  return q ? base + '?' + q : base;
}

/***** Output helpers *****/
function htmlOut(s) {
  return HtmlService.createHtmlOutput(String(s || ''));
}
function jsonOut(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
