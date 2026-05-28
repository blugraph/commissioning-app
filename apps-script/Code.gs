/**
 * Commissioning App – Backend
 * ============================
 * Deploy as a Google Apps Script Web App:
 *   Execute as: Me
 *   Who has access: Anyone
 *
 * Script Properties (Project Settings → Script Properties):
 *   SERVICE_ACCOUNT  =  <full contents of the service account JSON key file>
 *
 * After deploying, copy the Web App URL into the ⚙️ settings panel in the app.
 */

// ── Entry points ──────────────────────────────────────────────────

// GET  → used for copyTemplate (params always survive the redirect)
// POST → used for token requests only
function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action) || '';
    if (action === 'copyTemplate') return copyTemplateAction(e.parameter);
    return respond({ status: 'ok' });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function doPost(e) {
  try {
    var body = {};
    try {
      if (e && e.postData && e.postData.contents)
        body = JSON.parse(e.postData.contents);
    } catch(_) {}
    var action = body.action || '';
    if (action === 'copyTemplate') return copyTemplateAction(body);
    if (action === 'fillPhotos')   return fillPhotosAction(body);
    return getTokenAction();
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ── Action: copy template doc as script owner (counts against your quota) ──
//
// The service account cannot own files (it has no Drive storage).
// This action runs as the Apps Script owner, so the new doc is owned
// by your real Google account and uses your 100 GB quota.
// After copying, it grants the service account editor access so it can
// fill in the text markers and insert photos.

// Returns (or creates) a "Commissioning Reports" folder in the script
// owner's My Drive — always accessible since it's the same account.
function getOutputFolder() {
  var NAME = 'Commissioning Reports';
  var iter = DriveApp.getFoldersByName(NAME);
  if (iter.hasNext()) return iter.next();
  var folder = DriveApp.createFolder(NAME);
  console.log('getOutputFolder: created new folder ' + folder.getId());
  return folder;
}

function copyTemplateAction(params) {
  console.log('copyTemplateAction: templateId=' + params.templateId +
              ' filename=' + params.filename +
              ' parentId=' + params.parentId);

  var template = DriveApp.getFileById(params.templateId);

  // Use the caller-supplied folder if provided, otherwise fall back to the
  // self-managed "Commissioning Reports" folder in the script owner's My Drive.
  var folder;
  if (params.parentId) {
    folder = DriveApp.getFolderById(params.parentId);
  } else {
    folder = getOutputFolder();
  }

  var copy = template.makeCopy(params.filename, folder);
  copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  console.log('copyTemplateAction: created docId=' + copy.getId() +
              ' in folder=' + folder.getId());
  return respond({ docId: copy.getId(), url: copy.getUrl() });
}

// ── Action: embed photos directly into a Google Doc ──────────────
// Receives base64 JPEGs (already resized on the user's device) and
// uses DocumentApp to replace every in-place marker (P1-P16, S1-S4)
// with the corresponding image.  Each image is sized to fit its
// container (table cell or full text width) — no cropping.
//
// body = { docId, photos: { P1: 'data:image/jpeg;base64,…', … } }

function fillPhotosAction(body) {
  var docId  = body.docId;
  var photos = body.photos || {};
  if (!docId) return respond({ error: 'Missing docId' });

  var keys = Object.keys(photos);
  console.log('fillPhotosAction: docId=' + docId + ', photos=[' + keys.join(',') + ']');

  var doc     = DocumentApp.openById(docId);
  var docBody = doc.getBody();
  var errors  = [];

  for (var key in photos) {
    try {
      var b64   = photos[key].replace(/^data:image\/\w+;base64,/, '');
      if (!b64) { errors.push(key + ': empty base64 data'); continue; }
      var bytes = Utilities.base64Decode(b64);
      var blob  = Utilities.newBlob(bytes, 'image/jpeg', key + '.jpg');
      var found = replaceMarkerWithImage(docBody, key, blob);
      console.log('fillPhotosAction: ' + key + ' → ' + (found ? 'inserted ✓' : 'MARKER NOT FOUND'));
      if (!found) errors.push(key + ': marker not found in doc');
    } catch(err) {
      console.log('fillPhotosAction: ' + key + ' → error: ' + err.message);
      errors.push(key + ': ' + err.message);
    }
  }

  doc.saveAndClose();
  console.log('fillPhotosAction: done, errors=[' + errors.join('; ') + ']');
  return respond({ ok: true, errors: errors });
}

// Uses body.findText() — searches the ENTIRE document recursively,
// including nested tables, so no manual tree walking needed.
// Finds the paragraph whose full text (trimmed) equals the marker,
// removes it, then inserts the image sized to fit its container.
function replaceMarkerWithImage(body, marker, blob) {
  var hit = body.findText(marker);
  while (hit) {
    var textEl   = hit.getElement();          // Text element
    var para     = textEl.getParent();        // Paragraph containing the text
    var paraText = para.getText().trim();

    // Only replace if the entire paragraph is exactly this marker
    // (avoids matching "P1" inside "P10", "P11", etc.)
    if (paraText === marker) {
      var container = para.getParent();

      if (container.getType() === DocumentApp.ElementType.TABLE_CELL) {
        // ── Marker is inside a table cell ──────────────────────
        // Size the image to fit the actual column width so it never
        // overflows the cell and is never cropped.
        var cell    = container.asTableCell();
        var paraIdx = cell.getChildIndex(para);
        para.removeFromParent();
        var p = cell.insertParagraph(paraIdx, '');
        p.setAttributes({
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.SPACING_AFTER]:  0
        });
        var img = p.appendInlineImage(blob);
        sizeImage(img, cellMaxWidth(cell));

      } else {
        // ── Marker is a standalone paragraph in the body ────────
        var idx = body.getChildIndex(para);
        var p   = body.insertParagraph(idx, '');
        var img = p.appendInlineImage(blob);
        sizeImage(img, 460);                  // full A4 text width
        para.removeFromParent();
      }
      return true;
    }
    hit = body.findText(marker, hit);         // continue searching
  }
  return false;                               // marker not found
}

// Returns the usable inner width (points) of a table cell.
// Reads the column width via Table.getColumnWidth() and subtracts
// cell padding.  Falls back to 220 pt (safe for a 2-column A4 table).
function cellMaxWidth(cell) {
  try {
    var row    = cell.getParentRow();
    var table  = row.getParentTable();
    var colIdx = row.getChildIndex(cell);
    var colW   = table.getColumnWidth(colIdx);   // points, or null/0 if auto
    if (colW && colW > 0) {
      var pad = (cell.getPaddingLeft()  || 0)
              + (cell.getPaddingRight() || 0);
      return Math.max(colW - pad, 40);
    }
  } catch(e) { /* column width unavailable — use fallback */ }
  return 220; // safe default for a 2-column A4 layout
}

// Scale image to fit within maxW points, preserve aspect ratio — no cropping.
function sizeImage(img, maxW) {
  var limit = maxW || 460;
  var w = img.getWidth(), h = img.getHeight();
  if (w > limit) {
    img.setHeight(Math.round(h * limit / w));
    img.setWidth(limit);
  }
}

// ── Action: mint a service-account access token ───────────────────

function getTokenAction() {
  var sa    = JSON.parse(PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT'));
  var token = mintAccessToken(sa);
  return respond({ token: token, expires_in: 3600 });
}

function mintAccessToken(sa) {
  var now   = Math.floor(Date.now() / 1000);
  var scope = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
  ].join(' ');

  var header  = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  var payload = b64url(JSON.stringify({
    iss: sa.client_email,
    sub: sa.client_email,
    scope: scope,
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600,
  }));

  var unsigned  = header + '.' + payload;
  var sigBytes  = Utilities.computeRsaSha256Signature(unsigned, sa.private_key);
  var signature = b64url(sigBytes);
  var jwt       = unsigned + '.' + signature;

  var resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + jwt,
    muteHttpExceptions: true,
  });

  var data = JSON.parse(resp.getContentText());
  if (!data.access_token) throw new Error(data.error_description || data.error || 'Token exchange failed');
  return data.access_token;
}

// ── Helpers ───────────────────────────────────────────────────────

function b64url(input) {
  var bytes = (typeof input === 'string') ? Utilities.newBlob(input).getBytes() : input;
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
