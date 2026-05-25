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
    if (action === 'fillPhotos') return fillPhotosAction(body);
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

function copyTemplateAction(params) {
  var template = DriveApp.getFileById(params.templateId);
  var folder   = params.parentId
    ? DriveApp.getFolderById(params.parentId)
    : DriveApp.getRootFolder();

  var copy = template.makeCopy(params.filename, folder);

  // Make the doc editable by anyone with the link so the service account
  // (and the Docs/Drive API calls from the app) can fill in text and photos
  // immediately — no per-user permission propagation delay.
  copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  return respond({ docId: copy.getId(), url: copy.getUrl() });
}

// ── Action: embed photos directly into a Google Doc ──────────────
// Receives base64 JPEGs (already resized on the user's device) and
// uses DocumentApp to insert them in-place — no Drive temp files.
//
// body = { docId, photos: { P1: 'data:image/jpeg;base64,…', … } }

function fillPhotosAction(body) {
  var docId  = body.docId;
  var photos = body.photos || {};
  if (!docId) return respond({ error: 'Missing docId' });

  var doc    = DocumentApp.openById(docId);
  var docBody = doc.getBody();
  var errors  = [];

  for (var key in photos) {
    try {
      var b64   = photos[key].replace(/^data:image\/\w+;base64,/, '');
      var bytes = Utilities.base64Decode(b64);
      var blob  = Utilities.newBlob(bytes, 'image/jpeg', key + '.jpg');
      var found = replaceMarkerWithImage(docBody, key, blob);
      if (!found) errors.push(key + ': marker not found in doc');
    } catch(err) {
      errors.push(key + ': ' + err.message);
    }
  }

  doc.saveAndClose();
  return respond({ ok: true, errors: errors });
}

// Finds a cell or paragraph whose entire text equals the marker (e.g. "P1")
// and replaces its content with a sized inline image.
function replaceMarkerWithImage(body, marker, blob) {
  var re = new RegExp('^\\s*' + marker + '\\s*$');

  // ── Search tables ──────────────────────────────────────────────
  var numChildren = body.getNumChildren();
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
    var table = child.asTable();
    for (var r = 0; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      for (var c = 0; c < row.getNumCells(); c++) {
        var cell = row.getCell(c);
        if (!re.test(cell.getText())) continue;
        cell.clear();
        var para = cell.appendParagraph('');
        para.setAttributes({ [DocumentApp.Attribute.SPACING_BEFORE]: 0,
                             [DocumentApp.Attribute.SPACING_AFTER]:  0 });
        var img = para.appendInlineImage(blob);
        sizeImage(img);
        return true;
      }
    }
  }

  // ── Search standalone paragraphs ───────────────────────────────
  for (var j = 0; j < body.getNumChildren(); j++) {
    var el = body.getChild(j);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var para = el.asParagraph();
    if (!re.test(para.getText())) continue;
    var idx = body.getChildIndex(para);
    var newPara = body.insertParagraph(idx, '');
    var img = newPara.appendInlineImage(blob);
    sizeImage(img);
    para.removeFromParent();
    return true;
  }

  return false;
}

// Scale image to fit within the doc's text width (≈ 460 pt for A4).
function sizeImage(img) {
  var MAX_W = 460; // points  (~16.2 cm on A4 with default margins)
  var w = img.getWidth();
  var h = img.getHeight();
  if (w > MAX_W) {
    img.setHeight(Math.round(h * MAX_W / w));
    img.setWidth(MAX_W);
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
