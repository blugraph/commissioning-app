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
    if (action === 'uploadPhoto') return uploadPhotoAction(body);
    if (action === 'deleteFile')  return deleteFileAction(body);
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

// ── Action: upload a photo to Drive as script owner ──────────────
// Receives base64 image data, creates a Drive file owned by the real
// Google account (not the service account), makes it publicly readable
// so the Docs API can embed it, and returns the file ID.

function uploadPhotoAction(body) {
  var b64  = (body.data || '').replace(/^data:image\/\w+;base64,/, '');
  if (!b64) return respond({ error: 'No image data received' });
  var bytes = Utilities.base64Decode(b64);
  var blob  = Utilities.newBlob(bytes, 'image/jpeg', body.filename || 'photo.jpg');
  var file  = DriveApp.createFile(blob);
  // Publicly readable so Google Docs API can fetch it for inline embedding
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return respond({ fileId: file.getId() });
}

// ── Action: trash a temp photo file after embedding ───────────────

function deleteFileAction(body) {
  try {
    if (body.fileId) DriveApp.getFileById(body.fileId).setTrashed(true);
  } catch(_) {} // ignore if already gone
  return respond({ ok: true });
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
