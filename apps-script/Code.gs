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

function doGet(e) {
  return respond({ status: 'ok' });
}

function doPost(e) {
  try {
    var body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    switch (body.action) {
      case 'copyTemplate': return copyTemplateAction(body);
      default:             return getTokenAction();
    }
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

function copyTemplateAction(body) {
  var sa       = JSON.parse(PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT'));
  var template = DriveApp.getFileById(body.templateId);
  var folder   = body.parentId
    ? DriveApp.getFolderById(body.parentId)
    : DriveApp.getRootFolder();

  var copy = template.makeCopy(body.filename, folder);

  // Grant service account editor access so the frontend can fill the doc
  copy.addEditor(sa.client_email);

  return respond({ docId: copy.getId(), url: copy.getUrl() });
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
