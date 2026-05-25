/**
 * Commissioning App – Token Proxy
 * ================================
 * Deploy this as a Google Apps Script Web App:
 *   Execute as: Me
 *   Who has access: Anyone
 *
 * Setup:
 *   1. In the script editor: Project Settings → Script Properties
 *   2. Add property:  SERVICE_ACCOUNT  =  <paste full JSON key file contents>
 *   3. Deploy → New deployment → Web App → Execute as Me, Anyone → Deploy
 *   4. Copy the Web App URL into CFG.tokenEndpoint in index.html
 */

// ── CORS pre-flight ───────────────────────────────────────────────
function doGet(e) {
  return respond({ status: 'ok' });
}

// ── Token request ─────────────────────────────────────────────────
function doPost(e) {
  try {
    var sa = JSON.parse(
      PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT')
    );
    if (!sa || !sa.private_key) throw new Error('SERVICE_ACCOUNT property not configured.');

    var token = mintAccessToken(sa);
    return respond({ token: token, expires_in: 3600 });
  } catch (err) {
    return respond({ error: err.message }, 500);
  }
}

// ── Mint a service-account access token via JWT ───────────────────
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
  var bytes  = (typeof input === 'string') ? Utilities.newBlob(input).getBytes() : input;
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
