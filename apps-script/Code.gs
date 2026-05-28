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
// uses DocumentApp to:
//   1. Replace in-place markers (P1, P7 … S1-S4) throughout the doc
//   2. Append a "Photo Evidence" gallery page at the end with ALL
//      photos together in a compact 3-column grid, no cropping.
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

  // 1 ── Replace in-place markers throughout the document ──────────
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

  // 2 ── Append photo gallery page at the end ──────────────────────
  try {
    appendPhotoGallery(docBody, photos);
    console.log('fillPhotosAction: gallery page appended ✓');
  } catch(err) {
    console.log('fillPhotosAction: gallery error: ' + err.message);
    errors.push('gallery: ' + err.message);
  }

  doc.saveAndClose();
  console.log('fillPhotosAction: done, errors=[' + errors.join('; ') + ']');
  return respond({ ok: true, errors: errors });
}

// ── Append a photo gallery page at the end of the document ──────────
// All photos are placed in a 3-column grid.  Each image is scaled to
// fit the column width (no cropping); portrait shots remain portrait.
function appendPhotoGallery(docBody, photos) {

  // Canonical display order
  var ORDER  = ['P1','P2','P3','P4','P5','P6','P7',
                'P10','P11','P12','P13','P14','P15','P16',
                'S1','S2','S3','S4'];
  var LABELS = {
    P1:'Sensor / Dry Canal',     P2:'Water Level Meas.',
    P3:'Tape & Pipe',            P4:'Station View',
    P5:'Footing',                P6:'Internal',
    P7:'Panel Height',           P10:'Water Level Display',
    P11:'Station ID Label',      P12:'Metadata Screen',
    P13:'Website Dashboard',     P14:'Dual SIM Cards',
    P15:'3-Point Calibration',   P16:'Dashboard Screenshot',
    S1:'SMS Alert 1',            S2:'SMS Alert 2',
    S3:'SMS Alert 3',            S4:'SMS Alert 4'
  };

  // Build ordered list of keys that actually have photo data
  var keys = ORDER.filter(function(k) { return !!photos[k]; });
  // Append any extra keys not in ORDER (future-proof)
  Object.keys(photos).forEach(function(k) {
    if (ORDER.indexOf(k) === -1 && photos[k]) keys.push(k);
  });
  if (keys.length === 0) return;

  var COLS  = 3;
  var IMG_W = 140; // points — max image width per cell

  // ── Title on a new page ──────────────────────────────────────────
  var title = docBody.appendParagraph('Photo Evidence');
  title.setPageBreakBefore(true);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.editAsText().setFontSize(14).setBold(true);
  // Small breathing room below the heading
  title.setSpacingAfter(8);

  // ── 3-column photo table ─────────────────────────────────────────
  var numRows = Math.ceil(keys.length / COLS);
  var table = docBody.appendTable();

  for (var r = 0; r < numRows; r++) {
    var tr = table.appendTableRow();
    for (var c = 0; c < COLS; c++) {
      var idx  = r * COLS + c;
      var cell = tr.appendTableCell();
      cell.setPaddingTop(4);
      cell.setPaddingBottom(4);
      cell.setPaddingLeft(4);
      cell.setPaddingRight(4);

      if (idx < keys.length) {
        var key = keys[idx];

        // Label (reuse the empty paragraph that exists when cell is created)
        var labelPara = cell.getChild(0).asParagraph();
        labelPara.setText((LABELS[key] || key) + ' (' + key + ')');
        labelPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        labelPara.editAsText().setFontSize(7).setBold(true).setForegroundColor('#444444');
        labelPara.setSpacingAfter(2);

        // Photo — resize to column width, no cropping
        try {
          var b64   = photos[key].replace(/^data:image\/\w+;base64,/, '');
          var bytes = Utilities.base64Decode(b64);
          var blob  = Utilities.newBlob(bytes, 'image/jpeg', key + '.jpg');
          var imgPara = cell.appendParagraph('');
          imgPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          imgPara.setSpacingBefore(0);
          imgPara.setSpacingAfter(0);
          var img = imgPara.appendInlineImage(blob);
          sizeGalleryImg(img, IMG_W);
        } catch(err) {
          cell.appendParagraph('(error)');
          console.log('gallery ' + key + ': ' + err.message);
        }
      }
    }
  }
}

// Uses body.findText() — searches the ENTIRE document recursively,
// including nested tables, so no manual tree walking needed.
// Finds the paragraph whose full text (trimmed) equals the marker,
// clears it, then inserts the image directly in its place.
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
        // Remove ONLY the Pn paragraph – keep the title paragraph(s) above it
        var cell     = container.asTableCell();
        var paraIdx  = cell.getChildIndex(para);
        para.removeFromParent();              // remove only the "Pn" marker
        var p = cell.insertParagraph(paraIdx, '');
        p.setAttributes({
          [DocumentApp.Attribute.SPACING_BEFORE]: 0,
          [DocumentApp.Attribute.SPACING_AFTER]:  0
        });
        var img = p.appendInlineImage(blob);
        sizeImage(img);

      } else {
        // ── Marker is a standalone paragraph in the body ────────
        var idx = body.getChildIndex(para);
        var p   = body.insertParagraph(idx, '');
        var img = p.appendInlineImage(blob);
        sizeImage(img);
        para.removeFromParent();              // remove the Pn paragraph
      }
      return true;
    }
    hit = body.findText(marker, hit);         // continue searching
  }
  return false;                               // marker not found
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

// Scale gallery thumbnail to fit within maxW, preserve aspect ratio — no cropping.
function sizeGalleryImg(img, maxW) {
  var w = img.getWidth(), h = img.getHeight();
  if (w > maxW) {
    img.setHeight(Math.round(h * maxW / w));
    img.setWidth(maxW);
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
