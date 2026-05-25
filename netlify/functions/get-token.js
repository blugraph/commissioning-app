/**
 * Netlify Function: /api/get-token
 *
 * Mints a short-lived Google OAuth access token from a service account.
 * The service account JSON must be stored in the GOOGLE_SERVICE_ACCOUNT
 * environment variable in Netlify (Site → Environment variables).
 *
 * The frontend uses this token to call Google Drive, Sheets, and Docs APIs
 * directly — no user sign-in required.
 */

const crypto = require('crypto');

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
};

exports.handler = async (event) => {
  // CORS pre-flight
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers: CORS, body: '' };
  }

  try {
    const raw = process.env.GOOGLE_SERVICE_ACCOUNT;
    if (!raw) throw new Error('GOOGLE_SERVICE_ACCOUNT env var is not set.');

    const sa = JSON.parse(raw);

    // Build and sign the JWT claim set (RFC 7523 / Google OAuth2 service account)
    const now = Math.floor(Date.now() / 1000);
    const headerB64  = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
    const payloadB64 = b64url(JSON.stringify({
      iss: sa.client_email,
      sub: sa.client_email,
      scope: [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/spreadsheets.readonly',
      ].join(' '),
      aud: 'https://oauth2.googleapis.com/token',
      iat: now,
      exp: now + 3600,
    }));

    const unsigned  = `${headerB64}.${payloadB64}`;
    const signer    = crypto.createSign('RSA-SHA256');
    signer.update(unsigned);
    const signature = signer.sign(sa.private_key, 'base64url');
    const jwt       = `${unsigned}.${signature}`;

    // Exchange JWT for an access token
    const tokenResp = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
    });

    const tokenData = await tokenResp.json();

    if (!tokenResp.ok || !tokenData.access_token) {
      throw new Error(tokenData.error_description || tokenData.error || 'Token exchange failed');
    }

    return {
      statusCode: 200,
      headers: { ...CORS, 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: JSON.stringify({
        token:      tokenData.access_token,
        expires_in: tokenData.expires_in || 3600,
      }),
    };

  } catch (err) {
    console.error('get-token error:', err.message);
    return {
      statusCode: 500,
      headers: { ...CORS, 'Content-Type': 'application/json' },
      body: JSON.stringify({ error: err.message }),
    };
  }
};

function b64url(str) {
  return Buffer.from(str).toString('base64url');
}
