// ════════════════════════════════════════════════════════════
//  Cloudflare Worker — Machine Inspection Proxy
//  วางโค้ดนี้ใน Cloudflare Workers แล้ว Deploy
// ════════════════════════════════════════════════════════════

// ── ใส่ Apps Script URL ที่นี่ ──────────────────────────────
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin':  '*',
  'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export default {
  async fetch(request) {

    // Preflight CORS
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: CORS_HEADERS });
    }

    // Health check
    if (request.method === 'GET') {
      return new Response(
        JSON.stringify({ status: 'ok', service: 'Machine Inspection Proxy' }),
        { headers: { 'Content-Type': 'application/json', ...CORS_HEADERS } }
      );
    }

    // POST → ส่งไป Apps Script
    if (request.method === 'POST') {
      try {
        const body = await request.text();

        const gasResp = await fetch(APPS_SCRIPT_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body
        });

        const result = await gasResp.json().catch(() => ({ status: 'ok' }));

        return new Response(
          JSON.stringify({ status: 'ok', ...result }),
          { headers: { 'Content-Type': 'application/json', ...CORS_HEADERS } }
        );

      } catch(err) {
        return new Response(
          JSON.stringify({ status: 'error', message: err.message }),
          { status: 500, headers: { 'Content-Type': 'application/json', ...CORS_HEADERS } }
        );
      }
    }

    return new Response('Method Not Allowed', { status: 405, headers: CORS_HEADERS });
  }
};
