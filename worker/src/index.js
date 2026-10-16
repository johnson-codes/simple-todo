/**
 * Image API for the To-Do page, backed by Cloudflare R2.
 *
 *   POST   /images        body = raw image bytes, needs X-Passcode  -> { ok, key }
 *   GET    /images/<key>  public; keys are random UUIDs, so they can't be guessed
 *   DELETE /images/<key>  needs X-Passcode
 */
const MAX_BYTES = 10 * 1024 * 1024;
const TYPES = { jpg: 'image/jpeg', png: 'image/png', webp: 'image/webp', gif: 'image/gif' };
const KEY_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\.(jpg|png|webp|gif)$/;

export default {
  async fetch(request, env) {
    const origin = request.headers.get('Origin') || '';
    const allowed = (env.ALLOWED_ORIGINS || '').split(',').map(s => s.trim()).filter(Boolean);
    const cors = allowed.includes(origin)
      ? { 'Access-Control-Allow-Origin': origin, 'Vary': 'Origin' }
      : {};

    const reply = (status, body) => new Response(JSON.stringify(body), {
      status,
      headers: { 'Content-Type': 'application/json', ...cors },
    });

    if (request.method === 'OPTIONS') {
      if (!cors['Access-Control-Allow-Origin']) return new Response(null, { status: 403 });
      return new Response(null, {
        status: 204,
        headers: {
          ...cors,
          'Access-Control-Allow-Methods': 'POST, DELETE, GET',
          'Access-Control-Allow-Headers': 'Content-Type, X-Passcode',
          'Access-Control-Max-Age': '86400',
        },
      });
    }

    const url = new URL(request.url);
    const parts = url.pathname.split('/').filter(Boolean);
    if (parts[0] !== 'images' || parts.length > 2) return reply(404, { ok: false, error: 'Not found' });
    const key = parts[1];

    try {
      if (request.method === 'GET' && key) {
        if (!KEY_RE.test(key)) return new Response('Not found', { status: 404 });
        const obj = await env.IMAGES.get(key);
        if (!obj) return new Response('Not found', { status: 404 });
        return new Response(obj.body, {
          headers: {
            'Content-Type': TYPES[key.split('.').pop()],
            'Cache-Control': 'public, max-age=31536000, immutable',
            'X-Content-Type-Options': 'nosniff',
            'Content-Security-Policy': "default-src 'none'; sandbox",
            'Cross-Origin-Resource-Policy': 'cross-origin',
          },
        });
      }

      // Everything below changes storage, so it needs the passcode and a browser from an allowed site
      if (!cors['Access-Control-Allow-Origin']) return reply(403, { ok: false, error: 'Origin not allowed' });

      const ip = request.headers.get('CF-Connecting-IP') || 'unknown';
      if (env.LIMITER) {
        const { success } = await env.LIMITER.limit({ key: ip });
        if (!success) return reply(429, { ok: false, error: 'Too many requests, try again in a minute' });
      }

      if (!(await passcodeOk(request.headers.get('X-Passcode'), env.PASSCODE))) {
        return reply(401, { ok: false, error: 'Wrong passcode' });
      }

      if (request.method === 'POST' && !key) {
        const declared = Number(request.headers.get('Content-Length') || 0);
        if (declared > MAX_BYTES) return reply(413, { ok: false, error: 'Image must be 10 MB or smaller' });

        const bytes = new Uint8Array(await request.arrayBuffer());
        if (bytes.byteLength === 0) return reply(400, { ok: false, error: 'Empty upload' });
        if (bytes.byteLength > MAX_BYTES) return reply(413, { ok: false, error: 'Image must be 10 MB or smaller' });

        // Trust the file's actual bytes, not its name or declared type (this also keeps SVG out)
        const ext = sniff(bytes);
        if (!ext) return reply(415, { ok: false, error: 'Only JPEG, PNG, WebP or GIF images are allowed' });

        const newKey = `${crypto.randomUUID()}.${ext}`;
        await env.IMAGES.put(newKey, bytes, { httpMetadata: { contentType: TYPES[ext] } });
        return reply(200, { ok: true, key: newKey });
      }

      if (request.method === 'DELETE' && key) {
        if (!KEY_RE.test(key)) return reply(400, { ok: false, error: 'Invalid key' });
        await env.IMAGES.delete(key);
        return reply(200, { ok: true });
      }

      return reply(405, { ok: false, error: 'Method not allowed' });
    } catch (err) {
      console.error(err);  // details stay in Cloudflare logs
      return reply(500, { ok: false, error: 'Server error' });
    }
  },
};

function sniff(b) {
  if (b.length < 12) return null;
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'jpg';
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png';
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38) return 'gif';
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 &&
      b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return 'webp';
  return null;
}

// Compare SHA-256 digests in constant time so the passcode can't be guessed by timing
async function passcodeOk(given, expected) {
  if (!given || !expected) return false;
  const enc = new TextEncoder();
  const [a, b] = await Promise.all([
    crypto.subtle.digest('SHA-256', enc.encode(given)),
    crypto.subtle.digest('SHA-256', enc.encode(expected)),
  ]);
  return crypto.subtle.timingSafeEqual(a, b);
}
