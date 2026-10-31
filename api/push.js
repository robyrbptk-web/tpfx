const crypto = require('crypto');
const webpush = require('web-push');

const PUBLIC_KEY = String(process.env.VAPID_PUBLIC_KEY || '').trim();
const PRIVATE_KEY = String(process.env.VAPID_PRIVATE_KEY || '').trim();
const RELAY_SECRET = String(process.env.PUSH_RELAY_SECRET || '').trim();
const SUBJECT = String(process.env.VAPID_SUBJECT || 'mailto:admin@tpfx.app').trim();

const MAX_SUBS = 100;

const ALLOWED_HOSTS = [
  'fcm.googleapis.com',
  'android.googleapis.com',
  'web.push.apple.com',
  'updates.push.services.mozilla.com',
  'push.services.mozilla.com'
];

function isConfigured() {
  return !!(PUBLIC_KEY && PRIVATE_KEY && RELAY_SECRET);
}

function secretOk(value) {
  const provided = Buffer.from(String(value || ''));
  const expected = Buffer.from(RELAY_SECRET);
  if (!provided.length || provided.length !== expected.length) return false;
  return crypto.timingSafeEqual(provided, expected);
}

function clean(value, max) {
  return String(value == null ? '' : value).replace(/\s+/g, ' ').trim().slice(0, max);
}

function endpointAllowed(endpoint) {
  let url;
  try {
    url = new URL(endpoint);
  } catch (e) {
    return false;
  }
  if (url.protocol !== 'https:') return false;
  const host = url.hostname.toLowerCase();
  return ALLOWED_HOSTS.includes(host) || host.endsWith('.notify.windows.com');
}

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('Content-Type', 'application/json; charset=utf-8');

  if (req.method === 'GET') {
    const enabled = !!(PUBLIC_KEY && PRIVATE_KEY);
    return res.status(enabled ? 200 : 503).json({
      enabled: enabled,
      publicKey: enabled ? PUBLIC_KEY : ''
    });
  }

  if (req.method !== 'POST') {
    res.setHeader('Allow', 'GET, POST');
    return res.status(405).json({ error: 'Method not allowed' });
  }

  if (!isConfigured()) {
    return res.status(503).json({ error: 'Push relay belum dikonfigurasi' });
  }

  if (!secretOk(req.headers['x-push-secret'])) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  const body = req.body && typeof req.body === 'object' ? req.body : {};
  const note = body.notification && typeof body.notification === 'object' ? body.notification : {};
  const subs = Array.isArray(body.subs) ? body.subs.slice(0, MAX_SUBS) : [];

  if (!subs.length) {
    return res.status(200).json({ ok: true, sent: 0, failed: 0, gone: [] });
  }

  const payload = JSON.stringify({
    title: clean(note.title, 90) || 'TPFX Workspace',
    body: clean(note.body, 180),
    url: clean(note.url, 300) || './index.html',
    tag: clean(note.tag, 60) || 'tpfx',
    ts: Date.now()
  });

  webpush.setVapidDetails(SUBJECT, PUBLIC_KEY, PRIVATE_KEY);

  let sent = 0;
  let failed = 0;
  const gone = [];

  await Promise.all(subs.map(async sub => {
    const endpoint = sub && typeof sub.endpoint === 'string' ? sub.endpoint : '';
    const keys = sub && sub.keys;

    if (!endpoint || !endpointAllowed(endpoint) || !keys || !keys.p256dh || !keys.auth) {
      failed++;
      return;
    }

    try {
      await webpush.sendNotification(
        {
          endpoint: endpoint,
          keys: { p256dh: String(keys.p256dh), auth: String(keys.auth) }
        },
        payload,
        { TTL: 86400, urgency: 'high' }
      );
      sent++;
    } catch (e) {
      failed++;
      const status = Number(e && e.statusCode);
      if (status === 404 || status === 410) gone.push(endpoint);
    }
  }));

  return res.status(200).json({ ok: true, sent: sent, failed: failed, gone: gone });
};
