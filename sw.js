const NOTIFICATION_ICON = './pwa-icon-192.png';
const CREDS_DB = 'tpfx-push';
const CREDS_STORE = 'creds';
const CREDS_KEY = 'me';

self.addEventListener('install', () => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(self.clients.claim());
});

// Sengaja tanpa cache: aplikasi selalu memakai versi terbaru dari Vercel.
self.addEventListener('fetch', () => {});

function openCredsDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(CREDS_DB, 1);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(CREDS_STORE)) {
        request.result.createObjectStore(CREDS_STORE);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function writeCreds(creds) {
  const db = await openCredsDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(CREDS_STORE, 'readwrite');
    tx.objectStore(CREDS_STORE).put(creds, CREDS_KEY);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}

async function readCreds() {
  const db = await openCredsDb();
  const value = await new Promise((resolve, reject) => {
    const tx = db.transaction(CREDS_STORE, 'readonly');
    const request = tx.objectStore(CREDS_STORE).get(CREDS_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
  db.close();
  return value;
}

// Chrome mengganti endpoint push secara berkala. Tanpa pendaftaran ulang,
// notifikasi berhenti diam-diam, jadi worker menyimpan kredensial login
// untuk menyambung langganan sendiri tanpa harus membuka aplikasinya.
async function registerSubscription(sub) {
  const creds = await readCreds();
  if (!creds || !creds.token || !creds.api) return false;
  const payload = sub && typeof sub.toJSON === 'function' ? sub.toJSON() : sub;
  if (!payload || !payload.endpoint) return false;
  const res = await fetch(creds.api, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ fn: 'subscribePush', args: [creds.token, JSON.stringify(payload)] })
  });
  const data = await res.json().catch(() => null);
  return !!(data && data.result && data.result.ok);
}

async function resubscribePush() {
  const creds = await readCreds();
  if (!creds || !creds.token || !creds.key) throw new Error('kredensial push belum tersimpan');
  const sub = await self.registration.pushManager.subscribe({
    userVisibleOnly: true,
    applicationServerKey: creds.key
  });
  return registerSubscription(sub);
}

self.addEventListener('pushsubscriptionchange', event => {
  event.waitUntil((async () => {
    try {
      const incoming = event.data && event.data.newSubscription;
      const sub = incoming || await self.registration.pushManager.getSubscription();
      const ok = sub ? await registerSubscription(sub) : await resubscribePush();
      if (ok) return;
      throw new Error('langganan push gagal diperbarui');
    } catch (e) {
      await self.registration.showNotification('Notifikasi TPFX perlu disambungkan ulang', {
        body: 'Buka aplikasi TPFX sekali saja agar pesan dan panggilan masuk kembali ke layar HP Anda.',
        icon: NOTIFICATION_ICON,
        badge: NOTIFICATION_ICON,
        tag: 'tpfx-push-refresh',
        lang: 'id',
        data: { url: './index.html' }
      });
    }
  })());
});

self.addEventListener('push', event => {
  event.waitUntil((async () => {
    let data = {};
    try {
      data = event.data ? event.data.json() : {};
    } catch (e) {}

    const title = String(data.title || 'TPFX Workspace');
    const ts = Number(data.ts) || Date.now();
    const options = {
      body: String(data.body || 'Anda menerima pesan baru'),
      icon: NOTIFICATION_ICON,
      badge: NOTIFICATION_ICON,
      tag: String(data.tag || 'tpfx'),
      renotify: !!data.tag,
      data: { url: String(data.url || './index.html'), ts: ts },
      vibrate: [90, 40, 90],
      lang: 'id',
      dir: 'auto',
      silent: false,
      timestamp: ts
    };

    try {
      const list = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
      const target = new URL(String(data.url || './index.html'), self.registration.scope).href;
      if (String(data.tag || '').indexOf('tpfx-call-') === 0 && target.indexOf('#call=') >= 0) {
        list.forEach(client => {
          try { client.postMessage({ type: 'incoming-call', url: target }); } catch (e) {}
        });
      }
      if (list.some(client => client.focused) && options.tag !== 'tpfx-test') return;
    } catch (e) {}

    await self.registration.showNotification(title, options);
  })());
});

self.addEventListener('notificationclick', event => {
  event.notification.close();
  const url = (event.notification.data && event.notification.data.url) || './index.html';
  const target = new URL(url, self.registration.scope).href;

  event.waitUntil((async () => {
    const list = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const client of list) {
      if (new URL(client.url).origin !== new URL(target).origin) continue;
      try { await client.focus(); } catch (e) {}
      client.postMessage({ type: 'open-chat', url: target });
      return;
    }
    await self.clients.openWindow(target);
  })());
});

self.addEventListener('message', event => {
  const data = event.data || {};
  if (data.type === 'tpfx-push-creds' && data.token && data.key && data.api) {
    event.waitUntil(writeCreds({
      token: String(data.token),
      key: String(data.key),
      api: String(data.api),
      endpoint: String(data.endpoint || ''),
      saved: Date.now()
    }).catch(() => {}));
    return;
  }
  if (data.type !== 'notify' || !data.title) return;
  self.registration.showNotification(String(data.title), {
    body: String(data.body || ''),
    icon: NOTIFICATION_ICON,
    badge: NOTIFICATION_ICON,
    tag: String(data.tag || 'tpfx-local'),
    data: { url: String(data.url || './index.html'), ts: Date.now() },
    lang: 'id',
    dir: 'auto',
    silent: false,
    timestamp: Date.now()
  });
});
