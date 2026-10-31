const NOTIFICATION_ICON = './pwa-icon-192.png';

self.addEventListener('install', () => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(self.clients.claim());
});

// Sengaja tanpa cache: aplikasi selalu memakai versi terbaru dari Vercel.
self.addEventListener('fetch', () => {});

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
      if (list.some(client => client.focused)) return;
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
