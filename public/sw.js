/* Service worker: يخزّن التطبيق ليشتغل بدون إنترنت + يعرض الإشعارات */
const CACHE = 'myschedule-v5';
const SHELL = ['./', 'index.html', 'ics.js', 'config.js', 'manifest.webmanifest',
  'icons/icon.svg', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).catch(() => {}).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// الشبكة أولاً (عشان التحديثات توصل)، وإذا ما فيه إنترنت نستخدم النسخة المخزّنة
self.addEventListener('fetch', e => {
  const req = e.request;
  const u = new URL(req.url);
  if (req.method !== 'GET' || u.origin !== location.origin || u.pathname.startsWith('/api/') || u.pathname.startsWith('/proxy')) return;
  e.respondWith(
    fetch(req).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
      return res;
    }).catch(() => caches.match(req).then(r => r || caches.match('index.html')))
  );
});

self.addEventListener('notificationclick', e => {
  if (!(e.notification.data && e.notification.data.pin)) e.notification.close();
  e.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
      for (const c of list) { if ('focus' in c) return c.focus(); }
      return self.clients.openWindow('./');
    })
  );
});

// إشعار Push من الخادم (يصل حتى والتطبيق مقفل). الخادم يرسل إشارة بدون نص، ونجلب النص منه.
self.addEventListener('push', e => {
  e.waitUntil((async () => {
    let items = [];
    try {
      const sub = await self.registration.pushManager.getSubscription();
      if (sub) {
        const r = await fetch('/api/pending', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ endpoint: sub.endpoint }) });
        if (r.ok) items = (await r.json()).items || [];
      }
    } catch (_) {}
    if (!items.length) items = [{ title: 'جدولي', body: 'عندك تنبيه جديد · You have a new alert', tag: 'jadwali-generic' }];
    await Promise.all(items.map(n => self.registration.showNotification(n.title, {
      body: n.body, tag: n.tag, lang: n.lang, dir: n.dir,
      icon: 'icons/icon-192.png', badge: 'icons/icon-192.png', vibrate: [120, 60, 120], data: { url: './' }
    })));
  })());
});
