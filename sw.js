/* 金邊鍋貼 PWA service worker */
const VERSION = 'v1';
const SHELL = 'golden-edge-shell-' + VERSION;
const RUNTIME = 'golden-edge-runtime-' + VERSION;
const SHELL_FILES = ['./', 'index.html', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png'];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(SHELL).then(cache => cache.addAll(SHELL_FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== SHELL && k !== RUNTIME).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET' || !req.url.startsWith('http')) return;
  const sameOrigin = new URL(req.url).origin === self.location.origin;

  // 頁面與本站檔案：先連網拿最新版，離線時才用快取。
  if (req.mode === 'navigate' || sameOrigin) {
    event.respondWith(
      fetch(req).then(res => {
        if (res.ok) { const copy = res.clone(); caches.open(SHELL).then(c => c.put(req, copy)); }
        return res;
      }).catch(() => caches.match(req).then(hit => hit || (req.mode === 'navigate' ? caches.match('index.html') : Response.error())))
    );
    return;
  }

  // 外部圖片、字體與套件：先用快取，同時在背景更新。
  event.respondWith(
    caches.open(RUNTIME).then(cache => cache.match(req).then(hit => {
      const fresh = fetch(req).then(res => {
        if (res.ok || res.type === 'opaque') cache.put(req, res.clone());
        return res;
      }).catch(() => hit || Response.error());
      return hit || fresh;
    }))
  );
});
