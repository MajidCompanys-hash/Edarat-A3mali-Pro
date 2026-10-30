// إدارة أعمالي — يجعل التطبيق يفتح بدون إنترنت
const CACHE = "idarat-a3mali-v5";
const CORE = ["./", "./index.html", "./manifest.webmanifest", "./icons/icon-192.png", "./icons/icon-512.png", "./icons/apple-touch-icon.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET") return;
  // Google sign-in and Drive always go straight to the network
  const isFont = url.host === "fonts.googleapis.com" || url.host === "fonts.gstatic.com";
  if (url.origin !== location.origin && !isFont) return;
  // App page: network first so updates arrive, cache when offline
  if (e.request.mode === "navigate" || url.pathname.endsWith("index.html")) {
    e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put("./index.html", c)); return r; }).catch(() => caches.match("./index.html")));
    return;
  }
  // Icons and fonts: cache first
  {
    e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request).then(r => { if (r.ok || r.type === "opaque") { const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); } return r; })));
  }
});
