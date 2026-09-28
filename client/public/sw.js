const CACHE = "cherry-checklist-v4";
const BASE = self.registration.scope;
const CORE = [BASE, `${BASE}manifest.json`];
self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(CORE)));
  self.skipWaiting();
});
self.addEventListener("activate", (event) => {
  event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((key) => key !== CACHE).map((key) => caches.delete(key)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  const request = event.request;
  event.respondWith(caches.match(request).then((cached) => cached || fetch(request).then((response) => { const clone = response.clone(); caches.open(CACHE).then((cache) => cache.put(request, clone)); return response; }).catch(() => caches.match(BASE))));
});
