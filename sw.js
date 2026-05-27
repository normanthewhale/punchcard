const CACHE = "punchcard-v9";
const ASSETS = ["./index.html", "./manifest.json"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener("activate", e => {
  // Delete all old cache versions
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", e => {
  // Never intercept API calls
  if (e.request.url.includes("script.google.com")) return;
  // Always try network first, fall back to cache if offline
  e.respondWith(
    fetch(e.request, { cache: "no-cache" }).catch(() => caches.match(e.request))
  );
});