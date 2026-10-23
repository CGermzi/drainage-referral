// Keeps the app working without signal. The page comes from the internet when there is signal (so updates arrive)
// and from the copy saved on the phone when there isn't. Change VERSION when the CORE list changes.
const VERSION = "v2";
const CORE = ["./", "index.html", "manifest.webmanifest", "logo.svg", "icon-192.png", "icon-512.png", "apple-touch-icon.png",
  "reference/drainage-areas-1.jpg", "reference/drainage-areas-2.jpg", "reference/drainage-areas-3.jpg", "reference/drainage-poster.jpg"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET" || new URL(req.url).origin !== location.origin) return;   // feedback and map lookups go straight out
  if (req.mode === "navigate") {
    // the page: fresh when the internet answers within 3 s, otherwise the saved copy
    e.respondWith((async () => {
      const cache = await caches.open(VERSION);
      try {
        const res = await Promise.race([fetch(req), new Promise((_, fail) => setTimeout(fail, 3000))]);
        if (res.ok) cache.put("index.html", res.clone());
        return res;
      } catch {
        return (await cache.match("index.html")) || (await cache.match("./")) || Response.error();
      }
    })());
    return;
  }
  // icons and reference pictures: the saved copy straight away, refreshed in the background
  e.respondWith(caches.open(VERSION).then(async cache => {
    const saved = await cache.match(req);
    const fresh = fetch(req).then(res => { if (res.ok) cache.put(req, res.clone()); return res; }).catch(() => saved || Response.error());
    return saved || fresh;
  }));
});
