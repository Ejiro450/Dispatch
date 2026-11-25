// Keeps the app shell available offline. Calendar data always comes live from Microsoft.
// Every request checks GitHub for a newer copy first (no-cache = revalidate), so uploads show on the next open.
const VERSION = "eam-app-v30";
const SHELL = ["./", "index.html", "quotes.html", "quotes-bridge.js", "config.js", "msal-browser.min.js", "manifest.webmanifest", "pdf-lib.js", "logo.png", "geo-vic.js", "icons/icon-192.png", "icons/icon-512.png"];
self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== VERSION).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET" || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(
    fetch(req, { cache: "no-cache" }).then((res) => {
      if (res.ok) { const copy = res.clone(); caches.open(VERSION).then((c) => c.put(req, copy)); }
      return res;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then((hit) => hit || caches.match("index.html")))
  );
});
