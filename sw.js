// Service worker: szybkie uruchamianie i działanie ekranu bez sieci.
// Dane (produkty, wpisy) zawsze idą do Google Sheets, więc do zapisu potrzebny jest internet.
const CACHE = "magazyn-qr-v2";
const SHELL = ["./", "./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png",
  "https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/qrcode-generator/1.4.4/qrcode.min.js"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE)
    .then(c => Promise.all(SHELL.map(u => c.add(u).catch(() => {}))))
    .then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.hostname.endsWith("google.com") || url.hostname.endsWith("googleusercontent.com")) return;
  // najpierw sieć (świeża wersja), a bez sieci — kopia z pamięci
  e.respondWith(fetch(e.request).then(r => {
    if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
    return r;
  }).catch(() => caches.match(e.request)));
});
