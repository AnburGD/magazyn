// Service worker: szybkie uruchamianie i działanie ekranu bez sieci.
// Dane (produkty, wpisy) zawsze idą do Google Sheets, więc do zapisu potrzebny jest internet.
// Numer musi być taki sam jak APP_VER w index.html
const CACHE = "magazyn-qr-v22";
const SHELL = ["./", "./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png",
  "./lib/html5-qrcode.min.js", "./lib/qrcode.js", "./lib/jsQR.js"];

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
  // najpierw sieć (świeża wersja), a bez sieci — kopia z pamięci.
  // Własne pliki zawsze sprawdzamy na serwerze (no-cache), żeby nowa wersja nie czekała w pamięci telefonu.
  let req = e.request;
  if (url.origin === self.location.origin)
    req = e.request.mode === "navigate" ? new Request(e.request.url, { cache: "no-cache", credentials: "same-origin" }) : new Request(e.request, { cache: "no-cache" });
  e.respondWith(fetch(req).then(r => {
    if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
    return r;
  }).catch(() => caches.match(e.request)));
});
