/* Service worker: membuat aplikasi bisa dipasang & halaman jamaah tetap terbuka saat sinyal buruk.
 * Aman: hanya menyimpan aset statis dan halaman PUBLIK (tanda header X-App-Public: 1).
 * Halaman pengurus, data login, dan permintaan POST tidak pernah disimpan. */
const VERSION = "v3";
const STATIC = "static-" + VERSION;
const PAGES = "pages-" + VERSION;
const PRECACHE = ["/offline", "/icons/icon-192.png", "/icons/icon-512.png", "/images/logo-dkm.png"];
const PUBLIC_PATHS = ["/", "/jadwal-ramadan", "/info-masjid", "/donasi", "/transparansi"];
const ASSET_PREFIX = ["/vendor/", "/css/", "/images/", "/icons/", "/qris/"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(STATIC).then((cache) => cache.addAll(PRECACHE)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== STATIC && k !== PAGES).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

function isPublicPage(url) {
  return PUBLIC_PATHS.some((p) => url.pathname === p);
}

// Aset statis: tampilkan dari simpanan lebih dulu, perbarui di latar belakang
async function staleWhileRevalidate(request) {
  const cache = await caches.open(STATIC);
  const cached = await cache.match(request, { ignoreVary: true });
  const fresh = fetch(request)
    .then((response) => {
      if (response && response.ok) cache.put(request, response.clone());
      return response;
    })
    .catch(() => cached);
  return cached || fresh;
}

// Halaman publik: coba jaringan dulu (data terbaru), jika gagal/lambat pakai simpanan terakhir
async function networkFirst(request) {
  const cache = await caches.open(PAGES);
  try {
    const response = await Promise.race([
      fetch(request),
      new Promise((_, reject) => setTimeout(() => reject(new Error("timeout")), 6000)),
    ]);
    if (response && response.ok && response.headers.get("X-App-Public") === "1") {
      cache.put(request, response.clone());
    }
    return response;
  } catch (error) {
    const cached = await cache.match(request, { ignoreVary: true });
    if (cached) return cached;
    return (await caches.match("/offline", { ignoreVary: true })) || Response.error();
  }
}

self.addEventListener("fetch", (event) => {
  const { request } = event;
  if (request.method !== "GET") return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;
  if (url.pathname.startsWith("/uploads/") || url.pathname.startsWith("/auth")) return; // privat: biarkan lewat jaringan

  if (request.mode === "navigate") {
    if (isPublicPage(url)) {
      event.respondWith(networkFirst(request));
    } else {
      event.respondWith(fetch(request).catch(() => caches.match("/offline", { ignoreVary: true })));
    }
    return;
  }

  if (ASSET_PREFIX.some((p) => url.pathname.startsWith(p))) {
    event.respondWith(staleWhileRevalidate(request));
  }
});
