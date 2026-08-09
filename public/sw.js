// Service worker tối giản — chỉ để trình duyệt coi app là "installable" (PWA) và có
// cache dự phòng cơ bản. KHÔNG cache API Supabase (network-first cho mọi request khác
// app shell) để dữ liệu luôn mới, không bị lỗi hiển thị số liệu cũ.
const CACHE_NAME = "pl-eros-143-shell-v1";
const APP_SHELL = ["/", "/manifest.json", "/icons/icon-192.png", "/icons/icon-512.png"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  // Không can thiệp request tới Supabase hay API bên ngoài — luôn lấy dữ liệu mới nhất.
  if (event.request.method !== "GET" || url.origin !== self.location.origin) return;

  event.respondWith(
    fetch(event.request)
      .then((res) => {
        const resClone = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, resClone)).catch(() => {});
        return res;
      })
      .catch(() => caches.match(event.request))
  );
});
