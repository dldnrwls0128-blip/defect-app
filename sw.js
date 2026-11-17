/* 하자관리 서비스 워커 — 앱 셸 캐시, 오프라인 동작 (하자 데이터는 캐시하지 않음: IndexedDB에만 저장) */
const CACHE = "defect-app-v2";
const PRECACHE = ["./", "./index.html", "./manifest.webmanifest", "./css/app.css",
  "./js/app.js", "./js/db.js", "./js/data.js", "./js/importer.js", "./js/import-worker.js",
  "./js/exporter.js", "./js/backup.js", "./js/photo.js",
  "./icons/icon-192.png", "./icons/icon-512.png",
  "./vendor/jszip/jszip.min.js", "./vendor/xlsx/xlsx.mjs"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(PRECACHE)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

// 캐시 우선 + 백그라운드 갱신(현장 약전계에서도 즉시 열림). 새 버전은 다음 실행 때 반영.
self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const cached = await cache.match(req, { ignoreSearch: true });
    const net = fetch(req).then((res) => {
      if (res && res.ok) cache.put(req, res.clone());
      return res;
    }).catch(() => null);
    if (cached) { event.waitUntil(net); return cached; }
    const res = await net;
    if (res) return res;
    if (req.mode === "navigate") {
      const fb = await cache.match("./index.html");
      if (fb) return fb;
    }
    return new Response("offline", { status: 503 });
  })());
});
