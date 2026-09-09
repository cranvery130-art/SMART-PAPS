// 최소한의 서비스워커. 오프라인 캐싱 전략을 복잡하게 가져가는 대신, 설치(installable) 조건을
// 만족시키고 앱 셸(껍데기) 정도만 가볍게 캐시하는 데 목적을 둔다 — 기록 데이터 자체는 항상
// Firestore에서 최신 값을 받아오므로, 여기서 데이터까지 오프라인 캐싱하려 하지 않는다.
const CACHE_NAME = "smart-paps-shell-v1";
const SHELL_URLS = ["/", "/index.html", "/manifest.webmanifest"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_URLS)).catch(() => {})
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

// 네트워크 우선, 실패 시에만 캐시로 대체(항상 최신 코드를 우선 사용).
self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
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
