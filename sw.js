// 엘리야 대시보드 서비스워커 (v1 · 2026.07.29)
// 목적: PWA "앱 설치"가 가능하도록 서비스워커를 등록만 함.
// 캐싱은 하지 않는다 — 파일을 자주 수정하므로, 캐시하면 옛 버전이 떠서 혼란.
// 따라서 모든 요청은 네트워크 그대로 통과(항상 최신 파일 로드).
self.addEventListener('install', function(event){
  self.skipWaiting(); // 새 SW 즉시 활성화
});
self.addEventListener('activate', function(event){
  event.waitUntil(self.clients.claim());
});
self.addEventListener('fetch', function(event){
  // respondWith를 호출하지 않음 = 브라우저 기본 네트워크 요청(캐시 없음).
  // (fetch 핸들러 존재 자체가 설치 가능 조건을 만족시킴)
});
