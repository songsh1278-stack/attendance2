// 홈 화면 "앱 설치"가 가능하도록 하는 최소한의 서비스 워커. 데이터는 저장하지 않고 항상 인터넷에서 불러온다.
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function (e) { e.respondWith(fetch(e.request)); });
