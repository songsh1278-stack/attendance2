// 자동 생성 파일 (깃허브페이지_만들기.py). 직접 고치지 마세요.
// 화면은 항상 네트워크에서 새로 받고, 인터넷이 안 될 때만 저장본을 보여 줍니다.
var C = 'ya-offline';
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) { return Promise.all(ks.filter(function (k) { return k !== C; }).map(function (k) { return caches.delete(k); })); }).then(function () { return self.clients.claim(); }).then(function () { return self.clients.matchAll({ type: 'window' }); }).then(function (cs) { cs.forEach(function (c) { try { c.navigate(c.url); } catch (x) {} }); }));
});
self.addEventListener('fetch', function (e) {
  var r = e.request;
  if (r.method !== 'GET' || new URL(r.url).origin !== location.origin) return;
  e.respondWith(fetch(r, { cache: 'no-store' }).then(function (res) {
    if (res.ok) { var cp = res.clone(); caches.open(C).then(function (c) { c.put(r, cp); }); }
    return res;
  }).catch(function () { return caches.match(r).then(function (m) { return m || caches.match('index.html'); }); }));
});
