// 앱 화면(이 사이트의 파일)을 기기에 저장해 두고 바로 띄운다. 뒤에서 새 버전을 받아 두었다가 다음에 열 때 쓴다.
// 출석 데이터(구글 웹앱 요청)는 저장하지 않고 항상 인터넷으로 주고받는다.
var CACHE = 'attendance-v2';
var SHELL = ['./', 'manifest.webmanifest', 'icon-192.png', 'icon-180.png'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  var key = req.mode === 'navigate' ? './' : req;
  e.respondWith(caches.open(CACHE).then(function (c) {
    return c.match(key).then(function (hit) {
      var net = fetch(req, { cache: 'no-cache' }).then(function (res) {
        if (res && res.ok) c.put(key, res.clone());
        return res;
      });
      if (hit) { e.waitUntil(net.catch(function () {})); return hit; }
      return net;
    });
  }));
});
