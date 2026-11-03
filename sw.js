/* Service worker cho bản web Văn phòng của Bùi.
   CHỈ chạy khi trang được phục vụ qua https (hoặc localhost).
   Ở chế độ file://, trình duyệt không đăng ký được service worker — app vẫn chạy bình thường, chỉ không cài như PWA và không có bộ nhớ đệm ngoại tuyến. */
var CACHE = 'van-phong-bui-v4';
var CORE = ['./index.html', './style.css', './app.js', './manifest.json', './icon.svg'];
self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(CORE); }).catch(function () {}));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }));
});
self.addEventListener('fetch', function (e) {
  if (e.request.method !== 'GET') return;
  e.respondWith(caches.match(e.request).then(function (hit) {
    return hit || fetch(e.request).catch(function () { return caches.match('./index.html'); });
  }));
});
