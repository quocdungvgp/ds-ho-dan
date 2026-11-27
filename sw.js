/* Service worker — trang mở được khi mất sóng (thực địa).
   index.html: mạng trước, mất mạng thì bản đã lưu. Thư viện jsDelivr (đã ghim phiên bản): lưu 1 lần dùng mãi.
   Dữ liệu hộ + hàng đợi sửa do index.html tự giữ (localStorage / IndexedDB), không đi qua đây.
   Đổi BAN khi muốn bỏ hết bộ nhớ cũ. */
const BAN = 'dsho-v1';
self.addEventListener('install', e => { self.skipWaiting(); e.waitUntil(caches.open(BAN).then(c => c.addAll(['./', './index.html', './manifest.json']).catch(() => {}))); });
self.addEventListener('activate', e => e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== BAN).map(k => caches.delete(k)))).then(() => self.clients.claim())));
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET') return;
  if (u.hostname === 'cdn.jsdelivr.net' || u.hostname === 'fonts.gstatic.com' || u.hostname === 'fonts.googleapis.com') {
    e.respondWith(caches.match(r).then(c => c || fetch(r).then(res => { if (res.ok || res.type === 'opaque') { const cp = res.clone(); caches.open(BAN).then(ca => ca.put(r, cp)); } return res; })));
    return;
  }
  if (u.origin === location.origin) {
    e.respondWith(fetch(r).then(res => { if (res.ok) { const cp = res.clone(); caches.open(BAN).then(ca => ca.put(r, cp)); } return res; })
      .catch(() => caches.match(r, { ignoreSearch: true }).then(c => c || caches.match('./index.html'))));
  }
});
