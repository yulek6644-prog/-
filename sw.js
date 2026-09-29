/* Service worker: приложение открывается и сканирует даже без интернета.
 * При выпуске новой версии увеличьте CACHE. */
var CACHE = 'scan-to-sheet-v2';
var FILES = [
  './',
  'index.html',
  'style.css',
  'app.js',
  'manifest.webmanifest',
  'vendor/barcode-detector.js',
  'vendor/zxing_reader.wasm',
  'icons/icon.svg',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/icon-maskable-512.png'
];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE).then(function (c) { return c.addAll(FILES); }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (event) {
  var req = event.request;
  // Запросы к Google (отправка данных) не трогаем — только свои файлы.
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;

  // Сразу отдаём из кэша (быстро даже при плохой сети), а в фоне обновляем кэш —
  // новая версия приложения подхватится при следующем открытии.
  event.respondWith(
    caches.match(req, { ignoreSearch: true }).then(function (hit) {
      var network = fetch(req).then(function (res) {
        if (res && res.ok) {
          var copy = res.clone();
          caches.open(CACHE).then(function (c) { c.put(req, copy); });
        }
        return res;
      });
      if (hit) {
        event.waitUntil(network.catch(function () {}));
        return hit;
      }
      return network.catch(function () { return caches.match('index.html'); });
    })
  );
});
