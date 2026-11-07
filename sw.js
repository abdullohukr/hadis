/* Hadislar: offline help.
   - the hadith data (data/, h/), the search index (search/) and the fonts that were used (fonts/): served from the copy at once and
     refreshed in the background;
   - the pages, scripts and styles: ALWAYS from the network first (the panel reloads itself when a new version is out); the copy
     is used only when there is no network. */
var CACHE = 'hadis-data-v1', SHELL = 'hadis-shell-v1';
/* the shell of the panel, copied at once (else the first visit, made before the worker came, would leave nothing for
   offline); a file that fails is left out */
var SHELL_FILES = ['./', 'index.html', 'style.css', 'app.js', 'core.js', 'search.js', 'text.js', 'i18n.js', 'releases.js',
  'assets/brand/logo.svg?v=1', 'assets/brand/favicon.svg?v=1'];
/* the fonts of the interface, with the ?v= of style.css (the data cache matches the whole address) */
var FONT_FILES = ['fonts/SourceSans3-Uz.woff2?v=1', 'fonts/SourceSerif4-Uz.woff2?v=1', 'fonts/SourceSerif4-Uz-Italic.woff2?v=1', 'fonts/ScheherazadeNew-Regular.woff2?v=4', 'fonts/ScheherazadeNew-Bold.woff2?v=1'];
self.addEventListener('install', function (e) {
  self.skipWaiting();
  e.waitUntil(Promise.all([
    caches.open(SHELL).then(function (c) {
      return Promise.all(SHELL_FILES.map(function (f) { return c.add(new Request(f, { cache: 'no-cache' })).catch(function () {}); }));
    }),
    caches.open(CACHE).then(function (c) {
      return Promise.all(FONT_FILES.map(function (f) { return c.match(f).then(function (hit) { return hit || c.add(f).catch(function () {}); }); }));
    })
  ]));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.filter(function (k) { return k.indexOf('hadis-data-') === 0 && k !== CACHE && k !== SHELL; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
/* the pointers to the current versions (data/current.json, search/current.json, data/latin.json): ALWAYS from the network
   first (a stale one would send the panel to the files of an old version, and the site keeps only the current one); one
   copy each, without the ?t= of the address, for offline. When a pointer comes, the files of the other versions are
   taken out of the caches (h/v<old>/, search/v<old>/, h/v…/lat-<old>/: the cache does not grow with every release). */
var POINTER = /\/(?:data\/current|search\/current|data\/latin)\.json$/;
function prune(path, j) {
  var re = null, keep = '';
  if (/data\/current\.json$/.test(path) && j && j.version) { re = /\/h\/v([\w]+)\//; keep = j.version; }
  else if (/search\/current\.json$/.test(path) && j && (j.version || j.v)) { re = /\/search\/v([\w]+)\//; keep = j.version || j.v; }
  else if (/data\/latin\.json$/.test(path) && j && j.v) { re = /\/lat-([\w]+)\//; keep = j.v; }
  if (!re) return Promise.resolve();
  return caches.keys().then(function (names) {
    return Promise.all(names.filter(function (n) { return n.indexOf('hadis-') === 0; }).map(function (n) {
      return caches.open(n).then(function (c) {
        return c.keys().then(function (ks) {
          return Promise.all(ks.map(function (k) {
            var p = new URL(k.url).pathname, m = re.exec(p);
            if (m && m[1] !== keep || POINTER.test(p) && new URL(k.url).search) return c.delete(k);
            return null;
          }));
        });
      });
    }));
  });
}
self.addEventListener('fetch', function (e) {
  var r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || u.origin !== location.origin) return;
  if (POINTER.test(u.pathname)) {
    var key = u.origin + u.pathname;
    e.respondWith(fetch(r, { cache: 'no-store' }).then(function (res) {
      if (res && res.ok) {
        var cp = res.clone();
        e.waitUntil(caches.open(CACHE).then(function (c) {
          return c.put(key, cp.clone()).then(function () { return cp.json(); }).then(function (j) { return prune(u.pathname, j); });
        }).catch(function () {}));
      }
      return res;
    }).catch(function () { return caches.open(CACHE).then(function (c) { return c.match(key); }).then(function (hit) { return hit || Response.error(); }); }));
    return;
  }
  if (!/\/(data|h|search|fonts)\//.test(u.pathname)) {
    if (/\/(sw\.js|version\.json)$/.test(u.pathname) || !/\.(html|js|css|png|svg|woff2|webmanifest)$|\/$/.test(u.pathname)) return;
    e.respondWith(fetch(r).then(function (res) {
      // the new copy of a file (app.js?v=21) replaces the copies of its older versions (?v=20 …): the shell cache does
      // not grow with every release
      if (res && res.ok) { var cp = res.clone(); caches.open(SHELL).then(function (c) {
        return c.keys(r, { ignoreSearch: true }).then(function (ks) {
          return Promise.all(ks.filter(function (k) { return k.url !== r.url; }).map(function (k) { return c.delete(k); }));
        }).then(function () { return c.put(r, cp); });
      }).catch(function () {}); }
      return res;
    }).catch(function () { return caches.open(SHELL).then(function (c) { return c.match(r, { ignoreSearch: true }); }).then(function (hit) { return hit || Response.error(); }); }));
    return;
  }
  e.respondWith(caches.open(CACHE).then(function (c) {
    return c.match(r).then(function (hit) {
      var net = fetch(r).then(function (res) { if (res && res.ok) c.put(r, res.clone()).catch(function () {}); return res; });
      if (hit) { net.catch(function () {}); return hit; }
      return net;
    });
  }));
});
