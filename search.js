/* Search in the hadiths: the client of the static search index (Arabic + Uzbek), the encoder / decoder of its chunks
   and an in-memory builder. The general module of quranuz/web/search.js (format KQS1, docs/search-index-format.md of
   quranuz) with the number of documents as a parameter: a document = a hadith (0..N-1, the order of the books,
   web/data/books.json). No DOM: the panel, the Word add-in, the extension, Node (index builder, tests) and Workers use
   the same code. The text normalisation is HadisText (web/text.js), shared with the builder.

     var s = HadisSearch.Search({ fetch, bases?, version?, cache?, timeout?, text? });
       text(doc, srcId, signal) -> Promise<string | null>: the text of a hadith in a source ('ar' | 'uz', without HTML)
       for the phrase check and the highlighting (the ranges are offsets in that text); without it: no confirm step
     s.run(query, { scope, langs, selected, prefer, limit, local, signal, deadline, onProgress, onResults })
       -> Promise<[{d, score, tier, same?, hits: [{src, kind, ranges?}]}]> (never rejects; the array has .warnings []
          and .partial when a part of the data could not be loaded). tier: 3 the phrase, 2 the words near each other,
          1 every word, 0 not every word (see tiered); same: the phrase in the very forms of the query. A query of 2+
          words: the hadiths without every word go after all the others, and are left out when CUT_MIN have them all
          (opts.cut false: kept, the search of a quote)
       local(q) -> [{d, score, hits}]: results found without the index (a reference «Бухорий 1»), merged first   */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./text.js'));
  else root.HadisSearch = factory(root.HadisText);
})(typeof self !== 'undefined' ? self : this, function (T) {
  'use strict';

  var DEFAULT_LANGS = ['ar', 'uz'];
  // the site itself; the mirrors (jsDelivr, raw GitHub) are given by the caller (o.bases) once the domain is chosen
  var BASES = ['./'];
  var MAGIC = [0x4B, 0x51, 0x53, 0x31];          // "KQS1"
  var MAX_QUERY = 200, MAX_WORDS = 12, MAX_PAIRS = 24, MAX_POSTINGS = 600000;
  var LEN_STEP = 4;                               // len.bin: the number of index terms of a hadith / 4, at most 255

  /* ---------- bytes ---------- */
  function utf8(s) {
    if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(s);
    var out = [];
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      if (c >= 0xD800 && c < 0xDC00 && i + 1 < s.length) { c = 0x10000 + ((c - 0xD800) << 10) + (s.charCodeAt(++i) - 0xDC00); }
      if (c < 0x80) out.push(c);
      else if (c < 0x800) out.push(0xC0 | c >> 6, 0x80 | c & 63);
      else if (c < 0x10000) out.push(0xE0 | c >> 12, 0x80 | c >> 6 & 63, 0x80 | c & 63);
      else out.push(0xF0 | c >> 18, 0x80 | c >> 12 & 63, 0x80 | c >> 6 & 63, 0x80 | c & 63);
    }
    return new Uint8Array(out);
  }
  var decoder = typeof TextDecoder !== 'undefined' ? new TextDecoder('utf-8', { fatal: true }) : null;
  function fromUtf8(b, from, to) {
    if (decoder) return decoder.decode(b.subarray(from, to));
    var s = '';
    for (var i = from; i < to;) {
      var c = b[i++];
      if (c >= 0xF0) c = (c & 7) << 18 | (b[i++] & 63) << 12 | (b[i++] & 63) << 6 | b[i++] & 63;
      else if (c >= 0xE0) c = (c & 15) << 12 | (b[i++] & 63) << 6 | b[i++] & 63;
      else if (c >= 0xC0) c = (c & 31) << 6 | b[i++] & 63;
      if (c >= 0x10000) { c -= 0x10000; s += String.fromCharCode(0xD800 + (c >> 10), 0xDC00 + (c & 1023)); }
      else s += String.fromCharCode(c);
    }
    return s;
  }
  function Writer() { this.b = new Uint8Array(1024); this.n = 0; }
  Writer.prototype.room = function (k) {
    if (this.n + k <= this.b.length) return;
    var nb = new Uint8Array(Math.max(this.b.length * 2, this.n + k)); nb.set(this.b.subarray(0, this.n)); this.b = nb;
  };
  Writer.prototype.u = function (v) {            // unsigned LEB128
    this.room(5);
    while (v >= 0x80) { this.b[this.n++] = v & 0x7F | 0x80; v = Math.floor(v / 128); }
    this.b[this.n++] = v;
  };
  Writer.prototype.bytes = function (a) { this.room(a.length); this.b.set(a, this.n); this.n += a.length; };
  Writer.prototype.done = function () { return this.b.slice(0, this.n); };
  function readU(b, st) {                        // st = {p}
    var v = 0, mul = 1, c;
    do {
      if (st.p >= b.length) throw new Error('chunk: truncated');
      c = b[st.p++]; v += (c & 0x7F) * mul; mul *= 128;
      if (mul > 0x800000000) throw new Error('chunk: bad number');
    } while (c & 0x80);
    return v;
  }

  /* ---------- chunks ----------
     entries: [{term, df, pairs: [doc, src, doc, src, ...] (any order) | null, docs: [doc, ...] | null}]
     A term is written as a bitmap of the documents (no sources) when it has more than `bitmapPairs` pairs.
     opts.docs: the number of documents N (the size of a bitmap: ceil(N / 8) bytes).                          */
  function cmp(a, b) { return a < b ? -1 : a > b ? 1 : 0; }
  function defaultBitmapPairs(n) { return Math.max(64, Math.ceil(n / 8 / 2.5)); }   // where a bitmap gets smaller
  function needDocs(n) { if (!(n > 0)) throw new Error('the number of documents is required'); return n; }
  function encodeChunk(entries, opts) {
    opts = opts || {};
    var docs = needDocs(opts.docs), bitmapPairs = opts.bitmapPairs || defaultBitmapPairs(docs);
    var list = entries.slice().sort(function (x, y) { return cmp(x.term, y.term); });
    var w = new Writer(), prev = new Uint8Array(0);
    w.bytes(MAGIC);
    w.u(list.length);
    for (var i = 0; i < list.length; i++) {
      var e = list[i], tb = utf8(e.term), sh = 0;
      if (i && list[i - 1].term === e.term) throw new Error('chunk: duplicate term ' + e.term);
      while (sh < tb.length && sh < prev.length && tb[sh] === prev[sh]) sh++;
      w.u(sh); w.u(tb.length - sh); w.bytes(tb.subarray(sh));
      prev = tb;
      var body = new Writer(), mode = 0, df;
      if (e.pairs && e.pairs.length / 2 <= bitmapPairs) {
        var codes = [];
        for (var k = 0; k < e.pairs.length; k += 2) codes.push(e.pairs[k] * 65536 + e.pairs[k + 1]);
        codes.sort(function (a, b) { return a - b; });
        var last = 0, lastCode = -1, cnt = 0;
        for (var j = 0; j < codes.length; j++) {
          if (codes[j] === lastCode) continue;
          var dc = Math.floor(codes[j] / 65536), src = codes[j] % 65536;
          if (dc >= docs) throw new Error('chunk: document ' + dc + ' out of range');
          body.u(dc - last); body.u(src); last = dc; lastCode = codes[j]; cnt++;
        }
        df = cnt;
      } else {
        mode = 1;
        var bm = new Uint8Array(Math.ceil(docs / 8)), set = e.docs || [];
        if (e.pairs) { set = []; for (var m = 0; m < e.pairs.length; m += 2) set.push(e.pairs[m]); }
        for (var q = 0; q < set.length; q++) {
          if (set[q] >= docs) throw new Error('chunk: document ' + set[q] + ' out of range');
          bm[set[q] >> 3] |= 1 << (set[q] & 7);
        }
        body.bytes(bm);
        df = e.df || (e.pairs ? e.pairs.length / 2 : set.length);
      }
      var bb = body.done();
      w.u(df); w.u(mode); w.u(bb.length); w.bytes(bb);
    }
    return w.done();
  }

  /* -> {terms[], find(term) -> i | -1, prefix(p, max) -> [i], df(i), mode(i), postings(i) -> Uint32Array
         [doc, src, ...] (mode 0) | docs(i) -> [doc] (mode 1)}; throws on a broken chunk. n: the number of documents */
  function decodeChunk(buf, n) {
    var b = buf instanceof Uint8Array ? buf : new Uint8Array(buf), docs = needDocs(n);
    for (var i = 0; i < 4; i++) if (b[i] !== MAGIC[i]) throw new Error('chunk: not a search chunk');
    var st = { p: 4 }, cnt = readU(b, st);
    if (cnt > b.length) throw new Error('chunk: bad term count');
    var terms = new Array(cnt), dfs = new Array(cnt), modes = new Array(cnt), offs = new Array(cnt), lens = new Array(cnt);
    var prev = new Uint8Array(0);
    for (var k = 0; k < cnt; k++) {
      var sh = readU(b, st), sl = readU(b, st);
      if (sh > prev.length || st.p + sl > b.length) throw new Error('chunk: bad term');
      var tb = new Uint8Array(sh + sl);
      tb.set(prev.subarray(0, sh)); tb.set(b.subarray(st.p, st.p + sl), sh);
      st.p += sl; prev = tb;
      terms[k] = fromUtf8(tb, 0, tb.length);
      dfs[k] = readU(b, st); modes[k] = readU(b, st); lens[k] = readU(b, st); offs[k] = st.p;
      if (modes[k] > 1 || st.p + lens[k] > b.length) throw new Error('chunk: bad postings');
      if (modes[k] === 0 ? dfs[k] * 2 > lens[k] : lens[k] !== Math.ceil(docs / 8)) throw new Error('chunk: bad postings size');
      st.p += lens[k];
    }
    function find(t) {
      var lo = 0, hi = cnt - 1;
      while (lo <= hi) { var m = (lo + hi) >> 1; if (terms[m] < t) lo = m + 1; else if (terms[m] > t) hi = m - 1; else return m; }
      return -1;
    }
    return {
      terms: terms, bytes: b.length,
      find: find,
      prefix: function (p, max) {
        var lo = 0, hi = cnt, out = [];
        while (lo < hi) { var m = (lo + hi) >> 1; if (terms[m] < p) lo = m + 1; else hi = m; }
        for (var j = lo; j < cnt && out.length < (max || 50) && terms[j].slice(0, p.length) === p; j++) out.push(j);
        return out;
      },
      df: function (j) { return dfs[j]; },
      mode: function (j) { return modes[j]; },
      postings: function (j) {
        var s = { p: offs[j] }, end = offs[j] + lens[j], out = new Uint32Array(dfs[j] * 2), c = 0, dc = 0;
        while (s.p < end && c < out.length) { dc += readU(b, s); out[c++] = dc; out[c++] = readU(b, s); }
        if (dc >= docs) throw new Error('chunk: document out of range');
        return c === out.length ? out : out.subarray(0, c);
      },
      docs: function (j) {
        var out = [];
        for (var a = 0; a < docs; a++) if (b[offs[j] + (a >> 3)] & 1 << (a & 7)) out.push(a);
        return out;
      }
    };
  }

  /* the group of a term: by the script of its first letter (manifest.groups[<set>-<script>]) */
  function scriptKey(term) {
    var c = term.charCodeAt(0);
    if (c >= 0xD800 && c < 0xDC00) c = 0x10000 + ((c - 0xD800) << 10) + (term.charCodeAt(1) - 0xDC00);
    return (T.scriptOf(c) || 'Zyyy').toLowerCase();
  }

  /* ---------- the builder (the index of all hadiths is built in memory: tools/build_search_index.js) ----------
     b = new Builder({docs: N}); b.addSource({id, lang, kind?, name, script?}, texts[N] | function (doc) -> text);
     b.finish({groups?, chunkBytes?, bitmapPairs?, version?, meta?}) -> {version, manifest, files: {path: Uint8Array|string}}
     Arabic-language sources (lang 'ar') are indexed by the stem (HadisText.arStem), the others by the whole word.
     len.bin: one byte per document, the number of its index terms (all sources) / LEN_STEP: the length of BM25.  */
  function Builder(o) {
    o = o || {};
    this.docs = needDocs(o.docs);
    this.sources = []; this.post = new Map(); this.len = new Uint32Array(this.docs); this.pairs = 0; this.kept = {};
  }
  Builder.prototype.addSource = function (meta, texts) {
    var n = this.sources.length, stem = meta.lang === 'ar', post = this.post, sample = '';
    var get = typeof texts === 'function' ? texts : function (i) { return texts[i]; };
    if (n >= 65535) throw new Error('too many sources');
    for (var i = 0; i < this.docs; i++) {
      var t = get(i);
      if (!t) continue;
      t = T.stripHtml(t);
      if (sample.length < 2000) sample += ' ' + t;
      var seen = new Set(), list = T.terms(t, { stem: stem }), any = false;
      this.len[i] += list.length;
      if (meta.keep) (this.kept[n] || (this.kept[n] = []))[i] = T.tokenize(t).map(function (x) {   // for addPairs
        return x.pairs ? null : [T.arLight(x.t), T.arStem(x.t)];
      });
      for (var k = 0; k < list.length; k++) {
        var term = list[k];
        if (seen.has(term) || term.length > 60) continue;
        seen.add(term); any = true;
        var arr = post.get(term);
        if (!arr) post.set(term, arr = []);
        arr.push(i * 65536 + n);
      }
      if (any) this.pairs++;
    }
    this.sources.push({ id: meta.id, lang: meta.lang, script: meta.script || T.detectScript(sample) || 'Latn',
                        kind: meta.kind || 'text', stem: stem ? 1 : 0, name: meta.name || meta.id });
    return n;
  };
  /* The word pairs of an Arabic-language source (added with meta.keep): every two neighbouring words whose stems are both
     in at least minDf hadiths (a rarer word finds its few hadiths alone), as one term «lightA lightB» (HadisText.arLight: the article kept, و ف dropped) of a source
     {kind: 'pair', of: <the source>}. A phrase of frequent
     words (حدثنا محمد بن, قال رسول الله) has few hadiths where they stand next to each other: the client looks the pairs
     of the query up and puts those hadiths first. Not a source of its own in the results, not in len.bin / pairs.  */
  Builder.prototype.addPairs = function (meta, from, minDf) {
    var lists = this.kept[from], n = this.sources.length, post = this.post, df = new Map(), i, k;
    if (!lists) throw new Error('addPairs: the source was added without keep');
    if (n >= 65535) throw new Error('too many sources');
    for (i = 0; i < this.docs; i++) {
      if (!lists[i]) continue;
      var once = new Set(lists[i].map(function (x) { return x ? x[1] : ''; }));
      once.forEach(function (t) { df.set(t, (df.get(t) || 0) + 1); });
    }
    function arab(t) { return scriptKey(t) === 'arab'; }
    for (i = 0; i < this.docs; i++) {
      var l = lists[i], seen = new Set();
      if (!l) continue;
      for (k = 0; k + 1 < l.length; k++) {
        var a = l[k], b = l[k + 1];
        if (!a || !b || df.get(a[1]) < minDf || df.get(b[1]) < minDf || !arab(a[0]) || !arab(b[0])) continue;
        var term = a[0] + ' ' + b[0];
        if (seen.has(term) || term.length > 60) continue;
        seen.add(term);
        var arr = post.get(term);
        if (!arr) post.set(term, arr = []);
        arr.push(i * 65536 + n);
      }
    }
    delete this.kept[from];
    this.sources.push({ id: meta.id, lang: meta.lang, script: 'Arab', kind: 'pair', stem: 1, of: this.sources[from].id,
                        min: minDf, name: meta.name || meta.id });
    return n;
  };
  Builder.prototype.finish = function (o) {
    o = o || {};
    var self = this, chunkBytes = o.chunkBytes || 65536, bitmapPairs = o.bitmapPairs || defaultBitmapPairs(this.docs);
    var langs = {};
    this.sources.forEach(function (s) { langs[s.lang] = 1; });
    var sets = o.groups || [{ name: 'core', langs: Object.keys(langs).sort() }];
    var files = {}, groups = {};
    var terms = Array.from(this.post.keys()).sort(cmp);
    sets.forEach(function (g) {
      var allowed = {};
      self.sources.forEach(function (s, n) { if (g.langs.indexOf(s.lang) >= 0) allowed[n] = 1; });
      var byScript = {}, langsOf = {};
      terms.forEach(function (term) {
        var codes = self.post.get(term), pairs = [], key = scriptKey(term), ls = langsOf[key] || (langsOf[key] = {});
        for (var i = 0; i < codes.length; i++) {
          var src = codes[i] % 65536;
          if (allowed[src]) { pairs.push(Math.floor(codes[i] / 65536), src); ls[self.sources[src].lang] = 1; }
        }
        if (!pairs.length) return;
        (byScript[key] || (byScript[key] = [])).push({ term: term, pairs: pairs });
      });
      Object.keys(byScript).sort().forEach(function (sk) {
        var name = g.name + '-' + sk, list = byScript[sk], first = [], sizes = [], cur = [], size = 0;
        function flush() {
          if (!cur.length) return;
          var bin = encodeChunk(cur, { docs: self.docs, bitmapPairs: bitmapPairs });
          files[name + '/' + first.length + '.bin'] = bin;
          first.push(first.length ? cur[0].term : '');
          sizes.push(bin.length);
          cur = []; size = 0;
        }
        list.forEach(function (e) {
          var est = e.pairs.length / 2 > bitmapPairs ? Math.ceil(self.docs / 8) : e.pairs.length * 1.6;
          if (size && size + est > chunkBytes) flush();
          cur.push(e); size += est + e.term.length + 6;
        });
        flush();
        groups[name] = { set: g.name, langs: Object.keys(langsOf[sk]).sort(), script: sk, first: first, bytes: sizes };
      });
    });
    var lenBytes = new Uint8Array(this.docs), sum = 0, cnt = 0;
    for (var i = 0; i < this.docs; i++) {
      lenBytes[i] = Math.min(255, Math.round(this.len[i] / LEN_STEP));
      if (this.len[i]) { sum += lenBytes[i]; cnt++; }
    }
    files['len.bin'] = lenBytes;
    var manifest = { format: 1, kind: 'hadis', docs: this.docs, pairs: this.pairs, defaultLangs: DEFAULT_LANGS,
                     sources: this.sources, groups: groups,
                     len: { file: 'len.bin', step: LEN_STEP, avg: cnt ? Math.round(sum / cnt * 100) / 100 : 1 } };
    if (o.meta) Object.keys(o.meta).forEach(function (k) { if (!(k in manifest)) manifest[k] = o.meta[k]; });
    var version = o.version || hashFiles(files, JSON.stringify(manifest));
    manifest.version = version;
    var out = {};
    Object.keys(files).forEach(function (p) { out['search/v' + version + '/' + p] = files[p]; });
    out['search/v' + version + '/manifest.json'] = JSON.stringify(manifest);
    out['search/current.json'] = JSON.stringify({ version: version });
    return { version: version, manifest: manifest, files: out };
  };
  function hashFiles(files, extra) {                // FNV-1a over the paths and bytes: the version of the index
    var h = 0x811C9DC5;
    function mix(c) { h ^= c; h = Math.imul ? Math.imul(h, 0x01000193) >>> 0 : (h * 0x01000193) >>> 0; }
    Object.keys(files).sort().forEach(function (p) {
      var f = files[p], i;
      for (i = 0; i < p.length; i++) mix(p.charCodeAt(i) & 255);
      if (typeof f === 'string') for (i = 0; i < f.length; i++) mix(f.charCodeAt(i) & 255);
      else for (i = 0; i < f.length; i++) mix(f[i]);
    });
    for (var j = 0; j < extra.length; j++) mix(extra.charCodeAt(j) & 255);
    return ('0000000' + h.toString(16)).slice(-8);
  }

  /* ---------- loading: memory LRU -> Cache API -> mirrors (own site -> jsDelivr -> raw GitHub) ---------- */
  function now() { return Date.now(); }
  function delay(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function tick() { return typeof setTimeout === 'function' ? delay(0) : Promise.resolve(); }
  /* a newer query took over: not an error of the data (no warning, no "partial") */
  function abortError() { var e = new Error('cancelled'); e.name = 'AbortError'; e.cancelled = true; return e; }
  function isAbort(e) { return !!(e && e.cancelled); }
  /* p, or null after ms (a Cache API that never answers must not stop the search); late(): called on the timeout */
  function orNull(p, ms, late) {
    return new Promise(function (resolve) {
      var tm = setTimeout(function () { if (late) late(); resolve(null); }, ms);
      Promise.resolve(p).then(function (v) { clearTimeout(tm); resolve(v); }, function () { clearTimeout(tm); resolve(null); });
    });
  }
  function onAbort(signal, fn) {
    if (!signal || !signal.addEventListener) return function () {};
    signal.addEventListener('abort', fn);
    return function () { signal.removeEventListener('abort', fn); };
  }

  function Loader(o) {
    o = o || {};
    this.fetch = o.fetch || (typeof fetch === 'function' ? function (u, init) { return fetch(u, init); } : null);
    this.bases = (o.bases || (o.base ? [].concat(o.base) : BASES)).map(function (b) { return /\/$/.test(b) ? b : b + '/'; });
    this.timeout = o.timeout || 6000;
    this.retries = o.retries != null ? o.retries : 1;
    this.retryPause = o.retryPause != null ? o.retryPause : 400;
    this.deadMs = o.deadMs || 60000;
    this.max = o.concurrency || 4;
    this.memMax = Math.min(o.memoryBytes || 24e6, 30e6);
    this.cacheName = o.cache === false ? null : (typeof o.cache === 'string' ? o.cache : 'hadis-search-v1');
    this.cacheObj = o.cache && typeof o.cache === 'object' ? o.cache : null;
    this.mem = {}; this.order = []; this.memBytes = 0;
    this.inflight = {}; this.dead = {}; this.active = 0; this.queue = [];
    this.stats = { network: 0, cache: 0, memory: 0, failed: 0, bytes: 0 };
  }
  Loader.prototype.remember = function (key, value, bytes) {
    if (this.mem[key]) return;
    this.mem[key] = { v: value, n: bytes }; this.order.push(key); this.memBytes += bytes;
    while (this.memBytes > this.memMax && this.order.length > 1) {
      var k = this.order.shift(); this.memBytes -= this.mem[k].n; delete this.mem[k];
    }
  };
  Loader.prototype.recall = function (key) {
    var e = this.mem[key];
    if (!e) return undefined;
    var i = this.order.indexOf(key);
    if (i >= 0 && i !== this.order.length - 1) { this.order.splice(i, 1); this.order.push(key); }
    return e.v;
  };
  /* at most `max` requests at a time, in the order asked; want() false when its turn comes (every query that waited
     for it has been replaced by a newer one): it is dropped without a request and the next one goes */
  Loader.prototype.slot = function (fn, want) {
    var self = this;
    return new Promise(function (resolve, reject) {
      self.queue.push(function () {
        if (want && !want()) { reject(abortError()); return; }
        self.active++;
        var p;
        try { p = Promise.resolve(fn()); } catch (e) { p = Promise.reject(e); }
        p.then(function (v) { done(); resolve(v); }, function (e) { done(); reject(e); });
      });
      function done() { self.active--; self.pump(); }
      self.pump();
    });
  };
  Loader.prototype.pump = function () {
    while (this.active < this.max && this.queue.length) this.queue.shift()();
  };
  Loader.prototype.cache = function () {
    if (this.cacheObj) return Promise.resolve(this.cacheObj);
    if (!this.cacheName || typeof caches === 'undefined' || !caches.open) return Promise.resolve(null);
    var self = this;
    if (!this._cache) this._cache = orNull(caches.open(this.cacheName), 3000).then(function (c) { if (!c) self.cacheName = null; return c; });
    return this._cache;
  };
  Loader.prototype.noCache = function () { this.cacheObj = null; this.cacheName = null; };   // the Cache API hangs: go without it
  Loader.prototype.fetchOne = function (url) {
    var self = this, ctl = typeof AbortController !== 'undefined' ? new AbortController() : null, timer;
    return new Promise(function (resolve, reject) {
      timer = setTimeout(function () { if (ctl) try { ctl.abort(); } catch (e) { /* ignore */ } reject(new Error('timeout')); }, self.timeout);
      var p;
      try { p = self.fetch(url, ctl ? { signal: ctl.signal } : undefined); } catch (e) { p = Promise.reject(e); }
      Promise.resolve(p).then(function (r) {
        if (!r || !r.ok) { var er = new Error('HTTP ' + (r ? r.status : '?')); er.status = r ? r.status : 0; throw er; }
        return r.arrayBuffer().then(function (ab) { return { r: r, ab: ab }; });
      }).then(function (v) { clearTimeout(timer); resolve(v); }, function (e) { clearTimeout(timer); reject(e); });
    });
  };
  /* path relative to the bases; type 'bin' | 'json'; immutable: the answer may be kept in the Cache API forever.
     One load per file, shared by everybody who asks for it. signal (AbortSignal, optional): the asker gives up — its
     promise rejects at once (an AbortError, err.cancelled); the shared load goes on while somebody else waits for it
     or its request is already running (the answer is kept for later), a queued one nobody wants any more is dropped. */
  Loader.prototype.get = function (path, type, immutable, signal) {
    var self = this, key = type + ':' + path, m = this.recall(key);
    if (signal && signal.aborted) return Promise.reject(abortError());
    if (m !== undefined) { this.stats.memory++; return Promise.resolve(m); }
    var e = this.inflight[key] || this.load(key, path, type, immutable);
    if (!signal) { e.strong++; return e.p; }
    e.users++;
    return new Promise(function (resolve, reject) {
      var over = false, off = onAbort(signal, function () { end(); reject(abortError()); });
      function end() { if (over) return false; over = true; e.users--; off(); return true; }
      e.p.then(function (v) { if (end()) resolve(v); }, function (err) {
        if (!end()) return;
        // the shared load was dropped just before this asker joined it: start it again
        if (isAbort(err) && !signal.aborted) self.get(path, type, immutable, signal).then(resolve, reject);
        else reject(err);
      });
    });
  };
  Loader.prototype.load = function (key, path, type, immutable) {
    var self = this, e = { users: 0, strong: 0 };
    function want() {                               // somebody still waits for it (else: forget it at once)
      if (e.strong > 0 || e.users > 0) return true;
      if (self.inflight[key] === e) delete self.inflight[key];
      return false;
    }
    var cacheKey = 'https://hadis-search.invalid/' + path;
    function parse(ab) {
      if (type === 'bin') return { v: new Uint8Array(ab), n: ab.byteLength * 3 };
      var text = fromUtf8(new Uint8Array(ab), 0, ab.byteLength);
      return { v: JSON.parse(text), n: text.length * 2 + 64 };
    }
    var p = (immutable ? this.cache().then(function (c) {
      return c ? orNull(c.match(cacheKey).then(function (r) { return r ? r.arrayBuffer() : null; }), 3000, function () { self.noCache(); }) : null;
    }) : Promise.resolve(null)).then(function (ab) {
      if (ab) {
        try { var got = parse(ab); self.stats.cache++; return got; }
        catch (er) { self.cache().then(function (c) { if (c && c['delete']) c['delete'](cacheKey); }); }
      }
      if (!want()) throw abortError();
      return self.network(path, want).then(function (res) {
        var got = parse(res.ab);                     // throws on broken data: the caller skips this part
        self.stats.network++; self.stats.bytes += res.ab.byteLength;
        if (immutable) self.cache().then(function (c) {
          if (c && typeof Response !== 'undefined') return c.put(cacheKey, new Response(res.ab));
        }).catch(function () { /* full storage: no cache */ });
        return got;
      });
    }).then(function (got) {
      if (self.inflight[key] === e) delete self.inflight[key];
      self.remember(key, got.v, got.n);
      return got.v;
    }, function (er) {
      if (self.inflight[key] === e) delete self.inflight[key];
      if (!isAbort(er)) self.stats.failed++;
      throw er;
    });
    p.then(null, function () { /* the askers get the error; nobody left: not an unhandled rejection */ });
    e.p = p;
    this.inflight[key] = e;
    return e;
  };
  /* the mirrors in turn (404: the next one; down / slow: the next one, and that mirror is left alone for a while);
     want(): see slot() */
  Loader.prototype.network = function (path, want) {
    var self = this;
    if (!this.fetch) return Promise.reject(new Error('no fetch'));
    function round(attempt) {
      var t = now(), live = self.bases.filter(function (b) { return !(self.dead[b] > t); });
      if (!live.length) live = self.bases.slice(0, 1);          // all marked down: try the first one once
      var all404 = true, lastErr = null;
      function next(k) {
        if (k >= live.length) {
          if (!all404 && attempt < self.retries) return delay(self.retryPause * (attempt + 1)).then(function () { return round(attempt + 1); });
          var e = lastErr || new Error('not loaded'); e.path = path; throw e;
        }
        var base = live[k];
        return self.slot(function () {
          // a mirror found down while this request waited in the queue: straight to the next one
          while (k + 1 < live.length && self.dead[live[k]] > now()) k++;
          base = live[k];
          return self.fetchOne(base + path).then(null, function (e) {
            // down / slow: leave it alone for a while (marked before the slot is free: the next one in the queue skips it)
            if (e.status !== 404) self.dead[base] = now() + self.deadMs;
            throw e;
          });
        }, want).then(null, function (e) {
          if (isAbort(e)) throw e;                             // nobody waits for it any more: not the mirror's fault
          lastErr = e;
          if (e.status !== 404) all404 = false;                // 404: the file is not there (yet)
          return next(k + 1);                                  // either way: the next mirror
        });
      }
      return next(0);
    }
    return round(0);
  };


  /* ---------- the query ---------- */
  function cleanQuery(q) {
    q = String(q == null ? '' : q).replace(/[\u0000-\u001F\u007F]/g, ' ').replace(/\s+/g, ' ').trim();
    return q.length > MAX_QUERY ? q.slice(0, MAX_QUERY) : q;
  }
  function isCyr(t) { return T.scriptOf(t.charCodeAt(0)) === 'Cyrl'; }
  /* the base of the Uzbek forms of a word (HadisText.uzBase), null when the word is not Uzbek Cyrillic / too short */
  function baseOf(t, partial) { if (!isCyr(t)) return null; var b = T.uzBase(t, partial); return b.length >= 3 ? b : null; }
  /* query words: [{t, stem, prefix, pair, alts}]; lat: the words come from a Latin query (alts: sh = ш | сҳ ...) */
  function queryWords(q, prefixLast, lat) {
    var words = T.queryTerms(q).slice(0, MAX_WORDS), pairs = 0, out = [];
    words.forEach(function (w, i) {
      if (w.pairs) {
        w.pairs.forEach(function (p) { if (pairs++ < MAX_PAIRS) out.push({ t: p, stem: p, stems: [p], lights: [p], prefix: false, pair: true, alts: [p] }); });
      } else out.push({ t: w.t, stem: w.stem, stems: w.stems || [w.stem], lights: w.lights || [w.t], prefix: prefixLast && i === words.length - 1 && w.t.length >= 3,
                        pair: false, alts: lat && isCyr(w.t) ? T.latAlts(w.t) : [w.t] });
    });
    return out;
  }

  /* a query word -> test(token) -> 0 (no), 1 (the word: the same form, the same Arabic stem), 2 (another form: an Uzbek
     suffix, the rest of the last word being typed). stem: the text is Arabic-language; lat: see queryWords */
  function wordTest(w, last, stem, lat) {
    var pre = last && w.t.length >= 3, alts = lat && isCyr(w.t) ? T.latAlts(w.t) : [w.t];
    var bases = alts.map(function (a) { return baseOf(a, pre); }), stems = w.stems || [w.stem];
    return function (t) {
      var i;
      for (i = 0; i < alts.length; i++) if (t === alts[i]) return 1;
      if (T.scriptOf(t.charCodeAt(0)) === 'Arab') return stems.indexOf(T.arStem(t)) >= 0 && (stem || w.stem.length >= 3) ? 1 : 0;
      for (i = 0; i < alts.length; i++) {
        if (pre && t.slice(0, alts[i].length) === alts[i]) return 2;
        if (bases[i] && (t === bases[i] || T.uzInflected(bases[i], t))) return 2;
      }
      return 0;
    };
  }

  /* text normalised for a phrase check / highlighting: stem: the Arabic-language variant */
  function normFor(s, stem) {
    return T.tokenize(s).map(function (t) { return t.pairs ? t.t : stem && T.scriptOf(t.t.charCodeAt(0)) === 'Arab' ? T.arStem(t.t) : t.t; }).join(' ');
  }
  function phraseIn(text, q, stem) {
    var nq = normFor(q, stem);
    if (!nq) return false;
    var nt = ' ' + normFor(text, stem) + ' ';
    return nt.indexOf(' ' + nq) >= 0 || /[^\s]/.test(nq) && T.tokenize(q).some(function (t) { return t.pairs; }) && nt.indexOf(nq) >= 0;
  }

  /* how close the query words are in text: 3 the phrase (the words one after another, in order; other forms of the
     words count), 2 near (all of them within a window of the number of words + 3), 1 all of them, 0 not all */
  var NEAR = 3;
  function closeness(text, q, stem, lat) {
    var qw = T.queryTerms(q).slice(0, MAX_WORDS);
    if (!qw.length) return 0;
    if (qw.some(function (w) { return w.pairs; })) return phraseIn(text, q, stem) ? 3 : 0;
    var toks = T.tokenize(text), nW = qw.length;
    var tests = qw.map(function (w, i) { return wordTest(w, i === nW - 1, stem, lat); });
    var hits = [], pos = [];                          // hits[k]: the mask of the words token k is
    for (var k = 0; k < toks.length; k++) {
      var mk = 0;
      if (!toks[k].pairs) for (var j = 0; j < nW; j++) if (tests[j](toks[k].t)) mk |= 1 << j;
      hits.push(mk);
      if (mk) pos.push(k);
    }
    var all = (1 << nW) - 1, seen = 0;
    pos.forEach(function (p) { seen |= hits[p]; });
    if (seen !== all) return 0;
    if (nW === 1) return 3;
    for (var a = 0; a < pos.length; a++) {
      var s = pos[a], ok = true;
      for (var b = 0; b < nW && ok; b++) ok = s + b < hits.length && !!(hits[s + b] & 1 << b);
      if (ok) return 3;
    }
    // the shortest window that holds every word (two pointers over the token positions of the words)
    var cnt = new Array(nW), have = 0, lo = 0, best = Infinity;
    for (var c = 0; c < nW; c++) cnt[c] = 0;
    function add(p, d) {
      for (var j2 = 0; j2 < nW; j2++) if (hits[p] & 1 << j2) {
        if (d > 0 && cnt[j2]++ === 0) have++;
        if (d < 0 && --cnt[j2] === 0) have--;
      }
    }
    for (var hi = 0; hi < pos.length; hi++) {
      add(pos[hi], 1);
      while (have === nW) { best = Math.min(best, pos[hi] - pos[lo] + 1); add(pos[lo++], -1); }
    }
    return best <= nW + NEAR ? 2 : 1;
  }

  /* [[start, end]] of the query words in text (original offsets), for any script; lat: see queryWords */
  function highlight(text, q, script, lat) {
    text = String(text || '');
    var toks = T.tokenize(text, true), qw = T.queryTerms(cleanQuery(q)), ranges = [];
    if (!qw.length) return [];
    var stem = script === 'Arab' || script === 'ar';
    qw.forEach(function (w, wi) {
      var test = w.pairs ? null : wordTest(w, wi === qw.length - 1, stem, lat);
      toks.forEach(function (tk) {
        if (w.pairs || tk.pairs) {
          var from = 0, p;
          while ((p = tk.t.indexOf(w.t, from)) >= 0) {
            ranges.push([tk.a[p], tk.b[p + w.t.length - 1]]);
            from = p + Math.max(1, w.t.length);
          }
          return;
        }
        if (test(tk.t)) ranges.push([tk.a[0], tk.b[tk.t.length - 1]]);
      });
    });
    ranges.sort(function (x, y) { return x[0] - y[0] || x[1] - y[1]; });
    var out = [];
    ranges.forEach(function (r) {
      var l = out[out.length - 1];
      if (l && r[0] <= l[1]) l[1] = Math.max(l[1], r[1]); else out.push([r[0], r[1]]);
    });
    return out;
  }

  /* groups to read for the terms of one script: the cheapest that cover the wanted languages */
  function pickGroups(manifest, script, langs) {
    var want = {}, cands = [];
    langs.forEach(function (l) { want[l] = 1; });
    Object.keys(manifest.groups || {}).forEach(function (name) {
      var g = manifest.groups[name];
      if (g.script !== script) return;
      var hit = g.langs.filter(function (l) { return want[l]; }).length;
      if (!hit) return;
      var avg = g.bytes && g.bytes.length ? g.bytes.reduce(function (a, b) { return a + b; }, 0) / g.bytes.length : 32768;
      cands.push({ name: name, g: g, cost: 1 + avg / 65536 + 0.5 * (g.langs.length - hit) / g.langs.length });
    });
    var exist = {};
    cands.forEach(function (c) { c.g.langs.forEach(function (l) { if (want[l]) exist[l] = 1; }); });
    var left = Object.keys(exist).length, out = [], covered = {};
    while (left > 0) {
      var best = null, bestV = 0;
      cands.forEach(function (c) {
        var gain = c.g.langs.filter(function (l) { return exist[l] && !covered[l]; }).length;
        if (gain && gain / c.cost > bestV) { bestV = gain / c.cost; best = c; }
      });
      if (!best) break;
      best.g.langs.forEach(function (l) { if (exist[l] && !covered[l]) { covered[l] = 1; left--; } });
      out.push(best.name);
      cands.splice(cands.indexOf(best), 1);
    }
    return out;
  }
  function chunkOf(group, term) {                   // last chunk whose first term <= term
    var f = group.first, lo = 0, hi = f.length - 1;
    while (lo < hi) { var m = (lo + hi + 1) >> 1; if (f[m] <= term) lo = m; else hi = m - 1; }
    return lo;
  }

  /* ---------- the client ---------- */
  function Search(o) {
    if (!(this instanceof Search)) return new Search(o);
    o = o || {};
    this.o = o;
    this.loader = new Loader(o);
    this.version = o.version || null;
    this.root = o.root || 'search/';
    this.text = typeof o.text === 'function' ? o.text : null;
  }
  Search.prototype.manifest = function () {
    var self = this;
    if (this._manifest) return Promise.resolve(this._manifest);
    var v = this.version ? Promise.resolve(this.version) : this.loader.get(this.root + 'current.json', 'json', false).then(function (c) {
      if (!c || !c.version || !/^[\w.-]{1,64}$/.test(c.version)) throw new Error('bad current.json');
      return c.version;
    });
    return v.then(function (ver) {
      return self.loader.get(self.root + 'v' + ver + '/manifest.json', 'json', true).then(function (m) {
        if (!m || m.format !== 1 || !m.sources || !m.groups || !(m.docs > 0)) throw new Error('bad manifest');
        m.version = ver;
        // the lengths of the hadiths (BM25): without them every hadith counts as of average length
        var len = m.len && m.len.file ? self.loader.get(self.root + 'v' + ver + '/' + m.len.file, 'bin', true).then(function (b) {
          if (b && b.length === m.docs) m._len = b; else m._lenWarn = 'len: bad size';
        }, function (e) { m._lenWarn = 'len: ' + (e && e.message || e); }) : Promise.resolve();
        return len.then(function () { self._manifest = m; self.version = ver; return m; });
      });
    });
  };
  Search.prototype.chunk = function (m, group, n, signal) {
    var self = this, key = 'dec:' + group + '/' + n, have = this.loader.recall(key);
    if (have) return Promise.resolve(have);
    return this.loader.get(this.root + 'v' + m.version + '/' + group + '/' + n + '.bin', 'bin', true, signal).then(function (u8) {
      var d = decodeChunk(u8, m.docs);
      self.loader.remember(key, d, 64 + d.terms.length * 48);
      return d;
    });
  };

  /* which sources take part */
  function chooseSources(m, opts) {
    var scope = opts.scope || 'default', langs = {}, ids = {};
    (opts.selected || []).forEach(function (id) { ids[id] = 1; });
    if (scope === 'default') {
      (opts.langs || m.defaultLangs || DEFAULT_LANGS).forEach(function (l) { langs[l] = 1; });
      m.sources.forEach(function (s) { if (ids[s.id]) langs[s.lang] = 1; });
    }
    var out = [];
    m.sources.forEach(function (s, n) {
      if (s.kind === 'pair') return;                  // the word pairs of a source: looked up with it (_index)
      if (scope === 'selected' ? ids[s.id] : scope === 'all' ? true : langs[s.lang] || ids[s.id]) out.push(n);
    });
    return out;
  }

  /* opts.signal (AbortSignal): a newer query replaces this one: from then on no onResults / onProgress, its loads nobody
     else waits for are dropped, the promise resolves at once with [] (.cancelled). opts.deadline (ms, 15000): what has
     not come by then is left out (a warning, .partial): the search always ends. */
  var K1 = 1.2, B = 0.6;
  Search.prototype.run = function (query, opts) {
    var self = this;
    opts = opts || {};
    var warnings = [], limit = Math.max(1, Math.min(opts.limit || 50, 500));
    var q = cleanQuery(query), local = [], user = opts.signal || null;
    function cancelled() { return !!(user && user.aborted); }
    function progress(p) { if (opts.onProgress && !cancelled()) try { opts.onProgress(p); } catch (e) { /* the caller's problem */ } }
    function finish(list, partial) {
      list.warnings = warnings; list.partial = !!partial;
      return list;
    }
    function emit(list, stage) { if (opts.onResults && !cancelled()) try { opts.onResults(list, { stage: stage }); } catch (e) { /* the caller's problem */ } }
    if (!q) return Promise.resolve(finish([], false));
    // 1. at once: what the caller finds without the index (a reference: book + number)
    if (typeof opts.local === 'function') {
      try { local = (opts.local(q) || []).map(function (r) { return { d: r.d, score: r.score, hits: (r.hits || []).slice() }; }); }
      catch (e) { warnings.push('local: ' + e.message); }
      emit(local.slice(0, limit), 'local');
    }
    progress({ stage: 'local', done: 0, total: 1 });
    if (opts.scope === 'local') return Promise.resolve(finish(local.slice(0, limit), false));
    if (cancelled()) return Promise.resolve(gone());

    // the loads of this run: given up when the caller cancels or when the run is over (nobody needs them then)
    var ctl = typeof AbortController !== 'undefined' ? new AbortController() : null, signal = ctl ? ctl.signal : null;
    var t0 = now(), deadline = opts.deadline > 0 ? opts.deadline : 15000, settle = null;
    var stop = new Promise(function (resolve) { settle = resolve; });   // resolves on cancel: the run ends at once
    var unhook = onAbort(user, function () { settle(); if (ctl) ctl.abort(); });
    function gone() { var l = finish([], false); l.cancelled = true; return l; }
    function check(v) { if (cancelled()) throw abortError(); return v; }
    /* p, or an error 'slow' when the deadline of the whole run has passed */
    function within(p) {
      var left = deadline - (now() - t0);
      return new Promise(function (resolve, reject) {
        var tm = setTimeout(function () { reject(new Error('slow')); }, Math.max(0, left));
        p.then(function (v) { clearTimeout(tm); resolve(v); }, function (e) { clearTimeout(tm); reject(e); });
      });
    }
    var ctx = { signal: signal, within: within, cancelled: cancelled };

    var m, prefixLast = !/\s$/.test(String(query == null ? '' : query));
    var work = within(this.manifest()).then(function (man) {
      check(); m = man;
      if (m._lenWarn) warnings.push(m._lenWarn);
      return self._index(q, m, opts, warnings, progress, prefixLast, ctx);
    }).then(function (found) {
      check();
      var list = merge(local, found, limit);
      emit(clean(cut(q, list, opts).slice(0, limit), true), 'index');
      return self._confirm(q, m, list, opts, warnings, ctx);
    }).then(function (list) {
      check();
      list = clean(cut(q, list, opts).slice(0, limit), false);
      emit(list, 'final');
      progress({ stage: 'done', done: 1, total: 1 });
      return finish(list, warnings.length > 0);
    }, function (e) {
      if (isAbort(e) || cancelled()) return gone();
      warnings.push('index: ' + (e && e.message || e));
      var list = local.slice(0, limit);
      emit(list, 'final');
      return finish(list, true);
    }).then(null, function (e) {                       // never rejects
      if (cancelled()) return gone();
      warnings.push('search: ' + (e && e.message || e));
      return finish(local.slice(0, limit), true);
    });
    return Promise.race([work, stop.then(gone)]).then(function (list) {
      unhook(); if (ctl) ctl.abort();                  // over: the loads only this run waited for are dropped
      return list;
    });
  };

  /* A query of 2+ words (the owner, 2026-11-02: «one word of the query found in another hadith comes first»): the
     hadiths with every word (tier >= 1, the references of local: no tier) come first anyway (tiered); once CUT_MIN of
     them are there, the hadiths with a part of the words only (tier 0) are left out */
  var CUT_MIN = 3;
  function cut(q, list, opts) {
    if (opts.cut === false || T.queryTerms(q).length < 2) return list;
    var full = 0;
    list.forEach(function (r) { if (!(r.tier === 0)) full++; });
    return full >= CUT_MIN ? list.filter(function (r) { return r.tier !== 0; }) : list;
  }

  function merge(local, found, limit) {
    var by = {}, out = [];
    local.concat(found).forEach(function (r) {
      var e = by[r.d];
      if (!e) { by[r.d] = e = { d: r.d, score: r.score, tier: r.tier, hits: r.hits.slice() }; out.push(e); return; }
      e.score = Math.max(e.score, r.score) + 0.5;      // found both without the index and in it
      e.tier = e.tier == null || r.tier == null ? null : Math.max(e.tier, r.tier);
      r.hits.forEach(function (h) { if (!e.hits.some(function (x) { return x.src === h.src; })) e.hits.push(h); });
    });
    out.sort(function (x, y) { return y.score - x.score || x.d - y.d; });
    return out.slice(0, Math.max(limit, 50));
  }

  /* copies without the internal fields (copy: for the intermediate results, the final ones are cleaned in place) */
  function clean(list, copy) {
    return list.map(function (r) {
      var o = copy ? { d: r.d, score: r.score, hits: [] } : r;
      if (copy) { if (r.tier != null) o.tier = r.tier; if (r.same) o.same = true; }
      o.hits = (r.hits || []).map(function (h) {
        var x = {};
        for (var k in h) if (k.charAt(0) !== '_') x[k] = h[k];
        return x;
      });
      return o;
    });
  }

  Search.prototype._index = function (q, m, opts, warnings, progress, prefixLast, ctx) {
    var self = this, srcs = chooseSources(m, opts);
    ctx = ctx || { signal: null, within: function (p) { return p; }, cancelled: function () { return false; } };
    if (!srcs.length) return Promise.resolve([]);
    var allowed = {}, needStem = false, needFull = false;
    srcs.forEach(function (n) { var st = !!m.sources[n].stem; allowed[n] = st ? 2 : 1; if (st) needStem = true; else needFull = true; });
    var langs = {};
    srcs.forEach(function (n) { langs[m.sources[n].lang] = 1; });
    langs = Object.keys(langs);
    // the groups of the word pairs of the Arabic-language sources taking part
    var pairLangs = {};
    m.sources.forEach(function (s) {
      if (s.kind === 'pair' && srcs.some(function (n) { return m.sources[n].id === s.of; })) pairLangs[s.lang] = 1;
    });
    var pairGroups = Object.keys(pairLangs).length ? pickGroups(m, 'arab', Object.keys(pairLangs)) : [];
    var results = {};
    function runVariant(v) {
      var jobs = {}, words = queryWords(v.q, prefixLast && opts.prefix !== false, v.lat);
      words.forEach(function (w) {
        var forms = [];                                // {t, full, stem}: full: for whole-word sources, stem: Arabic-language
        if (needFull) w.alts.forEach(function (a) { forms.push({ t: a, full: true, stem: a === w.stem }); });
        if (needStem) w.stems.forEach(function (st) {
          if (!forms.some(function (f) { return f.t === st; })) forms.push({ t: st, full: false, stem: true });
        });
        w.at = [];
        w.exact = {};
        forms.forEach(function (f) {
          w.exact[f.t] = 1;
          // what to look up: the form (+ its longer forms: the last word being typed), the Uzbek forms of its base
          var looks = [{ t: f.t, how: w.prefix ? 1 : 0 }], base = f.full && !w.pair ? baseOf(f.t, w.prefix) : null;
          if (base) looks.push({ t: base, how: 2 });
          looks.forEach(function (lk) {
            pickGroups(m, scriptKey(lk.t), langs).forEach(function (g) {
              var gr = m.groups[g], n = chunkOf(gr, lk.t), key = g + '/' + n;
              jobs[key] = jobs[key] || { g: g, n: n };
              w.at.push({ t: lk.t, how: lk.how, key: key, full: f.full, stem: f.stem });
              // the longer forms may begin in the next chunks (at most 2 more)
              for (var nx = n + 1; lk.how && nx <= n + 2 && nx < gr.first.length && gr.first[nx].slice(0, lk.t.length) === lk.t; nx++) {
                key = g + '/' + nx;
                jobs[key] = jobs[key] || { g: g, n: nx };
                w.at.push({ t: lk.t, how: lk.how, key: key, full: f.full, stem: f.stem });
              }
            });
          });
        });
      });
      // the neighbouring Arabic words of the query as the pairs of the index (a phrase of frequent words)
      var pairs = [];
      if (pairGroups.length) for (var i = 0; i + 1 < words.length && pairs.length < MAX_WORDS; i++) {
        var a = words[i], b = words[i + 1];
        if (a.pair || b.pair || scriptKey(a.stem) !== 'arab' || scriptKey(b.stem) !== 'arab') continue;
        var pl = { at: [] };
        a.lights.forEach(function (sa) {
          b.lights.forEach(function (sb) {
            var t = sa + ' ' + sb, how = b.prefix ? 1 : 0;
            pairGroups.forEach(function (g) {
              var gr = m.groups[g], n = chunkOf(gr, t), key = g + '/' + n;
              jobs[key] = jobs[key] || { g: g, n: n };
              pl.at.push({ t: t, how: how, key: key });
              for (var nx = n + 1; how && nx <= n + 1 && nx < gr.first.length && gr.first[nx].slice(0, t.length) === t; nx++) {
                key = g + '/' + nx;
                jobs[key] = jobs[key] || { g: g, n: nx };
                pl.at.push({ t: t, how: how, key: key });
              }
            });
          });
        });
        pairs.push(pl);
      }
      var keys = Object.keys(jobs), chunks = {}, done = 0;
      progress({ stage: 'chunks', done: 0, total: keys.length });
      return Promise.all(keys.map(function (k) {
        return ctx.within(self.chunk(m, jobs[k].g, jobs[k].n, ctx.signal)).then(function (d) { chunks[k] = d; }, function (e) {
          if (!isAbort(e)) warnings.push('chunk ' + k + ': ' + (e && e.message || e));
        }).then(function () { progress({ stage: 'chunks', done: ++done, total: keys.length }); });
      })).then(tick).then(function () {
        if (ctx.cancelled()) throw abortError();
        scoreVariant({ v: v, words: words, pairs: pairs }, chunks, m, allowed, opts, results);
      });
    }
    function good() { var n = 0; for (var k in results) n++; return n; }
    // the query as typed first; the other alphabet (Latin <-> Cyrillic) when that found little, and always (with the
    // same weight) when no source is written in the alphabet of the query (Latin Uzbek against the Cyrillic text)
    var sc = T.detectScript(q), other = null, has = {};
    srcs.forEach(function (n) { has[m.sources[n].script] = 1; });
    if (sc === 'Latn' && /[a-z]/i.test(q)) other = T.latToCyr(q);
    if (sc === 'Cyrl') other = T.cyrToLat(q);
    var always = other && !has[sc], step = runVariant({ q: q, w: 1 });
    if (other && T.norm(other) !== T.norm(q)) {
      step = step.then(function () { if (always || good() < 3) return runVariant({ q: other, w: always ? 1 : 0.8, lat: sc === 'Latn' }); });
    }
    return step.then(function () {
      var list = Object.keys(results).map(function (k) { return results[k]; });
      list.sort(function (x, y) { return y.score - x.score || x.i - y.i; });
      // the hits of the hadiths found by frequent words only: every source taking part, those of the query's script first
      var qs = T.detectScript(q), every = srcs.slice().sort(function (x, y) {
        return (m.sources[y].script === qs) - (m.sources[x].script === qs) || x - y;
      });
      return list.slice(0, 200).map(function (r) {
        var hits = r.hits || every.map(function (n) { return { src: m.sources[n].id, kind: m.sources[n].kind, _n: n }; });
        return { d: r.i, score: Math.round(r.score * 100) / 100, tier: tierOf(r.score), hits: hits };
      });
    });
  };

  /* The score of a result = TIER * tier + the weight squeezed into [0, TIER): a higher tier always comes first.
     Tiers: 0 not every word, 1 every word (the index), 2 every word near each other, 3 the phrase (both by _confirm). */
  var TIER = 20, PAIR_W = 0.5;
  function tiered(tier, v) { return tier * TIER + TIER * v / (v + 10); }
  function tierOf(score) { return Math.floor(score / TIER); }

  /* BM25 without term frequencies (a term is in a hadith of a source or not; the length: len.bin). Every word in a
     (hadith, source) comes first (tier 1); then the hadiths with the words that carry half of the weight (idf) or more.
     A word found in another form (an Uzbek suffix, the rest of the last word) weighs 0.7.
     Memory stays small: only the postings of the rarest words make candidates, the frequent ones only add to them. */
  function scoreVariant(p, chunks, m, allowed, opts, results) {
    var words = p.words.slice(0, 30), nW = words.length;
    if (!nW) return;
    var N = m.pairs || m.docs * m.sources.length, total = 0, budget = MAX_POSTINGS, idf = [], lists = [];
    var LEN = m._len || null, avg = m.len && m.len.avg || 1;
    words.forEach(function (w, j) {
      var df = 0, ls = [], seen = {};
      w.at.forEach(function (at) {
        var d = chunks[at.key];
        if (!d) return;
        var ids = at.how === 1 ? d.prefix(at.t, 40) : at.how === 2 ? d.prefix(at.t, 4000).filter(function (id) {
          return T.uzInflected(at.t, d.terms[id]);
        }) : [];
        var exact = d.find(at.t);
        if (exact >= 0 && ids.indexOf(exact) < 0) ids.unshift(exact);
        ids.forEach(function (id) {
          var key = at.key + '|' + id + '|' + at.full + at.stem;
          if (seen[key]) return;
          seen[key] = 1;
          df += d.df(id);
          ls.push({ d: d, id: id, exact: !!w.exact[d.terms[id]], full: at.full, stem: at.stem });
        });
      });
      idf[j] = Math.log(1 + (N - df + 0.5) / (df + 0.5));
      total += idf[j];
      lists[j] = { ls: ls, df: df };
    });
    if (!total) return;
    // the words that make candidates: the rarest ones, until the rest weigh less than half
    var order = words.map(function (w, j) { return j; }).sort(function (x, y) { return lists[x].df - lists[y].df; });
    var makes = {}, rest = total;
    for (var o = 0; o < order.length; o++) {
      if (rest < total * 0.5) break;
      makes[order[o]] = 1; rest -= idf[order[o]];
    }
    // the masks of the words found (any form) and of the words found as typed, per (hadith, source) / per hadith
    var cand = new Map(), ex = new Map(), anyMask = new Map(), anyEx = new Map();
    order.forEach(function (j) {
      var bit = 1 << j;
      lists[j].ls.forEach(function (l) {
        if (l.d.mode(l.id) === 1) {
          l.d.docs(l.id).forEach(function (dc) {
            anyMask.set(dc, (anyMask.get(dc) | 0) | bit);
            if (l.exact) anyEx.set(dc, (anyEx.get(dc) | 0) | bit);
          });
          return;
        }
        if (budget <= 0) return;
        var ps = l.d.postings(l.id);
        budget -= ps.length / 2;
        for (var k = 0; k < ps.length; k += 2) {
          var src = ps[k + 1], kind = allowed[src];
          if (!kind || (kind === 2 ? !l.stem : !l.full)) continue;
          var code = ps[k] * 65536 + src, v = cand.get(code);
          if (v === undefined && !makes[j]) continue;
          cand.set(code, (v | 0) | bit);
          if (l.exact) ex.set(code, (ex.get(code) | 0) | bit);
        }
      });
    });
    // the hadiths where the query words stand next to each other (the pairs of the index): + PAIR_W * idf of a pair
    var near = new Map();
    (p.pairs || []).forEach(function (pl) {
      var df = 0, ls = [], seen = {};
      pl.at.forEach(function (at) {
        var d = chunks[at.key];
        if (!d) return;
        var ids = at.how === 1 ? d.prefix(at.t, 40) : [], exact = d.find(at.t);
        if (exact >= 0 && ids.indexOf(exact) < 0) ids.unshift(exact);
        ids.forEach(function (id) {
          if (seen[at.key + '|' + id]) return;
          seen[at.key + '|' + id] = 1; df += d.df(id); ls.push({ d: d, id: id });
        });
      });
      if (!df) return;
      var wt = PAIR_W * Math.log(1 + (m.docs - df + 0.5) / (df + 0.5)), got = new Set();
      ls.forEach(function (l) {
        function add(dc) { if (!got.has(dc)) { got.add(dc); near.set(dc, (near.get(dc) || 0) + wt); } }
        if (l.d.mode(l.id) === 1) { l.d.docs(l.id).forEach(add); return; }
        var ps = l.d.postings(l.id);
        for (var k = 0; k < ps.length; k += 2) add(ps[k]);
      });
    });
    var fullMask = nW >= 31 ? -1 : (1 << nW) - 1;
    function weigh(mask, emask) {
      var sc = 0;
      for (var j = 0; j < nW; j++) if (mask & 1 << j) sc += idf[j] * (emask & 1 << j ? 1 : 0.7);
      return sc;
    }
    var per = {};
    function consider(i, src, mask, emask) {
      var sc = weigh(mask, emask), whole = (mask & fullMask) === fullMask;
      if (!whole && (nW < 2 || sc < total * 0.5)) return;
      var len = LEN && LEN[i] ? LEN[i] / avg : 1, norm = (K1 + 1) / (1 + K1 * (1 - B + B * len));
      var s = src >= 0 ? m.sources[src] : null, v = sc * norm * (whole ? 1 : 0.8 * sc / total);
      if (whole && (!s || s.stem)) v += near.get(i) || 0;
      if (s && opts.prefer && opts.prefer.indexOf(s.id) >= 0) v += 0.5;
      var e = per[i] || (per[i] = { best: 0, bestFull: 0, full: 0, hits: [] });
      if (whole) { e.full++; if (v > e.bestFull) e.bestFull = v; }
      if (v > e.best) e.best = v;
      if (s) e.hits.push({ src: s.id, kind: s.kind, _n: src, v: v + (whole ? 1000 : 0) });
    }
    cand.forEach(function (mask, code) {
      var i = Math.floor(code / 65536);
      consider(i, code % 65536, mask | (anyMask.get(i) | 0), (ex.get(code) | 0) | (anyEx.get(i) | 0));
    });
    anyMask.forEach(function (mask, i) { if (!per[i]) consider(i, -1, mask, anyEx.get(i) | 0); });
    Object.keys(per).forEach(function (key) {
      var i = +key, e = per[key];
      if (!e.best) return;
      e.hits.sort(function (x, y) { return y.v - x.v; });
      var score = tiered(e.full ? 1 : 0, ((e.full ? e.bestFull : e.best) + 0.1 * Math.log(1 + e.full)) * p.v.w), r = results[i];
      if (!r || score > r.score) {
        // no hits: found by frequent words only (bitmaps: no sources), given every source in _index
        results[i] = { i: i, score: score, all: e.full, hits: e.hits.length ? e.hits.slice(0, 8).map(function (h) {
          return { src: h.src, kind: h.kind, _n: h._n };
        }) : null };
      }
    });
  }

  /* the texts of the first results (the first page of the panel: 20): the tier by how close the words are (the phrase,
     near, see closeness), and the ranges to highlight (needs o.text) */
  Search.prototype._confirm = function (q, m, list, opts, warnings, ctx) {
    var within = ctx ? ctx.within : function (p) { return p; }, signal = ctx ? ctx.signal : null;
    var self = this, top = list.slice(0, opts.confirm != null ? opts.confirm : 20);
    if (!this.text || !top.length) return Promise.resolve(list);
    var multi = T.tokenize(q).length > 1 || T.tokenize(q).some(function (t) { return t.pairs; });
    var sc = T.detectScript(q), alt = sc === 'Latn' ? T.latToCyr(q) : sc === 'Cyrl' ? T.cyrToLat(q) : null;
    var altLat = sc === 'Latn';
    var srcIndex = {};
    m.sources.forEach(function (s, n) { srcIndex[s.id] = n; });
    return Promise.all(top.map(function (r) {
      var close = 0, same = false;
      return Promise.all((r.hits || []).slice(0, 2).map(function (h) {
        var n = h._n != null ? h._n : srcIndex[h.src], s = m.sources[n];
        if (!s) return null;
        return within(Promise.resolve(self.text(r.d, s.id, signal))).then(function (text) {
          if (typeof text !== 'string') return;
          var how = s.stem ? 'ar' : s.script;
          h.ranges = highlight(text, q, how);
          if (!h.ranges.length && alt) h.ranges = highlight(text, alt, how, altLat);      // Latin query, Cyrillic text
          if (multi && close < 3) close = Math.max(close, closeness(text, q, !!s.stem), alt ? closeness(text, alt, !!s.stem, altLat) : 0);
          // the phrase in the very forms of the query (المسجد is not مسجدا): first among the phrases of the stem
          if (multi && close === 3 && !same) same = phraseIn(text, q, false) || !!alt && phraseIn(text, alt, false);
        }, function (e) { if (!isAbort(e)) warnings.push('text ' + r.d + '/' + s.id + ': ' + (e && e.message || e)); });
      })).then(function () {
        var tier = tierOf(r.score), inner = r.score - tier * TIER;
        if (close > tier && tier < 3) { r.score = close * TIER + inner; tier = close; }
        if (same && tier === 3) { r.score = tier * TIER + inner + (TIER - inner) * 0.5; r.same = true; }
        r.score = Math.round(r.score * 100) / 100;
        if (r.tier != null) r.tier = Math.max(r.tier, Math.min(tier, 3));
      });
    })).then(function () {
      list.sort(function (x, y) { return y.score - x.score || x.d - y.d; });
      return list;
    }).then(null, function (e) { warnings.push('confirm: ' + e.message); return list; });
  };

  return {
    Search: Search, Builder: Builder, Loader: Loader,
    encodeChunk: encodeChunk, decodeChunk: decodeChunk, scriptKey: scriptKey, chunkOf: chunkOf, pickGroups: pickGroups,
    highlight: highlight, phraseIn: phraseIn, closeness: closeness, norm: T.norm, terms: T.terms, queryTerms: T.queryTerms, detectScript: T.detectScript,
    DEFAULT_LANGS: DEFAULT_LANGS, BASES: BASES, LEN_STEP: LEN_STEP,
    limits: { query: MAX_QUERY, words: MAX_WORDS, pairs: MAX_PAIRS, postings: MAX_POSTINGS }
  };
});
