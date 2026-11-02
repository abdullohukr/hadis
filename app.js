/* Ҳадислар: the panel for the website and the Word add-in task pane (the same design and ways as Khatt al-Quran,
   ../quranuz/web/app.js). The hadiths come only from the data (web/h, web/data, web/search) through HadisCore.Data. */
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  /* an old Word (Office 2016/2019 with Internet Explorer inside) or browser: without fetch, Promise and normalize no hadith
     is ever loaded (an endless «loading»): it is said instead */
  if (typeof window.fetch !== 'function' || typeof window.Promise !== 'function' || typeof String.prototype.normalize !== 'function') {
    var oldLang = 'uz';
    try { oldLang = I18n.ofScript((JSON.parse(localStorage.getItem('hadis-settings') || '{}') || {}).script); } catch (e) {}
    var oldMsg = typeof window.I18n === 'function' ? new I18n(oldLang).t('oldWord') : '';
    document.body.className = '';
    document.body.innerHTML = '<p class="old-host" role="alert">' + oldMsg.replace(/[&<>"]/g, '') + '</p>';
    return;
  }
  var C = window.HadisCore, T = window.HadisText;
  var FMT_DEFAULT = {
    ar: { font: '', size: 18, b: false, i: false, u: false, color: '', align: 'both', line: 0 },
    uz: { font: '', size: 14, b: false, i: false, u: false, color: '', align: 'both', line: 0 },
    ref: { font: '', size: 0, b: false, i: false, u: false, color: '' },
    hukm: { font: '', size: 0, b: false, i: true, u: false, color: '' }
  };
  var DEFAULTS = {
    script: 'cyrl', theme: 'auto', tab: 'general',             // the interface is Uzbek in the alphabet of `script`:
    apos: 'mod',                                       // 'cyrl' | 'latn' (oʻ gʻ) | 'new' (2026: ö ğ ş ç); apos: ʻ ʼ | ‘ ’ | '
    withAr: true, withUz: true, order: 'ar', mode: 'sahabi', // from the Companion (owner 2026-10-31) | 'full' | 'matn'
    refUz: true, refAr: false, withBook: false, withHukm: true, withBob: false,
    withGrades: true, gradesMatn: false,               // the gradings of other scholars (hukm2); in «only matn»: off
    allGrades: false,                                  // every grading in the tahrij, not only the main one
    matnBold: true, arMatnBold: true, newPara: true, baskSrc: 'list',
    picBgOn: false, picBg: '#ffffff', picAction: 'ask', extInsert: 'ask',
    day: true,                                         // the «Кун ҳадиси» card at the top
    dayClose: 'ask',                                   // its cross: 'ask' | 'day' (hide until tomorrow) | 'off' (turn it off)
    dayHide: '',                                       // hidden for this local date (dayKey) only
    dayAt: 0,                                          // when day / dayHide were chosen: the newest of all the stores wins
    sortBy: 'books',                                   // the results of a search: 'books' (BOOK_ORDER, then the number) | 'score'
    smooth: true,                                      // smooth changes (never with prefers-reduced-motion)
    zoom: 1                                            // the size of the texts in the big window «Кенгайтириш»
  };
  var settings = loadSettings(), i18n = new I18n(uiLang());
  var t = function (k, v) { return i18n.t(k, v); };
  /* ?host=gdocs | wp | ext | tg: the panel inside Google Docs, WordPress, the browser extension, Telegram (see sendToHost) */
  var HOST = (location.search.match(/[?&]host=(gdocs|wp|ext|tg)\b/) || [])[1] || '';
  if (window.parent === window && HOST !== 'tg') HOST = '';
  var inWord = false, online = !!HOST, windows = /Win/.test(navigator.platform || '');
  var data = null, catalog = null, search = null;
  /* the hadith on the card: d, the loaded hadith (h), its parsed text (P); sel: the mode and the chosen words
     (sel.ar / sel.uz: null = all the text, or the parts [[from, to], …] in the order of the text, HadisCore.ranges) */
  var cur = { d: -1, h: null, P: null };
  var sel = { d: -1, mode: settings.mode, ar: null, uz: null };
  var clickStart = null, muted = {};
  var extOrigin = '';
  function fromExt(e) {
    if (HOST !== 'ext') return true;
    if (!/^(chrome|moz|safari-web)-extension:\/\//.test(e.origin || '') || (extOrigin && e.origin !== extOrigin)) return false;
    extOrigin = e.origin; return true;
  }
  function toParent(msg) {
    if (HOST !== 'ext') { window.parent.postMessage(msg, '*'); return true; }
    if (extOrigin) { window.parent.postMessage(msg, extOrigin); return true; }
    if (msg.type === 'khatt-ready') { window.parent.postMessage(msg, '*'); return true; }
    return false;
  }

  /* ---------- settings (localStorage) ---------- */
  function loadSettings() {
    var st = JSON.parse(JSON.stringify(DEFAULTS)), saved = null;
    try { saved = JSON.parse(localStorage.getItem('hadis-settings') || 'null'); } catch (e) {}
    if (saved) Object.keys(saved).forEach(function (k) { if (saved[k] !== undefined) st[k] = saved[k]; });
    // the alphabet where localStorage is not kept (Word for Mac, a frame of Google Docs / the extension): the cookie
    var c = prefCookie();
    if (c && (!saved || !saved.script)) { if (c.script) st.script = c.script; if (c.apos) st.apos = c.apos; }
    if (!saved || saved.dayAt == null) {                 // the old key of «hide for today» (before 2026-10-30)
      try { var old = localStorage.getItem('hadis-day-closed'); if (old) { st.dayHide = old; st.dayAt = 1; } } catch (e) {}
    }
    if (c) takeDayPref(st, c);
    sanitize(st);
    return st;
  }
  /* «Кун ҳадиси» turned off / hidden for today: kept like the alphabet (localStorage, the cookie, the host's storage);
     each copy has the time of the choice (dayAt), the newest one wins: a lost or old copy never brings the card back */
  function takeDayPref(st, p) {
    var at = +p.dayAt || 0;
    if (!(at > (+st.dayAt || 0)) || typeof p.day !== 'boolean' && p.day !== 0 && p.day !== 1) return false;
    st.day = !!p.day; st.dayHide = typeof p.dayHide === 'string' ? p.dayHide : ''; st.dayAt = at;
    return true;
  }
  function dayPref() { return { day: settings.day, dayHide: settings.dayHide || '', dayAt: settings.dayAt || 0 }; }
  /* the choice of the alphabet, kept twice more than in localStorage: a cookie of a year (the same site) and, inside
     a host (the extension: storage.sync, Google Docs: the properties of the user), the message khatt-pref {script, apos} */
  function prefCookie() {
    try {
      var m = /(?:^|;\s*)hadis-pref=([^;]*)/.exec(document.cookie || '');
      var o = m ? JSON.parse(decodeURIComponent(m[1])) : null;
      return o && typeof o === 'object' ? o : null;
    } catch (e) { return null; }
  }
  function savePref() {
    var p = dayPref(), o = { script: settings.script, apos: settings.apos, day: p.day ? 1 : 0, dayHide: p.dayHide, dayAt: p.dayAt };
    var v = encodeURIComponent(JSON.stringify(o));
    // Secure is refused on http:// (except localhost): there the cookie is kept without it
    var sec = location.protocol === 'https:' ? '; SameSite=None; Secure' : '; SameSite=Lax';
    try { document.cookie = 'hadis-pref=' + v + '; max-age=31536000; path=/' + sec; } catch (e) {}
    if (HOST && HOST !== 'tg') try { toParent({ type: 'khatt-pref', script: settings.script, apos: settings.apos, day: p.day, dayHide: p.dayHide, dayAt: p.dayAt }); } catch (e) {}
  }
  /* damaged or foreign data in the storage must never stop the panel */
  function sanitize(st) {
    var oneOf = function (v, list, d) { return list.indexOf(v) >= 0 ? v : d; };
    delete st.uiLang;                                  // the old choice of the interface language (now: Uzbek only)
    ['theme', 'tab', 'picBg'].forEach(function (k) { if (typeof st[k] !== 'string') st[k] = DEFAULTS[k]; });
    ['withAr', 'withUz', 'refUz', 'refAr', 'withBook', 'withHukm', 'withGrades', 'gradesMatn', 'allGrades', 'withBob', 'matnBold', 'arMatnBold', 'newPara', 'picBgOn', 'day', 'smooth']
      .forEach(function (k) { st[k] = typeof st[k] === 'boolean' ? st[k] : DEFAULTS[k]; });
    st.script = oneOf(st.script, ['cyrl', 'latn', 'new'], 'cyrl');
    st.apos = oneOf(st.apos, ['mod', 'typo', 'ascii'], 'mod');
    st.order = oneOf(st.order, ['ar', 'uz'], 'ar');
    // «Санад» (2026-10-31): the default is from the Companion; the old default «Санад ва матн» ('full') moves to it once
    if (st.modeV !== 2) { if (st.mode !== 'matn') st.mode = 'sahabi'; st.modeV = 2; }
    st.mode = oneOf(st.mode, C.MODES, 'sahabi');
    st.baskSrc = oneOf(st.baskSrc, ['list', 'each'], 'list');
    st.picAction = oneOf(st.picAction, ['ask', 'insert', 'copy'], 'ask');
    st.extInsert = oneOf(st.extInsert, ['ask', 'text', 'picture'], 'ask');
    st.dayClose = oneOf(st.dayClose, ['ask', 'day', 'off'], 'ask');
    st.dayHide = typeof st.dayHide === 'string' && /^\d{4}-\d{1,2}-\d{1,2}$/.test(st.dayHide) ? st.dayHide : '';
    st.dayAt = typeof st.dayAt === 'number' && st.dayAt >= 0 ? st.dayAt : 0;
    st.sortBy = oneOf(st.sortBy, ['books', 'score'], 'books');
    st.zoom = typeof st.zoom === 'number' && st.zoom >= 0.7 && st.zoom <= 2 ? Math.round(st.zoom * 10) / 10 : 1;
    if (!/^#[0-9a-f]{6}$/i.test(st.picBg)) st.picBg = DEFAULTS.picBg;
    var f = st.fmt && typeof st.fmt === 'object' && !Array.isArray(st.fmt) ? st.fmt : {};
    st.fmt = {};
    Object.keys(FMT_DEFAULT).forEach(function (k) {
      var d = FMT_DEFAULT[k], v = f[k] && typeof f[k] === 'object' ? f[k] : {}, o = {};
      Object.keys(d).forEach(function (p) { o[p] = typeof v[p] === typeof d[p] ? v[p] : d[p]; });
      if (o.color && !/^#[0-9a-f]{6}$/i.test(o.color)) o.color = '';
      if (!(o.size >= 0 && o.size <= 96)) o.size = d.size;
      if (o.align !== undefined) o.align = oneOf(o.align, ['left', 'center', 'right', 'both'], 'both');
      st.fmt[k] = o;
    });
  }
  function saveSettings() { try { localStorage.setItem('hadis-settings', JSON.stringify(settings)); } catch (e) {} }

  var MAC = /Mac|iPhone|iPad/.test(navigator.platform || '');
  function toast(msg, act) {
    var el = $('toast'), top = document.querySelector('dialog[open]') || document.body;
    if (el.parentNode !== top) top.appendChild(el);     // over the open window (the top layer), else under it
    if (MAC) msg = String(msg).replace(/(Ctrl|Strg)\+/g, '⌘');
    el.textContent = msg;
    if (act) {
      var b = document.createElement('button');
      b.type = 'button'; b.className = 'toast-act'; b.textContent = act.label;
      b.addEventListener('click', function () { el.classList.remove('show'); act.fn(); });
      el.appendChild(b);
    }
    el.classList.add('show');
    clearTimeout(toast.t); toast.t = setTimeout(function () { el.classList.remove('show'); }, act ? 7000 : 2200);
  }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function latn() { return settings.script !== 'cyrl'; }            // either Latin: the texts are made in the canonical one
  function uzText(s) { return latn() ? C.toLatin(s) : s; }
  function uzLang() { return latn() ? 'uz_latn' : 'uz'; }
  /* the panel speaks the canonical Latin (uz_latn); the 2026 Latin and the other apostrophes are made of it where the
     text leaves the code: the page (styleDom), the title, what is inserted or copied (fmt) */
  function uiLang() { return settings.script === 'cyrl' ? 'uz' : 'uz_latn'; }
  function styleOpts() { return { nw: settings.script === 'new', apos: latn() ? settings.apos : 'mod' }; }
  function styled() { return settings.script === 'new' || latn() && settings.apos !== 'mod'; }
  function latStyle(s) { return styled() ? T.latinStyle(s, styleOpts()) : s; }
  /* HadisCore.format in the style of the Latin chosen (the Uzbek paragraphs; Arabic has no Latin letters) */
  function fmt(h, b, o) {
    var out = C.format(h, b, o);
    if (!out || !styled()) return out;
    out.paras.forEach(function (p) { if (p.dir !== 'rtl') p.runs.forEach(function (r) { r.t = latStyle(r.t); }); });
    out.text = latStyle(out.text);
    if (typeof out.ref === 'string') out.ref = latStyle(out.ref);
    return out;
  }
  /* the hadith as it is shown: in the Latin alphabet with the Uzbek texts of tools/build_latin.js (Harfsoz, built on our
     machine; a block of 100 loaded only now), else as in the data (the letters are then changed by HadisCore.toLatin) */
  var latMeta = null;
  function latinMeta() {
    if (!latMeta) {
      latMeta = data.loader.get('data/latin.json?t=' + Math.floor(Date.now() / 3600000), 'json', false).then(function (m) {
        return m && m.data === data.version && /^[\w]{1,16}$/.test(m.v || '') ? m : null;
      });
      latMeta.then(null, function () { latMeta = null; });
    }
    return latMeta;
  }
  function hadithView(d, signal) {
    return data.hadith(d, signal).then(function (h) {
      if (!latn()) return h;
      var b = catalog.byId[h.b], off = h.d - b.doc[0], k = Math.floor(off / catalog.block);
      return latinMeta().then(function (m) {
        if (!m) return h;
        return data.loader.get('h/v' + data.version + '/lat-' + m.v + '/' + h.b + '/' + k + '.json', 'json', true, signal).then(function (arr) {
          var x = Array.isArray(arr) ? arr[off - k * catalog.block] : null, o = {};
          if (!x) return h;
          for (var key in h) o[key] = h[key];
          if (x.uz) o.uz = x.uz;
          if (x.bob) o.bob = x.bob;
          // the start at the Companion in the Latin text (tools/build_latin.js): its prefix in Latin, 0 — the whole text
          if (x.su !== undefined && h.st) o.st = { a: h.st.a, u: x.su || undefined };
          if (x.hk && Array.isArray(h.hukm)) o.hukm = h.hukm.map(function (y, i) {
            if (!x.hk[i]) return y;
            var z = {}; for (var kk in y) z[kk] = y[kk]; z.text = x.hk[i]; return z;
          });
          return o;
        });
      }).then(null, function () { return h; });
    });
  }
  /* names of the books in the chosen alphabet */
  function nameLang() { return uzLang(); }
  function shortName(id) { return C.bookName(id, nameLang(), 'short'); }
  function titleName(id) {
    var b = catalog && catalog.byId[id], L = nameLang(), nm = b && b.name;
    var v = nm && (L === 'uz' ? nm.uz : L === 'uz_latn' ? nm.uzl : nm[L]);
    return v || C.bookName(id, L, 'book');
  }
  function refText(d) {
    var r = data && data.refOf(d);
    return r ? shortName(r.book) + ' ' + C.refLabel(r) : '#' + (d + 1);
  }

  /* ---------- interface language: Uzbek, the alphabet follows settings.script ---------- */
  function applyI18n() {
    document.documentElement.lang = latn() ? 'uz-Latn' : 'uz-Cyrl';
    document.documentElement.dir = 'ltr';
    [].forEach.call(document.querySelectorAll('[data-i18n]'), function (el) { el.textContent = t(el.dataset.i18n); });
    [].forEach.call(document.querySelectorAll('[data-i18n-placeholder]'), function (el) {
      el.placeholder = t(el.dataset.i18nPlaceholder);
      // a field with no label of its own (the search): its name for a screen reader is the placeholder
      if (!el.hasAttribute('aria-labelledby') && !(el.id && document.querySelector('label[for="' + el.id + '"]')) && !el.closest('label')) el.setAttribute('aria-label', el.placeholder);
    });
    [].forEach.call(document.querySelectorAll('[data-i18n-title]'), function (el) {
      el.title = t(el.dataset.i18nTitle);
      if (el.hasAttribute('aria-label')) el.setAttribute('aria-label', el.title);
    });
    document.title = latStyle(t('appName'));
  }
  function refreshAll() {
    fillBooks(); render(); drawMarks(); fmtEditor(); refSample(); paintDay();
    if (catalog) $('status').textContent = '';
    if (searchQ.trim()) runSearch(); else statusLine();
  }

  /* ---------- themes: light (ink on paper), dark, paper (warm, for long reading), high contrast; «auto» = as the system.
     The themes of the older versions (ocean, forest, …) go to the nearest of these. ---------- */
  var THEMES = [['auto', '#ffffff', '#131210', '#b0361c'], ['light', '#ffffff', '#17140f', '#b0361c'], ['dark', '#131210', '#f1ece3', '#f08a68'],
    ['sepia', '#f3ead8', '#2e2418', '#9a3a1c'], ['contrast', '#000000', '#ffffff', '#ffd60a']];
  var DARK = { dark: 1, contrast: 1 };
  (function () {
    var old = { forest: 'dark', midnight: 'dark', graphite: 'dark', nord: 'dark', desert: 'sepia', ocean: 'light', rose: 'light', lavender: 'light' };
    if (old[settings.theme]) settings.theme = old[settings.theme];
    else if (!THEMES.some(function (x) { return x[0] === settings.theme; })) settings.theme = 'auto';
  })();
  function isDark() {
    if (settings.theme === 'auto') return !!(window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches);
    return !!DARK[settings.theme];
  }
  function applyTheme() {
    if (settings.theme && settings.theme !== 'auto') document.documentElement.setAttribute('data-theme', settings.theme);
    else document.documentElement.removeAttribute('data-theme');
    $('theme-btn').setAttribute('aria-pressed', isDark());
  }
  function fillThemes() {
    $('s-theme').innerHTML = THEMES.map(function (x) {
      var cap = x[0].charAt(0).toUpperCase() + x[0].slice(1);
      return '<button type="button" class="theme" role="radio" data-theme-id="' + x[0] + '" aria-checked="' + (settings.theme === x[0]) + '">' +
        '<span class="sw"><b style="background:' + x[1] + '"></b><b style="background:' + x[2] + '"></b><b style="background:' + x[3] + '"></b></span>' +
        '<span class="nm">' + esc(t('theme' + cap)) + '</span></button>';
    }).join('');
  }
  $('s-theme').addEventListener('click', function (e) {
    var b = e.target.closest('.theme'); if (!b) return;
    settings.theme = b.dataset.themeId; applyTheme(); saveSettings();
    [].forEach.call($('s-theme').children, function (x) { x.setAttribute('aria-checked', x === b); });
  });
  /* the moon in the top bar: dark <-> light at once */
  $('theme-btn').addEventListener('click', function () {
    settings.theme = isDark() ? 'light' : 'dark'; applyTheme(); saveSettings(); fillThemes();
  });
  if (window.matchMedia) try { matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyTheme); } catch (e) {}
  applyTheme();

  /* ---------- settings tabs ---------- */
  function showTab(name) {
    if (!document.querySelector('.tab[data-tab="' + name + '"]')) name = 'general';
    [].forEach.call(document.querySelectorAll('.tabs [data-tab]'), function (b) { b.setAttribute('aria-selected', b.dataset.tab === name); });
    [].forEach.call(document.querySelectorAll('section.tab'), function (s) { s.hidden = s.dataset.tab !== name; });
    settings.tab = name;
  }
  document.querySelector('.tabs').addEventListener('click', function (e) {
    var b = e.target.closest('[data-tab]'); if (!b) return;
    showTab(b.dataset.tab); saveSettings();
    document.querySelector('.sheet-body').scrollTop = 0;
  });

  /* ---------- the books and the number ---------- */
  function fillBooks() {
    if (!catalog) return;
    $('book').innerHTML = catalog.books.slice().sort(function (a, b) { return C.bookRank(a.id) - C.bookRank(b.id) || a.doc[0] - b.doc[0]; }).map(function (b) {
      return '<option value="' + b.id + '">' + esc(titleName(b.id)) + (b.label ? ' · ' + esc(uzOrUi(b.label)) : '') + '</option>';
    }).join('');
    var r = cur.d >= 0 && data.refOf(cur.d);
    if (r) $('book').value = r.book;
  }
  /* «Саҳиҳ» (books.json label, Uzbek Cyrillic): in Latin letters / the interface language */
  function uzOrUi(s) {
    if (s === 'Саҳиҳ') return t('labelSahih');
    return uzText(s);
  }
  $('book').addEventListener('change', function () {
    var b = catalog.byId[this.value]; if (!b) return;
    go(b.doc[0]);
  });
  function numFromInput() {
    var m = /^\s*(\d{1,5})(?:\s*[/:.]\s*(\d{1,4}))?\s*$/.exec(C.latinDigits($('num').value));
    if (!m) return;
    var book = $('book').value, d = data.docOf({ book: book, n: +m[1], s: m[2] ? +m[2] : null });
    if (d < 0) { toast(t('noNumber', { ref: shortName(book) + ' ' + m[1] + (m[2] ? '/' + m[2] : '') })); $('num').value = cur.d >= 0 ? C.refLabel(data.refOf(cur.d)) : ''; return; }
    go(d);
  }
  $('num').addEventListener('change', numFromInput);
  $('num').addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); numFromInput(); } });
  $('num').addEventListener('focus', function () { var x = this; setTimeout(function () { x.select(); }, 0); });

  /* ---------- the hadith on the card ---------- */
  var goSeq = 0;
  function go(d, keep) {
    if (!catalog || !(d >= 0 && d < catalog.docs)) return Promise.resolve();
    var seq = ++goSeq, card = document.querySelector('.selection');
    sel = { d: d, mode: keep && keep.mode || settings.mode, ar: C.ranges(keep && keep.ar), uz: C.ranges(keep && keep.uz) };
    clickStart = null;
    var r = data.refOf(d);
    if (r) { $('book').value = r.book; $('num').value = C.refLabel(r); }
    // the text on the card stays (dimmed) until the new one is there: the card does not shrink to «loading» and back
    var had = !!cur.P;
    cur = { d: d, h: null, P: had ? cur.P : null, old: had };
    if (had) { card.classList.add('busy'); titleOnly(); } else render();
    if (!keep || !keep.noHash) setHash(d);
    return (lastGo = hadithView(d).then(function (h) {
      if (seq !== goSeq) return;
      cur = { d: d, h: h, P: C.parseHadith(h, sel.mode) };
      card.classList.remove('busy');
      steady(render);
    }, function (e) {
      if (seq !== goSeq) return;
      card.classList.remove('busy'); cur.P = null;
      $('preview-ar').innerHTML = ''; $('preview-uz').innerHTML = '<p class="muted">' + esc(t('loadFailed')) + ' ' + esc(e && e.message || '') + '</p>';
    }));
  }
  var lastGo = null;                                   // the loading of the last go() (the first screen waits for it)
  function setHash(d) {
    var r = data.refOf(d); if (!r) return;
    var h = '#' + r.book + ':' + C.refLabel(r);
    if (location.hash !== h) try { if (typeof history.replaceState === 'function') history.replaceState(null, '', location.pathname + location.search + h); else location.hash = h; } catch (e) {}
  }
  function fromHash() {
    var m = /^#([a-z]+):(\d{1,5})(?:\/(\d{1,4}))?$/.exec(location.hash);
    if (m && catalog && catalog.byId[m[1]]) {
      var d = data.docOf({ book: m[1], n: +m[2], s: m[3] ? +m[3] : null });
      if (d >= 0) { if (d !== cur.d) go(d, { noHash: true }); return true; }
    }
    if (location.hash === '#settings') {
      try { history.replaceState(null, '', location.pathname + location.search); } catch (e) {}
      $('settings-btn').click(); return false;
    }
    return false;
  }
  window.addEventListener('hashchange', fromHash);
  function step(dir) { if (cur.d >= 0) go(Math.max(0, Math.min(catalog.docs - 1, cur.d + dir))); }
  $('prev').addEventListener('click', function () { step(-1); });
  $('next').addEventListener('click', function () { step(1); });

  /* the text of the card in the mode: from the Companion ('sahabi', 'matn') or all ('full') */
  function reparse() { if (cur.h) cur.P = C.parseHadith(cur.h, sel.mode); }
  /* the words of the two texts in the current mode (the selection counts in them) */
  function modeOf(side) {
    if (sel.mode !== 'matn' || !cur.P) return 'full';
    return C.hasMatn(cur.P[side].paras) ? 'matn' : 'full';
  }
  function wordList(side) { return cur.P ? C.words(cur.P[side].paras, modeOf(side)) : []; }
  function inParts(parts, k) { for (var i = 0; i < parts.length; i++) if (k >= parts[i][0] && k <= parts[i][1]) return i; return -1; }
  function wordState(side, k) {
    var r = sel[side];
    if (clickStart && clickStart.side === side && clickStart.k === k) return 'start';
    if (r && inParts(r, k) < 0) return 'off';
    if (r && r.length > 1) return 'on';
    return '';
  }
  function wordsHtml(side) {
    var list = wordList(side), html = '', para = -1, lat = side === 'uz' && latn();
    list.forEach(function (w, i) {
      if (w.para !== para) {
        html += side === 'uz' ? (para >= 0 ? '</p>' : '') + '<p>' : para >= 0 ? '<br>' : '';
        para = w.para;
      }
      if (w.gap) html += '<span class="gap">…</span> ';
      var inner = '', prev = ' ';
      w.parts.forEach(function (x) {
        var tx = esc(lat ? C.toLatin(x.t, prev) : x.t); prev = x.t.slice(-1);
        if (x.b) tx = '<b>' + tx + '</b>';
        if (x.i) tx = '<i>' + tx + '</i>';
        inner += tx;
      });
      html += '<span class="w' + (w.matn ? ' m' : '') + (wordState(side, i) ? ' ' + wordState(side, i) : '') + '" data-side="' + side + '" data-k="' + i + '">' + inner + '</span> ';
    });
    if (side === 'uz' && para >= 0) html += '</p>';
    return html;
  }
  function titleOnly() {
    var d = cur.d, r = d >= 0 ? data.refOf(d) : null;
    $('sel-title').innerHTML = r ? esc(shortName(r.book)) + ' <bdi>' + esc(C.refLabel(r)) + '</bdi> <span class="bk">· ' + esc(titleName(r.book)) + '</span>' : '';
  }
  /* A change of the card that may change its height (the mode, another alphabet, the next hadith): the card stays
     where it was on the screen. When the page gets shorter than the place it was scrolled to, the browser would pull
     it down: a spacer under the page holds the place and is cut away as soon as it is out of sight (scroll). */
  var spacer = 0;
  function setSpacer(px) { spacer = Math.max(0, Math.round(px)); document.body.style.paddingBottom = spacer ? spacer + 'px' : ''; }
  window.addEventListener('scroll', function () {
    if (!spacer) return;
    var se = document.scrollingElement || document.documentElement, below = se.scrollHeight - se.scrollTop - window.innerHeight;
    if (below > 0) setSpacer(spacer - below);
  }, { passive: true });
  function steady(fn) {
    var card = document.querySelector('.selection'), se = document.scrollingElement || document.documentElement;
    var top = card.getBoundingClientRect().top, y = se.scrollTop, bx = $('big-body'), by = bx ? bx.scrollTop : 0;
    fn();
    var want = y + card.getBoundingClientRect().top - top, max = se.scrollHeight - window.innerHeight;
    if (want > max + 1) setSpacer(spacer + want - max);
    if (Math.abs(se.scrollTop - want) > 1) se.scrollTop = want;
    if (bx && $('big').open) bx.scrollTop = by;
  }
  function render() {
    if (!catalog) return;
    titleOnly();
    [].forEach.call($('mode').children, function (b) { b.setAttribute('aria-pressed', b.dataset.mode === sel.mode); });
    [].forEach.call($('script').children, function (b) { b.setAttribute('aria-pressed', b.dataset.script === settings.script); });
    drawChips();
    if (!cur.P) {
      $('preview-ar').innerHTML = ''; $('preview-uz').innerHTML = '<p class="muted">' + esc(t('loading')) + '</p>';
      $('tahrij-box').innerHTML = ''; $('bob').hidden = true; $('matn-note').hidden = true;
      starState(); bigDraw(); return;
    }
    if (!cur.h) return;                                // the old text stays while the new one loads (go)
    var P = cur.P, b = catalog.byId[cur.h.b];
    // the chapter heading (taken from the end of the previous hadith by the builder)
    if (P.bob.length) {
      $('bob').innerHTML = P.bob.map(function (p) { return p.map(function (x) { var s = esc(uzText(x.t)); return x.b ? '<b>' + s + '</b>' : s; }).join(''); }).join('<br>');
      $('bob').hidden = false;
    } else $('bob').hidden = true;
    $('preview-ar').innerHTML = P.ar.paras.length ? wordsHtml('ar') : '<span class="muted small" dir="' + i18n.dir() + '">' + esc(t('noArabic')) + '</span>';
    $('preview-uz').innerHTML = P.uz.paras.length ? wordsHtml('uz') : '<p class="muted small">' + esc(t('noUzbek')) + '</p>';
    dimTexts();
    $('matn-note').hidden = !(sel.mode === 'matn' && (modeOf('ar') === 'full' && P.ar.paras.length || modeOf('uz') === 'full' && P.uz.paras.length));
    // tahrij and the words of the author of the book (only from the data, with who said them)
    var tj = '<div class="tj">' + esc(C.tahrij(cur.h, b, uzLang(), { withBook: settings.withBook })) + '</div>' +
      '<div class="tj-ar" lang="ar">' + esc(C.tahrij(cur.h, b, 'ar', { withBook: settings.withBook })) + '</div>';
    if (b && b.label) tj += '<div><span class="label">' + esc(t('bookLabel', { label: uzOrUi(b.label) })) + '</span> <span class="muted small">' + esc(t('sahihNote')) + '</span></div>';
    tj += riwayatHtml(cur.h);
    var hk = C.hukmOf(cur.h, b, uzLang()), hkAr = C.hukmOf(cur.h, b, 'ar').filter(function (x) { return x.lang === 'ar'; });
    function hkHtml(x) {
      var by = x.book ? shortName(x.book) : x.byName;
      return '<div class="hk"><span class="by">' + esc(t('hukmBy', { by: by })) + '</span><q' + (x.lang === 'ar' ? ' dir="rtl" lang="ar"' : '') + '>' + esc(x.text) + '</q></div>';
    }
    hk.forEach(function (x) { tj += hkHtml(x); });
    if (hk.length && hk[0].lang !== 'ar') hkAr.forEach(function (x) { tj += hkHtml(x); });
    // the gradings of other scholars: a box of its own under the words of the author (never mixed with them)
    var gr = C.gradesOf(cur.h, b, uzLang());
    if (gr.length) {
      tj += '<div class="g2"><span class="by">' + esc(t('grades')) + '</span>' + gr.map(function (x) {
        var src = x.src === 'matn' ? t('gradeFromText') : (x.src || '').replace(/^https?:\/\//, '');
        var tip = t('gradeSrc', { src: src, c: x.c == null ? '—' : String(x.c) });
        return '<div class="g2-r" title="' + esc(tip) + '"><b>' + esc(x.byName) + ':</b> ' + esc(x.g) +
          (x.ar ? ' <bdi class="g2-ar" dir="rtl" lang="ar">' + esc(x.ar) + '</bdi>' : '') +
          (x.work ? ' — <span class="g2-w">' + esc(x.work) + (x.ref != null ? ', ' + esc(String(x.ref)) : '') +
            (x.vp ? ' (' + esc(t('gradeVp', { vp: x.vp })) + ')' : '') + '</span>' : '') + '</div>';
      }).join('') + '</div>';
    }
    // no grading at all (at-Tabarani: none in the data yet): said honestly, never guessed (the owner, 2026-11-02);
    // the books «Саҳиҳ» (Bukhari, Muslim) have their label instead
    if (!hk.length && !gr.length && !(b && b.label)) tj += '<div class="hk no-hukm muted small" title="' + esc(t('noHukmHint')) + '">' + esc(t('noHukm')) + '</div>';
    tj += '<a class="src-link" href="sources.html#' + esc(cur.h.b) + '" target="_blank" rel="noopener">' + esc(t('sourceLink')) + '</a>';
    $('tahrij-box').innerHTML = tj;
    selState();
    fmtPreview();
    bigDraw();
  }
  /* The same hadith in the other Sahih (the field x of the data: [[book, n], …], tools/link_riwayat.py): «Муслимда:
     №1907» on the card, a link opens it. Only on the card: never in the inserted text or the tahrij (the owner). */
  function riwayatHtml(h) {
    var by = {}, order = [];
    (Array.isArray(h && h.x) ? h.x : []).forEach(function (y) {
      if (!Array.isArray(y) || !catalog || !catalog.byId[y[0]] || !(y[1] > 0)) return;
      if (!by[y[0]]) { by[y[0]] = []; order.push(y[0]); }
      if (by[y[0]].indexOf(+y[1]) < 0) by[y[0]].push(+y[1]);
    });
    return order.map(function (id) {
      return '<div class="xr" title="' + esc(t('alsoInHint')) + '">' + esc(t('alsoIn', { book: shortName(id) })) + ' ' + by[id].map(function (n, i) {
        return '<a class="xr-a" href="#' + esc(id) + ':' + n + '">' + (i ? '' : '№') + n + '</a>';
      }).join(', ') + '</div>';
    }).join('');
  }
  /* after a change of the chosen words: only the classes of the words (no new text, the height stays), the buttons */
  function paintWords(side) {
    [].forEach.call(document.querySelectorAll('[data-words="' + side + '"] .w'), function (el) {
      var st = wordState(side, +el.dataset.k);
      el.classList.toggle('off', st === 'off'); el.classList.toggle('start', st === 'start'); el.classList.toggle('on', st === 'on');
      el.classList.remove('drag');
    });
  }
  function selState() {
    var any = !!(sel.ar || sel.uz);
    // the button keeps its place (visibility), the line under the text does not change its height
    [$('reset-words'), $('big-reset')].forEach(function (b) { if (b) { b.classList.toggle('idle', !any); b.disabled = !any; } });
    var n = (sel.ar ? sel.ar.length : 0) + (sel.uz ? sel.uz.length : 0);
    $('pick-n').textContent = n > 1 ? t('partsN', { n: n }) : '';
    if ($('big-n')) $('big-n').textContent = $('pick-n').textContent;
    starState();
  }
  function selChanged(side) {
    if (side) paintWords(side); else { paintWords('ar'); paintWords('uz'); }
    selState(); dimTexts(); fmtPreview();
  }
  /* the parts that go in (chips on the card: a quick «not this time», the settings stay) */
  function drawChips() {
    [].forEach.call($('chips').children, function (b) {
      var k = b.dataset.part, on = settings[k] && !muted[k];
      b.classList.toggle('off', !on); b.setAttribute('aria-pressed', on);
      b.title = t('chipHint');
    });
  }
  $('chips').addEventListener('click', function (e) {
    var b = e.target.closest('.tchip'); if (!b) return;
    var k = b.dataset.part;
    if (!settings[k]) { settings[k] = true; muted[k] = false; saveSettings(); fillSettingsForm(); }
    else muted[k] = !muted[k];
    drawChips(); dimTexts(); fmtPreview();
  });
  /* a part chosen in one text only: the other text does not go in (the owner, 2026-11-02: «a part chosen, the whole
     hadith inserted» — the whole translation went in with a few Arabic words); it is dimmed on the card */
  function sideOut(s, side) { var o = side === 'ar' ? 'uz' : 'ar'; return !(s[o] && !s[side]); }
  function dimTexts() {
    ['ar', 'uz'].forEach(function (s) {
      var k = s === 'ar' ? 'withAr' : 'withUz', off = !settings[k] || !!muted[k] || !sideOut(sel, s);
      [].forEach.call(document.querySelectorAll('[data-words="' + s + '"]'), function (el) { el.classList.toggle('dim', off); });
    });
  }
  $('mode').addEventListener('click', function (e) {
    var b = e.target.closest('[data-mode]'); if (!b || b.dataset.mode === sel.mode) return;
    sel.mode = b.dataset.mode; sel.ar = sel.uz = null; clickStart = null;
    settings.mode = sel.mode; saveSettings(); $('s-mode').value = sel.mode;
    reparse();
    steady(render);
  });
  $('script').addEventListener('click', function (e) {
    var b = e.target.closest('[data-script]'); if (!b) return;
    setScript(b.dataset.script);
  });
  function setScript(s, quiet) {
    if (settings.script === s || ['cyrl', 'latn', 'new'].indexOf(s) < 0) return;
    var wasLat = latn();
    settings.script = s; saveSettings(); if (!quiet) savePref(); $('s-script').value = s;
    i18n.set(uiLang()); applyI18n(); fillThemes(); refreshAll(); styleDom();   // the interface follows the alphabet of the texts
    // the other alphabet of the hadith on the card and of the hadith of the day (the Latin texts of the build)
    if (wasLat !== latn()) {
      if (cur.d >= 0) go(cur.d, { mode: sel.mode, ar: sel.ar, uz: sel.uz, noHash: true });
      drawDay();
    }
  }
  function setApos(a, quiet) {
    if (settings.apos === a || ['mod', 'typo', 'ascii'].indexOf(a) < 0) return;
    settings.apos = a; saveSettings(); if (!quiet) savePref(); $('s-apos').value = a;
    applyI18n(); styleDom(); render();
  }

  /* ---------- the 2026 Latin and the apostrophes on the page ----------
     Every text node and title / placeholder / aria-label of the page in the canonical Latin is shown in the style of
     settings (HadisText.latinStyle), whoever wrote it (the strings, the hadith, the tahrij, a toast); its canonical text
     is remembered, so another style (or the canonical one) is made of it again. Not touched: [translate="no"], the
     fields, the scripts, the Arabic and Cyrillic texts (no Latin letters). */
  var styleMemo = typeof WeakMap === 'function' ? new WeakMap() : null, styleBusy = false;
  var STYLE_ATTRS = ['title', 'placeholder', 'aria-label'];
  function keepEl(el) { return !el || (el.closest && el.closest('[translate="no"], script, style, textarea, code')); }
  function styleText(node, opts) {
    var v = node.nodeValue, m = styleMemo.get(node), orig = m && v === m.conv ? m.orig : v;
    if (!/[A-Za-zʻʼ‘’]/.test(orig)) { if (m) styleMemo.delete(node); return; }
    var conv = opts ? T.latinStyle(orig, opts) : orig;
    if (conv !== v) node.nodeValue = conv;
    if (conv !== orig) styleMemo.set(node, { orig: orig, conv: conv }); else if (m) styleMemo.delete(node);
  }
  function styleAttrs(el, opts) {
    var m = styleMemo.get(el) || {};
    STYLE_ATTRS.forEach(function (a) {
      if (!el.hasAttribute(a)) return;
      var v = el.getAttribute(a), x = m[a], orig = x && v === x.conv ? x.orig : v;
      if (!/[A-Za-zʻʼ‘’]/.test(orig)) return;
      var conv = opts ? T.latinStyle(orig, opts) : orig;
      if (conv !== v) el.setAttribute(a, conv);
      m[a] = { orig: orig, conv: conv };
    });
    styleMemo.set(el, m);
  }
  function styleTree(root, opts) {
    if (root.nodeType === 3) { if (!keepEl(root.parentElement)) styleText(root, opts); return; }
    if (root.nodeType !== 1 || keepEl(root)) return;
    styleAttrs(root, opts);
    var w = document.createTreeWalker(root, 5, null), n;   // elements and texts
    while ((n = w.nextNode())) {
      if (n.nodeType === 1) { if (!keepEl(n)) styleAttrs(n, opts); }
      else if (!keepEl(n.parentElement)) styleText(n, opts);
    }
  }
  function styleDom() {
    if (!styleMemo || !document.body) return;
    styleBusy = true;
    try { styleTree(document.body, styled() ? styleOpts() : null); } finally { styleBusy = false; }
  }
  if (styleMemo && typeof MutationObserver === 'function') {
    new MutationObserver(function (list) {
      if (styleBusy || !styled()) return;
      styleBusy = true;
      try {
        var opts = styleOpts();
        list.forEach(function (r) {
          if (r.type === 'characterData') { if (!keepEl(r.target.parentElement)) styleText(r.target, opts); }
          else if (r.type === 'attributes') { if (!keepEl(r.target)) styleAttrs(r.target, opts); }
          else [].forEach.call(r.addedNodes, function (n) { styleTree(n, opts); });
        });
      } finally { styleBusy = false; }
    }).observe(document.body, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: STYLE_ATTRS });
  }
  function clearPick() { sel.ar = sel.uz = null; clickStart = null; selChanged(); }
  $('reset-words').addEventListener('click', clearPick);

  /* A part of a text (the Arabic and the Uzbek, each on its own): tap the first word, then the last one, or drag across
     the words. With Ctrl / ⌘ a part is added (several places: inserted in the order of the text, « … » between them;
     Ctrl / ⌘ on a chosen part takes it away); with Shift the nearest part grows up to the word. */
  function applyRange(side, a, z, how) {
    if (z < a) { var x = a; a = z; z = x; }
    var base = how === 'add' ? (clickStart && clickStart.side === side ? clickStart.base : sel[side]) || [] : [];
    sel[side] = C.ranges(base.concat([[a, z]])); clickStart = null; selChanged(side);
  }
  function clickWord(side, k, how) {
    var parts = sel[side] || [];
    if (clickStart && clickStart.side === side) {                    // the second tap closes the part begun
      var cs = clickStart; clickStart = null;
      sel[side] = C.ranges((cs.base || []).concat([[cs.k, k]])); selChanged(side); return;
    }
    if (how === 'extend' && parts.length) {
      var best = 0, dist = Infinity;
      parts.forEach(function (p, i) { var dd = k < p[0] ? p[0] - k : k > p[1] ? k - p[1] : 0; if (dd < dist) { dist = dd; best = i; } });
      var np = parts.map(function (p, i) { return i === best ? [Math.min(p[0], k), Math.max(p[1], k)] : p; });
      sel[side] = C.ranges(np); selChanged(side); return;
    }
    if (how === 'add' && parts.length) {
      var i = inParts(parts, k);
      if (i >= 0) {                                                   // Ctrl / ⌘ on a chosen part: it goes
        sel[side] = C.ranges(parts.filter(function (p, j) { return j !== i; })); selChanged(side); return;
      }
      clickStart = { side: side, k: k, base: parts.slice() };
      sel[side] = C.ranges(parts.concat([[k, k]])); selChanged(side); return;
    }
    if (clickStart && clickStart.side !== side) clickStart = null;     // the other text: its first word stays chosen
    clickStart = { side: side, k: k, base: [] };
    sel[side] = [[k, k]]; selChanged(side);
  }
  function howOf(e) { return MAC ? (e.metaKey ? 'add' : e.shiftKey ? 'extend' : '') : (e.ctrlKey || e.metaKey ? 'add' : e.shiftKey ? 'extend' : ''); }
  var drag = null;
  function wordAt(x, y, box) {
    var el = document.elementFromPoint(x, y);
    el = el && el.closest ? el.closest('.w[data-side]') : null;
    if (el && box && !box.contains(el)) return null;
    return el ? { side: el.dataset.side, k: +el.dataset.k } : null;
  }
  function paintDrag() {
    var lo = Math.min(drag.start.k, drag.end.k), hi = Math.max(drag.start.k, drag.end.k);
    [].forEach.call(drag.box.querySelectorAll('.w'), function (el) {
      var k = +el.dataset.k; el.classList.toggle('drag', k >= lo && k <= hi);
    });
  }
  function wordsBox(el) {
    el.addEventListener('pointerdown', function (e) {
      if (e.button > 0 || !cur.h) return;                           // the old text while the next hadith loads: not chosen
      var p = wordAt(e.clientX, e.clientY, el); if (!p) return;
      e.preventDefault();
      drag = { start: p, end: p, moved: false, box: el, how: howOf(e) };
    });
    // Ctrl+click on a Mac is the right button: no menu over the words
    el.addEventListener('contextmenu', function (e) { if (e.ctrlKey && e.target.closest('.w')) e.preventDefault(); });
  }
  [].forEach.call(document.querySelectorAll('[data-words]'), wordsBox);
  document.addEventListener('pointermove', function (e) {
    if (!drag) return;
    var p = wordAt(e.clientX, e.clientY, drag.box);
    if (!p || p.side !== drag.start.side || p.k === drag.end.k) return;
    drag.end = p; drag.moved = true;
    drag.box.classList.add('dragging');
    paintDrag();
  });
  document.addEventListener('pointerup', function () {
    if (!drag) return;
    var d = drag; drag = null;
    d.box.classList.remove('dragging');
    if (d.moved) applyRange(d.start.side, d.start.k, d.end.k, d.how === 'add' ? 'add' : ''); else clickWord(d.start.side, d.start.k, d.how);
  });
  /* A part selected the usual way, with the mouse or a finger (a selection of the browser: the card of «Кун ҳадиси»,
     old Word for Windows (IE 11 knows no user-select), a long press on a phone): the words it touches are the part.
     Read when Insert / Copy is pressed (pointerdown / mousedown / touchstart, before the press takes the selection away)
     or by Ctrl+Enter. The card: put into sel (seen on the card, the same as taps); the card of the day: this insert. */
  var NATIVE_BOXES = [['preview-ar', 'ar', 'card'], ['preview-uz', 'uz', 'card'], ['big-ar', 'ar', 'card'], ['big-uz', 'uz', 'card'],
                      ['day-ar', 'ar', 'day'], ['day-uz', 'uz', 'day']];
  function touches(rg, el) {                                       // the selection holds a letter of el
    try {
      var r = document.createRange(); r.selectNodeContents(el);
      if (rg.compareBoundaryPoints(Range.START_TO_START, r) > 0) r.setStart(rg.startContainer, rg.startOffset);
      if (rg.compareBoundaryPoints(Range.END_TO_END, r) < 0) r.setEnd(rg.endContainer, rg.endOffset);
      return /\S/.test(r.toString());
    } catch (e) { return false; }
  }
  function nativePick() {
    var s = window.getSelection ? window.getSelection() : null, out = null;
    if (!s || !s.rangeCount || s.isCollapsed) return null;          // (String(s) is '' under user-select: none)
    for (var i = 0; i < s.rangeCount; i++) {
      var rg = s.getRangeAt(i);
      NATIVE_BOXES.forEach(function (b) {
        var box = $(b[0]);
        if (!box || !rg.intersectsNode(box) || out && out.where !== b[2]) return;
        var ks = [];
        [].forEach.call(box.querySelectorAll('[data-k]'), function (el) { if (rg.intersectsNode(el) && touches(rg, el)) ks.push(+el.dataset.k); });
        if (!ks.length) return;
        out = out || { where: b[2], ar: null, uz: null };
        out[b[1]] = (out[b[1]] || []).concat([[Math.min.apply(null, ks), Math.max.apply(null, ks)]]);
      });
    }
    if (out) { out.ar = C.ranges(out.ar); out.uz = C.ranges(out.uz); }
    return out;
  }
  var nativeAt = null;                                             // {pick, t}: taken at the press of a button
  function takeNative() {
    var x = nativeAt; nativeAt = null;
    return x && Date.now() - x.t < 5000 ? x.pick : nativePick();
  }
  function nativeToCard(p) {
    if (!p || p.where !== 'card' || !cur.h) return;
    sel.ar = p.ar; sel.uz = p.uz; clickStart = null;
    try { window.getSelection().removeAllRanges(); } catch (e) {}
    selChanged();
  }
  ['pointerdown', 'mousedown', 'touchstart'].forEach(function (ev) {
    document.addEventListener(ev, function (e) {
      var b = e.target && e.target.closest ? e.target.closest('#insert, #copy, #big-insert, #big-copy, #day-insert, #day-copy') : null;
      if (!b) return;
      var p = nativePick();
      if (p || !nativeAt || Date.now() - nativeAt.t > 1000) nativeAt = { pick: p, t: Date.now() };
    }, true);
  });
  document.addEventListener('pointercancel', function () {
    if (!drag) return;
    drag.box.classList.remove('dragging'); drag = null; selChanged();
  });

  /* ---------- «Кенгайтириш»: the hadith big in the middle of the screen ----------
     The same words as on the card (the same selection, the classes painted in both), a bigger font (settings.zoom),
     its own scroll; «Қўйиш» / «Нусха» from here. Closed by Esc, the cross or a click outside it. */
  function bigDraw() {
    var dlg = $('big');
    if (!dlg || !dlg.open) return;
    var r = cur.d >= 0 ? data.refOf(cur.d) : null;
    $('big-title').innerHTML = $('sel-title').innerHTML;
    if (!cur.h || !cur.P) { $('big-ar').innerHTML = ''; $('big-uz').innerHTML = '<p class="muted">' + esc(t('loading')) + '</p>'; $('big-tahrij').innerHTML = ''; return; }
    var by = $('big-body').scrollTop;
    $('big-bob').innerHTML = $('bob').innerHTML; $('big-bob').hidden = $('bob').hidden;
    $('big-ar').innerHTML = cur.P.ar.paras.length ? wordsHtml('ar') : '';
    $('big-uz').innerHTML = cur.P.uz.paras.length ? wordsHtml('uz') : '';
    $('big-tahrij').innerHTML = $('tahrij-box').innerHTML;
    dimTexts(); bigZoom();
    $('big-body').scrollTop = by;
    if (!r) $('big-title').textContent = '';
  }
  function bigZoom() {
    $('big').style.setProperty('--zoom', String(settings.zoom || 1));
    $('big-minus').disabled = settings.zoom <= 0.7; $('big-plus').disabled = settings.zoom >= 2;
  }
  function bigOpen() {
    if (cur.d < 0) return;
    var dlg = $('big');
    if (!dlg.open) dlg.showModal();
    $('big-body').scrollTop = 0;
    bigDraw(); selState();
    $('big-close').focus();
  }
  function bigShut() { var dlg = $('big'); if (dlg.open) dlg.close(); }
  $('big-btn').addEventListener('click', bigOpen);
  $('big-close').addEventListener('click', bigShut);
  // a click on the dark outside (the dialog itself, not its box); a drag that only ends there does not close it
  var bigDown = null;
  $('big').addEventListener('pointerdown', function (e) { bigDown = e.target; });
  $('big').addEventListener('click', function (e) { if (e.target === this && bigDown === this) bigShut(); bigDown = null; });
  $('big').addEventListener('close', function () { clickStart = null; drag = null; $('big-btn').focus(); });
  $('big').addEventListener('keydown', function (e) {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); bigInsert(); }
  });
  $('big-reset').addEventListener('click', clearPick);
  function bigInsert() { var p = doInsert(); bigShut(); return p; }   // the toast and «undo» are seen in the panel
  $('big-insert').addEventListener('click', bigInsert);
  $('big-copy').addEventListener('click', function () { $('copy').click(); });
  function zoomBy(dz) { settings.zoom = Math.max(0.7, Math.min(2, Math.round(((settings.zoom || 1) + dz) * 10) / 10)); saveSettings(); bigZoom(); }
  $('big-minus').addEventListener('click', function () { zoomBy(-0.1); });
  $('big-plus').addEventListener('click', function () { zoomBy(0.1); });

  /* ---------- smooth changes (settings.smooth; never with prefers-reduced-motion) ---------- */
  var REDUCE = window.matchMedia ? matchMedia('(prefers-reduced-motion: reduce)') : null;
  function motion() { return !!settings.smooth && !(REDUCE && REDUCE.matches); }
  function applyMotion() {
    document.documentElement.classList.toggle('smooth', motion());
    document.documentElement.classList.toggle('no-motion', !motion());
  }
  function scrollBeh() { return motion() ? 'smooth' : 'auto'; }
  if (REDUCE) try { REDUCE.addEventListener('change', applyMotion); } catch (e) {}
  applyMotion();

  /* ---------- the result: HadisCore.format with the settings ---------- */
  function optsFor(s, extra) {
    var o = {
      lang: uzLang(), script: latn() ? 'latn' : 'cyrl', mode: s.mode,
      withAr: settings.withAr && !muted.withAr && sideOut(s, 'ar'), withUz: settings.withUz && !muted.withUz && sideOut(s, 'uz'), order: settings.order,
      arRange: C.ranges(s.ar), uzRange: C.ranges(s.uz),
      refUz: settings.refUz && !muted.refUz, refAr: settings.refAr, withBook: settings.withBook, withHukm: settings.withHukm,
      withGrades: settings.withGrades, gradesInMatn: settings.gradesMatn, allGrades: settings.allGrades,
      withBob: settings.withBob, matnBold: settings.matnBold, arMatnBold: settings.arMatnBold
    };
    if (extra) for (var k in extra) o[k] = extra[k];
    return o;
  }
  function current() {
    if (!cur.h) return null;
    return fmt(cur.h, catalog.byId[cur.h.b], optsFor(sel));
  }
  function ready() { return !!(cur.h && cur.d === sel.d); }

  /* ---------- formatting of each part: settings.fmt[key] = {font, size, b, i, u, color[, align, line]} ---------- */
  var FMT_LABEL = { ar: 'arabicText', uz: 'uzbekText', ref: 'tahrij', hukm: 'hukm' };
  var FMT_PARA = { ar: 1, uz: 1 };
  var fmtKey = 'ar';
  function arabicFont() { return settings.fmt.ar.font || (HOST || online ? 'Scheherazade New' : inWord ? 'Sakkal Majalla' : 'Scheherazade New'); }
  function paraKind(p) { return p.dir === 'rtl' ? 'ar' : 'uz'; }
  /* the style of a run: the paragraph's format, then the format of its role; the bold of the data (matn) stays */
  function runStyle(r, kind) {
    var F = settings.fmt, base = F[kind], s = { b: !!base.b || !!r.b, i: !!base.i || !!r.i, u: !!base.u, color: base.color || '', font: '', size: 0 };
    var x = r.role === 'ref' ? F.ref : r.role === 'hukm' || r.role === 'grade' ? F.hukm : null;
    if (x) { s.b = !!x.b; s.i = !!x.i || r.role === 'ref' && !!r.i; s.u = !!x.u; if (x.color) s.color = x.color; s.font = x.font; s.size = +x.size || 0; }
    if (r.role === 'bob') s.b = true;
    return s;
  }
  function paraFont(kind) { return kind === 'ar' ? arabicFont() : settings.fmt.uz.font; }
  function paraSize(kind) { return +settings.fmt[kind].size || (kind === 'ar' ? 18 : 14); }
  /* the runs of a paragraph with their style resolved; neighbours of the same style are one run. The data cuts the matn
     into words (a word, then the space after it: two runs of the same look), up to 4 000 runs in one paragraph, ~330 bytes
     each in OOXML: 1.2 MB that Word lays out for minutes. Joined: the same letters in the same style, 10-15 times less. */
  var joinRuns = true;
  function styledRuns(runs, kind) {
    var list = [];
    runs.forEach(function (r) {
      var s = runStyle(r, kind), k = [s.b, s.i, s.u, s.color, s.font, s.size].join('\u0001'), last = list[list.length - 1];
      if (joinRuns && last && last.k === k) last.t += r.t; else list.push({ t: r.t, s: s, k: k });
    });
    return list;
  }
  function runsHtml(runs, kind, pt) {
    return styledRuns(runs, kind).map(function (r) {
      var s = r.s;
      var st = 'font-weight:' + (s.b ? 'bold' : 'normal') + ';font-style:' + (s.i ? 'italic' : 'normal') +
        (s.u ? ';text-decoration:underline' : '') + (s.color ? ';color:' + s.color : '');
      if (s.font) st += ";font-family:'" + esc(s.font) + "'";
      if (s.size && pt) st += ';font-size:' + s.size + 'pt';
      else if (s.size) st += ';font-size:' + (s.size / paraSize(kind)).toFixed(3) + 'em';
      return '<span style="' + st + '">' + esc(r.t) + '</span>';
    }).join('');
  }
  function fmtEditor() {
    var F = settings.fmt, f = F[fmtKey];
    $('f-keys').innerHTML = Object.keys(FMT_LABEL).map(function (k) {
      return '<button type="button" role="tab" data-k="' + k + '" aria-selected="' + (k === fmtKey) + '">' + esc(t(FMT_LABEL[k])) + '</button>';
    }).join('');
    $('f-font').value = f.font || '';
    $('f-font').placeholder = fmtKey === 'ar' ? arabicFont() : t('docFont');
    if (fmtKey === 'ar') $('f-font').setAttribute('list', 'ar-fonts'); else $('f-font').removeAttribute('list');
    $('f-size').value = +f.size || '';
    $('f-size').placeholder = fmtKey === 'ar' ? 18 : fmtKey === 'uz' ? 14 : (+F.uz.size || 14);
    [].forEach.call($('f-style').children, function (b) { b.setAttribute('aria-pressed', !!f[b.dataset.p]); });
    $('f-color').value = f.color || '#000000';
    $('f-color-show').style.background = f.color || '';
    $('f-color-show').classList.toggle('auto', !f.color);
    document.querySelector('#settings .para-only').hidden = !FMT_PARA[fmtKey];
    if (FMT_PARA[fmtKey]) {
      [].forEach.call($('f-align').children, function (b) { b.setAttribute('aria-pressed', (f.align || 'both') === b.dataset.a); });
      $('f-line').value = String(+f.line || 0);
    }
    fmtPreview();
  }
  function fmtPreview() {
    if (!$('settings').open) return;
    var out = current();
    if (!out) { $('f-preview').innerHTML = ''; return; }
    $('f-preview').innerHTML = out.paras.map(function (p) {
      var kind = paraKind(p), f = settings.fmt[kind], font = paraFont(kind);
      return '<p dir="' + p.dir + '" class="' + (kind === 'ar' ? 'arabic' : '') + '" style="text-align:' + (f.align === 'both' ? 'justify' : f.align) +
        (+f.line ? ';line-height:' + (kind === 'ar' ? 1.45 * f.line : f.line) : '') + (font ? ";font-family:'" + esc(font) + "'" + (kind === 'ar' ? ",'Hadis Arabic'" : '') : '') + '">' +
        runsHtml(p.runs, kind, false) + '</p>';
    }).join('');
  }
  function fmtSet(prop, value) { settings.fmt[fmtKey][prop] = value; saveSettings(); fmtEditor(); }
  $('f-keys').addEventListener('click', function (e) { var b = e.target.closest('[data-k]'); if (b) { fmtKey = b.dataset.k; fmtEditor(); } });
  $('f-style').addEventListener('click', function (e) { var b = e.target.closest('[data-p]'); if (b) fmtSet(b.dataset.p, !settings.fmt[fmtKey][b.dataset.p]); });
  $('f-align').addEventListener('click', function (e) { var b = e.target.closest('[data-a]'); if (b) fmtSet('align', b.dataset.a); });
  $('f-line').addEventListener('change', function () { fmtSet('line', +this.value); });
  $('f-font').addEventListener('change', function () { fmtSet('font', this.value.trim()); });
  $('f-size').addEventListener('change', function () { fmtSet('size', Math.max(0, Math.min(96, +this.value || 0))); });
  $('f-color').addEventListener('input', function () { fmtSet('color', this.value); });
  $('f-color-reset').addEventListener('click', function () { fmtSet('color', ''); });

  /* ---------- output: Word (OOXML), clipboard (HTML + text) ---------- */
  function xmlEsc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function runXml(r, font, size, rtl, s) {
    var f = font ? '<w:rFonts w:ascii="' + xmlEsc(font) + '" w:hAnsi="' + xmlEsc(font) + '" w:cs="' + xmlEsc(font) + '"/>' : '';
    var v = function (on) { return on ? '' : ' w:val="0"'; };
    var c = s.color ? '<w:color w:val="' + s.color.replace('#', '').toUpperCase() + '"/>' : '';
    var sz = size ? '<w:sz w:val="' + Math.round(size * 2) + '"/><w:szCs w:val="' + Math.round(size * 2) + '"/>' : '';
    return '<w:r><w:rPr>' + f + '<w:b' + v(s.b) + '/><w:bCs' + v(s.b) + '/><w:i' + v(s.i) + '/><w:iCs' + v(s.i) + '/>' +
      '<w:u w:val="' + (s.u ? 'single' : 'none') + '"/>' + c + sz + '<w:rtl' + v(rtl) + '/>' + (rtl ? '<w:lang w:bidi="ar-SA"/>' : '') + '</w:rPr>' +
      '<w:t xml:space="preserve">' + xmlEsc(r.t) + '</w:t></w:r>';
  }
  /* in a right-to-left paragraph Word reads "left" as the start of the line: left and right are swapped there */
  function pPrXml(f, rtl) {
    var jc = f.align || 'both';
    if (rtl && (jc === 'left' || jc === 'right')) jc = jc === 'left' ? 'right' : 'left';
    return '<w:pPr><w:bidi' + (rtl ? '' : ' w:val="0"') + '/>' +
      (+f.line ? '<w:spacing w:line="' + Math.round(240 * f.line) + '" w:lineRule="auto"/>' : '') + '<w:jc w:val="' + jc + '"/></w:pPr>';
  }
  /* a paragraph with every run's font, size and style resolved: {dir, kind, f, runs: [{t, font, size, s}]} */
  function paraModel(p) {
    var kind = paraKind(p), pf = settings.fmt[kind];
    return { dir: p.dir, kind: kind, f: pf, runs: styledRuns(p.runs, kind).map(function (r) {
      return { t: r.t, font: r.s.font || paraFont(kind), size: r.s.size || paraSize(kind), s: r.s };
    }) };
  }
  function paraXml(p) {
    var m = paraModel(p), rtl = m.dir === 'rtl';
    return '<w:p>' + pPrXml(m.f, rtl) + m.runs.map(function (x) { return runXml(x, x.font, x.size, rtl, x.s); }).join('') + '</w:p>';
  }
  function pictureXml(pic) {
    var cx = Math.round(pic.wPt * 12700), cy = Math.round(pic.hPt * 12700), alt = xmlEsc(pic.alt).replace(/"/g, '&quot;');
    return '<w:p>' + pPrXml({ align: 'center' }, true) + '<w:r><w:drawing>' +
      '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
      '<wp:docPr id="1" name="Hadis" descr="' + alt + '"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="hadis.png" descr="' + alt + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="rIdHadis"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  }
  /* pic: the Arabic paragraphs as one picture (the others stay text) */
  function buildOoxml(out, pic) {
    var body = '', picDone = false;
    out.paras.forEach(function (p) {
      if (pic && p.dir === 'rtl' && p.kind === 'ar') { if (!picDone) body += pictureXml(pic); picDone = true; return; }
      body += paraXml(p);
    });
    // Word merges the LAST inserted paragraph into the paragraph at the cursor: an empty last paragraph takes that role
    body += '<w:p/>';
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships></pkg:xmlData></pkg:part>' +
      (pic ? '<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rIdHadis" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/hadis.png"/>' +
        '</Relationships></pkg:xmlData></pkg:part>' +
        '<pkg:part pkg:name="/word/media/hadis.png" pkg:contentType="image/png" pkg:compression="store"><pkg:binaryData>' + pic.png + '</pkg:binaryData></pkg:part>' : '') +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
      ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>' + body +
      '</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>';
  }
  function buildHtml(out, noArabic) {
    return out.paras.filter(function (p) { return !(noArabic && p.dir === 'rtl'); }).map(function (p) {
      var kind = paraKind(p), f = settings.fmt[kind], font = paraFont(kind);
      return '<p dir="' + p.dir + '" style="text-align:' + (f.align === 'both' ? 'justify' : f.align || 'justify') + (+f.line ? ';line-height:' + f.line : '') +
        (font ? ";font-family:'" + esc(font) + "'" : '') + ';font-size:' + paraSize(kind) + 'pt">' + runsHtml(p.runs, kind, true) + '</p>';
    }).join('');
  }
  /* Inserting into Word. Word.run cannot be stopped: a guard of INSERT_MS without progress tells what happens instead of
     an endless wait; one insert at a time (the panel stays locked until Word ends the one it does, so a click more never
     inserts twice); the HTML way only where Word has no OOXML (an error after a part went in is never inserted again);
     a big text goes in parts of PART_XML by paragraphs, each with its own sync, without track() on all of it. */
  var INSERT_MS = 45000, BIG_TEXT = 20000, BIG_XML = 300000, PART_XML = 100000, inserting = false;
  var NO_OOXML = { NotImplemented: 1, ApiNotFound: 1, InvalidArgument: 1 };
  function uzError(k) { var e = new Error(t(k)); e.uz = true; return e; }
  function busy(on) {
    inserting = on;
    document.body.classList.toggle('inserting', on);
    ['insert', 'big-insert', 'bask-insert'].forEach(function (id) { var b = $(id); if (b) b.setAttribute('aria-busy', on); });
  }
  /* start(kick): the work, kick() at each step done; rejected with «insertSlow» when no step ends in INSERT_MS */
  function guarded(start) {
    return new Promise(function (ok, no) {
      var tm = null, over = false;
      function kick() { clearTimeout(tm); tm = setTimeout(function () { over = true; no(uzError('insertSlow')); }, INSERT_MS); }
      kick();
      Promise.resolve().then(function () { return start(kick); }).then(function (v) {
        clearTimeout(tm);
        if (over) toast(t('inserted')); else ok(v);           // Word ended after all: said now, never inserted again
      }, function (e) { clearTimeout(tm); if (!over) no(e); });
    });
  }
  /* "New paragraph": the text goes where the cursor is, in paragraphs of its own (the paragraph is split there first) */
  function insertWordWith(out, pic, opts) {
    opts = opts || {};
    if (inserting) return Promise.reject(uzError('insertBusy'));
    var xml = opts.html ? '' : buildOoxml(out, pic);
    var big = !opts.html && (out.text.length > BIG_TEXT || xml.length > BIG_XML);
    busy(true);
    return guarded(function (kick) {
      return (big ? insertParts(out, pic, kick) : insertOnce(xml, opts, kick)).catch(function (e) {
        if (e && e.partly || !(e && NO_OOXML[e.code])) throw e;    // not because of OOXML or a part is in: not again
        kick();
        return Word.run(function (ctx) {                       // hosts without OOXML support
          return insertTarget(ctx).then(function (at) {           // a paragraph mark put before the error is not put twice
            return keepInserted(ctx, at.target.insertHtml(opts.html || buildHtml(out), 'Replace'), false);
          });
        });
      }).then(function (v) { busy(false); return v; }, function (e) { busy(false); throw e; });
    });
  }
  /* the place of the insert: the selection, or after a new paragraph mark; safe: nothing replaced or split (see doUndo) */
  function insertTarget(ctx) {
    var s = ctx.document.getSelection(), before = null;
    s.load('text');
    if (settings.newPara) { before = s.paragraphs.getFirst().getRange('Start').expandTo(s.getRange('Start')); before.load('text'); }
    return ctx.sync().then(function () {
      var target = s, safe = !s.text;
      if (before && before.text.replace(/\s+/g, '') !== '') { target = s.insertText('\n', 'Replace').getRange('After'); safe = false; }
      return { target: target, safe: safe };
    });
  }
  function insertOnce(xml, opts, kick) {
    return Word.run(function (ctx) {
      return insertTarget(ctx).then(function (at) {
        kick();
        var rng = opts.html ? at.target.insertHtml(opts.html, 'Replace') : at.target.insertOoxml(xml, 'Replace');
        return keepInserted(ctx, rng, at.safe);
      });
    });
  }
  /* the paragraphs in parts of at most PART_XML (a paragraph bigger than that is a part alone); the Arabic as a picture
     is one unit with the paragraphs it stands for */
  function insertChunks(out, pic) {
    var parts = [], cur = null, picAt = -1;
    out.paras.forEach(function (p) {
      var isPic = pic && p.dir === 'rtl' && p.kind === 'ar';
      if (isPic && picAt >= 0) { parts[picAt].paras.push(p); return; }
      var n = isPic ? pic.png.length + 2000 : paraXml(p).length;
      if (!cur || cur.size && cur.size + n > PART_XML) { cur = { paras: [], size: 0, pic: false }; parts.push(cur); }
      cur.paras.push(p); cur.size += n;
      if (isPic) { cur.pic = true; picAt = parts.length - 1; }
    });
    return parts;
  }
  function insertParts(out, pic, kick) {
    var parts = insertChunks(out, pic), done = 0;
    return Word.run(function (ctx) {
      return insertTarget(ctx).then(function (at) {
        var place = at.target;
        function next() {
          if (done >= parts.length) { setUndo(true, null); return; }      // no track() on all of it: Ctrl+Z undoes
          kick();
          // each part ends with an empty paragraph that joins the text after the cursor: the end of the part is the
          // start of that text, where the next part goes
          var rng = place.insertOoxml(buildOoxml({ paras: parts[done].paras }, parts[done].pic ? pic : null), 'Replace');
          place = rng.getRange('End');
          return ctx.sync().then(function () {
            done++;
            return new Promise(function (ok) { setTimeout(ok, 30); });    // Word's window gets its turn between the parts
          }).then(next);
        }
        return next();
      });
    }).catch(function (e) { if (done && e) e.partly = true; throw e; });
  }
  function keepInserted(ctx, rng, safe) {
    return ctx.sync().then(function () {
      try { rng.track(); rng.load('text'); } catch (e) { setUndo(true, null); return; }
      return ctx.sync().then(function () { setUndo(true, safe ? { rng: rng, text: rng.text } : null); }, function () { setUndo(true, null); });
    });
  }
  function copy(out) { return copyRaw(buildHtml(out), out.text); }
  function copyRaw(html, text) {
    if (navigator.clipboard && window.ClipboardItem) {
      return navigator.clipboard.write([new ClipboardItem({
        'text/html': new Blob([html], { type: 'text/html' }),
        'text/plain': new Blob([text], { type: 'text/plain' })
      })]).catch(function () { return fallbackCopy(text); });
    }
    return fallbackCopy(text);
  }
  function fallbackCopy(text) {
    var ta = document.createElement('textarea'); ta.value = text; document.body.appendChild(ta);
    ta.select(); var ok = false; try { ok = document.execCommand('copy'); } catch (e) {} ta.remove();
    return ok ? Promise.resolve() : Promise.reject(new Error('clipboard'));
  }
  function doInsert() {
    if (!ready()) { toast(t('loading')); return; }          // the hadith is still loading: say so instead of doing nothing
    nativeToCard(takeNative());
    var out = current();
    if (!out.paras.length) { toast(t('nothingToInsert')); return; }
    remember();
    return insertOut(out);
  }
  /* the text into the document (Word), the host (sendToHost) or onto the clipboard */
  function insertOut(out) {
    if (HOST) return sendToHost(out);
    return (inWord ? insertWordWith(out, null) : copy(out)).then(function () {
      toast(inWord ? t('inserted') : t('copiedPaste'), inWord && undoReady ? undoAct() : null);
    }).catch(errToast);
  }
  /* an error of ours (e.uz) is said as it is, the others after «Хато» */
  function errToast(e) {
    var m = e && e.message || (typeof e === 'string' ? e : '');
    toast(e && e.uz ? m : t('error') + (m ? ': ' + m : ''));
  }
  function undoAct() { return { label: t('undoShort'), fn: doUndo }; }
  $('insert').addEventListener('click', doInsert);
  $('copy').addEventListener('click', function () {
    if (!ready()) return;
    nativeToCard(takeNative());
    remember();
    copy(current()).then(function () { toast(t('copied')); }, function (e) { toast(t('error') + ': ' + e.message); });
  });

  /* Embedding hosts get the result as a message (the protocol of Khatt al-Quran, so the same host code serves both):
     html + text (web pages, editors) and doc, the paragraphs with resolved fonts and styles (Google Docs builds them).
     It is also put on the clipboard first, while the click still counts as the user's: where the host cannot insert, Ctrl+V. */
  var hostWait = null, hostCopied = null;
  function hostDoc(out) {
    return out.paras.map(paraModel).map(function (m) {
      return { dir: m.dir, kind: m.kind, align: m.f.align || 'both', line: +m.f.line || 0, runs: m.runs.map(function (x) {
        return { t: x.t, font: x.font || '', size: x.size, b: !!x.s.b, i: !!x.s.i, u: !!x.s.u, color: x.s.color || '' };
      }) };
    });
  }
  function sendToHost(out) {
    var copied = hostCopied = copy(out).then(function () { return true; }, function () { return false; });
    if (HOST === 'tg') return copied.then(function (ok) { toast(ok ? t('copiedPaste') : t('error')); });
    if (HOST === 'ext') return sendToExt(out, copied);
    window.parent.postMessage({ type: 'khatt-insert', v: 1, app: 'hadis', html: buildHtml(out), text: out.text, doc: hostDoc(out),
                                newPara: !!settings.newPara }, '*');
    clearTimeout(hostWait);
    hostWait = setTimeout(function () { copied.then(function (ok) { toast(ok ? t('copiedPaste') : t('error')); }); }, 4000);
  }
  var extLast = null;
  function sendToExt(out, copied, mode) {
    mode = mode || settings.extInsert || 'ask'; extLast = out;
    var msg = { type: 'khatt-insert', v: 1, app: 'hadis', html: buildHtml(out), text: out.text, doc: hostDoc(out),
                newPara: !!settings.newPara, mode: mode };
    clearTimeout(hostWait);
    return (mode === 'picture' ? arabicPicture(out).catch(function () { return null; }) : Promise.resolve(null)).then(function (pic) {
      if (pic) {
        msg.pic = 'data:image/png;base64,' + pic.png;
        msg.picHtml = '<p dir="rtl" style="text-align:center"><img src="' + msg.pic + '" alt="' + esc(pic.alt) + '" width="' +
          Math.round(pic.wPt * 4 / 3) + '" style="max-width:100%;height:auto"></p>';
        msg.restHtml = buildHtml(out, true);
      }
      toParent(msg);
      hostWait = setTimeout(function () { copied.then(function (ok) { toast(ok ? t('copiedPaste') : t('error')); }); }, 4000);
    });
  }
  window.addEventListener('message', function (e) {
    var d = e.data;
    if (!HOST || e.source !== window.parent || !d || !fromExt(e)) return;
    if (d.type === 'khatt-result') {
      clearTimeout(hostWait);
      if (d.undo) {
        clearTimeout(undoWait);
        if (d.ok) { setUndo(false); toast(t('undoDone')); }
        else if (d.error === 'gone') { setUndo(false); toast(t('undoGone')); }
        else toast(t(d.error === 'changed' ? 'undoChanged' : 'undoManual'));
        return;
      }
      if (d.ok && !d.ask) setUndo(true, null);
      if (HOST === 'ext' && d.ask) { sendToExt(extLast, hostCopied || Promise.resolve(false), 'text'); return; }
      if (d.ok) toast(t('inserted'), HOST === 'wp' || HOST === 'gdocs' ? undoAct() : null);
      else (hostCopied || Promise.resolve(false)).then(function (ok) { toast(ok ? t('copiedPaste') : t('error') + (d.error ? ': ' + d.error : '')); });
    } else if (d.type === 'khatt-search' && typeof d.q === 'string') searchFor(d.q);
    else if (d.type === 'khatt-pref') {
      if (HOST === 'ext' && /^(ask|text|picture)$/.test(d.extInsert || '')) { settings.extInsert = d.extInsert; saveSettings(); }
      // the alphabet kept by the host (the storage of the extension / of Google Docs): it wins over a lost localStorage
      if (d.apos) setApos(String(d.apos), true);
      if (d.script) setScript(String(d.script), true);
      // «Кун ҳадиси»: the host's copy is newer (this frame lost its storage): taken and kept here too; older: renewed
      if (d.dayAt != null && takeDayPref(settings, d)) {
        sanitize(settings); saveSettings(); savePref(); if ($('s-day')) $('s-day').checked = settings.day; drawDay();
      } else if (d.dayAt != null && (+d.dayAt || 0) < (settings.dayAt || 0)) savePref();
    }
  });
  if (HOST === 'ext') toParent({ type: 'khatt-ready' });
  function searchFor(q) {
    q = String(q || '').replace(/\s+/g, ' ').trim().slice(0, 200);
    if (!q) return;
    (function wait(n) {
      if (!catalog) { if (n < 300) setTimeout(function () { wait(n + 1); }, 200); return; }
      $('q').value = q; clearTimeout(searchTimer); runSearch(true);
    })(0);
  }
  /* text dragged onto the panel is searched */
  var dropTimer;
  function dropOff() { document.body.style.boxShadow = ''; }
  function dropTarget(e) {
    var dt = e.dataTransfer, el = e.target && e.target.closest ? e.target.closest('input, textarea, select, [contenteditable]') : null;
    return dt && [].indexOf.call(dt.types || [], 'text/plain') >= 0 && (!el || el.id === 'q');
  }
  window.addEventListener('dragover', function (e) {
    if (!dropTarget(e)) return;
    e.preventDefault(); e.dataTransfer.dropEffect = 'copy';
    document.body.style.boxShadow = 'inset 0 0 0 3px var(--accent)';
    clearTimeout(dropTimer); dropTimer = setTimeout(dropOff, 250);
  });
  window.addEventListener('drop', function (e) { if (!dropTarget(e)) return; e.preventDefault(); dropOff(); searchFor(e.dataTransfer.getData('text/plain')); });

  /* ---------- the Arabic as a picture (transparent unless a background is set) ---------- */
  /* PIC_MAX_H: the canvas of WebKit (Word for Mac) and of old Edge gives an empty picture past ~16 000 px: 3 000 pt is
     about 6 pages of Arabic, longer goes as text; PIC_FONT_MS: a font that does not come is not waited for */
  var PIC_WIDTH_PT = 450, PIC_SCALE = 4, PIC_MAX_H = 12000, PIC_FONT_MS = 3000;
  function arabicPicture(out) {
    var ps = out.paras.filter(function (p) { return p.dir === 'rtl'; });
    if (!ps.length) return Promise.reject(new Error(t('noArabic')));
    var toks = [], fontsToLoad = {}, F = settings.fmt.ar, family = "'" + arabicFont() + "','Hadis Arabic','Scheherazade New',serif";
    ps.forEach(function (p, pi) {
      if (pi > 0) toks.push({ br: true });
      p.runs.forEach(function (r) {
        var s = runStyle(r, 'ar'), px = (s.size || paraSize('ar')) * PIC_SCALE;
        var font = (s.i ? 'italic ' : '') + (s.b ? 'bold ' : '') + px + 'px ' + (s.font ? "'" + s.font + "'," : '') + family;
        fontsToLoad[font] = r.t.replace(/\s/g, '').slice(0, 8) || 'ا';
        r.t.split(/ +/).forEach(function (w, i) {
          if (i > 0 || (/^ /.test(r.t) && toks.length)) toks.push({ space: true, font: font, px: px });
          if (w) toks.push({ t: w, font: font, px: px, color: s.color || '#000000', u: s.u });
        });
      });
    });
    return Promise.all(Object.keys(fontsToLoad).map(function (f) {
      if (!document.fonts || !document.fonts.load) return null;
      return new Promise(function (ok) {
        setTimeout(ok, PIC_FONT_MS);
        document.fonts.load(f, fontsToLoad[f]).then(ok, ok);
      });
    })).then(function () {
      var W = PIC_WIDTH_PT * PIC_SCALE, cv = document.createElement('canvas'), ctx = cv.getContext('2d');
      var base = paraSize('ar') * PIC_SCALE, lineH = Math.round(base * (+F.line ? 1.45 * F.line : 2.1));
      toks.forEach(function (k) { if (k.br) return; ctx.font = k.font; k.w = ctx.measureText(k.space ? ' ' : k.t).width; });
      var lines = [], line = [], width = 0;
      function flush(last) { while (line.length && line[line.length - 1].space) width -= line.pop().w; lines.push({ toks: line, w: width, last: last }); line = []; width = 0; }
      toks.forEach(function (k) {
        if (k.br) { flush(true); return; }
        if (k.space) { if (line.length) { line.push(k); width += k.w; } return; }
        if (width + k.w > W && line.length) flush(false);
        line.push(k); width += k.w;
      });
      if (line.length) flush(true);
      if (lines.length * lineH > PIC_MAX_H) throw uzError('picTooLong');
      cv.width = W; cv.height = Math.max(1, lines.length) * lineH;
      if (settings.picBgOn) { ctx.fillStyle = settings.picBg; ctx.fillRect(0, 0, W, cv.height); }
      ctx.textBaseline = 'alphabetic'; ctx.direction = 'rtl'; ctx.textAlign = 'right';
      var align = F.align || 'both';
      lines.forEach(function (ln, i) {
        var gaps = ln.toks.filter(function (k) { return k.space; }).length;
        var extra = align === 'both' && !ln.last && gaps ? (W - ln.w) / gaps : 0, y = i * lineH + lineH * 0.7;
        var x = align === 'left' ? ln.w : align === 'center' ? (W + ln.w) / 2 : W;
        ln.toks.forEach(function (k) {
          if (k.space) { x -= k.w + extra; return; }
          ctx.font = k.font; ctx.fillStyle = k.color; ctx.fillText(k.t, x, y);
          if (k.u) ctx.fillRect(x - k.w, y + k.px * 0.18, k.w, Math.max(2, k.px * 0.05));
          x -= k.w;
        });
      });
      var png = cv.toDataURL('image/png').split(',')[1];
      if (!png) throw uzError('picTooLong');                   // the canvas too big for this browser: an empty picture
      return { png: png, wPt: PIC_WIDTH_PT, hPt: cv.height / PIC_SCALE,
               alt: ps.map(function (p) { return p.runs.map(function (r) { return r.t; }).join(''); }).join(' ') };
    });
  }
  function pictureBlob(out) {
    return arabicPicture(out).then(function (pic) {
      var bin = atob(pic.png), buf = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
      return new Blob([buf], { type: 'image/png' });
    });
  }
  function savePicture(blob) {
    var a = document.createElement('a'), url = URL.createObjectURL(blob), r = data.refOf(cur.d);
    a.href = url; a.download = 'hadis-' + (r ? r.book + '-' + C.refLabel(r).replace('/', '-') : cur.d) + '.png';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 10000);
    toast(t('pictureSaved'));
  }
  function copyPicture() {
    if (!ready()) return;
    remember();
    var blob = pictureBlob(current()), write;
    try {
      if (!navigator.clipboard || !navigator.clipboard.write || !window.ClipboardItem ||
          (ClipboardItem.supports && !ClipboardItem.supports('image/png'))) throw new Error('no image clipboard');
      write = navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
    } catch (e) { write = Promise.reject(e); }
    write.then(function () { toast(t('pictureCopied')); }, function () { return blob.then(savePicture); })
      .catch(function (e) { toast(t('error') + (e && e.message ? ': ' + e.message : '')); });
  }
  function insertPicture() {
    if (!ready()) return;
    remember();
    var out = current();
    arabicPicture(out).then(function (pic) { return insertWordWith(out, pic); })
      .then(function () { toast(t('inserted'), undoReady ? undoAct() : null); }, errToast);
  }
  $('copy-pic').addEventListener('click', function () {
    if (!ready()) return;
    if (!inWord || HOST || settings.picAction === 'copy') return copyPicture();
    if (settings.picAction === 'insert') return insertPicture();
    $('pic-remember').checked = false; $('pic-ask').returnValue = ''; $('pic-ask').showModal();
  });
  $('pic-ask').addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('button[value=insert], button[value=copy]');
    if (!b) return;
    e.preventDefault();
    var how = b.value;
    if ($('pic-remember').checked) { settings.picAction = how; saveSettings(); $('s-pic-action').value = how; }
    if (how === 'insert') insertPicture(); else copyPicture();
    $('pic-ask').close(how);
  });

  /* ---------- history, favourites, basket ----------
     A mark: {d, m: mode, a: [from, to] (Arabic words), u: [from, to] (Uzbek words), t: the start of the text, time}.
     Kept in localStorage (the same limits as Khatt al-Quran). */
  var MARK_KEY = { hist: 'hadis-history', fav: 'hadis-favorites', bask: 'hadis-basket' }, MARK_MAX = { hist: 50, fav: 500, bask: 100 }, marksOpen = '';
  var CLOSE_SVG = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6 6 18"/></svg>';
  function okPair(r) { return Array.isArray(r) && r.length === 2 && typeof r[0] === 'number' && r[0] >= 0 && (r[1] == null || r[1] >= r[0]); }
  function okRange(r) { return r == null || okPair(r) || Array.isArray(r) && r.length > 0 && r.length <= 50 && r.every(okPair); }
  /* the parts of a mark: one as [from, to] (the old form), several as [[from, to], …] */
  function packParts(p) { p = C.ranges(p); return !p ? null : p.length === 1 ? p[0] : p; }
  function okMark(x) { return !!x && typeof x.d === 'number' && x.d >= 0 && x.d < 200000 && okRange(x.a) && okRange(x.u); }
  function loadMarks(k) {
    try { var x = JSON.parse(localStorage.getItem(MARK_KEY[k]) || '[]'); return Array.isArray(x) ? x.filter(okMark).slice(0, MARK_MAX[k]) : []; }
    catch (e) { return []; }
  }
  var marks = { hist: loadMarks('hist'), fav: loadMarks('fav'), bask: loadMarks('bask') };
  function markId(x) { return x.d + ':' + (x.m || 'full') + ':' + (x.a ? x.a.join('-') : '') + ':' + (x.u ? x.u.join('-') : ''); }
  function storeMarks(k) { try { localStorage.setItem(MARK_KEY[k], JSON.stringify(marks[k])); } catch (e) {} }
  function saveMarks(k) { storeMarks(k); drawMarks(); starState(); dayStar(); }
  function addMark(k, x) {
    var id = markId(x);
    marks[k] = [x].concat(marks[k].filter(function (y) { return markId(y) !== id; })).slice(0, MARK_MAX[k]);
    saveMarks(k);
  }
  function selMark() {
    var x = { d: sel.d, time: Date.now() };
    if (sel.mode !== 'full') x.m = sel.mode;                     // none: the whole isnad (the marks before 2026-10-31)
    if (sel.ar) x.a = packParts(sel.ar);
    if (sel.uz) x.u = packParts(sel.uz);
    var w = cur.P ? C.words(cur.P.ar.paras, modeOf('ar')) : [], pa = C.ranges(x.a);
    if (pa) w = [].concat.apply([], pa.map(function (p) { return w.slice(p[0], p[1] + 1); }));
    x.t = w.slice(0, 12).map(function (y) { return y.text; }).join(' ').slice(0, 120);
    return x;
  }
  function remember() { if (ready()) addMark('hist', selMark()); }
  function isFav() { var id = markId(selMark()); return marks.fav.some(function (y) { return markId(y) === id; }); }
  function starState() {
    var on = cur.d >= 0 && isFav(), b = $('fav-toggle');
    b.setAttribute('aria-pressed', on); b.title = t(on ? 'favRemove' : 'favAdd'); b.setAttribute('aria-label', b.title);
  }
  $('fav-toggle').addEventListener('click', function () {
    if (!ready()) return;
    var x = selMark(), id = markId(x);
    if (isFav()) { marks.fav = marks.fav.filter(function (y) { return markId(y) !== id; }); saveMarks('fav'); }
    else addMark('fav', x);
  });
  function markTime(ms) {
    if (!ms) return '';
    var d = new Date(ms), today = d.toDateString() === new Date().toDateString();
    var o = today ? { hour: '2-digit', minute: '2-digit' } : { day: 'numeric', month: 'short' };
    try { return d.toLocaleString(i18n.code(), o); } catch (e) { return d.toLocaleString(undefined, o); }
  }
  function markLabel(x) {
    return refText(x.d) + (x.m === 'matn' ? ' · ' + t('modeMatn') : '') + (x.a || x.u ? ' · ' + t('part') : '');
  }
  function drawMarks() {
    $('hist-n').textContent = marks.hist.length || ''; $('fav-n').textContent = marks.fav.length || ''; $('bask-n').textContent = marks.bask.length || '';
    [['hist', 'history'], ['fav', 'favorites'], ['bask', 'basket']].forEach(function (x) {
      var b = $(x[0] + '-btn'), n = marks[x[0]].length;
      b.title = t(x[1]) + (n ? ' (' + n + ')' : ''); b.setAttribute('aria-label', b.title);
      b.setAttribute('aria-expanded', marksOpen === x[0]);
    });
    $('bask-acts').hidden = marksOpen !== 'bask' || !marks.bask.length;
    $('marks-panel').hidden = !marksOpen;
    if (!marksOpen) return;
    var list = marks[marksOpen];
    $('marks-list').innerHTML = list.map(function (x, i) {
      return '<li data-i="' + i + '"><button type="button" class="mk-go"><span class="mk-ref">' + (marksOpen === 'bask' ? (i + 1) + '. ' : '') + esc(markLabel(x)) +
        (marksOpen === 'hist' ? ' <span class="mk-time">' + esc(markTime(x.time)) + '</span>' : '') +
        '</span><span class="mk-ar" dir="rtl">' + esc(x.t || '') + '</span></button>' +
        '<button type="button" class="icon-btn mk-del" title="' + esc(t('remove')) + '" aria-label="' + esc(t('remove') + ': ' + markLabel(x)) + '">' + CLOSE_SVG + '</button></li>';
    }).join('');
    $('marks-empty').textContent = list.length ? '' : t(marksOpen === 'hist' ? 'histEmpty' : marksOpen === 'fav' ? 'favEmpty' : 'basketEmpty');
    $('marks-clear').hidden = !list.length;
  }
  ['hist', 'fav', 'bask'].forEach(function (k) {
    $(k + '-btn').addEventListener('click', function () {
      marksOpen = marksOpen === k ? '' : k; drawMarks();
      if (marksOpen) $('marks-panel').scrollIntoView({ block: 'nearest', behavior: scrollBeh() });
    });
  });
  function pickMark(x) { go(x.d, { mode: x.m || 'full', ar: x.a || null, uz: x.u || null }); }
  $('marks-list').addEventListener('click', function (e) {
    var li = e.target.closest('li'), k = marksOpen; if (!li || !k) return;
    var i = +li.dataset.i, x = marks[k][i]; if (!x) return;
    if (!e.target.closest('.mk-del')) { pickMark(x); return; }
    marks[k].splice(i, 1); saveMarks(k);
    var next = $('marks-list').querySelectorAll('.mk-del')[Math.min(i, marks[k].length - 1)];
    (next || $(k + '-btn')).focus();
  });
  $('marks-clear').addEventListener('click', function () {
    if (!marksOpen) return;
    marks[marksOpen] = []; saveMarks(marksOpen); $(marksOpen + '-btn').focus();
  });
  /* the basket: hadiths collected in order, inserted (or copied) together; the sources as one list at the end (or each
     hadith with its own tahrij, settings.baskSrc) */
  $('m-bask').addEventListener('click', function () {
    menuOpen(false); if (!ready()) return;
    var x = selMark(), id = markId(x);
    marks.bask = marks.bask.filter(function (y) { return markId(y) !== id; }).concat([x]).slice(-MARK_MAX.bask);
    saveMarks('bask'); toast(t('basketAdded'));
  });
  function basketOut() {
    var list = marks.bask.slice(), lst = settings.baskSrc === 'list';
    return Promise.all(list.map(function (x) { return hadithView(x.d); })).then(function (hs) {
      var paras = [], text = [], refs = [];
      hs.forEach(function (h, i) {
        var x = list[i], b = catalog.byId[h.b];
        var o = fmt(h, b, optsFor({ mode: x.m || 'full', ar: x.a || null, uz: x.u || null }, lst ? { refUz: false, refAr: false, mark: i + 1 } : null));
        paras = paras.concat(o.paras); text.push(o.text);
        // the same tahrij as after a hadith (web/core.js tahrij), without the brackets
        refs.push(['[' + (i + 1) + '] ', C.tahrij(h, b, uzLang(), { paren: false, withBook: settings.withBook, hukm: settings.withHukm,
          grades: settings.withGrades && (x.m !== 'matn' || settings.gradesMatn), all: settings.allGrades })]);
      });
      if (lst && refs.length) {
        paras.push({ dir: 'ltr', kind: 'tahrij', runs: [{ t: t('sources') + ':', b: true, i: false, role: '' }] });
        refs.forEach(function (r) {
          paras.push({ dir: 'ltr', kind: 'tahrij', runs: [{ t: r[0], b: false, i: false, role: '' }, { t: r[1], b: false, i: true, role: 'ref' }, { t: '.', b: false, i: false, role: '' }] });
        });
        text.push(t('sources') + ':\n' + refs.map(function (r) { return r[0] + r[1] + '.'; }).join('\n'));
      }
      return { paras: paras, text: text.join('\n\n') };
    });
  }
  function basketGo(copyOnly) {
    if (!catalog || !marks.bask.length) return;
    basketOut().then(function (out) {
      if (!copyOnly && inWord && !HOST) return insertWordWith(out, null).then(function () { toast(t('inserted'), undoReady ? undoAct() : null); });
      if (!copyOnly && HOST) return sendToHost(out);
      return copy(out).then(function () { toast(t('copied')); });
    }).catch(errToast);
  }
  $('bask-insert').addEventListener('click', function () { basketGo(false); });
  $('bask-copy').addEventListener('click', function () { basketGo(true); });

  /* Undo the last insert, as careful as in Khatt al-Quran: only when nothing was replaced or split and the inserted text
     is still exactly as it was put in. */
  var undoReady = false, undoRange = null, undoWait = null;
  function setUndo(on, word) {
    var old = undoRange;
    if (old && old !== word && old.rng && window.Word) {
      try { Word.run(old.rng, function (ctx) { old.rng.untrack(); return ctx.sync(); }).catch(function () {}); } catch (e) {}
    }
    undoReady = on; undoRange = on ? word || null : null; $('again').disabled = !on;
  }
  function doUndo() {
    if (!undoReady) return;
    if (HOST) {
      if (HOST !== 'wp' && HOST !== 'gdocs') { toast(t('undoManual')); return; }
      clearTimeout(undoWait); undoWait = setTimeout(function () { toast(t('undoManual')); }, 6000);
      window.parent.postMessage({ type: 'khatt-undo', v: 1 }, '*');
      return;
    }
    var u = undoRange;
    if (!inWord || !u || !window.Word) { toast(t('undoManual')); return; }
    Word.run(u.rng, function (ctx) {
      u.rng.load('text');
      return ctx.sync().then(function () {
        var now = u.rng.text;
        if (!String(now).replace(/\s+/g, '')) return 'gone';
        if (now !== u.text) return 'changed';
        u.rng.delete();
        return ctx.sync().then(function () { try { u.rng.untrack(); } catch (e) {} return 'ok'; });
      });
    }).then(function (r) {
      if (r === 'ok') { setUndo(false); toast(t('undoDone')); }
      else if (r === 'gone') { setUndo(false); toast(t('undoGone')); }
      else toast(t('undoChanged'));
    }, function () { toast(t('undoManual')); });
  }
  $('again').addEventListener('click', doUndo);

  /* ---------- "More": plain text, Markdown, a link, the basket, the book, a footnote (Word) ---------- */
  function menuOpen(on) {
    var m = $('more-menu'); m.hidden = !on; $('more-btn').setAttribute('aria-expanded', !!on);
    if (on) $('m-foot').hidden = !(inWord && window.Office && Office.context && Office.context.requirements && Office.context.requirements.isSetSupported('WordApi', '1.5'));
  }
  $('more-btn').addEventListener('click', function (e) { e.stopPropagation(); menuOpen($('more-menu').hidden); });
  document.addEventListener('click', function (e) { if (!$('more-menu').hidden && !e.target.closest('#more-menu')) menuOpen(false); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && !$('more-menu').hidden) { menuOpen(false); $('more-btn').focus(); } });
  function clipText(txt) {
    var done = function () { toast(t('copied')); };
    if (navigator.clipboard && navigator.clipboard.writeText) return navigator.clipboard.writeText(txt).then(done, function () { fallbackCopy(txt).then(done, function () { toast(t('error')); }); });
    fallbackCopy(txt).then(done, function () { toast(t('error')); });
  }
  function mdOf(out) {
    return String(out.text).split(/\n+/).filter(function (p) { return p.trim(); }).map(function (p) { return '> ' + p.trim(); }).join('\n>\n');
  }
  function refLink() { return location.origin + location.pathname + location.hash; }
  $('m-text').addEventListener('click', function () { menuOpen(false); if (ready()) { remember(); clipText(current().text); } });
  $('m-md').addEventListener('click', function () { menuOpen(false); if (ready()) { remember(); clipText(mdOf(current())); } });
  $('m-link').addEventListener('click', function () { menuOpen(false); if (cur.d >= 0) { setHash(cur.d); clipText(refLink()); } });
  $('m-foot').addEventListener('click', function () {
    menuOpen(false); if (!ready()) return;
    var ref = C.tahrij(cur.h, catalog.byId[cur.h.b], uzLang(), { paren: false, withBook: settings.withBook });
    Word.run(function (ctx) { ctx.document.getSelection().insertFootnote(ref); return ctx.sync(); })
      .then(function () { toast(t('inserted')); }, function () { toast(t('footnoteNo')); });
  });
  /* about the book: from books.json (the edition of the numbering: not chosen yet, see ПЛАН.md §5) */
  function bookInfo() {
    var id = cur.h ? cur.h.b : $('book').value, b = catalog.byId[id]; if (!b) return;
    $('info-title').textContent = titleName(id);
    var L = nameLang(), au = L === 'uz' ? b.author.uz : L === 'uz_latn' ? b.author.uzl : C.bookName(id, L, 'author');
    var rows = [[t('bookAuthor'), au + (b.author.ar ? ' · ' + b.author.ar : '')], [t('arTitle'), b.name.ar || '—'],
                [t('hadithCount'), String(b.count)], [t('numbersRange'), b.numbers[0] + '–' + b.numbers[1]],
                [t('edition'), b.edition ? (typeof b.edition === 'string' ? b.edition : b.edition.uz) : t('notGiven')]];
    if (b.label) rows.push([t('labelTitle'), uzOrUi(b.label)]);
    $('info-body').innerHTML = rows.map(function (r) { return '<dt>' + esc(r[0]) + '</dt><dd>' + esc(r[1]) + '</dd>'; }).join('') +
      '<p class="note2">' + esc(C.NO_HUKM[id] ? t('sahihNote') : t('hukmNote')) + '</p>' +
      (C.NO_HUKM[id] ? '' : '<p class="note2">' + esc(t('gradesHint')) + '</p>') +
      '<p class="note2"><a href="sources.html#' + esc(id) + '" target="_blank" rel="noopener">' + esc(t('sourceLink')) + '</a></p>';
    $('info').showModal();
  }
  $('info-btn').addEventListener('click', bookInfo);
  $('m-info').addEventListener('click', function () { menuOpen(false); bookInfo(); });

  /* ---------- «Check a quote»: a pasted text -> the hadith of the data it comes from, word by word (HadisCore.quote,
     findQuote, compareQuote: the letters, the vowel marks apart). «Fix» puts the words of the data in its place (Word:
     over the selection, elsewhere onto the clipboard), with the tahrij of the settings. Nothing but the data is shown. */
  var quoteSeq = 0, quoteState = null;              // {q, list: [{d, h, r}], k: the one shown}
  function quoteVerdict(text, cls) { return '<div class="verdict' + (cls ? ' ' + cls : '') + '" dir="' + i18n.dir() + '">' + esc(text) + '</div>'; }
  function quoteOpen() {
    menuOpen(false);
    var q = $('q').value.trim();
    if (!$('quote-in').value.trim() && q.split(/\s+/).length >= 3 && !C.parseRef(q)) $('quote-in').value = q;
    $('quote-doc').hidden = !inWord;
    $('quote-dlg').showModal(); $('quote-in').focus();
  }
  function quoteRun() {
    var q = C.quote($('quote-in').value), out = $('quote-out'), seq = ++quoteSeq;
    quoteState = null;
    if (q.words < 3) { out.innerHTML = quoteVerdict(t('quoteShort')); return; }
    if (!search || !catalog) { out.innerHTML = quoteVerdict(t('loading')); return; }
    out.innerHTML = quoteVerdict(t('quoteSearching'), 'muted');
    $('quote-run').disabled = true;
    C.findQuote(q, search, function (d, signal) { return data.hadith(d, signal); }).then(function (list) {
      if (seq !== quoteSeq) return;
      quoteState = { q: q, list: list, k: 0 };
      quoteDraw();
    }, function (e) {
      if (seq === quoteSeq) out.innerHTML = quoteVerdict(t('loadFailed') + ' ' + (e && e.message || ''), 'bad');
    }).then(function () { if (seq === quoteSeq) $('quote-run').disabled = false; });
  }
  function quoteOpsHtml(r) {
    var parts = [], skip = 0;
    function flush() { if (skip) { parts.push('<span class="q-skip" title="' + esc(t('quoteLegSkip')) + ': ' + skip + '">…</span>'); skip = 0; } }
    r.ops.forEach(function (o) {
      if (o.op === 'skip') { skip++; return; }
      flush();
      var a = esc(o.a || ''), q = esc(o.q || '');
      if (o.op === 'same') parts.push(a);
      else if (o.op === 'mark' || o.op === 'spell') parts.push('<mark class="q-mark" title="' + esc(t('quoteYours', { w: o.q })) + '">' + a + '</mark>');
      else if (o.op === 'typo' || o.op === 'sub') parts.push('<del class="q-bad">' + q + '</del> <ins class="q-ok">' + a + '</ins>');
      else if (o.op === 'miss') parts.push('<ins class="q-miss" title="' + esc(t('quoteLegMiss')) + '">' + a + '</ins>');
      else if (o.op === 'extra') parts.push('<del class="q-extra" title="' + esc(t('quoteLegExtra')) + '">' + q + '</del>');
    });
    flush();
    return (r.from > 0 ? '<span class="q-ctx">… </span>' : '') + parts.join(' ') + (r.to < r.n - 1 ? '<span class="q-ctx"> …</span>' : '');
  }
  function quoteDraw() {
    var st = quoteState, out = $('quote-out');
    if (!st) return;
    var x = st.list[st.k], r = x && x.r;
    if (!r || !r.found) { out.innerHTML = quoteVerdict(t('quoteNotFound'), 'bad'); return; }
    var uz = r.side === 'uz', soft = uz ? r.spell : r.mark, v;
    if (r.exact) v = t('quoteSame');
    else if (!r.diff) v = t(uz ? 'quoteSpellOnly' : 'quoteMarksOnly', { n: soft });
    else v = t('quoteDiff', { n: r.diff }) + (soft ? ' ' + t(uz ? 'quoteSpellToo' : 'quoteMarksToo', { n: soft }) : '');
    var notes = [];
    if (r.part) notes.push(t('quotePart', { a: r.from + 1, b: r.to + 1, n: r.n }));
    if (!uz && !r.marksChecked) notes.push(t('quoteNoMarks'));
    var leg = [['q-mark', uz ? 'quoteLegSpell' : 'quoteLegMark', soft], ['q-bad', 'quoteLegBad', r.typo + r.sub],
               ['q-miss', 'quoteLegMiss', r.miss], ['q-extra', 'quoteLegExtra', r.extra], ['q-skip', 'quoteLegSkip', r.skip]].filter(function (l) { return l[2]; });
    var others = st.list.map(function (y, i) { return { y: y, i: i }; }).filter(function (o) { return o.i !== st.k && o.y.r.found; }).slice(0, 4);
    out.innerHTML = quoteVerdict(v, r.exact ? 'ok' : r.diff ? 'bad' : 'warn') +
      notes.map(function (s) { return '<p class="small muted" dir="' + i18n.dir() + '">' + esc(s) + '</p>'; }).join('') +
      '<p class="q-ref" dir="' + i18n.dir() + '"><a href="#" id="quote-open" data-i18n-title="quoteOpen" title="' + esc(t('quoteOpen')) + '">' +
        esc(C.tahrij(x.h, catalog.byId[x.h.b], uzLang(), { withBook: settings.withBook })) + '</a></p>' +
      '<div class="q-text ' + (uz ? 'uz' : 'ar') + '" dir="' + (uz ? 'ltr' : 'rtl') + '"' + (uz ? '' : ' lang="ar"') + '>' + quoteOpsHtml(r) + '</div>' +
      (leg.length ? '<p class="q-legend small" dir="' + i18n.dir() + '">' + leg.map(function (l) { return '<bdi class="' + l[0] + '">' + esc(t(l[1])) + '</bdi>'; }).join(' ') + '</p>' : '') +
      '<p class="actions" dir="' + i18n.dir() + '"><button type="button" id="quote-fix" class="btn primary">' + esc(t('quoteFix')) + '</button>' +
        '<button type="button" id="quote-copy" class="btn">' + esc(t('copy')) + '</button></p>' +
      (others.length ? '<div class="q-others" dir="' + i18n.dir() + '"><span class="small muted">' + esc(t('quoteOthers')) + '</span>' + others.map(function (o) {
        return '<button type="button" class="link small" data-k="' + o.i + '">' + esc(refText(o.y.d)) + ' · ' + esc(t('quoteOtherDiff', { n: o.y.r.diff })) + '</button>';
      }).join('') + '</div>' : '');
  }
  function quoteOut(x) {                            // the words of the data the quote covers, as the panel inserts them
    var r = x.r, rg = [r.from, r.to];
    return fmt(x.h, catalog.byId[x.h.b], optsFor({ mode: 'full', ar: r.side === 'ar' ? rg : null, uz: r.side === 'uz' ? rg : null },
      // the alphabet of the quote for its words; the tahrij of an Arabic quote in the alphabet of the panel ('new' is
      // Latin for HadisCore.format: the 2026 letters are made by fmt())
      { withAr: r.side === 'ar', withUz: r.side === 'uz', withBob: false, script: r.side === 'uz' ? (r.lat ? 'latn' : 'cyrl') : latn() ? 'latn' : 'cyrl' }));
  }
  $('m-quote').addEventListener('click', quoteOpen);
  $('quote-run').addEventListener('click', quoteRun);
  $('quote-in').addEventListener('keydown', function (e) { if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); quoteRun(); } });
  $('quote-doc').addEventListener('click', function () {
    if (!inWord) return;
    Word.run(function (ctx) { var s = ctx.document.getSelection(); s.load('text'); return ctx.sync().then(function () { return s.text; }); })
      .then(function (txt) { if (String(txt || '').trim()) { $('quote-in').value = txt; quoteRun(); } else toast(t('quoteNoSel')); }, function () { toast(t('quoteNoSel')); });
  });
  $('quote-out').addEventListener('click', function (e) {
    var st = quoteState, x = st && st.list[st.k], el = e.target.closest('a, button');
    if (!x || !el) return;
    if (el.id === 'quote-open') {
      e.preventDefault();
      var rg = [x.r.from, x.r.to];
      $('quote-dlg').close();
      go(x.d, { mode: 'full', ar: x.r.side === 'ar' ? rg : null, uz: x.r.side === 'uz' ? rg : null });
    } else if (el.id === 'quote-fix') {
      var out = quoteOut(x);
      if (!out.paras.length) { toast(t('nothingToInsert')); return; }
      $('quote-dlg').close();
      insertOut(out);
    } else if (el.id === 'quote-copy') {
      copy(quoteOut(x)).then(function () { toast(t('copied')); }, function (er) { toast(t('error') + ': ' + (er && er.message || er)); });
    } else if (el.dataset.k != null) { st.k = +el.dataset.k; quoteDraw(); }
  });
  document.addEventListener('keydown', function (e) {
    var tag = (e.target.tagName || ''), typing = /^(INPUT|TEXTAREA|SELECT)$/.test(tag) || e.target.isContentEditable, dlg = document.querySelector('dialog[open]');
    if (e.key === '/' && !typing && !dlg && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); $('q').focus(); $('q').select(); }
    else if ((e.ctrlKey || e.metaKey) && e.key === 'Enter' && !dlg) { e.preventDefault(); doInsert(); }
    else if (e.altKey && (e.key === 'ArrowDown' || e.key === 'ArrowUp') && !dlg) { e.preventDefault(); step(e.key === 'ArrowDown' ? 1 : -1); }
    else if (e.key === 'Escape' && !typing && !dlg && $('more-menu').hidden && (sel.ar || sel.uz)) clearPick();
  });

  /* ---------- search ----------
     A reference («Бухорий 1», «Муслим 2564/2», «بخاري ١») is found at once from data/refs; any other text goes to the
     static index (web/search.js: Arabic with or without vowel marks, Uzbek in Cyrillic or Latin letters). Its results
     stream in (local, index, final) into the same list; the shown items get the hadith text with the words marked.
     A newer text cancels the older search at the first keystroke. */
  var PAGE = 20, shown = PAGE, results = [], byD = {}, searchQ = '', deepSeq = 0, deepState = null, searchTimer;
  function stopDeep() {
    deepSeq++;
    if (deepState && deepState.ctl) try { deepState.ctl.abort(); } catch (e) {}
    deepState = null;
  }
  function refDocs(ref) {
    var out = [];
    if (!ref) return out;
    for (var n = ref.n; n <= (ref.to || ref.n); n++) {
      var d = data.docOf({ book: ref.book, n: n, s: n === ref.n ? ref.s : null });
      if (d >= 0) out.push(d);
    }
    return out;
  }
  function runSearch(now) {
    var q = $('q').value;
    stopDeep(); searchQ = q;
    results = []; byD = {}; shown = PAGE;
    $('results').innerHTML = '';
    var ref = q.trim() ? C.parseRef(q) : null;
    if (ref) {
      refDocs(ref).forEach(function (d) { addResult({ d: d, hits: [] }); });
      if (results.length) go(results[0].d);
      deepNote('', false, 'done');
    } else if (q.trim().length >= 2 && search) {
      var st = deepState = { seq: deepSeq, q: q, done: false, ctl: typeof AbortController !== 'undefined' ? new AbortController() : null };
      deepNote(t('searching'), true, 'run');
      search.run(q, { limit: 100, signal: st.ctl && st.ctl.signal,
        onResults: function (list) { if (st === deepState) take(list); }
      }).then(function (list) {
        if (st !== deepState) return;
        st.done = true;
        take(list);
        deepNote(list && list.partial ? t('searchPartial') : '', false, 'done');
        statusLine();
      });
    } else deepNote('', false, 'done');
    drawResults(); statusLine();
  }
  function addResult(x) {
    var r = byD[x.d];
    if (!r) { r = byD[x.d] = { d: x.d, hits: [] }; results.push(r); }
    if (x.hits && x.hits.length) { r.hits = x.hits; r.frag = null; }
    if ('tier' in x) { r.tier = x.tier; r.same = !!x.same; }
    return r;
  }
  /* Every stage (index, final) comes in the order of the search (r.rank: the place by relevance); the list is put in
     the order of settings.sortBy (the ones the stage no longer has stay after it), the shown items are painted again in
     their new places. 'books': first by how fully the query is there (the levels of the search: the phrase in the very
     forms, the phrase, the words near each other, every word, a part of the words — the owner, 2026-11-02: a hadith with
     one word of the query must not come before the one with the whole query), in a level the books in BOOK_ORDER, in a
     book by the number (the three best are marked); 'score': as the search ranks. take(null): the same list in the
     order chosen now. */
  function levelOf(r) { return r.tier == null ? 9 : r.tier * 2 + (r.same ? 1 : 0); }
  function byBooks(list) {
    var rank = function (d) { var b = data.refOf(d); return b ? C.bookRank(b.book) : 99; };
    return list.map(function (r, i) { return { r: r, l: levelOf(r), k: rank(r.d), i: i }; })
      .sort(function (a, b) { return b.l - a.l || a.k - b.k || a.r.d - b.r.d || a.i - b.i; }).map(function (x) { return x.r; });
  }
  function take(list) {
    var order = [], seen = {};
    if (list) list.forEach(function (x, i) { var r = addResult(x); if (!seen[r.d]) { seen[r.d] = 1; r.rank = i; order.push(r); } });
    else results.forEach(function (r) { if (r.rank != null) { seen[r.d] = 1; order.push(r); } });
    order.sort(function (a, b) { return a.rank - b.rank; });
    if (settings.sortBy === 'books') order = byBooks(order);
    results.forEach(function (r) { if (!seen[r.d]) order.push(r); });
    var lis = $('results').children, moved = false;
    for (var i = 0; i < order.length && !moved; i++) moved = order[i] !== results[i];
    results = order;
    if (moved) {
      results.forEach(function (r) { r.li = null; });
      for (var k = 0; k < lis.length && k < results.length; k++) { results[k].li = lis[k]; results[k].html = ''; }
    }
    results.forEach(function (r) { if (r.li) paintResult(r); });
    drawResults(); statusLine();
  }
  function statusLine() {
    var q = searchQ.trim();
    $('status').textContent = !q ? '' : results.length ? t('results', { n: results.length }) : deepState && !deepState.done ? '' : t('notFound');
  }
  function deepNote(text, busy, state) {
    var el = $('deep-note');
    el.dataset.state = state || ''; el.dataset.q = searchQ;
    if (el.textContent !== text) el.textContent = text;
    el.classList.toggle('busy', !!busy && !!text);
  }
  function drawResults() {
    var list = $('results');
    for (var i = list.children.length; i < Math.min(shown, results.length); i++) {
      var li = document.createElement('li');
      li.dataset.i = i; results[i].li = li; results[i].html = '';
      list.appendChild(li);
      paintResult(results[i]);
    }
    [].forEach.call(list.children, function (li) { var r = results[+li.dataset.i]; li.classList.toggle('active', !!r && r.d === cur.d); });
    $('more').hidden = results.length <= shown;
    loadTexts();
  }
  function paintResult(r) {
    if (!r.li) return;
    var html = resultHtml(r);
    if (html !== r.html) { r.html = html; r.li.innerHTML = html; }
  }
  function hitOf(r, src) { return (r.hits || []).filter(function (h) { return h.src === src && h.ranges && h.ranges.length; })[0] || null; }
  function resultHtml(r) {
    var ref = data.refOf(r.d), best = settings.sortBy === 'books' && r.rank != null && r.rank < 3 && results.length > 3;
    var h = '<div class="ref"><b>' + esc(ref ? shortName(ref.book) + ' ' + C.refLabel(ref) : '#' + r.d) + '</b>' +
      (ref ? ' · ' + esc(titleName(ref.book)) : '') + (best ? ' <span class="best" title="' + esc(t('bestHint')) + '">' + esc(t('best')) + '</span>' : '') + '</div>';
    if (!r.h) return h + '<div class="small muted">' + esc(t('loading')) + '</div>';
    var ar = r.h.ar || '', uz = r.h.uz ? T.stripHtml(r.h.uz).replace(/\s+/g, ' ') : '';
    var ha = hitOf(r, 'ar'), hu = hitOf(r, 'uz');
    if (ar) h += '<div class="ar" dir="rtl" lang="ar">' + (ha ? fragmentHtml(r.h.ar, ha.ranges, false) : esc(ar.slice(0, 220))) + '</div>';
    if (uz) h += '<div class="uz">' + (hu ? fragmentHtml(T.stripHtml(r.h.uz), hu.ranges, latn()) : esc(uzText(uz.slice(0, 240)))) + '</div>';
    return h;
  }
  /* the text around the place with the most marked words; the text is escaped piece by piece, only <mark> is ours */
  function fragmentHtml(text, ranges, lat) {
    var MAX = 220, best = 0, bestN = 0;
    function out(s) { s = s.replace(/\s+/g, ' '); return esc(lat ? C.toLatin(s) : s); }
    ranges = ranges.slice().sort(function (a, b) { return a[0] - b[0]; });
    for (var i = 0; i < ranges.length; i++) {
      var n = 0;
      for (var j = i; j < ranges.length && ranges[j][1] <= ranges[i][0] + MAX - 60; j++) n += ranges[j][1] - ranges[j][0] > 2 ? 2 : 1;
      if (n > bestN) { bestN = n; best = i; }
    }
    var first = ranges.length ? ranges[best][0] : 0, start = Math.max(0, first - 60);
    if (start > 0) { var sp = text.lastIndexOf(' ', start); if (sp > 0 && first - sp < 120) start = sp + 1; }
    var end = Math.min(text.length, start + MAX);
    if (end < text.length) { var e2 = text.indexOf(' ', end); if (e2 > 0 && e2 - end < 30) end = e2; }
    var h = start > 0 ? '…' : '', pos = start;
    ranges.forEach(function (rg) {
      var a = Math.max(rg[0], pos), b = Math.min(rg[1], end);
      if (b <= a) return;
      h += out(text.slice(pos, a)) + '<mark>' + out(text.slice(a, b)) + '</mark>';
      pos = b;
    });
    return h + out(text.slice(pos, end)) + (end < text.length ? '…' : '');
  }
  /* the texts of the shown items, two loads at a time in the order of the list */
  var textActive = 0;
  function loadTexts() {
    var seq = deepSeq;
    while (textActive < 2) {
      var r = null;
      for (var k = 0; k < Math.min(shown, results.length); k++) if (!results[k].h && !results[k].loading) { r = results[k]; break; }
      if (!r) return;
      r.loading = true; textActive++;
      (function (r) {
        data.hadith(r.d).then(function (h) { r.h = h; paintResult(r); }, function () { r.loading = false; r.failed = true; })
          .then(function () { textActive--; if (seq === deepSeq || results.indexOf(r) >= 0) loadTexts(); });
      })(r);
    }
  }
  $('more').addEventListener('click', function () { shown += PAGE; drawResults(); });
  var RECENT_KEY = 'hadis-recent';
  function recentList() {
    try { var x = JSON.parse(localStorage.getItem(RECENT_KEY) || '[]'); return Array.isArray(x) ? x.filter(function (s) { return typeof s === 'string'; }).slice(0, 10) : []; } catch (e) { return []; }
  }
  function addRecent(q) {
    q = String(q || '').trim();
    if (q.length < 2 || q.length > 80) return;
    try { localStorage.setItem(RECENT_KEY, JSON.stringify([q].concat(recentList().filter(function (x) { return x !== q; })).slice(0, 10))); } catch (e) {}
    drawRecent();
  }
  function drawRecent() { $('recent-q').innerHTML = recentList().map(function (q) { return '<option value="' + esc(q) + '">'; }).join(''); }
  drawRecent();
  $('q').addEventListener('input', function () {
    clearTimeout(searchTimer); stopDeep(); deepNote('', false, 'wait');
    searchTimer = setTimeout(runSearch, 300);
  });
  $('q').addEventListener('keydown', function (e) {
    if (e.key !== 'Enter') return;
    clearTimeout(searchTimer); runSearch(true); addRecent($('q').value);
  });
  $('results').addEventListener('click', function (e) {
    var li = e.target.closest('li'); if (!li) return;
    var r = results[+li.dataset.i]; if (!r) return;
    addRecent($('q').value);
    go(r.d).then(function () { drawResults(); });
    [].forEach.call($('results').children, function (x) { x.classList.toggle('active', x === li); });
    if (document.body.classList.contains('narrow') || window.innerWidth <= 800) document.querySelector('.selection').scrollIntoView({ block: 'start', behavior: scrollBeh() });
  });

  /* ---------- «Кун ҳадиси»: the hadith of the day ----------
     The same for everyone on a calendar day, without a server and without chance: web/data/daily.json
     (tools/build_daily.py: short hadiths of Bukhari and Muslim with the Arabic and the Uzbek text and a matn) is a list
     of references; the day number (days since 1970 of the local date) times a prime step picks one. The card shows only
     the matn; insert, copy and the star take the hadith in the mode «only matn»; «Тўлиқ ҳадис» opens it on the card. */
  var DAY_CLOSED = 'hadis-day-closed', dayLoad = null, daySeq = 0, day = { d: -1, h: null, P: null };   // DAY_CLOSED: the old key
  function dayKey(dt) { return dt.getFullYear() + '-' + (dt.getMonth() + 1) + '-' + dt.getDate(); }
  /* shown: turned on and not hidden for this local date (a hide of another day is over) */
  function dayOn(dt) { return !!settings.day && settings.dayHide !== dayKey(dt || new Date()); }
  /* the user's choice: kept in every store at once (localStorage, the cookie, the host), with its time */
  function setDayPref(on, hide) {
    settings.day = !!on; settings.dayHide = hide || ''; settings.dayAt = Math.max(Date.now(), (settings.dayAt || 0) + 1);
    saveSettings(); savePref();
    try { localStorage.removeItem(DAY_CLOSED); } catch (e) {}
    if ($('s-day')) $('s-day').checked = settings.day;
  }
  function loadDaily() {
    if (!dayLoad) {
      dayLoad = data.loader.get('data/daily.json?v=' + encodeURIComponent((data.version || '') + '.' + myVersion()), 'json', true).then(function (j) {
        if (!j || !Array.isArray(j.list) || !j.list.length || !Array.isArray(j.books)) throw new Error('bad daily.json');
        return j;
      });
      dayLoad.then(null, function () { dayLoad = null; });
    }
    return dayLoad;
  }
  function dayDoc(j, dt) {
    var n = j.list.length, num = Math.floor(Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate()) / 864e5);
    var k = ((num * (+j.step || 7919)) % n + n) % n;
    for (var i = 0; i < n; i++) {                         // a reference the data does not have: the next one of the list
      var x = j.list[(k + i) % n], book = j.books[x && x[0]];
      var d = book && catalog.byId[book] ? data.docOf({ book: book, n: x[1], s: x[2] || 1 }) : -1;
      if (d >= 0) return d;
    }
    return -1;
  }
  function drawDay() {
    var card = $('day'), now = new Date(), seq = ++daySeq;
    card.dataset.day = dayKey(now);
    if (!catalog || !dayOn(now)) { card.hidden = true; return Promise.resolve(); }
    return loadDaily().then(function (j) {
      var d = dayDoc(j, now);
      if (d < 0) throw new Error('no hadith of the day');
      return day.d === d && day.h && day.lat === latn() ? day.h : hadithView(d);
    }).then(function (h) {
      if (seq !== daySeq) return;
      day = { d: h.d, h: h, P: C.parseHadith(h, 'matn'), lat: latn() };
      paintDay(); card.hidden = false;
    }).catch(function () { if (seq === daySeq) card.hidden = true; });
  }
  /* the words of the matn (the whole text where it has no matn: the Uzbek of Muslim has no bold); without the number
     of the translation at the start («295.», «533, 534.», «3м.») */
  function dayHtml(side) {
    var paras = day.P[side].paras, list = C.words(paras, C.hasMatn(paras) ? 'matn' : 'full'), lat = side === 'uz' && latn();
    if (side === 'uz') while (list.length && !list[0].matn && /^\d+[а-яa-z]?[.,]$/i.test(list[0].text)) list = list.slice(1);
    return list.map(function (w, i) {
      var prev = ' ', s = w.parts.map(function (x) { var y = lat ? C.toLatin(x.t, prev) : x.t; prev = x.t.slice(-1); return y; }).join('');
      // data-k: the number of the word in words() (the start without the number of the translation counts): a part
      // selected with the mouse (nativePick) is inserted
      return (i && w.gap ? '<span class="gap">…</span> ' : '') + '<span data-k="' + w.k + '">' + esc(s) + '</span>';
    }).join(' ');
  }
  function paintDay() {
    if (!day.h || !catalog || !data) return;
    var r = data.refOf(day.d);
    if (!r) return;
    $('day-ref').textContent = shortName(r.book) + ' ' + C.refLabel(r) + ' · ' + titleName(r.book);
    $('day-ref').href = '#' + r.book + ':' + C.refLabel(r);
    $('day-ar').innerHTML = day.P.ar.paras.length ? dayHtml('ar') : '';
    $('day-uz').innerHTML = day.P.uz.paras.length ? dayHtml('uz') : '';
    dayStar();
  }
  function dayOut() {
    var p = takeNative(), part = p && p.where === 'day' ? p : null;
    return fmt(day.h, catalog.byId[day.h.b], optsFor({ mode: 'matn', ar: part && part.ar, uz: part && part.uz }));
  }
  function dayMark() {
    var paras = day.P.ar.paras, w = C.words(paras, C.hasMatn(paras) ? 'matn' : 'full');
    return { d: day.d, m: 'matn', time: Date.now(), t: w.slice(0, 12).map(function (y) { return y.text; }).join(' ').slice(0, 120) };
  }
  function dayIsFav() { var id = markId({ d: day.d, m: 'matn' }); return day.h && marks.fav.some(function (y) { return markId(y) === id; }); }
  function dayStar() {
    var b = $('day-fav'), on = !!dayIsFav();
    b.setAttribute('aria-pressed', on); b.title = t(on ? 'favRemove' : 'favAdd'); b.setAttribute('aria-label', b.title);
  }
  $('day-fav').addEventListener('click', function () {
    if (!day.h) return;
    var id = markId({ d: day.d, m: 'matn' });
    if (dayIsFav()) { marks.fav = marks.fav.filter(function (y) { return markId(y) !== id; }); saveMarks('fav'); }
    else addMark('fav', dayMark());
  });
  $('day-insert').addEventListener('click', function () {
    if (!day.h) return;
    var out = dayOut();
    if (!out.paras.length) { toast(t('nothingToInsert')); return; }
    addMark('hist', dayMark());
    insertOut(out);
  });
  $('day-copy').addEventListener('click', function () {
    if (!day.h) return;
    addMark('hist', dayMark());
    copy(dayOut()).then(function () { toast(t('copied')); }, function (e) { toast(t('error') + ': ' + e.message); });
  });
  function dayFull(e) {
    if (e) e.preventDefault();
    if (!day.h) return;
    go(day.d, { mode: settings.mode === 'matn' ? 'sahabi' : settings.mode });
    document.querySelector('.selection').scrollIntoView({ block: 'start', behavior: scrollBeh() });
  }
  $('day-full').addEventListener('click', dayFull);
  $('day-ref').addEventListener('click', dayFull);
  /* the cross: the first time asks «turn it off or hide until tomorrow»; the answer is a setting (Settings, General) */
  function closeDay(how) {
    if (how === 'off') setDayPref(false, '');
    else setDayPref(true, dayKey(new Date()));
    daySeq++; $('day').hidden = true;
  }
  $('day-close').addEventListener('click', function () {
    if (settings.dayClose === 'off' || settings.dayClose === 'day') closeDay(settings.dayClose);
    else { $('day-ask').returnValue = ''; $('day-ask').showModal(); }
  });
  $('day-ask').addEventListener('close', function () {
    var how = $('day-ask').returnValue;
    if (how !== 'off' && how !== 'day') return;
    settings.dayClose = how; $('s-day-close').value = how;
    closeDay(how);
  });
  document.addEventListener('visibilitychange', function () {   // the panel stayed open past midnight
    if (!document.hidden && catalog && $('day').dataset.day !== dayKey(new Date())) drawDay();
  });

  /* ---------- settings dialog ---------- */
  var B = { withAr: 's-with-ar', withUz: 's-with-uz', withBob: 's-bob', matnBold: 's-matn-bold', arMatnBold: 's-ar-matn-bold',
            refUz: 's-ref-uz', refAr: 's-ref-ar', withBook: 's-with-book', withHukm: 's-hukm', newPara: 's-newpara', picBgOn: 's-picbg-on',
            withGrades: 's-grades', gradesMatn: 's-grades-matn', allGrades: 's-all-grades',
            day: 's-day', smooth: 's-smooth' };
  var SEL = { order: 's-order', mode: 's-mode', baskSrc: 's-bask-src', picAction: 's-pic-action', script: 's-script', apos: 's-apos', dayClose: 's-day-close',
              sortBy: 's-sort' };
  function fillSettingsForm() {
    Object.keys(B).forEach(function (k) { $(B[k]).checked = !!settings[k]; });
    Object.keys(SEL).forEach(function (k) { $(SEL[k]).value = settings[k]; });
    $('s-picbg').value = settings.picBg;
    refSample();
  }
  function refSample() {
    var h = { b: 'bukhari', n: 1 };
    $('s-ref-sample').innerHTML = '<em>' + esc(C.tahrij(h, catalog ? catalog.byId.bukhari : 'bukhari', uzLang(), { withBook: settings.withBook })) + '</em>.';   // <i> is the knob of .switch
  }
  Object.keys(B).forEach(function (k) {
    $(B[k]).addEventListener('change', function () {
      if (k === 'day') { setDayPref(this.checked, ''); drawDay(); return; }
      settings[k] = this.checked; muted[k] = false; saveSettings(); refSample();
      if (k === 'smooth') { applyMotion(); return; }
      render();
    });
  });
  Object.keys(SEL).forEach(function (k) {
    $(SEL[k]).addEventListener('change', function () {
      if (k === 'script') { setScript(this.value); return; }
      if (k === 'apos') { setApos(this.value); return; }
      settings[k] = this.value; saveSettings();
      if (k === 'sortBy') { take(null); return; }
      if (k === 'mode') { sel.mode = this.value; sel.ar = sel.uz = null; reparse(); }
      render();
    });
  });
  $('s-picbg').addEventListener('input', function () { settings.picBg = this.value; saveSettings(); });
  $('settings-btn').addEventListener('click', function () {
    fillSettingsForm(); fillThemes(); showTab(settings.tab); $('settings').showModal(); fmtEditor();
  });
  $('settings').addEventListener('close', function () { saveSettings(); render(); });
  var resetTimer;
  function resetIdle() { var b = $('s-reset'); clearTimeout(resetTimer); delete b.dataset.sure; b.textContent = t('resetSettings'); }
  $('s-reset').addEventListener('click', function () {
    var b = this;
    if (!b.dataset.sure) { b.dataset.sure = '1'; b.textContent = t('pressAgain'); resetTimer = setTimeout(resetIdle, 4000); return; }
    var keep = settings.script, keepApos = settings.apos;
    settings = JSON.parse(JSON.stringify(DEFAULTS)); settings.script = keep; settings.apos = keepApos;
    sanitize(settings); setDayPref(true, ''); resetIdle(); applyTheme(); applyMotion(); fillThemes(); fillSettingsForm(); fmtEditor(); render();
    drawDay();
    toast(t('resetDone'));
  });

  /* ---------- version, updates (version.json, releases.js) ---------- */
  function myVersion() { var m = document.querySelector('meta[name="hadis-version"]'); return m ? m.content : ''; }
  /* Word for Mac keeps the task pane in its cache: a newer version.json reloads the page at a new address */
  function checkVersion() {
    var mine = myVersion();
    if (!mine || !window.fetch || !/^https?:/.test(location.protocol)) return;
    fetch('version.json?t=' + Date.now(), { cache: 'no-store' }).then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      if (!d || !d.v || d.v === mine) return;
      try { if (sessionStorage.getItem('hadis-reload') === d.v) return; sessionStorage.setItem('hadis-reload', d.v); } catch (e) {}
      location.replace(location.pathname + '?v=' + encodeURIComponent(d.v) + (HOST ? '&host=' + HOST : '') + location.hash);
    }).catch(function () {});
  }
  checkVersion();
  function loadReleases(v) {
    return new Promise(function (ok, fail) {
      var s = document.createElement('script');
      s.src = 'releases.js?v=' + encodeURIComponent(v);
      s.onload = function () { s.remove(); window.HadisReleases ? ok(window.HadisReleases) : fail(); };
      s.onerror = function () { s.remove(); fail(); };
      document.head.appendChild(s);
    });
  }
  function showNews(list) {
    var L = uzLang();
    $('upd-news').innerHTML = list.slice(0, 3).map(function (r) {
      var title = r[2][L] || r[2].uz, items = r[3][L] || r[3].uz || [];
      return '<div class="rel"><b><span class="ver">' + esc(r[0]) + '</span>' + esc(title) + '<span class="date">' + esc(r[1]) + '</span></b>' +
        '<ul>' + items.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul></div>';
    }).join('');
  }
  function openUpdates() {
    var mine = myVersion(), st = $('upd-status'), now = $('upd-now');
    st.className = 'upd-status'; st.textContent = t('updChecking'); now.hidden = true;
    $('updates').showModal();
    fetch('version.json?t=' + Date.now(), { cache: 'no-store' }).then(function (r) { return r.ok ? r.json() : Promise.reject(); }).then(function (d) {
      return loadReleases(d.v).then(function (list) {
        showNews(list);
        if (d.v === mine) { st.textContent = t('updLatest', { v: list[0][0] }); return; }
        st.textContent = t('updNew', { v: list[0][0] }); st.classList.add('new'); now.hidden = false;
        now.onclick = function () {
          try { sessionStorage.setItem('hadis-reload', d.v); } catch (e) {}
          location.replace(location.pathname + '?v=' + encodeURIComponent(d.v) + (HOST ? '&host=' + HOST : ''));
        };
      });
    }).catch(function () { st.textContent = t('updOffline'); loadReleases(mine).then(showNews, function () {}); });
  }
  $('updates-btn').addEventListener('click', function (e) { e.preventDefault(); openUpdates(); });
  loadReleases(myVersion()).then(function (list) {
    var a = document.querySelector('#app-version a');
    a.textContent = 'v' + list[0][0];
    a.addEventListener('click', function (e) { e.preventDefault(); openUpdates(); });
    $('app-version').hidden = false;
  }, function () {});

  /* ---------- layout: the height of the top bar for the sticky search (as in Khatt al-Quran) ---------- */
  function layoutVars() {
    var tb = document.querySelector('.topbar');
    document.documentElement.style.setProperty('--topbar-h', tb.offsetHeight + 'px');
    document.body.classList.toggle('narrow', window.innerWidth <= 800);
  }
  window.addEventListener('resize', layoutVars);

  /* ---------- init ---------- */
  function init() {
    applyI18n(); fillThemes(); layoutVars(); styleDom();
    var own = /^https?:/.test(location.protocol) ? new URL('./', location.href).href : './';
    data = C.Data({ bases: [own] });
    search = HadisSearch.Search({ bases: [own], text: function (d, src, signal) {
      return data.hadith(d, signal).then(function (h) { return src === 'ar' ? h.ar || null : h.uz ? T.stripHtml(h.uz) : null; });
    } });
    /* the first screen: the card, «Кун ҳадиси» and the footer are shown together once the hadith and the day are there
       (body.starting: visibility only), so nothing on the screen jumps while they come; at most 2.5 s */
    var shown = false, parts = 0;
    function started() { if (!shown) { shown = true; document.body.classList.remove('starting'); } }
    function part() { if (++parts >= 2) started(); }
    setTimeout(started, 2500);
    data.books().then(function (c) {
      catalog = c;
      fillBooks();
      return data.allRefs();
    }).then(function () {
      $('status').textContent = '';
      var first = fromHash() ? lastGo : go(catalog.byId.bukhari ? catalog.byId.bukhari.doc[0] : 0);
      Promise.resolve(first).then(part, part);
      drawMarks(); drawDay().then(part, part);
      if (searchQ.trim()) runSearch();
    }).catch(function (e) { started(); $('status').textContent = t('loadFailed') + ' ' + (e && e.message || e); });
  }
  if (/[?&]debug\b/.test(location.search)) window.HadisDebug = { buildOoxml: buildOoxml, buildHtml: buildHtml, current: current, go: go,
    settings: function () { return settings; }, sel: function () { return sel; }, marks: function () { return marks; },
    setWord: function (v) { inWord = !!v; }, day: function () { return day; }, dayOf: function (dt) { return loadDaily().then(function (j) { return dayDoc(j, dt); }); },
    picture: function () { return arabicPicture(current()); }, basket: basketOut, hostDoc: function () { return hostDoc(current()); },
    results: function () { return results.map(function (r) { return { d: r.d, rank: r.rank, tier: r.tier, same: !!r.same }; }); }, search: function () { return search; },
    text: function () { var o = current(); return o ? o.text : ''; }, bookOf: function (d) { var r = data.refOf(d); return r ? r.book : ''; },
    // the checks of the insert into Word (tools/test-embed/word.mjs): the runs joined or not, the guard, a hadith in a mode
    join: function (on) { joinRuns = !!on; }, insertMs: function (ms) { INSERT_MS = ms; }, inserting: function () { return inserting; },
    insert: function (out, pic) { return insertWordWith(out, pic || null); }, chunks: function (out) { return insertChunks(out, null).map(function (x) { return x.paras.length; }); },
    hostDocOf: hostDoc, books: function () { return catalog.books; }, docOf: function (book, n, sub) { return data.docOf({ book: book, n: n, s: sub }); },
    formatOf: function (d, mode) {
      return hadithView(d).then(function (h) { return fmt(h, catalog.byId[h.b], optsFor({ mode: mode, ar: null, uz: null })); });
    } };

  if (window.Office && Office.onReady) {
    Office.onReady(function (info) {
      inWord = info && info.host === Office.HostType.Word;
      online = !!HOST || inWord && info.platform !== Office.PlatformType.PC && info.platform !== Office.PlatformType.Mac;
      if (inWord) windows = info.platform === Office.PlatformType.PC;
      if (inWord) { document.body.classList.add('office'); $('host-badge').hidden = false; layoutVars(); }
      render();
    });
  }
  if (HOST) {
    $('host-badge').textContent = { gdocs: 'Google Docs', wp: 'WordPress', ext: t('browser'), tg: 'Telegram' }[HOST];
    if (HOST === 'ext') $('host-badge').dataset.i18n = 'browser';
    $('host-badge').hidden = false;
    document.body.classList.add('embedded');
  }
  if ('serviceWorker' in navigator && /^https:|^http:\/\/localhost/.test(location.protocol + '//' + location.hostname)) {
    window.addEventListener('load', function () { setTimeout(function () { try { navigator.serviceWorker.register('sw.js').catch(function () {}); } catch (e) {} }, 3000); });
  }
  init();
})();
