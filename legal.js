/* The pages help, privacy, terms, install, releases, sources, 404: only Uzbek, in three alphabets — uz (Cyrillic, the main one),
   uz_latn (Latin) and uz_new (the 2026 Latin: ö ğ ş ç). The pages are written in uz and uz_latn ([data-l] blocks);
   uz_new is the uz_latn text through the table of letters HadisText.latinStyle (web/text.js), as in the panel.
   The alphabet: ?lang=uz_new | uz_latn | uz, else the alphabet of the panel (hadis-settings.script, else the cookie
   hadis-pref), else Cyrillic. The switch «Öz | Ўз | Oʼz» (#script) shows that alphabet and is remembered for the panel too.
   A page may set window.HadisPage = function (lang) {…} to draw its own part (releases.html, sources.html; lang: uz | uz_latn)
   and call window.HadisPageRedraw() when its data comes later. */
(function () {
  'use strict';
  var box = document.getElementById('script');
  var SCRIPT = { uz: 'cyrl', uz_latn: 'latn', uz_new: 'new' };
  function readSettings() { try { return JSON.parse(localStorage.getItem('hadis-settings') || '{}') || {}; } catch (e) { return {}; } }
  function readCookie() {
    try { var m = /(?:^|;\s*)hadis-pref=([^;]*)/.exec(document.cookie || ''); return m ? JSON.parse(decodeURIComponent(m[1])) || {} : {}; } catch (e) { return {}; }
  }
  function pref() { var st = readSettings(); return st.script ? st : readCookie(); }
  /* the text nodes of the page in the style of the Latin chosen; the canonical text is remembered for another switch */
  var memo = typeof WeakMap === 'function' ? new WeakMap() : null;
  function restyle(opts) {
    var T = window.HadisText;
    if (!memo || !T || !T.latinStyle || !document.body) return;
    var w = document.createTreeWalker(document.body, 4, null), n;
    while ((n = w.nextNode())) {
      var el = n.parentElement;
      if (!el || el.closest('#script, [translate="no"], script, style, code')) continue;
      var v = n.nodeValue, m = memo.get(n), orig = m && v === m.conv ? m.orig : v;
      if (!/[A-Za-zʻʼ‘’]/.test(orig)) continue;
      var conv = opts ? T.latinStyle(orig, opts) : orig;
      if (conv !== v) n.nodeValue = conv;
      memo.set(n, { orig: orig, conv: conv });
    }
  }
  var shown = 'uz';
  function show(l) {
    if (l !== 'uz_latn' && l !== 'uz_new') l = 'uz';
    shown = l;
    var base = l === 'uz' ? 'uz' : 'uz_latn', apos = pref().apos;
    document.documentElement.lang = base === 'uz_latn' ? 'uz-Latn' : 'uz-Cyrl';
    document.documentElement.dir = 'ltr';
    if (box) [].forEach.call(box.querySelectorAll('[data-s]'), function (b) { b.setAttribute('aria-pressed', String(b.dataset.s === l)); });
    restyle(null);                                   // back to the canonical text first (the page may draw its part again)
    [].forEach.call(document.querySelectorAll('[data-l]'), function (el) {
      el.style.display = el.dataset.l === base ? (el.tagName === 'SPAN' ? 'inline' : 'block') : 'none';
    });
    if (window.I18n) [].forEach.call(document.querySelectorAll('p.author, .site-head'), function (el) { I18n.localize(el, base); });
    if (typeof window.HadisPage === 'function') window.HadisPage(base);
    var nw = l === 'uz_new', ap = base === 'uz_latn' && /^(typo|ascii)$/.test(apos || '') ? apos : 'mod';
    if (nw || ap !== 'mod') restyle({ nw: nw, apos: ap });
    /* the title of the tab: «Ҳадислар — <the title of the page>» (the small line over the title is the name) */
    var h1 = document.querySelector('h1');
    if (h1) {
      var parts = [].filter.call(h1.querySelectorAll(':scope > [data-l], :scope > .pre > [data-l]'), function (x) { return x.style.display !== 'none'; })
        .map(function (x) { return x.textContent.replace(/\s+/g, ' ').trim(); });
      document.title = parts.length ? parts.join(' — ') : (h1.innerText || h1.textContent).replace(/\s+/g, ' ').trim();
    }
  }
  function remember(l) {
    var st = readSettings();
    st.script = SCRIPT[l] || 'cyrl';
    try { localStorage.setItem('hadis-settings', JSON.stringify(st)); } catch (e) {}
    var c = readCookie();
    c.script = st.script;
    try { document.cookie = 'hadis-pref=' + encodeURIComponent(JSON.stringify(c)) + '; max-age=31536000; path=/; SameSite=None; Secure'; } catch (e) {}
  }
  var q = (location.search.match(/[?&]lang=(\w+)/) || [])[1] || '';
  var lang = q === 'uz_new' || q === 'new' ? 'uz_new' : q === 'uz_latn' || q === 'uzl' || q === 'latn' ? 'uz_latn' : q === 'uz' || q === 'cyrl' ? 'uz' : '';
  if (!lang) { var s = pref().script; lang = s === 'latn' ? 'uz_latn' : s === 'new' ? 'uz_new' : 'uz'; }
  if (box) box.addEventListener('click', function (e) {
    var b = e.target.closest('[data-s]'); if (!b) return;
    show(b.dataset.s); remember(b.dataset.s);
    history.replaceState(null, '', '?lang=' + b.dataset.s);
  });
  show(lang);
  /* a page that draws its part later (sources.html, after its data came) draws it again in the alphabet shown */
  window.HadisPageRedraw = function () { show(shown); };
})();
