/* The search text of the hadiths: ONE normalisation for the index builder (Node, tools/build_search_index.js) and the
   query (browser, Word panel, extension, Worker). Taken unchanged from quranuz (web/core.js, QuranCore.text) — the same
   tokens, folds, Arabic stem and Latin <-> Cyrillic — plus two Uzbek folds: э -> е (the Latin "e" of "eshitdi") and
   ъ ь dropped (the Latin apostrophe of "ma'no" is dropped too). No DOM.
     HadisText.{norm, terms, queryTerms, tokenize, detectScript, scriptOf, arStem, arLight, latToCyr, cyrToLat, levenshtein,
                stripHtml, uzBase, uzInflected, latAlts, latinStyle, latinStyleHtml}                                                          */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.HadisText = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var AR_DIGITS = '٠١٢٣٤٥٦٧٨٩';


  /* ---------- Search text in any language and script ----------
     ONE normalisation for the index builder (node) and for the query (browser, Worker), see
     docs/search-index-format.md. Code point by code point (a memo per code point keeps the builder fast):
       NFKD + lower case; combining marks dropped after Latin, Greek, Cyrillic, Arabic, Hebrew letters and kept
       elsewhere (Indic vowel signs are letters of their block, never dropped); ZWJ / ZWNJ / tatweel dropped;
       Latin ı→i ß→ss ø→o æ→ae œ→oe ł→l đ ð→d þ→th ə→e, apostrophes joined (oʻzi → ozi; French l' d' qu' split);
       Cyrillic ё→е й→и э→е, ъ ь dropped (uzhadislar) and the loose level қ→к ғ→г ҳ→х ў→у ң→н ө→о ү ұ→у ә→а і→и һ→х ҷ→ч ҙ→з ҫ→с;
       Arabic script ٱ أ إ آ→ا, ى ی ے ې→ي, ة ہ ۀ ھ ە→ه, ک→ك, ؤ→و, ئ→ي, ء dropped, Arabic and Persian digits → 0-9;
       the ligatures of words ﷺ ﷻ ﷽ -> their words (ﷺ = صلى الله عليه وسلم typed in full).
     Tokens: words of space-separated scripts; Chinese / Japanese, Korean, Thai, Lao, Khmer, Burmese runs are
     cut into pairs of neighbouring characters (a run of one character is one token).                        */
  var SCRIPTS = [   // [first, last, script] of letters, sorted; anything else >= 0x80 is 'Zzzz' (a letter of no known script)
    [0xC0, 0x24F, 'Latn'], [0x250, 0x2AF, 'Latn'], [0x370, 0x3FF, 'Grek'], [0x400, 0x52F, 'Cyrl'], [0x530, 0x58F, 'Armn'],
    [0x590, 0x5FF, 'Hebr'], [0x600, 0x6FF, 'Arab'], [0x700, 0x74F, 'Syrc'], [0x750, 0x77F, 'Arab'], [0x780, 0x7BF, 'Thaa'],
    [0x7C0, 0x7FF, 'Nkoo'], [0x870, 0x8FF, 'Arab'], [0x900, 0x97F, 'Deva'], [0x980, 0x9FF, 'Beng'], [0xA00, 0xA7F, 'Guru'],
    [0xA80, 0xAFF, 'Gujr'], [0xB00, 0xB7F, 'Orya'], [0xB80, 0xBFF, 'Taml'], [0xC00, 0xC7F, 'Telu'], [0xC80, 0xCFF, 'Knda'],
    [0xD00, 0xD7F, 'Mlym'], [0xD80, 0xDFF, 'Sinh'], [0xE00, 0xE7F, 'Thai'], [0xE80, 0xEFF, 'Laoo'], [0xF00, 0xFFF, 'Tibt'],
    [0x1000, 0x109F, 'Mymr'], [0x10A0, 0x10FF, 'Geor'], [0x1100, 0x11FF, 'Hang'], [0x1200, 0x139F, 'Ethi'],
    [0x13A0, 0x13FF, 'Cher'], [0x1780, 0x17FF, 'Khmr'], [0x1C80, 0x1C8F, 'Cyrl'], [0x1C90, 0x1CBF, 'Geor'],
    [0x1E00, 0x1EFF, 'Latn'], [0x1F00, 0x1FFF, 'Grek'], [0x2C60, 0x2C7F, 'Latn'], [0x2D00, 0x2D2F, 'Geor'],
    [0x2D30, 0x2D7F, 'Tfng'], [0x2DE0, 0x2DFF, 'Cyrl'], [0x2E80, 0x2FDF, 'Hani'], [0x3005, 0x3007, 'Hani'],
    [0x3021, 0x3029, 'Hani'], [0x3031, 0x3035, 'Hani'], [0x3038, 0x303C, 'Hani'], [0x3041, 0x30FA, 'Hani'],
    [0x30FC, 0x30FF, 'Hani'], [0x3105, 0x312F, 'Hani'], [0x3131, 0x318F, 'Hang'], [0x31A0, 0x31BF, 'Hani'],
    [0x31F0, 0x31FF, 'Hani'], [0x3400, 0x4DBF, 'Hani'], [0x4E00, 0x9FFF, 'Hani'], [0xA640, 0xA69F, 'Cyrl'],
    [0xA720, 0xA7FF, 'Latn'], [0xA960, 0xA97F, 'Hang'], [0xAB30, 0xAB6F, 'Latn'], [0xAC00, 0xD7FF, 'Hang'],
    [0xF900, 0xFAFF, 'Hani'], [0xFB1D, 0xFB4F, 'Hebr'], [0xFB50, 0xFDFF, 'Arab'], [0xFE70, 0xFEFF, 'Arab'],
    [0x20000, 0x3FFFF, 'Hani']
  ];
  var PAIRS = { Hani: 1, Hang: 1, Thai: 1, Laoo: 1, Khmr: 1, Mymr: 1 };        // no spaces between words: pairs of letters
  var DROP_MARKS = { Latn: 1, Grek: 1, Cyrl: 1, Arab: 1, Hebr: 1, Armn: 1, Geor: 1, Zyyy: 1 };
  var MARK = [0x300, 0x36F, 0x483, 0x489, 0x591, 0x5BD, 0x5BF, 0x5BF, 0x5C1, 0x5C2, 0x5C4, 0x5C5, 0x5C7, 0x5C7,
    0x610, 0x61A, 0x64B, 0x65F, 0x670, 0x670, 0x6D6, 0x6DC, 0x6DF, 0x6E4, 0x6E7, 0x6E8, 0x6EA, 0x6ED, 0x8D3, 0x8E1,
    0x8E3, 0x8FF, 0x1AB0, 0x1AFF, 0x1DC0, 0x1DFF, 0x20D0, 0x20FF, 0xFE20, 0xFE2F];
  var SEPS = [0x80, 0xBF, 0xD7, 0xD7, 0xF7, 0xF7, 0x37E, 0x37E, 0x387, 0x387, 0x55A, 0x55F, 0x589, 0x58A,
    0x5BE, 0x5BE, 0x5C0, 0x5C0, 0x5C3, 0x5C3, 0x5C6, 0x5C6, 0x5F3, 0x5F4, 0x600, 0x60F, 0x61B, 0x61F, 0x66A, 0x66D,
    0x6D4, 0x6D4, 0x6DD, 0x6DE, 0x6E9, 0x6E9, 0x6FD, 0x6FE, 0x964, 0x965, 0x970, 0x970, 0xE2F, 0xE2F, 0xE4F, 0xE4F,
    0xE5A, 0xE5B, 0xF04, 0xF14, 0x104A, 0x104F, 0x10FB, 0x10FB, 0x1360, 0x1368, 0x166D, 0x166E, 0x17D4, 0x17DA,
    0x2000, 0x200A, 0x2010, 0x205F, 0x2070, 0x2BFF, 0x2E00, 0x2E7F, 0x3000, 0x3004, 0x3008, 0x3020, 0x3030, 0x3030,
    0x303D, 0x303F, 0x30FB, 0x30FB, 0xD800, 0xF8FF, 0xFD3E, 0xFD3F, 0xFE10, 0xFE1F, 0xFE30, 0xFE6F, 0xFF00, 0xFF0F,
    0xFF1A, 0xFF20, 0xFF3B, 0xFF40, 0xFF5B, 0xFF65, 0xFFF0, 0xFFFF, 0x1F000, 0x1FFFF, 0xE0000, 0x10FFFF];
  // dropped without breaking the word: joiners, soft hyphen, direction marks, tatweel, hamza, small waw / yeh
  var SKIP = { 0xAD: 1, 0x640: 1, 0x6E5: 1, 0x6E6: 1, 0x200B: 1, 0x200C: 1, 0x200D: 1, 0x200E: 1, 0x200F: 1, 0x2060: 1,
               0xFEFF: 1, 0x621: 1, 0x674: 1, 0x34F: 1 };
  var APOS = { 0x27: 1, 0x60: 1, 0xB4: 1, 0x2BB: 1, 0x2BC: 1, 0x2BE: 1, 0x2BF: 1, 0x2018: 1, 0x2019: 1, 0x2032: 1 };
  var ELISION = { l: 1, d: 1, j: 1, m: 1, n: 1, s: 1, t: 1, c: 1, qu: 1 };   // French / Italian l'homme, qu'il: two words
  var FOLD = {
    'ı': 'i', 'ß': 'ss', 'ø': 'o', 'æ': 'ae', 'œ': 'oe', 'ł': 'l', 'đ': 'd', 'ð': 'd', 'þ': 'th', 'ħ': 'h', 'ŧ': 't',
    'ĸ': 'k', 'ə': 'e', 'ɛ': 'e', 'ɔ': 'o', 'ŋ': 'n', 'ƴ': 'y', 'ɓ': 'b', 'ɗ': 'd', 'ƙ': 'k', 'ς': 'σ',
    'қ': 'к', 'ғ': 'г', 'ҳ': 'х', 'ң': 'н', 'ө': 'о', 'ү': 'у', 'ұ': 'у', 'ә': 'а', 'і': 'и', 'һ': 'х', 'ҷ': 'ч',
    'ҙ': 'з', 'ҫ': 'с', 'є': 'е', 'ґ': 'г', 'ҕ': 'г', 'ҝ': 'к', 'ҡ': 'к', 'ҭ': 'т', 'ӈ': 'н', 'ҥ': 'н', 'ӄ': 'к',
    'ӆ': 'л', 'ӊ': 'н', 'ӎ': 'м', 'ӌ': 'ч', 'ҹ': 'ч', 'ҵ': 'ц', 'ѣ': 'е', 'ѳ': 'ф', 'ѵ': 'и', 'ӏ': '', 'э': 'е', 'ъ': '', 'ь': '',
    'ٱ': 'ا', 'ى': 'ي', 'ی': 'ي', 'ے': 'ي', 'ې': 'ي', 'ۍ': 'ي', 'ێ': 'ي', 'ة': 'ه', 'ہ': 'ه', 'ۀ': 'ه', 'ھ': 'ه',
    'ە': 'ه', 'ۃ': 'ه', 'ک': 'ك', 'ڪ': 'ك', 'ؤ': 'و', 'ۅ': 'و', 'ۆ': 'و', 'ۇ': 'و', 'ۈ': 'و', 'ۉ': 'و', 'ۋ': 'و', 'ں': 'ن'
  };
  (function () { for (var d = 0; d < 10; d++) { FOLD[AR_DIGITS[d]] = String(d); FOLD['۰۱۲۳۴۵۶۷۸۹'[d]] = String(d); } })();

  function inRanges(r, c) {
    for (var i = 0; i < r.length; i += 2) { if (c < r[i]) return false; if (c <= r[i + 1]) return true; }
    return false;
  }
  function scriptOf(c) {
    if (c < 0x80) return c >= 0x30 && c <= 0x39 ? 'Zyyy' : ((c | 0x20) >= 0x61 && (c | 0x20) <= 0x7A ? 'Latn' : '');
    if (inRanges(SEPS, c)) return '';
    var lo = 0, hi = SCRIPTS.length - 1;
    while (lo <= hi) {
      var m = (lo + hi) >> 1, r = SCRIPTS[m];
      if (c < r[0]) hi = m - 1; else if (c > r[1]) lo = m + 1; else return r[2];
    }
    return 'Zzzz';
  }
  function fromCp(c) {
    if (c < 0x10000) return String.fromCharCode(c);
    c -= 0x10000; return String.fromCharCode(0xD800 + (c >> 10), 0xDC00 + (c & 0x3FF));
  }
  var hasNormalize = typeof ''.normalize === 'function';
  /* per code point: [text, kind, script]; kind: 's' separator, 'x' skipped, 'a' apostrophe, 'm' mark, 'w' word, 'p' pairs */
  var cpMemo = {};
  function cpInfo(c) {
    var e = cpMemo[c];
    if (e) return e;
    if (SKIP[c]) e = ['', 'x', ''];
    else if (APOS[c]) e = ['', 'a', ''];
    else if (inRanges(MARK, c)) e = [fromCp(c), 'm', ''];
    else {
      var ch = fromCp(c), d = ch;
      if (hasNormalize && c >= 0x80 && !(c >= 0xAC00 && c <= 0xD7A3)) d = ch.normalize('NFKD');
      d = d.toLowerCase();
      var out = '', sc = '', drop = true, words = [];
      for (var i = 0; i < d.length; i++) {
        var u = d.charCodeAt(i);
        // a ligature of words (ﷺ = صلى الله عليه وسلم, ﷻ, ﷽): the same words as when they are typed one by one
        if (u === 0x20) { if (out) words.push(out); out = ''; continue; }
        if (u >= 0xD800 && u < 0xDC00 && i + 1 < d.length) { out += d.substr(i, 2); i++; if (!sc) sc = 'Hani'; continue; }
        if (inRanges(MARK, u)) { if (!drop) out += d[i]; continue; }
        if (APOS[u] || SKIP[u]) continue;
        var s2 = scriptOf(u);
        if (!s2) continue;
        if (!sc) { sc = s2; drop = !!DROP_MARKS[sc]; }
        var f = FOLD[d[i]];
        out += f != null ? f : d[i];
      }
      if (out && words.length) words.push(out);
      if (!sc) e = ['', 's', ''];                                    // punctuation, symbols, spaces
      else if (words.length > 1) e = [words, 'n', sc];
      else if (words.length) e = [words[0], PAIRS[sc] ? 'p' : 'w', sc];
      else if (!out) e = ['', 'x', ''];
      else e = [out, PAIRS[sc] ? 'p' : 'w', sc];
    }
    cpMemo[c] = e;
    return e;
  }

  /* s -> tokens [{t, pairs, a: [orig start per char], b: [orig end per char]}] (a, b only with pos) */
  function tokenize(s, pos) {
    s = String(s == null ? '' : s);
    var toks = [], tok = '', kind = '', sc = '', a = null, b = null, i = 0, n = 1;
    function end() {
      if (tok) toks.push(pos ? { t: tok, pairs: kind === 'p', a: a, b: b } : { t: tok, pairs: kind === 'p' });
      tok = ''; kind = ''; sc = ''; if (pos) { a = []; b = []; }
    }
    function add(t) {
      tok += t;
      if (pos) for (var j = 0; j < t.length; j++) { a.push(i); b.push(i + n); }
    }
    if (pos) { a = []; b = []; }
    while (i < s.length) {
      var c = s.charCodeAt(i);
      n = 1;
      if (c >= 0xD800 && c < 0xDC00 && i + 1 < s.length) {
        var lo = s.charCodeAt(i + 1);
        if (lo >= 0xDC00 && lo < 0xE000) { c = 0x10000 + ((c - 0xD800) << 10) + (lo - 0xDC00); n = 2; }
      }
      var e = cpMemo[c] || cpInfo(c), k = e[1];
      if (k === 's') end();
      else if (k === 'a') { if (ELISION[tok] || tok.length > 2 && tok.slice(-2) === 'qu') end(); }
      else if (k === 'm') { if (tok && !DROP_MARKS[sc]) add(e[0]); }
      else if (k === 'n') {                            // several words in one character (ﷺ): each one a token
        end();
        for (var w = 0; w < e[0].length; w++) { kind = 'w'; sc = e[2]; add(e[0][w]); end(); }
      }
      else if (k !== 'x') {
        if (tok && (k !== kind || k === 'p' && e[2] !== sc)) end();
        kind = k; sc = e[2];
        add(e[0]);
      }
      i += n;
    }
    end();
    return toks;
  }

  function cps(t) { return t.match(/[\uD800-\uDBFF][\uDC00-\uDFFF]|[\s\S]/g) || []; }
  function pairsOf(t) {
    var c = cps(t), out = [];
    if (c.length === 1) return c;
    for (var i = 0; i + 1 < c.length; i++) out.push(c[i] + c[i + 1]);
    return out;
  }

  /* Arabic word without the attached و ف, ب ك ل, ال (the variant used for Arabic-language sources), then
     without alef and doubled letters (level 2 of normArabic): والرحمن, الرحمان, رحمن -> رحمن */
  function arStem(w) {
    var r = w;
    if ((r[0] === 'و' || r[0] === 'ف') && r.length >= 4) r = r.slice(1);
    if (r.length >= 5 && (r.slice(0, 3) === 'بال' || r.slice(0, 3) === 'كال')) r = r.slice(3);
    else if (r.length >= 4 && (r.slice(0, 2) === 'ال' || r.slice(0, 2) === 'لل')) r = r.slice(2);
    else if (r.length >= 5 && (r[0] === 'ب' || r[0] === 'ل' || r[0] === 'ك')) r = r.slice(1);
    var s = r.replace(/ا/g, '').replace(/(.)\1+/g, '$1');
    return s || w.replace(/(.)\1+/g, '$1');
  }
  /* the light stem: only without و ف and alef / doubled letters (the article and ب ك ل kept: المسجد is not مسجدا);
     the word pairs of the index (web/search.js Builder.addPairs) */
  function arLight(w) {
    var r = (w[0] === 'و' || w[0] === 'ف') && w.length >= 4 ? w.slice(1) : w;
    var s = r.replace(/ا/g, '').replace(/(.)\1+/g, '$1');
    return s || w.replace(/(.)\1+/g, '$1');
  }
  function isArabTok(t) { var c = t.charCodeAt(0); return c >= 0x620 && c <= 0x6FF || c >= 0x750 && c <= 0x77F || c >= 0x8A0 && c <= 0x8FF; }

  /* The script of a text: the one with the most letters (Hani covers Chinese and Japanese kana). */
  function detectScript(s) {
    s = String(s || '');
    var count = {}, best = '', n = 0;
    for (var i = 0; i < s.length && i < 400; i++) {
      var sc = scriptOf(s.charCodeAt(i));
      if (!sc || sc === 'Zyyy' || inRanges(MARK, s.charCodeAt(i))) continue;
      count[sc] = (count[sc] || 0) + 1;
      if (count[sc] > n) { n = count[sc]; best = sc; }
    }
    return best;
  }
  /* Normalised text: tokens separated by single spaces (pair-script runs stay whole). */
  function norm(script, s) {
    if (s === undefined) { s = script; }
    return tokenize(s).map(function (t) { return t.t; }).join(' ');
  }
  /* Index terms of a text, in order. opts.stem: the Arabic-language variant of Arabic-script words (arStem). */
  function terms(s, opts) {
    var stem = opts && opts.stem, out = [];
    tokenize(s).forEach(function (t) {
      if (t.pairs) out.push.apply(out, pairsOf(t.t));
      else out.push(stem && isArabTok(t.t) ? arStem(t.t) : t.t);
    });
    return out;
  }
  /* Query words: [{t, stem, stems, lights, pairs: [..] | null}] (t: the term in other sources, stem: in Arabic-language
     ones; stems: stem and the stems of the word with the other seat of a hamza typed in it, ؤ <-> ئ: مسؤول = مسئول;
     lights: the same for arLight) */
  var SEAT_MAX = 4;
  function seatAlts(raw, t) {
    var out = [t.t];
    for (var j = 0; j < t.t.length && out.length < SEAT_MAX; j++) {
      var o = raw.charAt(t.a[j]), next = raw.charAt(t.b[j]), c = t.t.charAt(j);
      if (c !== 'و' && c !== 'ي') continue;
      if (!(o === 'ؤ' || o === 'ئ' || next === '\u0654' && (o === 'و' || o === 'ي' || o === 'ى'))) continue;
      var to = c === 'و' ? 'ي' : 'و';
      out = out.concat(out.map(function (x) { return x.slice(0, j) + to + x.slice(j + 1); })).slice(0, SEAT_MAX);
    }
    return out;
  }
  function queryTerms(s) {
    var raw = String(s == null ? '' : s);
    return tokenize(raw, true).map(function (t) {
      if (t.pairs) return { t: t.t, stem: t.t, stems: [t.t], lights: [t.t], pairs: pairsOf(t.t) };
      if (!isArabTok(t.t)) return { t: t.t, stem: t.t, stems: [t.t], lights: [t.t], pairs: null };
      var stems = [], lights = [];
      seatAlts(raw, t).forEach(function (x) {
        var st = arStem(x), li = arLight(x);
        if (stems.indexOf(st) < 0) stems.push(st);
        if (lights.indexOf(li) < 0) lights.push(li);
      });
      return { t: t.t, stem: stems[0], stems: stems, lights: lights, pairs: null };
    });
  }

  /* Latin <-> Cyrillic (Uzbek, Kazakh, Tajik, Serbian ... are written in both): an extra variant of the query */
  var L2C = [['shch', 'щ'], ["yo'", 'йў'], ['yoʻ', 'йў'], ['yo‘', 'йў'], ['yo’', 'йў'], ["o'", 'ў'], ['oʻ', 'ў'], ['o‘', 'ў'], ['o’', 'ў'], ["g'", 'ғ'], ['gʻ', 'ғ'], ['g‘', 'ғ'], ['g’', 'ғ'],
    ['sh', 'ш'], ['ch', 'ч'], ['zh', 'ж'], ['kh', 'х'], ['ya', 'я'], ['yo', 'ё'], ['yu', 'ю'], ['ye', 'е'], ['ts', 'ц'],
    ["oʼ", 'ў'], ["gʼ", 'ғ'],
    ['ş', 'ш'], ['ç', 'ч'], ['ğ', 'ғ'], ['ı', 'ы'], ['ö', 'ў'], ['ü', 'ү'], ['ä', 'ә'], ['ñ', 'ң'], ['ū', 'ұ'], ['š', 'ш'],
    ['č', 'ч'], ['ž', 'ж'], ['ć', 'ч'], ['đ', 'ђ'],
    ['a', 'а'], ['b', 'б'], ['c', 'с'], ['d', 'д'], ['e', 'е'], ['f', 'ф'], ['g', 'г'], ['h', 'ҳ'], ['i', 'и'], ['j', 'ж'],
    ['k', 'к'], ['l', 'л'], ['m', 'м'], ['n', 'н'], ['o', 'о'], ['p', 'п'], ['q', 'қ'], ['r', 'р'], ['s', 'с'], ['t', 'т'],
    ['u', 'у'], ['v', 'в'], ['w', 'в'], ['x', 'х'], ['y', 'й'], ['z', 'з']];
  var C2L = { 'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo', 'ж': 'j', 'з': 'z', 'и': 'i', 'й': 'y',
    'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'x',
    'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'sh', 'ъ': "'", 'ы': 'i', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya', 'ў': "o'",
    'қ': 'q', 'ғ': "g'", 'ҳ': 'h', 'ң': 'ng', 'ө': 'o', 'ү': 'u', 'ұ': 'u', 'ә': 'a', 'і': 'i', 'һ': 'h', 'ҷ': 'j',
    'ӣ': 'i', 'ӯ': 'u', 'ј': 'j', 'љ': 'lj', 'њ': 'nj', 'ћ': 'c', 'ђ': 'dj', 'џ': 'dz', 'є': 'ye', 'ї': 'yi', 'ґ': 'g' };
  function latToCyr(s) {
    s = String(s || '').toLowerCase();
    var out = '';
    for (var i = 0; i < s.length;) {
      var hit = null;
      for (var k = 0; k < L2C.length; k++) if (s.substr(i, L2C[k][0].length) === L2C[k][0]) { hit = L2C[k]; break; }
      if (hit) { out += hit[1]; i += hit[0].length; } else { out += s[i]; i++; }
    }
    return out;
  }
  /* the normalised Cyrillic word of a Latin query word, with what its Latin letters may also stand for: sh = ш or сҳ
     (ashob = асҳоб, Ishoq = Исҳоқ), ts = ц or тс (otsiz = отсиз); at most 8 forms, the first one = the word itself */
  function latAlts(w) {
    var out = [w];
    for (var i = 0; i < w.length && out.length < 8; i++) {
      var to = w[i] === 'ш' ? 'сх' : w[i] === 'ц' ? 'тс' : null;
      if (!to) continue;
      out = out.concat(out.map(function (x) {
        var j = x.length - (w.length - i);                  // the same letter in a form changed before (from the end)
        return x.slice(0, j) + to + x.slice(j + 1);
      })).slice(0, 8);
    }
    return out;
  }
  function cyrToLat(s) {
    s = String(s || '').toLowerCase();
    var out = '';
    for (var i = 0; i < s.length; i++) {
      var c = s[i], l = C2L[c];
      if (c === 'е' && (i === 0 || /[\s\-аеёиоуўэюяы]/.test(s[i - 1]))) l = 'ye';
      out += l != null ? l : c;
    }
    return out;
  }

  /* Uzbek inflection (the folded Cyrillic of norm: қ ғ ҳ ў й ё э -> к г х у и е е, ъ ь dropped). The index keeps whole
     words; a query word also finds its forms with the suffixes of nouns: намоз -> намознинг, намози; калити -> калит,
     калитлари, калитини. uzBase(w): w without its suffixes (the shortest base of at least 3 letters, else w);
     uzInflected(base, t): t is base + a chain of at most 4 suffixes (the possessive ones after a vowel / a consonant).
     partial: the word is being typed, its last suffix may be cut (ниятла -> ният).                                    */
  var UZ_SUF = ['лар', 'нинг', 'ни', 'га', 'ка', 'да', 'та', 'дан', 'тан', 'гача', 'кача', 'даги', 'таги', 'ники', 'дек',
    'дай', 'ча', 'дир', 'ман', 'сан', 'миз', 'сиз', 'ли', 'лик', 'лиг', 'чи', 'ми', 'ку', 'на', 'нда', 'ндан', 'гина',
    'кина', 'ги', 'ки', 'у', 'ю'];
  var UZ_SUF_V = ['м', 'нг', 'миз', 'нгиз', 'си'], UZ_SUF_C = ['им', 'инг', 'и', 'имиз', 'ингиз'];
  var UZ_VOWEL = { 'а': 1, 'е': 1, 'и': 1, 'о': 1, 'у': 1, 'ю': 1, 'я': 1 };
  function isCyrTok(t) { var c = t.charCodeAt(0); return c >= 0x400 && c <= 0x52F; }
  function uzRest(prev, rest, depth, partial) {      // partial: rest may end inside a suffix (a word being typed)
    if (!rest) return true;
    if (depth >= 4) return false;
    var lists = [UZ_SUF, UZ_VOWEL[prev] ? UZ_SUF_V : UZ_SUF_C];
    for (var l = 0; l < 2; l++) {
      for (var i = 0; i < lists[l].length; i++) {
        var s = lists[l][i];
        if (partial && s.length > rest.length && s.slice(0, rest.length) === rest) return true;
        if (rest.slice(0, s.length) === s && uzRest(s[s.length - 1], rest.slice(s.length), depth + 1, partial)) return true;
      }
    }
    return false;
  }
  function uzInflected(base, t, partial) {
    return t.length > base.length && t.slice(0, base.length) === base && uzRest(base[base.length - 1], t.slice(base.length), 0, partial);
  }
  var baseMemo = {};
  function uzBase(w, partial) {
    if (w.length < 4 || !isCyrTok(w)) return w;
    var mk = (partial ? '+' : '') + w;
    if (baseMemo[mk]) return baseMemo[mk];
    var out = w;
    for (var k = 3; k < w.length; k++) if (uzInflected(w.slice(0, k), w, partial)) { out = w.slice(0, k); break; }
    return (baseMemo[mk] = out);
  }

  /* Levenshtein distance, max+1 as soon as it is over max */
  function levenshtein(a, b, max) {
    if (Math.abs(a.length - b.length) > max) return max + 1;
    var prev = [], cur = [], i, j;
    for (j = 0; j <= b.length; j++) prev[j] = j;
    for (i = 1; i <= a.length; i++) {
      cur[0] = i; var low = i;
      for (j = 1; j <= b.length; j++) {
        cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
        if (cur[j] < low) low = cur[j];
      }
      if (low > max) return max + 1;
      var t = prev; prev = cur; cur = t;
    }
    return prev[b.length];
  }

  /* ---------- the styles of the Uzbek Latin: one canonical text, the others made of it by a table of letters ----------
     The canonical Latin (the official alphabet of 1995: oʻ gʻ with ʻ U+02BB, the tutuq ʼ U+02BC; the Latin texts of
     the hadiths, tools/build_latin.js, and the interface uz_latn) -> latinStyle(s, {nw, apos}):
       nw: the 2026 Latin (the law OʻRQ-1184, as Harfsoz does it: src/engine.js markedOldToNew): sh -> ş, ch -> ç,
           oʻ -> ö, gʻ -> ğ, sʼh (с + ҳ: Isʼhoq) -> sh, any other ʻ -> ʼ; the letter case kept (Sh, SH -> Ş);
       apos: 'mod' (ʻ ʼ, as they are), 'typo' (‘ ’, the quotes of Word), 'ascii' (' for both).
     Left as they are: links, e-mails, domains, /commands and /paths, {placeholders}, foreign words (w, c without h, a few names: Chrome).
     latinStyleHtml: the same for the text between the tags of HTML.                                              */
  var OKINA = { 'ʻ': 1, '‘': 1, "'": 1, '`': 1 }, TUTUQ = { 'ʼ': 1, '’': 1, "'": 1 };
  var KEEP_LAT = /^(chrome|shift|scheherazade|microsoft|bukhari|sahih|cheap|search)$/i;
  var SKIP_TOK = /:\/\/|^www\.|^\/|@|[\w-]\.(?:org|uz|com|net|ru|io|app|dev|info)\b|\{\w+\}/i;
  function newWord(w) {
    var r = '';
    for (var i = 0; i < w.length; i++) {
      var c = w[i], n = w[i + 1] || '', n2 = w[i + 2] || '', lc = c.toLowerCase(), ln = n.toLowerCase();
      if (lc === 's' && TUTUQ[n] && n2.toLowerCase() === 'h') { r += c + n2; i += 2; continue; }
      if (lc === 's' && ln === 'h') { r += c === 'S' ? 'Ş' : 'ş'; i++; continue; }
      if (lc === 'c' && ln === 'h') { r += c === 'C' ? 'Ç' : 'ç'; i++; continue; }
      if ((lc === 'o' || lc === 'g') && OKINA[n]) { r += lc === 'o' ? (c === 'O' ? 'Ö' : 'ö') : (c === 'G' ? 'Ğ' : 'ğ'); i++; continue; }
      if ((c === 'ʻ' || c === '‘') && i > 0) { r += 'ʼ'; continue; }      // not an opening quote ‘…’
      r += c;
    }
    return r;
  }
  function latinStyle(s, o) {
    s = String(s == null ? '' : s);
    var nw = !!(o && o.nw), apos = o && o.apos || 'mod';
    if (!nw && apos === 'mod') return s;
    return s.replace(/[^\s{}]+|\{\w+\}/g, function (tok) {
      if (!/[A-Za-zʻʼ‘’]/.test(tok) || SKIP_TOK.test(tok)) return tok;
      if (nw) tok = tok.replace(/[A-Za-zʻʼ‘’'`]+/g, function (w) {
        var letters = w.replace(/[^A-Za-z]/g, '');
        return /[wW]|[cC](?![hH])/.test(letters) || KEEP_LAT.test(letters) ? w : newWord(w);
      });
      if (apos === 'typo') tok = tok.replace(/ʻ/g, '‘').replace(/ʼ/g, '’');
      else if (apos === 'ascii') tok = tok.replace(/[ʻʼ]/g, "'");
      return tok;
    });
  }
  function latinStyleHtml(html, o) {
    if (!(o && (o.nw || o.apos && o.apos !== 'mod'))) return String(html == null ? '' : html);
    return String(html == null ? '' : html).split(/(<[^>]*>|&[#\w]+;)/).map(function (part, i) {
      return i % 2 ? part : latinStyle(part, o);
    }).join('');
  }

  /* HTML -> text (tags dropped, the usual entities decoded) */
  function stripHtml(t) {
    return String(t || '').replace(/<sup[^>]*>[\s\S]*?<\/sup>/g, '').replace(/<a[^>]*>[\s\S]*?<\/a>/g, '')
      .replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  }

  return { norm: norm, terms: terms, queryTerms: queryTerms, tokenize: tokenize, detectScript: detectScript,
           scriptOf: scriptOf, arStem: arStem, arLight: arLight, latToCyr: latToCyr, cyrToLat: cyrToLat, levenshtein: levenshtein,
           stripHtml: stripHtml, uzBase: uzBase, uzInflected: uzInflected, latAlts: latAlts,
           latinStyle: latinStyle, latinStyleHtml: latinStyleHtml };
});
