/* Shared logic of the hadith panel: the website, the Word add-in, the extension, Google Docs, WordPress and Node (tests).
   No DOM. The texts of the hadiths come only from the data (web/h, web/data): nothing here adds a word to a hadith.

     HadisCore.BOOKS                         the 11 books: names in uz / uz_latn / ar (+ aliases for the parser)
     HadisCore.bookName(id, lang, kind)      kind: 'short' (Бухорий) | 'book' (Саҳиҳул Бухорий) | 'author' (Имом Бухорий)
     HadisCore.parseRef(q)                   «Бухорий 1», «Muslim 2564/2», «بخاري ١», «Б 1», «Термизий 5-7» -> {book, n, s?, to?} | null
     HadisCore.Refs(json)                    data/refs/<book>.json: number <-> doc
     HadisCore.parseHadith(h, mode)          {ar: {paras: [[{t, matn}]]}, uz: {paras: [[{t, b, i}]]}} (matn: «» / <b>); mode
                                             'sahabi' | 'matn': from the Companion (data `st`, tools/sahabi.js), 'full' / none: all
     HadisCore.MODES                         'sahabi' (default: from the Companion), 'full' (the whole isnad), 'matn'
     HadisCore.words(paras, mode)            the words of a text for the selection; mode 'matn': only the words of the Prophet ﷺ
     HadisCore.ranges(r, n)                  the chosen words ([from, to] or a list of parts) in order, joined, inside 0..n-1
     HadisCore.BOOK_ORDER, bookRank(id)      the order of the books in the results of a search and in the lists
     HadisCore.tahrij(h, book, lang, opts)   «(«Саҳиҳи Бухорий», 1)», «(«Абу Довуд», 2226; Ҳадиснинг даражаси: «саҳиҳ». Манба:
                                             «Саҳиҳи Абу Довуд», 2226)» (opts.grades), «(رواه البخاري ١)»
     HadisCore.tahrijGrades(h, book, lang, o) the gradings of a tahrij, the main one first (o.all: every one)
     HadisCore.emphasis(h)                   what is bold: the words of the Prophet ﷺ, else of the Companion (athar)
     HadisCore.fixQuotes(s)                  “ ” and paired " -> « »
     HadisCore.hukmOf(h, book, lang)         the words of the author of the book from the field `hukm` only (who said it, the quote)
     HadisCore.gradesOf(h, book, lang)       the gradings of other scholars from the field `hukm2` (Албоний, Арнаут, Заҳабий, Хайсамий …):
                                             who, the grade, the work, the number in it; never mixed with hukm (the author)
     HadisCore.format(h, book, opts)         {paras: [{dir, kind, runs: [{t, b, i, role}]}], text}
     HadisCore.Data(o)                       the client of the data: current.json, books.json, refs, blocks of 100 (cache, retries, mirrors)
     HadisCore.toLatin(s)                    Uzbek Cyrillic -> Latin (the case kept)
     HadisCore.quote(text), .compareQuote(q, h), .findQuote(q, search, hadith)
                                             «check a quote»: a pasted text against a hadith, word by word  */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./text.js'), require('./search.js'));
  else root.HadisCore = factory(root.HadisText, root.HadisSearch);
})(typeof self !== 'undefined' ? self : this, function (T, S) {
  'use strict';

  var AR_DIGITS = '٠١٢٣٤٥٦٧٨٩';
  function arNum(n) { return String(n).replace(/[0-9]/g, function (d) { return AR_DIGITS[+d]; }); }
  function latinDigits(s) {
    return String(s).replace(/[٠-٩]/g, function (d) { return AR_DIGITS.indexOf(d); })
      .replace(/[۰-۹]/g, function (d) { return '۰۱۲۳۴۵۶۷۸۹'.indexOf(d); });
  }

  /* ---------- the books ----------
     short: the name used in a reference (the author); book: the title of the collection (null: not known yet, see
     docs/data_report.md, Байҳақий and Табароний); author: with «Имом» where it is usual. aliases: other spellings for the
     parser (folded by fold(), so case, ҳ/х, ʻ/' and the Arabic vowel marks do not matter); abbr: the short signs.
     Languages: uz (Cyrillic), uz_latn, ar; a language that is missing falls back to uz. */
  var BOOKS = [
    { id: 'bukhari', short: { uz: 'Бухорий', uz_latn: 'Buxoriy', ar: 'البخاري' },
      book: { uz: 'Саҳиҳул Бухорий', uz_latn: 'Sahihul Buxoriy', ar: 'صحيح البخاري' },
      author: { uz: 'Имом Бухорий', uz_latn: 'Imom Buxoriy', ar: 'البخاري' },
      aliases: ['бухорий', 'бухори', 'бухари', 'бухарий', 'buxoriy', 'buxori', 'bukhari', 'bukhariy', 'buhari', 'boukhari', 'بخاري',
                'саҳиҳул бухорий', 'саҳиҳи бухорий', 'сахих бухари', 'sahih bukhari', 'sahihul buxoriy', 'صحيح بخاري'],
      abbr: ['б', 'b', 'خ'], sahih: true },
    { id: 'muslim', short: { uz: 'Муслим', uz_latn: 'Muslim', ar: 'مسلم' },
      book: { uz: 'Саҳиҳи Муслим', uz_latn: 'Sahihi Muslim', ar: 'صحيح مسلم' },
      author: { uz: 'Имом Муслим', uz_latn: 'Imom Muslim', ar: 'مسلم' },
      aliases: ['муслим', 'muslim', 'مسلم', 'саҳиҳи муслим', 'сахих муслим', 'sahih muslim', 'sahihi muslim', 'صحيح مسلم'],
      abbr: ['м', 'm', 'م'], sahih: true },
    { id: 'abudawud', short: { uz: 'Абу Довуд', uz_latn: 'Abu Dovud', ar: 'أبو داود' },
      book: { uz: 'Сунани Абу Довуд', uz_latn: 'Sunani Abu Dovud', ar: 'سنن أبي داود' },
      author: { uz: 'Абу Довуд', uz_latn: 'Abu Dovud', ar: 'أبو داود' },
      aliases: ['абу довуд', 'абу давуд', 'абу дауд', 'абудовуд', 'abu dovud', 'abu dawud', 'abu dawood', 'abu daud', 'abu davud',
                'abudovud', 'abudawud', 'ابو داود', 'ابي داود', 'سنن ابي داود', 'сунани абу довуд', 'sunan abi dawud'],
      abbr: ['ад', 'ad', 'د'] },
    { id: 'tirmidhi', short: { uz: 'Термизий', uz_latn: 'Termiziy', ar: 'الترمذي' },
      book: { uz: 'Сунани Термизий', uz_latn: 'Sunani Termiziy', ar: 'سنن الترمذي' },
      author: { uz: 'Имом Термизий', uz_latn: 'Imom Termiziy', ar: 'الترمذي' },
      aliases: ['термизий', 'термизи', 'тирмизи', 'тирмизий', 'termiziy', 'termizi', 'tirmiziy', 'tirmizi', 'tirmidhi', 'tirmidi',
                'ترمذي', 'سنن ترمذي', 'جامع ترمذي', 'сунани термизий', 'jami tirmidhi', 'sunan tirmidhi'],
      abbr: ['т', 't', 'ت'] },
    { id: 'ibnmajah', short: { uz: 'Ибн Можа', uz_latn: 'Ibn Moja', ar: 'ابن ماجه' },
      book: { uz: 'Сунани Ибн Можа', uz_latn: 'Sunani Ibn Moja', ar: 'سنن ابن ماجه' },
      author: { uz: 'Ибн Можа', uz_latn: 'Ibn Moja', ar: 'ابن ماجه' },
      aliases: ['ибн можа', 'ибн можжа', 'ибн маджа', 'ибн маджах', 'ibn moja', 'ibn majah', 'ibn maja', 'ibn majjah', 'ibnmajah',
                'ibnmoja', 'ибнможа', 'ابن ماجه', 'سنن ابن ماجه', 'сунани ибн можа', 'sunan ibn majah'],
      abbr: ['им', 'im', 'جه', 'ه'] },
    { id: 'nasai', short: { uz: 'Насоий', uz_latn: 'Nasoiy', ar: 'النسائي' },
      book: { uz: 'Сунани Насоий', uz_latn: 'Sunani Nasoiy', ar: 'سنن النسائي' },
      author: { uz: 'Имом Насоий', uz_latn: 'Imom Nasoiy', ar: 'النسائي' },
      aliases: ['насоий', 'насои', 'насаи', 'насаий', 'nasoiy', 'nasoi', 'nasai', 'nasaee', 'nasaiy', 'نسايي', 'نساي', 'سنن نسايي',
                'сунани насоий', 'sunan nasai'],
      abbr: ['н', 'n', 'ن', 'س'] },
    { id: 'ahmad', short: { uz: 'Аҳмад', uz_latn: 'Ahmad', ar: 'أحمد' },
      book: { uz: 'Муснади Аҳмад', uz_latn: 'Musnadi Ahmad', ar: 'مسند أحمد' },
      author: { uz: 'Имом Аҳмад', uz_latn: 'Imom Ahmad', ar: 'أحمد' },
      aliases: ['аҳмад', 'ахмад', 'ahmad', 'ahmed', 'احمد', 'مسند احمد', 'муснад', 'муснади аҳмад', 'муснад аҳмад', 'муснад ахмада',
                'musnad', 'musnadi ahmad', 'musnad ahmad', 'احمد بن حنبل', 'ибн ҳанбал', 'ibn hanbal'],
      abbr: ['а', 'a', 'حم'] },
    { id: 'hakim', short: { uz: 'Ҳоким', uz_latn: 'Hokim', ar: 'الحاكم' },
      book: { uz: 'Мустадрак', uz_latn: 'Mustadrak', ar: 'المستدرك على الصحيحين' },
      author: { uz: 'Ҳоким', uz_latn: 'Hokim', ar: 'الحاكم' },
      aliases: ['ҳоким', 'хоким', 'хаким', 'ҳаким', 'hokim', 'hakim', 'حاكم', 'мустадрак', 'mustadrak', 'مستدرك', 'مستدرك علي صحيحين'],
      abbr: ['ҳк', 'hk', 'ك'] },
    { id: 'bayhaqi', short: { uz: 'Байҳақий', uz_latn: 'Bayhaqiy', ar: 'البيهقي' },
      book: null,
      author: { uz: 'Байҳақий', uz_latn: 'Bayhaqiy', ar: 'البيهقي' },
      aliases: ['байҳақий', 'байхакий', 'байхаки', 'байҳақи', 'bayhaqiy', 'bayhaqi', 'baihaqi', 'bayhaki', 'بيهقي'],
      abbr: ['бй', 'by', 'هق'] },
    { id: 'tabarani', short: { uz: 'Табароний', uz_latn: 'Tabaroniy', ar: 'الطبراني' },
      book: null,
      author: { uz: 'Табароний', uz_latn: 'Tabaroniy', ar: 'الطبراني' },
      aliases: ['табароний', 'табарони', 'табарани', 'tabaroniy', 'tabaroni', 'tabarani', 'tabaraani', 'طبراني'],
      abbr: ['тб', 'tb', 'طب'] },
    { id: 'daraqutni', short: { uz: 'Дорақутний', uz_latn: 'Doraqutniy', ar: 'الدارقطني' },
      book: { uz: 'Сунани Дорақутний', uz_latn: 'Sunani Doraqutniy', ar: 'سنن الدارقطني' },
      author: { uz: 'Дорақутний', uz_latn: 'Doraqutniy', ar: 'الدارقطني' },
      aliases: ['дорақутний', 'доракутний', 'даракутни', 'дорақутни', 'doraqutniy', 'doraqutni', 'daraqutni', 'daraqutniy', 'darakutni',
                'دارقطني', 'سنن دارقطني'],
      abbr: ['дқ', 'dq', 'قط'] }
  ];
  var BY_ID = {};
  BOOKS.forEach(function (b, i) { b.order = i; BY_ID[b.id] = b; });

  function pick(obj, lang) {
    if (!obj) return null;
    if (obj[lang] != null) return obj[lang];
    var base = String(lang || '').split('_')[0];
    return obj[base] != null ? obj[base] : obj.uz;
  }
  /* kind: 'short' | 'book' | 'author'; a book without a known title ('book') gives its author's name */
  function bookName(id, lang, kind) {
    var b = BY_ID[id]; if (!b) return id;
    var o = kind === 'book' ? b.book || b.short : kind === 'author' ? b.author : b.short;
    return pick(o, lang || 'uz');
  }

  /* ---------- Uzbek Cyrillic -> Latin (the official alphabet; the case kept; ʻ U+02BB, ʼ U+02BC) ---------- */
  var C2L = { 'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo', 'ж': 'j', 'з': 'z', 'и': 'i', 'й': 'y',
    'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'x',
    'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'sh', 'ъ': 'ʼ', 'ы': 'i', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya', 'ў': 'oʻ',
    'қ': 'q', 'ғ': 'gʻ', 'ҳ': 'h' };
  var VOWEL_BEFORE = /[аеёиоуўэюяaeiou]/i;
  function isCyrUpper(c) { return c !== c.toLowerCase() && /[Ѐ-ӿ]/.test(c); }
  function toLatin(s, prev) {
    s = String(s == null ? '' : s);
    var out = '', last = prev || '';
    for (var i = 0; i < s.length; i++) {
      var c = s[i], lo = c.toLowerCase(), l = C2L[lo];
      if (l == null) { out += c; last = c; continue; }
      // е, ц: «ye» at the start of a word and after a vowel / ъ; «s» for ц between letters is not forced (loan words keep ts)
      if (lo === 'е' && (!last || !/[Ѐ-ӿa-z]/i.test(last) || VOWEL_BEFORE.test(last) || /[ъь]/i.test(last))) l = 'ye';
      if (isCyrUpper(c) && l) {
        var nx = s[i + 1] || '', nextUpper = isCyrUpper(nx) || (!/[Ѐ-ӿ]/.test(nx) && isCyrUpper(last));
        l = nextUpper ? l.toUpperCase() : l.charAt(0).toUpperCase() + l.slice(1);
      }
      out += l; last = c;
    }
    return out;
  }

  /* ---------- the reference parser ---------- */
  var HARAKAT = /[ؐ-ًؚ-ٰٟۖ-ۭـ]/g;
  function fold(s) {
    s = String(s || '').toLowerCase();
    return (typeof s.normalize === 'function' ? s.normalize('NFC') : s)          // an old webview (IE11) has no normalize
      .replace(HARAKAT, '').replace(/[أإآٱ]/g, 'ا').replace(/[ىی]/g, 'ي').replace(/ة/g, 'ه').replace(/ک/g, 'ك').replace(/[ئؤء]/g, '')
      .replace(/(^|\s)ال(?=\S)/g, '$1')                              // البخاري -> بخاري
      .replace(/ҳ/g, 'х').replace(/қ/g, 'к').replace(/ғ/g, 'г').replace(/ў/g, 'у').replace(/ё/g, 'е').replace(/[ъь]/g, '')
      .replace(/[ʻʼ‘’'`ʻʼ]/g, '')
      .replace(/(^|\s)(аль|ал|ан|ат|ад|ас|ар|al|an|at|ad|as|ar|el|ash|ash)-/g, '$1')
      .replace(/[-–—_.,;:()«»"“”\[\]]+/g, ' ').replace(/\s+/g, ' ').trim();
  }
  var PREFIX = /^(имом|имам|imom|imam|رواه|اخرجه)\s+/;
  var SUFFIX = /\s+(ривояти|ривоят қилган|rivoyati|rivoyat qilgan|передал|narrated|رواه)$/;
  var ALIAS = null;
  function aliasTable() {
    if (ALIAS) return ALIAS;
    ALIAS = { exact: {}, list: [] };
    BOOKS.forEach(function (b) {
      var names = b.aliases.slice();
      ['short', 'book', 'author'].forEach(function (k) { if (b[k]) for (var l in b[k]) names.push(b[k][l]); });
      names.forEach(function (n) {
        var f = fold(n).replace(PREFIX, '');
        if (!f) return;
        if (!ALIAS.exact[f]) { ALIAS.exact[f] = b.id; ALIAS.list.push([f, b.id]); }
        var lat = fold(T.cyrToLat(f));                               // the other alphabet: «бухорий» -> «buxoriy»
        if (lat && !ALIAS.exact[lat]) { ALIAS.exact[lat] = b.id; ALIAS.list.push([lat, b.id]); }
      });
      b.abbr.forEach(function (a) { ALIAS.exact['#' + fold(a)] = b.id; });
    });
    return ALIAS;
  }
  /* the name of a book from what the person typed (null: not a book) */
  function findBook(name) {
    var f = fold(name).replace(PREFIX, '').replace(SUFFIX, '');
    if (!f) return null;
    var A = aliasTable();
    if (A.exact[f]) return A.exact[f];
    if (A.exact['#' + f]) return A.exact['#' + f];
    var w = f.split(' ');
    // «Саҳиҳул Бухорий», «Сунани Термизий»: the last word alone; the first one («Муслим ривояти»)
    if (w.length > 1 && A.exact[w[w.length - 1]] && !A.exact['#' + w[w.length - 1]]) return A.exact[w[w.length - 1]];
    if (w.length > 1 && A.exact[w[0]] && w[0].length > 3) return A.exact[w[0]];
    // the beginning of a name (at least 3 letters), when only one book has it: «Бух 1», «Tirm 5»
    if (f.length >= 3) {
      var hit = null;
      for (var i = 0; i < A.list.length; i++) {
        if (A.list[i][0].indexOf(f) === 0) { if (hit && hit !== A.list[i][1]) return null; hit = A.list[i][1]; }
      }
      return hit;
    }
    return null;
  }
  /* «Бухорий 1» · «Бухорий, 1-ҳадис» · «(Имом Бухорий ривояти, 1-ҳадис)» · «Muslim 2564/2» · «Муслим 2564:2» · «بخاري ١» ·
     «رواه البخاري (١)» · «Б 1» · «Термизий 5-7» (a run of numbers, at most 50) · «1 Бухорий».  -> {book, n, s?, to?} | null */
  function parseRef(q) {
    var s = latinDigits(String(q || '')).replace(/[‎‏‪-‮]/g, '').trim();
    if (!s || s.length > 80) return null;
    s = s.replace(/^[(\[]+|[)\].]+$/g, '').trim();
    s = s.replace(/[-\s]*(ҳадис|хадис|hadis|hadith|ҳадиси|hadisi|№|no\.?|nr\.?|#)\s*$/i, '').trim();
    var m = /^(.*?[^\d\s,:#№/-])[\s,:#№(]*?(?:№|no\.?|#)?\s*\(?\s*(\d{1,5})(?:\s*[/:.]\s*(\d{1,4}))?(?:\s*[-–—]\s*(\d{1,5}))?\s*(?:-?\s*(?:ҳадис|хадис|hadis|hadith))?\s*\)?$/i.exec(s);
    var name, n, sub, to;
    if (m) { name = m[1]; n = +m[2]; sub = m[3] ? +m[3] : null; to = m[4] ? +m[4] : null; }
    else {
      m = /^(\d{1,5})(?:\s*[/:]\s*(\d{1,4}))?\s+(.+)$/.exec(s);                       // «1 Бухорий»
      if (!m) return null;
      name = m[3]; n = +m[1]; sub = m[2] ? +m[2] : null; to = null;
    }
    var book = findBook(name);
    if (!book || !(n >= 1)) return null;
    var r = { book: book, n: n };
    if (sub) r.s = sub;
    if (to && to > n && !sub) r.to = Math.min(to, n + 49);
    return r;
  }

  /* ---------- refs: data/refs/<book>.json {book, doc0, count, n: [first, delta, …], s: {offset: sub}} ---------- */
  function Refs(j) {
    if (!(this instanceof Refs)) return new Refs(j);
    if (!j || !Array.isArray(j.n) || j.n.length !== j.count) throw new Error('refs: bad file');
    this.book = j.book; this.doc0 = j.doc0; this.count = j.count;
    var nums = new Int32Array(j.count), at = {}, v = 0;
    for (var i = 0; i < j.count; i++) {
      v = i ? v + j.n[i] : j.n[0]; nums[i] = v;
      var sub = j.s && j.s[i] ? +j.s[i] : 1, key = v + '/' + sub;
      if (at[key] == null) at[key] = i;
      if (at[v] == null) at[v] = i;                         // a number without its sub: the first hadith of it
    }
    this.nums = nums; this.subs = j.s || {}; this.at = at;
  }
  Refs.prototype.doc = function (n, s) {                     // -> doc | -1
    var k = s ? this.at[n + '/' + s] : this.at[n];
    return k == null ? -1 : this.doc0 + k;
  };
  Refs.prototype.ref = function (d) {                        // -> {n, s?} of a doc of this book
    var k = d - this.doc0;
    if (k < 0 || k >= this.count) return null;
    var r = { n: this.nums[k] }, s = this.subs[k];
    if (s && +s !== 1) r.s = +s;
    return r;
  };
  Refs.prototype.has = function (d) { return d >= this.doc0 && d < this.doc0 + this.count; };
  function refLabel(r) { return r ? r.n + (r.s ? '/' + r.s : '') : ''; }

  /* ---------- the text of a hadith: runs, matn, words ----------
     ar: plain text; the words of the Prophet ﷺ (matn) are in «» (77 thousand hadiths). Lines: '\n'.
     uz: HTML with <p> <b> <i> only (tools/build_data.py); <b> = the words of the Prophet ﷺ. */
  function decodeEnt(s) {
    return s.replace(/&(amp|lt|gt|quot|#39|#x27|nbsp);/g, function (m, e) {
      return { amp: '&', lt: '<', gt: '>', quot: '"', '#39': "'", '#x27': "'", nbsp: ' ' }[e];
    });
  }
  function parseUz(html) {
    var paras = [], cur = null, b = 0, it = 0, re = /<(\/?)(p|b|i|strong|em|br)\b[^>]*>|([^<]+)|(<)/gi, m;
    function para() { if (!cur) { cur = []; paras.push(cur); } return cur; }
    while ((m = re.exec(String(html || '')))) {
      if (m[3] != null || m[4]) {
        var t = decodeEnt(m[3] != null ? m[3] : '<');
        if (!t) continue;
        var p = para(), last = p[p.length - 1], bb = b > 0, ii = it > 0;
        if (last && last.b === bb && last.i === ii) last.t += t; else p.push({ t: t, b: bb, i: ii });
        continue;
      }
      var close = !!m[1], tag = m[2].toLowerCase();
      if (tag === 'p') { cur = null; if (!close) para(); }
      else if (tag === 'br') { cur = null; }
      else if (tag === 'b' || tag === 'strong') b = Math.max(0, b + (close ? -1 : 1));
      else it = Math.max(0, it + (close ? -1 : 1));
    }
    return paras.map(function (p) {                         // spaces at the ends of a paragraph go
      if (p.length) { p[0].t = p[0].t.replace(/^\s+/, ''); p[p.length - 1].t = p[p.length - 1].t.replace(/\s+$/, ''); }
      return p.filter(function (r) { return r.t; });
    }).filter(function (p) { return p.length; });
  }
  function parseAr(text) {
    var lines = String(text || '').split(/\n+/), depth = 0;
    return lines.map(function (line) {
      var runs = [], cur = null;
      for (var i = 0; i < line.length; i++) {
        var c = line[i], matn;
        if (c === '«') { depth++; matn = true; }
        else if (c === '»') { matn = depth > 0; depth = Math.max(0, depth - 1); }
        else matn = depth > 0;
        if (cur && cur.matn === matn) cur.t += c; else { cur = { t: c, matn: matn }; runs.push(cur); }
      }
      if (runs.length) { runs[0].t = runs[0].t.replace(/^\s+/, ''); runs[runs.length - 1].t = runs[runs.length - 1].t.replace(/\s+$/, ''); }
      return runs.filter(function (r) { return r.t; });
    }).filter(function (p) { return p.length; });
  }
  function parseHadith(h, mode) {
    var P = { ar: { paras: parseAr(h && h.ar) }, uz: { paras: parseUz(h && h.uz) },
              bob: h && h.bob ? parseUz(h.bob) : [], bobAr: h && h.bob_ar ? parseAr(h.bob_ar) : [] };
    var st = h && h.st;
    if (st && mode && mode !== 'full') {
      P.ar.paras = cutParas(P.ar.paras, st.a, 'ar');
      P.uz.paras = cutParas(P.uz.paras, st.u, 'uz');
    }
    return P;
  }
  /* The start at the Companion (the owner, 2026-10-31): the field `st` of the data (tools/sahabi.js, at build time),
     {a: side, u: side}, side = [k, prefix, z, o]: the first k words go (words(…, 'full')), the prefix of the data is put
     before the rest («عَنْ أَبِي هُرَيْرَةَ», «Абу Ҳурайрадан ривоят қилинади:»), the last z «»» go (they closed a quote
     of the part left out), the first o «»» that close no quote of the rest go too. Latin: web/app.js takes `su` of the
     Latin block for `u`. A side without it: the whole text. */
  var MODES = ['sahabi', 'full', 'matn'];
  function cutParas(paras, c, kind) {
    if (!Array.isArray(c) || !paras.length) return paras;
    var k = c[0] || 0, pre = c[1], z = c[2] || 0, n = 0, out = [], started = !k;
    paras.forEach(function (p) {
      if (started) { out.push(p); return; }
      var text = p.map(function (r) { return r.t; }).join(''), re = /\S+/g, m, at = -1;
      while ((m = re.exec(text))) { if (n === k) { at = m.index; break; } n++; }
      if (at < 0) return;                                   // all the words of the paragraph go
      started = true;
      var pos = 0, q = [];
      p.forEach(function (r) {
        var a = pos, b = pos + r.t.length;
        pos = b;
        if (b <= at) return;
        var x = {}; for (var key in r) x[key] = r[key];
        if (a < at) x.t = r.t.slice(at - a);
        q.push(x);
      });
      out.push(q);
    });
    if (!started || !out.length) return paras;            // not this text (the data changed): all of it
    // never «« / »» from the cut (the owner, 2026-10-31): a quote opened again in the prefix before a word with its own
    // quote (the data before it: «…: ««Мен») is not opened again, the » that closed it goes like those of c[3]
    var o = c[3] || 0;
    if (typeof pre === 'string') pre = pre.replace(/«{2,}(?=[^\s«»])/, function (m) { o += m.length - 1; return '«'; });
    if (pre) out[0] = [kind === 'ar' ? { t: pre + ' ', matn: false } : { t: pre + ' ', b: false, i: false }].concat(out[0]);
    if (o > 0) {                                            // the first o » that close no quote of the text left
      var dq = 0;
      out = out.map(function (p) {
        return p.map(function (r) {
          var t = '';
          for (var ci = 0; ci < r.t.length; ci++) {
            var ch = r.t[ci];
            if (ch === '«') dq++;
            else if (ch === '»') { if (dq > 0) dq--; else if (o > 0) { o--; continue; } }
            t += ch;
          }
          if (t === r.t) return r;
          var y = {}; for (var kk in r) y[kk] = r[kk]; y.t = t; return y;
        }).filter(function (x) { return x.t; });
      });
    }
    for (var pi = out.length - 1; pi >= 0 && z > 0; pi--) {
      for (var ri = out[pi].length - 1; ri >= 0 && z > 0; ri--) {
        var r = out[pi][ri], t = r.t;
        for (var ci = t.length - 1; ci >= 0 && z > 0; ci--) if (t[ci] === '»') { t = t.slice(0, ci) + t.slice(ci + 1); z--; }
        if (t !== r.t) { var y = {}; for (var kk in r) y[kk] = r[kk]; y.t = t; out[pi][ri] = y; }
      }
      out[pi] = out[pi].filter(function (x) { return x.t; });
    }
    return out.filter(function (p) { return p.length; });
  }
  /* The words of a text: [{k, para, parts: [{t, b, i, matn}]}] (spaces between them are not words). A word is matn when
     most of its letters are matn (Arabic: inside «», Uzbek: bold). mode 'matn': only those words; `gap` marks the first
     word after a left-out part (shown and inserted as « … »). */
  function words(paras, mode) {
    var out = [], k = 0;
    paras.forEach(function (p, pi) {
      var cur = null;
      p.forEach(function (r) {
        var bits = r.t.split(/(\s+)/);
        bits.forEach(function (bit) {
          if (!bit) return;
          if (/^\s+$/.test(bit)) { cur = null; return; }
          var part = { t: bit, b: !!r.b, i: !!r.i, e: !!r.e, matn: r.matn != null ? !!r.matn : !!r.b };
          if (!cur) { cur = { para: pi, parts: [] }; out.push(cur); }
          cur.parts.push(part);
        });
      });
    });
    out.forEach(function (w) {
      var a = 0, b = 0;
      w.parts.forEach(function (x) { var n = x.t.replace(/[«»"“”.,:;!?،؛()\[\]-]/g, '').length || x.t.length; if (x.matn) a += n; else b += n; });
      w.matn = a > b;
      w.text = w.parts.map(function (x) { return x.t; }).join('');
    });
    out.forEach(function (w, i) { w.idx = i; });
    var list = mode === 'matn' ? out.filter(function (w) { return w.matn; }) : out, prevIdx = -1;
    list.forEach(function (w) { w.gap = mode === 'matn' && prevIdx >= 0 && w.idx !== prevIdx + 1; prevIdx = w.idx; w.k = k++; });
    return list;
  }
  /* The chosen words of a text: [from, to] (to null: the same word) or a list of such parts (Ctrl/Cmd+click) ->
     the parts in the order of the text, inside 0..n-1, the ones that touch or overlap joined; null: all the text. */
  function ranges(r, n) {
    if (!Array.isArray(r) || !r.length) return null;
    var list = typeof r[0] === 'number' ? [r] : r, out = [];
    list.forEach(function (x) {
      if (!Array.isArray(x) || typeof x[0] !== 'number') return;
      var a = x[0], z = x[1] == null ? a : x[1];
      if (z < a) { var y = a; a = z; z = y; }
      if (n != null) { if (a > n - 1) return; a = Math.max(0, a); z = Math.min(z, n - 1); }
      out.push([a, z]);
    });
    out.sort(function (p, q) { return p[0] - q[0] || p[1] - q[1]; });
    var res = [];
    out.forEach(function (x) {
      var l = res[res.length - 1];
      if (l && x[0] <= l[1] + 1) l[1] = Math.max(l[1], x[1]); else res.push([x[0], x[1]]);
    });
    return res.length ? res : null;
  }
  /* The order of the books in the results of a search and in the lists (the owner, 2026-10-30): the six books, Ahmad,
     al-Bayhaqi, ad-Darimi (not in the data yet: its place is kept), al-Hakim, at-Tabarani, ad-Daraqutni. A book not in
     the list comes after them in the order of the data. */
  var BOOK_ORDER = ['bukhari', 'muslim', 'abudawud', 'tirmidhi', 'nasai', 'ibnmajah', 'ahmad', 'bayhaqi', 'darimi', 'hakim', 'tabarani', 'daraqutni'];
  function bookRank(id) { var i = BOOK_ORDER.indexOf(id); return i < 0 ? BOOK_ORDER.length : i; }
  function hasMatn(paras) { return words(paras, 'full').some(function (w) { return w.matn; }); }

  /* ---------- quotes: always «» in the Uzbek text (the owner, 2026-10-30) ----------
     “ ” (and „ “) -> « »; straight " only when they pair up (an opening one after a space / the start / a bracket, a
     closing one elsewhere, as many of each): an odd or a mixed-up set is left as it is. The same as fix_quotes() of
     tools/build_data.py. A character for a character: the positions in the text do not move. */
  function fixQuotes(s) {
    s = String(s == null ? '' : s);
    if (!/["“”„]/.test(s)) return s;
    s = /„/.test(s) ? s.replace(/„/g, '«').replace(/“/g, '»') : s.replace(/“/g, '«');
    s = s.replace(/”/g, '»');
    var n = 0, open = 0, out = '';
    for (var i = 0; i < s.length; i++) if (s[i] === '"') n++;
    if (!n || n % 2) return s;
    var marks = [];
    for (i = 0; i < s.length; i++) {
      if (s[i] !== '"') continue;
      var o = i === 0 || /[\s(\[«—–-]/.test(s[i - 1]);
      marks.push(o); open += o ? 1 : -1;
      if (open < 0) return s;
    }
    if (open !== 0) return s;
    var k = 0;
    for (i = 0; i < s.length; i++) out += s[i] === '"' ? (marks[k++] ? '«' : '»') : s[i];
    return out;
  }
  /* the same over the paragraphs of a text (the pairs may cross the runs of <b>, <i>) */
  function fixQuotesParas(paras) {
    var all = paras.map(function (p) { return p.map(function (r) { return r.t; }).join(''); }).join('\n');
    if (!/["“”„]/.test(all)) return paras;
    var fixed = fixQuotes(all), at = 0;
    return paras.map(function (p) {
      var q = p.map(function (r) {
        var x = {}; for (var k in r) x[k] = r[k];
        x.t = fixed.substr(at, r.t.length); at += r.t.length;
        return x;
      });
      at += 1;                                              // '\n'
      return q;
    });
  }

  /* ---------- bold: whose words ----------
     The words of the Prophet ﷺ in bold, only them; when the hadith has none of them (an athar), the words of the
     Companion / the Successor (the matn after the isnad). Only from the text of the data:
       uz: the <b> of the data (the translators marked the words of the Prophet ﷺ) — taken as it is; without it: a «…»
           after «Расулуллоҳ соллаллоҳу алайҳи васаллам (шундай) дедилар:» / before «…», дедилар» with him in the sentence;
       ar: a «…» after «قال رسول الله ﷺ:», «أن النبي ﷺ قال:», «سمعت رسول الله ﷺ يقول:» …; without «» at all: the end of
           the text after the only such words, when no one else speaks after them.
     An athar: ar — the text after the first «قال:» that follows the last «حدثنا / أخبرنا / سمعت …» (up to the words of
     the author «قال أبو عيسى …», a new line); uz — the only «…» of the text after «… айтади:».
     Not sure — nothing: words of the Prophet ﷺ seen (قال رسول الله ﷺ, «соллаллоҳу алайҳи васаллам … дедилар») but not
     found where they begin and end, unbalanced «», two «…» of narrators.
     -> {kind: 'nabi' | 'athar' | '', ar: [[from, to)] | null, uz: [[from, to)] | null, why}: ranges over the text of
        the paragraphs joined by '\n' (uz with the quotes fixed). */
  var AR_SKIP = /[ؐ-ًؚ-ٰٟۖ-ۭـ‌‍‎‏]/;
  function arNorm(s) {                                     // without the vowel marks, one alif, ي, ه; map: to the text
    var out = '', map = [];
    for (var i = 0; i < s.length; i++) {
      var c = s[i];
      if (AR_SKIP.test(c)) continue;
      out += /[أإآٱ]/.test(c) ? 'ا' : c === 'ى' ? 'ي' : c === 'ة' ? 'ه' : c;
      map.push(i);
    }
    map.push(s.length);
    return { s: out, map: map };
  }
  var SAW_AR = '(?:ﷺ|صلي الله عليه وسلم|صلي الله عليه وعلي اله وسلم)';
  var NABI_AR = '(?:رسول الله|رسول اللّه|نبي الله|النبي|رسول|نبي)';
  var B_AR = '(?:^|[\\s،,:؛.(])';
  var TO_AR = '(?:\\s+(?:لنا|لي|لهم|له|لها|لهن|للناس))?';
  var SAY_AR = '(?:[وف]?قال|[وف]?يقول|كان يقول|[وف]?نادي|[وف]?ينادي)';
  var CONJ_AR = '(?:(?:انه|وهو|ثم|قد|كان)\\s+)*';
  var R_AR_PRE = new RegExp(B_AR + '(?:[وف]?قال|[وف]?يقول|سمعت|سمعنا|عن|به|[وف]?نادي)' + TO_AR + '\\s+' + NABI_AR + '\\s*' + SAW_AR + '\\s*[:،,]?\\s*$');
  var R_AR_POST = new RegExp(B_AR + NABI_AR + '\\s*' + SAW_AR + '\\s*،?\\s*' + CONJ_AR + SAY_AR + TO_AR + '\\s*[:،,]?\\s*$');
  var R_AR_SEEN = new RegExp(B_AR + '(?:' + NABI_AR + '\\s*' + SAW_AR + '\\s*،?\\s*' + CONJ_AR + '(?:[وف]?قال|[وف]?يقول)|(?:[وف]?قال|[وف]?يقول|سمعت)' + TO_AR + '\\s+' + NABI_AR + '\\s*' + SAW_AR + ')');
  var R_AR_INTRO_G = new RegExp('(?:' + B_AR + '(?:[وف]?قال|[وف]?يقول)' + TO_AR + '\\s+' + NABI_AR + '\\s*' + SAW_AR + '|' + B_AR + NABI_AR + '\\s*' + SAW_AR + '\\s*،?\\s*' + CONJ_AR + '(?:[وف]?قال|[وف]?يقول)' + TO_AR + ')\\s*:', 'g');
  var R_AR_SPEAK_G = /(?:^|[\s،,.])(?:[وف]?(?:قال|قالت|قالا|قالوا|يقول|تقول|يقولان))(?:\s+[^\s:«»\n]+){0,4}?\s*:/g;
  var R_AR_CHAIN_AT = /^[\s،,]*(?:[وف]?(?:حدثنا|حدثني|حدثتني|حدثتنا|حدثه|حدثهم|اخبرنا|اخبرني|اخبرتني|اخبرتنا|اخبره|انبانا|انباني|ثنا|نا|سمعت|سمعنا)|عن|ح)(?=[\s،,:]|$)/;
  var R_AR_SPEECH = /(?:^|[\s،,.])(?:[وف]?(?:قال|قالت|قالا|قالوا|يقول|تقول|قلت|سمعت)|حدثنا|اخبرنا)(?=[\s،,:]|$)/;
  var R_AR_NOTE = /(?:^|[\s،,.«])(?:قال\s+(?:ابو\s+عيسي|ابو\s+داود|ابو\s+عبد\s+الرحمن|ابو\s+عبد\s+الله|ابو\s+بكر|الحاكم|الشيخ|البيهقي|ابن\s+ماجه)|(?:و?هذا\s+(?:حديث|اسناد|الحديث))|صحيح\s+علي\s+شرط|حديث\s+(?:[^\s.]+\s+){0,5}?حديث\s+(?:حسن|صحيح|غريب)|(?:و?العمل\s+علي\s+هذا)|[وف]?(?:حدثنا|حدثني|حدثتني|حدثتنا|اخبرنا|اخبرني|اخبرتني|اخبرتنا|انبانا|ثنا)(?=[\s،,:]))/;
  // the narrator just before «قال:» (the end of the isnad): «عن فلان بن فلان رضي الله عنه قال:», «أن فلانا قال:»
  var R_AR_NARRATOR = /(?:^|[\s،,(])(?:[وف]?(?:حدثنا|حدثني|اخبرنا|اخبرني|انبانا|انباني|ثنا|نا|سمعت|سمعنا)|عن|عنه|ان|انه|انها)\s+(?:[^\s:«»]+\s+){0,8}?[^\s:«»]+\s*[،,]?\s*$/;
  var FORM_UZ = '(?:соллаллоҳу\\s+алайҳи\\s+(?:ва\\s*)?саллам|саллаллоҳу\\s+алайҳи\\s+(?:ва\\s*)?саллам|алайҳиссалом|ﷺ)';
  var SUBJ_UZ = '(?:' + FORM_UZ + '(?:нинг)?|(?:^|[\\s,(])у\\s+зот(?:нинг)?)';
  var HVERB_UZ = '(?:[^\\s«»:]*(?:дилар|эдилар|ардилар|адилар|йдилар)(?:-да|-ку)?|(?:деган|айтган)лар(?:-да)?|(?:деган|айтган|айтаётган|деяётган)лар[^\\s«»:]*\\s+(?:эшит|кўр|айт)[^\\s«»:]*)';
  var R_UZ_PRE = new RegExp(SUBJ_UZ + '(?=[\\s:,])(?:\\s+[^\\s«»:]+){0,5}?\\s+' + HVERB_UZ + '\\s*:\\s*$', 'i');
  var R_UZ_COLON = new RegExp('(?:' + FORM_UZ + '|(?:^|[\\s,(])у\\s+зот)(?:\\s+[^\\s«»:,]+){0,4}?\\s*:\\s*$', 'i');
  var R_UZ_BARE = new RegExp(FORM_UZ + '\\s*:\\s*$', 'i');
  var R_UZ_SUBJ = new RegExp(SUBJ_UZ + '(?=[\\s,:]|$)', 'i');
  var SAYS_UZ = '(?:дедилар|айтдилар|дер\\s+эдилар|айтар\\s+эдилар|дердилар|айтардилар|дейдилар|айтадилар|деганлар\\S*|айтганлар\\S*|' +
                'сўрадилар|сўрардилар|буюрдилар|марҳамат\\s+қилдилар|нидо\\s+қилдилар|дуо\\s+қилдилар|жавоб\\s+бердилар|қайтардилар)';
  var R_UZ_AFTER = new RegExp('^[\\s,.!?]*(?:(?:деб|дея)\\s+)?(?:[^\\s«».:]+\\s+){0,3}?' + SAYS_UZ + '(?=[\\s.,;:!?»-]|$)', 'i');
  var R_UZ_SEEN = new RegExp(SUBJ_UZ + '[^.!?«»\\n]{0,120}?(?:дедилар|айтдилар|дер\\s+эдилар|айтар\\s+эдилар|дердилар|айтардилар|дейдилар|айтадилар|деганлар|айтганлар|айтаётганлар|марҳамат\\s+қилдилар|буюрдилар|сўрадилар|нидо\\s+қилдилар)|' + FORM_UZ + '\\s*:\\s*«', 'i');

  /* the «» pairs of a text: [{o, c, depth}] (o, c: the positions of « and »); null when they do not balance.
     lenient: the first « of the text left open (the translation often does not close the quote of the narrator) is
     taken as closed at the end ({virtual: true}); anything else that does not balance: null. */
  function quotePairs(s, lenient) {
    var st = [], out = [];
    for (var i = 0; i < s.length; i++) {
      if (s[i] === '«') st.push(i);
      else if (s[i] === '»') {
        if (!st.length) return null;
        var o = st.pop();
        out.push({ o: o, c: i, depth: st.length });
      }
    }
    if (st.length === 1 && lenient && st[0] === s.indexOf('«')) out.push({ o: st[0], c: s.length, depth: 0, virtual: true });
    else if (st.length) return null;
    return out.sort(function (a, b) { return a.o - b.o; });
  }
  /* the text before a quote back to the previous «, » (and, for uz, the end of a sentence) */
  function leadOf(s, o, sentence) {
    var a = o - 1;
    while (a >= 0 && s[a] !== '«' && s[a] !== '»' && s[a] !== '\n' && !(sentence && /[.!?]/.test(s[a]) && /\s/.test(s[a + 1] || ''))) a--;
    return s.slice(a + 1, o);
  }
  function outer(list) {                                   // the pairs that are not inside an other one of the list
    var res = [];
    list.sort(function (a, b) { return a.o - b.o; }).forEach(function (p) {
      var l = res[res.length - 1];
      if (!l || p.o > l.c) res.push(p);
    });
    return res;
  }
  function arSpans(text) {
    var N = arNorm(text), s = N.s, map = N.map, res = { nabi: [], seen: R_AR_SEEN.test(s), athar: null, unbalanced: false };
    var pairs = s.indexOf('«') >= 0 || s.indexOf('»') >= 0 ? quotePairs(s) : [];
    if (pairs === null) { res.unbalanced = true; pairs = []; }
    pairs.forEach(function (p) {
      var lead = leadOf(s, p.o, false);
      if (R_AR_PRE.test(lead) || R_AR_POST.test(lead)) res.nabi.push(p);
    });
    res.nabi = outer(res.nabi).map(function (p) { return [map[p.o], map[p.c] + 1]; });
    if (!pairs.length && !res.unbalanced) {
      // no «» at all (Аҳмад): the end of the text after the only «قال رسول الله ﷺ:», when no one speaks after it
      var m, intro = [];
      R_AR_INTRO_G.lastIndex = 0;
      while ((m = R_AR_INTRO_G.exec(s))) intro.push(m.index + m[0].length);
      if (intro.length === 1) {
        var from = intro[0], rest = s.slice(from), cut = rest.search(/\n/);
        if (cut >= 0) rest = rest.slice(0, cut);
        if (!R_AR_SPEECH.test(rest) && !R_AR_NOTE.test(rest) && rest.trim().split(/\s+/).length >= 3) {
          var end = from + rest.replace(/\s+$/, '').length;
          while (from < end && /\s/.test(s[from])) from++;
          if (end > from) res.nabi.push([map[from], map[end - 1] + 1]);
        }
      }
    }
    // an athar: after the first «قال:» that is not followed by the isnad again («قال: حدثنا», «قال: سمعت فلانا يقول:»)
    R_AR_SPEAK_G.lastIndex = 0;
    while ((m = R_AR_SPEAK_G.exec(s)) && R_AR_CHAIN_AT.test(s.slice(m.index + m[0].length, m.index + m[0].length + 40))) {}
    var before = m ? s.slice(Math.max(0, m.index - 140), m.index + 1).replace(/[\s،,.]$/, '') : '';
    var nm = before && R_AR_NARRATOR.exec(before);
    // the narrator, not a sentence about the Prophet ﷺ («أنه صلى مع النبي ﷺ …، قال:»); «زوج النبي ﷺ» is a name
    if (nm && new RegExp(SAW_AR).test(nm[0].replace(new RegExp('زوج\\s+' + NABI_AR + '\\s*' + SAW_AR, 'g'), ''))) nm = null;
    if (m && !res.unbalanced && nm) {
      var a = m.index + m[0].length, z = s.length, nl = s.indexOf('\n', a), note = s.slice(a).search(R_AR_NOTE);
      if (nl >= 0) z = nl;
      if (note >= 0 && a + note < z) z = a + note;
      while (a < z && /\s/.test(s[a])) a++;
      while (z > a && /[\s،,]/.test(s[z - 1])) z--;
      // a quote cut in two by the end: not sure
      // a region that opens with «: up to its », the words after it are a narrator's note («وقال يونس …»)
      if (s[a] === '«') {
        var qp = quotePairs(s.slice(a, z)) || [];
        qp.forEach(function (p) { if (p.o === 0 && /^[\s.،,]*(?:و?قال|وفي|وزاد|ورواه|وكذا|ح[\s)])/.test(s.slice(a + p.c + 1, z))) z = a + p.c + 1; });
      }
      var inside = s.slice(a, z), q = inside.length ? quotePairs(inside) : [];
      if (q && inside.split(/\s+/).length >= 3) res.athar = [map[a], map[z - 1] + 1];
    }
    return res;
  }
  function uzSpans(s) {
    var res = { nabi: [], seen: R_UZ_SEEN.test(s), athar: null, unbalanced: false };
    var pairs = s.indexOf('«') >= 0 || s.indexOf('»') >= 0 ? quotePairs(s, true) : [];
    if (pairs === null) { res.unbalanced = true; return res; }
    pairs.forEach(function (p) {
      if (p.virtual) return;
      var lead = leadOf(s, p.o, false), sent = leadOf(s, p.o, true), after = s.slice(p.c + 1, p.c + 80);
      var nabi = R_UZ_PRE.test(lead) || R_UZ_BARE.test(lead) ||
                 R_UZ_COLON.test(lead) && R_UZ_AFTER.test(after) ||
                 !/:\s*$/.test(lead) && R_UZ_SUBJ.test(sent) && R_UZ_AFTER.test(after);
      if (nabi) res.nabi.push(p);
    });
    res.nabi = outer(res.nabi).map(function (p) { return [p.o, p.c + 1]; });
    var top = pairs.filter(function (p) { return p.depth === 0; });
    // «… ривоят қилинади:» and the quote on the next line too
    if (top.length === 1 && !top[0].virtual && /:\s*$/.test(s.slice(Math.max(0, top[0].o - 300), top[0].o))) res.athar = [top[0].o, top[0].c + 1];
    return res;
  }
  function flat(paras) { return paras.map(function (p) { return p.map(function (r) { return r.t; }).join(''); }).join('\n'); }
  function emphasis(h, P) {
    P = P || parseHadith(h);
    var uzP = fixQuotesParas(P.uz.paras), arText = flat(P.ar.paras), uzText = flat(uzP);
    var data = [], pos = 0;                                // the <b> of the data
    uzP.forEach(function (p) {
      p.forEach(function (r) {
        if (r.b) { var l = data[data.length - 1]; if (l && l[1] === pos) l[1] = pos + r.t.length; else data.push([pos, pos + r.t.length]); }
        pos += r.t.length;
      });
      pos += 1;
    });
    var A = arText ? arSpans(arText) : { nabi: [], seen: false, athar: null }, U = uzText ? uzSpans(uzText) : { nabi: [], seen: false, athar: null };
    var nabi = data.length || A.nabi.length || U.nabi.length || A.seen || U.seen;
    if (nabi) {
      var ar = A.nabi.length ? A.nabi : null, uz = data.length ? data : U.nabi.length ? U.nabi : null;
      return { kind: ar || uz ? 'nabi' : '', ar: ar, uz: uz, why: data.length ? 'uz-data' : ar || uz ? 'found' : 'seen-not-found' };
    }
    var ar2 = A.athar ? [A.athar] : null, uz2 = U.athar ? [U.athar] : null;
    return { kind: ar2 || uz2 ? 'athar' : '', ar: ar2, uz: uz2, why: ar2 || uz2 ? 'athar' : 'none' };
  }
  /* the paragraphs with every run cut where the ranges begin and end; e: in a range (bold) */
  function markRuns(paras, ranges) {
    var pos = 0;
    return paras.map(function (p) {
      var out = [];
      p.forEach(function (r) {
        var t = r.t, i = 0;
        while (i < t.length) {
          var at = pos + i, on = false, next = pos + t.length;
          (ranges || []).forEach(function (g) {
            if (at >= g[0] && at < g[1]) { on = true; next = Math.min(next, g[1]); }
            else if (g[0] > at) next = Math.min(next, g[0]);
          });
          var len = Math.max(1, Math.min(t.length - i, next - at)), x = {};
          for (var k in r) x[k] = r[k];
          x.t = t.substr(i, len); x.e = on;
          out.push(x); i += len;
        }
        pos += t.length;
      });
      pos += 1;
      return out;
    });
  }

  /* ---------- tahrij and the words of the author ----------
     book: an entry of web/data/books.json (author, name, edition, label) or only an id; lang: an i18n code.
     The form of the owner (2026-10-30), uz (Latin the same in Latin letters), in italics in the text, the full stop after:
       («Саҳиҳи Бухорий», 1) · («Саҳиҳи Муслим», 1)                     Bukhari, Muslim: the number, never a grade
       («Абу Довуд», 2226; Ҳадиснинг даражаси: «саҳиҳ». Манба: «Саҳиҳи Абу Довуд», 2226)
       («Абу Довуд», 2226)                                              no grading in the data
     The grade and the number in the work: only hukm2 of the data (gradesOf) or the words of the author of the book
     (hukm, the number of the hadith); the title of the work: WORKS (= data/gradings.json works[].tahrij_uz …).
     opts.grades: with the grade (default off here, on in format), opts.hukm: false — not the words of the author,
     opts.all: every grading («; » between them) instead of the main one (Албоний, Арнаут, Хайсамий / Заҳабий / Бусирий,
     the author), opts.withBook: the title of the collection («Сунани Абу Довуд»), opts.paren: false — no brackets.
     ar: «(رواه البخاري ١)». */
  var TAHRIJ_WORDS = {
    uz: { grade: 'Ҳадиснинг даражаси', source: 'Манба' },
    uz_latn: { grade: 'Hadisning darajasi', source: 'Manba' }
  };
  /* the titles of the works of the gradings in a tahrij: not the scholar, his book. Albani: «Саҳиҳи …» for a sound
     grade (саҳиҳ, ҳасан), «Заифи …» for a weak one; neither (мақтуъ alone): the title of the set (uz). The same as
     data/gradings.json works (tools/test_format.js checks it). */
  var WORKS = {
    'albani-abudawud': { uz: 'Саҳиҳ ва заиф Сунани Абу Довуд', t: ['Саҳиҳи Абу Довуд', 'Sahihi Abu Dovud'], z: ['Заифи Абу Довуд', 'Zaifi Abu Dovud'] },
    'albani-tirmidhi': { uz: 'Саҳиҳ ва заиф Сунани Термизий', t: ['Саҳиҳи Термизий', 'Sahihi Termiziy'], z: ['Заифи Термизий', 'Zaifi Termiziy'] },
    'albani-nasai': { uz: 'Саҳиҳ ва заиф Сунани Насоий', t: ['Саҳиҳи Насоий', 'Sahihi Nasoiy'], z: ['Заифи Насоий', 'Zaifi Nasoiy'] },
    'albani-ibnmajah': { uz: 'Саҳиҳ ва заиф Сунани Ибн Можа', t: ['Саҳиҳи Ибн Можа', 'Sahihi Ibn Moja'], z: ['Заифи Ибн Можа', 'Zaifi Ibn Moja'] },
    'arnaut-abudawud': { uz: 'Сунани Абу Довуд, Арнаут таҳқиқи', t: ['Абу Довуд тахрижи', 'Abu Dovud taxriji'] },
    'arnaut-ibnmajah': { uz: 'Сунани Ибн Можа, Арнаут таҳқиқи', t: ['Ибн Можа тахрижи', 'Ibn Moja taxriji'] },
    'arnaut-ahmad': { uz: 'Муснади Аҳмад, Арнаут таҳқиқи', t: ['Муснад тахрижи', 'Musnad taxriji'] },
    'dhahabi-talkhis': { uz: 'Талхисул Мустадрак', t: ['Талхисул Мустадрак', 'Talxisul Mustadrak'] },
    'busiri-zawaid': { uz: 'аз-Завоид (Мисбоҳуз зужожа)', t: ['Мисбоҳуз зужожа', 'Misbohuz zujoja'] },
    'haythami-majma': { uz: 'Мажмаъуз завоид', t: ['Мажмаъуз завоид', 'Majmaʼuz zavoid'] },
    // the same works read from the Shamela copies of the owner (tools/match_shamela.py): the full edition in data/gradings.json
    'sh-arnaut-ahmad': { uz: 'Муснади Аҳмад, Арнаут таҳқиқи («Рисола»)', t: ['Муснад тахрижи', 'Musnad taxriji'] },
    'sh-arnaut-abudawud': { uz: 'Сунани Абу Довуд, Арнаут таҳқиқи', t: ['Абу Довуд тахрижи', 'Abu Dovud taxriji'] },
    'sh-arnaut-ibnmajah': { uz: 'Сунани Ибн Можа, Арнаут таҳқиқи', t: ['Ибн Можа тахрижи', 'Ibn Moja taxriji'] },
    'sh-haythami-majma': { uz: 'Мажмаъуз завоид (ал-Қудсий нашри)', t: ['Мажмаъуз завоид', 'Majmaʼuz zavoid'] },
    'sh-albani-sahihjami': { uz: 'Саҳиҳул жомеъ ас-сағир', t: ['Саҳиҳул жомеъ', 'Sahihul jomeʼ'] },
    'sh-albani-daifjami': { uz: 'Заифул жомеъ ас-сағир', t: ['Заифул жомеъ', 'Zaiful jomeʼ'] }
  };
  var WORK_BY_UZ = {};
  Object.keys(WORKS).forEach(function (k) { WORK_BY_UZ[WORKS[k].uz] = k; });
  var GRADE_RANK = { 'Албоний': 0, 'Арнаут': 1, 'Хайсамий': 2, 'Заҳабий': 2, 'Бусирий': 2 };
  var WEAK = /заиф|мункар|шоз|мавзуъ|маълул|ботил|мурсал|воҳий/i, SOUND = /саҳиҳ|ҳасан/i;
  function workTitle(x, lang) {
    var w = x.work || {}, id = w.id || WORK_BY_UZ[w.uz], W = WORKS[id], latn = lang === 'uz_latn';
    if (!W) return latn ? toLatin(w.uz || '') : (w.uz || '');
    var weak = WEAK.test(x.g || ''), sound = !weak && SOUND.test(x.g || '');
    if (W.z) { if (weak) return W.z[latn ? 1 : 0]; if (!sound) return latn ? toLatin(W.uz) : W.uz; }
    return W.t[latn ? 1 : 0];
  }
  function numOf(h) { return h.n + (h.s && h.s !== 1 ? '/' + h.s : ''); }
  /* the name of the book in a tahrij: «Саҳиҳи Бухорий», «Саҳиҳи Муслим», else the author's name («Абу Довуд») or,
     withBook, the title of the collection («Сунани Абу Довуд») */
  var REF_NAME = { bukhari: { uz: 'Саҳиҳи Бухорий', uz_latn: 'Sahihi Buxoriy' }, muslim: { uz: 'Саҳиҳи Муслим', uz_latn: 'Sahihi Muslim' },
    ahmad: { uz: 'Муснади Аҳмад', uz_latn: 'Musnadi Ahmad', ar: 'مسند أحمد' } };   // owner 2026-10-30: Ahmad is written «Муснади Аҳмад» in a tahrij
  function bookOf(book) { return typeof book === 'string' ? { id: book } : book || {}; }
  function refName(b, lang, withBook) {
    if (REF_NAME[b.id]) return pick(REF_NAME[b.id], lang);
    if (withBook) { var t = titleName(b, lang); if (t) return t; }
    return bookName(b.id, lang, 'short');
  }
  /* the gradings of a tahrij, the main one first: [{by, g, work, ref, own}] (own: the words of the author of the book) */
  function tahrijGrades(h, book, lang, opts) {
    opts = opts || {};
    var b = bookOf(book), id = b.id || h.b, latn = lang === 'uz_latn';
    if (!b.id) b = { id: id };
    if (NO_HUKM[id]) return [];
    var list = [];
    (Array.isArray(h.hukm2) ? h.hukm2 : []).forEach(function (x, i) {
      if (!x || !x.by || !x.g) return;
      list.push({ by: x.by, byName: latn ? toLatin(x.by) : x.by, g: fixQuotes(latn ? toLatin(x.g) : x.g).replace(/[\s.]+$/, ''),
                  work: workTitle(x, lang), ref: x.ref == null ? null : x.ref, own: false,
                  rank: GRADE_RANK[x.by] != null ? GRADE_RANK[x.by] : 2, i: i });
    });
    if (opts.hukm !== false) {
      var seen = {}, title = titleName(b, lang) || bookName(id, lang, 'short');
      hukmOf(h, b, lang).forEach(function (x, i) {
        if (x.lang !== 'uz') return;                       // an Arabic quote does not go into an Uzbek sentence
        var g = fixQuotes(x.text).replace(/[\s.]+$/, '');
        if (!g || seen[g]) return;
        seen[g] = 1;
        list.push({ by: x.by, byName: x.byName, g: g, work: title, ref: numOf(h), own: true, rank: 3, i: 100 + i });
      });
    }
    list.sort(function (x, y) { return x.rank - y.rank || x.i - y.i; });
    return opts.all ? list : list.slice(0, 1);
  }
  function tahrij(h, book, lang, opts) {
    opts = opts || {};
    var b = bookOf(book), id = b.id || h.b;
    if (!b.id) b = { id: id };
    if (lang === 'ar') {
      var anum = arNum(h.n) + (h.s && h.s !== 1 ? '/' + arNum(h.s) : '');
      var book2 = opts.withBook ? titleName(b, 'ar') : null;
      var sa = 'رواه ' + authorName(b, 'ar') + (book2 ? ' في «' + book2 + '»' : '') + ' ' + anum;
      return opts.paren === false ? sa : '(' + sa + ')';
    }
    var L = TAHRIJ_WORDS[lang] ? lang : 'uz', Wd = TAHRIJ_WORDS[L];
    var s = '«' + refName(b, L, opts.withBook) + '», ' + numOf(h);
    if (opts.grades) {
      tahrijGrades(h, b, L, opts).forEach(function (x) {
        s += '; ' + Wd.grade + ': «' + x.g + '». ' + Wd.source + ': «' + x.work + '»' + (x.ref != null && x.ref !== '' ? ', ' + x.ref : '');
      });
    }
    return opts.paren === false ? s : '(' + s + ')';
  }
  function authorName(b, lang) {
    if (b.author) { if (lang === 'uz' && b.author.uz) return b.author.uz; if (lang === 'uz_latn' && b.author.uzl) return b.author.uzl; if (lang === 'ar' && b.author.ar) return b.author.ar; }
    return bookName(b.id, lang, 'author');
  }
  function titleName(b, lang) {
    var nm = b.name || {}, v = lang === 'uz' ? nm.uz : lang === 'uz_latn' ? nm.uzl : nm[lang];
    if (v) return v;
    var core = BY_ID[b.id];
    return core && core.book ? pick(core.book, lang) : null;
  }
  /* Bukhari and Muslim: never a grade on a hadith (the book is «Саҳиҳ»). Others: only the quotes of hukm[] of the data,
     each with who said it. lang 'ar': the Arabic quotes, else the Uzbek ones (Latin for uz_latn), else the Arabic ones.
     -> [{by, byName, text, lang, f, at}] */
  var NO_HUKM = { bukhari: 1, muslim: 1 };
  var BY_AUTHOR = {};
  BOOKS.forEach(function (b) { BY_AUTHOR[b.short.uz] = b.id; BY_AUTHOR[b.author.uz] = b.id; });
  function hukmOf(h, book, lang) {
    var id = bookOf(book).id || h.b;
    if (NO_HUKM[id] || !h || !Array.isArray(h.hukm) || !h.hukm.length) return [];
    var list = h.hukm.filter(function (x) { return x && x.by && x.text; });
    var want = lang === 'ar' ? 'ar' : 'uz';
    var mine = list.filter(function (x) { return x.lang === want; });
    if (!mine.length) mine = list.filter(function (x) { return x.lang !== want; });
    return mine.map(function (x) {
      var bid = BY_AUTHOR[x.by], text = x.text;
      if (x.lang === 'uz' && lang === 'uz_latn') text = toLatin(text);
      return { by: x.by, book: bid || null, byName: bid ? bookName(bid, lang, 'short') : (lang === 'uz_latn' ? toLatin(x.by) : x.by),
               text: text, lang: x.lang, f: x.f, at: x.at };
    });
  }

  /* The gradings of other scholars (hukm2 of the data, tools/match_gradings.py): only what a source says, with the name
     of the scholar, his work and the number given there (or none). lang 'ar': the Arabic names and words.
     -> [{by, byName, g, ar, work, ref, vp, src, c, text}] (text: «Албоний: саҳиҳ (Саҳиҳ ва заиф Сунани Абу Довуд, 429)»;
     vp: «2/5» the volume/page of the hadith in the edition of the work, when the data has it) */
  var GRADER_AR = { 'Албоний': 'الألباني', 'Арнаут': 'الأرنؤوط', 'Заҳабий': 'الذهبي', 'Бусирий': 'البوصيري', 'Хайсамий': 'الهيثمي' };
  function gradesOf(h, book, lang) {
    var id = bookOf(book).id || (h && h.b);
    if (NO_HUKM[id] || !h || !Array.isArray(h.hukm2) || !h.hukm2.length) return [];
    // Arabic: only the gradings whose Arabic wording is known (the set may give only an English word)
    return h.hukm2.filter(function (x) { return x && x.by && x.g && (lang !== 'ar' || x.ar); }).map(function (x) {
      var w = x.work || {}, ar = lang === 'ar', latn = lang === 'uz_latn';
      var byName = ar ? (GRADER_AR[x.by] || x.by) : latn ? toLatin(x.by) : x.by;
      var g = ar ? x.ar : latn ? toLatin(x.g) : x.g;
      var work = ar ? (w.ar || w.uz || '') : latn ? toLatin(w.uz || '') : (w.uz || '');
      var ref = x.ref == null ? '' : (ar ? arNum(x.ref) : String(x.ref));
      var where = work ? work + (ref ? (ar ? ' (' + ref + ')' : ', ' + ref) : '') : '';
      var text = ar ? 'قال ' + byName + ': ' + g + (where ? ' – ' + where : '')
                    : byName + ': ' + g + (where ? ' (' + where + ')' : '');
      return { by: x.by, byName: byName, g: g, ar: x.ar || null, work: work, ref: x.ref == null ? null : x.ref,
               vp: x.vp || null, src: x.src || null, c: x.c, text: text };
    });
  }

  /* ---------- the result of an insert ----------
     opts = { lang ('uz' | 'uz_latn' | …: the interface, for the tahrij of the paragraphs that are not Arabic),
              script ('cyrl' | 'latn': the Uzbek text), withAr, withUz, order ('ar' | 'uz': which comes first),
              mode ('sahabi' (default: from the Companion, data `st`) | 'full' (all the isnad) | 'matn'), arRange / uzRange ([from, to] word numbers of words(…, mode), a list of such parts
                     — joined by « … » in the order of the text, see ranges() — or null: all),
              refUz (the tahrij after the Uzbek text, the form of tahrij(): «(«Абу Довуд», 2226; Ҳадиснинг даражаси: …)»
                     in italics, the full stop after it; default on), refAr (the Arabic tahrij «(رواه البخاري ١)» after
                     the Arabic text; then the words of the author and the gradings in Arabic after it when there is no
                     Uzbek text); both off: no tahrij,
              withGrades (the grade in the tahrij; default on), allGrades (every grading, not only the main one),
              gradesInMatn (also in the mode 'matn'; default off), withHukm (the words of the author of the book as a
              grade when there is no other; default on), withBook, withBob,
              matnBold (bold: the words of the Prophet ﷺ, else of the Companion — emphasis(); default on),
              arMatnBold (the same in the Arabic text; default off), quotes (['«', '»']), mark (a number «[1]») }
     -> {paras: [{dir, kind: 'ar' | 'uz' | 'bob' | 'tahrij', runs: [{t, b, i, role}]}], text, ref, emphasis, grades}
     roles: 'matn', 'ref' (tahrij), 'hukm', 'grade', 'gap' (« … »), 'mark'.                                     */
  function format(h, book, opts) {
    opts = opts || {};
    var P = parseHadith(h, opts.mode || 'sahabi'), paras = [], latn = opts.script === 'latn', b = bookOf(book);
    if (!b.id) b = { id: h.b };
    P.uz.paras = fixQuotesParas(P.uz.paras);
    var E = emphasis(h, P);
    P.ar.paras = markRuns(P.ar.paras, E.ar);
    P.uz.paras = markRuns(P.uz.paras, E.uz);
    var withAr = opts.withAr !== false && P.ar.paras.length, withUz = opts.withUz !== false && P.uz.paras.length;
    var uzLang = latn ? 'uz_latn' : 'uz';
    var grades = opts.withGrades !== false && !(opts.mode === 'matn' && !opts.gradesInMatn);
    var tOpts = { withBook: opts.withBook, grades: grades, all: !!opts.allGrades, hukm: opts.withHukm !== false };
    function textRuns(paras0, range, kind) {
      var list = words(paras0, opts.mode === 'matn' && hasMatn(paras0) ? 'matn' : 'full');
      var parts = ranges(range, list.length);
      if (parts && list.length) {
        // several parts (Ctrl/Cmd+click): in the order of the text, « … » between them
        var all = list; list = [];
        parts.forEach(function (rg, j) {
          var sl = all.slice(rg[0], rg[1] + 1);
          if (j > 0 && sl.length) sl[0].cut = true;
          list = list.concat(sl);
        });
      }
      var out = [], lastPara = null;
      list.forEach(function (w, i) {
        var prev = ' ';
        if (i > 0) {
          if (w.gap || w.cut) out.push({ t: ' … ', b: false, i: false, role: 'gap' });
          else if (w.para !== lastPara) out.push({ t: '\n', b: false, i: false, role: 'br' });
          else out.push({ t: ' ', b: false, i: false });
        }
        w.parts.forEach(function (x) {
          var t = kind === 'uz' && latn ? toLatin(x.t, prev) : x.t;
          prev = x.t.slice(-1);
          var bold = opts.matnBold === false ? false : kind === 'uz' ? x.e : x.e && !!opts.arMatnBold;
          out.push({ t: t, b: bold, i: kind === 'uz' ? x.i : false, role: x.matn ? 'matn' : '' });
        });
        lastPara = w.para;
      });
      // a space between two bold words is bold too (one run: «<b>a b</b>»)
      for (var k = 1; k + 1 < out.length; k++) if (out[k].t === ' ' && !out[k].role && out[k - 1].b && out[k + 1].b) out[k].b = true;
      // the text in paragraphs of its own («\n» runs split it)
      var res = [[]];
      merge(out).forEach(function (r) {
        if (r.role === 'br') { res.push([]); return; }
        res[res.length - 1].push(r);
      });
      return res.filter(function (p) { return p.length; });
    }
    function plain(runs) { return runs.map(function (r) { return r.t; }).join(''); }
    /* the end of the inserted text before the tahrij: no comma / colon (the data goes on after it) and no full stop
       (it comes after the brackets: «… матн («Саҳиҳи Бухорий», 1).»); «…», «!», «?» stay */
    function trimEnd(runs) {
      runs = runs.slice();
      while (runs.length) {
        var l = runs[runs.length - 1], t = l.t.replace(/[\s,،;؛:]+$/, '');
        t = t.replace(/([^.\s])(\s*\.)(?:\s+\.)+$/, '$1$2');  // «… сўзлар. .» of the data: one full stop
        if (/[^.]\.$|^\.$/.test(t)) t = t.slice(0, -1).replace(/[\s,،;؛:]+$/, '');
        if (t) { runs[runs.length - 1] = { t: t, b: l.b, i: l.i, role: l.role }; break; }
        runs.pop();
      }
      return runs;
    }
    function refRuns(lang) {
      var s = tahrij(h, b, lang, tOpts);
      return [{ t: s, b: false, i: lang !== 'ar', role: 'ref' }, { t: '.', b: false, i: false, role: '' }];
    }
    var arParas = withAr ? textRuns(P.ar.paras, opts.arRange, 'ar') : [];
    var uzParas = withUz ? textRuns(P.uz.paras, opts.uzRange, 'uz') : [];
    /* a last paragraph of nothing but «.» or a line of dots «. . . .» (separators of the data, ~1 300 in Аҳмад): the
       tahrij goes after the paragraph before it */
    function dropEmptyEnd(list) {
      while (list.length > 1 && (!trimEnd(list[list.length - 1]).length || /^[\s.…·–—-]*$/.test(plain(list[list.length - 1])))) list.pop();
    }
    if (opts.refUz !== false) dropEmptyEnd(uzParas);
    if (opts.refAr) dropEmptyEnd(arParas);
    var arAll = arParas.map(plain).join(' ');
    // the Arabic form only: the words of the author that are not in the inserted text, the gradings in Arabic
    function hukmRuns(lang, inside) {
      if (opts.withHukm === false) return [];
      var q = opts.quotes || ['«', '»'];
      return hukmOf(h, b, lang).filter(function (x) { return inside.indexOf(x.text) < 0; }).map(function (x) {
        return { t: ' ' + x.byName + ': ' + q[0] + x.text.replace(/[\s.]+$/, '') + q[1] + '.', b: false, i: false, role: 'hukm' };
      });
    }
    function gradeRuns(lang) {
      if (!grades) return [];
      return gradesOf(h, b, lang).map(function (x) { return { t: ' ' + x.text + '.', b: false, i: false, role: 'grade' }; });
    }
    var mark = opts.mark ? { t: '[' + opts.mark + '] ', b: false, i: false, role: 'mark' } : null;
    if (opts.withBob && (P.bob.length || P.bobAr.length)) {
      if (withAr && P.bobAr.length && !withUz) paras.push({ dir: 'rtl', kind: 'bob', runs: [{ t: P.bobAr.map(function (p) { return plain(p); }).join(' '), b: true, i: false, role: 'bob' }] });
      if (withUz && P.bob.length) {
        var bt = fixQuotes(P.bob.map(function (p) { return plain(p); }).join(' '));
        paras.push({ dir: 'ltr', kind: 'bob', runs: [{ t: latn ? toLatin(bt) : bt, b: true, i: false, role: 'bob' }] });
      }
    }
    var arBlock = arParas.map(function (runs, i) {
      runs = i === arParas.length - 1 && opts.refAr ? trimEnd(runs) : runs.slice();
      if (i === arParas.length - 1 && opts.refAr) {
        runs.push({ t: ' ', b: false, i: false });
        runs = runs.concat(refRuns('ar'));
        if (!withUz) runs = runs.concat(hukmRuns('ar', arAll), gradeRuns('ar'));
      }
      return { dir: 'rtl', kind: 'ar', runs: merge(runs) };
    });
    /* the translation of Ибн Можа ends with the grade of al-Albani («…]. Саҳиҳ», ~3 600): when the tahrij gives the
       very same grade, it is not said twice */
    var mainGrade = grades && opts.refUz !== false ? tahrijGrades(h, b, uzLang, tOpts)[0] : null;
    function dropGrade(runs) {
      if (!mainGrade || !runs.length) return runs;
      var l = runs[runs.length - 1], m = /(^|[.\]»!?]\s+|^\s*)([^\s.\]»]+(?: [^\s.\]»]+){0,2})\.?\s*$/.exec(l.t);
      if (!m || m[2].toLowerCase() !== mainGrade.g.toLowerCase() || m[2].charAt(0) === m[2].charAt(0).toLowerCase()) return runs;
      runs = runs.slice(0, -1);
      var rest = l.t.slice(0, m.index + m[1].length);
      if (rest.trim()) runs.push({ t: rest, b: l.b, i: l.i, role: l.role });
      return runs;
    }
    if (mainGrade && uzParas.length && !opts.uzRange) {
      var lastP = dropGrade(uzParas[uzParas.length - 1]);
      if (lastP.length) uzParas[uzParas.length - 1] = lastP; else if (uzParas.length > 1) uzParas.pop();
    }
    var uzBlock = uzParas.map(function (runs, i) {
      runs = i === uzParas.length - 1 && opts.refUz !== false ? trimEnd(runs) : runs.slice();
      if (i === uzParas.length - 1 && opts.refUz !== false) {
        runs.push({ t: ' ', b: false, i: false });
        runs = runs.concat(refRuns(uzLang));
      }
      return { dir: 'ltr', kind: 'uz', runs: merge(runs) };
    });
    var first = opts.order === 'uz' ? uzBlock : arBlock, second = opts.order === 'uz' ? arBlock : uzBlock;
    paras = paras.concat(first, second);
    // nothing to put the tahrij after (no Uzbek text, the Arabic reference off): a paragraph of its own
    if (!uzBlock.length && !opts.refAr && opts.refUz !== false && (arBlock.length || paras.length)) {
      paras.push({ dir: 'ltr', kind: 'tahrij', runs: merge(refRuns(uzLang)) });
    }
    if (mark && paras.length) paras[0].runs.unshift(mark);
    var text = paras.map(function (p) { return plain(p.runs); }).join('\n');
    return { paras: paras, text: text, ref: tahrij(h, b, uzLang, { paren: false, withBook: opts.withBook }), emphasis: E,
             grades: grades && (opts.refUz !== false || opts.refAr) ? tahrijGrades(h, b, uzLang, tOpts) : [] };
  }
  function merge(runs) {
    var out = [];
    runs.forEach(function (r) {
      if (!r.t) return;
      var l = out[out.length - 1];
      if (l && l.b === r.b && l.i === r.i && (l.role || '') === (r.role || '') && r.role !== 'br') l.t += r.t;
      else out.push({ t: r.t, b: !!r.b, i: !!r.i, role: r.role || '' });
    });
    return out;
  }

  /* ---------- the client of the data ----------
     web/data/current.json {version} -> web/data/books.json (?v=) -> web/data/refs/<book>.json (?v=), web/h/v<V>/<book>/<k>.json
     The loader of web/search.js (one load per file, the Cache API for the versioned files, retries, mirrors one after
     another, a mirror that is down is left alone for a while, AbortSignal). o: {bases, fetch, cache, timeout, retries}. */
  function Data(o) {
    if (!(this instanceof Data)) return new Data(o);
    o = o || {};
    var lo = {}; for (var k in o) lo[k] = o[k];
    if (lo.cache === undefined) lo.cache = 'hadis-data-v1';
    if (!lo.memoryBytes) lo.memoryBytes = 24e6;
    this.loader = new S.Loader(lo);
    this.version = o.version || null;
    this.refs = {};
    this._books = null;
  }
  Data.prototype.current = function () {
    var self = this;
    if (this.version) return Promise.resolve(this.version);
    if (!this._cur) this._cur = this.loader.get('data/current.json?t=' + Math.floor(Date.now() / 60000), 'json', false).then(function (c) {
      if (!c || !/^[\w.-]{1,64}$/.test(c.version || '')) throw new Error('bad current.json');
      self.version = c.version; return c.version;
    }, function (e) { self._cur = null; throw e; });
    return this._cur;
  };
  Data.prototype.books = function () {
    var self = this;
    if (this._books) return this._books;
    this._books = this.current().then(function (v) {
      return self.loader.get('data/books.json?v=' + v, 'json', true);
    }).then(function (b) {
      if (!b || !Array.isArray(b.books) || !(b.docs > 0) || !(b.block > 0)) throw new Error('bad books.json');
      if (b.version && self.version && b.version !== self.version) throw new Error('books.json is of another version');
      b.byId = {};
      b.books.forEach(function (x) { b.byId[x.id] = x; });
      self.catalog = b;
      return b;
    });
    this._books.then(null, function () { self._books = null; });
    return this._books;
  };
  Data.prototype.bookOfDoc = function (d) {
    var c = this.catalog; if (!c) return null;
    for (var i = 0; i < c.books.length; i++) { var b = c.books[i]; if (d >= b.doc[0] && d <= b.doc[1]) return b; }
    return null;
  };
  Data.prototype.loadRefs = function (id, signal) {
    var self = this;
    if (this.refs[id]) return Promise.resolve(this.refs[id]);
    return this.books().then(function () {
      return self.loader.get('data/refs/' + id + '.json?v=' + self.version, 'json', true, signal);
    }).then(function (j) { var r = new Refs(j); self.refs[id] = r; return r; });
  };
  Data.prototype.allRefs = function () {
    var self = this;
    return this.books().then(function (c) { return Promise.all(c.books.map(function (b) { return self.loadRefs(b.id).catch(function () { return null; }); })); })
      .then(function () { return self.refs; });
  };
  Data.prototype.refOf = function (d) {                     // {book, n, s?} (refs of its book loaded) | null
    var b = this.bookOfDoc(d); if (!b) return null;
    var r = this.refs[b.id] && this.refs[b.id].ref(d);
    return r ? { book: b.id, n: r.n, s: r.s } : null;
  };
  Data.prototype.docOf = function (ref) {                   // -> doc | -1 (the refs must be loaded)
    var r = ref && this.refs[ref.book];
    return r ? r.doc(ref.n, ref.s) : -1;
  };
  Data.prototype.block = function (bookId, k, signal) {
    var self = this;
    return this.books().then(function () {
      return self.loader.get('h/v' + self.version + '/' + bookId + '/' + k + '.json', 'json', true, signal);
    });
  };
  Data.prototype.hadith = function (d, signal) {
    var self = this;
    return this.books().then(function (c) {
      var b = self.bookOfDoc(d);
      if (!b) throw new Error('no hadith ' + d);
      var off = d - b.doc[0], k = Math.floor(off / c.block);
      return self.block(b.id, k, signal).then(function (arr) {
        var h = Array.isArray(arr) ? arr[off - k * c.block] : null;
        if (!h || h.d !== d) throw new Error('block ' + b.id + '/' + k + ': hadith ' + d + ' is not there');
        return h;
      });
    });
  };

  /* ---------- «Check a quote»: a pasted text against a hadith of the data, word by word ----------
     quote(text) -> {side: 'ar' | 'uz' | '', lat, toks, words, marks, queries}   (side '': no Arabic / Uzbek letters)
     compare(q, h) -> {side, lat, from, to, n, ops: [{op, a, q, w}], same, mark, spell, typo, sub, miss, extra, skip,
                       diff, matched, total, marksChecked, found, exact, part}
       ops: 'same'; 'mark' the letters agree, the vowel marks do not (Arabic); 'spell' the same word in another spelling
       (Uzbek: к / қ, х / ҳ, у / ў, the apostrophe); 'typo' a letter or two, a prefix (و ف ب ل ال); 'sub' another word;
       'miss' a word of the hadith not in the quote; 'extra' a word of the quote not in the hadith; 'skip' left out on
       purpose (« … » in the quote). a: the word of the hadith, q: the word of the quote, w: its index in
       words(paras, 'full') (from, to: the part of the hadith the quote covers, the range for format).
       The Arabic letters are compared without the vowel marks (ٱ أ إ آ = ا, ى = ي, ة = ه, ء dropped: HadisText);
       the vowel marks only where the quote has them (a partly vowelled quote is not wrong), sukun not at all.
       Nothing is added to the text: a, the words of the hadith, come from the data only.                        */
  var Q_MAX = 400, Q_GAP = /^(?:\.{2,}|…+)$/;
  var SAW = T.norm('ﷺ').replace(/ /g, '');      // ﷺ = صلى الله عليه وسلم
  function qKey(s) { return T.norm(s).replace(/ /g, ''); }
  function arMarks(raw) {                           // the vowel marks of every letter of the key, sorted
    var s = typeof raw.normalize === 'function' ? raw.normalize('NFD') : raw, out = [], at = -1, any = false;
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      if (c >= 0x64B && c <= 0x651) { if (at >= 0) { out[at] += s[i]; any = true; } continue; }
      var k = qKey(s[i]);
      if (!k) { if (T.scriptOf(c) === 'Arab' && !(c >= 0x652 && c <= 0x65F) && c !== 0x670) at = -1; continue; }
      for (var j = 0; j < k.length; j++) out.push('');
      at = out.length - 1;
    }
    return { m: out.map(function (x) { return x.split('').sort().join(''); }), any: any };
  }
  function uzExact(s) {                              // the letters as written (lower case, one apostrophe)
    var out = '';
    s = String(s).toLowerCase();
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i), sc = T.scriptOf(c);
      if (/['`´ʻʼʽ‘’′]/.test(s[i])) out += "'";
      else if (sc === 'Latn' || sc === 'Cyrl' || sc === 'Zyyy') out += s[i];
    }
    return out.replace(/^'+|'+$/g, '');
  }
  /* one word -> a token {raw, key, x (exact form / vowel marks), w0, w1}; null when it has no letters */
  function qTok(raw, side, lat, w) {
    var shown = side === 'uz' && lat && w != null ? toLatin(raw) : raw;   // a Cyrillic word of the hadith for a Latin quote
    var key = qKey(shown);
    if (!key) return null;
    var tk = { raw: raw, shown: shown, key: key, w0: w, w1: w };
    if (side === 'ar') { var mk = arMarks(shown); tk.marks = mk.m.length === key.length ? mk : { m: [], any: false }; }
    else tk.exact = uzExact(shown);
    return tk;
  }
  function mergeSaw(toks) {                          // «صلى الله عليه وسلم» written out = ﷺ (the data has both)
    var out = [];
    for (var i = 0; i < toks.length; i++) {
      var x = toks[i];
      if (x && !x.gap && i + 3 < toks.length && x.key === SAW.slice(0, x.key.length)) {
        var k = '', j = i;
        while (j < toks.length && j < i + 5 && toks[j] && !toks[j].gap && k.length < SAW.length) k += toks[j++].key;
        if (k === SAW && j - i > 1) {
          var raw = toks.slice(i, j).map(function (y) { return y.raw; }).join(' ');
          out.push({ raw: raw, shown: toks.slice(i, j).map(function (y) { return y.shown; }).join(' '), key: SAW,
                     w0: x.w0, w1: toks[j - 1].w1, marks: { m: [], any: false }, exact: SAW });
          i = j - 1; continue;
        }
      }
      out.push(x);
    }
    return out;
  }
  function quote(text) {
    text = String(text == null ? '' : text).replace(/\.{3,}|…/g, ' … ').replace(/\s+/g, ' ').trim();
    var sc = T.detectScript(text), side = sc === 'Arab' ? 'ar' : sc === 'Cyrl' || sc === 'Latn' ? 'uz' : '', lat = sc === 'Latn';
    var toks = [];
    if (side) text.split(' ').forEach(function (raw) {
      if (Q_GAP.test(raw)) { if (toks.length && !toks[toks.length - 1].gap) toks.push({ gap: true, raw: '…' }); return; }
      var tk = qTok(raw, side, lat, null);
      if (tk) toks.push(tk);
    });
    if (side === 'ar') toks = mergeSaw(toks);
    while (toks.length && toks[0].gap) toks.shift();
    while (toks.length && toks[toks.length - 1].gap) toks.pop();
    toks = toks.slice(0, Q_MAX);
    var words = toks.filter(function (x) { return !x.gap; });
    // the queries of the search: the beginning, the end and the middle of a long quote (the vowel marks dropped)
    var plain = words.map(function (x) { return x.raw.replace(/[ً-ٰٟۖ-ۭـ]/g, ''); }), queries = [];
    function add(a) { var q = plain.slice(a, a + 10).join(' '); if (q && queries.indexOf(q + ' ') < 0) queries.push(q + ' '); }
    add(0);
    if (plain.length > 14) add(plain.length - 10);
    if (plain.length > 30) add(Math.floor(plain.length / 2) - 5);
    return { side: side, lat: lat, toks: toks, words: words.length, queries: queries,
             marks: side === 'ar' && words.some(function (x) { return x.marks.any; }) };
  }
  function hadithToks(h, side, lat) {
    var P = side === 'ar' ? parseAr(h && h.ar) : parseUz(h && h.uz), out = [];
    words(P, 'full').forEach(function (w) { var tk = qTok(w.text, side, lat, w.k); if (tk) out.push(tk); });
    return side === 'ar' ? mergeSaw(out) : out;
  }
  /* two words with different keys: 'typo' (close) | 'sub' */
  function qClose(a, b, side) {
    if (side === 'ar') { var sa = T.arStem(a), sb = T.arStem(b); if (sa === sb && sa.length >= 2) return true; }
    var L = Math.max(a.length, b.length), lim = L >= 7 ? 2 : L >= 3 ? 1 : 0;
    return lim > 0 && T.levenshtein(a, b, lim) <= lim;
  }
  function sameMarks(a, q) {                          // q's marks are among a's, letter by letter
    if (!q.marks || !q.marks.any || !a.marks || !a.marks.any) return true;
    for (var i = 0; i < q.marks.m.length; i++) {
      var x = q.marks.m[i], y = a.marks.m[i] || '';
      for (var j = 0; j < x.length; j++) if (y.indexOf(x[j]) < 0) return false;
    }
    return true;
  }
  var C_TYPO = 0.6, C_SUB = 1.2, C_GAP = 1;
  function compare(q, h) {
    var side = q.side, lat = q.lat, A = side ? hadithToks(h, side, lat) : [], B = q.toks;
    var res = { side: side, lat: lat, from: -1, to: -1, n: A.length ? A[A.length - 1].w1 + 1 : 0, ops: [], same: 0, mark: 0,
                spell: 0, typo: 0, sub: 0, miss: 0, extra: 0, skip: 0, diff: 0, matched: 0, total: q.words,
                marksChecked: !!q.marks, found: false, exact: false, part: false };
    if (!A.length || !q.words) return res;
    // the part of the hadith to align with: all of it, or around the words of the quote (a long hadith, a short quote)
    var lo = 0, hi = A.length, m = B.length, hasGap = B.some(function (x) { return x.gap; });
    if (!hasGap && A.length > 3 * m + 40) {
      var set = {}, W = m + 10, sum = 0, best = -1, at = 0, i;
      B.forEach(function (x) { set[x.key] = 1; });
      var hit = A.map(function (x) { return set[x.key] ? 1 : 0; });
      for (i = 0; i < A.length; i++) {
        sum += hit[i]; if (i >= W) sum -= hit[i - W];
        if (sum > best) { best = sum; at = Math.max(0, i - W + 1); }
      }
      lo = Math.max(0, at - m); hi = Math.min(A.length, at + W + m);
    }
    var a = A.slice(lo, hi), n = a.length, cols = m + 1;
    var D = new Float64Array((n + 1) * cols), P = new Uint8Array((n + 1) * cols), memo = {};
    function sub(x, y) {
      if (x.key === y.key) return 0;
      var k = x.key + '\u0001' + y.key;
      if (memo[k] == null) memo[k] = qClose(x.key, y.key, side) ? C_TYPO : C_SUB;
      return memo[k];
    }
    // semi-global: the quote all of it, the hadith from any word to any word; « … » in the quote takes words for free
    for (var j = 1; j <= m; j++) { D[j] = D[j - 1] + (B[j - 1].gap ? 0 : C_GAP); P[j] = 3; }
    for (var r = 1; r <= n; r++) {
      var row = r * cols, up = row - cols;
      D[row] = 0; P[row] = 0;
      for (j = 1; j <= m; j++) {
        var y = B[j - 1], best2, how;
        if (y.gap) {
          best2 = D[row + j - 1]; how = 3;
          if (D[up + j] < best2) { best2 = D[up + j]; how = 2; }
        } else {
          best2 = D[up + j - 1] + sub(a[r - 1], y); how = 1;
          if (D[up + j] + C_GAP < best2) { best2 = D[up + j] + C_GAP; how = 2; }
          if (D[row + j - 1] + C_GAP < best2) { best2 = D[row + j - 1] + C_GAP; how = 3; }
        }
        D[row + j] = best2; P[row + j] = how;
      }
    }
    var end = 0;
    for (r = 1; r <= n; r++) if (D[r * cols + m] < D[end * cols + m]) end = r;
    var ops = [];
    r = end; j = m;
    while (j > 0) {
      var p = P[r * cols + j];
      if (p === 1) {
        var x = a[r - 1], yq = B[j - 1], op;
        if (x.key === yq.key) op = side === 'ar' ? (sameMarks(x, yq) ? 'same' : 'mark') : (x.exact === yq.exact ? 'same' : 'spell');
        else op = sub(x, yq) === C_TYPO ? 'typo' : 'sub';
        ops.push({ op: op, a: x.shown, q: yq.raw, w: x.w0, w1: x.w1 }); r--; j--;
      } else if (p === 2) {
        ops.push({ op: B[j - 1].gap ? 'skip' : 'miss', a: a[r - 1].shown, q: null, w: a[r - 1].w0, w1: a[r - 1].w1 }); r--;
      } else if (p === 3) {
        if (!B[j - 1].gap) ops.push({ op: 'extra', a: null, q: B[j - 1].raw, w: -1 });
        j--;
      } else break;
    }
    ops.reverse();
    ops.forEach(function (o) { res[o.op]++; });
    var ws = ops.filter(function (o) { return o.w >= 0; });
    res.ops = ops;
    res.diff = res.typo + res.sub + res.miss + res.extra;
    res.matched = res.same + res.mark + res.spell;
    if (ws.length) { res.from = ws[0].w; res.to = ws[ws.length - 1].w1; }
    res.found = res.matched >= Math.max(Math.min(2, q.words), Math.ceil(q.words * 0.5));
    res.exact = res.found && !res.diff && !res.mark && !res.spell;
    res.part = res.found && (res.from > A[0].w0 || res.to < A[A.length - 1].w1);
    return res;
  }
  /* which of the compared hadiths fits the quote best: higher first */
  function quoteRank(r) { return r.found ? 1000 + r.matched - r.diff - 0.1 * (r.mark + r.spell) - 0.001 * (r.n - (r.to - r.from + 1)) : r.matched - r.diff; }
  /* the hadiths of the data a quote may come from: search (HadisSearch.Search) with the queries of the quote in its
     language, hadith(d, signal) -> Promise<hadith> (Data.hadith), then compare each one.
     -> Promise<[{d, h, r}]>, the best first (r.found false: nothing like the quote) */
  function findQuote(q, search, hadith, opts) {
    opts = opts || {};
    if (!q.side || !q.queries.length) return Promise.resolve([]);
    var max = opts.max || 12;
    return Promise.all(q.queries.map(function (s) {
      return search.run(s, { langs: [q.side], limit: 8, prefix: false, confirm: 8, cut: false, signal: opts.signal }).then(function (l) { return l || []; });
    })).then(function (lists) {
      var seen = {}, docs = [];
      for (var i = 0; i < 8 && docs.length < max; i++) lists.forEach(function (l) {   // in turn: the best of every query first
        if (l[i] && !seen[l[i].d] && docs.length < max) { seen[l[i].d] = 1; docs.push(l[i].d); }
      });
      return Promise.all(docs.map(function (d) {
        return Promise.resolve(hadith(d, opts.signal)).then(function (h) { return { d: d, h: h, r: compare(q, h) }; }, function () { return null; });
      }));
    }).then(function (list) {
      list = list.filter(Boolean);
      list.forEach(function (x, i) { x.i = i; x.rank = quoteRank(x.r); });
      list.sort(function (x, y) { return y.rank - x.rank || x.i - y.i; });
      return list;
    });
  }

  /* the hosts the panel works in (as in quranuz): text everywhere, a picture where the host takes one */
  var HOSTS = {
    word:       { text: true, picture: true },
    wordOnline: { text: true, picture: true },
    gdocs:      { text: true, picture: false },
    wp:         { text: true, picture: false },
    ext:        { text: true, picture: false },
    telegram:   { text: true, picture: true },
    mcp:        { text: true, picture: false }
  };

  return { BOOKS: BOOKS, book: function (id) { return BY_ID[id] || null; }, bookName: bookName, findBook: findBook, fold: fold,
           parseRef: parseRef, Refs: Refs, refLabel: refLabel, parseHadith: parseHadith, MODES: MODES, cutParas: cutParas, parseUz: parseUz, parseAr: parseAr,
           words: words, ranges: ranges, BOOK_ORDER: BOOK_ORDER, bookRank: bookRank, hasMatn: hasMatn, tahrij: tahrij, tahrijGrades: tahrijGrades, WORKS: WORKS, hukmOf: hukmOf, gradesOf: gradesOf,
           format: format, emphasis: emphasis, fixQuotes: fixQuotes, Data: Data, toLatin: toLatin,
           arNum: arNum, latinDigits: latinDigits, HOSTS: HOSTS, quote: quote, compareQuote: compare, quoteRank: quoteRank, findQuote: findQuote, stripHtml: T.stripHtml, NO_HUKM: NO_HUKM };
});
