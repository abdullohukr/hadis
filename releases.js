/* Ҳадислар releases, newest first: [version, date, title {uz, uz_latn}, changes {uz, uz_latn}].
   Read by the panel's Updates window (app.js); the same format as Khatt al-Quran (quranuz/web/releases.js). */
window.HadisReleases = [
  ['0.9.4', '2026-10-31', { uz: 'Word’га қўйиш тезроқ ва ишончлироқ', uz_latn: 'Wordga qoʻyish tezroq va ishonchliroq' }, {
    uz: ['Узун ҳадислар Word’га бир неча баравар тезроқ қўйилади: бир хил кўринишдаги матн бўлакларга бўлинмайди (кўриниши ўша-ўша).',
         'Жуда узун ҳадис қисмларга бўлиб қўйилади, Word «қотиб» қолмайди. Word жавоб бермаса, 45 сониядан кейин бу ҳақда айтилади; қўйиш тугамагунча такрор босиш ҳадисни иккинчи марта қўймайди.',
         'Эски Word’да (Office 2016/2019) панел «Юкланмоқда»да тўхтаб қолмай, нима кераклигини айтади. Арабча матн расм учун жуда узун бўлса, бу ҳақда айтилади.'],
    uz_latn: ['Uzun hadislar Wordga bir necha barobar tezroq qoʻyiladi: bir xil koʻrinishdagi matn boʻlaklarga boʻlinmaydi (koʻrinishi oʻsha-oʻsha).',
         'Juda uzun hadis qismlarga boʻlib qoʻyiladi, Word «qotib» qolmaydi. Word javob bermasa, 45 soniyadan keyin bu haqda aytiladi; qoʻyish tugamaguncha takror bosish hadisni ikkinchi marta qoʻymaydi.',
         'Eski Wordda (Office 2016/2019) panel «Yuklanmoqda»da toʻxtab qolmay, nima kerakligini aytadi. Arabcha matn rasm uchun juda uzun boʻlsa, bu haqda aytiladi.'] }],
  ['0.9.3', '2026-10-31', { uz: 'Янги кўриниш ва «Манбалар» саҳифаси', uz_latn: 'Yangi koʻrinish va «Manbalar» sahifasi' }, {
    uz: ['Панель ва сайт янгича: китоб саҳифасидек — оқ қоғоз, қора сиёҳ ва битта қизил ранг (қўлёзмалардаги «таҳмир» чизиғи), соя ва хира ойналар ўрнига аниқ чизиқлар. Word панелида (320–420 px) ҳамма жой матнга берилди.',
         'Ўзбекча матн учун китоб шрифти (Source Serif 4) ва интерфейс шрифти (Source Sans 3) — ўзбек кирилли (Ў Ғ Қ Ҳ) ва лотинининг (ʻ ʼ, ö ğ ş ç) ҳамма ҳарфлари билан; арабчаси — Scheherazade New. Ҳаммаси сайтнинг ўзидан юкланади.',
         'Ёруғ, қоронғи, «Қоғоз» ва «Контраст» мавзулари; матн ва фон орасидаги фарқ WCAG AA дан кам эмас. Клавиатура билан ишлаганда фокус аниқ кўринади.',
         'Янги «Манбалар» саҳифаси: 11 китобнинг нашри, таҳқиқи, ношири, жилдлари, рақамланиши ва Шамиладаги нашрга ҳавола; ҳукм берган олимлар ва асарлари. Ҳадис остида «Манба ҳақида» ҳаволаси.',
         '«Ўрнатиш» саҳифаси янгиланди: Mac ва Windows учун онлайн ва офлайн ўрнатувчилар, Терминал ва PowerShell буйруқлари («Нусха олиш» тугмаси билан), SHA256, кенгайтма, Google Docs, WordPress ва Telegram бот.',
         'Янги белги (ВАҚТИНЧА): «ح» ҳарфи ва унинг устидаги қизил чизиқ.'],
    uz_latn: ['Panel va sayt yangicha: kitob sahifasidek — oq qogʻoz, qora siyoh va bitta qizil rang (qoʻlyozmalardagi «tahmir» chizigʻi), soya va xira oynalar oʻrniga aniq chiziqlar. Word panelida (320–420 px) hamma joy matnga berildi.',
              'Oʻzbekcha matn uchun kitob shrifti (Source Serif 4) va interfeys shrifti (Source Sans 3) — oʻzbek kirilli (Ў Ғ Қ Ҳ) va lotinining (ʻ ʼ, ö ğ ş ç) hamma harflari bilan; arabchasi — Scheherazade New. Hammasi saytning oʻzidan yuklanadi.',
              'Yorugʻ, qorongʻi, «Qogʻoz» va «Kontrast» mavzulari; matn va fon orasidagi farq WCAG AA dan kam emas. Klaviatura bilan ishlaganda fokus aniq koʻrinadi.',
              'Yangi «Manbalar» sahifasi: 11 kitobning nashri, tahqiqi, noshiri, jildlari, raqamlanishi va Shamiladagi nashrga havola; hukm bergan olimlar va asarlari. Hadis ostida «Manba haqida» havolasi.',
              '«Oʻrnatish» sahifasi yangilandi: Mac va Windows uchun onlayn va oflayn oʻrnatuvchilar, Terminal va PowerShell buyruqlari («Nusxa olish» tugmasi bilan), SHA256, kengaytma, Google Docs, WordPress va Telegram bot.',
              'Yangi belgi (VAQTINCHA): «ح» harfi va uning ustidagi qizil chiziq.']
  }],
  ['0.9.2', '2026-10-31', { uz: 'Ҳадис саҳобийдан бошланади', uz_latn: 'Hadis sahobiydan boshlanadi' }, {
    uz: ['Ҳадис энди одатда уни Пайғамбар соллаллоҳу алайҳи васалламдан эшитган саҳобийдан бошлаб қўйилади: «عَنْ أَبِي هُرَيْرَةَ …», «Абу Ҳурайрадан ривоят қилинади: …». Санаднинг олдинги қисми (китоб муаллифининг устозларидан тобеъингача) қўйилмайди.',
         '«Санад» танлови: «Саҳобийдан» (одатда), «Тўлиқ санад» (аввалгидек), «Фақат матн». Word, Google Docs, нусха олиш, сават, расм ва кенгайтмада бир хил; Telegram бот ва MCP ҳам саҳобийдан бошлайди.',
         'Таржима матни ўзгармайди: фақат олдинги санад қисми тушиб қолади; саҳобий номи таржиманинг ўзидан олинади. Тахриж ва қалин ёзув аввалгидек.'],
    uz_latn: ['Hadis endi odatda uni Paygʻambar sollallohu alayhi vasallamdan eshitgan sahobiydan boshlab qoʻyiladi: «عَنْ أَبِي هُرَيْرَةَ …», «Abu Hurayradan rivoyat qilinadi: …». Sanadning oldingi qismi (kitob muallifining ustozlaridan tobeʼingacha) qoʻyilmaydi.',
              '«Sanad» tanlovi: «Sahobiydan» (odatda), «Toʻliq sanad» (avvalgidek), «Faqat matn». Word, Google Docs, nusxa olish, savat, rasm va kengaytmada bir xil; Telegram bot va MCP ham sahobiydan boshlaydi.',
              'Tarjima matni oʻzgarmaydi: faqat oldingi sanad qismi tushib qoladi; sahobiy nomi tarjimaning oʻzidan olinadi. Taxrij va qalin yozuv avvalgidek.']
  }],
  ['0.9.1', '2026-10-31', { uz: 'Тахриж ҳадисдан кейин, телефон ва офлайн тузатишлари', uz_latn: 'Taxrij hadisdan keyin, telefon va oflayn tuzatishlari' }, {
    uz: ['Тахриж энди боб сарлавҳасидан кейин эмас, ҳадис матнидан кейин туради (Аҳмад, Байҳақий, Ибн Можа, Термизий ва бошқаларда ~1 070 ҳадис): сарлавҳа кейинги ҳадисга ўтказилди.',
         'Ибн Можа таржимаси охиридаги даража («… Саҳиҳ») тахрижда ҳам бўлса, икки марта ёзилмайди; матн охиридаги «. .» ва нуқталар қатори тахриждан олдин олиб ташланади.',
         'Телефонда (360 px) карточка экрандан чиқиб кетмайди; қидирув майдонининг экран ўқувчи учун номи бор.',
         'Офлайн: панель биринчи ташрифдан кейин ҳам интернетсиз очилади; янги маълумот чиққанда эски қидирув индексига мурожаат қилинмайди, эски версиялар кэшдан тозаланади.',
         'Word лентасидаги «Ёрдам» ва «Янгиликлар» тугмалари учун саҳифа ва иконкалар қўшилди.',
         'Telegram бот ва MCP: «Öz» (янги лотин) ёзуви.'],
    uz_latn: ['Taxrij endi bob sarlavhasidan keyin emas, hadis matnidan keyin turadi (Ahmad, Bayhaqiy, Ibn Moja, Termiziy va boshqalarda ~1 070 hadis): sarlavha keyingi hadisga oʻtkazildi.',
              'Ibn Moja tarjimasi oxiridagi daraja («… Sahih») taxrijda ham boʻlsa, ikki marta yozilmaydi; matn oxiridagi «. .» va nuqtalar qatori taxrijdan oldin olib tashlanadi.',
              'Telefonda (360 px) kartochka ekrandan chiqib ketmaydi; qidiruv maydonining ekran oʻquvchi uchun nomi bor.',
              'Oflayn: panel birinchi tashrifdan keyin ham internetsiz ochiladi; yangi maʼlumot chiqqanda eski qidiruv indeksiga murojaat qilinmaydi, eski versiyalar keshdan tozalanadi.',
              'Word lentasidagi «Yordam» va «Yangiliklar» tugmalari uchun sahifa va ikonkalar qoʻshildi.',
              'Telegram bot va MCP: «Öz» (yangi lotin) yozuvi.']
  }],
  ['0.9.0', '2026-10-30', { uz: '«Кенгайтириш», бир неча парча ва китоблар тартиби', uz_latn: '«Kengaytirish», bir necha parcha va kitoblar tartibi' }, {
    uz: ['«Кенгайтириш»: ҳадис ва таржима экран ўртасида катта ҳарфда очилади (A− / A+), узун ҳадис ичида варақланади; сўзлар шу ерда танланади, «Қўйиш» ва «Нусха олиш» ҳам шу ерда. Esc, × ёки ташқарига босиш билан ёпилади.',
         'Парча танлаш: арабча ва ўзбекча матнда ҳам. Ctrl (Mac: ⌘) билан бир неча жойдан танланади — тартиби билан, ораси «…» билан қўйилади, тахриж қолади; Shift — парчани узайтириш; «Танловни бекор қилиш».',
         'Сўз танлаганда ва режим алмашганда саҳифа энди сакрамайди. Созламада «Равон ўтишлар» (тизимда «ҳаракатни камайтириш» ёқилган бўлса — ўчиқ).',
         '«Кун ҳадиси»ни ёпиш ёки ўчириш энди Word, Google Docs ва кенгайтмада ҳам, янгиланишдан кейин ҳам сақланади; «Бугунга яшириш» — эртагача.',
         'Қидирув натижалари китоблар тартибида: Бухорий, Муслим, Абу Довуд, Термизий, Насоий, Ибн Можа, Аҳмад, Байҳақий, Ҳоким, Табароний, Дорақутний; китоб ичида рақам бўйича. Энг мос учтаси белгиланади. Созламада «Саралаш: мослик бўйича».',
         'Номи ҳамма жойда «Ҳадислар» (Word, кенгайтма, Google Docs, WordPress).'],
    uz_latn: ['«Kengaytirish»: hadis va tarjima ekran oʻrtasida katta harfda ochiladi (A− / A+), uzun hadis ichida varaqlanadi; soʻzlar shu yerda tanlanadi, «Qoʻyish» va «Nusxa olish» ham shu yerda. Esc, × yoki tashqariga bosish bilan yopiladi.',
              'Parcha tanlash: arabcha va oʻzbekcha matnda ham. Ctrl (Mac: ⌘) bilan bir necha joydan tanlanadi — tartibi bilan, orasi «…» bilan qoʻyiladi, taxrij qoladi; Shift — parchani uzaytirish; «Tanlovni bekor qilish».',
              'Soʻz tanlaganda va rejim almashganda sahifa endi sakramaydi. Sozlamada «Ravon oʻtishlar» (tizimda «harakatni kamaytirish» yoqilgan boʻlsa — oʻchiq).',
              '«Kun hadisi»ni yopish yoki oʻchirish endi Word, Google Docs va kengaytmada ham, yangilanishdan keyin ham saqlanadi; «Bugunga yashirish» — ertagacha.',
              'Qidiruv natijalari kitoblar tartibida: Buxoriy, Muslim, Abu Dovud, Termiziy, Nasoiy, Ibn Moja, Ahmad, Bayhaqiy, Hokim, Tabaroniy, Doraqutniy; kitob ichida raqam boʻyicha. Eng mos uchtasi belgilanadi. Sozlamada «Saralash: moslik boʻyicha».',
              'Nomi hamma joyda «Hadislar» (Word, kengaytma, Google Docs, WordPress).']
  }],
  ['0.8.0', '2026-10-30', { uz: 'Уч ёзув ва арабча қидирув', uz_latn: 'Uch yozuv va arabcha qidiruv' }, {
    uz: ['Ёзувни танлаш тепада: «Öz» — янги лотин (ö ğ ş ç), «Ўз» — кирилл, «Oʼz» — лотин. Созламаларда «Стандарт ёзув» ва лотиндаги апостроф шакли. Танлов Word, Google Docs ва кенгайтмада ҳам сақланади.',
         'Ҳадисларнинг лотинча матни Harfsoz қоидалари ва луғати билан тайёрланган (89 215 ҳадис); тахриж, китоб ва олим номлари, баҳолар, «Кун ҳадиси» ва саҳифалар ҳам танланган ёзувда. Қидирув учала ёзувдаги сўровни тушунади.',
         'Арабча қидирув ҳаракат, алиф (أ إ آ ٱ / ا), ى/ي, ة/ه, ؤ/ئ, татвил, Қуръон белгилари, ﷺ ва «صلى الله عليه وسلم», бошқа бўшлиқларга қарамайди; тез-тез учрайдиган сўзлардан иборат ибора ҳам топилади.'],
    uz_latn: ['Yozuvni tanlash tepada: «Öz» — yangi lotin (ö ğ ş ç), «Ўз» — kirill, «Oʼz» — lotin. Sozlamalarda «Standart yozuv» va lotindagi apostrof shakli. Tanlov Word, Google Docs va kengaytmada ham saqlanadi.',
              'Hadislarning lotincha matni Harfsoz qoidalari va lugʻati bilan tayyorlangan (89 215 hadis); taxrij, kitob va olim nomlari, baholar, «Kun hadisi» va sahifalar ham tanlangan yozuvda. Qidiruv uchala yozuvdagi soʻrovni tushunadi.',
              'Arabcha qidiruv harakat, alif (أ إ آ ٱ / ا), ى/ي, ة/ه, ؤ/ئ, tatvil, Qurʼon belgilari, ﷺ va «صلى الله عليه وسلم», boshqa boʻshliqlarga qaramaydi; tez-tez uchraydigan soʻzlardan iborat ibora ham topiladi.']
  }],
  ['0.7.0', '2026-10-30', { uz: 'Тахрижнинг янги шакли ва қалин ёзув', uz_latn: 'Taxrijning yangi shakli va qalin yozuv' }, {
    uz: ['Тахриж ҳадисдан кейин, қавс ичида ва курсивда, нуқта қавсдан кейин: «… («Саҳиҳи Бухорий», 1).», «… («Абу Довуд», 2226; Ҳадиснинг даражаси: «саҳиҳ». Манба: «Саҳиҳи Абу Довуд», 2226).»',
         'Даража фақат маълумотдан: Албоний («Саҳиҳи …» / «Заифи …»), Арнаут («Муснад тахрижи», «Абу Довуд тахрижи»), Хайсамий, Заҳабий, Бусирий, улар бўлмаса — китоб муаллифининг ўз сўзи. Бухорий ва Муслимда даража қўйилмайди. Созламада «Барча баҳолар».',
         'Қалин ёзув: Пайғамбар ﷺ сўзлари бўлса — фақат улар (арабчада ҳам); бўлмаса — саҳоба ёки тобеъийнинг сўзлари. Чегараси аниқ бўлмаса, ҳеч нарса қалин қилинмайди.',
         'Таржимадаги қўштирноқлар доим «…». Арабча тахриж: «(رواه البخاري ١)».'],
    uz_latn: ['Taxrij hadisdan keyin, qavs ichida va kursivda, nuqta qavsdan keyin: «… («Sahihi Buxoriy», 1).», «… («Abu Dovud», 2226; Hadisning darajasi: «sahih». Manba: «Sahihi Abu Dovud», 2226).»',
              'Daraja faqat maʼlumotdan: Alboniy («Sahihi …» / «Zaifi …»), Arnaut («Musnad taxriji», «Abu Dovud taxriji»), Xaysamiy, Zahabiy, Busiriy, ular boʻlmasa — kitob muallifining oʻz soʻzi. Buxoriy va Muslimda daraja qoʻyilmaydi. Sozlamada «Barcha baholar».',
              'Qalin yozuv: Paygʻambar ﷺ soʻzlari boʻlsa — faqat ular (arabchada ham); boʻlmasa — sahoba yoki tobeʼiyning soʻzlari. Chegarasi aniq boʻlmasa, hech narsa qalin qilinmaydi.',
              'Tarjimadagi qoʻshtirnoqlar doim «…». Arabcha taxrij: «(رواه البخاري ١)».']
  }],
  ['0.5.0', '2026-10-30', { uz: 'Олимлар баҳоси', uz_latn: 'Olimlar bahosi' }, {
    uz: ['Карточкада тахриж остида «Бошқа олимлар баҳоси»: Албоний ва Арнаут (Абу Довуд, Термизий, Насоий, Ибн Можа), Заҳабий («Талхис», Ҳоким), Бусирий («Завоид», Ибн Можа) — олим номи, баҳо (арабчаси билан), асар ва рақам. Муаллифнинг ўз сўзидан алоҳида.',
         'Баҳолар очиқ манбадан, ҳадис матни бўйича солиштирилган (рақам фақат тасдиқ); ишончсиз мослик қўйилмаган. 22 600 дан ортиқ ҳадисда баҳо бор; Аҳмад, Байҳақий, Табароний, Дорақутний учун ҳозирча йўқ.',
         'Қўйишда тахриждан кейин қўшилади; «Фақат матн»да қўшилмайди. Созламалар: «Олимлар баҳосини қўшиш», «„Фақат матн“да ҳам …».'],
    uz_latn: ['Kartochkada taxrij ostida «Boshqa olimlar bahosi»: Alboniy va Arnaut (Abu Dovud, Termiziy, Nasoiy, Ibn Moja), Zahabiy («Talxis», Hokim), Busiriy («Zavoid», Ibn Moja) — olim nomi, baho (arabchasi bilan), asar va raqam. Muallifning oʻz soʻzidan alohida.',
              'Baholar ochiq manbadan, hadis matni boʻyicha solishtirilgan (raqam faqat tasdiq); ishonchsiz moslik qoʻyilmagan. 22 600 dan ortiq hadisda baho bor; Ahmad, Bayhaqiy, Tabaroniy, Doraqutniy uchun hozircha yoʻq.',
              'Qoʻyishda taxrijdan keyin qoʻshiladi; «Faqat matn»da qoʻshilmaydi. Sozlamalar: «Olimlar bahosini qoʻshish», «„Faqat matn“da ham …».']
  }],
  ['0.4.0', '2026-10-29', { uz: 'Фақат ўзбек тили', uz_latn: 'Faqat oʻzbek tili' }, {
    uz: ['Интерфейс энди фақат ўзбекча: кирилл (асосий) ёки лотин. Ёзувни юқоридаги «Кир | Lot» тугмаси ёки Созламалардаги «Ёзув» алмаштиради — интерфейс ва таржима бир хил ёзувда бўлади.',
         'Бошқа тиллар олиб ташланди: панель, ёрдам, махфийлик, шартлар, ўрнатиш ва нашрлар саҳифалари, браузер кенгайтмаси, WordPress плагини, Telegram бот ва MCP сервер.',
         'Ҳадисларнинг арабча матни ва «⁧رواه البخاري (١)⁩» шаклидаги арабча тахриж аввалгидек қолди.'],
    uz_latn: ['Interfeys endi faqat oʻzbekcha: kirill (asosiy) yoki lotin. Yozuvni yuqoridagi «Kir | Lot» tugmasi yoki Sozlamalardagi «Yozuv» almashtiradi — interfeys va tarjima bir xil yozuvda boʻladi.',
              'Boshqa tillar olib tashlandi: panel, yordam, maxfiylik, shartlar, oʻrnatish va nashrlar sahifalari, brauzer kengaytmasi, WordPress plagini, Telegram bot va MCP server.',
              'Hadislarning arabcha matni va «⁧رواه البخاري (١)⁩» shaklidagi arabcha taxrij avvalgidek qoldi.']
  }],
  ['0.3.0', '2026-10-29', { uz: 'Иқтибосни текшириш', uz_latn: 'Iqtibosni tekshirish' }, {
    uz: ['«Яна» менюсида «Иқтибосни текшириш»: ҳадис матнини қўйинг (арабча ҳаракатли ёки ҳаракатсиз, ўзбекча кирилл ёки лотин) — панель уни базадан топади ва сўзма-сўз солиштиради.',
         'Фарқлар белгиланади: ортиқча, тушиб қолган, хато ёки бошқа сўз; ҳаракат (ўзбекчада имло) фарқи ҳарфлардан алоҳида. Баҳо: «тўлиқ мос» ёки «N та сўз фарқ қилади».',
         'Тахриж ҳаволаси ҳадисни очади; «Базадаги матн билан алмаштириш» тўғри матнни тахриж билан қўяди (Word’да белгиланган жой ўрнига).'],
    uz_latn: ['«Yana» menyusida «Iqtibosni tekshirish»: hadis matnini qoʻying (arabcha harakatli yoki harakatsiz, oʻzbekcha kirill yoki lotin) — panel uni bazadan topadi va soʻzma-soʻz solishtiradi.',
              'Farqlar belgilanadi: ortiqcha, tushib qolgan, xato yoki boshqa soʻz; harakat (oʻzbekchada imlo) farqi harflardan alohida. Baho: «toʻliq mos» yoki «N ta soʻz farq qiladi».',
              'Taxrij havolasi hadisni ochadi; «Bazadagi matn bilan almashtirish» toʻgʻri matnni taxrij bilan qoʻyadi (Word’da belgilangan joy oʻrniga).']
  }],
  ['0.2.0', '2026-10-29', { uz: '«Кун ҳадиси»', uz_latn: '«Kun hadisi»' }, {
    uz: ['Бош саҳифада «Кун ҳадиси»: ҳар куни Саҳиҳул Бухорий ёки Саҳиҳи Муслимдан битта қисқа ҳадис (фақат матн), бир кунда ҳамма учун бир хил.',
         'Қўйиш, нусха олиш, сараланганларга қўшиш; «Тўлиқ ҳадис» уни санади билан очади.',
         'Ёпганда: бугунга яшириш ёки бутунлай ўчириш (Созламалар, «Асосий»).'],
    uz_latn: ['Bosh sahifada «Kun hadisi»: har kuni Sahihul Buxoriy yoki Sahihi Muslimdan bitta qisqa hadis (faqat matn), bir kunda hamma uchun bir xil.',
              'Qoʻyish, nusxa olish, saralanganlarga qoʻshish; «Toʻliq hadis» uni sanadi bilan ochadi.',
              'Yopganda: bugunga yashirish yoki butunlay oʻchirish (Sozlamalar, «Asosiy»).']
  }],
  ['0.1.0', '2026-10-28', { uz: 'Биринчи синов нашри', uz_latn: 'Birinchi sinov nashri' }, {
    uz: ['11 китоб, 89 215 ҳадис: арабча матн (ҳаракатлари билан) ва ўзбекча таржима.',
         'Қидирув: «Бухорий 1», «Муслим 2564/2», «بخاري ١» каби ҳавола, арабча (ҳаракатли ёки ҳаракатсиз) ва ўзбекча (кирилл ёки лотин) матн бўйича.',
         'Тахриж: «(Имом Бухорий ривояти, 1-ҳадис)», «⁧رواه البخاري (١)⁩»; ҳукм фақат китоб муаллифининг ўз сўзидан, ким айтгани билан.',
         'Фақат матн ёки санад билан, сўзма-сўз парча танлаш; Word, Google Docs, WordPress ва браузерга қўйиш, нусха олиш, расм.',
         'Тарих, сараланганлар, сават (манбалар умумий рўйхати билан), кирилл / лотин, 13 мавзу.',
         'Интерфейс 22 тилда. Ёрдам, махфийлик, шартлар, ўрнатиш ва нашрлар саҳифалари.'],
    uz_latn: ['11 kitob, 89 215 hadis: arabcha matn (harakatlari bilan) va oʻzbekcha tarjima.',
              'Qidiruv: «Buxoriy 1», «Muslim 2564/2», «بخاري ١» kabi havola, arabcha (harakatli yoki harakatsiz) va oʻzbekcha (kirill yoki lotin) matn boʻyicha.',
              'Taxrij: «(Imom Buxoriy rivoyati, 1-hadis)», «⁧رواه البخاري (١)⁩»; hukm faqat kitob muallifining oʻz soʻzidan, kim aytgani bilan.',
              'Faqat matn yoki sanad bilan, soʻzma-soʻz parcha tanlash; Word, Google Docs, WordPress va brauzerga qoʻyish, nusxa olish, rasm.',
              'Tarix, saralanganlar, savat (manbalar umumiy roʻyxati bilan), kirill / lotin, 13 mavzu.',
         'Interfeys 22 tilda. Yordam, maxfiylik, shartlar, oʻrnatish va nashrlar sahifalari.']
  }]
];
