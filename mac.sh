#!/bin/bash
# Ҳадислар — Word учун қўшимчани Mac'га ўрнатиш (администратор ҳуқуқи керак эмас).
# Terminal'да:   curl -fsSL https://hadis.abdulloh.org/mac.sh | bash
# Тузатиш (тоза ўрнатиш, Word ишламаса/қотса):  curl -fsSL https://hadis.abdulloh.org/mac.sh | bash -s -- --clean
#                (--dry-run: фақат рўйхат, ҳеч нарса кўчирилмайди)
# Ўчириш:        curl -fsSL https://hadis.abdulloh.org/mac-uninstall.sh | bash
# Офлайн/синов:  HADIS_SRC=<папка>  (ичида manifest.xml, ихтиёрий ScheherazadeNew-*.ttf)
# Нима қилади: (1) Word'нинг wef папкасига manifest.xml'ни HadisUz.xml номи билан қўяди, (2) "Scheherazade New"
# шрифти йўқ бўлса ~/Library/Fonts'га ўрнатади (арабча матн учун, лицензия OFL). Word'ни ёпиб қайта очинг.
# --clean: аввал қўшимчанинг ҳамма изларини (фақат ўзиники: Id ва домен бўйича) Корзинадаги заҳира папкага
# кўчиради, кейин одатдагидек ўрнатади. Ҳужжатлар, бошқа қўшимчалар, бошқа сайтлар ва шрифтга тегмайди.
set -e
SITE="${HADIS_SITE:-https://hadis.abdulloh.org}"
MIRROR="https://abdullohukr.github.io/hadis"
HOST="hadis.abdulloh.org"                      # панел очиладиган сайт (addin/manifest.xml SourceLocation)
ID="044de12a-604e-4a28-aa7c-acb3b787f180"      # addin/manifest.xml <Id> (tools/build_installers.sh текширади)
SRC="$HADIS_SRC"
WORD="$HOME/Library/Containers/com.microsoft.Word/Data"
WEF="$WORD/Documents/wef"
WORDPROC="${HADIS_WORD_PROC:-Microsoft Word}"  # синов учун бошқа ном (Word'га тегмасдан)
CLEAN="" DRY=""
for a in "$@"; do
  case "$a" in
    --clean) CLEAN=1 ;;
    --dry-run) CLEAN=1 DRY=1 ;;
    *) printf 'Номаълум параметр: %s (бор: --clean, --dry-run)\n' "$a"; exit 3 ;;
  esac
done
TMP="$(mktemp -d)"; trap 'rm -rf "$TMP"' EXIT
say() { printf '%s\n' "$*"; }
[ "$(uname)" = "Darwin" ] || { say "Бу скрипт фақат macOS учун."; exit 1; }
[ "$(id -u)" != 0 ] || { say "Буни sudo'сиз, оддий фойдаланувчи сифатида ишга туширинг."; exit 1; }
[ -d "/Applications/Microsoft Word.app" ] || say "Диққат: Microsoft Word.app топилмади (/Applications). Ўрнатишни давом эттираман."
word_running() { pgrep -x "$WORDPROC" >/dev/null 2>&1; }

get() {  # get <файл> <чиқиш>
  if [ -n "$SRC" ] && [ -f "$SRC/$1" ]; then cp -f "$SRC/$1" "$2"; return 0; fi
  if [ -n "$SRC" ] && [ -f "$SRC/${1##*/}" ]; then cp -f "$SRC/${1##*/}" "$2"; return 0; fi   # fonts/X.ttf или просто X.ttf
  curl -fsSL --retry 3 --retry-delay 2 "$SITE/$1" -o "$2" 2>/dev/null || curl -fsSL --retry 3 --retry-delay 2 "$MIRROR/$1" -o "$2"
}

# ff: find, «Interrupted system call» бўлса яна (Word очиқ бўлса папкалар ўзгариб туради); такрорлар кейин олиб ташланади
ff() { local i; for i in 1 2 3; do find "$@" 2>/dev/null && return 0; done; return 0; }

# ---- тузатиш: қўшимчанинг изларини заҳирага кўчириш (ўчирмайди) ----
clean() {
  set +e   # излаш (find/grep) «топилмади» деса ҳам давом этади
  if [ -z "$DRY" ] && word_running; then
    say "Word очиқ. Тузатишдан олдин уни тўлиқ ёпинг (Cmd+Q). Сақланмаган ҳужжатларни аввал сақланг."
    if ! { exec 3</dev/tty; } 2>/dev/null; then say "Word'ни ёпиб, буйруқни қайта ишга туширинг."; exit 1; fi
    while word_running; do
      printf "Word'ни ёпиб, Enter'ни босинг (бекор қилиш: q ва Enter): "
      read -r ans <&3 || exit 1
      case "$ans" in q|Q|й|Й) say "Бекор қилинди, ҳеч нарса ўзгармади."; exit 1 ;; esac
    done
    exec 3<&-
  fi
  local stamp; stamp="$(date +%Y%m%d-%H%M%S)"
  BK="$HOME/.Trash/Hadislar-backup-$stamp"
  if [ -z "$DRY" ] && ! mkdir -p "$BK" 2>/dev/null; then BK="$HOME/Hadislar-backup-$stamp"; mkdir -p "$BK"; fi
  local L="$WORD/Library" LIST="$TMP/list" OSF="$HOME/Library/Containers/com.Microsoft.OsfWebHost/Data/Library"
  : > "$LIST"
  # 1) Word'нинг wef папкасидаги бизнинг манифестлар (HadisUz.xml ва эски номлар: ичида Id бўлса)
  if [ -d "$WEF" ]; then
    ff "$WEF" -maxdepth 1 -type f \( -name 'HadisUz.xml*' -o -name '*.xml' \) -print0 | while IFS= read -r -d '' f; do
      case "${f##*/}" in HadisUz.xml*) printf '%s\n' "$f" ;; *) grep -qiF "$ID" "$f" 2>/dev/null && printf '%s\n' "$f" ;; esac
    done >> "$LIST"
  fi
  # 2) Office қўшимчалар кеши (Wef): номида Id бор файл/папкалар; ичида Id ёки домен бор кичик файллар
  #    (умумий кўрсаткичлар AppCommands, AggregatedCache, MetaOS — бошқа қўшимчаларники ҳам, тегилмайди)
  for W in "$L/Application Support/Microsoft/Office/16.0/Wef"; do
    [ -d "$W" ] || continue
    ff "$W" -iname "*$ID*" -prune -print >> "$LIST"
    ff "$W" \( -path "$W/AppCommands" -o -path "$W/AggregatedCache" -o -path "$W/MetaOS" -o -iname "*$ID*" \) -prune -o \
      -type f -size -1024k -print0 | while IFS= read -r -d '' f; do
      grep -qaiF -e "$ID" -e "$HOST" "$f" 2>/dev/null && printf '%s\n' "$f"
    done >> "$LIST"
  done
  # 3) панел маълумотлари (WebKit): localStorage, IndexedDB, Cache Storage, Service Worker — фақат шу сайт папкалари
  #    (WebsiteData/Default/<..>/<..>/origin ичида «https hadis.abdulloh.org»), эски номлар https_<сайт>_0*,
  #    тармоқ кешидан шу сайт ёзувлари
  for R in "$L/WebKit" "$L/Caches/WebKit" "$OSF/WebKit" "$OSF/Caches/WebKit"; do
    [ -d "$R" ] || continue
    ff "$R" -type f -name origin -size -8k -print0 | while IFS= read -r -d '' f; do
      perl -0777 -ne 'exit(/\x01\Q'"$HOST"'\E\x00/ ? 0 : 1)' "$f" 2>/dev/null && printf '%s\n' "${f%/origin}"
    done >> "$LIST"
    ff "$R" \( -name "https_${HOST}_*" -o -name "http_${HOST}_*" \) -prune -print >> "$LIST"
    [ -d "$R/NetworkCache" ] && ff "$R/NetworkCache" -path '*/Records/*' -type f -print0 | while IFS= read -r -d '' f; do
      grep -qaF "://$HOST/" "$f" 2>/dev/null && printf '%s\n' "$f"
    done >> "$LIST"
  done
  # рўйхат: такрорсиз, бир-бирининг ичидагилар олиб ташланади
  LC_ALL=C sort -u "$LIST" | awk 'NF { for (i = 1; i <= n; i++) if (index($0, k[i] "/") == 1) next; k[++n] = $0; print }' > "$LIST.u"
  local n=0 rel
  say "Тузатиш: «Ҳадислар» қўшимчасининг излари$([ -n "$DRY" ] && printf ' (синов: ҳеч нарса кўчирилмайди)'):"
  while IFS= read -r p; do
    [ -e "$p" ] || continue
    rel="${p#"$HOME"/}"; say "  - ~/$rel"; n=$((n + 1))
    [ -n "$DRY" ] && continue
    mkdir -p "$BK/$(dirname "$rel")" && mv "$p" "$BK/$rel" || { say "    ! кўчирилмади"; continue; }
    case "$p" in */WebsiteData/Default/*/*) rmdir "$(dirname "$p")" 2>/dev/null || true ;; esac   # бўш қолган юқори папка
  done < "$LIST.u"
  # 4) Word'нинг HTTP кеши (Cache.db, SQLite, бошқа сайтлар ҳам шу ерда): фақат шу сайт сатрлари; файл олдин заҳирага
  local DB="$L/Caches/com.microsoft.Word/Cache.db" q rows
  q="request_key LIKE 'https://$HOST/%' OR request_key LIKE 'http://$HOST/%'"
  if [ -f "$DB" ] && command -v sqlite3 >/dev/null; then
    rows="$(sqlite3 -readonly "$DB" "SELECT count(*) FROM cfurl_cache_response WHERE $q" 2>/dev/null || echo 0)"
    if [ "${rows:-0}" -gt 0 ]; then
      say "  - ~/${DB#"$HOME"/}: шу сайтнинг $rows та ёзуви"; n=$((n + 1))
      if [ -z "$DRY" ]; then
        local d="$BK/${DB#"$HOME"/}"; mkdir -p "$(dirname "$d")"
        for x in "$DB" "$DB-wal" "$DB-shm"; do [ -f "$x" ] && cp -p "$x" "$(dirname "$d")/"; done
        sqlite3 "$DB" "SELECT receiver_data FROM cfurl_cache_receiver_data WHERE isDataOnFS = 1 AND entry_ID IN (SELECT entry_ID FROM cfurl_cache_response WHERE $q)" 2>/dev/null |
          while IFS= read -r fsd; do
            case "$fsd" in ''|*/*) continue ;; esac
            [ -f "$(dirname "$DB")/fsCachedData/$fsd" ] && mv "$(dirname "$DB")/fsCachedData/$fsd" "$(dirname "$d")/" || true
          done
        sqlite3 "$DB" "BEGIN; DELETE FROM cfurl_cache_blob_data WHERE entry_ID IN (SELECT entry_ID FROM cfurl_cache_response WHERE $q);
          DELETE FROM cfurl_cache_receiver_data WHERE entry_ID IN (SELECT entry_ID FROM cfurl_cache_response WHERE $q);
          DELETE FROM cfurl_cache_response WHERE $q; COMMIT;" || say "! Cache.db тозаланмади (заҳира нусхаси бор)"
      fi
    fi
  fi
  if [ "$n" = 0 ]; then say "  (ҳеч нарса топилмади — изи йўқ)"; [ -n "$DRY" ] || rmdir "$BK" 2>/dev/null || true
  elif [ -z "$DRY" ]; then say "✓ $n та нарса заҳирага кўчирилди: $BK"; say "  (тиклаш: ўша папкадаги файлларни ~/ ичидаги жойига қайтаринг; керак бўлмаса Корзинани тозаланг)"; fi
  say "Тегилмади: ҳужжатларингиз, бошқа қўшимчалар ва сайтлар маълумотлари, Scheherazade New шрифти, Word куки файли (умумий)."
  say ""
  set -e
}

if [ -n "$CLEAN" ]; then clean; [ -z "$DRY" ] || exit 0; say "Энди тоза ўрнатиш:"; fi
if word_running; then say "Диққат: Word очиқ. Ўрнатгандан кейин уни ёпиб, қайта очинг."; fi

# 1) қўшимча манифести
get manifest.xml "$TMP/manifest.xml" || { say "Хато: manifest.xml юкланмади."; exit 1; }
grep -q "<OfficeApp" "$TMP/manifest.xml" || { say "Хато: manifest.xml яроқсиз."; exit 1; }
mkdir -p "$WEF"
cp -f "$TMP/manifest.xml" "$WEF/HadisUz.xml"
NC="$WORD/Library/Caches/WebKit/NetworkCache"   # Word'нинг тармоқ кешидан фақат шу сайтнинг эски ёзувлари
[ ! -d "$NC" ] || ff "$NC" -path '*/Records/*' -type f -exec grep -qaF "://$HOST/" {} \; -print | while IFS= read -r f; do rm -f "$f"; done
say "✓ Word қўшимчаси: $WEF/HadisUz.xml"

# 2) шрифт
if fc-list 2>/dev/null | grep -qi "Scheherazade New" || ls "$HOME/Library/Fonts" /Library/Fonts 2>/dev/null | grep -qi "ScheherazadeNew"; then
  say "✓ Scheherazade New шрифти аллақачон бор"
else
  mkdir -p "$HOME/Library/Fonts"
  for f in ScheherazadeNew-Regular.ttf ScheherazadeNew-Bold.ttf; do
    if get "fonts/$f" "$TMP/$f"; then cp -f "$TMP/$f" "$HOME/Library/Fonts/$f"; else say "! $f юкланмади"; fi
  done
  say "✓ Scheherazade New шрифти ўрнатилди"
fi

say ""
say "Тайёр. Word'ни тўлиқ ёпинг (Cmd+Q) ва қайта очинг. Лентада «Ҳадислар» ойнаси чиқади."
say "Агар кўринмаса: Word → Insert (Қўшиш) → Add-ins → My Add-ins → Shared Folder / Developer Add-ins."
