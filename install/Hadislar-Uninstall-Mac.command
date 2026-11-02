#!/bin/bash
# Ҳадислар қўшимчасини Mac'дан олиб ташлаш (шрифтга тегмайди).
#   curl -fsSL https://hadis.abdulloh.org/mac-uninstall.sh | bash
F="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef/HadisUz.xml"
if [ -f "$F" ]; then mkdir -p "$HOME/.Trash" && mv "$F" "$HOME/.Trash/HadisUz.xml.$(date +%s)"; echo "✓ олиб ташланди (Корзинага)"; else echo "Қўшимча ўрнатилмаган."; fi
echo "Word'ни ёпиб қайта очинг."
if [ -t 0 ]; then echo; read -r -p "Ёпиш учун Enter'ни босинг. " _; fi
