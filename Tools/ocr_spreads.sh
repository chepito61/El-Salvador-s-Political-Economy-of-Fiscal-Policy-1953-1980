#!/bin/bash
# Narrative pages: two-page spreads scanned sideways. Usage: ocr_spreads.sh NN  (run from the work dir with p-NN.png at 200 dpi)
# Picks 90 or 270 by Spanish stop-word count, rotates, splits into halves, OCRs each half.
n=$1; best=-1; mkdir -p h
for a in 90 270; do convert p-$n.png -rotate $a -resize 50% /tmp/q${n}_$a.png; tesseract /tmp/q${n}_$a.png /tmp/q${n}_$a -l spa --psm 6 >/dev/null 2>&1
c=$(grep -o -w -i -E "de|la|el|en|los|las|del|que|por|con|para|total|gastos" /tmp/q${n}_$a.txt | wc -l); if [ $c -gt $best ]; then best=$c; ba=$a; fi; rm -f /tmp/q${n}_$a.*; done
echo $ba > h/ang$n; convert p-$n.png -rotate $ba h/r$n.png
w=$(identify -format %w h/r$n.png); hh=$(identify -format %h h/r$n.png); half=$((w/2))
convert h/r$n.png -crop ${half}x${hh}+0+0 +repage h/${n}L.png; convert h/r$n.png -crop $((w-half))x${hh}+${half}+0 +repage h/${n}R.png
tesseract h/${n}L.png h/${n}L -l spa --psm 6 >/dev/null 2>&1; tesseract h/${n}R.png h/${n}R -l spa --psm 6 >/dev/null 2>&1
