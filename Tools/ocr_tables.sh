#!/bin/bash
# Statement pages (Estados): one wide table per page, scanned sideways. Usage: ocr_tables.sh NN
# Picks 90 or 270 by count of number patterns and keywords; keeps rotated image t/RNN.png and text t/fNN.txt.
n=$1; best=-1; mkdir -p t
for a in 90 270; do convert p-$n.png -rotate $a t/r${n}_$a.png; tesseract t/r${n}_$a.png t/x${n}_$a -l spa --psm 6 >/dev/null 2>&1
c=$(grep -o -E "[0-9]{1,3}\.[0-9]{3}\.[0-9]{2}\b|\b(DE|LA|EL|DEL|TOTAL|RAMO|Impuesto|de|la|del|Programa)\b" t/x${n}_$a.txt | wc -l); if [ $c -gt $best ]; then best=$c; ba=$a; fi; done
cp t/x${n}_$ba.txt t/f$n.txt; mv t/r${n}_$ba.png t/R$n.png; rm -f t/r${n}_*.png
