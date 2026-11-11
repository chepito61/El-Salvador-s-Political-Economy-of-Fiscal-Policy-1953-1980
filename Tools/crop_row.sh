#!/bin/bash
# Crop the band around the first word matching a regex, to read the figures on the image.
# Usage: crop_row.sh NN 'regex' out.png [pixels_above] [pixels_below]   (uses t/RNN.png)
# Trust the tsv y-coordinate, not the displayed scale of the image.
n=$1; w=$2; o=$3; a=${4:-40}; b=${5:-60}
[ -f t/R$n.tsv ] || tesseract t/R$n.png t/R$n -l spa --psm 6 tsv >/dev/null 2>&1
y=$(awk -F'\t' -v w="$w" 'NR>1 && $12 ~ w {print $8; exit}' t/R$n.tsv)
W=$(identify -format %w t/R$n.png); echo "y=$y"; convert t/R$n.png -crop ${W}x$((a+b))+0+$((y-a)) +repage $o
