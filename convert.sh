#!/usr/bin/env bash

set -euo pipefail

args="$*"

for file in $args; do
  b=`basename $file .pptx`
  ../pptx-to-yaml.py $file ${b}.yaml ${b}_figures
  ../yaml-to-md.py ${b}.yaml ${b}.md
  for svg_file in ${b}_figures/*.svg; do
    if [ -x $svg_file ] ; then
      svg_base=`basename $svg_file .svg`
      inkscape --export-eps=${b}_figures/${svg_base}.eps $svg_file
    fi
  done
  sed -i.bak '/\!\[/ s/\.svg)/.eps)/' ${b}.md 
  pandoc -t beamer -o ${b}.tex ${b}.md
  pandoc -t beamer -o ${b}.pdf ${b}.md || true
done

