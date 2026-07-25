#!/bin/sh

for i in $(ls ./*.svg)
do
#    file=$(echo "${i}" | sed -e 's/_.*.svg//g')
    file=$(echo ${i%.svg})
#    echo $file
    icona1=$file"_16.bmp"
    icona2=$file"_16h.bmp"
    icona3=$file"_26.bmp"
    icona4=$file"_26h.bmp"
    cp $i ../$icona1
    cp $i ../$icona2
    cp $i ../$icona3
    cp $i ../$icona4
done
