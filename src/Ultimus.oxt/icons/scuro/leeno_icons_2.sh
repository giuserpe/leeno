#!/bin/bash

for file in ./*.svg
do
    echo ../${file%_*}_{1,2}6{,h}.bmp | xargs -n 1 cp "$file"
done
