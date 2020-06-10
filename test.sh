#!/bin/bash

# execution_path="`dirname \"$0\"`"
# execution_path="`( cd \"$execution_path\" && pwd )`"
# addon_src_path=$execution_path/src/Ultimus.oxt
# addon_bin_path=$execution_path/bin
# 
# addon_name=LeenO
# 
# unopkg_bin=/usr/bin/unopkg
oocalc_bin=/usr/lib/libreoffice/program/scalc
oowriter_bin=/usr/lib/libreoffice/program/swriter
# 
# addon_files=$addon_src_path/*
# oxt_file=$addon_name.oxt
# 
# #remove any previous package
# rm $addon_bin_path/$oxt_file
# 
# #create the add-on package
# cd $addon_src_path
# zip -r $addon_bin_path/$oxt_file *
# 
# #remove the previous extension
# $unopkg_bin remove $oxt_file
# 
# #add the new one
# echo "sì" | $unopkg_bin add $addon_bin_path/$oxt_file

#enable debug log
#export PYUNO_LOGLEVEL=CALL
export PYUNO_LOGLEVEL=ARGS
export PYUNO_LOGTARGET=stdout

#launch LO CALC
$oocalc_bin
