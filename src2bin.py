#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse, os, zipfile

with open('src/Ultimus.oxt/leeno_version_code', 'r') as file:
    last_version = file.read().rstrip()

parser = argparse.ArgumentParser(description='Process some integers.')
parser.add_argument('-e', '--ext', default='oxt')
parser.add_argument('-V', '--version', default=last_version)
args = parser.parse_args()

def compress():
    """
    Compress all directory under src in LeenO.oxt extension for LibreOffice
    """
    for rootname in os.walk('src').__next__()[1]:
        archname = 'LeenO.oxt'
        fileName, fileExtension = os.path.splitext(archname)
        if args.version:
            fileName += '_%s' % args.version
        if args.ext:
            fileExtension = args.ext
        archname = '%s.%s' % (fileName, fileExtension)
        archpath = os.path.join('bin', archname)
        if not os.path.isdir('bin'):
            os.mkdir('bin')
        with zipfile.ZipFile(archpath, "w") as archive:
            n = 0
            for dirname, dirnames, filenames in os.walk(os.path.join('src', rootname)):
                dirpath = os.path.relpath(dirname, os.path.join('src', rootname))
                if dirpath != '.':
                    archive.write(dirname, dirpath)
                for filename in filenames:
                    filepath = os.path.join(dirname, filename)
                    relpath = os.path.relpath(filepath, os.path.join('src', rootname))
                    archive.write(filepath, relpath)
    print ('''\n\nLeenO - Computo metrico assistito
Copyright (C) 2014-2019 Giuseppe Vizziello

Software Libero per computi metrici estimativi

Questa estensione si basa su UltimusFree di Bartolomeo Aimar
ed è distribuita con licenza LGPL

File di installazione generato correttamente: ''' + archpath +'\n')

if __name__=="__main__":
    compress()
