#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse, os, zipfile

parser = argparse.ArgumentParser(description='Process some integers.')
parser.add_argument('-e', '--ext', default='oxt')
parser.add_argument('-V', '--version', default='')
args = parser.parse_args()

def compress():
    """
    Compress all directory under src in a zip archive under bin
    """
    for rootname in os.walk('src').next()[1]:
        archname = rootname
        fileName, fileExtension = os.path.splitext(archname)
        if args.version:
            fileName += '-%s' % args.version
        if args.ext:
            fileExtension = args.ext
        archname = '%s.%s' % (fileName, fileExtension)
        archpath = os.path.join('bin', archname)
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

if __name__=="__main__":
    compress()