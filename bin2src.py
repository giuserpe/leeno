#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os, zipfile

def extract():
    """
    Extract all archives in bin directory under src.
    """
    for dirname, dirnames, filenames in os.walk('bin'):            
        for filename in filenames:
            directory = filename
            path = os.path.join('src', directory)
            filepath = os.path.join(dirname, filename)
            if not os.path.exists(path):
                os.makedirs(path)
            with zipfile.ZipFile(filepath, "r") as f:
                f.extractall(path)

if __name__=="__main__":
    extract()