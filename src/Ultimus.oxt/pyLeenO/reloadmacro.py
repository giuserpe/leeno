# -*- coding: utf-8 -*-
"""
file: reloadmacro.py
comment: this is a macro to reload and run other macros during development.
questo file deve stare in C:\Users\giuserpe\AppData\Roaming\LibreOffice\4\user\Scripts
o dir equivalente
"""
def Macro2ReloadAndRun(*args):
 ThisDoc = XSCRIPTCONTEXT.getDocument()
 import pyleeno
     try:#reload is builtin in python 2.x 
        reload(pyleeno)#reload source module
    except NameError:#reload is in the imp module in python 3.x
        from imp import reload
        reload(pyleeno)#reload source module
 from pyleeno import xmlsix2ods
 xmlsix2ods(ThisDoc)
g_exportedScripts = Macro2ReloadAndRun,
