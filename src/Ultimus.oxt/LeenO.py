#!/usr/bin/env python
# -*- Mode: Python; coding: utf-8; indent-tabs-mode: nil; tab-width: 4 -*-
########################################################################
# LeenO - Computo Metrico
# Template assistito per la compilazione di Computi Metrici Estimativi
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
# Il codice contenuto in questo modulo è parte integrante dell'estensione LeenO
# Vi sarò grato se vorrete segnalarmi i malfunzionamenti (veri o presunti)
# Sono inoltre graditi suggerimenti in merito alle gestione della
# Contabilità Lavori e per l'ottimizzazione del codice.
########################################################################
#~ __author__="Giuseppe Vizziello"
import os, unohelper, pyuno, logging, shutil, base64, sys, uno

import traceback
#~ try:
    #~ import uno
    #~ import unohelper
    #~ import io
    #~ from com.sun.star.task import XJobExecutor
    #~ import sys
    #~ import os, unohelper, pyuno, logging, shutil, base64, sys, uno

    #~ import traceback
#~ except ImportError:
    #~ print("LeenO : import error")

def LeenO_path(arg=None):
    #~ '''Restituisce il percorso di installazione di LeenO.oxt'''
    ctx = uno.getComponentContext()
    pir = ctx.getValueByName('/singletons/com.sun.star.deployment.PackageInformationProvider')
    expath = pir.getPackageLocation('org.giuseppe-vizziello.leeno')
    return expath
########################################################################
# ~class pyLeenO( unohelper.Base, XJobExecutor ):
    # ~oxt_path = uno.fileUrlToSystemPath(LeenO_path())
    # ~sys.path.insert(0, oxt_path)
    # ~ chi(sys.path.insert(0, 'pyLeenO/'))
sys.path.insert(0, (LeenO_path()+'/pyLeenO'))
import pyleeno

    # ~def __init__( self, ctx ):
        # ~self.ctx = ctx
#~ print(LeenO_path()+'/pyLeenO')
def trigger( self, args ):
    #~ chi(arg)
    if arg=='DlgMain':
        # ~from pyLeenO.pyleeno import DlgMain
        #~ pyleeno.DlgMain()
        DlgMain()
# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(None,                       # UNO object class
                                        "org.giuseppe-vizziello.leeno", # implemenation name
                                        ("org.giuseppe-vizziello.leeno",),)     # list of implemented service
