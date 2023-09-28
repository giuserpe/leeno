#!/usr/bin/env python3
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

#~ MsgBox('''Per segnalare questo problema,
#~ contatta il canale Telegram
#~ https://t.me/leeno_computometrico''', 'ERRORE!')

#~ documentazione ufficiale: https://api.libreoffice.org/
import locale
import codecs
import configparser
import collections
#~ import subprocess
#~ import psutil
import os, unohelper, pyuno, logging, shutil, base64, sys, uno
import time
import copy
from multiprocessing import Process, freeze_support
import threading
# cos'e' il namespace:
# http://www.html.it/articoli/il-misterioso-mondo-dei-namespaces-1/
from datetime import datetime, date
from com.sun.star.beans import PropertyValue
from xml.etree.ElementTree import ElementTree, Element, SubElement, Comment, tostring
#~ from com.sun.star.table.CellContentType import TEXT, EMPTY, VALUE, FORMULA
from com.sun.star.sheet.CellFlags import (VALUE, DATETIME, STRING,
                                          ANNOTATION, FORMULA, HARDATTR,
                                          OBJECTS, EDITATTR, FORMATTED)

import LeenoUtils
import pyleeno as PL
import SheetUtils
########################################################################


def creaGiornale():
    desktop = LeenoUtils.getDesktop()
    opz = PropertyValue()
    opz.Name = 'AsTemplate'
    opz.Value = True
    document = desktop.loadComponentFromURL(
        PL.LeenO_path() + '/template/leeno/Giornale_Lavori.ots', "_blank", 0,
        (opz, ))
    return document
    
    
########################################################################


def nuovo_giorno():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName('GIORNALE')
    # ~PL.GotoSheet('GIORNALE')
    #~ '''Imposta salti di pagina'''
    for i in range (5, SheetUtils.getLastUsedRow(oSheet)):
        if 'Data:' in oSheet.getCellByPosition(0, i).String:
            oSheet.getCellByPosition(0, i).Rows.IsStartOfNewPage = True
    #~ '''Inserisce unovo giorno'''
    oRangeAddress=oDoc.NamedRanges.giornale_bianco.ReferredCells.RangeAddress
    pin = SheetUtils.getLastUsedRow(oSheet) +1
    oCellAddress = oSheet.getCellByPosition(0, pin).getCellAddress()
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
    oSheet.copyRange(oCellAddress, oRangeAddress)
    PL._gotoCella(0, pin+1)
    oSheet.getCellByPosition(1, pin).Formula = '=NOW()'
    oSheet.getCellByPosition(0, pin).String = 'Data: ' + oSheet.getCellByPosition(1, pin).String
    oSheet.getCellByPosition(1, pin).String = ''
    PL._gotoCella(0, pin+2)
    #~ '''Raggruppa le righe per giorni.'''
    oSheet = oDoc.getSheets().getByName('GIORNALE')
    PL.GotoSheet('GIORNALE')
    test = SheetUtils.getLastUsedRow(oSheet)
    coppie = list()
    x = 2
    for i in range (3, test):
        if 'Data:' in oSheet.getCellByPosition(0, i).String:
            coppie.append([x, (i-1)])
            x = i +1
    # ~coppie.append(test)
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    oCellRangeAddr.StartColumn = 0
    oCellRangeAddr.EndColumn = 1
    oCellRangeAddr.StartRow = 0
    oCellRangeAddr.EndRow = test
    oSheet.ungroup(oCellRangeAddr, 1)
    oSheet.ungroup(oCellRangeAddr, 1)
    oSheet.getCellRangeByPosition(0, 1, 0, test).Rows.IsVisible = True
    # ~coppie.pop(-1)
    for el in  coppie:
        # ~oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.StartColumn = 0
        oCellRangeAddr.EndColumn = 1
        oCellRangeAddr.StartRow = el[0]
        oCellRangeAddr.EndRow = el[1]
        oSheet.group(oCellRangeAddr, 1)
        oSheet.getCellRangeByPosition(0, el[0], 1, el[1]).Rows.IsVisible = False
    elementi = list()
    for i in range (3, test):
        if oSheet.getCellByPosition(0, i).CellStyle == 'titolo':
            elementi.append([x, (i-1)])
            x = i +1
    # L'USO DEL FOR QUI DI SEGUITO RENDE IL FILE IRRECUPERABILE!!!
    # ~for el in elementi:
        # ~oCellRangeAddr.StartRow = el[0]
        # ~oCellRangeAddr.EndRow = el[1]
        # ~oSheet.group(oCellRangeAddr, 1)
    oDoc.CurrentController.setFirstVisibleRow(0)
    
    
def MENU_nuovo_giorno():
    '''
    Apre un nuovo giornale lavori o inserisce nuovo giorno
    '''
    oDoc = LeenoUtils.getDocument()
    if oDoc.getSheets().hasByName('GIORNALE_BIANCO'):
        PL.GotoSheet('GIORNALE')
    else:
        creaGiornale()
    nuovo_giorno()
    return
