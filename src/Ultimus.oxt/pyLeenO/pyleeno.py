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
########################################################################
# https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=27805&p=127383
import random
from com.sun.star.script.provider import XScriptProviderFactory
    
from com.sun.star.script.provider import XScriptProvider
def barra_di_stato(testo='', valore=0):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oProgressBar = oDoc.CurrentController.Frame.createStatusIndicator()
    oProgressBar.start('', 100)
    oProgressBar.Value = valore
    oProgressBar.Text = testo
def Xray(myObject):
    # Taken from http://www.oooforum.org/forum/viewtopic.phtml?t=23577
    xCompCont = XSCRIPTCONTEXT.getComponentContext()
    sm = xCompCont.ServiceManager
    mspf = sm.createInstance("com.sun.star.script.provider.MasterScriptProviderFactory")
    scriptPro = mspf.createScriptProvider("")
    Xscript = scriptPro.getScript("vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application")
    Xscript.invoke((myObject,), None, None)

def basic_LeenO(funcname, *args):
    xCompCont = XSCRIPTCONTEXT.getComponentContext()
    sm = xCompCont.ServiceManager
    mspf = sm.createInstance("com.sun.star.script.provider.MasterScriptProviderFactory")
    scriptPro = mspf.createScriptProvider("")
    Xscript = scriptPro.getScript("vnd.sun.star.script:UltimusFree2." + funcname + "?language=Basic&location=application")
    Result = Xscript.invoke(args, None, None)
    return Result[0]
########################################################################
def LeenO_path(arg=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    pir = ctx.getValueByName('/singletons/com.sun.star.deployment.PackageInformationProvider')
    expath = pir.getPackageLocation('org.giuseppe-vizziello.leeno')
    return expath
########################################################################
class New_file:
    '''Crea un nuovo computo.'''
    def __init__(self):#, computo):
        pass
    def computo(arg=1):
        '''arg  { integer } : 1 mostra il dialogo di salvataggio file'''
        desktop = XSCRIPTCONTEXT.getDesktop()
        opz = PropertyValue()
        opz.Name = 'AsTemplate'
        opz.Value = True
        document = desktop.loadComponentFromURL(LeenO_path()+'/template/leeno/Computo_LeenO.ots', "_blank", 0, (opz,))
        if arg == 1:
            MsgBox('''Prima di procedere è consigliabile salvare il lavoro.
Provvedi subito a dare un nome al file di computo...''', 'Dai un nome al file...')
            salva_come()
            autoexec()
        #~ salva_come()
        return document
    def usobollo():
        desktop = XSCRIPTCONTEXT.getDesktop()
        opz = PropertyValue()
        opz.Name = 'AsTemplate'
        opz.Value = True
        document = desktop.loadComponentFromURL(LeenO_path()+'/template/offmisc/UsoBollo.ott', "_blank", 0, (opz,))
        return document
########################################################################
def nuovo_computo(arg=None):
    New_file.computo()
########################################################################
def nuovo_usobollo(arg=None):
    New_file.usobollo()
########################################################################
def invia_voce_ep(arg=None):
    '''
    Invia le voci di prezzario selezionate da un elenco prezzi all'Elenco Prezzi del
    Documento di Contabilità Corrente DCC. Trasferisce anche le Analisi di Prezzo.
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    #~ oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
    SR = oRangeAddress.StartRow
    ER = oRangeAddress.EndRow
    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, SR, getLastUsedCell(oSheet).EndColumn, ER))
    lista = list()
    for el in range(SR, ER+1):
        if oSheet.getCellByPosition(1, el).Type.value == 'FORMULA':
            lista.append(oSheet.getCellByPosition(0, el).String)
    try:
        fpartenza = uno.fileUrlToSystemPath(oDoc.getURL())
    except:
        MsgBox("E' necessario prima salvare il file di fpartenza.", "Attenzione!")
        salva_come()
        fpartenza = uno.fileUrlToSystemPath(oDoc.getURL())
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, ".uno:Copy", "", 0, list())
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
    if sUltimus == '':
        MsgBox("E' necessario impostare il Documento di contabilità Corrente.", "Attenzione!")
        return
    _gotoDoc(sUltimus)
    ddcDoc = XSCRIPTCONTEXT.getDocument()
    dccSheet = ddcDoc.getSheets().getByName('Elenco Prezzi')
    dccSheet.IsVisible = True
    ddcDoc.CurrentController.setActiveSheet(dccSheet)

    dccSheet.getRows().insertByIndex(3, ER-SR+1)
    
    ddcDoc.CurrentController.select(dccSheet.getCellByPosition(0, 3))
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx )
    dispatchHelper.executeDispatch(oFrame, ".uno:Paste", "", 0, list())
    ddcDoc.CurrentController.select(ddcDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
    #~ doppioni()
    _gotoDoc(fpartenza)
    oDoc = XSCRIPTCONTEXT.getDocument()

    if len(lista) > 0:
        if oDoc.getSheets().hasByName('tmp_DCC') == False:
            sheet = oDoc.createInstance("com.sun.star.sheet.Spreadsheet")
            tmp = oDoc.Sheets.insertByName('tmp_DCC', sheet)
        tmp = oDoc.getSheets().getByName('tmp_DCC')
        oSheet = oDoc.getSheets().getByName('Analisi di Prezzo')
        oDoc.CurrentController.setActiveSheet(oSheet)
        for el in lista:
            celle = Circoscrive_Analisi(uFindStringCol(el, 0, oSheet))
            oRangeAddress = celle.getRangeAddress()
            oCellAddress = tmp.getCellByPosition(0, getLastUsedCell(tmp).EndRow).getCellAddress()
            tmp.copyRange(oCellAddress, oRangeAddress)
        
        nuove_righe = getLastUsedCell(tmp).EndRow+1
        analisi = tmp.getCellRangeByPosition(0, 0, getLastUsedCell(tmp).EndColumn, getLastUsedCell(tmp).EndRow)
        
        oDoc.CurrentController.select(analisi)
        ctx = XSCRIPTCONTEXT.getComponentContext()
        desktop = XSCRIPTCONTEXT.getDesktop()
        oFrame = desktop.getCurrentFrame()
        dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
        dispatchHelper.executeDispatch(oFrame, ".uno:Copy", "", 0, list())
        oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
        
        _gotoDoc(sUltimus)
        ddcDoc = XSCRIPTCONTEXT.getDocument()
        if ddcDoc.getSheets().hasByName('Analisi di Prezzo') == False:
            inizializza_analisi()
        dccSheet = ddcDoc.getSheets().getByName('Analisi di Prezzo')
        lrow = getLastUsedCell(dccSheet).EndRow
        
        dccSheet.getRows().insertByIndex(lrow, nuove_righe)
        ddcDoc.CurrentController.select(dccSheet.getCellByPosition(0, lrow))
    
        ctx = XSCRIPTCONTEXT.getComponentContext()
        desktop = XSCRIPTCONTEXT.getDesktop()
        oFrame = desktop.getCurrentFrame()
        dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
        dispatchHelper.executeDispatch(oFrame, ".uno:Paste", "", 0, list())
        ddcDoc.CurrentController.select(ddcDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
        doppioni()
        ddcDoc.CurrentController.setActiveSheet(ddcDoc.getSheets().getByName('Elenco Prezzi'))

        _gotoDoc(fpartenza)
        oDoc.Sheets.removeByName('tmp_DCC')
    _gotoDoc(sUltimus)
    oDoc.CurrentController.setActiveSheet(oDoc.getSheets().getByName('Elenco Prezzi'))
    dccSheet = ddcDoc.getSheets().getByName('Elenco Prezzi')
    dccSheet.IsVisible = True
    ddcDoc.CurrentController.setActiveSheet(dccSheet)
    formule = list()
    for n in range(3, ER-SR+1+3):
        formule.append(['=IF(ISERROR(N'+str(n+1)+'/$N$2);"--";N'+str(n+1)+'/$N$2)',
                        '=SUMIF(AA;A'+str(n+1)+';BB)',
                        '=SUMIF(AA;A'+str(n+1)+';cEuro)'])
    oRange = dccSheet.getCellRangeByPosition(11, 3, 13, ER-SR+3)
    formule = tuple(formule)
    oRange.setFormulaArray(formule)
    if conf.read(path_conf, 'Generale', 'torna_a_ep') == '1':
        _gotoDoc(fpartenza)
########################################################################
def _gotoDoc(sUrl):
    '''
    sUrl  { string } : nome del file
    porta il focus su di un determinato documento
    '''
    sUrl = uno.systemPathToFileUrl(sUrl)
    #~ target = XSCRIPTCONTEXT.getDesktop().loadComponentFromURL(sUrl, "_default", 0, list())
    #~ target.getCurrentController().Frame.ContainerWindow.toFront()
    #~ target.getCurrentController().Frame.activate()
    if sys.platform == 'linux' or sys.platform == 'darwin':
        oDialogo_attesa = dlg_attesa()
        attesa().start() #mostra il dialogo
        target = XSCRIPTCONTEXT.getDesktop().loadComponentFromURL(sUrl, "_default", 0, list())
        target.getCurrentController().Frame.ContainerWindow.toFront()
        target.getCurrentController().Frame.activate()
        oDialogo_attesa.endExecute()
    elif sys.platform == 'win32':
        #~ target = XSCRIPTCONTEXT.getDesktop().loadComponentFromURL(sUrl, "_default", 0, list())
        #~ target.getCurrentController().Frame.ContainerWindow.toFront()
        #~ target.getCurrentController().Frame.activate()
        
        desktop = XSCRIPTCONTEXT.getDesktop()
        oFocus = uno.createUnoStruct('com.sun.star.awt.FocusEvent')
        target = desktop.loadComponentFromURL(sUrl, "_default", 0, list())
        target.getCurrentController().getFrame().focusGained(oFocus)
########################################################################
def oggi():
    '''
    restituisce la data di oggi
    '''
    return('/'.join(reversed(str(datetime.now()).split(' ')[0].split('-'))))
import distutils.dir_util
########################################################################
def copia_sorgente_per_git(arg=None):
    '''
    fa una copia della directory del codice nel repository locale ed apre una shell per la commit
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    try:
        if oDoc.getSheets().getByName('S1').getCellByPosition(7,338).String == '':
            src_oxt ='_LeenO'
        else:
            src_oxt = oDoc.getSheets().getByName('S1').getCellByPosition(7,338).String
    except:
        pass
    make_pack(bar=1)
    oxt_path = uno.fileUrlToSystemPath(LeenO_path())
    if sys.platform == 'linux' or sys.platform == 'darwin':
        pass
        dest = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt'
        
        #~ os.system('nemo /media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt')
        os.system('cd /media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt && gnome-terminal && gitk &')

    elif sys.platform == 'win32':
        if not os.path.exists('w:/_dwg/ULTIMUSFREE/_SRC/leeno/src/'):
            try:
                os.makedirs(os.getenv("HOMEPATH") +'\\'+ src_oxt +'\\leeno\\src\\Ultimus.oxt\\')
            except FileExistsError:
                pass
            dest = os.getenv("HOMEDRIVE") + os.getenv("HOMEPATH") +'\\'+ src_oxt +'\\leeno\\src\\Ultimus.oxt\\'
        else:
            dest = 'w:/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt'
        
            #~ os.system('explorer.exe w:\\_dwg\\ULTIMUSFREE\\_SRC\\leeno\\src\\Ultimus.oxt\\')
            os.system('w: && cd w:/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt && "C:/Program Files/Git/git-bash.exe" && "C:/Program Files/Git/cmd/gitk.exe"')
    distutils.dir_util.copy_tree(oxt_path, dest)
    return
########################################################################
def debugs(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    desktop = XSCRIPTCONTEXT.getDesktop()
    ctx = XSCRIPTCONTEXT.getComponentContext()
    oSheet = oDoc.CurrentController.ActiveSheet
    

    oSheet.getCellByPosition(1,7).String = oDoc.getURL()
    oSheet.getCellByPosition(1,8).String = dir(os).__str__()
    oSheet.getCellByPosition(1, 9).String = dir(uno.__package__.title.__name__).__str__()
    oSheet.getCellByPosition(1, 10).String = dir(unohelper).__str__()
    #~ oSheet.getCellRangeByName('A12').String = sys.__doc__
    #~ oSheet.getCellRangeByName('A13').String = dir(uno).__str__()
    #~ oSheet.getCellRangeByName('A14').String = dir(unohelper).__str__()
    #~ oSheet.getCellRangeByName('A15').String = dir(pyuno).__str__()
    #~ oSheet.getCellRangeByName('A11').String = dir(pyuno).__str__()
    #~ n = 1
    #~ for el in dir(uno):
        #~ oSheet.getCellRangeByName('A' + str(n)).String = 'uno.' +
    #~ oSheet.getCellRangeByName('A12').String = sys.__doc__

########################################################################
def Inser_SottoCapitolo(arg=None):
    Ins_Categorie(2)

def Inser_SottoCapitolo_arg(lrow, sTesto): #
    '''
    lrow    { double } : id della riga di inerimento
    sTesto  { string } : titolo della sottocategoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in('COMPUTO', 'VARIANTE'):
        return

    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Default': lrow -= 2#se oltre la riga rossa
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Riga_rossa_Chiudi': lrow -= 1#se riga rossa
    insRows(lrow, 1)
    oSheet.getCellByPosition(2, lrow).String = sTesto
# inserisco i valori e le formule
    oSheet.getCellRangeByPosition(0, lrow, 41, lrow).CellStyle = 'livello2 valuta'
    oSheet.getCellRangeByPosition(2, lrow, 17, lrow).CellStyle = 'livello2_'
    oSheet.getCellRangeByPosition(18, lrow, 18, lrow).CellStyle = 'livello2 scritta mini'
    oSheet.getCellRangeByPosition(24, lrow, 24, lrow).CellStyle = 'livello2 valuta mini %'
    oSheet.getCellRangeByPosition(29, lrow, 29, lrow).CellStyle = 'livello2 valuta mini %'
    oSheet.getCellRangeByPosition(30, lrow, 30, lrow).CellStyle = 'livello2 valuta mini'
    oSheet.getCellRangeByPosition(31, lrow, 33, lrow).CellStyle = 'livello2_'
    oSheet.getCellRangeByPosition(2, lrow, 11, lrow).merge(True)
    #~ oSheet.getCellByPosition(1, lrow).Formula = '=AF' + str(lrow+1) + '''&"."&''' + 'AG' + str(lrow+1)
    # rinumero e ricalcolo
    ocellBaseA = oSheet.getCellByPosition(1, lrow)
    ocellBaseR = oSheet.getCellByPosition(31, lrow)

    lrowProvv = lrow-1
    while oSheet.getCellByPosition(32, lrowProvv).CellStyle != 'livello2 valuta':
        if lrowProvv > 4:
            lrowProvv -=1
        else:
            break
    oSheet.getCellByPosition(32, lrow).Value = oSheet.getCellByPosition(1 , lrowProvv).Value + 1
    lrowProvv = lrow-1
    while oSheet.getCellByPosition(31, lrowProvv).CellStyle != 'Livello-1-scritta':
        if lrowProvv > 4:
            lrowProvv -=1
        else:
            break
    oSheet.getCellByPosition(31, lrow).Value = oSheet.getCellByPosition(1 , lrowProvv).Value
    #~ SubSum_Cap(lrow)

########################################################################
def Ins_Categorie(n):
    '''
    n    { int } : livello della categoria
    0 = SuperCategoria
    1 = Categoria
    2 = SubCategoria
    '''
    #~ datarif = datetime.now()

    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    row = Range2Cell()[1]
    if oSheet.getCellByPosition(0, row).CellStyle in stili_computo:
        lrow = next_voice(row, 1)
    elif oSheet.getCellByPosition(0, row).CellStyle in noVoce:
        lrow = row+1
    else:
        return
    sTesto = ''
    if n==0:
        sTesto = 'Inserisci il titolo per la Supercategoria'
    elif n==1:
        sTesto = 'Inserisci il titolo per la Categoria'
    elif n==2:
        sTesto = 'Inserisci il titolo per la Sottocategoria'
    sString = InputBox('', sTesto)
    if sString == None or sString == '':
        return
    oDoc.CurrentController.ZoomValue = 400
    if n==0:
        Inser_SuperCapitolo_arg(lrow, sString)
    elif n==1:
        Inser_Capitolo_arg(lrow, sString)
    elif n==2:
        Inser_SottoCapitolo_arg(lrow, sString)

    _gotoCella(2, lrow)
    Rinumera_TUTTI_Capitoli2()
    oDoc.CurrentController.ZoomValue = 100
    oDoc.CurrentController.setFirstVisibleColumn(0)
    oDoc.CurrentController.setFirstVisibleRow(lrow-5)
    #~ MsgBox('eseguita in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!','')
    
########################################################################
def Inser_SuperCapitolo(arg=None):
    Ins_Categorie(0)

def Inser_SuperCapitolo_arg(lrow, sTesto='Super Categoria'): #
    '''
    lrow    { double } : id della riga di inerimento
    sTesto  { string } : titolo della categoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in('COMPUTO', 'VARIANTE'):
        return
    #~ lrow = Range2Cell()[1]
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Default': lrow -= 2#se oltre la riga rossa
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Riga_rossa_Chiudi': lrow -= 1#se riga rossa
    insRows(lrow, 1)
    oSheet.getCellByPosition(2, lrow).String = sTesto
    # inserisco i valori e le formule
    oSheet.getCellRangeByPosition(0, lrow, 41, lrow).CellStyle = 'Livello-0-scritta'
    oSheet.getCellRangeByPosition(2, lrow, 17, lrow).CellStyle = 'Livello-0-scritta mini'
    oSheet.getCellRangeByPosition(18, lrow, 18, lrow).CellStyle = 'Livello-0-scritta mini val'
    oSheet.getCellRangeByPosition(24, lrow, 24, lrow).CellStyle = 'Livello-0-scritta mini %'
    oSheet.getCellRangeByPosition(29, lrow, 29, lrow).CellStyle = 'Livello-0-scritta mini %'
    oSheet.getCellRangeByPosition(30, lrow, 30, lrow).CellStyle = 'Livello-0-scritta mini val'
    oSheet.getCellRangeByPosition(2, lrow, 11, lrow).merge(True)
    # rinumero e ricalcolo
    ocellBaseA = oSheet.getCellByPosition(1, lrow)
    ocellBaseR = oSheet.getCellByPosition(31, lrow)
    lrowProvv = lrow-1
    while oSheet.getCellByPosition(31, lrowProvv).CellStyle != 'Livello-0-scritta':
        if lrowProvv > 4:
            lrowProvv -=1
        else:
            break
    oSheet.getCellByPosition(31, lrow).Value = oSheet.getCellByPosition(1 , lrowProvv).Value + 1
########################################################################
def Inser_Capitolo(arg=None):
    Ins_Categorie(1)

def Inser_Capitolo_arg(lrow, sTesto='Categoria'): #
    '''
    lrow    { double } : id della riga di inerimento
    sTesto  { string } : titolo della categoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in('COMPUTO', 'VARIANTE'):
        return
    #~ lrow = Range2Cell()[1]
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Default': lrow -= 2#se oltre la riga rossa
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Riga_rossa_Chiudi': lrow -= 1#se riga rossa
    insRows(lrow, 1)
    oSheet.getCellByPosition(2, lrow).String = sTesto
    # inserisco i valori e le formule
    oSheet.getCellRangeByPosition(0, lrow, 41, lrow).CellStyle = 'Livello-1-scritta'
    oSheet.getCellRangeByPosition(2, lrow, 17, lrow).CellStyle = 'Livello-1-scritta mini'
    oSheet.getCellRangeByPosition(18, lrow, 18, lrow).CellStyle = 'Livello-1-scritta mini val'
    oSheet.getCellRangeByPosition(24, lrow, 24, lrow).CellStyle = 'Livello-1-scritta mini %'
    oSheet.getCellRangeByPosition(29, lrow, 29, lrow).CellStyle = 'Livello-1-scritta mini %'
    oSheet.getCellRangeByPosition(30, lrow, 30, lrow).CellStyle = 'Livello-1-scritta mini val'
    oSheet.getCellRangeByPosition(2, lrow, 11, lrow).merge(True)
    # rinumero e ricalcolo
    ocellBaseA = oSheet.getCellByPosition(1, lrow)
    ocellBaseR = oSheet.getCellByPosition(31, lrow)
    lrowProvv = lrow-1
    while oSheet.getCellByPosition(31, lrowProvv).CellStyle != 'Livello-1-scritta':
        if lrowProvv > 4:
            lrowProvv -=1
        else:
            break
    oSheet.getCellByPosition(31, lrow).Value = oSheet.getCellByPosition(1 , lrowProvv).Value + 1
########################################################################
def Rinumera_TUTTI_Capitoli2(arg=None):
    Sincronizza_SottoCap_Tag_Capitolo_Cor()# sistemo gli idcat voce per voce
    Tutti_Subtotali()# ricalcola i totali di categorie e subcategorie

def Tutti_Subtotali(arg=None):
    '''ricalcola i subtotali di categorie e subcategorie'''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in('COMPUTO', 'VARIANTE'):
        return
    for n in range(0, ultima_voce(oSheet)+1):
        if oSheet.getCellByPosition(0, n).CellStyle == 'Livello-0-scritta':
            SubSum_SuperCap(n)
        if oSheet.getCellByPosition(0, n).CellStyle == 'Livello-1-scritta':
            SubSum_Cap(n)
        if oSheet.getCellByPosition(0, n).CellStyle == 'livello2 valuta':
            SubSum_SottoCap(n)
# TOTALI GENERALI
    lrow = ultima_voce(oSheet)+1
    for x in (1, lrow):
        oSheet.getCellByPosition(17, x).Formula = '=SUBTOTAL(9;R4:R' + str(lrow+1) + ')'
        oSheet.getCellByPosition(18, x).Formula = '=SUBTOTAL(9;S4:S' + str(lrow+1) + ')'
        oSheet.getCellByPosition(30, x).Formula = '=SUBTOTAL(9;AE4:AE' + str(lrow+1) + ')'
        oSheet.getCellByPosition(36, x).Formula = '=SUBTOTAL(9;AK4:AK' + str(lrow+1) + ')'
########################################################################
def SubSum_SuperCap(lrow):
    '''
    lrow    { double } : id della riga di inerimento
    inserisce i dati nella riga di SuperCategoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in('COMPUTO', 'VARIANTE'):
        return
    #~ lrow = Range2Cell()[1]
    lrowE = ultima_voce(oSheet)+2
    nextCap = lrowE
    for n in range(lrow+1, lrowE):
        if oSheet.getCellByPosition(18, n).CellStyle in('Livello-0-scritta mini val', 'Comp TOTALI'):
            #~ MsgBox(oSheet.getCellByPosition(18, n).CellStyle,'')
            nextCap = n + 1
            break
    #~ oDoc.enableAutomaticCalculation(False)
    oSheet.getCellByPosition(18, lrow).Formula = '=SUBTOTAL(9;S' + str(lrow + 1) + ':S' + str(nextCap) + ')'
    oSheet.getCellByPosition(18, lrow).CellStyle = 'Livello-0-scritta mini val'
    oSheet.getCellByPosition(24, lrow).Formula = '=S' + str(lrow + 1) + '/S' + str(lrowE)
    oSheet.getCellByPosition(24, lrow).CellStyle = 'Livello-0-scritta mini %'
    oSheet.getCellByPosition(28, lrow).Formula = '=SUBTOTAL(9;AC' + str(lrow + 1) + ':AC' + str(nextCap) + ')'
    oSheet.getCellByPosition(29, lrow).Formula = '=AE' + str(lrow + 1) + '/S' + str(lrowE)
    oSheet.getCellByPosition(29, lrow).CellStyle = 'Livello-0-scritta mini %'
    oSheet.getCellByPosition(30, lrow).Formula = '=SUBTOTAL(9;AE' + str(lrow + 1) + ':AE' + str(nextCap) + ')'
    oSheet.getCellByPosition(30, lrow).CellStyle = 'Livello-0-scritta mini val'
    #~ oDoc.enableAutomaticCalculation(True)
########################################################################
def SubSum_SottoCap(lrow):
    '''
    lrow    { double } : id della riga di inerimento
    inserisce i dati nella riga di subcategoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in('COMPUTO', 'VARIANTE'):
        return
    #lrow = 0#Range2Cell()[1]
    lrowE = ultima_voce(oSheet)+2
    nextCap = lrowE
    for n in range(lrow+1, lrowE):
        if oSheet.getCellByPosition(18, n).CellStyle in('livello2 scritta mini', 'Livello-0-scritta mini val', 'Livello-1-scritta mini val', 'Comp TOTALI'):
            nextCap = n + 1
            break
    oSheet.getCellByPosition(18, lrow).Formula = '=SUBTOTAL(9;S' + str(lrow + 1) + ':S' + str(nextCap) + ')'
    oSheet.getCellByPosition(18, lrow).CellStyle = 'livello2 scritta mini'
    oSheet.getCellByPosition(24, lrow).Formula = '=S' + str(lrow + 1) + '/S' + str(lrowE)
    oSheet.getCellByPosition(24, lrow).CellStyle = 'livello2 valuta mini %'
    oSheet.getCellByPosition(28, lrow).Formula = '=SUBTOTAL(9;AC' + str(lrow + 1) + ':AC' + str(nextCap) + ')'
    oSheet.getCellByPosition(28, lrow).CellStyle = 'livello2 scritta mini'
    oSheet.getCellByPosition(29, lrow).Formula = '=AE' + str(lrow + 1) + '/S' + str(lrowE)
    oSheet.getCellByPosition(29, lrow).CellStyle = 'livello2 valuta mini %'
    oSheet.getCellByPosition(30, lrow).Formula = '=SUBTOTAL(9;AE' + str(lrow + 1) + ':AE' + str(nextCap) + ')'
    oSheet.getCellByPosition(30, lrow).CellStyle = 'livello2 valuta mini'
########################################################################
def SubSum_Cap(lrow):
    '''
    lrow    { double } : id della riga di inerimento
    inserisce i dati nella riga di categoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in('COMPUTO', 'VARIANTE'):
        return
    #~ lrow = Range2Cell()[1]
    lrowE = ultima_voce(oSheet)+2
    nextCap = lrowE
    for n in range(lrow+1, lrowE):
        if oSheet.getCellByPosition(18, n).CellStyle in('Livello-1-scritta mini val','Livello-0-scritta mini val',  'Comp TOTALI'):
            #~ MsgBox(oSheet.getCellByPosition(18, n).CellStyle,'')
            nextCap = n + 1
            break
    #~ oDoc.enableAutomaticCalculation(False)
    oSheet.getCellByPosition(18, lrow).Formula = '=SUBTOTAL(9;S' + str(lrow + 1) + ':S' + str(nextCap) + ')'
    oSheet.getCellByPosition(18, lrow).CellStyle = 'Livello-1-scritta mini val'
    oSheet.getCellByPosition(24, lrow).Formula = '=S' + str(lrow + 1) + '/S' + str(lrowE)
    oSheet.getCellByPosition(24, lrow).CellStyle = 'Livello-1-scritta mini %'
    oSheet.getCellByPosition(28, lrow).Formula = '=SUBTOTAL(9;AC' + str(lrow + 1) + ':AC' + str(nextCap) + ')'
    oSheet.getCellByPosition(29, lrow).Formula = '=AE' + str(lrow + 1) + '/S' + str(lrowE)
    oSheet.getCellByPosition(29, lrow).CellStyle = 'Livello-1-scritta mini %'
    oSheet.getCellByPosition(30, lrow).Formula = '=SUBTOTAL(9;AE' + str(lrow + 1) + ':AE' + str(nextCap) + ')'
    oSheet.getCellByPosition(30, lrow).CellStyle = 'Livello-1-scritta mini val'
    #~ oDoc.enableAutomaticCalculation(True)
########################################################################
def debug_delay(n=None):
    '''
    sCella  { string } : stringa di default nella casella di testo
    t       { string } : titolo del dialogo
    Viasualizza un dialogo di richiesta testo
    '''

    psm = uno.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDialog1 = dp.createDialog("vnd.sun.star.script:UltimusFree2.DlgAttesa?language=Basic&location=application")
    oDialog1Model = oDialog1.Model

    oDialog1Model.Title = 'tiolo'
    
    if n==1:
        oDialog1.execute()
        #~ chi(oDialog1)
    elif n==0:
        oDialog1.endDialog()
        oDialog1.endExecute()
        

########################################################################
def Sincronizza_SottoCap_Tag_Capitolo_Cor(arg=None):
    '''
    lrow    { double } : id della riga di inerimento
    sincronizza il categoria e sottocategorie
    '''
    datarif = datetime.now()
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in('COMPUTO', 'VARIANTE'):
        return
#    lrow = Range2Cell()[1]
    lastRow = ultima_voce(oSheet)+1

    listasbcat = list()
    listacat = list()
    listaspcat = list()
    for lrow in range(0,lastRow): # 
        if oSheet.getCellByPosition(2, lrow).CellStyle == 'livello2_': #SUB CATEGORIA
            if oSheet.getCellByPosition(2, lrow).String not in listasbcat:
                listasbcat.append((oSheet.getCellByPosition(2, lrow).String))
            try:
                oSheet.getCellByPosition(31, lrow).Value = idspcat
            except:
                pass
            try:
                oSheet.getCellByPosition(32, lrow).Value = idcat
            except:
                pass
            idsbcat = listasbcat.index(oSheet.getCellByPosition(2, lrow).String) +1
            oSheet.getCellByPosition(33, lrow).Value = idsbcat
            oSheet.getCellByPosition(1, lrow).Formula = '=AF' + str(lrow+1) +'&"."&AG' + str(lrow+1) + '&"."&AH' + str(lrow+1)
            
        elif oSheet.getCellByPosition(2, lrow).CellStyle == 'Livello-1-scritta mini': #CATEGORIA
            if oSheet.getCellByPosition(2, lrow).String not in listacat:
                listacat.append((oSheet.getCellByPosition(2, lrow).String))
                
                idsbcat = None
                
            try:
                oSheet.getCellByPosition(31, lrow).Value = idspcat
            except:
                pass
            idcat = listacat.index(oSheet.getCellByPosition(2, lrow).String) +1
            oSheet.getCellByPosition(32, lrow).Value = idcat
            oSheet.getCellByPosition(1, lrow).Formula = '=AF' + str(lrow+1) +'&"."&AG' + str(lrow+1)

        elif oSheet.getCellByPosition(2, lrow).CellStyle == 'Livello-0-scritta mini': #SUPER CATEGORIA
            if oSheet.getCellByPosition(2, lrow).String not in listaspcat:
                listaspcat.append((oSheet.getCellByPosition(2, lrow).String))
                
                idcat = idsbcat = None
                
            idspcat = listaspcat.index(oSheet.getCellByPosition(2, lrow).String) +1
            oSheet.getCellByPosition(31, lrow).Value = idspcat
            oSheet.getCellByPosition(1, lrow).Formula = '=AF' + str(lrow+1)
            
        elif oSheet.getCellByPosition(33, lrow).CellStyle == 'compTagRiservato': #CATEGORIA
            try:
                oSheet.getCellByPosition(33, lrow).Value = idsbcat
            except:
                oSheet.getCellByPosition(33, lrow).Value = 0
            try:
                oSheet.getCellByPosition(32, lrow).Value = idcat
            except:
                oSheet.getCellByPosition(32, lrow).Value = 0
            try:
                oSheet.getCellByPosition(31, lrow).Value = idspcat
            except:
                oSheet.getCellByPosition(31, lrow).Value = 0

    #~ MsgBox('Importazione eseguita con successo\n in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!','')
    
########################################################################
def insRows(lrow, nrighe): #forse inutile
    '''
    lrow    { double }  : id della riga di inerimento
    lrow    { integer } : numero di nuove righe da inserire

    Inserisce nrighe nella posizione lrow - alternativo a
    oSheet.getRows().insertByIndex(lrow, 1)
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    iSheet = oSheet.RangeAddress.Sheet
    #~ oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    #~ lrow = Range2Cell()[1]
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    oCellRangeAddr.StartColumn = 0
    oCellRangeAddr.EndColumn = 0
    oCellRangeAddr.StartRow = lrow
    oCellRangeAddr.EndRow = lrow + nrighe - 1
    oSheet.insertCells(oCellRangeAddr, 3)   # com.sun.star.sheet.CellInsertMode.ROW
########################################################################
def ultima_voce(oSheet):
    #~ oDoc = XSCRIPTCONTEXT.getDocument()
    #~ oSheet = oDoc.CurrentController.ActiveSheet
    nRow = getLastUsedCell(oSheet).EndRow
    for n in reversed(range(0, nRow)):
        if oSheet.getCellByPosition(0, n).CellStyle in('EP-aS', 'EP-Cs', 'An-sfondo-basso Att End', 'Comp End Attributo',
                                                        'Comp End Attributo_R', 'comp Int_colonna', 'comp Int_colonna_R_prima',
                                                        'Livello-0-scritta', 'Livello-1-scritta', 'livello2 valuta'):
            break
    return n
########################################################################
def uFindStringCol(sString, nCol, oSheet):
    '''
    sString { string }  : stringa da cercare
    nCol    { integer } : indice di colonna
    oSheet  { object }  :

    Trova la prima ricorrenza di una stringa(sString) nella
    colonna nCol di un foglio di calcolo(oSheet) e restituisce
    in numero di riga
    '''
    oCell = oSheet.getCellByPosition(0,0)
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.gotoEndOfUsedArea(True)
    aAddress = oCursor.RangeAddress
    for nRow in range(0, aAddress.EndRow+1):
        if sString in oSheet.getCellByPosition(nCol,nRow).String:
            return(nRow)
########################################################################
def uFindString(sString, oSheet):
    '''
    sString { string }  : stringa da cercare
    oSheet  { object }  :

    Trova la prima ricorrenza di una stringa(sString) riga
    per riga in un foglio di calcolo(oSheet) e restituisce
    una tupla(IDcolonna, IDriga)
    '''
    oCell = oSheet.getCellByPosition(0,0)
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.gotoEndOfUsedArea(True)
    aAddress = oCursor.RangeAddress
    for nRow in range(0, aAddress.EndRow+1):
        for nCol in range(0, aAddress.EndColumn+1):
    # ritocco di +Daniele Zambelli:
            if sString in oSheet.getCellByPosition(nCol,nRow).String:
                return(nCol,nRow)
########################################################################
def join_sheets(arg=None):
    '''
    unisci fogli
    serve per unire tanti fogli in un unico foglio
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    lista_fogli = oDoc.Sheets.ElementNames
    if oDoc.getSheets().hasByName('unione_fogli') == False:
        sheet = oDoc.createInstance("com.sun.star.sheet.Spreadsheet")
        unione = oDoc.Sheets.insertByName('unione_fogli', sheet)
        unione = oDoc.getSheets().getByName('unione_fogli')
        for el in lista_fogli:
            oSheet = oDoc.getSheets().getByName(el)
            oRangeAddress = oSheet.getCellRangeByPosition(0,0,(getLastUsedCell(oSheet).EndColumn),(getLastUsedCell(oSheet).EndRow)).getRangeAddress()
            oCellAddress = unione.getCellByPosition(0, getLastUsedCell(unione).EndRow+1).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
        MsgBox('Unione dei fogli eseguita.','Avviso')
    else:
        unione = oDoc.getSheets().getByName('unione_fogli')
        MsgBox('Il foglio "unione_fogli" è già esistente, quindi non procedo.','Avviso!')
    oDoc.CurrentController.setActiveSheet(unione)
########################################################################
def copia_sheet(nSheet, tag):
    '''
    nSheet   { string } : nome sheet
    tag      { string } : stringa di tag
    duplica copia sheet corrente di fianco a destra
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    #~ nSheet = 'COMPUTO'
    oSheet = oDoc.getSheets().getByName(nSheet)
    idSheet = oSheet.RangeAddress.Sheet + 1
    if oDoc.getSheets().hasByName(nSheet +'_'+ tag) == True:
        MsgBox('La tabella di nome '+ nSheet +'_'+ tag + 'è già presente.', 'ATTENZIONE! Impossibile procedere.')
        return
    else:
        oDoc.Sheets.copyByName(nSheet, nSheet +'_'+ tag, idSheet)
        oSheet = oDoc.getSheets().getByName(nSheet +'_'+ tag)
        oDoc.CurrentController.setActiveSheet(oSheet)
        #~ oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
########################################################################
def debugpuliscixxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx():
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    for lrow in reversed(range(0, ultima_voce(oSheet))):
        if oSheet.getCellByPosition(31,lrow).CellStyle == 'compTagG' :
            oSheet.getCellByPosition(31,lrow).String = ''
            oSheet.getCellByPosition(32,lrow).String = ''
            oSheet.getCellByPosition(33,lrow).String = ''
            oSheet.getCellByPosition(34,lrow).String = ''
            oSheet.getCellByPosition(35,lrow).String = ''
    _gotoSheet('S5')
    oSheet = oDoc.CurrentController.ActiveSheet
    for lrow in reversed(range(0, ultima_voce(oSheet))):
        if oSheet.getCellByPosition(31,lrow).CellStyle == 'compTagG' :
            oSheet.getCellByPosition(31,lrow).String = ''
            oSheet.getCellByPosition(32,lrow).String = ''
            oSheet.getCellByPosition(33,lrow).String = ''
            oSheet.getCellByPosition(34,lrow).String = ''
            oSheet.getCellByPosition(35,lrow).String = ''
    _gotoSheet('VARIANTE')
    oSheet = oDoc.CurrentController.ActiveSheet
    for lrow in reversed(range(0, ultima_voce(oSheet))):
        if oSheet.getCellByPosition(31,lrow).CellStyle == 'compTagG' :
            oSheet.getCellByPosition(31,lrow).String = ''
            oSheet.getCellByPosition(32,lrow).String = ''
            oSheet.getCellByPosition(33,lrow).String = ''
            oSheet.getCellByPosition(34,lrow).String = ''
            oSheet.getCellByPosition(35,lrow).String = ''
def Filtra_computo(nSheet, nCol, sString):
    '''
    nSheet   { string } : nome Sheet
    ncol     { integer } : colonna di tag
    sString  { string } : stringa di tag
    crea una nuova sheet contenente le sole voci filtrate
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    copia_sheet(nSheet, sString)
    oSheet = oDoc.CurrentController.ActiveSheet
    for lrow in reversed(range(0, ultima_voce(oSheet))):
        try:
            sStRange = Circoscrive_Voce_Computo_Att(lrow)
            sopra = sStRange.RangeAddress.StartRow
            sotto = sStRange.RangeAddress.EndRow
            if nCol ==1:
                test=sopra+1
            else:
                test=sotto
            if sString != oSheet.getCellByPosition(nCol,test).String:
                oSheet.getRows().removeByIndex(sopra, sotto-sopra+1)
                lrow =next_voice(lrow,0)
        except:
            lrow =next_voice(lrow,0)
    for lrow in range(3, getLastUsedCell(oSheet).EndRow):
        if oSheet.getCellByPosition(18,lrow).CellStyle == 'Livello-1-scritta mini val' and \
        oSheet.getCellByPosition(18,lrow).Value == 0 or \
        oSheet.getCellByPosition(18,lrow).CellStyle == 'livello2 scritta mini' and \
        oSheet.getCellByPosition(18,lrow).Value == 0:

            oSheet.getRows().removeByIndex(lrow, 1)

    #~ iCellAttr =(oDoc.createInstance("com.sun.star.sheet.CellFlags.OBJECTS"))
    flags = OBJECTS
    oSheet.getCellRangeByPosition(0,0,42,0).clearContents(flags) #cancello gli oggetti
    oDoc.CurrentController.select(oSheet.getCellByPosition(0,3))
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
########################################################################
def Vai_a_filtro(arg=None):
    _gotoSheet('S3')
    _primaCella(0,1)
########################################################################
def Filtra_Computo_Cap(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7,8).String
    sString = oSheet.getCellByPosition(7,10).String
    Filtra_computo(nSheet, 31, sString)
########################################################################
def Filtra_Computo_SottCap(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7, 8).String
    sString = oSheet.getCellByPosition(7, 12).String
    Filtra_computo(nSheet, 32, sString)
########################################################################
def Filtra_Computo_A(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7, 8).String
    sString = oSheet.getCellByPosition(7, 14).String
    Filtra_computo(nSheet, 33, sString)
########################################################################
def Filtra_Computo_B(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7, 8).String
    sString = oSheet.getCellByPosition(7, 16).String
    Filtra_computo(nSheet, 34, sString)
########################################################################
def Filtra_Computo_C(arg=None): #filtra in base al codice di prezzo
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7, 8).String
    sString = oSheet.getCellByPosition(7, 20).String
    Filtra_computo(nSheet, 1, sString)
########################################################################
def Vai_a_M1(arg=None):
    _gotoSheet('M1', 85)
    _primaCella(0,0)
########################################################################
def Vai_a_S2(arg=None):
    _gotoSheet('S2')
########################################################################
def Vai_a_S1(arg=None):
    _gotoSheet('S1')
    _primaCella(0,190)
########################################################################
def Vai_a_ElencoPrezzi(arg=None):
    _gotoSheet('Elenco Prezzi')
########################################################################
def Vai_a_Computo(arg=None):
    _gotoSheet('COMPUTO')
########################################################################
def Vai_a_Variabili(arg=None):
    _gotoSheet('S1', 85)
    _primaCella(6,289)
########################################################################
def Vai_a_Scorciatoie(arg=None):
    _gotoSheet('Scorciatoie')
    _primaCella(0,0)
########################################################################
def Vai_a_SegnaVoci(arg=None):
    _gotoSheet('S3',100)
    _primaCella(37,4)
########################################################################
def _gotoSheet(nSheet, fattore=100):
    '''
    nSheet   { string } : nome Sheet
    attiva e seleziona una sheet
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.Sheets.getByName(nSheet)
    oSheet.IsVisible = True
    oDoc.CurrentController.setActiveSheet(oSheet)
    #~ oDoc.CurrentController.ZoomValue = fattore

     #~ oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
########################################################################
def _primaCella(IDcol=0, IDrow=0):
    '''
    IDcol   { integer } : id colonna
    IDrow   { integer } : id riga
    settaggio prima cella visibile(IDcol, IDrow)
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oDoc.CurrentController.setFirstVisibleColumn(IDcol)
    oDoc.CurrentController.setFirstVisibleRow(IDrow)
    return
########################################################################
def ordina_col(ncol):
    '''
    ncol   { integer } : id colonna
    ordina i dati secondo la colonna con id ncol
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    oProp = []
    oProp0 = PropertyValue()
    oProp0.Name = 'ByRows'
    oProp0.Value = True
    oProp1 = PropertyValue()
    oProp1.Name = 'HasHeader'
    oProp1.Value = False
    oProp2 = PropertyValue()
    oProp2.Name = 'CaseSensitive'
    oProp2.Value = False
    oProp3 = PropertyValue()
    oProp3.Name = 'NaturalSort'
    oProp3.Value = False
    oProp4 = PropertyValue()
    oProp4.Name = 'IncludeAttribs'
    oProp4.Value = True
    oProp5 = PropertyValue()
    oProp5.Name = 'UserDefIndex'
    oProp5.Value = 0
    oProp6 = PropertyValue()
    oProp6.Name = 'Col1'
    oProp6.Value = ncol
    oProp7 = PropertyValue()
    oProp7.Name = 'Ascending1'
    oProp7.Value = True
    oProp.append(oProp0)
    oProp.append(oProp1)
    oProp.append(oProp2)
    oProp.append(oProp3)
    oProp.append(oProp4)
    oProp.append(oProp5)
    oProp.append(oProp6)
    oProp.append(oProp7)
    properties = tuple(oProp)
    dispatchHelper.executeDispatch(oFrame, '.uno:DataSort', '', 0, properties)
########################################################################
def setTabColor(colore):
    '''
    colore   { integer } : id colore
    attribuisce al foglio corrente un colore a scelta
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    oProp = PropertyValue()
    oProp.Name = 'TabBgColor'
    oProp.Value = colore
    properties =(oProp,)
    dispatchHelper.executeDispatch(oFrame, '.uno:SetTabBgColor', '', 0, properties)
########################################################################
def show_sheets(x=True):
    '''
    x   { boolean } : True = ON, False = OFF
    
    Mastra/nasconde tutte le tabelle ad escluzione di COMPUTO ed Elenco Prezzi
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheets = list(oDoc.getSheets().getElementNames())
    oSheets.remove('COMPUTO')
    oSheets.remove('Elenco Prezzi')
    for nome in oSheets:
        oSheet = oDoc.getSheets().getByName(nome)
        oSheet.IsVisible = x
def nascondi_sheets(arg=None):
    show_sheets(False)
########################################################################
def salva_come(nomefile=None):
    '''
    nomefile   { string } : nome del file di destinazione
    Se presente l'argomento nomefile, salva il file corrente in nomefile.
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    
    oProp = []
    if nomefile != None:
        nomefile = uno.systemPathToFileUrl(nomefile)
        oProp0 = PropertyValue()
        oProp0.Name = "URL"
        oProp0.Value = nomefile
        oProp.append(oProp0)

    oProp1 = PropertyValue()
    oProp1.Name = "FilterName"
    oProp1.Value = "calc8"
    oProp.append(oProp1)
    
    properties = tuple(oProp)

    dispatchHelper.executeDispatch(oFrame, ".uno:SaveAs", "", 0, properties)
########################################################################
def _gotoCella(IDcol=0, IDrow=0):
    '''
    IDcol   { integer } : id colonna
    IDrow   { integer } : id riga

    muove il cursore nelle cella(IDcol, IDrow)
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    
    oDoc.CurrentController.select(oSheet.getCellByPosition(IDcol, IDrow))
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
    return
########################################################################
def adatta_altezza_riga(nSheet=None):
    '''
    nSheet   { string } : nSheet della sheet
    imposta l'altezza ottimale delle celle
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if nSheet == None:
        nSheet = oSheet.Name
    oDoc.getSheets().hasByName(nSheet)
    oSheet.getCellRangeByPosition(0, 0, getLastUsedCell(oSheet).EndColumn, getLastUsedCell(oSheet).EndRow).Rows.OptimalHeight = True
    if oSheet.Name in('Elenco Prezzi', 'VARIANTE', 'COMPUTO', 'CONTABILITA'):
        oSheet.getCellByPosition(0, 2).Rows.Height = 800
########################################################################
# elenco prezzi ########################################################
#~ def debug(arg=None):
def scelta_viste(arg=None):
    '''
    Gestisce i dialoghi del menù viste nelle tabelle di Analisi di Prezzo,
    Elenco Prezzi, COMPUTO, VARIANTE, CONTABILITA'
    Genera i raffronti tra COMPUTO e VARIANTE e CONTABILITA'
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    psm = uno.getComponentContext().ServiceManager
    dp = psm.createInstance('com.sun.star.awt.DialogProvider')
    if oSheet.Name in('VARIANTE', 'COMPUTO'):
        oDialog1 = dp.createDialog('vnd.sun.star.script:UltimusFree2.DialogViste_A?language=Basic&location=application')
        oDialog1Model = oDialog1.Model
        if oSheet.getColumns().getByIndex(5).Columns.IsVisible  == True: oDialog1.getControl('CBMis').State = 1
        if oSheet.getColumns().getByIndex(17).Columns.IsVisible  == True: oDialog1.getControl('CBSic').State = 1
        if oSheet.getColumns().getByIndex(29).Columns.IsVisible  == True: oDialog1.getControl('CBMdo').State = 1
        if oSheet.getColumns().getByIndex(31).Columns.IsVisible  == True: oDialog1.getControl('CBCat').State = 1
        #~ if oSheet.getColumns().getByIndex(34).Columns.IsVisible  == True: oDialog1.getControl('CBTag').State = 1
        if oSheet.getColumns().getByIndex(38).Columns.IsVisible  == True: oDialog1.getControl('CBFig').State = 1
        oDialog1.execute()
        
        if oDialog1.getControl('OBTerra').State == True:
            computo_terra_terra()
            oDialog1.getControl('CBSic').State = 0
            oDialog1.getControl('CBMdo').State = 0
            oDialog1.getControl('CBCat').State = 0
            #~ oDialog1.getControl('CBTag').State = 0
            oDialog1.getControl('CBFig').State = 0

        if oDialog1.getControl('CBMdo').State == 0: #manodopera
            oSheet.getColumns().getByIndex(28).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(29).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(30).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(28).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(29).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(30).Columns.IsVisible = True
        
        if oDialog1.getControl('CBCat').State == 0: #categorie
            oSheet.getColumns().getByIndex(31).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(32).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(33).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(31).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(32).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(33).Columns.IsVisible = True

        #~ if oDialog1.getControl("CBTag").State == 0: #TAG
            #~ oSheet.getColumns().getByIndex(34).Columns.IsVisible = False
        #~ else:
            #~ oSheet.getColumns().getByIndex(34).Columns.IsVisible = True

        if oDialog1.getControl("CBSic").State == 0: #sicurezza
            oSheet.getColumns().getByIndex(17).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(17).Columns.IsVisible = True

        if oDialog1.getControl("CBFig").State == 0: #figure
            oSheet.getColumns().getByIndex(38).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(38).Columns.IsVisible = True
            
        if oDialog1.getControl("CBMis").State == 0: #misure
            oSheet.getColumns().getByIndex(5).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(6).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(7).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(8).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(5).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(6).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(7).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(8).Columns.IsVisible = True            

        if oDialog1.getControl("CBDet").State == 0: #
            basic_LeenO('Magic.Formula_magica_off')
        else:
            basic_LeenO('Magic.Formula_magica_aggiorna')
            
    elif oSheet.Name in('Elenco Prezzi'):
        oDialog1 = dp.createDialog("vnd.sun.star.script:UltimusFree2.DialogViste_EP?language=Basic&location=application")
        oDialog1Model = oDialog1.Model
        if oSheet.getColumns().getByIndex(3).Columns.IsVisible  == True: oDialog1.getControl('CBSic').State = 1
        if oSheet.getColumns().getByIndex(5).Columns.IsVisible  == True: oDialog1.getControl('CBMdo').State = 1
        if oSheet.getCellByPosition(1, 3).Rows.OptimalHeight == False: oDialog1.getControl('CBDesc').State = 1
        if oSheet.getColumns().getByIndex(7).Columns.IsVisible  == True: oDialog1.getControl('CBOrig').State = 1
        if oSheet.getColumns().getByIndex(8).Columns.IsVisible  == True: oDialog1.getControl('CBUsa').State = 1
        oDialog1.execute()

        if oDialog1.getControl("CBSic").State == 0: #sicurezza
            oSheet.getColumns().getByIndex(3).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(3).Columns.IsVisible = True

        if oDialog1.getControl("CBMdo").State == 0: #sicurezza
            oSheet.getColumns().getByIndex(5).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(6).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(5).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(6).Columns.IsVisible = True

        if oDialog1.getControl("CBDesc").State == 1: #descrizione
            oSheet.getColumns().getByIndex(3).Columns.IsVisible = False
            oSheet.getCellByPosition(1, 3).Rows.OptimalHeight
            basic_LeenO('Strutture.Tronca_altezza_voci_ep')
        elif oDialog1.getControl("CBDesc").State == 0: adatta_altezza_riga(oSheet.Name)

        if oDialog1.getControl("CBOrig").State == 0: #origine
            oSheet.getColumns().getByIndex(7).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(7).Columns.IsVisible = True
        
        if oDialog1.getControl("CBSom").State == 1:
            genera_sommario()

        #~ if oDialog1.getControl("CBUsa").State == 0: #usato
            #~ oSheet.getColumns().getByIndex(8).Columns.IsVisible = False
            #~ oSheet.getColumns().getByIndex(9).Columns.IsVisible = False
        #~ else:
            #~ oSheet.getColumns().getByIndex(8).Columns.IsVisible = True
            #~ oSheet.getColumns().getByIndex(9).Columns.IsVisible = True
            
        oRangeAddress=oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
        SR = oRangeAddress.StartRow+1
        ER = oRangeAddress.EndRow-1
        oSheet.ungroup(oRangeAddress,0) #colonne
        oSheet.ungroup(oRangeAddress,1) #righe
        oSheet.getCellRangeByPosition(15, 0, 26, 0).Columns.IsVisible = True
        oSheet.getCellRangeByPosition(23 , SR, 25, ER).CellStyle = 'EP statistiche'
        oSheet.getCellRangeByPosition(26, SR, 26, ER+1).CellStyle = 'EP-mezzo %'
        
        formule = list()
        if oDialog1.getControl("ComVar").State == True: #Computo - Variante
            genera_sommario()
            oRangeAddress.StartColumn = 19
            oRangeAddress.EndColumn = 22
            oSheet.group(oRangeAddress,0)
            oSheet.getCellByPosition(23, 0).String = 'Computo - Variante'
            for n in range(4, ultima_voce(oSheet)+2):
                formule.append(['=IF(Q' + str(n) + '-M' + str(n) + '=0;"--";Q' + str(n) + '-M' + str(n) + ')',
                                '=IF(R' + str(n) + '-N' + str(n) + '>0;R' + str(n) + '-N' + str(n) + ';"")',
                                '=IF(R' + str(n) + '-N' + str(n) + '<0;N' + str(n) + '-R' + str(n) + ';"")',
'=IFERROR(IFS(AND(N' + str(n) + '>R' + str(n) + ';R' + str(n) + '=0);-100;AND(N' + str(n) + '<R' + str(n) + ';N' + str(n) + '=0);100;N' + str(n) + '=R' + str(n) + ';"--";N' + str(n) + '>R' + str(n) + ';-(N' + str(n) + '-R' + str(n) + ')/N' + str(n) + ';N' + str(n) + '<R' + str(n) + ';-(N' + str(n) + '-R' + str(n) + ')/N' + str(n) + ')/100;"--")'])

            n += 1
            for el in(1, ER+1):
                oSheet.getCellByPosition(26, el).Formula = '=IFERROR(IFS(AND(N' + str(n) + '>R' + str(n) + ';R' + str(n) + '=0);-100;AND(N' + str(n) + '<R' + str(n) + ';N' + str(n) + '=0);100;N' + str(n) + '=R' + str(n) + ';"--";N' + str(n) + '>R' + str(n) + ';-(N' + str(n) + '-R' + str(n) + ')/N' + str(n) + ';N' + str(n) + '<R' + str(n) + ';-(N' + str(n) + '-R' + str(n) + ')/N' + str(n) + ')/100;"--")'
            oRange = oSheet.getCellRangeByPosition(23, 3, 26, ultima_voce(oSheet))
            formule = tuple(formule)
            oRange.setFormulaArray(formule)

        if oDialog1.getControl("ComCon").State == True: #Computo - Contabilità
            genera_sommario()
            oRangeAddress.StartColumn = 15
            oRangeAddress.EndColumn = 18
            oSheet.group(oRangeAddress,0)
            oSheet.getCellByPosition(23, 0).String = 'Computo - Contabilità'
            for n in range(4, ultima_voce(oSheet)+2):
                formule.append(['=IF(U' + str(n) + '-M' + str(n) + '=0;"--";U' + str(n) + '-M' + str(n) + ')',
                                '=IF(V' + str(n) + '-N' + str(n) + '>0;V' + str(n) + '-N' + str(n) + ';"")',
                                '=IF(V' + str(n) + '-N' + str(n) + '<0;N' + str(n) + '-V' + str(n) + ';"")',
'=IFERROR(IFS(AND(N' + str(n) + '>V' + str(n) + ';V' + str(n) + '=0);-100;AND(N' + str(n) + '<V' + str(n) + ';N' + str(n) + '=0);100;N' + str(n) + '=V' + str(n) + ';"--";N' + str(n) + '>V' + str(n) + ';-(N' + str(n) + '-V' + str(n) + ')/N' + str(n) + ';N' + str(n) + '<V' + str(n) + ';-(N' + str(n) + '-V' + str(n) + ')/N' + str(n) + ')/100;"--")'])
            n += 1
            for el in(1, ER+1):
                oSheet.getCellByPosition(26, el).Formula = '=IFERROR(IFS(AND(N' + str(n) + '>V' + str(n) + ';V' + str(n) + '=0);-100;AND(N' + str(n) + '<V' + str(n) + ';N' + str(n) + '=0);100;N' + str(n) + '=V' + str(n) + ';"--";N' + str(n) + '>V' + str(n) + ';-(N' + str(n) + '-V' + str(n) + ')/N' + str(n) + ';N' + str(n) + '<V' + str(n) + ';-(N' + str(n) + '-V' + str(n) + ')/N' + str(n) + ')/100;"--")'
            oRange = oSheet.getCellRangeByPosition(23, 3, 26, ultima_voce(oSheet))
            formule = tuple(formule)
            oRange.setFormulaArray(formule)

        if oDialog1.getControl("VarCon").State == True: #Variante - Contabilità
            genera_sommario()
            oRangeAddress.StartColumn = 11
            oRangeAddress.EndColumn = 14
            oSheet.group(oRangeAddress, 0)
            oSheet.getCellByPosition(23, 0).String = 'Variante - Contabilità'
            for n in range(4, ultima_voce(oSheet)+2):
                formule.append(['=IF(U' + str(n) + '-Q' + str(n) + '=0;"--";U' + str(n) + '-Q' + str(n) + ')',
                                '=IF(V' + str(n) + '-R' + str(n) + '>0;V' + str(n) + '-R' + str(n) + ';"")',
                                '=IF(V' + str(n) + '-R' + str(n) + '<0;R' + str(n) + '-V' + str(n) + ';"")',
'=IFERROR(IFS(AND(R' + str(n) + '>V' + str(n) + ';V' + str(n) + '=0);-100;AND(R' + str(n) + '<V' + str(n) + ';R' + str(n) + '=0);100;R' + str(n) + '=V' + str(n) + ';"--";R' + str(n) + '>V' + str(n) + ';-(R' + str(n) + '-V' + str(n) + ')/R' + str(n) + ';R' + str(n) + '<V' + str(n) + ';-(R' + str(n) + '-V' + str(n) + ')/R' + str(n) + ')/100;"--")'])
            n += 1
            for el in(1, ER+1):
                oSheet.getCellByPosition(26, el).Formula = '=IFERROR(IFS(AND(R' + str(n) + '>V' + str(n) + ';V' + str(n) + '=0);-100;AND(R' + str(n) + '<V' + str(n) + ';R' + str(n) + '=0);100;R' + str(n) + '=V' + str(n) + ';"--";R' + str(n) + '>V' + str(n) + ';-(R' + str(n) + '-V' + str(n) + ')/R' + str(n) + ';R' + str(n) + '<V' + str(n) + ';-(R' + str(n) + '-V' + str(n) + ')/R' + str(n) + ')/100;"--")'
            oRange = oSheet.getCellRangeByPosition(23, 3, 26, ultima_voce(oSheet))
            formule = tuple(formule)
            oRange.setFormulaArray(formule)
        for el in(11, 15, 19, 26):
            oSheet.getCellRangeByPosition(el, 3, el, ultima_voce(oSheet)).CellStyle = 'EP-mezzo %'
        for el in(12, 16, 20, 23):
            oSheet.getCellRangeByPosition(el, 3, el, ultima_voce(oSheet)).CellStyle = 'EP statistiche_q'
        for el in(13, 17, 21, 24, 25):
            oSheet.getCellRangeByPosition(el, 3, el, ultima_voce(oSheet)).CellStyle = 'EP statistiche'
        if DlgSiNo("Nascondo eventuali righe con scostamento nullo?") == 2:
            errori =('#DIV/0!', '--')
            hide_error(errori, 26)
            
    elif oSheet.Name in('Analisi di Prezzo'):
        oDialog1 = dp.createDialog("vnd.sun.star.script:UltimusFree2.DialogViste_AN?language=Basic&location=application")
        oDialog1Model = oDialog1.Model
        if  oSheet.getCellByPosition(1, 2).Rows.OptimalHeight == False: oDialog1.getControl("CBDesc").State = 1 #descrizione breve
        oDialog1.execute()
        
        if  oSheet.getCellByPosition(1, 2).Rows.OptimalHeight == True and oDialog1.getControl("CBDesc").State == 1: #descrizione breve
            basic_LeenO('Strutture.Tronca_Altezza_Analisi')
        elif oDialog1.getControl("CBDesc").State == 0: adatta_altezza_riga(oSheet.Name)
        
    elif oSheet.Name in('CONTABILITA', 'Registro', 'SAL'):
        oDialog1 = dp.createDialog("vnd.sun.star.script:UltimusFree2.Dialogviste_N?language=Basic&location=application")
        oDialog1Model = oDialog1.Model
        oDialog1.execute()
    #~ MsgBox('Operazione eseguita con successo!','')
########################################################################
class genera_sommario_th(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        genera_sommario_run()
def genera_sommario(arg=None):
#~ def debug(arg=None):
    genera_sommario_th().start()
#~ ###
def genera_sommario_run(arg=None):
    '''
    sostituisce la sub Rifa_AA_BB_Computo
    serve a generare i sommari in Elenco Prezzi
    '''
    #~ oDialogo_attesa = dlg_attesa()
    #~ attesa().start() #mostra il dialogo
    refresh(0)

    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('COMPUTO')
    lRow = getLastUsedCell(oSheet).EndRow
    rifa_nomearea('COMPUTO', '$AJ$3:$AJ$' + str(lRow), 'AA')
    rifa_nomearea('COMPUTO', '$N$3:$N$'  + str(lRow), "BB")
    rifa_nomearea('COMPUTO', '$AK$3:$AK$' + str(lRow), "cEuro")

    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    if oDoc.getSheets().hasByName('VARIANTE') == True:
        rifa_nomearea('VARIANTE', '$AJ$3:$AJ$' + str(lRow), 'varAA')
        rifa_nomearea('VARIANTE', '$N$3:$N$'  + str(lRow), "varBB")
        rifa_nomearea('VARIANTE', '$AK$3:$AK$' + str(lRow), "varEuro")

    if oDoc.getSheets().hasByName('CONTABILITA') == True:
        lRow = getLastUsedCell(oDoc.getSheets().getByName('CONTABILITA')).EndRow
        rifa_nomearea('CONTABILITA', '$AJ$3:$AJ$' + str(lRow), 'GG')
        rifa_nomearea('CONTABILITA', '$S$3:$S$'  + str(lRow), "G1G1")
        rifa_nomearea('CONTABILITA', '$AK$3:$AK$' + str(lRow), "conEuro")
        
    formule = list()
    for n in range(4, ultima_voce(oSheet)+2):
        stringa =(['=N' + str(n) + '/$N$2',
                        '=SUMIF(AA;A' + str(n) + ';BB)',
                        '=SUMIF(AA;A' + str(n) + ';cEuro)',
                        '', '', '', '', '', '', '', ''])
        if oDoc.getSheets().hasByName('VARIANTE') == True:
            stringa =(['=N' + str(n) + '/$N$2',
                        '=SUMIF(AA;A' + str(n) + ';BB)',
                        '=SUMIF(AA;A' + str(n) + ';cEuro)',
                        '',
                        '=R' + str(n) + '/$R$2',
                        '=SUMIF(varAA;A' + str(n) + ';varBB)',
                        '=SUMIF(varAA;A' + str(n) + ';varEuro)',
                        '', '', '',
                        ''])
            if oDoc.getSheets().hasByName('CONTABILITA') == True:
                stringa =(['=N' + str(n) + '/$N$2',
                            '=SUMIF(AA;A' + str(n) + ';BB)',
                            '=SUMIF(AA;A' + str(n) + ';cEuro)',
                            '',
                            '=R' + str(n) + '/$R$2',
                            '=SUMIF(varAA;A' + str(n) + ';varBB)',
                            '=SUMIF(varAA;A' + str(n) + ';varEuro)',
                            '',
                            '=V' + str(n) + '/$V$2',
                            '=SUMIF(GG;A' + str(n) + ';G1G1)',
                            '=SUMIF(GG;A' + str(n) + ';conEuro)'])
        elif oDoc.getSheets().hasByName('CONTABILITA') == True:
            stringa =(['=N' + str(n) + '/$N$2',
                        '=SUMIF(AA;A' + str(n) + ';BB)',
                        '=SUMIF(AA;A' + str(n) + ';cEuro)',
                        '',
                        '', '', '',
                        '',
                        '=V' + str(n) + '/$V$2',
                        '=SUMIF(GG;A' + str(n) + ';G1G1)',
                        '=SUMIF(GG;A' + str(n) + ';conEuro)'])
        formule.append(stringa)
    oRange = oSheet.getCellRangeByPosition(11, 3, 21, ultima_voce(oSheet))
    formule = tuple(formule)
    oRange.setFormulaArray(formule)
    refresh(1)
    adatta_altezza_riga(oSheet.Name)
    #~ oDialogo_attesa.endExecute() #chiude il dialogo
########################################################################
def riordina_ElencoPrezzi(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oRangeAddress=oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    IS = oRangeAddress.Sheet
    SC = oRangeAddress.StartColumn
    EC = oRangeAddress.EndColumn
    SR = oRangeAddress.StartRow+1
    ER = oRangeAddress.EndRow-1
    if ER < SR:
        try:
            uFindStringCol('Fine elenco', 0, oSheet)
        except TypeError:
            inserisci_Riga_rossa()
        test = str(uFindStringCol('Fine elenco', 0, oSheet))
        rifa_nomearea('Elenco Prezzi', "$A$3:$AF$" + test, 'elenco_prezzi')
        rifa_nomearea('Elenco Prezzi', "$A$3:$A$" + test, 'Lista')
        oRangeAddress=oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
        ER = oRangeAddress.EndRow-1
    oRange = oSheet.getCellRangeByPosition(SC, SR, EC, ER)
    oDoc.CurrentController.select(oRange)
    ordina_col(1)
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
########################################################################
# L'ATTIVAZIONE DELLA CLASS doppioni_th INTERFERISCE CON invia_voce_ep()
#~ class doppioni_th(threading.Thread):
    #~ def __init__(self):
        #~ threading.Thread.__init__(self)
    #~ def run(self):
        #~ doppioni_run()
#~ def doppioni(arg=None):
    #~ doppioni_th().start()
###
#~ def doppioni_run(arg=None):
def doppioni(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    '''
    Cancella eventuali voci che si ripetono in Elenco Prezzi
    '''
    #~ oDialogo_attes = dlg_attesa()
    #~ attesa().start() #mostra il dialogo
    oDoc.CurrentController.ZoomValue = 400
    refresh(0)
    if oDoc.getSheets().hasByName('Analisi di Prezzo') == True:
        lista_tariffe_analisi = list()
        oSheet = oDoc.getSheets().getByName('Analisi di Prezzo')
        for n in range(0, ultima_voce(oSheet)+1):
            if oSheet.getCellByPosition(0, n).CellStyle == 'An-1_sigla':
                lista_tariffe_analisi.append(oSheet.getCellByPosition(0, n).String)
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')

    oRangeAddress=oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    SR = oRangeAddress.StartRow+1
    ER = oRangeAddress.EndRow-1
    oRange = oSheet.getCellRangeByPosition(0, SR, 7, ER)
    
    if oDoc.getSheets().hasByName('Analisi di Prezzo') == True:
        for i in reversed(range(SR, ER)):
            if oSheet.getCellByPosition(0, i).String in lista_tariffe_analisi:
                oSheet.getRows().removeByIndex(i, 1)
    oRangeAddress=oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    SR = oRangeAddress.StartRow+1
    ER = oRangeAddress.EndRow-1
    oRange = oSheet.getCellRangeByPosition(0, SR, 7, ER)
    lista_come_array = tuple(set(oRange.getDataArray()))
    oSheet.getRows().removeByIndex(SR, ER-SR+1)
    lista_tar = list()
    oSheet.getRows().insertByIndex(SR, len(set(lista_come_array)))
    for el in set(lista_come_array):
        lista_tar.append(el[0])
    colonne_lista = len(lista_come_array[0]) # numero di colonne necessarie per ospitare i dati
    righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati
    oRange = oSheet.getCellRangeByPosition( 0,
                                            3,
                                            colonne_lista + 0 - 1, # l'indice parte da 0
                                            righe_lista + 3 - 1)
    oRange.setDataArray(lista_come_array)
    oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellStyle = "EP-aS"
    oSheet.getCellRangeByPosition(1, 3, 1, righe_lista + 3 - 1).CellStyle = "EP-a"
    oSheet.getCellRangeByPosition(2, 3, 7, righe_lista + 3 - 1).CellStyle = "EP-mezzo"
    oSheet.getCellRangeByPosition(5, 3, 5, righe_lista + 3 - 1).CellStyle = "EP-mezzo %"
    oSheet.getCellRangeByPosition(8, 3, 9, righe_lista + 3 - 1).CellStyle = "EP-sfondo"

    oSheet.getCellRangeByPosition(11, 3, 11, righe_lista + 3 - 1).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition(12, 3, 12, righe_lista + 3 - 1).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition(13, 3, 13, righe_lista + 3 - 1).CellStyle = 'EP statistiche_Contab_q'
    #~ oDoc.CurrentController.select(oRange)
    #~ ordina_col(1)
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
    if oDoc.getSheets().hasByName('Analisi di Prezzo') == True:
        tante_analisi_in_ep()
    riordina_ElencoPrezzi()
    refresh(1)
    oDoc.CurrentController.ZoomValue = 100
    adatta_altezza_riga(oSheet.Name)
    #~ oDialogo_attesa.endExecute() #chiude il dialogo
    if len(set(lista_tar)) != len(set(lista_come_array)):
        MsgBox('Probabilmente ci sono ancora 2 o più voci\nche hanno lo stesso Codice Articolo. Controlla.', 'Attenzione!')
########################################################################
# Scrive un file.
def XPWE_out(arg=None):
    '''
    esporta il documento in formato XPWE
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oDialogo_attesa = dlg_attesa('LETTURA DATI IN CORSO...')
    attesa().start() #mostra il dialogo

    if oDoc.getSheets().hasByName('S2') == False:
        MsgBox('Puoi usare questo comando da un file di computo esistente.','Avviso!')
        return
    top = Element('PweDocumento')
#~ dati generali
    PweDatiGenerali = SubElement(top,'PweDatiGenerali')
    PweMisurazioni = SubElement(top,'PweMisurazioni')
    PweDGProgetto = SubElement(PweDatiGenerali,'PweDGProgetto')
    PweDGDatiGenerali = SubElement(PweDGProgetto,'PweDGDatiGenerali')
    PercPrezzi = SubElement(PweDGDatiGenerali,'PercPrezzi')
    PercPrezzi.text = '0'

    Comune = SubElement(PweDGDatiGenerali,'Comune')
    Provincia = SubElement(PweDGDatiGenerali,'Provincia')
    Oggetto = SubElement(PweDGDatiGenerali,'Oggetto')
    Committente = SubElement(PweDGDatiGenerali,'Committente')
    Impresa = SubElement(PweDGDatiGenerali,'Impresa')
    ParteOpera = SubElement(PweDGDatiGenerali,'ParteOpera')
#~  leggo i dati generali
    oSheet = oDoc.getSheets().getByName('S2')
    Comune.text = oSheet.getCellByPosition(2, 3).String
    Provincia.text = ''
    Oggetto.text = oSheet.getCellByPosition(2, 2).String
    Committente.text = oSheet.getCellByPosition(2, 5).String
    Impresa.text = oSheet.getCellByPosition(2, 16).String
    ParteOpera.text = ''
#~ Capitoli e Categorie
    PweDGCapitoliCategorie = SubElement(PweDatiGenerali,'PweDGCapitoliCategorie')

#~ SuperCategorie
    oSheet = oDoc.getSheets().getByName(arg)
    lastRow = ultima_voce(oSheet)+1
    # evito di esportare in SuperCategorie perché inutile, almeno per ora
    listaspcat = list()
    PweDGSuperCategorie = SubElement(PweDGCapitoliCategorie,'PweDGSuperCategorie')
    for n in range(0, lastRow):
        if oSheet.getCellByPosition(1, n).CellStyle == 'Livello-0-scritta':
            desc = oSheet.getCellByPosition(2, n).String
            if desc not in listaspcat:
                listaspcat.append(desc)
                idID = str(listaspcat.index(desc) +1)

            #~ PweDGSuperCategorie = SubElement(PweDGCapitoliCategorie,'PweDGSuperCategorie')
                DGSuperCategorieItem = SubElement(PweDGSuperCategorie,'DGSuperCategorieItem')
                DesSintetica = SubElement(DGSuperCategorieItem,'DesSintetica')
            
                DesEstesa = SubElement(DGSuperCategorieItem,'DesEstesa')
                DataInit = SubElement(DGSuperCategorieItem,'DataInit')
                Durata = SubElement(DGSuperCategorieItem,'Durata')
                CodFase = SubElement(DGSuperCategorieItem,'CodFase')
                Percentuale = SubElement(DGSuperCategorieItem,'Percentuale')
                Codice = SubElement(DGSuperCategorieItem,'Codice')

                DGSuperCategorieItem.set('ID', idID)
                DesSintetica.text = desc
                DataInit.text = oggi()
                Durata.text = '0'
                Percentuale.text = '0'

#~ Categorie
    listaCat = list()
    PweDGCategorie = SubElement(PweDGCapitoliCategorie,'PweDGCategorie')
    for n in range(0,lastRow):
        if oSheet.getCellByPosition(2, n).CellStyle == 'Livello-1-scritta mini':
            desc = oSheet.getCellByPosition(2, n).String
            if desc not in listaCat:
                listaCat.append(desc)
                idID = str(listaCat.index(desc) +1)

                #~ PweDGCategorie = SubElement(PweDGCapitoliCategorie,'PweDGCategorie')
                DGCategorieItem = SubElement(PweDGCategorie,'DGCategorieItem')
                DesSintetica = SubElement(DGCategorieItem,'DesSintetica')
                
                DesEstesa = SubElement(DGCategorieItem,'DesEstesa')
                DataInit = SubElement(DGCategorieItem,'DataInit')
                Durata = SubElement(DGCategorieItem,'Durata')
                CodFase = SubElement(DGCategorieItem,'CodFase')
                Percentuale = SubElement(DGCategorieItem,'Percentuale')
                Codice = SubElement(DGCategorieItem,'Codice')

                DGCategorieItem.set('ID', idID)
                DesSintetica.text = desc
                DataInit.text = oggi()
                Durata.text = '0'
                Percentuale.text = '0'

#~ SubCategorie
    listasbCat = list()
    PweDGSubCategorie = SubElement(PweDGCapitoliCategorie,'PweDGSubCategorie')
    for n in range(0,lastRow):
        if oSheet.getCellByPosition(2, n).CellStyle == 'livello2_':
            desc = oSheet.getCellByPosition(2, n).String
            if desc not in listasbCat:
                listasbCat.append(desc)
                idID = str(listasbCat.index(desc) +1)

                #~ PweDGSubCategorie = SubElement(PweDGCapitoliCategorie,'PweDGSubCategorie')
                DGSubCategorieItem = SubElement(PweDGSubCategorie,'DGSubCategorieItem')
                DesSintetica = SubElement(DGSubCategorieItem,'DesSintetica')

                DesEstesa = SubElement(DGSubCategorieItem,'DesEstesa')
                DataInit = SubElement(DGSubCategorieItem,'DataInit')
                Durata = SubElement(DGSubCategorieItem,'Durata')
                CodFase = SubElement(DGSubCategorieItem,'CodFase')
                Percentuale = SubElement(DGSubCategorieItem,'Percentuale')
                Codice = SubElement(DGSubCategorieItem,'Codice')

                DGSubCategorieItem.set('ID', idID)
                DesSintetica.text = desc
                DataInit.text = oggi()
                Durata.text = '0'
                Percentuale.text = '0'

#~ Moduli
    PweDGModuli = SubElement(PweDatiGenerali,'PweDGModuli')
    PweDGAnalisi = SubElement(PweDGModuli,'PweDGAnalisi')
    SpeseUtili = SubElement(PweDGAnalisi,'SpeseUtili')
    SpeseGenerali = SubElement(PweDGAnalisi,'SpeseGenerali')
    UtiliImpresa = SubElement(PweDGAnalisi,'UtiliImpresa')
    OneriAccessoriSc = SubElement(PweDGAnalisi,'OneriAccessoriSc')
    ConfQuantita = SubElement(PweDGAnalisi,'ConfQuantita')

    oSheet = oDoc.getSheets().getByName('S1')
    if oSheet.getCellByPosition(7,322).Value ==0: # se 0: Spese e Utili Accorpati
        SpeseUtili.text = '1'
    else:
        SpeseUtili.text = '-1'
        
    UtiliImpresa.text = oSheet.getCellByPosition(7,320).String[:-1].replace(',','.')
    OneriAccessoriSc.text = oSheet.getCellByPosition(7,318).String[:-1].replace(',','.')
    SpeseGenerali.text = oSheet.getCellByPosition(7,319).String[:-1].replace(',','.')

#~ Elenco Prezzi
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    PweElencoPrezzi = SubElement(PweMisurazioni,'PweElencoPrezzi')
    diz_ep = dict()
    lista_AP = list()
    for n in range(3, getLastUsedCell(oSheet).EndRow):
        barra_di_stato(str(n) + ' ' + str(getLastUsedCell(oSheet).EndRow))
        if oSheet.getCellByPosition(1, n).Type.value == 'TEXT' and \
        oSheet.getCellByPosition(2, n).Type.value == 'TEXT':
            EPItem = SubElement(PweElencoPrezzi,'EPItem')
            EPItem.set('ID', str(n))
            TipoEP = SubElement(EPItem,'TipoEP')
            TipoEP.text = '0'
            Tariffa = SubElement(EPItem,'Tariffa')
            id_tar = str(n)
            Tariffa.text = oSheet.getCellByPosition(0, n).String
            diz_ep[oSheet.getCellByPosition(0, n).String] = id_tar
            Articolo = SubElement(EPItem,'Articolo')
            Articolo.text = ''
            DesEstesa = SubElement(EPItem,'DesEstesa')
            DesEstesa.text = oSheet.getCellByPosition(1, n).String
            DesRidotta = SubElement(EPItem,'DesRidotta')
            if len(DesEstesa.text) > 120:
                DesRidotta.text = DesEstesa.text[:60] + ' ... ' + DesEstesa.text[-60:]
            else:
                DesRidotta.text = DesEstesa.text
            DesBreve = SubElement(EPItem,'DesBreve')
            if len(DesEstesa.text) > 60:
                DesBreve.text = DesEstesa.text[:30] + ' ... ' + DesEstesa.text[-30:]
            else:
                DesBreve.text = DesEstesa.text
            UnMisura = SubElement(EPItem,'UnMisura')
            UnMisura.text = oSheet.getCellByPosition(2, n).String
            Prezzo1 = SubElement(EPItem,'Prezzo1')
            Prezzo1.text = str(oSheet.getCellByPosition(4, n).Value)
            Prezzo2 = SubElement(EPItem,'Prezzo2')
            Prezzo2.text = '0'
            Prezzo3 = SubElement(EPItem,'Prezzo3')
            Prezzo3.text = '0'
            Prezzo4 = SubElement(EPItem,'Prezzo4')
            Prezzo4.text = '0'
            Prezzo5 = SubElement(EPItem,'Prezzo5')
            Prezzo5.text = '0'
            IDSpCap = SubElement(EPItem,'IDSpCap')
            IDSpCap.text = '0'
            IDCap = SubElement(EPItem,'IDCap')
            IDCap.text = '0'
            IDSbCap = SubElement(EPItem,'IDSbCap')
            IDSbCap.text = '0'
            Flags = SubElement(EPItem,'Flags')
            if oSheet.getCellByPosition(8, n).String  == '(AP)':
                Flags.text = '131072'
            else:
                Flags.text = '0'
            Data = SubElement(EPItem,'Data')
            Data.text = '30/12/1899'
            AdrInternet = SubElement(EPItem,'AdrInternet')
            AdrInternet.text = ''
            PweEPAnalisi = SubElement(EPItem,'PweEPAnalisi')
            xlo_sic = SubElement(EPItem,'xlo_sic')
            if oSheet.getCellByPosition(3, n).Value == 0.0:
                xlo_sic.text = ''
            else:
                xlo_sic.text = str(oSheet.getCellByPosition(3, n).Value)
            xlo_mdop = SubElement(EPItem,'xlo_mdop')
            if oSheet.getCellByPosition(5, n).Value == 0.0:
                xlo_mdop.text = ''
            else:
                xlo_mdop.text = str(oSheet.getCellByPosition(5, n).Value)
            xlo_mdo = SubElement(EPItem,'xlo_mdo')
            if oSheet.getCellByPosition(6, n).Value == 0.0:
                xlo_mdo.text = ''
            else:
                xlo_mdo.text = str(oSheet.getCellByPosition(6, n).Value)
        elif oSheet.getCellByPosition(1, n).Type.value == 'FORMULA' and \
        oSheet.getCellByPosition(2, n).Type.value == 'FORMULA':
            lista_AP.append(oSheet.getCellByPosition(0, n).String)
#Analisi di prezzo
    if len(lista_AP) != 0:
        k = n+1
        oSheet = oDoc.getSheets().getByName('Analisi di Prezzo')
        for el in lista_AP:
            try:
                n =(uFindStringCol(el, 0, oSheet))
                EPItem = SubElement(PweElencoPrezzi,'EPItem')
                EPItem.set('ID', str(k))
                TipoEP = SubElement(EPItem,'TipoEP')
                TipoEP.text = '0'
                Tariffa = SubElement(EPItem,'Tariffa')
                id_tar = str(k)
                Tariffa.text = oSheet.getCellByPosition(0, n).String
                diz_ep[oSheet.getCellByPosition(0, n).String] = id_tar
                Articolo = SubElement(EPItem,'Articolo')
                Articolo.text = ''
                DesEstesa = SubElement(EPItem,'DesEstesa')
                DesEstesa.text = oSheet.getCellByPosition(1, n).String
                DesRidotta = SubElement(EPItem,'DesRidotta')
                if len(DesEstesa.text) > 120:
                    DesRidotta.text = DesEstesa.text[:60] + ' ... ' + DesEstesa.text[-60:]
                else:
                    DesRidotta.text = DesEstesa.text
                DesBreve = SubElement(EPItem,'DesBreve')
                if len(DesEstesa.text) > 60:
                    DesBreve.text = DesEstesa.text[:30] + ' ... ' + DesEstesa.text[-30:]
                else:
                    DesBreve.text = DesEstesa.text

                UnMisura = SubElement(EPItem,'UnMisura')
                UnMisura.text = oSheet.getCellByPosition(2, n).String
                Prezzo1 = SubElement(EPItem,'Prezzo1')
                Prezzo1.text = str(oSheet.getCellByPosition(6, n).Value)
                Prezzo2 = SubElement(EPItem,'Prezzo2')
                Prezzo2.text = '0'
                Prezzo3 = SubElement(EPItem,'Prezzo3')
                Prezzo3.text = '0'
                Prezzo4 = SubElement(EPItem,'Prezzo4')
                Prezzo4.text = '0'
                Prezzo5 = SubElement(EPItem,'Prezzo5')
                Prezzo5.text = '0'
                IDSpCap = SubElement(EPItem,'IDSpCap')
                IDSpCap.text = '0'
                IDCap = SubElement(EPItem,'IDCap')
                IDCap.text = '0'
                IDSbCap = SubElement(EPItem,'IDSbCap')
                IDSbCap.text = '0'
                Flags = SubElement(EPItem,'Flags')
                Flags.text = '131072'
                Data = SubElement(EPItem,'Data')
                Data.text = '30/12/1899'
                AdrInternet = SubElement(EPItem,'AdrInternet')
                AdrInternet.text = ''
                PweEPAnalisi = SubElement(EPItem,'PweEPAnalisi')
                PweEPAR = SubElement(PweEPAnalisi,'PweEPAR')
                nEPARItem = 2
                for x in range(n, n+100):
                    if oSheet.getCellByPosition(0, x).CellStyle == 'An-lavoraz-Cod-sx' and \
                    oSheet.getCellByPosition(1, x).String != '':
                        EPARItem = SubElement(PweEPAR,'EPARItem')
                        EPARItem.set('ID', str(nEPARItem))
                        nEPARItem += 1
                        Tipo = SubElement(EPARItem,'Tipo')
                        Tipo.text = '1'
                        IDEP = SubElement(EPARItem,'IDEP')
                        IDEP.text = diz_ep.get(oSheet.getCellByPosition(0, x).String)
                        if IDEP.text == None:
                            IDEP.text ='-2'
                        Descrizione = SubElement(EPARItem,'Descrizione')
                        if '=IF(' in oSheet.getCellByPosition(1, x).String:
                            Descrizione.text = ''
                        else:
                            Descrizione.text = oSheet.getCellByPosition(1, x).String
                        Misura = SubElement(EPARItem,'Misura')
                        Misura.text = oSheet.getCellByPosition(2, x).String
                        Qt = SubElement(EPARItem,'Qt')
                        Qt.text = oSheet.getCellByPosition(3, x).String.replace(',','.')
                        Prezzo = SubElement(EPARItem,'Prezzo')
                        Prezzo.text = str(oSheet.getCellByPosition(4, x).Value).replace(',','.')
                        FieldCTL = SubElement(EPARItem,'FieldCTL')
                        FieldCTL.text = '0'
                    elif oSheet.getCellByPosition(0, x).CellStyle == 'An-sfondo-basso Att End':
                        break

                xlo_sic = SubElement(EPItem,'xlo_sic')
                if oSheet.getCellByPosition(10, n).Value == 0.0:
                    xlo_sic.text = ''
                else:
                    xlo_sic.text = str(oSheet.getCellByPosition(10, n).Value)
                    
                xlo_mdop = SubElement(EPItem,'xlo_mdop')
                if oSheet.getCellByPosition(8, n).Value == 0.0:
                    xlo_mdop.text = ''
                else:
                    xlo_mdop.text = str(oSheet.getCellByPosition(5, n).Value)
                
                xlo_mdo = SubElement(EPItem,'xlo_mdo')
                if oSheet.getCellByPosition(9, n).Value == 0.0:
                    xlo_mdo.text = ''
                else:
                    xlo_mdo.text = str(oSheet.getCellByPosition(9, n).Value)
                k += 1
            except:
                pass
#COMPUTO/VARIANTE
    oSheet = oDoc.getSheets().getByName(arg)
    PweVociComputo = SubElement(PweMisurazioni,'PweVociComputo')
    oDoc.CurrentController.setActiveSheet(oSheet)
    #~ oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
    nVCItem = 2
    for n in range(0, ultima_voce(oSheet)):
        if oSheet.getCellByPosition(0, n).CellStyle == 'Comp Start Attributo':
            sStRange = Circoscrive_Voce_Computo_Att(n)
            sStRange.RangeAddress
            sopra = sStRange.RangeAddress.StartRow
            sotto = sStRange.RangeAddress.EndRow

            VCItem = SubElement(PweVociComputo,'VCItem')
            VCItem.set('ID', str(nVCItem))
            nVCItem += 1

            IDEP = SubElement(VCItem,'IDEP')
            IDEP.text = diz_ep.get(oSheet.getCellByPosition(1, sopra+1).String)
##########################
            Quantita = SubElement(VCItem,'Quantita')
            Quantita.text = oSheet.getCellByPosition(9, sotto).String
##########################
            DataMis = SubElement(VCItem,'DataMis')
            DataMis.text = oggi() #'26/12/1952'#'28/09/2013'###
            vFlags = SubElement(VCItem,'Flags')
            vFlags.text = '0'
##########################
            IDSpCat = SubElement(VCItem,'IDSpCat')
            IDSpCat.text = str(oSheet.getCellByPosition(31, sotto).String)
            if IDSpCat.text == '':
                IDSpCat.text = '0'
##########################
            IDCat = SubElement(VCItem,'IDCat')
            IDCat.text = str(oSheet.getCellByPosition(32, sotto).String)
            if IDCat.text == '':
                IDCat.text = '0'
##########################
            IDSbCat = SubElement(VCItem,'IDSbCat')
            IDSbCat.text = str(oSheet.getCellByPosition(33, sotto).String)
            if IDSbCat.text == '':
                IDSbCat.text = '0'
##########################
            PweVCMisure = SubElement(VCItem,'PweVCMisure')
            for m in range(sopra+2, sotto):
                RGItem = SubElement(PweVCMisure,'RGItem')
                x = 2
                RGItem.set('ID', str(x))
                x += 1
##########################
                IDVV = SubElement(RGItem,'IDVV')
                IDVV.text = '-2'
##########################
                Descrizione = SubElement(RGItem,'Descrizione')
                Descrizione.text = oSheet.getCellByPosition(2, m).String
##########################
                PartiUguali = SubElement(RGItem,'PartiUguali')
                PartiUguali.text = valuta_cella(oSheet.getCellByPosition(5, m))
##########################
                Lunghezza = SubElement(RGItem,'Lunghezza')
                Lunghezza.text = valuta_cella(oSheet.getCellByPosition(6, m))
##########################
                Larghezza = SubElement(RGItem,'Larghezza')
                Larghezza.text = valuta_cella(oSheet.getCellByPosition(7, m))
##########################
                HPeso = SubElement(RGItem,'HPeso')
                HPeso.text = valuta_cella(oSheet.getCellByPosition(8, m))
##########################
                Quantita = SubElement(RGItem,'Quantita')
                Quantita.text = str(oSheet.getCellByPosition(9, m).Value)
##########################
                Flags = SubElement(RGItem,'Flags')
                if '-' in Quantita.text:
                    Flags.text = '1'
                elif "Parziale [" in oSheet.getCellByPosition(8, m).String:
                    Flags.text = '2'
                    HPeso.text = ''
                elif 'PARTITA IN CONTO PROVVISORIO' in Descrizione.text:
                    Flags.text = '16'
                else:
                    Flags.text = '0'
##########################
                if 'DETRAE LA PARTITA IN CONTO PROVVISORIO' in Descrizione.text:
                    Flags.text = '32'
                if '- vedi voce n. ' in Descrizione.text:
                    IDVV.text = str(int(Descrizione.text.split(' - vedi voce n. ')[1].split(' ')[0])+1)
                    Flags.text = '32768'
                    PartiUguali.text =''
                    if '-' in Quantita.text:
                        Flags.text = '32769'
            n = sotto+1
##########################
    oDialogo_attesa.endExecute()
    out_file = filedia('Salva con nome...', '*.xpwe', 1)
    try:
        if out_file.split('.')[-1].upper() != 'XPWE':
            out_file = out_file + '-'+ arg + '.xlo.xpwe'
    except AttributeError:
        return
    riga = str(tostring(top, encoding="unicode"))
    #~ if len(lista_AP) != 0:
        #~ riga = riga.replace('<PweDatiGenerali>','<Fgs>131072</Fgs><PweDatiGenerali>')
    try:
        of = codecs.open(out_file,'w','utf-8')
        of.write(riga)
        MsgBox('Esportazione in formato XPWE eseguita con successo\nsul file ' + out_file + '!','Avviso.')
    except:
        MsgBox('Esportazione non eseguita!\n\nVerifica che il file di destinazione non sia già in uso!','E R R O R E !')
########################################################################
def firme_in_calce_run(arg=None):
    oDialogo_attesa = dlg_attesa()# avvia il diaolgo di attesa che viene chiuso alla fine con 
    '''
    Inserisce(in COMPUTO o VARIANTE) un riepilogo delle categorie
    ed i dati necessari alle firme
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()

    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in('Analisi di Prezzo', 'Elenco Prezzi'):
        lRowF = ultima_voce(oSheet)+1
        oDoc.CurrentController.setFirstVisibleRow(lRowF-1)
        lRowE = getLastUsedCell(oSheet).EndRow
        for i in range(lRowF, getLastUsedCell(oSheet).EndRow+1):
            if oSheet.getCellByPosition(0, i).CellStyle == "Riga_rossa_Chiudi":
                lRowE = i
                break
        if lRowE > lRowF+1:
            oSheet.getRows().removeByIndex(lRowF, lRowE-lRowF)
        riga_corrente = lRowF+1
        oSheet.getRows().insertByIndex(lRowF, 15)
        oSheet.getCellRangeByPosition(0,lRowF,100,lRowF+15-1).CellStyle = "Ultimus_centro"
    #~ raggruppo i righi di mirura
        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        oCellRangeAddr.StartColumn = 0
        oCellRangeAddr.EndColumn = 0
        oCellRangeAddr.StartRow = lRowF
        oCellRangeAddr.EndRow = lRowF+15-1
        oSheet.group(oCellRangeAddr, 1)
        
#~ INSERISCI LA DATA E IL PROGETTISTA
        oSheet.getCellByPosition(1 , riga_corrente+3).Formula = '=CONCATENATE("Data, ";TEXT(NOW();"DD/MM/YYYY"))'
    #~ consolido il risultato
        oRange = oSheet.getCellByPosition(1 , riga_corrente+3)
        flags =(oDoc.createInstance('com.sun.star.sheet.CellFlags.FORMULA'))
        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)
        oSheet.getCellRangeByPosition(1,riga_corrente+3,1,riga_corrente+3).CellStyle = 'ULTIMUS'
        oSheet.getCellByPosition(1 , riga_corrente+5).Formula = 'Il progettista'
        oSheet.getCellByPosition(1 , riga_corrente+6).Formula = '=CONCATENATE($S2.$C$13)'

    if oSheet.Name in('COMPUTO', 'VARIANTE', 'CompuM_NoP'):
        oDoc.CurrentController.ZoomValue = 400

        attesa().start()
        lRowF = ultima_voce(oSheet)+2

        oDoc.CurrentController.setFirstVisibleRow(lRowF-2)
        lRowE = getLastUsedCell(oSheet).EndRow
        for i in range(lRowF, getLastUsedCell(oSheet).EndRow+1):
            if oSheet.getCellByPosition(0, i).CellStyle == "Riga_rossa_Chiudi":
                lRowE = i
                break
        if lRowE > lRowF+1:
            oSheet.getRows().removeByIndex(lRowF, lRowE-lRowF)
        riga_corrente = lRowF+2
        if oDoc.getSheets().hasByName('S2') == True:
            ii = 11
            vv = 18
            ac = 28
            ad = 29
            ae = 30
            ss = 41
            col ='S'
        else:
            ii = 8
            vv = 9
            ss = 9
            col ='J'
        #~ mri(oSheet.getCellByPosition(lRowF,0).Rows)
        oSheet.getRows().insertByIndex(lRowF, 17)
        oSheet.getCellRangeByPosition(0, lRowF, ss, lRowF+17-1).CellStyle = 'ULTIMUS'
        # raggruppo i righi di mirura
        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        oCellRangeAddr.StartColumn = 0
        oCellRangeAddr.EndColumn = 0
        oCellRangeAddr.StartRow = lRowF
        oCellRangeAddr.EndRow = lRowF+17-1
        oSheet.group(oCellRangeAddr, 1)

    #~ INSERIMENTO TITOLO
        oSheet.getCellByPosition(2 , riga_corrente).String = 'Riepilogo strutturale delle Categorie'
        oSheet.getCellByPosition(ii , riga_corrente).String = 'Incidenze %'
        oSheet.getCellByPosition(vv , riga_corrente).String = 'Importi €'
        oSheet.getCellByPosition(ac , riga_corrente).String = 'Materiali\ne Noli €'
        oSheet.getCellByPosition(ad , riga_corrente).String = 'Incidenza\nMDO %'
        oSheet.getCellByPosition(ae , riga_corrente).String = 'Importo\nMDO €'
        inizio_gruppo = riga_corrente
        riga_corrente += 1
        for i in range(0, lRowF):
            if oSheet.getCellByPosition(1 , i).CellStyle == 'Livello-0-scritta':
                #~ chi(riga_corrente)
                oSheet.getRows().insertByIndex(riga_corrente,1)
                oSheet.getCellByPosition(1 , riga_corrente).Formula = '=B' + str(i+1) 
                oSheet.getCellByPosition(1 , riga_corrente).CellStyle = 'Ultimus_destra'
                oSheet.getCellByPosition(2 , riga_corrente).Formula = '=C' + str(i+1)
                #~ chi(formulaSCat)
                oSheet.getCellByPosition(ii , riga_corrente).Formula = '=' + col + str(riga_corrente+1) + '/' + col + str(lRowF) + '*100'
                oSheet.getCellByPosition(ii, riga_corrente).CellStyle = 'Ultimus %'
                oSheet.getCellByPosition(vv , riga_corrente).Formula = '='+ col + str(i+1) 
                oSheet.getCellRangeByPosition(vv , riga_corrente, ae , riga_corrente).CellStyle = 'Ultimus_totali'
                oSheet.getCellByPosition(ac , riga_corrente).Formula = '=AC'+ str(i+1)
                oSheet.getCellByPosition(ad , riga_corrente).Formula = '=AD'+ str(i+1) + '*100'
                oSheet.getCellByPosition(ad, riga_corrente).CellStyle = 'Ultimus %'
                oSheet.getCellByPosition(ae , riga_corrente).Formula = '=AE'+ str(i+1)
                riga_corrente += 1
            elif oSheet.getCellByPosition(1 , i).CellStyle == 'Livello-1-scritta':
                #~ chi(riga_corrente)
                oSheet.getRows().insertByIndex(riga_corrente,1)
                oSheet.getCellByPosition(1 , riga_corrente).Formula = '=B' + str(i+1) 
                oSheet.getCellByPosition(1 , riga_corrente).CellStyle = 'Ultimus_destra'
                oSheet.getCellByPosition(2 , riga_corrente).Formula = '=CONCATENATE("   ";C' + str(i+1) + ')'
                #~ chi(formulaSCat)
                oSheet.getCellByPosition(ii , riga_corrente).Formula = '=' + col + str(riga_corrente+1) + '/' + col + str(lRowF) + '*100'
                oSheet.getCellByPosition(ii, riga_corrente).CellStyle = 'Ultimus %'
                oSheet.getCellByPosition(vv , riga_corrente).Formula = '='+ col + str(i+1) 
                oSheet.getCellByPosition(vv , riga_corrente).CellStyle = 'Ultimus_bordo'
                oSheet.getCellByPosition(ac , riga_corrente).Formula = '=AC'+ str(i+1)
                oSheet.getCellByPosition(ad , riga_corrente).Formula = '=AD'+ str(i+1) + '*100'
                oSheet.getCellByPosition(ad, riga_corrente).CellStyle = 'Ultimus %'
                oSheet.getCellByPosition(ae , riga_corrente).Formula = '=AE'+ str(i+1)
                riga_corrente += 1
            elif oSheet.getCellByPosition(1 , i).CellStyle == 'livello2 valuta':
                #~ chi(riga_corrente)
                oSheet.getRows().insertByIndex(riga_corrente,1)
                oSheet.getCellByPosition(1 , riga_corrente).Formula = '=B' + str(i+1) 
                oSheet.getCellByPosition(1 , riga_corrente).CellStyle = 'Ultimus_destra'
                oSheet.getCellByPosition(2 , riga_corrente).Formula = '=CONCATENATE("      ";C' + str(i+1) + ')'
                #~ chi(formulaSCat)
                oSheet.getCellByPosition(ii , riga_corrente).Formula = '=' + col + str(riga_corrente+1) + '/' + col + str(lRowF) + '*100'
                oSheet.getCellByPosition(ii, riga_corrente).CellStyle = 'Ultimus %'
                oSheet.getCellByPosition(vv , riga_corrente).Formula = '='+ col + str(i+1) 
                oSheet.getCellByPosition(vv , riga_corrente).CellStyle = 'ULTIMUS'
                oSheet.getCellByPosition(ac , riga_corrente).Formula = '=AC'+ str(i+1)
                oSheet.getCellByPosition(ad , riga_corrente).Formula = '=AD'+ str(i+1) + '*100'
                oSheet.getCellByPosition(ad, riga_corrente).CellStyle = 'Ultimus %'
                oSheet.getCellByPosition(ae , riga_corrente).Formula = '=AE'+ str(i+1)
                riga_corrente += 1
        #~ riga_corrente +=1
     
        oSheet.getCellRangeByPosition(2,inizio_gruppo,ae,inizio_gruppo).CellStyle = "Ultimus_centro"

        oSheet.getCellByPosition(2 , riga_corrente).String= 'T O T A L E   €'
        oSheet.getCellByPosition(2 , riga_corrente).CellStyle = 'Ultimus_destra'
        oSheet.getCellByPosition(vv , riga_corrente).Formula = '=' + col + str(lRowF) 
        oSheet.getCellByPosition(vv , riga_corrente).CellStyle = 'Ultimus_Bordo_sotto'
        oSheet.getCellByPosition(ac , riga_corrente).Formula = '=AC' + str(lRowF)
        oSheet.getCellByPosition(ac , riga_corrente).CellStyle = 'Ultimus_Bordo_sotto'
        oSheet.getCellByPosition(ae , riga_corrente).Formula = '=AE' + str(lRowF)
        oSheet.getCellByPosition(ae , riga_corrente).CellStyle = 'Ultimus_Bordo_sotto'
        oSheet.getCellByPosition(ad , riga_corrente).Formula = '=AD' + str(lRowF) + '*100'
        fine_gruppo = riga_corrente
    #~ DATA
        oSheet.getCellByPosition(2 , riga_corrente+3).Formula = '=CONCATENATE("Data, ";TEXT(NOW();"DD/MM/YYYY"))'
    #~ consolido il risultato
        oRange = oSheet.getCellByPosition(2 , riga_corrente+3)
        flags =(oDoc.createInstance('com.sun.star.sheet.CellFlags.FORMULA'))
        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)
        
        oSheet.getCellByPosition(2 , riga_corrente+5).Formula = 'Il Progettista'
        oSheet.getCellByPosition(2 , riga_corrente+6).Formula = '=CONCATENATE($S2.$C$13)'
        oSheet.getCellRangeByPosition(2 , riga_corrente+5, 2 , riga_corrente+6).CellStyle = 'Ultimus_centro'

        ###  inserisco il salto pagina in cima al riepilogo
        oDoc.CurrentController.select(oSheet.getCellByPosition(0, lRowF))
        ctx = XSCRIPTCONTEXT.getComponentContext()
        desktop = XSCRIPTCONTEXT.getDesktop()
        oFrame = desktop.getCurrentFrame()

        dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
        dispatchHelper.executeDispatch(oFrame, ".uno:InsertRowBreak", "", 0, list())
        oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
        ###
        #~ oSheet.getCellByPosition(lRowF,0).Rows.IsManualPageBreak = True

    oDialogo_attesa.endExecute()
    oDoc.CurrentController.ZoomValue = 100
########################################################################
def next_voice(lrow, n=1):
    '''
    lrow { double }   : riga di riferimento
    n    { integer }  : se 0 sposta prima della voce corrente
                        se 1 sposta dopo della voce corrente
    sposta il cursore prima o dopo la voce corrente restituento un idrow
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ n =0
    #~ lrow = Range2Cell()[1]
    fine = ultima_voce(oSheet)+1
    if lrow >= fine:
        return lrow

    if oSheet.getCellByPosition(0, lrow).CellStyle in stili_computo + stili_contab:
        if n==0:
            sopra = Circoscrive_Voce_Computo_Att(lrow).RangeAddress.StartRow
            lrow = sopra
        elif n==1:
            sotto = Circoscrive_Voce_Computo_Att(lrow).RangeAddress.EndRow
            lrow = sotto+1
    elif oSheet.getCellByPosition(0, lrow).CellStyle in ('Ultimus_centro_bordi_lati',):
        for y in range(lrow, getLastUsedCell(oSheet).EndRow+1):
            if oSheet.getCellByPosition(0, y).CellStyle != 'Ultimus_centro_bordi_lati':
                lrow = y
                break
    elif oSheet.getCellByPosition(0, lrow).CellStyle in noVoce:
        lrow +=1
    else:
        return
    return lrow
########################################################################
def cancella_analisi_da_ep(arg=None):
    '''
    cancella le voci in Elenco Prezzi che derivano da analisi
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
    lista_an = list()
    for i in range(0, getLastUsedCell(oSheet).EndRow):
        if oSheet.getCellByPosition(0, i).CellStyle == 'An-1_sigla':
            codice = oSheet.getCellByPosition(0, i).String
            lista_an.append(oSheet.getCellByPosition(0, i).String)
    oSheet = oDoc.Sheets.getByName('Elenco Prezzi')
    for i in reversed(range(0, getLastUsedCell(oSheet).EndRow)):
        if oSheet.getCellByPosition(0, i).String in lista_an:
            oSheet.getRows().removeByIndex(i, 1)
###
def analisi_in_ElencoPrezzi(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    try:
        oSheet = oDoc.CurrentController.ActiveSheet
        if oSheet.Name != 'Analisi di Prezzo':
            return
        oDoc.enableAutomaticCalculation(False) # blocco il calcolo automatico
        sStRange = Circoscrive_Analisi(Range2Cell()[1])
        riga = sStRange.RangeAddress.StartRow + 1
        
        codice = oSheet.getCellByPosition(0, riga).String
        
        oSheet = oDoc.Sheets.getByName('Elenco Prezzi')
        oDoc.CurrentController.setActiveSheet(oSheet)
        
        oSheet.getRows().insertByIndex(3,1)

        oSheet.getCellByPosition(0,3).CellStyle = 'EP-Cs'
        oSheet.getCellByPosition(1,3).CellStyle = 'EP-C'
        oSheet.getCellRangeByPosition(2,3,8,3).CellStyle = 'EP-C mezzo'
        oSheet.getCellByPosition(5,3).CellStyle = 'EP-C mezzo %'
        oSheet.getCellByPosition(9,3).CellStyle = 'EP-sfondo'
        oSheet.getCellByPosition(10,3).CellStyle = 'Default'
        oSheet.getCellByPosition(11,3).CellStyle = 'EP-mezzo %'
        oSheet.getCellByPosition(12,3).CellStyle = 'EP statistiche_q'
        oSheet.getCellByPosition(13,3).CellStyle = 'EP statistiche_Contab_q'

        oSheet.getCellByPosition(0,3).String = codice

        oSheet.getCellByPosition(1,3).Formula = "=$'Analisi di Prezzo'.B" + str(riga+1)
        oSheet.getCellByPosition(2,3).Formula = "=$'Analisi di Prezzo'.C" + str(riga+1)
        oSheet.getCellByPosition(3,3).Formula = "=$'Analisi di Prezzo'.K" + str(riga+1)
        oSheet.getCellByPosition(4,3).Formula = "=$'Analisi di Prezzo'.G" + str(riga+1)
        oSheet.getCellByPosition(5,3).Formula = "=$'Analisi di Prezzo'.I" + str(riga+1)
        oSheet.getCellByPosition(6,3).Formula = "=$'Analisi di Prezzo'.J" + str(riga+1)
        oSheet.getCellByPosition(7,3).Formula = "=$'Analisi di Prezzo'.A" + str(riga+1)
        oSheet.getCellByPosition(8,3).String = "(AP)"
        oSheet.getCellByPosition(11,3).Formula = "=N4/$N$2"
        oSheet.getCellByPosition(12,3).Formula = "=SUMIF(AA;A4;BB)"
        oSheet.getCellByPosition(13,3).Formula = "=SUMIF(AA;A4;cEuro)"
        oDoc.enableAutomaticCalculation(True)  # sblocco il calcolo automatico
        _gotoCella(1, 3)
    except:
        oDoc.enableAutomaticCalculation(True)
########################################################################
def tante_analisi_in_ep(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    lista_analisi = list()
    oSheet = oDoc.getSheets().getByName('Analisi di prezzo')
    voce = list()
    idx = 4
    for n in range(0, ultima_voce(oSheet)+1):
        if oSheet.getCellByPosition(0, n).CellStyle == 'An-1_sigla' and oSheet.getCellByPosition(1, n).String != '<<<Scrivi la descrizione della nuova voce da analizzare   ':
            voce =(oSheet.getCellByPosition(0, n).String,
                "=$'Analisi di Prezzo'.B" + str(n+1),
                "=$'Analisi di Prezzo'.C" + str(n+1),
                "=$'Analisi di Prezzo'.K" + str(n+1),
                "=$'Analisi di Prezzo'.G" + str(n+1),
                "=$'Analisi di Prezzo'.I" + str(n+1),
                "=F"+ str(idx)+"*E"+ str(idx),
                "=$'Analisi di Prezzo'.A" + str(n+1),
            )
            lista_analisi.append(voce)
            idx += 1
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    if len(lista_analisi) !=0:
        oSheet.getRows().insertByIndex(3,len(lista_analisi))
    else:
        return
    oRange = oSheet.getCellRangeByPosition(0, 3, 7, 3+len(lista_analisi)-1)
    lista_come_array = tuple(lista_analisi)
    
    oSheet.getCellRangeByPosition(11, 3, 11, 3+len(lista_analisi)-1).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition(12, 3, 12, 3+len(lista_analisi)-1).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition(13, 3, 13, 3+len(lista_analisi)-1).CellStyle = 'EP statistiche_Contab_q'
    
    oRange.setDataArray(lista_come_array) #setFrmulaArray() sarebbe meglio, ma mi fa storie sul codice articolo
    for y in range(3, 3+len(lista_analisi)):
        for x in range(1, 8): #evito il codice articolo, altrimenti me lo converte in numero
            oSheet.getCellByPosition(x, y).Formula = oSheet.getCellByPosition(x, y).String
    oSheet.getCellRangeByPosition(0, 3, 7, 3+len(lista_analisi)-1).CellStyle = 'EP-C mezzo'
    oSheet.getCellRangeByPosition(0, 3, 0, 3+len(lista_analisi)-1).CellStyle = 'EP-Cs'
    oSheet.getCellRangeByPosition(1, 3, 1, 3+len(lista_analisi)-1).CellStyle = 'EP-C'
    oSheet.getCellRangeByPosition(5, 3, 5, 3+len(lista_analisi)-1).CellStyle = 'EP-C mezzo %'
    #~ MsgBox('Trasferite ' + str(len(lista_analisi)) + ' analisi di prezzo in Elenco Prezzi.', 'Avviso')
########################################################################
def Circoscrive_Analisi(lrow):
    '''
    lrow    { double }  : riga di riferimento per
                        la selezione dell'intera voce
    Circoscrive una voce di analisi
    partendo dalla posizione corrente del cursore
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ lrow = Range2Cell()[1]
    #~ return
    if oSheet.getCellByPosition(0, lrow).CellStyle in stili_analisi:
        if oSheet.getCellByPosition(0, lrow).CellStyle == stili_analisi[0]:
            lrowS=lrow
        else:
            while oSheet.getCellByPosition(0, lrow).CellStyle != stili_analisi[0]:
                lrow = lrow-1
            lrowS=lrow
        lrow = lrowS
        while oSheet.getCellByPosition(0, lrow).CellStyle != stili_analisi[-1]:
            lrow=lrow+1
        lrowE=lrow+1
    celle=oSheet.getCellRangeByPosition(0,lrowS,250,lrowE)
    return celle
def Circoscrive_Voce_Computo_Att(lrow):
    '''
    lrow    { double }  : riga di riferimento per
                        la selezione dell'intera voce

    Circoscrive una voce di computo, variante o contabilità
    partendo dalla posizione corrente del cursore
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ lrow = Range2Cell()[1]
    #~ if oSheet.Name in('VARIANTE', 'COMPUTO','CONTABILITA'):
    if oSheet.getCellByPosition(0, lrow).CellStyle in('comp progress', 'comp 10 s', 'Comp Start Attributo', 'Comp End Attributo', 'Comp Start Attributo_R', 'comp 10 s_R', 'Comp End Attributo_R', 'Livello-1-scritta', 'livello2 valuta'):
        while oSheet.getCellByPosition(0, lrow).CellStyle not in('Comp End Attributo', 'Comp End Attributo_R'):
            lrow +=1
        lrowE=lrow
        while oSheet.getCellByPosition(0, lrow).CellStyle not in('Comp Start Attributo', 'Comp Start Attributo_R'):
            lrow -=1
        lrowS=lrow
    celle=oSheet.getCellRangeByPosition(0,lrowS,250,lrowE)
    return celle
########################################################################
def ColumnNumberToName(oSheet,cColumnNumb):
    '''Trasforma IDcolonna in Nome'''
    #~ oDoc = XSCRIPTCONTEXT.getDocument()
    #~ oSheet = oDoc.CurrentController.ActiveSheet
    oColumns = oSheet.getColumns()
    oColumn = oColumns.getByIndex(cColumnNumb).Name
    return oColumn
########################################################################
def ColumnNameToNumber(oSheet,cColumnName):
    '''Trasforma il nome colonna in IDcolonna'''
    #~ oDoc = XSCRIPTCONTEXT.getDocument()
    #~ oSheet = oDoc.CurrentController.ActiveSheet
    oColumns = oSheet.getColumns()
    oColumn = oColumns.getByName(cColumnName)
    oRangeAddress = oColumn.getRangeAddress()
    nColumn = oRangeAddress.StartColumn
    return nColumn
########################################################################
def azzera_voce(arg=None):
    '''
    Azzera la quantità di una voce e ne raggruppa le relative righe
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in('COMPUTO', 'VARIANTE'):
#########################################
        try:
            sRow = oDoc.getCurrentSelection().getRangeAddresses()[0].StartRow
            eRow = oDoc.getCurrentSelection().getRangeAddresses()[0].EndRow

        except:
            sRow = oDoc.getCurrentSelection().getRangeAddress().StartRow
            eRow = oDoc.getCurrentSelection().getRangeAddress().EndRow
        sStRange = Circoscrive_Voce_Computo_Att(sRow)
        sStRange.RangeAddress
        sRow = sStRange.RangeAddress.StartRow
        sStRange = Circoscrive_Voce_Computo_Att(eRow)
        sStRange.RangeAddress
        inizio = sStRange.RangeAddress.StartRow
        eRow = sStRange.RangeAddress.EndRow+1
        
        lrow = sRow
        fini = list()
        for x in range(sRow, eRow):
            if oSheet.getCellByPosition(0, x).CellStyle in('Comp End Attributo', 'Comp End Attributo_R'):
                fini.append(x)
    idx = 0
    for lrow in fini:
        lrow += idx
        try:
            sStRange = Circoscrive_Voce_Computo_Att(lrow)
            sStRange.RangeAddress
            inizio = sStRange.RangeAddress.StartRow
            fine = sStRange.RangeAddress.EndRow

            _gotoCella(2, fine-1)
            if oSheet.getCellByPosition(2, fine-1).String == '*** VOCE AZZERATA ***':
                ### elimino il colore di sfondo
                oSheet.getCellRangeByPosition(0, inizio, 250, fine).clearContents(HARDATTR)
                raggruppa_righe_voce(lrow, 0)
                oSheet.getRows().removeByIndex(fine-1, 1)
                fine -=1
                _gotoCella(2, fine-1)
                idx -= 1
            else:
                Copia_riga_Ent()
                oSheet.getCellByPosition(2, fine).String = '*** VOCE AZZERATA ***'
                oSheet.getCellByPosition(5, fine).Formula = '=-SUBTOTAL(9;J' + str(inizio+1) + ':J' + str(fine) + ')'
                ### cambio il colore di sfondo
                oDoc.CurrentController.select(sStRange)
                raggruppa_righe_voce (lrow, 1)
                ctx = XSCRIPTCONTEXT.getComponentContext()
                desktop = XSCRIPTCONTEXT.getDesktop()
                oFrame = desktop.getCurrentFrame()
                dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
                oProp = PropertyValue()
                oProp.Name = 'BackgroundColor'
                oProp.Value = 15066597
                properties =(oProp,)
                dispatchHelper.executeDispatch(oFrame, '.uno:BackgroundColor', '', 0, properties)
                _gotoCella(2, fine)
                ###   
            lrow = Range2Cell()[1]
            lrow = next_voice(lrow, 1)
        except:
            pass
    return
########################################################################
def elimina_voci_azzerate(arg=None):
    '''
    Elimina le voci in cui compare la dicitura '*** VOCE AZZERATA ***'
    in COMPUTO o in VARIANTE, senza chiedere conferma
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        if oSheet.Name in('COMPUTO', 'VARIANTE'):
            ER = getLastUsedCell(oSheet).EndRow
            for lrow in reversed(range(0, ER)):
                if oSheet.getCellByPosition(2, lrow).String == '*** VOCE AZZERATA ***':
                    elimina_voce(lRow=lrow, msg=0)
            numera_voci(1)
    except:
        return
########################################################################
def raggruppa_righe_voce (lrow, flag=1):
    '''
    Raggruppa le righe che compongono una singola voce.
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ lrow = Range2Cell()[1]
    if oSheet.Name in('COMPUTO', 'VARIANTE'):
        sStRange = Circoscrive_Voce_Computo_Att (lrow)
        sStRange.RangeAddress

        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        oCellRangeAddr.StartColumn = sStRange.RangeAddress.StartColumn
        oCellRangeAddr.EndColumn = sStRange.RangeAddress.EndColumn
        oCellRangeAddr.StartRow = sStRange.RangeAddress.StartRow
        oCellRangeAddr.EndRow = sStRange.RangeAddress.EndRow
        if flag == 1:
            oSheet.group(oCellRangeAddr, 1)
        else:
            oSheet.ungroup(oCellRangeAddr, 1)
########################################################################
#~ def debug(arg=None):
def nasconde_voci_azzerate(arg=None):
    '''
    Nasconde le voci in cui compare la dicitura '*** VOCE AZZERATA ***'
    in COMPUTO o in VARIANTE.
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        if oSheet.Name in('COMPUTO', 'VARIANTE'):
            oSheet.clearOutline()
            ER = getLastUsedCell(oSheet).EndRow
            for lrow in reversed(range(0, ER)):
                if oSheet.getCellByPosition(2, lrow).String == '*** VOCE AZZERATA ***':
                    raggruppa_righe_voce(lrow, 1)
    except:
        return
########################################################################
def elimina_voce(arg=None, lRow=None, msg=1):
    '''
    lRow { long }  : numero riga
    msg  { bit }   : 1 chiedi conferma con messaggio
                     0 egegui senza conferma
    Elimina una voce in COMPUTO, VARIANTE, CONTABILITA o Analisi di Prezzo
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if lRow == None:
        lRow = Range2Cell()[1]
    try:
        if oSheet.Name in('COMPUTO', 'VARIANTE', 'CONTABILITA'):
            sStRange = Circoscrive_Voce_Computo_Att(lRow)
        elif oSheet.Name == 'Analisi di Prezzo':
            sStRange = Circoscrive_Analisi(lRow)
    except:
        return
    sStRange.RangeAddress
    SR = sStRange.RangeAddress.StartRow
    ER = sStRange.RangeAddress.EndRow
    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, SR, 250, ER))
    if msg==1:
        if DlgSiNo("""OPERAZIONE NON ANNULLABILE!
        
Stai per eliminare la voce selezionata.
Vuoi Procedere?
 """,'AVVISO!') ==2:
            oSheet.getRows().removeByIndex(SR, ER-SR+1)
            numera_voci(0)
        else:
            return
    elif msg==0:
        oSheet.getRows().removeByIndex(SR, ER-SR+1)
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
########################################################################
def copia_riga_computo(lrow):
    '''
    Inserisce una nuova riga di misurazione nel computo
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ lrow = Range2Cell()[1]
    stile = oSheet.getCellByPosition(1, lrow).CellStyle
    if stile in('comp Art-EP', 'comp Art-EP_R', 'Comp-Bianche in mezzo'):#'Comp-Bianche in mezzo Descr', 'comp 1-a', 'comp sotto centro'):# <stili computo
        sStRange = Circoscrive_Voce_Computo_Att(lrow)
        sStRange.RangeAddress
        sopra = sStRange.RangeAddress.StartRow
        sotto = sStRange.RangeAddress.EndRow
        lrow = lrow+1 # PER INSERIMENTO SOTTO RIGA CORRENTE
        oSheet.getRows().insertByIndex(lrow,1)
# imposto gli stili
        oSheet.getCellRangeByPosition(5, lrow, 7, lrow,).CellStyle = 'comp 1-a'
        oSheet.getCellByPosition(0, lrow).CellStyle = 'comp 10 s'
        oSheet.getCellByPosition(1, lrow).CellStyle = 'Comp-Bianche in mezzo'
        oSheet.getCellByPosition(2, lrow).CellStyle = 'comp 1-a'
        oSheet.getCellRangeByPosition(3, lrow, 4, lrow).CellStyle = 'Comp-Bianche in mezzo bordate_R'
        oSheet.getCellByPosition(8, lrow).CellStyle = 'comp 1-a peso'
        oSheet.getCellByPosition(9, lrow).CellStyle = 'Blu'
# ci metto le formule
        oSheet.getCellByPosition(9, lrow).Formula = '=IF(PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + ')=0;"";PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + '))'
        oSheet.getCellByPosition(10 , lrow).Formula = ""
        #~ _gotoCella(2, lrow)
        oDoc.CurrentController.select(oSheet.getCellByPosition(2, lrow))
        oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
def copia_riga_contab(lrow):
    '''
    Inserisce una nuova riga di misurazione in contabilità
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ lrow = Range2Cell()[1]
    stile = oSheet.getCellByPosition(1, lrow).CellStyle
    if  oSheet.getCellByPosition(1, lrow+1).CellStyle == 'comp sotto Bianche_R':
        return
    if stile in('comp Art-EP_R', 'Data_bianca', 'Comp-Bianche in mezzo_R'):
        sStRange = Circoscrive_Voce_Computo_Att(lrow)
        sStRange.RangeAddress
        sopra = sStRange.RangeAddress.StartRow
        sotto = sStRange.RangeAddress.EndRow
        lrow = lrow+1 # PER INSERIMENTO SOTTO RIGA CORRENTE
        #~ if  oSheet.getCellByPosition(2, lrow).CellStyle == 'comp sotto centro_R':
            #~ lrow = lrow-1
        oSheet.getRows().insertByIndex(lrow,1)
    # imposto gli stili
        oSheet.getCellByPosition(1, lrow).CellStyle = 'Comp-Bianche in mezzo_R'
        oSheet.getCellByPosition(2, lrow).CellStyle = 'comp 1-a'
        oSheet.getCellRangeByPosition(5, lrow, 7, lrow).CellStyle = 'comp 1-a'
        oSheet.getCellRangeByPosition(11, lrow, 23, lrow).CellStyle = 'Comp-Bianche in mezzo_R'
        oSheet.getCellByPosition(8, lrow).CellStyle = 'comp 1-a peso'
        oSheet.getCellRangeByPosition(9, lrow, 11, lrow).CellStyle = 'Comp-Variante'
    # ci metto le formule
        oSheet.getCellByPosition(9, lrow).Formula = '=IF(PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + ')<=0;"";PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + '))'
        oSheet.getCellByPosition(11, lrow).Formula = '=IF(PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + ')>=0;"";PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + ')*-1)'
    # preserva la data di misura
        if oSheet.getCellByPosition(1, lrow+1).CellStyle == 'Data_bianca':
            oRangeAddress = oSheet.getCellByPosition(1, lrow+1).getRangeAddress()
            oCellAddress = oSheet.getCellByPosition(1,lrow).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
            oSheet.getCellByPosition(1, lrow+1).String = ""
            oSheet.getCellByPosition(1, lrow+1).CellStyle = 'Comp-Bianche in mezzo_R'
        oDoc.CurrentController.select(oSheet.getCellByPosition(2, lrow))
def copia_riga_analisi(lrow):
    '''
    Inserisce una nuova riga di misurazione in analisi di prezzo
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    stile = oSheet.getCellByPosition(0, lrow).CellStyle
    if stile in('An-lavoraz-desc', 'An-lavoraz-Cod-sx'):
        lrow=lrow+1
        oSheet.getRows().insertByIndex(lrow,1)
    # imposto gli stili
        oSheet.getCellByPosition(0, lrow).CellStyle = 'An-lavoraz-Cod-sx'
        oSheet.getCellRangeByPosition(1, lrow, 5, lrow).CellStyle = 'An-lavoraz-generica'
        oSheet.getCellByPosition(3, lrow).CellStyle = 'An-lavoraz-input'
        oSheet.getCellByPosition(6, lrow).CellStyle = 'An-senza'
        oSheet.getCellByPosition(7, lrow).CellStyle = 'An-senza-DX'
    # ci metto le formule
        #~ oDoc.enableAutomaticCalculation(False)
        oSheet.getCellByPosition(1, lrow).Formula = '=IF(A' + str(lrow+1) + '="";"";CONCATENATE("  ";VLOOKUP(A' + str(lrow+1) + ';elenco_prezzi;2;FALSE());' '))'
        oSheet.getCellByPosition(2, lrow).Formula = '=IF(A' + str(lrow+1) + '="";"";VLOOKUP(A' + str(lrow+1) + ';elenco_prezzi;3;FALSE()))'
        oSheet.getCellByPosition(3, lrow).Value = 0
        oSheet.getCellByPosition(4, lrow).Formula = '=IF(A' + str(lrow+1) + '="";0;VLOOKUP(A' + str(lrow+1) + ';elenco_prezzi;5;FALSE()))'
        oSheet.getCellByPosition(5, lrow).Formula = '=D' + str(lrow+1) + '*E' + str(lrow+1)
        oSheet.getCellByPosition(8, lrow).Formula = '=IF(A' + str(lrow+1) + '="";"";IF(VLOOKUP(A' + str(lrow+1) + ';elenco_prezzi;6;FALSE())="";"";(VLOOKUP(A' + str(lrow+1) + ';elenco_prezzi;6;FALSE()))))'
        oSheet.getCellByPosition(9, lrow).Formula = '=IF(I' + str(lrow+1) + '="";"";I' + str(lrow+1) + '*F' + str(lrow+1) + ')'
        #~ oDoc.enableAutomaticCalculation(True)
    # preserva il Pesca
        if oSheet.getCellByPosition(1, lrow-1).CellStyle == 'An-lavoraz-dx-senza-bordi':
            oRangeAddress = oSheet.getCellByPosition(0, lrow+1).getRangeAddress()
            oCellAddress = oSheet.getCellByPosition(0,lrow).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
        oSheet.getCellByPosition(0, lrow).String = 'Cod. Art.?'
    oDoc.CurrentController.select(oSheet.getCellByPosition(1, lrow))
########################################################################
def Copia_riga_Ent(arg=None): #Aggiungi Componente - capisce su quale tipologia di tabelle è
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    nome_sheet = oSheet.Name
    if nome_sheet in('COMPUTO', 'VARIANTE'):
        copia_riga_computo(lrow)
    elif nome_sheet == 'CONTABILITA':
        copia_riga_contab(lrow)
    elif nome_sheet == 'Analisi di Prezzo':
        copia_riga_analisi(lrow)
########################################################################
def cerca_partenza(arg=None):
    '''
    Conserva, nella variabile globale 'partenza', il nome del foglio [0] e l'id
    della riga di codice prezzo componente [1], il flag '#reg' solo per la contbailità.
    partenza = (nome_foglio, id_rcodice, flag_contabilità)
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    global partenza
    if oSheet.getCellByPosition(0, lrow).CellStyle in stili_computo: #COMPUTO, VARIANTE
        sStRange = Circoscrive_Voce_Computo_Att(lrow)
        partenza =(oSheet.Name, sStRange.RangeAddress.StartRow+1)
    elif oSheet.getCellByPosition(0, lrow).CellStyle in stili_contab: #CONTABILITA
        sStRange = Circoscrive_Voce_Computo_Att(lrow)
        partenza =(oSheet.Name, sStRange.RangeAddress.StartRow+1, oSheet.getCellByPosition(22, sStRange.RangeAddress.StartRow+1).String)
    elif oSheet.getCellByPosition(0, lrow).CellStyle in('An-lavoraz-Cod-sx'): #ANALISI
        partenza =(oSheet.Name, lrow)
    return partenza
########################################################################
sblocca_computo = 0
def pesca_cod(arg=None):
    '''
    Permette di scegliere il codice per la voce di COMPUTO o VARIANTE o CONTABILITA dall'Elenco Prezzi.
    Capisce quando la voce nel libretto delle misure è già registrata o nel documento ci sono già atti contabili emessi.
    '''
    global sblocca_computo
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    if oSheet.getCellByPosition(0, lrow).CellStyle not in stili_computo + stili_contab + stili_analisi + stili_elenco:
        return
    if oSheet.Name in('Analisi di Prezzo'):
        cerca_partenza()
        _gotoSheet('Elenco Prezzi')
    if oSheet.Name in('CONTABILITA'):
        cerca_partenza()
        if oSheet.getCellByPosition(1, partenza[1]).String != 'Cod. Art.?':
            basic_LeenO('Cerca_Rior.cerca_in_elenco')
            return
        try:
            if partenza[2] == '#reg':
                if DlgSiNo("""Cambiando il Codice Articolo di questa voce, comprometterai
la validità degli atti contabili già emessi.

VUOI PROCEDERE?

Scegliendo Sì sarai costretto a rigenerarli!""", 'Voce già registrata!') ==3:
                    return
                else:
                    _gotoSheet('Elenco Prezzi')
            else:
                _gotoSheet('Elenco Prezzi')
            partenza[2]
        except TypeError:
            return
    if oSheet.Name in('COMPUTO', 'VARIANTE'):
        if oDoc.NamedRanges.hasByName("#Lib#1") == True:
            if sblocca_computo == 0:
                if DlgSiNo("Risulta già registrato un SAL. VUOI PROCEDERE COMUQUE?",'ATTENZIONE!') ==3:
                    return
                else:
                    sblocca_computo = 1
        cerca_partenza()
        if oSheet.getCellByPosition(1, partenza[1]).String != 'Cod. Art.?':
            basic_LeenO('Cerca_Rior.cerca_in_elenco')
            return
        _gotoSheet('Elenco Prezzi')
    if oSheet.Name in('Elenco Prezzi'):
        try:
            lrow = Range2Cell()[1]
            codice = oSheet.getCellByPosition(0, lrow).String
            _gotoSheet(partenza[0])
            oSheet = oDoc.CurrentController.ActiveSheet
            if partenza[0] == 'Analisi di Prezzo':
                oSheet.getCellByPosition(0, partenza[1]).String = codice
                _gotoCella(3, partenza[1])
            else:
                oSheet.getCellByPosition(1, partenza[1]).String = codice
                _gotoCella(2, partenza[1]+1)
        except NameError:
            return
########################################################################
def ricicla_misure(arg=None):
#~ def debug(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name != 'CONTABILITA': return
    lrow = Range2Cell()[1]
    #~ chi(lrow)
    if oSheet.getCellByPosition(0, lrow).CellStyle not in stili_contab + ('Comp TOTALI', 'Ultimus_centro_bordi_lati',):
        return
    chi(next_voice(lrow))
    #~ cerca_partenza()
    #~ chi(partenza)
    #~ lrowE = ultima_voce(oSheet)+1
    #~ chi(lrowE)
########################################################################
def inverti_segno(arg=None):
    '''
    Inverte il segno delle formule di quantità nei righi di misurazione selezionati.
    Funziona solo in COMPUTO e VARIANTE.
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in('COMPUTO', 'VARIANTE'):
        try:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
        except AttributeError:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
        SR = oRangeAddress.StartRow
        ER = oRangeAddress.EndRow
        for lrow in range(SR, ER+1):
            if oSheet.getCellByPosition(2, lrow).CellStyle == 'comp 1-a':
                if '-' in oSheet.getCellByPosition(9, lrow).Formula:
                    oSheet.getCellByPosition(9, lrow).Formula = '=IF(PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + ')=0;"";PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + '))'
                    for x in range (2, 8):
                        oSheet.getCellByPosition(x, lrow).CharColor = -1
                else:
                    oSheet.getCellByPosition(9, lrow).Formula = '=IF(PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + ')=0;"";-PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + '))'
                    for x in range (2, 8):
                        oSheet.getCellByPosition(x, lrow).CharColor = 16724787
########################################################################
def valuta_cella(oCell):
    '''
    Estrae qualsiasi valore da una cella, restituendo una strigna, indipendentemente dal tipo originario.
    oCell       { object }  : cella da validare
    '''
    if oCell.Type.value == 'FORMULA':
        try:
            eval(oCell.Formula.split('=')[-1])
            valore = oCell.Formula.split('=')[-1]
        except:
            try:
                valore = str(oSheet.getCellRangeByName(oCell.Formula.split('=')[-1]).Value)
            except:
                valore = str(oCell.Value)
    elif oCell.Type.value == 'VALUE':
        valore = str(oCell.Value)
    elif oCell.Type.value == 'TEXT':
        valore = str(oCell.String)
    elif oCell.Type.value == 'EMPTY':
        valore = ''
    #~ if valore == ' ': valore = ''
    return valore
def debug_validation(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ mri(oDoc.CurrentSelection.Validation)
    
    oSheet.getCellRangeByName('L1').String = 'Ricicla da:'
    oSheet.getCellRangeByName('L1').CellStyle = 'Reg_prog'
    oCell= oSheet.getCellRangeByName('N1')
    if oCell.String not in("COMPUTO", "VARIANTE", 'Scegli origine'):
        oCell.CellStyle = 'Menu_sfondo _input_grasBig'
        valida_cella(oCell, '"COMPUTO";"VARIANTE"',titoloInput='Scegli...', msgInput='COMPUTO o VARIANTE', err=True)
        oCell.String ='Scegli...'
    
def valida_cella(oCell, lista_val, titoloInput='', msgInput='', err= False ):
    '''
    Validità lista valori
    Imposta un elenco di valori a cascata, da cui scegliere.
    oCell       { object }  : cella da validare
    lista_val   { string }  : lista dei valori in questa forma: '"UNO";"DUE";"TRE"'
    titoloInput { string }  : titolo del suggerimento che compare passando il cursore sulla cella
    msgInput    { string }  : suggerimento che compare passando il cursore sulla cella
    err         { boolean } : permette di abilitare il messaggio di errore per input non validi
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    oTabVal = oCell.getPropertyValue("Validation")
    oTabVal.setPropertyValue('ConditionOperator', 1)

    oTabVal.setPropertyValue("ShowInputMessage", True) 
    oTabVal.setPropertyValue("InputTitle", titoloInput)
    oTabVal.setPropertyValue("InputMessage", msgInput) 
    oTabVal.setPropertyValue("ErrorMessage", "ERRORE: Questo valore non è consentito.")
    oTabVal.setPropertyValue("ShowErrorMessage", err)
    oTabVal.ErrorAlertStyle = uno.Enum("com.sun.star.sheet.ValidationAlertStyle", "STOP")
    oTabVal.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")
    oTabVal.Operator = uno.Enum("com.sun.star.sheet.ConditionOperator", "EQUAL")
    oTabVal.setFormula1(lista_val)
    oCell.setPropertyValue("Validation", oTabVal)

def debug_ConditionalFormat(arg=None):
#~ def debug(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oCell= oDoc.CurrentSelection
    oSheet = oDoc.CurrentController.ActiveSheet

    i =oCell.RangeAddress.StartRow
    n =oCell.Rows.Count
    oSheet.getRows().removeByIndex(i, n)
    #~ mri(oCell)#.ConditionalFormat)

########################################################################

def debug_tipo_di_valore(arg=None):
#~ def debug(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    chi(oSheet.getCellByPosition(Range2Cell()[0],Range2Cell()[10]).Type.value)
    #~ if oSheet.getCellByPosition(2, 5).Type.value == 'FORMULA':
        #~ MsgBox(oSheet.getCellByPosition(9, 5).Formula)
########################################################################
def debugclip(arg=None):
    import pyperclip
    #~ mri(XSCRIPTCONTEXT.getComponentContext())
    sText = 'sticazzi'
    #create SystemClipboard instance
    oClip = createUnoService("com.sun.star.datatransfer.clipboard.SystemClipboard")
    oClipContents = oClip.getContents()
    flavors = oClipContents.getTransferDataFlavors()
    mri(oClip)
    #~ for i in flavors:
        #~ aDataFlavor = flavors(i)
        #~ chi(aDataFlavor)
        
    return
    #~ createUnoService =(XSCRIPTCONTEXT.getComponentContext().getServiceManager().createInstance)
    #~ oTR = createUnoListener("Tr_", "com.sun.star.datatransfer.XTransferable")
    oClip.setContents( oTR, None )
    sTxtCString = sText
    oClip.flushClipboard()
########################################################################
def copy_clip(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()

    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    dispatchHelper.executeDispatch(oFrame, ".uno:Copy", "", 0, list())
########################################################################
def paste_clip(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()

    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    dispatchHelper.executeDispatch(oFrame, ".uno:Paste", "", 0, list())
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect

########################################################################
def copia_celle_visibili(arg=None):
    '''
    A partire dalla selezione di un range di celle in cui alcune righe e/o
    colonne sono nascoste, mette in clipboard solo il contenuto delle celle
    visibili.
    Liberamente ispirato a "Copy only visible cells" http://bit.ly/2j3bfq2
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
    IS = oRangeAddress.Sheet
    SC = oRangeAddress.StartColumn
    EC = oRangeAddress.EndColumn
    SR = oRangeAddress.StartRow
    ER = oRangeAddress.EndRow
    if EC == 1023:
        EC = getLastUsedCell(oSheet).EndColumn
    if ER == 1048575:
        ER = getLastUsedCell(oSheet).EndRow
    righe = list()
    colonne = list()
    i = 0
    for nRow in range(SR, ER+1):
        if oSheet.getCellByPosition(SR, nRow).Rows.IsVisible == False:
            righe.append(i)
        i += 1
    i = 0
    for nCol in range(SC, EC+1):
        if oSheet.getCellByPosition(nCol, nRow).Columns.IsVisible == False:
            colonne.append(i)
        i += 1

    if oDoc.getSheets().hasByName('tmp_clip') == False:
        sheet = oDoc.createInstance("com.sun.star.sheet.Spreadsheet")
        tmp = oDoc.Sheets.insertByName('tmp_clip', sheet)
    tmp = oDoc.getSheets().getByName('tmp_clip')    

    oCellAddress = tmp.getCellByPosition(0,0).getCellAddress()
    tmp.copyRange(oCellAddress, oRangeAddress)
    
    for i in reversed(righe):
        tmp.getRows().removeByIndex(i, 1)
    for i in reversed(colonne):
        tmp.getColumns().removeByIndex(i, 1)

    oRange = tmp.getCellRangeByPosition(0,0, EC-SC-len(colonne), ER-SR-len(righe))
    oDoc.CurrentController.select(oRange)

    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    dispatchHelper.executeDispatch(oFrame, ".uno:Copy", "", 0, list())
    oDoc.Sheets.removeByName('tmp_clip')
    oDoc.CurrentController.setActiveSheet(oSheet)
    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(SC, SR, EC, ER))
# Range2Cell ###########################################################
def Range2Cell():
    '''
    Restituisce la tupla(IDcolonna, IDriga) della posizione corrnete
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        if oDoc.getCurrentSelection().getRangeAddresses()[0]:
            nRow = oDoc.getCurrentSelection().getRangeAddresses()[0].StartRow
            nCol = oDoc.getCurrentSelection().getRangeAddresses()[0].StartColumn
    except AttributeError:
        nRow = oDoc.getCurrentSelection().getRangeAddress().StartRow
        nCol = oDoc.getCurrentSelection().getRangeAddress().StartColumn
    return(nCol,nRow)
########################################################################
# restituisce l'ID dell'ultima riga usata
def getLastUsedCell(oSheet):
    '''
    Restitusce l'indirizzo dell'ultima cella usata
    in forma di oggetto
    '''
    oCell = oSheet.getCellByPosition(0, 0)
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.gotoEndOfUsedArea(True)
    aAddress = oCursor.RangeAddress
    return aAddress#.EndColumn, aAddress.EndRow)
########################################################################
# numera le voci di computo o contabilità
def numera_voci(bit=1):#
    '''
    bit { integer }  : 1 rinumera tutto
                       0 rinumera dalla voce corrente in giù
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    if oDoc.getSheets().getByName('S1').getCellRangeByName('H335').Value != 0:
        _gotoSheet('S1')
        _gotoCella(7, 334)
        MsgBox('''La rinumerazione delle voci è disabilitata.
Questo dipende dalla Variabile Generale qui evidenziata.''', "Avviso!")
        return
    oSheet = oDoc.CurrentController.ActiveSheet
    lastRow = getLastUsedCell(oSheet).EndRow+1
    lrow = Range2Cell()[1]
    n = 1
    if bit==0:
        for x in reversed(range(0, lrow)):
            if oSheet.getCellByPosition(1,x).CellStyle in('comp Art-EP', 'comp Art-EP_R'):
                n = oSheet.getCellByPosition(0,x).Value +1
                break
        for row in range(lrow,lastRow):
            if oSheet.getCellByPosition(1,row).CellStyle in('comp Art-EP', 'comp Art-EP_R'):
                oSheet.getCellByPosition(0,row).Value = n
                n +=1
    if bit==1:
        for row in range(0,lastRow):
            if oSheet.getCellByPosition(1,row).CellStyle in('comp Art-EP', 'comp Art-EP_R'):
                oSheet.getCellByPosition(0,row).Value = n
                n = n+1
########################################################################
def refresh(arg=1):
    '''
    Abilita / disabilita il refresh per accelerare le procedure
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    if arg == 0:
        #~ oDoc.CurrentController.ZoomValue = 400
        oDoc.enableAutomaticCalculation(False) # blocco il calcolo automatico
        #~ oDoc.addActionLock()
        #~ oDoc.removeActionLock()
        #~ oDoc.lockControllers #disattiva l'eco a schermo
    elif arg == 1:
        #~ oDoc.CurrentController.ZoomValue = 100
        oDoc.enableAutomaticCalculation(True) # sblocco il calcolo automatico
        #~ oDoc.removeActionLock()
        #~ oDoc.unlockControllers #attiva l'eco a schermo
########################################################################
def debug(arg=None):
#~ def richiesta_offerta(arg=None):
    '''Crea la Lista Lavorazioni e Forniture dall'Elenco Prezzi,
per la formulazione dell'offerta'''
    oDoc = XSCRIPTCONTEXT.getDocument()
    _gotoSheet('Elenco Prezzi')
    #~ genera_sommario()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        oDoc.Sheets.copyByName(oSheet.Name,'Elenco Prezzi', 5)
    except:
        pass
    nSheet = oDoc.getSheets().getByIndex(5).Name
    _gotoSheet(nSheet)
    setTabColor(10079487)
    oSheet = oDoc.CurrentController.ActiveSheet
    fine = getLastUsedCell(oSheet).EndRow+1
    oRange = oSheet.getCellRangeByPosition (12,3,12, fine)
    aSaveData = oRange.getDataArray()
    oRange = oSheet.getCellRangeByPosition (3,3,3, fine)
    oRange.CellStyle = 'EP statistiche_q'
    oRange.setDataArray(aSaveData)
    
    oSheet.getCellByPosition(3 , 2).String = 'Quantità\na Computo'
    oSheet.getCellByPosition(5 , 2).String = 'Prezzo Unitario\nin lettere'
    oSheet.getCellByPosition(6 , 2).String = 'Importo'
    oSheet.Columns.removeByIndex(7, 100)
    oSheet.getColumns().getByName("D").IsVisible = True
    oSheet.getColumns().getByName("F").IsVisible = True
    oSheet.getColumns().getByName("G").IsVisible = True
    oSheet.getColumns().getByName("A").Columns.Width = 1600
    oSheet.getColumns().getByName("B").Columns.Width = 8000
    oSheet.getColumns().getByName("C").Columns.Width = 1200
    oSheet.getColumns().getByName("D").Columns.Width = 1600
    oSheet.getColumns().getByName("E").Columns.Width = 1500
    oSheet.getColumns().getByName("F").Columns.Width = 4000
    oSheet.getColumns().getByName("G").Columns.Width = 1800
    oDoc.CurrentController.freezeAtPosition(0, 1)
    
    formule = list()
    for x in range(3,getLastUsedCell(oSheet).EndRow-1):
        #~ oSheet.getCellByPosition(6,x).Formula ='=IF(E' + str(x+1) + '<>"";D' + str(x+1) + '*E' + str(x+1) + ';""'
        formule.append(['=IF(E' + str(x+1) + '<>"";D' + str(x+1) + '*E' + str(x+1) + ';""'])
    oSheet.getCellRangeByPosition (6,3,6,len(formule)+2).CellBackColor = 15757935
    oRange = oSheet.getCellRangeByPosition (6,3,6,len(formule)+2)
    formule = tuple(formule)
    oRange.setFormulaArray(formule)
    #~ return
    
    oSheet.getCellRangeByPosition(4, 3, 4, fine).clearContents(VALUE + DATETIME + STRING +
                                          ANNOTATION + FORMULA + HARDATTR +
                                          OBJECTS + EDITATTR + FORMATTED)

    oSheet.getCellRangeByPosition(0, fine-1, 100, fine+1).clearContents(VALUE + FORMULA + STRING)

    oSheet.Columns.insertByIndex(0, 1)
    oSrc = oSheet.getCellRangeByPosition(1,0,1, fine).RangeAddress
    oDest = oSheet.getCellByPosition(0,0 ).CellAddress
    oSheet.copyRange(oDest, oSrc)
    oSheet.getCellByPosition(0,2).String="N."
    for x in range(3, fine-1):
        oSheet.getCellByPosition(0,x).Value = x-2
    oSheet.getColumns().getByName("A").Columns.Width = 650

    oSheet.getCellByPosition(7,fine).Formula="=SUBTOTAL(9;H2:H"+ str(fine+1) +")"
    oSheet.getCellByPosition(2,fine).String="TOTALE COMPUTO"
    oSheet.getCellRangeByPosition(0,fine,7,fine).CellStyle="Comp TOTALI"
    oSheet.Rows.removeByIndex(fine-1, 1)
    oSheet.Rows.removeByIndex(0, 2)

    oSheet.getCellByPosition(2,fine+3).String="(diconsi euro - in lettere)"
    oSheet.getCellRangeByPosition (2,fine+3,6,fine+3).CellStyle="List-intest_med_c"

    oSheet.getCellByPosition(2,fine+5).String="Pari a Ribasso del ___________%"

    oSheet.getCellByPosition(2,fine+8).String="(ribasso in lettere)"
    oSheet.getCellRangeByPosition (2,fine+8,6,fine+8).CellStyle="List-intest_med_c"
    
    # INSERISCI LA DATA E L'OFFERENTE
    oSheet.getCellByPosition(2, fine+10).Formula = '=CONCATENATE("Data, ";TEXT(NOW();"DD/MM/YYYY"))'
    oSheet.getCellRangeByPosition (2,fine+10,2,fine+10).CellStyle = "Ultimus"
    oSheet.getCellByPosition(2, fine+12).String = "L'OFFERENTE"
    oSheet.getCellByPosition(2, fine+12).CellStyle = 'centro_grassetto'
    oSheet.getCellByPosition(2, fine+13).String= '(timbro e firma)'
    oSheet.getCellByPosition(2, fine+13).CellStyle = 'centro_corsivo'
    
    # CONSOLIDA LA DATA	
    oRange = oSheet.getCellRangeByPosition (2,fine+10,2,fine+10)
    #~ Flags = com.sun.star.sheet.CellFlags.FORMULA
    aSaveData = oRange.getDataArray()
    oRange.setDataArray(aSaveData)

    oSheet.getCellRangeByPosition(0, 0, getLastUsedCell(oSheet).EndColumn, getLastUsedCell(oSheet).EndRow).CellBackColor = -1

    _gotoCella(0, 1)
    adatta_altezza_riga(nSheet)
    
    oSheet.PageStyle = 'PageStyle_COMPUTO_A4'
    pagestyle = oDoc.StyleFamilies.getByName('PageStyles').getByName('PageStyle_COMPUTO_A4')
    pagestyle.HeaderIsOn =  True
    left = pagestyle.RightPageHeaderContent.LeftText.Text
    
    pagestyle.HeaderIsOn= True
    oHContent=pagestyle.RightPageHeaderContent
    oHContent.LeftText.String = uno.fileUrlToSystemPath(oDoc.getURL())
    oHContent.CenterText.String=''
    oHContent.RightText.String = tempo = ''.join(''.join(''.join(str(datetime.now()).split('.')[0].split(' ')).split('-')).split(':'))

    pagestyle.RightPageHeaderContent=oHContent
    return
########################################################################
def ins_voce_elenco(arg=None):
    '''
    Inserisce una nuova riga voce in Elenco Prezzi
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    refresh(0) 
    oSheet = oDoc.CurrentController.ActiveSheet
    _gotoCella(0,3)
    oSheet.getRows().insertByIndex(3,1)
    
    oSheet.getCellByPosition(0, 3).CellStyle = "EP-aS"
    oSheet.getCellByPosition(1, 3).CellStyle = "EP-a"
    oSheet.getCellRangeByPosition(2, 3, 7, 3).CellStyle = "EP-mezzo"
    oSheet.getCellRangeByPosition(8, 3, 9, 3).CellStyle = "EP-sfondo"
    for el in(5, 11, 15, 19, 26):
        oSheet.getCellByPosition(el, 3).CellStyle = "EP-mezzo %"

    for el in(12, 16, 20, 21):#(12, 16, 20):
        oSheet.getCellByPosition(el, 3).CellStyle = 'EP statistiche_q'

    for el in(13, 17, 23, 24, 25):#(12, 16, 20):
        oSheet.getCellByPosition(el, 3).CellStyle = 'EP statistiche'

    oSheet.getCellRangeByPosition(0, 3, 26, 3).clearContents(HARDATTR)
    oSheet.getCellByPosition(11, 3).Formula = '=IF(ISERROR(N4/$N$2);"--";N4/$N$2)'
    #~ oSheet.getCellByPosition(11, 3).Formula = '=N4/$N$2'
    oSheet.getCellByPosition(12, 3).Formula = '=SUMIF(AA;A4;BB)'
    oSheet.getCellByPosition(13, 3).Formula = '=SUMIF(AA;A4;cEuro)'

    #copio le formule dalla riga sotto
    oRangeAddress = oSheet.getCellRangeByPosition(15, 4, 26, 4).getRangeAddress()
    oCellAddress = oSheet.getCellByPosition(15,3).getCellAddress()
    oSheet.copyRange(oCellAddress, oRangeAddress)
    oCell = oSheet.getCellByPosition(2, 3)
    valida_cella(oCell, '"cad";"corpo";"dm";"dm²";"dm³";"kg";"lt";"m";"m²";"m³";"q";"t";"',titoloInput='Scegli...', msgInput='Unità di misura')
    refresh(1)
########################################################################
# nuova_voce ###########################################################
def ins_voce_computo_grezza(lrow):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #lrow = Range2Cell()[1]
########################################################################
# questo sistema eviterebbe l'uso della sheet S5 da cui copiare i range campione
# potrei svuotare la S5 ma allungando di molto il codice per la generazione della voce
# per ora lascio perdere
    # inserisco le righe ed imposto gli stili
    #~ insRows(lrow,4) #inserisco le righe
    #~ oSheet.getCellByPosition(0,lrow).CellStyle = 'Comp Start Attributo'
    #~ oSheet.getCellRangeByPosition(0,lrow,30,lrow).CellStyle = 'Comp-Bianche sopra'
    #~ oSheet.getCellByPosition(2,lrow).CellStyle = 'Comp-Bianche sopraS'
    #~
    #~ oSheet.getCellByPosition(0,lrow+1).CellStyle = 'comp progress'
    #~ oSheet.getCellByPosition(1,lrow+1).CellStyle = 'comp Art-EP'
    #~ oSheet.getCellRangeByPosition(2,lrow+1,8,lrow+1).CellStyle = 'Comp-Bianche in mezzo Descr'
    #~ oSheet.getCellRangeByPosition(2,lrow+1,8,lrow+1).merge(True)
########################################################################
## vado alla vecchia maniera ## copio il range di righe computo da S5 ##
    oSheetto = oDoc.getSheets().getByName('S5')
    #~ oRangeAddress = oSheetto.getCellRangeByName('$A$9:$AR$12').getRangeAddress()
    oRangeAddress = oSheetto.getCellRangeByPosition(0, 8, 42, 11).getRangeAddress()
    oCellAddress = oSheet.getCellByPosition(0,lrow).getCellAddress()
    oSheet.getRows().insertByIndex(lrow, 4)#~ insRows(lrow,4) #inserisco le righe
    oSheet.copyRange(oCellAddress, oRangeAddress)
########################################################################
# controllo la presenza di voci abbreviate e nel caso adatto la formula
    for i in range(3, 10):
        if '=IF(LEN(VLOOKUP(B' in oSheet.getCellByPosition(2, i).getFormula():

            #~ primi = oDoc.Sheets.getByName('S1').getCellByPosition(7,336).Value #S1.H337
            #~ ultimi = oDoc.Sheets.getByName('S1').getCellByPosition(7,337).Value #S1.H338
            #~ sformula = '=IF(LEN(VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;FALSE()))<' + str(primi+ultimi) + ';VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;FALSE());CONCATENATE(LEFT(VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;FALSE());160);" [...] ";RIGHT(VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;FALSE());' + str(ultimi) + ')))'
            sformula = '=IF(LEN(VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;FALSE()))<($S1.$H$337+$S1.H338);VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;FALSE());CONCATENATE(LEFT(VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;FALSE());$S1.$H$337);" [...] ";RIGHT(VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;FALSE());$S1.$H$338)))'

            oSheet.getCellByPosition(2, lrow+1).Formula = sformula
            break
########################################################################
# raggruppo i righi di mirura
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    oCellRangeAddr.StartColumn = 0
    oCellRangeAddr.EndColumn = 0
    oCellRangeAddr.StartRow = lrow+2
    oCellRangeAddr.EndRow = lrow+2
    oSheet.group(oCellRangeAddr, 1)
########################################################################
# correggo alcune formule
    oSheet.getCellByPosition(13,lrow+3).Formula ='=J'+str(lrow+4)
    oSheet.getCellByPosition(35,lrow+3).Formula ='=B'+str(lrow+2)

    if oSheet.getCellByPosition(31, lrow-1).CellStyle in('livello2 valuta', 'Livello-0-scritta', 'Livello-1-scritta', 'compTagRiservato'):
        oSheet.getCellByPosition(31, lrow+3).Value = oSheet.getCellByPosition(31, lrow-1).Value
        oSheet.getCellByPosition(32, lrow+3).Value = oSheet.getCellByPosition(32, lrow-1).Value
        oSheet.getCellByPosition(33, lrow+3).Value = oSheet.getCellByPosition(33, lrow-1).Value
########################################################################
    _gotoCella(1,lrow+1)
########################################################################
# ins_voce_computo #####################################################
def ins_voce_computo(arg=None): #TROPPO LENTA
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    if oSheet.getCellByPosition(0, lrow).CellStyle in(noVoce + stili_computo):
        lrow = next_voice(lrow, 1)
    else:
        return
    ins_voce_computo_grezza(lrow)
    numera_voci(0)
    if conf.read(path_conf, 'Generale', 'pesca_auto') == '1':
        pesca_cod()
########################################################################
# leeno.conf  ##########################################################
def leeno_conf(arg=None):
    '''
    Visualizza il menù di configurazione
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    try:
        oSheet = oDoc.getSheets().getByName('S1')
    except:
        return
    psm = uno.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDlg_config = dp.createDialog("vnd.sun.star.script:UltimusFree2.Dlg_config?language=Basic&location=application")
    oDialog1Model = oDlg_config.Model
    try:
        if conf.read(path_conf, 'Generale', 'visualizza_tabelle_extra') == '1': oDlg_config.getControl('CheckBox2').State = 1
        if conf.read(path_conf, 'Generale', 'pesca_auto') == '1': oDlg_config.getControl('CheckBox1').State = 1 #pesca codice automatico
        sString = oDlg_config.getControl('TextField1')
        sString.Text = conf.read(path_conf, 'Generale', 'altezza_celle')
        
        sString = oDlg_config.getControl("ComboBox1")
        sString.Text = conf.read(path_conf, 'Generale', 'visualizza') #visualizza all'avvio
        
        sString = oDlg_config.getControl("ComboBox2") #spostamento ad INVIO
        if conf.read(path_conf, 'Generale', 'movedirection')== '1':
            sString.Text = 'A DESTRA' 
        elif conf.read(path_conf, 'Generale', 'movedirection')== '0':
            sString.Text = 'IN BASSO' 
        
        sString = oDlg_config.getControl('TextField5')
        sString.Text =oSheet.getCellRangeByName('S1.H319').Value * 100 #sicurezza
        sString = oDlg_config.getControl('TextField6')
        sString.Text =oSheet.getCellRangeByName('S1.H320').Value * 100 #spese_generali

        sString = oDlg_config.getControl('TextField7')
        sString.Text =oSheet.getCellRangeByName('S1.H321').Value * 100 #utile_impresa
        
        #accorpa_spese_utili
        if oSheet.getCellRangeByName('S1.H321').Value == 1: oDlg_config.getControl('CheckBox4').State = 1

        sString = oDlg_config.getControl('TextField8')
        sString.Text =oSheet.getCellRangeByName('S1.H324').Value * 100 #sconto
        
        sString = oDlg_config.getControl('TextField9')
        sString.Text =oSheet.getCellRangeByName('S1.H326').Value * 100 #maggiorazione
        
        # fullscreen
        
        oLayout = oDoc.CurrentController.getFrame().LayoutManager
        if oLayout.isElementVisible('private:resource/toolbar/standardbar') == False:
            oDlg_config.getControl('CheckBox3').State = 1
        
        sString = oDlg_config.getControl('TextField10')
        sString.Text =oSheet.getCellRangeByName('S1.H337').Value #inizio_voci_abbreviate

        sString = oDlg_config.getControl('TextField11')
        sString.Text =oSheet.getCellRangeByName('S1.H338').Value #fine_voci_abbreviate
        
        # riga_bianca_categorie
        if oSheet.getCellRangeByName('S1.H334').Value == 1: oDlg_config.getControl('CheckBox5').State = 1
        
        # voci_senza_numerazione
        if oSheet.getCellRangeByName('S1.H335').Value == 1:
            oDlg_config.getControl('CheckBox6').State = 0
        else:
            oDlg_config.getControl('CheckBox6').State = 1
            
        # voci_senza_numerazione
        if conf.read(path_conf, 'Generale', 'torna_a_ep') == '1': oDlg_config.getControl('CheckBox8').State = 1

        
        # Contabilità abilita
        if oSheet.getCellRangeByName('S1.H328').Value == 1: oDlg_config.getControl('CheckBox7').State = 1
        sString = oDlg_config.getControl('TextField13')
        sString.Text = conf.read(path_conf, 'Contabilità', 'idxSAL')
    except:
        config_default()
    oDlg_config.execute()
    
    if oDlg_config.getControl('CheckBox3').State == 1:
        toolbar_switch(0)
    else:
        toolbar_switch(1)
 
    conf.write(path_conf, 'Generale', 'visualizza', oDlg_config.getControl('ComboBox1').getText())
    
    ctx = XSCRIPTCONTEXT.getComponentContext()
    oGSheetSettings = ctx.ServiceManager.createInstanceWithContext("com.sun.star.sheet.GlobalSheetSettings", ctx)
    if oDlg_config.getControl('ComboBox2').getText() == 'IN BASSO':
        conf.write(path_conf, 'Generale', 'movedirection', '0')
        oGSheetSettings.MoveDirection = 0
    else:
        conf.write(path_conf, 'Generale', 'movedirection', '1')
        oGSheetSettings.MoveDirection = 1
    conf.write(path_conf, 'Generale', 'altezza_celle', oDlg_config.getControl('TextField1').getText())
    conf.write(path_conf, 'Generale', 'visualizza_tabelle_extra', str(oDlg_config.getControl('CheckBox2').State))
    conf.write(path_conf, 'Generale', 'pesca_auto', str(oDlg_config.getControl('CheckBox1').State))
    conf.write(path_conf, 'Generale', 'torna_a_ep', str(oDlg_config.getControl('CheckBox8').State)) #torna su prezzario
        


    conf.write(path_conf, 'Computo', 'riga_bianca_categorie', str(oDlg_config.getControl('CheckBox5').State))
    #~ conf.write(path_conf, 'Computo', 'voci_senza_numerazione', str(oDlg_config.getControl('CheckBox6').State))
    if oDlg_config.getControl('CheckBox6').State == 1:
        oSheet.getCellRangeByName('S1.H335').Value = 0
    else:
        oSheet.getCellRangeByName('S1.H335').Value = 1
    conf.write(path_conf, 'Computo', 'inizio_voci_abbreviate', oDlg_config.getControl('TextField10').getText())
    conf.write(path_conf, 'Computo', 'fine_voci_abbreviate', oDlg_config.getControl('TextField11').getText())

    conf.write(path_conf, 'Contabilità', 'abilita', str(oDlg_config.getControl('CheckBox7').State))
    conf.write(path_conf, 'Contabilità', 'idxSAL', oDlg_config.getControl('TextField13').getText())

########################################################################
#percorso di ricerca di leeno.conf
if sys.platform == 'win32':
    path_conf = os.getenv("HOMEDRIVE") + os.getenv("HOMEPATH") + '/.config/leeno/leeno.conf'
else:
    path_conf = os.getenv("HOME") + '/.config/leeno/leeno.conf'
class conf:
    '''
    path    { string }: indirizzo del file di configurazione
    section { string }: sezione
    option  { string }: opzione
    value   { string }: valore
    '''
    def __init__(self, path=path_conf):
        #~ config = configparser.SafeConfigParser()
        #~ config.read(path) 
        #~ self.path = path
        pass
    def write(path, section, option, value):
        """
        Scrive i dati nel file di configurazione.
        http://www.programcreek.com/python/example/1033/ConfigParser.SafeConfigParser
        Write the specified Section.Option to the config file specified by path.
        Replace any previous value.  If the path doesn't exist, create it.
        Also add the option the the in-memory config.
        """
        config = configparser.SafeConfigParser()
        config.read(path)

        if not config.has_section(section):
            config.add_section(section)
        config.set(section, option, value)
        
        fp = open(path, 'w')
        config.write(fp)
        fp.close()
        
    def read(path, section, option):
        '''
        https://pymotw.com/2/ConfigParser/
        Legge i dati dal file di configurazione.
        '''
        config = configparser.SafeConfigParser()
        config.read(path)
        return config.get(section, option)
        
    def diz(path):
        '''
        Legge tutto il file di configurazione e restituisce un dizionario.
        '''
        my_config_parser_dict = {s:dict(config.items(s)) for s in config.sections()}
        return my_config_parser_dict
    
def config_default(arg=None):
    '''
    Imposta i parametri di default in leeno.conf
    '''
    parametri = (
    ('Zoom', 'fattore', '100'),
    ('Zoom', 'fattore_ottimale', '81'),
    ('Zoom', 'fullscreen', '0'),
    ('Generale', 'visualizza', 'Menù Principale'),
    ('Generale', 'altezza_celle', '1.25'),
    ('Generale', 'visualizza_tabelle_extra', '1'),
    ('Generale', 'pesca_auto', '1'),
    ('Generale', 'movedirection', '0'),
    ('Computo', 'riga_bianca_categorie', '1'),
    #~ ('Computo', 'voci_senza_numerazione', '0'),
    ('Computo', 'inizio_voci_abbreviate', '160'),
    ('Computo', 'fine_voci_abbreviate', '100'),
    ('Contabilità', 'abilita', '0'),
    ('Contabilità', 'idxSAL', '30')
    )
    for el in parametri:
        try:
            conf.read(path_conf, el[0], el[1])
        except:
            conf.write(path_conf, el[0], el[1], el[2])

    #~ leeno_conf()
########################################################################
def nuova_voce_scelta(arg=None): #assegnato a ctrl-shift-n
#~ def debug(arg=None):
    '''
    Contestualizza in ogni tabella l'inserimento delle voci.
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in('COMPUTO', 'VARIANTE'):
        ins_voce_computo()
    elif oSheet.Name =='Analisi di Prezzo':
        inizializza_analisi()
    elif oSheet.Name =='CONTABILITA':
        ins_voce_contab()
    elif oSheet.Name =='Elenco Prezzi':
        ins_voce_elenco()
    
# nuova_voce_contab  ##################################################
def ins_voce_contab(arg=None):
    '''
    Inserisce una nuova voce in CONTABILITA.
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    nome = oSheet.Name
    stile = oSheet.getCellByPosition( 0, lrow).CellStyle
    nSal = 1
    if stile == 'comp Int_colonna_R_prima':
        lrow += 1
    elif stile =='Ultimus_centro_bordi_lati':
        i = lrow
        while i != 0:
            if oSheet.getCellByPosition(23, i).Value != 0:
                nSal = int(oSheet.getCellByPosition(23, i).Value)
                break
            i -= 1
        while oSheet.getCellByPosition( 0, lrow).CellStyle == stile:
            lrow += 1
        if oSheet.getCellByPosition( 0, lrow).CellStyle == 'uuuuu':
            lrow += 1
            #~ nSal += 1
        #~ else
    elif stile == 'Comp TOTALI':
        pass
    elif stile in stili_contab:
        sStRange = Circoscrive_Voce_Computo_Att(lrow)
        nSal = int(oSheet.getCellByPosition(23, sStRange.RangeAddress.StartRow + 1).Value)
        if oSheet.getCellByPosition(22, sStRange.RangeAddress.StartRow + 1).String == '#reg':
            if DlgSiNo("""Inserendo qui una nuova voce, comprometterai
la validità degli atti contabili già emessi.

VUOI PROCEDERE?

Scegliendo Sì sarai costretto a rigenerarli!
Scegliendo No, potrai decidere una diversa posizione di inserimento.""", 'Voce già registrata!') ==3:
                return

        data = oSheet.getCellByPosition(1, sStRange.RangeAddress.StartRow + 2).Value
        lrow = next_voice(lrow)
    else:
        return
    oSheetto = oDoc.getSheets().getByName('S5')
    oRangeAddress = oSheetto.getCellRangeByPosition(0, 22, 48, 26).getRangeAddress()
    oCellAddress = oSheet.getCellByPosition(0,lrow).getCellAddress()
    oSheet.getRows().insertByIndex(lrow, 5) #inserisco le righe
    oSheet.copyRange(oCellAddress, oRangeAddress)
    oSheet.getCellRangeByPosition(0, lrow, 48, lrow+5).Rows.OptimalHeight = True
    _gotoCella(1, lrow+1)

    #~ if(oSheet.getCellByPosition(0,lrow).queryIntersection(oSheet.getCellRangeByName('#Lib#'+str(nSal)).getRangeAddress())):
        #~ chi('appartiene')
    #~ else:
        #~ chi('nooooo')
    #~ return

    sStRange = Circoscrive_Voce_Computo_Att(lrow)
    sopra = sStRange.RangeAddress.StartRow

    try:
        oSheet.getCellByPosition(1, sopra+2).Value = data
    except:
        oSheet.getCellByPosition(1, sopra+2).Value = date.today().toordinal()-693594

########################################################################
# controllo la presenza di voci abbreviate e nel caso adatto la formula
    for i in range(3, 10):
        if '=IF(LEN(VLOOKUP(B' in oSheet.getCellByPosition(2, i).getFormula():
            sformula = '=IF(LEN(VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;FALSE()))<($S1.$H$337+$S1.H338);VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;FALSE());CONCATENATE(LEFT(VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;FALSE());$S1.$H$337);" [...] ";RIGHT(VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;FALSE());$S1.$H$338)))'
            oSheet.getCellByPosition(2, lrow+1).Formula = sformula
            break
########################################################################
# raggruppo i righi di mirura
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    oCellRangeAddr.StartColumn = 0
    oCellRangeAddr.EndColumn = 0
    oCellRangeAddr.StartRow = lrow+2
    oCellRangeAddr.EndRow = lrow+2
    oSheet.group(oCellRangeAddr, 1)
########################################################################

    if oDoc.NamedRanges.hasByName('#Lib#'+str(nSal)):
        if lrow -1 == oSheet.getCellRangeByName('#Lib#'+str(nSal)).getRangeAddress().EndRow:
            nSal += 1
    
    oSheet.getCellByPosition(23, sopra + 1).Value = nSal
    oSheet.getCellByPosition(23, sopra + 1).CellStyle = 'Sal'
    
    oSheet.getCellByPosition(35, sopra+4).Formula = '=B'+ str(sopra+2)
    oSheet.getCellByPosition(36, sopra+4).Formula = '=IF(ISERROR(P'+ str(sopra +5) +');"";IF(P' + str(sopra+5) +'<>"";P' + str(sopra +5) + ';""))'
    oSheet.getCellByPosition(36, sopra+4).CellStyle = "comp -controolo"
    
    numera_voci(0)
    if conf.read(path_conf, 'Generale', 'pesca_auto') == '1':
        pesca_cod()

########################################################################
# attiva contabilità  ##################################################
def attiva_contabilita(arg=None):
#~ def debug(arg=None):
    '''Se presente, attiva e visualizza le tabelle di contabilità'''

    oDoc = XSCRIPTCONTEXT.getDocument()
    if oDoc.Sheets.hasByName('S1'):
        oDoc.Sheets.getByName('S1').getCellByPosition(7,327).Value = 1
        if oDoc.Sheets.hasByName('CONTABILITA'):
            for el in('Registro', 'SAL','CONTABILITA'):
                if oDoc.Sheets.hasByName(el):_gotoSheet(el)
        else:
            oDoc.Sheets.insertNewByName('CONTABILITA',3)
            _gotoSheet('CONTABILITA')
            svuota_contabilita()
            ins_voce_contab()
            set_larghezza_colonne()
        _gotoSheet('CONTABILITA')
########################################################################
# svuota contabilità  ##################################################
def svuota_contabilita(arg=None):
#~ def debug(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    for n in range(1 ,20):
        if oDoc.NamedRanges.hasByName('#Lib#'+str(n)) == True:
            oDoc.NamedRanges.removeByName('#Lib#'+str(n))
            oDoc.NamedRanges.removeByName('#SAL#'+str(n))
            oDoc.NamedRanges.removeByName('#Reg#'+str(n))
    for el in('Registro', 'SAL', 'CONTABILITA'):
        if oDoc.Sheets.hasByName(el):
            oDoc.Sheets.removeByName(el)
    
    oDoc.Sheets.insertNewByName('CONTABILITA',3)
    oSheet = oDoc.Sheets.getByName('CONTABILITA')

    _gotoSheet('CONTABILITA')
    setTabColor(16757935)
    oSheet.getCellRangeByName('C1').String = 'CONTABILITA'
    oSheet.getCellRangeByName('C1').CellStyle = 'comp Int_colonna'
    oSheet.getCellRangeByName('C1').CellBackColor = 16757935
    oSheet.getCellByPosition(0,2).String = 'N.'
    oSheet.getCellByPosition(1,2).String = 'Articolo\nData'
    oSheet.getCellByPosition(2,2).String = 'LAVORAZIONI\nO PROVVISTE'
    oSheet.getCellByPosition(5,2).String = 'P.U.\nCoeff.'
    oSheet.getCellByPosition(6,2).String = 'Lung.'
    oSheet.getCellByPosition(7,2).String = 'Larg.'
    oSheet.getCellByPosition(8,2).String = 'Alt.\nPeso'
    oSheet.getCellByPosition(9,2).String = 'Quantità\nPositive'
    oSheet.getCellByPosition(11,2).String = 'Quantità\nNegative'
    oSheet.getCellByPosition(13,2).String = 'Prezzo\nunitario'
    oSheet.getCellByPosition(15,2).String = 'Importi'
    oSheet.getCellByPosition(17,2).String = 'Sicurezza\ninclusa'
    oSheet.getCellByPosition(18,2).String = 'Serve per avere le quantità\nrealizzate "pulite" e sommabili'
    oSheet.getCellByPosition(19,2).String = 'Lib.\nN.'
    oSheet.getCellByPosition(20,2).String = 'Lib.\nP.'
    oSheet.getCellByPosition(22,2).String = 'flag'
    oSheet.getCellByPosition(23,2).String = 'SAL\nN.'
    oSheet.getCellByPosition(25,2).String = 'Importi\nSAL parziali'
    oSheet.getCellByPosition(27,2).String = 'Sicurezza\nunitaria'
    oSheet.getCellByPosition(28,2).String = 'Materiali\ne Noli €'
    oSheet.getCellByPosition(29,2).String = 'Incidenza\nMdO %'
    oSheet.getCellByPosition(30,2).String = 'Importo\nMdO'
    oSheet.getCellByPosition(31,2).String = 'Super Cat'
    oSheet.getCellByPosition(32,2).String = 'Cat'
    oSheet.getCellByPosition(33,2).String = 'Sub Cat'
    #~ oSheet.getCellByPosition(34,2).String = 'tag B'sub Scrivi_header_moduli
    #~ oSheet.getCellByPosition(35,2).String = 'tag C'
    oSheet.getCellByPosition(36,2).String = 'Importi\nsenza errori'
    oSheet.getCellByPosition(0,2).Rows.Height = 800
    #~ colore colonne riga di intestazione
    oSheet.getCellRangeByPosition(0, 2, 36 , 2).CellStyle = 'comp Int_colonna_R'
    oSheet.getCellByPosition(0, 2).CellStyle = 'comp Int_colonna_R_prima'
    oSheet.getCellByPosition(18, 2).CellStyle = 'COnt_noP'
    oSheet.getCellRangeByPosition(0,0,0,3).Rows.OptimalHeight = True
    #~ riga di controllo importo
    oSheet.getCellRangeByPosition(0, 1, 36 , 1).CellStyle = 'comp In testa'
    oSheet.getCellByPosition(2,1).String = 'QUESTA RIGA NON VIENE STAMPATA'
    oSheet.getCellRangeByPosition(0, 1, 1, 1).merge(True)
    oSheet.getCellByPosition(13,1).String = 'TOTALE:'
    oSheet.getCellByPosition(20,1).String = 'SAL SUCCESSIVO:'
    
    
    oSheet.getCellByPosition(25, 1).Formula = '=$P$2-SUBTOTAL(9;$P$2:$P$2)'
    #~ 'pippi
    oSheet.getCellByPosition(15,1).Formula='=SUBTOTAL(9;P3:P4)' #importo lavori
    oSheet.getCellByPosition(0,1).Formula='=AK2' #importo lavori
    oSheet.getCellByPosition(17,1).Formula='=SUBTOTAL(9;R3:R4)' #importo sicurezza
    

    oSheet.getCellByPosition(28,1).Formula='=SUBTOTAL(9;AC3:AC4)' #importo materiali
    oSheet.getCellByPosition(29,1).Formula='=AE2/Z2'  #Incidenza manodopera %
    oSheet.getCellByPosition(29, 1).CellStyle = 'Comp TOTALI %'
    oSheet.getCellByPosition(30,1).Formula='=SUBTOTAL(9;AE3:AE4)' #importo manodopera
    oSheet.getCellByPosition(36,1).Formula='=SUBTOTAL(9;AK3:AK4)' #importo certo


    #~ rem riga del totale
    oSheet.getCellByPosition(2,3).String = 'T O T A L E'
    oSheet.getCellByPosition(15,3).Formula='=SUBTOTAL(9;P3:P4)' #importo lavori
    oSheet.getCellByPosition(17,3).Formula='=SUBTOTAL(9;R3:R4)' #importo sicurezza
    oSheet.getCellByPosition(30,3).Formula='=SUBTOTAL(9;AE3:AE4)' #importo manodopera
    oSheet.getCellRangeByPosition(0, 3, 36 , 3).CellStyle = 'Comp TOTALI'
    #~ rem riga rossa
    oSheet.getCellByPosition(0,4).String = 'Fine Computo'
    oSheet.getCellRangeByPosition(0, 4, 36 , 4).CellStyle = 'Riga_rossa_Chiudi'
    _gotoCella(0, 2)
    set_larghezza_colonne()
########################################################################
# inizializza_analisi ##################################################
def inizializza_analisi(arg=None):
    '''
    Se non presente, crea il foglio 'Analisi di Prezzo' ed inserisce la prima scheda
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    rifa_nomearea('S5', '$B$108:$P$133', 'blocco_analisi')
    if oDoc.getSheets().hasByName('Analisi di Prezzo') == False:
        oDoc.getSheets().insertNewByName('Analisi di Prezzo',1)
        oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
        oSheet.getCellRangeByPosition(0,0,15,0).CellStyle = 'Analisi_Sfondo'
        oSheet.getCellByPosition(0, 1).Value = 0
        oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
        oDoc.CurrentController.setActiveSheet(oSheet)
        setTabColor(12189608)
        oRangeAddress=oDoc.NamedRanges.blocco_analisi.ReferredCells.RangeAddress
        oCellAddress = oSheet.getCellByPosition(0, getLastUsedCell(oSheet).EndRow).getCellAddress()
        oDoc.CurrentController.select(oSheet.getCellByPosition(0,2))
        oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
        set_larghezza_colonne()
    else:
        oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
        oDoc.CurrentController.setActiveSheet(oSheet)
        if oDoc.getSheets().getByName('Analisi di Prezzo').IsVisible == False:
            #~ oDoc.getSheets().getByName('Analisi di Prezzo').IsVisible = True
            _gotoSheet('Analisi di Prezzo')
            return
        lrow = Range2Cell()[1]
        urow = getLastUsedCell(oSheet).EndRow
        if lrow >= urow:
            lrow = ultima_voce(oSheet)-5
        for n in range(lrow ,getLastUsedCell(oSheet).EndRow):
            if oSheet.getCellByPosition(0, n).CellStyle == 'An-sfondo-basso Att End':
                break 
        oRangeAddress=oDoc.NamedRanges.blocco_analisi.ReferredCells.RangeAddress
        oSheet.getRows().insertByIndex(n+2,26)
        oCellAddress = oSheet.getCellByPosition(0,n+2).getCellAddress()
        oDoc.CurrentController.select(oSheet.getCellByPosition(0,n+2+1))
        oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
    oSheet.copyRange(oCellAddress, oRangeAddress)
    inserisci_Riga_rossa()
    return
########################################################################
def inserisci_Riga_rossa(arg=None):
    '''
    Inserisce la riga rossa di chiusura degli elaborati
    Questa riga è un rigerimento per varie operazioni
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    nome = oSheet.Name
    if nome in('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        lrow = ultima_voce(oSheet) + 2
        for n in range(lrow, getLastUsedCell(oSheet).EndRow+2):
            if oSheet.getCellByPosition(0,n).CellStyle == 'Riga_rossa_Chiudi':
                return
        oSheet.getRows().insertByIndex(lrow,1)
        oSheet.getCellByPosition(0, lrow).String = 'Fine Computo'
        oSheet.getCellRangeByPosition(0,lrow,34,lrow).CellStyle='Riga_rossa_Chiudi'
    elif nome == 'Analisi di Prezzo':
        lrow = ultima_voce(oSheet) + 2
        oSheet.getCellByPosition( 0, lrow).String = 'Fine ANALISI'
        oSheet.getCellRangeByPosition(0,lrow,10,lrow).CellStyle='Riga_rossa_Chiudi' 
    elif nome == 'Elenco Prezzi':
        lrow = ultima_voce(oSheet) + 2
        oSheet.getCellByPosition( 0, lrow).String = 'Fine elenco'
        oSheet.getCellRangeByPosition(0,lrow,26,lrow).CellStyle='Riga_rossa_Chiudi' 
    oSheet.getCellByPosition(2, lrow).String = 'Questa riga NON deve essere cancellata, MAI!!!(ma può rimanere tranquillamente NASCOSTA!)'
########################################################################
# rifa_nomearea ########################################################
def rifa_nomearea(sSheet, sRange, sName):
    '''
    Definisce o ridefinisce un'area di dati a cui far riferimento
    sSheet = nome del foglio, es.: 'S5'
    sRange = individuazione del range di celle, es.:'$B$89:$L$89'
    sName = nome da attribuire all'area scelta, es.: "manodopera"
    '''
    sPath = "$'" + sSheet + "'." + sRange
    oDoc = XSCRIPTCONTEXT.getDocument()
    oRanges = oDoc.NamedRanges
    oCellAddress = oDoc.Sheets.getByName(sSheet).getCellRangeByName('A1').getCellAddress()
    if oRanges.hasByName(sName):
        oRanges.removeByName(sName)
    oRanges.addNewByName(sName,sPath,oCellAddress,0)
########################################################################
def struct_colore(l):
    '''
    Mette in vista struttura secondo il colore
    l { integer } : specifica il livello di categoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()

    #~ oDoc.CurrentController.ZoomValue = 400

    oSheet = oDoc.CurrentController.ActiveSheet
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    hriga = oSheet.getCellRangeByName('B4').CharHeight * 65
    #~ giallo(16777072,16777120,16777168)
    #~ verde(9502608,13696976,15794160)
    #~ viola(12632319,13684991,15790335)
    col0 = 16724787 #riga_rossa
    col1 = 16777072
    col2 = 16777120
    col3 = 16777168
# attribuisce i colori
    for y in range(3, getLastUsedCell(oSheet).EndRow):
        if oSheet.getCellByPosition(0, y).String == '':
            oSheet.getCellByPosition(0, y).CellBackColor = col3
        elif len(oSheet.getCellByPosition(0, y).String.split('.')) == 2:
            oSheet.getCellByPosition(0, y).CellBackColor = col2
        elif len(oSheet.getCellByPosition(0, y).String.split('.')) == 1:
            oSheet.getCellByPosition(0, y).CellBackColor = col1
    if l == 0:
        colore = col1
        myrange =(col1, col0)
    elif l == 1:
        colore = col2
        myrange =(col1, col2, col0)
    elif l == 2:
        colore = col3
        myrange =(col1, col2, col3, col0)
   
        for n in(3, 5, 7):
            oCellRangeAddr.StartColumn = n
            oCellRangeAddr.EndColumn = n
            oSheet.group(oCellRangeAddr,0)
            oSheet.getCellRangeByPosition(n, 0, n, 0).Columns.IsVisible=False
    test = ultima_voce(oSheet)+2
    lista = list()
    for n in range(0, test):
        if oSheet.getCellByPosition(0, n).CellBackColor == colore:
            oSheet.getCellByPosition(0,n).Rows.Height = hriga
            sopra = n+1
            for n in range(sopra+1, test):
                if oSheet.getCellByPosition(0, n).CellBackColor in myrange:
                    sotto = n-1
                    lista.append((sopra, sotto))
                    break
    for el in lista:
        oCellRangeAddr.StartRow = el[0]
        oCellRangeAddr.EndRow = el[1]
        oSheet.group(oCellRangeAddr,1)
        oSheet.getCellRangeByPosition(0, el[0], 0, el[1]).Rows.IsVisible=False
    oDoc.CurrentController.ZoomValue = 100
    return
########################################################################
def struttura_Elenco(arg=None):
    '''
    Dà una tonalità di colore, diverso dal colore dello stile cella, alle righe
    che non hanno il prezzo, come i titoli di capitolo e sottocapitolo.
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    #~ chi(oSheet.getCellRangeByName('A14976').CellBackColor)
    #~ return
    struct_colore(0) #attribuisce i colori
    struct_colore(1)
    struct_colore(2)
    return
###########################################
###########################################
###########################################
    col1 = 16777072
    col2 = 16777120
    col3 = 16777168
    oDoc = XSCRIPTCONTEXT.getDocument()
    #~ oDoc.CurrentController.ZoomValue = 400

    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    lista = list()
    test = getLastUsedCell(oSheet).EndRow-1
    for n in range(3, test):#
        if oSheet.getCellByPosition(4, n).String == '':
            #~ pass
            oSheet.getCellRangeByPosition(0, n, 0, n).CellBackColor = col2
            #~ oCellRangeAddr.StartRow = n
            #~ oCellRangeAddr.EndRow = n
            #~ oSheet.group(oCellRangeAddr,1)
    #~ oDoc.CurrentController.ZoomValue = 100

    cap = list()
    for y in range(3, getLastUsedCell(oSheet).EndRow):
        if len(oSheet.getCellByPosition(0, y).String)== 1:
            #~ oCellRangeAddr.StartRow = y
            #~ oCellRangeAddr.EndRow = n
            #~ oSheet.group(oCellRangeAddr,1)
            cap.append(y)
    cap.append(getLastUsedCell(oSheet).EndRow)
    #~ chi (cap)
    test = copy.deepcopy(cap)
    a = cap.pop(0)
    #~ chi(test)
    for el in test:
        try:
            b = cap.pop(0)
            oCellRangeAddr.StartRow = a +1
            oCellRangeAddr.EndRow = b -1
            oSheet.group(oCellRangeAddr,1)
            oSheet.getCellRangeByPosition(0, a+1, 0, b-1).Rows.IsVisible = False
            oSheet.getCellByPosition(0,el).Rows.Height = 520
        except IndexError:
            return
        #~ chi(a)
        a = b

########################################################################
# XML_toscana_import ###################################################
def XML_toscana_import(arg=None):
    '''
    Importazione di un prezzario XML della regione Toscana 
    in tabella Elenco Prezzi del template COMPUTO.
    '''
    MsgBox('Questa operazione potrebbe richiedere del tempo.','Avviso')

    try:
        filename = filedia('Scegli il file XML Toscana da importare', '*.xml')
        if filename == None: return
    except:
        return
    New_file.computo(0)
    # effettua il parsing del file XML
    tree = ElementTree()
    tree.parse(filename)
    
    # ottieni l'item root
    root = tree.getroot()
    iter = tree.getiterator()

    PRT = '{' + str(iter[0].getchildren()[0]).split('}')[0].split('{')[-1] + '}' # xmlns
    # nome del prezzario
    intestazione = root.find(PRT+'intestazione')
    titolo = 'Prezzario '+ intestazione.get('autore') + ' - ' + intestazione[0].get('area') +' '+ intestazione[0].get('anno')
    licenza = intestazione[1].get('descrizione').split(':')[0] +' '+ intestazione[1].get('tipo')
    titolo = titolo + '\nCopyright: ' + licenza  + '\nhttp://prezzariollpp.regione.toscana.it'

    Contenuto = root.find(PRT+'Contenuto')

    voci = root.getchildren()[1]

    tipo_lista = list()
    cap_lista = list()
    lista_articoli = list()
    lista_cap = list()
    lista_subcap = list()
    for el in voci:
        if el.tag == PRT+'Articolo':
            codice = el.get('codice')
            codicesp = codice.split('.')
        
        voce = el.getchildren()[2].text
        articolo = el.getchildren()[3].text
        if articolo == None:
            desc_voce = voce
        else:
            desc_voce = voce + ' ' + articolo
        udm = el.getchildren()[4].text

        try:
            sic = float(el.getchildren()[-1][-4].get('valore'))
        except IndexError:
            sic =''
        prezzo = float(el.getchildren()[5].text)
        try:
            mdo = float(el.getchildren()[-1][-1].get('percentuale'))/100
            mdoE = mdo * prezzo
        except IndexError:
            mdo =''
            mdoE = ''
        if codicesp[0] not in tipo_lista:
            tipo_lista.append(codicesp[0])
            cap =(codicesp[0], el.getchildren()[0].text, '', '', '', '', '')
            lista_cap.append(cap)
        if codicesp[0]+'.'+codicesp[1] not in cap_lista:
            cap_lista.append(codicesp[0]+'.'+codicesp[1])
            cap =(codicesp[0]+'.'+codicesp[1], el.getchildren()[1].text, '', '', '', '', '', '')
            lista_subcap.append(cap)
        voceel =(codice, desc_voce, udm, sic, prezzo, mdo, mdoE)
        lista_articoli.append(voceel)
# compilo ##############################################################
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = titolo
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oSheet.getCellByPosition(1, 0).String = titolo
    oSheet.getCellByPosition(2, 0).String = '''ATTENZIONE!
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.

Si consiglia una attenta lettura delle note informative disponibili sul sito istituzionale ufficiale prima di accedere al prezzario.'''
    oSheet.getCellByPosition(1, 0).CellStyle = 'EP-mezzo'
    n = 0

    for el in (lista_articoli, lista_cap, lista_subcap):
        oSheet.getRows().insertByIndex(4, len(el))
        lista_come_array = tuple(el)
        # Parametrizzo il range di celle a seconda della dimensione della lista
        scarto_colonne = 0 # numero colonne da saltare a partire da sinistra
        scarto_righe = 4 # numero righe da saltare a partire dall'alto
        colonne_lista = len(lista_come_array[1]) # numero di colonne necessarie per ospitare i dati
        righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati
        oRange = oSheet.getCellRangeByPosition( 0, 4, colonne_lista + 0 - 1, righe_lista + 4 - 1)
        oRange.setDataArray(lista_come_array)
        #~ oSheet.getRows().removeByIndex(3, 1)
        oDoc.CurrentController.setActiveSheet(oSheet)

        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellStyle = "EP-aS"
        oSheet.getCellRangeByPosition(1, 3, 1, righe_lista + 3 - 1).CellStyle = "EP-a"
        oSheet.getCellRangeByPosition(2, 3, 7, righe_lista + 3 - 1).CellStyle = "EP-mezzo"
        oSheet.getCellRangeByPosition(5, 3, 5, righe_lista + 3 - 1).CellStyle = "EP-mezzo %"
        oSheet.getCellRangeByPosition(8, 3, 9, righe_lista + 3 - 1).CellStyle = "EP-sfondo"
        oSheet.getCellRangeByPosition(11, 3, 11, righe_lista + 3 - 1).CellStyle = 'EP-mezzo %'
        oSheet.getCellRangeByPosition(12, 3, 12, righe_lista + 3 - 1).CellStyle = 'EP statistiche_q'
        oSheet.getCellRangeByPosition(13, 3, 13, righe_lista + 3 - 1).CellStyle = 'EP statistiche'
        if n == 1: 
            oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellBackColor = 16777120
        elif n == 2:
            oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellBackColor = 16777168
        n += 1
    #~ set_larghezza_colonne()
    toolbar_vedi()
    adatta_altezza_riga('Elenco Prezzi')
    riordina_ElencoPrezzi()
    struttura_Elenco()
    
    dest = filename[0:-4]+ '.ods'
    salva_come(dest)
    MsgBox('''
Importazione eseguita con successo!

ATTENZIONE:
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.

N.B.: Si consiglia una attenta lettura delle note informative disponibili sul sito istituzionale ufficiale prima di accedere al Prezzario.

    ''','ATTENZIONE!')
#~ ########################################################################
def fuf(arg=None):
    ''' Traduce un particolare formato DAT usato in falegnameria - non c'entra un tubo con LeenO.
        E' solo una cortesia per un amico.'''
    bak_timestamp()
    filename = filedia('Scegli il file XML-SIX da importare', '*.dat')
    riga = list()
    try:
        f = open(filename, 'r')
    except TypeError:
        return
    ordini = list()
    riga =('Codice', 'Descrizione articolo', 'Quantità', 'Data consegna','Conto lavoro', 'Prezzo(€)')
    ordini.append(riga)
    
    for row in f:
        art =row[:15]
        if art[0:4] not in('HEAD', 'FEET'):
            art = art[4:]
            des =row[22:62]
            num = 1 #row[72:78].replace(' ','')
            car =row[78:87]
            dataC =row[96:104]
            dataC = '=DATE('+ dataC[:4]+';'+dataC[4:6]+';'+dataC[6:] + ')'
            clav =row[120:130]
            prz =row[142:-1]
            riga =(art, des, num, dataC, clav, float(prz.strip()))
            ordini.append(riga)

    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lista_come_array = tuple(ordini)
    colonne_lista = len(lista_come_array[0]) # numero di colonne necessarie per ospitare i dati
    righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati

    oRange = oSheet.getCellRangeByPosition( 0,
                                            0,
                                            colonne_lista -1, # l'indice parte da 0
                                            righe_lista -1)
    oRange.setFormulaArray(lista_come_array)
    
    #~ oDoc.CurrentController.select(oSheet.getCellRangeByPosition(3, 1, 3, getLastUsedCell(oSheet).EndRow+1))
    
    oSheet.getCellRangeByPosition(0, 0, getLastUsedCell(oSheet).EndColumn, getLastUsedCell(oSheet).EndRow).Columns.OptimalWidth = True

    return
    copy_clip()

    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    oProp = []
    oProp0 = PropertyValue()
    oProp0.Name = 'Flags'
    oProp0.Value = 'D'
    oProp1 = PropertyValue()
    oProp1.Name = 'FormulaCommand'
    oProp1.Value = 0
    oProp2 = PropertyValue()
    oProp2.Name = 'SkipEmptyCells'
    oProp2.Value = False
    oProp3 = PropertyValue()
    oProp3.Name = 'Transpose'
    oProp3.Value = False
    oProp4 = PropertyValue()
    oProp4.Name = 'AsLink'
    oProp4.Value = False
    oProp5 = PropertyValue()
    oProp5.Name = 'MoveMode'
    oProp5.Value = 4
    oProp.append(oProp0)
    oProp.append(oProp1)
    oProp.append(oProp2)
    oProp.append(oProp3)
    oProp.append(oProp4)
    oProp.append(oProp5)
    properties = tuple(oProp)
    #~ _gotoCella(6,1)

    dispatchHelper.executeDispatch(oFrame, '.uno:InsertContents', '', 0, properties)
    #~ paste_clip()
    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, 1, 5, getLastUsedCell(oSheet).EndRow+1))

    ordina_col(3)
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
    
#~ ########################################################################
# XML_import_ep ########################################################
def XML_import_ep(arg=None):
    MsgBox('Questa operazione potrebbe richiedere del tempo.','Avviso')
    New_file.computo(0)
    '''
    Routine di importazione di un prezzario XML-SIX in tabella Elenco Prezzi
    del template COMPUTO.
    '''
    try:
        filename = filedia('Scegli il file XML-SIX da importare', '*.xml')
        if filename == None:
            return
        oDialogo_attesa = dlg_attesa()
        attesa().start() #mostra il dialogo
    except:
        return

    datarif = datetime.now()
    # inizializzazioe delle variabili
    lista_articoli = list() # lista in cui memorizzare gli articoli da importare
    diz_um = dict() # array per le unità di misura
    # stringhe per descrizioni articoli
    desc_breve = str()
    desc_estesa = str()
    # effettua il parsing del file XML
    tree = ElementTree()
    if filename == None:
        return
    tree.parse(filename)
    # ottieni l'item root
    root = tree.getroot()
    logging.debug(list(root))
    # effettua il parsing di tutti gli elemnti dell'albero
    iter = tree.getiterator()
    listaSOA = []
    articolo = []
    articolo_modificato =()
    lingua_scelta = 'it'
########################################################################
    # nome del prezzario
    prezzario = root.find('{six.xsd}prezzario')
    if len(prezzario.findall('{six.xsd}przDescrizione')) == 2:
        if prezzario.findall('{six.xsd}przDescrizione')[0].get('lingua') == lingua_scelta:
            nome = prezzario.findall('{six.xsd}przDescrizione')[0].get('breve')
        else:
            nome = prezzario.findall('{six.xsd}przDescrizione')[1].get('breve')
    else:
        nome = prezzario.findall('{six.xsd}przDescrizione')[0].get('breve')
########################################################################
    madre = ''
    for elem in iter:
        # esegui le verifiche sulla root dell'XML
        if elem.tag == '{six.xsd}intestazione':
            intestazioneId= elem.get('intestazioneId')
            lingua= elem.get('lingua')
            separatore= elem.get('separatore')
            separatoreParametri= elem.get('separatoreParametri')
            valuta= elem.get('valuta')
            autore= elem.get('autore')
            versione= elem.get('versione')
            # inserisci i dati generali
            #~ self.update_dati_generali(nome=None, cliente=None,
                                       #~ redattore=autore,
                                       #~ ricarico=1,
                                       #~ manodopera=None,
                                       #~ sicurezza=None,
                                       #~ indirizzo=None,
                                       #~ comune=None, provincia=None,
                                       #~ valuta=valuta)
        elif elem.tag == '{six.xsd}categoriaSOA':
            soaId = elem.get('soaId')
            soaCategoria = elem.get('soaCategoria')
            soaDescrizione = elem.find('{six.xsd}soaDescrizione')
            if soaDescrizione != None:
                breveSOA = soaDescrizione.get('breve')
            voceSOA =(soaCategoria, soaId, breveSOA)
            listaSOA.append(voceSOA)
        elif elem.tag == '{six.xsd}prezzario':
            prezzarioId = elem.get('prezzarioId')
            przId= elem.get('przId')
            livelli_struttura= len(elem.get('prdStruttura').split('.'))
            categoriaPrezzario= elem.get('categoriaPrezzario')
########################################################################
        elif elem.tag == '{six.xsd}unitaDiMisura':
            um_id= elem.get('unitaDiMisuraId')
            um_sim= elem.get('simbolo')
            um_dec= elem.get('decimali')
            # crea il dizionario dell'unita di misura
########################################################################
            #~ unità di misura
            unita_misura = ''
            try:
                if len(elem.findall('{six.xsd}udmDescrizione')) == 1:
                    unita_misura = elem.findall('{six.xsd}udmDescrizione')[0].get('breve')
                else:
                    if elem.findall('{six.xsd}udmDescrizione')[1].get('lingua') == lingua_scelta:
                        idx = 1 #ITALIANO
                    else:
                        idx = 0 #TEDESCO
                    unita_misura = elem.findall('{six.xsd}udmDescrizione')[idx].get('breve')
            except IndexError:
                pass
            diz_um[um_id] = unita_misura
########################################################################
        # se il tag è un prodotto fa parte degli articoli da analizzare
        elif elem.tag == '{six.xsd}prodotto':
            prod_id = elem.get('prodottoId')
            if prod_id is not None:
                prod_id = int(prod_id)
            tariffa= elem.get('prdId')
            voce = elem.get('voce')

            sic = elem.get('onereSicurezza')
            if sic != None:
                sicurezza = float(sic)
            else:
                sicurezza = ''
########################################################################
            if diz_um.get(elem.get('unitaDiMisuraId')) != None:
                unita_misura = diz_um.get(elem.get('unitaDiMisuraId'))
            else:
                unita_misura = ''
########################################################################
            # verifica e ricava le sottosezioni
            sub_mdo = elem.find('{six.xsd}incidenzaManodopera')
            if sub_mdo != None:
                mdo = float(sub_mdo.text)
            else:
                mdo =''
########################################################################
            #~ chi(elem.findall('{six.xsd}prdDescrizione')[0].get('breve'))
            #~ return
            try:
                if len(elem.findall('{six.xsd}prdDescrizione')) == 1:
                    desc_breve = elem.findall('{six.xsd}prdDescrizione')[0].get('breve')
                    desc_estesa = elem.findall('{six.xsd}prdDescrizione')[0].get('estesa')
                else:
                #descrizione voce
                    if elem.findall('{six.xsd}prdDescrizione')[0].get('lingua') == lingua_scelta:
                        idx = 0 #ITALIANO
                    else:
                        idx = 1 #TEDESCO
                        idx = 0 #ITALIANO
                    desc_breve = elem.findall('{six.xsd}prdDescrizione')[idx].get('breve')
                    desc_estesa = elem.findall('{six.xsd}prdDescrizione')[idx].get('estesa')
            except:
                pass
            if desc_breve == None: desc_breve = ''
            if desc_estesa == None: desc_estesa = ''
            if len(desc_breve) > len(desc_estesa): desc_voce = desc_breve
            else: desc_voce = desc_estesa
########################################################################
            sub_quot = elem.find('{six.xsd}prdQuotazione')
            if sub_quot != None:
                list_nr = sub_quot.get('listaQuotazioneId')
                if sub_quot.get('valore') != None:
                    valore = float(sub_quot.get('valore'))
                if valore == 0: valore = ''
                if sub_quot.get('quantita') is not None: quantita = float(sub_quot.get('quantita')) #SERVE DAVVERO???
                if desc_voce[:2] == '- ': desc_voce=desc_voce[2:]
                desc_voce = madre + '\n- ' + desc_voce
            else:
                madre = desc_voce
                valore = ''
                quantita = ''
            elem_7 = ''
            elem_11 = ''
            if mdo != '' and mdo != 0: elem_7 = mdo/100
            if sicurezza != '' and valore != '': elem_11 = valore*sicurezza/100
            # Nota che ora articolo_modificato non è più una lista ma una tupla,
            # riguardo al motivo, vedi commenti in basso
            articolo_modificato = (tariffa,          #2  colonna
                                    desc_voce,        #4  colonna
                                    unita_misura,     #6  colonna
                                    '',
                                    valore,           #7  prezzo
                                    elem_7,           #8  mdo %
                                    elem_11)          #11 sicurezza %
            lista_articoli.append(articolo_modificato)
# compilo ##############################################################
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = nome
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oSheet.getCellByPosition(1, 1).String = nome
    oSheet.getRows().insertByIndex(4, len(lista_articoli))

    lista_come_array = tuple(lista_articoli)
    # Parametrizzo il range di celle a seconda della dimensione della lista
    scarto_colonne = 0 # numero colonne da saltare a partire da sinistra
    scarto_righe = 4 # numero righe da saltare a partire dall'alto
    colonne_lista = len(lista_come_array[1]) # numero di colonne necessarie per ospitare i dati
    righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati
    oRange = oSheet.getCellRangeByPosition( scarto_colonne,
                                            scarto_righe,
                                            colonne_lista + scarto_colonne - 1, # l'indice parte da 0
                                            righe_lista + scarto_righe - 1)
    oRange.setDataArray(lista_come_array)
    oSheet.getRows().removeByIndex(3, 1)
    oDoc.CurrentController.setActiveSheet(oSheet)
    #~ struttura_Elenco()
    oDialogo_attesa.endExecute()
    MsgBox('Importazione eseguita con successo!','')
    autoexec()
# XML_import ###########################################################
########################################################################
def XML_import_multi(arg=None):
    MsgBox("L'importazione dati dal formato XML-SIX potrebbe richiedere del tempo.", 'Avviso')
    '''
    Routine di importazione di un prezzario XML-SIX in tabella Elenco Prezzi
    del template COMPUTO.
    Tratta da PreventARES https://launchpad.net/preventares
    di <Davide Vescovini> <davide.vescovini@gmail.com>
    *Versione bilingue*
    '''
    New_file.computo(0)
    try:
        filename = filedia('Scegli il file XML-SIX da importare', '*.xml')
        if filename == None:
            return
        oDialogo_attesa = dlg_attesa()
        attesa().start() #mostra il dialogo
    except:
        return
        
    datarif = datetime.now()
    # inizializzazioe delle variabili
    lista_articoli = list() # lista in cui memorizzare gli articoli da importare
    diz_um = dict() # array per le unità di misura
    # stringhe per descrizioni articoli
    desc_breve = str()
    desc_estesa = str()
    # effettua il parsing del file XML
    tree = ElementTree()
    if filename == None:
        return
    tree.parse(filename)
    # ottieni l'item root
    root = tree.getroot()
    logging.debug(list(root))
    # effettua il parsing di tutti gli elemnti dell'albero XML
    iter = tree.getiterator()
    listaSOA = []
    articolo = []
    lingua_scelta = 'it'
########################################################################
    # nome del prezzario
    prezzario = root.find('{six.xsd}prezzario')
    if len(prezzario.findall('{six.xsd}przDescrizione')) == 2:
        if prezzario.findall('{six.xsd}przDescrizione')[0].get('lingua') == lingua_scelta:
            nome1 = prezzario.findall('{six.xsd}przDescrizione')[0].get('breve')
            nome2 = prezzario.findall('{six.xsd}przDescrizione')[1].get('breve')
        else:
            nome1 = prezzario.findall('{six.xsd}przDescrizione')[1].get('breve')
            nome2 = prezzario.findall('{six.xsd}przDescrizione')[0].get('breve')
        nome=nome1+'\n§\n'+nome2
    else:
        nome = prezzario.findall('{six.xsd}przDescrizione')[0].get('breve')
########################################################################
    suffB_IT, suffE_IT, suffB_DE, suffE_DE = '', '', '', ''
    test = True
    madre = ''
    for elem in iter:
        # esegui le verifiche sulla root dell'XML
        if elem.tag == '{six.xsd}intestazione':
            intestazioneId= elem.get('intestazioneId')
            lingua= elem.get('lingua')
            separatore= elem.get('separatore')
            separatoreParametri= elem.get('separatoreParametri')
            valuta= elem.get('valuta')
            autore= elem.get('autore')
            versione= elem.get('versione')
        elif elem.tag == '{six.xsd}categoriaSOA':
            soaId = elem.get('soaId')
            soaCategoria = elem.get('soaCategoria')
            soaDescrizione = elem.find('{six.xsd}soaDescrizione')
            if soaDescrizione != None:
                breveSOA = soaDescrizione.get('breve')
            voceSOA =(soaCategoria, soaId, breveSOA)
            listaSOA.append(voceSOA)
        elif elem.tag == '{six.xsd}prezzario':
            prezzarioId = elem.get('prezzarioId')
            przId= elem.get('przId')
            livelli_struttura= len(elem.get('prdStruttura').split('.'))
            categoriaPrezzario= elem.get('categoriaPrezzario')
########################################################################
        elif elem.tag == '{six.xsd}unitaDiMisura':
            um_id= elem.get('unitaDiMisuraId')
            um_sim= elem.get('simbolo')
            um_dec= elem.get('decimali')
            # crea il dizionario dell'unita di misura
########################################################################
            #~ unità di misura
            unita_misura = ''
            try:
                if len(elem.findall('{six.xsd}udmDescrizione')) == 1:
                    unita_misura = elem.findall('{six.xsd}udmDescrizione')[0].get('breve')
                else:
                    if elem.findall('{six.xsd}udmDescrizione')[1].get('lingua') == lingua_scelta:
                        unita_misura1 = elem.findall('{six.xsd}udmDescrizione')[1].get('breve')
                        unita_misura2 = elem.findall('{six.xsd}udmDescrizione')[0].get('breve')
                    else:
                        unita_misura1 = elem.findall('{six.xsd}udmDescrizione')[0].get('breve')
                        unita_misura2 = elem.findall('{six.xsd}udmDescrizione')[1].get('breve')
                if unita_misura != None:
                    unita_misura = unita_misura1 +' § '+ unita_misura2
            except IndexError:
                pass
            diz_um[um_id] = unita_misura
########################################################################
        # se il tag è un prodotto fa parte degli articoli da analizzare
        elif elem.tag == '{six.xsd}prodotto':

            prod_id = elem.get('prodottoId')
            if prod_id is not None:
                prod_id = int(prod_id)
            tariffa= elem.get('prdId')
            sic = elem.get('onereSicurezza')
            if sic != None:
                sicurezza = float(sic)
            else:
                sicurezza = ''
########################################################################
            if diz_um.get(elem.get('unitaDiMisuraId')) != None:
                unita_misura = diz_um.get(elem.get('unitaDiMisuraId'))
            else:
                unita_misura = ''
########################################################################
            # verifica e ricava le sottosezioni
            sub_mdo = elem.find('{six.xsd}incidenzaManodopera')
            if sub_mdo != None:
                mdo = float(sub_mdo.text)
            else:
                mdo =''
########################################################################
            # descrizione voci
            desc_estesa1, desc_estesa2 = '', ''
            if test == 0:
                test = 1
                suffB_IT = suffB_IT + ' '
                suffE_IT = suffE_IT + ' '
                suffB_DE = suffB_DE + ' '
                suffE_DE = suffE_DE + ' '
            #~ try:
            if len(elem.findall('{six.xsd}prdDescrizione')) == 1:
                desc_breve = elem.findall('{six.xsd}prdDescrizione')[0].get('breve')
                desc_estesa = elem.findall('{six.xsd}prdDescrizione')[0].get('estesa')
            else:
        #descrizione voce
                if elem.findall('{six.xsd}prdDescrizione')[0].get('lingua') == lingua_scelta:
                    desc_breve1  = elem.findall('{six.xsd}prdDescrizione')[0].get('breve')
                    desc_breve2  = elem.findall('{six.xsd}prdDescrizione')[1].get('breve')
                    desc_estesa1 = elem.findall('{six.xsd}prdDescrizione')[0].get('estesa')
                    desc_estesa2 = elem.findall('{six.xsd}prdDescrizione')[1].get('estesa')
                else:
                    desc_breve1  = elem.findall('{six.xsd}prdDescrizione')[1].get('breve')
                    desc_breve2  = elem.findall('{six.xsd}prdDescrizione')[0].get('breve')
                    desc_estesa1 = elem.findall('{six.xsd}prdDescrizione')[1].get('estesa')
                    desc_estesa2 = elem.findall('{six.xsd}prdDescrizione')[0].get('estesa')
                if desc_breve1 == None:
                    desc_breve1 = ''
                if desc_breve2 == None:
                    desc_breve2 = ''
                if desc_estesa1 == None:
                    desc_estesa1 = ''
                if desc_estesa2 == None:
                    desc_estesa2 = ''
                desc_breve = suffB_IT + desc_breve1.strip() +'\n§\n'+ suffB_DE + desc_breve2.strip()
                desc_estesa = suffE_IT + desc_estesa1.strip() +'\n§\n'+ suffE_DE + desc_estesa2.strip()
            if len(desc_breve) > len(desc_estesa):
                desc_voce = desc_breve
            else:
                desc_voce = desc_estesa
            #~ except IndexError:
                #~ pass
########################################################################
            sub_quot = elem.find('{six.xsd}prdQuotazione')
            if sub_quot != None:
                list_nr = sub_quot.get('listaQuotazioneId')
                if sub_quot.get('valore') != None:
                    valore = float(sub_quot.get('valore'))
                if valore == 0:
                    valore = ''
                if sub_quot.get('quantita') is not None: #SERVE DAVVERO???
                    quantita = float(sub_quot.get('quantita'))
            else:
                test = 0
                suffB_IT, suffB_DE, suffE_IT, suffE_DE = desc_breve1, desc_breve2, desc_estesa1, desc_estesa2
                valore = ''
                quantita = ''
            vuoto = ''
            elem_7 = ''
            elem_11 = ''
            articolo_modificato = (tariffa,          #2  colonna
                                    desc_voce,        #4  colonna
                                    unita_misura,     #6  colonna
                                    vuoto,
                                    valore,           #7  prezzo
                                    elem_7,           #8  mdo %
                                    elem_11)          #11 sicurezza %
            lista_articoli.append(articolo_modificato)
# compilo la tabella ###################################################
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = nome
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oSheet.getCellByPosition(1, 1).String = nome
    oSheet.getRows().insertByIndex(4, len(lista_articoli))

    lista_come_array = tuple(lista_articoli)
    # Parametrizzo il range di celle a seconda della dimensione della lista
    scarto_colonne = 0 # numero colonne da saltare a partire da sinistra
    scarto_righe = 4 # numero righe da saltare a partire dall'alto
    colonne_lista = len(lista_come_array[1]) # numero di colonne necessarie per ospitare i dati
    righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati
    oRange = oSheet.getCellRangeByPosition( scarto_colonne,
                                            scarto_righe,
                                            colonne_lista + scarto_colonne - 1, # l'indice parte da 0
                                            righe_lista + scarto_righe - 1)
    oRange.setDataArray(lista_come_array)
    oSheet.getRows().removeByIndex(3, 1)
    oDoc.CurrentController.setActiveSheet(oSheet)
    struttura_Elenco()
    oDialogo_attesa.endExecute()
    MsgBox('Importazione eseguita con successo!','')
    autoexec()
# XML_import_multi ###################################################
########################################################################
class importa_listino_leeno_th(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        importa_listino_leeno_run()
def importa_listino_leeno(arg=None):
#~ def debug (arg=None):
    importa_listino_leeno_th().start()
###
def importa_listino_leeno_run(arg=None):
    '''
    Esegue la conversione di un listino(formato LeenO) in template Computo
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ giallo(16777072,16777120,16777168)
    #~ verde(9502608,13696976,15794160)
    #~ viola(12632319,13684991,15790335)
    lista_articoli = list()
    nome = oSheet.getCellByPosition(2, 0).String
    chi(8)
    test = uFindStringCol('ATTENZIONE!', 5, oSheet)+1
    assembla = DlgSiNo('''Il riconoscimento di descrizioni e sottodescrizioni
dipende dalla colorazione di sfondo delle righe.

Nel caso in cui questa fosse alterata, il risultato finale
della conversione potrebbe essere inatteso.

Considera anche la possibilità di recuperare il formato XML(SIX)
di questo prezzario dal sito ufficiale dell'ente che lo rilascia.        

Vuoi assemblare descrizioni e sottodescrizioni?''', 'Richiesta')

    orig = oDoc.getURL()
    dest0 = orig[0:-4]+ '_new.ods'

    orig = uno.fileUrlToSystemPath(LeenO_path()+'/template/leeno/Computo_LeenO.ots')
    dest = uno.fileUrlToSystemPath(dest0)

    shutil.copyfile(orig, dest)
    oDialogo_attesa = dlg_attesa()
    attesa().start() #mostra il dialogo
    madre = ''
    for el in range(test, getLastUsedCell(oSheet).EndRow+1):
        tariffa = oSheet.getCellByPosition(2, el).String
        descrizione = oSheet.getCellByPosition(4, el).String
        um = oSheet.getCellByPosition(6, el).String
        sic = oSheet.getCellByPosition(11, el).String
        prezzo = oSheet.getCellByPosition(7, el).String
        mdo_p = oSheet.getCellByPosition(8, el).String
        mdo = oSheet.getCellByPosition(9, el).String
        if oSheet.getCellByPosition(2, el).CellBackColor in(16777072,16777120,9502608,13696976,12632319,13684991):
            articolo =(tariffa,
                        descrizione,
                        um,
                        sic,
                        prezzo,
                        mdo_p,
                        mdo,)
        elif oSheet.getCellByPosition(2, el).CellBackColor in(16777168,15794160,15790335):
            if assembla ==2: madre = descrizione
            articolo =(tariffa,
                        descrizione,
                        um,
                        sic,
                        prezzo,
                        mdo_p,
                        mdo,)
        else:
            if madre == '':
                descrizione = oSheet.getCellByPosition(4, el).String
            else:
                descrizione = madre + ' \n- ' + oSheet.getCellByPosition(4, el).String
            articolo =(tariffa,
                        descrizione,
                        um,
                        sic,
                        prezzo,
                        mdo_p,
                        mdo,)
        lista_articoli.append(articolo)
    oDialogo_attesa.endExecute()
    _gotoDoc(dest) #vado sul nuovo file
# compilo la tabella ###################################################
    oDoc = XSCRIPTCONTEXT.getDocument()
    oDialogo_attesa = dlg_attesa()
    attesa().start() #mostra il dialogo
    
    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = nome
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oSheet.getCellByPosition(1, 1).String = nome

    oSheet.getRows().insertByIndex(4, len(lista_articoli))
    lista_come_array = tuple(lista_articoli)
    # Parametrizzo il range di celle a seconda della dimensione della lista
    colonne_lista = len(lista_come_array[1]) # numero di colonne necessarie per ospitare i dati
    righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati
    oRange = oSheet.getCellRangeByPosition( 0,
                                            4,
                                            colonne_lista - 1, # l'indice parte da 0
                                            righe_lista + 4 - 1)
    oRange.setDataArray(lista_come_array)
    oSheet.getRows().removeByIndex(3, 1)
    oDoc.CurrentController.setActiveSheet(oSheet)
    oDialogo_attesa.endExecute()
    procedo = DlgSiNo('''Vuoi mettere in ordine la visualizzazione del prezzario?     

Le righe senza prezzo avranno una tonalità di sfondo
diversa dalle altre e potranno essere facilmente nascoste.

Questa operazione potrebbe richiedere del tempo.''', 'Richiesta...')
    if procedo ==2:
        attesa().start() #mostra il dialogo
        struttura_Elenco()
        oDialogo_attesa.endExecute()
    MsgBox('Conversione eseguita con successo!','')
    autoexec()
   
########################################################################
def importa_stili(arg=None):
    '''
    Importa tutti gli stili da un documento di riferimento. Se non è
    selezionato, il file di rifetimento è il template di leenO.
    '''
    if DlgSiNo('''Questa operazione sovrascriverà gli stili
del documento attivo se già presenti!

Se non scegli un file di riferimento, saranno
importati gli stili di default di LeenO.

Vuoi continuare?''', 'Importa Stili in blocco?') == 3: return
    filename = filedia('Scegli il file di riferimento...', '*.ods')
    if filename == None:
        #~ desktop = XSCRIPTCONTEXT.getDesktop()
        filename = LeenO_path()+'/template/leeno/Computo_LeenO.ots'
    else:
        filename = uno.systemPathToFileUrl(filename)
    oDoc = XSCRIPTCONTEXT.getDocument()
    oDoc.getStyleFamilies().loadStylesFromURL(filename,list())
    for el in oDoc.Sheets.ElementNames:
        oDoc.CurrentController.setActiveSheet(oDoc.getSheets().getByName(el))
        adatta_altezza_riga(el)
    try:
        _gotoSheet('Elenco Prezzi')
    except:
        pass
########################################################################
def parziale(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    if oSheet.Name in('COMPUTO','VARIANTE', 'CONTABILITA'):
        parziale_core(lrow)
        parziale_verifica()
def parziale_core(lrow):
    #~ lrow = 324
    if lrow == 0: return
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    sStRange = Circoscrive_Voce_Computo_Att(lrow)
    sopra = sStRange.RangeAddress.StartRow
    sotto = sStRange.RangeAddress.EndRow
    if oSheet.Name in('COMPUTO','VARIANTE'):
        if oSheet.getCellByPosition(0, lrow).CellStyle == 'comp 10 s' and \
        oSheet.getCellByPosition(1, lrow).CellStyle == 'Comp-Bianche in mezzo' and \
        oSheet.getCellByPosition(2, lrow).CellStyle == 'comp 1-a' or \
        oSheet.getCellByPosition(0, lrow).CellStyle == 'Comp End Attributo':
            oSheet.getRows().insertByIndex(lrow, 1)
        elif 'Parziale [' in(oSheet.getCellByPosition(8, lrow).String):
                pass
        else:
            return
        oSheet.getCellByPosition(1, lrow).CellStyle = 'Comp-Bianche in mezzo'
        oSheet.getCellRangeByPosition(2, lrow, 7, lrow).CellStyle = 'comp sotto centro'
        oSheet.getCellByPosition(8, lrow).CellStyle = 'comp sotto BiancheS'
        oSheet.getCellByPosition(9, lrow).CellStyle = 'Comp-Variante num sotto'
        oSheet.getCellByPosition(8, lrow).Formula = '''=CONCATENATE("Parziale [";VLOOKUP(B'''+ str(sopra+2) + ''';elenco_prezzi;3;FALSE());"]")'''
        for i in reversed(range(0, lrow)):
            if oSheet.getCellByPosition(9, i-1).CellStyle in('vuote2', 'Comp-Variante num sotto'):
                i
                break
        oSheet.getCellByPosition(9, lrow).Formula = "=SUBTOTAL(9;J" + str(i) + ":J" + str(lrow+1) + ")"
    if oSheet.Name in('CONTABILITA'): MsgBox('Contatta il canale Telegram https://t.me/joinchat/AAAAAEFGWSw-p_N6tUt0FA')
###
def parziale_verifica(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    #~ if oSheet.Name in('COMPUTO','VARIANTE', 'CONTABILITA'):
    sStRange = Circoscrive_Voce_Computo_Att(lrow)
    sopra = sStRange.RangeAddress.StartRow+2
    sotto = sStRange.RangeAddress.EndRow
    for n in range(sopra, sotto):
        if 'Parziale [' in(oSheet.getCellByPosition(8, n).String):
            parziale_core(n)
    #~ chi(oDoc.CurrentSelection.CellBackColor)


########################################################################
# abs2name ############################################################
def abs2name(nCol, nRow):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    idvoce = oSheet.getCellByPosition(nCol, nRow).AbsoluteName.split('$')
    return idvoce[2]+idvoce[3]
########################################################################
# vedi_voce_xpwe ############################################################
def vedi_voce_xpwe(riga_corrente,vRif,flags=''):
    """(riga d'inserimento, riga di riferimento)"""
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    sStRange = Circoscrive_Voce_Computo_Att(vRif)
    sStRange.RangeAddress
    idv = sStRange.RangeAddress.StartRow +1
    sotto = sStRange.RangeAddress.EndRow
    art = abs2name(1, idv)
    idvoce = abs2name(0, idv)
    quantity = abs2name(9, sotto)
    um = 'VLOOKUP(' + art + ';elenco_prezzi;3;FALSE())'
    oSheet.getCellByPosition(2, riga_corrente).Formula='=CONCATENATE("";" - vedi voce n. ";TEXT(' + idvoce +';"@");" - art. ";' + art + ';" [";' + um + ';"]"'
    if flags in('32768', '32769', '32801'):
        #~ oSheet.getCellByPosition(5, riga_corrente).Formula='=-' + quantity
    #~ else:
        oSheet.getCellByPosition(5, riga_corrente).Formula='=' + quantity
########################################################################
def strall(el, n=3):
    #~ el ='o'
    while len(el) < n:
        el = '0' + el
    return el

########################################################################
def converti_stringhe(arg=None):
    '''
    Converte in numeri le stinghe selezionate.
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    try:
        sRow = oDoc.getCurrentSelection().getRangeAddresses()[0].StartRow
        sCol = oDoc.getCurrentSelection().getRangeAddresses()[0].StartColumn
        eRow = oDoc.getCurrentSelection().getRangeAddresses()[0].EndRow
        eCol = oDoc.getCurrentSelection().getRangeAddresses()[0].EndColumn
    except AttributeError:
        sRow = oDoc.getCurrentSelection().getRangeAddress().StartRow
        sCol = oDoc.getCurrentSelection().getRangeAddress().StartColumn
        eRow = oDoc.getCurrentSelection().getRangeAddress().EndRow
        eCol = oDoc.getCurrentSelection().getRangeAddress().EndColumn
    oRange = oSheet.getCellRangeByPosition(sCol, sRow, eCol, eRow)
    for y in range(sCol, eCol+1):
        for x in range(sRow, eRow+1):
            try:
                oSheet.getCellByPosition(y, x).Value = float(oSheet.getCellByPosition(y, x).String.replace(',','.'))
            except:
                pass
    return
# XPWE_in ##########################################################
def XPWE_in(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    refresh(0)
    oDialogo_attesa = dlg_attesa('Caricamento dei dati...')
    if oDoc.getSheets().hasByName('S2') == False:
        MsgBox('Puoi usare questo comando da un file di computo nuovo o già esistente.','ATTENZIONE!')
        return
    else:
        MsgBox("Il contenuto dell'archivio XPWE sarà aggiunto a questo file.",'Avviso!')
    _gotoSheet('COMPUTO')
    oDoc.CurrentController.select(oDoc.getSheets().hasByName('COMPUTO')) # per evitare che lo script parta da un altro documento
    filename = filedia('Scegli il file XPWE da importare...','*.xpwe')#'*.xpwe')
    '''xml auto indent: http://www.freeformatter.com/xml-formatter.html'''
    # inizializzazione delle variabili
    datarif = datetime.now()
    lista_articoli = list() # lista in cui memorizzare gli articoli da importare
    diz_ep = dict() # array per le voci di elenco prezzi
    # effettua il parsing del file XML
    tree = ElementTree()
    if filename == 'Cancel' or filename == '':
        return
    try:
        tree.parse(filename)
    except TypeError:
        return
    except PermissionError:
        MsgBox('Accertati che il nome del file sia corretto.', 'ATTENZIONE! Impossibile procedere.')
        return
    # ottieni l'item root
    root = tree.getroot()
    logging.debug(list(root))
    # effettua il parsing di tutti gli elementi dell'albero XML
    iter = tree.getiterator()
    if root.find('FileNameDocumento'):
        nome_file = root.find('FileNameDocumento').text
    else:
        nome_file = "nome_file"

###
    dati = root.find('PweDatiGenerali')
    DatiGenerali = dati.getchildren()[0][0]
    percprezzi = DatiGenerali[0].text
    comune = DatiGenerali[1].text
    provincia = DatiGenerali[2].text
    oggetto = DatiGenerali[3].text
    committente = DatiGenerali[4].text
    impresa = DatiGenerali[5].text
    parteopera = DatiGenerali[6].text
###
#PweDGCapitoliCategorie
    try:
        CapCat = dati.find('PweDGCapitoliCategorie')
###
#PweDGSuperCapitoli
        lista_supcap = list()
        if CapCat.find('PweDGSuperCapitoli'):
            PweDGSuperCapitoli = CapCat.find('PweDGSuperCapitoli').getchildren()
            for elem in PweDGSuperCapitoli:
                id_sc = elem.get('ID')
                codice = elem.find('Codice').text
                try:
                    codice = elem.find('Codice').text
                except AttributeError:
                    codice = ''
                dessintetica = elem.find('DesSintetica').text
                percentuale = elem.find('Percentuale').text
                diz = dict()
                diz['id_sc'] = id_sc
                diz['codice'] = codice
                diz['dessintetica'] = dessintetica
                diz['percentuale'] = percentuale
                lista_supcap.append(diz)
###
#PweDGCapitoli
        lista_cap = list()
        if CapCat.find('PweDGCapitoli'):
            PweDGCapitoli = CapCat.find('PweDGCapitoli').getchildren()
            for elem in PweDGCapitoli:
                id_sc = elem.get('ID')
                codice = elem.find('Codice').text
                try:
                    codice = elem.find('Codice').text
                except AttributeError:
                    codice = ''
                dessintetica = elem.find('DesSintetica').text
                percentuale = elem.find('Percentuale').text
                diz = dict()
                diz['id_sc'] = id_sc
                diz['codice'] = codice
                diz['dessintetica'] = dessintetica
                diz['percentuale'] = percentuale
                lista_cap.append(diz)
###
#PweDGSubCapitoli
        lista_subcap = list()
        if CapCat.find('PweDGSubCapitoli'):
            PweDGSubCapitoli = CapCat.find('PweDGSubCapitoli').getchildren()
            for elem in PweDGSubCapitoli:
                id_sc = elem.get('ID')
                codice = elem.find('Codice').text
                try:
                    codice = elem.find('Codice').text
                except AttributeError:
                    codice = ''
                dessintetica = elem.find('DesSintetica').text
                percentuale = elem.find('Percentuale').text
                diz = dict()
                diz['id_sc'] = id_sc
                diz['codice'] = codice
                diz['dessintetica'] = dessintetica
                diz['percentuale'] = percentuale
                lista_subcap.append(diz)
###
#PweDGSuperCategorie
        lista_supcat = list()
        if CapCat.find('PweDGSuperCategorie'):
            PweDGSuperCategorie = CapCat.find('PweDGSuperCategorie').getchildren()
            for elem in PweDGSuperCategorie:
                id_sc = elem.get('ID')
                dessintetica = elem.find('DesSintetica').text
                try:
                    percentuale = elem.find('Percentuale').text
                except AttributeError:
                    percentuale = '0'
                supcat =(id_sc, dessintetica, percentuale)
                lista_supcat.append(supcat)
            #~ MsgBox(str(lista_supcat),'') ; return
###
#PweDGCategorie
        lista_cat = list()
        if CapCat.find('PweDGCategorie'):
            PweDGCategorie = CapCat.find('PweDGCategorie').getchildren()
            for elem in PweDGCategorie:
                id_sc = elem.get('ID')
                dessintetica = elem.find('DesSintetica').text
                try:
                    percentuale = elem.find('Percentuale').text
                except AttributeError:
                    percentuale = '0'
                cat =(id_sc, dessintetica, percentuale)
                lista_cat.append(cat)
            #~ MsgBox(str(lista_cat),'')
###
#PweDGSubCategorie
        lista_subcat = list()
        if CapCat.find('PweDGSubCategorie'):
            PweDGSubCategorie = CapCat.find('PweDGSubCategorie').getchildren()
            for elem in PweDGSubCategorie:
                id_sc = elem.get('ID')
                dessintetica = elem.find('DesSintetica').text
                try:
                    percentuale = elem.find('Percentuale').text
                except AttributeError:
                    percentuale = '0'
                subcat =(id_sc, dessintetica, percentuale)
                lista_subcat.append(subcat)
            #~ MsgBox(str(lista_subcat),'') ; return
    except AttributeError:
        pass
###
#PweDGWBS
    try:
        PweDGWBS = dati.find('PweDGWBS')
        pass
    except AttributeError:
        pass
###
    try:
        PweDGModuli = dati.getchildren()[2][0].getchildren()    #PweDGModuli
        speseutili = PweDGModuli[0].text
        spesegenerali = PweDGModuli[1].text
        utiliimpresa = PweDGModuli[2].text
        oneriaccessorisc = PweDGModuli[3].text
        ConfQuantita = PweDGModuli[4].text
    except IndexError:
        pass
###
    try:
        PweDGModuli = dati.getchildren()[2][0].getchildren()    #PweDGModuli
        speseutili = PweDGModuli[0].text
        spesegenerali = PweDGModuli[1].text
        utiliimpresa = PweDGModuli[2].text
        oneriaccessorisc = PweDGModuli[3].text
        ConfQuantita = PweDGModuli[4].text
    except IndexError:
        pass
###
    try:
        PweDGConfigurazione = dati.getchildren()[3][0].getchildren()    #PweDGConfigurazione
        Divisa = PweDGConfigurazione[0].text
        ConversioniIN = PweDGConfigurazione[1].text
        FattoreConversione = PweDGConfigurazione[2].text
        Cambio = PweDGConfigurazione[3].text
        PartiUguali = PweDGConfigurazione[4].text
        PartiUguali = PweDGConfigurazione[5].text
        Larghezza = PweDGConfigurazione[6].text
        HPeso = PweDGConfigurazione[7].text
        Quantita = PweDGConfigurazione[8].text
        Prezzi = PweDGConfigurazione[9].text
        PrezziTotale = PweDGConfigurazione[10].text
        ConvPrezzi= PweDGConfigurazione[11].text
        ConvPrezziTotale = PweDGConfigurazione[12].text
        IncidenzaPercentuale = PweDGConfigurazione[13].text
        Aliquote = PweDGConfigurazione[14].text
    except IndexError:
        pass
    
###
    misurazioni = root.find('PweMisurazioni')
    PweElencoPrezzi = misurazioni.getchildren()[0]
###
# leggo l'elenco prezzi ################################################
    epitems = PweElencoPrezzi.findall('EPItem')
    dict_articoli = dict()
    lista_articoli = list()
    lista_analisi = list()
    lista_tariffe_analisi = list()
    for elem in epitems:
        id_ep = elem.get('ID')
        diz_ep = dict()
        tipoep = elem.find('TipoEP').text
        if elem.find('Tariffa').text != None:
            tariffa = elem.find('Tariffa').text
        else:
            tariffa = ''
        articolo = elem.find('Articolo').text
        desridotta = elem.find('DesRidotta').text
        destestesa = elem.find('DesEstesa').text#.strip()
        try:
            desridotta = elem.find('DesBreve').text
        except AttributeError:
            pass
        try:
            desbreve = elem.find('DesBreve').text
        except AttributeError:
            desbreve = ''

        if elem.find('UnMisura').text != None:
            unmisura = elem.find('UnMisura').text
        else:
            unmisura = ''
        prezzo1 = elem.find('Prezzo1').text
        prezzo2 = elem.find('Prezzo2').text
        prezzo3 = elem.find('Prezzo3').text
        prezzo4 = elem.find('Prezzo4').text
        prezzo5 = elem.find('Prezzo5').text
        try:
            idspcap = elem.find('IDSpCap').text
        except AttributeError:
            idspcap = ''
        try:
            idcap = elem.find('IDCap').text
        except AttributeError:
            idcap = ''
        try:
            idsbcap = elem.find('IDSbCap').text
        except AttributeError:
            idsbcap = ''
        try:
            flags = elem.find('Flags').text
        except AttributeError:
            flags = ''
        try:
            data = elem.find('Data').text
        except AttributeError:
            data = ''
            
        xlo_sic = ''
        xlo_mdop = ''
        xlo_mdo = ''

        try:
            xlo_sic = float(elem.find('xlo_sic').text)
            xlo_mdop = float(elem.find('xlo_mdop').text)
            xlo_mdo = float(elem.find('xlo_mdo').text)
        except: # AttributeError TypeError:
            pass
        try:
            adrinternet = elem.find('AdrInternet').text
        except AttributeError:
            adrinternet = ''
        if elem.find('PweEPAnalisi').text == None:
            pweepanalisi = ''
        else:
            pweepanalisi = elem.find('PweEPAnalisi').text
        #~ chi(pweepanalisi)
        diz_ep['tipoep'] = tipoep
        diz_ep['tariffa'] = tariffa
        diz_ep['articolo'] = articolo
        diz_ep['desridotta'] = desridotta
        diz_ep['destestesa'] = destestesa
        diz_ep['desridotta'] = desridotta
        diz_ep['desbreve'] = desbreve
        diz_ep['unmisura'] = unmisura
        diz_ep['prezzo1'] = prezzo1
        diz_ep['prezzo2'] = prezzo2
        diz_ep['prezzo3'] = prezzo3
        diz_ep['prezzo4'] = prezzo4
        diz_ep['prezzo5'] = prezzo5
        diz_ep['idspcap'] = idspcap
        diz_ep['idcap'] = idcap
        diz_ep['flags'] = flags
        diz_ep['data'] = data
        diz_ep['adrinternet'] = adrinternet
        #~ diz_ep['pweepanalisi'] = pweepanalisi
        diz_ep['xlo_sic'] = xlo_sic
        diz_ep['xlo_mdop'] = xlo_mdop
        diz_ep['xlo_mdo'] = xlo_mdo
        dict_articoli[id_ep] = diz_ep
        articolo_modificato = (tariffa,
                                    destestesa,
                                    unmisura,
                                    xlo_sic,
                                    float(prezzo1),
                                    xlo_mdop,
                                    xlo_mdo)
        lista_articoli.append(articolo_modificato)
### leggo analisi di prezzo
        pweepanalisi = elem.find('PweEPAnalisi')
        PweEPAR = pweepanalisi.find('PweEPAR')
        if PweEPAR != None:
            EPARItem = PweEPAR.findall('EPARItem')
            analisi = list()
            for el in EPARItem:
                id_an = el.get('ID')
                an_tipo = el.find('Tipo').text
                id_ep = el.find('IDEP').text
                an_des = el.find('Descrizione').text
                an_um = el.find('Misura').text
                an_qt = el.find('Qt').text.replace(' ','')
                an_pr = el.find('Prezzo').text.replace(' ','')
                an_fld = el.find('FieldCTL').text
                an_rigo =(id_ep, an_des, an_um, an_qt, an_pr)
                analisi.append(an_rigo)
            lista_analisi.append([tariffa, destestesa, unmisura, analisi, prezzo1])
            lista_tariffe_analisi.append(tariffa)
# leggo voci di misurazione e righe ####################################
    lista_misure = list()
    try:
        PweVociComputo = misurazioni.getchildren()[1]
        vcitems = PweVociComputo.findall('VCItem')
        prova_l = list()
        for elem in vcitems:
            diz_misura = dict()
            id_vc = elem.get('ID')
            id_ep = elem.find('IDEP').text
            quantita = elem.find('Quantita').text
            try:
                datamis = elem.find('DataMis').text
            except AttributeError:
                datamis = ''
            try:
                flags = elem.find('Flags').text
            except AttributeError:
                flags = ''
            try:
                idspcat = elem.find('IDSpCat').text
            except AttributeError:
                idspcat = ''
            try:
                idcat = elem.find('IDCat').text
            except AttributeError:
                idcat = ''
            try:
                idsbcat = elem.find('IDSbCat').text
            except AttributeError:
                idsbcat = ''
            try:
                CodiceWBS = elem.find('CodiceWBS').text
            except AttributeError:
                CodiceWBS = ''
            righi_mis = elem.getchildren()[-1].findall('RGItem')
            lista_rig = list()
            riga_misura =()
            lista_righe = list()#[]
            new_id_l = list()

            for el in righi_mis:
                #~ diz_rig = dict()
                rgitem = el.get('ID')
                idvv = el.find('IDVV').text
                if el.find('Descrizione').text != None:
                    descrizione = el.find('Descrizione').text
                else:
                    descrizione = ''
                partiuguali = el.find('PartiUguali').text
                lunghezza = el.find('Lunghezza').text
                larghezza = el.find('Larghezza').text
                hpeso = el.find('HPeso').text
                quantita = el.find('Quantita').text
                flags = el.find('Flags').text
                riga_misura = (descrizione,
                                '',
                                '',
                                partiuguali,
                                lunghezza,
                                larghezza,
                                hpeso,
                                quantita,
                                flags,
                                idvv,
                                )
                mia = []
                mia.append(riga_misura[0])
                for el in riga_misura[1:]:
                    if el == None:
                        el = ''
                    else:
                        try:
                            el = float(el)
                        except ValueError:
                            if el != '':
                                el = '=' + el.replace('.',',')
                    mia.append(el)
                lista_righe.append(riga_misura)
            diz_misura['id_vc'] = id_vc
            diz_misura['id_ep'] = id_ep
            diz_misura['quantita'] = quantita
            diz_misura['datamis'] = datamis
            diz_misura['flags'] = flags
            diz_misura['idspcat'] = idspcat
            diz_misura['idcat'] = idcat
            diz_misura['idsbcat'] = idsbcat
            diz_misura['lista_rig'] = lista_righe

            new_id = strall(idspcat) +'.'+ strall(idcat) +'.'+ strall(idsbcat)
            new_id_l =(new_id, diz_misura)
            prova_l.append(new_id_l)
            lista_misure.append(diz_misura)
    except IndexError:
        MsgBox("""Nel file scelto non risultano esserci voci di misurazione,
perciò saranno importate le sole voci di Elenco Prezzi.

Si tenga conto che:
    - sarà importato solo il "Prezzo 1" dell'elenco;
    - il formato XPWE non conserva alcuni dati come
      le incidenze di sicurezza e di manodopera!""",'ATTENZIONE!')
        pass
    if len(lista_misure) != 0:
        if DlgSiNo("""Vuoi tentare un riordino delle voci secondo la stuttura delle Categorie?

    Scegliendo Sì, nel caso in cui il file di origine risulti particolarmente disordinato, riceverai un messaggio che ti indica come intervenire.

    Se il risultato finale non dovesse andar bene, puoi ripetere l'importazione senza il riordino delle voci rispondendo No a questa domanda.""", "Richiesta") ==2:
            riordine = sorted(prova_l, key=lambda el: el[0])
            lista_misure = list()
            for el in riordine:
                lista_misure.append(el[1])
    attesa().start()
###
# compilo Anagrafica generale ##########################################
    #~ New_file.computo()
# compilo Anagrafica generale ##########################################
    oSheet = oDoc.getSheets().getByName('S2')
    if oggetto != None:
        oSheet.getCellByPosition(2,2).String = oggetto
    if comune != None:
        oSheet.getCellByPosition(2,3).String = comune
    if committente != None:
        oSheet.getCellByPosition(2,5).String = committente
    if impresa != None:
        oSheet.getCellByPosition(3,16).String = impresa
###
    try:
        oSheet = oDoc.getSheets().getByName('S1')
        oSheet.getCellByPosition(7,318).Value = float(oneriaccessorisc)/100
        oSheet.getCellByPosition(7,319).Value = float(spesegenerali)/100
        oSheet.getCellByPosition(7,320).Value = float(utiliimpresa)/100
    except:
        pass
    oDoc.CurrentController.ZoomValue = 400

# compilo Elenco Prezzi ################################################
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    # Siccome setDataArray pretende una tupla(array 1D) o una tupla di tuple(array 2D)
    # trasformo la lista_articoli da una lista di tuple a una tupla di tuple
    lista_come_array = tuple(lista_articoli)
    # Parametrizzo il range di celle a seconda della dimensione della lista
    colonne_lista = len(lista_come_array[0]) # numero di colonne necessarie per ospitare i dati
    righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati

    oSheet.getRows().insertByIndex(3, righe_lista)
    oRange = oSheet.getCellRangeByPosition( 0,
                                            3,
                                            colonne_lista -1, # l'indice parte da 0
                                            righe_lista +3-1)
    oRange.setDataArray(lista_come_array)
    lrow = getLastUsedCell(oSheet).EndRow -1 
    oSheet.getCellRangeByPosition(0, 3, 0, lrow).CellStyle = "EP-aS"
    oSheet.getCellRangeByPosition(1, 3, 1, lrow).CellStyle = "EP-a"
    oSheet.getCellRangeByPosition(2, 3, 7, lrow).CellStyle = "EP-mezzo"
    oSheet.getCellRangeByPosition(5, 3, 5, lrow).CellStyle = "EP-mezzo %"
    oSheet.getCellRangeByPosition(8, 3, 9, lrow).CellStyle = "EP-sfondo"

    oSheet.getCellRangeByPosition(11, 3, 11, lrow).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition(12, 3, 12, lrow).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition(13, 3, 13, lrow).CellStyle = 'EP statistiche_Contab_q'
# aggiungo i capitoli alla lista delle voci ############################
    #~ giallo(16777072,16777120,16777168)
    #~ verde(9502608,13696976,15794160)
    #~ viola(12632319,13684991,15790335)
    col1 = 16777072
    col2 = 16777120
    col3 = 16777168
    capitoli = list()
# SUPERCAPITOLI
    try:
        for el in lista_supcap:
            tariffa = el.get('codice')
            if tariffa != None:
                destestesa = el.get('dessintetica')
                titolo = (tariffa,
                                        destestesa,
                                        '',
                                        '',
                                        '',
                                        '',
                                        '')
                capitoli.append(titolo)
        lista_come_array = tuple(capitoli)
        colonne_lista = len(lista_come_array[0]) # numero di colonne necessarie per ospitare i dati
        righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati

        oSheet.getRows().insertByIndex(3, righe_lista)
        oRange = oSheet.getCellRangeByPosition( 0,
                                                3,
                                                colonne_lista -1, # l'indice parte da 0
                                                righe_lista +3-1)
        oRange.setDataArray(lista_come_array)
        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista +3-1).CellStyle = "EP-aS"
        oSheet.getCellRangeByPosition(1, 3, 1, righe_lista +3-1).CellStyle = "EP-a"
        oSheet.getCellRangeByPosition(2, 3, 7, righe_lista +3-1).CellStyle = "EP-mezzo"
        oSheet.getCellRangeByPosition(5, 3, 5, righe_lista +3-1).CellStyle = "EP-mezzo %"
        oSheet.getCellRangeByPosition(8, 3, 9, righe_lista +3-1).CellStyle = "EP-sfondo"

        oSheet.getCellRangeByPosition(11, 3, 11, righe_lista +3-1).CellStyle = 'EP-mezzo %'
        oSheet.getCellRangeByPosition(12, 3, 12, righe_lista +3-1).CellStyle = 'EP statistiche_q'
        oSheet.getCellRangeByPosition(13, 3, 13, righe_lista +3-1).CellStyle = 'EP statistiche_Contab_q'
        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellBackColor = col1
    except:
        pass
# CAPITOLI
    capitoli = list()
    try:
        for el in lista_cap: # + lista_subcap:
            tariffa = el.get('codice')
            if tariffa != None:
                destestesa = el.get('dessintetica')
                titolo = (tariffa,
                                        destestesa,
                                        '',
                                        '',
                                        '',
                                        '',
                                        '')
                capitoli.append(titolo)
        lista_come_array = tuple(capitoli)
        colonne_lista = len(lista_come_array[0]) # numero di colonne necessarie per ospitare i dati
        righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati

        oSheet.getRows().insertByIndex(3, righe_lista)
        oRange = oSheet.getCellRangeByPosition( 0,
                                                3,
                                                colonne_lista -1, # l'indice parte da 0
                                                righe_lista +3-1)
        oRange.setDataArray(lista_come_array)
        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista +3-1).CellStyle = "EP-aS"
        oSheet.getCellRangeByPosition(1, 3, 1, righe_lista +3-1).CellStyle = "EP-a"
        oSheet.getCellRangeByPosition(2, 3, 7, righe_lista +3-1).CellStyle = "EP-mezzo"
        oSheet.getCellRangeByPosition(5, 3, 5, righe_lista +3-1).CellStyle = "EP-mezzo %"
        oSheet.getCellRangeByPosition(8, 3, 9, righe_lista +3-1).CellStyle = "EP-sfondo"

        oSheet.getCellRangeByPosition(11, 3, 11, righe_lista +3-1).CellStyle = 'EP-mezzo %'
        oSheet.getCellRangeByPosition(12, 3, 12, righe_lista +3-1).CellStyle = 'EP statistiche_q'
        oSheet.getCellRangeByPosition(13, 3, 13, righe_lista +3-1).CellStyle = 'EP statistiche_Contab_q'
        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellBackColor = col2
    except:
        pass
# SUBCAPITOLI
    capitoli = list()
    try:
        for el in lista_subcap:
            tariffa = el.get('codice')
            if tariffa != None:
                destestesa = el.get('dessintetica')
                titolo = (tariffa,
                                        destestesa,
                                        '',
                                        '',
                                        '',
                                        '',
                                        '')
                capitoli.append(titolo)
        lista_come_array = tuple(capitoli)
        colonne_lista = len(lista_come_array[0]) # numero di colonne necessarie per ospitare i dati
        righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati

        oSheet.getRows().insertByIndex(4, righe_lista)
        oRange = oSheet.getCellRangeByPosition( 0,
                                                3,
                                                colonne_lista -1, # l'indice parte da 0
                                                righe_lista +3-1)
        oRange.setDataArray(lista_come_array)
        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista +3-1).CellStyle = "EP-aS"
        oSheet.getCellRangeByPosition(1, 3, 1, righe_lista +3-1).CellStyle = "EP-a"
        oSheet.getCellRangeByPosition(2, 3, 7, righe_lista +3-1).CellStyle = "EP-mezzo"
        oSheet.getCellRangeByPosition(5, 3, 5, righe_lista +3-1).CellStyle = "EP-mezzo %"
        oSheet.getCellRangeByPosition(8, 3, 9, righe_lista +3-1).CellStyle = "EP-sfondo"

        oSheet.getCellRangeByPosition(11, 3, 11, righe_lista +3-1).CellStyle = 'EP-mezzo %'
        oSheet.getCellRangeByPosition(12, 3, 12, righe_lista +3-1).CellStyle = 'EP statistiche_q'
        oSheet.getCellRangeByPosition(13, 3, 13, righe_lista +3-1).CellStyle = 'EP statistiche_Contab_q'
        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellBackColor = col3
    except:
        pass
    for el in(11, 15, 19, 26):
        oSheet.getCellRangeByPosition(el, 3, el, ultima_voce(oSheet)).CellStyle = 'EP-mezzo %'
    for el in(12, 16, 20, 23):
        oSheet.getCellRangeByPosition(el, 3, el, ultima_voce(oSheet)).CellStyle = 'EP statistiche_q'
    for el in(13, 17, 21, 24, 25):
        oSheet.getCellRangeByPosition(el, 3, el, ultima_voce(oSheet)).CellStyle = 'EP statistiche'
    #~ adatta_altezza_riga('Elenco Prezzi')
    riordina_ElencoPrezzi()
    struttura_Elenco()

### elimino le voci che hanno analisi
    for i in reversed(range(3, getLastUsedCell(oSheet).EndRow)):
        if oSheet.getCellByPosition(0, i).String in lista_tariffe_analisi:
            oSheet.getRows().removeByIndex(i, 1)
    if len(lista_misure) == 0:
        #~ MsgBox('Importazione eseguita con successo in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!        \n\nImporto € ' + oSheet.getCellByPosition(0, 1).String ,'')
        MsgBox("Importate n."+ str(len(lista_articoli)) +" voci dall'elenco prezzi\ndel file: " + filename, 'Avviso')
        oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
        oDoc.CurrentController.setActiveSheet(oSheet)
        oDoc.CurrentController.ZoomValue = 100
        oDialogo_attesa.endExecute()
        return
###
# Compilo Analisi di prezzo ############################################
    #~ if len(lista_analisi) !=0:
    inizializza_analisi()
    if len(lista_analisi) !=0:
        oSheet = oDoc.getSheets().getByName('Analisi di Prezzo')
        for el in lista_analisi:
            sStRange = Circoscrive_Analisi(Range2Cell()[1])
            lrow = sStRange.RangeAddress.StartRow + 1
            oSheet.getCellByPosition(0, lrow).String = el[0]
            oSheet.getCellByPosition(1, lrow).String = el[1]
            oSheet.getCellByPosition(2, lrow).String = el[2]
            oSheet.getCellByPosition(6, lrow).Value = el[4]
            n = lrow + 2
            y = 0
            for x in el[3]:
                copia_riga_analisi(n)
                try:
                    oSheet.getCellByPosition(0, n).String = dict_articoli.get(el[3][y][0]).get('tariffa')
                except:
                    oSheet.getCellByPosition(0, n).String = '--'
                    oSheet.getCellByPosition(1, n).String = el[3][y][1]
                    oSheet.getCellByPosition(2, n).String = el[3][y][2]
                    oSheet.getCellByPosition(4, n).Value = el[3][y][4]
                    oSheet.getCellByPosition(8, n).Value = 0
                oSheet.getCellByPosition(3, n).Value = el[3][y][3]
                y += 1
                n += 1
            oSheet.getRows().removeByIndex(n, 3)
            oSheet.getCellByPosition(0, n+2).String = ''
            oSheet.getCellByPosition(0, n+5).String = ''
            oSheet.getCellByPosition(0, n+8).String = ''
            oSheet.getCellByPosition(0, n+11).String = ''
            inizializza_analisi()
    #~ #basic_LeenO('Voci_Sposta.elimina_voce') #rinvia a basic
    tante_analisi_in_ep()
# Inserisco i dati nel COMPUTO #########################################
    if arg == 'VARIANTE':
        basic_LeenO('Computo.genera_variante')
    
    oDoc.CurrentController.ZoomValue = 400
    oSheet = oDoc.getSheets().getByName(arg)
    if oSheet.getCellByPosition(1, 4).String == 'Cod. Art.?':
        oSheet.getRows().removeByIndex(3, 4)
    oDoc.CurrentController.select(oSheet)
    iSheet_num = oSheet.RangeAddress.Sheet
###
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = oSheet.RangeAddress.Sheet # recupero l'index del foglio
    diz_vv = dict()

    testspcat = '0'
    testcat = '0'
    testsbcat = '0'
    x = 1
    for el in lista_misure:
        idspcat = el.get('idspcat')
        idcat = el.get('idcat')
        idsbcat = el.get('idsbcat')

        lrow = ultima_voce(oSheet) + 1
#~ inserisco le categorie
        try:
            if idspcat != testspcat:
                testspcat = idspcat
                testcat = '0'
                Inser_SuperCapitolo_arg(lrow, lista_supcat[eval(idspcat)-1][1])
                lrow = lrow + 2
        except UnboundLocalError:
            pass
        try:
            if idcat != testcat:
                testcat = idcat
                testsbcat = '0'
                Inser_Capitolo_arg(lrow, lista_cat[eval(idcat)-1][1])
                lrow = lrow + 2
        except UnboundLocalError:
            pass
        try:
            if idsbcat != testsbcat:
                testsbcat = idsbcat
                
                Inser_SottoCapitolo_arg(lrow, lista_subcat[eval(idsbcat)-1][1])
        except UnboundLocalError:
            pass
        lrow = ultima_voce(oSheet) + 1
        ins_voce_computo_grezza(lrow)
        ID = el.get('id_ep')
        id_vc = el.get('id_vc')

        try:
            oSheet.getCellByPosition(1, lrow+1).String = dict_articoli.get(ID).get('tariffa')
        except:
            pass

        diz_vv[id_vc] = lrow+1
        oSheet.getCellByPosition(0, lrow+1).String = str(x)
        x = x+1
        SC = 2
        SR = lrow + 2 + 1
        nrighe = len(el.get('lista_rig')) - 1

        #~ chi(el.get('lista_rig'))
        #~ return
        if nrighe > -1:
            EC = SC + len(el.get('lista_rig')[0])
            ER = SR + nrighe

            if nrighe > 0:
                oSheet.getRows().insertByIndex(SR, nrighe)

            oRangeAddress = oSheet.getCellRangeByPosition(0, SR-1, 250, SR-1).getRangeAddress()

            for n in range(SR, SR+nrighe):
                oCellAddress = oSheet.getCellByPosition(0, n).getCellAddress()
                oSheet.copyRange(oCellAddress, oRangeAddress)

            oCellRangeAddr.StartColumn = SC
            oCellRangeAddr.StartRow = SR
            oCellRangeAddr.EndColumn = EC
            oCellRangeAddr.EndRow = ER

        ###
        # INSERISCO PRIMA SOLO LE RIGHE SE NO MI FA CASINO

    # metodo veloce, ma ignora le formule
    # va bene se lista_righe viene convertito come tupla
            SR = SR - 1
            for mis in el.get('lista_rig'):
                if mis[0] != None: #descrizione
                    descrizione = mis[0].strip()
                    oSheet.getCellByPosition(2, SR).String = descrizione
                else:
                    descrizione =''

                if mis[3] != None: #parti uguali
                    try:
                        oSheet.getCellByPosition(5, SR).Value = float(mis[3].replace(',','.'))
                    except ValueError:
                        oSheet.getCellByPosition(5, SR).Formula = '=' + str(mis[3]).split('=')[-1] # tolgo evenutali '=' in eccesso
                if mis[4] != None: #lunghezza
                    try:
                        oSheet.getCellByPosition(6, SR).Value = float(mis[4].replace(',','.'))
                    except ValueError:
                        oSheet.getCellByPosition(6, SR).Formula = '=' + str(mis[4]).split('=')[-1] # tolgo evenutali '=' in eccesso
                if mis[5] != None: #larghezza
                    try:
                        oSheet.getCellByPosition(7, SR).Value = float(mis[5].replace(',','.'))
                    except ValueError:
                        oSheet.getCellByPosition(7, SR).Formula = '=' + str(mis[5]).split('=')[-1] # tolgo evenutali '=' in eccesso
                if mis[6] != None: #HPESO
                    try:
                        oSheet.getCellByPosition(8, SR).Value = float(mis[6].replace(',','.'))
                        
                    except:
                        oSheet.getCellByPosition(8, SR).Formula = '=' + str(mis[6]).split('=')[-1] # tolgo evenutali '=' in eccesso
                if mis[8] == '2':
                    parziale_core(SR)
                    oSheet.getRows().removeByIndex(SR+1, 1)
                    descrizione =''
                    
                if '-' in mis[7]:
                    for x in range(5, 8):
                        try:
                            if oSheet.getCellByPosition(x, SR).Value != 0:
                                oSheet.getCellByPosition(x, SR).Value = abs(oSheet.getCellByPosition(x, SR).Value)
                        except:
                            pass
                    oSheet.getCellByPosition(9, SR).Formula = '=IF(PRODUCT(F' + str(SR+1) + ':I' + str(SR+1) + ')=0;"";-PRODUCT(F' + str(SR+1) + ':I' + str(SR+1) + '))'

                if oSheet.getCellByPosition(5, SR).Type.value == 'FORMULA':
                    va = oSheet.getCellByPosition(5, SR).Formula
                else:
                    va = oSheet.getCellByPosition(5, SR).Value

                if oSheet.getCellByPosition(6, SR).Type.value == 'FORMULA':
                    vb = oSheet.getCellByPosition(6, SR).Formula
                else:
                    vb = oSheet.getCellByPosition(6, SR).Value
                    
                if oSheet.getCellByPosition(7, SR).Type.value == 'FORMULA':
                    vc = oSheet.getCellByPosition(7, SR).Formula
                else:
                    vc = oSheet.getCellByPosition(7, SR).Value

                if oSheet.getCellByPosition(8, SR).Type.value == 'FORMULA':
                    vd = oSheet.getCellByPosition(8, SR).Formula
                else:
                    vd = oSheet.getCellByPosition(8, SR).Value

                if mis[3] == None:
                    va =''
                else:
                    if '^' in mis[3]:
                        va = eval(mis[3].replace('^','**'))
                    else:
                        va = eval(mis[3])
                lista_n = list()
                if mis[9] != '-2':
                    for el in (va, vb, vc, vd):
                        if el != 0 : lista_n.append(el)
                    vedi = diz_vv.get(mis[9])
                    try:
                        vedi_voce_xpwe(SR, vedi, mis[8])
                    except:
                        MsgBox("""Il file di origine è particolarmente disordinato.
Riordinando il computo trovo riferimenti a voci non ancora inserite.

Al termine dell'impotazione controlla la voce con tariffa """ + dict_articoli.get(ID).get('tariffa') +
"""\nalla riga n.""" + str(lrow+2) + """ del foglio, evidenziata qui a sinistra.""", 'Attenzione!')
                    x = 0
                    if len(lista_n) != 0:
                        for n in lista_n:
                            try: 
                                float(n)
                                oSheet.getCellByPosition(8-x, SR).Value = n
                            except:
                                oSheet.getCellByPosition(8-x, SR).Formula = n
                            x +=1
                SR = SR+1
    numera_voci()

    try:
        Rinumera_TUTTI_Capitoli2()
    except:
        pass
    oDoc.CurrentController.ZoomValue = 100
    refresh(1)
    #~ MsgBox('Importazione eseguita con successo in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!        \n\nImporto € ' + oSheet.getCellByPosition(0, 1).String ,'')
    oDialogo_attesa.endExecute()
    doppioni()
    MsgBox('Importazione eseguita con successo!','')

# XPWE_in ##########################################################
########################################################################
#VARIABILI GLOBALI:#####################################################
########################################################################
Lmajor= 3 #'INCOMPATIBILITA'
Lminor= 17 #'NUOVE FUNZIONALITA'
Lsubv= "2.dev" #'CORREZIONE BUGS
noVoce =('Livello-0-scritta', 'Livello-1-scritta', 'livello2 valuta', 'comp Int_colonna', 'Ultimus_centro_bordi_lati')
stili_computo =('Comp Start Attributo', 'comp progress', 'comp 10 s','Comp End Attributo')
stili_contab =('Comp Start Attributo_R', 'comp 10 s_R','Comp End Attributo_R')
stili_analisi =('An.1v-Att Start', 'An-1_sigla', 'An-lavoraz-desc', 'An-lavoraz-Cod-sx', 'An-lavoraz-desc-CEN', 'An-sfondo-basso Att End')
stili_elenco =('EP-Cs', 'EP-aS')
createUnoService =(
        XSCRIPTCONTEXT
        .getComponentContext()
        .getServiceManager()
        .createInstance
                    )
GetmyToolBarNames =('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar',
    'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ELENCO',
    'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ANALISI',
    'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_COMPUTO',
    'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CONTABILITA',)
#
sUltimus = ''
########################################################################
def ssUltimus(arg=None):
    oDlgMain.endExecute()
    '''
    Scrive la variabile globale che individua il Documento di Contabilità Corrente(DCC)
    che è il file a cui giungono le voci di prezzo inviate da altri file
    '''
    global sUltimus
    oDoc = XSCRIPTCONTEXT.getDocument()
    if oDoc.getSheets().hasByName('M1') == False:
        return
    if len(oDoc.getURL()) == 0:
        MsgBox('''Prima di procedere, devi salvare il lavoro!
Provvedi subito a dare un nome al file di computo...''', 'Dai un nome al file...')
        salva_come()
        autoexec()
    try:
        sUltimus = uno.fileUrlToSystemPath(oDoc.getURL())
    except:
        return
    oSheet = oDoc.getSheets().getByName('M1')
    
    oSheet.getCellByPosition(2,27).String = sUltimus
    DlgMain()
    return
########################################################################
def debugnn(sCella='', t=''):
    mri(XSCRIPTCONTEXT.getDocument())
    '''
    sCella  { string } : stringa di default nella casella di testo
    t       { string } : titolo del dialogo
    Viasualizza un dialogo di richiesta testo
    '''
    chi(sys)
    return
    psm = uno.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDialog1 = dp.createDialog("vnd.sun.star.script:UltimusFree2.DlgFile?language=Basic&location=application")
    oDialog1Model = oDialog1.Model

    oDialog1Model.Title = t

    oString = oDialog1.getControl("FileControl1")
    chi(oString)
    oString.Text = sCella
    #~ oDialog1.execute()
    if oDialog1.execute()==0:
        return
    else:
        MsgBox(oString.Text)
        return oString.Text

########################################################################
def filedia(titolo='Scegli il file...', est='*.*', mode=0):
    """
    titolo  { string }  : titolo del FilePicker
    est     { string }  : filtro di visualizzazione file
    mode    { integer } : modalità di gestione del file

    Apri file:  `mode in(0, 6, 7, 8, 9)`
    Salva file: `mode in(1, 2, 3, 4, 5, 10)`
    see:('''http://api.libreoffice.org/docs/idl/ref/
            namespacecom_1_1sun_1_1star_1_1ui_1_1
            dialogs_1_1TemplateDescription.html''' )
    see:('''http://stackoverflow.com/questions/30840736/
        libreoffice-how-to-create-a-file-dialog-via-python-macro''')
    """
    estensioni = {'*.*'   : 'Tutti i file(*.*)',
                '*.odt' : 'Writer(*.odt)',
                '*.ods' : 'Calc(*.ods)',
                '*.odb' : 'Base(*.odb)',
                '*.odg' : 'Draw(*.odg)',
                '*.odp' : 'Impress(*.odp)',
                '*.odf' : 'Math(*.odf)',
                '*.xpwe': 'Primus(*.xpwe)',
                '*.xml' : 'XML(*.xml)',
                '*.dat' : 'dat(*.dat)',
                }
    try:
        oFilePicker = createUnoService( "com.sun.star.ui.dialogs.OfficeFilePicker" )
        oFilePicker.initialize(( mode,) )
        oFilePicker.Title = titolo

        app = estensioni.get(est)
        oFilePicker.appendFilter(app, est)
        if oFilePicker.execute():
            oDisp = uno.fileUrlToSystemPath(oFilePicker.getFiles()[0])
        return oDisp
    except:
        MsgBox('Il file non è stato selezionato', 'ATTENZIONE!')
        return

########################################################################
import traceback
from com.sun.star.awt import Rectangle

def filedia_(titolo=''):
# http://openoffice3.web.fc2.com/Python_Macro_Calc.html#OOoCCB01 #
    try:
        oCtx = uno.getComponentContext()
        oServiceManager = oCtx.ServiceManager
        oDlgModel = oServiceManager.createInstance('com.sun.star.awt.UnoControlDialogModel')
        # Size of Dialog
        oDlgWth = 200
        oDlgHgt = 75
        oDlgModel.Width = oDlgWth
        oDlgModel.Height = oDlgHgt
        oDlgModel.PositionX = 150
        oDlgModel.PositionY = 200
        oDlgModel.BackgroundColor = 0xafd2fc
        # Title of Dialog
        oDlgModel.Title = titolo #'Scegli il file da convertire...'
        #
        # ***** [ OK / Cancel  Button 設定 ] *****
        # OK Button 仕様
        oModel = oDlgModel.createInstance('com.sun.star.awt.UnoControlButtonModel')
        oTabIndex = 0
        oModel.Name = 'OkBtn'
        oModel.TabIndex = oTabIndex
        oModel.PositionX = oDlgWth/2 - 50
        oModel.PositionY = oDlgHgt - 20
        oModel.Width = 40
        oModel.Height = 15
        oModel.Label = u'OK'
        oModel.PushButtonType = 1       # 1 : OK
        # Dialog Modelの仕様に Step Button の仕様を設定
        oDlgModel.insertByName('OkBtn', oModel)
        # ***** [ OK / Cancel  Button 設定 ] *****
        #
        # ***** [ OK / Cancel  Button 設定 ] *****
        # OK Button 仕様
        oModel = oDlgModel.createInstance('com.sun.star.awt.UnoControlButtonModel')
        oTabIndex = 0
        oModel.Name = 'AnnullaBtn'
        oModel.TabIndex = oTabIndex
        oModel.PositionX = oDlgWth/2 + 10
        oModel.PositionY = oDlgHgt - 20
        oModel.Width = 40
        oModel.Height = 15
        oModel.Label = u'Annulla'
        oModel.PushButtonType = 2       # 1 : CANCEL
        # Dialog Modelの仕様に Step Button の仕様を設定
        oDlgModel.insertByName('AnnullaBtn', oModel)
        # ***** [ OK / Cancel  Button Setting ] *****
        #
        # ***** [ FileCntrol Setting ] *****
        # FileCntrol specifiche
        oModel = oDlgModel.createInstance('com.sun.star.awt.UnoControlFileControlModel')
        oTabIndex = oTabIndex + 1
        oModel.Name = 'FileCtrl'
        oModel.TabIndex = oTabIndex
        oModel.PositionX = 10
        oModel.PositionY = oDlgHgt - 60
        oModel.Width = oDlgWth - 20
        oModel.Height = 15
        oModel.Enabled = 1
        oModel.Border = 1
        # Dialog Model FileCntrol
        oDlgModel.insertByName('FileCtrl', oModel)
        # ***** [ FileCntrol 設定 ] *****
        #
        # Create the dialog and set the model
        oDlg = oServiceManager.createInstance('com.sun.star.awt.UnoControlDialog')
        oDlg.setModel(oDlgModel)
        #
        # Create a window and then tell the dialog to use the created window.
        oWindow = oServiceManager.createInstance('com.sun.star.awt.Toolkit')
        oDlg.createPeer(oWindow,None)        # None : OK / none : NG
        #
        # Dialogの表示実行
        oClick = oDlg.execute()
        if oClick == 1:
            oSelFile = oDlgModel.getByName('FileCtrl')
            oDisp = oSelFile.Text
        else:
            oDisp = u'Cancel' # 'Cancel' è il risultato del tasto
        # End Dialog
        oDlg.endExecute()
    except:
        oDisp = traceback.format_exc(sys.exc_info()[2])
    finally:
        return oDisp

########################################################################
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_OK, DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, DEFAULT_BUTTON_YES, DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

#rif.: https://wiki.openoffice.org/wiki/PythonDialogBox

def chi(s): # s = oggetto
    '''
    s    { object }  : oggetto da interrogare

    mostra un dialog che indica il tipo di oggetto ed i metodi ad esso applicabili
    '''
    doc = XSCRIPTCONTEXT.getDocument()
    parentwin = doc.CurrentController.Frame.ContainerWindow
    s1 = str(s) + '\n\n' + str(dir(s).__str__())
    MessageBox(parentwin, str(s1), str(type(s)), 'infobox')
    
def DlgSiNo(s,t='Titolo'): # s = messaggio | t = titolo
    '''
    Visualizza il menù di scelta sì/no
    restituisce 2 per sì e 3 per no
    '''
    doc = XSCRIPTCONTEXT.getDocument()
    parentwin = doc.CurrentController.Frame.ContainerWindow
    #~ s = 'This a message'
    #~ t = 'Title of the box'
    #~ MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
    return MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO + DEFAULT_BUTTON_NO)
    

def MsgBox(s,t=''): # s = messaggio | t = titolo
    doc = XSCRIPTCONTEXT.getDocument()
    parentwin = doc.CurrentController.Frame.ContainerWindow
    #~ s = 'This a message'
    #~ t = 'Title of the box'
    #~ res = MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO_CANCEL + DEFAULT_BUTTON_NO)
    #~ chi(res)
    #~ return
    #~ s = res
    #~ t = 'Titolo'
    if t == None:
        t='messaggio'
    MessageBox(parentwin, str(s), t, 'infobox')

# Show a message box with the UNO based toolkit
def MessageBox(ParentWin, MsgText, MsgTitle, MsgType=MESSAGEBOX, MsgButtons=BUTTONS_OK):
    ctx = uno.getComponentContext()
    sm = ctx.ServiceManager
    sv = sm.createInstanceWithContext('com.sun.star.awt.Toolkit', ctx)
    myBox = sv.createMessageBox(ParentWin, MsgType, MsgButtons, MsgTitle, MsgText)
    return myBox.execute()
# [　入手元　]

def mri(target):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    mri = ctx.ServiceManager.createInstanceWithContext('mytools.Mri',ctx)
    mri.inspect(target)
    MsgBox('MRI in corso...','avviso')

########################################################################
#import pdb; pdb.set_trace() #debugger
########################################################################
#codice di Manuele Pesenti #############################################
########################################################################
def getFormula(n, a, b):
    """
    n { integer } : posizione cella
    a  { string } : primo parametro da sostituire
    b  { string } : secondo parametro da sostituire
    """
    v = dict(n=n, a=a, b=b)
    formulas = {
        18: '=SUBTOTAL(9;S%(a)s:S%(b)s)',
        24: '=S%(a)s/S%(b)s',
        29: '=AE%(a)s/S%(b)s',
        30: '=SUBTOTAL(9;AE%(a)s:AE%(b)s)'
    }
    return formulas[n] % v

def getCellStyle(l, n):
    """
    l { integer } : livello(1 o 2)
    n { integer } : posizione cella
    """
    styles = {
        2: {
            18: 'livello2 scritta mini',
            24: 'livello2 valuta mini %',
            29: 'livello2 valuta mini %',
            30: 'livello2 valuta mini'
        },
        1: {
            18: 'Livello-1-scritta mini val',
            24: 'Livello-1-scritta mini %',
            29: 'Livello-1-scritta mini %',
            30: 'Livello-1-scritta mini val'
        }
    }
    return styles[l][n]

def SubSum(lrow, sub=False):
    """ Inserisce i dati nella riga
    sub { boolean } : specifica se sotto-categoria
    """
    if sub:
        myrange =('livello2 scritta mini', 'Livello-1-scritta minival', 'Comp TOTALI',)
        level = 2
    else:
        myrange =('Livello-1-scritta mini val', 'Comp TOTALI',)
        level = 1

    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in('COMPUTO', 'VARIANTE'):
        return
    lrowE = ultima_voce(oSheet)+1
    nextCap = lrowE
    for n in range(lrow+1, lrowE):
        if oSheet.getCellByPosition(18, n).CellStyle in myrange:
            nextCap = n + 1
            break
    for n,a,b in((18, lrow+1, nextCap,),(24, lrow+1, lrowE+1,),(29, lrow+1, lrowE+1,),(30, lrow+1, nextCap,),):
        oSheet.getCellByPosition(n, lrow).Formula = getFormula(n, a, b)
        Sheet.getCellByPosition(18, lrow).CellStyle = getCellStyle(level, n)
########################################################################
# GESTIONE DELLE VISTE IN STRUTTURA ####################################
########################################################################
def filtra_codice(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    lrow = Range2Cell()[1]
    myrange =('Comp End Attributo', 'Comp TOTALI',)
    if oSheet.getCellByPosition(0, lrow).CellStyle in(stili_computo + stili_contab) :
        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        sStRange = Circoscrive_Voce_Computo_Att(lrow)
        sopra = sStRange.RangeAddress.StartRow
        voce = oSheet.getCellByPosition(1, sopra+1).String
    else:
        MsgBox('Devi prima selezionare una voce di misurazione.','Avviso!')
        return
    fine = ultima_voce(oSheet)+1
    lista_pt = list()
    _gotoCella(0, 0)

    for n in range(0, fine):
        if oSheet.getCellByPosition(0, n).CellStyle in('Comp Start Attributo','Comp Start Attributo_R'):
            sStRange = Circoscrive_Voce_Computo_Att(n)
            sopra = sStRange.RangeAddress.StartRow
            sotto = sStRange.RangeAddress.EndRow
            if oSheet.getCellByPosition(1, sopra+1).String != voce:
                lista_pt.append((sopra, sotto))
                #~ lista_pt.append((sopra+2, sotto-1))
    for el in lista_pt:
        oCellRangeAddr.StartRow = el[0]
        oCellRangeAddr.EndRow = el[1]
        oSheet.group(oCellRangeAddr,1)
        oSheet.getCellRangeByPosition(0, el[0], 0, el[1]).Rows.IsVisible=False
    _gotoCella(0, lrow)
    MsgBox('Filtro attivato in base al codice!','Codice voce: ' + voce)

def struttura_ComputoM(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    Rinumera_TUTTI_Capitoli2()
    struct(0)
    struct(1)
    struct(2)
    struct(3)

def struttura_Analisi(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    struct(4)

def struct(l):
    ''' mette in vista struttura secondo categorie
    l { integer } : specifica il livello di categoria
    ### COMPUTO/VARIANTE ###
    0 = super-categoria
    1 = categoria
    2 = sotto-categoria
    3 = intera voce di misurazione
    ### ANALISI ###
    4 = simile all'elenco prezzi
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet

    if l == 0:
        stile = 'Livello-0-scritta'
        myrange =('Livello-0-scritta', 'Comp TOTALI',)
        Dsopra = 1
        Dsotto = 1
    elif l == 1:
        stile = 'Livello-1-scritta'
        myrange =('Livello-1-scritta', 'Livello-0-scritta', 'Comp TOTALI',)
        Dsopra = 1
        Dsotto = 1
    elif l == 2:
        stile = 'livello2 valuta'
        myrange =('livello2 valuta','Livello-1-scritta', 'Livello-0-scritta', 'Comp TOTALI',)
        Dsopra = 1
        Dsotto = 1
    elif l == 3:
        stile = 'Comp Start Attributo'
        myrange =('Comp End Attributo', 'Comp TOTALI',)
        Dsopra = 2
        Dsotto = 1

    elif l == 4: #Analisi di Prezzo
        stile = 'An-1_sigla'
        myrange =('An.1v-Att Start', 'Analisi_Sfondo',)
        Dsopra = 1
        Dsotto = -1
        for n in(3, 5, 7):
            oCellRangeAddr.StartColumn = n
            oCellRangeAddr.EndColumn = n
            oSheet.group(oCellRangeAddr,0)
            oSheet.getCellRangeByPosition(n, 0, n, 0).Columns.IsVisible=False

    test = ultima_voce(oSheet)+2
    lista_cat = list()
    for n in range(0, test):
        if oSheet.getCellByPosition(0, n).CellStyle == stile:
            sopra = n+Dsopra
            for n in range(sopra+1, test):
                if oSheet.getCellByPosition(0, n).CellStyle in myrange:
                    sotto = n-Dsotto
                    lista_cat.append((sopra, sotto))
                    break
    for el in lista_cat:
        oCellRangeAddr.StartRow = el[0]
        oCellRangeAddr.EndRow = el[1]
        oSheet.group(oCellRangeAddr,1)
        oSheet.getCellRangeByPosition(0, el[0], 0, el[1]).Rows.IsVisible=False
########################################################################
def autoexec_off(arg=None):
    bak_timestamp()
    toolbar_switch(1)
    #~ private:resource/toolbar/standardbar
    sUltimus = ''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('M1')
    oSheet.getCellByPosition(2,27).String = ''#sUltimus

def autoexec(arg=None):
    '''
    questa è richiamata da New_File()
    '''
    ctx = XSCRIPTCONTEXT.getComponentContext()
    oGSheetSettings = ctx.ServiceManager.createInstanceWithContext("com.sun.star.sheet.GlobalSheetSettings", ctx)
    oGSheetSettings.UsePrinterMetrics = True #Usa i parametri della stampante per la formattazione del testo
#Crea ed imposta leeno.conf SOLO SE NON PRESENTE.
    if sys.platform == 'win32':
        path = os.getenv("HOMEDRIVE") + os.getenv("HOMEPATH")
    else:
        path = os.getenv("HOME")
    if not os.path.exists(path_conf):
        os.makedirs(path_conf[:-11])
        config_default()
    try:
        if conf.read(path_conf, 'Generale', 'movedirection') == '0':
            oGSheetSettings.MoveDirection = 0
        else:
            oGSheetSettings.MoveDirection = 1
    except:
        config_default()
    oDoc = XSCRIPTCONTEXT.getDocument()
#~ RegularExpressions and Wildcards are mutually exclusive, only one can have the value TRUE.
#~ If both are set to TRUE via API calls then the last one set takes precedence.
    oDoc.Wildcards = False
    oDoc.RegularExpressions = True
    try:
        oSheet = oDoc.getSheets().getByName('S1')
        oSheet.getCellByPosition(7, 290).Value = oDoc.getDocumentProperties().getUserDefinedProperties().Versione
        oSheet.getCellByPosition(7,193).Value = r_version_code()
        oSheet.getCellByPosition(8,193).Value = Lmajor
        oSheet.getCellByPosition(9,193).Value = Lminor
        oSheet.getCellByPosition(10,193).String = Lsubv
        oSheet.getCellByPosition(7,295).Value = Lmajor
        oSheet.getCellByPosition(8,295).Value = Lminor
        oSheet.getCellByPosition(9,295).String = Lsubv
        adegua_tmpl() #esegue degli aggiustamenti del template
        toolbar_vedi()
    except:
        #~ chi("autoexec py")
        return
# scegli cosa visualizzare all'avvio:
    vedi = conf.read(path_conf, 'Generale', 'visualizza')
    if vedi == 'Menù Principale':
        DlgMain()
    elif vedi == 'Dati Generali':
        Vai_a_Variabili()
    elif vedi in('Elenco Prezzi', 'COMPUTO'):
        _gotoSheet(vedi)
#
########################################################################
def computo_terra_terra(arg=None):
    '''
    Settaggio base di configuazione colonne in COMPUTO e VARIANTE
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.getCellRangeByPosition(33,0,1023,0).Columns.IsVisible = False
    set_larghezza_colonne()
########################################################################
def viste_nuove(sValori):
    '''
    sValori { string } : una tringa di configurazione della visibilità colonne
    permette di visualizzare/nascondere un set di colonne
    T = visualizza
    F = nasconde
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    n = 0
    for el in sValori:
        if el == 'T':
            oSheet.getCellByPosition(n, 2).Columns.IsVisible = True
        elif el == 'F':
            oSheet.getCellByPosition(n, 2).Columns.IsVisible = False
        n += 1
########################################################################
def set_larghezza_colonne(arg=None):
    '''
    regola la larghezza delle colonne a seconda della sheet
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name == 'Analisi di Prezzo':
        oSheet.getColumns().getByName('A').Columns.Width = 2100
        oSheet.getColumns().getByName('B').Columns.Width = 12000
        oSheet.getColumns().getByName('C').Columns.Width = 1600 
        oSheet.getColumns().getByName('D').Columns.Width = 2000
        oSheet.getColumns().getByName('E').Columns.Width = 3400
        oSheet.getColumns().getByName('F').Columns.Width = 3400
        oSheet.getColumns().getByName('G').Columns.Width = 2700
        oSheet.getColumns().getByName('H').Columns.Width = 2700
        oSheet.getColumns().getByName('I').Columns.Width = 2000
        oSheet.getColumns().getByName('J').Columns.Width = 2000
        oSheet.getColumns().getByName('K').Columns.Width = 2000
        oDoc.CurrentController.freezeAtPosition(0, 2)
    if oSheet.Name == 'CONTABILITA':
        viste_nuove('TTTFFTTTTTFTFTFTFTFTTFTTFTFTTTTFFFFFF')
        oSheet.getCellRangeByPosition(13,0,1023,0).Columns.Width = 1900 # larghezza colonne importi
        oSheet.getCellRangeByPosition(19,0,23,0).Columns.Width = 1000 # larghezza colonne importi
        oSheet.getCellRangeByPosition(51,0,1023,0).Columns.IsVisible = False # nascondi colonne
        oSheet.getColumns().getByName('A').Columns.Width = 600
        oSheet.getColumns().getByName('B').Columns.Width = 1500
        oSheet.getColumns().getByName('C').Columns.Width = 6300 #7800
        oSheet.getColumns().getByName('F').Columns.Width = 1300
        oSheet.getColumns().getByName('G').Columns.Width = 1300
        oSheet.getColumns().getByName('H').Columns.Width = 1300
        oSheet.getColumns().getByName('I').Columns.Width = 1300
        oSheet.getColumns().getByName('J').Columns.Width = 1700
        oSheet.getColumns().getByName('L').Columns.Width = 1700
        oSheet.getColumns().getByName('N').Columns.Width = 1900
        oSheet.getColumns().getByName('P').Columns.Width = 1900
        oSheet.getColumns().getByName('T').Columns.Width = 1000
        oSheet.getColumns().getByName('U').Columns.Width = 1000
        oSheet.getColumns().getByName('W').Columns.Width = 1000
        oSheet.getColumns().getByName('X').Columns.Width = 1000
        oSheet.getColumns().getByName('Z').Columns.Width = 1900
        oSheet.getColumns().getByName('AC').Columns.Width = 1700
        oSheet.getColumns().getByName('AD').Columns.Width = 1700
        oSheet.getColumns().getByName('AE').Columns.Width = 1700
        oSheet.getColumns().getByName('AX').Columns.Width = 1900
        oSheet.getColumns().getByName('AY').Columns.Width = 1900
        oDoc.CurrentController.freezeAtPosition(0, 3)
    if oSheet.Name in('COMPUTO', 'VARIANTE'):
        oSheet.getColumns().getByName('A').Columns.Width = 600
        oSheet.getColumns().getByName('B').Columns.Width = 1500
        oSheet.getColumns().getByName('C').Columns.Width = 6300 #7800
        oSheet.getColumns().getByName('F').Columns.Width = 1500
        oSheet.getColumns().getByName('G').Columns.Width = 1300
        oSheet.getColumns().getByName('H').Columns.Width = 1300
        oSheet.getColumns().getByName('I').Columns.Width = 1300
        oSheet.getColumns().getByName('J').Columns.Width = 1700
        oSheet.getColumns().getByName('L').Columns.Width = 1700
        oSheet.getColumns().getByName('S').Columns.Width = 1700
        oSheet.getColumns().getByName('AC').Columns.Width = 1700
        oSheet.getColumns().getByName('AD').Columns.Width = 1700
        oSheet.getColumns().getByName('AE').Columns.Width = 1700
        oDoc.CurrentController.freezeAtPosition(0, 3)
        viste_nuove('TTTFFTTTTTFTFFFFFFTFFFFFFFFFFFFFFFFFFFFFFFFFTT')
    if oSheet.Name == 'Elenco Prezzi':
        oSheet.getColumns().getByName('A').Columns.Width = 1600
        oSheet.getColumns().getByName('B').Columns.Width = 10000
        oSheet.getColumns().getByName('C').Columns.Width = 1500
        oSheet.getColumns().getByName('D').Columns.Width = 1500
        oSheet.getColumns().getByName('E').Columns.Width = 1600
        oSheet.getColumns().getByName('F').Columns.Width = 1500
        oSheet.getColumns().getByName('G').Columns.Width = 1500
        oSheet.getColumns().getByName('H').Columns.Width = 1600
        oSheet.getColumns().getByName('I').Columns.Width = 1200
        oSheet.getColumns().getByName('J').Columns.Width = 1200
        oSheet.getColumns().getByName('K').Columns.Width = 100
        oSheet.getColumns().getByName('L').Columns.Width = 1600
        oSheet.getColumns().getByName('M').Columns.Width = 1600
        oSheet.getColumns().getByName('N').Columns.Width = 1600
        oSheet.getColumns().getByName('O').Columns.Width = 100
        oSheet.getColumns().getByName('P').Columns.Width = 1600
        oSheet.getColumns().getByName('Q').Columns.Width = 1600
        oSheet.getColumns().getByName('R').Columns.Width = 1600
        oSheet.getColumns().getByName('S').Columns.Width = 100
        oSheet.getColumns().getByName('T').Columns.Width = 1600
        oSheet.getColumns().getByName('U').Columns.Width = 1600
        oSheet.getColumns().getByName('V').Columns.Width = 1600
        oSheet.getColumns().getByName('W').Columns.Width = 100
        oSheet.getColumns().getByName('X').Columns.Width = 1600
        oSheet.getColumns().getByName('Y').Columns.Width = 1600
        oSheet.getColumns().getByName('Z').Columns.Width = 1600
        oSheet.getColumns().getByName('AA').Columns.Width = 1600
        oDoc.CurrentController.freezeAtPosition(0, 3)
    adatta_altezza_riga(oSheet.Name)
########################################################################
#~ class adegua_tmpl_th(threading.Thread):
    #~ def __init__(self):
        #~ threading.Thread.__init__(self)
    #~ def run(self):
        #~ adegua_tmpl_run()
def adegua_tmpl(arg=None):
    #~ adegua_tmpl_th().start()
#~ def debug(arg=None):
    '''
    Mantengo la compatibilità con le vecchie versioni del template:
    - dal 200 parte di autoexec è in python
    - dal 203(LeenO 3.14.0 ha templ 202) introdotta la Super Categoria con nuovi stili di cella;
        sostituita la colonna "Tag A" con "Tag Super Cat"
    - dal 207 introdotta la colonna dei materiali in computo e contabilità
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
#qui le cose da cambiare comunque
    
    oSheet = oDoc.getSheets().getByName('S1')
    flags = VALUE + DATETIME + STRING + ANNOTATION + FORMULA + OBJECTS + EDITATTR # FORMATTED + HARDATTR 
    #cancello da S! le variabili gestite in leeno.conf
    for x in (333, 335):
        oSheet.getCellRangeByPosition(6, x, 30, x).clearContents(flags)

# di seguito ci metto tutte le variabili aggiunte dopo la prima introduzione di /.config/leeno/leeno.conf
    try:
        conf.read(path_conf, 'Contabilità', 'idxSAL')
    except:
        conf.write(path_conf, 'Contabilità', 'idxSAL', '30') #numero massimo di SAL possibili
    try:
        conf.read(path_conf, 'Generale', 'pesca_auto')
    except:
        conf.write(path_conf, 'Generale', 'pesca_auto', '1') #abilita il pesca dopo inserimento nuova voce
    try:
        conf.read(path_conf, 'Generale', 'movedirection')
    except:
        conf.write(path_conf, 'Generale', 'movedirection', '0') #muove il cursore in basso

    # cambiare stile http://bit.ly/2cDcCJI
    
    ver_tmpl = oDoc.getDocumentProperties().getUserDefinedProperties().Versione
    if ver_tmpl > 200:
        basic_LeenO('_variabili.autoexec') #rinvia a autoexec in basic
    if ver_tmpl < 207:
        if DlgSiNo('''Vuoi procedere con l'adeguamento di questo file
alla versione corrente di LeenO?

In caso affermativo dovrai attendere il completamento
dell'operazione che terminerà con un messaggio di avviso.
''', "Richiesta") !=2:
            MsgBox('''Non avendo effettuato l'adeguamento del lavoro alla versione corrente di LeenO, potresti avere dei malfunzionamenti!''', 'Avviso!')
            return
        oDialogo_attesa = dlg_attesa("Adeguamento alla versione corrente di LeenO...")
        oDoc.CurrentController.ZoomValue = 400
        attesa().start() #mostra il dialogo

#~ adeguo gli stili secondo il template corrente
        sUrl = LeenO_path()+'/template/leeno/Computo_LeenO.ots'
        styles = oDoc.getStyleFamilies()
        styles.loadStylesFromURL(sUrl, list())
        
        oSheet.getCellByPosition(7, 290).Value = oDoc.getDocumentProperties().getUserDefinedProperties().Versione = 207
        for el in oDoc.Sheets.ElementNames:
            oDoc.getSheets().getByName(el).IsVisible = True
            oDoc.CurrentController.setActiveSheet(oDoc.getSheets().getByName(el))
            adatta_altezza_riga(el)
            oDoc.getSheets().getByName(el).IsVisible = False
        _gotoSheet('S5')
        oSheet = oDoc.getSheets().getByName('S5')
        oSheet.getCellByPosition(28,11).Formula = '=S12-AE12'
        oSheet.getCellByPosition(28,11).CellStyle = 'Comp-sotto euri'
        oSheet.getCellByPosition(28,26).Formula = '=P27-AE27'
        oSheet.getCellByPosition(28,26).CellStyle = 'Comp-sotto euri'
        for el in('CONTABILITA', 'VARIANTE', 'COMPUTO'):
            if oDoc.getSheets().hasByName(el) == True:
                _gotoSheet(el)
                oSheet = oDoc.getSheets().getByName(el)
                if oSheet.Name != 'CONTABILITA': Rinumera_TUTTI_Capitoli2()
                oSheet.getCellByPosition(31,2).String = 'Super Cat'
                oSheet.getCellByPosition(32,2).String = 'Cat'
                oSheet.getCellByPosition(33,2).String = 'Sub Cat'
                oSheet.getCellByPosition(28,2).String = 'Materiali\ne Noli €'
                n = ultima_voce(oSheet)
                oSheet.getCellByPosition(28,n+1).Formula = '=SUBTOTAL(9;AC3:AC'+ str(n+2)
                lrow = 0
                while lrow < n:
                    try:
                        sStRange = Circoscrive_Voce_Computo_Att(lrow)
                        sotto = sStRange.RangeAddress.EndRow
                        if oSheet.Name == 'CONTABILITA':
                            oSheet.getCellByPosition(28,sotto).Formula = '=P' + str(sotto+1) + '-AE' + str(sotto+1)
                        else:
                            oSheet.getCellByPosition(28,sotto).Formula = '=S' + str(sotto+1) + '-AE' + str(sotto+1)
                        oSheet.getCellByPosition(28,sotto).CellStyle = 'Comp-sotto euri'
                        lrow =next_voice(lrow,1)
                    except:
                        lrow += 1
        oDoc.getSheets().getByName('S1').IsVisible = False
        oDialogo_attesa.endExecute() #chiude il dialogo
        
        oDoc.CurrentController.ZoomValue = 80
        MsgBox("Adeguamento del file completato con successo.", "Avviso")
#~ ########################################################################
def r_version_code(arg=None):
    if os.altsep:
        code_file = uno.fileUrlToSystemPath(LeenO_path() + os.altsep + 'leeno_version_code')
    else:
        code_file = uno.fileUrlToSystemPath(LeenO_path() + os.sep + 'leeno_version_code')
    f = open(code_file, 'r')
    return f.readline().split('-')[-1]
########################################################################
def XPWE_export_run(arg=None ):
    '''
    Viasualizza il menù export/import XPWE
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    psm = uno.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDlgXLO = dp.createDialog("vnd.sun.star.script:UltimusFree2.Dialog_XLO?language=Basic&location=application")
    oDialog1Model = oDlgXLO.Model
    oDlgXLO.Title = 'Menù export XPWE'
    if oDlgXLO.execute() ==1:
        if oDlgXLO.getControl("CME_XLO").State == True:
            XPWE_out('COMPUTO')
        elif  oDlgXLO.getControl("VAR_XLO").State == True:
            XPWE_out('VARIANTE')
        elif  oDlgXLO.getControl("CON_XLO").State == True:
            XPWE_out('CONTABILITA')
########################################################################
def XPWE_import_run(arg=None ):
    '''
    Viasualizza il menù export/import XPWE
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    psm = uno.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDlgXLO = dp.createDialog("vnd.sun.star.script:UltimusFree2.Dialog_XLO?language=Basic&location=application")
    oDialog1Model = oDlgXLO.Model
    oDlgXLO.Title = 'Menù import XPWE'
    if oDlgXLO.execute() ==1:
        if oDlgXLO.getControl("CME_XLO").State == True:
            XPWE_in('COMPUTO')
        elif  oDlgXLO.getControl("VAR_XLO").State == True:
            XPWE_in('VARIANTE')
        elif  oDlgXLO.getControl("CON_XLO").State == True:
            XPWE_in('CONTABILITA')

########################################################################
def DlgMain(arg=None):
    '''
    Visualizza il menù principale
    '''
    bak_timestamp() # fa il backup del file
    oDoc = XSCRIPTCONTEXT.getDocument()
    psm = uno.getComponentContext().ServiceManager
    oSheet = oDoc.CurrentController.ActiveSheet
    if oDoc.getSheets().hasByName('S2') == False:
        for bar in GetmyToolBarNames:
            toolbar_on(bar, 0)
        if len(oDoc.getURL())==0 and \
        getLastUsedCell(oSheet).EndColumn ==0 and \
        getLastUsedCell(oSheet).EndRow ==0:
            oDoc.close(True)
        New_file.computo()
    toolbar_vedi()
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    global oDlgMain
    oDlgMain = dp.createDialog("vnd.sun.star.script:UltimusFree2.DlgMain?language=Basic&location=application")
    oDialog1Model = oDlgMain.Model
    oDlgMain.Title = 'Menù Principale(Ctrl+0)'
    
    sUrl = LeenO_path()+'/icons/Immagine.png'
    oDlgMain.getModel().ImageControl1.ImageURL=sUrl

    if os.altsep:
        code_file = uno.fileUrlToSystemPath(LeenO_path() + os.altsep + 'leeno_version_code')
    else:
        code_file = uno.fileUrlToSystemPath(LeenO_path() + os.sep + 'leeno_version_code')
    f = open(code_file, 'r')
    
    sString = oDlgMain.getControl("Label12")
    sString.Text = f.readline()
    
    sString = oDlgMain.getControl("Label_DDC")
    sString.Text = sUltimus #oSheet.getCellByPosition(2,27).String

    sString = oDlgMain.getControl("Label1")
    sString.Text = str(Lmajor) +'.'+ str(Lminor) +'.'+ Lsubv

    sString = oDlgMain.getControl("Label2")
    try:
        oSheet = oDoc.Sheets.getByName('S1')
    except:
        return
    sString.Text = oDoc.getDocumentProperties().getUserDefinedProperties().Versione #oSheet.getCellByPosition(7, 290).String
    try:
        oSheet = oDoc.Sheets.getByName('COMPUTO')
        sString = oDlgMain.getControl("Label8")
        sString.Text = "€ {:,.2f}".format(oSheet.getCellByPosition(18, 1).Value)
    except:
        pass
    try:
        oSheet = oDoc.Sheets.getByName('VARIANTE')
        sString = oDlgMain.getControl("Label5")
        sString.Text = "€ {:,.2f}".format(oSheet.getCellByPosition(18, 1).Value)
    except:
        pass
    try:
        oSheet = oDoc.Sheets.getByName('CONTABILITA')
        sString = oDlgMain.getControl("Label9")
        sString.Text = "€ {:,.2f}".format(oSheet.getCellByPosition(15, 1).Value)
    except:
        pass
    sString = oDlgMain.getControl("ComboBox1")
    
    sString.Text = conf.read(path_conf, 'Generale', 'visualizza')
    
    oDlgMain.execute()
    sString = oDlgMain.getControl("ComboBox1")
    conf.write(path_conf, 'Generale', 'visualizza', sString.getText())
    return
########################################################################
def InputBox(sCella='', t=''):
    '''
    sCella  { string } : stringa di default nella casella di testo
    t       { string } : titolo del dialogo
    Viasualizza un dialogo di richiesta testo
    '''

    psm = uno.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDialog1 = dp.createDialog("vnd.sun.star.script:UltimusFree2.DlgTesto?language=Basic&location=application")
    oDialog1Model = oDialog1.Model

    oDialog1Model.Title = t

    sString = oDialog1.getControl("TextField1")
    sString.Text = sCella

    if oDialog1.execute()==0:
        return
    else:
        return sString.Text

import zipfile
########################################################################
def hide_error(lErrori, icol):
    '''
    lErrori  { tuple } : nome dell'errore es.: '#DIV/0!'
    icol { integer } : indice di colonna della riga da nascondere
    Viasualizza o nascondi una toolbar
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oDoc.CurrentController.ZoomValue = 400
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    n = 3
    test = ultima_voce(oSheet)+1
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    for i in range(n, test):
        for el in lErrori:
            if oSheet.getCellByPosition(icol, i).String == el:
                oCellRangeAddr.StartRow = i
                oCellRangeAddr.EndRow = i
                oSheet.group(oCellRangeAddr,1)
                oSheet.getCellByPosition(0, i).Rows.IsVisible = False
    oDoc.CurrentController.ZoomValue = 80
########################################################################
def bak_timestamp(arg=None):
    '''
    fa il backup del file di lavoro, partendo dall'ultimo salvataggio certo,
    in una directory con nome "/percorso_file/leeno-bk/"
    viene avviato con DlgMain()
    '''
    tempo = ''.join(''.join(''.join(str(datetime.now()).split('.')[0].split(' ')).split('-')).split(':'))
    oDoc = XSCRIPTCONTEXT.getDocument()

    orig = oDoc.getURL()
    dest = '.'.join(os.path.basename(orig).split('.')[0:-1])+ '-' + tempo + '.ods'
    dir_bak = os.path.dirname(oDoc.getURL()) + '/leeno-bk/'
    if len(orig) ==0:
        return
    orig = uno.fileUrlToSystemPath(orig)
    dir_bak = uno.fileUrlToSystemPath(dir_bak)
    dest = uno.fileUrlToSystemPath(dest)
    if not os.path.exists(dir_bak):
        os.makedirs(dir_bak)
    shutil.copyfile(orig, dir_bak + dest)
    return
########################################################################
# Scrive un file.
def w_version_code(arg=None):
    '''
    scrive versione e timestamp nel file leeno_version_code
    '''
    tempo = ''.join(''.join(''.join(str(datetime.now()).split('.')[0].split(' ')).split('-')).split(':'))

    if os.altsep:
        out_file = uno.fileUrlToSystemPath(LeenO_path() + os.altsep + 'leeno_version_code')
    else:
        out_file = uno.fileUrlToSystemPath(LeenO_path() + os.sep + 'leeno_version_code')
        
    of = open(out_file,'w')
    of.write(str(Lmajor) +'.'+ str(Lminor) +'.'+ Lsubv +'-'+ tempo[:-2])
    of.close()
    return str(Lmajor) +'.'+ str(Lminor) +'.'+ Lsubv +'-'+ tempo[:-2]
########################################################################
def toolbar_vedi(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    try:
        oLayout = oDoc.CurrentController.getFrame().LayoutManager

        if oDoc.getSheets().getByName('S1').getCellByPosition(7,316).Value == 0:
            for bar in GetmyToolBarNames: #toolbar sempre visibili
                toolbar_on(bar)
        else:
            for bar in GetmyToolBarNames: #toolbar contestualizzate
                toolbar_on(bar, 0)
        #~ oLayout.hideElement("private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV")
        toolbar_ordina()
        oLayout.showElement("private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar")
        nSheet = oDoc.CurrentController.ActiveSheet.Name

        if nSheet == 'Elenco Prezzi':
            toolbar_on('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ELENCO')
        elif nSheet == 'Analisi di Prezzo':
            toolbar_on('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ANALISI')
        elif nSheet == 'CONTABILITA':
            toolbar_on('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CONTABILITA')
        elif nSheet in('COMPUTO','VARIANTE'):
            toolbar_on('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_COMPUTO')
    except:
        pass
def toolbar_switch(arg=1):
#~ def debug(arg=None):
    '''Nasconde o mostra le toolbar di Libreoffice.'''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oLayout = oDoc.CurrentController.getFrame().LayoutManager
    for el in oLayout.Elements:
        if el.ResourceURL not in GetmyToolBarNames +('private:resource/menubar/menubar', 'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV', 'private:resource/toolbar/findbar','private:resource/statusbar/statusbar',):
            #~ if oLayout.isElementVisible(el.ResourceURL):
            if arg == 0:
                oLayout.hideElement(el.ResourceURL)
            else:
                oLayout.showElement(el.ResourceURL)
    return
    #~ private:resource/toolbar/standardbar
def toolbar_on(toolbarURL, flag=1):
    '''
    toolbarURL  { string } : indirizzo toolbar
    flag { integer } : 1 = acceso; 0 = spento
    Viasualizza o nascondi una toolbar
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oLayout = oDoc.CurrentController.getFrame().LayoutManager
    if flag == 0:
        oLayout.hideElement(toolbarURL)
    else:
        oLayout.showElement(toolbarURL)
#######################################################################
from com.sun.star.awt import Point
def toolbar_ordina(arg=None):
    #~ https://www.openoffice.org/api/docs/common/ref/com/sun/star/ui/DockingArea.html
    oDoc = XSCRIPTCONTEXT.getDocument()
    oLayout = oDoc.CurrentController.getFrame().LayoutManager
    i = 0
    for bar in GetmyToolBarNames:
        oLayout.dockWindow(bar, 'DOCKINGAREA_TOP', Point(i, 4))
        i += 1
    oLayout.dockWindow('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV', 'DOCKINGAREA_RIGHT', Point(0, 0))
#######################################################################
def make_pack(arg=None, bar=0):
    '''
    bar { integer } : toolbar 0=spenta 1=accesa
    Pacchettizza l'estensione in duplice copia: LeenO.oxt e LeenO-yyyymmddhhmm.oxt
    in una directory precisa(per ora - da parametrizzare)
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    try:
        if oDoc.getSheets().getByName('S1').getCellByPosition(7,338).String == '':
            src_oxt ='_LeenO'
        else:
            src_oxt = oDoc.getSheets().getByName('S1').getCellByPosition(7,338).String
    except:
        pass
    tempo = w_version_code()
    if bar == 0:
        oDoc = XSCRIPTCONTEXT.getDocument()
        for bar in GetmyToolBarNames: #toolbar sempre visibili
            toolbar_on(bar, 0)
        oLayout = oDoc.CurrentController.getFrame().LayoutManager
        oLayout.hideElement("private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV")
    oxt_path = uno.fileUrlToSystemPath(LeenO_path())
    if sys.platform == 'linux' or sys.platform == 'darwin':
        nomeZip2= '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/OXT/LeenO-' + tempo + '.oxt'
        nomeZip = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/OXT/LeenO.oxt'
        os.system('nemo /media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/OXT')
    elif sys.platform == 'win32':
        if not os.path.exists('w:/_dwg/ULTIMUSFREE/_SRC/OXT/'):
            try:
                os.makedirs(os.getenv("HOMEPATH") +'/'+ src_oxt +'/')
            except FileExistsError:
                pass
            nomeZip2= os.getenv("HOMEPATH") +'/'+ src_oxt +'/OXT/LeenO-' + tempo + '.oxt'
            nomeZip = os.getenv("HOMEPATH") +'/'+ src_oxt +'/OXT/LeenO.oxt'
            os.system('explorer.exe ' + os.getenv("HOMEPATH") +'\\'+ src_oxt +'\\OXT\\')
        else:
            nomeZip2= 'w:/_dwg/ULTIMUSFREE/_SRC/OXT/LeenO-' + tempo + '.oxt'
            nomeZip = 'w:/_dwg/ULTIMUSFREE/_SRC/OXT/LeenO.oxt'
            os.system('explorer.exe w:\\_dwg\\ULTIMUSFREE\\_SRC\\OXT\\')
    
    shutil.make_archive(nomeZip2, 'zip', oxt_path)
    shutil.move(nomeZip2 + '.zip', nomeZip2)
    shutil.copyfile(nomeZip2, nomeZip)
    #~ chi(os.getenv("HOMEPATH") +'\\'+ src_oxt +'\\OXT\\')
#######################################################################
def dlg_attesa(msg=''):
    '''
    definisce la variabile globale oDialogo_attesa
    che va gestita così negli script:
    
    oDialogo_attesa = dlg_attesa()
    attesa().start() #mostra il dialogo
    ...
    oDialogo_attesa.endExecute() #chiude il dialogo
    '''
    psm = uno.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    global oDialogo_attesa
    oDialogo_attesa = dp.createDialog("vnd.sun.star.script:UltimusFree2.DlgAttesa?language=Basic&location=application")

    oDialog1Model = oDialogo_attesa.Model # oDialogo_attesa è una variabile generale
    
    sString = oDialogo_attesa.getControl("Label2")
    sString.Text = msg #'ATTENDI...'
    oDialogo_attesa.Title = 'Operazione in corso...'
    sUrl = LeenO_path()+'/icons/attendi.png'
    oDialogo_attesa.getModel().ImageControl1.ImageURL=sUrl
    return oDialogo_attesa
#~ #

class attesa(threading.Thread):
    #~ http://bit.ly/2fzfsT7
    '''avvia il dialogo di attesa'''
    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
        oDialogo_attesa.execute()
        return
########################################################################
class firme_in_calce_th(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        firme_in_calce_run()
def firme_in_calce(arg=None):
    firme_in_calce_th().start()
########################################################################
class XPWE_import_th(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        XPWE_import_run()
def XPWE_import(arg=None):
    XPWE_import_th().start()
########################################################################
class XPWE_export_th(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        XPWE_export_run()
def XPWE_export(arg=None):
    XPWE_export_th().start()
########################################################################
class inserisci_nuova_riga_con_descrizione_th(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        oDialogo_attesa = dlg_attesa()
        oDoc = XSCRIPTCONTEXT.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        if oSheet.Name not in('COMPUTO', 'VARIANTE'):
            return
        descrizione = InputBox(t='inserisci una descrizione per la nuova riga')
        attesa().start() #mostra il dialogo
        
        oDoc.CurrentController.ZoomValue = 400
        i =0
        while(i < getLastUsedCell(oSheet).EndRow):

            if oSheet.getCellByPosition(2, i ).CellStyle == 'comp 1-a':
                sStRange = Circoscrive_Voce_Computo_Att(i)
                qui = sStRange.RangeAddress.StartRow+1

                i = sotto = sStRange.RangeAddress.EndRow+3
                oDoc.CurrentController.select(oSheet.getCellByPosition(2, qui ))
                Copia_riga_Ent()
                oSheet.getCellByPosition(2, qui+1 ).String = descrizione
                next_voice(sotto)

                oDoc.CurrentController.select(oSheet.getCellByPosition(2, i ))
            i += 1
        oDialogo_attesa.endExecute() #chiude il dialogo
        oDoc.CurrentController.ZoomValue = 100
def inserisci_nuova_riga_con_descrizione(arg=None):
    '''
    inserisce, all'inizio di ogni voce di computo o variante,
    una nuova riga con una descrizione a scelta
    '''
    inserisci_nuova_riga_con_descrizione_th().start()
########################################################################
def ctrl_d(arg=None):
    '''
    Copia il valore della prima cella superiore utile.
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oCell= oDoc.CurrentSelection
    oSheet = oDoc.CurrentController.ActiveSheet
    x = Range2Cell()[0]
    lrow = Range2Cell()[1]
    y = lrow-1
    try:
        while oSheet.getCellByPosition(x, y).Type.value == 'EMPTY':
            y -= 1
    except:
        return
    oDoc.CurrentController.select(oSheet.getCellByPosition(x, y))
    copy_clip()
    oDoc.CurrentController.select(oCell)
    paste_clip()
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
########################################################################
def taglia_x(arg=None):
    '''
    taglia il contenuto della selezione
    senza cancellare la formattazione delle celle
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()

    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    dispatchHelper.executeDispatch(oFrame, ".uno:Copy", "", 0, list())

    try:
        sRow = oDoc.getCurrentSelection().getRangeAddresses()[0].StartRow
        sCol = oDoc.getCurrentSelection().getRangeAddresses()[0].StartColumn
        eRow = oDoc.getCurrentSelection().getRangeAddresses()[0].EndRow
        eCol = oDoc.getCurrentSelection().getRangeAddresses()[0].EndColumn
    except AttributeError:
        sRow = oDoc.getCurrentSelection().getRangeAddress().StartRow
        sCol = oDoc.getCurrentSelection().getRangeAddress().StartColumn
        eRow = oDoc.getCurrentSelection().getRangeAddress().EndRow
        eCol = oDoc.getCurrentSelection().getRangeAddress().EndColumn
    oRange = oSheet.getCellRangeByPosition(sCol, sRow, eCol, eRow)
    flags = VALUE + DATETIME + STRING + ANNOTATION + FORMULA + OBJECTS + EDITATTR # FORMATTED + HARDATTR 
    oSheet.getCellRangeByPosition(sCol, sRow, eCol, eRow).clearContents(flags)
########################################################################
def debug_mt(arg=None): #COMUNE DI MATERA
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    #~ chi(oSheet.getCellRangeByName('B278').Type.value)
    #~ return
    #~ chi(oSheet.getCellRangeByName('a6').CellBackColor)# 
    #~ return
 

    #~ oSheet.getCellRangeByName('Y254')
    
    #~ return
    #~ col1 = 16777072 #16771481
    #~ col2 = 16777120 #16771501
    #~ for y in reversed(range(3, getLastUsedCell(oSheet).EndRow)):
        #~ if oSheet.getCellByPosition(0, y).CellBackColor == 16771481:
            #~ oSheet.getCellByPosition(0, y).CellBackColor = 16777072
            #~ oSheet.getCellRangeByPosition(1, y, 26, y).clearContents(HARDATTR)
        #~ if oSheet.getCellByPosition(0, y).CellBackColor == 16771501:
            #~ oSheet.getCellByPosition(0, y).CellBackColor = 16777120
            #~ oSheet.getCellRangeByPosition(1, y, 26, y).clearContents(HARDATTR)
    #~ return
    #~ for y in range(3, getLastUsedCell(oSheet).EndRow):
        #~ for x in (29, 30):
            #~ oSheet.getCellByPosition(x, y).String= oSheet.getCellByPosition(x, y).String.replace(' ','\n')
    #~ chi(len(oSheet.getCellRangeByName('A6').String.split('.')))
    #~ return
# SALTA SULLE CELLE 
    #~ for y in range(Range2Cell()[1]+1, getLastUsedCell(oSheet).EndRow):
    for y in reversed(range(0, getLastUsedCell(oSheet).EndRow+1)):
        if  oSheet.getCellByPosition(0, y).CellStyle in ('Livello-0-scritta', 'Livello-1-scritta'):
            oSheet.getRows().removeByIndex(y, 1)

            #~ testo = oSheet.getCellByPosition(2, y).String.split('- art. ')[1]

            #~ oSheet.getCellByPosition(2, y).String = '- vedi voce art. ' + testo
            #~ oSheet.getCellByPosition(5, y).Value = oSheet.getCellByPosition(5, y).Value
            #~ _gotoCella(2, y)
            #~ chi (len(oSheet.getCellByPosition(0, y).String.split('.')))
            #~ _gotoCella(4, y)
            
            #~ return
            #~ oSheet.getCellByPosition(6, y).Value = oSheet.getCellByPosition(5, y).Value / 100
    
            
             
        #~ if len (oSheet.getCellByPosition(2, y).String) > 5:
            #~ oSheet.getCellByPosition(4, y).String = ''
        #~ for x in range(3, 3):
        #~ if oSheet.getCellByPosition(2, y).Type.value == 'TEXT' and oSheet.getCellByPosition(3, y).Type.value == 'TEXT':
            #~ oSheet.getCellByPosition(1, y).String = oSheet.getCellByPosition(1, y).String +' '+oSheet.getCellByPosition(2, y).String
            #~ oSheet.getCellByPosition(2, y).String = oSheet.getCellByPosition(3, y).String
            #~ oSheet.getCellByPosition(3, y).Value = oSheet.getCellByPosition(4, y).Value
            #~ oSheet.getCellByPosition(4, y).Value = oSheet.getCellByPosition(5, y).Value
            #~ oSheet.getCellByPosition(5, y).String = ''
            
            #~ if oSheet.getCellByPosition(x, y).getIsMerged() == True:
            
            #~ return
    chi("fine")
    return

# SPALMA I VALORI
    #~ for y in range(0, getLastUsedCell(oSheet).EndRow):
        #~ if oSheet.getCellByPosition(0, y).Type.value =='VALUE':
            #~ valore = oSheet.getCellByPosition(0, y).Value
        #~ else:
            #~ oSheet.getCellByPosition(0, y).Value = valore

# COLORA VALORI DIFFERENTI
    #~ for y in range(3, getLastUsedCell(oSheet).EndRow):
        #~ if oSheet.getCellByPosition(10, y).String != oSheet.getCellByPosition(11, y).String:
            #~ if oSheet.getCellByPosition(11, y).String != '':
                #~ oSheet.getCellByPosition(10, y).CellBackColor = 16777113
                #~ oSheet.getCellByPosition(11, y).CellBackColor = 16777113

#~ # SOSTITUZIONI
    test = getLastUsedCell(oSheet).EndRow+1
    #~ for y in range(Range2Cell()[1]+1, test):
    for y in range(3, test):
        if '-' in oSheet.getCellByPosition(13, y).String:
            oSheet.getCellByPosition(13, y).String = oSheet.getCellByPosition(13, y).String.replace('-','/')
    return
# inserisce numero tabella
    #~ for y in range(0, getLastUsedCell(oSheet).EndRow):
        #~ if oSheet.getCellByPosition(0, y).CellBackColor == 16777113:
            #~ oSheet.getCellByPosition(13, y).Formula = '=VLOOKUP(A'+ str(y+1) + ';strade;2;0)'
            #~ oSheet.getCellByPosition(13, y).String = oSheet.getCellByPosition(13, y).String
            #~ oSheet.getCellRangeByPosition(0, y, 11, y).merge(True)

    #~ return  
#~ RECUPERA VIE
    #~ vie = list()
    #~ n = 0
    #~ for y in range(0, getLastUsedCell(oSheet).EndRow):
        #~ if oSheet.getCellByPosition(0, y).CellBackColor == 16777113:
            #~ oSheet.getCellByPosition(12, y).CellBackColor = 16777113
            #~ n += 1
            #~ testo = oSheet.getCellByPosition(0, y).String
            #~ num = oSheet.getCellByPosition(13, y).Value
            #~ el =(n, testo, num)
            #~ vie.append(el)
    
    #~ oSheet = oDoc.getSheets().getByName('VIE')
    #~ oRange = oSheet.getCellRangeByPosition(0, 1, len(vie[0])-1, len(vie))
    #~ lista_come_array = tuple(vie)
    #~ oRange.setDataArray(lista_come_array)
#~ crea via e numero
    #~ for y in range(0, getLastUsedCell(oSheet).EndRow+1):
        #~ if oSheet.getCellByPosition(0, y).CellBackColor == 16777113:
            #~ testo = oSheet.getCellByPosition(0, y).String
        #~ else:
            #~ try:
                #~ if oSheet.getCellByPosition(2, y).String != '':
                    #~ oSheet.getCellByPosition(12, y).String = testo + ', ' + oSheet.getCellByPosition(2, y).String.upper()
                #~ else:
                    #~ oSheet.getCellByPosition(12, y).String = ''
            #~ except:
                #~ pass
#~ elimina '/' finale
    #~ for y in range(0, getLastUsedCell(oSheet).EndRow):
        #~ try:
            #~ if oSheet.getCellByPosition(9, y).String[-1] == '/':
                #~ oSheet.getCellByPosition(9, y).String = oSheet.getCellByPosition(9, y).String[:-1]
        #~ except:
            #~ pass
            
    #~ return
 #~ INSERISCI PARTICELLE
    #~ for y in reversed(range(3, getLastUsedCell(oSheet).EndRow+1)):
        #~ if oSheet.getCellByPosition(10, y).String != '':
            #~ if oSheet.getCellByPosition(11, y).String != '':
                #~ oSheet.getCellByPosition(0, y).String = oSheet.getCellByPosition(10, y).String + '/' + oSheet.getCellByPosition(11, y).String
            #~ else:
                #~ oSheet.getCellByPosition(0, y).String = oSheet.getCellByPosition(10, y).String
    
#~ ricerca graffate
    #~ for y in reversed(range(3, getLastUsedCell(oSheet).EndRow)):
        #~ if '\n' in oSheet.getCellByPosition(6, y).String:
            #~ particelle = oSheet.getCellByPosition(6, y).String.split('\n')
            #~ sub = oSheet.getCellByPosition(7, y).String.split('\n')

            #~ while len(sub) < len(particelle):
                #~ sub.append('')
            #~ oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, y, 5, y))
            #~ copy_clip()
            #~ oSheet.getRows().insertByIndex(y+1, len(particelle)-1)
            #~ oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, y+1, 0, y+len(particelle)-1))
            #~ paste_clip()
            #~ for n in range(0, len(particelle)):
                #~ oSheet.getCellByPosition(6, y+n).String = particelle[n]
                #~ oSheet.getCellByPosition(7, y+n).String = sub[n]
                #~ oSheet.getCellByPosition(15, y+n).String = particelle[0]+ '/' + sub[0]
    #~ chi(sub)
#RAGGRUPPA LE RIGHE SECONDO IL COLoRE
    #~ oDoc = XSCRIPTCONTEXT.getDocument()
    #~ oSheet = oDoc.CurrentController.ActiveSheet
    #~ iSheet = oSheet.RangeAddress.Sheet
    #~ oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    #~ oCellRangeAddr.Sheet = iSheet
    #~ lista = list()
    #~ test = getLastUsedCell(oSheet).EndRow-1
    #~ for n in range(0, test):
        #~ if oSheet.getCellByPosition(0, n).CellBackColor == 16777113:
            #~ sopra = n+1
            
            #~ for n in range(sopra+1, test):
                #~ if oSheet.getCellByPosition(0, n).CellBackColor == 16777113:
                    #~ sotto = n-1
                    #~ lista.append((sopra, sotto))

                    #~ break
    #~ for el in lista:
        #~ oCellRangeAddr.StartRow = el[0]
        #~ oCellRangeAddr.EndRow = el[1]
        #~ oSheet.group(oCellRangeAddr,1)
        #~ oSheet.getCellRangeByPosition(0, el[0], 0, el[1]).Rows.IsVisible=False
    #~ return
########################################################################
# ELENCO DEGLI SCRIPT VISUALIZZATI NEL SELETTORE DI MACRO              #
g_exportedScripts = attiva_contabilita,
########################################################################
########################################################################
# ... here is the python script code
# this must be added to every script file(the
# name org.openoffice.script.DummyImplementationForPythonScripts should be changed to something
# different(must be unique within an office installation !)
# --- faked component, dummy to allow registration with unopkg, no functionality expected
#~ import unohelper
# questo mi consente di inserire i comandi python in Accelerators.xcu
# vedi pag.264 di "Manuel du programmeur oBasic"
# <<< vedi in description.xml
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(None, "org.giuseppe-vizziello.leeno",("org.giuseppe-vizziello.leeno",),)
########################################################################
