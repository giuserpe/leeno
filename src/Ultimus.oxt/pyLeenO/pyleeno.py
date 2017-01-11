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
import locale
import codecs
#~ import subprocess
import os, sys, uno, unohelper, pyuno, logging, shutil, base64
import time
from multiprocessing import Process, freeze_support
import threading
# cos'e' il namespace:
# http://www.html.it/articoli/il-misterioso-mondo-dei-namespaces-1/
from datetime import datetime, date
from com.sun.star.beans import PropertyValue
from xml.etree.ElementTree import ElementTree, Element, SubElement, Comment, tostring
#~ from com.sun.star.table.CellContentType import TEXT, EMPTY, VALUE, FORMULA
from com.sun.star.sheet.CellFlags import (VALUE, DATETIME, STRING,
    ANNOTATION, FORMULA, HARDATTR, OBJECTS, EDITATTR, FORMATTED)
########################################################################
# https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=27805&p=127383
import random
from com.sun.star.script.provider import XScriptProviderFactory
from com.sun.star.script.provider import XScriptProvider

def Xray(myObject):
    # Taken from http://www.oooforum.org/forum/viewtopic.phtml?t=23577
    xCompCont = XSCRIPTCONTEXT.getComponentContext()
    sm = xCompCont.ServiceManager
    mspf = sm.createInstance("com.sun.star.script.provider.MasterScriptProviderFactory")
    scriptPro = mspf.createScriptProvider("")
    Xscript = scriptPro.getScript("vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application")
    Xscript.invoke((myObject,), None, None)

def Lib_LeenO(funcname,*args):
    xCompCont = XSCRIPTCONTEXT.getComponentContext()
    sm = xCompCont.ServiceManager
    mspf = sm.createInstance("com.sun.star.script.provider.MasterScriptProviderFactory")
    scriptPro = mspf.createScriptProvider("");
    Xscript = scriptPro.getScript("vnd.sun.star.script:UltimusFree2." + funcname + "?language=Basic&location=application")
    Result=Xscript.invoke(args,None,None)
    return Result[0]
########################################################################
def LeenO_path(arg=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    pir = ctx.getValueByName('/singletons/com.sun.star.deployment.PackageInformationProvider')
    expath=pir.getPackageLocation('org.giuseppe-vizziello.leeno')
    return (expath)
########################################################################
#~ class New_File:
    #~ ''' Crea un nuovo computo o un nuovo listino '''
    #~ def __init__(self):
        #~ self.desktop = XSCRIPTCONTEXT.getDesktop()
        #~ self.opz = PropertyValue()
        #~ self.opz.Name = 'AsTemplate'
        #~ self.opz.Value = True
    #~ def loadComponent(self, filename):
        #~ path = os.path.join(LeenO_path(), 'template', 'leeno', filename)
        #~ return self.desktop.loadComponentFromURL(path, '_blank', 0, (self.opz,))
    #~ def computo(self):
        #~ return self.loadComponent('Computo_LeenO.ots')
    #~ def listino(self):
        #~ return self.loadComponent('Listino_LeenO.ots')
########################################################################
class New_file:
    '''Crea un nuovo computo o un nuovo listino.'''
    def __init__(self):#, computo, listino):
        pass
    def computo():
        desktop = XSCRIPTCONTEXT.getDesktop()
        opz = PropertyValue()
        opz.Name = 'AsTemplate'
        opz.Value = True
        document = desktop.loadComponentFromURL(LeenO_path()+'/template/leeno/Computo_LeenO.ots', "_blank", 0, (opz,))
        MsgBox('''Prima di procedere è consigliabile salvare il lavoro.
Provvedi subito a dare un nome al file di computo...''', 'Dai un nome al file...')
        salva_come()
        autoexec()
        return (document)
    def listino():
        desktop = XSCRIPTCONTEXT.getDesktop()
        opz = PropertyValue()
        opz.Name = 'AsTemplate'
        opz.Value = True
        document = desktop.loadComponentFromURL(LeenO_path()+'/template/leeno/Listino_LeenO.ots', "_blank", 0, (opz,))
        autoexec()
        return (document)
    def usobollo():
        desktop = XSCRIPTCONTEXT.getDesktop()
        opz = PropertyValue()
        opz.Name = 'AsTemplate'
        opz.Value = True
        document = desktop.loadComponentFromURL(LeenO_path()+'/template/offmisc/UsoBollo.ott', "_blank", 0, (opz,))
        return (document)
########################################################################
def nuovo_computo (arg=None):
    New_file.computo()
########################################################################
def nuovo_listino (arg=None):
    New_file.listino()
########################################################################
def nuovo_usobollo (arg=None):
    New_file.usobollo()
########################################################################
    #~ oDoc = XSCRIPTCONTEXT.getDocument()
    #~ path = uno.fileUrlToSystemPath(oDoc.getURL())
    #~ url = uno.systemPathToFileUrl(path)
    #~ chi(uno.sys)
    #~ chi(url)
########################################################################
def voce_voce(arg=None):
#~ def debug (arg=None):
    '''
    Invia una voce di prezzario da un elenco prezzi all'Elenco Prezzi del
    Documento di Contabilità Corrente DCC
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    #~ oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]

    oRangeAddress = oSheet.getCellRangeByPosition(0, lrow, getLastUsedCell(oSheet).EndColumn, lrow).getRangeAddress()
    
    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, lrow, getLastUsedCell(oSheet).EndColumn, lrow))
    partenza = uno.fileUrlToSystemPath(oDoc.getURL())
        
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    dispatchHelper.executeDispatch(oFrame, ".uno:Copy", "", 0, list())
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect

    if sUltimus == '':
        MsgBox("E' necessario impostare il Documento di contabilità Corrente.", "Attenzione!")
        return
    _gotoDoc(sUltimus)
    
    ddcDoc = XSCRIPTCONTEXT.getDocument()
    #~ chi(ddcDoc.getURL())
    dccSheet = ddcDoc.getSheets().getByName('Elenco Prezzi')
    dccSheet.IsVisible = True
    ddcDoc.CurrentController.setActiveSheet(dccSheet)
    dccSheet.getRows().insertByIndex(3, 1)
    ddcDoc.CurrentController.select(dccSheet.getCellByPosition(0, 3))
    #~ chi('ppp')
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    dispatchHelper.executeDispatch(oFrame, ".uno:Paste", "", 0, list())
    ddcDoc.CurrentController.select(ddcDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect

    #~ _gotoDoc(partenza)

########################################################################
def _gotoDoc(sUrl):
    '''
    sUrl  { string } : nome del file
    porta il focus su di un determinato documento
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    path = oDoc.getURL()
    desktop = XSCRIPTCONTEXT.getDesktop()
    opz = PropertyValue()
    sUrl = uno.systemPathToFileUrl(sUrl)
    target = desktop.loadComponentFromURL(sUrl, "_default", 0, (opz,))
    oFocus = uno.createUnoStruct('com.sun.star.awt.FocusEvent')
    target.getCurrentController().getFrame().focusGained(oFocus)
    
########################################################################
def oggi():
    '''
    restituisce la data di oggi
    '''
    return ('/'.join(reversed(str(datetime.now()).split(' ')[0].split('-'))))
import distutils.dir_util
########################################################################
def copia_sorgente_per_git(arg=None):
    '''
    fa una copia della directory del codice nel repository locale ed apre una shell per la commit
    '''
    make_pack(bar=1)
    oxt_path = uno.fileUrlToSystemPath(LeenO_path())
    if sys.platform == 'linux' or sys.platform == 'darwin':
        dest = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt'
        
        os.system('nemo /media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt')
        os.system('cd /media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt && gnome-terminal && gitk &')
        
    elif sys.platform == 'win32':
        dest = 'w:/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt'
        
        os.system('explorer.exe w:\\_dwg\\ULTIMUSFREE\\_SRC\\leeno\\src\\Ultimus.oxt\\')
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

def Inser_SottoCapitolo_arg (lrow, sTesto): #
    '''
    lrow    { double } : id della riga di inerimento
    sTesto  { string } : titolo della sottocategoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
        return
    #~ lrow = Range2Cell()[1]
    #~ sTesto = 'prova'
    style = oSheet.getCellByPosition(1, lrow).CellStyle
    #~ if style in ('comp Int_colonna', 'Livello-1-scritta', 'livello2 valuta', 'Comp TOTALI',
                #~ 'Comp-Bianche sopra', 'comp Art-EP', 'comp Art-EP_R','Comp-Bianche in mezzo', 'comp sotto Bianche'):
        #~ if style in ('comp Int_colonna', 'Livello-1-scritta', 'livello2 valuta'):
            #~ lrow += 1
        #~ elif style in ('Comp-Bianche sopra', 'comp Art-EP','comp Art-EP_R', 'Comp-Bianche in mezzo', 'comp sotto Bianche'):
            #~ sStRange = Circoscrive_Voce_Computo_Att (lrow)
            #~ lrow = sStRange.RangeAddress.EndRow+1
    if oDoc.getSheets().getByName('S1').getCellByPosition(7,333).Value == 1: #con riga bianca
        insRows(lrow, 2)
        oSheet.getCellRangeByPosition(0, lrow, 41, lrow).CellStyle = 'livello-1-sopra'
        lrow += 1
        oSheet.getCellByPosition(2, lrow).String = sTesto
    else:
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

    if oDoc.getSheets().getByName('S1').getCellByPosition(7,305).Value == 1:
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
    #~ SubSum_Cap (lrow)

########################################################################
def Ins_Categorie(n):
    '''
    n    { int } : livello della categoria
    0 = SuperCategoria
    1 = Categoria
    2 = SubCategoria
    '''
    sTesto = ''
    if n==0:
        sTesto = 'Inserisci il titolo per la Supercategoria'
    elif n==1:
        sTesto = 'Inserisci il titolo per la Categoria'
    elif n==2:
        sTesto = 'Inserisci il titolo per la Sottocategoria'
    sString = InputBox('', sTesto)
    if sString ==None:
        return

    oDoc = XSCRIPTCONTEXT.getDocument()
    #~ oDoc.CurrentController.ZoomValue = 400
    oSheet = oDoc.CurrentController.ActiveSheet
    row = Range2Cell()[1]
    if oSheet.getCellByPosition(0, row).CellStyle in siVoce:
        lrow = next_voice(row, 1)
    elif oSheet.getCellByPosition(0, row).CellStyle in noVoce:
        lrow = row+1
    else:
        return

    if n==0:
        Inser_SuperCapitolo_arg (lrow, sString)
    elif n==1:
        Inser_Capitolo_arg (lrow, sString)
    elif n==2:
        Inser_SottoCapitolo_arg (lrow, sString)

    if oDoc.getSheets().getByName('S1').getCellByPosition(7,333).Value == 1: #con riga bianca
        _gotoCella(2, lrow+1)
    else:
        _gotoCella(2, lrow)
    Rinumera_TUTTI_Capitoli2
    oDoc.CurrentController.ZoomValue = 100
    oDoc.CurrentController.setFirstVisibleColumn(0)
    oDoc.CurrentController.setFirstVisibleRow(lrow-5)
    
########################################################################
def Inser_SuperCapitolo(arg=None):
    Ins_Categorie(0)

def Inser_SuperCapitolo_arg (lrow, sTesto='Super Categoria'): #
    '''
    lrow    { double } : id della riga di inerimento
    sTesto  { string } : titolo della categoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
        return
    #~ lrow = Range2Cell()[1]
    style = oSheet.getCellByPosition(1, lrow).CellStyle
    if oDoc.getSheets().getByName('S1').getCellByPosition(7,333).Value == 1: #con riga bianca
        insRows(lrow, 2)
        oSheet.getCellRangeByPosition(0, lrow, 41, lrow).CellStyle = 'livello-1-sopra'
        lrow += 1
        oSheet.getCellByPosition(2, lrow).String = sTesto
    else:
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
    if oDoc.getSheets().getByName('S1').getCellByPosition(7,305).Value == 1:
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

def Inser_Capitolo_arg (lrow, sTesto='Categoria'): #
    '''
    lrow    { double } : id della riga di inerimento
    sTesto  { string } : titolo della categoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
        return
    #~ lrow = Range2Cell()[1]
    style = oSheet.getCellByPosition(1, lrow).CellStyle
    if oDoc.getSheets().getByName('S1').getCellByPosition(7,333).Value == 1: #con riga bianca
        insRows(lrow, 2)
        oSheet.getCellRangeByPosition(0, lrow, 41, lrow).CellStyle = 'livello-1-sopra'
        lrow += 1
        oSheet.getCellByPosition(2, lrow).String = sTesto
    else:
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
    if oDoc.getSheets().getByName('S1').getCellByPosition(7,305).Value == 1:
        lrowProvv = lrow-1
        while oSheet.getCellByPosition(31, lrowProvv).CellStyle != 'Livello-1-scritta':
            if lrowProvv > 4:
                lrowProvv -=1
            else:
                break
        oSheet.getCellByPosition(31, lrow).Value = oSheet.getCellByPosition(1 , lrowProvv).Value + 1
########################################################################
def Rinumera_TUTTI_Capitoli2(arg=None):
    Tutti_Subtotali()# ricalcola i totali di categorie e subcategorie
    Sincronizza_SottoCap_Tag_Capitolo_Cor()# sistemo gli idcat voce per voce

def Tutti_Subtotali(arg=None):
    '''ricalcola i subtotali di categorie e subcategorie'''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
        return
    for n in range (0, ultima_voce(oSheet)+1):
        if oSheet.getCellByPosition(0, n).CellStyle == 'Livello-0-scritta':
            SubSum_SuperCap (n)
        if oSheet.getCellByPosition(0, n).CellStyle == 'Livello-1-scritta':
            SubSum_Cap (n)
        if oSheet.getCellByPosition(0, n).CellStyle == 'livello2 valuta':
            SubSum_SottoCap (n)
########################################################################
def SubSum_SuperCap (lrow):
    '''
    lrow    { double } : id della riga di inerimento
    inserisce i dati nella riga di SuperCategoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
        return
    #~ lrow = Range2Cell()[1]
    lrowE = ultima_voce(oSheet)+2
    nextCap = lrowE
    for n in range (lrow+1, lrowE):
        if oSheet.getCellByPosition(18, n).CellStyle in ('Livello-0-scritta mini val', 'Comp TOTALI'):
            #~ MsgBox(oSheet.getCellByPosition(18, n).CellStyle,'')
            nextCap = n + 1
            break
    #~ oDoc.enableAutomaticCalculation(False)
    oSheet.getCellByPosition(18, lrow).Formula = '=SUBTOTAL(9;S' + str(lrow + 1) + ':S' + str(nextCap) + ')'
    oSheet.getCellByPosition(18, lrow).CellStyle = 'Livello-0-scritta mini val'
    oSheet.getCellByPosition(24, lrow).Formula = '=S' + str(lrow + 1) + '/S' + str(lrowE+1)
    oSheet.getCellByPosition(24, lrow).CellStyle = 'Livello-0-scritta mini %'
    oSheet.getCellByPosition(29, lrow).Formula = '=AE' + str(lrow + 1) + '/S' + str(lrowE+1)
    oSheet.getCellByPosition(29, lrow).CellStyle = 'Livello-0-scritta mini %'
    oSheet.getCellByPosition(30, lrow).Formula = '=SUBTOTAL(9;AE' + str(lrow + 1) + ':AE' + str(nextCap) + ')'
    oSheet.getCellByPosition(30, lrow).CellStyle = 'Livello-0-scritta mini val'
    #~ oDoc.enableAutomaticCalculation(True)
########################################################################
def SubSum_SottoCap (lrow):
    '''
    lrow    { double } : id della riga di inerimento
    inserisce i dati nella riga di subcategoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
        return
    #lrow = 0#Range2Cell()[1]
    lrowE = ultima_voce(oSheet)+2
    nextCap = lrowE
    for n in range (lrow+1, lrowE):
        if oSheet.getCellByPosition(18, n).CellStyle in ('livello2 scritta mini', 'Livello-0-scritta mini val', 'Livello-1-scritta mini val', 'Comp TOTALI'):
            nextCap = n + 1
            break
    oSheet.getCellByPosition(18, lrow).Formula = '=SUBTOTAL(9;S' + str(lrow + 1) + ':S' + str(nextCap) + ')'
    oSheet.getCellByPosition(18, lrow).CellStyle = 'livello2 scritta mini'
    oSheet.getCellByPosition(24, lrow).Formula = '=S' + str(lrow + 1) + '/S' + str(lrowE+1)
    oSheet.getCellByPosition(24, lrow).CellStyle = 'livello2 valuta mini %'
    oSheet.getCellByPosition(29, lrow).Formula = '=AE' + str(lrow + 1) + '/S' + str(lrowE+1)
    oSheet.getCellByPosition(29, lrow).CellStyle = 'livello2 valuta mini %'
    oSheet.getCellByPosition(30, lrow).Formula = '=SUBTOTAL(9;AE' + str(lrow + 1) + ':AE' + str(nextCap) + ')'
    oSheet.getCellByPosition(30, lrow).CellStyle = 'livello2 valuta mini'
########################################################################
def SubSum_Cap (lrow):
    '''
    lrow    { double } : id della riga di inerimento
    inserisce i dati nella riga di categoria
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
        return
    #~ lrow = Range2Cell()[1]
    lrowE = ultima_voce(oSheet)+2
    nextCap = lrowE
    for n in range (lrow+1, lrowE):
        if oSheet.getCellByPosition(18, n).CellStyle in ('Livello-1-scritta mini val','Livello-0-scritta mini val',  'Comp TOTALI'):
            #~ MsgBox(oSheet.getCellByPosition(18, n).CellStyle,'')
            nextCap = n + 1
            break
    #~ oDoc.enableAutomaticCalculation(False)
    oSheet.getCellByPosition(18, lrow).Formula = '=SUBTOTAL(9;S' + str(lrow + 1) + ':S' + str(nextCap) + ')'
    oSheet.getCellByPosition(18, lrow).CellStyle = 'Livello-1-scritta mini val'
    oSheet.getCellByPosition(24, lrow).Formula = '=S' + str(lrow + 1) + '/S' + str(lrowE+1)
    oSheet.getCellByPosition(24, lrow).CellStyle = 'Livello-1-scritta mini %'
    oSheet.getCellByPosition(29, lrow).Formula = '=AE' + str(lrow + 1) + '/S' + str(lrowE+1)
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
def Sincronizza_SottoCap_Tag_Capitolo_Cor (arg=None):
    '''
    lrow    { double } : id della riga di inerimento
    sincronizza il categoria e sottocategorie
    '''
    datarif = datetime.now()
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oDoc.getSheets().getByName('S1').getCellByPosition(7,304).Value == 0: #se 1 aggiorna gli indici
        return
    if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
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
def ultima_voce (oSheet):
    #~ oDoc = XSCRIPTCONTEXT.getDocument()
    #~ oSheet = oDoc.CurrentController.ActiveSheet
    nRow = getLastUsedCell(oSheet).EndRow
    for n in reversed(range(0, nRow)):
        if oSheet.getCellByPosition(0, n).CellStyle in ('EP-aS', 'EP-Cs', 'An-sfondo-basso Att End', 'Comp End Attributo',
                                                        'Comp End Attributo_R', 'comp Int_colonna', 'comp Int_colonna_R_prima',
                                                        'Livello-0-scritta', 'Livello-1-scritta', 'livello2 valuta'):
            break
    return n
########################################################################
def uFindString (sString, oSheet):
    '''
    sString { string }  : stringa da cercare
    oSheet  { object }  :

    Trova la prima ricorrenza di una stringa (sString) riga
    per riga in un foglio di calcolo (oSheet) e restituisce
    una tupla (IDcolonna, IDriga)
    '''
    oCell = oSheet.getCellByPosition(0,0)
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.gotoEndOfUsedArea(True)
    aAddress = oCursor.RangeAddress
    for nRow in range(0, aAddress.EndRow+1):
        for nCol in range(0, aAddress.EndColumn+1):
    # ritocco di +Daniele Zambelli:
            if sString in oSheet.getCellByPosition(nCol,nRow).String:
                return (nCol,nRow)
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
def copia_sheet (nSheet, tag):
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
    for lrow in reversed(range(0, ultima_voce (oSheet))):
        if oSheet.getCellByPosition(31,lrow).CellStyle == 'compTagG' :
            oSheet.getCellByPosition(31,lrow).String = ''
            oSheet.getCellByPosition(32,lrow).String = ''
            oSheet.getCellByPosition(33,lrow).String = ''
            oSheet.getCellByPosition(34,lrow).String = ''
            oSheet.getCellByPosition(35,lrow).String = ''
    _gotoSheet('S5')
    oSheet = oDoc.CurrentController.ActiveSheet
    for lrow in reversed(range(0, ultima_voce (oSheet))):
        if oSheet.getCellByPosition(31,lrow).CellStyle == 'compTagG' :
            oSheet.getCellByPosition(31,lrow).String = ''
            oSheet.getCellByPosition(32,lrow).String = ''
            oSheet.getCellByPosition(33,lrow).String = ''
            oSheet.getCellByPosition(34,lrow).String = ''
            oSheet.getCellByPosition(35,lrow).String = ''
    _gotoSheet('VARIANTE')
    oSheet = oDoc.CurrentController.ActiveSheet
    for lrow in reversed(range(0, ultima_voce (oSheet))):
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
    copia_sheet (nSheet, sString)
    oSheet = oDoc.CurrentController.ActiveSheet
    for lrow in reversed(range(0, ultima_voce (oSheet))):
        try:
            sStRange = Circoscrive_Voce_Computo_Att (lrow)
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

    #~ iCellAttr = (oDoc.createInstance("com.sun.star.sheet.CellFlags.OBJECTS"))
    flags = OBJECTS
    oSheet.getCellRangeByPosition (0,0,42,0).clearContents(flags) #cancello gli oggetti
    oDoc.CurrentController.select(oSheet.getCellByPosition(0,3))
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
########################################################################
def Vai_a_Filtro (arg=None):
    _gotoSheet('S3')
    _primaCella(0,1)
########################################################################
def Filtra_Computo_Cap (arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7,8).String
    sString = oSheet.getCellByPosition(7,10).String
    Filtra_computo(nSheet, 31, sString)
########################################################################
def Filtra_Computo_SottCap (arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7,8).String
    sString = oSheet.getCellByPosition(7,12).String
    Filtra_computo(nSheet, 32, sString)
########################################################################
def Filtra_Computo_A (arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7,8).String
    sString = oSheet.getCellByPosition(7,14).String
    Filtra_computo(nSheet, 33, sString)
########################################################################
def Filtra_Computo_B (arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7,8).String
    sString = oSheet.getCellByPosition(7,16).String
    Filtra_computo(nSheet, 34, sString)
########################################################################
def Filtra_Computo_C (arg=None): #filtra in base al codice di prezzo
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7,8).String
    sString = oSheet.getCellByPosition(7,20).String
    Filtra_computo(nSheet, 1, sString)
########################################################################
def Vai_a_M1 (arg=None):
    _gotoSheet ('M1', 85)
    _primaCella(0,0)
########################################################################
def Vai_a_S2 (arg=None):
    _gotoSheet ('S2')
########################################################################
def Vai_a_S1 (arg=None):
    _gotoSheet ('S1')
    _primaCella(0,190)
########################################################################
def Vai_a_ElencoPrezzi (arg=None):
    _gotoSheet ('Elenco Prezzi')
########################################################################
def Vai_a_Computo (arg=None):
    _gotoSheet ('COMPUTO')
########################################################################
def Vai_a_Variabili (arg=None):
    _gotoSheet ('S1', 85)
    _primaCella(6,289)
########################################################################
def Vai_a_Scorciatoie (arg=None):
    _gotoSheet ('Scorciatoie')
    _primaCella(0,0)
########################################################################
def Vai_a_SegnaVoci (arg=None):
    _gotoSheet ('S3',100)
    _primaCella(37,4)
########################################################################
def _gotoSheet (nSheet, fattore=100):
    '''
    nSheet   { string } : nome Sheet
    attiva e seleziona una sheet
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.Sheets.getByName(nSheet)
    oSheet.IsVisible = True
    oDoc.CurrentController.setActiveSheet(oSheet)
    oDoc.CurrentController.ZoomValue = fattore

     #~ oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
########################################################################
def _primaCella (IDcol=0, IDrow=0):
    '''
    IDcol   { integer } : id colonna
    IDrow   { integer } : id riga
    settaggio prima cella visibile (IDcol, IDrow)
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oDoc.CurrentController.setFirstVisibleColumn(IDcol)
    oDoc.CurrentController.setFirstVisibleRow(IDrow)
    return
########################################################################
def setTabColor (colore):
    '''
    colore   { integer } : id colonna
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
    properties = (oProp,)
    dispatchHelper.executeDispatch(oFrame, '.uno:SetTabBgColor', '', 0, properties)
########################################################################
def salva_come (arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    oProp = PropertyValue()
    oProp.Name = "FilterName"
    oProp.Value = "calc8"
    properties = (oProp,)
    dispatchHelper.executeDispatch(oFrame, ".uno:SaveAs", "", 0, properties)
########################################################################
def _gotoCella (IDcol=0, IDrow=0):
    '''
    IDcol   { integer } : id colonna
    IDrow   { integer } : id riga

    muove il cursore nelle cella (IDcol, IDrow)
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    
    oDoc.CurrentController.select(oSheet.getCellByPosition(IDcol, IDrow))
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
    return
########################################################################
def adatta_altezza_riga (nome=None):
    '''
    nome   { string } : nome della sheet
    imposta l'altezza ottimale delle celle
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if nome == None:
        nome = oSheet.Name
    oDoc.getSheets().hasByName(nome)
    oSheet.getCellRangeByPosition(0, 0, getLastUsedCell(oSheet).EndColumn, getLastUsedCell(oSheet).EndRow).Rows.OptimalHeight = True
    if oSheet.Name in ('Elenco Prezzi', 'VARIANTE', 'COMPUTO', 'CONTABILITA'):
        oSheet.getCellByPosition(0, 2).Rows.Height = 800
########################################################################
# doppioni #############################################################
def doppioni(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    ###
    lista_voci = list()
    voce = list()
    for n in range (3, ultima_voce(oSheet)+1):
        voce = ( oSheet.getCellByPosition(0, n).String,
            oSheet.getCellByPosition(1, n).String,
            oSheet.getCellByPosition(2, n).String,
            oSheet.getCellByPosition(3, n).Value,
            oSheet.getCellByPosition(4, n).Value,
            oSheet.getCellByPosition(5, n).Value,
            oSheet.getCellByPosition(6, n).Value,
            oSheet.getCellByPosition(7, n).Value,
        )
        lista_voci.append(voce)
    oSheet.getRows().removeByIndex(4, ultima_voce(oSheet)-3) # lascio una riga per conservare gli stili
    oSheet.getRows().insertByIndex(4, len(set(lista_voci))-1)
    
    #~ lista_voci = set (lista_voci)
    lista_come_array = tuple (set (lista_voci))

    scarto_colonne = 0 # numero colonne da saltare a partire da sinistra
    scarto_righe = 3 # numero righe da saltare a partire dall'alto
    colonne_lista = len(lista_come_array[1]) # numero di colonne necessarie per ospitare i dati
    righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati
    oRange = oSheet.getCellRangeByPosition( 0,
                                            3,
                                            colonne_lista + 0 - 1, # l'indice parte da 0
                                            righe_lista + 3 - 1)
    oRange.setDataArray(lista_come_array)
# doppioni #############################################################
########################################################################
# Scrive un file.
def XPWE_out(arg=None):
    '''
    esporta il documento in formato XPWE
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oDialogo_attesa = dlg_attesa()
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
    for n in range (0, lastRow):
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
    diz_ep = dict ()
    lista_AP = list ()
    for n in range (0, getLastUsedCell(oSheet).EndRow):
        if oSheet.getCellByPosition(0, n).CellStyle in ('EP-aS', 'EP-Cs') and \
        oSheet.getCellByPosition(8, n).String  != '(AP)':
        #~ if oSheet.getCellByPosition(0, n).CellStyle in ('EP-aS', 'EP-Cs'):
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
        elif oSheet.getCellByPosition(8, n).String  == '(AP)':
            lista_AP.append(oSheet.getCellByPosition(0, n).String)
#Analisi di prezzo
    if len(lista_AP) != 0:
        k = n+1
        oSheet = oDoc.getSheets().getByName('Analisi di Prezzo')
        for el in lista_AP:
            try:
                n = (uFindString(el, oSheet)[-1])
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
                for x in range (n, n+100):
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
    for n in range (0, ultima_voce(oSheet)):
        if oSheet.getCellByPosition(0, n).CellStyle == 'Comp Start Attributo':
            sStRange = Circoscrive_Voce_Computo_Att (n)
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
            IDSpCat.text = str (oSheet.getCellByPosition(31, sotto).String)
            if IDSpCat.text == '':
                IDSpCat.text = '0'
##########################
            IDCat = SubElement(VCItem,'IDCat')
            IDCat.text = str (oSheet.getCellByPosition(32, sotto).String)
            if IDCat.text == '':
                IDCat.text = '0'
##########################
            IDSbCat = SubElement(VCItem,'IDSbCat')
            IDSbCat.text = str (oSheet.getCellByPosition(33, sotto).String)
            if IDSbCat.text == '':
                IDSbCat.text = '0'
##########################
            PweVCMisure = SubElement(VCItem,'PweVCMisure')
            for m in range (sopra+2, sotto):
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
                try:
                    eval(oSheet.getCellByPosition(5, m).Formula.split('=')[-1].replace('^','**').replace(',','.').replace('1/2','1.0/2'))
                    PartiUguali.text = oSheet.getCellByPosition(5, m).Formula.split('=')[-1]
                except:
                    if oSheet.getCellByPosition(5, m).Value !=0:
                        PartiUguali.text = str(oSheet.getCellByPosition(5, m).Value)
                    else:
                        PartiUguali.text = ''
                if PartiUguali.text == ' ':
                    PartiUguali.text = ''
##########################
                Lunghezza = SubElement(RGItem,'Lunghezza')
                if oSheet.getCellByPosition(6, m).Formula.split('=')[-1] == None:
                    Lunghezza.text = oSheet.getCellByPosition(6, m).String
                else:
                    Lunghezza.text = str(oSheet.getCellByPosition(6, m).Formula.split('=')[-1])
                if Lunghezza.text == ' ':
                    Lunghezza.text = ''
##########################
                Larghezza = SubElement(RGItem,'Larghezza')
                if oSheet.getCellByPosition(7, m).Formula.split('=')[-1] == None:
                    Larghezza.text = oSheet.getCellByPosition(7, m).String
                else:
                    Larghezza.text = str(oSheet.getCellByPosition(7, m).Formula.split('=')[-1])
                if Larghezza.text == ' ':
                    Larghezza.text = ''
##########################
                HPeso = SubElement(RGItem,'HPeso')
                if oSheet.getCellByPosition(8, m).Formula.split('=')[-1] == None:
                    HPeso.text = oSheet.getCellByPosition(8, m).Formula
                else:
                    HPeso.text = str(oSheet.getCellByPosition(8, m).Formula.split('=')[-1])
                if HPeso.text == ' ':
                    HPeso.text = ''
##########################
                Quantita = SubElement(RGItem,'Quantita')
                Quantita.text = str(oSheet.getCellByPosition(9, m).Value)
##########################
                Flags = SubElement(RGItem,'Flags')
                if "Parziale [" in oSheet.getCellByPosition(8, m).String:
                    Flags.text = '2'
                    HPeso.text = ''
                elif 'PARTITA IN CONTO PROVVISORIO' in Descrizione.text:
                    Flags.text = '16'
                else:
                    Flags.text = '0'
##########################
                if 'DETRAE LA PARTITA IN CONTO PROVVISORIO' in Descrizione.text:
                    Flags.text = '32'
                if ' - vedi voce n. ' in Descrizione.text:
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
def firme_in_calce_run (arg=None):
    oDialogo_attesa = dlg_attesa()# avvia il diaolgo di attesa che viene chiuso alla fine con 
    '''
    Inserisce (in COMPUTO o VARIANTE) un riepilogo delle categorie
    ed i dati necessari alle firme
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()

    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in ('Analisi di Prezzo', 'Elenco Prezzi'):
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
        oSheet.getCellRangeByPosition (0,lRowF,100,lRowF+15-1).CellStyle = "Ultimus_centro"
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
        flags = (oDoc.createInstance('com.sun.star.sheet.CellFlags.FORMULA'))
        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)
        oSheet.getCellRangeByPosition (1,riga_corrente+3,1,riga_corrente+3).CellStyle = 'ULTIMUS'
        oSheet.getCellByPosition(1 , riga_corrente+5).Formula = 'Il progettista'
        oSheet.getCellByPosition(1 , riga_corrente+6).Formula = '=CONCATENATE("(";$S2.$C$13;")")'

    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
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
        oSheet.getCellByPosition(2 , riga_corrente).String = 'Riepilogo Categorie'
        oSheet.getCellByPosition(ii , riga_corrente).String = 'Incidenze %'
        oSheet.getCellByPosition(vv , riga_corrente).String = 'Importi €'
        inizio_gruppo = riga_corrente
        riga_corrente += 1
        for i in range (0, lRowF):
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
                oSheet.getCellByPosition(vv , riga_corrente).CellStyle = 'Ultimus_totali'
                riga_corrente += 1
            elif oSheet.getCellByPosition(1 , i).CellStyle == 'Livello-1-scritta':
                #~ chi(riga_corrente)
                oSheet.getRows().insertByIndex(riga_corrente,1)
                oSheet.getCellByPosition(1 , riga_corrente).Formula = '=B' + str(i+1) 
                oSheet.getCellByPosition(1 , riga_corrente).CellStyle = 'Ultimus_destra'
                oSheet.getCellByPosition(2 , riga_corrente).Formula = '=CONCATENATE ("   ";C' + str(i+1) + ')'
                #~ chi(formulaSCat)
                oSheet.getCellByPosition(ii , riga_corrente).Formula = '=' + col + str(riga_corrente+1) + '/' + col + str(lRowF) + '*100'
                oSheet.getCellByPosition(ii, riga_corrente).CellStyle = 'Ultimus %'
                oSheet.getCellByPosition(vv , riga_corrente).Formula = '='+ col + str(i+1) 
                oSheet.getCellByPosition(vv , riga_corrente).CellStyle = 'Ultimus_bordo'
                riga_corrente += 1
            elif oSheet.getCellByPosition(1 , i).CellStyle == 'livello2 valuta':
                #~ chi(riga_corrente)
                oSheet.getRows().insertByIndex(riga_corrente,1)
                oSheet.getCellByPosition(1 , riga_corrente).Formula = '=B' + str(i+1) 
                oSheet.getCellByPosition(1 , riga_corrente).CellStyle = 'Ultimus_destra'
                oSheet.getCellByPosition(2 , riga_corrente).Formula = '=CONCATENATE ("      ";C' + str(i+1) + ')'
                #~ chi(formulaSCat)
                oSheet.getCellByPosition(ii , riga_corrente).Formula = '=' + col + str(riga_corrente+1) + '/' + col + str(lRowF) + '*100'
                oSheet.getCellByPosition(ii, riga_corrente).CellStyle = 'Ultimus %'
                oSheet.getCellByPosition(vv , riga_corrente).Formula = '='+ col + str(i+1) 
                oSheet.getCellByPosition(vv , riga_corrente).CellStyle = 'ULTIMUS'
                riga_corrente += 1
        #~ riga_corrente +=1
     
        oSheet.getCellRangeByPosition (2,inizio_gruppo,vv,inizio_gruppo).CellStyle = "Ultimus_centro"

        oSheet.getCellByPosition(2 , riga_corrente).String= 'T O T A L E   €'
        oSheet.getCellByPosition(2 , riga_corrente).CellStyle = 'Ultimus_destra'
        oSheet.getCellByPosition(vv , riga_corrente).Formula = '=' + col + str(lRowF) 
        oSheet.getCellByPosition(vv , riga_corrente).CellStyle = 'Ultimus_Bordo_sotto'
        fine_gruppo = riga_corrente
    #~ DATA
        oSheet.getCellByPosition(2 , riga_corrente+3).Formula = '=CONCATENATE("Data, ";TEXT(NOW();"DD/MM/YYYY"))'
    #~ consolido il risultato
        oRange = oSheet.getCellByPosition(2 , riga_corrente+3)
        flags = (oDoc.createInstance('com.sun.star.sheet.CellFlags.FORMULA'))
        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)
        
        oSheet.getCellByPosition(2 , riga_corrente+5).Formula = 'Il Progettista'
        oSheet.getCellByPosition(2 , riga_corrente+6).Formula = '=CONCATENATE ("(";$S2.$C$13;")")'
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

def next_voice (lrow, n=1):
    '''
    lrow { double }   : riga di riferimento
    n    { integer }  : se 0 sposta prima della voce corrente
                        se 1 sposta dopo della voce corrente
    sposta il cursore prima o dopola voce corrente restituento un idrow
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ n =0
    #~ lrow = Range2Cell()[1]
    fine = ultima_voce(oSheet)+1
    if lrow >= fine:
        return lrow

    if oSheet.getCellByPosition(0, lrow).CellStyle in siVoce:
        if n==0:
            sopra = Circoscrive_Voce_Computo_Att (lrow).RangeAddress.StartRow
            lrow = sopra
        elif n==1:
            sotto = Circoscrive_Voce_Computo_Att (lrow).RangeAddress.EndRow
            lrow = sotto+1
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
def ANALISI_IN_ELENCOPREZZI (arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    try:
        oSheet = oDoc.CurrentController.ActiveSheet
        if oSheet.Name != 'Analisi di Prezzo':
            return
        oDoc.enableAutomaticCalculation(False) # blocco il calcolo automatico
        sStRange = Circoscrive_Analisi (Range2Cell()[1])
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
        _gotoCella (1, 3)
    except:
        oDoc.enableAutomaticCalculation(True)
    
def Circoscrive_Analisi (lrow):
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
        if oSheet.getCellByPosition (0, lrow).CellStyle == stili_analisi[0]:
            lrowS=lrow
        else:
            while oSheet.getCellByPosition(0, lrow).CellStyle != stili_analisi[0]:
                lrow = lrow-1
            lrowS=lrow
        lrow = lrowS
        while oSheet.getCellByPosition (0, lrow).CellStyle != stili_analisi[-1]:
            lrow=lrow+1
        lrowE=lrow
    celle=oSheet.getCellRangeByPosition(0,lrowS,250,lrowE)
    return celle
def Circoscrive_Voce_Computo_Att (lrow):
    '''
    lrow    { double }  : riga di riferimento per
                        la selezione dell'intera voce

    Circoscrive una voce di computo, variante o contabilità
    partendo dalla posizione corrente del cursore
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ lrow = Range2Cell()[1]
    #~ if oSheet.Name in ('VARIANTE', 'COMPUTO','CONTABILITA'):
    if oSheet.getCellByPosition(0, lrow).CellStyle in ('comp progress', 'comp 10 s', 'Comp Start Attributo', 'Comp End Attributo', 'Comp Start Attributo_R', 'comp 10 s_R', 'Comp End Attributo_R', 'Livello-1-scritta', 'livello2 valuta'):
        if oSheet.getCellByPosition (0, lrow).CellStyle in ('Comp Start Attributo', 'Comp Start Attributo_R'):
            lrowS=lrow
        else:
            while oSheet.getCellByPosition(0, lrow).CellStyle not in ('Comp Start Attributo', 'Comp Start Attributo_R'):
                lrow = lrow-1
            lrowS=lrow
        lrow = lrowS
        ### cerco l'ultima riga
        while oSheet.getCellByPosition (0, lrow).CellStyle not in ('Comp End Attributo', 'Comp End Attributo_R'):
            lrow=lrow+1
        lrowE=lrow
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
    Azzera la quantità di una voce
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        lrow = Range2Cell()[1]
        sStRange = Circoscrive_Voce_Computo_Att (lrow)
        sStRange.RangeAddress
        inizio = sStRange.RangeAddress.StartRow
        fine = sStRange.RangeAddress.EndRow
        _gotoCella(2, fine-1)
        if oSheet.getCellByPosition(2, fine-1).String == '*** VOCE AZZERATA ***':
            ### elimino il colore di sfondo
            oSheet.getCellRangeByPosition(0, inizio, 250, fine).clearContents(HARDATTR)
            
            oSheet.getRows().removeByIndex(fine-1, 1)
            _gotoCella(2, fine-2)
        else:
            Copia_riga_Ent()
            oSheet.getCellByPosition(2, fine).String = '*** VOCE AZZERATA ***'
            oSheet.getCellByPosition(5, fine).Formula = '=-SUBTOTAL(9;J' + str(inizio+1) + ':J' + str(fine) + ')'
            ### cambio il colore di sfondo
            oDoc.CurrentController.select(sStRange)
            ctx = XSCRIPTCONTEXT.getComponentContext()
            desktop = XSCRIPTCONTEXT.getDesktop()
            oFrame = desktop.getCurrentFrame()
            dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
            oProp = PropertyValue()
            oProp.Name = 'BackgroundColor'
            oProp.Value = 8421504
            properties = (oProp,)
            dispatchHelper.executeDispatch(oFrame, '.uno:BackgroundColor', '', 0, properties)
            ###
########################################################################
def copia_riga_computo(lrow):
    '''
    Inserisce una nuova riga di misurazione nel computo
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ lrow = Range2Cell()[1]
    stile = oSheet.getCellByPosition(1, lrow).CellStyle
    if stile in ('comp Art-EP', 'comp Art-EP_R', 'Comp-Bianche in mezzo'):#'Comp-Bianche in mezzo Descr', 'comp 1-a', 'comp sotto centro'):# <stili computo
        sStRange = Circoscrive_Voce_Computo_Att (lrow)
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
    if stile in ('comp Art-EP_R', 'Data_bianca', 'Comp-Bianche in mezzo_R'):
        sStRange = Circoscrive_Voce_Computo_Att (lrow)
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
    if stile in ('An-lavoraz-desc', 'An-lavoraz-Cod-sx'):
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
def Copia_riga_Ent(arg=None): #Aggiungi Componente - capisce su quale tipologia di tabelle è
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    nome_sheet = oSheet.Name
    if nome_sheet in ('COMPUTO', 'VARIANTE'):
        copia_riga_computo(lrow)
    elif nome_sheet == 'CONTABILITA':
        copia_riga_contab(lrow)
    elif nome_sheet == 'Analisi di Prezzo':
        copia_riga_analisi(lrow)
########################################################################
def debug_tipo_di_valore(arg=None):
#~ def debug(arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.getCellByPosition(2, 5).Type.value == 'FORMULA':
        MsgBox(oSheet.getCellByPosition(9, 5).Formula)
########################################################################
def debug_clipboard(arg=None):
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
    #~ createUnoService = (XSCRIPTCONTEXT.getComponentContext().getServiceManager().createInstance)
    #~ oTR = createUnoListener("Tr_", "com.sun.star.datatransfer.XTransferable")
    oClip.setContents( oTR, None )
    sTxtCString = sText
    oClip.flushClipboard()
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
def Range2Cell ():
    '''
    Restituisce la tupla (IDcolonna, IDriga) della posizione corrnete
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
    return (nCol,nRow)
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
def Numera_Voci(bit=1):#
    '''
    bit { integer }  : 1 rinumera tutto
                       0 rinumera dalla voce corrente in giù
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lastRow = getLastUsedCell(oSheet).EndRow+1
    lrow = Range2Cell()[1]
    n = 1
    if bit==0:
        for x in reversed(range(0, lrow)):
            if oSheet.getCellByPosition(0, x).CellStyle == 'comp progress':
                n = oSheet.getCellByPosition (0,x).Value +1
                break
        for row in range(lrow,lastRow):
            if oSheet.getCellByPosition (1,row).CellStyle == 'comp Art-EP' or oSheet.getCellByPosition (1,row).CellStyle == 'comp Art-EP_R':
                oSheet.getCellByPosition (0,row).Value = n
                n +=1
    if bit==1:
        for row in range(0,lastRow):
            if oSheet.getCellByPosition (1,row).CellStyle == 'comp Art-EP' or oSheet.getCellByPosition (1,row).CellStyle == 'comp Art-EP_R':
                oSheet.getCellByPosition (0,row).Value = n
                n = n+1
########################################################################
# ins_voce_computo #####################################################
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
    #~ oSheet.getCellByPosition (0,lrow).CellStyle = 'Comp Start Attributo'
    #~ oSheet.getCellRangeByPosition (0,lrow,30,lrow).CellStyle = 'Comp-Bianche sopra'
    #~ oSheet.getCellByPosition (2,lrow).CellStyle = 'Comp-Bianche sopraS'
    #~
    #~ oSheet.getCellByPosition (0,lrow+1).CellStyle = 'comp progress'
    #~ oSheet.getCellByPosition (1,lrow+1).CellStyle = 'comp Art-EP'
    #~ oSheet.getCellRangeByPosition (2,lrow+1,8,lrow+1).CellStyle = 'Comp-Bianche in mezzo Descr'
    #~ oSheet.getCellRangeByPosition (2,lrow+1,8,lrow+1).merge(True)
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
    for i in range (3, 10):
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
# sistemo i LINK dei tagG nelle righe sopra al tag vero e prorio...
    #~ oSheet.getCellByPosition(31, lrow+2).Formula = '=AF$'+str(lrow+4)
    #~ oSheet.getCellByPosition(32, lrow+2).Formula = '=AG$'+str(lrow+4)
    #~ oSheet.getCellByPosition(33, lrow+2).Formula = '=AH$'+str(lrow+4)
    #~ oSheet.getCellByPosition(34, lrow+2).Formula = '=AI$'+str(lrow+4)
    #~ oSheet.getCellByPosition(35, lrow+2).Formula = '=AJ$'+str(lrow+4)
    #~ oSheet.getCellByPosition(31, lrow+1).Formula = '=AF$'+str(lrow+4)
    #~ oSheet.getCellByPosition(32, lrow+1).Formula = '=AG$'+str(lrow+4)
    #~ oSheet.getCellByPosition(33, lrow+1).Formula = '=AH$'+str(lrow+4)
    #~ oSheet.getCellByPosition(34, lrow+1).Formula = '=AI$'+str(lrow+4)
    #~ oSheet.getCellByPosition(35, lrow+1).Formula = '=AJ$'+str(lrow+4)
    #~ oSheet.getCellByPosition(31, lrow).Formula = '=AF$'+str(lrow+4)
    #~ oSheet.getCellByPosition(32, lrow).Formula = '=AG$'+str(lrow+4)
    #~ oSheet.getCellByPosition(33, lrow).Formula = '=AH$'+str(lrow+4)
    #~ oSheet.getCellByPosition(34, lrow).Formula = '=AI$'+str(lrow+4)
    #~ oSheet.getCellByPosition(35, lrow).Formula = '=AJ$'+str(lrow+4)
    if oSheet.getCellByPosition(31, lrow-1).CellStyle in ('livello2 valuta', 'Livello-0-scritta', 'Livello-1-scritta', 'compTagRiservato'):
        oSheet.getCellByPosition(31, lrow+3).Value = oSheet.getCellByPosition(31, lrow-1).Value
        oSheet.getCellByPosition(32, lrow+3).Value = oSheet.getCellByPosition(32, lrow-1).Value
        oSheet.getCellByPosition(33, lrow+3).Value = oSheet.getCellByPosition(33, lrow-1).Value
    #~ celle=oSheet.getCellRangeByPosition(0, lrow, 43,lrow+3)# 'seleziona la cella
    #~ oDoc.CurrentController.select(celle)
    #~ celle.Rows.OptimalHeight = True
########################################################################
    _gotoCella(1,lrow+1)
########################################################################
# ins_voce_computo #####################################################
def ins_voce_computo(arg=None): #TROPPO LENTA
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    if oSheet.getCellByPosition(0, lrow).CellStyle in (noVoce + siVoce):
        lrow = next_voice(lrow, 1)
    else:
        return
    ins_voce_computo_grezza(lrow)
    Numera_Voci(0)
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
        setTabColor (12189608)
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
            _gotoSheet ('Analisi di Prezzo')
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
def inserisci_Riga_rossa (arg=None):
    '''
    Inserisce la riga rossa di chiusura degli elaborati
    Questa riga è un rigerimento per varie operazioni
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    nome = oSheet.Name
    if nome in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
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
        lrow = ultima_voce(oSheet) + 1
        oSheet.getCellByPosition( 0, lrow).String = 'Fine elenco'
        oSheet.getCellRangeByPosition(0,lrow,26,lrow).CellStyle='Riga_rossa_Chiudi' 
    oSheet.getCellByPosition(2, lrow).String = 'Questa riga NON deve essere cancellata, MAI!!! (ma può rimanere tranquillamente NASCOSTA!)'
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
def struttura_Elenco (arg=None):
    '''
    Dà una tonalità di colore, diverso dal colore dello stile cella, alle righe
    che non hanno il prezzo, come i titoli di capitolo e sottocapitolo.
    '''
    col1 = 16771481
    col2 = 16771501
    col3 = 16771521 #chiaro - sfondo celle elenco prezzi
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    lista = list()
    test = getLastUsedCell(oSheet).EndRow
    for n in range (3, test):#
        if len(oSheet.getCellByPosition(0, n).String.split('.')) == 2:
            oSheet.getCellRangeByPosition(0, n, 26, n).CellBackColor = col1
            sopra = n+1
            for n in range (sopra+1, test):
                if len(oSheet.getCellByPosition(0, n).String.split('.')) == 2:
                    sotto = n-1
                    lista.append((sopra, sotto))
                    break
        if len(oSheet.getCellByPosition(0, n).String.split('.')) == 4 and \
        oSheet.getCellByPosition(4, n).String == '':
            oSheet.getCellRangeByPosition(0, n, 26, n).CellBackColor = col2
            oCellRangeAddr.StartRow = n
            oCellRangeAddr.EndRow = n
            oSheet.group(oCellRangeAddr,1)
    for el in lista:
        oCellRangeAddr.StartRow = el[0]
        oCellRangeAddr.EndRow = el[1]
        oSheet.group(oCellRangeAddr,1)
        #~ oSheet.getCellRangeByPosition(0, el[0], 0, el[1]).Rows.IsVisible=False
    return
    #~ la parte che segue è servita a riordinare le descrizioni del prezzario Umbria 2016
        #~ if len(oSheet.getCellByPosition(2, n).String.split('.')) == 4 and \
        #~ oSheet.getCellByPosition(7, n).String == '':
            #~ des0 = oSheet.getCellByPosition(4, n).String
            #~ des = ''
        #~ if len(oSheet.getCellByPosition(2, n).String.split('.')) == 4 and \
        #~ oSheet.getCellByPosition(7, n).String != '':
            #~ des = des0 +'\n- ' + oSheet.getCellByPosition(4, n).String
            #~ oSheet.getCellByPosition(13, n).String = des
########################################################################
# XML_import ###########################################################
def XML_import (arg=None):
    New_file.listino()
    '''
    Routine di importazione di un prezziario XML formato SIX. Molto
    liberamente tratta da PreventARES https://launchpad.net/preventares
    di <Davide Vescovini> <davide.vescovini@gmail.com>
    '''
    try:
        filename = filedia('Scegli il file XML-SIX da importare', '*.xml')
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
    # effettua il parsing di tutti gli elemnti dell'albero XMLsub nuova_voce_computo_at
    iter = tree.getiterator()
    listaSOA = []
    articolo = []
    articolo_modificato = ()
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
            #~ self.update_dati_generali (nome=None, cliente=None,
                                       #~ redattore=autore,
                                       #~ ricarico=1,
                                       #~ manodopera=None,
                                       #~ sicurezza=None,
                                       #~ indirizzo=None,
                                       #~ comune=None, provincia=None,
                                       #~ valuta=valuta)
        elif elem.tag == '{six.xsd}categoriaSOA':
            soaId = elem.get('soaId')
            soaCategoria = elem.get ('soaCategoria')
            soaDescrizione = elem.find('{six.xsd}soaDescrizione')
            if soaDescrizione != None:
                breveSOA = soaDescrizione.get('breve')
            voceSOA = (soaCategoria, soaId, breveSOA)
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
                if len (elem.findall('{six.xsd}udmDescrizione')) == 1:
                    #~ unita_misura = elem.getchildren()[0].get('breve')
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
            try:
                if len (elem.findall('{six.xsd}prdDescrizione')) == 1:
                    desc_breve = elem.findall('{six.xsd}prdDescrizione')[0].get('breve')
                    desc_estesa = elem.findall('{six.xsd}prdDescrizione')[0].get('estesa')
                else:
            #descrizione voce
                    if elem.findall('{six.xsd}prdDescrizione')[0].get('lingua') == lingua_scelta:
                        idx = 0 #ITALIANO
                    else:
                        idx = 1 #TEDESCO
                    desc_breve = elem.findall('{six.xsd}prdDescrizione')[idx].get('breve')
                    desc_estesa = elem.findall('{six.xsd}prdDescrizione')[idx].get('estesa')
                if desc_breve == None:
                    desc_breve = ''
                if desc_estesa == None:
                    desc_estesa = ''
                if len(desc_breve) > len (desc_estesa):
                    desc_voce = desc_breve
                else:
                    desc_voce = desc_estesa
            except IndexError:
                pass
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
                valore = ''
                quantita = ''
#~ Modifiche introdotte da Valerio De Angelis che ringrazio
            # Riarrangio i dati di ogni articolo così da formare una tupla 1D
            # l'idea è creare un array 2D e caricarlo direttamente nel foglio in una singola operazione
            vuoto = ''
            elem_7 = ''
            elem_11 = ''
            if mdo != '' and mdo != 0:
                elem_7 = mdo/100
            if sicurezza != '' and valore != '':
                elem_11 = valore*sicurezza/100
            # Nota che ora articolo_modificato non è più una lista ma una tupla,
            # riguardo al motivo, vedi commenti in basso
            articolo_modificato =  (prod_id,          #0  colonna
                                    vuoto,            #1  colonna
                                    tariffa,          #2  colonna
                                    vuoto,            #3  colonna
                                    desc_voce,        #4  colonna
                                    vuoto,            #5  colonna
                                    unita_misura,     #6  colonna
                                    valore,           #7  colonna
                                    elem_7,           #8  colonna %
                                    vuoto,            #9  colonna
                                    vuoto,            #10 colonna
                                    elem_11)          #11 colonna %
            lista_articoli.append(articolo_modificato)
# compilo la tabella ###################################################
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('Listino')
    if oSheet.getCellByPosition(5, 4).String == '': #se la cella F5 è vuota, elimina riga
        oRangeAddress = oSheet.getCellRangeByPosition(0,4,1,4).getRangeAddress()
        oSheet.removeRange(oRangeAddress, 3) # Mode.ROWS
# nome del prezzario ###################################################
    oSheet.getCellByPosition(2, 0).String = '.\n' + nome
    # Siccome setDataArray pretende una tupla (array 1D) o una tupla di tuple (array 2D)
    # trasformo la lista_articoli da una lista di tuple a una tupla di tuple
    lista_come_array = tuple(lista_articoli)
    # Parametrizzo il range di celle a seconda della dimensione della lista
    scarto_colonne = 0 # numero colonne da saltare a partire da sinistra
    scarto_righe = 5 # numero righe da saltare a partire dall'alto
    colonne_lista = len(lista_come_array[1]) # numero di colonne necessarie per ospitare i dati
    righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati
    oRange = oSheet.getCellRangeByPosition( scarto_colonne,
                                            scarto_righe,
                                            colonne_lista + scarto_colonne - 1, # l'indice parte da 0
                                            righe_lista + scarto_righe - 1)
    oRange.setDataArray(lista_come_array)
    oSheet.getCellRangeByPosition (0,scarto_righe,5,righe_lista + scarto_righe - 1).CellStyle = 'List-stringa-sin'
    oSheet.getCellRangeByPosition (6,scarto_righe,6,righe_lista + scarto_righe - 1).CellStyle = 'List-stringa-centro'
    oSheet.getCellRangeByPosition (7,scarto_righe,11,righe_lista + scarto_righe - 1).CellStyle = 'List-num-euro'
    oSheet.getCellRangeByPosition (8,scarto_righe,8,righe_lista + scarto_righe - 1).CellStyle = 'List-%'
    oSheet.getCellRangeByPosition (10,scarto_righe,10,righe_lista + scarto_righe - 1).CellStyle = 'List-%'
    MsgBox('Importazione eseguita con successo\n in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!','')
# XML_import ###########################################################
########################################################################
def XML_import_BOLZANO (arg=None):
    '''
    Routine di importazione di un prezziario XML formato SIX. Molto
    liberamente tratta da PreventARES https://launchpad.net/preventares
    di <Davide Vescovini> <davide.vescovini@gmail.com>
    *Versione bilingue*
    '''
    New_file.listino()
    filename = filedia('Scegli il file XML-SIX da convertire...','*.xml')
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
            soaCategoria = elem.get ('soaCategoria')
            soaDescrizione = elem.find('{six.xsd}soaDescrizione')
            if soaDescrizione != None:
                breveSOA = soaDescrizione.get('breve')
            voceSOA = (soaCategoria, soaId, breveSOA)
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
            #~ try:
            if len (elem.findall('{six.xsd}udmDescrizione')) == 1:
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
            #~ except IndexError:
                #~ pass
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
            if len (elem.findall('{six.xsd}prdDescrizione')) == 1:
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
            if len(desc_breve) > len (desc_estesa):
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
            articolo_modificato =  (prod_id,          #0  colonna
                                    vuoto,            #1  colonna
                                    tariffa,          #2  colonna
                                    vuoto,            #3  colonna
                                    desc_voce,        #4  colonna
                                    vuoto,            #5  colonna
                                    unita_misura,     #6  colonna
                                    valore,           #7  colonna
                                    elem_7,           #8  colonna %
                                    vuoto,            #9  colonna
                                    vuoto,            #10 colonna
                                    elem_11)          #11 colonna %
            lista_articoli.append(articolo_modificato)
# compilo la tabella ###################################################
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('Listino')
    oSheet.getCellByPosition(6, 0).Value = 1 #Livello di accodamento
    if oSheet.getCellByPosition(5, 4).String == '': #se la cella F5 è vuota, elimina riga
        oRangeAddress = oSheet.getCellRangeByPosition (0,4,1,4).getRangeAddress()
        oSheet.removeRange(oRangeAddress, 3) # Mode.ROWS
# nome del prezzario ###################################################
    oSheet.getCellByPosition(2, 0).String = '.\n' + nome
    # Siccome setDataArray pretende una tupla (array 1D) o una tupla di tuple (array 2D)
    # trasformo la lista_articoli da una lista di tuple a una tupla di tuple
    lista_come_array = tuple(lista_articoli)
    # Parametrizzo il range di celle a seconda della dimensione della lista
    scarto_colonne = 0 # numero colonne da saltare a partire da sinistra
    scarto_righe = 5 # numero righe da saltare a partire dall'alto
    colonne_lista = len(lista_come_array[1]) # numero di colonne necessarie per ospitare i dati
    righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati
    oRange = oSheet.getCellRangeByPosition( scarto_colonne,
                                            scarto_righe,
                                            colonne_lista + scarto_colonne - 1, # l'indice parte da 0
                                            righe_lista + scarto_righe - 1)
    oRange.setDataArray(lista_come_array)
    oSheet.getCellRangeByPosition (0,scarto_righe,5,righe_lista + scarto_righe - 1).CellStyle = 'List-stringa-sin'
    oSheet.getCellRangeByPosition (6,scarto_righe,6,righe_lista + scarto_righe - 1).CellStyle = 'List-stringa-centro'
    oSheet.getCellRangeByPosition (7,scarto_righe,11,righe_lista + scarto_righe - 1).CellStyle = 'List-num-euro'
    oSheet.getCellRangeByPosition (8,scarto_righe,8,righe_lista + scarto_righe - 1).CellStyle = 'List-%'
    oSheet.getCellRangeByPosition (10,scarto_righe,10,righe_lista + scarto_righe - 1).CellStyle = 'List-%'
    MsgBox('Importazione eseguita con successo\n in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!','')
# XML_import_BOLZANO ###################################################
########################################################################
# parziale_core ########################################################
def parziale_core(lrow):
    #~ lrow = 7
    if lrow == 0:
        return
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    sStRange = Circoscrive_Voce_Computo_Att (lrow)
    #~ sStRange.RangeAddress
    sopra = sStRange.RangeAddress.StartRow
    sotto = sStRange.RangeAddress.EndRow

    if oSheet.Name in ('COMPUTO','VARIANTE'):
        if oSheet.getCellByPosition (0, lrow).CellStyle == 'comp 10 s' and \
        oSheet.getCellByPosition (1, lrow).CellStyle == 'Comp-Bianche in mezzo' and \
        oSheet.getCellByPosition (2, lrow).CellStyle == 'comp 1-a' or \
        oSheet.getCellByPosition (0, lrow).CellStyle == 'Comp End Attributo':
            oSheet.getRows().insertByIndex(lrow, 1)
            oSheet.getCellByPosition (1, lrow).CellStyle = 'Comp-Bianche in mezzo'
            oSheet.getCellRangeByPosition (2, lrow, 7, lrow).CellStyle = 'comp sotto centro'
            oSheet.getCellByPosition (8, lrow).CellStyle = 'comp sotto BiancheS'
            oSheet.getCellByPosition (9, lrow).CellStyle = 'Comp-Variante num sotto'

            #~ oSheet.getCellByPosition(31, lrow).Formula ='=AF$' + str(sotto+2)
            #~ oSheet.getCellByPosition(32, lrow).Formula ='=AG$' + str(sotto+2)
            #~ oSheet.getCellByPosition(33, lrow).Formula ='=AH$' + str(sotto+2)
            #~ oSheet.getCellByPosition(34, lrow).Formula ='=AI$' + str(sotto+2)
            #~ oSheet.getCellByPosition(35, lrow).Formula ='=AJ$' + str(sotto+2)

            oSheet.getCellByPosition (8, lrow).Formula = '''=CONCATENATE("Parziale [";VLOOKUP(B'''+ str(sopra+2) + ''';elenco_prezzi;3;FALSE());"]")'''

            for i in reversed(range(0, lrow)):
                if oSheet.getCellByPosition (9, i-1).CellStyle in ('vuote2', 'Comp-Variante num sotto'):
                    i
                    break

            oSheet.getCellByPosition(9, lrow).Formula = "=SUBTOTAL(9;J" + str(i) + ":J" + str(lrow+1) + ")"
########################################################################
# abs2name ############################################################
def abs2name(nCol, nRow):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    idvoce = oSheet.getCellByPosition(nCol, nRow).AbsoluteName.split('$')
    return idvoce[2]+idvoce[3]
########################################################################
# vedi_voce ############################################################
def vedi_voce(riga_corrente,vRif,flags=''):
    """(riga d'inserimento, riga di riferimento)"""
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    sStRange = Circoscrive_Voce_Computo_Att (vRif)
    sStRange.RangeAddress
    idv = sStRange.RangeAddress.StartRow +1
    sotto = sStRange.RangeAddress.EndRow
    art = abs2name (1, idv)
    idvoce = abs2name (0, idv)
    quantity = abs2name (9, sotto)
    um = 'VLOOKUP(' + art + ';elenco_prezzi;3;FALSE())'
    oSheet.getCellByPosition(2, riga_corrente).Formula='=CONCATENATE("";" - vedi voce n. ";TEXT(' + idvoce +';"@");" - art. ";' + art + ';" [";' + um + ';"]"'
    if flags in ('32769', '32801'):
        oSheet.getCellByPosition(5, riga_corrente).Formula='=-' + quantity
    else:
        oSheet.getCellByPosition(5, riga_corrente).Formula='=' + quantity
########################################################################
def strall (el, n=3):
    #~ el ='o'
    while len(el) < n:
        el = '0' + el
    return el

# XPWE_in ##########################################################
def XPWE_in(arg=None): #(filename):
    oDoc = XSCRIPTCONTEXT.getDocument()
    ###
    #~ oDoc.enableAutomaticCalculation(False) # blocco il calcolo automatico
    #~ oDoc.addActionLock
    #~ oDoc.lockControllers #disattiva l'eco a schermo
    ###
    oDialogo_attesa = dlg_attesa()
    if oDoc.getSheets().hasByName('S2') == False:
        MsgBox('Puoi usare questo comando da un file di computo nuovo o già esistente.','Avviso!')
        return
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
    CapCat = dati.find('PweDGCapitoliCategorie')
###
#PweDGSuperCapitoli
    if CapCat.find('PweDGSuperCapitoli'):
        PweDGSuperCapitoli = CapCat.find('PweDGSuperCapitoli').getchildren()
        lista_supcap = list()
        for elem in PweDGSuperCapitoli:
            id_sc = elem.get('ID')
            dessintetica = elem.find('DesSintetica').text
            percentuale = elem.find('Percentuale').text
            diz = dict ()
            diz['id_sc'] = id_sc
            diz['dessintetica'] = dessintetica
            diz['percentuale'] = percentuale
            lista_supcap.append(diz)
###
#PweDGCapitoli
    if CapCat.find('PweDGCapitoli'):
        PweDGCapitoli = CapCat.find('PweDGCapitoli').getchildren()
        lista_cap = list()
        for elem in PweDGCapitoli:
            id_sc = elem.get('ID')
            dessintetica = elem.find('DesSintetica').text
            percentuale = elem.find('Percentuale').text
            diz = dict ()
            diz['id_sc'] = id_sc
            diz['dessintetica'] = dessintetica
            diz['percentuale'] = percentuale
            lista_cap.append(diz)
###
#PweDGSubCapitoli
    if CapCat.find('PweDGSubCapitoli'):
        PweDGSubCapitoli = CapCat.find('PweDGSubCapitoli').getchildren()
        lista_subcap = list()
        for elem in PweDGSubCapitoli:
            id_sc = elem.get('ID')
            dessintetica = elem.find('DesSintetica').text
            percentuale = elem.find('Percentuale').text
            diz = dict ()
            diz['id_sc'] = id_sc
            diz['dessintetica'] = dessintetica
            diz['percentuale'] = percentuale
            lista_subcap.append(diz)
###
#PweDGSuperCategorie
    if CapCat.find('PweDGSuperCategorie'):
        PweDGSuperCategorie = CapCat.find('PweDGSuperCategorie').getchildren()
        lista_supcat = list()
        for elem in PweDGSuperCategorie:
            id_sc = elem.get('ID')
            dessintetica = elem.find('DesSintetica').text
            try:
                percentuale = elem.find('Percentuale').text
            except AttributeError:
                percentuale = '0'
            supcat = (id_sc, dessintetica, percentuale)
            lista_supcat.append(supcat)
        #~ MsgBox(str(lista_supcat),'') ; return
###
#PweDGCategorie
    if CapCat.find('PweDGCategorie'):
        PweDGCategorie = CapCat.find('PweDGCategorie').getchildren()
        lista_cat = list()
        for elem in PweDGCategorie:
            id_sc = elem.get('ID')
            dessintetica = elem.find('DesSintetica').text
            try:
                percentuale = elem.find('Percentuale').text
            except AttributeError:
                percentuale = '0'
            cat = (id_sc, dessintetica, percentuale)
            lista_cat.append(cat)
        #~ MsgBox(str(lista_cat),'')
###
#PweDGSubCategorie
    if CapCat.find('PweDGSubCategorie'):
        PweDGSubCategorie = CapCat.find('PweDGSubCategorie').getchildren()
        lista_subcat = list()
        for elem in PweDGSubCategorie:
            id_sc = elem.get('ID')
            dessintetica = elem.find('DesSintetica').text
            try:
                percentuale = elem.find('Percentuale').text
            except AttributeError:
                percentuale = '0'
            subcat = (id_sc, dessintetica, percentuale)
            lista_subcat.append(subcat)
        #~ MsgBox(str(lista_subcat),'') ; return
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
        desridotta = elem.find('DesBreve').text
        desbreve = elem.find('DesBreve').text
        if elem.find('UnMisura').text != None:
            unmisura = elem.find('UnMisura').text
        else:
            unmisura = ''
        prezzo1 = elem.find('Prezzo1').text
        prezzo2 = elem.find('Prezzo2').text
        prezzo3 = elem.find('Prezzo3').text
        prezzo4 = elem.find('Prezzo4').text
        prezzo5 = elem.find('Prezzo5').text
        idspcap = elem.find('IDSpCap').text
        idcap = elem.find('IDCap').text
        idsbcap = elem.find('IDSbCap').text
        flags = elem.find('Flags').text
        data = elem.find('Data').text

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
        #~ chi (pweepanalisi)
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
        articolo_modificato =  (tariffa,
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
            analisi = list ()
            for el in EPARItem:
                id_an = el.get('ID')
                an_tipo = el.find('Tipo').text
                id_ep = el.find('IDEP').text
                an_des = el.find('Descrizione').text
                an_um = el.find('Misura').text
                an_qt = el.find('Qt').text.replace(' ','')
                an_pr = el.find('Prezzo').text.replace(' ','')
                an_fld = el.find('FieldCTL').text
                an_rigo = (id_ep, an_des, an_um, an_qt, an_pr)
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
            datamis = elem.find('DataMis').text
            flags = elem.find('Flags').text
            idspcat = elem.find('IDSpCat').text
            idcat = elem.find('IDCat').text
            idsbcat = elem.find('IDSbCat').text
            righi_mis = elem.getchildren()[-1].findall('RGItem')
            lista_rig = list()
            riga_misura = ()
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
                riga_misura =  (descrizione,
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
            new_id_l = (new_id, diz_misura)
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
        oSheet.getCellByPosition (2,2).String = oggetto
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
    except UnboundLocalError:
        pass
# compilo Elenco Prezzi ################################################
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    # Siccome setDataArray pretende una tupla (array 1D) o una tupla di tuple (array 2D)
    # trasformo la lista_articoli da una lista di tuple a una tupla di tuple
    lista_come_array = tuple(lista_articoli)
    # Parametrizzo il range di celle a seconda della dimensione della lista
    scarto_colonne = 0 # numero colonne da saltare a partire da sinistra
    scarto_righe = 3 # numero righe da saltare a partire dall'alto
    colonne_lista = len(lista_come_array[1]) # numero di colonne necessarie per ospitare i dati
    righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati

    oSheet.getRows().insertByIndex(4, righe_lista -1)

    oRange = oSheet.getCellRangeByPosition( scarto_colonne,
                                            scarto_righe,
                                            colonne_lista + scarto_colonne - 1, # l'indice parte da 0
                                            righe_lista + scarto_righe - 1)
    oRange.setDataArray(lista_come_array)
    doppioni()
### elimino le voci che hanno analisi
    for i in reversed(range(3, getLastUsedCell(oSheet).EndRow)):
        if oSheet.getCellByPosition(0, i).String in lista_tariffe_analisi:
            oSheet.getRows().removeByIndex(i, 1)
            
    if len(lista_misure) == 0:
        MsgBox("Importate n."+ str(len(lista_articoli)) +" voci dall'elenco prezzi\ndel file: " + filename, 'Avviso')
        oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
        oDoc.CurrentController.setActiveSheet(oSheet)
        oDoc.CurrentController.ZoomValue = 100
        oDialogo_attesa.endExecute()
        return
###
# Compilo Analisi di prezzo ############################################
    oDoc.CurrentController.ZoomValue = 400
    if len (lista_analisi) !=0:
        inizializza_analisi()
        oSheet = oDoc.getSheets().getByName('Analisi di Prezzo')
        for el in lista_analisi:
            sStRange = Circoscrive_Analisi (Range2Cell()[1])
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
    #~ Lib_LeenO('Voci_Sposta.elimina_voce') #rinvia a basic
    Lib_LeenO('Analisi.tante_analisi_in_ep') #rinvia a basic
# Inserisco i dati nel COMPUTO #########################################
    if arg == 'VARIANTE':
        Lib_LeenO('Computo.genera_variante')
    oSheet = oDoc.getSheets().getByName(arg)
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
        if nrighe > -1:
            EC = SC + len(el.get('lista_rig')[0])
            ER = SR + nrighe

            if nrighe > 0:
                oSheet.getRows().insertByIndex(SR, nrighe)

            oRangeAddress = oSheet.getCellRangeByPosition(0, SR-1, 250, SR-1).getRangeAddress()

            for n in range (SR, SR+nrighe):
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

                if mis[4] != None: #lunghezza
                    if any(o in mis[4] for o in ('+', '*', '/', '-','^',)):
                        oSheet.getCellByPosition(6, SR).Formula = '=' + str(mis[4]).split('=')[-1] # tolgo evenutali '=' in eccesso
                    else:
                        oSheet.getCellByPosition(6, SR).Value = eval(mis[4].replace(',','.'))
                else:
                    pass

                if mis[5] != None: #larghezza
                    if any(o in mis[5] for o in ('+', '*', '/', '-', '^',)):
                        oSheet.getCellByPosition(7, SR).Formula = '=' + str(mis[5]).split('=')[-1] # tolgo evenutali '=' in eccesso
                    else:
                        try:
                            eval(mis[5])
                            oSheet.getCellByPosition(7, SR).Value = eval(mis[5].replace(',','.'))
                        except:
                            oSheet.getCellByPosition(7, SR).Value = mis[5].replace(',','.')

                if mis[6] != None: #HPESO
                    if any(o in mis[6] for o in ('+', '*', '/', '-', '^',)):
                        oSheet.getCellByPosition(8, SR).Formula = '=' + str(mis[6]).split('=')[-1] # tolgo evenutali '=' in eccesso
                    else:
                        oSheet.getCellByPosition(8, SR).Value = eval(mis[6])
                if mis[8] == '2':
                    parziale_core(SR)
                    oSheet.getRows().removeByIndex(SR+1, 1)
                    descrizione =''

                va = oSheet.getCellByPosition(5, SR).Value
                vb = oSheet.getCellByPosition(6, SR).Value
                vc = oSheet.getCellByPosition(7, SR).Value
                vd = oSheet.getCellByPosition(8, SR).Value

                if mis[3] == None:
                    va =1
                else:
                    if '^' in mis[3]:
                        va = eval(mis[3].replace('^','**'))
                    else:
                        va = eval(mis[3])
                if vb ==0:
                    vb =1
                if vc ==0:
                    vc =1
                if vd ==0:
                    vd =1
                try:
                    if mis[3] != None: #partiuguali
                        if '-' in mis[7] and va*vb*vc*vd >0: #quantità
                            pu = '-1*(' + str(mis[3]) +')'
                        else:
                            pu = str(mis[3])
                        if any(o in pu for o in ('+', '*', '/', '-', '^',)):
                            oSheet.getCellByPosition(5, SR).Formula = '=' + pu.split('=')[-1] # tolgo evenutali '=' in eccesso

                        else:
                            oSheet.getCellByPosition(5, SR).Value = eval(pu)

                    if '-' in mis[7] and va*vb*vc*vd >0: #quantità
                        if mis[3] != None: #partiuguali
                            oSheet.getCellByPosition(5, SR).Formula =  '=-1*(' + str(mis[3]) +')'
                        else:
                            oSheet.getCellByPosition(5, SR).Value = eval('-1')

                    if mis[9] != '-2':
                        vedi = diz_vv.get(mis[9])
                        try:
                            vedi_voce(SR, vedi, mis[8])
                        except:
                            MsgBox("""Il file di origine è particolarmente disordinato.
Riordinando il computo trovo riferimenti a voci non ancora inserite.

Al termine dell'impotazione controlla la voce con tariffa """ + dict_articoli.get(ID).get('tariffa') +
"""\nalla riga n.""" + str(lrow+2) + """ del foglio, evidenziata qui a sinistra.""", 'Attenzione!')
                        lista_n = [va, vb, vc, vd]
                        lista_p = list()
                        if va*vb*vc*vd != 1:
                            for n in lista_n:
                                if n != 1:
                                    lista_p.append(n)
                                else:
                                    lista_p.append(0)
                            x = 0

                        lista_p.sort()

                        lista_p = lista_p [1:]
                        for n in reversed(lista_p):
                            if n== 0:
                                oSheet.getCellByPosition(8-x, SR).String = ''
                            else:
                                oSheet.getCellByPosition(8-x, SR).Value = n
                            x +=1

                except TypeError:
                    pass
                SR = SR+1
    Numera_Voci()
    try:
        Rinumera_TUTTI_Capitoli2()
    except:
        pass
    oDoc.CurrentController.ZoomValue = 100
    oDialogo_attesa.endExecute()
    #~ oDoc.enableAutomaticCalculation(True) # abilito il calcolo automatico
    MsgBox('Importazione eseguita con successo in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!        \n\nImporto € ' + oSheet.getCellByPosition(0, 1).String ,'')
    #~ MsgBox('Importazione eseguita con successo!','')
# XPWE_in ##########################################################
########################################################################
#VARIABILI GLOBALI:
Lmajor= 3 #'INCOMPATIBILITA'
Lminor= 14 #'NUOVE FUNZIONALITA'
Lsubv= "1.dev"#'CORREZIONE BUGS
noVoce = ('Livello-0-scritta', 'Livello-1-scritta', 'livello2 valuta', 'comp Int_colonna')
siVoce = ('Comp Start Attributo', 'comp progress', 'comp 10 s','Comp End Attributo', )
siVoce_R = ('Comp Start Attributo_R', 'comp 10 s_R','Comp End Attributo_R', )
stili_analisi = ('An.1v-Att Start', 'An-1_sigla', 'An-lavoraz-desc', 'An-lavoraz-Cod-sx', 'An-lavoraz-desc-CEN', 'An-sfondo-basso Att End', )
createUnoService = (
        XSCRIPTCONTEXT
        .getComponentContext()
        .getServiceManager()
        .createInstance
                    )
GetmyToolBarNames = ('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar', 'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ELENCO','private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ANALISI', 'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_COMPUTO', 'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CONTABILITA', )
#
sUltimus = ''
def ssUltimus (arg=None):
    '''
    Scrive la variabile globale che individua il Documento di Contabilità Corrente (DCC)
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
    return
########################################################################
def debugnn (sCella='', t=''):
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
def filedia (titolo='Scegli il file...', est='*.*',  mode=0):
    """
    titolo  { string }  : titolo del FilePicker
    est     { string }  : filtro di visualizzazione file
    mode    { integer } : modalità di gestione del file

    Apri file:  `mode in (0, 6, 7, 8, 9)`
    Salva file: `mode in (1, 2, 3, 4, 5, 10)`
    see: ('''http://api.libreoffice.org/docs/idl/ref/
            namespacecom_1_1sun_1_1star_1_1ui_1_1
            dialogs_1_1TemplateDescription.html''' )
    see: ('''http://stackoverflow.com/questions/30840736/
        libreoffice-how-to-create-a-file-dialog-via-python-macro''')
    """
    estensioni = {'*.*'   : 'Tutti i file (*.*)',
                '*.odt' : 'Writer (*.odt)',
                '*.ods' : 'Calc (*.ods)',
                '*.odb' : 'Base (*.odb)',
                '*.odg' : 'Draw (*.odg)',
                '*.odp' : 'Impress (*.odp)',
                '*.odf' : 'Math (*.odf)',
                '*.xpwe': 'Primus (*.xpwe)',
                '*.xml' : 'XML (*.xml)'
                }
    try:
        oFilePicker = createUnoService( "com.sun.star.ui.dialogs.OfficeFilePicker" )
        oFilePicker.initialize( ( mode,) )
        oFilePicker.Title = titolo

        app = estensioni.get(est)
        oFilePicker.appendFilter (app, est)
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
    #~ chi (res)
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
    l { integer } : livello (1 o 2)
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
        myrange = ('livello2 scritta mini', 'Livello-1-scritta minival', 'Comp TOTALI',)
        level = 2
    else:
        myrange = ('Livello-1-scritta mini val', 'Comp TOTALI',)
        level = 1

    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
        return
    lrowE = ultima_voce(oSheet)+1
    nextCap = lrowE
    for n in range (lrow+1, lrowE):
        if oSheet.getCellByPosition(18, n).CellStyle in myrange:
            nextCap = n + 1
            break
    for n,a,b in ((18, lrow+1, nextCap,), (24, lrow+1, lrowE+1,), (29, lrow+1, lrowE+1,), (30, lrow+1, nextCap,),):
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
    myrange = ('Comp End Attributo', 'Comp TOTALI',)
    if oSheet.getCellByPosition(0, lrow).CellStyle in (siVoce + siVoce_R) :
        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        sStRange = Circoscrive_Voce_Computo_Att (lrow)
        sopra = sStRange.RangeAddress.StartRow
        voce = oSheet.getCellByPosition(1, sopra+1).String
    else:
        MsgBox('Devi prima selezionare una voce di misurazione.','Avviso!')
        return
    fine = ultima_voce(oSheet)+1
    lista_pt = list()
    _gotoCella(0, 0)

    for n in range(0, fine):
        if oSheet.getCellByPosition(0, n).CellStyle in ('Comp Start Attributo','Comp Start Attributo_R'):
            sStRange = Circoscrive_Voce_Computo_Att (n)
            sopra = sStRange.RangeAddress.StartRow
            sotto = sStRange.RangeAddress.EndRow
            if oSheet.getCellByPosition(1, sopra+1).String != voce:
                lista_pt.append((sopra, sotto))
                lista_pt.append((sopra+2, sotto-1))
    #~ MsgBox(lista_pt)
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
        myrange = ('Livello-0-scritta', 'Comp TOTALI',)
        Dsopra = 1
        Dsotto = 1
    elif l == 1:
        stile = 'Livello-1-scritta'
        myrange = ('Livello-1-scritta', 'Livello-0-scritta', 'Comp TOTALI',)
        Dsopra = 1
        Dsotto = 1
    elif l == 2:
        stile = 'livello2 valuta'
        myrange = ('livello2 valuta','Livello-1-scritta', 'Livello-0-scritta', 'Comp TOTALI',)
        Dsopra = 1
        Dsotto = 1
    elif l == 3:
        stile = 'Comp Start Attributo'
        myrange = ('Comp End Attributo', 'Comp TOTALI',)
        Dsopra = 2
        Dsotto = 1

    elif l == 4: #Analisi di Prezzo
        stile = 'An-1_sigla'
        myrange = ('An.1v-Att Start', 'Analisi_Sfondo',)
        Dsopra = 1
        Dsotto = -1
        for n in (3, 5, 7):
            oCellRangeAddr.StartColumn = n
            oCellRangeAddr.EndColumn = n
            oSheet.group(oCellRangeAddr,0)
            oSheet.getCellRangeByPosition(n, 0, n, 0).Columns.IsVisible=False

    test = ultima_voce(oSheet)+2
    lista_cat = list()
    for n in range (0, test):
        if oSheet.getCellByPosition(0, n).CellStyle == stile:
            sopra = n+Dsopra
            for n in range (sopra+1, test):
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
def autoexec (arg=None):
    '''
    questa è richiamata da New_File()
    '''
    #~ chi("autoexec py")
    oDoc = XSCRIPTCONTEXT.getDocument()
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
        DlgMain()
    except:
        #~ chi("autoexec py")
        return
########################################################################
def computo_terra_terra (arg=None):
    '''
    Settaggio base di configuazione colonne in COMPUTO e VARIANTE
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.getCellRangeByPosition(33,0,1023,0).Columns.IsVisible = False
    set_larghezza_colonne()
########################################################################
def viste_nuove (sValori):
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
def set_larghezza_colonne (arg=None):
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
        viste_nuove('TTTFFTTTTTFTFTFTFTFTTFTTFTFTFTTFFFFFF')
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
        oSheet.getColumns().getByName('AX').Columns.Width = 1900
        oSheet.getColumns().getByName('AY').Columns.Width = 1900
        oDoc.CurrentController.freezeAtPosition(0, 3)
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
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
        oDoc.CurrentController.freezeAtPosition(0, 3)
        viste_nuove('TTTFFTTTTTFTFFFFFFTFFFFFFFFFFFFFFFFFFFFFFFFFTT')
    if oSheet.Name == 'Elenco Prezzi':
        oSheet.getColumns().getByName('A').Columns.Width = 1600
        oSheet.getColumns().getByName('B').Columns.Width = 10000
        oSheet.getColumns().getByName('C').Columns.Width = 1500
        oSheet.getColumns().getByName('D').Columns.Width = 2300
        oSheet.getColumns().getByName('E').Columns.Width = 1600
        oSheet.getColumns().getByName('F').Columns.Width = 2300
        oSheet.getColumns().getByName('G').Columns.Width = 2300
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
#~ class adegua_tmpl_th (threading.Thread):
    #~ def __init__(self):
        #~ threading.Thread.__init__(self)
    #~ def run(self):
        #~ adegua_tmpl_run()
#~ def adegua_tmpl (arg=None):
    #~ adegua_tmpl_th().start()
##########
#~ def adegua_tmpl_run (arg=None):
def adegua_tmpl (arg=None):
    '''
    Mantengo la compatibilità con le vecchie versioni del template:
    - dal 200 parte di autoexec è in python
    - dal 203 (LeenO 3.14.0 ha templ 202) introdotta la Super Categoria con nuovi stili di cella;
        sostituita la colonna "Tag A" con "Tag Super Cat"
    '''
    # cambiare stile http://bit.ly/2cDcCJI
    oDoc = XSCRIPTCONTEXT.getDocument()
    ver_tmpl = oDoc.getDocumentProperties().getUserDefinedProperties().Versione
    if ver_tmpl > 200:
        Lib_LeenO('_variabili.autoexec') #rinvia a autoexec in basic
    if ver_tmpl < 203:
        if DlgSiNo("Vuoi procedere con l'adeguamento di questo file alla versione corrente di LeenO?", "Richiesta") ==2:
            #~ oDialogo_attesa = dlg_attesa()
            #~ attesa().start() #mostra il dialogo
#~ adeguo gli stili secondo il template corrente
            sUrl = LeenO_path()+'/template/leeno/Computo_LeenO.ots'
            styles = oDoc.getStyleFamilies()
            styles.loadStylesFromURL(sUrl, list())
            Lib_LeenO('computo.inizializza_computo') #sovrascrive le intestazioni di tabella del computo 
            oSheet = oDoc.getSheets().getByName('S1')
            oSheet.getCellByPosition(7, 290).Value = oDoc.getDocumentProperties().getUserDefinedProperties().Versione = 203
            for el in oDoc.Sheets.ElementNames:
                oDoc.getSheets().getByName(el).IsVisible = True
                oDoc.CurrentController.setActiveSheet(oDoc.getSheets().getByName(el))
                adatta_altezza_riga(el)
                oDoc.getSheets().getByName(el).IsVisible = False
            _gotoSheet ('COMPUTO')
            oDoc.getSheets().getByName('S1').IsVisible = False
            #~ oDialogo_attesa.endExecute() #chiude il dialogo
            #~ oDlgMain.endExecute()
            MsgBox("Adeguamento del file completato con successo.", "Avviso")
        else:
            MsgBox('''Non avendo effettuato l'adeguamento del lavoro alla versione corrente di LeenO, potresti avere dei malfunzionamenti!''', 'Avviso!')
#~ ########################################################################
def r_version_code(arg=None):
    if os.altsep:
        code_file = uno.fileUrlToSystemPath(LeenO_path() + os.altsep + 'leeno_version_code')
    else:
        code_file = uno.fileUrlToSystemPath(LeenO_path() + os.sep + 'leeno_version_code')
    f = open(code_file, 'r')
    return f.readline().split('-')[-1]
########################################################################
def XPWE_export_run (arg=None ):
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
def XPWE_import_run (arg=None ):
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
    Viasualizza il menù principale
    '''
    bak_timestamp() # fa il backup del file
    oDoc = XSCRIPTCONTEXT.getDocument()
    psm = uno.getComponentContext().ServiceManager
    oSheet = oDoc.CurrentController.ActiveSheet
    if oDoc.getSheets().hasByName('S2') == False:
        for bar in GetmyToolBarNames:
            toolbar_on (bar, 0)
        if len(oDoc.getURL())==0 and \
        getLastUsedCell(oSheet).EndColumn ==0 and \
        getLastUsedCell(oSheet).EndRow ==0:
            oDoc.close(True)
        New_file.computo()
    toolbar_vedi()
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDlgMain = dp.createDialog("vnd.sun.star.script:UltimusFree2.DlgMain?language=Basic&location=application")
    oDialog1Model = oDlgMain.Model
    oDlgMain.Title = 'Menù Principale (Ctrl+0)'
    
    sUrl = LeenO_path()+'/icons/Immagine.png'
    oDlgMain.getModel().ImageControl1.ImageURL=sUrl

    if os.altsep:
        code_file = uno.fileUrlToSystemPath(LeenO_path() + os.altsep + 'leeno_version_code')
    else:
        code_file = uno.fileUrlToSystemPath(LeenO_path() + os.sep + 'leeno_version_code')
    f = open(code_file, 'r')
    
    sString = oDlgMain.getControl("Label12")
    sString.Text = f.readline()

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

    oDlgMain.execute()
    return
########################################################################
def InputBox (sCella='', t=''):
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
def hide_error (lErrori, irow):
    '''
    lErrori  { list } : nome dell'errore es.: '#DIV/0!'
    irow { integer } : indice della riga da nascondere
    Viasualizza o nascondi una toolbar
    '''
    #~ attesa().start()
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
            if oSheet.getCellByPosition (irow, i).String == el:
                oCellRangeAddr.StartRow = i
                oCellRangeAddr.EndRow = i
                oSheet.group(oCellRangeAddr,1)
                oSheet.getCellByPosition (0, i).Rows.IsVisible = True
    oDialogo_attesa.endExecute()
    oDoc.CurrentController.ZoomValue = 100
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
def toolbar_vedi (arg=None):
    oDoc = XSCRIPTCONTEXT.getDocument()
    try:
        oLayout = oDoc.CurrentController.getFrame().LayoutManager

        if oDoc.getSheets().getByName('S1').getCellByPosition(7,316).Value == 0:
            for bar in GetmyToolBarNames: #toolbar sempre visibili
                toolbar_on (bar)
        else:
            for bar in GetmyToolBarNames: #toolbar contestualizzate
                toolbar_on (bar, 0)
        #~ oLayout.hideElement("private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV")
        toolbar_ordina()
        oLayout.showElement("private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar")
        nSheet = oDoc.CurrentController.ActiveSheet.Name

        if nSheet == 'Elenco Prezzi':
            toolbar_on ('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ELENCO')
        elif nSheet == 'Analisi di Prezzo':
            toolbar_on ('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ANALISI')
        elif nSheet == 'CONTABILITA':
            toolbar_on ('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CONTABILITA')
        elif nSheet in ('COMPUTO','VARIANTE'):
            toolbar_on ('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_COMPUTO')
    except:
        pass

def toolbar_on (toolbarURL, flag=1):
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
def toolbar_ordina (arg=None):
    #~ https://www.openoffice.org/api/docs/common/ref/com/sun/star/ui/DockingArea.html
    oDoc = XSCRIPTCONTEXT.getDocument()
    oLayout = oDoc.CurrentController.getFrame().LayoutManager
    i = 0
    for bar in GetmyToolBarNames:
        oLayout.dockWindow(bar, 'DOCKINGAREA_TOP', Point(i, 4))
        i += 1
    oLayout.dockWindow('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV', 'DOCKINGAREA_RIGHT', Point(0, 0))
#######################################################################
def make_pack (arg=None, bar=0):
    '''
    bar { integer } : toolbar 0=spenta 1=accesa
    Pacchettizza l'estensione in duplice copia: LeenO.oxt e LeenO-yyyymmddhhmm.oxt
    in una directory precisa (per ora - da parametrizzare)
    '''
    tempo = w_version_code()
    if bar == 0:
        oDoc = XSCRIPTCONTEXT.getDocument()
        for bar in GetmyToolBarNames: #toolbar sempre visibili
            toolbar_on (bar, 0)
        oLayout = oDoc.CurrentController.getFrame().LayoutManager
        oLayout.hideElement("private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV")
    oxt_path = uno.fileUrlToSystemPath(LeenO_path())
    if sys.platform == 'linux' or sys.platform == 'darwin':
        nomeZip2= '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/OXT/LeenO-' + tempo + '.oxt'
        nomeZip = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/OXT/LeenO.oxt'
        os.system('nemo /media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/OXT')

    elif sys.platform == 'win32':
        nomeZip2= 'w:/_dwg/ULTIMUSFREE/_SRC/OXT/LeenO-' + tempo + '.oxt'
        nomeZip = 'w:/_dwg/ULTIMUSFREE/_SRC/OXT/LeenO.oxt'
        os.system('explorer.exe w:\\_dwg\\ULTIMUSFREE\\_SRC\\OXT\\')
    
    shutil.make_archive(nomeZip2, 'zip', oxt_path)
    shutil.move(nomeZip2 + '.zip', nomeZip2)
    shutil.copyfile (nomeZip2, nomeZip)
#######################################################################
def dlg_attesa(arg=None):
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
    return oDialogo_attesa
#~ #
class attesa (threading.Thread):
    #~ http://bit.ly/2fzfsT7
    '''avvia il dialogo di attesa'''
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        oDialog1Model = oDialogo_attesa.Model # oDialogo_attesa è una variabile generale
        oDialog1Model.Title = 'O   p   e   r   a   z   i   o   n   e       i   n       c   o   r   s   o    .   .   .'
        sUrl = LeenO_path()+'/icons/Immagine.png'
        oDialogo_attesa.getModel().ImageControl1.ImageURL=sUrl
        oDialogo_attesa.execute()
        return
########################################################################
class firme_in_calce_th (threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        firme_in_calce_run ()
def firme_in_calce (arg=None):
    firme_in_calce_th().start()
########################################################################
class XPWE_import_th (threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        XPWE_import_run ()
def XPWE_import (arg=None):
    XPWE_import_th().start()
########################################################################
class XPWE_export_th (threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        XPWE_export_run()
def XPWE_export (arg=None):
    XPWE_export_th().start()
########################################################################
class debug_th (threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):

        oDoc = XSCRIPTCONTEXT.getDocument()
        oDoc.enableAutomaticCalculation(False) # blocco il calcolo automatico

        oSheet = oDoc.CurrentController.ActiveSheet
        for i in reversed(range(3, getLastUsedCell(oSheet).EndRow)):
            if oSheet.getCellByPosition(13, i).Value  == 0:
                oSheet.getRows().removeByIndex(i, 1)
        oDoc.enableAutomaticCalculation(True) # riavvio il calcolo automatico
        #~ oDialogo_attesa.endExecute()
def cancella_righe (arg=None):
    '''
    Questa serve a cancellare righe con valori particolari;
    è molto lenta perché Calc ricalcola i valori ad ogni cancellazione.
    Conviene inibire il ricalcolo
    '''
    debug_th().start()
########################################################################
class inserisci_nuova_riga_con_descrizione_th (threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        oDialogo_attesa = dlg_attesa()
        
        oDoc = XSCRIPTCONTEXT.getDocument()
        if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
            return
        
        descrizione = InputBox(t='inserisci una descrizione per la nuova riga')
        attesa().start() #mostra il dialogo
        
        oDoc.CurrentController.ZoomValue = 400
        oSheet = oDoc.CurrentController.ActiveSheet
        i =0
        while (i < getLastUsedCell(oSheet).EndRow):

            if oSheet.getCellByPosition(2, i ).CellStyle == 'comp 1-a':
                sStRange = Circoscrive_Voce_Computo_Att (i)
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
def inserisci_nuova_riga_con_descrizione (arg=None):
    '''
    inserisce, all'inizio di ogni voce di computo o variante,
    una nuova riga con una descrizione a scelta
    '''
    inserisci_nuova_riga_con_descrizione_th().start()
########################################################################
class nascondi_err_th (threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
    def run(self):
        errori = ('#DIV/0!')
        hide_error(errori, 26)
def nascondi_err (arg=None):
    nascondi_err_th().start()
#~ def nascondi_err(arg=None):
    #~ if DlgSiNo("Nascondo eventuali righe a '#DIV/0!' nell'ultima colonna?") == 1:
    #~ errori = ('#DIV/0!', '--')
    #~ hide_error(errori, 26)
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
# ELENCO DEGLI SCRIPT VISUALIZZATI NEL SELETTORE DI MACRO              #
#~ g_exportedScripts = Copia_riga_Ent, doppioni, DlgMain, filtra_codice, Filtra_Computo_A, Filtra_Computo_B, Filtra_Computo_C, Filtra_Computo_Cap, Filtra_Computo_SottCap, Filtra_computo, Ins_Categorie, ins_voce_computo, Inser_Capitolo, Inser_SottoCapitolo, Numera_Voci, Rinumera_TUTTI_Capitoli2, Sincronizza_SottoCap_Tag_Capitolo_Cor, struttura_Analisi, struttura_ComputoM, SubSum, Tutti_Subtotali, Vai_a_M1, XML_import_BOLZANO, XML_import, XPWE_export, XPWE_import, Vai_a_ElencoPrezzi, Vai_a_Computo, Vai_a_Variabili, Vai_a_Scorciatoie, Vai_a_S2, Vai_a_Filtro, Vai_a_SegnaVoci, nuovo_computo, nuovo_listino, nuovo_usobollo, toolbar_vedi, ANALISI_IN_ELENCOPREZZI, Vai_a_S1, autoexec, nascondi_err, azzera_voce, inizializza_analisi, computo_terra_terra,
########################################################################
########################################################################
# ... here is the python script code
# this must be added to every script file (the
# name org.openoffice.script.DummyImplementationForPythonScripts should be changed to something
# different (must be unique within an office installation !)
# --- faked component, dummy to allow registration with unopkg, no functionality expected
#~ import unohelper
# questo mi consente di inserire i comandi python in Accelerators.xcu
# vedi pag.264 di "Manuel du programmeur oBasic"
# <<< vedi in description.xml
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(None, "org.giuseppe-vizziello.leeno", ("org.giuseppe-vizziello.leeno",),)
########################################################################
