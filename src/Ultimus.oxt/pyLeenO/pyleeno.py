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
import os, sys, uno, unohelper, pyuno, logging, shutil
# cos'e' il namespace:
# http://www.html.it/articoli/il-misterioso-mondo-dei-namespaces-1/
from datetime import datetime, date
from com.sun.star.beans import PropertyValue
from xml.etree.ElementTree import ElementTree, Element, SubElement, Comment, tostring
########################################################################
def LeenO_path():
    ctx = XSCRIPTCONTEXT.getComponentContext()
    pir = ctx.getValueByName('/singletons/com.sun.star.deployment.PackageInformationProvider')
    expath=pir.getPackageLocation('org.giuseppe-vizziello.leeno')
    return (expath)
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
        return (document)
    def listino():
        desktop = XSCRIPTCONTEXT.getDesktop()
        opz = PropertyValue()
        opz.Name = 'AsTemplate'
        opz.Value = True
        document = desktop.loadComponentFromURL(LeenO_path()+'/template/leeno/Listino_LeenO.ots', "_blank", 0, (opz,))
        return (document)
import shutil

def oggi():
    '''
    restituisce la data di oggi
    '''
    return ('/'.join(reversed(str(datetime.now()).split(' ')[0].split('-'))))
def debuggfe():
    #~ oDoc = XSCRIPTCONTEXT.getDocument()
    #~ path = oDoc.getURL()
    #~ bak = '.'.join(path.split('.')[:-1]) + '-backup.ods'
    #~ tempo = ''.join(''.join(''.join(str(datetime.now()).split('.')[0].split(' ')).split('-')).split(':'))
    #~ dafg = (str(datetime.now()).split(' ')[0].split('-'))
    return ('/'.join(reversed(str(datetime.now()).split(' ')[0].split('-'))))
    #~ chi(dafg)
    return
    dest = ''.join([path.split('.')[0], '-', tempo, '.ods'])
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.getCellByPosition(1,0).String = bak
    oSheet.getCellByPosition(1,1).String = path
    oSheet.getCellByPosition(1,2).String = dest
    shutil.copyfile (path, dest)
    #~ oSheet.getCellByPosition(1,3).String = path.split('.')[0]
    #~ oSheet.getCellByPosition(1,4).String = tempo
    #~ oSheet.getCellByPosition(1,5).String = path.split('.')[:-1]
    #~ oSheet.getCellByPosition(1,6).String = path.split('.')[:-1]

    
def fdebug():
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
    if style in ('comp Int_colonna', 'Livello-1-scritta', 'livello2 valuta', 'Comp TOTALI',
                'Comp-Bianche sopra', 'comp Art-EP', 'comp Art-EP_R','Comp-Bianche in mezzo', 'comp sotto Bianche'):
        if style in ('comp Int_colonna', 'Livello-1-scritta', 'livello2 valuta'):
            lrow += 1
        elif style in ('Comp-Bianche sopra', 'comp Art-EP','comp Art-EP_R', 'Comp-Bianche in mezzo', 'comp sotto Bianche'):
            sStRange = Circoscrive_Voce_Computo_Att (lrow)
            lrow = sStRange.RangeAddress.EndRow+1
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
        oSheet.getCellRangeByPosition(31, lrow, 32, lrow).CellStyle = 'livello2_'
        oSheet.getCellRangeByPosition(2, lrow, 11, lrow).merge(True)
        oSheet.getCellByPosition(1, lrow).Formula = '=AF' + str(lrow+1) + '''&"."&''' + 'AG' + str(lrow+1)
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
    #~ sTesto = 'prova'
    style = oSheet.getCellByPosition(1, lrow).CellStyle
    if style in ('comp Int_colonna', 'Livello-1-scritta', 'livello2 valuta', 'Comp TOTALI',
                'Comp-Bianche sopra', 'comp Art-EP', 'comp Art-EP_R','Comp-Bianche in mezzo', 'comp sotto Bianche'):
        if style in ('comp Int_colonna', 'Livello-1-scritta', 'livello2 valuta'):
            lrow += 1
        #~ elif style in ('Comp TOTALI'):
            #~ lrow -= 1
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
        oSheet.getCellRangeByPosition(0, lrow, 41, lrow).CellStyle = 'Livello-1-scritta'
        oSheet.getCellRangeByPosition(2, lrow, 17, lrow).CellStyle = 'Livello-1-scritta mini'
        oSheet.getCellRangeByPosition(18, lrow, 18, lrow).CellStyle = 'Livello-1-scritta mini val'
        oSheet.getCellRangeByPosition(24, lrow, 24, lrow).CellStyle = 'Livello-1-scritta mini %'
        oSheet.getCellRangeByPosition(29, lrow, 29, lrow).CellStyle = 'Livello-1-scritta mini %'
        oSheet.getCellRangeByPosition(30, lrow, 30, lrow).CellStyle = 'Livello-1-scritta mini val'
        oSheet.getCellRangeByPosition(2, lrow, 11, lrow).merge(True)
        oSheet.getCellByPosition(1, lrow).Formula = '=AF' + str(lrow+1)
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
    #~ oSheet.getCellRangeByPosition(2, lrow-1, 11, lrow).Rows.OptimalHeight = True
    #~ SubSum_Cap (lrow)
########################################################################
def Rinumera_TUTTI_Capitoli2():
    Tutti_Subtotali()# ricalcola i totali di categorie e subcategorie
    Sincronizza_SottoCap_Tag_Capitolo_Cor()# sistemo gli idcat voce per voce
    
def Tutti_Subtotali():
    '''ricalcola i subtotali di categorie e subcategorie'''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
        return
    for n in range (0, ultima_voce(oSheet)+1):
        if oSheet.getCellByPosition(0, n).CellStyle == 'Livello-1-scritta':
            SubSum_Cap (n)
        if oSheet.getCellByPosition(0, n).CellStyle == 'livello2 valuta':
            SubSum_SottoCap (n)
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
    lrowE = ultima_voce(oSheet)+1
    nextCap = lrowE
    for n in range (lrow+1, lrowE):
        #~ if oSheet.getCellByPosition(18, n).CellStyle in ('livello2 scritta mini', 'Livello-1-scritta mini val', 'Comp TOTALI'):
        if oSheet.getCellByPosition(18, n).CellStyle in ('livello2 scritta mini', 'Livello-1-scritta mini val', 'Comp TOTALI'):
            #~ MsgBox(oSheet.getCellByPosition(18, n).CellStyle,'')
            nextCap = n + 1
            break
    #~ MsgBox(str(nextCap),'')
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
    lrowE = ultima_voce(oSheet)+1
    nextCap = lrowE
    for n in range (lrow+1, lrowE):
        if oSheet.getCellByPosition(18, n).CellStyle in ('Livello-1-scritta mini val', 'Comp TOTALI'):
            #~ MsgBox(oSheet.getCellByPosition(18, n).CellStyle,'')
            nextCap = n + 1
            break
    oSheet.getCellByPosition(18, lrow).Formula = '=SUBTOTAL(9;S' + str(lrow + 1) + ':S' + str(nextCap) + ')'
    oSheet.getCellByPosition(18, lrow).CellStyle = 'Livello-1-scritta mini val'
    oSheet.getCellByPosition(24, lrow).Formula = '=S' + str(lrow + 1) + '/S' + str(lrowE+1)
    oSheet.getCellByPosition(24, lrow).CellStyle = 'Livello-1-scritta mini %'
    oSheet.getCellByPosition(29, lrow).Formula = '=AE' + str(lrow + 1) + '/S' + str(lrowE+1)
    oSheet.getCellByPosition(29, lrow).CellStyle = 'Livello-1-scritta mini %'
    oSheet.getCellByPosition(30, lrow).Formula = '=SUBTOTAL(9;AE' + str(lrow + 1) + ':AE' + str(nextCap) + ')'
    oSheet.getCellByPosition(30, lrow).CellStyle = 'Livello-1-scritta mini val'
########################################################################
def Sincronizza_SottoCap_Tag_Capitolo_Cor():
    '''
    lrow    { double } : id della riga di inerimento
    sincronizza il categoria e sottocategorie
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oDoc.getSheets().getByName('S1').getCellByPosition(7,304).Value == 0: #se 1 aggiorna gli indici
        return
    if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
        return
#    lrow = Range2Cell()[1]
    lastRow = ultima_voce(oSheet)+1
    
    lista = list()
    for lrow in range(0,lastRow):
        if oSheet.getCellByPosition(2, lrow).CellStyle == 'Livello-1-scritta mini':
            if oSheet.getCellByPosition(2, lrow).String not in lista:
                lista.append((oSheet.getCellByPosition(2, lrow).String))
            idcat = lista.index(oSheet.getCellByPosition(2, lrow).String) +1
            oSheet.getCellByPosition(31, lrow).Value = idcat
        if oSheet.getCellByPosition(31, lrow).CellStyle in ('compTagRiservato',
                                                            'livello2_'):
            try:
                oSheet.getCellByPosition(31, lrow).Value = idcat
            except:
                oSheet.getCellByPosition(31, lrow).Value = 0

    #~ MsgBox(str(lista),'')

    lista = list()
    for lrow in range(0,lastRow):
        if oSheet.getCellByPosition(2, lrow).CellStyle == 'livello2_':
            if oSheet.getCellByPosition(2, lrow).String not in lista:
                lista.append((oSheet.getCellByPosition(2, lrow).String))
            idsbcat = lista.index(oSheet.getCellByPosition(2, lrow).String) +1
            oSheet.getCellByPosition(32, lrow).Value = idsbcat
        if oSheet.getCellByPosition(31, lrow).CellStyle in ('compTagRiservato', 'livello2_'):
            try:
                oSheet.getCellByPosition(32, lrow).Value = idsbcat
            except:
                oSheet.getCellByPosition(32, lrow).Value = 0
        elif oSheet.getCellByPosition(31, lrow).CellStyle in ('Livello-1-scritta'):
            #~ oSheet.getCellByPosition(32, lrow).Value = 0
            idsbcat = 0

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
    #~ MsgBox(nRow,'')
    for n in reversed(range(0, nRow)):
        if oSheet.getCellByPosition(0, n).CellStyle in ('EP-aS', 'An-sfondo-basso Att End', 'Comp End Attributo',
                                                        'Comp End Attributo_R', 'comp Int_colonna', 'comp Int_colonna_R_prima',
                                                        'Livello-1-scritta', 'livello2 valuta'):
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
from com.sun.star.beans import PropertyValue
def _gotoCella (IDcol, IDrow):
    '''
    IDcol   { integer } : id colonna
    IDrow   { integer } : id riga

    muove il cursore nelle cella (IDcol, IDrow)
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    oProp = PropertyValue()
    oProp.Name = 'ToPoint'
    oProp.Value = ColumnNumberToName(oSheet, IDcol)+str(IDrow+1)
    properties = (oProp,)
    dispatchHelper.executeDispatch(oFrame, '.uno:GoToCell', '', 0, properties )
########################################################################
def Adatta_Altezza_riga ():
    '''questa sembra inefficace
    meglio usare qualcosa tipo:
    oSheet.getCellRangeByPosition(2, lrow-1, 11, lrow).Rows.OptimalHeight = True'''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ctx = XSCRIPTCONTEXT.getComponentContext()
    desktop = XSCRIPTCONTEXT.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.DispatchHelper', ctx )
    oProp = PropertyValue()
    oProp.Name = 'aExtraHeight'
    oProp.Value = 0
    properties = (oProp,)
    dispatchHelper.executeDispatch(oFrame, '.uno:SetOptimalRowHeight', '', 0, properties)
    if oSheet.Name in ('Elenco Prezzi', 'VARIANTE', 'COMPUTO', 'CONTABILITA'):
        oSheet.getCellByPosition(0, 2).Rows.Height = 800
########################################################################
# doppioni #############################################################
def doppioni():
    '''
    Elimina i doppioni nell'elenco prezzi
    basandosi solo sul confronto dei codici di prezzo
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    ###
    #~ oDoc.addActionLock
    #~ oDoc.lockControllers
    ###
    lista_voci = list()
    diz_ep = dict()
    for n in range (0, ultima_voce(oSheet)+1):
        if oSheet.getCellByPosition(0, n).CellStyle == 'EP-aS':
            cod = oSheet.getCellByPosition(0, n).String
            des = oSheet.getCellByPosition(1, n).String
            um = oSheet.getCellByPosition(2, n).String
            sic = oSheet.getCellByPosition(3, n).Value
            pr = oSheet.getCellByPosition(4, n).Value
            iMDO = oSheet.getCellByPosition(5, n).Value
            mdo = oSheet.getCellByPosition(6, n).Value
            cor = oSheet.getCellByPosition(7, n).String
            voce = (cod, des, um, sic, pr, iMDO, mdo, cor)
            diz_ep[cod] = voce
    for voce in diz_ep.items():
        lista_voci.append(voce[-1])
    oSheet.getRows().removeByIndex(3, ultima_voce(oSheet)-2)
    lista_voci.sort()
    lista_come_array = tuple(lista_voci) 

    # Parametrizzo il range di celle a seconda della dimensione della lista
    colonne_lista = len(lista_come_array[1]) # numero di colonne necessarie per ospitare i dati
    righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati

    oSheet.getRows().insertByIndex(3, righe_lista)

    oRange = oSheet.getCellRangeByPosition( 0, 
                                            3, 
                                            colonne_lista - 1, # l'indice parte da 0
                                            righe_lista + 3 - 1)
    oRange.setDataArray(lista_come_array)

#~ SISTEMO GLI STILI
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = oSheet.RangeAddress.Sheet # recupero l'index del foglio
    SC = oCellRangeAddr.StartColumn = 0
    SR = oCellRangeAddr.StartRow = 3
    EC = oCellRangeAddr.EndColumn = 0
    ER = oCellRangeAddr.EndRow = 3 + righe_lista - 1
    
    oSheet.getCellRangeByPosition (0, SR, 7, ER).CellStyle = 'EP-aS'
    oSheet.getCellRangeByPosition (1, SR, 1, ER).CellStyle = 'EP-a'
    oSheet.getCellRangeByPosition (2, SR, 6, ER).CellStyle = 'EP-mezzo'
    oSheet.getCellRangeByPosition (5, SR, 5, ER).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition (8, SR, 9, ER).CellStyle = 'EP-sfondo'
    oSheet.getCellRangeByPosition (10, SR, 10, ER).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition (11, SR, 11, ER).CellStyle = 'EP statistiche'
    oSheet.getCellRangeByPosition (13, SR, 13, ER).CellStyle = 'EP statistiche_Contab_q'
    oSheet.getCellRangeByPosition (14, SR, 14, ER).CellStyle = 'EP statistiche_Contab'
    ###
    #~ oDoc.removeActionLock
    #~ oDoc.unlockControllers
    ###
    
# doppioni #############################################################
########################################################################
# Scrive un file.
def XPWE_export():
    '''
    esporta il documento in formato XPWE
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    if oDoc.getSheets().hasByName('S2') == False:
        MsgBox('Puoi usare questo comando da un file di computo esistente.','Avviso!')
        return
    lista_righe = list()
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
    oSheet = oDoc.getSheets().getByName('COMPUTO')
    lastRow = ultima_voce(oSheet)+1
    # evito di esportare in SuperCategorie perché inutile, almeno per ora
    #~ for n in range (0, ultima_voce(oSheet)):
        #~ if oSheet.getCellByPosition(1, n).CellStyle == 'Livello-1-scritta':
            #~ idID = oSheet.getCellByPosition(1, n).String
            #~ desc = oSheet.getCellByPosition(2, n).String
            
            #~ PweDGSuperCategorie = SubElement(PweDGCapitoliCategorie,'PweDGSuperCategorie')
            #~ DGSuperCategorieItem = SubElement(PweDGSuperCategorie,'DGSuperCategorieItem')
            #~ DesSintetica = SubElement(DGSuperCategorieItem,'DesSintetica')
            
            #~ DGSuperCategorieItem.set('ID', idID)
            #~ DesSintetica.text = desc

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

#~ Elenco Prezzi
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    PweElencoPrezzi = SubElement(PweMisurazioni,'PweElencoPrezzi')
    diz_ep = dict ()
    for n in range (0, getLastUsedCell(oSheet).EndRow):
        if oSheet.getCellByPosition(0, n).CellStyle == 'EP-aS':
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
            DesRidotta = SubElement(EPItem,'DesRidotta')
            DesRidotta.text = ''
            DesEstesa = SubElement(EPItem,'DesEstesa')
            DesEstesa.text = oSheet.getCellByPosition(1, n).String
            DesBreve = SubElement(EPItem,'DesBreve')
            DesBreve.text = ''
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
            Flags.text = '0'
            Data = SubElement(EPItem,'Data')
            Data.text = '30/12/1899'
            AdrInternet = SubElement(EPItem,'AdrInternet')
            AdrInternet.text = ''
            PweEPAnalisi = SubElement(EPItem,'PweEPAnalisi')
            PweEPAnalisi.text = ''
    #~ COMPUTO
    oSheet = oDoc.getSheets().getByName('COMPUTO')
    PweVociComputo = SubElement(PweMisurazioni,'PweVociComputo')
    oDoc.CurrentController.select(oSheet)
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
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
            IDSpCat.text = ''
##########################
            IDCat = SubElement(VCItem,'IDCat')
            IDCat.text = oSheet.getCellByPosition(31, sotto).String
##########################
            IDSbCat = SubElement(VCItem,'IDSbCat')
            IDSbCat.text = oSheet.getCellByPosition(32, sotto).String
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
##########################
                Lunghezza = SubElement(RGItem,'Lunghezza')
                if oSheet.getCellByPosition(6, m).Formula.split('=')[-1] == None:
                    Lunghezza.text = oSheet.getCellByPosition(6, m).String
                else:
                    Lunghezza.text = str(oSheet.getCellByPosition(6, m).Formula.split('=')[-1])
                try:
                    int(oSheet.getCellByPosition(6, m).Formula[1])
                except:
                    Lunghezza.text = oSheet.getCellByPosition(6, m).String
##########################
                Larghezza = SubElement(RGItem,'Larghezza')
                if oSheet.getCellByPosition(7, m).Formula.split('=')[-1] == None:
                    Larghezza.text = oSheet.getCellByPosition(7, m).String
                else:
                    Larghezza.text = str(oSheet.getCellByPosition(7, m).Formula.split('=')[-1])
                try:
                    int(oSheet.getCellByPosition(7, m).Formula[1])
                except:
                    Larghezza.text = oSheet.getCellByPosition(7, m).String
##########################
                HPeso = SubElement(RGItem,'HPeso')
                if oSheet.getCellByPosition(8, m).Formula.split('=')[-1] == None:
                    HPeso.text = oSheet.getCellByPosition(8, m).Formula
                else:
                    HPeso.text = str(oSheet.getCellByPosition(8, m).Formula.split('=')[-1])
                try:
                    int(oSheet.getCellByPosition(8, m).Formula[1])
                except:
                    HPeso.text = oSheet.getCellByPosition(8, m).Formula
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
                    if '-' in Quantita:
                        Flags.text = '32769'
            n = sotto+1
##########################
    out_file = filedia('Salva con nome...', '*.xpwe', 1)

    if out_file.split('.')[-1].upper() != 'XPWE':
        out_file = out_file + '.xpwe'

    riga = str(tostring(top, encoding="unicode"))
    of = codecs.open(out_file,'w','utf-8')
    of.write(riga)
    MsgBox('Esportazione in formato XPWE eseguita con successo\nsul file ' + out_file + '!','Avviso.')
########################################################################
def debug___():
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    n = next_voice(lrow, 1)
    MsgBox(n)
    #~ _gotoCella(0, n)
    
def next_voice (lrow, n=1):
    '''
    lrow { double }   : riga di riferimento
    n    { integer }  : se 0 inserisce prima della voce corrente
                        se 1 inserisce dopo della voce corrente
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
    if oSheet.Name in ('VARIANTE', 'COMPUTO','CONTABILITA'):
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
def copia_riga_computo(lrow):
    '''
    Inserisce una nuova riga di misurazione nel computo
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ lrow = Range2Cell()[1]
    stile = oSheet.getCellByPosition(2, lrow).CellStyle
    if stile in ('Comp-Bianche in mezzo Descr', 'comp 1-a', 'comp sotto centro'):# <stili computo
        sStRange = Circoscrive_Voce_Computo_Att (lrow)
        sStRange.RangeAddress
        sopra = sStRange.RangeAddress.StartRow
        sotto = sStRange.RangeAddress.EndRow
        if stile == 'Comp-Bianche in mezzo Descr' or stile == 'comp 1-a':
            lrow = lrow+1 # PER INSERIMENTO SOTTO RIGA CORRENTE
        if stile == 'comp sotto centro':
            pass
        oSheet.getRows().insertByIndex(lrow,1)
# immissione tags cat/subcat
        oSheet.getCellByPosition(31, lrow).Formula = '=AF$' +str(sotto+2)
        oSheet.getCellByPosition(32, lrow).Formula = '=AG$' +str(sotto+2)
        oSheet.getCellByPosition(33, lrow).Formula = '=AH$' +str(sotto+2)
        oSheet.getCellByPosition(34, lrow).Formula = '=AI$' +str(sotto+2)
        oSheet.getCellByPosition(35, lrow).Formula = '=AJ$' +str(sotto+2)
# imposto gli stili
        oSheet.getCellRangeByPosition(5, lrow, 7, lrow,).CellStyle = 'comp 1-a'
        oSheet.getCellByPosition(0, lrow).CellStyle = 'comp 10 s'
        oSheet.getCellByPosition(1, lrow).CellStyle = 'Comp-Bianche in mezzo'
        oSheet.getCellByPosition(2, lrow).CellStyle = 'comp 1-a'
        oSheet.getCellRangeByPosition(3, lrow, 4, lrow).CellStyle = 'Comp-Bianche in mezzo bordate_R'
        oSheet.getCellByPosition(8, lrow).CellStyle = 'comp 1-a peso'
        oSheet.getCellByPosition(9, lrow).CellStyle = 'Blu'
# ci metto le formule
        oSheet.getCellByPosition(9, lrow).Formula = '=IF(PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + ')=0;'';PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + '))'
        oSheet.getCellByPosition(10 , lrow).Formula = ''
        oDoc.CurrentController.select(oSheet.getCellByPosition(2, lrow))
def copia_riga_contab(lrow):
    '''
    Inserisce una nuova riga di misurazione in contabilità
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ lrow = Range2Cell()[1]
    stile = oSheet.getCellByPosition(1, lrow).CellStyle
    if stile in ('comp Art-EP_R', 'Data_bianca', 'Comp-Bianche in mezzo_R'):
        sStRange = Circoscrive_Voce_Computo_Att (lrow)
        sStRange.RangeAddress
        sopra = sStRange.RangeAddress.StartRow
        sotto = sStRange.RangeAddress.EndRow
        lrow = lrow+1 # PER INSERIMENTO SOTTO RIGA CORRENTE
        if  oSheet.getCellByPosition(2, lrow).CellStyle == 'comp sotto centro_R':
            lrow = lrow-1
        oSheet.getRows().insertByIndex(lrow,1)
    # immissione tags cat/subcat
        oSheet.getCellByPosition(31, lrow).Formula = '=AF$' +str(sotto+2)
        oSheet.getCellByPosition(32, lrow).Formula = '=AG$' +str(sotto+2)
        oSheet.getCellByPosition(33, lrow).Formula = '=AH$' +str(sotto+2)
        oSheet.getCellByPosition(34, lrow).Formula = '=AI$' +str(sotto+2)
        oSheet.getCellByPosition(35, lrow).Formula = '=AJ$' +str(sotto+2)
    # imposto gli stili
        oSheet.getCellByPosition(1, lrow).CellStyle = 'Comp-Bianche in mezzo_R'
        oSheet.getCellByPosition(2, lrow).CellStyle = 'comp 1-a'
        oSheet.getCellRangeByPosition(5, lrow, 7, lrow).CellStyle = 'comp 1-a'
        oSheet.getCellRangeByPosition(11, lrow, 23, lrow).CellStyle = 'Comp-Bianche in mezzo_R'
        oSheet.getCellByPosition(8, lrow).CellStyle = 'comp 1-a peso'
        oSheet.getCellRangeByPosition(9, lrow, 11, lrow).CellStyle = 'Comp-Variante'
    # ci metto le formule
        oSheet.getCellByPosition(9, lrow).Formula = '=IF(PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + ')<=0;'';PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + '))'
        oSheet.getCellByPosition(11, lrow).Formula = '=IF(PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + ')>=0;'';PRODUCT(F' + str(lrow+1) + ':I' + str(lrow+1) + ')*-1)'
    # preserva la data di misura
        if oSheet.getCellByPosition(1, lrow+1).CellStyle == 'Data_bianca':
            oRangeAddress = oSheet.getCellByPosition(1, lrow+1).getRangeAddress()
            oCellAddress = oSheet.getCellByPosition(1,lrow).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
            oSheet.getCellByPosition(1, lrow+1).String = ''
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
        oSheet.getCellByPosition(1, lrow).Formula = '=IF(A' + str(lrow+1) + '='';'';CONCATENATE('  ';VLOOKUP(A' + str(lrow+1) + ';elenco_prezzi;2;FALSE());' '))'
        oSheet.getCellByPosition(2, lrow).Formula = '=IF(A' + str(lrow+1) + '='';'';VLOOKUP(A' + str(lrow+1) + ';elenco_prezzi;3;FALSE()))'
        oSheet.getCellByPosition(4, lrow).Formula = '=IF(A' + str(lrow+1) + '='';0;VLOOKUP(A' + str(lrow+1) + ';elenco_prezzi;5;FALSE()))'
        oSheet.getCellByPosition(5, lrow).Formula = '=D' + str(lrow+1) + '*E' + str(lrow+1)
        oSheet.getCellByPosition(8, lrow).Formula = '=IF(A' + str(lrow+1) + '='';'';IF(VLOOKUP(A' + str(lrow+1) + ';elenco_prezzi;6;FALSE())='';'';(VLOOKUP(A' + str(lrow+1) + ';elenco_prezzi;6;FALSE()))))'
        oSheet.getCellByPosition(9, lrow).Formula = '=IF(I' + str(lrow+1) + '='';'';I' + str(lrow+1) + '*F' + str(lrow+1) + ')'
    # preserva il Pesca
        if oSheet.getCellByPosition(1, lrow-1).CellStyle == 'An-lavoraz-dx-senza-bordi':
            oRangeAddress = oSheet.getCellByPosition(0, lrow+1).getRangeAddress()
            oCellAddress = oSheet.getCellByPosition(0,lrow).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
        oSheet.getCellByPosition(0, lrow).String = 'Cod. Art.?'
    oDoc.CurrentController.select(oSheet.getCellByPosition(1, lrow))
def Copia_riga_Ent(): #Aggiungi Componente - capisce su quale tipologia di tabelle è
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    nome_sheet = oSheet.Name
    if nome_sheet == 'COMPUTO':
        copia_riga_computo(lrow)
    elif nome_sheet == 'CONTABILITA':
        copia_riga_contab(lrow)
    elif nome_sheet == 'Analisi di Prezzo':
        copia_riga_analisi(lrow)
########################################################################
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
    oSheet.getRows().insertByIndex(lrow,4)#~ insRows(lrow,4) #inserisco le righe
    oSheet.copyRange(oCellAddress, oRangeAddress)
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
# sistemo i LINK dei tagG nelle righe sopra al tag vero e prorio...
    oSheet.getCellByPosition(31, lrow+2).Formula = '=AF$'+str(lrow+4)
    oSheet.getCellByPosition(32, lrow+2).Formula = '=AG$'+str(lrow+4)
    oSheet.getCellByPosition(33, lrow+2).Formula = '=AH$'+str(lrow+4)
    oSheet.getCellByPosition(34, lrow+2).Formula = '=AI$'+str(lrow+4)
    oSheet.getCellByPosition(35, lrow+2).Formula = '=AJ$'+str(lrow+4)
    oSheet.getCellByPosition(31, lrow+1).Formula = '=AF$'+str(lrow+4)
    oSheet.getCellByPosition(32, lrow+1).Formula = '=AG$'+str(lrow+4)
    oSheet.getCellByPosition(33, lrow+1).Formula = '=AH$'+str(lrow+4)
    oSheet.getCellByPosition(34, lrow+1).Formula = '=AI$'+str(lrow+4)
    oSheet.getCellByPosition(35, lrow+1).Formula = '=AJ$'+str(lrow+4)
    oSheet.getCellByPosition(31, lrow).Formula = '=AF$'+str(lrow+4)
    oSheet.getCellByPosition(32, lrow).Formula = '=AG$'+str(lrow+4)
    oSheet.getCellByPosition(33, lrow).Formula = '=AH$'+str(lrow+4)
    oSheet.getCellByPosition(34, lrow).Formula = '=AI$'+str(lrow+4)
    oSheet.getCellByPosition(35, lrow).Formula = '=AJ$'+str(lrow+4)
    if oSheet.getCellByPosition(31, lrow-1).CellStyle in ('livello2 valuta', 'Livello-1-scritta', 'compTagRiservato'):
        oSheet.getCellByPosition(31, lrow+3).Value = oSheet.getCellByPosition(31, lrow-1).Value
        oSheet.getCellByPosition(32, lrow+3).Value = oSheet.getCellByPosition(32, lrow-1).Value
    #~ celle=oSheet.getCellRangeByPosition(0, lrow, 43,lrow+3)# 'seleziona la cella
    #~ oDoc.CurrentController.select(celle)
    #~ celle.Rows.OptimalHeight = True
########################################################################
    _gotoCella(1,lrow+1)
########################################################################
# ins_voce_computo #####################################################
def ins_voce_computo(): #TROPPO LENTA
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
########################################################################
# XML_import ###########################################################
def XML_import ():
    New_file.listino()
    '''
    Routine di importazione di un prezziario XML formato SIX. Molto
    liberamente tratta da PreventARES https://launchpad.net/preventares
    di <Davide Vescovini> <davide.vescovini@gmail.com>
    '''
    filename = filedia('Scegli il file XML-SIX da importare', '*.xml')
    datarif = datetime.now()
    # inizializzazioe delle variabili
    lista_articoli = list() # lista in cui memorizzare gli articoli da importare
    diz_um = dict() # array per le unità di misura
    # stringhe per descrizioni articoli
    desc_breve = str()
    desc_estesa = str()
    # effettua il parsing del file XML
    tree = ElementTree()
    if filename == 'Cancel' or filename == '':
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
def XML_import_BOLZANO ():
    New_file.listino()
    '''
    Routine di importazione di un prezziario XML formato SIX. Molto
    liberamente tratta da PreventARES https://launchpad.net/preventares
    di <Davide Vescovini> <davide.vescovini@gmail.com>
    *Versione bilingue*
    '''
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
    if filename == 'Cancel' or filename == '':
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
            oSheet.getCellRangeByPosition (2, lrow, 7, lrow).CellStyle = 'comp sotto centro'
            oSheet.getCellByPosition (8, lrow).CellStyle = 'comp sotto BiancheS'
            oSheet.getCellByPosition (9, lrow).CellStyle = 'Comp-Variante num sotto'

            oSheet.getCellByPosition(31, lrow).Formula ='=AF$' + str(sotto+2)
            oSheet.getCellByPosition(32, lrow).Formula ='=AG$' + str(sotto+2)
            oSheet.getCellByPosition(33, lrow).Formula ='=AH$' + str(sotto+2)
            oSheet.getCellByPosition(34, lrow).Formula ='=AI$' + str(sotto+2)
            oSheet.getCellByPosition(35, lrow).Formula ='=AJ$' + str(sotto+2)
            
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
def vedi_voce(riga_corrente,vRif):
    """(riga d'inserimento, riga di riferimento)"""
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ riga_corrente = Range2Cell()[1]

    #~ if oSheet.getCellByPosition(2, riga_corrente).CellStyle != 'comp 1-a':
        #~ MsgBox(oSheet.getCellByPosition(2, riga_corrente).CellStyle,str(type(riga_corrente)))
        #~ return

    #~ copia_riga_computo(riga_corrente+1)
    #~ vRif = 6
    sStRange = Circoscrive_Voce_Computo_Att (vRif)
    sStRange.RangeAddress
    idv = sStRange.RangeAddress.StartRow +1
    sotto = sStRange.RangeAddress.EndRow
    art = abs2name (1, idv)
    idvoce = abs2name (0, idv)
    quantity = abs2name (9, sotto)
    um = 'VLOOKUP(' + art + ';elenco_prezzi;3;FALSE())'
    #~ MsgBox(str(um),'um')
    oSheet.getCellByPosition(2, riga_corrente).Formula='=CONCATENATE("";" - vedi voce n. ";TEXT(' + idvoce +';"@");" - art. ";' + art + ';"[";' + um + ';"]"'
    oSheet.getCellByPosition(5, riga_corrente).Formula='=' + quantity
########################################################################
# XPWE_import ##########################################################
def XPWE_import(): #(filename):
    oDoc = XSCRIPTCONTEXT.getDocument()
    ###
    oDoc.addActionLock
    oDoc.lockControllers
    ###
    if oDoc.getSheets().hasByName('S2') == False:
        MsgBox('Puoi usare questo comando da un file di computo nuovo o già esistente.','Avviso!')
        return
    #~ filename = filedia('Scegli il file XPWE da importare...')
    filename = filedia('Scegli il file XPWE da importare...','*.xpwe')
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
    except PermissionError:
        MsgBox('Accertati che il nome del file sia corretto.', 'ATTENZIONE! Impossibile procedere.')
        return
    # ottieni l'item root
    root = tree.getroot()
    logging.debug(list(root))
    # effettua il parsing di tutti gli elemnti dell'albero XML
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
        adrinternet = elem.find('AdrInternet').text
        pweepanalisi = elem.find('PweEPAnalisi').text
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
        diz_ep['pweepanalisi'] = pweepanalisi
        dict_articoli[id_ep] = diz_ep
        lista_articoli.append
        articolo_modificato =  (tariffa,
                                    destestesa,
                                    unmisura,
                                    '',
                                    float(prezzo1))
        lista_articoli.append(articolo_modificato)
###
# leggo voci di misurazione e righe ####################################
    lista_misure = list()
    try:
        PweVociComputo = misurazioni.getchildren()[1]
        vcitems = PweVociComputo.findall('VCItem')
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
            #~ MsgBox(idspcat,'')
            for el in righi_mis:
                diz_rig = dict()
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
                diz_rig['rgitem'] = rgitem
                diz_rig['idvv'] = idvv
                diz_rig['descrizione'] = descrizione

                if partiuguali !=None:
                    diz_rig['partiuguali'] = partiuguali.replace('.',',')
                else:
                    diz_rig['partiuguali'] = partiuguali

                if lunghezza !=None:
                    diz_rig['lunghezza'] = lunghezza.replace('.',',')
                else:
                    diz_rig['lunghezza'] = lunghezza

                if larghezza !=None:
                    diz_rig['larghezza'] = larghezza.replace('.',',')
                else:
                    diz_rig['larghezza'] = larghezza

                if hpeso !=None:
                    diz_rig['hpeso'] = hpeso.replace('.',',')
                else:
                    diz_rig['hpeso'] = hpeso

                if quantita !=None:
                    diz_rig['quantita'] = quantita.replace('.',',')
                else:
                    diz_rig['quantita'] = quantita

                diz_rig['flags'] = flags
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
            lista_misure.append(diz_misura)
    except IndexError:
        MsgBox("""Nel file scelto non risultano esserci voci di misurazione,
perciò saranno importate le sole voci di Elenco Prezzi.

Si tenga conto che:
    - sarà importato solo il "Prezzo 1" dell'elenco;
    - il formato XPWE non conserva alcuni dati come
      le incidenze di sicurezza e di manodopera!""",'ATTENZIONE!')
        pass
    #~ articoli = open ('/home/giuserpe/.config/libreoffice/4/user/uno_packages/cache/uno_packages/luds59ep.tmp_/LeenO-3.11.3.dev-150714180321.oxt/pyLeenO/articoli.txt', 'w')
    #~ print (str(lista_articoli), file=articoli)
    #~ articoli.close()
    #~ misure = open ('/home/giuserpe/.config/libreoffice/4/user/uno_packages/cache/uno_packages/luds59ep.tmp_/LeenO-3.11.3.dev-150714180321.oxt/pyLeenO/misure.txt', 'w')
    #~ print (str(lista_misure), file=misure)
    #~ misure.close()
    #~ MsgBox('ho stampato', '')
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

    oSheet.getRows().insertByIndex(3, righe_lista)

    oRange = oSheet.getCellRangeByPosition( scarto_colonne, 
                                            scarto_righe, 
                                            colonne_lista + scarto_colonne - 1, # l'indice parte da 0
                                            righe_lista + scarto_righe - 1)

    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = oSheet.RangeAddress.Sheet # recupero l'index del foglio

    SC = oCellRangeAddr.StartColumn = 0
    SR = oCellRangeAddr.StartRow = 3
    EC = oCellRangeAddr.EndColumn = 0
    ER = oCellRangeAddr.EndRow = 3 + righe_lista - 1

    oRange.setDataArray(lista_come_array)
#~ SISTEMO GLI STILI
    oSheet.getCellRangeByPosition (0, SR, 7, ER).CellStyle = 'EP-aS'
    oSheet.getCellRangeByPosition (1, SR, 1, ER).CellStyle = 'EP-a'
    oSheet.getCellRangeByPosition (2, SR, 6, ER).CellStyle = 'EP-mezzo'
    oSheet.getCellRangeByPosition (5, SR, 5, ER).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition (8, SR, 9, ER).CellStyle = 'EP-sfondo'
    oSheet.getCellRangeByPosition (10, SR, 10, ER).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition (11, SR, 11, ER).CellStyle = 'EP statistiche'
    oSheet.getCellRangeByPosition (13, SR, 13, ER).CellStyle = 'EP statistiche_Contab_q'
    oSheet.getCellRangeByPosition (14, SR, 14, ER).CellStyle = 'EP statistiche_Contab'
    #~ return
###
    if len(lista_misure) == 0:
        MsgBox("Importate n."+ str(len(lista_articoli)) +" voci dall'elenco prezzi\ndel file: " + filename, 'Avviso')
        oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
        oDoc.CurrentController.select(oSheet)
        oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
        oDoc.CurrentController.ZoomValue = 100
        return
###
# Inserisco i dati nel COMPUTO #########################################
    oSheet = oDoc.getSheets().getByName('COMPUTO')
    oDoc.CurrentController.select(oSheet)
    oDoc.CurrentController.ZoomValue = 400
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
        if idcat != testcat:
            testcat = idcat
            #~ MsgBox(lista_cat)
            Inser_Capitolo_arg(lrow, lista_cat[eval(idcat)-1][1])
            lrow = lrow + 2
            Inser_SottoCapitolo_arg(lrow, lista_subcat[eval(idsbcat)-1][1])
            #~ if idsbcat != testsbcat:
                #~ testsbcat = idsbcat
                #~ Inser_SottoCapitolo_arg(lrow, lista_subcat[eval(idsbcat)-1][1])

        if idsbcat != testsbcat:
            testsbcat = idsbcat
            Inser_SottoCapitolo_arg(lrow, lista_subcat[eval(idsbcat)-1][1])

        lrow = ultima_voce(oSheet) + 1
        ins_voce_computo_grezza(lrow)
        ID = el.get('id_ep')
        id_vc = el.get('id_vc')
        oSheet.getCellByPosition(1, lrow+1).String = dict_articoli.get(ID).get('tariffa')
        diz_vv[id_vc] = lrow+1
        oSheet.getCellByPosition(0, lrow+1).String = str(x)
        x = x+1
        SC = 2
        SR = lrow + 2 + 1
        nrighe = len(el.get('lista_rig')) - 1
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
# va bene se lista_righe viene convertito come tupla alla riga 1363
        SR = SR - 1
        #~ MsgBox(str(eval('1,1')),str(eval('1,1')))

        for mis in el.get('lista_rig'):
            #~ MsgBox(str(mis[9]),'idvv')

            if mis[0] != None: #descrizione
                descrizione = mis[0].strip()
                oSheet.getCellByPosition(2, SR).String = descrizione
            else:
                descrizione =''

            if mis[4] != None: #lunghezza
                if any(o in mis[4] for o in ('+', '*', '/', '-',)):
                    oSheet.getCellByPosition(6, SR).Formula = '=' + str(mis[4])
                else:
                    #~ MsgBox(id_vc,'idvoce')
                    #~ MsgBox(mis[4],str(eval(mis[4])))
                    #~ MsgBox(str(mis),'')
                    oSheet.getCellByPosition(6, SR).Value = eval(mis[4].replace(',','.'))
            else:
                pass
                    
            if mis[5] != None: #larghezza
                if any(o in mis[5] for o in ('+', '*', '/', '-', )):
                    oSheet.getCellByPosition(7, SR).Formula = '=' + str(mis[5])
                else:
                    try:
                        eval(mis[5])
                        oSheet.getCellByPosition(7, SR).Value = eval(mis[5].replace(',','.'))
                    except:
                        MsgBox(str(type(mis[5])),'') ; return

            if mis[6] != None: #HPESO
                if any(o in mis[6] for o in ('+', '*', '/', '-', )):
                    oSheet.getCellByPosition(8, SR).Formula = '=' + str(mis[6])
                else:
                    oSheet.getCellByPosition(8, SR).Value = eval(mis[6])

            if mis[8] == '2':
                #~ oRangeAddress = oSheet.getCellRangeByPosition(0, SR+1, 1, SR+1).getRangeAddress()
                #~ oSheet.removeRange(oRangeAddress, 3) # Mode.ROWS
                #~ MsgBox('parziale','') #; return
                parziale_core(SR)
                oSheet.getRows().removeByIndex(SR+1, 1)
                descrizione =''
                
            #~ MsgBox(str(SR),'SR')
            
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
                        oSheet.getCellByPosition(5, SR).Formula = '=' + pu
                    else:
                        oSheet.getCellByPosition(5, SR).Value = eval(pu)

                if '-' in mis[7] and va*vb*vc*vd >0: #quantità
                    if mis[3] != None: #partiuguali
                        oSheet.getCellByPosition(5, SR).Formula =  '=-1*(' + str(mis[3]) +')'
                    else:
                        oSheet.getCellByPosition(5, SR).Value = eval('-1')

                if mis[9] != '-2':
                    vedi = diz_vv.get(mis[9])
                    #~ MsgBox(vedi,str(SR))
                    vedi_voce(SR, vedi)
                    if va*vb*vc*vd != 1:
                        oSheet.getCellByPosition(8, SR).Value = va*vb*vc*vd
                    if '-' in mis[7] and oSheet.getCellByPosition(5, SR).Value > 0:
                        oSheet.getCellByPosition(8, SR).Value = va*vb*vc*vd*-1
                    oSheet.getCellByPosition(7, SR).String = ''
                    oSheet.getCellByPosition(6, SR).String = ''

            except TypeError:
                pass
            SR = SR+1
    Numera_Voci()
    oDoc.CurrentController.ZoomValue = 100
    ###
    try:
        Tutti_Subtotali()# ricalcola i totali di categorie e subcategorie
        Sincronizza_SottoCap_Tag_Capitolo_Cor()# sistemo gli idcat voce per voce
    except:
        pass
    MsgBox('Importazione eseguita con successo\n in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!','')
    #~ MsgBox(str(diz_vv),str(type(diz_vv)))

    #~ prin('')
    #~ MsgBox(str((datetime.now() - date).total_seconds()),'')
# XPWE_import ##########################################################
########################################################################
#VARIABILI GLOBALI:
noVoce = ('Livello-1-scritta', 'livello2 valuta', 'comp Int_colonna')
siVoce = ('Comp Start Attributo', 'comp progress', 'comp 10 s','Comp End Attributo', )
siVoce_R = ('Comp Start Attributo_R', 'comp 10 s_R','Comp End Attributo_R', )

createUnoService = (
        XSCRIPTCONTEXT
        .getComponentContext()
        .getServiceManager()
        .createInstance 
                    )
########################################################################
def filedia(titolo='Scegli il file...', est='*.*',  mode=0):
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
            oDisp = oFilePicker.getFiles()[0]
    except:
        MsgBox('Il file non è stato selezionato', 'ATTENZIONE!') ; return
    return oDisp.split('///')[-1].replace('%20',' ')
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
        MsgBox(oDisp,'')
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
    
def MsgBox(s,t=''): # s = messaggio | t = titolo
    doc = XSCRIPTCONTEXT.getDocument()
    parentwin = doc.CurrentController.Frame.ContainerWindow
    #~ s = 'This a message'
    #~ t = 'Title of the box'
    #~ res = MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO_CANCEL + DEFAULT_BUTTON_NO)

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

#g_exportedScripts = TestMessageBox,
########################################################################
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
def filtra_codice():
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
        #~ sStRange.RangeAddress
        sopra = sStRange.RangeAddress.StartRow
        voce = oSheet.getCellByPosition(1, sopra+1).String
        #~ MsgBox(voce)
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

def struttura_ComputoM():
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    struct(1)
    struct(2)
    struct(3)
    struct(4)

def struct(l):
    ''' mette in vista struttura secondo categorie
    l { integer } : specifica il livello di categoria
    1 = categoria
    2 = sotto-categoria
    3 = intera voce di misurazione
    4 = righi di misurazione di ogni voce
    '''
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if l == 1:
        stile = 'Livello-1-scritta'
        myrange = ('Livello-1-scritta', 'Comp TOTALI',)
        Dsopra = 1
        Dsotto = 1
    elif l == 2:
        stile = 'livello2 valuta'
        myrange = ('livello2 valuta','Livello-1-scritta', 'Comp TOTALI',)
        Dsopra = 1
        Dsotto = 1
    
    elif l == 3:
        stile = 'Comp Start Attributo'
        myrange = ('Comp End Attributo', 'Comp TOTALI',)
        Dsopra = 0
        Dsotto = 0
    elif l == 4:
        stile = 'Comp Start Attributo'
        myrange = ('Comp End Attributo', 'Comp TOTALI',)
        Dsopra = 2
        Dsotto = 1
    
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
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
