#!
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
import os, sys, uno, unohelper, pyuno, logging, shutil
# cos'e' il namespace:
# http://www.html.it/articoli/il-misterioso-mondo-dei-namespaces-1/
from datetime import datetime
from com.sun.star.beans import PropertyValue
from xml.etree.ElementTree import ElementTree
########################################################################
def LeenO_path():
    ctx = XSCRIPTCONTEXT.getComponentContext()
    pir = ctx.getValueByName('/singletons/com.sun.star.deployment.PackageInformationProvider')
    expath=pir.getPackageLocation('org.giuseppe-vizziello.leeno')
    return (expath)
class New_file:
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
def xxdebug():
    New_file.computo()
    New_file.listino()
def debughelp_uno():
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
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
def insRows(lrow, nrighe): #forse inutile
    '''Inserisce nrighe nella posizione lrow - alternativo a
    oSheet.getRows().insertByIndex(lrow,1)'''
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
    oCellRangeAddr.EndRow = lrow+4-1
    oSheet.insertCells(oCellRangeAddr, 3)   # com.sun.star.sheet.CellInsertMode.ROW
########################################################################
def uFindString (sString, oSheet):
    '''Trova la prima ricorrenza di una stringa (sString) riga per riga
    in un foglio di calcolo (oSheet) e restituisce una tupla (IDcolonna, IDriga)'''
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
def _gotoCella (IDcol,IDrow):
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
def Circoscrive_Voce_Computo_Att (lrow):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ lrow = Range2Cell()[1]
    if oSheet.Name in ( 'COMPUTO','CONTABILITA'):
        #~ stile=oSheet.getCellByPosition(0, lrow).CellStyle
        if oSheet.getCellByPosition(0, lrow).CellStyle in ('comp progress', 'comp 10 s', 'Comp Start Attributo', 'Comp End Attributo', 'Comp Start Attributo_R', 'comp 10 s_R', 'Comp End Attributo_R'):
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
    #~ oDoc.CurrentController.select(celle)
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
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ lrow = Range2Cell()[1]
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
def Range2Cell():
    '''Partendo da una selezione qualsiasi restituisce una tupla (IDcolonna, IDriga)'''
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
    oCell = oSheet.getCellByPosition(0, 0)
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.gotoEndOfUsedArea(True)
    aAddress = oCursor.RangeAddress
    return aAddress#.EndColumn, aAddress.EndRow)
########################################################################
# numera le voci di computo o contabilità
def Numera_Voci():
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    n , lastRow = 1 , getLastUsedCell(oSheet).EndRow+1
    for lrow in range(0,lastRow):
        if oSheet.getCellByPosition (1,lrow).CellStyle == 'comp Art-EP' or oSheet.getCellByPosition (1,lrow).CellStyle == 'comp Art-EP_R':
            oSheet.getCellByPosition (0,lrow).Value = n
            n = n+1
########################################################################
# ins_voce_computo #####################################################
def ins_voce_computo_grezza(lrow):
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #~ lrow = Range2Cell()[1]
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
    celle=oSheet.getCellRangeByPosition(0, lrow, 43,lrow+3)# 'seleziona la cella
    oDoc.CurrentController.select(celle)
    celle.Rows.OptimalHeight = True
########################################################################
    _gotoCella(1,lrow+1)
########################################################################
# ins_voce_computo #####################################################
def ins_voce_computo():
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    eRow = uFindString ('TOTALI COMPUTO', oSheet)[1]
    #~ MsgBox(str(eRow),'eRow')
    cella = oSheet.getCellByPosition(1,lrow)
    if lrow <= 3:
        lrow = 3
    elif oSheet.getCellByPosition(1, lrow).CellStyle == 'Livello-1-scritta' or oSheet.getCellByPosition(1, lrow).CellStyle == 'livello2 valuta':
        lrow = lrow+1
    elif lrow >= eRow:
        lrow = eRow
    elif cella.CellStyle != 'comp sotto Bianche':
        while cella.CellStyle != 'comp sotto Bianche':
            lrow = lrow+1
            cella = oSheet.getCellByPosition(1,lrow)
        lrow = lrow+1
        #~ if cella.CellStyle == 'comp sotto Bianche':
            #~ lrow = lrow+1
    elif cella.CellStyle == 'comp sotto Bianche':
        lrow = lrow+1
    ins_voce_computo_grezza(lrow)
    Numera_Voci()
########################################################################
########################################################################
# XML_import ###########################################################
def XML_import (): #(filename):
    date1 = datetime.now()
    '''Routine di importazione di un prezziario XML formato SIX. Molto
    liberamente tratta da PreventARES https://launchpad.net/preventares
    di <Davide Vescovini> <davide.vescovini@gmail.com>'''
    filename = filedia('Scegli il file XML-SIX da convertire...')
    #~ MsgBox(filename,'')
    # inizializzazioe delle variabili
    lista_articoli = list() # lista in cui memorizzare gli articoli da importare
    diz_um = dict() # array per le unità di misura
    # stringhe per descrizioni articoli
    desc_breve = str()
    desc_estesa = str()
    # effettua il parsing del file XML
    tree = ElementTree()
    tree.parse(filename)
    # ottieni l'item root
    root = tree.getroot()
    logging.debug(list(root))
    # effettua il parsing di tutti gli elemnti dell'albero XMLsub nuova_voce_computo_at
    iter = tree.getiterator()
    listaSOA = []
    articolo = []
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
        #~ elif elem.tag == '{six.xsd}przDescrizione':
            #~ logging.debug(elem.get('breve'))
            #~ # inserisci il titolo del prezziario
            #~ nome=elem.get('breve')
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
            articolo = (prod_id,            #0
                        tariffa,            #1
                        desc_voce,        #2
                        desc_estesa,        #3 non usata
                        unita_misura,       #4
                        valore,             #5
                        quantita,           #6
                        mdo,                #7
                        sicurezza)         #8 %
            lista_articoli.append(articolo)
# compilo la tabella ###################################################
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('Listino')
    #~ if oSheet.getCellRangeByName('F5').String == '': #se la cella F5 è vuota, elimina riga
    if oSheet.getCellByPosition(5, 4).String == '': #se la cella F5 è vuota, elimina riga
        #~ oRangeAddress = oSheet.getCellRangeByName('A5:B5').getRangeAddress()
        oRangeAddress = oSheet.getCellRangeByPosition(0,4,1,4).getRangeAddress()
        oSheet.removeRange(oRangeAddress, 3) # Mode.ROWS
# nome del prezzario ###################################################
    oSheet.getCellByPosition(2, 0).String = nome
    scarto=6 #primo rigo dati
    for riga, elem in enumerate(lista_articoli):
# ciclo for ottimizzato da Dabniele Zambelli https://plus.google.com/u/0/107607056036247518268/about
        riga += scarto
        oSheet.getCellByPosition(0, riga).String = elem[0]    #prod_id
        oSheet.getCellByPosition(2, riga).String = elem[1]    #tariffa
        oSheet.getCellByPosition(4, riga).String = elem[2]    #desc_voce
        oSheet.getCellByPosition(6, riga).String = elem[4]    #unita_misura
        if elem[5] != '':
            oSheet.getCellByPosition(7, riga).Value = elem[5] #valore
        if elem[7] != '' and elem[7] != 0:
            oSheet.getCellByPosition(8, riga).Value = elem[7] /100 #manodopera % (lo stile % di LibreOffice moltiplica per 100)
        if elem[8] != '':
            oSheet.getCellByPosition(11, riga).Value = float (elem[5] * elem[8] / 100) #sicurezza
    MsgBox(str(date1 - datetime.now()),'')
# XML_import ###########################################################
########################################################################
def debug():#XML_import_BOLZANO ():
    date1 = datetime.now()
    New_file.listino()
    '''Routine di importazione di un prezziario XML formato SIX. Molto
    liberamente tratta da PreventARES https://launchpad.net/preventares
    di <Davide Vescovini> <davide.vescovini@gmail.com>
    Versione bilingue'''
    filename = filedia('Scegli il file XML-SIX da convertire...')
    #~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/Bolzano/2015/prova_ZIX.xml'
    # inizializzazioe delle variabili
    lista_articoli = list() # lista in cui memorizzare gli articoli da importare
    diz_um = dict() # array per le unità di misura
    # stringhe per descrizioni articoli
    desc_breve = str()
    desc_estesa = str()
    # effettua il parsing del file XML
    tree = ElementTree()
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
            try:
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
            try:
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
                test = 0
                suffB_IT, suffB_DE, suffE_IT, suffE_DE = desc_breve1, desc_breve2, desc_estesa1, desc_estesa2
                valore = ''
                quantita = ''
            articolo = (prod_id,            #0
                        tariffa,            #1
                        desc_voce,        #2
                        desc_estesa,        #3 non usata
                        unita_misura,       #4
                        valore,             #5
                        quantita,           #6
                        mdo,                #7
                        sicurezza)         #8 %
            lista_articoli.append(articolo)
# compilo la tabella ###################################################
    #~ MsgBox (str(date1 - datetime.now()), '')
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('Listino')
    oSheet.getCellByPosition(6, 0).Value = 1 #Livello di accodamento
    if oSheet.getCellByPosition(5, 4).String == '': #se la cella F5 è vuota, elimina riga
        oRangeAddress = oSheet.getCellRangeByPosition (0,4,1,4).getRangeAddress()
        oSheet.removeRange(oRangeAddress, 3) # Mode.ROWS
# nome del prezzario ###################################################
    oSheet.getCellByPosition(2, 0).String = nome
    scarto=6 #primo rigo dati
    for riga, elem in enumerate(lista_articoli):
# ciclo for ottimizzato da Dabniele Zambelli https://plus.google.com/u/0/107607056036247518268/about
        riga += scarto
        oSheet.getCellByPosition(0, riga).String = elem[0]    #prod_id
        oSheet.getCellByPosition(2, riga).String = elem[1]    #tariffa
        oSheet.getCellByPosition(4, riga).String = elem[2]    #desc_voce
        oSheet.getCellByPosition(6, riga).String = elem[4]    #unita_misura
        if elem[5] != '':
            oSheet.getCellByPosition(7, riga).Value = elem[5] #valore
        if elem[7] != '' and elem[7] != 0:
            oSheet.getCellByPosition(8, riga).Value = elem[7] /100 #manodopera % (lo stile % di LibreOffice moltiplica per 100)
        if elem[8] != '':
            oSheet.getCellByPosition(11, riga).Value = float (elem[5] * elem[8] / 100) #sicurezza
    eRow = uFindString ('ATTENZIONE!', oSheet)[1]+1
    lastRow = getLastUsedCell(oSheet).EndRow+1
    oSheet.getCellRangeByPosition (0,eRow,5,lastRow).CellStyle = 'List-stringa-sin'
    oSheet.getCellRangeByPosition (6,eRow,6,lastRow).CellStyle = 'List-stringa-centro'
    oSheet.getCellRangeByPosition (7,eRow,11,lastRow).CellStyle = 'List-num-euro'
    oSheet.getCellRangeByPosition (8,eRow,8,lastRow).CellStyle = 'List-%'
    oSheet.getCellRangeByPosition (10,eRow,10,lastRow).CellStyle = 'List-%'
    MsgBox (str(date1 - datetime.now()), '')
# XML_import_BOLZANO ###################################################
########################################################################
# XPWE_import ##########################################################
def debugXPWE_import (): #(filename):
    #~ filename = filedia('Scegli il file XPWE da importare...')
    #~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/xpwe/xpwe_prova.xpwe'
    #~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/_Prezzari/2005/da_pwe/Esempio_Progetto_CorpoMisura.xpwe'
    #~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/Sicilia/sicilia2013.xpwe'
    filename = 'W:\\_dwg\\ULTIMUSFREE\\xpwe\\berlingieri.xpwe'
    #~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/xpwe/berlingieri.xpwe'
    #~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/xpwe/prova_vedi_voce.xpwe'
    #~ filename = 'W:\\BAR\\S. Francesco di assisi Gravina\\primus\\Campanile_san_francesco.xpwe'
    #~ filename = filedia('Scegli il file XPWE da importare...')
    '''xml auto indent: http://www.freeformatter.com/xml-formatter.html'''
    #~ filename = filedia('Scegli il file XML-SIX da convertire...')
    # inizializzazioe delle variabili
    date1 = datetime.now()
    lista_articoli = list() # lista in cui memorizzare gli articoli da importare
    diz_ep = dict() # array per le voci di elenco prezzi
    # effettua il parsing del file XML
    tree = ElementTree()
    tree.parse(filename)
    # ottieni l'item root
    root = tree.getroot()
    logging.debug(list(root))
    # effettua il parsing di tutti gli elemnti dell'albero XML
    iter = tree.getiterator()
    nome_file = root.find('FileNameDocumento').text
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
        lista_cap = list()
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
            percentuale = elem.find('Percentuale').text
            diz = dict ()
            diz['id_sc'] = id_sc
            diz['dessintetica'] = dessintetica
            diz['percentuale'] = percentuale
            lista_supcat.append(diz)
###
#PweDGCategorie
    if CapCat.find('PweDGCategorie'):
        PweDGCategorie = CapCat.find('PweDGCategorie').getchildren()
        lista_cat = list()
        for elem in PweDGCategorie:
            id_sc = elem.get('ID')
            dessintetica = elem.find('DesSintetica').text
            percentuale = elem.find('Percentuale').text
            diz = dict ()
            diz['id_sc'] = id_sc
            diz['dessintetica'] = dessintetica
            diz['percentuale'] = percentuale
            lista_cat.append(diz)
###
#PweDGSubCategorie
    if CapCat.find('PweDGSubCategorie'):
        PweDGSubCategorie = CapCat.find('PweDGSubCategorie').getchildren()
        lista_subcat = list()
        for elem in PweDGSubCategorie:
            id_sc = elem.get('ID')
            dessintetica = elem.find('DesSintetica').text
            percentuale = elem.find('Percentuale').text
            diz = dict ()
            diz['id_sc'] = id_sc
            diz['dessintetica'] = dessintetica
            diz['percentuale'] = percentuale
            lista_subcat.append(diz)
###
    PweDGModuli = dati.getchildren()[2][0].getchildren()    #PweDGModuli
    speseutili = PweDGModuli[0].text
    spesegenerali = PweDGModuli[1].text
    utiliimpresa = PweDGModuli[2].text
    oneriaccessorisc = PweDGModuli[3].text
###
    misurazioni = root.find('PweMisurazioni')
    PweElencoPrezzi = misurazioni.getchildren()[0]
###
# leggo l'elenco prezzi ################################################
    epitems = PweElencoPrezzi.findall('EPItem')
    lista_articoli = dict()
    for elem in epitems:
        id_ep = elem.get('ID')
        diz_ep = dict()
        tipoep = elem.find('TipoEP').text
        tariffa = elem.find('Tariffa').text
        articolo = elem.find('Articolo').text
        desridotta = elem.find('DesRidotta').text
        destestesa = elem.find('DesEstesa').text
        desridotta = elem.find('DesBreve').text
        desbreve = elem.find('DesBreve').text
        unmisura = elem.find('UnMisura').text
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
        lista_articoli[id_ep] = diz_ep
###
# leggo voci di misurazione e righe ####################################
    #~ try:
    PweVociComputo = misurazioni.getchildren()[1]
    vcitems = PweVociComputo.findall('VCItem')
    lista_misure = list()
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
        for el in righi_mis:
            diz_rig = dict()
            rgitem = el.get('ID')
            idvv = el.find('IDVV').text
            descrizione = el.find('Descrizione').text
            partiuguali = el.find('PartiUguali').text
            lunghezza = el.find('Lunghezza').text
            larghezza = el.find('Larghezza').text
            hpeso = el.find('HPeso').text
            quantita = el.find('Quantita').text
            flags = el.find('Flags').text
            diz_rig['rgitem'] = rgitem
            diz_rig['idvv'] = idvv
            diz_rig['descrizione'] = descrizione
            diz_rig['partiuguali'] = partiuguali
            diz_rig['lunghezza'] = lunghezza
            diz_rig['larghezza'] = larghezza
            diz_rig['hpeso'] = hpeso
            diz_rig['quantita'] = quantita
            diz_rig['flags'] = flags
            lista_rig.append(diz_rig)
        diz_misura['id_vc'] = id_vc
        diz_misura['id_ep'] = id_ep
        diz_misura['quantita'] = quantita
        diz_misura['datamis'] = datamis
        diz_misura['flags'] = flags
        diz_misura['idspcat'] = idspcat
        diz_misura['idcat'] = idcat
        diz_misura['idsbcat'] = idsbcat
        diz_misura['lista_rig'] = lista_rig
        #~ vcitem = (id_vc, id_ep, quantita, datamis, flags, idspcat, idcat, idsbcat)
        #~ lista_misure[id_vc] = vcitem
        lista_misure.append(diz_misura)
    #~ except IndexError:
        #~ MsgBox('questo risulta essere un elenco prezzi senza voci di misurazione','ATTENZIONE!')
        #~ pass
    #~ articoli = open ('/home/giuserpe/.config/libreoffice/4/user/uno_packages/cache/uno_packages/luds59ep.tmp_/LeenO-3.11.3.dev-150714180321.oxt/pyLeenO/articoli.txt', 'w')
    #~ print (str(lista_articoli), file=articoli)
    #~ articoli.close()
    #~ misure = open ('/home/giuserpe/.config/libreoffice/4/user/uno_packages/cache/uno_packages/luds59ep.tmp_/LeenO-3.11.3.dev-150714180321.oxt/pyLeenO/misure.txt', 'w')
    #~ print (str(lista_misure), file=misure)
    #~ misure.close()
    #~ MsgBox('ho stampato', '')
###
    oDoc = XSCRIPTCONTEXT.getDocument()
# compilo Anagrafica generale ##########################################
    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition (2,2).String = oggetto
    oSheet.getCellByPosition(2,3).String = comune
    oSheet.getCellByPosition(2,5).String = committente
    if impresa != None:
        oSheet.getCellByPosition(3,16).String = impresa
# compilo Elenco Prezzi ################################################
    oSheet = oDoc.getSheets().getByName('S1')
    oSheet.getCellByPosition(7,318).Value = float(oneriaccessorisc)/100
    oSheet.getCellByPosition(7,319).Value = float(spesegenerali)/100
    oSheet.getCellByPosition(7,320).Value = float(utiliimpresa)/100
# compilo Elenco Prezzi ################################################
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = 0
    oCellRangeAddr.StartColumn = 0
    oCellRangeAddr.StartRow = 3
    oCellRangeAddr.EndColumn = 0
    oCellRangeAddr.EndRow = 3
    lrow=4 #primo rigo dati
    idxcol=0
###
# inserisco prima solo le righe se no mi fa casino
    for elem in lista_articoli:
        oSheet.insertCells(oCellRangeAddr, 3)   # com.sun.star.sheet.CellInsertMode.ROW
        oSheet.getCellByPosition(0,3).CellStyle = 'EP-aS'
        oSheet.getCellByPosition(1,3).CellStyle = 'EP-a'
        oSheet.getCellRangeByPosition(2,3,7,3).CellStyle = 'EP-mezzo'
        oSheet.getCellByPosition(5, 3).CellStyle = 'EP-mezzo %'
        oSheet.getCellRangeByPosition (8,3,9,3).CellStyle = 'EP-sfondo'
###
# sommario computo
        oSheet.getCellByPosition(10, 3).Formula =  '=sumif(AA;A4;BB)'
        oSheet.getCellByPosition(10, 3).CellStyle = 'EP statistiche_q'
        oSheet.getCellByPosition(11, 3).Formula = '=K4*E4'
        oSheet.getCellByPosition(11, 3).CellStyle = 'EP statistiche'
###
# sommario contabilità
        oSheet.getCellByPosition(13, 3).CellStyle = 'EP statistiche_Contab_q'
        oSheet.getCellByPosition(13, 3).Formula = '=SUMIF(GG;A4;G1G1)'
        oSheet.getCellByPosition(14, 3).CellStyle = 'EP statistiche_Contab'
        oSheet.getCellByPosition(14, 3).Formula = '=N4*E4'
        lrow=lrow+1
    lrow=4 #primo rigo dati
###
# inserisco le voci di Elenco Prezzi
    for elem in lista_articoli.items():
        oSheet.getCellByPosition(0, lrow-1).String = elem[1].get('tariffa')
        oSheet.getCellByPosition(1, lrow-1).String = elem[1].get('destestesa').strip()
        if elem[1].get('unmisura') != None:
            oSheet.getCellByPosition(2, lrow-1).String = elem[1].get('unmisura')
        oSheet.getCellByPosition(4, lrow-1).Value = elem[1].get('prezzo1')
        lrow=lrow+1
###
# Inserisco i dati nel COMPUTO #########################################
    #~ oSheet=ThisComponent.currentController.activeSheet
    oSheet = oDoc.getSheets().getByName('COMPUTO')
    oDoc.CurrentController.select(oSheet)
    iSheet_num = oSheet.RangeAddress.Sheet
###
    #~ lista_misure.reverse()
    lrow=3 #primo rigo dati
    for el in lista_misure:
        ins_voce_computo_grezza(lrow)
        ID = el.get('id_ep')
        oSheet.getCellByPosition(1, lrow+1).String = lista_articoli.get(ID).get('tariffa')
        lmis=lrow+2
        for mis in el.get('lista_rig'):
            if mis.get('descrizione') != None:
                oSheet.getCellByPosition(2, lmis).String = mis.get('descrizione').strip()
            if mis.get('partiuguali') != None:
                oSheet.getCellByPosition(5, lmis).Value = eval(mis.get('partiuguali'))
            if mis.get('lunghezza') != None:
                oSheet.getCellByPosition(6, lmis).Value = eval(mis.get('lunghezza'))
            if mis.get('larghezza') != None:
                oSheet.getCellByPosition(7, lmis).Value = eval(mis.get('larghezza'))
            if mis.get('hpeso') != None:
                oSheet.getCellByPosition(8, lmis).Value = eval(mis.get('hpeso'))
            lmis = lmis+1
            copia_riga_computo(lmis)
        lrow=lmis+2
    Numera_Voci()
    MsgBox(str(date1 - datetime.now()),'')
# XPWE_import ##########################################################
########################################################################
########################################################################
########################################################################
import traceback
from com.sun.star.awt import Rectangle
def filedia(titolo):
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
def MsgBox(s,t): # s = messaggio | t = titolo
    doc = XSCRIPTCONTEXT.getDocument()
    parentwin = doc.CurrentController.Frame.ContainerWindow
    #~ s = 'This a message'
    #~ t = 'Title of the box'
    #~ res = MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO_CANCEL + DEFAULT_BUTTON_NO)
 
    #~ s = res
    #~ t = 'Titolo'
    if t == None:
        t='messaggio'
    MessageBox(parentwin, s, t, 'infobox')

# Show a message box with the UNO based toolkit
def MessageBox(ParentWin, MsgText, MsgTitle, MsgType=MESSAGEBOX, MsgButtons=BUTTONS_OK):
    ctx = uno.getComponentContext()
    sm = ctx.ServiceManager
    sv = sm.createInstanceWithContext('com.sun.star.awt.Toolkit', ctx) 
    myBox = sv.createMessageBox(ParentWin, MsgType, MsgButtons, MsgTitle, MsgText)
    return myBox.execute()
 
#g_exportedScripts = TestMessageBox,
########################################################################
########################################################################
#import pdb; pdb.set_trace() #debugger
