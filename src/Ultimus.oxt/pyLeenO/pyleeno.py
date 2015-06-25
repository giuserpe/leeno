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
import os, sys, uno, unohelper, pyuno
# cos'e' il namespace:
# http://www.html.it/articoli/il-misterioso-mondo-dei-namespaces-1/
import logging
from xml.etree.ElementTree import ElementTree
########################################################################
def leeno_path():
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.getCellRangeByName('a1').String = dir(uno.getComponentContext).__str__()
    oSheet.getCellRangeByName('a2').String = dir(uno.getComponentContext.__name__).__str__()
    #~ strin = ''
    #~ if 'win' in sys.platform:
        #~ lst = sys.path[0].split('\\')
        #~ lstr = len(lst)
        #~ for el in lst:# range(0,lstr-1):
            #~ strin = strin + el + '\\'
        #~ oSheet.getCellRangeByName('a2').String = strin
def debughelp_uno():
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.getCellRangeByName('B9').String = dir(os).__str__()
    oSheet.getCellRangeByName('B10').String = dir(uno.__package__.title.__name__).__str__()
    oSheet.getCellRangeByName('B11').String = dir(unohelper).__str__()
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
def insRows(lrow, nrighe):
    """Inserisce nrighe nella posizione lrow"""
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
    #~ MsgBox(str(lrow),'lrow')
########################################################################
def uFindString (sString, oSheet):
    """Trova la prima ricorrenza di una stringa (sString) riga per riga
    in un foglio di calcolo (oSheet) e restituisce una tupla (IDcolonna, IDriga)"""
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
def ColumnNumberToName(oSheet,cColumnNumb):
    """Trasforma IDcolonna in Nome"""
    #~ oDoc = XSCRIPTCONTEXT.getDocument()
    #~ oSheet = oDoc.CurrentController.ActiveSheet
    oColumns = oSheet.getColumns()
    oColumn = oColumns.getByIndex(cColumnNumb).Name
    return oColumn
########################################################################
def ColumnNameToNumber(oSheet,cColumnName):
    """Trasforma il nome colonna in IDcolonna"""
    #~ oDoc = XSCRIPTCONTEXT.getDocument()
    #~ oSheet = oDoc.CurrentController.ActiveSheet
    oColumns = oSheet.getColumns()
    oColumn = oColumns.getByName(cColumnName)
    oRangeAddress = oColumn.getRangeAddress()
    nColumn = oRangeAddress.StartColumn
    return nColumn
########################################################################
# Range2Cell ###########################################################
def Range2Cell():
    """Partendo da una selezione qualsiasi restituisce una tupla (IDcolonna, IDriga)"""
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
    n , lastRow = 4 , getLastUsedCell(oSheet).EndRow+1
    for lrow in range(0,lastRow):
        if oSheet.getCellByPosition (1,lrow).CellStyle == 'comp Art-EP' or oSheet.getCellByPosition (1,lrow).CellStyle == 'comp Art-EP_R':
            oSheet.getCellByPosition (0,lrow).Value = n
            n = n+1
########################################################################
# ins_voce_computo #####################################################
def debug(): #ins_voce_computo ():
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = Range2Cell()[1]
    eRow = uFindString ('TOTALI COMPUTO', oSheet)[1]
    #~ MsgBox(str(eRow),'eRow')
    cella = oSheet.getCellByPosition(1,lrow)
    if lrow <= 3:
        lrow = 3
    elif oSheet.getCellByPosition(1, lrow).CellStyle == "Livello-1-scritta" or oSheet.getCellByPosition(1, lrow).CellStyle == "livello2 valuta":
        lrow = lrow+1
    elif lrow >= eRow:
        lrow = eRow
    elif cella.CellStyle != "Comp-Bianche sopra":
        while cella.CellStyle != "Comp-Bianche sopra": #comp sotto Bianche per inserimento SOTTO
            lrow = lrow-1
            cella = oSheet.getCellByPosition(1,lrow)
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
##### vado alla vecchia maniera ########################################
##### copio il range di righe computo da S5 ############################
    oSheetto = oDoc.getSheets().getByName('S5')
    oRangeAddress = oSheetto.getCellRangeByName('$A$9:$AR$12').getRangeAddress()
    oCellAddress = oSheet.getCellByPosition(0,lrow).getCellAddress()
    insRows(lrow,4) #inserisco le righe
    oSheet.copyRange(oCellAddress, oRangeAddress)
########################################################################
# correggo alcune formule
    oSheet.getCellByPosition(13,lrow+3).Formula ="=J"+str(lrow+4)
# sistemo i LINK dei tagG nelle righe sopra al tag vero e prorio...
    oSheet.getCellByPosition(31, lrow+2).Formula = "=AF$"+str(lrow+4)
    oSheet.getCellByPosition(32, lrow+2).Formula = "=AG$"+str(lrow+4)
    oSheet.getCellByPosition(33, lrow+2).Formula = "=AH$"+str(lrow+4)
    oSheet.getCellByPosition(34, lrow+2).Formula = "=AI$"+str(lrow+4)
    oSheet.getCellByPosition(35, lrow+2).Formula = "=AJ$"+str(lrow+4)
    oSheet.getCellByPosition(31, lrow+1).Formula = "=AF$"+str(lrow+4)
    oSheet.getCellByPosition(32, lrow+1).Formula = "=AG$"+str(lrow+4)
    oSheet.getCellByPosition(33, lrow+1).Formula = "=AH$"+str(lrow+4)
    oSheet.getCellByPosition(34, lrow+1).Formula = "=AI$"+str(lrow+4)
    oSheet.getCellByPosition(35, lrow+1).Formula = "=AJ$"+str(lrow+4)
    oSheet.getCellByPosition(31, lrow).Formula = "=AF$"+str(lrow+4)
    oSheet.getCellByPosition(32, lrow).Formula = "=AG$"+str(lrow+4)
    oSheet.getCellByPosition(33, lrow).Formula = "=AH$"+str(lrow+4)
    oSheet.getCellByPosition(34, lrow).Formula = "=AI$"+str(lrow+4)
    oSheet.getCellByPosition(35, lrow).Formula = "=AJ$"+str(lrow+4)
    celle=oSheet.getCellRangeByPosition(0, lrow, 43,lrow+3)# 'seleziona la cella
    oDoc.CurrentController.select(celle)
    celle.Rows.OptimalHeight = True
########################################################################
    Numera_Voci()
    _gotoCella(1,lrow+1)

########################################################################
########################################################################
# XML_import ###########################################################
def XML_import (): #(filename):
    """Routine di importazione di un prezziario XML formato SIX. Molto
    liberamente tratta da PreventARES https://launchpad.net/preventares
    di <Davide Vescovini> <davide.vescovini@gmail.com>"""
    filename = filedia('Scegli il file XML-SIX da convertire...')
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
    if len(prezzario.findall("{six.xsd}przDescrizione")) == 2:
        if prezzario.findall("{six.xsd}przDescrizione")[0].get('lingua') == lingua_scelta:
            nome = prezzario.findall("{six.xsd}przDescrizione")[0].get('breve')
        else:
            nome = prezzario.findall("{six.xsd}przDescrizione")[1].get('breve')
    else:
        nome = prezzario.findall("{six.xsd}przDescrizione")[0].get('breve')
########################################################################
    for elem in iter:
        # esegui le verifiche sulla root dell'XML
        if elem.tag == "{six.xsd}intestazione":
            intestazioneId= elem.get("intestazioneId")
            lingua= elem.get("lingua")
            separatore= elem.get("separatore")
            separatoreParametri= elem.get("separatoreParametri")
            valuta= elem.get("valuta")
            autore= elem.get("autore")
            versione= elem.get("versione")
            # inserisci i dati generali
            #~ self.update_dati_generali (nome=None, cliente=None,
                                       #~ redattore=autore,
                                       #~ ricarico=1,
                                       #~ manodopera=None,
                                       #~ sicurezza=None,
                                       #~ indirizzo=None,
                                       #~ comune=None, provincia=None,
                                       #~ valuta=valuta)
        elif elem.tag == "{six.xsd}categoriaSOA":
            soaId = elem.get("soaId")
            soaCategoria = elem.get ("soaCategoria")
            soaDescrizione = elem.find("{six.xsd}soaDescrizione")
            if soaDescrizione != None:
                breveSOA = soaDescrizione.get("breve")
            voceSOA = (soaCategoria, soaId, breveSOA)
            listaSOA.append(voceSOA)
        elif elem.tag == "{six.xsd}prezzario":
            prezzarioId = elem.get("prezzarioId")
            przId= elem.get("przId")
            livelli_struttura= len(elem.get("prdStruttura").split("."))
            categoriaPrezzario= elem.get("categoriaPrezzario")
        #~ elif elem.tag == "{six.xsd}przDescrizione":
            #~ logging.debug(elem.get("breve"))
            #~ # inserisci il titolo del prezziario
            #~ nome=elem.get("breve")
########################################################################
        elif elem.tag == "{six.xsd}unitaDiMisura":
            um_id= elem.get("unitaDiMisuraId")
            um_sim= elem.get("simbolo")
            um_dec= elem.get("decimali")
            # crea il dizionario dell'unita di misura
            #~ diz_um[um_id] = um_sim
########################################################################
            #~ udmDescrizione = elem.find("{six.xsd}udmDescrizione")
            #~ if udmDescrizione != None:
                #~ udmDescrizione.get("breve")
########################################################################
            #~ unità di misura
            unita_misura = ""
            try:
                if len (elem.findall("{six.xsd}udmDescrizione")) == 1:
                    #~ unita_misura = elem.getchildren()[0].get('breve')
                    unita_misura = elem.findall("{six.xsd}udmDescrizione")[0].get('breve')
                else:
                    if elem.findall("{six.xsd}udmDescrizione")[1].get('lingua') == lingua_scelta:
                        idx = 1 #ITALIANO
                    else:
                        idx = 0 #TEDESCO
                    unita_misura = elem.findall("{six.xsd}udmDescrizione")[idx].get('breve')
            except IndexError:
                pass
            diz_um[um_id] = unita_misura
########################################################################
        # se il tag è un prodotto fa parte degli articoli da analizzare
        elif elem.tag == "{six.xsd}prodotto":
            prod_id = elem.get("prodottoId")
            if prod_id is not None:
                prod_id = int(prod_id)
            tariffa= elem.get("prdId")
            sic = elem.get("onereSicurezza")
            if sic != None:
                sicurezza = float(sic)
            else:
                sicurezza = ""
########################################################################
            if diz_um.get(elem.get("unitaDiMisuraId")) != None:
                unita_misura = diz_um.get(elem.get("unitaDiMisuraId"))
            else:
                unita_misura = ""
########################################################################
            # verifica e ricava le sottosezioni
            sub_mdo = elem.find("{six.xsd}incidenzaManodopera")
            if sub_mdo != None:
                mdo = float(sub_mdo.text)
            else:
                mdo =""
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
                    desc_breve = ""
                if desc_estesa == None:
                    desc_estesa = ""
                if len(desc_breve) > len (desc_estesa):
                    desc_voce = desc_breve
                else:
                    desc_voce = desc_estesa
            except IndexError:
                pass
########################################################################
            sub_quot = elem.find("{six.xsd}prdQuotazione")
            if sub_quot != None:
                list_nr = sub_quot.get("listaQuotazioneId")
                if sub_quot.get("valore") != None:
                    valore = float(sub_quot.get("valore"))
                if valore == 0:
                    valore = ""
                if sub_quot.get("quantita") is not None: #SERVE DAVVERO???
                    quantita = float(sub_quot.get("quantita"))
            else:
                valore = ""
                quantita = ""
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
    if oSheet.getCellRangeByName('F5').String == "": #se la cella F5 è vuota, elimina riga
        oRangeAddress = oSheet.getCellRangeByName('A5:B5').getRangeAddress()
        oSheet.removeRange(oRangeAddress, 3) # Mode.ROWS
# nome del prezzario ###################################################
    oSheet.getCellRangeByName('C1').String = nome
    i=6 #primo rigo dati
    for elem in lista_articoli:
        try:
            oSheet.getCellRangeByName('A'+ str (i)).String = elem[0]    #prod_id
            oSheet.getCellRangeByName('C'+ str (i)).String = elem[1]    #tariffa
            oSheet.getCellRangeByName('E'+ str (i)).String = elem[2]    #desc_voce
            #~ if elem[3] !="":
                #~ oSheet.getCellRangeByName('E'+ str (i)).String = elem[3]    #desc_estesa
            oSheet.getCellRangeByName('G'+ str (i)).String = elem[4]    #unita_misura
            if elem[5] == "":
                oSheet.getCellRangeByName('H'+ str (i)).String = ""
            else:
                oSheet.getCellRangeByName('H'+ str (i)).Value = elem[5] #valore
            if elem[7] == "" or elem[7] == 0:
                oSheet.getCellRangeByName('I'+ str (i)).String = ""
            else:
                oSheet.getCellRangeByName('I'+ str (i)).Value = elem[7] /100 #manodopera % (lo stile % di LibreOffice moltiplica per 100)
                #~ oSheet.getCellRangeByName('J'+ str (i)).Value = float (elem[5] * elem[7] / 100) #manodopera
            if elem[8] == "":
                oSheet.getCellRangeByName('L'+ str (i)).String = ""
            else:
                #~ oSheet.getCellRangeByName('K'+ str (i)).Value = elem[8] /100 #sicurezza % (lo stile % di LibreOffice moltiplica per 100)
                oSheet.getCellRangeByName('L'+ str (i)).Value = float (elem[5] * elem[8] / 100) #sicurezza
        except IndexError:
            pass
        i=i+1
    #~ listaSOA.sort()
    #~ for elem in listaSOA:
        #~ oSheet.getCellRangeByName('B'+ str (i)).String = elem[0]
        #~ oSheet.getCellRangeByName('C'+ str (i)).String = elem[1]    #Codice
        #~ oSheet.getCellRangeByName('E'+ str (i)).String = elem[2]    #Codice
        #~ i=i+1

# XML_import ###########################################################
########################################################################
########################################################################
# XPWE_import ##########################################################
def XPWE_import (): #(filename):
    #~ filename = filedia('Scegli il file XPWE da importare...')
    #~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/xpwe/xpwe_prova.xpwe'
    #~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/_Prezzari/2005/da_pwe/Esempio_Progetto_CorpoMisura.xpwe'
    #~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/Sicilia/sicilia2013.xpwe'
    #~ filename = 'W:\\_dwg\\ULTIMUSFREE\\xpwe\\berlingieri.xpwe'
    filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/xpwe/berlingieri.xpwe'
    """xml auto indent: http://www.freeformatter.com/xml-formatter.html"""
    #~ filename = filedia('Scegli il file XML-SIX da convertire...')
    # inizializzazioe delle variabili
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
########################################################################
    dati = root.find('PweDatiGenerali')
    DatiGenerali = dati.getchildren()[0][0]
    percprezzi = DatiGenerali[0].text
    comune = DatiGenerali[1].text
    provincia = DatiGenerali[2].text
    oggetto = DatiGenerali[3].text
    committente = DatiGenerali[4].text
    impresa = DatiGenerali[5].text
    parteopera = DatiGenerali[6].text
########################################################################
#PweDGCapitoliCategorie
    PweDGSuperCapitoli  = dati.getchildren()[1][0].getchildren()    #PweDGSuperCapitoli
    PweDGCapitoli       = dati.getchildren()[1][1].getchildren()    #PweDGCapitoli
    PweDGSuperCategorie = dati.getchildren()[1][2].getchildren()    #PweDGSuperCategorie
    PweDGCategorie      = dati.getchildren()[1][3].getchildren()    #PweDGCategorie
    PweDGSubCategorie   = dati.getchildren()[1][4].getchildren()    #PweDGSubCategorie
########################################################################
#PweDGSuperCapitoli
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
########################################################################
#PweDGCapitoli
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
########################################################################
#PweDGSuperCategorie
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
########################################################################
#PweDGCategorie
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
########################################################################
#PweDGSubCategorie
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
########################################################################
    PweDGModuli = dati.getchildren()[2][0].getchildren()    #PweDGModuli
    speseutili = PweDGModuli[0].text
    spesegenerali = PweDGModuli[1].text
    utiliimpresa = PweDGModuli[2].text
    oneriaccessorisc = PweDGModuli[3].text
########################################################################
    misurazioni = root.find('PweMisurazioni')
    PweElencoPrezzi = misurazioni.getchildren()[0]
########################################################################
# leggo l'elenco prezzi ################################################
    epitems = PweElencoPrezzi.findall('EPItem')
    lista_articoli = list()
    for elem in epitems:
        diz_ep = dict()
        id_ep = elem.get('ID')
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
        diz_ep['id_ep'] = id_ep
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
        #~ epitem = (tipoep, tariffa, articolo, desridotta, destestesa, desridotta, desbreve, prezzo1, prezzo2, prezzo3, prezzo4, prezzo5, idspcap, idcap, flags, data, adrinternet, pweepanalisi)
        lista_articoli.append(diz_ep)
########################################################################
# leggo voci di misurazione e righe ####################################
    try:
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
                hpeso = el.find('Larghezza').text
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
    except IndexError:
        MsgBox("questo risulta essere un elenco prezzi senza voci di misurazione","ATTENZIONE!")
        pass
########################################################################
    oDoc = XSCRIPTCONTEXT.getDocument()
# compilo Anagrafica generale ##########################################
    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellRangeByName('C3').String = oggetto
    oSheet.getCellRangeByName('C4').String = comune
    oSheet.getCellRangeByName('C6').String = committente
    oSheet.getCellRangeByName('C17').String = impresa
# compilo Elenco Prezzi ################################################
    oSheet = oDoc.getSheets().getByName('S1')
    oSheet.getCellRangeByName('H319').Value = float(oneriaccessorisc)/100
    oSheet.getCellRangeByName('H320').Value = float(spesegenerali)/100
    oSheet.getCellRangeByName('H321').Value = float(utiliimpresa)/100
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
########################################################################
# inserisco prima solo le righe se no mi fa casino
    for elem in lista_articoli:
        try:
            oSheet.insertCells(oCellRangeAddr, 3)   # com.sun.star.sheet.CellInsertMode.ROW
            oSheet.getCellRangeByName('A4').CellStyle = "EP-aS"
            oSheet.getCellRangeByName('B4').CellStyle = "EP-a"
            oSheet.getCellRangeByName('C4').CellStyle = "EP-mezzo"
            oSheet.getCellRangeByName('D4').CellStyle = "EP-mezzo"
            oSheet.getCellRangeByName('E4').CellStyle = "EP-mezzo"
            oSheet.getCellRangeByName('F4').CellStyle = "EP-mezzo %"
            oSheet.getCellRangeByName('G4').CellStyle = "EP-mezzo"
            oSheet.getCellRangeByName('H4').CellStyle = "EP-mezzo"
            oSheet.getCellRangeByName('I4').CellStyle = "EP-sfondo"
            oSheet.getCellRangeByName('J4').CellStyle = "EP-sfondo"
########################################################################
# sommario computo
            oSheet.getCellRangeByName('K4').Formula =  "=sumif(AA;A4;BB)"
            oSheet.getCellRangeByName('K4').CellStyle = "EP statistiche_q"
            oSheet.getCellRangeByName('L4').Formula = "=K4*E4"
            oSheet.getCellRangeByName('L4').CellStyle = "EP statistiche"
########################################################################
# sommario contabilità
            oSheet.getCellRangeByName('N4').CellStyle = "EP statistiche_Contab_q"
            oSheet.getCellRangeByName('N4').Formula = "=SUMIF(GG;A4;G1G1)"
            oSheet.getCellRangeByName('O4').CellStyle = "EP statistiche_Contab"
            oSheet.getCellRangeByName('O4').Formula = "=N4*E4"
        except IndexError:
            pass
        lrow=lrow+1
    lrow=4 #primo rigo dati
########################################################################
# inserisco le voci di Elenco Prezzi
    for elem in lista_articoli:
        try:
            oSheet.getCellRangeByName('A'+ str (lrow)).String = elem['tariffa']
            oSheet.getCellRangeByName('B'+ str (lrow)).String = elem['destestesa']
            if elem['unmisura'] == None:
                oSheet.getCellRangeByName('C'+ str (lrow)).String = ""
            else:
                oSheet.getCellRangeByName('C'+ str (lrow)).String = elem['unmisura']
            oSheet.getCellRangeByName('E'+ str (lrow)).Value = elem['prezzo1']
        except IndexError:
            pass
        lrow=lrow+1
########################################################################
# Inserisco i dati nel COMPUTO #########################################
    #~ oSheet=ThisComponent.currentController.activeSheet
    oSheet = oDoc.getSheets().getByName('COMPUTO')
    iSheet_num = oSheet.RangeAddress.Sheet
    #~ sLinkSheetName = thisComponent.Sheets.getByIndex(iSheet_num).getName()
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet_num
    oCellRangeAddr.StartColumn = 0
    oCellRangeAddr.EndColumn = 0
    oCellRangeAddr.StartRow = 3
    oCellRangeAddr.EndRow = 3+len(lista_misure)*4
    lrow=4 #primo rigo dati
    oSheet.insertCells(oCellRangeAddr, 3)   # com.sun.star.sheet.CellInsertMode.ROW
########################################################################
    #~ for el in lista_misure:


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
            oDisp = u'Cancel' # "Cancel" è il risultato del tasto
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
    #~ s = "This a message"
    #~ t = "Title of the box"
    #~ res = MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO_CANCEL + DEFAULT_BUTTON_NO)
 
    #~ s = res
    #~ t = "Titolo"
    if t == None:
        t='messaggio'
    MessageBox(parentwin, s, t, "infobox")

# Show a message box with the UNO based toolkit
def MessageBox(ParentWin, MsgText, MsgTitle, MsgType=MESSAGEBOX, MsgButtons=BUTTONS_OK):
    ctx = uno.getComponentContext()
    sm = ctx.ServiceManager
    sv = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx) 
    myBox = sv.createMessageBox(ParentWin, MsgType, MsgButtons, MsgTitle, MsgText)
    return myBox.execute()
 
#g_exportedScripts = TestMessageBox,
########################################################################
########################################################################
#import pdb; pdb.set_trace() #debugger
