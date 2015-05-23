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
import os, sys, uno, unohelper
# cos'e' il namespace:
# http://www.html.it/articoli/il-misterioso-mondo-dei-namespaces-1/
import logging
from xml.etree.ElementTree import ElementTree

#~ filename = 'W:/_dwg/ULTIMUSFREE/elenchi/abruzzo/2014/AB2014_03-11-2014.xml'
#~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/abruzzo/2014/AB2014_03-11-2014.xml'
#~ filename = "W:/_dwg/ULTIMUSFREE/elenchi/Calabria/2013/Urbanizzazioni_2013_UCU2.xml"

#~ filename = 'W:/_dwg/ULTIMUSFREE/elenchi/Bolzano/2014/HBED14_OpereEdili_20141103_XmlSix.xml'
#~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/Bolzano/2014/HBED14_OpereEdili_20141103_XmlSix.xml'
#~ filename = 'W:/_dwg/ULTIMUSFREE/elenchi/Piemonte/Cuneo/2015/Prezzario_15BD.xml'

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
    # effettua il parsing di tutti gli elemnti dell'albero XML
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



########################################################################

import os
import uno
import sys
import traceback
from com.sun.star.awt import Rectangle
#
#~ def oTest():
    #~ filedia('Scegli il file da convertire...')

def filedia(titolo):
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
        oModel.PositionX = oDlgWth/2 - 20
        oModel.PositionY = oDlgHgt - 20
        oModel.Width = 40
        oModel.Height = 15
        oModel.Label = u'OK'
        oModel.PushButtonType = 1       # 1 : OK
        # Dialog Modelの仕様に Step Button の仕様を設定
        oDlgModel.insertByName('OkBtn', oModel)
        # ***** [ OK / Cancel  Button 設定 ] *****
        #
        # ***** [ FileCntrol 設定 ] *****
        # FileCntrol 仕様
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
        # Dialog Modelの仕様に FileCntrol の仕様を設定
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
            oDisp = u'Cancelされました。'
        # End Dialog
        oDlg.endExecute()
    except:
        oDisp = traceback.format_exc(sys.exc_info()[2])
    finally:
        return oDisp
        #~ oDoc = XSCRIPTCONTEXT.getDocument()
        #~ oSheet = oDoc.getSheets().getByIndex(0)
        #~ oSheet.getCellRangeByName('A1').String = oDisp # nome file



#g_exportedScripts = xmlsix2ods, import_XML,
########################################################################
#import pdb; pdb.set_trace() #debugger
