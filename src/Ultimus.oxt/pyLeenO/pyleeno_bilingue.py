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

def XML_import ():
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
            nome1 = prezzario.findall("{six.xsd}przDescrizione")[0].get('breve')
            nome2 = prezzario.findall("{six.xsd}przDescrizione")[1].get('breve')
        else:
            nome1 = prezzario.findall("{six.xsd}przDescrizione")[1].get('breve')
            nome2 = prezzario.findall("{six.xsd}przDescrizione")[0].get('breve')
        nome=nome1+"§"+nome2
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
########################################################################
        elif elem.tag == "{six.xsd}unitaDiMisura":
            um_id= elem.get("unitaDiMisuraId")
            um_sim= elem.get("simbolo")
            um_dec= elem.get("decimali")
            # crea il dizionario dell'unita di misura
########################################################################
            #~ unità di misura
            unita_misura = ""
            try:
                if len (elem.findall("{six.xsd}udmDescrizione")) == 1:
                    unita_misura = elem.findall("{six.xsd}udmDescrizione")[0].get('breve')
                else:
                    if elem.findall("{six.xsd}udmDescrizione")[1].get('lingua') == lingua_scelta:
                        unita_misura1 = elem.findall("{six.xsd}udmDescrizione")[1].get('breve')
                        unita_misura2 = elem.findall("{six.xsd}udmDescrizione")[0].get('breve')
                    else:
                        unita_misura1 = elem.findall("{six.xsd}udmDescrizione")[0].get('breve')
                        unita_misura2 = elem.findall("{six.xsd}udmDescrizione")[1].get('breve')
                if unita_misura == None:
                    unita_misura = unita_misura1+" § "+unita_misura2
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
            # descrizione voci
            desc_estesa1, desc_estesa2 = "", ""
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
                        desc_breve1 = ""
                    if desc_breve2 == None:
                        desc_breve2 = ""
                    if desc_estesa1 == None:
                        desc_estesa1 = ""
                    if desc_estesa2 == None:
                        desc_estesa2 = ""
                    desc_breve = desc_breve1 +"§"+ desc_breve2
                    desc_estesa = desc_estesa1 +"§"+ desc_estesa2
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
# riferimento: http://openoffice3.web.fc2.com/index.html
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
########################################################################
import uno
 
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_OK, DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, DEFAULT_BUTTON_YES, DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE
 
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
 
#rif.: https://wiki.openoffice.org/wiki/PythonDialogBox

def MsgB(s,t): # s = messaggio | t = titolo
    doc = XSCRIPTCONTEXT.getDocument()
    parentwin = doc.CurrentController.Frame.ContainerWindow
    #~ s = "This a message"
    #~ t = "Title of the box"
    #~ res = MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO_CANCEL + DEFAULT_BUTTON_NO)
 
    #~ s = res
    #~ t = "Titolo"

    #~ MessageBox(parentwin, s, t, "infobox")
 
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


#g_exportedScripts = xmlsix2ods, import_XML,
########################################################################
#import pdb; pdb.set_trace() #debugger
