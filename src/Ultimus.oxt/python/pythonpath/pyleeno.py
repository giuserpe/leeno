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

# MsgBox('''Per segnalare questo problema,
# contatta il canale Telegram
# https://t.me/leeno_computometrico''', 'ERRORE!')

# documentazione ufficiale: https://api.libreoffice.org/
# import pydevd

    # funzioni per misurare la velocità delle macro
    # datarif = datetime.now()
    # DLG.chi('eseguita in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!')

# from scriptforge import CreateScriptService
from datetime import datetime, date
from xml.etree.ElementTree import Element, SubElement, tostring
from collections import OrderedDict
# from com.sun.star.task import XStatusIndicator


# import distutils.dir_util

import codecs
import subprocess
# import psutil
import re
import traceback
import threading
import time
# import csv

from com.sun.star.awt.FontWeight import BOLD, NORMAL

import os
import shutil
import sys
import uno
import unohelper
import zipfile
import inspect

import SheetUtils
import LeenoUtils
import LeenoSheetUtils
import LeenoToolbars as Toolbars
import LeenoFormat
import LeenoComputo
import LeenoContab
import LeenoGiornale
import LeenoGlobals
import LeenoAnalysis
import LeenoDialogs as DLG
import PersistUtils as PU
import LeenoEvents
import LeenoSettings as LS
import LeenoPdf as LPdf
import DocUtils

import LeenoConfig
cfg = LeenoConfig.Config()

import Dialogs

# cos'e' il namespace:
# http://www.html.it/articoli/il-misterioso-mondo-dei-namespaces-1/

# from com.sun.star.lang import Locale
from com.sun.star.beans import PropertyValue
# from com.sun.star.table.CellContentType import TEXT, EMPTY, VALUE, FORMULA
from com.sun.star.table.CellHoriJustify import RIGHT
from com.sun.star.awt.FontSlant import ITALIC, NONE
from com.sun.star.sheet.CellFlags import \
    VALUE, DATETIME, STRING, ANNOTATION, FORMULA, HARDATTR, OBJECTS, EDITATTR, FORMATTED

from com.sun.star.beans.PropertyAttribute import \
    MAYBEVOID, REMOVEABLE, MAYBEDEFAULT

########################################################################
# https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=27805&p=127383

########################################################################
# IMPORT DEI MODULI SEPARATI DI LEENO
########################################################################

def basic_LeenO(funcname, *args):
    '''Richiama funzioni definite in Basic'''

    xCompCont = LeenoUtils.getComponentContext()
    sm = xCompCont.ServiceManager
    mspf = sm.createInstance("com.sun.star.script.provider.MasterScriptProviderFactory")
    scriptPro = mspf.createScriptProvider("")
    Xscript = scriptPro.getScript(
        f"vnd.sun.star.script:UltimusFree2.{funcname}?language=Basic&location=application")
    Result = Xscript.invoke(args, None, None)
    return Result[0]


# leeno.conf
def MENU_leeno_conf():
    '''
    Visualizza il menù di configurazione
    '''
    oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('S1'):
        Toolbars.AllOff()
        return
    psm = LeenoUtils.getServiceManager()
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDlg_config = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.Dlg_config?language=Basic&location=application"
    )
    # oDialog1Model = oDlg_config.Model

    oSheets = list(oDoc.getSheets().getElementNames())
    # for nome in ('M1', 'S1', 'S2', 'S4', 'S5', 'Elenco Prezzi', 'COMPUTO'):
    for nome in ('M1', 'S1', 'S2', 'S5', 'Elenco Prezzi', 'COMPUTO'):
        oSheets.remove(nome)
    for nome in oSheets:
        oSheet = oDoc.getSheets().getByName(nome)
        # visualizzazione fogli
        if not oSheet.IsVisible: 
            oDlg_config.getControl('CheckBox2').State = 0
            test = 0
            break
        oDlg_config.getControl('CheckBox2').State = 1
        test = 1
    if oDoc.getSheets().getByName("copyright_LeenO").IsVisible:
        oDlg_config.getControl('CheckBox2').State = 1

    # precisione come mostrato
    if cfg.read('Generale', 'precisione_come_mostrato') == 'True':
        oDoc.CalcAsShown = True
        oDlg_config.getControl('CheckBox4').State = 1  

    if cfg.read('Generale', 'pesca_auto') == '1':
        oDlg_config.getControl('CheckBox1').State = 1  # pesca codice automatico

    # if cfg.read('Generale', 'pesca_auto') == '1':
    #     oDlg_config.getControl('CheckBox1').State = 1  # pesca codice automatico

    if cfg.read('Generale', 'toolbar_contestuali') == '1':
        oDlg_config.getControl('CheckBox6').State = 1

    oSheet = oDoc.getSheets().getByName('S5')
    # descrizione_in_una_colonna
    if not oSheet.getCellRangeByName('C9').IsMerged:
        oDlg_config.getControl('CheckBox5').State = 1
    else:
        oDlg_config.getControl('CheckBox5').State = 0

    sString = oDlg_config.getControl('ComboBox6')
    sString.Text = cfg.read('Generale', 'altezza_celle')

    sString = oDlg_config.getControl("ComboBox2")  # spostamento ad INVIO
    if cfg.read('Generale', 'movedirection') == '1':
        sString.Text = 'A DESTRA'
    elif cfg.read('Generale', 'movedirection') == '0':
        sString.Text = 'IN BASSO'
    oSheet = oDoc.getSheets().getByName('S1')

    # fullscreen
    oLayout = oDoc.CurrentController.getFrame().LayoutManager
    if not oLayout.isElementVisible('private:resource/toolbar/standardbar'):
        oDlg_config.getControl('CheckBox3').State = 1

    sString = oDlg_config.getControl('TextField14')
    sString.Text = oSheet.getCellRangeByName(
        'S1.H334').String  # vedi_voce_breve
    # sString = oDlg_config.getControl('TextField4')
    # sString.Text = oSheet.getCellRangeByName(
        # 'S1.H335').String  # cont_inizio_voci_abbreviate
    if oDoc.NamedRanges.hasByName("_Lib_1"):
        sString.setEnable(False)
    # sString = oDlg_config.getControl('TextField12')
    # sString.Text = oSheet.getCellRangeByName(
        # 'S1.H336').String  # cont_fine_voci_abbreviate
    if oDoc.NamedRanges.hasByName("_Lib_1"):
        sString.setEnable(False)

    if cfg.read('Generale', 'torna_a_ep') == '1':
        oDlg_config.getControl('CheckBox8').State = 1

    sString = oDlg_config.getControl('ComboBox4')
    sString.Text = cfg.read('Generale', 'copie_backup')
    if int(cfg.read('Generale', 'copie_backup')) != 0:
        sString = oDlg_config.getControl('ComboBox5')
        sString.Text = cfg.read('Generale', 'pausa_backup')
    # else:
        # oDlg_config.getControl('ComboBox5').setEnable(False)
        # oDlg_config.execute()
        # DLG.chi(oDlg_config.getControl('ComboBox5'))


    # sString = oDlg_config.getControl('costo_medio_mdo')
    # sString.Text = cfg.read('Computo', 'costo_medio_mdo')

    # sString = oDlg_config.getControl('addetti_mdo')
    # sString.Text = cfg.read('Computo', 'addetti_mdo')

    # MOSTRA IL DIALOGO
    oDlg_config.execute()


    if oDlg_config.getControl('CheckBox2').State != test:
        if oDlg_config.getControl('CheckBox2').State == 1:
            show_sheets(True)
        else:
            show_sheets(False)

    if oDlg_config.getControl('ComboBox1').getText() == 'Chiaro':
        nuove_icone(True)
    elif oDlg_config.getControl('ComboBox1').getText() == 'Scuro':
        nuove_icone(False)

    if oDlg_config.getControl('CheckBox3').State == 1:
        Toolbars.Switch(False)
    else:
        Toolbars.Switch(True)

    if oDlg_config.getControl('CheckBox4').State == 1:
        cfg.write('Generale', 'precisione_come_mostrato', 'True')
        oDoc.CalcAsShown = True
    else:
        cfg.write('Generale', 'precisione_come_mostrato', 'False')
        oDoc.CalcAsShown = False


    ctx = LeenoUtils.getComponentContext()
    oGSheetSettings = ctx.ServiceManager.createInstanceWithContext("com.sun.star.sheet.GlobalSheetSettings", ctx)
    if oDlg_config.getControl('ComboBox2').getText() == 'IN BASSO':
        cfg.write('Generale', 'movedirection', '0')
        oGSheetSettings.MoveDirection = 0
    else:
        cfg.write('Generale', 'movedirection', '1')
        oGSheetSettings.MoveDirection = 1
    cfg.write('Generale', 'altezza_celle', oDlg_config.getControl('ComboBox6').getText())

    cfg.write('Generale', 'pesca_auto', str(oDlg_config.getControl('CheckBox1').State))
    cfg.write('Generale', 'descrizione_in_una_colonna', str(oDlg_config.getControl('CheckBox5').State))
    cfg.write('Generale', 'toolbar_contestuali', str(oDlg_config.getControl('CheckBox6').State))
    Toolbars.Vedi()
    if oDlg_config.getControl('CheckBox5').State == 1:
        descrizione_in_una_colonna(False)
    else:
        descrizione_in_una_colonna(True)
    # torna su prezzario
    cfg.write('Generale', 'torna_a_ep', str(oDlg_config.getControl('CheckBox8').State))

    # il salvataggio anche su leeno.conf serve alla funzione voce_breve()

    if oDlg_config.getControl('TextField14').getText() != '10000':
        cfg.write('Generale', 'vedi_voce_breve', oDlg_config.getControl('TextField14').getText())
    oSheet.getCellRangeByName('S1.H334').Value = float(oDlg_config.getControl('TextField14').getText())

    # if oDlg_config.getControl('TextField4').getText() != '10000':
        # cfg.write('Contabilita', 'cont_inizio_voci_abbreviate', oDlg_config.getControl('TextField4').getText())
    # oSheet.getCellRangeByName('S1.H335').Value = float(oDlg_config.getControl('TextField4').getText())

    # if oDlg_config.getControl('TextField12').getText() != '10000':
        # cfg.write('Contabilita', 'cont_fine_voci_abbreviate', oDlg_config.getControl('TextField12').getText())
    # oSheet.getCellRangeByName('S1.H336').Value = float(oDlg_config.getControl('TextField12').getText())
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)

    cfg.write('Generale', 'copie_backup', oDlg_config.getControl('ComboBox4').getText())
    cfg.write('Generale', 'pausa_backup', oDlg_config.getControl('ComboBox5').getText())

    # cfg.write('Computo', 'costo_medio_mdo', oDlg_config.getControl('costo_medio_mdo').getText())
    # cfg.write('Computo', 'addetti_mdo', oDlg_config.getControl('addetti_mdo').getText())
    autorun()

########################################################################


def LeenO_path():
    '''Restituisce il percorso di installazione di LeenO.oxt'''
    ctx = LeenoUtils.getComponentContext()
    pir = ctx.getValueByName(
        '/singletons/com.sun.star.deployment.PackageInformationProvider')
    expath = pir.getPackageLocation('org.giuseppe-vizziello.leeno')
    return expath


########################################################################

def creaComputo(arg=1):
    '''arg  { integer } : 1 mostra il dialogo di salvataggio file'''
    desktop = LeenoUtils.getDesktop()
    opz = PropertyValue()
    opz.Name = 'AsTemplate'
    opz.Value = True

    if not os.path.exists(LeenO_path() + '/template/leeno/Computo_LeenO.ods'):
        document = desktop.loadComponentFromURL(
            LeenO_path() + '/template/leeno/Computo_LeenO.ods', "_blank", 0,
            (opz, ))
    else:
        document = desktop.loadComponentFromURL(
            LeenO_path() + '/template/leeno/Computo_LeenO.ods', "_blank", 0,
            (opz, ))

    autoexec()
    if arg == 1:
        IconType = "error"
        Title = 'ATTENZIONE!'
        Text='''
Prima di procedere è meglio dare un nome al file.

Lavorando su un file senza nome
potresti avere dei malfunzionamenti.
'''
        Dialogs.NotifyDialog(IconType = IconType, Title = Title, Text = Text)
        DlgMain()
    return document


def creaUsobollo():
    desktop = LeenoUtils.getDesktop()
    opz = PropertyValue()
    opz.Name = 'AsTemplate'
    opz.Value = True
    document = desktop.loadComponentFromURL(
        LeenO_path() + '/template/offmisc/UsoBollo.ott', "_blank", 0,
        (opz, ))
    return document


########################################################################


def MENU_nuovo_computo():
    '''Crea un nuovo computo vuoto.'''
    creaComputo()


########################################################################


def MENU_nuovo_usobollo():
    '''Crea un nuovo documento in formato uso bollo.'''
    creaUsobollo()


########################################################################

def invia_voce_interno():
    '''
    Invia le voci di Elenco Prezzi verso uno degli altri elaborati.
    Richiede comunque la scelta del DP
    '''
    LeenoUtils.DocumentRefresh(False)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    elenco = seleziona()
    codici = [oSheet.getCellByPosition(0, el).String for el in elenco]
    meta = oSheet.getCellRangeByName('C2').String

    if meta == 'VARIANTE':
        genera_variante()
    elif meta == 'CONTABILITA':
        LeenoContab.attiva_contabilita()
        # ins_voce_contab()
    elif meta == 'COMPUTO':
        GotoSheet(meta)
    else:
        LeenoUtils.DocumentRefresh(True)
        Dialogs.Exclamation(Title='AVVISO!',
    Text='''Per procedere devi prima scegliere,
dalla cella "C2", l'elaborato a cui
inviare le voci di prezzo selezionate.

Se l'elaborato è già esistente,
assicurati di aver scelto anche
la posizione di destinazione.''')
        _gotoCella(2, 1)
        return
    oSheet = oDoc.getSheets().getByName(meta)
    for el in codici:
        if oSheet.Name == 'CONTABILITA':
            GotoSheet(meta)
            ins_voce_contab(cod=el)
        else:
            LeenoComputo.ins_voce_computo(cod=el)
        lrow = SheetUtils.getLastUsedRow(oSheet)
    LeenoUtils.DocumentRefresh(True)
    return

###############################################################################
def MENU_invia_voce():
    with LeenoUtils.DocumentRefreshContext(False):
        stato = cfg.read('Generale', 'pesca_auto')
        cfg.write('Generale', 'pesca_auto', 0)

        invia_voce()

        cfg.write('Generale', 'pesca_auto', stato)
    
def invia_voce():
    '''
    Invia le voci di computo, elenco prezzi e analisi, con costi elementari,
    dal documento corrente al Documento Principale.
    '''
    # LeenoUtils.DocumentRefresh(False)

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_cat = LeenoUtils.getGlobalVar('stili_cat')
    stili_contab = LeenoUtils.getGlobalVar('stili_contab')

    nSheet = oSheet.Name
    fpartenza = uno.fileUrlToSystemPath(oDoc.getURL())
    if fpartenza == LeenoUtils.getGlobalVar('sUltimus'):
        if nSheet == 'Elenco Prezzi':
            invia_voce_interno()
            return
        else:
            Dialogs.Exclamation(Title='ATTENZIONE!',
                Text="Questo file coincide con il Documento Principale (DP).")
        return
    elif LeenoUtils.getGlobalVar('sUltimus') == '':
        Dialogs.Exclamation(Title='ATTENZIONE!',
            Text="E' necessario impostare il Documento Principale (DP).")
        return
    nSheetDCC = getDCCSheet()
    # arrivo - Documento Principale
    DP = LeenoUtils.getGlobalVar('sUltimus')
    ddcDoc = LeenoUtils.findOpenDocument(DP)
    lrow = LeggiPosizioneCorrente()[1]


    # DLG.chi(1)

    def getAnalisi(oSheet):
        try:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
        except AttributeError:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
        el_y = []
        try:
            len(oRangeAddress)
            for el in oRangeAddress:
                el_y.append((el.StartRow, el.EndRow))
        except TypeError:
            el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))
        lista = []
        for y in el_y:
            for el in range(y[0], y[1] + 1):
                lista.append(el)
        analisi = []
        for y in lista:
            if oSheet.getCellByPosition(1, y).Type.value == 'FORMULA':
                analisi.append(oSheet.getCellByPosition(0, y).String)
        return (analisi, lista)

    # def Circoscrive_Analisi(lrow):
    #     # oDoc = LeenoUtils.getDocument()
    #     # oSheet = oDoc.CurrentController.ActiveSheet
    #     stili_analisi = LeenoUtils.getGlobalVar('stili_analisi')
    #     if oSheet.getCellByPosition(0, lrow).CellStyle in stili_analisi:
    #         for el in reversed(range(0, lrow)):
    #             if oSheet.getCellByPosition(0,
    #                                         el).CellStyle == 'An.1v-Att Start':
    #                 SR = el
    #                 break
    #         for el in range(lrow, SheetUtils.getUsedArea(oSheet).EndRow):
    #             if oSheet.getCellByPosition(
    #                     0, el).CellStyle == 'Analisi_Sfondo':
    #                 ER = el
    #                 break
    #     celle = oSheet.getCellRangeByPosition(0, SR, 250, ER)
    #     return celle

    def recupera_voce (codice_da_cercare):
        '''
        recupra la voce di prezzo dal foglio di partenza
        e la inserisce nel foglio di destinazione
        '''
        oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
        if SheetUtils.getLastUsedRow(oDoc.getSheets().getByName('COMPUTO')) > 20:
            row = SheetUtils.uFindString(codice_da_cercare, oSheet)[1]
        else:
            row = LeggiPosizioneCorrente()[1]

        range_src = f'A{row+1}:G{row+1}'
        data = oSheet.getCellRangeByName(range_src).FormulaArray

        dccSheet = ddcDoc.getSheets().getByName('Elenco Prezzi')
        dccSheet.getRows().insertByIndex(4, 1)
        range_dest = f'A5:G5'
        dccSheet.getCellRangeByName(range_dest).FormulaArray = data
        ddcDoc.CurrentController.setFirstVisibleRow(4)

    # DLG.chi(1)

    # partenza
    if oSheet.Name == 'Elenco Prezzi':
        voce_da_inviare = oSheet.getCellByPosition(0, lrow).String
        dccSheet = ddcDoc.getSheets().getByName('Elenco Prezzi')
        # verifica presenza codice in EP
        cerca_in_elenco_prezzi = SheetUtils.uFindString(voce_da_inviare, dccSheet)
        if not cerca_in_elenco_prezzi:
            recupera_voce(voce_da_inviare)
        if nSheetDCC in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
            _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
            MENU_nuova_voce_scelta()
            lrow = LeggiPosizioneCorrente()
            dccSheet = ddcDoc.getSheets().getByName(nSheetDCC)
            dccSheet.getCellByPosition(lrow[0], lrow[1]).String = voce_da_inviare
            dccSheet.getCellByPosition(lrow[0]+1, lrow[1]).CellBackColor = 14942166
            _gotoCella(lrow[0]+1, lrow[1]+1)
        return
 
    # partenza
    if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        dv = LeenoComputo.DatiVoce(oSheet, lrow)
        art = dv.art
        ER = dv.ER
        SR = dv.SR

        range_src = f'A{SR+1}:AZ{ER+1}'

        data = oSheet.getCellRangeByName(range_src).FormulaArray

        # oSheet.getCellRangeByPosition(30, SR, 30, ER).CellBackColor = 15757935
        oSheet.getCellByPosition(1, SR +1).CellBackColor = 14942166

        # seleziona()
        if nSheetDCC in ('Analisi di Prezzo'):
            Dialogs.Exclamation(Title = 'ATTENZIONE!',
            Text='Il foglio di destinazione non è corretto.')
            oDoc.CurrentController.select(
                oDoc.createInstance(
                    "com.sun.star.sheet.SheetCellRanges"))  # unselect
            return
        noVoce = LeenoUtils.getGlobalVar('noVoce')
        if nSheetDCC in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
            comando('Copy')
            prossima = LeenoSheetUtils.prossimaVoce(oSheet, ER, 1)
            _gotoCella(0, prossima)
    # arrivo
            # DP = LeenoUtils.getGlobalVar('sUltimus')
            # ddcDoc = LeenoUtils.findOpenDocument(DP)
            dccSheet = ddcDoc.getSheets().getByName(nSheetDCC)
            dccSheet.getCellByPosition(1, SR + 1).CellBackColor = 14942166
            _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
            lrow = LeggiPosizioneCorrente()[1]

            # dccv = LeenoComputo.DatiVoce(dccSheet, lrow)

            if dccSheet.getCellByPosition(0, lrow).CellStyle in (noVoce + stili_computo + stili_contab):
                lrow += 1
            else:
                return
            if dccSheet.getCellByPosition(0, lrow).CellStyle not in stili_computo + stili_contab + stili_cat:
                Dialogs.Exclamation(Title = 'ATTENZIONE!',
                Text='''La posizione di destinazione non è corretta.
    I nomi dei fogli di partenza e di arrivo devo essere coincidenti.''')
                # unselect
                oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
                return
            else:
                try:
                    lrow = LeenoSheetUtils.prossimaVoce(dccSheet, LeggiPosizioneCorrente()[1], 1)
                except Exception as e:
                    DLG.errore(e)

            MENU_nuova_voce_scelta() # inserisce la nuova voce

            _gotoCella(0, lrow + 2) # posiziona sulla prima riga delle misure della voce

            if dccSheet.getCellByPosition(0, LeggiPosizioneCorrente()[1]).CellStyle == 'Comp End Attributo':
                _gotoCella(0, lrow + 1) # posiziona sulla prima riga delle misure della voce

            Copia_riga_Ent(ER - SR - 3) # inserisce le righe per le misure

            _gotoCella(0, lrow) # posiziona sulla riga della voce

            if lrow == 4:
                range_dest = f'A{lrow}:AZ{lrow+ER-SR}'
            else:
                range_dest = f'A{lrow+1}:AZ{lrow+1+ER-SR}'
            
            dccSheet.getCellRangeByName(range_dest).FormulaArray = data

            rigenera_voce(lrow)
            rigenera_parziali(False)
           
            # se nella voce inserita la descrizione non risulta presente
            # la voce di prezzo viene presa dal foglio di partenza
            
            art = dv.art
            ddcSheet = ddcDoc.getSheets().getByName('Elenco Prezzi')
            
            cerca_in_elenco_prezzi = SheetUtils.uFindString(art, ddcSheet)
            # recupera_voce(art)

            if not cerca_in_elenco_prezzi:
                recupera_voce(art)

            Menu_adattaAltezzaRiga()
  
        if nSheetDCC in ('Elenco Prezzi'):
            # DLG.MsgBox("Non è possibile inviare voci da un COMPUTO all'Elenco Prezzi.")
            Dialogs.Exclamation(Title = 'ATTENZIONE!',
            Text="Non è possibile inviare voci da un COMPUTO all'Elenco Prezzi.")
            return
        oDoc.CurrentController.select(
            oDoc.createInstance(
                "com.sun.star.sheet.SheetCellRanges"))  # unselect


    try:
        len(analisi)

        selezione = []
        lista = []
        _gotoDoc(fpartenza)
        oDoc = LeenoUtils.getDocument()
        GotoSheet('Analisi di Prezzo')
        ranges = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")
        ranges.addRangeAddresses(selezione_analisi, True)
        oDoc.CurrentController.select(ranges)

        comando('Copy')

        _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
        inizializza_analisi()
        _gotoCella(0, 0)
        paste_clip(insCells=1)
        tante_analisi_in_ep()
    except Exception:
        pass

    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect
    _gotoDoc(fpartenza)
    GotoSheet(nSheet)
    _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    GotoSheet(nSheetDCC)
    if nSheetDCC in ('COMPUTO', 'VARIANTE'):
        lrow = LeggiPosizioneCorrente()[1]
        if dccSheet.getCellByPosition(0, lrow).CellStyle == 'comp progress':
            ddcDoc.getSheets().getByName(nSheetDCC).getCellByPosition(1, lrow).CellBackColor = 14942166
            _gotoCella(2, lrow)
        else:
            ddcDoc.getSheets().getByName(nSheetDCC).getCellByPosition(1, lrow + 1).CellBackColor = 14942166
            _gotoCella(2, lrow + 1)
    try:
        oSheet = oDoc.getSheets().getByName(nSheetDCC)
        LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    except Exception as e:
        #  DLG.errore(e)
        pass
    # torno su partenza
    if cfg.read('Generale', 'torna_a_ep') == '1':
        _gotoDoc(fpartenza)


########################################################################


def codice_voce(lrow, cod=None):
    '''
    lrow    { int } : id della riga
    cod  { string } : codice del prezzo
    Se cod è assente, restituisce il codice della voce,
    altrimenti glielo assegna.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #  lrow = LeggiPosizioneCorrente()[1]
    if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        try:
            sopra = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
        except:
            Dialogs.Exclamation(Title = 'ATTENZIONE!',
                Text='''La posizione di partenza non è corretta.''')
            return
    elif oSheet.Name in ('Analisi di Prezzo'):
        sopra = Circoscrive_Analisi(lrow).RangeAddress.StartRow + 1
    if cod is None:
        return oSheet.getCellByPosition(1, sopra + 1).String
    else:
        oSheet.getCellByPosition(1, sopra + 1).String = cod

########################################################################


def _gotoDoc(sUrl):
    '''
    sUrl  { string } : nome del file
    porta il focus su di un determinato documento
    '''
    sUrl = uno.systemPathToFileUrl(sUrl)
    if sys.platform == 'win32':
        desktop = LeenoUtils.getDesktop()
        oFocus = uno.createUnoStruct('com.sun.star.awt.FocusEvent')
        target = desktop.loadComponentFromURL(sUrl, "_default", 0, [])
        target.getCurrentController().getFrame().focusGained(oFocus)

    # if sys.platform == 'linux' or sys.platform == 'darwin':
    else:
        target = LeenoUtils.getDesktop().loadComponentFromURL(
            sUrl, "_default", 0, [])
        target.getCurrentController().Frame.ContainerWindow.toFront()
        target.getCurrentController().Frame.activate()
    return target


########################################################################


def getDCCSheet():
    '''
    sUrl  { string } : nome del file
    porta il focus su di un determinato documento
    '''
    oDoc = LeenoUtils.getDocument()
    fpartenza = uno.fileUrlToSystemPath(oDoc.getURL())
    _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
    sUltimus_sheet = LeenoUtils.getDocument().CurrentController.ActiveSheet.Name
    _gotoDoc(fpartenza)
    LeenoUtils.setGlobalVar('sUltimus_sheet', sUltimus_sheet)
    return sUltimus_sheet


########################################################################


def oggi():
    '''
    restituisce la data di oggi
    '''
    return datetime.now().strftime('%d/%m/%Y')


########################################################################


def MENU_copia_sorgente_per_git():
    '''
    Fa una copia della directory del codice nel repository locale ed apre una shell per la commit.
    '''
    oDoc = LeenoUtils.getDocument()
    src_oxt = ''

    try:
        if oDoc.getSheets().getByName('S1').getCellByPosition(7, 338).String == '':
            src_oxt = '_LeenO'
        else:
            src_oxt = oDoc.getSheets().getByName('S1').getCellByPosition(7, 338).String
    except Exception:
        pass

    make_pack(bar=1)

    dest = LeenoGlobals.dest()

    if os.name == 'nt':
        subprocess.Popen(f'w: && cd {dest} && "W:/programmi/PortableGit/git-bash.exe"', shell=True, stdout=subprocess.PIPE)
    else:
        comandi = f'cd {dest} && mate-terminal && gitk &'
        if not processo('wish'):
            subprocess.Popen(comandi, shell=True, stdout=subprocess.PIPE)

    return


########################################################################

def cerca_path_valido():
    if 'giuserpe' in os.getlogin():
        # Try multiple possible paths
        possible_paths = [
            # "C:\\Users\\DELL\\AppData\\Local\\Programs\\cursor\\Cursor.exe",
            os.path.expanduser("~\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"),
            "C:\\Program Files\\Microsoft VS Code\\Code.exe",
            "C:\\Program Files (x86)\\Microsoft VS Code\\Code.exe",
            "C:\\Users\\giuserpe\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
            "C:\\Users\\DELL\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
        ]

        editor_path = None
        for path in possible_paths:
            if os.path.exists(path):
                editor_path = path
                break

        if editor_path is None:
            raise FileNotFoundError("Impossibile trovare VS Code. Assicurati che sia installato.")
        return editor_path

def apri_con_editor(full_file_path, line_number):
    # Imposta il percorso di VSCodium per Windows
    # if os.path.exists("C:\\Users\\giuserpe\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"):
    #     editor_path = "C:\\Users\\giuserpe\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
    # else:
    #     editor_path = "C:\\Users\\DELL\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"

    # if os.path.exists("C:\\Users\\giuserpe\\AppData\\Local\\Programs\\cursor\\Cursor.exe"):
    #     editor_path = "C:\\Users\\giuserpe\\AppData\\Local\\Programs\\cursor\\Cursor.exe"
    # else:
    #     editor_path = "C:\\Users\\DELL\\AppData\\Local\\Programs\\cursor\\Cursor.exe"

    editor_path = cerca_path_valido()

    # Controlla se il file esiste
    if not os.path.exists(full_file_path):
        DLG.chi(f"File non trovato: {full_file_path}")
        return

    # Controlla che il numero di riga sia valido
    if not isinstance(line_number, int) or line_number < 1:
        DLG.chi("Numero di riga non valido. Deve essere un intero maggiore di 0.")
        return

    # Costruisci il comando per aprire il file con VSCodium alla linea specifica
    comando = f'"{editor_path}" --goto "{full_file_path}:{line_number}"'

    # dest = LeenoGlobals.dest()

    # Prova ad aprire il file con VSCodium
    try:
        subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except Exception as e:
        DLG.chi(f"Errore durante l'apertura del file con VSCodium: {e}")


def MENU_avvia_IDE():
    '''
    Avvia la modifica di pyleeno.py con geany
    '''
    avvia_IDE()

def avvia_IDE():
    '''Avvia la modifica di pyleeno.py con geany o VSCodium'''
    oDoc = LeenoUtils.getDocument()
    Toolbars.On("private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV", 1)
    try:
        if oDoc.getSheets().getByName('S1').getCellByPosition(
                7, 338).String == '':
            src_oxt = '_LeenO'
        else:
            src_oxt = oDoc.getSheets().getByName('S1').getCellByPosition(
                7, 338).String
    except Exception:
        pass

    dest = LeenoGlobals.dest()

    # apri_con_editor(f'{dest}/python/pythonpath', 1)
    apri_con_editor(f'{dest}', 1)
    # apri_con_editor(f'{dest}/python/pythonpath/pyleeno.py', 1)
    basic_LeenO('PY_bridge.avvia_IDE')
    return


########################################################################


def Ins_Categorie(n):
    '''
    n    { int } : livello della categoria
    0 = SuperCategoria
    1 = Categoria
    2 = SubCategoria
    '''
    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        # LeenoUtils.DocumentRefresh(False)
        oSheet = oDoc.CurrentController.ActiveSheet

        stili_computo = LeenoUtils.getGlobalVar('stili_computo')
        stili_contab = LeenoUtils.getGlobalVar('stili_contab')
        noVoce = LeenoUtils.getGlobalVar('noVoce')

        row = LeggiPosizioneCorrente()[1]
        if oSheet.getCellByPosition(0,row).CellStyle in stili_computo + stili_contab:
            lrow = LeenoSheetUtils.prossimaVoce(oSheet, row, 1)
        elif oSheet.getCellByPosition(0, row).CellStyle in noVoce:
            lrow = row + 1
        else:
            # LeenoUtils.DocumentRefresh(True)
            return
        sTesto = ''
        if n == 0:
            sTesto = 'Inserisci il titolo per la Supercategoria'
        elif n == 1:
            sTesto = 'Inserisci il titolo per la Categoria'
        elif n == 2:
            sTesto = 'Inserisci il titolo per la Sottocategoria'
        sString = InputBox('', sTesto)
        if sString is None or sString == '':
            # LeenoUtils.DocumentRefresh(True)
            return

        if n == 0:
            LeenoSheetUtils.inserSuperCapitolo(oSheet, lrow, sString)
        elif n == 1:
            LeenoSheetUtils.inserCapitolo(oSheet, lrow, sString)
        elif n == 2:
            LeenoSheetUtils.inserSottoCapitolo(oSheet, lrow, sString)

        _gotoCella(2, lrow)
        Rinumera_TUTTI_Capitoli2(oSheet)
        oDoc.enableAutomaticCalculation(True)
        oDoc.CurrentController.setFirstVisibleColumn(0)
        oDoc.CurrentController.setFirstVisibleRow(lrow - 5)
        # LeenoUtils.DocumentRefresh(True)
        LeenoSheetUtils.adattaAltezzaRiga()


########################################################################


def MENU_Inser_SuperCapitolo():
    '''
    @@ DA DOCUMENTARE
    '''
#     Inser_SuperCapitolo()


# def Inser_SuperCapitolo():
#     '''
#     @@ DA DOCUMENTARE
    # '''
    Ins_Categorie(0)

########################################################################

def MENU_Inser_Capitolo():
    '''
    @@ DA DOCUMENTARE
    '''
#     Inser_Capitolo()


# def Inser_Capitolo():
#     '''
#     @@ DA DOCUMENTARE
    # '''
    Ins_Categorie(1)


########################################################################


def MENU_Inser_SottoCapitolo():
    '''
    @@ DA DOCUMENTARE
    '''
#     Inser_SottoCapitolo()


# def Inser_SottoCapitolo():
#     '''
#     @@ DA DOCUMENTARE
#     '''
    Ins_Categorie(2)

########################################################################


def numera_voci():
    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        Rinumera_TUTTI_Capitoli2(oSheet)


def Rinumera_TUTTI_Capitoli2(oSheet):
    # sistemo gli idcat voce per voce
    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(False)
    Sincronizza_SottoCap_Tag_Capitolo_Cor(oSheet)

    # ricalcola i totali di categorie e subcategorie
    Tutti_Subtotali(oSheet)
    oDoc.enableAutomaticCalculation(True)


def Tutti_Subtotali(oSheet):
    '''ricalcola i subtotali di categorie e subcategorie'''

    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
    for n in range(0, LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1):
        if oSheet.getCellByPosition(0, n).CellStyle == 'Livello-0-scritta':
            SubSum_SuperCap(n)
        if oSheet.getCellByPosition(0, n).CellStyle == 'Livello-1-scritta':
            SubSum_Cap(n)
        if oSheet.getCellByPosition(0, n).CellStyle == 'livello2 valuta':
            SubSum_SottoCap(n)

    # TOTALI GENERALI
    lrow = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1
    oSheet.getCellByPosition(
        17, 1).Formula = '=SUBTOTAL(9;R4:R' + str(lrow + 1) + ')'
    oSheet.getCellByPosition(
        17, lrow).Formula = '=SUBTOTAL(9;R4:R' + str(lrow + 1) + ')'

    if oSheet.Name != 'CONTABILITA':
        oSheet.getCellByPosition(
            18, 1).Formula = '=SUBTOTAL(9;S3:S' + str(lrow + 1) + ')'

    oSheet.getCellByPosition(
        18, lrow).Formula = '=SUBTOTAL(9;S3:S' + str(lrow + 1) + ')'

    oSheet.getCellByPosition(
        28, lrow).Formula = '=SUBTOTAL(9;AC4:AC' + str(lrow + 1) + ')'
    oSheet.getCellByPosition(
        28, 1).Formula = '=SUBTOTAL(9;AC4:AC' + str(lrow + 1) + ')'

    oSheet.getCellByPosition(
        30, lrow).Formula = '=SUBTOTAL(9;AE4:AE' + str(lrow + 1) + ')'
    oSheet.getCellByPosition(
        30, 1).Formula = '=SUBTOTAL(9;AE4:AE' + str(lrow + 1) + ')'
    oSheet.getCellByPosition(
        36, lrow).Formula = '=SUBTOTAL(9;AK4:AK' + str(lrow + 1) + ')'
    oSheet.getCellByPosition(
        36, 1).Formula = '=SUBTOTAL(9;AK4:AK' + str(lrow + 1) + ')'


########################################################################


def SubSum_SuperCap(lrow):
    '''
    lrow    { double } : id della riga di inserimento
    inserisce i dati nella riga di SuperCategoria
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
    # lrow = LeggiPosizioneCorrente()[1]
    lrowE = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2
    nextCap = lrowE
    for n in range(lrow + 1, lrowE):
        if oSheet.getCellByPosition(
                18,
                n).CellStyle in ('Livello-0-scritta mini val', 'Comp TOTALI'):
            # MsgBox(oSheet.getCellByPosition(18, n).CellStyle,'')
            nextCap = n + 1
            break
    # oDoc.enableAutomaticCalculation(False)
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        oSheet.getCellByPosition(
            18, lrow
        ).Formula = '=SUBTOTAL(9;S' + str(lrow + 1) + ':S' + str(nextCap) + ')'
        oSheet.getCellByPosition(
            24, lrow).Formula = '=S' + str(lrow + 1) + '/S' + str(lrowE)
        oSheet.getCellByPosition(28, lrow).Formula = '=SUBTOTAL(9;AC' + str(
            lrow + 1) + ':AC' + str(nextCap) + ')'
        oSheet.getCellByPosition(
            29, lrow).Formula = '=AE' + str(lrow + 1) + '/S' + str(lrow + 1)
        oSheet.getCellByPosition(30, lrow).Formula = '=SUBTOTAL(9;AE' + str(
            lrow + 1) + ':AE' + str(nextCap) + ')'
        oSheet.getCellByPosition(18,
                                 lrow).CellStyle = 'Livello-0-scritta mini val'
        oSheet.getCellByPosition(24,
                                 lrow).CellStyle = 'Livello-0-scritta mini %'
        oSheet.getCellByPosition(29,
                                 lrow).CellStyle = 'Livello-0-scritta mini %'
        oSheet.getCellByPosition(30,
                                 lrow).CellStyle = 'Livello-0-scritta mini val'
    if oSheet.Name in ('CONTABILITA'):
        oSheet.getCellByPosition(15, lrow).Formula = '=SUBTOTAL(9;P' + str(
            lrow + 1) + ':P' + str(nextCap) + ')'  # IMPORTO
        oSheet.getCellByPosition(
            16, lrow).Formula = '=P' + str(lrow + 1) + '/P' + str(
                lrowE)  # incidenza sul totale
        oSheet.getCellByPosition(28, lrow).Formula = '=SUBTOTAL(9;AC' + str(
            lrow + 1) + ':AC' + str(nextCap) + ')'
        oSheet.getCellByPosition(
            29, lrow).Formula = '=AE' + str(lrow + 1) + '/P' + str(lrow + 1)
        oSheet.getCellByPosition(30, lrow).Formula = '=SUBTOTAL(9;AE' + str(
            lrow + 1) + ':AE' + str(nextCap) + ')'
        oSheet.getCellByPosition(15,
                                 lrow).CellStyle = 'Livello-0-scritta mini val'
        oSheet.getCellByPosition(16,
                                 lrow).CellStyle = 'Livello-0-scritta mini %'
        oSheet.getCellByPosition(29,
                                 lrow).CellStyle = 'Livello-0-scritta mini %'
        oSheet.getCellByPosition(28,
                                 lrow).CellStyle = 'Livello-0-scritta mini val'
        oSheet.getCellByPosition(30,
                                 lrow).CellStyle = 'Livello-0-scritta mini val'
    # oDoc.enableAutomaticCalculation(True)


########################################################################


def SubSum_SottoCap(lrow):
    '''
    lrow    { double } : id della riga di inserimento
    inserisce i dati nella riga di subcategoria
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
    # lrow = 0#LeggiPosizioneCorrente()[1]
    lrowE = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2
    nextCap = lrowE
    for n in range(lrow + 1, lrowE):
        if oSheet.getCellByPosition(
                18,
                n).CellStyle in ('livello2 scritta mini',
                                 'Livello-0-scritta mini val',
                                 'Livello-1-scritta mini val', 'Comp TOTALI'):
            nextCap = n + 1
            break
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        oSheet.getCellByPosition(
            18, lrow
        ).Formula = '=SUBTOTAL(9;S' + str(lrow + 1) + ':S' + str(nextCap) + ')'
        oSheet.getCellByPosition(
            24, lrow).Formula = '=S' + str(lrow + 1) + '/S' + str(lrowE)
        oSheet.getCellByPosition(28, lrow).Formula = '=SUBTOTAL(9;AC' + str(
            lrow + 1) + ':AC' + str(nextCap) + ')'
        oSheet.getCellByPosition(
            29, lrow).Formula = '=AE' + str(lrow + 1) + '/S' + str(lrow + 1)
        oSheet.getCellByPosition(30, lrow).Formula = '=SUBTOTAL(9;AE' + str(
            lrow + 1) + ':AE' + str(nextCap) + ')'
        oSheet.getCellByPosition(18, lrow).CellStyle = 'livello2 scritta mini'
        oSheet.getCellByPosition(24, lrow).CellStyle = 'livello2 valuta mini %'
        oSheet.getCellByPosition(28, lrow).CellStyle = 'livello2 scritta mini'
        oSheet.getCellByPosition(29, lrow).CellStyle = 'livello2 valuta mini %'
        oSheet.getCellByPosition(30, lrow).CellStyle = 'livello2 valuta mini'
    if oSheet.Name in ('CONTABILITA'):
        oSheet.getCellByPosition(15, lrow).Formula = '=SUBTOTAL(9;P' + str(
            lrow + 1) + ':P' + str(nextCap) + ')'  # IMPORTO
        oSheet.getCellByPosition(
            16, lrow).Formula = '=P' + str(lrow + 1) + '/P' + str(
                lrowE)  # incidenza sul totale
        oSheet.getCellByPosition(28, lrow).Formula = '=SUBTOTAL(9;AC' + str(
            lrow + 1) + ':AC' + str(nextCap) + ')'
        oSheet.getCellByPosition(
            29, lrow).Formula = '=AE' + str(lrow + 1) + '/P' + str(lrow + 1)
        oSheet.getCellByPosition(30, lrow).Formula = '=SUBTOTAL(9;AE' + str(
            lrow + 1) + ':AE' + str(nextCap) + ')'
        oSheet.getCellByPosition(15, lrow).CellStyle = 'livello2 scritta mini'
        oSheet.getCellByPosition(16, lrow).CellStyle = 'livello2 valuta mini %'
        oSheet.getCellByPosition(29, lrow).CellStyle = 'livello2 valuta mini %'
        oSheet.getCellByPosition(28, lrow).CellStyle = 'livello2 scritta mini'
        oSheet.getCellByPosition(30, lrow).CellStyle = 'livello2 scritta mini'


########################################################################


def SubSum_Cap(lrow):
    '''
    lrow    { double } : id della riga di inserimento
    inserisce i dati nella riga di categoria
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
    # lrow = LeggiPosizioneCorrente()[1]
    lrowE = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2
    nextCap = lrowE
    for n in range(lrow + 1, lrowE):
        if oSheet.getCellByPosition(
                18,
                n).CellStyle in ('Livello-1-scritta mini val',
                                 'Livello-0-scritta mini val', 'Comp TOTALI'):
            # MsgBox(oSheet.getCellByPosition(18, n).CellStyle,'')
            nextCap = n + 1
            break
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        oSheet.getCellByPosition(18, lrow).Formula = '=SUBTOTAL(9;S' + str(
            lrow + 1) + ':S' + str(nextCap) + ')'  # IMPORTO
        oSheet.getCellByPosition(
            24, lrow).Formula = '=S' + str(lrow + 1) + '/S' + str(lrowE)
        oSheet.getCellByPosition(28, lrow).Formula = '=SUBTOTAL(9;AC' + str(
            lrow + 1) + ':AC' + str(nextCap) + ')'
        oSheet.getCellByPosition(
            29, lrow).Formula = '=AE' + str(lrow + 1) + '/S' + str(lrow + 1)
        oSheet.getCellByPosition(30, lrow).Formula = '=SUBTOTAL(9;AE' + str(
            lrow + 1) + ':AE' + str(nextCap) + ')'
        oSheet.getCellByPosition(18,
                                 lrow).CellStyle = 'Livello-1-scritta mini val'
        oSheet.getCellByPosition(24,
                                 lrow).CellStyle = 'Livello-1-scritta mini %'
        oSheet.getCellByPosition(29,
                                 lrow).CellStyle = 'Livello-1-scritta mini %'
        oSheet.getCellByPosition(30,
                                 lrow).CellStyle = 'Livello-1-scritta mini val'
    if oSheet.Name in ('CONTABILITA'):
        oSheet.getCellByPosition(15, lrow).Formula = '=SUBTOTAL(9;P' + str(
            lrow + 1) + ':P' + str(nextCap) + ')'  # IMPORTO
        oSheet.getCellByPosition(
            16, lrow).Formula = '=P' + str(lrow + 1) + '/P' + str(
                lrowE)  # incidenza sul totale
        oSheet.getCellByPosition(28, lrow).Formula = '=SUBTOTAL(9;AC' + str(
            lrow + 1) + ':AC' + str(nextCap) + ')'
        oSheet.getCellByPosition(
            29, lrow).Formula = '=AE' + str(lrow + 1) + '/P' + str(lrow + 1)
        oSheet.getCellByPosition(30, lrow).Formula = '=SUBTOTAL(9;AE' + str(
            lrow + 1) + ':AE' + str(nextCap) + ')'
        oSheet.getCellByPosition(15,
                                 lrow).CellStyle = 'Livello-1-scritta mini val'
        oSheet.getCellByPosition(16,
                                 lrow).CellStyle = 'Livello-1-scritta mini %'
        oSheet.getCellByPosition(29,
                                 lrow).CellStyle = 'Livello-1-scritta mini %'
        oSheet.getCellByPosition(28,
                                 lrow).CellStyle = 'Livello-1-scritta mini val'
        oSheet.getCellByPosition(30,
                                 lrow).CellStyle = 'Livello-1-scritta mini val'


########################################################################


def Sincronizza_SottoCap_Tag_Capitolo_Cor(oSheet):
    '''
    lrow    { double } : id della riga di inserimento
    sincronizza categoria e sottocategorie
    '''
    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
    
    # Ottieni l'indicatore di stato
    oDoc = LeenoUtils.getDocument()

    
    lastRow = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1

    listasbcat = []
    listacat = []
    listaspcat = []
    
    # Inizializza la progressbar
    indicator = oDoc.getCurrentController().getStatusIndicator()
    if indicator is not None:
        indicator.start("Sincronizzazione categorie...", lastRow)
    
    try:
        for lrow in range(0, lastRow):
            # Aggiorna la progressbar
            if indicator is not None:
                indicator.setValue(lrow)
            
            # SUB CATEGORIA
            if oSheet.getCellByPosition(2, lrow).CellStyle == 'livello2_':
                if oSheet.getCellByPosition(2, lrow).String not in listasbcat:
                    listasbcat.append((oSheet.getCellByPosition(2, lrow).String))
                try:
                    oSheet.getCellByPosition(31, lrow).Value = idspcat
                except Exception:
                    pass
                try:
                    oSheet.getCellByPosition(32, lrow).Value = idcat
                except Exception:
                    pass
                idsbcat = listasbcat.index(oSheet.getCellByPosition(2, lrow).String) + 1
                oSheet.getCellByPosition(33, lrow).Value = idsbcat
                oSheet.getCellByPosition(
                    1, lrow).Formula = '=AF' + str(lrow + 1) + '&"."&AG' + str(
                        lrow + 1) + '&"."&AH' + str(lrow + 1)

            # CATEGORIA
            elif oSheet.getCellByPosition(2, lrow).CellStyle == 'Livello-1-scritta mini':
                if oSheet.getCellByPosition(2, lrow).String not in listacat:
                    listacat.append((oSheet.getCellByPosition(2, lrow).String))
                    idsbcat = None

                try:
                    oSheet.getCellByPosition(31, lrow).Value = idspcat
                except Exception:
                    pass
                idcat = listacat.index(oSheet.getCellByPosition(2,
                                                                lrow).String) + 1
                oSheet.getCellByPosition(32, lrow).Value = idcat
                oSheet.getCellByPosition(
                    1, lrow).Formula = '=AF' + str(lrow +
                                                   1) + '&"."&AG' + str(lrow + 1)

            # SUPER CATEGORIA
            elif oSheet.getCellByPosition(
                    2,
                    lrow).CellStyle == 'Livello-0-scritta mini':
                if oSheet.getCellByPosition(2, lrow).String not in listaspcat:
                    listaspcat.append((oSheet.getCellByPosition(2, lrow).String))
                    idcat = idsbcat = None

                idspcat = listaspcat.index(
                    oSheet.getCellByPosition(2, lrow).String) + 1
                oSheet.getCellByPosition(31, lrow).Value = idspcat
                oSheet.getCellByPosition(1, lrow).Formula = '=AF' + str(lrow + 1)

            # CATEGORIA
            elif oSheet.getCellByPosition(
                    33, lrow).CellStyle == 'compTagRiservato':
                try:
                    oSheet.getCellByPosition(33, lrow).Value = idsbcat
                except Exception:
                    oSheet.getCellByPosition(33, lrow).Value = 0
                try:
                    oSheet.getCellByPosition(32, lrow).Value = idcat
                except Exception:
                    oSheet.getCellByPosition(32, lrow).Value = 0
                try:
                    oSheet.getCellByPosition(31, lrow).Value = idspcat
                except Exception:
                    oSheet.getCellByPosition(31, lrow).Value = 0
                    
    finally:
        # Assicurati che la progressbar venga chiusa anche in caso di eccezioni
        if indicator is not None:
            indicator.end()

########################################################################
# MENU_unisci_fogli moved to SheetUtils.py
########################################################################


def MENU_mostra_fogli():
    '''Mostra tutti i foglio fogli'''
    oDoc = LeenoUtils.getDocument()
    lista_fogli = oDoc.Sheets.ElementNames
    for el in lista_fogli:
        oDoc.getSheets().getByName(el).IsVisible = True


########################################################################


def MENU_mostra_fogli_principali():
    '''
    Mostra tutti i foglio fogli
    '''
    mostra_fogli_principali()


def mostra_fogli_principali():
    '''Mostra tutti i foglio fogli'''
    oDoc = LeenoUtils.getDocument()
    lista_fogli = oDoc.Sheets.ElementNames
    for el in lista_fogli:
        oDoc.getSheets().getByName(el).IsVisible = True
        for nome in ('cP_', 'cT_', 'M1', 'S1', 'S2', 'S5', 'QUADRO ECONOMICO',
                     '_LeenO', 'Scorciatoie'):
            if nome in el:
                oDoc.getSheets().getByName(el).IsVisible = False


########################################################################


def MENU_mostra_tabs_contab():
    '''Mostra tutti i foglio fogli'''
    oDoc = LeenoUtils.getDocument()
    lista_fogli = oDoc.Sheets.ElementNames
    sproteggi_sheet_TUTTE()
    for el in lista_fogli:
        oDoc.getSheets().getByName(el).IsVisible = True
        for nome in ('cP_', 'M1', 'S1', 'S2', 'S5', 'QUADRO ECONOMICO',
                     '_LeenO', 'Scorciatoie'):
            if nome in el:
                oDoc.getSheets().getByName(el).IsVisible = False


########################################################################


def MENU_mostra_tabs_computo():
    '''Mostra tutti i foglio fogli'''
    oDoc = LeenoUtils.getDocument()
    lista_fogli = oDoc.Sheets.ElementNames
    sproteggi_sheet_TUTTE()
    for el in lista_fogli:
        oDoc.getSheets().getByName(el).IsVisible = True
        for nome in ('cT_', 'M1', 'S1', 'S2', 'S5', 'QUADRO ECONOMICO',
                     '_LeenO', 'Scorciatoie'):
            if nome in el:
                oDoc.getSheets().getByName(el).IsVisible = False


########################################################################


def copia_sheet(nSheet, tag="_copia"):
    '''
    nSheet   { string } : nome sheet
    tag      { string } : stringa di tag
    duplica copia sheet corrente di fianco a destra
    '''
    oDoc = LeenoUtils.getDocument()
    # nSheet = 'COMPUTO'
    oSheet = oDoc.getSheets().getByName(nSheet)
    idSheet = oSheet.RangeAddress.Sheet + 1
    if oDoc.getSheets().hasByName(nSheet + tag):
        DLG.MsgBox(f'La tabella di nome {nSheet} {tag} è già presente.', 'ATTENZIONE! Impossibile procedere.')
        return
    else:
        oDoc.Sheets.copyByName(nSheet, nSheet + tag, idSheet)
        oSheet = oDoc.getSheets().getByName(nSheet + tag)
        oDoc.CurrentController.setActiveSheet(oSheet)
        # oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect


########################################################################
def copia_sheet_consolida():
    '''Copia il foglio corrente e ne consolida il contenuto'''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    copia_sheet(oSheet.Name)

    comando("SelectAll")
    comando("Copy")
    paste_clip(insCells=0, pastevalue=True)
    LeenoUtils.DocumentRefresh(True)
    return
########################################################################


def Filtra_computo(nSheet, nCol, sString):
    # "SERVE?"
    '''
    nSheet   { string } : nome Sheet
    ncol     { integer } : colonna di tag
    sString  { string } : stringa di tag
    crea una nuova sheet contenente le sole voci filtrate
    '''
    oDoc = LeenoUtils.getDocument()
    copia_sheet(nSheet, sString)
    oSheet = oDoc.CurrentController.ActiveSheet
    for lrow in reversed(range(0, LeenoSheetUtils.cercaUltimaVoce(oSheet))):
        try:
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
            sopra = sStRange.RangeAddress.StartRow
            sotto = sStRange.RangeAddress.EndRow
            if nCol == 1:
                test = sopra + 1
            else:
                test = sotto
            if sString != oSheet.getCellByPosition(nCol, test).String:
                oSheet.getRows().removeByIndex(sopra, sotto - sopra + 1)
                lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 0)
        except Exception:
            lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 0)
    for lrow in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
        if(oSheet.getCellByPosition(18, lrow).CellStyle == 'Livello-1-scritta mini val' and
           oSheet.getCellByPosition(18, lrow).Value == 0 or
           oSheet.getCellByPosition(18, lrow).CellStyle == 'livello2 scritta mini' and
           oSheet.getCellByPosition(18, lrow).Value == 0):

            oSheet.getRows().removeByIndex(lrow, 1)
    return
    # iCellAttr =(oDoc.createInstance("com.sun.star.sheet.CellFlags.OBJECTS"))
    flags = OBJECTS
    oSheet.getCellRangeByPosition(0, 0, 42, 0).clearContents(
        flags)  # cancello gli oggetti
    oDoc.CurrentController.select(oSheet.getCellByPosition(0, 3))
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect


########################################################################


def vai_a_M1():
    chiudi_dialoghi()
    GotoSheet('M1', 85)
    _primaCella(0, 0)


########################################################################


def vai_a_S2():
    chiudi_dialoghi()
    GotoSheet('S2')
    _primaCella(0, 0)


########################################################################


def vai_a_S1():
    chiudi_dialoghi()
    GotoSheet('S1')
    _primaCella(0, 190)


########################################################################


def vai_a_ElencoPrezzi(event=None):
    chiudi_dialoghi()
    GotoSheet('Elenco Prezzi')


########################################################################


def vai_a_Computo():
    chiudi_dialoghi()
    GotoSheet('COMPUTO')


########################################################################


def vai_a_variabili():
    chiudi_dialoghi()
    GotoSheet('S1', 85)
    _primaCella(6, 289)


########################################################################


def vai_a_Scorciatoie():
    chiudi_dialoghi()
    GotoSheet('Scorciatoie')
    _primaCella(0, 0)


########################################################################


def GotoSheet(nSheet, fattore=100):
    '''
    nSheet   { string } : nome Sheet
    attiva e seleziona una sheet
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.Sheets.getByName(nSheet)
    # oDoc.getCurrentSelection().getCellAddress().Sheet

    oSheet.IsVisible = True
    oDoc.CurrentController.setActiveSheet(oSheet)
    # oDoc.CurrentController.ZoomValue = fattore
    # oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect


########################################################################


def _primaCella(IDcol=0, IDrow=0):
    '''
    IDcol   { integer } : id colonna
    IDrow   { integer } : id riga
    settaggio prima cella visibile(IDcol, IDrow)
    '''
    oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet
    oDoc.CurrentController.setFirstVisibleColumn(IDcol)
    oDoc.CurrentController.setFirstVisibleRow(IDrow)
    return


########################################################################


def ordina_col(ncol):
    '''
    ncol   { integer } : id colonna
    ordina i dati secondo la colonna con id ncol
    '''

    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
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

def MENU_sproteggi_sheet_TUTTE():
    '''
    Sprotegge e riordina tutti fogli del documento.
    '''
    sproteggi_sheet_TUTTE()


def sproteggi_sheet_TUTTE():
    '''
    Sprotegge e riordina tutti i fogli del documento.
    '''
    # Ottieni il documento corrente
    oDoc = LeenoUtils.getDocument()

    # Sproteggi tutti i fogli del documento
    for nome in oDoc.Sheets.ElementNames:
        oSheet = oDoc.getSheets().getByName(nome)
        oSheet.unprotect('')  # Rimuovi la protezione con una password vuota

    # Specifica l'ordine desiderato dei fogli
    ordine_fogli = ["Analisi di Prezzo", "Elenco Prezzi", "COMPUTO", "VARIANTE", "CONTABILITA", "M1", "S1", "S2", "S5", "copyright_LeenO"]

    # Sposta i fogli nell'ordine specificato
    for posizione, nome_foglio in enumerate(ordine_fogli):
        if oDoc.Sheets.hasByName(nome_foglio):
            oDoc.Sheets.moveByName(nome_foglio, posizione)


########################################################################


def setPreview():
    '''
    Attiva/disattiva il preview
    '''
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    oProp = PropertyValue()
    properties = (oProp, )
    dispatchHelper.executeDispatch(oFrame, '.uno:PrintPreview', '', 0, properties)
    return


########################################################################


def setTabColor(colore):
    '''
    colore   { integer } : id colore
    attribuisce al foglio corrente un colore a scelta
    '''
    # oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    oProp = PropertyValue()
    oProp.Name = 'TabBgColor'
    oProp.Value = colore
    properties = (oProp, )
    dispatchHelper.executeDispatch(oFrame, '.uno:SetTabBgColor', '', 0,
                                   properties)


########################################################################


def txt_Format(stile):
    '''
    Forza la formattazione della cella
    '''
    # oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    oProp = PropertyValue()
    oProp.Name = stile
    oProp.Value = True
    properties = (oProp, )
    dispatchHelper.executeDispatch(oFrame, '.uno:' + stile, '', 0, properties)


########################################################################


def show_sheets(x=True):
    '''
    x   { boolean } : True = ON, False = OFF

    Mastra/nasconde tutte le tabelle ad esclusione di COMPUTO ed Elenco Prezzi
    '''
    oDoc = LeenoUtils.getDocument()
    oSheets = list(oDoc.getSheets().getElementNames())
    # for nome in ('M1', 'S1', 'S2', 'S5', 'Elenco Prezzi', 'COMPUTO'):
    for nome in ('Elenco Prezzi', 'COMPUTO'):
        oSheets.remove(nome)
    # oSheets.remove('Elenco Prezzi')
    # oSheets.remove('COMPUTO')
    for nome in oSheets:
        oSheet = oDoc.getSheets().getByName(nome)
        oSheet.IsVisible = x
    for nome in ('COMPUTO', 'VARIANTE', 'Elenco Prezzi', 'CONTABILITA',
                 'Analisi di Prezzo'):
        try:
            oSheet = oDoc.getSheets().getByName(nome)
            oSheet.IsVisible = True
        except Exception:
            pass


def nascondi_sheets():
    show_sheets(False)


########################################################################


def salva_come(nomefile=None):
    '''
    nomefile   { string } : nome del file di destinazione
    Se presente l'argomento nomefile, salva il file corrente in nomefile.
    '''
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)

    oProp = []
    if nomefile is not None:
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

def sposta_cursore(destra = 1, basso = 1):

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # posizione corrente
    IDcol = LeggiPosizioneCorrente()[0]
    IDrow = LeggiPosizioneCorrente()[1]

    # nuova posizione
    new_col = IDcol + destra
    new_row = IDrow + basso

    # nuova cella
    oDoc.CurrentController.select(oSheet.getCellByPosition(new_col, new_row))
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
    return


########################################################################


def _gotoCella(IDcol=0, IDrow=0):
    '''
    IDcol   { integer } : id colonna
    IDrow   { integer } : id riga

    muove il cursore nelle cella(IDcol, IDrow)
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    oDoc.CurrentController.select(oSheet.getCellByPosition(IDcol, IDrow))
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
    return


########################################################################


def loVersion():
    '''
    Legge il numero di versione di LibreOffice.
    '''
    # sAccess = LeenoUtils.createUnoService(
    #    "com.sun.star.configuration.ConfigurationAccess")
    aConfigProvider = LeenoUtils.createUnoService(
        "com.sun.star.configuration.ConfigurationProvider")
    arg = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    arg.Name = "nodepath"
    arg.Value = '/org.openoffice.Setup/Product'
    return aConfigProvider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationAccess",
        (arg, )).ooSetupVersionAboutBox

########################################################################

def Menu_adattaAltezzaRiga():
    # oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet
    LeenoSheetUtils.adattaAltezzaRiga()

########################################################################

def MENU_voce_breve():
    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        if oSheet.Name in ('Analisi di Prezzo'):
            voce_breve_an()
        elif oSheet.Name in ('Elenco Prezzi'):
            voce_breve_ep()
        elif oSheet.Name in ("COMPUTO", "VARIANTE", "CONTABILITA"):
            voce_breve()

########################################################################


def voce_breve():
    '''
    Cambia il numero di caratteri visualizzati per la descrizione voce in COMPUTO,
    CONTABILITA E VARIANTE.
    '''
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet
    # Definizione dei fogli di lavoro da modificare
    # nome_foglio = oDoc.CurrentController.ActiveSheet.getName()

    if oDoc.NamedRanges.hasByName("_Lib_1"):
        Dialogs.Exclamation(
            Title='ATTENZIONE!', Text="Risulta già registrato un SAL.\n\nIl foglio CONTABILITA sarà ignorato.")
        fogli_lavoro = ['COMPUTO']
    else:
        fogli_lavoro = ['COMPUTO', 'CONTABILITA']
        oDoc.getSheets().getByName('S1').getCellRangeByName(
            'S1.H335').Value = oDoc.getSheets().getByName('S1').getCellRangeByName('S1.H337').Value
        oDoc.getSheets().getByName('S1').getCellRangeByName(
            'S1.H336').Value = oDoc.getSheets().getByName('S1').getCellRangeByName('S1.H338').Value

    oSheet = oDoc.getSheets().getByName('S1')
    for nome_foglio in fogli_lavoro:
        if nome_foglio == 'CONTABILITA':
            if oSheet.getCellRangeByName('S1.H335').Value < 10000:
                cfg.write('Contabilita', 'cont_inizio_voci_abbreviate',
                          oSheet.getCellRangeByName('S1.H335').String)
                oSheet.getCellRangeByName('S1.H335').Value = 10000
            else:
                oSheet.getCellRangeByName('S1.H335').Value = int(
                    cfg.read('Contabilita', 'cont_inizio_voci_abbreviate'))
            if oSheet.getCellRangeByName('S1.H336').Value < 10000:
                cfg.write('Contabilita', 'cont_fine_voci_abbreviate',
                          oSheet.getCellRangeByName('S1.H336').String)
                oSheet.getCellRangeByName('S1.H336').Value = 10000
            else:
                oSheet.getCellRangeByName('S1.H336').Value = int(
                    cfg.read('Contabilita', 'cont_fine_voci_abbreviate'))

        else:
            if oSheet.getCellRangeByName('S1.H337').Value < 10000:
                cfg.write('Computo', 'inizio_voci_abbreviate',
                          oSheet.getCellRangeByName('S1.H337').String)
                oSheet.getCellRangeByName('S1.H337').Value = 10000
            else:
                oSheet.getCellRangeByName('S1.H337').Value = int(
                    cfg.read('Computo', 'inizio_voci_abbreviate'))
            if oSheet.getCellRangeByName('S1.H338').Value < 10000:
                cfg.write('Computo', 'fine_voci_abbreviate',
                          oSheet.getCellRangeByName('S1.H338').String)
                oSheet.getCellRangeByName('S1.H338').Value = 10000
            else:
                oSheet.getCellRangeByName('S1.H338').Value = int(
                    cfg.read('Computo', 'fine_voci_abbreviate'))

    for el in ("COMPUTO", "VARIANTE", "CONTABILITA", "Elenco Prezzi", "Analisi di Prezzo"):
        try:
            LeenoSheetUtils.adattaAltezzaRiga(oDoc.getSheets().getByName(el))
            pass
        except:
            pass

########################################################################
def MENU_prefisso_VDS_():
    '''
    Duplica la voce di Elenco Prezzi corrente aggiunge il prefisso 'VDS_'
    e individuandola come Voce Della Sicurezza
    '''
    oDoc = LeenoUtils.getDocument()

    pref = "VDS_"
    # pref = "NP_"

    def vds_ep():
        oSheet = oDoc.CurrentController.ActiveSheet
        lrow = LeggiPosizioneCorrente()[1]
        if pref in  oSheet.getCellByPosition(0, lrow).String:
            #  Dialogs.Info(Title = 'Infomazione', Text = 'Voce della sicurezza già esistente')
            LeenoUtils.DocumentRefresh(True)
            return
        oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, lrow, 9, lrow))
        comando('Copy')
        MENU_nuova_voce_scelta()
        paste_clip(pastevalue = False)
        oSheet.getCellRangeByName("A4").String = pref + oSheet.getCellRangeByName("A4").String
        oSheet.getCellRangeByName("A4").CellBackColor = 14942166

        LeenoUtils.DocumentRefresh(False)
    #  oSheet = oDoc.CurrentController.ActiveSheet

    # LeenoUtils.DocumentRefresh(False)
    try:
        oSheet = oDoc.CurrentController.ActiveSheet
        if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
            try:
                sRow = oDoc.getCurrentSelection().getRangeAddresses(
                )[0].StartRow
                eRow = oDoc.getCurrentSelection().getRangeAddresses()[0].EndRow

            except Exception:
                sRow = oDoc.getCurrentSelection().getRangeAddress().StartRow
                eRow = oDoc.getCurrentSelection().getRangeAddress().EndRow
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, sRow)
            # sStRange.RangeAddress
            sRow = sStRange.RangeAddress.StartRow
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, eRow)
            try:
                sStRange.RangeAddress
            except Exception:
                return
            inizio = sStRange.RangeAddress.StartRow
            eRow = sStRange.RangeAddress.EndRow + 1
            lrow = sRow
            fini = []
            for x in range(sRow, eRow):
                if oSheet.getCellByPosition(
                        0, x).CellStyle == 'Comp End Attributo':
                    fini.append(x)
                elif oSheet.getCellByPosition(
                        0, x).CellStyle == 'Comp End Attributo_R':
                    fini.append(x - 2)
        idx = 0
        for lrow in reversed(fini):
            lrow += idx
            try:
                sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
                sStRange.RangeAddress
                inizio = sStRange.RangeAddress.StartRow
                fine = sStRange.RangeAddress.EndRow
                oSheet.getCellByPosition(1, inizio + 1).CellBackColor = 14942166
                if oSheet.Name == 'CONTABILITA':
                    fine -= 1
                _gotoCella(2, fine - 1)
               
                if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
                    # voce = (num, art, desc, um, quantP, prezzo, importo, sic, mdo)
                    # art = LeenoComputo.datiVoceComputo(oSheet, lrow)[1]
                    pesca_cod()
                    vds_ep()
                    pesca_cod()

                elif oSheet.Name == 'Elenco Prezzi':
                    vds_ep()

                    ###
                lrow = LeggiPosizioneCorrente()[1]
                lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 1)
            except Exception:
                pass
        # numera_voci(1)
    except Exception:
        pass
    #  _gotoCella(1, fine +3)
    LeenoUtils.DocumentRefresh(True)

########################################################################


def MENU_prefisso_codice():
    '''
    Aggiunge prefisso al Codice Articolo
    '''

    chiudi_dialoghi()
    if Dialogs.YesNoDialog(IconType="question", Title='AVVISO!',
    Text='''Questo comando aggiunge un prefisso a scelta
ai SOLI codici di prezzo selezionati,
oppure a tutti se non ne sono stati selezionati.

Procedo?''') == 1:
        pass
    else:
        return

    LeenoUtils.DocumentRefresh(False)

    testo = ''
    prefisso = InputBox(
        testo, t='Inserisci il prefisso per il Codice Articolo (es: "BAS22/1_").')
    if prefisso in (None, '', ' '):
        return
    oDoc = LeenoUtils.getDocument()
    stili = {
    'Elenco Prezzi': (0, 'EP-aS'),
    'COMPUTO': (1, 'comp Art-EP_R'),
    'VARIANTE': (1, 'comp Art-EP_R'),
    'CONTABILITA': (1, 'comp Art-EP_R'),
    'Analisi di Prezzo': (0, 'An-lavoraz-Cod-sx')
    }

    #cattura l'elenco voci selezionate
    lsubst = list()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in ('Elenco Prezzi', 'Analisi di Prezzo'):
        pass
    else:
        LeenoUtils.DocumentRefresh(True)
        return
    try:
        selezioni = oDoc.getCurrentSelection().getRangeAddresses()
        for el in selezioni:
            sRow = el.StartRow
            eRow = el.EndRow + 1
            lsubst.append(range(sRow, eRow))
    except AttributeError:
        el = oDoc.getCurrentSelection().getRangeAddress()
        sRow = el.StartRow
        eRow = el.EndRow + 1
        lsubst.append(range(sRow, eRow))
        LeenoUtils.DocumentRefresh(True)

    #lista dei codici da sostituire
    lista = []
    for el in lsubst:
        for y in el:
            lista.append(oSheet.getCellByPosition(0, y).String)
    if len(lista) == 1:
        lista = []
        lrow = SheetUtils.uFindStringCol('Fine elenco', 0, oSheet)
        for y in range(0, lrow):
            lista.append(oSheet.getCellByPosition(0, y).String)
    
    i = 0
    for el in ('Elenco Prezzi', 'COMPUTO', 'VARIANTE', 'CONTABILITA', 'Analisi di Prezzo'):
        i = i + 1
        try:
            oSheet = oDoc.getSheets().getByName(el)
            lrow = SheetUtils.getLastUsedRow(oSheet)
            x = stili[el][0]
            stile = stili[el][1]
            for y in range(0, lrow):
                codice = oSheet.getCellByPosition(x, y).String
                if codice in lista:
                    if oSheet.getCellByPosition(x, y).CellStyle == stile and \
                        oSheet.getCellByPosition(x, y).String != "000":
                        if codice.startswith('VDS_'):
                            nuovo_codice = 'VDS_' + prefisso + '_' + codice[4:]
                        else:
                            nuovo_codice = prefisso + '_' + codice
                        oSheet.getCellByPosition(x, y).String = nuovo_codice
        except:
            LeenoUtils.DocumentRefresh(True)
            pass
    LeenoUtils.DocumentRefresh(True)

########################################################################


def nascondi_voci_zero(lcol = None):
    '''
    Nasconde le voci il cui valore della colonna corrente è pari a zero.
    '''
    # chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if lcol is None:
        lcol = LeggiPosizioneCorrente()[0]

    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet

    ER = SheetUtils.getLastUsedRow(oSheet)

    for i in reversed(range(3, ER)):
        if oSheet.getCellByPosition(lcol, i).Value == 0 and \
            oSheet.getCellByPosition(lcol, i).CellStyle not in ('ULTIMUS', 'Ultimus_centro'):
            oCellRangeAddr.StartRow = i
            oCellRangeAddr.EndRow = i
            oSheet.ungroup(oCellRangeAddr, 1)
            oSheet.group(oCellRangeAddr, 1)
            oSheet.getCellRangeByPosition(lcol, i, lcol, i).Rows.IsVisible = False


########################################################################
def Cancel():
    return -1

def cancella_voci_non_usate():
    '''
    Cancella le voci di prezzo non utilizzate.
    '''
    with LeenoUtils.DocumentRefreshContext(False):
        chiudi_dialoghi()

        # if Dialogs.YesNoDialog(Title='AVVISO!',
        if Dialogs.YesNoDialog(IconType="question", Title='AVVISO!',
        Text='''Questo comando ripulisce l'Elenco Prezzi
    dalle voci non utilizzate in nessuno degli altri elaborati.

    La procedura potrebbe richiedere del tempo.

    Vuoi procedere comunque?''') == 0:
            return
        oDoc = LeenoUtils.getDocument()
        # oDoc.enableAutomaticCalculation(False)
        oSheet = oDoc.CurrentController.ActiveSheet

        oRange = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
        SRep = oRange.StartRow + 1
        ERep = oRange.EndRow
        lista_prezzi = []
        # prende l'elenco dal foglio Elenco Prezzi
        for n in range(SRep, ERep):
            lista_prezzi.append(oSheet.getCellByPosition(0, n).String)
        # attiva la progressbar
        indicator = oDoc.getCurrentController().getStatusIndicator()
        n = 0
        # prende l'elenco da tutti gli altri fogli
        if 'ELENCO DEI COSTI ELEMENTARI' in lista_prezzi:
            lista_prezzi.remove('ELENCO DEI COSTI ELEMENTARI')
        lista = []
        for tab in ('COMPUTO', 'Analisi di Prezzo', 'VARIANTE', 'CONTABILITA'):
            try:
                oSheet = oDoc.getSheets().getByName(tab)
                ER = SheetUtils.getLastUsedRow(oSheet)
                indicator.start("Eliminazione delle voci in corso...", ER)

                if tab == 'Analisi di Prezzo':
                    stile = 'An-lavoraz-Cod-sx'
                    for n in range(0, ER):
                        indicator.setValue(n)
                        cell = oSheet.getCellByPosition(0, n)
                        if cell.CellStyle == stile:
                            lista.append(cell.String)
                else:
                    stile = 'comp Art-EP_R'
                    for n in range(0, ER):
                        cell = oSheet.getCellByPosition(1, n)
                        if cell.CellStyle == stile:
                            lista.append(cell.String)
            except Exception as e:
                # DLG.errore(e)
                pass
        indicator.start("Eliminazione delle voci in corso...", 5)  # 100 = max progresso
        indicator.setValue(2)

        da_cancellare = set(lista_prezzi).difference(set(lista))
        oSheet = oDoc.CurrentController.ActiveSheet
        oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        indicator.setValue(3)
        struttura_off()
        struttura_off()
        struttura_off()
        indicator.setValue(4)

        for n in reversed(range(SRep, ERep)):
            cell_0 = oSheet.getCellByPosition(0, n).String
            cell_1 = oSheet.getCellByPosition(1, n).String
            cell_4 = oSheet.getCellByPosition(4, n).String

            if cell_0 in da_cancellare or (cell_0 == '' and cell_1 == '' and cell_4 == ''):
                oSheet.Rows.removeByIndex(n, 1)

        indicator.setValue(5)
        indicator.end()
        _gotoCella(0, 3)
        Dialogs.Info(Title = 'Ricerca conclusa', Text=f"Eliminate {len(da_cancellare)} voci dall'elenco prezzi.")


########################################################################


def voce_breve_an():
    '''
    Ottimizza l'altezza delle celle di Analisi di Prezzo o visualizza solo
    tre righe della descrizione.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ER = SheetUtils.getLastUsedRow(oSheet)

    if not oSheet.getCellRangeByName('B3').Rows.OptimalHeight:
        LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    else:
        altezza_base = oSheet.getCellRangeByName('B3').CharHeight * 66 / 3 * 2
        nr_descrizione = float(cfg.read('Generale', 'altezza_celle'))

        hriga = 100 + altezza_base * nr_descrizione  # Calcola l'altezza desiderata

        for el in range(0, ER):
            if oSheet.getCellByPosition(1, el).CellStyle == 'An-1-descr_':
                oSheet.getCellByPosition(1, el).Rows.Height = hriga


########################################################################


def voce_breve_ep():
    '''
    Ottimizza l'altezza delle celle di Elenco Prezzi o visualizza solo
    tre righe della descrizione.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    oRange = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    SR = oRange.StartRow + 1
    ER = oRange.EndRow

    cell_1_3 = oSheet.getCellByPosition(1, 3)

    if not cell_1_3.Rows.OptimalHeight:
        LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    else:
        altezza_base = oSheet.getCellRangeByName('B4').CharHeight * 64 / 3 * 2
        nr_descrizione = float(cfg.read('Generale', 'altezza_celle'))
        hriga = 100 + altezza_base * nr_descrizione  # Calcola l'altezza desiderata

        oSheet.getCellRangeByPosition(0, SR, 0, ER).Rows.Height = hriga



########################################################################

def scelta_viste():
    with LeenoUtils.DocumentRefreshContext(False):
        scelta_viste_run()

def scelta_viste_run():
    '''
    Gestisce i dialoghi del menù viste nelle tabelle di Analisi di Prezzo,
    Elenco Prezzi, COMPUTO, VARIANTE, CONTABILITA'
    Genera i raffronti tra COMPUTO e VARIANTE e CONTABILITA'
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    n = SheetUtils.getLastUsedRow(oSheet)

    LeenoSheetUtils.memorizza_posizione()

    psm = LeenoUtils.getComponentContext().ServiceManager
    dp = psm.createInstance('com.sun.star.awt.DialogProvider')
    global oDialog1
    if oSheet.Name in ('VARIANTE', 'COMPUTO'):
        oDialog1 = dp.createDialog(
            'vnd.sun.star.script:UltimusFree2.DialogViste_A?language=Basic&location=application'
        )
        # configuro il dialogo
        oDialog1.getControl('Dettaglio').State = cfg.read('Generale', 'dettaglio')
        if oSheet.getColumns().getByIndex(5).Columns.IsVisible:
            oDialog1.getControl('CBMis').State = 1
        if oSheet.getColumns().getByIndex(17).Columns.IsVisible:
            oDialog1.getControl('CBSic').State = 1
        if oSheet.getColumns().getByIndex(28).Columns.IsVisible:
            oDialog1.getControl('CBMat').State = 1
        
        try:
            if oSheet.getColumns().getByIndex(29).Columns.IsVisible:
                oDialog1.getControl('CBMdo').State = True
        except:
            pass

        if oSheet.getColumns().getByIndex(31).Columns.IsVisible:
            oDialog1.getControl('CBCat').State = 1
        if oSheet.getColumns().getByIndex(38).Columns.IsVisible:
            oDialog1.getControl('CBFig').State = 1

        sString = oDialog1.getControl('TextField10')
        sString.Text = oDoc.getSheets().getByName('S1').getCellRangeByName(
            'H337').Value  # inizio_voci_abbreviate
        sString = oDialog1.getControl('TextField11')

        sString.Text = oDoc.getSheets().getByName('S1').getCellRangeByName(
            'H338').Value  # fine_voci_abbreviate
        oDialog1.getControl('ComboVISTA').setText('Scegli...')

# VARIANTE
        if oSheet.Name == 'VARIANTE':
            oDialog1.getControl('CommandButton7').setEnable(0)

        # descrizione_in_una_colonna
        for i in range (10):
            if oSheet.getCellByPosition(2, i).CellStyle == 'Comp-Bianche in mezzo Descr' and \
                oSheet.getCellByPosition(2, i).IsMerged:
                oDialog1.getControl('Descrizione_condensata').State = 0
                break
            else:
                oDialog1.getControl('Descrizione_condensata').State = 1

        if oDialog1.execute() == 0:
            return

        # leggo le scelte
        # il salvataggio anche su leeno.conf serve alla funzione voce_breve()
        if oDialog1.getControl('TextField10').getText() != '10000':
            cfg.write('Computo', 'inizio_voci_abbreviate', oDialog1.getControl('TextField10').getText())
        oDoc.getSheets().getByName('S1').getCellRangeByName('H337').Value = float(oDialog1.getControl('TextField10').getText())

        if oDialog1.getControl('TextField11').getText() != '10000':
            cfg.write('Computo', 'fine_voci_abbreviate', oDialog1.getControl('TextField11').getText())
        oDoc.getSheets().getByName('S1').getCellRangeByName('H338').Value = float(oDialog1.getControl('TextField11').getText())

        if oDialog1.getControl("CBMis").State == 0:  # misure
            for el in range (5, 9):
                oSheet.getColumns().getByIndex(el).Columns.IsVisible = False
            #copia la formula dell'UM nella colonna C
            for el in range(4, n):
                if oSheet.getCellByPosition(2, el).CellStyle == "comp sotto centro":
                    oSheet.getCellByPosition(2, el).Formula = oSheet.getCellByPosition(8, el).Formula
                    oDoc.StyleFamilies.getByName("CellStyles").getByName(
                            'comp sotto centro').HoriJustify = RIGHT
                    oDoc.StyleFamilies.getByName("CellStyles").getByName(
                            'comp sotto centro').CharPosture = NONE
        else:
            for el in range (5, 9):
                oSheet.getColumns().getByIndex(el).Columns.IsVisible = True

        if oDialog1.getControl('Descrizione_condensata').State == 1:
            descrizione_in_una_colonna(False)
        else:
            descrizione_in_una_colonna(True)

        if oDialog1.getControl('CBMat').State == 0:  # materiali
            oSheet.getColumns().getByIndex(28).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(28).Columns.IsVisible = True

        if oDialog1.getControl('CBCat').State == 0:  # categorie
            oSheet.getColumns().getByIndex(31).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(32).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(33).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(31).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(32).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(33).Columns.IsVisible = True

        if oDialog1.getControl("CBSic").State == 0:  # sicurezza
            oSheet.getColumns().getByIndex(17).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(17).Columns.IsVisible = True

        if oDialog1.getControl("CBFig").State == 0:  # figure
            oSheet.getColumns().getByIndex(38).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(38).Columns.IsVisible = True

        if oDialog1.getControl('Dettaglio').State == 0:  #
            cfg.write('Generale', 'dettaglio', '0')
            dettaglio_misure(0)
        else:
            cfg.write('Generale', 'dettaglio', '1')
            dettaglio_misure(0)
            dettaglio_misure(1)
        
        Menu_adattaAltezzaRiga()

        if oDialog1.getControl('ComboVISTA').getText() == 'Predefinita':
            vSintetica(False)
            vista_configurazione('terra_terra')
        elif oDialog1.getControl('ComboVISTA').getText() == 'Incidenza Manodopera':
            vista_configurazione('mdo')
        elif oDialog1.getControl('ComboVISTA').getText() == 'Sintetica':
            vSintetica(True)

# CONTABILITA
    elif oSheet.Name in ('CONTABILITA', 'Registro', 'SAL'):
        GotoSheet('CONTABILITA')
        oSheet = oDoc.CurrentController.ActiveSheet

        oDialog1 = dp.createDialog(
            "vnd.sun.star.script:UltimusFree2.Dialogviste_N?language=Basic&location=application"
        )

        oDialog1.getControl('ComboVISTA').setText('Scegli...')
    
        # Inizio voce
        sString = oDialog1.getControl('TextField3')
        sString.Text = oDoc.getSheets().getByName('S1').getCellRangeByName(
            'H335').Value  # cont_inizio_voci_abbreviate

        # Fine voce
        sString = oDialog1.getControl('TextField2')
        sString.Text = oDoc.getSheets().getByName('S1').getCellRangeByName(
            'H336').Value  # cont_fine_voci_abbreviate

        # descrizione_in_una_colonna
        for i in range (10):
            if oSheet.getCellByPosition(2, i).CellStyle == 'Comp-Bianche in mezzo Descr_R' and \
                oSheet.getCellByPosition(2, i).IsMerged:
                oDialog1.getControl('Descrizione_condensata').State = 0
                break
            else:
                oDialog1.getControl('Descrizione_condensata').State = 1

        listaSal = LeenoContab.ultimo_sal()

    # Mostra SAL n.
        oDialog1.getControl('ComboBox1').addItems(listaSal, 1)
        try:
            oDialog1.getControl('ComboBox1').Text = listaSal[-1]
        except:
            pass
        if len(listaSal) != 0:
            oDialog1.getControl('RimuoviSAL').Label = "Elimina atti SAL n. " + str(len(listaSal))
            oDialog1.getControl('GeneraAtti').Label = 'Genera atti SAL n. ' + str(len(listaSal)+1)
        else:
            oDialog1.getControl('RimuoviSAL').Enable = False
            oDialog1.getControl('RimuoviSAL').Label = "Nessun SAL da rimuovere"
            oDialog1.getControl('SituazioneContabile').Enable = False
            oDialog1.getControl('GeneraAtti').Label = 'Genera atti SAL n. 1'

        if oSheet.getCellRangeByName('A4').CellStyle == 'Comp TOTALI':
            oDialog1.getControl('EliminaMisure').Enable = False

        if oSheet.getCellByPosition(0, SheetUtils.uFindStringCol('T O T A L E', 2, oSheet) - 1).CellStyle == 'Ultimus_centro_bordi_lati':
            oDialog1.getControl('GeneraAtti').Enable = False
            oDialog1.getControl('GeneraAtti').Label = 'Nessun SAL da generare'

        # Ricicla voci da
        sString = oDialog1.getControl('ComboBox3')
        sString.Text = cfg.read('Contabilita', 'ricicla_da')

        oDialog1.getControl('Dettaglio').State = cfg.read('Generale', 'dettaglio')

        if oDialog1.execute() == 0:
            return

        if oDialog1.getControl('Descrizione_condensata').State == 1:
            descrizione_in_una_colonna(False)
        else:
            descrizione_in_una_colonna(True)
 
        # il salvataggio anche su leeno.conf serve alla funzione voce_breve()
        if oDialog1.getControl('TextField3').getText() != '10000':
            cfg.write('Contabilita', 'cont_inizio_voci_abbreviate', oDialog1.getControl('TextField3').getText())
        oDoc.getSheets().getByName('S1').getCellRangeByName('H335').Value = float(oDialog1.getControl('TextField3').getText())

        if oDialog1.getControl('TextField2').getText() != '10000':
            cfg.write('Contabilita', 'cont_fine_voci_abbreviate', oDialog1.getControl('TextField2').getText())
        oDoc.getSheets().getByName('S1').getCellRangeByName('H336').Value = float(oDialog1.getControl('TextField2').getText())

        if oDialog1.getControl('ComboBox3').getText() in ('COMPUTO', '&305.Dlg_config.ComboBox3.Text'):
            cfg.write('Contabilita', 'ricicla_da', 'COMPUTO')
        else:
            cfg.write('Contabilita', 'ricicla_da', 'VARIANTE')

        if oDialog1.getControl('Dettaglio').State == 0:
            cfg.write('Generale', 'dettaglio', '0')
            dettaglio_misure(0)
        else:
            cfg.write('Generale', 'dettaglio', '1')
            dettaglio_misure(0)
            dettaglio_misure(1)

        try:
            nSal = int(oDialog1.getControl('ComboBox1').getText())
            LeenoContab.mostra_sal(nSal)
        except Exception as e:
            pass

        if oDialog1.getControl('ComboVISTA').getText() == 'Predefinita':
            vSintetica(False)
            vista_configurazione('terra_terra')
        elif oDialog1.getControl('ComboVISTA').getText() == 'Semplificata':
            vista_configurazione('Semplificata')
        elif oDialog1.getControl('ComboVISTA').getText() == 'Sintetica':
            vSintetica(True)

#Elenco Prezzi
    elif oSheet.Name in ('Elenco Prezzi'):
        oCellRangeAddr = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
        oDialog1 = dp.createDialog(
            "vnd.sun.star.script:UltimusFree2.DialogViste_EP?language=Basic&location=application"
        )

        if oSheet.getColumns().getByIndex(3).Columns.IsVisible:
            oDialog1.getControl('CBSic').State = True
        if oSheet.getColumns().getByIndex(5).Columns.IsVisible:
            oDialog1.getControl('CBMdo').State = True
        if oSheet.getColumns().getByIndex(7).Columns.IsVisible:
            oDialog1.getControl('CBOrig').State = True

        costo_elem_row = SheetUtils.uFindStringCol('ELENCO DEI COSTI ELEMENTARI', 1, oSheet)
        if costo_elem_row:
            oDialog1.getControl('Titolo_COSTI').Enable = False

        if oDialog1.execute() == 1:
            if oDialog1.getControl("CBSic").State == 0:  # sicurezza
                oSheet.getColumns().getByIndex(3).Columns.IsVisible = False
            else:
                oSheet.getColumns().getByIndex(3).Columns.IsVisible = True

            if oDialog1.getControl("CBMdo").State == 0:  # manodopera
                oSheet.getColumns().getByIndex(5).Columns.IsVisible = False
                oSheet.getColumns().getByIndex(6).Columns.IsVisible = False
            else:
                oSheet.getColumns().getByIndex(5).Columns.IsVisible = True
                oSheet.getColumns().getByIndex(6).Columns.IsVisible = True

            if oDialog1.getControl("CBOrig").State == 0:  # origine
                oSheet.getColumns().getByIndex(7).Columns.IsVisible = False
            else:
                oSheet.getColumns().getByIndex(7).Columns.IsVisible = True

            if oDialog1.getControl("CBSom").State == 1:
                genera_sommario()

            oRangeAddress = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
            SR = oRangeAddress.StartRow + 1
            ER = oRangeAddress.EndRow - 1

            formule = []

            oSheet.getCellRangeByName('X2').Formula = '=IF(N(U2)>N(T2); N(U2)-N(T2); "")'
            oSheet.getCellRangeByName('Y2').Formula = '=IF(N(T2)>N(U2); N(T2)-N(U2); "")'
            ultima_voce = LeenoSheetUtils.cercaUltimaVoce(oSheet)
    
            if oDialog1.getControl('ComboRAFFRONTO').getText() == 'Computo - Variante':
                inizializza_elenco()
                genera_sommario()
                oSheet.getCellRangeByName('Z2').Formula = '=IFERROR(LET(t;N(T2);u;N(U2);IF(AND(t=0;u=0);"--";IFS(u=0;-1;t=0;1;t=u;"--";t>u;-(t-u)/t;t<u;(u-t)/t)));"--")'

                oSheet.getCellRangeByName('X1').String = 'COMPUTO_VARIANTE'
                LeenoSheetUtils.setLarghezzaColonne(oSheet)
                for n in range(4, ultima_voce + 2):
                    formule.append([
                        f'=IF(N(U{n})>N(T{n}); N(U{n})-N(T{n}); "")',
                        f'=IF(N(T{n})>N(U{n}); N(T{n})-N(U{n}); "")',
                        f'=IFERROR(LET(t;N(T{n});u;N(U{n});IF(AND(t=0;u=0);"--";IFS(u=0;-1;t=0;1;t=u;"--";t>u;-(t-u)/t;t<u;(u-t)/t)));"--")',
                    ])

                n += 1
                oRange = oSheet.getCellRangeByPosition(23, 3, 25, ultima_voce)
                formule = tuple(formule)
                oRange.setFormulaArray(formule)

                oSheet.getCellRangeByName(f'Z{n}').Formula = f'=IFERROR(LET(t;N(T{n});u;N(U{n});IF(AND(t=0;u=0);"--";IFS(u=0;-1;t=0;1;t=u;"--";t>u;-(t-u)/t;t<u;(u-t)/t)));"--")'

            if oDialog1.getControl('ComboRAFFRONTO').getText() == 'Computo - Contabilità':
                inizializza_elenco()
                genera_sommario()
                oSheet.getCellRangeByName('Z2').Formula = '=IFERROR(LET(t;N(T2);u;N(V2);IF(AND(t=0;u=0);"--";IFS(u=0;-1;t=0;1;t=u;"--";t>u;-(t-u)/t;t<u;(u-t)/t)));"--")'

                oSheet.getCellRangeByName ('X1').String = 'COMPUTO_CONTABILITÀ'
                LeenoSheetUtils.setLarghezzaColonne(oSheet)
                for n in range(4, ultima_voce + 2):
                    formule.append([
                        f'=IF(N(V{n})>N(T{n}); N(V{n})-N(T{n}); "")',
                        f'=IF(N(T{n})>N(V{n}); N(T{n})-N(V{n}); "")',
                        f'=IFERROR(LET(t;N(T{n});u;N(V{n});IF(AND(t=0;u=0);"--";IFS(u=0;-1;t=0;1;t=u;"--";t>u;-(t-u)/t;t<u;(u-t)/t)));"--")',
                    ])

                n += 1
                oRange = oSheet.getCellRangeByPosition(23, 3, 25, ultima_voce)
                formule = tuple(formule)
                oRange.setFormulaArray(formule)

                if oRangeAddress.StartColumn != 0:
                    if DLG.DlgSiNo(
                            "Nascondo eventuali voci non ancora contabilizzate?"
                    ) == 2:
                        for el in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
                            if oSheet.getCellByPosition(20, el).Value == 0:
                                oCellRangeAddr.StartRow = el
                                oCellRangeAddr.EndRow = el
                                oSheet.group(oCellRangeAddr, 1)
                                oSheet.getCellRangeByPosition(
                                    0, el, 1, el).Rows.IsVisible = False

                oSheet.getCellRangeByName(f'Z{n}').Formula = f'=IFERROR(LET(t;N(T{n});u;N(V{n});IF(AND(t=0;u=0);"--";IFS(u=0;-1;t=0;1;t=u;"--";t>u;-(t-u)/t;t<u;(u-t)/t)));"--")'

            if oDialog1.getControl('ComboRAFFRONTO').getText() == 'Variante - Contabilità':
                inizializza_elenco()
                genera_sommario()
                oSheet.getCellRangeByName('Z2').Formula = '=IFERROR(LET(t;N(U2);u;N(V2);IF(AND(t=0;u=0);"--";IFS(u=0;-1;t=0;1;t=u;"--";t>u;-(t-u)/t;t<u;(u-t)/t)));"--")'

                oSheet.getCellRangeByName ('X1').String = 'VARIANTE_CONTABILITÀ'
                LeenoSheetUtils.setLarghezzaColonne(oSheet)
                for n in range(4, ultima_voce + 2):
                    formule.append([
                        f'=IF(N(V{n})>N(U{n}); N(V{n})-N(U{n}); "")',
                        f'=IF(N(U{n})>N(V{n}); N(U{n})-N(V{n}); "")',
                        f'=IFERROR(LET(t;N(U{n});u;N(V{n});IF(AND(t=0;u=0);"--";IFS(u=0;-1;t=0;1;t=u;"--";t>u;-(t-u)/t;t<u;(u-t)/t)));"--")',
                    ])
                n += 1
                oRange = oSheet.getCellRangeByPosition(23, 3, 25, ultima_voce)
                formule = tuple(formule)
                oRange.setFormulaArray(formule)

                oSheet.getCellRangeByName(f'Z{n}').Formula = f'=IFERROR(LET(t;N(U{n});u;N(V{n});IF(AND(t=0;u=0);"--";IFS(u=0;-1;t=0;1;t=u;"--";t>u;-(t-u)/t;t<u;(u-t)/t)));"--")'

            LeenoSheetUtils.inserisciRigaRossa(oSheet)
        else:
            return oDialog1.dispose()
        # evidenzia le quantità eccedenti il VI/V
        for el in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
            if oSheet.getCellByPosition(
                    26,
                    el).Value >= 0.2 or oSheet.getCellByPosition(
                        26, el).String == '20,00%':
                oSheet.getCellRangeByPosition(
                    0, el, 1, el).CellBackColor = 16777175

        oDoc.CurrentController.select(oSheet.getCellRangeByName('Z2'))
        comando('Copy')
        oDoc.CurrentController.select(
            oSheet.getCellRangeByPosition(25, 3, 25, ER+1))
        paste_format()

        _primaCella()
        oSheet.getCellRangeByPosition(11, 3, 13, ER+1).CellBackColor = 16777175

        oSheet.getCellRangeByName(f'A{n+1}:Z{n+1}').CharWeight = BOLD

# Analisi di Prezzo
    elif oSheet.Name in ('Analisi di Prezzo'):
        oDialog1 = dp.createDialog(
            "vnd.sun.star.script:UltimusFree2.DialogViste_AN?language=Basic&location=application"
        )

        oS1 = oDoc.getSheets().getByName('S1')

        sString = oDialog1.getControl('TextField5')
        # sString.Text = cfg.read('Analisi', 'sicurezza')
        sString.Text = oS1.getCellRangeByName('S1.H319').Value * 100  # sicurezza
        
        sString = oDialog1.getControl('TextField6')
        # sString.Text = cfg.read('Analisi', 'spese_generali')
        sString.Text = oS1.getCellRangeByName('S1.H320').Value * 100  # spese_generali

        sString = oDialog1.getControl('TextField7')
        # sString.Text = cfg.read('Analisi', 'utile_impresa')
        sString.Text = oS1.getCellRangeByName('S1.H321').Value * 100  # utile_impresa

        # accorpa_spese_utili
        if oS1.getCellRangeByName('S1.H323').Value == 1:
            oDialog1.getControl('CheckBox4').State = 1
        sString = oDialog1.getControl('TextField8')
        sString.Text = oS1.getCellRangeByName('S1.H324').Value * 100  # sconto
        sString = oDialog1.getControl('TextField9')
        sString.Text = oS1.getCellRangeByName(
            'S1.H326').Value * 100  # maggiorazione

        if oDialog1.execute() == 0:
            return

        
        sicurezza = oDialog1.getControl('TextField5').getText().replace(',', '.')
        oS1.getCellRangeByName('S1.H319').Value = float(sicurezza) / 100  # sicurezza
        cfg.write('Analisi', 'sicurezza', '0')
        
        spese_generali = oDialog1.getControl('TextField6').getText().replace(',', '.')
        oS1.getCellRangeByName('S1.H320').Value = float(spese_generali) / 100  # spese_generali
        cfg.write('Analisi', 'spese_generali', spese_generali)

        utile_impresa = oDialog1.getControl('TextField7').getText().replace(',', '.')
        oS1.getCellRangeByName('S1.H321').Value = float(utile_impresa) / 100  # utile_impresa
        cfg.write('Analisi', 'utile_impresa', utile_impresa)

        oS1.getCellRangeByName('S1.H323').Value = oDialog1.getControl(
            'CheckBox4').State
        oS1.getCellRangeByName('S1.H324').Value = float(
            oDialog1.getControl('TextField8').getText().replace(
                ',', '.')) / 100  # sconto
        oS1.getCellRangeByName('S1.H326').Value = float(
            oDialog1.getControl('TextField9').getText().replace(
                ',', '.')) / 100  # maggiorazione

        # accorpa_spese_utili
        if oS1.getCellRangeByName('S1.H323').Value == 1:
            oDialog1.getControl('CheckBox4').State = 1
        sString = oDialog1.getControl('TextField8')
        sString.Text = oS1.getCellRangeByName('S1.H324').Value * 100  # sconto
        sString = oDialog1.getControl('TextField9')
        sString.Text = oS1.getCellRangeByName(
            'S1.H326').Value * 100  # maggiorazione

    fissa()
    LeenoSheetUtils.ripristina_posizione()


########################################################################


def genera_variante():
    '''
    Genera il foglio di VARIANTE a partire dal COMPUTO
    @@@ MODIFICA IN CORSO CON 'LeenoVariante.generaVariante'
    '''
    # chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('VARIANTE'):
        if oDoc.NamedRanges.hasByName("AA"):
            oDoc.NamedRanges.removeByName("AA")
            oDoc.NamedRanges.removeByName("BB")
        oDoc.Sheets.copyByName('COMPUTO', 'VARIANTE', 4)
        oSheet = oDoc.getSheets().getByName('COMPUTO')
        lrow = SheetUtils.getUsedArea(oSheet).EndRow
        SheetUtils.NominaArea(oDoc, 'COMPUTO', '$AJ$3:$AJ$' + str(lrow), 'AA')
        SheetUtils.NominaArea(oDoc, 'COMPUTO', '$N$3:$N$' + str(lrow), "BB")
        SheetUtils.NominaArea(oDoc, 'COMPUTO', '$AK$3:$AK$' + str(lrow), "cEuro")
        oSheet = oDoc.getSheets().getByName('VARIANTE')
        GotoSheet('VARIANTE')
        setTabColor(16777175)
        oSheet.getCellByPosition(2, 0).Formula = '=RIGHT(CELL("FILENAME"; A1); LEN(CELL("FILENAME"; A1)) - FIND("$"; CELL("FILENAME"; A1)))'
        oSheet.getCellByPosition(2, 0).CellStyle = "comp Int_colonna"
        oSheet.getCellRangeByName("C1").CellBackColor = 16777175
        oSheet.getCellRangeByPosition(0, 2, 42, 2).CellBackColor = 16777175
        if DLG.DlgSiNo(
                """Vuoi svuotare la VARIANTE appena generata?

Se decidi di continuare, cancellerai tutte le voci di
misurazione già presenti in questo elaborato.
Cancello le voci di misurazione?
 """, 'ATTENZIONE!') == 2:
            lrow = SheetUtils.uFindStringCol('TOTALI COMPUTO', 2, oSheet) - 3
            oSheet.Rows.removeByIndex(3, lrow)
            _gotoCella(0, 2)
            LeenoComputo.ins_voce_computo()
            oSheet = oDoc.getSheets().getByName('VARIANTE')
            LeenoSheetUtils.adattaAltezzaRiga(oSheet)
            if SheetUtils.uFindStringCol('Riepilogo strutturale delle Categorie', 2, oSheet):
                row = SheetUtils.uFindStringCol('Riepilogo strutturale delle Categorie', 2, oSheet)
                _gotoCella(0, row)
                LeenoSheetUtils.elimina_voce(row, msg = 0)
                _gotoCella(1, 4)
    #  else:
    GotoSheet('VARIANTE')
    ScriviNomeDocumentoPrincipale()
    Menu_adattaAltezzaRiga()
    LeenoEvents.assegna()


########################################################################

def genera_sommario():
    '''
    Genera i sommari in Elenco Prezzi
    '''
    struttura_off()

    oDoc = LeenoUtils.getDocument()
    
    sistema_aree()

    with LeenoUtils.DocumentRefreshContext(False):

        # oSheet = oDoc.getSheets().getByName('COMPUTO')
        # lrow = SheetUtils.getUsedArea(oSheet).EndRow
        # SheetUtils.NominaArea(oDoc, 'COMPUTO', '$AJ$3:$AJ$' + str(lrow), 'AA')
        # SheetUtils.NominaArea(oDoc, 'COMPUTO', '$N$3:$N$' + str(lrow), "BB")
        # SheetUtils.NominaArea(oDoc, 'COMPUTO', '$AK$3:$AK$' + str(lrow), "cEuro")

        # if oDoc.getSheets().hasByName('VARIANTE'):
        #     oSheet = oDoc.getSheets().getByName('VARIANTE')
        #     lrow = SheetUtils.getUsedArea(oSheet).EndRow
        #     SheetUtils.NominaArea(oDoc, 'VARIANTE', '$AJ$3:$AJ$' + str(lrow), 'varAA')
        #     SheetUtils.NominaArea(oDoc, 'VARIANTE', '$N$3:$N$' + str(lrow), "varBB")
        #     SheetUtils.NominaArea(oDoc, 'VARIANTE', '$AK$3:$AK$' + str(lrow), "varEuro")

        # if oDoc.getSheets().hasByName('CONTABILITA'):
        #     oSheet = oDoc.getSheets().getByName('CONTABILITA')
        #     lrow = SheetUtils.getUsedArea(oSheet).EndRow
        #     lrow = SheetUtils.getUsedArea(
        #         oDoc.getSheets().getByName('CONTABILITA')).EndRow
        #     SheetUtils.NominaArea(oDoc, 'CONTABILITA', '$AJ$3:$AJ$' + str(lrow), 'GG')
        #     # SheetUtils.NominaArea(oDoc, 'CONTABILITA', '$S$3:$S$' + str(lrow), "G1G1")
        #     SheetUtils.NominaArea(oDoc, 'CONTABILITA', '$J$3:$J$' + str(lrow), "G1G1")
        #     SheetUtils.NominaArea(oDoc, 'CONTABILITA', '$AK$3:$AK$' + str(lrow), "conEuro")

        formule = []
        oSheet = oDoc.getSheets().getByName('Elenco Prezzi')

        # attiva la progressbar
        indicator = oDoc.getCurrentController().getStatusIndicator()
        ultima_voce = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1
        indicator.start("Genera sommario...", ultima_voce)  # 100 = max progresso

        for n in range(4, ultima_voce + 1):
            indicator.Value = n
            # Controlla lo stile della cella nella prima colonna
            cell = oSheet.getCellByPosition(0, n-1)  # -1 perché gli indici partono da 0
            cell_style = cell.CellStyle
            
            if cell_style == "Ultimus_centro":
                # Se lo stile è "Ultimus_centro", inserisci formula vuota
                stringa = [""] * 11
            else:
                stringa = [
                    # Quantità Computo
                    f"=LET(s; SUMIF(AA; A{n}; BB); IF(s; s; \"--\"))",
                    # Quantità Variante
                    f'=LET(s; SUMIF(varAA; A{n}; varBB);IFERROR(IF(s; s; "--"); "--"))',
                    # Quantità Contabilità
                    f'=LET(s; SUMIF(GG; A{n}; G1G1);IFERROR(IF(s; s; "--"); "--"))',
                    '',
                    # Scostamento Variante Computo
                    f'=LET(s; IF(M{n}="--"; 0; VALUE(M{n})) - IF(L{n}="--"; 0; VALUE(L{n})); IF(s; s; "--"))',
                    # Scostamento Contabilità Computo
                    f'=LET(s; IF(N{n}="--"; 0; VALUE(N{n})) - IF(L{n}="--"; 0; VALUE(L{n})); IF(s; s; "--"))',
                    # Scostamento Variante Computo
                    f'=LET(s; IF(N{n}="--"; 0; VALUE(N{n})) - IF(M{n}="--"; 0; VALUE(M{n})); IF(s; s; "--"))',
                    '',
                    # Importi Computo
                    f'=LET( s; IF(L{n}="--"; 0; VALUE(L{n}))*E{n}; risultato; IF(C{n}="%"; s/100; s);IF(risultato; risultato; "") )',
                    # Importi Variante
                    f'=LET( s; IF(M{n}="--"; 0; VALUE(M{n}))*E{n}; risultato; IF(C{n}="%"; s/100; s);IF(risultato; risultato; "") )',
                    # Importi Contabilità
                    f'=LET( s; IF(N{n}="--"; 0; VALUE(N{n}))*E{n}; risultato; IF(C{n}="%"; s/100; s);IF(risultato; risultato; "") )',

                ]
            formule.append(stringa)
        indicator.end()
        oRange = oSheet.getCellRangeByPosition(11, 3, 21, LeenoSheetUtils.cercaUltimaVoce(oSheet))
        formule = tuple(formule)
        oRange.setFormulaArray(formule)
########################################################################

def riordina_ElencoPrezzi():
    MENU_riordina_ElencoPrezzi()

def MENU_riordina_ElencoPrezzi():
    """
    Riordina l'Elenco Prezzi secondo l'ordine alfabetico dei codici di prezzo.
    """
    with LeenoUtils.DocumentRefreshContext(False):

        oDoc = LeenoUtils.getDocument()

        # Ottieni il foglio di lavoro
        oSheet = oDoc.CurrentController.ActiveSheet

        if oSheet.Name != 'Elenco Prezzi':
            return

        if SheetUtils.uFindStringCol('Fine elenco', 0, oSheet) is None:
            LeenoSheetUtils.inserisciRigaRossa(oSheet)
        
        last_row = str(SheetUtils.uFindStringCol('Fine elenco', 0, oSheet) + 1)
        SheetUtils.NominaArea(oDoc, 'Elenco Prezzi', f"$A$3:$AF${last_row}", 'elenco_prezzi')
        SheetUtils.NominaArea(oDoc, 'Elenco Prezzi', f"$A$3:$A${last_row}", 'Lista')
        oRangeAddress = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress

        start_row = oRangeAddress.StartRow + 1
        start_col = 0
        end_col = oRangeAddress.EndColumn
        end_row = oRangeAddress.EndRow - 1

        if start_row == end_row:
            return

        def ordina_intervallo(sheet, start_row, end_row, start_col, end_col):
            """
            Ordina un intervallo di celle per la prima colonna.
            """
            oRange = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
            SheetUtils.simpleSortColumn(oRange, 0, True)

        try:
            # Trova il limite della prima sezione
            costo_elem_row = SheetUtils.uFindStringCol('ELENCO DEI COSTI ELEMENTARI', 1, oSheet)
            
            if costo_elem_row is None:
                # Ordina tutto l'intervallo se la stringa non viene trovata
                ordina_intervallo(oSheet, start_row, end_row, start_col, end_col)
            else:
                # Ordina la prima sezione
                ordina_intervallo(oSheet, start_row, costo_elem_row - 1, start_col, end_col)
                # Ordina la seconda sezione
                ordina_intervallo(oSheet, costo_elem_row + 1, end_row, start_col, end_col)
        except Exception as e:
            DLG.errore(e)
        Menu_adattaAltezzaRiga()
########################################################################


def MENU_doppioni():
    # Inizializza la progress bar
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    indicator = oDoc.getCurrentController().getStatusIndicator()
    if indicator:
        indicator.start("Elaborazione in corso...", 100)  # 100 = max progresso
    
    try:
        with LeenoUtils.DocumentRefreshContext(False):
            inizializza_elenco()
            # Fase 1: Elimina doppioni (20%)
            if indicator:
                indicator.Text = "Eliminazione voci doppie..."
                indicator.Value = 20
            EliminaVociDoppieElencoPrezzi()

            # Fase 2: Analisi EP (40%)
            if indicator:
                indicator.Text = "Analisi in corso..."
                indicator.Value = 40
            tante_analisi_in_ep()

            # Fase 3: Genera sommario (60%)
            if indicator:
                indicator.Text = "Generazione sommario..."
                indicator.Value = 60
            genera_sommario()
            LeenoSheetUtils.setLarghezzaColonne(oSheet)

            # Fase 4: Riordina EP (80%)
            if indicator:
                indicator.Text = "Riordino Elenco Prezzi..."
                indicator.Value = 80
            riordina_ElencoPrezzi()

            # Fase 5: Sistema stili (100%)
            if indicator:
                indicator.Text = "Applicazione stili..."
                indicator.Value = 100
            sistema_stili()

        LeenoSheetUtils.adattaAltezzaRiga()

    finally:
        if indicator:
            indicator.end()  # Chiude la progress bar


# def EliminaVociDoppieElencoPrezzi():
#     """
#     Rimuove dall'elenco prezzi:
#     1. Voci duplicate (stessa chiave fino a MAX_COMPARE_COLS, esclusa colonna 1)
#        - Per duplicati: mantiene righe con markup (col5) o prima riga
#     2. Tutte le voci che contengono "(AP)" nella colonna 8
#     3. Voci già presenti in Analisi
#     """
#     LeenoSheetUtils.memorizza_posizione()

#     # --- CONFIGURAZIONE ---
#     MARKUP_COL = 5             # Colonna markup/note
#     EXCLUDE_COL = 8            # Colonna per esclusione "(AP)"
#     MAX_COMPARE_COLS = 5       # Colonne per confronto duplicati
#     TOTAL_COLS = 14            # Totale colonne elenco
#     AP_FLAG = "(AP)"           # Testo da cercare per esclusione

#     oDoc = LeenoUtils.getDocument()
#     # LeenoUtils.DocumentRefresh(False)

#     # Recupera dati da Elenco Prezzi
#     # if not oDoc.NamedRanges.hasByName('elenco_prezzi'):
#     #     LeenoUtils.DocumentRefresh(True)
#     #     return

#     # oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
#     oSheet = oDoc.CurrentController.ActiveSheet
#     named_range = oDoc.NamedRanges.getByName('elenco_prezzi').ReferredCells.RangeAddress
#     start_row, end_row = named_range.StartRow + 1, named_range.EndRow - 1

#     if end_row <= start_row:
#         # LeenoUtils.DocumentRefresh(True)
#         return
#     # Lettura dati
#     data_range = oSheet.getCellRangeByPosition(0, start_row, TOTAL_COLS - 1, end_row)
#     full_data = data_range.getDataArray()

#     # Filtraggio dati
#     filtered_data = []
#     for row in full_data:
#         # Escludi se contiene "(AP)" nella colonna 8
#         if EXCLUDE_COL < len(row) and AP_FLAG in str(row[EXCLUDE_COL]):
#             continue
            
#         filtered_data.append(row)

#     # Eliminazione duplicati
#     groups = OrderedDict()
#     for row in filtered_data:
#         key = tuple(row[i] for i in range(MAX_COMPARE_COLS) if i != 1)
#         if key not in groups:
#             groups[key] = []
#         groups[key].append(row)

#     clean_data = []
#     for rows in groups.values():
#         with_markup = [r for r in rows if r[MARKUP_COL] not in ('', None)]
        
#         if with_markup:
#             clean_data.append(with_markup[0])
#         else:
#             clean_data.append(rows[0])

#     # Scrittura risultati
#     if len(clean_data) != len(full_data):
#         oSheet.getRows().removeByIndex(start_row, end_row - start_row + 1)
#         oSheet.getRows().insertByIndex(start_row, len(clean_data))
        
#         output_range = oSheet.getCellRangeByPosition(
#             0, start_row, 
#             TOTAL_COLS - 1, 
#             start_row + len(clean_data) - 1
#         )
#         output_range.setDataArray(clean_data)
#     # LeenoUtils.DocumentRefresh(True)
#     LeenoSheetUtils.ripristina_posizione()

def EliminaVociDoppieElencoPrezzi():
    """
    Funzione per la pulizia dell'elenco prezzi che:
    1. Rimuove i duplicati basandosi sulle prime MAX_COMPARE_COLS colonne (esclusa colonna 1)
       - Conserva le righe con markup (colonna 5) se presenti, altrimenti mantiene la prima occorrenza
    2. Elimina tutte le voci contenenti "(AP)" nella colonna 8
    3. Preserva l'integrità dei dati mantenendo la struttura originale
    
    Restituisce:
        bool: True se l'operazione è riuscita, False in caso di errore
    """
    try:
        # Memorizza la posizione corrente per ripristinarla al termine
        LeenoSheetUtils.memorizza_posizione()

        # --- CONFIGURAZIONE ---
        COLONNA_MARKUP = 5          # Colonna contenente markup/note
        COLONNA_ESCLUSIONE = 8      # Colonna per l'esclusione voci con "(AP)"
        MAX_COLONNE_CONFRONTO = 6   # Numero colonne per il controllo duplicati
        TOTALE_COLONNE = 14         # Numero totale colonne nell'elenco
        TESTO_ESCLUSIONE = "(AP)"    # Testo da cercare per escludere voci

        # Ottieni il documento e il foglio attivo
        oDoc = LeenoUtils.getDocument()
        indicator = oDoc.getCurrentController().getStatusIndicator()
        foglio = oDoc.CurrentController.ActiveSheet
        
        # Verifica che esista l'intervallo nominato 'elenco_prezzi'
        if not oDoc.NamedRanges.hasByName('elenco_prezzi'):
            raise Exception("Intervallo nominato 'elenco_prezzi' non trovato")
            
        # Ottieni l'intervallo di dati da elaborare
        oSheet = oDoc.NamedRanges.getByName('elenco_prezzi').ReferredCells.RangeAddress
        riga_inizio, riga_fine = oSheet.StartRow + 1, oSheet.EndRow - 1

        # Se non ci sono dati da elaborare, termina
        if riga_fine <= riga_inizio:
            return True

        # Leggi tutti i dati dall'elenco prezzi
        intervallo_dati = foglio.getCellRangeByPosition(
            0, riga_inizio, 
            TOTALE_COLONNE - 1, riga_fine
        )
        
        # Inizializza la progress bar
        total_rows = riga_fine - riga_inizio + 1
        indicator.start("Pulizia elenco prezzi in corso...", total_rows)
        
        dati_completi = intervallo_dati.getDataArray()
        indicator.setValue(10)  # 10% completato dopo lettura dati

        # Filtra i dati rimuovendo le voci con "(AP)"
        dati_filtrati = []
        for idx, riga in enumerate(dati_completi):
            # Aggiorna la progress bar
            progress = 10 + int(30 * idx / len(dati_completi))
            indicator.setValue(progress)
            
            # Escludi la riga se contiene "(AP)" nella colonna specificata
            if len(riga) > COLONNA_ESCLUSIONE and TESTO_ESCLUSIONE in str(riga[COLONNA_ESCLUSIONE]):
                continue
            dati_filtrati.append(riga)

        indicator.setValue(40)  # 40% completato dopo filtro

        # Elimina i duplicati
        gruppi = OrderedDict()
        for idx, riga in enumerate(dati_filtrati):
            # Aggiorna la progress bar
            progress = 40 + int(30 * idx / len(dati_filtrati))
            indicator.setValue(progress)
            
            # Crea una chiave unica escludendo la colonna 1
            chiave = tuple(riga[i] for i in range(MAX_COLONNE_CONFRONTO) if i != 1)
            if chiave not in gruppi:
                gruppi[chiave] = []
            gruppi[chiave].append(riga)

        indicator.setValue(70)  # 70% completato dopo raggruppamento

        # Prepara i dati puliti
        dati_puliti = []
        for idx, gruppo in enumerate(gruppi.values()):
            # Aggiorna la progress bar
            progress = 70 + int(20 * idx / len(gruppi))
            indicator.setValue(progress)
            
            # Cerca righe con markup
            righe_con_markup = [r for r in gruppo if r[COLONNA_MARKUP] not in ('', None)]
            # Aggiungi la riga con markup se esiste, altrimenti la prima del gruppo
            dati_puliti.append(righe_con_markup[0] if righe_con_markup else gruppo[0])

        indicator.setValue(90)  # 90% completato dopo preparazione dati

        # Scrivi i risultati solo se ci sono stati cambiamenti
        if len(dati_puliti) != len(dati_completi):
            # Rimuovi e reinserisci le righe per mantenere la formattazione
            foglio.getRows().removeByIndex(riga_inizio, riga_fine - riga_inizio + 1)
            foglio.getRows().insertByIndex(riga_inizio, len(dati_puliti))
            
            # Scrivi i dati puliti
            intervallo_output = foglio.getCellRangeByPosition(
                0, riga_inizio, 
                TOTALE_COLONNE - 1, 
                riga_inizio + len(dati_puliti) - 1
            )
            intervallo_output.setDataArray(dati_puliti)
            
        indicator.setValue(100)  # 100% completato
        indicator.end()  # Chiudi la progress bar
        
        # Ripristina la posizione originale
        LeenoSheetUtils.ripristina_posizione()
        return True
    
    except Exception as e:
        if 'indicator' in locals():
            indicator.end()  # Assicurati di chiudere la progress bar in caso di errore
        DLG.chi(f"Errore durante l'elaborazione: {str(e)}")
        LeenoSheetUtils.ripristina_posizione()
        return False

########################################################################
# Scrive un file.
def XPWE_out(elaborato, out_file):
    with LeenoUtils.DocumentRefreshContext(False):
        XPWE_out_run(elaborato, out_file)

def XPWE_out_run(elaborato, out_file):
    '''
    esporta il documento in formato XPWE

    elaborato { string } : nome del foglio da esportare
    out_file  { string } : nome base del file

    il nome file risulterà out_file-elaborato.xpwe
    '''
    if cfg.read('Generale', 'dettaglio') == '1':
        # dettaglio = 1
        cfg.write('Generale', 'dettaglio', '0')
        dettaglio_misure(0)
    # else:
    #     dettaglio = 0

    oDoc = LeenoUtils.getDocument()
    # oDoc.enableAutomaticCalculation(False)
    if cfg.read('Generale', 'dettaglio') == '1':
        dettaglio_misure(0)
    numera_voci(1)
    top = Element('PweDocumento')
    #  intestazioni
    CopyRight = SubElement(top, 'CopyRight')
    CopyRight.text = 'Copyright ACCA software S.p.A.'
    TipoDocumento = SubElement(top, 'TipoDocumento')
    TipoDocumento.text = '1'
    # impostando in TipoDocumento.text a 2, in Primus
    # si abilitano funzionalità altrimenti indisponibili
    if elaborato == 'CONTABILITA':
        TipoDocumento.text = '2'
    if TipoDocumento.text != '2':
        if Dialogs.YesNoDialog(
            Title='',
            Text= 'Abilitando la contabilità nel formato XPWE,\n'
            'Primus potrà riconoscere e gestire correttamente le Voci della Sicurezza.\n\n'
            'Vuoi abilitare la contabilità nel file esportato?') == 1:
            TipoDocumento.text = '2'

    # attiva la progressbar
    indicator = oDoc.getCurrentController().getStatusIndicator()
    if indicator:
        indicator.start(f'Esportazione di {elaborato} in corso...', 7)  # max progresso
        indicator.Text = f'Esportazione di {elaborato} in corso...'
        

    TipoFormato = SubElement(top, 'TipoFormato')
    TipoFormato.text = 'XMLPwe'
    Versione = SubElement(top, 'Versione')
    Versione.text = ''
    SourceVersione = SubElement(top, 'SourceVersione')

    release = (
       str(LeenoUtils.getGlobalVar('Lmajor')) + '.' +
       str(LeenoUtils.getGlobalVar('Lminor')) + '.' +
       LeenoUtils.getGlobalVar('Lsubv')
    )

    SourceVersione.text = release
    SourceNome = SubElement(top, 'SourceNome')
    SourceNome.text = 'LeenO.org'
    FileNameDocumento = SubElement(top, 'FileNameDocumento')
    #  dati generali
    PweDatiGenerali = SubElement(top, 'PweDatiGenerali')
    PweMisurazioni = SubElement(top, 'PweMisurazioni')
    PweDGProgetto = SubElement(PweDatiGenerali, 'PweDGProgetto')
    PweDGDatiGenerali = SubElement(PweDGProgetto, 'PweDGDatiGenerali')
    PercPrezzi = SubElement(PweDGDatiGenerali, 'PercPrezzi')
    PercPrezzi.text = '0'

    Comune = SubElement(PweDGDatiGenerali, 'Comune')
    Provincia = SubElement(PweDGDatiGenerali, 'Provincia')
    Oggetto = SubElement(PweDGDatiGenerali, 'Oggetto')
    Committente = SubElement(PweDGDatiGenerali, 'Committente')
    Impresa = SubElement(PweDGDatiGenerali, 'Impresa')
    ParteOpera = SubElement(PweDGDatiGenerali, 'ParteOpera')
    #   leggo i dati generali
    oSheet = oDoc.getSheets().getByName('S2')
    Comune.text = oSheet.getCellByPosition(2, 3).String
    Provincia.text = ''
    Oggetto.text = oSheet.getCellByPosition(2, 2).String
    Committente.text = oSheet.getCellByPosition(2, 5).String
    Impresa.text = oSheet.getCellByPosition(2, 16).String
    ParteOpera.text = ''
    #  Capitoli e Categorie
    PweDGCapitoliCategorie = SubElement(PweDatiGenerali,
                                        'PweDGCapitoliCategorie')
    #  SuperCategorie
    oSheet = oDoc.getSheets().getByName(elaborato)
    lastRow = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1
    # evito di esportare in SuperCategorie perché inutile, almeno per ora
    listaspcat = []
    PweDGSuperCategorie = SubElement(PweDGCapitoliCategorie,
                                     'PweDGSuperCategorie')
    indicator.Value = 1
    for n in range(0, lastRow):
        if oSheet.getCellByPosition(1, n).CellStyle == 'Livello-0-scritta':
            desc = oSheet.getCellByPosition(2, n).String
            if desc not in listaspcat:
                listaspcat.append(desc)
                idID = str(listaspcat.index(desc) + 1)

                #  PweDGSuperCategorie = SubElement(PweDGCapitoliCategorie,'PweDGSuperCategorie')
                DGSuperCategorieItem = SubElement(PweDGSuperCategorie,
                                                  'DGSuperCategorieItem')
                DesSintetica = SubElement(DGSuperCategorieItem, 'DesSintetica')

                DesEstesa = SubElement(DGSuperCategorieItem, 'DesEstesa')
                DataInit = SubElement(DGSuperCategorieItem, 'DataInit')
                Durata = SubElement(DGSuperCategorieItem, 'Durata')
                # CodFase = SubElement(DGSuperCategorieItem, 'CodFase')
                Percentuale = SubElement(DGSuperCategorieItem, 'Percentuale')
                # Codice = SubElement(DGSuperCategorieItem, 'Codice')

                DGSuperCategorieItem.set('ID', idID)
                DesSintetica.text = desc
                DataInit.text = oggi()
                Durata.text = '0'
                Percentuale.text = '0'

#  Categorie
    listaCat = []
    PweDGCategorie = SubElement(PweDGCapitoliCategorie, 'PweDGCategorie')
    indicator.Value = 2
    for n in range(0, lastRow):
        if oSheet.getCellByPosition(2,
                                    n).CellStyle == 'Livello-1-scritta mini':
            desc = oSheet.getCellByPosition(2, n).String
            if desc not in listaCat:
                listaCat.append(desc)
                idID = str(listaCat.index(desc) + 1)

                #  PweDGCategorie = SubElement(PweDGCapitoliCategorie,'PweDGCategorie')
                DGCategorieItem = SubElement(PweDGCategorie, 'DGCategorieItem')
                DesSintetica = SubElement(DGCategorieItem, 'DesSintetica')

                DesEstesa = SubElement(DGCategorieItem, 'DesEstesa')
                DataInit = SubElement(DGCategorieItem, 'DataInit')
                Durata = SubElement(DGCategorieItem, 'Durata')
                # CodFase = SubElement(DGCategorieItem, 'CodFase')
                Percentuale = SubElement(DGCategorieItem, 'Percentuale')
                # Codice = SubElement(DGCategorieItem, 'Codice')

                DGCategorieItem.set('ID', idID)
                DesSintetica.text = desc
                DataInit.text = oggi()
                Durata.text = '0'
                Percentuale.text = '0'

#  SubCategorie
    listasbCat = []
    PweDGSubCategorie = SubElement(PweDGCapitoliCategorie, 'PweDGSubCategorie')
    indicator.Value = 3
    for n in range(0, lastRow):
        if oSheet.getCellByPosition(2, n).CellStyle == 'livello2_':
            desc = oSheet.getCellByPosition(2, n).String
            if desc not in listasbCat:
                listasbCat.append(desc)
                idID = str(listasbCat.index(desc) + 1)

                #  PweDGSubCategorie = SubElement(PweDGCapitoliCategorie,'PweDGSubCategorie')
                DGSubCategorieItem = SubElement(PweDGSubCategorie,
                                                'DGSubCategorieItem')
                DesSintetica = SubElement(DGSubCategorieItem, 'DesSintetica')

                DesEstesa = SubElement(DGSubCategorieItem, 'DesEstesa')
                DataInit = SubElement(DGSubCategorieItem, 'DataInit')
                Durata = SubElement(DGSubCategorieItem, 'Durata')
                # CodFase = SubElement(DGSubCategorieItem, 'CodFase')
                Percentuale = SubElement(DGSubCategorieItem, 'Percentuale')
                # Codice = SubElement(DGSubCategorieItem, 'Codice')

                DGSubCategorieItem.set('ID', idID)
                DesSintetica.text = desc
                DataInit.text = oggi()
                Durata.text = '0'
                Percentuale.text = '0'

#  Moduli
    PweDGModuli = SubElement(PweDatiGenerali, 'PweDGModuli')
    PweDGAnalisi = SubElement(PweDGModuli, 'PweDGAnalisi')
    SpeseUtili = SubElement(PweDGAnalisi, 'SpeseUtili')
    SpeseGenerali = SubElement(PweDGAnalisi, 'SpeseGenerali')
    UtiliImpresa = SubElement(PweDGAnalisi, 'UtiliImpresa')
    OneriAccessoriSc = SubElement(PweDGAnalisi, 'OneriAccessoriSc')
    # ConfQuantita = SubElement(PweDGAnalisi, 'ConfQuantita')

    oSheet = oDoc.getSheets().getByName('S1')
    if oSheet.getCellByPosition(
            7, 322).Value == 0:  # se 0: Spese e Utili Accorpati
        SpeseUtili.text = '1'
    else:
        SpeseUtili.text = '-1'

    UtiliImpresa.text = oSheet.getCellByPosition(7, 320).String[:-1].replace(
        ',', '.')
    OneriAccessoriSc.text = oSheet.getCellByPosition(7,
                                                     318).String[:-1].replace(
                                                         ',', '.')
    SpeseGenerali.text = oSheet.getCellByPosition(7, 319).String[:-1].replace(
        ',', '.')

    #  Configurazioni
    PU = str(len(LeenoFormat.getFormatString('comp 1-a PU').split(',')[-1]))
    LUN = str(len(LeenoFormat.getFormatString('comp 1-a LUNG').split(',')[-1]))
    LAR = str(len(LeenoFormat.getFormatString('comp 1-a LARG').split(',')[-1]))
    PES = str(len(LeenoFormat.getFormatString('comp 1-a peso').split(',')[-1]))
    QUA = str(len(LeenoFormat.getFormatString('Blu').split(',')[-1]))
    PR = str(len(LeenoFormat.getFormatString('comp sotto Unitario').split(',')[-1]))
    TOT = str(len(LeenoFormat.getFormatString('An-1v-dx').split(',')[-1]))
    PweDGConfigurazione = SubElement(PweDatiGenerali, 'PweDGConfigurazione')
    PweDGConfigNumeri = SubElement(PweDGConfigurazione, 'PweDGConfigNumeri')
    Divisa = SubElement(PweDGConfigNumeri, 'Divisa')
    Divisa.text = 'euro'
    ConversioniIN = SubElement(PweDGConfigNumeri, 'ConversioniIN')
    ConversioniIN.text = 'lire'
    FattoreConversione = SubElement(PweDGConfigNumeri, 'FattoreConversione')
    FattoreConversione.text = '1936.27'
    Cambio = SubElement(PweDGConfigNumeri, 'Cambio')
    Cambio.text = '1'
    PartiUguali = SubElement(PweDGConfigNumeri, 'PartiUguali')
    PartiUguali.text = '9.' + PU + '|0'
    Lunghezza = SubElement(PweDGConfigNumeri, 'Lunghezza')
    Lunghezza.text = '9.' + LUN + '|0'
    Larghezza = SubElement(PweDGConfigNumeri, 'Larghezza')
    Larghezza.text = '9.' + LAR + '|0'
    HPeso = SubElement(PweDGConfigNumeri, 'HPeso')
    HPeso.text = '9.' + PES + '|0'
    Quantita = SubElement(PweDGConfigNumeri, 'Quantita')
    Quantita.text = '10.' + QUA + '|1'
    Prezzi = SubElement(PweDGConfigNumeri, 'Prezzi')
    Prezzi.text = '10.' + PR + '|1'
    PrezziTotale = SubElement(PweDGConfigNumeri, 'PrezziTotale')
    PrezziTotale.text = '14.' + TOT + '|1'
    ConvPrezzi = SubElement(PweDGConfigNumeri, 'ConvPrezzi')
    ConvPrezzi.text = '11.0|1'
    ConvPrezziTotale = SubElement(PweDGConfigNumeri, 'ConvPrezziTotale')
    ConvPrezziTotale.text = '15.0|1'
    IncidenzaPercentuale = SubElement(PweDGConfigNumeri,
                                      'IncidenzaPercentuale')
    IncidenzaPercentuale.text = '7.3|0'
    Aliquote = SubElement(PweDGConfigNumeri, 'Aliquote')
    Aliquote.text = '7.3|0'

    # if dettaglio == 1:
    #     dettaglio_misure(1)
    #     cfg.write('Generale', 'dettaglio', '1')
#  Elenco Prezzi
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    PweElencoPrezzi = SubElement(PweMisurazioni, 'PweElencoPrezzi')
    diz_ep = {}
    lista_AP = []
    indicator.Value = 4
    listaspcap = []
    listacap = []
    listasbcap = []
    #  giallo(16777072,16777120,16777168)
    for n in range(3, SheetUtils.getUsedArea(oSheet).EndRow):

      #  SuperCapitoli
        if oSheet.getCellByPosition(0, n).CellBackColor == 16777072 and \
        oSheet.getCellByPosition(0, n).String != '000':
            cod = oSheet.getCellByPosition(0, n).String
            desc = oSheet.getCellByPosition(1, n).String
            if desc not in listaspcap:
                listaspcap.append(desc)
                IDSpCap = str(listaspcap.index(desc) + 1)

                PweDGSuperCapitoli = SubElement(PweDGCapitoliCategorie,'PweDGSuperCapitoli')
                DGSuperCapitoliItem = SubElement(PweDGSuperCapitoli,
                                                  'DGSuperCapitoliItem')
                DesSintetica = SubElement(DGSuperCapitoliItem, 'DesSintetica')

                DesEstesa = SubElement(DGSuperCapitoliItem, 'DesEstesa')
                DataInit = SubElement(DGSuperCapitoliItem, 'DataInit')
                Durata = SubElement(DGSuperCapitoliItem, 'Durata')
                # CodFase = SubElement(DGSuperCapitoliItem, 'CodFase')
                Percentuale = SubElement(DGSuperCapitoliItem, 'Percentuale')
                Codice = SubElement(DGSuperCapitoliItem, 'Codice')

                DGSuperCapitoliItem.set('ID', IDSpCap)
                DesSintetica.text = desc
                Codice.text = cod
                DataInit.text = '' #oggi()
                Durata.text = '0'
                Percentuale.text = '0'

      #  Capitoli
        if oSheet.getCellByPosition(0, n).CellBackColor == 16777120:
            cod = oSheet.getCellByPosition(0, n).String
            desc = oSheet.getCellByPosition(1, n).String
            if desc not in listacap:
                listacap.append(desc)
                IDCap = str(listacap.index(desc) + 1)

                PweDGCapitoli = SubElement(PweDGCapitoliCategorie,'PweDGCapitoli')
                DGCapitoliItem = SubElement(PweDGCapitoli,
                                                  'DGCapitoliItem')
                DesSintetica = SubElement(DGCapitoliItem, 'DesSintetica')

                DesEstesa = SubElement(DGCapitoliItem, 'DesEstesa')
                DataInit = SubElement(DGCapitoliItem, 'DataInit')
                Durata = SubElement(DGCapitoliItem, 'Durata')
                # CodFase = SubElement(DGCapitoliItem, 'CodFase')
                Percentuale = SubElement(DGCapitoliItem, 'Percentuale')
                Codice = SubElement(DGCapitoliItem, 'Codice')

                DGCapitoliItem.set('ID', IDCap)
                DesSintetica.text = desc
                Codice.text = cod
                DataInit.text = '' #oggi()
                Durata.text = '0'
                Percentuale.text = '0'

      #  SubCapitoli
        if oSheet.getCellByPosition(0, n).CellBackColor == 16777168:
            cod = oSheet.getCellByPosition(0, n).String
            desc = oSheet.getCellByPosition(1, n).String
            if desc not in listasbcap:
                listasbcap.append(desc)
                IDSbCap = str(listasbcap.index(desc) + 1)

                PweDGSubCapitoli = SubElement(PweDGCapitoliCategorie,'PweDGSubCapitoli')
                DGSubCapitoliItem = SubElement(PweDGSubCapitoli,
                                                  'DGSubCapitoliItem')
                DesSintetica = SubElement(DGSubCapitoliItem, 'DesSintetica')

                DesEstesa = SubElement(DGSubCapitoliItem, 'DesEstesa')
                DataInit = SubElement(DGSubCapitoliItem, 'DataInit')
                Durata = SubElement(DGSubCapitoliItem, 'Durata')
                # CodFase = SubElement(DGSubCapitoliItem, 'CodFase')
                Percentuale = SubElement(DGSubCapitoliItem, 'Percentuale')
                Codice = SubElement(DGSubCapitoliItem, 'Codice')

                DGSubCapitoliItem.set('ID', IDSbCap)
                DesSintetica.text = desc
                Codice.text = cod
                DataInit.text = '' #oggi()
                Durata.text = '0'
                Percentuale.text = '0'


    #voci di prezzo
        if(oSheet.getCellByPosition(1, n).Type.value == 'FORMULA' and
           oSheet.getCellByPosition(2, n).Type.value == 'FORMULA'):
            lista_AP.append(oSheet.getCellByPosition(0, n).String)
        elif(oSheet.getCellByPosition(1, n).Type.value == 'TEXT' and
             oSheet.getCellByPosition(2, n).Type.value == 'TEXT'):
            EPItem = SubElement(PweElencoPrezzi, 'EPItem')
            EPItem.set('ID', str(n))
            TipoEP = SubElement(EPItem, 'TipoEP')
            TipoEP.text = '0'
            Tariffa = SubElement(EPItem, 'Tariffa')
            id_tar = str(n)
            Tariffa.text = oSheet.getCellByPosition(0, n).String
            diz_ep[oSheet.getCellByPosition(0, n).String] = id_tar
            Articolo = SubElement(EPItem, 'Articolo')
            Articolo.text = ''
            DesRidotta = SubElement(EPItem, 'DesRidotta')
            DesEstesa = SubElement(EPItem, 'DesEstesa')
            DesEstesa.text = oSheet.getCellByPosition(1, n).String
            if len(DesEstesa.text) > 120:
                DesRidotta.text = DesEstesa.text[:
                                                 60] + ' ... ' + DesEstesa.text[
                                                     -60:]
            else:
                DesRidotta.text = DesEstesa.text
            DesBreve = SubElement(EPItem, 'DesBreve')
            if len(DesEstesa.text) > 60:
                DesBreve.text = DesEstesa.text[:30] + ' ... ' + DesEstesa.text[
                    -30:]
            else:
                DesBreve.text = DesEstesa.text
            UnMisura = SubElement(EPItem, 'UnMisura')
            UnMisura.text = oSheet.getCellByPosition(2, n).String
            Prezzo1 = SubElement(EPItem, 'Prezzo1')
            Prezzo1.text = str(oSheet.getCellByPosition(4, n).Value)
            Prezzo2 = SubElement(EPItem, 'Prezzo2')
            Prezzo2.text = '0'
            Prezzo3 = SubElement(EPItem, 'Prezzo3')
            Prezzo3.text = '0'
            Prezzo4 = SubElement(EPItem, 'Prezzo4')
            Prezzo4.text = '0'
            Prezzo5 = SubElement(EPItem, 'Prezzo5')
            Prezzo5.text = '0'

            try:
                SubElement(EPItem, 'IDSpCap').text = IDSpCap
            except:
                SubElement(EPItem, 'IDSpCap').text = '0'
            try:
                SubElement(EPItem, 'IDCap').text = IDCap
            except:
                SubElement(EPItem, 'IDCap').text = '0'
            try:
                SubElement(EPItem, 'IDSbCap').text = IDSbCap
            except:
                SubElement(EPItem, 'IDSbCap').text = '0'

            Flags = SubElement(EPItem, 'Flags')
            if oSheet.getCellByPosition(8, n).String == '(AP)':
                Flags.text = '131072'
            elif 'VDS_' in oSheet.getCellByPosition(0, n).String:
                Flags.text = '134217728'
                Tariffa.text = Tariffa.text.split('VDS_')[-1]
            else:
                Flags.text = '0'
            Data = SubElement(EPItem, 'Data')
            Data.text = '30/12/1899'
            AdrInternet = SubElement(EPItem, 'AdrInternet')
            AdrInternet.text = ''
            PweEPAnalisi = SubElement(EPItem, 'PweEPAnalisi')

            IncSIC = SubElement(EPItem, 'IncSIC')
            if oSheet.getCellByPosition(3, n).Value == 0.0:
                IncSIC.text = ''
            else:
                IncSIC.text = str(oSheet.getCellByPosition(3, n).Value * 100)

            IncMDO = SubElement(EPItem, 'IncMDO')
            if oSheet.getCellByPosition(5, n).Value == 0.0:
                IncMDO.text = ''
            else:
                IncMDO.text = str(oSheet.getCellByPosition(5, n).Value * 100)

            IncMAT = SubElement(EPItem, 'IncMAT')
            if oSheet.getCellByPosition(6, n).Value == 0.0:
                IncMAT.text = ''
            else:
                IncMAT.text = str(oSheet.getCellByPosition(6, n).Value * 100)

            IncATTR = SubElement(EPItem, 'IncATTR')
            if oSheet.getCellByPosition(7, n).Value == 0.0:
                IncATTR.text = ''
            else:
                IncATTR.text = str(oSheet.getCellByPosition(7, n).Value * 100)

    # Analisi di prezzo
    indicator.Value = 5
    if len(lista_AP) != 0:
        lista_AP = list(set(lista_AP))
        oSheet = oDoc.getSheets().getByName('Analisi di Prezzo')
        k = n + 1
        for el in lista_AP:
            try:
                m = SheetUtils.uFindStringCol(el, 0, oSheet)
                EPItem = SubElement(PweElencoPrezzi, 'EPItem')
                EPItem.set('ID', str(k))
                TipoEP = SubElement(EPItem, 'TipoEP')
                TipoEP.text = '0'
                Tariffa = SubElement(EPItem, 'Tariffa')
                id_tar = str(k)
                Tariffa.text = oSheet.getCellByPosition(0, m).String
                diz_ep[oSheet.getCellByPosition(0, m).String] = id_tar
                Articolo = SubElement(EPItem, 'Articolo')
                Articolo.text = ''
                DesRidotta = SubElement(EPItem, 'DesRidotta')
                DesEstesa = SubElement(EPItem, 'DesEstesa')
                DesEstesa.text = oSheet.getCellByPosition(1, m).String
                if len(DesEstesa.text) > 120:
                    DesRidotta.text = DesEstesa.text[:
                                                     60] + ' ... ' + DesEstesa.text[
                                                         -60:]
                else:
                    DesRidotta.text = DesEstesa.text
                DesBreve = SubElement(EPItem, 'DesBreve')
                if len(DesEstesa.text) > 60:
                    DesBreve.text = DesEstesa.text[:
                                                   30] + ' ... ' + DesEstesa.text[
                                                       -30:]
                else:
                    DesBreve.text = DesEstesa.text
                UnMisura = SubElement(EPItem, 'UnMisura')
                UnMisura.text = oSheet.getCellByPosition(2, m).String
                Prezzo1 = SubElement(EPItem, 'Prezzo1')
                Prezzo1.text = str(oSheet.getCellByPosition(6, m).Value)
                Prezzo2 = SubElement(EPItem, 'Prezzo2')
                Prezzo2.text = '0'
                Prezzo3 = SubElement(EPItem, 'Prezzo3')
                Prezzo3.text = '0'
                Prezzo4 = SubElement(EPItem, 'Prezzo4')
                Prezzo4.text = '0'
                Prezzo5 = SubElement(EPItem, 'Prezzo5')
                Prezzo5.text = '0'
                IDSpCap = SubElement(EPItem, 'IDSpCap')
                IDSpCap.text = '0'
                IDCap = SubElement(EPItem, 'IDCap')
                IDCap.text = '0'
                IDSbCap = SubElement(EPItem, 'IDSbCap')
                IDSbCap.text = '0'
                Flags = SubElement(EPItem, 'Flags')
                Flags.text = '131072'
                Data = SubElement(EPItem, 'Data')
                Data.text = '30/12/1899'
                AdrInternet = SubElement(EPItem, 'AdrInternet')
                AdrInternet.text = ''
                PweEPAnalisi = SubElement(EPItem, 'PweEPAnalisi')
                PweEPAR = SubElement(PweEPAnalisi, 'PweEPAR')
                nEPARItem = 2
                for x in range(m, m + 100):
                    if oSheet.getCellByPosition(
                            0, x).CellStyle == 'An-lavoraz-desc':
                        EPARItem = SubElement(PweEPAR, 'EPARItem')
                        EPARItem.set('ID', str(nEPARItem))
                        nEPARItem += 1
                        Tipo = SubElement(EPARItem, 'Tipo')
                        Tipo.text = '0'
                        IDEP = SubElement(EPARItem, 'IDEP')
                        IDEP.text = diz_ep.get(
                            oSheet.getCellByPosition(0, x).String)
                        if IDEP.text is None:
                            IDEP.text = '-2'
                        Descrizione = SubElement(EPARItem, 'Descrizione')
                        if '=IF(' in oSheet.getCellByPosition(1, x).String:
                            Descrizione.text = ''
                        else:
                            Descrizione.text = oSheet.getCellByPosition(
                                1, x).String
                        Misura = SubElement(EPARItem, 'Misura')
                        Misura.text = ''
                        Qt = SubElement(EPARItem, 'Qt')
                        Qt.text = ''
                        Prezzo = SubElement(EPARItem, 'Prezzo')
                        Prezzo.text = ''
                        FieldCTL = SubElement(EPARItem, 'FieldCTL')
                        FieldCTL.text = '0'
                    if(oSheet.getCellByPosition(0, x).CellStyle == 'An-lavoraz-Cod-sx' and
                       oSheet.getCellByPosition(1, x).String != ''):
                        EPARItem = SubElement(PweEPAR, 'EPARItem')
                        EPARItem.set('ID', str(nEPARItem))
                        nEPARItem += 1
                        Tipo = SubElement(EPARItem, 'Tipo')
                        Tipo.text = '1'
                        IDEP = SubElement(EPARItem, 'IDEP')
                        IDEP.text = diz_ep.get(
                            oSheet.getCellByPosition(0, x).String)
                        if IDEP.text is None:
                            IDEP.text = '-2'
                        Descrizione = SubElement(EPARItem, 'Descrizione')
                        if '=IF(' in oSheet.getCellByPosition(1, x).String:
                            Descrizione.text = ''
                        else:
                            Descrizione.text = oSheet.getCellByPosition(
                                1, x).String
                        Misura = SubElement(EPARItem, 'Misura')
                        Misura.text = oSheet.getCellByPosition(2, x).String
                        Qt = SubElement(EPARItem, 'Qt')
                        Qt.text = oSheet.getCellByPosition(3,
                                                           x).String.replace(
                                                               ',', '.')
                        Prezzo = SubElement(EPARItem, 'Prezzo')
                        Prezzo.text = str(
                            oSheet.getCellByPosition(4, x).Value).replace(
                                ',', '.')
                        FieldCTL = SubElement(EPARItem, 'FieldCTL')
                        FieldCTL.text = '0'
                    elif oSheet.getCellByPosition(
                            0, x).CellStyle == 'An-sfondo-basso Att End':
                        break

                IncSIC = SubElement(EPItem, 'IncSIC')
                if oSheet.getCellByPosition(10, m).Value == 0.0:
                    IncSIC.text = ''
                else:
                    IncSIC.text = str(oSheet.getCellByPosition(10, m).Value)

                IncMDO = SubElement(EPItem, 'IncMDO')
                # oDoc.CurrentController.select(oSheet.getCellByPosition(8, m))
                # DLG.chi(oSheet.getCellByPosition(8, m).AbsoluteName)
                if oSheet.getCellByPosition(8, m).Value == 0.0:
                    IncMDO.text = ''
                else:
                    IncMDO.text = str(
                        oSheet.getCellByPosition(8, m).Value * 100)
                k += 1
            except Exception:
                pass

    if elaborato == 'Elenco_Prezzi':
        pass
    else:
        # COMPUTO/VARIANTE/CONTABILITA
        oSheet = oDoc.getSheets().getByName(elaborato)
        PweVociComputo = SubElement(PweMisurazioni, 'PweVociComputo')
        oDoc.CurrentController.setActiveSheet(oSheet)
        Rinumera_TUTTI_Capitoli2(oSheet)
        nVCItem = 2
        indicator.Value = 6
        indicator.start(f'Esportazione di {elaborato} in corso...', LeenoSheetUtils.cercaUltimaVoce(oSheet))  # max progresso

        for n in range(0, LeenoSheetUtils.cercaUltimaVoce(oSheet)):
            indicator.Value = n            
            if oSheet.getCellByPosition(0,
                                        n).CellStyle in ('Comp Start Attributo',
                                                         'Comp Start Attributo_R'):
                sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, n)
                sStRange.RangeAddress
                sopra = sStRange.RangeAddress.StartRow
                sotto = sStRange.RangeAddress.EndRow
                if elaborato == 'CONTABILITA':
                    sotto -= 1

                voce = LeenoComputo.datiVoceComputo(oSheet, sopra) # voce = (num, art, desc, um, quantP, prezzo, importo, sic, mdo)

                VCItem = SubElement(PweVociComputo, 'VCItem')
                VCItem.set('ID', str(nVCItem))
                nVCItem += 1

                IDEP = SubElement(VCItem, 'IDEP')
                IDEP.text = diz_ep.get(
                    oSheet.getCellByPosition(1, sopra + 1).String)
                ##########################
                Quantita = SubElement(VCItem, 'Quantita')
                Quantita.text = oSheet.getCellByPosition(9, sotto).String
                ##########################
                DataMis = SubElement(VCItem, 'DataMis')
                if elaborato == 'CONTABILITA':
                    DataMis.text = oSheet.getCellByPosition(1, sopra + 2).String
                else:
                    DataMis.text = oggi()  # 26/12/1952'#'28/09/2013'###
                vFlags = SubElement(VCItem, 'Flags')
                vFlags.text = '0'
                if 'VDS_' in voce[1]:
                    vFlags.text = '134217728'
                ##########################
                IDSpCat = SubElement(VCItem, 'IDSpCat')
                IDSpCat.text = str(oSheet.getCellByPosition(31, sotto).String)
                if elaborato == 'CONTABILITA':
                    IDSpCat.text = str(oSheet.getCellByPosition(31, sotto + 1).String)
                if IDSpCat.text == '':
                    IDSpCat.text = '0'
                # #########################
                IDCat = SubElement(VCItem, 'IDCat')
                IDCat.text = str(oSheet.getCellByPosition(32, sotto).String)
                if elaborato == 'CONTABILITA':
                    IDCat.text = str(oSheet.getCellByPosition(32, sotto + 1).String)
                if IDCat.text == '':
                    IDCat.text = '0'
                # #########################
                IDSbCat = SubElement(VCItem, 'IDSbCat')
                IDSbCat.text = str(oSheet.getCellByPosition(33, sotto).String)
                if elaborato == 'CONTABILITA':
                    IDSbCat.text = str(oSheet.getCellByPosition(33, sotto + 1).String)
                if IDSbCat.text == '':
                    IDSbCat.text = '0'
                # #########################
                PweVCMisure = SubElement(VCItem, 'PweVCMisure')
                x = 2
                for m in range(sopra + 2, sotto):
                    RGItem = SubElement(PweVCMisure, 'RGItem')
                    RGItem.set('ID', str(x))
                    x += 1
                    # #########################
                    IDVV = SubElement(RGItem, 'IDVV')
                    IDVV.text = '-2'
                    ##########################
                    Descrizione = SubElement(RGItem, 'Descrizione')
                    Descrizione.text = oSheet.getCellByPosition(2, m).String
                    # #########################
                    PartiUguali = SubElement(RGItem, 'PartiUguali')
                    PartiUguali.text = valuta_cella(oSheet.getCellByPosition(5, m))
                    # #########################
                    Lunghezza = SubElement(RGItem, 'Lunghezza')
                    Lunghezza.text = valuta_cella(oSheet.getCellByPosition(6, m))
                    # #########################
                    Larghezza = SubElement(RGItem, 'Larghezza')
                    Larghezza.text = valuta_cella(oSheet.getCellByPosition(7, m))
                    # #########################
                    HPeso = SubElement(RGItem, 'HPeso')
                    HPeso.text = valuta_cella(oSheet.getCellByPosition(8, m))
                    # #########################
                    Quantita = SubElement(RGItem, 'Quantita')
                    Quantita.text = str(oSheet.getCellByPosition(9, m).Value)
                    # se negativa in CONTABILITA:
                        # quando vedi_voce guarda ad un valore negativo
                    if oSheet.getCellByPosition(4, m).Value < 0:
                        test = True
                    if elaborato == 'CONTABILITA':
                        if oSheet.getCellByPosition(11, m).Value != 0:
                            Quantita.text = '-' + str(oSheet.getCellByPosition(11, m).Value)
                    # #########################
                    Flags = SubElement(RGItem, 'Flags')
                    if '*** VOCE AZZERATA ***' in Descrizione.text:
                        PartiUguali.text = str(
                            abs(float(valuta_cella(oSheet.getCellByPosition(5,
                                                                            m)))))
                        Flags.text = '1'
                    elif '-' in Quantita.text or oSheet.getCellByPosition(
                            11, m).Value != 0:
                        Flags.text = '1'
                    elif "Parziale [" in oSheet.getCellByPosition(8, m).String:
                        Flags.text = '2'
                        HPeso.text = ''
                    elif 'PARTITA IN CONTO PROVVISORIO' in Descrizione.text or \
                        'PARTITA PROVVISORIA' in Descrizione.text:
                        Flags.text = '16'
                    else:
                        Flags.text = '0'
                    # #########################
                    if 'DETRAE LA PARTITA IN CONTO PROVVISORIO' in Descrizione.text or \
                        'SI DETRAE PARTITA PROVVISORIA' in Descrizione.text:
                        Flags.text = '32'
                    if '- vedi voce n.' in Descrizione.text:
                        IDVV.text = str(
                            int(
                                Descrizione.text.split('- vedi voce n.')[1].split(
                                    ' ')[0]) + 1)
                        Flags.text = '32768'
                        Descrizione.text = ''
                        #  PartiUguali.text =''
                        if oSheet.getCellByPosition(4, m).Value < 0 and \
                            oSheet.getCellByPosition(11, m).Value != 0:
                                Flags.text = '32768'
                        if oSheet.getCellByPosition(4, m).Value > 0 and \
                            oSheet.getCellByPosition(11, m).Value != 0:
                                Flags.text = '32769'
                        if oSheet.getCellByPosition(4, m).Value > 0 and \
                            oSheet.getCellByPosition(10, m).Value != 0:
                                Flags.text = '32768'
                        if elaborato in ('COMPUTO', 'VARIANTE'):
                            if  oSheet.getCellByPosition(9, m).Value < 0:
                                Flags.text = '32769'
                n = sotto + 1

    # #########################
    # out_file = Dialogs.FileSelect('Salva con nome...', '*.xpwe', 1)
    # out_file = uno.fileUrlToSystemPath(oDoc.getURL())
    # DLG.mri (uno.fileUrlToSystemPath(oDoc.getURL()))
    # chi(out_file)
    if cfg.read('Generale', 'dettaglio') == '1':
        dettaglio_misure(1)
    try:
        if out_file.split('.')[-1].upper() != 'XPWE':
            out_file = out_file + '-' + elaborato + '.xpwe'
        FileNameDocumento.text = out_file
    except AttributeError:
        return
    riga = str(tostring(top, encoding="unicode"))
    #  if len(lista_AP) != 0:
    #  riga = riga.replace('<PweDatiGenerali>','<Fgs>131072</Fgs><PweDatiGenerali>')
    indicator.end()
    try:
        of = codecs.open(out_file, 'w', 'utf-8')
        of.write(riga)
        of.close()
        Dialogs.Exclamation(Title = 'INFORMAZIONE',

        Text=f'Esportazione in formato XPWE eseguita con successo sul file:\n\n {out_file}'
        '\n\n----\n\n'
        'Il formato XPWE è un formato XML di interscambio per Primus di ACCA.\n\n'
        'Prima di utilizzare questo file in Primus, assicurarsi che le percentuali\n'
        'di Spese Generali e Utile d\'Impresa siano impostate correttamente, in modo da\n'
        'garantire la corretta elaborazione dei dati.')
    except IOError:
        Dialogs.Exclamation(Title = 'E R R O R E !',
            Text='''               Esportazione non eseguita!
Verifica che il file di destinazione non sia già in uso!''')

########################################################################

def MENU_firme_in_calce():
    with LeenoUtils.DocumentRefreshContext(False):
        firme_in_calce(lrowF=None)

def firme_in_calce(lrowF=None):
    '''
    Inserisce(in COMPUTO o VARIANTE) un riepilogo delle categorie
    ed i dati necessari alle firme
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet_S2 = oDoc.getSheets().getByName('S2')

    datafirme = oSheet_S2.getCellRangeByName('$S2.C4').String.split(' ')[-1]

    if datafirme == "":
        datafirme="Data,"
    else:
        datafirme = datafirme + ", "

    if oSheet.Name == "CONTABILITA":
        if lrowF == None:
            lrowF = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2
        oSheet.getRows().insertByIndex(lrowF, 11)
        riga_corrente = lrowF + 1

    # INSERISCI LA DATA E IL PROGETTISTA
        # DLG.chi(datafirme)
        oSheet.getCellByPosition(2 , riga_corrente).Formula = (
            '=CONCATENATE("' + datafirme + '";TEXT(NOW();"GG/mm/aaaa"))')

        oSheet.getCellByPosition(2 , riga_corrente + 2).Formula = (
            "L'Impresa esecutrice\n(" + oSheet_S2.getCellByPosition(
                2, 16).String + ")")
        oSheet.getCellByPosition(2 , riga_corrente + 6).Formula = (
            "Il Direttore dei Lavori\n(" + oSheet_S2.getCellByPosition(
                2, 15).String + ")")
        comando('CalculateHard')
# rem CONSOLIDA LA DATA
        oRange = oSheet.getCellRangeByPosition (2, riga_corrente, 40, riga_corrente)
        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)

    if oSheet.Name in ("Registro", "SAL"):
        if lrowF == None:
            lrowF = SheetUtils.getLastUsedRow(oSheet)

        oSheet.getRows().insertByIndex(lrowF, 13)
        riga_corrente = lrowF + 1
        oSheet.getCellByPosition(1 , riga_corrente).Formula = '=CONCATENATE("' + datafirme + '";TEXT(NOW();"GG/mm/aaaa"))'
        comando('CalculateHard')
        oRange = oSheet.getCellRangeByPosition (1, riga_corrente, 40, riga_corrente)
        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)

        oSheet.getCellByPosition(1, riga_corrente + 2).Formula = (
            "L'Impresa esecutrice\n(" + oSheet_S2.getCellRangeByName(
                '$S2.C17').String + ")")

        oSheet.getCellByPosition(1, riga_corrente + 6).Formula = (
            "Il Direttore dei Lavori\n(" + oSheet_S2.getCellRangeByName(
            '$S2.C16').String + ")")
        oSheet.getCellRangeByPosition (0, riga_corrente + 2, 5,riga_corrente + 6).Rows.OptimalHeight = True
        if oSheet.Name == "SAL":
            return
        nSal = 1
        for i in reversed(range(2, 50)):
            if oDoc.NamedRanges.hasByName("_Lib_" + str(i)):
                nSal = i
                break
        oSheet.getCellByPosition(1, riga_corrente + 10).Formula = (
            # '=CONCATENATE("In data ";TEXT(NOW();"DD/MM/YYYY");" è stato emesso il CERTIFICATO DI PAGAMENTO n.' + str(nSal) + ' per un importo di €")')
            '=CONCATENATE("In data __/__/____ è stato emesso il CERTIFICATO DI PAGAMENTO n.' + str(nSal) + ' per un importo di €")')
        comando('CalculateHard')

        oRange = oSheet.getCellRangeByPosition (1, riga_corrente + 10, 40, riga_corrente + 10)

        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)

        oSheet.getCellByPosition(1 , riga_corrente + 12).Formula = (
            "Il Direttore dei Lavori\n(" + oSheet_S2.getCellRangeByName(
                '$S2.C16').String + ")")

    # if oSheet.Name in ('Analisi di Prezzo', 'Elenco Prezzi'):
    #     # Configurazione iniziale
    #     if lrowF is None:
    #         lrowF = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1
        
    #     oDoc.CurrentController.setFirstVisibleRow(lrowF - 1)
        
    #     # Trova ultima riga da processare
    #     used_area = SheetUtils.getUsedArea(oSheet)
    #     lrowE = used_area.EndRow
        
    #     # Cerca riga rossa di chiusura
    #     for i in range(lrowF, lrowE + 1):
    #         if oSheet.getCellByPosition(0, i).CellStyle == "Riga_rossa_Chiudi":
    #             lrowE = i
    #             break
    #     # Elimina righe se necessario
    #     if lrowE > lrowF + 1:
    #         oSheet.getRows().removeByIndex(lrowF, lrowE - lrowF)
        
    #     # Inserimento nuove righe
    #     NUM_RIGHE = 15
    #     # riga_corrente = lrowF + 1
    #     riga_corrente = lrowF
    #     first_new_row = lrowF -1
    #     last_new_row = lrowF + NUM_RIGHE - 1


    #     oSheet.getRows().insertByIndex(lrowF -1, NUM_RIGHE)
        
    #     # Formattazione celle
    #     oSheet.getCellRangeByPosition(0, first_new_row, 25, last_new_row -1).CellStyle = "Ultimus_centro"
    #     oSheet.getCellRangeByPosition(11, last_new_row, 25, last_new_row).CellStyle = "Comp-Bianche in mezzo Descr_R"
        
    #     # Raggruppamento righe
    #     range_addr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    #     range_addr.Sheet = oSheet.RangeAddress.Sheet
    #     range_addr.StartColumn = 0
    #     range_addr.EndColumn = 0
    #     range_addr.StartRow = first_new_row
    #     range_addr.EndRow = last_new_row - 1
    #     oSheet.group(range_addr, 1)
        
    #     # Inserimento dati
    #     # Data
    #     data_row = riga_corrente + 3
    #     data_cell = oSheet.getCellByPosition(1, data_row)
    #     data_cell.Formula = '=CONCATENATE("Data, ";TEXT(NOW();"GG/MM/AAAA"))'
    #     comando('CalculateHard')
        
    #     # Consolida formula della data
    #     data_array = data_cell.getDataArray()
    #     data_cell.setDataArray(data_array)
    #     oSheet.getCellRangeByPosition(1, data_row, 1, data_row).CellStyle = 'ULTIMUS'
        
    #     # Tecnico
    #     oSheet.getCellByPosition(1, riga_corrente + 5).Formula = 'Il Tecnico'
    #     oSheet.getCellByPosition(1, riga_corrente + 6).Formula = '=CONCATENATE($S2.$C$13)'
    if oSheet.Name in ('Analisi di Prezzo', 'Elenco Prezzi'):
        if lrowF == None:
            lrowF = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1
        oDoc.CurrentController.setFirstVisibleRow(lrowF - 1)
        lrowE = SheetUtils.getUsedArea(oSheet).EndRow
        for i in range(lrowF, SheetUtils.getUsedArea(oSheet).EndRow + 1):
            if oSheet.getCellByPosition(0, i).CellStyle == "Riga_rossa_Chiudi":
                lrowE = i
                break
        if lrowE > lrowF + 1:
            oSheet.getRows().removeByIndex(lrowF, lrowE - lrowF)
        riga_corrente = lrowF + 1
        oSheet.getRows().insertByIndex(lrowF, 15)
        oSheet.getCellRangeByPosition(0, lrowF, 25, lrowF + 15 -
                                      1).CellStyle = "Ultimus_centro"
        oSheet.getCellRangeByPosition(0, lrowF + 15 - 1, 25, lrowF + 15 -
                                      1).CellStyle = "Comp-Bianche in mezzo Descr_R"
        # raggruppo i righi di mirura
        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct(
            'com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        oCellRangeAddr.StartColumn = 0
        oCellRangeAddr.EndColumn = 0
        oCellRangeAddr.StartRow = lrowF
        oCellRangeAddr.EndRow = lrowF + 15 - 1
        oSheet.group(oCellRangeAddr, 1)

        # INSERISCI LA DATA E IL PROGETTISTA
        oSheet.getCellByPosition(
            1, riga_corrente +
            3).Formula = '=CONCATENATE("Data, ";TEXT(NOW();"GG/MM/AAAA"))'
        comando('CalculateHard')
        #  consolido il risultato
        oRange = oSheet.getCellByPosition(1, riga_corrente + 3)
        # flags = (oDoc.createInstance('com.sun.star.sheet.CellFlags.FORMULA'))
        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)
        oSheet.getCellRangeByPosition(1, riga_corrente + 3, 1,
                                      riga_corrente + 3).CellStyle = 'ULTIMUS'
        oSheet.getCellByPosition(1,
                                 riga_corrente + 5).Formula = 'Il Tecnico'
        oSheet.getCellByPosition(
            1, riga_corrente + 6
        ).Formula = '=CONCATENATE($S2.$C$13)'  # senza concatenate, se la cella di origine è vuota il risultato è '0,00'

    if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CompuM_NoP'):
        if lrowF == None:
            lrowF = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2
        oDoc.CurrentController.setFirstVisibleRow(lrowF - 2)
        lrowE = SheetUtils.getUsedArea(oSheet).EndRow
        for i in range(lrowF, SheetUtils.getUsedArea(oSheet).EndRow + 1):
            if oSheet.getCellByPosition(0, i).CellStyle == "Riga_rossa_Chiudi":
                lrowE = i
                break
        if lrowE > lrowF + 1:
            oSheet.getRows().removeByIndex(lrowF, lrowE - lrowF)
        riga_corrente = lrowF + 2
        if oDoc.getSheets().hasByName('S2'):
            ii = 11
            vv = 18
            ac = 28
            ad = 29
            ae = 30
            ss = 41
            col = 'S'

        else:
            ii = 8
            vv = 9
            ss = 9
            col = 'J'
        oSheet.getRows().insertByIndex(lrowF, 17)
        oSheet.getCellRangeByPosition(0, lrowF, ss,
                                      lrowF + 17 - 1).CellStyle = 'ULTIMUS'
        # raggruppo i righi di mirura
        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct(
            'com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        oCellRangeAddr.StartColumn = 0
        oCellRangeAddr.EndColumn = 0
        oCellRangeAddr.StartRow = lrowF
        oCellRangeAddr.EndRow = lrowF + 17 - 1
        oSheet.group(oCellRangeAddr, 1)

        #  INSERIMENTO TITOLO
        oSheet.getCellByPosition(
            2, riga_corrente).String = 'Riepilogo strutturale delle Categorie'
        oSheet.getCellByPosition(ii, riga_corrente).String = 'Incidenze %'
        oSheet.getCellByPosition(vv, riga_corrente).String = 'Importi €'
        oSheet.getCellByPosition(ac,
                                 riga_corrente).String = 'Materiali\ne Noli €'
        oSheet.getCellByPosition(ad, riga_corrente).String = 'Incidenza\nMDO %'
        oSheet.getCellByPosition(ae, riga_corrente).String = 'Importo\nMDO €'

        #  TITOLI PER CRONOPROGRAMMA
        oSheet.getCellByPosition(45, riga_corrente -1).CellStyle = 'comp 1-a peso'
        # costo_medio_mdo = cfg.read('Computo', 'costo_medio_mdo')
        # if costo_medio_mdo == '':
        #     costo_medio_mdo = '0,00'
        # oSheet.getCellByPosition(45, riga_corrente -1).Value = float(costo_medio_mdo.replace(',', '.'))

        oSheet.Annotations.insertNew(oSheet.getCellByPosition(45, riga_corrente -1).CellAddress,
    'Inserisci il costo orario medio della Manodopera.')
        oSheet.getCellByPosition(47, riga_corrente -1).CellStyle = 'comp 1-a peso'
        # addetti_mdo = cfg.read('Computo', 'addetti_mdo')
        # oSheet.getCellByPosition(47, riga_corrente -1).Value = addetti_mdo

        oSheet.Annotations.insertNew(oSheet.getCellByPosition(47, riga_corrente -1).CellAddress,
    'Inserisci il numero di componenti della squadra di operai.')
        nOperai = "AV$" + str(riga_corrente)
        mdo = riga_corrente
        oSheet.getCellByPosition(45, riga_corrente).Formula = '=concatenate("MDO\nmedia\n€ ";AT' + str(mdo) + ';"/h")'
        oSheet.getCellByPosition(46, riga_corrente).Formula = '=concatenate("EC U/GG.\nMDO media\n€ ";AT' + str(mdo) + ';"/h")'
        oSheet.getCellByPosition(47, riga_corrente).String = '(Gl)\nGiorni\nlavorativi'
        oSheet.getCellByPosition(48, riga_corrente).String = '(Gs)\nGiorni\nimproduttivi'
        oSheet.getCellByPosition(49, riga_corrente).String = 'GG\nTempo di\nesecuzione'
        oSheet.getCellRangeByPosition(45, riga_corrente,49,
                                      riga_corrente).CellStyle = 'Ultimus_centro'
        inizio_gruppo = riga_corrente
        riga_corrente += 1

    # attiva la progressbar
        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start("Composizione del riepilogo strutturale...", lrowF)

        categorie = {}
        for i in range(0, lrowF):
            indicator.Value = i

            if oSheet.getCellByPosition(1, i).CellStyle == 'Livello-0-scritta':
                categorie[oSheet.getCellByPosition(1, i).String] = oSheet.getCellByPosition(2, i).String

            elif oSheet.getCellByPosition(1, i).CellStyle == 'Livello-1-scritta':
                 categorie[oSheet.getCellByPosition(1, i).String] = oSheet.getCellByPosition(2, i).String

            elif oSheet.getCellByPosition(1, i).CellStyle == 'livello2 valuta':
                categorie[oSheet.getCellByPosition(1, i).String] = oSheet.getCellByPosition(2, i).String

        # Funzione per ordinare le chiavi in ordine naturale
        def chiave_naturale(chiave):
            return [int(parte) if parte.isdigit() else float(parte) for parte in chiave.split('.')]

        # Ordinamento del dizionario
        categorie_ordinate = dict(sorted(categorie.items(), key=lambda item: chiave_naturale(item[0])))

        def applica_stili_e_formule(sheet, row, parametri):
            # Applicazione degli stili
            sheet.getCellRangeByPosition(0, row, 49, row).CellStyle = parametri['range_style']
            sheet.getCellByPosition(1, row).CellStyle = parametri['cell_style_1']
            sheet.getCellByPosition(parametri['ii'], row).CellStyle = parametri['cell_style_ii']
            sheet.getCellByPosition(parametri['vv'], row).CellStyle = parametri['cell_style_vv']
            sheet.getCellByPosition(parametri['ad'], row).CellStyle = parametri['cell_style_ad']
            
            # Formule cronoprogramma
            sheet.getCellByPosition(45, row).Formula = f"=AE{row+1}/AT{parametri['mdo']}"
            sheet.getCellByPosition(46, row).Formula = f"=AT{row+1}/8"
            sheet.getCellByPosition(47, row).Formula = f"=AU{row+1}/{parametri['n_operai']}"
            sheet.getCellByPosition(48, row).Formula = f"=AW{row+2}/AV{row+2}*AV{row+1}"
            sheet.getCellByPosition(49, row).Formula = f"=SUM(AW{row+1}:AV{row+1})"

        # Parametri per i tre livelli
        parametri_livello1 = {
            'range_style': 'ULTIMUS_1',
            'cell_style_1': 'Ultimus_destra_1',
            'cell_style_ii': 'Ultimus %_1',
            'cell_style_vv': 'Ultimus_totali_1',
            'cell_style_ad': 'Ultimus %_1',
            'mdo': mdo,
            'n_operai': nOperai,
            'ii': ii,
            'vv': vv,
            'ad': ad,
        }

        parametri_livello2 = {
            'range_style': 'ULTIMUS_2',
            'cell_style_1': 'Ultimus_destra',
            'cell_style_ii': 'Ultimus %',
            'cell_style_vv': 'Ultimus_bordo',
            'cell_style_ad': 'Ultimus %',
            'mdo': mdo,
            'n_operai': nOperai,
            'ii': ii,
            'vv': vv,
            'ad': ad,
        }

        parametri_livello3 = {
            'range_style': 'ULTIMUS_3',
            'cell_style_1': 'Ultimus_destra_3',
            'cell_style_ii': 'Ultimus %_3',
            'cell_style_vv': 'ULTIMUS_3',
            'cell_style_ad': 'Ultimus %_3',
            'mdo': mdo,
            'n_operai': nOperai,
            'ii': ii,
            'vv': vv,
            'ad': ad,
        }

        # Applicazione
        oRow = SheetUtils.uFindString("TOTALI COMPUTO", oSheet)[1] +1

        for key, value in categorie_ordinate.items():
            oSheet.getRows().insertByIndex(riga_corrente, 1)
            oSheet.getCellByPosition(1, riga_corrente).String = key
            oSheet.getCellByPosition(2, riga_corrente).String = value
            oSheet.getCellByPosition(11, riga_corrente).Formula = f'=S{riga_corrente + 1}/S{oRow}*100'
            oSheet.getCellByPosition(18, riga_corrente).Formula = f'=SUMIF($B$2:$B${lrowF}; B{riga_corrente + 1}; $S$2:$S${lrowF})'
            oSheet.getCellByPosition(29, riga_corrente).Formula = f'=AE{riga_corrente + 1}/S{riga_corrente + 1}*100'
            oSheet.getCellByPosition(30, riga_corrente).Formula = f'=SUMIF($B$2:$B${lrowF}; B{riga_corrente + 1}; AE$2:AE${lrowF})'

            livello = len(key.split('.'))
            if livello == 1:
                applica_stili_e_formule(oSheet, riga_corrente, parametri_livello1)
            elif livello == 2:
                applica_stili_e_formule(oSheet, riga_corrente, parametri_livello2)
            elif livello == 3:
                applica_stili_e_formule(oSheet, riga_corrente, parametri_livello3)

            riga_corrente += 1

        indicator.end()
        oSheet.getCellRangeByPosition(
            2, inizio_gruppo, ae, inizio_gruppo).CellStyle = "Ultimus_centro"
        oSheet.getCellByPosition(ii, riga_corrente).Value = 100
        oSheet.getCellByPosition(2, riga_corrente).CellStyle = 'Ultimus_destra'
        oSheet.getCellByPosition(ii, riga_corrente).CellStyle = 'Ultimus %_1'
        oSheet.getCellByPosition(
            vv, riga_corrente).Formula = '=' + col + str(lrowF)
        oSheet.getCellByPosition(
            vv, riga_corrente).CellStyle = 'Ultimus_Bordo_sotto'
        oSheet.getCellByPosition(ac,
                                 riga_corrente).Formula = '=AC' + str(lrowF)
        oSheet.getCellByPosition(
            ac, riga_corrente).CellStyle = 'Ultimus_Bordo_sotto'
        oSheet.getCellByPosition(ae,
                                 riga_corrente).Formula = '=AE' + str(lrowF)
        oSheet.getCellByPosition(
            ae, riga_corrente).CellStyle = 'Ultimus_Bordo_sotto'
        oSheet.getCellByPosition(
            ad, riga_corrente).Formula = '=AD' + str(lrowF) + '*100'
        oSheet.getCellByPosition(
            2, riga_corrente).String = '          T O T A L E   €'
        oSheet.getCellByPosition(2, riga_corrente).CellStyle = 'ULTIMUS_1'

        # lettura dati cronoprogramma
        oSheet.getCellRangeByPosition(
            45, riga_corrente, 49,
            riga_corrente).CellStyle = 'ULTIMUS_1'
        oSheet.getCellByPosition(
            45, riga_corrente).Formula = '=AE' + str(riga_corrente +1) + '/AT' + str(mdo)
        oSheet.getCellByPosition(
            46, riga_corrente).Formula = '=AT' + str(riga_corrente +1) + '/8'
        operai = oSheet.getCellByPosition(47, mdo).Value
        oSheet.getCellByPosition(
            47, riga_corrente).Formula = '=AU' + str(riga_corrente +1) + '/'+ nOperai
        
        oSheet.getCellByPosition(48, riga_corrente).CellStyle = 'comp 1-a peso'
        oSheet.getCellByPosition(48, riga_corrente).Formula = f'=AV{riga_corrente +1}*115/365'
        oSheet.Annotations.insertNew(oSheet.getCellByPosition(48, riga_corrente).CellAddress,
    "Giorni improduttivi previsti per il cantiere, riposi settimanali inclusi (indicativamente 115 giorni all’anno).")

        oSheet.getCellByPosition(
            49, riga_corrente).Formula = '=sum(AW' + str(riga_corrente +1) + ":AV" + str(riga_corrente +1) +')'

        # imposta 0 cifre decimali
        oDoc.CurrentController.select(oSheet.getCellRangeByPosition(46, mdo -1, 49, riga_corrente))
        ctx = LeenoUtils.getComponentContext()
        desktop = LeenoUtils.getDesktop()
        oFrame = desktop.getCurrentFrame()

        oProp = PropertyValue()
        oProp.Name = 'NumberFormatValue'
        oProp.Value = 1
        properties = (oProp, )
        dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
            'com.sun.star.frame.DispatchHelper', ctx)
        dispatchHelper.executeDispatch(oFrame, '.uno:NumberFormatValue', '', 0, properties)

        #  iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct(
            'com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        oCellRangeAddr.StartColumn = 44
        oCellRangeAddr.EndColumn = 49
        oCellRangeAddr.StartRow = mdo
        oCellRangeAddr.EndRow = riga_corrente
        oSheet.ungroup(oCellRangeAddr, 0)
        oSheet.group(oCellRangeAddr, 0)
        #  if flag == 1:

        #  else:
            #  oSheet.ungroup(oCellRangeAddr, 1)

        # fine_gruppo = riga_corrente
        #  DATA
        oSheet.getCellByPosition(
            2, riga_corrente +
            3).Formula = '=CONCATENATE("Data, ";TEXT(NOW();"GG/MM/AAAA"))'
        #  consolido il risultato
        oRange = oSheet.getCellByPosition(2, riga_corrente + 3)
        # flags = (oDoc.createInstance('com.sun.star.sheet.CellFlags.FORMULA'))
        comando('CalculateHard')
        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)

        oSheet.getCellByPosition(2,
                                 riga_corrente + 5).Formula = 'Il Tecnico'
        oSheet.getCellByPosition(
            2, riga_corrente + 6
        ).Formula = '=CONCATENATE($S2.$C$13)'  # senza concatenate, se la cella di origine è vuota il risultato è '0,00'
        oSheet.getCellRangeByPosition(2, riga_corrente + 5, 2, riga_corrente +
                                      6).CellStyle = 'Ultimus_centro'

        # inserisco il salto pagina in cima al riepilogo
        oDoc.CurrentController.select(oSheet.getCellByPosition(0, lrowF))
        ctx = LeenoUtils.getComponentContext()
        desktop = LeenoUtils.getDesktop()
        oFrame = desktop.getCurrentFrame()
        dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
            'com.sun.star.frame.DispatchHelper', ctx)
        dispatchHelper.executeDispatch(oFrame, ".uno:InsertRowBreak", "", 0, [])
        oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))

        #  oSheet.getCellByPosition(lrowF,0).Rows.IsManualPageBreak = True
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)

def gantt():
    # Ottieni il documento corrente e prepara il percorso del file di output
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    if oSheet.Name not in ("COMPUTO", "VARIANTE"):
        Dialogs.Exclamation(Title='Avviso!',
        Text= '''L'esportazione in formato CSV per GanttProject\npuò avvenire dal COMPUTO o dalla VARIANTE.    ''')
        #  GotoSheet("COMPUTO")
        return
    try:
        sRow = SheetUtils.uFindStringCol('Riepilogo', 2, oSheet, start=2, equal=0, up=False) + 1
    except Exception as e:
        Dialogs.Exclamation(Title='Informazione',
        Text= "L'esportazione in formato CSV può avvenire solo\nin presenza del Riepilogo strutturale delle Categorie.")
        return

    out_file = uno.fileUrlToSystemPath(oDoc.getURL()).rsplit('.', 1)[0] + '-' + oSheet.Name + '_gantt.csv'

    sRow = SheetUtils.uFindStringCol('Riepilogo', 2, oSheet, start=2, equal=0, up=False) + 1
    eRow = SheetUtils.uFindStringCol('T O T A L E', 2, oSheet, start=sRow, equal=0, up=False)
    dati = [(
        "ID", "Nome", "Data d'inizio", "Data di fine", "Durata", "Completamento",
        "Costo", "Coordinatore", "Predecessori", "Numero dello schema", "Risorse",
        "Assignments", "Colore attività", "Link Web", "Note"
    )]

    ID = 1

    nome = oSheet.getCellByPosition(2, eRow).String.replace("€","").replace(" ","")
    durata = int(oSheet.getCellByPosition(49, eRow).Value)
    costo = oSheet.getCellByPosition(18, eRow).String.replace(".","").replace(",",".")
    schema = oSheet.getCellByPosition(1, eRow).String
    dati.append((ID, nome, '', '', durata, '', costo, '', '', schema, '', '', '', '', ''))

    for r in range(sRow, eRow):
        ID += 1
        nome = oSheet.getCellByPosition(2, r).String
        durata = int(oSheet.getCellByPosition(49, r).Value)
        costo = oSheet.getCellByPosition(18, r).String.replace(".","").replace(",",".")
        schema = oSheet.getCellByPosition(1, r).String
        dati.append((ID, nome, '', '', durata, '', costo, '', '', schema, '', '', '', '', ''))

    # Scrivi i dati in un file CSV
    try:
        with open(out_file, 'w', newline='') as file:
            for row in dati:
                # Converti ogni tupla di riga in una stringa separata da virgole
                file.write(','.join(map(str, row)) + "\n")
    except Exception as e:
        Dialogs.Exclamation(Title='Avviso!',
        Text= f'''Errore: {e}\nPrima di esportazione nel formato CSV\nè necessario generare il riepilogo delle categoirie.''')
        return

    Dialogs.Info(Title = 'Avviso.',
    Text='Il file:\n\n' + out_file + '\n\nè pronto per essere importato in GanttProject.' )

    return
########################################################################

def cancella_analisi_da_ep():
    '''
    cancella le voci in Elenco Prezzi che derivano da analisi
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
    lista_an = []
    for i in range(0, SheetUtils.getUsedArea(oSheet).EndRow):
        if oSheet.getCellByPosition(0, i).CellStyle == 'An-1_sigla':
            # codice = oSheet.getCellByPosition(0, i).String
            lista_an.append(oSheet.getCellByPosition(0, i).String)
    oSheet = oDoc.Sheets.getByName('Elenco Prezzi')
    for i in reversed(range(0, SheetUtils.getUsedArea(oSheet).EndRow)):
        if oSheet.getCellByPosition(0, i).String in lista_an:
            oSheet.getRows().removeByIndex(i, 1)


def MENU_analisi_in_ElencoPrezzi():
    '''
    Invia l'analisi in Elenco Prezzi.
    '''
    oDoc = LeenoUtils.getDocument()
    try:
        oSheet = oDoc.CurrentController.ActiveSheet
        if oSheet.Name != 'Analisi di Prezzo':
            return
        oDoc.enableAutomaticCalculation(False)  # blocco il calcolo automatico
        sStRange = Circoscrive_Analisi(LeggiPosizioneCorrente()[1])
        riga = sStRange.RangeAddress.StartRow + 2

        codice = oSheet.getCellByPosition(0, riga - 1).String

        oSheet = oDoc.Sheets.getByName('Elenco Prezzi')
        oDoc.CurrentController.setActiveSheet(oSheet)

        oSheet.getRows().insertByIndex(4, 1)

        oSheet.getCellByPosition(0, 4).String = codice
        oSheet.getCellByPosition(
            1, 4).Formula = "=$'Analisi di Prezzo'.B" + str(riga)
        oSheet.getCellByPosition(
            2, 4).Formula = "=$'Analisi di Prezzo'.C" + str(riga)
        oSheet.getCellByPosition(
            3, 4).Formula = "=$'Analisi di Prezzo'.K" + str(riga)
        oSheet.getCellByPosition(
            4, 4).Formula = "=$'Analisi di Prezzo'.G" + str(riga)
        oSheet.getCellByPosition(
            5, 4).Formula = "=$'Analisi di Prezzo'.I" + str(riga)
        oSheet.getCellByPosition(
            6, 4).Formula = "=$'Analisi di Prezzo'.J" + str(riga)
        oSheet.getCellByPosition(
            7, 4).Formula = "=$'Analisi di Prezzo'.A" + str(riga)
        oSheet.getCellByPosition(8, 4).String = "(AP)"
        oSheet.getCellByPosition(11, 4).Formula = "=N4/$N$2"
        oSheet.getCellByPosition(12, 4).Formula = "=SUMIF(AA;A4;BB)"
        oSheet.getCellByPosition(13, 4).Formula = "=SUMIF(AA;A4;cEuro)"
        oDoc.enableAutomaticCalculation(True)  # sblocco il calcolo automatico
        _gotoCella(1, 4)
    except Exception:
        pass

    oDoc.enableAutomaticCalculation(True)
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)


########################################################################
def tante_analisi_in_ep():
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    # oDoc.enableAutomaticCalculation(False)  # Disabilita calcolo
    with LeenoUtils.DocumentRefreshContext(False):

        # 1. Prepara dati dalla sheet Analisi
        from com.sun.star.container import NoSuchElementException
        try:
            src_sheet = oDoc.getSheets().getByName('Analisi di prezzo')
        except NoSuchElementException:
            return
        last_row = SheetUtils.getUsedArea(src_sheet).EndRow
        SheetUtils.NominaArea(oDoc, 'Analisi di Prezzo', f'$A$3:$K${last_row}', 'analisi')
        
        # Estrai dati in blocco
        data_range = src_sheet.getCellRangeByPosition(0, 0, 10, last_row)
        data = data_range.getDataArray()
        
        # 2. Costruisci lista analisi
        lista_analisi = []
        target_row = 4  # Inizia dalla riga 4 in Elenco Prezzi
        
        for n, row in enumerate(data):
            if (n < len(data) and 
                src_sheet.getCellByPosition(0, n).CellStyle == 'An-1_sigla' and 
                (row[1] if len(row)>1 else '') != '<<<Scrivi la descrizione della nuova voce da analizzare   '):
                
                lista_analisi.append([
                    row[0] if len(row)>0 else '',  # Codice
                    f"=$'Analisi di Prezzo'.B{n+1}",
                    f"=$'Analisi di Prezzo'.C{n+1}",
                    f"=$'Analisi di Prezzo'.K{n+1}",
                    f"=$'Analisi di Prezzo'.G{n+1}",
                    f"=$'Analisi di Prezzo'.I{n+1}",
                    "",
                    f"=$'Analisi di Prezzo'.A{n+1}",
                    "(AP)",
                    '',
                    '',
                    f"=N{target_row+1}/$N$2",
                    f"=SUMIF(AA;A{target_row};BB)",
                    f"=SUMIF(AA;A{target_row};cEuro)"
                ])
                target_row += 1

        # 3. Inserisci in Elenco Prezzi
        if not lista_analisi:
            # oDoc.enableAutomaticCalculation(True)
            return
            
        dest_sheet = oDoc.getSheets().getByName('Elenco Prezzi')
        dest_sheet.getRows().insertByIndex(4, len(lista_analisi))
        
        # Scrivi in blocco
        oRange = dest_sheet.getCellRangeByPosition(0, 4, 13, 4 + len(lista_analisi) - 1)
        oRange.setDataArray(lista_analisi)
 
        oRange.clearContents(HARDATTR)
        
        # Converti formule (evitando la colonna 0)
        for y in range(4, 4 + len(lista_analisi)):
            for x in range(1, 14):  # Colonne da 1 a 13
                dest_sheet.getCellByPosition(x, y).Formula = dest_sheet.getCellByPosition(x, y).String
        
        # Finalizza
        # oDoc.enableAutomaticCalculation(True)
        GotoSheet('Elenco Prezzi')
        # LeenoSheetUtils.adattaAltezzaRiga(dest_sheet)

########################################################################
def Circoscrive_Analisi(lrow):
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoAnalysis.circoscriveAnalisi'
    lrow    { int }  : riga di riferimento per
                        la selezione dell'intera voce
    Circoscrive una voce di analisi
    partendo dalla posizione corrente del cursore
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    stili_analisi = LeenoUtils.getGlobalVar('stili_analisi')
    if oSheet.getCellByPosition(0, lrow).CellStyle in stili_analisi:
        for el in reversed(range(0, lrow)):
            # DLG.chi(oSheet.getCellByPosition(0, el).CellStyle)
            if oSheet.getCellByPosition(0, el).CellStyle == 'An.1v-Att Start':
                SR = el
                break
        for el in range(lrow, SheetUtils.getUsedArea(oSheet).EndRow + 1):
            if oSheet.getCellByPosition(0, el).CellStyle == 'Analisi_Sfondo':
                ER = el
                break
    celle = oSheet.getCellRangeByPosition(0, SR, 250, ER)
    return celle

########################################################################
def ColumnNumberToName(oSheet, cColumnNumb):
    '''Trasforma IDcolonna in Nome'''
    #  oDoc = LeenoUtils.getDocument()
    #  oSheet = oDoc.CurrentController.ActiveSheet
    oColumns = oSheet.getColumns()
    oColumn = oColumns.getByIndex(cColumnNumb).Name
    return oColumn


########################################################################
def ColumnNameToNumber(oSheet, cColumnName):
    '''Trasforma il nome colonna in IDcolonna'''
    #  oDoc = LeenoUtils.getDocument()
    #  oSheet = oDoc.CurrentController.ActiveSheet
    oColumns = oSheet.getColumns()
    oColumn = oColumns.getByName(cColumnName)
    oRangeAddress = oColumn.getRangeAddress()
    nColumn = oRangeAddress.StartColumn
    return nColumn


########################################################################
def MENU_azzera_voce():
    '''
    Azzera la quantità di una voce e ne raggruppa le relative righe
    '''
    oDoc = LeenoUtils.getDocument()
    LeenoUtils.DocumentRefresh(False)

    try:
        oSheet = oDoc.CurrentController.ActiveSheet
        if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
            try:
                sRow = oDoc.getCurrentSelection().getRangeAddresses(
                )[0].StartRow
                eRow = oDoc.getCurrentSelection().getRangeAddresses()[0].EndRow

            except Exception:
                sRow = oDoc.getCurrentSelection().getRangeAddress().StartRow
                eRow = oDoc.getCurrentSelection().getRangeAddress().EndRow
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, sRow)
            sStRange.RangeAddress
            sRow = sStRange.RangeAddress.StartRow
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, eRow)
            try:
                sStRange.RangeAddress
            except Exception:
                return
            inizio = sStRange.RangeAddress.StartRow
            eRow = sStRange.RangeAddress.EndRow + 1

            lrow = sRow
            fini = []
            for x in range(sRow, eRow):
                if oSheet.getCellByPosition(
                        0, x).CellStyle == 'Comp End Attributo':
                    fini.append(x)
                elif oSheet.getCellByPosition(
                        0, x).CellStyle == 'Comp End Attributo_R':
                    fini.append(x - 2)
        idx = 0
        for lrow in reversed(fini):
            lrow += idx
            try:
                sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
                sStRange.RangeAddress
                inizio = sStRange.RangeAddress.StartRow
                fine = sStRange.RangeAddress.EndRow
                if oSheet.Name == 'CONTABILITA':
                    fine -= 1
                _gotoCella(2, fine - 1)
                if '*** VOCE AZZERATA ***' in oSheet.getCellByPosition(2, fine - 1).String:
                    # elimino il colore di sfondo
                    if oSheet.Name == 'CONTABILITA':
                        oSheet.getCellRangeByPosition(
                            0, inizio, 250, fine + 1).clearContents(HARDATTR)
                    else:
                        oSheet.getCellRangeByPosition(
                            0, inizio, 250, fine).clearContents(HARDATTR)
                    raggruppa_righe_voce(lrow, 0)
                    oSheet.getRows().removeByIndex(fine - 1, 1)
                    fine -= 1
                    _gotoCella(2, fine - 1)
                    idx -= 1
                else:
                    Copia_riga_Ent()
                    oSheet.getCellByPosition(2, fine).String = '*** VOCE AZZERATA ***'
                    if oSheet.Name == 'CONTABILITA':
                        oSheet.getCellByPosition(
                            5, fine).Formula = '=SUBTOTAL(9;J' + str(
                                inizio + 1) + ':J' + str(
                                    fine) + ')-SUBTOTAL(9;L' + str(
                                        inizio + 1) + ':L' + str(fine) + ')'
                    else:
                        oSheet.getCellByPosition(
                            5, fine).Formula = '=SUBTOTAL(9;J' + str(
                                inizio + 1) + ':J' + str(fine) + ')'
                    inverti_segno()
                    # cambio il colore di sfondo
                    oDoc.CurrentController.select(sStRange)
                    raggruppa_righe_voce(lrow, 1)
                    ctx = LeenoUtils.getComponentContext()
                    desktop = LeenoUtils.getDesktop()
                    oFrame = desktop.getCurrentFrame()
                    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
                        'com.sun.star.frame.DispatchHelper', ctx)
                    oProp = PropertyValue()
                    oProp.Name = 'BackgroundColor'
                    oProp.Value = 15066597
                    properties = (oProp, )
                    dispatchHelper.executeDispatch(oFrame, '.uno:BackgroundColor', '', 0, properties)
                    _gotoCella(2, fine)
                    ###
                lrow = LeggiPosizioneCorrente()[1]
                lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 1)
            except Exception:
                pass
        # numera_voci(1)
    except Exception:
        pass
    #  _gotoCella(1, fine +3)
    LeenoUtils.DocumentRefresh(True)


########################################################################
def MENU_elimina_voci_azzerate():
    '''
    Elimina le voci in cui compare la dicitura '*** VOCE AZZERATA ***'
    in COMPUTO o in VARIANTE, senza chiedere conferma
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
            ER = SheetUtils.getUsedArea(oSheet).EndRow
                # attiva la progressbar
            # progress = Dialogs.Progress(Title='Esecuzione in corso...', Text="Cancellazione voci azzerate")
            indicator = oDoc.getCurrentController().getStatusIndicator()
            indicator.start("Cancellazione voci azzerate...", LeenoSheetUtils.cercaUltimaVoce(oSheet))
            n = 0
            # progress.setLimits(0, LeenoSheetUtils.cercaUltimaVoce(oSheet))
            indicator.setValue(n)
            for lrow in reversed(range(0, ER)):
                n += 1
                indicator.setValue(n)
                # if oSheet.getCellByPosition(
                        # 2, lrow).String == '*** VOCE AZZERATA ***':
                if '*** VOCE AZZERATA ***' in oSheet.getCellByPosition(2, lrow).String:
                    LeenoSheetUtils.eliminaVoce(oSheet, lrow)

            numera_voci(1)
            indicator.end()
    except Exception:
        return


########################################################################
def raggruppa_righe_voce(lrow, flag=1):
    '''
    Raggruppa le righe che compongono una singola voce.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #  lrow = LeggiPosizioneCorrente()[1]
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        sStRange.RangeAddress

        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct(
            'com.sun.star.table.CellRangeAddress')
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
def MENU_nasconde_voci_azzerate():
    '''
    Nasconde le voci in cui compare la dicitura '*** VOCE AZZERATA ***'
    in COMPUTO o in VARIANTE.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        if oSheet.Name in ('COMPUTO', 'VARIANTE'):
            ER = SheetUtils.getUsedArea(oSheet).EndRow
            for lrow in reversed(range(0, ER)):
                if '*** VOCE AZZERATA ***' in oSheet.getCellByPosition(2, lrow).String:
                    raggruppa_righe_voce(lrow, 1)
    except Exception:
        return


########################################################################
def seleziona(lrow=None):
    '''
    Seleziona voci intere
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    if lrow == None:
        lrow = LeggiPosizioneCorrente()[1]

        try:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
        except AttributeError:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
    if oSheet.Name in ('Elenco Prezzi'):

        el_y = []
        lista_y = []
        try:
            len(oRangeAddress)
            for el in oRangeAddress:
                el_y.append((el.StartRow, el.EndRow))
        except TypeError:
            el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))
        for y in el_y:
            for el in range(y[0], y[1] + 1):
                lista_y.append(el)


    if oSheet.Name in ('COMPUTO', 'VARIANTE', 'Analisi di Prezzo'):
        try:
            if lrow is not None:
                SR = oRangeAddress.StartRow
                SR = LeenoComputo.circoscriveVoceComputo(oSheet, SR).RangeAddress.StartRow
            else:
                SR = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
        except AttributeError:
            # DLG.MsgBox('La selezione deve essere contigua.', 'ATTENZIONE!')
            Dialogs.Exclamation(Title = 'ATTENZIONE!',
            Text='''La selezione deve essere contigua.''')
            return 0
        if lrow is not None:
            ER = oRangeAddress.EndRow
            ER = LeenoComputo.circoscriveVoceComputo(oSheet, ER).RangeAddress.EndRow
        else:
            ER = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow
        lista_y = [SR, ER]
    # if oSheet.Name == 'Analisi di Prezzo':
        # try:
            # oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
        # except AttributeError:
            # oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
        # try:
            # if lrow is not None:
                # SR = oRangeAddress.StartRow
                # SR = LeenoComputo.circoscriveVoceComputo(oSheet, SR).RangeAddress.StartRow
            # else:
                # SR = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
        # except AttributeError:
            # DLG.MsgBox('La selezione deve essere contigua.', 'ATTENZIONE!')
            # return 0
        # if lrow is not None:
            # ER = oRangeAddress.EndRow
            # ER = LeenoComputo.circoscriveVoceComputo(oSheet, ER).RangeAddress.EndRow
        # else:
            # ER = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow
    if oSheet.Name == 'CONTABILITA':
        partenza = cerca_partenza()
        if partenza[2] == '#reg':
            sblocca_cont()
            if LeenoUtils.getGlobalVar('sblocca_computo') == 0:
                return
            pass
        else:
            pass
        try:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
        except AttributeError:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
        try:
            if lrow is not None:
                SR = oRangeAddress.StartRow
                SR = LeenoComputo.circoscriveVoceComputo(oSheet, SR).RangeAddress.StartRow
            else:
                SR = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
        except AttributeError:
            # DLG.MsgBox('La selezione deve essere contigua.', 'ATTENZIONE!')
            Dialogs.Exclamation(Title = 'ATTENZIONE!',
            Text='''La selezione deve essere contigua.''')
            return 0
        if lrow is not None:
            ER = oRangeAddress.EndRow
            ER = LeenoComputo.circoscriveVoceComputo(oSheet, ER).RangeAddress.EndRow
        else:
            ER = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow
        lista_y = [SR, ER]
    return lista_y


########################################################################
def seleziona_voce(lrow=None):
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoSheetUtils.selezionaVoce'
    Restituisce inizio e fine riga di una voce in COMPUTO, VARIANTE,
    CONTABILITA o Analisi di Prezzo
    lrow { long }  : numero riga
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if lrow is None or lrow == 0:
        lrow = LeggiPosizioneCorrente()[1]
    if oSheet.Name in ('Elenco Prezzi'):
        return
    try:
        if oSheet.Name in ('COMPUTO', 'VARIANTE'):
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        elif oSheet.Name == 'Analisi di Prezzo':
            sStRange = Circoscrive_Analisi(lrow)
        ###
        if oSheet.Name == 'CONTABILITA':
            partenza = cerca_partenza()
            if partenza[2] == '#reg':
                sblocca_cont()
                if LeenoUtils.getGlobalVar('sblocca_computo') == 0:
                    return
                pass
            else:
                pass
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        ###
    except Exception:
        return
    try:
        sStRange.RangeAddress
        SR = sStRange.RangeAddress.StartRow
        ER = sStRange.RangeAddress.EndRow
    except:
        return
    return (SR, ER)


########################################################################


def MENU_elimina_voce():
    LeenoUtils.DocumentRefresh(False)
    LeenoSheetUtils.elimina_voce()
    LeenoUtils.DocumentRefresh(True)


########################################################################


def MENU_elimina_righe():
    '''
    Elimina le righe selezionate anche se non contigue.
    '''
    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(False)
    oSheet = oDoc.CurrentController.ActiveSheet

    if oSheet.Name == 'Elenco Prezzi':
        Dialogs.Info(Title = 'Info', Text="""Per eliminare una o più voci dall'Elenco Prezzi
devi selezionarle ed utilizzare il comando 'Elimina righe' di Calc.""")
        return

    if oSheet.Name not in ('COMPUTO', 'CONTABILITA', 'VARIANTE', 'Analisi di Prezzo'):
        return

    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
    el_y = []
    lista_y = []
    try:
        len(oRangeAddress)
        for el in oRangeAddress:
            el_y.append((el.StartRow, el.EndRow))
    except TypeError:
        el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))
    for y in el_y:
        for el in range(y[0], y[1] + 1):
            lista_y.append(el)
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    rigen = False
    for y in reversed(lista_y):
        if oSheet.getCellByPosition(2, y).CellStyle not in ('An-lavoraz-generica',
                                                            'An-lavoraz-Cod-sx',
                                                            'comp 1-a',
                                                            'comp 1-a ROSSO',
                                                            'comp sotto centro',
                                                            'EP-mezzo',
                                                            'Livello-0-scritta mini',
                                                            'Livello-1-scritta mini',
                                                            'livello2_') or \
        'Somma positivi e negativi [' in oSheet.getCellByPosition(8, y).String or \
        'SOMMANO' in oSheet.getCellByPosition(8, y).String:
            pass
        else:
            if oSheet.getCellByPosition(2, y).CellStyle in ('comp sotto centro'):
                rigen = True
            if oSheet.getCellByPosition(1, y).CellStyle == 'Data_bianca':
                oCellAddress = oSheet.getCellByPosition(1, y+1).getCellAddress()
                oCellRangeAddr.Sheet = oSheet.RangeAddress.Sheet
                oCellRangeAddr.StartColumn = 1
                oCellRangeAddr.StartRow = y
                oCellRangeAddr.EndColumn = 1
                oCellRangeAddr.EndRow = y
                oSheet.copyRange(oCellAddress, oCellRangeAddr)
            stile = oSheet.getCellByPosition(2, y).CellStyle
            oSheet.getRows().removeByIndex(y, 1)
            if stile in ('Livello-0-scritta mini', 'Livello-1-scritta mini', 'livello2_'):
                Rinumera_TUTTI_Capitoli2(oSheet)
    if rigen:
        rigenera_parziali(False)
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
    #  oDoc.enableAutomaticCalculation(True)
    LeenoUtils.DocumentRefresh(True)

########################################################################
def copia_riga_computo(lrow):
    '''
    Inserisce una nuova riga di misurazione nel computo
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    # lrow = LeggiPosizioneCorrente()[1]
    stile = oSheet.getCellByPosition(1, lrow).CellStyle
    if stile in (
            'comp Art-EP', 'comp Art-EP_R', 'Comp-Bianche in mezzo'
    ):  # Comp-Bianche in mezzo Descr', 'comp 1-a', 'comp sotto centro'):# <stili computo
        lrow = lrow + 1  # PER INSERIMENTO SOTTO RIGA CORRENTE
        oSheet.getRows().insertByIndex(lrow, 1)
        # imposto gli stili
        oSheet.getCellRangeByPosition(
            5,
            lrow,
            7,
            lrow,
        ).CellStyle = 'comp 1-a'
        oSheet.getCellByPosition(0, lrow).CellStyle = 'comp 10 s'
        oSheet.getCellByPosition(1, lrow).CellStyle = 'Comp-Bianche in mezzo'
        oSheet.getCellByPosition(2, lrow).CellStyle = 'comp 1-a'
        oSheet.getCellRangeByPosition(
            3, lrow, 4, lrow).CellStyle = 'Comp-Bianche in mezzo bordate_R'
        oSheet.getCellByPosition(5, lrow).CellStyle = 'comp 1-a PU'
        oSheet.getCellByPosition(6, lrow).CellStyle = 'comp 1-a LUNG'
        oSheet.getCellByPosition(7, lrow).CellStyle = 'comp 1-a LARG'
        oSheet.getCellByPosition(8, lrow).CellStyle = 'comp 1-a peso'
        oSheet.getCellByPosition(9, lrow).CellStyle = 'Blu'
        # ci metto le formule
        oSheet.getCellByPosition(
            9, lrow).Formula = '=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(
                lrow + 1) + ')=0;"";PRODUCT(E' + str(lrow +
                                                     1) + ':I' + str(lrow +
                                                                     1) + '))'
        # Ottieni l'intervallo delle righe e collassa il gruppo
        #  r_addr = oSheet.getCellRangeByPosition(0, 0, 0, lrow).RangeAddress
        #  oSheet.group(r_addr, 1)  # Raggruppa le righe
        #  oSheet.hideDetail(r_addr)  # Collassa il gruppo
        _gotoCella(2, lrow)

    # LeenoUtils.DocumentRefresh(True)
        # oDoc.CurrentController.select(oSheet.getCellByPosition(2, lrow))
        # oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))


def copia_riga_contab(lrow):
    '''
    Inserisce una nuova riga di misurazione in contabilità
    '''
    oDoc = LeenoUtils.getDocument()
    # vado alla vecchia maniera ## copio il range di righe computo da S5 ##
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheetto = oDoc.getSheets().getByName('S5')
    oRangeAddress = oSheetto.getCellRangeByPosition(0, 24, 42, 24).getRangeAddress()

    stile = oSheet.getCellByPosition(1, lrow).CellStyle

    if oSheet.getCellByPosition(1, lrow + 1).CellStyle == 'comp sotto Bianche_R':
        return

    if stile in ('comp Art-EP_R', 'Data_bianca', 'Comp-Bianche in mezzo_R'):
        lrow = lrow + 1  # Inserisci sotto la riga corrente

        oCellAddress = oSheet.getCellByPosition(0, lrow).getCellAddress()


        if oSheet.isProtected():
            # oDoc.unprotect("password")  # Sostituisci con la password corretta
            # oppure
            oSheet.unprotect("password")

        oSheet.getRows().insertByIndex(lrow, 1)

        # DLG.chi(lrow)
        # try:
        #     oSheet.getRows().insertByIndex(lrow, 1)
        # except Exception as e:
        #     # Mostra un messaggio più informativo
        #     DLG.errore(f"Impossibile inserire riga {lrow}: {str(e)}")
        #     return


        oSheet.copyRange(oCellAddress, oRangeAddress)

        if stile == 'comp Art-EP_R':
            oRangeAddress = oSheet.getCellByPosition(1, lrow + 1).getRangeAddress()
            oCellAddress = oSheet.getCellByPosition(1, lrow).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
            oSheet.getCellByPosition(1, lrow + 1).String = ""
            oSheet.getCellByPosition(1, lrow + 1).CellStyle = 'Comp-Bianche in mezzo_R'
        else:
            oSheet.getCellByPosition(1, lrow).CellStyle = 'Comp-Bianche in mezzo_R'
    # Esempio di utilizzo di hideDetail()
    r_addr = oSheet.getCellRangeByPosition(0, 0, 0, lrow).RangeAddress

    _gotoCella(2, lrow)

    LeenoUtils.DocumentRefresh(True)



def copia_riga_analisi(lrow):
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoAnalysis.copiaRigaAnalisi'
    Inserisce una nuova riga di misurazione in analisi di prezzo
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    stile = oSheet.getCellByPosition(0, lrow).CellStyle
    if stile in ('An-lavoraz-desc', 'An-lavoraz-Cod-sx'):
        lrow = lrow + 1
        oSheet.getRows().insertByIndex(lrow, 1)
        # imposto gli stili
        oSheet.getCellByPosition(0, lrow).CellStyle = 'An-lavoraz-Cod-sx'
        oSheet.getCellRangeByPosition(1, lrow, 5,
                                      lrow).CellStyle = 'An-lavoraz-generica'
        oSheet.getCellByPosition(3, lrow).CellStyle = 'An-lavoraz-input'
        oSheet.getCellByPosition(6, lrow).CellStyle = 'An-senza'
        oSheet.getCellByPosition(7, lrow).CellStyle = 'An-senza-DX'
        # ci metto le formule
        #  oDoc.enableAutomaticCalculation(False)
        oSheet.getCellByPosition(1, lrow).Formula = '=IF(A' + str(
            lrow + 1) + '="";"";CONCATENATE("  ";VLOOKUP(A' + str(
                lrow + 1) + ';elenco_prezzi;2;FALSE());' '))'
        oSheet.getCellByPosition(
            2,
            lrow).Formula = '=IF(A' + str(lrow + 1) + '="";"";VLOOKUP(A' + str(
                lrow + 1) + ';elenco_prezzi;3;FALSE()))'
        oSheet.getCellByPosition(3, lrow).Value = 0
        oSheet.getCellByPosition(
            4,
            lrow).Formula = '=IF(A' + str(lrow + 1) + '="";0;VLOOKUP(A' + str(
                lrow + 1) + ';elenco_prezzi;5;FALSE()))'
        oSheet.getCellByPosition(
            5, lrow).Formula = '=D' + str(lrow + 1) + '*E' + str(lrow + 1)
        oSheet.getCellByPosition(
            8, lrow
        ).Formula = '=IF(A' + str(lrow + 1) + '="";"";IF(VLOOKUP(A' + str(
            lrow + 1) + ';elenco_prezzi;6;FALSE())="";"";(VLOOKUP(A' + str(
                lrow + 1) + ';elenco_prezzi;6;FALSE()))))'
        oSheet.getCellByPosition(9, lrow).Formula = '=IF(I' + str(
            lrow + 1) + '="";"";I' + str(lrow + 1) + '*F' + str(lrow + 1) + ')'
        #  oDoc.enableAutomaticCalculation(True)
        # preserva il Pesca
        if oSheet.getCellByPosition(
                1, lrow - 1).CellStyle == 'An-lavoraz-dx-senza-bordi':
            oRangeAddress = oSheet.getCellByPosition(0, lrow +
                                                     1).getRangeAddress()
            oCellAddress = oSheet.getCellByPosition(0, lrow).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
        # oSheet.getCellByPosition(0, lrow).String = 'Cod. Art.?'
    _gotoCella(0, lrow)
    if LeenoConfig.Config().read('Generale', 'pesca_auto') == '1':
        pesca_cod()

########################################################################


def MENU_Copia_riga_Ent():
    '''
    @@ DA DOCUMENTARE
    '''
    # with LeenoUtils.DocumentRefreshContext(False):
    Copia_riga_Ent()
        # LeenoSheetUtils.adattaAltezzaRiga()

def Copia_riga_Ent(num_righe=1):
    """
    Aggiunge una o tante righe di misurazione.
    num_righe { int }: Numero di righe da aggiungere (default: 1).
    """
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nome_sheet = oSheet.Name

    # Se le colonne di misura sono nascoste, vengono visualizzate
    col_misura = oSheet.getColumns()
    if not col_misura.getByIndex(5).IsVisible:
        n = SheetUtils.getLastUsedRow(oSheet)
        for el in range(4, n):
            cell = oSheet.getCellByPosition(2, el)
            if cell.CellStyle == "comp sotto centro":
                cell.Formula = ''
        for el in range(5, 8):
            col_misura.getByIndex(el).IsVisible = True

    lrow = LeggiPosizioneCorrente()[1]
    dettaglio_attivo = cfg.read('Generale', 'dettaglio') == '1'

    azioni = {
        'COMPUTO': copia_riga_computo,
        'VARIANTE': copia_riga_computo,
        'CONTABILITA': copia_riga_contab,
        'Analisi di Prezzo': copia_riga_analisi,
        'Elenco Prezzi': MENU_nuova_voce_scelta,
    }

    if nome_sheet in azioni:
        for _ in range(num_righe):
            if dettaglio_attivo and nome_sheet in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
                dettaglio_misura_rigo()
            azioni[nome_sheet](lrow)
            lrow += 1
    oSheet.getCellRangeByPosition(0, lrow, 0, lrow).Rows.OptimalHeight = True

########################################################################
def cerca_partenza():
    '''
    Conserva, nella variabile globale 'partenza', il nome del foglio [0] e l'id
    della riga di codice prezzo componente [1], il flag '#reg' solo per la contabilità.
    partenza = (nome_foglio, id_rcodice, flag_contabilità)
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]

    partenza = LeenoSheetUtils.cercaPartenza(oSheet, lrow)
    LeenoUtils.setGlobalVar('partenza', partenza)
    return partenza


def sblocca_cont():
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoContab.sbloccaContabilita'
    Controlla che non ci siano atti contabili registrati e dà il consenso a procedere.
    '''
    partenza = LeenoUtils.getGlobalVar('partenza')
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in ('CONTABILITA'):
        partenza = cerca_partenza()
        # DLG.chi(partenza[2])
        # DLG.chi(LeenoUtils.getGlobalVar('sblocca_computo'))
        if LeenoUtils.getGlobalVar('sblocca_computo') == 1:
            pass
        else:
            if partenza[2] == '':
                pass
            if partenza[2] == '#reg':
                if Dialogs.YesNoDialog(Title='Avviso: Voce già registrata!',

                Text= """Lavorando in questo punto del foglio,
comprometterai la validità degli atti contabili già emessi.

Vuoi procedere?

SCEGLIENDO SÌ DOVRAI NECESSARIAMENTE RIGENERARLI!""") == 0:
                    pass
                else:
                    LeenoUtils.setGlobalVar('sblocca_computo', 1)
        # DLG.chi(LeenoUtils.getGlobalVar('sblocca_computo'))


########################################################################


def MENU_cerca_in_elenco():
    '''
    Evidenzia il codice di elenco prezzi della voce corrente.
    '''
    cerca_in_elenco()


def cerca_in_elenco():
    '''
    Evidenzia il codice di elenco prezzi della voce corrente.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]
    if oSheet.Name in ('COMPUTO', 'CONTABILITA', 'VARIANTE', 'Registro',
                       'Analisi di Prezzo', 'SAL'):
        if oSheet.Name == 'Analisi di Prezzo':
            if oSheet.getCellByPosition(
                    0, lrow).CellStyle in ('An-lavoraz-Cod-sx', 'An-1_sigla'):
                codice_da_cercare = oSheet.getCellByPosition(0, lrow).String
            else:
                return
        elif oSheet.Name in ('Registro','SAL'):
            codice_da_cercare =oSheet.getCellByPosition(0, lrow).String.split('\n')[1]
        else:
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
            sopra = sStRange.RangeAddress.StartRow
            codice_da_cercare = oSheet.getCellByPosition(1, sopra + 1).String
        oSheet = oDoc.getSheets().getByName("Elenco Prezzi")
        oSheet.IsVisible = True
        GotoSheet('Elenco Prezzi')
    elif oSheet.Name in ('Elenco Prezzi'):
        if oSheet.getCellByPosition(1, lrow).Type.value == 'FORMULA':
            codice_da_cercare = oSheet.getCellByPosition(0, lrow).String
        else:
            return
        oSheet = oDoc.getSheets().getByName("Analisi di Prezzo")
        oSheet.IsVisible = True
        GotoSheet('Analisi di Prezzo')

    if codice_da_cercare == "Cod. Art.?":
        return
    if codice_da_cercare != '':
        oCell = SheetUtils.uFindString(codice_da_cercare, oSheet)
        try:
            oDoc.CurrentController.select(
                oSheet.getCellRangeByPosition(oCell[0], oCell[1], 30, oCell[1]))
        except:
            _gotoCella(1,  2)
    return

########################################################################


def MENU_pesca_cod():
    '''
    @@ DA DOCUMENTARE
    '''
    pesca_cod()


def pesca_cod():
    '''
    Permette di scegliere il codice per la voce di COMPUTO o VARIANTE o CONTABILITA dall'Elenco Prezzi.
    Capisce quando la voce nel libretto delle misure è già registrata o nel documento ci sono già atti contabili emessi.
    '''
    partenza = LeenoUtils.getGlobalVar('partenza')
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]

    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_contab = LeenoUtils.getGlobalVar('stili_contab')
    stili_analisi = LeenoUtils.getGlobalVar('stili_analisi')
    stili_elenco = LeenoUtils.getGlobalVar('stili_elenco')

    if oSheet.getCellByPosition(0, lrow).CellStyle not in stili_computo + stili_contab + stili_analisi + stili_elenco:
        return
    if oSheet.Name in ('Analisi di Prezzo'):
        test = oSheet.getCellByPosition(0, lrow).String
        partenza = cerca_partenza()
        cerca_in_elenco()
        GotoSheet('Elenco Prezzi')
        try:
            if test == '':
                oSheet = oDoc.CurrentController.ActiveSheet
                y = SheetUtils.uFindStringCol(
                    'ELENCO DEI COSTI ELEMENTARI', 1, oSheet) + 1
                _gotoCella(0, y)
            return
        except:
            pass

###

    if oSheet.Name in ('CONTABILITA'):
        # controllo che non ci siano atti registrati
        partenza = cerca_partenza()
        if partenza[2] == '#reg':
            sblocca_cont()
            if LeenoUtils.getGlobalVar('sblocca_computo') == 0:
                return
            pass
        else:
            pass
        ###
###
    if oSheet.Name in ('COMPUTO', 'VARIANTE') or 'LISTA' in oSheet.Name.upper():
        if oDoc.NamedRanges.hasByName("_Lib_1"):
            if LeenoUtils.getGlobalVar('sblocca_computo') == 0:
                if DLG.DlgSiNo(
                        "Risulta già registrato un SAL. VUOI PROCEDERE COMUQUE?",
                        'ATTENZIONE!') == 3:
                    return
                if Dialogs.YesNoDialog(IconType="question",Title='ATTENZIONE!',
                Text="Risulta già registrato un SAL."
                    "Vuoi procedere comunque?") == 0:
                    return
                else:
                    LeenoUtils.setGlobalVar('sblocca_computo', 1)
        partenza = cerca_partenza()
    if oSheet.getCellByPosition(1, partenza[1]).String != 'Cod. Art.?':
        cerca_in_elenco()
    GotoSheet('Elenco Prezzi')
    ###
    if oSheet.Name in ('Elenco Prezzi'):
        try:
            lrow = LeggiPosizioneCorrente()[1]
            codice = oSheet.getCellByPosition(0, lrow).String
            GotoSheet(partenza[0])
            oSheet = oDoc.CurrentController.ActiveSheet
            if partenza[0] == 'Analisi di Prezzo':
                oSheet.getCellByPosition(0, partenza[1]).String = codice
                _gotoCella(3, partenza[1])
            else:
                oSheet.getCellByPosition(1, partenza[1]).String = codice
                _gotoCella(2, partenza[1] + 1)
        except NameError:
            return


########################################################################
def MENU_ricicla_misure():
    '''
    In CONTABILITA consente l'inserimento di nuove voci di misurazione
    partendo da voci già inserite in COMPUTO o VARIANTE.
    '''
    LeenoUtils.DocumentRefresh(False)
    oDoc = LeenoUtils.getDocument()

    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name == 'CONTABILITA':
        try:
            # controllo che non ci siano atti registrati
            partenza = cerca_partenza()
            if partenza[2] == '#reg':
                sblocca_cont()
                if LeenoUtils.getGlobalVar('sblocca_computo') == 0:
                    return
                pass
            else:
                pass
            ###
        except Exception:
            pass
        lrow = LeggiPosizioneCorrente()[1]
        lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, saltaCat=True)

        stili_contab = LeenoUtils.getGlobalVar('stili_contab')

        if oSheet.getCellByPosition(0, lrow).CellStyle not in stili_contab + (
                'comp Int_colonna_R_prima', ):
            return
        ins_voce_contab(arg=0)
        partenza = cerca_partenza()
        try:
            GotoSheet(cfg.read('Contabilita', 'ricicla_da'))
        except:
            cfg.write('Contabilita', 'ricicla_da', 'COMPUTO')
            GotoSheet(cfg.read('Contabilita', 'ricicla_da'))
        LeenoUtils.DocumentRefresh(True)
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        lrow = LeggiPosizioneCorrente()[1]
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        sopra = sStRange.RangeAddress.StartRow + 2
        sotto = sStRange.RangeAddress.EndRow - 1

        lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 1, True)
        _gotoCella(2, lrow + 1)

        oSrc = oSheet.getCellRangeByPosition(2, sopra, 8,
                                             sotto).getRangeAddress()
        oSheet.getCellByPosition(2, sopra - 1).CellBackColor = 14942166
        partenza = LeenoUtils.getGlobalVar('partenza')
        if partenza is None:
            return
        oDest = oDoc.getSheets().getByName('CONTABILITA')
        oCellAddress = oDest.getCellByPosition(2, partenza[1] + 1).getCellAddress()
        GotoSheet('CONTABILITA')

        if sotto != sopra:
            oDest.getRows().insertByIndex(partenza[1] + 2, sotto - sopra)
            oDest.getCellRangeByPosition(1, partenza[1] + 2, 1, partenza[1] +
                sotto - sopra +1).CellStyle = 'Comp-Bianche in mezzo_R'

        oDest.copyRange(oCellAddress, oSrc)
        oDest.getCellByPosition(
            1, partenza[1]).String = oSheet.getCellByPosition(1, sopra - 1).String
        oDest.getCellByPosition(2, partenza[1]).CellBackColor = 14942166
        rigenera_voce(partenza[1])

        _gotoCella(2, partenza[1] + 1)

    LeenoUtils.DocumentRefresh(True)

    LeenoSheetUtils.adattaAltezzaRiga(oDoc.CurrentController.ActiveSheet)


def MENU_inverti_segno():
    inverti_segno()

def inverti_segno():
    '''
    Inverte il segno delle formule di quantità nei righi di misurazione selezionati.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # estrae il range o i ranges selezionati
    # (possono essere più di uno)
    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()

    # inserisce in una lista le righe di inizio e fine
    # di ogni range come touples ((inizio, fine), (inizio, fine)...)
    el_y = []
    try:
        len(oRangeAddress)
        for el in oRangeAddress:
            el_y.append((el.StartRow, el.EndRow))
    except TypeError:
        el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))

    # estrate tutte le righe incluse nel o nei range(s)
    # e le inserisce in una lista di righe
    lista = []
    for y in el_y:
        for el in range(y[0], y[1] + 1):
            lista.append(el)

    # va ad eseguire il lavoro su ogni riga della lista
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        for lrow in lista:
            if 'comp 1-a' in oSheet.getCellByPosition(2, lrow).CellStyle:
                if 'ROSSO' in oSheet.getCellByPosition(2, lrow).CellStyle:
                    # se VediVoce
                    oSheet.getCellByPosition(9, lrow).Formula = (
                       '=IF(PRODUCT(E' + str(lrow + 1) + ':I' +
                       str(lrow + 1) + ')=0;"";PRODUCT(E' +
                       str(lrow + 1) + ':I' +
                       str(lrow + 1) + '))')

                    for x in range(2, 10):
                        oSheet.getCellByPosition(x, lrow).CellStyle = (
                        oSheet.getCellByPosition(x, lrow).CellStyle.split(' ROSSO')[0])
                else:
                    # se VediVoce
                    oSheet.getCellByPosition(9, lrow).Formula = (
                       '=IF(PRODUCT(E' + str(lrow + 1) + ':I' +
                       str(lrow + 1) + ')=0;"";-PRODUCT(E' +
                       str(lrow + 1) + ':I' + str(lrow + 1) + '))')

                    for x in range(2, 10):
                        oSheet.getCellByPosition(x, lrow).CellStyle = (
                        oSheet.getCellByPosition(x, lrow).CellStyle + ' ROSSO')

    elif oSheet.Name in ('CONTABILITA'):
        for lrow in lista:
            if 'comp 1-a' in oSheet.getCellByPosition(2, lrow).CellStyle:
                formula1 = oSheet.getCellByPosition(9, lrow).Formula
                formula2 = oSheet.getCellByPosition(11, lrow).Formula
                oSheet.getCellByPosition(11, lrow).Formula = formula1
                oSheet.getCellByPosition(9, lrow).Formula = formula2
                if oSheet.getCellByPosition(11, lrow).Value > 0:
                    for x in range(2, 12):
                        oSheet.getCellByPosition(x, lrow).CellStyle = (
                        oSheet.getCellByPosition(x, lrow).CellStyle + ' ROSSO')
                else:
                    for x in range(2, 12):
                        oSheet.getCellByPosition(
                            x, lrow).CellStyle = (
                            oSheet.getCellByPosition(x, lrow).CellStyle.split(' ROSSO')[0])


########################################################################
def valuta_cella(oCell):
    '''
    Estrae qualsiasi valore da una cella, restituendo una stringa, indipendentemente dal tipo originario.
    oCell       { object }  : cella da validare
    '''
    if oCell.Type.value == 'FORMULA':
        if re.search('[a-zA-Z]', oCell.Formula):
            valore = str(oCell.Value)
        else:
            valore = oCell.Formula.split('=')[-1]
    elif oCell.Type.value == 'VALUE':
        valore = str(oCell.Value)
    elif oCell.Type.value == 'TEXT':
        valore = str(oCell.String)
    elif oCell.Type.value == 'EMPTY':
        valore = ''
    if valore == ' ':
        valore = ''
    return valore


########################################################################
def dettaglio_misura_rigo():
    '''
    Indica il dettaglio delle misure nel rigo di descrizione quando
    incontra delle formule nei valori immessi.
    bit { integer }  : 1 inserisce i dettagli
                       0 cancella i dettagli
    '''
    oDoc = LeenoUtils.getDocument()
    LeenoUtils.DocumentRefresh(False)
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]
    if ' ►' in oSheet.getCellByPosition(2, lrow).String:
        oSheet.getCellByPosition(2, lrow).String = oSheet.getCellByPosition(
            2, lrow).String.split(' ►')[0]
    # if oSheet.getCellByPosition(2, lrow).CellStyle in (
            # 'comp 1-a'
    if 'comp 1-a' in oSheet.getCellByPosition(2, lrow).CellStyle and \
    "*** VOCE AZZERATA ***" not in oSheet.getCellByPosition(2, lrow).String:
        for el in range(5, 9):
            if oSheet.getCellByPosition(el, lrow).Type.value == 'FORMULA':
                stringa = ''
                break
            else:
                stringa = None

        if stringa == '':
            for el in range(5, 9):
                if oSheet.getCellByPosition(el, lrow).Type.value == 'FORMULA':
                    if '$' not in oSheet.getCellByPosition(el, lrow).Formula:
                        try:
                            eval(
                                oSheet.getCellByPosition(
                                    el, lrow).Formula.split('=')[1].replace(
                                        '^', '**'))
                            # eval(oSheet.getCellByPosition(el, lrow).Formula.split('=')[1])
                            stringa = stringa + '(' + oSheet.getCellByPosition(
                                el, lrow).Formula.split('=')[-1] + ')*'
                        except Exception:
                            stringa = stringa + '(' + oSheet.getCellByPosition(
                                el, lrow).String.split('=')[-1] + ')*'
                            pass
                else:
                    stringa = stringa + '*' + str(
                        oSheet.getCellByPosition(el, lrow).String) + '*'
            while '**' in stringa:
                stringa = stringa.replace('**', '*')
            if stringa[0] == '*':
                stringa = stringa[1:-1]
            else:
                stringa = stringa[0:-1]
            stringa = ' ►' + stringa # + ')'
            if oSheet.getCellByPosition(2, lrow).Type.value != 'FORMULA':
                oSheet.getCellByPosition(
                    2, lrow).String = oSheet.getCellByPosition(
                        2, lrow).String + stringa.replace('.', ',')
    LeenoUtils.DocumentRefresh(True)


########################################################################
def dettaglio_misure(bit):
    '''
    Indica il dettaglio delle misure nel rigo di descrizione quando
    incontra delle formule nei valori immessi.
    bit { integer }  : 1 inserisce i dettagli
                       0 cancella i dettagli
    '''
    # qui il refresh lascia il foglio in freeze
    LeenoUtils.DocumentRefresh(False)
    oDoc = LeenoUtils.getDocument()
    try:
        oSheet = oDoc.CurrentController.ActiveSheet
    except Exception:
        return
    ER = SheetUtils.getUsedArea(oSheet).EndRow

    if bit == 1:
        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start('Rigenerazione in corso...', LeenoSheetUtils.cercaUltimaVoce(oSheet))
        indicator.setValue(0)
        

        for lrow in range(0, ER):
            indicator.setValue(lrow)
            if 'comp 1-a' in oSheet.getCellByPosition(2, lrow).CellStyle and \
            "*** VOCE AZZERATA ***" not in oSheet.getCellByPosition(
                    2, lrow).String:
                for el in range(5, 9):
                    if oSheet.getCellByPosition(el, lrow).Type.value == 'FORMULA':
                        stringa = ''
                        break
                    else:
                        stringa = None
                if stringa == '':
                    for el in range(5, 9):
                        if oSheet.getCellByPosition(
                                el, lrow).Type.value == 'FORMULA':
                            if '$' not in oSheet.getCellByPosition(
                                    el, lrow).Formula:
                                try:
                                    eval(
                                        oSheet.getCellByPosition(
                                            el, lrow).Formula.split('=')
                                        [1].replace('^', '**'))
                                    stringa = stringa + '(' + oSheet.getCellByPosition(
                                        el, lrow).Formula.split('=')[-1] + ')*'
                                except Exception:
                                    stringa = stringa + '(' + oSheet.getCellByPosition(
                                        el, lrow).String.split('=')[-1] + ')*'
                                    pass
                        else:
                            stringa = stringa + '*' + str(
                                oSheet.getCellByPosition(el,
                                                         lrow).String) + '*'
                    while '**' in stringa:
                        stringa = stringa.replace('**', '*')
                    if stringa[0] == '*':
                        stringa = stringa[1:-1]
                    else:
                        stringa = stringa[0:-1]
                    stringa = ' ►' + stringa #+ ')'
                    if oSheet.getCellByPosition(2,
                                                lrow).Type.value != 'FORMULA':
                        oSheet.getCellByPosition(
                            2, lrow).String = oSheet.getCellByPosition(
                                2, lrow).String + stringa.replace('.', ',')
        indicator.end()
    else:
        for lrow in range(0, ER):
            if ' ►' in oSheet.getCellByPosition(2, lrow).String:
                oSheet.getCellByPosition(
                    2, lrow).String = oSheet.getCellByPosition(
                        2, lrow).String.split(' ►')[0]

    LeenoUtils.DocumentRefresh(True)
    # LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    return


########################################################################
def debug_validation():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #  DLG.mri(oDoc.CurrentSelection.Validation)

    oSheet.getCellRangeByName('L1').String = 'Ricicla da:'
    oSheet.getCellRangeByName('L1').CellStyle = 'Reg_prog'
    oCell = oSheet.getCellRangeByName('N1')
    if oCell.String not in ("COMPUTO", "VARIANTE", 'Scegli origine'):
        oCell.CellStyle = 'Menu_sfondo _input_grasBig'
        valida_cella(oCell,
                     '"COMPUTO";"VARIANTE"',
                     titoloInput='Scegli...',
                     msgInput='COMPUTO o VARIANTE',
                     err=True)
        oCell.String = 'Scegli...'


def valida_cella(oCell, lista_val, titoloInput='', msgInput='', err=False):
    '''
    Validità lista valori
    Imposta un elenco di valori a cascata, da cui scegliere.
    oCell       { object }  : cella da validare
    lista_val   { string }  : lista dei valori in questa forma: '"UNO";"DUE";"TRE"'
    titoloInput { string }  : titolo del suggerimento che compare passando il cursore sulla cella
    msgInput    { string }  : suggerimento che compare passando il cursore sulla cella
    err         { boolean } : permette di abilitare il messaggio di errore per input non validi
    '''
    # oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet

    oTabVal = oCell.getPropertyValue("Validation")
    oTabVal.setPropertyValue('ConditionOperator', 1)

    oTabVal.setPropertyValue("ShowInputMessage", True)
    oTabVal.setPropertyValue("InputTitle", titoloInput)
    oTabVal.setPropertyValue("InputMessage", msgInput)
    oTabVal.setPropertyValue("ErrorMessage",
                             "ERRORE: Questo valore non è consentito.")
    oTabVal.setPropertyValue("ShowErrorMessage", err)
    oTabVal.ErrorAlertStyle = uno.Enum(
        "com.sun.star.sheet.ValidationAlertStyle", "STOP")
    oTabVal.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")
    oTabVal.Operator = uno.Enum("com.sun.star.sheet.ConditionOperator",
                                "EQUAL")
    oTabVal.setFormula1(lista_val)
    oCell.setPropertyValue("Validation", oTabVal)


def debug_ConditionalFormat():
    oDoc = LeenoUtils.getDocument()
    oCell = oDoc.CurrentSelection
    oSheet = oDoc.CurrentController.ActiveSheet

    i = oCell.RangeAddress.StartRow
    n = oCell.Rows.Count
    oSheet.getRows().removeByIndex(i, n)


########################################################################
def comando(cmd):
    '''
    Esegue un comando di menù.
    cmd       { string }  : nome del comando di menù

    Elenco comandi:
    'DeletePrintArea'       = Cancella l'area di stampa
    'ShowDependents'        = Mostra le celle dipendenti
    'ClearArrowDependents'  = elimina frecce celle dipendenti
    'Undo'                  = Annulla ultimo comando
    'CalculateHard'         = Ricalcolo incondizionato
    'Save'                  = Salva il file
    'NumberFormatDecimal'   =
    'ConvertFormulaToValue' = Converti formula in valore
    'DataBarFormatDialog'   = Formato Barra dati
    '''
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, ".uno:" + cmd, "", 0,
                                   [])


########################################################################
# SheetUtils.visualizza_PageBreak moved to SheetUtils.py
########################################################################


def delete(arg):
    '''
    Elimina righe o colonne.
    arg       { string }  : 'R' per righe
                            'C' per colonne
    '''
    # oDoc = LeenoUtils.getDocument()
    #  oSheet = oDoc.CurrentController.ActiveSheet
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    oProp = PropertyValue()
    oProp.Name = 'Flags'
    oProp.Value = arg
    properties = (oProp, )

    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, ".uno:DeleteCell", "", 0, properties)

def delete_cells(direction='up'):
    '''
    Elimina le celle selezionate e sposta le celle adiacenti nella direzione specificata.

    Args:
        direction (str): Direzione di spostamento delle celle:
                         - 'up' o 'u' (default): sposta le celle in alto
                         - 'left' o 'l': sposta le celle a sinistra

    Raises:
        ValueError: Se la direzione specificata non è valida
    '''
    # Validazione dell'input
    direction = direction.lower()
    if direction not in ('up', 'u', 'left', 'l'):
        raise ValueError("Direzione non valida. Usare 'up'/'u' o 'left'/'l'")

    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()

    # Prepara le proprietà per il comando
    # 1. Imposta l'azione di eliminazione celle con spostamento
    prop1 = PropertyValue()
    prop1.Name = 'Flags'
    prop1.Value = 'S'  # 'S' = Modalità spostamento celle

    # 2. Specifica la direzione di spostamento
    prop2 = PropertyValue()
    prop2.Name = 'ToRight'
    prop2.Value = (direction in ('left', 'l'))  # True per sinistra, False per alto

    properties = (prop1, prop2)

    # Esegui il comando di eliminazione
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(
        oFrame, 
        ".uno:DeleteCell",
        "", 
        0, 
        properties
    )
########################################################################
def paste_clip(insCells=0, pastevalue=False, noformat=False):
    '''
    Incolla il contenuto della clipboard.
    insCells       { bit }  : con 1 inserisce nuove righe
    pastevalue     { boolean }  : con True non incolla le formule
    noformat       { boolean }  : con True non incolla i formati
    '''
    oDoc = LeenoUtils.getDocument()
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    oProp = []

    if pastevalue:
        if noformat:
            oProp.append(crea_property_value('Flags', 'SV'))  # Solo Numeri e Testo (senza formato)
        else:
            oProp.append(crea_property_value('Flags', 'SVTD'))  # Numeri, Testo, Data e ora, Formati
    else:
        if noformat:
            oProp.append(crea_property_value('Flags', 'SV'))  # Solo Numeri e Testo (senza formato)
        else:
            oProp.append(crea_property_value('Flags', 'A'))  # Tutto

    oProp.append(crea_property_value('FormulaCommand', 0))
    oProp.append(crea_property_value('SkipEmptyCells', False))
    oProp.append(crea_property_value('Transpose', False))
    oProp.append(crea_property_value('AsLink', False))

    # insert mode ON
    if insCells == 1:
        oProp.append(crea_property_value('MoveMode', 0))
        #  oProp.append(crea_property_value('MoveMode', 4))  # per inserire intere righe

    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(
        oFrame, '.uno:InsertContents', '', 0, tuple(oProp))
    oDoc.CurrentController.select(oDoc.createInstance(
        "com.sun.star.sheet.SheetCellRanges"))  # unselect



########################################################################
def paste_format():
    '''
    Incolla solo il formato cella
    '''
    oDoc = LeenoUtils.getDocument()
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    oProp = []
    oProp0 = PropertyValue()
    oProp0.Name = 'Flags'
    oProp0.Value = 'T'
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
    oProp.append(oProp0)
    oProp.append(oProp1)
    oProp.append(oProp2)
    oProp.append(oProp3)
    oProp.append(oProp4)
    properties = tuple(oProp)
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, '.uno:InsertContents', '', 0, properties)
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect


########################################################################
def MENU_copia_celle_visibili():
    with LeenoUtils.DocumentRefreshContext(False):
        copia_celle_visibili()


def copia_celle_visibili():
    '''
    A partire dalla selezione di un range di celle in cui alcune righe e/o
    colonne sono nascoste, mette in clipboard solo il contenuto delle celle
    visibili.
    Liberamente ispirato a "Copy only visible cells" http://bit.ly/2j3bfq2
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
    # IS = oRangeAddress.Sheet
    SC = oRangeAddress.StartColumn
    EC = oRangeAddress.EndColumn
    SR = oRangeAddress.StartRow
    ER = oRangeAddress.EndRow
    if EC == 1023:
        EC = SheetUtils.getUsedArea(oSheet).EndColumn
    if ER == 1048575:
        ER = SheetUtils.getUsedArea(oSheet).EndRow
    righe = []
    colonne = []
    i = 0
    for nRow in range(SR, ER + 1):
        if not oSheet.getCellByPosition(SR, nRow).Rows.IsVisible:
            righe.append(i)
        i += 1
    i = 0
    for nCol in range(SC, EC + 1):
        if not oSheet.getCellByPosition(nCol, nRow).Columns.IsVisible:
            colonne.append(i)
        i += 1

    if not oDoc.getSheets().hasByName('tmp_clip'):
        sheet = oDoc.createInstance("com.sun.star.sheet.Spreadsheet")
        tmp = oDoc.Sheets.insertByName('tmp_clip', sheet)
    tmp = oDoc.getSheets().getByName('tmp_clip')

    oCellAddress = tmp.getCellByPosition(0, 0).getCellAddress()
    tmp.copyRange(oCellAddress, oRangeAddress)

    for i in reversed(righe):
        tmp.getRows().removeByIndex(i, 1)
    for i in reversed(colonne):
        tmp.getColumns().removeByIndex(i, 1)

    oRange = tmp.getCellRangeByPosition(0, 0, EC - SC - len(colonne),
                                        ER - SR - len(righe))
    oDoc.CurrentController.select(oRange)

    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, ".uno:Copy", "", 0, [])
    oDoc.Sheets.removeByName('tmp_clip')
    oDoc.CurrentController.setActiveSheet(oSheet)
    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(SC, SR, EC, ER))


# LeggiPosizioneCorrente ###########################################################
def LeggiPosizioneCorrente():
    '''
    Restituisce la tupla (IDcolonna, IDriga, NameSheet) della posizione corrente
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        if oDoc.getCurrentSelection().getRangeAddresses()[0]:
            nRow = oDoc.getCurrentSelection().getRangeAddresses()[0].StartRow
            nCol = oDoc.getCurrentSelection().getRangeAddresses(
            )[0].StartColumn
    except AttributeError:
        nRow = oDoc.getCurrentSelection().getRangeAddress().StartRow
        nCol = oDoc.getCurrentSelection().getRangeAddress().StartColumn
    return (nCol, nRow, oSheet.Name)

########################################################################
# numera le voci di computo o contabilità


def MENU_numera_voci():
    '''
    Comando di menu per numera_voci()
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    LeenoSheetUtils.numeraVoci(oSheet, 4, True)
    Rinumera_TUTTI_Capitoli2(oSheet)


def numera_voci(bit=1):  #
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoSheetUtils.numeraVoci'
    bit { integer }  : 1 rinumera tutto
                       0 rinumera dalla voce corrente in giù
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lastRow = SheetUtils.getUsedArea(oSheet).EndRow + 1
    lrow = LeggiPosizioneCorrente()[1]
    n = 1

    if bit == 0:
        for x in reversed(range(0, lrow)):
            if(
               oSheet.getCellByPosition(1, x).CellStyle in ('comp Art-EP', 'comp Art-EP_R') and
               oSheet.getCellByPosition(1, x).CellBackColor != 15066597):
                n = oSheet.getCellByPosition(0, x).Value + 1
                break
        for row in range(lrow, lastRow):
            if oSheet.getCellByPosition(1, row).CellBackColor == 15066597:
                oSheet.getCellByPosition(0, row).String = ''
            elif oSheet.getCellByPosition(1,row).CellStyle in ('comp Art-EP', 'comp Art-EP_R'):
                oSheet.getCellByPosition(0, row).Value = n
                n += 1
    if bit == 1:
        for row in range(0, lastRow):
            # if oSheet.getCellByPosition(1,row).CellBackColor == 15066597:
            # oSheet.getCellByPosition(0,row).String = ''
            # elif oSheet.getCellByPosition(1,row).CellStyle in('comp Art-EP', 'comp Art-EP_R'):
            # oSheet.getCellByPosition(0,row).Value = n
            # n = n+1
            if oSheet.getCellByPosition(1, row).CellStyle in ('comp Art-EP','comp Art-EP_R'):
                oSheet.getCellByPosition(0, row).Value = n
                n = n + 1
            # oSheet.getCellByPosition(0,row).Value = n
            # n = n+1


########################################################################

def richiesta_offerta():
    '''Crea la Lista Lavorazioni e Forniture dall'Elenco Prezzi,
per la formulazione dell'offerta'''
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    GotoSheet('Elenco Prezzi')
    genera_sommario()
    oSheet = oDoc.CurrentController.ActiveSheet
    idSheet = oSheet.RangeAddress.Sheet + 1
    if oDoc.getSheets().hasByName('Richiesta offerta'):
        Dialogs.Exclamation(Title = 'ATTENZIONE!',
        Text=f'La tabella di nome Richiesta offerta è già presente.')
        return
    else:
        oDoc.Sheets.copyByName('Elenco Prezzi', 'Richiesta offerta', idSheet)
    nSheet = 'Richiesta offerta'
    GotoSheet(nSheet)
    setTabColor(10079487)
    oSheet = oDoc.CurrentController.ActiveSheet
    fine = SheetUtils.getUsedArea(oSheet).EndRow + 1
    oRange = oSheet.getCellRangeByPosition(12, 3, 12, fine)
    aSaveData = oRange.getDataArray()
    oRange = oSheet.getCellRangeByPosition(3, 3, 3, fine)

    oRange.CellStyle = 'EP statistiche_q'
    oRange.setDataArray(aSaveData)
    oSheet.getCellByPosition(3, 2).String = 'Quantità\na Computo'
    oSheet.getCellByPosition(5, 2).String = 'Prezzo Unitario\nin lettere'
    oSheet.getCellByPosition(6, 2).String = 'Importo'

    colonne_visibili = ["D", "F", "G"]
    larghezze_colonne = {
        "A": 1600,
        "B": 8000,
        "C": 1200,
        "D": 1600,
        "E": 1500,
        "F": 4000,
        "G": 1800
    }

    for colonna in colonne_visibili:
        oSheet.getColumns().getByName(colonna).IsVisible = True

    for colonna, larghezza in larghezze_colonne.items():
        oSheet.getColumns().getByName(colonna).Columns.Width = larghezza

    oDoc.CurrentController.freezeAtPosition(0, 1)

    formule = []
    for x in range(3, SheetUtils.getUsedArea(oSheet).EndRow - 1):
        formule.append([
            f'=IF(E{x+1}<>""; D{x+1}*E{x+1}; "")'
        ])
    oSheet.getCellRangeByPosition(6, 3, 6,
                                  len(formule) + 2).CellBackColor = 15757935
    oRange = oSheet.getCellRangeByPosition(6, 3, 6, len(formule) + 2)
    formule = tuple(formule)
    oRange.setFormulaArray(formule)

    oSheet.getCellRangeByPosition(
        5, 3, 5,
        fine).clearContents(VALUE + DATETIME + STRING + ANNOTATION + FORMULA +
                            HARDATTR + OBJECTS + EDITATTR + FORMATTED)

    oSheet.getCellRangeByPosition(4, 3, 4, fine + 1).clearContents(
        VALUE + FORMULA + STRING)  # cancella prezzi unitari
    oSheet.getCellRangeByPosition(0, fine - 1, 100, fine +
                                  1).clearContents(VALUE + FORMULA + STRING)
    
    #copio le quantità dalla colonna computo
    oSrc = oSheet.getCellRangeByPosition(11, 3, 11, fine).getDataArray()  # Ottieni i dati come matrice
    oDest = oSheet.getCellRangeByPosition(3, 3, 3, fine)
    oDest.setDataArray(oSrc)

    oSheet.Columns.insertByIndex(0, 1)
    oSrc = oSheet.getCellRangeByPosition(1, 0, 1, fine).RangeAddress
    oDest = oSheet.getCellByPosition(0, 0).CellAddress
    oSheet.copyRange(oDest, oSrc)
    oSheet.getCellByPosition(0, 2).String = "N."
    for x in range(3, fine - 1):
        oSheet.getCellByPosition(0, x).Value = x - 2

    oSheet.getCellRangeByPosition(3, 1, 7, fine).CellStyle = "EP statistiche_q"
    oSheet.getCellRangeByPosition(0, 1, 0, fine).CellStyle = "EP-aS"
    for y in (2, 3):
        for x in range(3, fine - 1):
            if oSheet.getCellByPosition(y, x).Type.value == 'FORMULA':
                oSheet.getCellByPosition(y,
                                         x).String = oSheet.getCellByPosition(
                                             y, x).String
    for x in range(3, fine - 1):
        if oSheet.getCellByPosition(5, x).Type.value == 'FORMULA':
            oSheet.getCellByPosition(5, x).Value = oSheet.getCellByPosition(
                5, x).Value
    oSheet.getColumns().getByName("A").Columns.Width = 650

    oSheet.getCellByPosition(
        7, fine).Formula = "=SUBTOTAL(9;H2:H" + str(fine + 1) + ")"
    oSheet.getCellByPosition(2, fine).String = "TOTALE COMPUTO"
    oSheet.getCellRangeByPosition(0, fine, 7, fine).CellStyle = "Comp TOTALI"

    oSheet.Rows.removeByIndex(fine - 2, 2)
    oSheet.Rows.removeByIndex(0, 2)
    oSheet.getCellByPosition(2,
                             fine + 3).String = "(diconsi euro - in lettere)"
    oSheet.getCellRangeByPosition(2, fine + 3, 6,
                                  fine + 3).CellStyle = "List-intest_med_c"
    oSheet.getCellByPosition(2, fine +
                             5).String = "Pari a Ribasso del ___________%"
    oSheet.getCellByPosition(2, fine + 8).String = "(ribasso in lettere)"
    oSheet.getCellRangeByPosition(2, fine + 8, 6,
                                  fine + 8).CellStyle = "List-intest_med_c"
    # INSERISCI LA DATA E L'OFFERENTE
    oSheet.getCellByPosition(
        2,
        fine + 10).Formula = '=CONCATENATE("Data, ";TEXT(NOW();"GG/MM/AAAA"))'
    oSheet.getCellRangeByPosition(2, fine + 10, 2,
                                  fine + 10).CellStyle = "Ultimus"
    oSheet.getCellByPosition(2, fine + 12).String = "L'OFFERENTE"
    oSheet.getCellByPosition(2, fine + 12).CellStyle = 'centro_grassetto'
    oSheet.getCellByPosition(2, fine + 13).String = '(timbro e firma)'
    oSheet.getCellByPosition(2, fine + 13).CellStyle = 'centro_corsivo'

    # CONSOLIDA LA DATA
    oRange = oSheet.getCellRangeByPosition(2, fine + 10, 2, fine + 10)
    #  Flags = com.sun.star.sheet.CellFlags.FORMULA
    aSaveData = oRange.getDataArray()
    oRange.setDataArray(aSaveData)
    oSheet.getCellRangeByPosition(
        0, 0,
        SheetUtils.getUsedArea(oSheet).EndColumn,
        SheetUtils.getUsedArea(oSheet).EndRow).CellBackColor = -1
    # imposta stile pagina ed intestazioni
    oSheet.PageStyle = 'PageStyle_COMPUTO_A4'
    pagestyle = oDoc.StyleFamilies.getByName('PageStyles').getByName(
        'PageStyle_COMPUTO_A4')
    pagestyle.HeaderIsOn = True
    # left = pagestyle.RightPageHeaderContent.LeftText.Text

    pagestyle.HeaderIsOn = True
    oHContent = pagestyle.RightPageHeaderContent
    filename = ''  # uno.fileUrlToSystemPath(oDoc.getURL())
    if len(filename) > 50:
        filename = filename[:20] + ' ... ' + filename[-20:]
    oHContent.LeftText.String = filename
    oHContent.CenterText.String = ''
    oHContent.RightText.String = ''.join(''.join(''.join(
        str(datetime.now()).split('.')[0].split(' ')).split('-')).split(':'))
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    pagestyle.RightPageHeaderContent = oHContent
    _gotoCella(0, 1)
    oSheet.Columns.removeByIndex(8, 50)
    oSheet.getCellRangeByPosition(0, 0, 7, 0).CellStyle = "EP-a -Top"
    return


########################################################################
def ins_voce_elenco():
    '''
    Inserisce una nuova riga voce in Elenco Prezzi
    '''
    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(False)

    oSheet = oDoc.CurrentController.ActiveSheet
    _gotoCella(0, 3)
    oSheet.getRows().insertByIndex(3, 1)

    oSheet.getCellByPosition(0, 3).CellStyle = "EP-aS"
    oSheet.getCellByPosition(1, 3).CellStyle = "EP-a"
    oSheet.getCellRangeByPosition(2, 3, 7, 3).CellStyle = "EP-mezzo"
    oSheet.getCellRangeByPosition(8, 3, 9, 3).CellStyle = "EP-sfondo"
    for el in (5, 11, 15, 19, 25):
        oSheet.getCellByPosition(el, 3).CellStyle = "EP-mezzo %"

    for el in (12, 16, 20, 21):  # (12, 16, 20):
        oSheet.getCellByPosition(el, 3).CellStyle = 'EP statistiche_q'

    for el in (13, 17, 23, 24, 25):  # (12, 16, 20):
        oSheet.getCellByPosition(el, 3).CellStyle = 'EP statistiche'

    oSheet.getCellRangeByPosition(0, 3, 26, 3).clearContents(HARDATTR)
    oSheet.getCellByPosition(11,
                             3).Formula = '=LET(s; SUMIF(AA; A4; BB); IF(s <> 0; s; "--"))'
    #  oSheet.getCellByPosition(11, 3).Formula = '=N4/$N$2'
    oSheet.getCellByPosition(12, 3).Formula = '=SUMIF(AA;A4;BB)'
    oSheet.getCellByPosition(13, 3).Formula = '=SUMIF(AA;A4;cEuro)'

    # copio le formule dalla riga sotto
    oRangeAddress = oSheet.getCellRangeByPosition(15, 4, 26,
                                                  4).getRangeAddress()
    oCellAddress = oSheet.getCellByPosition(15, 3).getCellAddress()
    oSheet.copyRange(oCellAddress, oRangeAddress)
    oCell = oSheet.getCellByPosition(2, 3)
    #  valida_cella(
        #  oCell,
        #  '"cad";"corpo";"dm";"dm²";"dm³";"kg";"lt";"m";"m²";"m³";"q";"t";""',
        #  titoloInput='Scegli...',
        #  msgInput='Unità di misura')
    oDoc.enableAutomaticCalculation(True)


########################################################################
def rigenera_voce(lrow=None):
    '''
    Ripristina/ricalcola le formule di descrizione e somma di una voce.
    in COMPUTO, VARIANTE e CONTABILITA
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    except Exception:
        return
    sopra = sStRange.RangeAddress.StartRow
    sotto = sStRange.RangeAddress.EndRow
    # attiva la progressbar
#    progress = Dialogs.Progress(Title='Rigenerazione in corso...', Text="Formule")
#    progress.setLimits(0, sotto - sopra)
#    k = 0
#    progress.setValue(k)
#    progress.show()

    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
#        progress.setValue(10)
        # oSheet.getCellByPosition(1, sopra + 1).CellStyle = 'comp Art-EP_R'
        oSheet.getCellByPosition(2, sopra + 1).Formula = (
                f'=IF(LEN(VLOOKUP(B{sopra+2};elenco_prezzi;2;FALSE()))<($S1.$H$337+$S1.$H$338);'
                f'VLOOKUP(B{sopra+2};elenco_prezzi;2;FALSE());'
                f'CONCATENATE(LEFT(VLOOKUP(B{sopra+2};elenco_prezzi;2;FALSE());$S1.$H$337);'
                f'" [...] ";'
                f'RIGHT(VLOOKUP(B{sopra+2};elenco_prezzi;2;FALSE());$S1.$H$338)))'
            )

        oSheet.getCellByPosition(
            8, sotto).Formula = f'=CONCATENATE("SOMMANO [";VLOOKUP(B{sopra + 2};elenco_prezzi;3;FALSE());"]")'

        oSheet.getCellByPosition(9, sotto).Formula = f'=SUBTOTAL(9;J{sopra +2}:J{sotto +1})'

        oSheet.getCellByPosition(0, sotto).String = ''
        oSheet.getCellByPosition(1, sotto).String = ''
        oSheet.getCellByPosition(2, sotto).String = ''

        oSheet.getCellByPosition(11, sotto).Formula = f'=VLOOKUP(B{sopra + 2};elenco_prezzi;5;FALSE())'
        oSheet.getCellByPosition(13, sotto).Formula = f'=J{sotto + 1}'
        oSheet.getCellByPosition(
            17,
            sotto).Formula = f'=AB{sotto + 1}*J{sotto + 1}'

        oSheet.getCellByPosition(
            18, sotto).Formula = (
                            f'=IF(VLOOKUP(B{sopra + 2};elenco_prezzi;3;FALSE())="%"'
                            f';J{sotto + 1}*L{sotto + 1}/100;J{sotto + 1}*L{sotto + 1})'
                        )
        oSheet.getCellByPosition(27, sotto).Formula = f'=VLOOKUP(B{sopra + 2};elenco_prezzi;4;FALSE())'
        oSheet.getCellByPosition(28, sotto).Formula = f'=S{sotto + 1}-AE{sotto + 1}'
        oSheet.getCellByPosition(29, sotto).Formula = f'=VLOOKUP(B{sopra + 2};elenco_prezzi;6;FALSE())'
        oSheet.getCellByPosition(30, sotto).Formula = f'=IF(AD{sotto + 1}<>""; PRODUCT(AD{sotto + 1}*S{sotto + 1}))'
        oSheet.getCellByPosition(35, sotto).Formula = f'=B{sopra + 2}'
        oSheet.getCellByPosition(36, sotto).Formula = f'=IF(ISERROR(S{sotto + 1});"";IF(S{sotto + 1}<>"";S{sotto + 1};""))'

        formule = []
        for n in range (sopra + 2, sotto):
#            k += 1
#            progress.setValue(k)

            # elimina i collegamenti esterni
            if oSheet.getCellByPosition(2, n).CellStyle == 'comp 1-a' or \
                oSheet.getCellByPosition(2, n).CellStyle == 'comp 1-a ROSSO' and \
                "'" in oSheet.getCellByPosition(2, n).Formula:
                ff = oSheet.getCellByPosition(2, n).Formula.split("'")
                if len(ff) > 1:
                    oSheet.getCellByPosition(2, n).Formula = ff[0] + ff[-1][1:]

            rosso = 0
            for x in range (5, 8):
                if 'ROSSO' in oSheet.getCellByPosition(x, n).CellStyle:
                    rosso = 1
                    break
            if rosso == 1:
                formula = [f'=IF(PRODUCT(E{n+1}:I{n+1})=0;"";-PRODUCT(E{n+1}:I{n+1}))']
            else:
                formula = [f'=IF(PRODUCT(E{n+1}:I{n+1})=0;"";PRODUCT(E{n+1}:I{n+1}))']
            if oSheet.getCellByPosition(4, n).Value < 0:
                formula = [f'=IF(PRODUCT(E{n+1}:I{n+1})=0;"";PRODUCT(E{n+1}:I{n+1}))']
            formule.append(formula)
            if 'Parziale [' in oSheet.getCellByPosition(8, n).Formula:
                oSheet.getCellRangeByPosition(2, n, 7,
                                      n).CellStyle = 'comp sotto centro'
                oSheet.getCellByPosition(8, n).CellStyle = 'comp sotto BiancheS'
                oSheet.getCellByPosition(9, n).CellStyle = 'Comp-Variante num sotto'

        oRange = oSheet.getCellRangeByPosition(9, sopra + 2, 9, sotto - 1)
        formule = tuple(formule)
        # oDoc.CurrentController.select(oRange)
        # DLG.chi(formule)
        oRange.setFormulaArray(formule)

    if oSheet.Name in ('CONTABILITA'):
#        progress.setValue(10)
        oSheet.getCellByPosition(
            2, sopra + 1
        ).Formula = '=IF(LEN(VLOOKUP(B' + str(
            sopra + 2
        ) + ';elenco_prezzi;2;FALSE()))<($S1.$H$335+$S1.$H$336);VLOOKUP(B' + str(
            sopra + 2
        ) + ';elenco_prezzi;2;FALSE());CONCATENATE(LEFT(VLOOKUP(B' + str(
            sopra + 2
        ) + ';elenco_prezzi;2;FALSE());$S1.$H$335);" [...] ";RIGHT(VLOOKUP(B' + str(
            sopra + 2) + ';elenco_prezzi;2;FALSE());$S1.$H$336)))'
        oSheet.getCellByPosition(
            8, sotto - 1
        ).Formula = '=CONCATENATE("Somma positivi e negativi [";VLOOKUP(B' + str(
            sopra + 2) + ';elenco_prezzi;3;FALSE());"]")'
        oSheet.getCellByPosition(
            8, sotto).Formula = '=CONCATENATE("SOMMANO [";VLOOKUP(B' + str(
                sopra + 2) + ';elenco_prezzi;3;FALSE());"]")'
        oSheet.getCellByPosition(
            9, sotto -
            1).Formula = '=IF(SUBTOTAL(9;J' + str(sopra + 2) + ':J' + str(
                sotto) + ')<0;"";SUBTOTAL(9;J' + str(
                    sopra + 2) + ':J' + str(sotto) + '))'
        oSheet.getCellByPosition(
            11, sotto -
            1).Formula = '=IF(SUBTOTAL(9;L' + str(sopra + 2) + ':L' + str(
                sotto) + ')<0;"";SUBTOTAL(9;L' + str(
                    sopra + 2) + ':L' + str(sotto) + '))'
        oSheet.getCellByPosition(
            9, sotto).Formula = '=J' + str(sotto) + '-L' + str(sotto)
        oSheet.getCellByPosition(13, sotto).Formula = '=VLOOKUP(B' + str(
            sopra + 2) + ';elenco_prezzi;5;FALSE())'
        oSheet.getCellByPosition(
            15, sotto).Formula = '=IF(VLOOKUP(B' + str(
                sopra + 2) + ';elenco_prezzi;3;FALSE())="%";J' + str(
                    sotto + 1) + '*N' + str(sotto + 1) + '/100;J' + str(
                        sotto + 1) + '*N' + str(sotto + 1) + ')'
        oSheet.getCellByPosition(
            17,
            sotto).Formula = '=J' + str(sotto + 1) + '*AB' + str(sotto + 1)

        # oSheet.getCellByPosition(
        #     18,
        #     sotto).Formula = f'=J{sotto}'

        oSheet.getCellByPosition(
            35,
            sotto).Formula = f'=B{sopra + 2}'

        oSheet.getCellByPosition(27, sotto).Formula = '=VLOOKUP(B' + str(
            sopra + 2) + ';elenco_prezzi;4;FALSE())'
        oSheet.getCellByPosition(
            28,
            sotto).Formula = '=P' + str(sotto + 1) + '-AE' + str(sotto + 1)
        oSheet.getCellByPosition(29, sotto).Formula = '=VLOOKUP(B' + str(
            sopra + 2) + ';elenco_prezzi;6;FALSE())'
        oSheet.getCellByPosition(
            30, sotto
        ).Formula = '=IF(AD' + str(sotto + 1) + '<>""; PRODUCT(AD' + str(
            sotto + 1) + '*P' + str(sotto + 1) + '))'
        oSheet.getCellRangeByName('A2').Formula = '=P2'
        oSheet.getCellByPosition(9, sotto -
                                 1).CellStyle = 'Comp-Variante num sotto'
        formule = []
        for n in range (sopra + 2, sotto - 1):
#            k += 1
#            progress.setValue(k)

            # elimina i collegamenti esterni
            if oSheet.getCellByPosition(2, n).CellStyle == 'comp 1-a' or \
                oSheet.getCellByPosition(2, n).CellStyle == 'comp 1-a ROSSO' and \
                "'" in oSheet.getCellByPosition(2, n).Formula:
                ff = oSheet.getCellByPosition(2, n).Formula.split("'")
                if len(ff) > 1:
                    oSheet.getCellByPosition(2, n).Formula = ff[0] + ff[-1][1:]

            rosso = 0
            for x in range (5, 8):
                if 'ROSSO' in oSheet.getCellByPosition(x, n).CellStyle:
                    rosso = 1
                    break
            if rosso == 1:
                formule.append (['=IF(PRODUCT(E' + str(n + 1) + ':I' +
                                str(n + 1) + ')>=0;"";PRODUCT(E' +
                                str(n + 1) + ':I' + str(n + 1) + ')*-1)', '',
                                '=IF(PRODUCT(E' + str(n+1) + ':I' +
                                str(n+1) + ')<=0;"";PRODUCT(E' + str(
                                n + 1) + ':I' + str(n+1) + '))'])
            else:
                formule.append (['=IF(PRODUCT(E' + str(n + 1) + ':I' +
                                str(n + 1) + ')<=0;"";PRODUCT(E' +
                                str(n + 1) + ':I' + str(n + 1) + '))', '',
                                '=IF(PRODUCT(E' + str(n+1) + ':I' +
                                str(n+1) + ')>=0;"";PRODUCT(E' + str(
                                n + 1) + ':I' + str(n+1) + ')*-1)'])
            if 'Parziale [' in oSheet.getCellByPosition(8, n).Formula:
                oSheet.getCellRangeByPosition(2, n, 7,
                                            n).CellStyle = "comp sotto centro"
                oSheet.getCellByPosition(8, n).CellStyle = "comp sotto BiancheS"
                oSheet.getCellByPosition(9, n).CellStyle = "Comp-Variante num sotto"

        oRange = oSheet.getCellRangeByPosition(9, sopra + 2, 11, sotto - 2)
        formule = tuple(formule)
        oRange.setFormulaArray(formule)
#    progress.hide()



########################################################################
def rigenera_tutte(arg=None, ):
    '''
    Ripristina le formule in tutto il foglio
    '''
    LeenoSheetUtils.memorizza_posizione()
    with LeenoUtils.DocumentRefreshContext(False):

        chiudi_dialoghi()
        oDoc = LeenoUtils.getDocument()

        riordina_ElencoPrezzi()

        oSheet = oDoc.CurrentController.ActiveSheet
        nome = oSheet.Name
        stili_cat = LeenoUtils.getGlobalVar('stili_cat')


        # attiva la progressbar
        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start(f'Rigenerazione di {nome} in corso...', LeenoSheetUtils.cercaUltimaVoce(oSheet))
        indicator.setValue(0)
        if nome in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
            try:
                oSheet = oDoc.Sheets.getByName(nome)
                row = LeenoSheetUtils.prossimaVoce(oSheet, 0, 1, True)
                oDoc.CurrentController.select(oSheet.getCellByPosition(0, row))
                last = LeenoSheetUtils.cercaUltimaVoce(oSheet)
                while row < last:
                    indicator.setValue(row)
                    rigenera_voce(row)
                    # sistema_stili(row)
                    row = LeenoSheetUtils.prossimaVoce(oSheet, row, 1, True)
            except Exception:
                pass
        rigenera_parziali(True)
        Rinumera_TUTTI_Capitoli2(oSheet)
        numera_voci()
        fissa()
        indicator.end()
    LeenoSheetUtils.ripristina_posizione()


########################################################################

def sistema_stili(lrow=None):
    '''
    Ripristina stili di cella per una singola voce in COMPUTO, VARIANTE e CONTABILITA.
    Ottimizzato per ridurre le chiamate al foglio di calcolo.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    sheet_name = oSheet.Name  # Memorizza il nome per evitare chiamate ripetute



    def _apply_elenco_prezzi_styles(oSheet):
        '''Applica gli stili al foglio "Elenco Prezzi" in modo ottimizzato.
        
        Args:
            oSheet: Il foglio di lavoro su cui applicare gli stili
        '''
        # Ottieni l'intervallo di riferimento e calcola le righe
        oRangeAddress = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
        start_row = oRangeAddress.StartRow + 1  # +1 per offset
        end_row = oRangeAddress.EndRow
        last_row = LeenoSheetUtils.cercaUltimaVoce(oSheet)
        # Definizione degli stili in un dizionario per maggiore chiarezza
        style_ranges = {
            'EP-mezzo': {
                'cols': (2, 3, 4, 6, 7, 8, 9),
                'rows': (start_row, end_row)
            },
            'EP-mezzo %': {
                'cols': (5, 25),
                'rows': (start_row, end_row)
            },
            'EP-aS': {
                'cols': (0, 0),
                'rows': (3, last_row)
            },
            'EP-a': {
                'cols': (1, 1),
                'rows': (3, last_row)
            },
            'EP statistiche_q': [
                {'cols': (11, 13), 'rows': (start_row, end_row)},
                {'cols': (15, 18), 'rows': (start_row, end_row)},
                {'cols': (19, 21), 'rows': (start_row, end_row)},
                {'cols': (23, 24), 'rows': (start_row, end_row)}
            ],
            'Riga_rossa_Chiudi': {
                'cols': (0, 9),  # Da colonna 0 a colonna 9
                'rows': (last_row, last_row)
            }
        }

        # Applicazione degli stili
        for style_name, ranges in style_ranges.items():
            if isinstance(ranges, list):
                # Gestione di range multipli per lo stesso stile
                for range_def in ranges:
                    cols = range_def['cols']
                    rows = range_def['rows']
                    oSheet.getCellRangeByPosition(
                        cols[0], rows[0], cols[1], rows[1]
                    ).CellStyle = style_name
                
            else:
                # Gestione di range singoli
                if isinstance(ranges['cols'], tuple) and len(ranges['cols']) > 2:
                    # Range di colonne multiple (per EP-mezzo)
                    for col in ranges['cols']:
                        oSheet.getCellRangeByPosition(
                            col, ranges['rows'][0], col, ranges['rows'][1]
                        ).CellStyle = style_name
                else:
                    # Range continuo (incluso Riga_rossa_Chiudi)
                    oSheet.getCellRangeByPosition(
                        ranges['cols'][0], ranges['rows'][0],
                        ranges['cols'][1], ranges['rows'][1]
                    ).CellStyle = style_name
        return oSheet
        
        
        # Intestazioni
        headers = {11: 'COMPUTO', 15: 'VARIANTE', 19: 'CONTABILITA'}
        for col, text in headers.items():
            oSheet.getCellByPosition(col, 0).String = text
        
        # Visibilità colonne
        oSheet.getCellRangeByPosition(11, 0, 26, 0).Columns.IsVisible = True
        
        # Applicazione stili in batch per righe 3-last_row
        percent_style_cols = (11, 15, 19, 26)
        qty_style_cols = (12, 16, 20, 23)
        stats_style_cols = (13, 17, 21, 24, 25)
        
        for col in percent_style_cols:
            oSheet.getCellRangeByPosition(col, 3, col, last_row).CellStyle = 'EP-mezzo %'
        
        for col in qty_style_cols:
            oSheet.getCellRangeByPosition(col, 3, col, last_row).CellStyle = 'EP statistiche_q'
        
        for col in stats_style_cols:
            oSheet.getCellRangeByPosition(col, 3, col, last_row).CellStyle = 'EP statistiche'
        


    def _apply_computo_variante_contabilita_styles(oSheet, lrow):
        '''Gestisce gli stili per COMPUTO, VARIANTE e CONTABILITA.'''
        try:
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        except Exception:
            return
        
        sopra = sStRange.RangeAddress.StartRow
        sotto = sStRange.RangeAddress.EndRow
        sheet_name = oSheet.Name

        if sheet_name in ('COMPUTO', 'VARIANTE'):
            _apply_computo_variante_styles(oSheet, sopra, sotto)
        elif sheet_name == 'CONTABILITA':
            _apply_contabilita_styles(oSheet, sopra, sotto)


    def _apply_computo_variante_styles(oSheet, sopra, sotto):
        '''Applica stili per COMPUTO e VARIANTE.'''
        for x in range(sopra + 1, sotto - 1):
            cell_style = oSheet.getCellByPosition(2, x).CellStyle
            if 'comp 1-a' not in cell_style:
                continue
            
            is_negative = oSheet.getCellByPosition(9, x).Value < 0
            style_suffix = ' ROSSO' if is_negative else ''
            
            oSheet.getCellByPosition(9, x).CellStyle = 'Blu'
            oSheet.getCellByPosition(2, x).CellStyle = f'comp 1-a{style_suffix}'
            oSheet.getCellByPosition(5, x).CellStyle = f'comp 1-a PU{style_suffix}'
            oSheet.getCellByPosition(6, x).CellStyle = f'comp 1-a LUNG{style_suffix}'
            oSheet.getCellByPosition(7, x).CellStyle = f'comp 1-a LARG{style_suffix}'
            oSheet.getCellByPosition(8, x).CellStyle = f'comp 1-a peso{style_suffix}'


    def _apply_contabilita_styles(oSheet, sopra, sotto):
        '''Gestisce stili per CONTABILITA.'''
        # Stili fissi per righe specifiche
        oSheet.getCellByPosition(9, sopra + 1).CellStyle = 'vuote2'
        oSheet.getCellByPosition(11, sopra + 1).CellStyle = 'Comp-Bianche in mezzo_R'
        oSheet.getCellByPosition(9, sotto - 1).CellStyle = 'Comp-Variante num sotto'
        oSheet.getCellByPosition(9, sotto).CellStyle = 'Comp-Variante num sotto'
        oSheet.getCellByPosition(13, sotto).CellStyle = 'comp sotto Unitario'
        oSheet.getCellByPosition(15, sotto).CellStyle = 'comp sotto Euro Originale'
        oSheet.getCellByPosition(17, sotto).CellStyle = 'comp sotto Euro Originale'
        oSheet.getCellByPosition(11, sotto - 1).CellStyle = 'Comp-Variante num sotto ROSSO'
        oSheet.getCellByPosition(11, sotto).CellStyle = 'comp sotto centro_R'
        oSheet.getCellByPosition(28, sotto).CellStyle = 'Comp-sotto euri'

        # Applicazione stili dinamici
        for x in range(sopra + 1, sotto):
            if 'comp 1-a' in oSheet.getCellByPosition(2, x).CellStyle:
                oSheet.getCellByPosition(9, x).CellStyle = 'Blu'
            elif oSheet.getCellByPosition(2, x).CellStyle == 'comp sotto centro':
                oSheet.getCellByPosition(9, x).CellStyle = 'Comp-Variante num sotto'

        # Gestione formule e stili condizionali
        for x in range(sopra + 2, sotto - 1):
            test = any(oSheet.getCellByPosition(y, x).String != '' for y in range(2, 8))
            rosso = any('ROSSO' in oSheet.getCellByPosition(y, x).CellStyle for y in range(2, 8))
            
            if test and not rosso:
                _apply_standard_formula_and_style(oSheet, x)
            elif test and rosso:
                _apply_red_formula_and_style(oSheet, x)


    def _apply_standard_formula_and_style(oSheet, x):
        '''Applica stile standard e formula.'''
        oSheet.getCellByPosition(9, x).Formula = f'=IF(PRODUCT(E{x+1}:I{x+1})=0;"";PRODUCT(E{x+1}:I{x+1}))'
        oSheet.getCellByPosition(2, x).CellStyle = 'comp 1-a'
        oSheet.getCellByPosition(5, x).CellStyle = 'comp 1-a PU'
        oSheet.getCellByPosition(6, x).CellStyle = 'comp 1-a LUNG'
        oSheet.getCellByPosition(7, x).CellStyle = 'comp 1-a LARG'
        oSheet.getCellByPosition(8, x).CellStyle = 'comp 1-a peso'
        oSheet.getCellByPosition(11, x).String = ''


    def _apply_red_formula_and_style(oSheet, x):
        '''Applica stile rosso e formula.'''
        oSheet.getCellByPosition(11, x).Formula = f'=IF(PRODUCT(E{x+1}:I{x+1})=0;"";PRODUCT(E{x+1}:I{x+1}))'
        oSheet.getCellByPosition(2, x).CellStyle = 'comp 1-a ROSSO'
        oSheet.getCellByPosition(5, x).CellStyle = 'comp 1-a PU ROSSO'
        oSheet.getCellByPosition(6, x).CellStyle = 'comp 1-a LUNG ROSSO'
        oSheet.getCellByPosition(7, x).CellStyle = 'comp 1-a LARG ROSSO'
        oSheet.getCellByPosition(8, x).CellStyle = 'comp 1-a peso ROSSO'
        oSheet.getCellByPosition(9, x).String = ''

    with LeenoUtils.DocumentRefreshContext(False):
        if sheet_name == 'Elenco Prezzi':
            _apply_elenco_prezzi_styles(oSheet)
        elif sheet_name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
            _apply_computo_variante_contabilita_styles(oSheet, lrow)
    
########################################################################

def rigenera_parziali (arg=False):
    '''
    arg { boolean }: Se False rigenera solo voce corrente
    Rigenera i parziali di tutte le voci
    '''
    oDoc = LeenoUtils.getDocument()
    LeenoUtils.DocumentRefresh(False)
    oSheet = oDoc.CurrentController.ActiveSheet

    if oSheet.Name not in ('COMPUTO', 'CONTABILITA', 'VARIANTE'):
        return

    sopra = 4
    sotto = SheetUtils.getLastUsedRow(oSheet) + 1
    lrow = LeggiPosizioneCorrente()[1]

    if arg == False:
        try:
            sopra = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
        except:
            return
        sotto = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow

    # attiva la progressbar
    indicator = oDoc.getCurrentController().getStatusIndicator()
    indicator.start('Rigenerazione parziali in corso...', sotto - sopra)
    n = 0
    indicator.setValue(n)
    if lrow == True:
        sopra = lrow
    for i in range(sopra, sotto):
        n += 1
        indicator.setValue(n)
        if 'Parziale [' in oSheet.getCellByPosition(8, i).Formula:
            parziale_core(oSheet, i)
    # oDoc.enableAutomaticCalculation(True)
    LeenoUtils.DocumentRefresh(True)
    indicator.end()
    return


########################################################################
def MENU_nuova_voce_scelta():  # assegnato a ctrl-shift-n
    '''
    Contestualizza in ogni tabella l'inserimento delle voci.
    '''
    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        # oDoc.enableAutomaticCalculation(False)
        oSheet = oDoc.CurrentController.ActiveSheet
    #    lrow = LeggiPosizioneCorrente()[1]

        if oSheet.Name in ('COMPUTO', 'VARIANTE'):
            LeenoComputo.ins_voce_computo()
        elif oSheet.Name == 'Analisi di Prezzo':
            inizializza_analisi()
        elif oSheet.Name == 'CONTABILITA':
            # LeenoContab.insertVoceContabilita(oSheet, lrow)  <<< non va
            ins_voce_contab()
        elif oSheet.Name == 'Elenco Prezzi':
            ins_voce_elenco()
        elif oDoc.getSheets().hasByName('GIORNALE_BIANCO'):
            LeenoGiornale.MENU_nuovo_giorno()


# nuova_voce_contab  ##################################################
def ins_voce_contab(lrow=0, arg=1, cod=None):
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoContab.insertVoceContabilita
    Inserisce una nuova voce in CONTABILITA.
    '''
    oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet
    oSheet = oDoc.Sheets.getByName('CONTABILITA')

    stili_contab = LeenoUtils.getGlobalVar('stili_contab')
    stili_cat = LeenoUtils.getGlobalVar('stili_cat')

    if lrow == 0:
        lrow = LeggiPosizioneCorrente()[1]
        if oSheet.getCellByPosition(0, lrow + 1).CellStyle == 'uuuuu':
            return

    try:
        # controllo che non ci siano atti registrati
        partenza = cerca_partenza()
        if partenza[2] == '#reg':
            sblocca_cont()
            if LeenoUtils.getGlobalVar('sblocca_computo') == 0:
                return
        else:
            pass
        ###
    except Exception:
        pass

    stile = oSheet.getCellByPosition(0, lrow).CellStyle
    nSal = 0
    if stile == 'Ultimus_centro_bordi_lati':
        i = lrow
        while i != 0:
            if oSheet.getCellByPosition(23, i).Value != 0:
                nSal = int(oSheet.getCellByPosition(23, i).Value)
                break
            i -= 1
        while oSheet.getCellByPosition(0, lrow).CellStyle == stile:
            lrow += 1
        if oSheet.getCellByPosition(0, lrow).CellStyle == 'uuuuu':
            lrow += 1
            #  nSal += 1
        #  else
    elif stile == 'Comp TOTALI':
        pass
    if stile in stili_cat:
        lrow += 1
    elif stile in (stili_contab):
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        nSal = int(oSheet.getCellByPosition(23, sStRange.RangeAddress.StartRow + 1).Value)
        lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow)

    oSheetto = oDoc.getSheets().getByName('S5')
    oRangeAddress = oSheetto.getCellRangeByPosition(0, 22, 48, 26).getRangeAddress()
    oCellAddress = oSheet.getCellByPosition(0, lrow).getCellAddress()
    
    oSheet.getRows().insertByIndex(lrow, 5)  # inserisco le righe
    oSheet.copyRange(oCellAddress, oRangeAddress)
    
    _gotoCella(1, lrow + 1)

    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    sopra = sStRange.RangeAddress.StartRow
    data = 0.0
    for n in reversed(range(0, sopra)):
        if oSheet.getCellByPosition(
                1, n).CellStyle == 'Ultimus_centro_bordi_lati':
            break
        if oSheet.getCellByPosition(1, n).CellStyle == 'Data_bianca':
            data = oSheet.getCellByPosition(1, n).Value
            break
    if data == 0.0:
        oSheet.getCellByPosition(1, sopra +
                                 2).Value = date.today().toordinal() - 693594
    else:
        oSheet.getCellByPosition(1, sopra + 2).Value = data
########################################################################
#  sformula = '=IF(LEN(VLOOKUP(B' + str(lrow+2) + ';elenco_prezzi;2;
# FALSE()))<($S1.$H$337+$S1.H338);VLOOKUP(B' + str(lrow+2) +
# ';elenco_prezzi;2;FALSE());CONCATENATE(LEFT(VLOOKUP(B' + str(lrow+2) +
# ';elenco_prezzi;2;FALSE());$S1.$H$337);" [...] ";RIGHT(VLOOKUP(B' +
# str(lrow+2) + ';elenco_prezzi;2;FALSE());$S1.$H$338)))'
#  oSheet.getCellByPosition(2, lrow+1).Formula = sformula
########################################################################
# raggruppo i righi di mirura
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    oCellRangeAddr.StartColumn = 0
    oCellRangeAddr.EndColumn = 0
    oCellRangeAddr.StartRow = lrow + 2
    oCellRangeAddr.EndRow = lrow + 2
    oSheet.group(oCellRangeAddr, 1)
########################################################################

    if oDoc.NamedRanges.hasByName('_Lib_' + str(nSal)):
        if lrow - 1 == oSheet.getCellRangeByName(
                '_Lib_' + str(nSal)).getRangeAddress().EndRow:
            nSal += 1

    oSheet.getCellByPosition(23, sopra + 1).Value = nSal
    oSheet.getCellByPosition(23, sopra + 1).CellStyle = 'Sal'

    oSheet.getCellByPosition(35, sopra + 4).Formula = '=B' + str(sopra + 2)
    oSheet.getCellByPosition(
        36, sopra +
        4).Formula = '=IF(ISERROR(P' + str(sopra + 5) + ');"";IF(P' + str(
            sopra + 5) + '<>"";P' + str(sopra + 5) + ';""))'
    oSheet.getCellByPosition(36, sopra + 4).CellStyle = "comp -controolo"
    if cod:
        oSheet.getCellByPosition(1, sopra + 1).String = cod
    numera_voci(0)
    if cfg.read('Generale', 'pesca_auto') == '1':
        if arg == 0:
            return
        pesca_cod()


########################################################################
def stileCelleElencoPrezzi(oSheet, startRow, endRow, color=None):
    '''Applica gli stili alle celle del foglio Elenco Prezzi in modo ottimizzato.
    
    Args:
        oSheet: Il foglio di lavoro su cui applicare gli stili
        startRow: Riga di inizio dell'intervallo
        endRow: Riga di fine dell'intervallo
        color: Colore opzionale da applicare (non implementato in questa versione)
    '''
    # Mappatura degli stili in formato {stile: lista di tuple (col_start, col_end)}
    style_map = {
        'EP-aS': [(0, 0)],
        'EP-a': [(1, 1)],
        'EP-mezzo': [(2, 4), (6, 7)],
        # 'EP-mezzo %': [(5, 5), (11, 11), (15, 15), (19, 19), (25, 25)],
        'EP-mezzo %': [(23, 25)],
        'EP-sfondo': [(8, 9)],
        # 'EP statistiche_q': [(12, 12), (16, 16), (20, 20), (23, 23)],
        'EP statistiche_q': [(11, 13), (15, 17), (19, 21), (23, 24)],
        'EP statistiche': [(13, 13), (17, 17), (21, 21), (25, 25)]
    }
    # Applica gli stili in batch
    for style_name, ranges in style_map.items():
        for col_start, col_end in ranges:
            oSheet.getCellRangeByPosition(
                col_start, startRow, 
                col_end, endRow
            ).CellStyle = style_name
    if color is not None:
        oSheet.getCellRangeByPosition(0, startRow, 0, endRow).CellBackColor = color

def inizializza_elenco():
    '''
    Riscrive le intestazioni di colonna e le formule dei totali in Elenco Prezzi.
    Versione ottimizzata per performance e leggibilità.
    '''
    with LeenoUtils.DocumentRefreshContext(False):
        # 1. Inizializzazione e configurazioni
        chiudi_dialoghi()
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.Sheets.getByName('Elenco Prezzi')
        oSheet.Columns.removeByIndex(26, 1)

        # Configurazioni costanti
        STILI = {
            'intestazione': 'EP-a -Top',
            'testa': 'comp In testa',
            'percentuale': 'EP-mezzo %',
            'contab': 'EP statistiche_q',
            'default': 'Default'
        }
        
        #ridefinisce area nominata per precauzione
        last_row = LeenoSheetUtils.cercaUltimaVoce(oSheet) +2
        SheetUtils.NominaArea(oDoc, 'Elenco Prezzi', f"$A$3:$AF${last_row}", 'elenco_prezzi')
        SheetUtils.NominaArea(oDoc, 'Elenco Prezzi', f"$A$3:$A${last_row}", 'Lista')

        # 2. Pulisci contenuti
        oCellRangeAddr = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
        ER = oCellRangeAddr.EndRow
        EC = oCellRangeAddr.EndColumn
        oSheet.getCellRangeByPosition(11, 3, EC, ER).clearContents(STRING + VALUE + FORMULA)
        
        # 3. Configurazione base foglio
        oDoc.CurrentController.freezeAtPosition(0, 3)
        oSheet.getCellRangeByPosition(0, 0, 100, 0).CellStyle = STILI['default']
        LeenoSheetUtils.setLarghezzaColonne(oSheet)
        
        # 4. Intestazioni colonne (ottimizzato con mapping)
        INTESTAZIONI = {
            'L1': '', 'P1': '', 'T1': '',
            'B2': 'QUESTA RIGA NON VIENE STAMPATA',
            'A3': 'Codice\nArticolo',
            'B3': 'DESCRIZIONE DEI LAVORI\nE DELLE SOMMINISTRAZIONI',
            'C3': 'Unità\ndi misura',
            'D3': 'Sicurezza\ninclusa',
            'E3': 'Prezzo\nunitario',
            'F3': 'Incidenza\nMdO',
            'G3': 'Importo\nMdO',
            'H3': 'Codice di origine',
            'L3': 'Quantità\nComputo',
            'M3': 'Quantità\nVariante',
            'N3': 'Quantità\nContabilità',
            'P3': 'Scostamento\nVariante\nComputo',
            'Q3': 'Scostamento\nContabilità\nComputo',
            'R3': 'Scostamento\nContabilità\nVariante',
            'T3': 'Importi\nComputo',
            'U3': 'Importi\nVariante',
            'V3': 'Importi\nContabilità',
            'X3': 'IMPORTI\nin più',
            'Y3': 'IMPORTI\nin meno',
            'Z3': 'VAR. %',
        }
        
        # Applica intestazioni in batch
        for cell, text in INTESTAZIONI.items():
            oSheet.getCellRangeByName(f"'Elenco Prezzi'.{cell}").String = text
        
        # 5. Applicazione stili (ottimizzato)
        STILI_RANGE = {
            STILI['testa']: ["'Elenco Prezzi'.A2:Y2"],
            STILI['percentuale']: ["'Elenco Prezzi'.Z2"],
            STILI['intestazione']: ["'Elenco Prezzi'.A3:Z3"],
        }
        
        for style, ranges in STILI_RANGE.items():
            for range_name in ranges:
                oSheet.getCellRangeByName(range_name).CellStyle = style
        
        oSheet.getCellRangeByName('I1:J1').Columns.IsVisible = False
        
        # 6. Configurazione totali
        try:
            y = SheetUtils.uFindStringCol('Fine elenco', 0, oSheet)
        except:
            MENU_inserisci_Riga_rossa()
            y = SheetUtils.uFindStringCol('Fine elenco', 0, oSheet)
        
        # Formule totali (ottimizzate)
        FORMULE_TOTALI = {
            'T2': f'=IF(SUBTOTAL(9;T3:T{y})=0;"--";SUBTOTAL(9;T3:T{y}))',
            'U2': f'=IF(SUBTOTAL(9;U3:U{y})=0;"--";SUBTOTAL(9;U3:U{y}))',
            'V2': f'=IF(SUBTOTAL(9;V3:V{y})=0;"--";SUBTOTAL(9;V3:V{y}))',
            'X2': f'=IF(SUBTOTAL(9;X3:X{y})=0;"--";SUBTOTAL(9;X3:X{y}))',
            'Y2': f'=IF(SUBTOTAL(9;Y3:Y{y})=0;"--";SUBTOTAL(9;Y3:Y{y}))',
        }
        
        for cell, formula in FORMULE_TOTALI.items():
            oSheet.getCellRangeByName(cell).Formula = formula
        
        # 7. Righe di totale finali
        TOTALI_FINALI = {
            15: 'TOTALE',
            19: f'=IF(SUBTOTAL(9;T3:T{y})=0;"--";SUBTOTAL(9;T3:T{y}))',
            20: f'=IF(SUBTOTAL(9;U3:U{y})=0;"--";SUBTOTAL(9;U3:U{y}))',
            21: f'=IF(SUBTOTAL(9;V3:V{y})=0;"--";SUBTOTAL(9;V3:V{y}))',
            23: f'=IF(SUBTOTAL(9;X3:X{y})=0;"--";SUBTOTAL(9;X3:X{y}))',
            24: f'=IF(SUBTOTAL(9;Y3:Y{y})=0;"--";SUBTOTAL(9;Y3:Y{y}))',
            24: f'=IF(SUBTOTAL(9;Y3:Y{y})=0;"--";SUBTOTAL(9;Y3:Y{y}))',
        }
        oSheet.getCellRangeByName(f'L{y+1}:N{y+1}').merge(True)
        oSheet.getCellRangeByName(f'P{y+1}:R{y+1}').merge(True)

        for col, value in TOTALI_FINALI.items():
            oSheet.getCellByPosition(col, y).Formula = value if isinstance(value, str) and value.startswith('=') else value
        
        oSheet.getCellRangeByPosition(10, y, 25, y).CellStyle = STILI['contab']
        
        # 8. Pulizia finale e stili
        y += 1
        for col in ('K', 'O', 'S', 'W'):
            oSheet.getCellRangeByName(f'{col}2:{col}{y}').CellStyle = STILI['default']
        
        oSheet.getCellRangeByPosition(3, 3, 250, y + 10).clearContents(HARDATTR)
        
        # 9. Applicazione stili alle colonne (ottimizzato)
        STILI_COLONNE = {
            'EP-aS': [(0, 0)],
            'EP-a': [(1, 1)],
            'EP-mezzo': [(2, 4), (6, 7)],
            'EP-mezzo %': [(23, 25)],
            'EP-sfondo': [(8, 9)],
            'EP statistiche_q': [(11, 13), (15, 17), (19, 21), (23, 24)]
            # 'EP statistiche': [(13, 13), (17, 17), (21, 21), (25, 25)]
        }
        
        for style_name, ranges in STILI_COLONNE.items():
            for col_start, col_end in ranges:
                oSheet.getCellRangeByPosition(
                    col_start, 3, 
                    col_end, y - 2
                ).CellStyle = style_name



def inserisci_ElencoCosti():
    '''
    Inserisci titolo 'ELENCO DEI COSTI ELEMENTARI' in fondo a Elenco Prezzi.
    '''
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = SheetUtils.uFindStringCol('Fine elenco', 0, oSheet)

    oSheet.getRows().insertByIndex(lrow, 2)
    oSheet.getCellRangeByPosition(0, lrow , 26, lrow).CellStyle = 'EP-a -Top'
    oSheet.getCellRangeByPosition(0, lrow , 26, lrow).Rows.Height = 1000
    oSheet.getCellByPosition(1, lrow).String = 'ELENCO DEI COSTI ELEMENTARI'
    _gotoCella(0, lrow + 1)
    return

########################################################################
def inizializza_computo():
    '''
    Riscrive le intestazioni di colonna e le formule dei totali in COMPUTO.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.Sheets.getByName('COMPUTO')
    oDoc.CurrentController.setActiveSheet(oSheet)
    oDoc.CurrentController.freezeAtPosition(0, 3)

    lRow = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1
    #  arg = "cancella"
    #  if arg == "cancella":
    #  oSheet.Rows.removeByIndex(0, SheetUtils.getUsedArea(oSheet).EndRow+1)

    if LeenoSheetUtils.cercaUltimaVoce(oSheet) == 0:
        oRow = SheetUtils.uFindString("TOTALI COMPUTO", oSheet)
        try:
            lRowE = oRow.CellAddress.Row
        except Exception:
            lRowE = 3
    else:
        lRowE = lRow

    Flags = STRING + VALUE + FORMULA
    oSheet.getCellRangeByPosition(0, 0, 100, 2).clearContents(Flags)
    oSheet.getCellRangeByPosition(12, 0, 16, lRowE).clearContents(Flags)
    oSheet.getCellRangeByPosition(22, 0, 23, lRowE).clearContents(Flags)
    oSheet.getCellRangeByPosition(0, 0, 100, 0).CellStyle = "Default"
    oSheet.getCellRangeByPosition(0, 0, 100, 0).clearContents(HARDATTR)
    oSheet.getCellRangeByPosition(44, 0, 100, lRowE + 10).CellStyle = "Default"
    #  oSheet.getCellRangeByName('C1').CellStyle = 'comp Int_colonna'
    #  oSheet.getCellRangeByName('C1').CellBackColor = 16762855
    #  oSheet.getCellRangeByName('C1').String = 'COMPUTO'

    oSheet.getCellByPosition(2, 1).String = 'QUESTA RIGA NON VIENE STAMPATA'
    oSheet.getCellByPosition(
        17, 1).Formula = '=SUBTOTAL(9;R:R)'  # sicurezza
    oSheet.getCellByPosition(
        18,
        1).Formula = '=SUBTOTAL(9;S:S)'  # importo lavori
    oSheet.getCellByPosition(0, 1).Formula = '=AK2'

    oSheet.getCellByPosition(
        28,
        1).Formula = '=SUBTOTAL(9;AC:AC)'  # importo materiali

    oSheet.getCellByPosition(29,
                             1).Formula = '=AE2/S2'  # Incidenza manodopera %
    oSheet.getCellByPosition(29, 1).CellStyle = "Comp TOTALI %"
    oSheet.getCellByPosition(
        30,
        1).Formula = '=SUBTOTAL(9;AE:AE)'  # importo manodopera
    oSheet.getCellByPosition(36, 1).Formula = '=SUBTOTAL(9;AK:AK)'  # totale computo sole voci senza errori

    oSheet.getCellRangeByPosition(0, 1, 43, 1).CellStyle = "comp In testa"
    oSheet.getCellRangeByPosition(0, 0, 43, 2).merge(False)
    oSheet.getCellRangeByPosition(0, 1, 1, 1).merge(True)

    #  rem riga di intestazione
    oSheet.getCellByPosition(0, 2).String = 'N.'
    oSheet.getCellByPosition(1, 2).String = 'Articolo\nData'
    oSheet.getCellByPosition(
        2, 2).String = 'DESIGNAZIONE DEI LAVORI\nE DELLE SOMMINISTRAZIONI'
    oSheet.getCellByPosition(5, 2).String = 'P.U.\nCoeff.'
    oSheet.getCellByPosition(6, 2).String = 'Lung.'
    oSheet.getCellByPosition(7, 2).String = 'Larg.'
    oSheet.getCellByPosition(8, 2).String = 'Alt.\nPeso'
    oSheet.getCellByPosition(9, 2).String = 'Quantità'
    oSheet.getCellByPosition(11, 2).String = 'Prezzo\nunitario'
    oSheet.getCellByPosition(
        13, 2
    ).String = 'Serve per avere le quantità\nrealizzate "pulite" e sommabili'
    oSheet.getCellByPosition(17, 2).String = 'di cui\nsicurezza'
    oSheet.getCellByPosition(18, 2).String = 'Importo €'
    oSheet.getCellByPosition(
        24, 2
    ).String = 'Incidenza\nsul totale'  # POTREBBE SERVIRE PER INDICARE L'INCIDENZA DI OGNI SINGOLA VOCE
    oSheet.getCellByPosition(27, 2).String = 'Sicurezza\nunitaria'
    oSheet.getCellByPosition(28, 2).String = 'Materiali\ne Noli €'
    oSheet.getCellByPosition(29, 2).String = 'Incidenza\nMdO %'
    oSheet.getCellByPosition(30, 2).String = 'Importo\nMdO'
    oSheet.getCellByPosition(31, 2).String = 'Super Cat'
    oSheet.getCellByPosition(32, 2).String = 'Cat'
    oSheet.getCellByPosition(33, 2).String = 'Sub Cat'
    oSheet.getCellByPosition(34, 2).String = 'tag B'
    oSheet.getCellByPosition(35, 2).String = 'tag C'
    oSheet.getCellByPosition(36, 2).String = 'senza errori'
    oSheet.getCellByPosition(38, 2).String = 'Figure e\nannotazioni'
    oSheet.getCellByPosition(
        43, 2).String = 'riservato per annotare\nil numero della voce'
    oSheet.getCellRangeByPosition(0, 2, 43, 2).CellStyle = 'comp Int_colonna'
    oSheet.getCellByPosition(13, 2).CellStyle = 'COnt_noP'
    oSheet.getCellByPosition(19, 2).CellStyle = 'COnt_noP'
    oSheet.getCellByPosition(36, 2).CellStyle = 'COnt_noP'
    oSheet.getCellByPosition(43, 2).CellStyle = 'COnt_noP'
    # oCell = oSheet.getCellRangeByPosition(0, 0, 43, 2)

    #  rem riga del totale
    oSheet.getCellByPosition(2, lRowE).String = "TOTALI COMPUTO"
    oSheet.getCellByPosition(
        17,
        lRowE).Formula = "=SUBTOTAL(9;R:R)"  # importo sicurezza

        # lRowE).Formula = "=SUBTOTAL(9;R3:R" + str(lRowE +
                                                  # 1) + ")"  # importo sicurezza
    oSheet.getCellByPosition(
        18, lRowE).Formula = "=SUBTOTAL(9;S:S)"  # importo lavori
        # 18, lRowE).Formula = "=SUBTOTAL(9;S3:S" + str(lRowE +
                                                      # 1) + ")"  # importo lavori

    oSheet.getCellByPosition(29, lRowE).Formula = "=AE" + str(
        lRowE + 1) + "/S" + str(lRowE + 1) + ""  # Incidenza manodopera %
    oSheet.getCellByPosition(30, lRowE).Formula = "=SUBTOTAL(9;AE:AE)"  # importo manodopera
    # oSheet.getCellByPosition(30, lRowE).Formula = "=SUBTOTAL(9;AE3:AE" + str(
        # lRowE + 1) + ")"  # importo manodopera
    oSheet.getCellByPosition(36, lRowE).Formula = "=SUBTOTAL(9;AK:Ak)"  # totale computo sole voci senza errori
    # oSheet.getCellByPosition(36, lRowE).Formula = "=SUBTOTAL(9;AK3:AK" + str(
        # lRowE + 1) + ")"  # totale computo sole voci senza errori
    oSheet.getCellRangeByPosition(0, lRowE, 36,
                                  lRowE).CellStyle = "Comp TOTALI"
    oSheet.getCellByPosition(24, lRowE).CellStyle = "Comp TOTALI %"
    oSheet.getCellByPosition(29, lRowE).CellStyle = "Comp TOTALI %"

    LeenoSheetUtils.inserisciRigaRossa(oSheet)

    oSheet = oDoc.Sheets.getByName('S1')
    oSheet.getCellByPosition(9, 190).Formula = "=$COMPUTO.$S$2"
    oSheet = oDoc.Sheets.getByName('M1')
    oSheet.getCellByPosition(3, 0).Formula = "=$COMPUTO.$S$2"
    oSheet = oDoc.Sheets.getByName('S2')
    oSheet.getCellByPosition(4, 0).Formula = "=$COMPUTO.$S$2"
    LeenoSheetUtils.setLarghezzaColonne(oSheet)
    setTabColor(16762855)


########################################################################
def inizializza_analisi():
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoAnalysis.inizializzaAnalisi'
    Se non presente, crea il foglio 'Analisi di Prezzo' ed inserisce la prima scheda
    '''
    with LeenoUtils.DocumentRefreshContext(False):
        chiudi_dialoghi()
        oDoc = LeenoUtils.getDocument()
        SheetUtils.NominaArea(oDoc, 'S5', '$B$108:$P$133', 'blocco_analisi')
        if not oDoc.getSheets().hasByName('Analisi di Prezzo'):
            oDoc.getSheets().insertNewByName('Analisi di Prezzo', 1)
            oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
            oSheet.getCellRangeByPosition(0, 0, 15, 0).CellStyle = 'Analisi_Sfondo'
            oSheet.getCellByPosition(0, 1).Value = 0
            oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
            oDoc.CurrentController.setActiveSheet(oSheet)
            setTabColor(12189608)
            oRangeAddress = oDoc.NamedRanges.blocco_analisi.ReferredCells.RangeAddress
            oCellAddress = oSheet.getCellByPosition(
                0,
                SheetUtils.getUsedArea(oSheet).EndRow).getCellAddress()
            oDoc.CurrentController.select(oSheet.getCellByPosition(0, 2))
            oDoc.CurrentController.select(
                oDoc.createInstance(
                    "com.sun.star.sheet.SheetCellRanges"))  # unselect
            LeenoSheetUtils.setLarghezzaColonne(oSheet)

            LeenoEvents.assegna()
            LeenoSheetUtils.inserisciRigaRossa(oSheet)
            ScriviNomeDocumentoPrincipale()
        else:
            GotoSheet('Analisi di Prezzo')
            oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
            oDoc.CurrentController.setActiveSheet(oSheet)
            lrow = LeggiPosizioneCorrente()[1]
            urow = SheetUtils.getUsedArea(oSheet).EndRow
            if lrow >= urow:
                lrow = LeenoSheetUtils.cercaUltimaVoce(oSheet) - 5
            for n in range(lrow, SheetUtils.getUsedArea(oSheet).EndRow):
                if oSheet.getCellByPosition(
                        0, n).CellStyle == 'An-sfondo-basso Att End':
                    break
            oRangeAddress = oDoc.NamedRanges.blocco_analisi.ReferredCells.RangeAddress
            oSheet.getRows().insertByIndex(n + 2, 26)
            oCellAddress = oSheet.getCellByPosition(0, n + 2).getCellAddress()
            oDoc.CurrentController.select(oSheet.getCellByPosition(0, n + 2 + 1))
            oDoc.CurrentController.select(
                oDoc.createInstance(
                    "com.sun.star.sheet.SheetCellRanges"))  # unselect
        oSheet.copyRange(oCellAddress, oRangeAddress)
    LeenoSheetUtils.memorizza_posizione()
    MENU_struttura_on()
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    LeenoSheetUtils.ripristina_posizione()



########################################################################


def MENU_inserisci_Riga_rossa():
    '''
    Inserisce la riga rossa di chiusura degli elaborati
    Questa riga è un riferimento per varie operazioni
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    LeenoSheetUtils.inserisciRigaRossa(oSheet)

########################################################################
def struct_colore(level):
    '''
    Mette in vista struttura secondo il colore
    level { integer } : specifica il livello di categoria
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    # hriga = oSheet.getCellRangeByName('B4').CharHeight * 65
    #  giallo(16777072,16777120,16777168)
    #  verde(9502608,13696976,15794160)
    #  viola(12632319,13684991,15790335)
    col0 = 16724787  # riga_rossa
    col1 = 16777072  # scuro
    col2 = 16777120  # medio
    col3 = 16777168  # chiaro
    col4 = 12632319  # viola
    # attribuisce i colori
    for y in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
        if oSheet.getCellByPosition(0, y).String == '':
            oSheet.getCellByPosition(0, y).CellBackColor = col4
        elif len(oSheet.getCellByPosition(0, y).String.split('.')) == 3:
            oSheet.getCellByPosition(0, y).CellBackColor = col3
        elif len(oSheet.getCellByPosition(0, y).String.split('.')) == 2:
            oSheet.getCellByPosition(0, y).CellBackColor = col2
        elif len(oSheet.getCellByPosition(0, y).String.split('.')) == 1:
            oSheet.getCellByPosition(0, y).CellBackColor = col1
    if level == 0:
        colore = col1
        myrange = (col1, col0)
    elif level == 1:
        colore = col2
        myrange = (col1, col2, col0)
    elif level == 2:
        colore = col3
        myrange = (col1, col2, col3, col0)
    elif level == 3:
        colore = col0
        myrange = (col1, col2, col3, col4, col0)

    test = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2
    # attiva la progressbar

    indicator = oDoc.getCurrentController().getStatusIndicator()
    if indicator:
        indicator.start("Creazione struttura in corso...", test)  # 100 = max progresso


    lista = []
    for n in range(0, test):
        indicator.Value = n
        if oSheet.getCellByPosition(0, n).CellBackColor == colore:
            sopra = n + 1
            for n in range(sopra + 1, test):
                if oSheet.getCellByPosition(0, n).CellBackColor in myrange:
                    sotto = n - 1
                    lista.append((sopra, sotto))
                    break
    for el in lista:
        oCellRangeAddr.StartRow = el[0]
        oCellRangeAddr.EndRow = el[1]
        oSheet.group(oCellRangeAddr, 1)
    indicator.end()
    return

########################################################################
def struttura_Elenco():
    '''
    Dà una tonalità di colore, diverso dal colore dello stile cella, alle righe
    che non hanno il prezzo, come i titoli di capitolo e sottocapitolo.
    '''
    chiudi_dialoghi()

    # if Dialogs.YesNoDialog(Title='AVVISO!',
    # Text='''Adesso puoi dare ai titoli di capitolo e sottocapitolo
# una tonalità di colore che ne migliora la leggibilità, ma
# il risultato finale dipende dalla struttura dei codici di voce.

# L'operazione potrebbe richiedere del tempo e
# LibreOffice potrebbe sembrare bloccato!

# Vuoi procedere comunque?''') == 0:
        # return

    # LeenoUtils.DocumentRefresh(False)
    with LeenoUtils.DocumentRefreshContext(False):

        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet

        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        for n in (3, 7):
            oCellRangeAddr.StartColumn = n
            oCellRangeAddr.EndColumn = n
            oSheet.group(oCellRangeAddr, 0)

        for i in reversed(range(0, 3)):
            struct_colore(i) # attribuisce i colori

        return

########################################################################

def colora_vecchio_elenco():
    '''
    @@ DA DOCUMENTARE
    '''
    LeenoUtils.DocumentRefresh(False)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #  giallo(16777072,16777120,16777168)
    #  verde(9502608,13696976,15794160)
    #  viola(12632319,13684991,15790335)
    col1 = 16777072
    col2 = 16777120
    col3 = 16777168
    try:
        inizio = SheetUtils.uFindStringCol('ATTENZIONE!', 5, oSheet) + 1
    except:
        inizio = 5
    fine = SheetUtils.getUsedArea(oSheet).EndRow + 1
    for el in range(inizio, fine):
        if len(oSheet.getCellByPosition(2, el).String.split('.')) == 1:
            oSheet.getCellByPosition(2, el).CellBackColor = col1
        if len(oSheet.getCellByPosition(2, el).String.split('.')) == 2:
            oSheet.getCellByPosition(2, el).CellBackColor = col2
        if len(oSheet.getCellByPosition(2, el).String.split('.')) == 3:
            oSheet.getCellByPosition(2, el).CellBackColor = col3
    LeenoUtils.DocumentRefresh(True)


########################################################################


def crea_property_value(name, value):
    """
    Crea e restituisce un oggetto PropertyValue con il nome e il valore specificati.
    name : string : Nome della proprietà
    value : any : Valore della proprietà
    """
    from com.sun.star.beans import PropertyValue
    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


########################################################################


def MENU_importa_stili():
    '''
    Importa tutti gli stili da un documento di riferimento. Se non è
    selezionato, il file di riferimento è il template di leenO.
    '''
    with LeenoUtils.DocumentRefreshContext(False):

        if Dialogs.YesNoDialog(IconType="question",
            Title='Vuoi sostituire gli stili del documento?',
            Text='''

    ► Scegli "Sì" per sostituire gli stili del documento
        selezionando un file di riferimento (facoltativo).
        Se non selezioni alcun file, verranno applicati
        gli stili predefiniti di LeenO.

    ► Scegli "No" per mantenere gli stili attuali.
        
    '''
        ) == 0:
            return
        # Mostra una finestra di dialogo per selezionare il file di riferimento
        filename = Dialogs.FileSelect('Scegli il file di riferimento...', '*.ods')
        if filename is None:
            if not os.path.exists(LeenO_path() + '/template/leeno/Computo_LeenO.ods'):
                filename = LeenO_path() + '/template/leeno/Computo_LeenO.ods'
            else:
                filename = LeenO_path() + '/template/leeno/Computo_LeenO.ods'

        # Carica il documento di riferimento per ottenere gli stili
        rifDoc = DocUtils.loadDocument(filename, Hidden=True)
        stili_celle = {}
        elencoStili = rifDoc.StyleFamilies.getByName("CellStyles").ElementNames

        # Salva il formato numerico di tutti gli stili di cella
        for el in elencoStili:
            try:
                style = rifDoc.StyleFamilies.getByName("CellStyles").getByName(el)
                num = style.NumberFormat
                stili_celle[el] = rifDoc.NumberFormats.getByKey(num).FormatString
            except Exception as e:
                DLG.chi(f"Errore durante il salvataggio del formato numerico per lo stile {el}: {e}")

        # Chiudi il documento di riferimento
        rifDoc.close(True)

        # Ottieni il documento corrente
        oDoc = LeenoUtils.getDocument()
        nome = oDoc.CurrentController.ActiveSheet.Name

        # Carica gli stili dal file di riferimento nel documento corrente
        try:
            oDoc.StyleFamilies.loadStylesFromURL(filename, [])
        except Exception as e:
            DLG.chi(f"Errore durante il caricamento degli stili da {filename}: {e}")
            return
        #  oDoc.lockControllers()
        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start("Importazione stili in corso...", len(stili_celle))  # 100 = max progresso

        # Ripristina il formato numerico di tutte le celle salvate
        for n, el in enumerate(stili_celle.keys(), start=1):
            indicator.Value = n
            try:
                style = oDoc.StyleFamilies.getByName("CellStyles").getByName(el)
                style.NumberFormat = LeenoFormat.getNumFormat(stili_celle[el])
            except Exception as e:
                pass
                #  DLG.chi(f"Errore durante il ripristino del formato numerico per lo stile {el}: {e}")

        # Nascondi la finestra di progresso
        indicator.end()
        #  oDoc.unlockControllers()

        # Torna al foglio originale e riabilita il refresh automatico
        GotoSheet(nome)


########################################################################
def MENU_parziale():
    '''
    Inserisce una riga con l'indicazione della somma parziale.
    '''
    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        lrow = LeggiPosizioneCorrente()[1]
        if oSheet.getCellByPosition(1, lrow-1).CellStyle in ('comp Art-EP_R') or \
            lrow == 0:
            return
        if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
            parziale_core(oSheet, lrow)
            rigenera_parziali(False)
        LeenoSheetUtils.adattaAltezzaRiga(oSheet)


###
def parziale_core(oSheet, lrow):
    '''
    lrow    { double } : id della riga di inserimento
    '''
    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    sopra = sStRange.RangeAddress.StartRow
    # sotto = sStRange.RangeAddress.EndRow
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        if(oSheet.getCellByPosition(2, lrow).String == '' and
            oSheet.getCellByPosition(9, lrow).String == ''):
                pass
        if(oSheet.getCellByPosition(0, lrow).CellStyle == 'comp 10 s' and
           oSheet.getCellByPosition(1, lrow).CellStyle == 'Comp-Bianche in mezzo' and
           oSheet.getCellByPosition(2, lrow).CellStyle == 'comp 1-a' or
           oSheet.getCellByPosition(2, lrow).CellStyle == 'comp 1-a ROSSO' or
           oSheet.getCellByPosition(0, lrow).CellStyle == 'Comp End Attributo'):
            oSheet.getRows().insertByIndex(lrow, 1)

        oSheet.getCellByPosition(1, lrow).CellStyle = 'Comp-Bianche in mezzo'
        oSheet.getCellRangeByPosition(2, lrow, 7,
                                      lrow).CellStyle = 'comp sotto centro'
        oSheet.getCellByPosition(8, lrow).CellStyle = 'comp sotto BiancheS'
        oSheet.getCellByPosition(9, lrow).CellStyle = 'Comp-Variante num sotto'
        oSheet.getCellByPosition(
            8, lrow).Formula = '''=CONCATENATE("Parziale [";VLOOKUP(B''' + str(
                sopra + 2) + ''';elenco_prezzi;3;FALSE());"]")'''
        for i in reversed(range(0, lrow)):
            if oSheet.getCellByPosition(9, i - 1).CellStyle in ('vuote2', 'Comp-Variante num sotto'):
                # i
                break
        oSheet.getCellByPosition(9, lrow).Formula = "=SUBTOTAL(9;J" + str(i) + ":J" + str(lrow + 1) + ")"

    if oSheet.Name in ('CONTABILITA'):
        if(oSheet.getCellByPosition(2, lrow).String == '' and
            oSheet.getCellByPosition(9, lrow).String == '' and
            oSheet.getCellByPosition(11, lrow).String == ''):
                pass

        elif(oSheet.getCellByPosition(0, lrow).CellStyle == "comp 10 s_R" and
             oSheet.getCellByPosition(1, lrow).CellStyle == "Comp-Bianche in mezzo_R" and
             oSheet.getCellByPosition(2, lrow).CellStyle == "comp 1-a" or
             oSheet.getCellByPosition(2, lrow).CellStyle == 'comp 1-a ROSSO' or
             'Somma positivi e negativi [' in oSheet.getCellByPosition(8, lrow).String):
             oSheet.getRows().insertByIndex(lrow, 1)
        elif(oSheet.getCellByPosition(0, lrow).CellStyle == "Comp End Attributo_R" or
             oSheet.getCellByPosition(1, lrow).CellStyle == "Data_bianca" or
             oSheet.getCellByPosition(1, lrow).CellStyle == "comp Art-EP_R"):
            return

        oSheet.getCellByPosition(2, lrow).CellStyle = "comp sotto centro"
        oSheet.getCellRangeByPosition(5, lrow, 7,
                                      lrow).CellStyle = "comp sotto centro"
        oSheet.getCellByPosition(8, lrow).CellStyle = "comp sotto BiancheS"
        oSheet.getCellByPosition(9, lrow).CellStyle = "Comp-Variante num sotto"
        oSheet.getCellByPosition(
            8, lrow).Formula = '=CONCATENATE("Parziale [";VLOOKUP(B' + str(
                sopra + 2) + ';elenco_prezzi;3;FALSE());"]")'

        i = lrow
        while i > 0:
            if oSheet.getCellByPosition(
                    9,
                    i - 1).CellStyle in ('vuote2', 'Comp-Variante num sotto'):
                da = i
                break
            i -= 1
        oSheet.getCellByPosition(
            9, lrow).Formula = '=SUBTOTAL(9;J' + str(da) + ':J' + str(
                lrow + 1) + ')-SUBTOTAL(9;L' + str(da) + ':L' + str(lrow +
                                                                    1) + ')'


########################################################################
def vedi_voce_xpwe(oSheet, lrow, vRif):
    """
    (riga d'inserimento, riga di riferimento)
    """
    with LeenoUtils.DocumentRefreshContext(False):
        try:
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, vRif)
            # sStRange.RangeAddress
            idv = sStRange.RangeAddress.StartRow + 1
            sotto = sStRange.RangeAddress.EndRow
            art = 'B$' + str(idv + 1)
            idvoce = 'A$' + str(idv + 1)
            des = 'C$' + str(idv + 1)
            quantity = 'J$' + str(sotto + 1)
            val = oSheet.getCellByPosition(9, sotto).String
            um = 'VLOOKUP(' + art + ';elenco_prezzi;3;FALSE())'

            #  if oSheet.Name == 'CONTABILITA':
            #  sformula = '=CONCATENATE("";"- vedi voce n.";TEXT(' + idvoce +';"@");" - art. ";' + art + ';" [";' + um + ';"]"'
            #  else:
            sformula = ('=CONCATENATE("";"- vedi voce n.";TEXT(' +
                        idvoce + ';"@");" - art. ";' +
                        art +
                        ';" - ";LEFT(' +
                        des +
                        ';$S1.$H$334);"... [";' +
                        um + ';" ";TEXT(' +
                        quantity + ';"0,00");"]";)')
            oSheet.getCellByPosition(2, lrow).Formula = sformula
            # aggiunge commento, annotazione
            oSheet.Annotations.insertNew(oSheet.getCellByPosition(2, lrow).CellAddress,
            'Se non usi questo rigo di misura per il "Vedi voce precedente", eliminalo ed aggiungine uno nuovo.')
            oSheet.getCellByPosition(4, lrow).Formula = '=' + quantity
            if '-' in val:
                # if oSheet.Name == 'CONTABILITA':
                    # oSheet.getCellByPosition(11, lrow).Formula = (
                        # '=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(lrow + 1) +
                        # ')>=0;"";PRODUCT(E' + str(lrow + 1) + ':I' +
                        # str(lrow + 1) + ')*-1)')
                    # oSheet.getCellByPosition(9, lrow).String = ''
                # if oSheet.Name in ('COMPUTO', 'VARIANTE'):
                    # oSheet.getCellByPosition(9, lrow).Formula = (
                        # '=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(lrow + 1) +
                        # ')=0;"";PRODUCT(E' + str(lrow + 1) + ':I' + str(lrow + 1) + '))')
                for x in range(2, 12):
                    oSheet.getCellByPosition(x, lrow).CellStyle = (
                    oSheet.getCellByPosition(x, lrow).CellStyle + ' ROSSO')
                return '-'
        finally:
            sStRange = None
            return

########################################################################
def MENU_vedi_voce():
    '''
    Inserisce un riferimento a voce precedente sulla riga corrente.
    '''
    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        lrow = LeggiPosizioneCorrente()[1]
        if oSheet.getCellByPosition(2, lrow).String not in ('#N/A', '#RIF!'):
            if oSheet.getCellByPosition(2, lrow).Type.value != 'EMPTY':
                if oSheet.Name in ('COMPUTO', 'VARIANTE'):
                    copia_riga_computo(lrow)
                elif oSheet.Name in ('CONTABILITA'):
                    copia_riga_contab(lrow)
                lrow += 1
        if oSheet.getCellByPosition(2, lrow).CellStyle == 'comp 1-a':
            to = basic_LeenO('ListenersSelectRange.getRange',
                            "Seleziona voce di riferimento o indica n. d'ordine")
            if oSheet.Name not in to:
                to = '$' + oSheet.Name + '.$C$' + str(SheetUtils.uFindStringCol(to, 0, oSheet))
            # try:
            to = int(to.split('$')[-1]) - 1
            # except ValueError:
            #     LeenoUtils.DocumentRefresh(True)
            #     return
            _gotoCella(2, lrow)
            # focus = oDoc.CurrentController.getFirstVisibleRow
            if to < lrow:
                vedi_voce_xpwe(oSheet, lrow, to)
    LeenoSheetUtils.adattaAltezzaRiga()


def strall(el, n=3, pos=0):
    '''
    Allunga una stringa fino a n.
    el  { string }   : stringa di partenza
    n   { int }      : numero di caratteri da aggiungere
    pos { int }      : 0 = prefisso; 1 = suffisso

    '''
    #  el ='o'
    if pos == 0:
        el = n * '0' + el
    else:
        el = el + n * '0'
    return el


########################################################################


def setFormatoNumeri(valore):
    '''
    valore   { integer } : id formato
    attribuisce alla selezione di celle un formato numerico a scelta
    valore = 36 (dd/mm/yyyy)
    '''

    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    oProp = PropertyValue()
    oProp.Name = 'NumberFormatValue'
    oProp.Value = valore
    properties = (oProp, )
    dispatchHelper.executeDispatch(oFrame, '.uno:NumberFormatValue', '', 0,
                                   properties)

def MENU_converti_stringhe():
    '''
    Converte in numeri le stringhe o viceversa.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
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
    if '/' in oSheet.getCellByPosition(sCol, sRow).String:
                setFormatoNumeri(36) #imposta il formato numerico data dd/mm/yyyy
    for y in range(sCol, eCol + 1):
        for x in range(sRow, eRow + 1):
            try:
                if oSheet.getCellByPosition(y, x).Type.value == 'TEXT':
                    if '/' in oSheet.getCellByPosition(y, x).String:
                        oSheet.getCellByPosition(y, x).Formula = '=DATEVALUE("' + oSheet.getCellByPosition(y, x).String + '")'
                        oSheet.getCellByPosition(y, x).Value = oSheet.getCellByPosition(y, x).Value
                    else:
                        oSheet.getCellByPosition(y, x).Value = float(
                            oSheet.getCellByPosition(y,
                                                x).String.replace(',', '.'))
                else:
                    oSheet.getCellByPosition(
                        y, x).String = oSheet.getCellByPosition(y, x).String
            except Exception:
                pass
    return LeenoUtils.DocumentRefresh(True)



########################################################################


def ssUltimus():
    '''
    Scrive la variabile globale che individua il Documento Principale (DCC)
    che è il file a cui giungono le voci di prezzo inviate da altri file
    '''
    # chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('M1'):
        return
    try:
        LeenoUtils.getGlobalVar('oDlgMain').endExecute()
    except NameError:
        pass
    if len(oDoc.getURL()) == 0:
        Dialogs.Exclamation(Title = 'ATTENZIONE!',
        Text='''Prima di procedere, devi salvare il lavoro!
Provvedi subito a dare un nome al file.''')
        salva_come()
        autoexec()
    try:
        LeenoUtils.setGlobalVar('sUltimus', uno.fileUrlToSystemPath(oDoc.getURL()))
    except Exception:
        pass
    # DlgMain()
    return

########################################################################
# import pdb; pdb.set_trace() #debugger
########################################################################
# codice di Manuele Pesenti #############################################
########################################################################
def get_Formula(n, a, b):
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


def getCellStyle(level, pos):
    """
    level { integer } : livello(1 o 2)
    pos { integer } : posizione cella
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
    return styles[level][pos]


def SubSum(lrow, sub=False):
    """ Inserisce i dati nella riga
    sub { boolean } : specifica se sotto-categoria
    """
    if sub:
        myrange = (
            'livello2 scritta mini',
            'Livello-1-scritta minival',
            'Comp TOTALI',
        )
        level = 2
    else:
        myrange = (
            'Livello-1-scritta mini val',
            'Comp TOTALI',
        )
        level = 1

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
        return
    lrowE = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1
    nextCap = lrowE
    for n in range(lrow + 1, lrowE):
        if oSheet.getCellByPosition(18, n).CellStyle in myrange:
            nextCap = n + 1
            break
    for n, a, b in (
        (
            18,
            lrow + 1,
            nextCap,
        ),
        (
            24,
            lrow + 1,
            lrowE + 1,
        ),
        (
            29,
            lrow + 1,
            lrowE + 1,
        ),
        (
            30,
            lrow + 1,
            nextCap,
        ),
    ):
        oSheet.getCellByPosition(n, lrow).Formula = get_Formula(n, a, b)
        oSheet.getCellByPosition(18, lrow).CellStyle = getCellStyle(level, n)


########################################################################
# GESTIONE DELLE VISTE IN STRUTTURA ####################################
########################################################################


def MENU_filtra_codice():
    '''
    Applica un filtro di visualizzazione sulla base del codice di voce selezionata.
    Lanciando il comando da Elenco Prezzi, il comportamento è regolato dal valore presente nella cella 'C2'
    '''

    # per filtrare la prossima voce
    #  oDoc = LeenoUtils.getDocument()
    #  lrow = LeggiPosizioneCorrente()[1]
    #  oSheet = oDoc.CurrentController.ActiveSheet
    #  _gotoCella(2, LeenoSheetUtils.prossimaVoce(oSheet, lrow))
    with LeenoUtils.DocumentRefreshContext(False):
        LeenoSheetUtils.memorizza_posizione()
        filtra_codice()
        LeenoSheetUtils.ripristina_posizione()

def filtra_codice(voce=None):
    '''
    Applica un filtro di visualizzazione sulla base del codice di voce selezionata.
    Lanciando il comando da Elenco Prezzi, il comportamento è regolato dal valore presente nella cella 'C2'
    '''
    oDoc = LeenoUtils.getDocument()

    oSheet = oDoc.CurrentController.ActiveSheet

    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_contab = LeenoUtils.getGlobalVar('stili_contab')

    if oSheet.Name == "Elenco Prezzi":
        oCell = oSheet.getCellRangeByName('C2')
        voce = oDoc.Sheets.getByName('Elenco Prezzi').getCellByPosition(
            0, LeggiPosizioneCorrente()[1]).String

        # colora la descrizione scelta
        # oDoc.Sheets.getByName('Elenco Prezzi').getCellByPosition(
            # 1, LeggiPosizioneCorrente()[1]).CellBackColor = 16777120

        if oCell.String == '<DIALOGO>' or oCell.String == '':
            try:
                elaborato = DLG.ScegliElaborato('Ricerca di ' + voce)
                GotoSheet(elaborato)
            except Exception:
                return
        else:
            elaborato = oSheet.getCellByPosition(2, 1).String
            try:
                GotoSheet(elaborato)
            except Exception:
                return
        oSheet = oDoc.Sheets.getByName(elaborato)
        _gotoCella(0, 6)
        LeenoSheetUtils.prossimaVoce(oSheet, LeggiPosizioneCorrente()[1], 1)
    oSheet.clearOutline()
    lrow = LeggiPosizioneCorrente()[1]
    if oSheet.getCellByPosition(0, lrow).CellStyle in (stili_computo + stili_contab):
        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct(
            'com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        sopra = sStRange.RangeAddress.StartRow
        if not voce:
            voce = oSheet.getCellByPosition(1, sopra + 1).String
    else:
        # DLG.MsgBox('Devi prima selezionare una voce di misurazione.', 'Avviso!')
        Dialogs.Exclamation(Title = 'ATTENZIONE!',
        Text='''Devi prima selezionare una voce di misurazione.''')
        return
    fine = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1

    indicator = oDoc.getCurrentController().getStatusIndicator()
    indicator.start('Applicazione filtro...', fine)
    indicator.setValue(0)

    qui = None
    lista_pt = []
    _gotoCella(0, 0)
    for n in range(0, fine):
        indicator.setValue(n)
        if oSheet.getCellByPosition(0,
                                    n).CellStyle in ('Comp Start Attributo',
                                                     'Comp Start Attributo_R'):
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, n)
            sopra = sStRange.RangeAddress.StartRow
            sotto = sStRange.RangeAddress.EndRow
            if oSheet.getCellByPosition(1, sopra + 1).String != voce:
                lista_pt.append((sopra, sotto))
            else:

                # # colora lo sfondo della voce filtrata
                # oSheet.getCellRangeByPosition(0, sopra, 40, sotto).CellBackColor = 16777120

                if qui == None:
                    qui = sopra + 1
    indicator.setValue(fine)
    for el in lista_pt:
        oCellRangeAddr.StartRow = el[0]
        oCellRangeAddr.EndRow = el[1]
        oSheet.group(oCellRangeAddr, 1)
        oSheet.getCellRangeByPosition(0, el[0], 0,
                                      el[1]).Rows.IsVisible = False

    # LeenoSheetUtils.adattaAltezzaRiga(oSheet)

    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    oCellRangeAddr.StartColumn = 29
    oCellRangeAddr.EndColumn = 30
    oSheet.group(oCellRangeAddr, 0)
    oSheet.getCellRangeByPosition(29, 0, 30, 0).Columns.IsVisible = False

    try:
        _gotoCella(1, qui)
    except:
        struttura_off()
        indicator.end()
        GotoSheet("Elenco Prezzi")
        Dialogs.Exclamation(Title = 'Ricerca conclusa', Text='Nessuna corrispondenza trovata')
    indicator.end()

########################################################################

def MENU_struttura_on():
    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        # LeenoUtils.DocumentRefresh(False)
        oSheet = oDoc.CurrentController.ActiveSheet

        if oSheet.Name in ('COMPUTO', 'VARIANTE'):
            struttura_ComputoM()
        elif oSheet.Name == 'Elenco Prezzi':
            struttura_off()
            struttura_Elenco()
        elif oSheet.Name == 'Analisi di Prezzo':
            struttura_Analisi()
        elif oSheet.Name in ('CONTABILITA', 'Registro', 'SAL'):
            LeenoContab.struttura_CONTAB()
        # LeenoUtils.DocumentRefresh(True)

def struttura_ComputoM():
    '''
    Configura la struttura del foglio di computo metrico nascondendo e raggruppando
    colonne specifiche per una visualizzazione ottimizzata.
    
    La funzione esegue le seguenti operazioni:
    1. Nasconde le colonne AC e AD (indici 29-30) e le raggruppa al livello 0
    2. Raggruppa le colonne di misura (F-I, indici 5-8) al livello 0 e le rende visibili
    3. Mostra un indicatore di progresso durante l'operazione
    4. Esegue la funzione struct() per 4 livelli di struttura
    
    Returns:
        None
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()

    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    oCellRangeAddr.StartColumn = 29
    oCellRangeAddr.EndColumn = 30
    oSheet.group(oCellRangeAddr, 0)
    oSheet.getCellRangeByPosition(29, 0, 30, 0).Columns.IsVisible = False

    # Raggruppa le colonne di misura e mostra
    oCellRangeAddr.StartColumn = 5
    oCellRangeAddr.EndColumn = 8
    oSheet.group(oCellRangeAddr, 0)
    for i in range(5, 9):
        oSheet.getColumns().getByIndex(i).Columns.IsVisible = True

    # # attiva la prog
    indicator = oDoc.getCurrentController().getStatusIndicator()
    indicator.start('Creazione vista struttura in corso...', 4)
    for n in range(0, 4):
        indicator.Value = n
        struct(n)
    indicator.end()


def struttura_Analisi():
    '''
    Configura la struttura del foglio di analisi dei prezzi raggruppando
    le righe in base allo stile delle celle nella colonna A.
    
    La funzione esegue le seguenti operazioni:
    1. Rimuove tutti gli interruzioni di pagina manuali
    2. Pulisce la struttura esistente del foglio
    3. Identifica e raggruppa i blocchi di righe che NON hanno stile 'An-1_sigla'
       nella colonna A, creando gruppi annidabili (livello 1)
    4. I gruppi vengono creati tra le righe che non hanno lo stile specificato,
       interrotti quando viene trovata una riga con stile 'An-1_sigla'
    
    Il raggruppamento avviene su tutte le colonne del foglio (dalla A all'ultima colonna)
    
    Returns:
        None
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.removeAllManualPageBreaks()

    oSheet.clearOutline()

    lrow = SheetUtils.getLastUsedRow(oSheet)
    
    start_group = None
    
    for row in range(3, lrow + 1):
        cell_style = oSheet.getCellByPosition(0, row).CellStyle
        
        oRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oRangeAddr.Sheet = oSheet.RangeAddress.Sheet
        
        # Se la cella NON ha lo stile 'An-1_sigla', inizia/continua un gruppo
        if cell_style != 'An-1_sigla':
            if start_group is None:
                start_group = row  # Inizio del blocco
        else:
            # Se troviamo una cella con lo stile 'An-1_sigla' e c'è un gruppo aperto, chiudilo
            if start_group is not None:
                oRangeAddr.StartColumn = 0
                oRangeAddr.EndColumn = oSheet.Columns.Count - 1
                oRangeAddr.StartRow = start_group
                oRangeAddr.EndRow = row - 1  # Fino alla riga precedente
                
                oSheet.group(oRangeAddr, 1)  # Raggruppa
                # oSheet.Rows.getByIndex(start_group).IsVisible = False  # Chiudi il gruppo
                start_group = None  # Resetta
    
    # Gestisci l'ultimo gruppo se rimasto aperto
    if start_group is not None:
        oRangeAddr.StartRow = start_group
        oRangeAddr.EndRow = lrow
        oSheet.group(oRangeAddr, 1)
        oSheet.Rows.getByIndex(start_group).IsVisible = False

def MENU_struttura_off():
    '''
    Cancella la vista in struttura
    '''
    struttura_off()


def struttura_off():
    '''
    Cancella la vista in struttura
    '''
    with LeenoUtils.DocumentRefreshContext(False):
        # LeenoSheetUtils.memorizza_posizione()
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        oSheet.clearOutline()
        # LeenoSheetUtils.ripristina_posizione()



def struct(level, vedi = True):
    ''' mette in vista struttura secondo categorie
    level { integer } : specifica il livello di categoria
    ### COMPUTO/VARIANTE ###
    0 = super-categoria
    1 = categoria
    2 = sotto-categoria
    3 = intera voce di misurazione
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet

    if level == 0: # superCategorie
        stile = 'Livello-0-scritta'
        myrange = (
            'Livello-0-scritta',
            'Comp TOTALI',
        )
        Dsopra = 1
        Dsotto = 1
    elif level == 1: # Categorie
        stile = 'Livello-1-scritta'
        myrange = (
            'Livello-1-scritta',
            'Livello-0-scritta',
            'Comp TOTALI',
        )
        Dsopra = 1
        Dsotto = 1
    elif level == 2: # subCategorie
        stile = 'livello2 valuta'
        myrange = (
            'livello2 valuta',
            'Livello-1-scritta',
            'Livello-0-scritta',
            'Comp TOTALI',
        )
        Dsopra = 1
        Dsotto = 1
    elif level == 3: # misure
        if oSheet.Name == 'CONTABILITA':
            stile = 'Comp Start Attributo_R'
        elif oSheet.Name in ('COMPUTO', 'VARIANTE'):
            stile = 'Comp Start Attributo'

        if oSheet.Name == 'CONTABILITA':
            r1 = 'Comp End Attributo_R'
        elif oSheet.Name in ('COMPUTO', 'VARIANTE'):
            r1 = 'Comp End Attributo'
        myrange = (
            r1,
            'Comp TOTALI',
        )
        Dsopra = 2
        Dsotto = 1

    elif level == 7: # riepilogo
        stile = 'ULTIMUS_3'
        myrange = (
            'ULTIMUS_1',
            'ULTIMUS_2',
            'ULTIMUS_3',
            'ULTIMUS',
        )
        Dsopra = 0
        Dsotto = 0

    elif level == 6: # riepilogo
        stile = 'ULTIMUS_2'
        myrange = (
            'ULTIMUS_1',
            'ULTIMUS_2',
            # 'ULTIMUS_3',
            'ULTIMUS',
        )
        Dsopra = 1
        Dsotto = 1

    elif level == 5: # riepilogo
        stile = 'ULTIMUS_1'
        myrange = (
            'ULTIMUS_1',
            'ULTIMUS',
        )
        Dsopra = 1
        Dsotto = 1
        # for n in(3, 5, 7):
        # oCellRangeAddr.StartColumn = n
        # oCellRangeAddr.EndColumn = n
        # oSheet.group(oCellRangeAddr,0)
        # oSheet.getCellRangeByPosition(n, 0, n, 0).Columns.IsVisible=False

    # test = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2
    test = getLastUsedCell(oSheet).EndRow
    lista_cat = []
    for n in range(0, test):
        if oSheet.getCellByPosition(0, n).CellStyle == stile:
            sopra = n + Dsopra
            for n in range(sopra + 1, test):
                if oSheet.getCellByPosition(0, n).CellStyle in myrange:
                    sotto = n - Dsotto
                    lista_cat.append((sopra, sotto))
                    break
    for el in lista_cat:
        oCellRangeAddr.StartRow = el[0]
        oCellRangeAddr.EndRow = el[1]
        oSheet.group(oCellRangeAddr, 1)
        if vedi == False:
            oSheet.getCellRangeByPosition(0, el[0], 0, el[1]).Rows.IsVisible = False


########################################################################
def MENU_apri_manuale():
    '''
    @@ DA DOCUMENTARE
    '''
    apri = LeenoUtils.createUnoService("com.sun.star.system.SystemShellExecute")
    apri.execute(LeenO_path() + '/MANUALE_LeenO.pdf', "", 0)


########################################################################
def autoexec_off():
    '''
    @@ DA DOCUMENTARE
    '''
    LeenoUtils.DocumentRefresh(False)
    Toolbars.Switch(True)
    oDoc = LeenoUtils.getDocument()
    Toolbars.AllOff()
    LeenoEvents.pulisci()
    Toolbars.On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV', 0)
    for el in ('COMPUTO', 'VARIANTE', 'Elenco Prezzi', 'CONTABILITA',
               'Analisi di Prezzo'):
        try:
            oSheet = oDoc.Sheets.getByName(el)
            oSheet.getCellRangeByName("A1:AT1").clearContents(EDITATTR +
                                                              FORMATTED +
                                                              HARDATTR)
            oSheet.getCellRangeByName("A1").String = ''
        except Exception:
            pass
    LeenoUtils.DocumentRefresh(True)

########################################################################
class trun(threading.Thread):
    '''Avvia processi automatici ad intervalli definiti di tempo'''
    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
        while True:
            minuti = 60 * int(cfg.read('Generale', 'pausa_backup'))
            time.sleep(minuti)
            bak()


def autorun():
    '''
    @@ DA DOCUMENTARE
    '''
    # global utsave
    if int(cfg.read('Generale', 'copie_backup')) != 0:
        utsave = trun()
        utsave._stop()
        utsave.start()
    else:
        try:
            utsave._stop()
            utsave = False
        except:
            pass

def sistema_aree():
    """
    Aggiorna le aree denominate nei fogli COMPUTO, VARIANTE e CONTABILITA.
    Le aree partono sempre dalla riga 3 fino all'ultima riga utilizzata.
    """
    oDoc = LeenoUtils.getDocument()

    # Mappa fogli → (colonna, nome area)
    fogli_aree = {
        'COMPUTO': [
            ('AJ', 'AA'),
            # ('N', 'BB'),
            ('J', 'BB'),
            ('AK', 'cEuro')
        ],
        'VARIANTE': [
            ('AJ', 'varAA'),
            # ('N', 'varBB'),
            ('J', 'varBB'),
            ('AK', 'varEuro')
        ],
        'CONTABILITA': [
            ('AJ', 'GG'),
            ('J', 'G1G1'),
            ('AK', 'conEuro')
        ],
    }

    with LeenoUtils.DocumentRefreshContext(False):
        for nome_foglio, aree in fogli_aree.items():
            if not oDoc.getSheets().hasByName(nome_foglio):
                continue  # Salta se il foglio non esiste

            oSheet = oDoc.getSheets().getByName(nome_foglio)
            lrow = SheetUtils.getUsedArea(oSheet).EndRow

            # Evita di creare aree vuote
            if lrow < 3:
                continue

            for col, nome_area in aree:
                range_str = f'${col}$3:${col}${lrow}'
                SheetUtils.NominaArea(oDoc, nome_foglio, range_str, nome_area)


########################################################################
def autoexec():
    with LeenoUtils.DocumentRefreshContext(False):
        autoexec_run()

def autoexec_run():
    '''
    questa è richiamata da creaComputo()
    '''
    LeenoUtils.DocumentRefresh(False)
    #  LS.importa_stili_pagina_non_presenti() #troppo lenta con file grossi
    LeenoEvents.pulisci()
    inizializza()
    LeenoEvents.assegna()
    
    SheetUtils.remove_bad_ranges()
    SheetUtils.FixNamedArea()

    # rinvia a autoexec in basic
    basic_LeenO('_variabili.autoexec')
    bak0()
    autorun()
    sistema_aree()
    ctx = LeenoUtils.getComponentContext()
    oGSheetSettings = ctx.ServiceManager.createInstanceWithContext("com.sun.star.sheet.GlobalSheetSettings", ctx)
    # Usa i parametri della stampante per la formattazione del testo
    oGSheetSettings.UsePrinterMetrics = True

    # attiva 'copia di backup', ma dall'apertura successiva di LibreOffice
    node = GetRegistryKeyContent("/org.openoffice.Office.Common/Save/Document", True)
    node.CreateBackup = True
    node.commitChanges()
    uso = int(cfg.read('Generale', 'conta_usi')) + 1

    if uso == 10 or (uso % 50) == 0:
        dlg_donazioni()
    cfg.write('Generale', 'conta_usi', str(uso))
    if cfg.read('Generale', 'movedirection') == '0':
        oGSheetSettings.MoveDirection = 0
    else:
        oGSheetSettings.MoveDirection = 1
    oDoc = LeenoUtils.getDocument()
    #  RegularExpressions Wildcards are mutually exclusive, only one can have the value TRUE.
    #  If both are set to TRUE via API calls then the last one set takes precedence.
    try:
        oDoc.Wildcards = False
    except Exception:
        pass
    oDoc.RegularExpressions = False
    # oDoc.CalcAsShown = True  # precisione come mostrato
    adegua_tmpl()  # esegue degli aggiustamenti del template
    oSheet = oDoc.CurrentController.ActiveSheet
    for nome in ('VARIANTE', 'CONTABILITA', 'COMPUTO'):
        try:
            GotoSheet(nome)
            subst_str(' >(', ' ►(')
        except Exception:
            pass
    GotoSheet(oSheet.Name)
    Toolbars.Vedi()
    ScriviNomeDocumentoPrincipale()

    dp()

    LeenoUtils.DocumentRefresh(True)
    if len(oDoc.getURL()) != 0:
        # scegli cosa visualizzare all'avvio:
        vedi = cfg.read('Generale', 'dialogo')
        if vedi == '1':
                DlgMain()

def dp():
    d = {
        'COMPUTO': 'F1',
        'VARIANTE': 'F1',
        'Elenco Prezzi': 'A1',
        'CONTABILITA': 'F1',
        'Analisi di Prezzo': 'A1'
    }
    oDoc = LeenoUtils.getDocument()
    for el in d.keys():
        try:
            oSheet = oDoc.Sheets.getByName(el)
            if LeenoUtils.getGlobalVar('sUltimus') == uno.fileUrlToSystemPath(oDoc.getURL()):
                oSheet.getCellRangeByName(
                    "A1:AT1").CellBackColor = 16773632  # 13434777 giallo
                oSheet.getCellRangeByName(
                    d[el]).String = 'DP: Questo documento'
            else:
                oSheet.getCellRangeByName(
                    "A1:AT1").clearContents(HARDATTR)
                oSheet.getCellRangeByName(
                    d[el]).String = 'DP:' + LeenoUtils.getGlobalVar('sUltimus')

        except Exception as e:
            #  DLG.chi(f"Errore durante l'accesso al foglio '{el}': {e}")
            pass

########################################################################

def vista_configurazione(tipo_configurazione):
    '''
    Configurazione base di colonne in base al tipo (terra_terra o mdo).
    
    tipo_configurazione: str
        - "terra_terra" per configurazione standard COMPUTO e VARIANTE
        - "mdo" per configurazione manodopera
    '''
    struttura_off()
    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        # vRow = oDoc.CurrentController.getFirstVisibleRow()
        LeenoSheetUtils.memorizza_posizione()

        # Verifica il nome del foglio e imposta la colonna iniziale
        if oSheet.Name in ('COMPUTO', 'VARIANTE'):
            col = 46
        elif oSheet.Name == 'CONTABILITA':
            col = 39
        else:
            raise ValueError(f"Nome del foglio non gestito: {oSheet.Name}")

        ncol = oSheet.getColumns().getCount()
        iSheet = oSheet.RangeAddress.Sheet

        # # Funzione helper per creare CellRangeAddress
        def create_cell_range(sheet, start_col, end_col):
            addr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
            addr.Sheet = sheet
            addr.StartColumn = start_col
            addr.EndColumn = end_col
            return addr

        # # Raggruppo le colonne principali
        # oCellRangeAddr = create_cell_range(iSheet, col, ncol)
        # oSheet.group(oCellRangeAddr, 0)

        n = SheetUtils.getLastUsedRow(oSheet)
    # Configurazione specifica per "terra_terra" o "mdo"
        if tipo_configurazione == "terra_terra":
            struct(3)
            # Raggruppa le colonne di MDO e nasconde
            oCellRangeAddr = create_cell_range(iSheet, 29, 30)
            oSheet.group(oCellRangeAddr, 0)
            oSheet.getColumns().getByIndex(28).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(29).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(30).Columns.IsVisible = False

            # Raggruppa le colonne di misura e mostra
            oCellRangeAddr = create_cell_range(iSheet, 5, 8)
            oSheet.group(oCellRangeAddr, 0)
            for i in range(5, 9):
                oSheet.getColumns().getByIndex(i).Columns.IsVisible = True
            LeenoSheetUtils.setLarghezzaColonne(oSheet)

            for el in range(4, n):
                if oSheet.getCellByPosition(2, el).CellStyle == "comp sotto centro" or \
                    oSheet.getCellByPosition(2, el).CellStyle == "comp sotto centro_R":
                    for i in range(3):
                        oSheet.getCellByPosition(i, el).String = ''

        elif tipo_configurazione == "mdo":
            struct(3, vedi=False)

            # Raggruppa le colonne di MDO e mostra
            oCellRangeAddr = create_cell_range(iSheet, 29, 30)
            oSheet.group(oCellRangeAddr, 0)
            oSheet.getColumns().getByIndex(28).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(29).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(30).Columns.IsVisible = True

            # Raggruppa le colonne di misura e nasconde
            oCellRangeAddr = create_cell_range(iSheet, 5, 8)
            oSheet.group(oCellRangeAddr, 0)
            for i in range(5, 9):
                oSheet.getColumns().getByIndex(i).Columns.IsVisible = False
        
            for el in range(4, n):
                if oSheet.getCellByPosition(2, el).CellStyle == "comp sotto centro" or \
                    oSheet.getCellByPosition(2, el).CellStyle == "comp sotto centro_R":
                    oSheet.getCellByPosition(2, el).Formula = oSheet.getCellByPosition(8, el).Formula
            oSheet.getColumns().getByIndex(2).Columns.Width = 7900

        elif tipo_configurazione == "Semplificata":
            oCellRangeAddr = create_cell_range(iSheet, 17, 27)
            # oSheet.group(oCellRangeAddr, 0)
            oSheet.getCellRangeByPosition(17, 0, 27, 0).Columns.IsVisible = False
            for i in range(17, 28):
                oSheet.getColumns().getByIndex(i).Columns.IsVisible = False

            oCellRangeAddr = create_cell_range(iSheet, 31, 37)
            # oSheet.group(oCellRangeAddr, 0)
            oSheet.getCellRangeByPosition(31, 0, 37, 0).Columns.IsVisible = False

            oCellRangeAddr = create_cell_range(iSheet, 29, 30)
            # oSheet.group(oCellRangeAddr, 0)
            oSheet.getCellRangeByPosition(29, 0, 30, 0).Columns.IsVisible = False
            # for i in range(17, 28):
            #     oSheet.getColumns().getByIndex(i).Columns.IsVisible = False
        
            oCellRangeAddr = create_cell_range(iSheet, 28, 28)
            # oSheet.group(oCellRangeAddr, 0)
            oSheet.getColumns().getByIndex(28).Columns.IsVisible = False



        # Ripristina la prima riga visibile e aggiorna larghezza colonne
        # oDoc.CurrentController.setFirstVisibleRow(vRow)
        LeenoSheetUtils.ripristina_posizione()

def vista_terra_terra():
    vista_configurazione('terra_terra')

def vista_mdo():
    vista_configurazione('mdo')


########################################################################
def Menu_vSintetica():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    flag = True
    for i in range(1000):
        if oSheet.getCellByPosition(0, i).CellStyle == "Comp End Attributo" and \
            oSheet.getCellByPosition(0, i).String != '':
            flag = False
            break
    vSintetica(flag)

def vSintetica(flag = True):
    struttura_off()
    with LeenoUtils.DocumentRefreshContext(False):
        """
        Vista sintetica su COMPUTO, VARIANTE o CONTABILITA.
        Aggiorna le voci in base al foglio attivo e nasconde quelle con valori zero.
        """

        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet

        if oSheet.Name in ('COMPUTO', 'VARIANTE'):
            uRiga = SheetUtils.uFindStringCol('TOTALI COMPUTO', 2, oSheet)
            lcol = 18
            col_start = 29
            col_end = 30
        elif oSheet.Name == 'CONTABILITA':
            uRiga = SheetUtils.uFindStringCol('T O T A L E', 2, oSheet)
            lcol = 15
            col_start = 19
            col_end = 30

        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet

        i = 0
        uscita = False
        # attiva la progressbar
        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start(f'Elaborazione voci di {oSheet.Name}...', uRiga)

        while (i < uRiga):
            i = LeenoSheetUtils.prossimaVoce(oSheet, i)
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, i)
            try:
                sotto = sStRange.RangeAddress.EndRow
                sopra = sStRange.RangeAddress.StartRow
            except:
                break
            if oSheet.Name in ('COMPUTO', 'VARIANTE'):
                dati = LeenoComputo.datiVoceComputo(oSheet, i + 1)
            elif oSheet.Name == 'CONTABILITA':
                    dati = LeenoComputo.datiVoceComputo(oSheet, i + 1)[1]

            if flag:
                for col in range(3):
                    oSheet.getCellByPosition(col, sotto).Formula = f'{dati[col]}'
                oSheet.getCellRangeByPosition(2, sotto, 7, sotto).merge(True)
                oSheet.getCellByPosition(8, sotto).Formula = f'=CONCATENATE("[";VLOOKUP(B{sopra + 2};elenco_prezzi;3;FALSE());"]")'
                oSheet.getRows().getByIndex(sotto).Rows.Height = 750

                oCellRangeAddr.StartRow = sopra
                oCellRangeAddr.EndRow = sotto -1
                oSheet.group(oCellRangeAddr, 1)
                oSheet.getCellRangeByPosition(lcol, sopra, lcol, sotto -1).Rows.IsVisible = False

            else:
                oSheet.getCellRangeByPosition(2, sotto, 8, sotto).merge(False)
                if "SOMMANO [" in oSheet.getCellByPosition(8, sotto).String:
                    break                
                oSheet.getCellByPosition(8, sotto).Formula = f'=CONCATENATE("SOMMANO [";VLOOKUP(B{sopra + 2};elenco_prezzi;3;FALSE());"]")'
                for col in range(3):
                    oSheet.getCellByPosition(col, sotto).String = ''
            if uscita:
                # DLG.chi(sotto)
                break
            i += 1
            indicator.Value = i
        indicator.end()

        if flag:
            oCellRangeAddr.StartColumn = col_start
            oCellRangeAddr.EndColumn = col_end
            oSheet.group(oCellRangeAddr, 0)
            for col in range(28, 31):
                oSheet.getColumns().getByIndex(col).Columns.IsVisible = False
            Dialogs.Exclamation(Title='AVVISO!',
                Text='''QUESTA È SOLO UNA MODALITÀ DI VISUALIZZAZIONE!

    Apportare modifiche in questa modalità potrebbe
    compromettere l'integrità dei dati contenuti in questo foglio.

    Si ricorda che le celle con sfondo bianco, in genere,
    non sono destinate all'inserimento di dati.''')
        return


########################################################################


def catalogo_stili_cella():
    '''
    Apre un nuovo foglio e vi inserisce tutti gli stili di cella
    con relativo esempio
    '''
    oDoc = LeenoUtils.getDocument()
    sty = oDoc.StyleFamilies.getByName("CellStyles").getElementNames()
    if oDoc.Sheets.hasByName("stili"):
        oSheet = oDoc.Sheets.getByName("stili")
    else:
        sheet = oDoc.createInstance("com.sun.star.sheet.Spreadsheet")
        oDoc.Sheets.insertByName('stili', sheet)
        oSheet = oDoc.Sheets.getByName("stili")
    GotoSheet("stili")
    # attiva la progressbar
    indicator = oDoc.getCurrentController().getStatusIndicator()
    indicator.start('Creazione catalogo stili di cella in corso...', len(sty))
    i = 0
    sty = sorted(sty)
    for el in sty:
        oSheet.getCellByPosition( 0, i).String = el
        oSheet.getCellByPosition( 1, i).CellStyle = el
        oSheet.getCellByPosition( 3, i).CellStyle = el
        oSheet.getCellByPosition( 1, i).Value = -2000
        oSheet.getCellByPosition( 3, i).String = "LeenO"
        i += 1
        indicator.setValue(i)
    indicator.end()


def elimina_stili_cella():
    '''
    Elimina gli stili di cella non utilizzati.
    '''
    LeenoUtils.DocumentRefresh(False)
    oDoc = LeenoUtils.getDocument()
    stili = oDoc.StyleFamilies.getByName('CellStyles').getElementNames()

    # Crea una lista di stili non utilizzati
    stili_da_elim = [el for el in stili if not oDoc.StyleFamilies.getByName('CellStyles').getByName(el).isInUse()]
    #  stili_da_elim = stili # RIMUOVI TUTTI!!!

    # Rimuovi gli stili non utilizzati
    n = 0
    for el in stili_da_elim:
        oDoc.StyleFamilies.getByName('CellStyles').removeByName(el)
        n += 1
    Dialogs.Exclamation(Title = 'ATTENZIONE!', Text=f'Eliminati {n} stili di cella!')
    LeenoUtils.DocumentRefresh(True)

def elenca_stili_foglio():
    '''
    Restituisce l'elenco di tutti gli stili di cella applicati alle celle nel foglio corrente.
    '''
    try:
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        
        # Set per tenere traccia degli stili unici applicati
        stili_applicati = set()

        # Ottieni l'area utilizzata nel foglio
        area_utilizzata = SheetUtils.getUsedArea(oSheet)
        row = area_utilizzata.EndRow
        col = area_utilizzata.EndColumn

        # Itera sulle celle dell'area utilizzata
        for riga in range(row + 1):  # Includi l'ultima riga
            for colonna in range(col + 1):  # Includi l'ultima colonna
                cella = oSheet.getCellByPosition(colonna, riga)
                stile = cella.CellStyle
                if stile:  # Controlla se lo stile non è vuoto
                    stili_applicati.add(stile)

        # Converti il set in una lista
        lista_stili_applicati = list(stili_applicati)

        # Mostra o restituisci la lista degli stili applicati
        #  DLG.chi(f'Stili di cella applicati: {", ".join(lista_stili_applicati)}')
        return lista_stili_applicati

    except Exception as e:
        DLG.errore(e)
        return []


def elimina_stile():
    '''
    Elimina lo stile della cella selezionata.
    '''
    stili_utili = elenca_stili_foglio().append('Default')
    try:
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        selezione = oDoc.getCurrentSelection()
        stile = selezione.CellStyle
        DLG.chi(stile)
        selezione.CellStyle = 'Default' # Assegna lo stile predefinito alla cella
        # Rimuovi lo stile
        if stile not in stili_utili:
            oDoc.StyleFamilies.getByName('CellStyles').removeByName(stile)
    except Exception as e:
        DLG.errore(e)
        pass

    # Ottieni la posizione attuale del cursore
    cella_corrente = selezione.getCellAddress()
    nuova_riga = cella_corrente.Row + 1  # Sposta di una riga in basso

    # Assicurati di non uscire dall'intervallo delle righe del foglio
    if nuova_riga < oSheet.Rows.Count:
        # Sposta il cursore alla cella nella stessa colonna ma una riga sotto
        oDoc.CurrentController.select(oSheet.getCellByPosition(cella_corrente.Column, nuova_riga))

########################################################################
def inizializza():
    '''
    Inserisce tutti i dati e gli stili per preparare il lavoro.
    lanciata in autoexec()
    '''
    oDoc = LeenoUtils.getDocument()

    #  oDoc.IsUndoEnabled = False
    oDoc.getSheets().getByName('copyright_LeenO').getCellRangeByName(
        'A3').String = '# © 2001-2013 Bartolomeo Aimar - © 2014-' + str(
            datetime.now().year) + ' Giuseppe Vizziello'

    oUDP = oDoc.getDocumentProperties().getUserDefinedProperties()
    oSheet = oDoc.getSheets().getByName('S1')

    oSheet.getCellRangeByName('G219').String = 'Copyright 2014-' + str(datetime.now().year)

    # allow non-numeric version codes (example: testing)
    rvc = version_code.read().split('-')[0]
    if isinstance(rvc, str):
        oSheet.getCellRangeByName('H194').String = rvc
    else:
        oSheet.getCellRangeByName('H194').Value = rvc

    oSheet.getCellRangeByName('I194').Value = LeenoUtils.getGlobalVar('Lmajor')
    oSheet.getCellRangeByName('J194').Value = LeenoUtils.getGlobalVar('Lminor')
    oSheet.getCellRangeByName('H291').Value = oUDP.Versione
    oSheet.getCellRangeByName('I291').String = oUDP.Versione_LeenO.split('.')[0]
    oSheet.getCellRangeByName('J291').String = oUDP.Versione_LeenO.split('.')[1]

    oSheet.getCellRangeByName('H295').String = oUDP.Versione_LeenO.split('.')[0]
    oSheet.getCellRangeByName('I295').String = oUDP.Versione_LeenO.split('.')[1]
    oSheet.getCellRangeByName('J295').String = oUDP.Versione_LeenO.split('.')[2]

    oSheet.getCellRangeByName('K194').String = LeenoUtils.getGlobalVar('Lsubv')
    oSheet.getCellRangeByName('H296').Value = LeenoUtils.getGlobalVar('Lmajor')
    oSheet.getCellRangeByName('I296').Value = LeenoUtils.getGlobalVar('Lminor')
    oSheet.getCellRangeByName('J296').String = LeenoUtils.getGlobalVar('Lsubv')

    if oDoc.getSheets().hasByName('CONTABILITA'):
        oSheet.getCellRangeByName('H328').Value = 1
    else:
        oSheet.getCellRangeByName('H328').Value = 0

# inizializza la lista di scelta in elenco Prezzi
    oCell = oDoc.getSheets().getByName('Elenco Prezzi').getCellRangeByName('C2')
    valida_cella(oCell,
                 '"<DIALOGO>";"COMPUTO";"VARIANTE";"CONTABILITA"',
                 titoloInput='Scegli...',
                 msgInput='Applica Filtra Codice a...',
                 err=True)
    oCell.String = "<DIALOGO>"
    oCell.CellStyle = 'EP-aS'
    oCell = oDoc.getSheets().getByName('Elenco Prezzi').getCellRangeByName('C1')
    oCell.String = "Applica Filtro a:"
    oCell.CellStyle = 'EP-aS'
# inizializza la lista di scelta per la copertona cP_Cop
    oCell = oDoc.getSheets().getByName('cP_Cop').getCellRangeByName('B19')
    # if oCell.String == '';
    valida_cella(oCell,
                 '"ANALISI DI PREZZO";"ELENCO PREZZI";"ELENCO PREZZI E COSTI ELEMENTARI";\
                 "COMPUTO METRICO";"PERIZIA DI VARIANTE";"LIBRETTO DELLE MISURE";"REGISTRO DI CONTABILITÀ";\
                 "S.A.L. A TUTTO IL"',
                 titoloInput='Scegli...',
                 msgInput='Titolo della copertina...',
                 err=False)
    # oCell.String = ""
    # Indica qual è il Documento Principale
    ScriviNomeDocumentoPrincipale()
    # nascondi_sheets()


########################################################################
def adegua_tmpl():
    '''
    Mantengo la compatibilità con le vecchie versioni del template:
    - dal 200 parte di autoexec è in python
    - dal 203(LeenO 3.14.0 ha templ 202) introdotta la Super Categoria con nuovi stili di cella;
        sostituita la colonna "Tag A" con "Tag Super Cat"
    - dal 207 introdotta la colonna dei materiali in computo e contabilità
    - dal 209 cambia il nome di proprietà del file in "Versione_LeenO"
    - dal 211 cambiano le formule del prezzo unitario e dell'importo in Computo e Contabilità
    - dal 212 vengono cancellate le celle che indicano il DCC nel foglio M1
    - dal 213 sposta il VediVoce nella colonna E
    - dal 214 assegna un'approssimazione diversa per ognuno dei valori di misurazione
    - dal 215 adegua del formule degli importi ai prezzi in %
    - dal 216 aggiorna le formule in CONTABILITA
    - dal 217 aggiorna le formule in COMPUTO
    '''
    # LeenoUtils.DocumentRefresh(False)
    oDoc = LeenoUtils.getDocument()
    # oDoc.enableAutomaticCalculation(False)
    # LE VARIABILI NUOVE VANNO AGGIUNTE IN config_default()
    # cambiare stile http://bit.ly/2cDcCJI
    ver_tmpl = oDoc.getDocumentProperties().getUserDefinedProperties().Versione
    if ver_tmpl > 200:
        basic_LeenO('_variabili.autoexec')  # rinvia a autoexec in basic

    adegua_a = 217  # VERSIONE CORRENTE

    if ver_tmpl < adegua_a:
        if Dialogs.YesNoDialog(Title='Informazione',
        Text= '''Vuoi procedere con l'adeguamento di questo file
alla versione di LeenO installata?''') == 0:
            Dialogs.Exclamation(Title = 'Avviso!',
                         Text='''Non avendo effettuato l'adeguamento del file alla versione
di LeenO installata, potresti avere dei malfunzionamenti!''')

            return
        sproteggi_sheet_TUTTE()
        if oDoc.getSheets().hasByName('S4'):
            oDoc.Sheets.removeByName('S4')
        # attiva la progressbar

        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start('Adeguamento del lavoro in corso...', 10)
        indicator.setValue(0)

        ############
        # aggiungi stili di cella
        indicator.setValue(1)
        for el in ('comp 1-a PU', 'comp 1-a LUNG', 'comp 1-a LARG',
                   'comp 1-a peso', 'comp 1-a', 'Blu',
                   'Comp-Variante num sotto'):
            oStileCella = oDoc.createInstance("com.sun.star.style.CellStyle")
            if not oDoc.StyleFamilies.getByName('CellStyles').hasByName(el):
                oDoc.StyleFamilies.getByName('CellStyles').insertByName(
                    el, oStileCella)
                oStileCella.ParentStyle = 'comp 1-a'
        indicator.setValue(2)
        for el in ('comp 1-a PU', 'comp 1-a LUNG', 'comp 1-a LARG',
                   'comp 1-a peso', 'comp 1-a', 'Blu',
                   'Comp-Variante num sotto'):
            oStileCella = oDoc.createInstance("com.sun.star.style.CellStyle")
            if not oDoc.StyleFamilies.getByName("CellStyles").hasByName(el + ' ROSSO'):
                oDoc.StyleFamilies.getByName('CellStyles').insertByName(
                    el + ' ROSSO', oStileCella)
                oStileCella.ParentStyle = 'comp 1-a'
                oDoc.StyleFamilies.getByName("CellStyles").getByName(
                    el + ' ROSSO').CharColor = 16711680
############
# copia gli stili di cella dal template, ma non va perché tocca lavorare sulla FormatString - quando imparerò
#  sUrl = LeenO_path()+'/template/leeno/Computo_LeenO.ods'
#  styles = oDoc.getStyleFamilies()
#  styles.loadStylesFromURL(sUrl, [])
############
        indicator.setValue(3)
        oSheet = oDoc.getSheets().getByName('S1')
        oSheet.getCellRangeByName('S1.H291').Value = \
            oDoc.getDocumentProperties().getUserDefinedProperties().Versione = adegua_a
        for el in oDoc.Sheets.ElementNames:
            oDoc.getSheets().getByName(el).IsVisible = True
            oDoc.CurrentController.setActiveSheet(oDoc.getSheets().getByName(el))
            oDoc.getSheets().getByName(el).IsVisible = False
        # dal template 212
        flags = VALUE + DATETIME + STRING + ANNOTATION + FORMULA + OBJECTS + EDITATTR  # FORMATTED + HARDATTR
        indicator.setValue(4)
        GotoSheet('M1')
        oSheet = oDoc.getSheets().getByName('M1')
        oSheet.getCellRangeByName('B23:E30').clearContents(flags)
        oSheet.getCellRangeByName('B23:E30').CellStyle = 'M1 scritte noP'
        # dal template 208
        # > adegua le formule delle descrizioni di voci
        GotoSheet('S1')
        oSheet = oDoc.getSheets().getByName('S1')
        oSheet.getCellRangeByName(
            'G334'
        ).String = '[Computo e Variante] Vedi Voce: PRIMI caratteri della voce'
        oSheet.getCellRangeByName('H334').Value = 50
        oSheet.getCellRangeByName(
            'I334'
        ).String = "Quanti caratteri della descrizione vuoi visualizzare usando il Vedi Voce?"
        oSheet.getCellRangeByName(
            'G335'
        ).String = '[Contabilità] Descrizioni abbreviate: PRIMI caratteri della voce'
        oSheet.getCellRangeByName('H335').Value = 100
        oSheet.getCellRangeByName(
            'I335'
        ).String = "Quanti caratteri vuoi visualizzare partendo dall'INIZIO della descrizione?"
        oSheet.getCellRangeByName(
            'G336'
        ).String = 'Descrizioni abbreviate: primi caratteri della voce'
        oSheet.getCellRangeByName('H336').Value = 120
        oSheet.getCellRangeByName(
            'I336'
        ).String = "[Contabilità] Descrizioni abbreviate: ULTIMI caratteri della voce"
        oSheet.getCellRangeByName(
            'G337'
        ).String = 'Descrizioni abbreviate: primi caratteri della voce'
        oSheet.getCellRangeByName('H337').Value = 100
        oSheet.getCellRangeByName(
            'I337'
        ).String = "Quanti caratteri vuoi visualizzare partendo dall'INIZIO della descrizione?"
        oSheet.getCellRangeByName(
            'G338'
        ).String = 'Descrizioni abbreviate: ultimi caratteri della voce'
        oSheet.getCellRangeByName('H338').Value = 120
        oSheet.getCellRangeByName(
            'I338'
        ).String = "Quanti caratteri vuoi visualizzare partendo dalla FINE della descrizione?"
        oSheet.getCellRangeByName('L25').String = ''
        oSheet.getCellRangeByName('G297:G338').CellStyle = 'Setvar b'
        oSheet.getCellRangeByName('H297:H338').CellStyle = 'Setvar C'
        oSheet.getCellRangeByName('I297:I338').CellStyle = 'Setvar D'
        oSheet.getCellRangeByName('H319:H326').CellStyle = 'Setvar C_3'
        oSheet.getCellRangeByName('H311').CellStyle = 'Setvar C_3'
        oSheet.getCellRangeByName('H323').CellStyle = 'Setvar C'
        oDoc.StyleFamilies.getByName("CellStyles").getByName(
            'Setvar C_3').NumberFormat = LeenoFormat.getNumFormat('0,00%')  # percentuale
        # < adegua le formule delle descrizioni di voci
        # dal 209 cambia nome di custom propierty
        oUDP = oDoc.getDocumentProperties().getUserDefinedProperties()
        if oUDP.getPropertySetInfo().hasPropertyByName("Versione LeenO"):
            oUDP.removeProperty('Versione LeenO')
        if oUDP.getPropertySetInfo().hasPropertyByName("Versione_LeenO"):
            oUDP.removeProperty('Versione_LeenO')
        oUDP.addProperty('Versione_LeenO',
                         MAYBEVOID + REMOVEABLE + MAYBEDEFAULT,
                         str(LeenoUtils.getGlobalVar('Lmajor')) + '.' +
                         str(LeenoUtils.getGlobalVar('Lminor')) + '.x')
        indicator.setValue(5)
        for el in ('COMPUTO', 'VARIANTE'):
            if oDoc.getSheets().hasByName(el):
                GotoSheet(el)
                oSheet = oDoc.getSheets().getByName(el)
                # sposto il vedivoce nella colonna E
                fine = SheetUtils.getUsedArea(oSheet).EndRow
                oSheet.getCellRangeByPosition(3, 0, 4,
                                              fine).clearContents(HARDATTR)
                for n in range(0, fine):
                    if '=CONCATENATE("' in oSheet.getCellByPosition(
                            2, n).Formula and oSheet.getCellByPosition(
                                4, n).Type.value == 'EMPTY':
                        oSheet.getCellByPosition(
                            4, n).Formula = oSheet.getCellByPosition(5,
                                                                     n).Formula
                        oSheet.getCellByPosition(5, n).String = ''
                        oSheet.getCellByPosition(
                            9, n
                        ).Formula = '=IF(PRODUCT(E' + str(n + 1) + ':I' + str(
                            n + 1) + ')=0;"";PRODUCT(E' + str(
                                n + 1) + ':I' + str(n + 1) + '))'
            # sposto il vedivoce nella colonna E/
                oSheet.getCellByPosition(31, 2).String = 'Super Cat'
                oSheet.getCellByPosition(32, 2).String = 'Cat'
                oSheet.getCellByPosition(33, 2).String = 'Sub Cat'
                oSheet.getCellByPosition(28, 2).String = 'Materiali\ne Noli €'
                lrow = 4
                while lrow < n:
                    oDoc.CurrentController.select(
                        oSheet.getCellByPosition(0, lrow))
                    sistema_stili()
                    lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 1)
                    lrow += 1
                rigenera_tutte() # affido la rigenerazione delle formule al menu Viste
                # 214 aggiorna stili di cella per ogni colonna
                test = SheetUtils.getUsedArea(oSheet).EndRow + 1
                for y in range(0, test):
                    # aggiorna formula vedi voce #214
                    if ver_tmpl > 214:
                        break
                    if(oSheet.getCellByPosition(2, y).Type.value == 'FORMULA' and
                       oSheet.getCellByPosition(2, y).CellStyle == 'comp 1-a' and
                       oSheet.getCellByPosition(5, y).Type.value == 'FORMULA'):
                        try:
                            vRif = int(
                                oSheet.getCellByPosition(
                                    5, y).Formula.split('=J$')[-1]) - 1
                        except Exception:
                            vRif = int(
                                oSheet.getCellByPosition(
                                    5, y).Formula.split('=J')[-1]) - 1
                        if oSheet.getCellByPosition(9, y).Value < 0:
                            _gotoCella(2, y)
                            inverti = 1
                        oSheet.getCellByPosition(5, y).String = ''
                        vedi_voce_xpwe(oSheet, y, vRif)
                        try:
                            inverti
                            inverti_segno()
                        except Exception:
                            pass
                    if '=J' in oSheet.getCellByPosition(5, y).Formula:
                        if '$' in oSheet.getCellByPosition(5, y).Formula:
                            n = oSheet.getCellByPosition(
                                5, y).Formula.split('$')[1]
                        else:
                            n = oSheet.getCellByPosition(
                                5, y).Formula.split('J')[1]
                        oSheet.getCellByPosition(5, y).Formula = '=J$' + n
#  contatta il canale Telegram
#  https://t.me/leeno_computometrico''', 'AVVISO!')
        indicator.setValue(6)
        GotoSheet('S5')
        oSheet = oDoc.getSheets().getByName('S5')
        oSheet.getCellRangeByPosition(
            0, 0, 250,
            SheetUtils.getUsedArea(oSheet).EndRow).clearContents(EDITATTR +
                                                          FORMATTED + HARDATTR)
        oSheet.getCellRangeByName('C10').Formula = \
            '=IF(LEN(VLOOKUP(B10;elenco_prezzi;2;FALSE()))<($S1.$H$337+$S1.$H$338);\
            VLOOKUP(B10;elenco_prezzi;2;FALSE());CONCATENATE(LEFT(VLOOKUP(B10;\
            elenco_prezzi;2;FALSE());$S1.$H$337);" [...] ";RIGHT(VLOOKUP(B10;\
            elenco_prezzi;2;FALSE());$S1.$H$338)))'
        oSheet.getCellRangeByName('C24').Formula = \
            '=IF(LEN(VLOOKUP(B24;elenco_prezzi;2;FALSE()))<($S1.$H$335+$S1.$H$336);VLOOKUP(B24;\
            elenco_prezzi;2;FALSE());CONCATENATE(LEFT(VLOOKUP(B24;elenco_prezzi;\
            2;FALSE());$S1.$H$335);" [...] ";RIGHT(VLOOKUP(B24;elenco_prezzi;2;\
            FALSE());$S1.$H$336)))'
        oSheet.getCellRangeByName('I24').CellStyle = 'Comp-Bianche in mezzo_R'
        oSheet.getCellRangeByName('S12').Formula = '=IF(VLOOKUP(B10;elenco_prezzi;3;FALSE())="%";J12*L12/100;J12*L12)'
        oSheet.getCellRangeByName('P27').Formula = '=IF(VLOOKUP(B24;elenco_prezzi;3;FALSE())="%";J27*N27/100;J27*N27)'
        #
        oSheet.getCellRangeByName('AC12').Formula = '=S12-AE12'
        oSheet.getCellRangeByName('AC12').CellStyle = 'Comp-sotto euri'
        oSheet.getCellRangeByName('AC27').Formula = '=P27-AE27'
        oSheet.getCellRangeByName('AC27').CellStyle = 'Comp-sotto euri'

        oSheet.getCellRangeByName('J11').Formula = '=IF(PRODUCT(E11:I11)=0;"";PRODUCT(E11:I11))'
        oSheet.getCellRangeByName('J25').CellStyle = 'Blu'
        oSheet.getCellRangeByName('J25').Formula = '=IF(PRODUCT(E25:I25)<=0;"";PRODUCT(E25:I25))'

        oSheet.getCellRangeByName('L25').CellStyle = 'Blu ROSSO'
        oSheet.getCellRangeByName('L25').Formula = '=IF(PRODUCT(E25:I25)>=0;"";PRODUCT(E25:I25)*-1)'

        oSheet.getCellRangeByName('J26').Formula = '=IF(SUBTOTAL(9;J24:J26)<0;"";SUBTOTAL(9;J24:J26))'
        oSheet.getCellRangeByName('L26').Formula = '=IF(SUBTOTAL(9;L24:L26)<0;"";SUBTOTAL(9;L24:L26))'
        oSheet.getCellRangeByName('L26').CellStyle = 'Comp-Variante num sotto ROSSO'

        # CONTABILITA CONTABILITA CONTABILITA CONTABILITA CONTABILITA
        indicator.setValue(7)
        if oDoc.getSheets().hasByName('CONTABILITA'):
            GotoSheet('CONTABILITA')
            oSheet = oDoc.getSheets().getByName('CONTABILITA')
            # sposto il vedivoce nella colonna E
            fine = SheetUtils.getUsedArea(oSheet).EndRow + 1
            oSheet.getCellRangeByPosition(3, 0, 4,
                                          fine).clearContents(HARDATTR)
            for n in range(0, fine):
                if '=CONCATENATE("' in oSheet.getCellByPosition(
                        2, n).Formula and oSheet.getCellByPosition(
                            4, n).Type.value == 'EMPTY':
                    oSheet.getCellByPosition(
                        4, n).Formula = oSheet.getCellByPosition(5, n).Formula
                    oSheet.getCellByPosition(5, n).String = ''
                    oSheet.getCellByPosition(
                        9,
                        n).Formula = '=IF(PRODUCT(E' + str(n + 1) + ':I' + str(
                            n + 1) + ')=0;"";PRODUCT(E' + str(
                                n + 1) + ':I' + str(n + 1) + '))'
        # sposto il vedivoce nella colonna E/
            n = LeenoSheetUtils.cercaUltimaVoce(oSheet)
            oSheet.getCellByPosition(
                28, n + 1).Formula = '=SUBTOTAL(9;AC3:AC' + str(n + 2)
            # rigenera_tutte() affido la rigenerazione delle formule al menu Viste
            lrow = 4
            while lrow < n:
                oDoc.CurrentController.select(oSheet.getCellByPosition(0, lrow))
                sistema_stili()
                lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 1)
                lrow += 1
        for el in oDoc.Sheets.ElementNames:
            oDoc.CurrentController.setActiveSheet(
                oDoc.getSheets().getByName(el))
            oSheet = oDoc.getSheets().getByName(el)
            LeenoSheetUtils.adattaAltezzaRiga(oSheet)
#        oDialogo_attesa.endExecute()  # chiude il dialogo
        indicator.setValue(8)
        mostra_fogli_principali()
#    if Dialogs.YesNoDialog(Title='Informazione',
#        Text= '''Vuoi procedere con la rigenerazione di tutte le formule di ogni foglio?
#        Questo richidere del tempo.''') == 1:
        for el in ('COMPUTO',  'VARIANTE',  'CONTABILITA'):
            try:
                oSheet = oDoc.getSheets().getByName(el)
                GotoSheet(el)
                rigenera_tutte()
            except:
                pass
        GotoSheet('COMPUTO')
        indicator.hide()
        Dialogs.Info(Title = 'Avviso', Text='Adeguamento del file completato con successo.')
    LeenoUtils.DocumentRefresh(True)

########################################################################


def XPWE_export_run():
    '''
    Visualizza il menù export/import XPWE
    '''
    oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('S2'):
        return
    psm = LeenoUtils.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    Dialog_XPWE = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.Dialog_XPWE?language=Basic&location=application"
    )
    oSheet = oDoc.CurrentController.ActiveSheet
    # oDialog1Model = Dialog_XPWE.Model
    for el in ("COMPUTO", "VARIANTE", "CONTABILITA", "Elenco Prezzi"):
        try:
            importo = oDoc.getSheets().getByName(el).getCellRangeByName(
                'A2').String
            if el == 'COMPUTO':
                Dialog_XPWE.getControl(el).Label = 'Computo:     ' + importo
            if el == 'VARIANTE':
                Dialog_XPWE.getControl(el).Label = 'Variante:    ' + importo
            if el == 'CONTABILITA':
                Dialog_XPWE.getControl(el).Label = 'Contabilità: ' + importo
            if el == 'ELENCO PREZZI':
                Dialog_XPWE.getControl(el).Label = 'Elenco Prezzi'
            Dialog_XPWE.getControl(el).Enable = True
        except Exception:
            Dialog_XPWE.getControl(el).Enable = False
    Dialog_XPWE.Title = 'Esportazione XPWE'
    try:
        Dialog_XPWE.getControl(oSheet.Name).State = True
    except Exception:
        pass
    Dialog_XPWE.getControl(
        'FileControl1'
    ).Text = 'C:\\tmp\\prova.txt'  # uno.fileUrlToSystemPath(oDoc.getURL())
    # systemPathToFileUrl
    lista = []
    #  Dialog_XPWE.execute()
    # try:
    # Dialog_XPWE.execute()
    # except Exception:
    # pass
    if Dialog_XPWE.execute() == 1:
        for el in ("COMPUTO", "VARIANTE", "CONTABILITA", "Elenco Prezzi"):
            if Dialog_XPWE.getControl(el).State == 1:
                lista.append(el)
    out_file = Dialogs.FileSelect('Salva con nome...', '*.xpwe', 1)
    if out_file == '':
        return
    testo = '\n'
    for el in lista:
        XPWE_out(el, out_file)
        testo += f'● {out_file}-{el}.xpwe\n\n'

    # Dialogs.Exclamation(Title = 'AVVISO!',
    # Text='Il formato XPWE è un formato XML di interscambio per Primus di ACCA.\n\n'
    # 'Prima di utilizzare questo file in Primus, assicurarsi che le percentuali\n'
    # 'di Spese Generali e Utile d\'Impresa siano impostate correttamente, in modo da\n'
    # 'garantire la corretta elaborazione dei dati.')

########################################################################


def chiudi_dialoghi(event=None):
    '''
    @@ DA DOCUMENTARE
    '''

    try:
        oDialog1.endExecute()
    except:
        pass
    try:
        oDlgMain.endExecute()
    except:
        pass
    
    # return
    if event:
        event.Source.Context.endExecute()
    return


def chiudi_dialoghi(event=None):
    '''
    Chiude in modo sicuro eventuali finestre di dialogo aperte.
    '''
    def chiudi(dlg):
        try:
            dlg.endExecute()
        except Exception:
            pass

    # Chiude i dialoghi globali, se presenti
    for nome in ("oDialog1", "oDlgMain", "oDlgSiNo"):
        dlg = globals().get(nome)
        # DLG.mri(f'Chiudo dialogo: {dlg}')
        if dlg is not None:
            chiudi(dlg)

    # Chiude il dialogo associato all'evento, se presente
    if event:
        try:
            ctx = event.Source.Context
            chiudi(ctx)
        except Exception:
            pass
    return




########################################################################
def ScriviNomeDocumentoPrincipale():
    '''
    Indica qual è il Documento Principale
    '''
    return

    # legge il percorso del documento principale
    sUltimus = LeenoUtils.getGlobalVar('sUltimus')

    oDoc = LeenoUtils.getDocument()

    # se si sta lavorando sul Documento Principale, non fa nulla
    try:
        if sUltimus == uno.fileUrlToSystemPath(oDoc.getURL()):
            return
    except Exception:
        # file senza nome
        return

    # mah... non vedo come potrebbe esserco questo errore
    # a meno di non aver eliminato il controller
    try:
        oSheet = oDoc.CurrentController.ActiveSheet
    except AttributeError:
        return

    d = {
        'COMPUTO': 'F1',
        'VARIANTE': 'F1',
        'Elenco Prezzi': 'A1',
        'CONTABILITA': 'F1',
        'Analisi di Prezzo': 'A1'
    }
    for el in d.keys():
        try:
            oSheet = oDoc.Sheets.getByName(el)
            oSheet.getCellRangeByName(d[el]).String = 'DP: ' + LeenoUtils.getGlobalVar('sUltimus')
            oSheet.getCellRangeByName("A1:AT1").clearContents(EDITATTR + FORMATTED + HARDATTR)

        except Exception:
            pass

########################################################################

def ods2pdf(oDoc, sFile):
    '''
    Genera il PDF del file ODS partendo dalle aree di stampa impostate.
    oDoc    { object } : documento da esportare.
    sFile   { string } : nome del file di destinazione.
    '''
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    oProp = []
    oProp0 = PropertyValue()
    oProp0.Name = "URL"
    oProp0.Value = sFile # "file:///C:/TMP/000.pdf"
    oProp1 = PropertyValue()
    oProp1.Name = "FilterName"
    oProp1.Value = "calc_pdf_Export"
    oProp2 = PropertyValue()
    oProp2.Name = "FilterData"
    oProp2.Value = ()
    oProp.append(oProp0)
    oProp.append(oProp1)
    oProp.append(oProp2)
    properties = tuple(oProp)
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, '.uno:ExportToPDF', '', 0, properties)
    return

########################################################################

def DlgPDF():
    oDoc = LeenoUtils.getDocument()
    psm = LeenoUtils.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDlgPDF = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.DlgPDF?language=Basic&location=application"
    )
    sUrl = LeenO_path() + '/python/pythonpath/Icons-Big/preview.png'
    oDlgPDF.getModel().ImageControl1.ImageURL = sUrl

    oSheet = oDoc.getSheets().getByName('S2')

    progetto    = oSheet.getCellRangeByName('$S2.C3').String
    localit     = oSheet.getCellRangeByName('$S2.C4').String
    data_prg    = oSheet.getCellRangeByName('$S2.C5').String
    committente = oSheet.getCellRangeByName('$S2.C6').String
    sponsor     = oSheet.getCellRangeByName('$S2.C7').String
    rup         = oSheet.getCellRangeByName('$S2.C12').String
    progettista = oSheet.getCellRangeByName('$S2.C13').String
    dl          = oSheet.getCellRangeByName('$S2.C16').String
    formulas = {
        'progetto'    : '[PROGETTO]',
        'localit'     : '[LOCALITÀ]',
        'data_prg'    : '[DATA_PROGETTO]',
        'committente' : '[COMMITTENTE-STAZIONE APPALTANTE]',
        'sponsor'     : '[FINANZIATORE]',
        'rup'         : '[RESPONSABILE_PROCEDIMENTO]',
        'progettista' : '[PROGETTISTA]',
        'dl'          : '[DIRETTORE_LAVORI]',
    }
    # DLG.chi(list(formulas.values()))
    # return

    lista = list(formulas.values())

    oDlgPDF.getControl("ComboBox1").addItems(lista, 1)

    oDlgPDF.execute()


def DlgMain():
    '''
    Visualizza il menù principale DlgMain
    '''
    LeenoUtils.DocumentRefresh(True)
    with LeenoUtils.DocumentRefreshContext(False):

        LeenoSheetUtils.memorizza_posizione()

        oDoc = LeenoUtils.getDocument()
        oDoc.unlockControllers()
        psm = LeenoUtils.getComponentContext().ServiceManager
        oSheet = oDoc.CurrentController.ActiveSheet
        if not oDoc.getSheets().hasByName('S2'):
            Toolbars.AllOff()
            if(len(oDoc.getURL()) == 0 and
            SheetUtils.getUsedArea(oSheet).EndColumn == 0 and
            SheetUtils.getUsedArea(oSheet).EndRow == 0):
                oDoc.close(True)
            creaComputo()
        Toolbars.Vedi()
        dp = psm.createInstance("com.sun.star.awt.DialogProvider")
        global oDlgMain
        oDlgMain = dp.createDialog(
            "vnd.sun.star.script:UltimusFree2.DlgMain?language=Basic&location=application"
        )
        LeenoUtils.setGlobalVar('oDlgMain', oDlgMain)

        oDlgMain.Title = 'Menù Principale (Ctrl+0)'

        sUrl = LeenO_path() + '/icons/Immagine.png'
        oDlgMain.getModel().ImageControl1.ImageURL = sUrl

        sString = oDlgMain.getControl("CommandButton13")
        try:
            if LeenoUtils.getGlobalVar('sUltimus') == uno.fileUrlToSystemPath(oDoc.getURL()):
                sString.setEnable(False)
            else:
                sString.setEnable(True)
        except Exception:
            pass

        sString = oDlgMain.getControl("Label12")
        sString.Text = version_code.read()[6:]
        sString = oDlgMain.getControl("Label_DDC")
        sString.Text = LeenoUtils.getGlobalVar('sUltimus')

        sString = oDlgMain.getControl("Label1")
        sString.Text = (
            str(LeenoUtils.getGlobalVar('Lmajor')) + '.' +
            str(LeenoUtils.getGlobalVar('Lminor')) + '.' +
            LeenoUtils.getGlobalVar('Lsubv'))

        sString = oDlgMain.getControl("Label2")
        try:
            oSheet = oDoc.Sheets.getByName('S1')
        except Exception:
            return
        sString.Text = oDoc.getDocumentProperties().getUserDefinedProperties(
        ).Versione  # oSheet.getCellByPosition(7, 290).String
        # sString = oDlgMain.getControl("Label14") # Oggetto del lavoro
        sString = oDlgMain.getControl("TextField1")  # Oggetto del lavoro

        sString.Text = oDoc.Sheets.getByName('S2').getCellRangeByName('C3').String
        try:
            oSheet = oDoc.Sheets.getByName('COMPUTO')
            sString = oDlgMain.getControl("Label8")
            sString.Text = "€ {:,.2f}".format(
                oSheet.getCellByPosition(18, 1).Value)
        except Exception:
            pass
        try:
            oSheet = oDoc.Sheets.getByName('VARIANTE')
            sString = oDlgMain.getControl("Label5")
            sString.Text = "€ {:,.2f}".format(
                oSheet.getCellByPosition(18, 1).Value)
        except Exception:
            pass
        try:
            oSheet = oDoc.Sheets.getByName('CONTABILITA')
            sString = oDlgMain.getControl("Label9")
            sString.Text = "€ {:,.2f}".format(
                oSheet.getCellByPosition(15, 1).Value)
        except Exception:
            pass
        oDlgMain.getControl('CheckBox1').State = int(
            cfg.read('Generale', 'dialogo'))
        LeenoEvents.assegna()
        oDlgMain.execute()
        sString = oDlgMain.getControl("Label_DDC").Text
        if oDlgMain.getControl('CheckBox1').State == 1:
            cfg.write('Generale', 'dialogo', '1')
        else:
            cfg.write('Generale', 'dialogo', '0')
        oDoc.Sheets.getByName('S2').getCellRangeByName(
            'C3').String = oDlgMain.getControl("TextField1").Text

        d = {
            'COMPUTO': 'F1',
            'VARIANTE': 'F1',
            'Elenco Prezzi': 'A1',
            'CONTABILITA': 'F1',
            'Analisi di Prezzo': 'A1'
        }
        for el in d.keys():
            try:
                oSheet = oDoc.Sheets.getByName(el)
                if LeenoUtils.getGlobalVar('sUltimus') == uno.fileUrlToSystemPath(oDoc.getURL()):
                    oSheet.getCellRangeByName(
                        "A1:AT1").CellBackColor = 16773632  # 13434777 giallo
                    oSheet.getCellRangeByName(
                        d[el]).String = 'DP: Questo documento'
                else:
                    oSheet.getCellRangeByName(
                        "A1:AT1").clearContents(HARDATTR)
                    oSheet.getCellRangeByName(
                        d[el]).String = 'DP:' + LeenoUtils.getGlobalVar('sUltimus')

            except Exception:
                pass
        fissa()
        LeenoSheetUtils.ripristina_posizione()
        return


########################################################################
def InputBox(sCella='', t=''):
    '''
    sCella  { string } : stringa di default nella casella di testo
    t       { string } : titolo del dialogo
    Visualizza un dialogo di richiesta testo
    '''

    psm = LeenoUtils.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDialog1 = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.DlgTesto?language=Basic&location=application"
    )
    oDialog1Model = oDialog1.Model

    oDialog1Model.Title = t

    sString = oDialog1.getControl("TextField1")
    sString.Text = sCella

    if oDialog1.execute() == 0:
        return ''
    return sString.Text


########################################################################
def hide_error(lErrori, icol):
    '''
    lErrori  { tuple } : nome dell'errore es.: '#DIV/0!'
    icol { integer } : indice di colonna della riga da nascondere
    Visualizza o nascondi una toolbar
    '''
    oDoc = LeenoUtils.getDocument()
    LeenoUtils.DocumentRefresh(False)
    oSheet = oDoc.CurrentController.ActiveSheet
    #  oSheet.clearOutline()
    n = 3
    test = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    for i in range(n, test):
        for el in lErrori:
            if oSheet.getCellByPosition(icol, i).String == el:
                oCellRangeAddr.StartRow = i
                oCellRangeAddr.EndRow = i
                oSheet.group(oCellRangeAddr, 1)
                oSheet.getCellByPosition(0, i).Rows.IsVisible = False
    LeenoUtils.DocumentRefresh(True)

########################################################################
def bak0():
    '''
    Fa il backup del file di lavoro all'apertura.
    '''
    # tempo = ''.join(''.join(''.join(str(datetime.now()).split('.')[0].split(' ')).split('-')).split(':'))[:12]
    oDoc = LeenoUtils.getDocument()
    orig = oDoc.getURL()
    dest = '.'.join(os.path.basename(orig).split('.')[0:-1]) + '.bak.ods'
    dir_bak = os.path.dirname(oDoc.getURL()) + '/leeno-bk/'
    try:
        dir_bak = uno.fileUrlToSystemPath(dir_bak)
    except:
        pass
    # filename = '.'.join(os.path.basename(orig).split('.')[0:-1]) + '-'
    if len(orig) == 0:
        return
    if not os.path.exists(dir_bak):
        os.makedirs(dir_bak)
    orig = uno.fileUrlToSystemPath(orig)
    dest = uno.fileUrlToSystemPath(dest)
    if os.path.exists(dir_bak + dest):
        shutil.copyfile(dir_bak + dest, dir_bak + dest + '.old')
    shutil.copyfile(orig, dir_bak + dest)


########################################################################
def bak():
    '''
    Esegue un numero definito di backup durante il lavoro.
    '''
    tempo = ''.join(''.join(''.join(
        str(datetime.now()).split('.')[0].split(' ')).split('-')).split(
            ':'))[:12]
    oDoc = LeenoUtils.getDocument()
    orig = oDoc.getURL()
    dest = '.'.join(
        os.path.basename(orig).split('.')[0:-1]) + '-' + tempo + '.ods'
    dir_bak = os.path.dirname(oDoc.getURL()) + '/leeno-bk/'
    filename = '.'.join(os.path.basename(orig).split('.')[0:-1]) + '-'
    if len(orig) == 0:
        return
    if not os.path.exists(uno.fileUrlToSystemPath(dir_bak)):
        os.makedirs(uno.fileUrlToSystemPath(dir_bak))
    orig = uno.fileUrlToSystemPath(orig)
    dest = uno.fileUrlToSystemPath(dest)
    oDoc.storeToURL(dir_bak + dest, [])
    lista = os.listdir(uno.fileUrlToSystemPath(dir_bak))
    n = 0
    nb = int(cfg.read('Generale', 'copie_backup'))  # numero di copie)
    for el in reversed(lista):
        if filename in el:
            if n > nb - 1:
                os.remove(uno.fileUrlToSystemPath(dir_bak) + el)
            n += 1
    return


########################################################################

class version_code:
    """ Gestisce il nome del file OXT in leeno_version_code"""

    def __init__ (self):
        """ Class initialiser """
        pass

    def read ():

        if os.altsep:
            code_file = uno.fileUrlToSystemPath(LeenO_path() + os.altsep +
                                                'leeno_version_code')
        else:
            code_file = uno.fileUrlToSystemPath(LeenO_path() + os.sep +
                                                'leeno_version_code')
        f = open(code_file, 'r')
        return f.readline()

    def write ():

        if os.altsep:
            code_file = uno.fileUrlToSystemPath(LeenO_path() + os.altsep +
                                                'leeno_version_code')
        else:
            code_file = uno.fileUrlToSystemPath(LeenO_path() + os.sep +
                                                'leeno_version_code')
        f = open(code_file, 'r')
        Ldev = str (int(f.readline().split('LeenO-')[1].split('-')[0].split('.')[-1]) + 1)
        tempo = ''.join(''.join(''.join(str(datetime.now()).split('.')[0].split(' ')).split('-')).split(':'))
        of = open(code_file, 'w')

        new = (
            'LeenO-' +
            str(LeenoUtils.getGlobalVar('Lmajor')) + '.' +
            str(LeenoUtils.getGlobalVar('Lminor')) + '.' +
            LeenoUtils.getGlobalVar('Lsubv').split('.')[0] + '.' +
            Ldev + '-TESTING-' +
            tempo[:-6])
        of.write(new)
        of.close()
        return new

def description_upd():
    '''
    Aggiorna il valore di versione del file description.xml
    '''
    if os.altsep:
        desc_file = uno.fileUrlToSystemPath(LeenO_path() + os.altsep +
                                            'description.xml')
    else:
        desc_file = uno.fileUrlToSystemPath(LeenO_path() + os.sep +
                                            'description.xml')
    f = open(desc_file, 'r')
    oxt_name = version_code.read()

    new = []
    for el in f.readlines():
        if '<version value=' in el:
            el.split('''"''')
            el = el.split('''"''')[0] +'''"'''+ oxt_name[6:100] +'''"'''+ el.split('''"''')[2]
        new.append(el)

    str_join = ''.join(new)

    of = open(desc_file, 'w')
    of.write(str_join)
    of.close()
    return



########################################################################
def MENU_grid_switch():
    '''Mostra / nasconde griglia'''
    oDoc = LeenoUtils.getDocument()
    oDoc.CurrentController.ShowGrid = not oDoc.CurrentController.ShowGrid


def MENU_make_pack():
    '''
    Produce il pacchetto installabile
    '''
    make_pack()


def make_pack(bar=0):
    '''
    bar { integer } : toolbar 0=spenta 1=accesa
    Pacchettizza l'estensione in duplice copia: LeenO.oxt e LeenO-x.xx.x.xxx-TESTING-yyyymmdd.oxt
    in una directory precisa (da parametrizzare...)
    '''
    LeenoUtils.DocumentRefresh(False)
    oDoc = LeenoUtils.getDocument()
    try:
        if oDoc.getSheets().getByName('S1').getCellByPosition(
                7, 338).String == '':
            src_oxt = '_LeenO'
        else:
            src_oxt = oDoc.getSheets().getByName('S1').getCellByPosition(
                7, 338).String
    except Exception:
        pass
    oxt_name = version_code.write()
    description_upd()  # aggiorna description.xml - da disattivare prima del rilascio
    if bar == 0:
        oDoc = LeenoUtils.getDocument()
        Toolbars.AllOff()
    oxt_path = uno.fileUrlToSystemPath(LeenO_path())

    if os.name == 'nt':
        if not os.path.exists('w:/_dwg/ULTIMUSFREE/_SRC/OXT/'):
            try:
                os.makedirs(os.getenv("HOMEPATH") + '/' + src_oxt + '/')
            except FileExistsError:
                pass
            nomeZip2 = os.getenv(
                "HOMEPATH") + '/' + src_oxt + '/OXT/' + oxt_name + '.oxt'
            subprocess.Popen('explorer.exe ' + os.getenv("HOMEPATH") + '\\' +
                             src_oxt + '\\OXT\\',
                             shell=True,
                             stdout=subprocess.PIPE)
        else:
            nomeZip2 = 'w:/_dwg/ULTIMUSFREE/_SRC/OXT/' + oxt_name + '.oxt'
            subprocess.Popen('explorer.exe w:\\_dwg\\ULTIMUSFREE\\_SRC\\OXT\\',
                             shell=True,
                             stdout=subprocess.PIPE)

    # if sys.platform == 'linux' or sys.platform == 'darwin':
    else:
        dest = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt'
        if not os.path.exists(dest):
            try:
                dest = os.getenv(
                    "HOME") + '/' + src_oxt + '/leeno/src/Ultimus.oxt/'
                os.makedirs(dest)
                os.makedirs(os.getenv("HOME") + '/' + src_oxt + '/_SRC/OXT')
            except FileExistsError:
                pass
            nomeZip2 = os.getenv(
                "HOME") + '/' + src_oxt + '/_SRC/OXT/' + oxt_name + '.oxt'
            subprocess.Popen('caja ' + os.getenv("HOME") + '/' + src_oxt +
                             '/_SRC/OXT',
                             shell=True,
                             stdout=subprocess.PIPE)

        else:
            nomeZip2 = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/OXT/' + oxt_name + '.oxt'
            subprocess.Popen(
                'caja /media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/OXT/',
                shell=True,
                stdout=subprocess.PIPE)

    # Creazione manuale dello ZIP escludendo `.mypy_cache` e `_pycache_`
    with zipfile.ZipFile(nomeZip2 + '.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(oxt_path):
            # Escludi le directory `.mypy_cache` e `_pycache_`
            for exclude in ['.mypy_cache', '__pycache__', '.venv']:
                if exclude in dirs:
                    dirs.remove(exclude)
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, oxt_path)
                zipf.write(file_path, arcname)

    # Rinomina l'archivio in `.oxt`
    shutil.move(nomeZip2 + '.zip', nomeZip2)
    LeenoUtils.DocumentRefresh(True)


#######################################################################
def dlg_donazioni():
    '''
    @@ DA DOCUMENTARE
    '''
    psm = LeenoUtils.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDialog1 = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.DlgDonazioni?language=Basic&location=application"
    )
    # oDialog1Model = oDialog1.Model
    sUrl = LeenO_path() + '/icons/pizza.png'
    oDialog1.getModel().ImageControl1.ImageURL = sUrl
    if oDialog1.execute() == 0:
        return


########################################################################
def donazioni():
    '''
    @@ DA DOCUMENTARE
    '''
    apri = LeenoUtils.createUnoService("com.sun.star.system.SystemShellExecute")
    apri.execute("https://leeno.org/donazioni/", "", 0)

########################################################################
class XPWE_export_th(threading.Thread):
    '''
    @@ DA DOCUMENTARE
    '''
    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
        XPWE_export_run()


def MENU_XPWE_export():
    '''
    @@ DA DOCUMENTARE
    '''
    XPWE_export_th().start()


########################################################################
class inserisci_nuova_riga_con_descrizione_th(threading.Thread):
    '''
    @@ DA DOCUMENTARE
    '''
    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
#        oDialogo_attesa = DLG.dlg_attesa()
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet

        # attiva la progressbar
        progress = Dialogs.Progress(Title='Inserimrnto in corso...', Text="Lettura dati")
        
        progress.setLimits(0, SheetUtils.getUsedArea(oSheet).EndRow)
        progress.setValue(0)

        if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
            return
        descrizione = InputBox(t='Inserisci una descrizione per la nuova riga')
        progress.show()
        i = 0
        while (i < SheetUtils.getUsedArea(oSheet).EndRow):
            progress.setValue(i)
            if oSheet.getCellByPosition(2, i).CellStyle == 'comp 1-a':
                sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, i)
                qui = sStRange.RangeAddress.StartRow + 1

                i = sotto = sStRange.RangeAddress.EndRow + 3
                oDoc.CurrentController.select(oSheet.getCellByPosition(2, qui))
                Copia_riga_Ent()
                oSheet.getCellByPosition(2, qui + 1).String = descrizione
                LeenoSheetUtils.prossimaVoce(oSheet, sotto)
            i += 1
        progress.hide()


def MENU_inserisci_nuova_riga_con_descrizione():
    '''
    inserisce, all'inizio di ogni voce di computo o variante,
    una nuova riga con una descrizione a scelta
    '''
    inserisci_nuova_riga_con_descrizione_th().start()


def MENU_elenco_puntato_misure():
    '''
    Aggiunge il trattino (-) alle righe di misura
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
    lrow = SheetUtils.getLastUsedRow(oSheet)
    for el in range(3, lrow):
        cell = oSheet.getCellByPosition(2, el)
        cell_string = cell.String
        cell_style = cell.CellStyle

        is_not_dash_space = cell_string[0:2] != '- '
        is_comp_1_a = cell_style in ('comp 1-a', 'comp 1-a ROSSO')
        is_not_empty = cell_string != ''
        is_not_arrow = cell_string[0:2] != ' ►'
        is_not_ = cell_string[0:2] != ' ?'
        is_not_VOCE_AZZERATA = cell_string != '*** VOCE AZZERATA ***'
        is_not_partita = "PARTITA PROVVISORIA" not in cell_string

        if is_not_ and is_not_partita and is_not_dash_space and is_comp_1_a and is_not_empty and is_not_arrow and is_not_VOCE_AZZERATA:
            cell.String = '- ' + cell.String
    return


########################################################################
def ctrl_d():
    '''
    Copia il valore della prima cella superiore utile.
    '''
    oDoc = LeenoUtils.getDocument()
    oCell = oDoc.CurrentSelection
    oSheet = oDoc.CurrentController.ActiveSheet
    x = LeggiPosizioneCorrente()[0]
    lrow = LeggiPosizioneCorrente()[1]
    y = lrow - 1
    try:
        while oSheet.getCellByPosition(x, y).Type.value == 'EMPTY':
            y -= 1
    except Exception:
        return
    oDoc.CurrentController.select(oSheet.getCellByPosition(x, y))
    comando('Copy')
    oDoc.CurrentController.select(oCell)
    paste_clip(insCell=0)
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect


########################################################################
def MENU_taglia_x():
    '''
    taglia il contenuto della selezione
    senza cancellare la formattazione delle celle
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    comando('Copy')

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
    # oRange = oSheet.getCellRangeByPosition(sCol, sRow, eCol, eRow)
    flags = VALUE + DATETIME + STRING + ANNOTATION + FORMULA + OBJECTS + EDITATTR  # FORMATTED + HARDATTR
    oSheet.getCellRangeByPosition(sCol, sRow, eCol, eRow).clearContents(flags)


########################################################################

def calendario_mensile():
    '''
    Colora le colonne del sabato e della domenica, oltre alle festività,
    nel file ../PRIVATO/LeenO/extra/calendario.ods che potrei implementare
    in LeenO per la gestione delle ore in economia o del diagramma di Gantt.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName('elenco festività')
    oRangeAddress = oDoc.NamedRanges.feste.ReferredCells.RangeAddress
    SR = oRangeAddress.StartRow
    ER = oRangeAddress.EndRow
    lFeste = []
    for x in range(SR, ER):
        if oSheet.getCellByPosition(0, x).Value != 0:
            lFeste.append(oSheet.getCellByPosition(0, x).String)
    oSheet = oDoc.getSheets().getByName('CALENDARIO')
    test = SheetUtils.getUsedArea(oSheet).EndColumn + 1
    slist = []
    for x in range(0, test):
        if oSheet.getCellByPosition(
                x, 3).String == 's' or oSheet.getCellByPosition(
                    x, 3).String == 'd':
            slist.append(x)
    for x in range(0, test):
        if oSheet.getCellByPosition(x, 1).String in lFeste:
            slist.append(x)

    for x in range(2, SheetUtils.getUsedArea(oSheet).EndColumn + 1):
        for y in range(1, SheetUtils.getUsedArea(oSheet).EndRow + 1):
            if x in slist:
                oSheet.getCellByPosition(x, y).CellStyle = 'ok'
            else:
                oSheet.getCellByPosition(x, y).CellStyle = 'tabella'

########################################################################

def calendario_liste():
    LeenoUtils.DocumentRefresh(False)
    def rgb(r, g, b):
        return 256*256*r + 256*g + b
    '''
    Colora le colonne del sabato e della domenica
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if not "LISTA" in oSheet.Name.upper():
        return

    # slist = []
    # for x in range(3, 35):
        # if oSheet.getCellByPosition(
                # x, 3).String == 's' or oSheet.getCellByPosition(
                    # x, 3).String == 'd':
            # slist.append(x)
    slist = [x for x in range(3, 35) if oSheet.getCellByPosition(x, 3).String.lower() in ['s', 'd']]
    ER = SheetUtils.uFindStringCol('C', 0, oSheet)
    n = 1
    for y in range(3, ER):
        if oSheet.getCellByPosition(0, y).Type.value == 'VALUE' or oSheet.getCellByPosition(0, y).String == '':
            oSheet.getCellByPosition(0, y).Value = n
            oSheet.getCellByPosition(35, y).Formula = '=SUBTOTAL(9;E' + str(y + 1) + ':AI' + str(y + 1) + ')'  # somma ore
            oSheet.getCellByPosition(36, y).Formula = '=IFERROR(VLOOKUP(B' + str(y + 1) + ';elenco_prezzi;5;FALSE());"")'
            oSheet.getCellByPosition(37, y).Formula = '=IFERROR(AJ' + str(y + 1) + '*AK' + str(y + 1) + ';"")'
            n += 1
        if oSheet.getCellByPosition(4, y).IsMerged == True:
            ER = y -1
            break
    oSheet.getCellRangeByPosition(4, 3, 34, ER).CellBackColor = -1
    for y in slist:
        # oSheet.getCellRangeByPosition(y, 3, y, ER).CellBackColor = rgb (238,238,238)
        oSheet.getCellRangeByPosition(y, 3, y, ER).CellStyle = 'ok'
    LeenoEvents.assegna()
    LeenoUtils.DocumentRefresh(True)
    LeenoUtils.DocumentRefresh(True)


########################################################################
def dal_al():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    for y in range(3, 100):
        if oSheet.getCellByPosition(4, y).IsMerged == True:
            ER = y -1
            break
    found = True
    for y in range(4, ER + 1):
        for x in (4, 34):
            if found:
                break  
            if oSheet.getCellByPosition(x, y).Type.value != 'EMPTY':
                DLG.chi(oSheet.getCellByPosition(x, y).String)
                found = True  
                break



########################################################################
def clean_text(desc):
    # Rimuove caratteri non stampabili
    desc = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', desc)
    sostituzioni = {
        # "\n":" ",
        "&Agrave;": "À",
        "&#192;": "À",
        "&Egrave;": "È",
        "&#200;": "È",
        "&Igrave;": "Ì",
        "&#204;": "Ì",
        "&Ograve;": "Ò",
        "&#210;": "Ò",
        "&Ugrave;": "Ù",
        "&#217;": "Ù",
        "&agrave;": "à",
        "&#224;": "à",
        "&egrave;": "è",
        "&#232;": "è",
        "&igrave;": "ì",
        "&#236;": "ì",
        "&ograve;": "ò",
        "&#242;": "ò",
        "&ugrave;": "ù",
        "&#249;": "ù",
        '\t': ' ',
        'Ã¨': 'è',
        'Â°': '°',
        'Ã': 'à',
        ' $': '',
        'Ó': 'à',
        'Þ': 'é',
        '&#x13;': '',
        '&#xD;&#xA;': '',
        '&#xA;': '',
        '&apos;': "'",
        '&#x3;&#x1;': '',
        '- -': '- ',
        '—': '-',
        '–': '-',
        '\n- -': '\n-',
        '\n \n': '\n',
        '\n ': '\n',
        '': '\n',
    }

    # Esegue tutte le sostituzioni
    for old, new in sostituzioni.items():
        desc = desc.replace(old, new)

    # Rimuove spazi multipli con una singola regex
    desc = re.sub(r' +', ' ', desc)

    # Rimuove righe vuote multiple
    desc = re.sub(r'\n+', '\n', desc)
    desc = re.sub(r'^\s*-+\s*$', '', desc, flags=re.MULTILINE)

    # Rimuove spazi all'inizio e alla fine
    desc = desc.strip()

    return desc

def clean_text_file(filename):
    '''
    Pulisce il testo di un file da caratteri non stampabili e cattive codifiche.
    '''
    # Legge il file
    with open(filename, 'r', encoding='utf-8') as file:
        contenuto = file.read()

    # Pulisce il contenuto
    contenuto_pulito = clean_text(contenuto)

    # Scrive il file pulito (opzionale)
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(contenuto_pulito)
    return contenuto_pulito

def sistema_cose():
    '''
    Ripulisce il testo da capoversi, spazi multipli e cattive codifiche.
    '''
    LeenoUtils.DocumentRefresh(False)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lcol = LeggiPosizioneCorrente()[0]

    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()

    el_y = []
    lista_y = []

    try:
        len(oRangeAddress)
        for el in oRangeAddress:
            el_y.append((el.StartRow, el.EndRow))
    except TypeError:
        el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))

    for y in el_y:
        for el in reversed(range(y[0], y[1] + 1)):
            lista_y.append(el)

    for y in lista_y:
        cell = oSheet.getCellByPosition(lcol, y)
        if cell.Type.value == 'TEXT':
            cell.String = clean_text(cell.String)

    Menu_adattaAltezzaRiga()
    LeenoUtils.DocumentRefresh(True)


########################################################################


def descrizione_in_una_colonna(flag=False):
    '''
    Questa funzione consente di estendere su più colonne o ridurre ad una colonna lo spazio
    occupato dalla descrizione di voce in COMPUTO, VARIANTE e CONTABILITA.

    Args:
        flag (bool, optional): Se True, effettua l'unione delle celle. Se False, annulla l'unione.

    '''
    oDoc = LeenoUtils.getDocument()
    LeenoUtils.DocumentRefresh(False)

    fogli_lavoro = ['S5', 'COMPUTO', 'VARIANTE', 'CONTABILITA']

    if oDoc.NamedRanges.hasByName("_Lib_1"):
        Dialogs.Exclamation(Title='ATTENZIONE!', Text="Risulta già registrato un SAL.\n\nIl foglio CONTABILITA sarà ignorato.")
        fogli_lavoro.remove('CONTABILITA')

    for nome_foglio in fogli_lavoro:
        if oDoc.getSheets().hasByName(nome_foglio):
            oSheet = oDoc.getSheets().getByName(nome_foglio)

            # Definisci le righe iniziali in base al foglio
            inizio_righe = [3]
            fine_riga = SheetUtils.getUsedArea(oSheet).EndRow
            if nome_foglio == 'S5':
                inizio_righe = [3, 9, 23]
                if oDoc.NamedRanges.hasByName("_Lib_1"):
                    inizio_righe.remove(23)
                    fine_riga = 15
            for inizio_riga in inizio_righe:
                for y in range(inizio_riga, fine_riga + 1):
                    cell_style = oSheet.getCellByPosition(2, y).CellStyle
                    cell_range = oSheet.getCellRangeByPosition(2, y, 8, y)

                    if cell_style in ('Comp-Bianche sopraS', 'Comp-Bianche in mezzo Descr', 'Comp-Bianche sopra_R', 'Comp-Bianche in mezzo Descr_R'):
                        cell_range.merge(flag)

        # Menu_adattaAltezzaRiga()
    LeenoUtils.DocumentRefresh(True)


########################################################################
def MENU_numera_colonna():
    '''
    Comando di menu per numera_colonna()
    '''
    with LeenoUtils.DocumentRefreshContext(False):
        numera_colonna()

def numera_colonna():
    '''Inserisce l'indice di colonna nelle prime 100 colonne del rigo selezionato
Associato a Ctrl+Shift+C'''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]
    for x in range(0, 50):
        if oSheet.getCellByPosition(x, lrow).Type.value == 'EMPTY':
            larg = oSheet.getCellByPosition(x, lrow).Columns.Width
            # oSheet.getCellByPosition(x, lrow).Value = larg
            oSheet.getCellByPosition(x, lrow).Formula = '=" " & CELL("col")-1'
            oSheet.getCellByPosition(x, lrow).HoriJustify = 'CENTER'
        elif oSheet.getCellByPosition(x, lrow).Formula == '=" " & CELL("col")-1':
            oSheet.getCellByPosition(x, lrow).String = ''
            oSheet.getCellByPosition(x, lrow).HoriJustify = 'STANDARD'


########################################################################
def subst_str(cerca, sostituisce):
    '''
    Sostituisce stringhe di testi nel foglio corrente
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ReplaceDescriptor = oSheet.createReplaceDescriptor()
    ReplaceDescriptor.SearchString = cerca
    ReplaceDescriptor.ReplaceString = sostituisce
    oSheet.replaceAll(ReplaceDescriptor)


########################################################################
def processo(arg):
    '''Verifica l'esistenza di un processo di sistema'''
    ps = subprocess.Popen("ps -A", shell=True, stdout=subprocess.PIPE)
    #  arg = 'soffice'
    if arg in (str(ps.stdout.read())):
        return True
    return False


########################################################################
def GetRegistryKeyContent(sKeyName, bForUpdate):
    '''Dà accesso alla configurazione utente di LibreOffice'''
    oConfigProvider = LeenoUtils.createUnoService(
        "com.sun.star.configuration.ConfigurationProvider")
    arg = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    arg.Name = "nodepath"
    arg.Value = sKeyName
    if bForUpdate:
        GetRegistryKeyContent = oConfigProvider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationUpdateAccess", (arg, ))
    else:
        GetRegistryKeyContent = oConfigProvider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationAccess", (arg, ))
    return GetRegistryKeyContent


########################################################################


def DelPrintArea ():
    '''
    Cancella area di stampa di tutti i fogli ad esclusione di quello
    corrente e del foglio cP_Cop
    '''
    LeenoUtils.DocumentRefresh(True)
    oDoc = LeenoUtils.getDocument()
    nome = oDoc.CurrentController.ActiveSheet.Name
    lista_fogli = oDoc.Sheets.ElementNames
    for el in lista_fogli:
        if el not in (nome, 'cP_Cop'):
            oSheet = oDoc.getSheets().getByName(el)
            oSheet.setPrintAreas(())
    return


########################################################################


def set_area_stampa():
    ''' Imposta area di stampa in relazione all'elaborato da produrre'''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.removeAllManualPageBreaks()
    ER = SheetUtils.getLastUsedRow(oSheet)

    iSheet = oSheet.RangeAddress.Sheet
    oTitles = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oTitles.Sheet = iSheet

    if oSheet.Name in ("VARIANTE", "COMPUTO", "COMPUTO_print", 'Elenco Prezzi', 'CONTABILITA'):

        oSheet.getCellByPosition(0, 2).Rows.Height = 800
        SR = 2
        EC = 41
        # riga da ripetere
        oTitles.StartRow = 2
        oTitles.EndRow = 2
        oSheet.setTitleRows(oTitles)
        oSheet.setPrintTitleRows(True)
        if oSheet.Name == 'Elenco Prezzi':
            EC = 6
            ER -= 1
        if oSheet.Name == 'CONTABILITA':
            EC = 15
            if oDoc.NamedRanges.hasByName('_Lib_1'):
                return
    elif oSheet.Name in ('Analisi di Prezzo'):
        EC = 6
        SR = 1
        ER -= 1
        oSheet.setPrintTitleRows(False)
    elif oSheet.Name in ('cP_Cop'):
        EC = 7
        SR = 0
        ER -= 1
        oSheet.setPrintTitleRows(False)
    elif oSheet.Name in ('SAL'):
        SR = 2
        EC = 5
    else:
        SR = 0
        ER = SheetUtils.getLastUsedRow(oSheet)
        EC = SheetUtils.getLastUsedColumn(oSheet)
# imposta area di stampa
    oStampa = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oStampa.Sheet = iSheet
    oStampa.StartColumn = 0
    oStampa.StartRow = SR
    oStampa.EndColumn = EC
    oStampa.EndRow = ER
    oSheet.setPrintAreas((oStampa,))

    LS.setPageStyle()

########################################################################


def MENU_sistema_pagine(msg = True):
    '''
    msg { boolean } : se True mostra dialogo e lancia anteprima

    Configura intestazioni e pie' di pagina degli stili di stampa
    e propone un'anteprima di stampa
    '''
    oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('M1'):
        return

    # se il preview è già attivo, ferma tutto
    try:
        set_area_stampa()
    except:
        return
    if msg:
        if Dialogs.YesNoDialog(IconType="question",Title='AVVISO!',
            Text='''Vuoi attribuire il colore bianco allo sfondo delle celle?
Le formattazioni dirette impostate durante il lavoro andranno perse.

Per ripristinare i colori, tipici dei fogli di LeenO, basterà selezionare
le celle ed usare "CTRL+M".

Prima di procedere, vuoi il fondo bianco in tutte le celle?''') == 1:
            LeenoSheetUtils.SbiancaCellePrintArea()

    oSheet = oDoc.CurrentController.ActiveSheet
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    SheetUtils.visualizza_PageBreak()
    oSheet.removeAllManualPageBreaks()

    #  committente = oDoc.NamedRanges.Super_ego_8.ReferredCells.String
    oggetto = oDoc.getSheets().getByName('S2').getCellRangeByName("C3").String + '\n\n'
    committente = "Committente: " + oDoc.getSheets().getByName('S2').getCellRangeByName("C6").String
    luogo = '\n' + oSheet.Name
    if oSheet.Name == 'COMPUTO':
        luogo = '\nComputo Metrico Estimativo'
    elif oSheet.Name == 'VARIANTE':
        luogo = '\nPerizia di Variante'

    if oSheet.Name == 'COMPUTO' and oSheet.getColumns().getByName("AD").Columns.IsVisible == True:

        luogo = luogo + ' - Incidenza MdO'
    # luogo = '\nLocalità: ' + oDoc.getSheets().getByName('S2').getCellRangeByName("C4").String

    ###
    #  oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByName('PageStyle_COMPUTO_A4')
    #  DLG.mri(oAktPage)
    #  return
    ###
    if cfg.read('Generale', 'dettaglio') == '1':
        dettaglio_misure(0)
        dettaglio_misure(1)
    else:
        dettaglio_misure(0)
    for n in range(0, oDoc.StyleFamilies.getByName('PageStyles').Count):
        oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByIndex(n)

        # chi((n , oAktPage.DisplayName))
        if oAktPage.DisplayName == 'Page_Style_COPERTINE':
            oAktPage.HeaderIsOn = False
            oAktPage.FooterIsOn = False
        else:
            oAktPage.HeaderIsOn = True
            oAktPage.FooterIsOn = True

        # margini della pagina
        oAktPage.TopMargin = 1000
        oAktPage.BottomMargin = 1350
        oAktPage.LeftMargin = 1000
        oAktPage.RightMargin = 1000

        oAktPage.FooterLeftMargin = 0
        oAktPage.FooterRightMargin = 0
        oAktPage.HeaderLeftMargin = 0
        oAktPage.HeaderRightMargin = 0

        oAktPage.HeaderBodyDistance = 0
        oAktPage.FooterBodyDistance = 0

        # Adatto lo zoom alla larghezza pagina
        oAktPage.PageScale = 0
        oAktPage.CenterHorizontally = True
        oAktPage.ScaleToPagesX = 1
        oAktPage.ScaleToPagesY = 0
        


        if oAktPage.DisplayName in ('PageStyle_Analisi di Prezzo',
                                    'PageStyle_COMPUTO_A4',
                                    'PageStyle_Elenco Prezzi'):
            htxt = 8.0 / 100 * oAktPage.PageScale
            if oSheet.Name == 'Analisi di Prezzo':
                htxt = 9.0 / 100 * oAktPage.PageScale
            # if oAktPage.DisplayName in ('PageStyle_Analisi di Prezzo'):
            #     htxt = 10.0
            #azzera i bordi
            bordo = oAktPage.TopBorder
            bordo.LineWidth = 0
            bordo.OuterLineWidth = 0
            oAktPage.TopBorder = bordo

            bordo = oAktPage.BottomBorder
            bordo.LineWidth = 0
            bordo.OuterLineWidth = 0
            oAktPage.BottomBorder = bordo

            bordo = oAktPage.RightBorder
            bordo.LineWidth = 0
            bordo.OuterLineWidth = 0
            oAktPage.RightBorder = bordo

            bordo = oAktPage.LeftBorder
            bordo.LineWidth = 0
            bordo.OuterLineWidth = 0
            oAktPage.LeftBorder = bordo

            # HEADER
            oHeader = oAktPage.RightPageHeaderContent
            # oAktPage.PageScale = 95
            oHLText = oHeader.LeftText.Text.String = committente
            oHRText = oHeader.LeftText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            oHRText = oHeader.LeftText.Text.Text.CharHeight = htxt

            oHLText = oHeader.CenterText.Text.String = oggetto
            oHRText = oHeader.CenterText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            oHRText = oHeader.CenterText.Text.Text.CharHeight = htxt

            oHRText = oHeader.RightText.Text.String = luogo
            oHRText = oHeader.RightText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            oHRText = oHeader.RightText.Text.Text.CharHeight = htxt

            oAktPage.RightPageHeaderContent = oHeader
            # FOOTER
            oFooter = oAktPage.RightPageFooterContent
            oHLText = oFooter.CenterText.Text.String = ''
            nomefile = oDoc.getURL().replace('%20',' ')
            oHLText = oFooter.LeftText.Text.String = "\nrealizzato con LeenO: " + os.path.basename(nomefile)
            oHLText = oFooter.LeftText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            oHLText = oFooter.LeftText.Text.Text.CharHeight = htxt * 0.5
            oHLText = oFooter.RightText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            oHLText = oFooter.RightText.Text.Text.CharHeight = htxt
            oAktPage.RightPageFooterContent = oFooter

        if oAktPage.DisplayName == 'Page_Style_Libretto_Misure2':
            scelta_viste()

    try:
        if oDoc.CurrentController.ActiveSheet.Name in ('COMPUTO', 'VARIANTE',
                                                       'CONTABILITA',
                                                       'Elenco Prezzi'):

            _gotoCella(0, 3)
        if oDoc.CurrentController.ActiveSheet.Name in ('Analisi di Prezzo'):
            LeenoAnalysis.MENU_impagina_analisi()
            _gotoCella(0, 2)
        if msg:
            setPreview()
    except Exception:
        pass
        # bordo lato destro in attesa di LibreOffice 6.2
        # bordo = oAktPage.RightBorder
        # bordo.LineWidth = 0
        # bordo.OuterLineWidth = 0
        # oAktPage.RightBorder = bordo
    # last = SheetUtils.getUsedArea(oSheet).EndRow

    # oSheet.getCellRangeByPosition(1, 0, 41, last).Rows.OptimalHeight = True
    LeenoUtils.DocumentRefresh(True)
    return


########################################################################
def fissa():
    '''
    Fissa prima riga e colonna nel foglio attivo.
    '''
    oDoc = LeenoUtils.getDocument()
    # vRow = oDoc.CurrentController.getFirstVisibleRow()
    # LeenoSheetUtils.memorizza_posizione()

    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA', 'Elenco Prezzi'):
        _gotoCella(0, 3)
        oDoc.CurrentController.freezeAtPosition(0, 3)
    elif oSheet.Name in ('Analisi di Prezzo'):
        _gotoCella(0, 2)
        oDoc.CurrentController.freezeAtPosition(0, 2)
    elif oSheet.Name in ('Registro', 'SAL'):
        _gotoCella(0, 1)
        oDoc.CurrentController.freezeAtPosition(0, 1)
    # oDoc.CurrentController.setFirstVisibleRow(vRow)
    # LeenoSheetUtils.ripristina_posizione()

    # _gotoCella(lcol, lrow)


########################################################################


def trova_ricorrenze():
    '''
    Consente la visualizzazione selettiva delle voci di COMPUTO che fanno
    capo alla stessa voce di Elenco Prezzi.
    '''
    chiudi_dialoghi()

    def ricorrenze():
        '''Trova i codici di prezzo ricorrenti nel COMPUTO'''
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        struttura_off()
        last = SheetUtils.getUsedArea(oSheet).EndRow
        lista = []
        for n in range(3, last):
            if oSheet.getCellByPosition(1, n).CellStyle == 'comp Art-EP_R':
                lista.append(oSheet.getCellByPosition(1, n).String)
        unici = (set(lista))
        for el in unici:
            lista.remove(el)
        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct(
            'com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        # lrow = 0
        for n in range(0, last):
            if oSheet.getCellByPosition(1, n).CellStyle == 'comp Art-EP_R':
                if oSheet.getCellByPosition(1, n).String not in lista:
                    oRange = LeenoComputo.circoscriveVoceComputo(oSheet, n).RangeAddress
                    oCellRangeAddr.StartRow = oRange.StartRow
                    oCellRangeAddr.EndRow = oRange.EndRow
                    oSheet.group(oCellRangeAddr, 1)
        lista = list(set(lista))
        lista.sort()
        return lista

    #  try:
    #  lista_ricorrenze
    #  except Exception:
    lista_ricorrenze = ricorrenze()
    LeenoUtils.setGlobalVar('lista_ricorrenze', lista_ricorrenze)
    if len(lista_ricorrenze) == 0:
        # DLG.MsgBox('Non ci sono voci di prezzo ricorrenti.', 'Informazione')
        Dialogs.Info(Title = 'Informazione',
        Text="Non ci sono voci di prezzo ricorrenti.")
        return
    psm = LeenoUtils.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDlg = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.DlgLista?language=Basic&location=application"
    )
    # oDialog1Model = oDlg.Model
    #  oDlg.Title = 'Si ripetono '+ str(len(lista_ricorrenze)) + ' voci di prezzo'
    oDlg.Title = 'Seleziona il codice e dai OK...'
    oDlg.getControl('ListBox1').addItems(lista_ricorrenze, 0)
    if oDlg.execute() == 0:
        return
    filtra_codice(oDlg.getControl('ListBox1').SelectedItem)
    #  if oDlg.getControl('CheckBox1').State == 1:
    #  oDlg.getControl('ListBox1').removeItems(0, len(lista_ricorrenze))
    #  lista_ricorrenze = ricorrenze()
    #  oDlg.getControl('ListBox1').addItems(lista_ricorrenze, 0)
    #  oDlg.getControl('CheckBox1').State = 0
    #  oDlg.execute()

########################################################################


def trova_np():
    '''
    Raggruppa le righe in modo da rendere evidenti i nuovi prezzi
    e aggiunge il prefisso NPxx_ al loro codice (se confermato dall'utente),
    propagando la modifica in tutti i fogli.
    Se il codice ha già il prefisso VDS_, mantienilo e aggiungi NPxx_ dopo.
    Rileva la presenza di prefissi NPxx_ esistenti e informa l'utente.
    Permette di scegliere se confrontare COMPUTO con VARIANTE o con CONTABILITÀ.
    '''
    with LeenoUtils.DocumentRefreshContext(False):
        chiudi_dialoghi()
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet

        struttura_off()
        genera_sommario()

        # Scelta del tipo di confronto per l'individuazione dei nuovi prezzi
        # Usa YesNoDialog: Yes -> VARIANTE, No -> CONTABILITÀ
        if Dialogs.YesNoDialog(IconType="question",Title='Individuazione nuovi prezzi',
                            Text='Vuoi confrontare COMPUTO con VARIANTE?\n\n(Sì = VARIANTE, No = CONTABILITÀ)') == 1:
            confronto_col = 20  # Colonna VARIANTE
            confronto_nome = 'VARIANTE'
        else:
            confronto_col = 21  # Colonna CONTABILITÀ
            confronto_nome = 'CONTABILITÀ'

        # Chiedi all'utente se vuole aggiungere il prefisso NP
        if Dialogs.YesNoDialog(IconType="question",Title='Nuovi Prezzi',
            Text='Vuoi aggiungere il prefisso "NPxx_"\n\nai codici dei nuovi prezzi?') == 1:
            add_prefix = True
        else:
            add_prefix = False

        lrow = SheetUtils.getUsedArea(oSheet).EndRow
        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start('Elaborazione in corso...', lrow)

        oCellRangeAddr = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
        var = 0
        cont = 0
        i = 3
        np_counter = 1
        code_mappings = {}

        # Prima passata per trovare il numero progressivo più alto tra i prefissi NP esistenti
        max_np = 0
        np_codes_found = []
        for el in range(3, lrow):
            current_code = oSheet.getCellByPosition(0, el).String

            # Usa regex per rilevare NPxx_ o VDS_NPxx_
            match_np = re.match(r'^(VDS_)?NP(\d{2})_', current_code)
            if match_np:
                np_num = int(match_np.group(2))
                if np_num > max_np:
                    max_np = np_num
                np_codes_found.append(current_code)

        # Informa l'utente se sono stati trovati prefissi NP esistenti
        if max_np > 0 and add_prefix:
            message = f"Sono stati trovati {len(np_codes_found)} codici con prefisso NPxx_ esistente.\n"
            message += f"Il numero progressivo più alto trovato è NP{max_np:02d}_.\n\n"
            message += "Vuoi iniziare la numerazione da NP{:02d}_?".format(max_np + 1)

            if Dialogs.YesNoDialog(Title='Prefissi NP esistenti trovati',
                                Text=message) == 1:
                np_counter = max_np + 1
            else:
                # L'utente vuole resettare la numerazione
                if Dialogs.YesNoDialog(IconType="question",Title='Conferma',
                                    Text='Vuoi resettare la numerazione partendo da NP01_?') == 1:
                    np_counter = 1
                else:
                    add_prefix = False
                    Dialogs.Info(Title='Operazione annullata', Text='Non verranno aggiunti nuovi prefissi NP.')

        # Seconda passata per applicare i nuovi prefissi
        for el in range(3, lrow):
            i += 1
            indicator.Value = i
            val19 = oSheet.getCellByPosition(19, el).Value
            val_confronto = oSheet.getCellByPosition(confronto_col, el).Value

            if val19 == 0 and val_confronto > 0:
                current_code = oSheet.getCellByPosition(0, el).String

                # Verifica se il codice ha già un prefisso NPxx_
                has_np_prefix = bool(re.match(r'^(VDS_)?NP\d{2}', current_code))

                if add_prefix and not has_np_prefix:
                    # Gestione del prefisso VDS_
                    if current_code.startswith('VDS_'):
                        new_code = f'VDS_NP{np_counter:02d}_{current_code[4:]}'  # Mantiene VDS_ e aggiunge NPxx_
                    else:
                        new_code = f'NP{np_counter:02d}_{current_code}'

                    code_mappings[current_code] = new_code
                    oSheet.getCellByPosition(0, el).String = new_code
                    np_counter += 1

                # Evidenzia e somma i nuovi prezzi
                oSheet.getCellByPosition(confronto_col, el).CellBackColor = 16770000
                if confronto_col == 21:
                    cont += val_confronto
                else:
                    var += val_confronto
            else:
                oCellRangeAddr.StartRow = el
                oCellRangeAddr.EndRow = el
                oSheet.group(oCellRangeAddr, 1)
                oSheet.getCellRangeByPosition(0, el, 1, el).Rows.IsVisible = True

        if add_prefix and code_mappings:
            indicator.end()
            indicator.start('Propago i nuovi codici...', len(oDoc.Sheets))

            for sheet_index, sheet in enumerate(oDoc.Sheets):
                indicator.Value = sheet_index + 1
                if sheet.Name != oSheet.Name:
                    used_range = SheetUtils.getUsedArea(sheet)
                    if used_range:
                        # Determina la colonna corretta in base al nome del foglio
                        search_col = 1 if sheet.Name in ["COMPUTO", "VARIANTE", "CONTABILITA"] else 0

                        for row in range(used_range.StartRow, used_range.EndRow + 1):
                            cell_value = sheet.getCellByPosition(search_col, row).String
                            if cell_value in code_mappings:
                                sheet.getCellByPosition(search_col, row).String = code_mappings[cell_value]

        indicator.end()

        # Messaggio finale coerente con la scelta di confronto
        if confronto_col == 21 and cont > 0:
            Dialogs.Info(
                Title='Informazione',
                Text=f'''Il totale delle variazioni\nin CONTABILITÀ è di € {cont:,.2f}'''
                    .replace(',', 'X').replace('.', ',').replace('X', '.')
            )
        elif confronto_col == 20 and var > 0:
            Dialogs.Info(
                Title='Informazione',
                Text=f'''Il totale delle variazioni\nin VARIANTE è di € {var:,.2f}'''
                    .replace(',', 'X').replace('.', ',').replace('X', '.')
            )

        LeenoUtils.DocumentRefresh(True)

        return

###############################################################################
###############################################################################
###############################################################################


def _generate_new_code(current_code, prefix):
    '''Genera il nuovo codice mantenendo eventuale prefisso VDS_'''
    if current_code.startswith('VDS_'):
        return f'VDS_{prefix}{current_code[4:]}'
    return f'{prefix}{current_code}'

def _handle_non_new_price_row(oSheet, oCellRangeAddr, row):
    '''Gestisce le righe che non sono nuovi prezzi'''
    oCellRangeAddr.StartRow = row
    oCellRangeAddr.EndRow = row
    oSheet.group(oCellRangeAddr, 1)
    oSheet.getCellRangeByPosition(0, row, 1, row).Rows.IsVisible = True

def _propagate_code_changes(oDoc, current_sheet_name, code_mappings, indicator):
    '''Propaga i cambiamenti di codice a tutti i fogli'''
    indicator = oDoc.getCurrentController().getStatusIndicator()
    indicator.start('Propago i nuovi codici...', len(oDoc.Sheets)-1)    
    
    for sheet_index, sheet in enumerate(oDoc.Sheets):
        if sheet.Name != current_sheet_name:
            indicator.Value = sheet_index
            _update_sheet_codes(sheet, code_mappings)

def _update_sheet_codes(sheet, code_mappings):
    '''Aggiorna i codici in un singolo foglio'''
    used_range = SheetUtils.getUsedArea(sheet)
    if used_range:
        search_col = 1 if sheet.Name in {"COMPUTO", "VARIANTE", "CONTABILITA"} else 0
        
        for row in range(used_range.StartRow, used_range.EndRow + 1):
            cell = sheet.getCellByPosition(search_col, row)
            if cell.String in code_mappings:
                cell.String = code_mappings[cell.String]

def _show_results(var_total, cont_total):
    '''Mostra i dialoghi con i risultati'''
    if var_total > 0:
        Dialogs.Info(
            Title='Informazione',
            Text=f"Il totale delle variazioni\nin VARIANTE è di € {var_total:,.2f}"
                .replace(',', 'X').replace('.', ',').replace('X', '.')
        )

    if cont_total > 0:
        Dialogs.Info(
            Title='Informazione',
            Text=f"Il totale delle variazioni\nin CONTABILITÀ è di € {cont_total:,.2f}"
                .replace(',', 'X').replace('.', ',').replace('X', '.')
        )

###############################################################################
###############################################################################
###############################################################################
########################################################################
def MENU_hl():
    '''
    Sostituisce hyperlink alla stringa nella colonna in cui è la cella
    selezionata, se questa è un indirizzo di file o cartella ctrl-shift-h
    '''

    # Ottieni il documento corrente e il foglio attivo
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # Ottieni la posizione corrente della cella
    lcol = LeggiPosizioneCorrente()[0]
    lrow= LeggiPosizioneCorrente()[1]

    if oSheet.getCellByPosition(lcol, lrow).Type.value == 'EMPTY':
        comando("Paste")

    # Itera sulle righe del foglio, partendo dall'ultima e andando verso l'alto
    for el in reversed(range(0, SheetUtils.getUsedArea(oSheet).EndRow + 1)):
        try:
            # Ottieni la stringa nella cella corrente
            cell_string = oSheet.getCellByPosition(lcol, el).String

            # Verifica se la stringa rappresenta un indirizzo di file o cartella
            #  if cell_string[1] == ':' or cell_string[2] == ':' or cell_string[0:1] == '\\':
            if ':' in cell_string :
                cell_string = cell_string.replace('"', '')
                # Costruisci la formula per l'iperlink
                hyperlink_formula = '=HYPERLINK("' + cell_string + '";"►►►")' # >>>
                # Applica la formula all'interno della cella
                oSheet.getCellByPosition(lcol, el).Formula = hyperlink_formula
            elif '@' in cell_string :
                cell_string = cell_string.replace('"', '')
                # Costruisci la formula per l'iperlink
                hyperlink_formula = '=HYPERLINK("mailto:' + cell_string + '";"@>>")'
                # Applica la formula all'interno della cella
                oSheet.getCellByPosition(lcol, el).Formula = hyperlink_formula
        except Exception as e:
            # DLG.errore(e)
            pass



########################################################################
def MENU_filtro_descrizione():
    '''
    Raggruppa e nasconde tutte le voci di misura in cui non compare
    la stringa cercata.
    '''
    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet

        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        fine = SheetUtils.getUsedArea(oSheet).EndRow + 1
        el_y = []
        if 'comp 1-a' in oDoc.getCurrentSelection().CellStyle or \
        'Descr' in oDoc.getCurrentSelection().CellStyle:
            testo = oDoc.getCurrentSelection().String
        else:
            testo = ''
        descrizione = InputBox(
            testo, t='Inserisci la descrizione da cercare o OK per conferma.')
        if descrizione in (None, '', ' '):
            struttura_off()
            oSheet.getCellRangeByPosition(2, 0, 2, 1048575).clearContents(HARDATTR)
            LeenoUtils.DocumentRefresh(True)
            return

        struttura_off()
        oSheet.getCellRangeByPosition(2, 0, 2, 1048575).clearContents(HARDATTR)

        y = 4
        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start('Applicazione filtro...', fine)
        
        indicator.setValue(0)
        lRow = SheetUtils.sStrColtoList(descrizione, 2, oSheet, y)
        if len(lRow) == 0:
            indicator.end()
            # DLG.MsgBox('''Testo non trovato.''', 'ATTENZIONE!')
            Dialogs.Exclamation (Title = 'ATTENZIONE!',
            Text="Testo non trovato.")

            return
        el_y = []
        for y in lRow:
            indicator.setValue(y)
            oSheet.getCellByPosition(2, y).CellBackColor = 15757935
            el_y.append(seleziona_voce(y))
        lista_y = []
        lista_y.append(2)
        for el in el_y:
            y = el[0]
            indicator.setValue(y)
            lista_y.append(y)
            y = el[1]
            lista_y.append(y)
        if oSheet.Name == 'CONTABILITA':
            lista_y.append(fine - 2)
        else:
            lista_y.append(fine - 3)
        i = 0
        while len(lista_y[i:]) > 1:
            SR = lista_y[i:][0] + 1
            ER = lista_y[i:][1]
            if ER > SR:
                oCellRangeAddr.StartRow = SR
                oCellRangeAddr.EndRow = ER - 1
                oSheet.group(oCellRangeAddr, 1)
                oSheet.getCellRangeByPosition(0, SR, 0,
                                            ER - 1).Rows.IsVisible = False
            i += 2
        _gotoCella(2, lRow[0])
        indicator.end()

########################################################################

def somma():
    '''
    Mostra la somme dei valori di una selezione su singola colonna. Utile quando la
    finestra di Calc non mostra la statusbar.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lcol = LeggiPosizioneCorrente()[0]
    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
    el_y = []
    lista_y = []
    try:
        len(oRangeAddress)
        for el in oRangeAddress:
            el_y.append((el.StartRow, el.EndRow))
    except TypeError:
        el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))
    for y in el_y:
        for el in range(y[0], y[1] + 1):
            lista_y.append(el)
    somma = []
    for y in lista_y:
        somma.append(oSheet.getCellByPosition(lcol, y).Value)
    DLG.chi(sum(somma))


########################################################################

def calendario():
    '''
    Mostra un calendario da cui selezionare la data e la restituisce
    in formato gg/mm/aaaa.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    x = LeggiPosizioneCorrente()[0]
    y = LeggiPosizioneCorrente()[1]
    testo = Dialogs.pickDate()
    lst = str(testo).split('-')
    try:
        testo = lst[2] + '/' + lst[1] + '/' + lst[0]
    except:
        pass

    return testo


def PdfDlg():
    # dimensione verticale dei checkbox == dimensione bottoni
    #dummy, hItems = Dialogs.getButtonSize('', Icon="Icons-24x24/settings.png")
    nWidth, hItems = Dialogs.getEditBox('aa')

    # dimensione dell'icona col PDF
    imgW = Dialogs.getBigIconSize()[0] * 2

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheets = list(oDoc.getSheets().getElementNames())


    return Dialogs.Dialog(Title='Esportazione PDFPDFPDFPDFPDF',  Horz=False, CanClose=True,  Items=[
        Dialogs.HSizer(Items=[
            Dialogs.VSizer(Items=[
                Dialogs.Spacer(),
                Dialogs.ImageControl(Image='Icons-Big/pdf.png', MinWidth=imgW / 10),
                Dialogs.Spacer(),
            ]),
            # Dialogs.VSizer(Items=[
                # Dialogs.FixedText(Text='Elenco_prova'),
                # Dialogs.Spacer(),
                # Dialogs.Edit(Id='npElencoPrezzi', Align=1, FixedHeight=hItems, FixedWidth=nWidth),
                # Dialogs.Spacer(),
                # Dialogs.Edit(Id='npComputoMetrico', Align=1, FixedHeight=hItems, FixedWidth=nWidth),
                # Dialogs.Spacer(),
                # Dialogs.Edit(Id='npCostiManodopera', Align=1, FixedHeight=hItems, FixedWidth=nWidth),
                # Dialogs.Spacer(),
                # Dialogs.Edit(Id='npQuadroEconomico', Align=1, FixedHeight=hItems, FixedWidth=nWidth),
            # ]),
            Dialogs.Spacer(),
            Dialogs.VSizer(Items=[
                Dialogs.FixedText(Text='Oggetto'),
                Dialogs.Spacer(),
                Dialogs.ListBox(List=oSheets, FixedHeight=hItems * 1, FixedWidth=nWidth * 6),
                # Dialogs.Spacer(),
                # Dialogs.CheckBox(Id="cbElencoPrezzi", Label="Elenco prezzi", FixedHeight=hItems),
                # Dialogs.Spacer(),
                # Dialogs.CheckBox(Id="cbComputoMetrico", Label="Computo metrico", FixedHeight=hItems),
                # Dialogs.Spacer(),
                # Dialogs.CheckBox(Id="cbCostiManodopera", Label="Costi manodopera", FixedHeight=hItems),
                # Dialogs.Spacer(),
                # Dialogs.CheckBox(Id="cbQuadroEconomico", Label="Quadro economico", FixedHeight=hItems),
            ]),
            Dialogs.Spacer(),
        ]),
        Dialogs.Spacer(),
        Dialogs.Spacer(),
        Dialogs.FixedText(Text='Cartella di destinazione:'),
        Dialogs.Spacer(),
        Dialogs.PathControl(Id="pathEdit"),
        Dialogs.Spacer(),
        Dialogs.HSizer(Items=[
            Dialogs.Spacer(),
            Dialogs.Button(Label='Ok', MinWidth=Dialogs.MINBTNWIDTH, Icon='Icons-24x24/ok.png',  RetVal=1),
            Dialogs.Spacer(),
            Dialogs.Button(Label='Annulla', MinWidth=Dialogs.MINBTNWIDTH, Icon='Icons-24x24/cancel.png',  RetVal=-1),
            Dialogs.Spacer()
        ])
    ])

########################################################################

def tempo():
    '''
    in attesa di tradurre gina_furbetta_2 dal basic
    '''
    stringa = ''.join(''.join(''.join(
        str(datetime.now()).split('.')[0].split(' ')).split('-')).split(
            ':'))[:12]
    return stringa


def stampa_PDF():
    DelPrintArea()
    set_area_stampa()
    # tempo = ''.join(''.join(''.join(
        # str(datetime.now()).split('.')[0].split(' ')).split('-')).split(
            # ':'))[:12]
    oDoc = LeenoUtils.getDocument()
    orig = oDoc.getURL()
    dest = orig.split('.')[0] + '-' + tempo() + '.pdf'
    ods2pdf(oDoc, dest)
    # DLG.chi(dest)
    # rem ----------------------------------------------------------------------


########################################################################
########################################################################
########################################################################
def inputbox(message, title="", default="", x=None, y=None):
    """ Shows dialog with input box.
        @param message message to show on the dialog
        @param title window title
        @param default default value
        @param x dialog positio in twips, pass y also
        @param y dialog position in twips, pass y also
        @return string if OK button pushed, otherwise zero length string
    """
    WIDTH = 600
    HORI_MARGIN = VERT_MARGIN = 8
    BUTTON_WIDTH = 100
    BUTTON_HEIGHT = 26
    HORI_SEP = VERT_SEP = 8
    LABEL_HEIGHT = BUTTON_HEIGHT * 2 + 5
    EDIT_HEIGHT = 24
    HEIGHT = VERT_MARGIN * 2 + LABEL_HEIGHT + VERT_SEP + EDIT_HEIGHT
    import uno
    from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE
    from com.sun.star.awt.PushButtonType import OK, CANCEL
    from com.sun.star.util.MeasureUnit import TWIP
    ctx = uno.getComponentContext()
    def create(name):
        return ctx.getServiceManager().createInstanceWithContext(name, ctx)
    dialog = create("com.sun.star.awt.UnoControlDialog")
    dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
    dialog.setModel(dialog_model)
    dialog.setVisible(False)
    dialog.setTitle(title)
    dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)
    def add(name, type, x_, y_, width_, height_, props):
        model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + type + "Model")
        dialog_model.insertByName(name, model)
        control = dialog.getControl(name)
        control.setPosSize(x_, y_, width_, height_, POSSIZE)
        for key, value in props.items():
            setattr(model, key, value)
    label_width = WIDTH - BUTTON_WIDTH - HORI_SEP - HORI_MARGIN * 2
    add("label", "FixedText", HORI_MARGIN, VERT_MARGIN, label_width, LABEL_HEIGHT,
        {"Label": str(message), "NoLabel": True})
    add("btn_ok", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN,
            BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})
    add("btn_cancel", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN + BUTTON_HEIGHT + 5,
            BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": CANCEL})
    add("edit", "Edit", HORI_MARGIN, LABEL_HEIGHT + VERT_MARGIN + VERT_SEP,
            WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(default)})
    frame = create("com.sun.star.frame.Desktop").getCurrentFrame()
    window = frame.getContainerWindow() if frame else None
    dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
    if not x is None and not y is None:
        ps = dialog.convertSizeToPixel(uno.createUnoStruct("com.sun.star.awt.Size", x, y), TWIP)
        _x, _y = ps.Width, ps.Height
    elif window:
        ps = window.getPosSize()
        _x = ps.Width / 2 - WIDTH / 2
        _y = ps.Height / 2 - HEIGHT / 2
    dialog.setPosSize(_x, _y, 0, 0, POS)
    edit = dialog.getControl("edit")
    edit.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(default))))
    edit.setFocus()
    ret = edit.getModel().Text if dialog.execute() else ""
    dialog.dispose()
    return ret

########################################################################


def nuove_icone(chiaro = True):
    '''
    Imposta il tema chiaro/scuro delle toolbar di LeenO
    Copia le icone SVG nella cartella ../icons/
    '''
    if chiaro == True:
        fn = 'svg'
        e = 'Chiaro'
    else:
        fn = 'scuro'
        e = 'Scuro'
    fname = uno.fileUrlToSystemPath(LeenO_path()) + '/icons/' + fn
    files_in_folder = os.listdir(fname)

    # attiva la progress bar
    Title=f'Impostazione del tema {e} di LeenO...'

    progress = Dialogs.Progress(Title=f'Impostazione del tema {e} di LeenO...', Text="Lettura dati")
    progress.setLimits(0, len(files_in_folder))
    progress.show()

    step = 0
    progress.setValue(step)
    for el in files_in_folder:
        file_path = fname + '/' + el
        # Estrai il nome del file senza estensione
        fn = os.path.splitext(os.path.basename(file_path))[0]

        file_name = uno.fileUrlToSystemPath(LeenO_path()) + '/icons/' + fn

        # Crea i nomi dei nuovi file BMP
        bmp_26h = file_name + "_26h.bmp"
        bmp_26 = file_name + "_26.bmp"
        bmp_16h = file_name + "_16h.bmp"
        bmp_16 = file_name + "_16.bmp"

        # Copia il file SVG selezionato nei nuovi file BMP
        try:
            shutil.copy(file_path, bmp_26h)
        except Exception as e:
            DLG.errore(e)
        shutil.copy(file_path, bmp_26)
        shutil.copy(file_path, bmp_16h)
        shutil.copy(file_path, bmp_16)
        step += 1
        progress.setValue(step)
    progress.hide()

    Dialogs.Info(Title='Info', Text=f"Tema {e} di LeenO impostato con successo. Riavviare LibreOffice!")
    return

def celle_colorate(flag = False):
    "*** DA COPLETARE ***"
    '''
    Conta le celle colorate di giallo e inserisce il risultato a destra
    dell'ultima colonna selezionata.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    # DLG.chi(oSheet.getCellRangeByName('J5').CellBackColor)
    # return

    oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()

    SR = oRangeAddress.StartRow
    ER = oRangeAddress.EndRow
    SC = oRangeAddress.StartColumn
    EC = oRangeAddress.EndColumn
    color = 0
    for y in range(SR, ER+1):
        somma = 0
        if flag == False:
            for x in range(SC, EC+1):
                if oSheet.getCellByPosition(x, y).CellBackColor == color:
                    somma += 1
            if somma >0:
                oSheet.getCellByPosition(x+1, y).Value = somma
        elif flag == 9:
            for x in range(SC, EC+1):
                if oSheet.getCellByPosition(x, y).CellBackColor == color:
                    somma += 1
            if somma >0:
                oSheet.getCellByPosition(x+1, y).Value = somma
    return

import LeenoTabelle


########################################################################

# def create_progress_bar(title='', steps=100):
#     '''
#     Crea e mostra la barra di avanzamento di Calc.

#     Args:
#         title (str): Il titolo della barra di avanzamento.
#         steps (int): Il numero totale di passi della barra di avanzamento.

#     Returns:
#         progressBar: L'oggetto creato.
#     '''

#     caller_frame = inspect.stack()[1]
#     line_number = caller_frame.lineno
#     full_file_path = caller_frame.filename  # Ottieni il percorso completo
#     function_name = caller_frame.function  # Nome della funzione chiamante
#     file_name = os.path.basename(full_file_path)  # Solo il nome del file

    
#     # title = f"{self._text} ({percent})"
#     if 'giuserpe' in os.getlogin():
#         title = title + f"    Funzione: {function_name}()    Linea: {line_number}    File:{file_name}"

#     desktop = LeenoUtils.getDesktop()
#     model = desktop.getCurrentComponent()
#     controller = model.getCurrentController()
#     frame = controller.getFrame()

#     # Crea un nuovo oggetto indicatore di stato
#     progressBar = frame.createStatusIndicator()

#     # Mostra la barra di avanzamento e imposta il titolo
#     progressBar.start(title, steps)
#     # progressBar.Value = 0

#     return progressBar



########################################################################

def sposta_voce(lrow=None, msg=1):
    '''
    Sposta la voce selezionata in una posizione indicata dall'utente.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    
    SR, ER = seleziona_voce()
    
    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, SR, 250, ER))
    comando('Copy')

    # Richiede all'utente di selezionare una voce di riferimento o indicare il numero d'ordine
    to = basic_LeenO('ListenersSelectRange.getRange', "Seleziona voce di riferimento o indica nuovo numero d'ordine")

    # Se la voce non è riferita al foglio attivo, cerchiamo la colonna corrispondente
    if oSheet.Name not in to:
        if to == '1':
            to = '$' + oSheet.Name + '.$C$3'
        else:
            # to = str(int(to) - 1)
            col_num = SheetUtils.uFindStringCol(to, 0, oSheet, equal=1)
            to = '$' + oSheet.Name + '.$C$' + str(col_num)
    
    try:
        to_row = int(to.split('$')[-1]) - 1
    except ValueError:
        # Se non è possibile convertire, ricarica il documento e interrompe l'esecuzione
        LeenoUtils.DocumentRefresh(True)
        return

    # Trova la prossima voce disponibile nel foglio
    lrow = LeenoSheetUtils.prossimaVoce(oSheet, to_row, 1, True)
    
    # Vai alla cella individuata e incolla i dati copiati
    _gotoCella(0, lrow)
    paste_clip(insCells=1)
    
    if to_row < SR:
        add = ER - SR
        SR = SR + add + 1
        ER = ER + add + 1

    # Rimuovi la voce originale (se necessario)
    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, SR, 250, ER))
    comando('DeleteRows')  # Delezione delle righe originali

    try:
        add
        _gotoCella(1, lrow + 1)
    except:
        _gotoCella(1, lrow - ER + SR)

    lrow = LeggiPosizioneCorrente()[1]

    oDoc.CurrentController.setFirstVisibleRow(lrow - 8)
    numera_voci()
    return




def copia_stili_celle(sheet_src, range_src, sheet_dest, range_dest):
    datarif = datetime.now()

    '''
    Copia gli stili di un range da un foglio all'altro.
    '''
    cell_range_src = sheet_src.getCellRangeByName(range_src)
    cell_range_dest = sheet_dest.getCellRangeByName(range_dest)

    # Ottieni dimensioni range
    rows = cell_range_src.RangeAddress.EndRow - cell_range_src.RangeAddress.StartRow + 1
    cols = cell_range_src.RangeAddress.EndColumn - cell_range_src.RangeAddress.StartColumn + 1

    # attiva la progressbar

    progress = Dialogs.Progress(Title='copia...', Text="Lettura dati")
    indicator = LeenoUtils.getDocument().getCurrentController().getStatusIndicator()
    indicator.start('Copia formattazione celle...', rows)
    indicator.setValue(0)

    # Copia formattazione, bordi e allineamento
    for r in range(rows):
        indicator.setValue(r)
        for c in range(cols):
            cell_src = sheet_src.getCellByPosition(
                cell_range_src.RangeAddress.StartColumn + c,
                cell_range_src.RangeAddress.StartRow + r
            )
            cell_dest = sheet_dest.getCellByPosition(
                cell_range_dest.RangeAddress.StartColumn + c,
                cell_range_dest.RangeAddress.StartRow + r
            )

            cell_dest.CellStyle = cell_src.CellStyle
    indicator.end()
    DLG.chi('eseguita in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!')
    return

########################################################################
########################################################################


def DocumentRefresh(boo):
    oDoc = LeenoUtils.getDocument()
    if boo == True:
        oDoc.enableAutomaticCalculation(True)
        oDoc.removeActionLock()
        oDoc.unlockControllers()

    elif boo == False:
        oDoc.enableAutomaticCalculation(False)
        oDoc.lockControllers()
        oDoc.addActionLock()
########################################################################
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
def trova_colore_cella():
    oDoc = LeenoUtils.getDocument()
    active_cell = oDoc.CurrentSelection
    DLG.chi(active_cell.CellBackColor)
    return

def cerca_rosso_mancante():
    """
    Trova la prima riga in cui una cella con stile 'ROSSO' 
    (colonne 2, 5, 6, 8) non ha 'ROSSO' in colonna 9. 
    Sposta il cursore su quella cella e termina.
    """
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    urow = getLastUsedCell(oSheet).EndRow + 1
    lrow = LeggiPosizioneCorrente()[1] + 1
    for el in range(lrow, urow):
        if 'ROSSO' in oSheet.getCellByPosition(2, el).CellStyle or \
            'ROSSO' in oSheet.getCellByPosition(5, el).CellStyle or \
            'ROSSO' in oSheet.getCellByPosition(6, el).CellStyle or \
            'ROSSO' in oSheet.getCellByPosition(6, el).CellStyle or \
            'ROSSO' in oSheet.getCellByPosition(8, el).CellStyle:
            if 'ROSSO' not in oSheet.getCellByPosition(9, el).CellStyle:
                _gotoCella(9, el)
                return
            else:
                pass
    return


def Main_Riordina_Analisi_Alfabetico():
    DLG.chi('in allestimento...')
    return
    chiudi_dialoghi()
    articoli = []
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.Sheets.getByName("Analisi di Prezzo")
    lLastUrow = SheetUtils.getLastUsedRow(oSheet)  # ultima riga editata
    
    oDoc.CurrentController.select(oSheet.getCellByPosition(0, 3))
    
    # Trovo il punto di inserimento
    lrowDest = 2
    for i in range(11):  # 0 to 10
        cell = oSheet.getCellByPosition(0, i)
        if cell.CellStyle == "An.1v-Att Start" and cell.String == "COD./N.":
            lrowDest = i - 1  # punto di inserimento prima della prima scheda
            break
    
    # Recupero i codici presenti dalle ANALISI DI PREZZO
    for i in range(lLastUrow + 1):
        cell = oSheet.getCellByPosition(0, i)
        if cell.CellStyle == "An-1_sigla":
            art = cell.String  # articolo
            
            # Verifica doppioni
            if art in articoli:
                msg = f"Mi fermo! Il codice:\n\t\t\t\t\t\t{art}\nè presente più volte. Correggi e ripeti il comando."
                Dialogs.MessageBox(msg, "Avviso!", "OK")
                return
            
            articoli.append(art)
    
    # Riordino la lista alfabeticamente
    articoli.sort()
    
    # Process each item in sorted order
    for el in articoli:
        for i in range(lLastUrow + 1):
            if oSheet.getCellByPosition(0, i).String == el:  # trovo l'inizio della scheda
                inizio = i - 1
                
                # Trovo la fine della scheda
                for x in range(i, i + 101):
                    if oSheet.getCellByPosition(0, x).String == "----":
                        fine = x + 1
                        i = x + 1
                        nrighe = fine - inizio  # ampiezza in righe della scheda
                        
                        # Inserisci spazio per la scheda
                        oSheet.getRows().insertByIndex(lrowDest, nrighe + 1)
                        
                        # Seleziona e copia le righe
                        selezione = oSheet.getCellRangeByPosition(0, inizio + nrighe, 250, fine + nrighe)
                        oDoc.CurrentController.select(selezione)
                        oDest = oSheet.getCellByPosition(0, lrowDest).CellAddress
                        oSheet.copyRange(oDest, selezione.RangeAddress)
                        
                        # Cancella la vecchia scheda
                        oSheet.getRows().removeByIndex(inizio + nrighe, nrighe + 1)
                        
                        lrowDest += nrighe + 1
                        break
                break
def applica_barre_dati():
    try:
        # Ottieni il documento e il foglio attivo
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        
        # Crea l'oggetto DataBar
        data_bar = oDoc.createInstance("com.sun.star.sheet.DataBar")
        
        # Imposta le proprietà usando setPropertyValue
        data_bar.setPropertyValue("Color", 0x0000FF)  # Blu
        data_bar.setPropertyValue("MinLength", 0)
        data_bar.setPropertyValue("MaxLength", 100)
        
        # Ottieni il range di celle
        cell_range = oSheet.getCellRangeByName("E1:E10")
        
        # Aggiungi la formattazione condizionale
        cond_format = cell_range.ConditionalFormat
        cond_format.addNew(("com.sun.star.sheet.DataBar",), data_bar)
        
        DLG.chi("Barre dati applicate con successo!")
    except Exception as e:
        DLG.chi(f"Errore durante l'applicazione delle barre dati: {str(e)}")

########################################################################

def MENU_export_selected_range_to_odt():
    # with LeenoUtils.DocumentRefreshContext(False):
    export_selected_range_to_odt()


def export_selected_range_to_odt():
    """
    Esporta l'intervallo di celle selezionato in Calc in un nuovo documento Writer (ODT).
    Solo righe e colonne visibili, tabulazione a destra con puntini e paragrafi giustificati.
    """
    try:
        SEPARATORS = {
            0: ": ",
            1: "\rAl ",
            2: "\t€ ",
        }

        oDoc = LeenoUtils.getDocument()
        selection = oDoc.getCurrentSelection()

        if not selection.supportsService("com.sun.star.sheet.SheetCellRange"):
            DLG.chi("Seleziona un range di celle prima di eseguire la macro!")
            return

        output_path = Dialogs.FileSelect('Salva con nome...', '*.odt', 1)
        if not output_path:
            return
        if not output_path.endswith('.odt'):
            output_path += '.odt'

        desktop = LeenoUtils.getDesktop()
        writer_doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
        writer_text = writer_doc.Text

        cursor = writer_text.createTextCursor()
        try:
            cursor.ParaAdjust = 2              # Giustificato (BLOCK)
            cursor.ParaFirstLineIndent = 300   # Rientro prima riga di 0.3 cm
        except Exception:
            pass

        try:
            page_styles = writer_doc.getStyleFamilies().getByName("PageStyles")
            page_style_name = writer_doc.CurrentController.ViewCursor.PageStyle
            page_style = page_styles.getByName(page_style_name)
            page_width = int(getattr(page_style, "Width", 21000))
            left_margin = int(getattr(page_style, "LeftMargin", 2000))
            right_margin = int(getattr(page_style, "RightMargin", 2000))
        except Exception:
            page_width = 21000
            left_margin = 2000
            right_margin = 2000

        usable_width = page_width - left_margin - right_margin
        tab_position = left_margin + usable_width

        try:
            tabstop = uno.createUnoStruct("com.sun.star.style.TabStop")
            tabstop.Position = int(tab_position)
            tabstop.Alignment = 2
            tabstop.FillChar = ord('.')
            cursor.ParaTabStops = (tabstop,)
        except Exception:
            pass

        rows = selection.getRows()
        cols = selection.getColumns()
        visible_cols = [i for i in range(cols.getCount()) if cols.getByIndex(i).IsVisible]

        for row_idx in range(rows.getCount()):
            row = rows.getByIndex(row_idx)
            if not row.IsVisible:
                continue

            for col_pos, col_idx in enumerate(visible_cols):
                cell = selection.getCellByPosition(col_idx, row_idx)
                raw_value = ""
                try:
                    raw_value = (cell.getString() or "").strip()
                except:
                    raw_value = ""
                if not raw_value:
                    try:
                        v = cell.getValue()
                        raw_value = str(v) if v != 0 else ""
                    except:
                        raw_value = ""

                cell_value = raw_value.replace('\n', ' ').replace('\r', ' ')

                try:
                    if getattr(cell, "getType", None) and cell.getType().value == 2 and cell.getValue() != 0:
                        cell_value = f"{cell.getValue():,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                except:
                    pass

                # ✅ applica convert_number_string() solo all'ultima colonna visibile
                if col_idx == visible_cols[-1]:
                    if "," in cell_value or "." in cell_value:
                        converted = LeenoUtils.convert_number_string(cell_value)
                        if converted and converted != cell_value:
                            cell_value = f"{cell_value} (euro {converted})."

                try:
                    cursor.setPropertyValue("CharWeight", 150 if col_pos != 1 else 100)
                except:
                    pass

                writer_text.insertString(cursor, cell_value, False)

                if col_pos in SEPARATORS and col_pos < len(visible_cols) - 1:
                    try:
                        cursor.setPropertyValue("CharWeight", 150)
                    except:
                        pass
                    writer_text.insertString(cursor, SEPARATORS[col_pos], False)
                elif col_pos < len(visible_cols) - 1:
                    writer_text.insertString(cursor, " ", False)

            # try:
            #     cursor.ParaAdjust = 3
            #     cursor.ParaTopMargin = 200
            #     cursor.ParaBottomMargin = 200
            # except:
            #     pass

            writer_text.insertControlCharacter(
                cursor,
                uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK"),
                False
            )

        writer_doc.storeToURL(
            uno.systemPathToFileUrl(output_path),
            (PropertyValue("FilterName", 0, "writer8", 0),)
        )

        Dialogs.Info(
            Title='Informazione',
            Text=f"File creato con successo:\\n{output_path}"
        )

    except Exception as e:
        DLG.chi(f"Errore durante l'esportazione:\\n{str(e)}")
########################################################################

def struttura_Registro():
    '''
    Configura la struttura del Registro raggruppando
    le righe in base alla presenza della stringa "VDS_" nella colonna A.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.removeAllManualPageBreaks()

    oSheet.clearOutline()

    lrow = SheetUtils.getLastUsedRow(oSheet)
    
    start_group = None
    
    for row in range(3, lrow + 1):
        cell = oSheet.getCellByPosition(0, row)
        
        # Ottieni il valore della cella come stringa (gestisce tutti i tipi di dati)
        cell_value = ""
        try:
            cell_value = cell.getString()
        except:
            try:
                cell_value = str(cell.getValue())
            except:
                cell_value = ""
        
        oRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oRangeAddr.Sheet = oSheet.RangeAddress.Sheet
        
        # Se la cella contiene "VDS_", inizia/continua un gruppo
        if "VDS_" in cell_value:
            if start_group is None:
                start_group = row  # Inizio del blocco
        else:
            # Se troviamo una cella senza "VDS_" e c'è un gruppo aperto, chiudilo
            if start_group is not None:
                oRangeAddr.StartColumn = 0
                oRangeAddr.EndColumn = oSheet.Columns.Count - 1
                oRangeAddr.StartRow = start_group
                oRangeAddr.EndRow = row - 1
                
                oSheet.group(oRangeAddr, 1)
                start_group = None
    
def struttura_Registro():
    '''
    Configura la struttura del Registro raggruppando
    le righe in base alla presenza della stringa "VDS_" nella colonna A
    e raggruppa separatamente le celle senza "VDS_" con stile specifico.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.removeAllManualPageBreaks()

    oSheet.clearOutline()

    lrow = SheetUtils.getLastUsedRow(oSheet)
    
    start_group_vds = None  # Per raggruppare celle con "VDS_"
    start_group_no_vds = None  # Per raggruppare celle senza "VDS_" e con stile specifico
    
    for row in range(3, lrow + 1):
        cell = oSheet.getCellByPosition(0, row)
        
        # Ottieni il valore della cella come stringa
        cell_value = ""
        try:
            cell_value = cell.getString()
        except:
            try:
                cell_value = str(cell.getValue())
            except:
                cell_value = ""
        
        # Ottieni lo stile della cella
        cell_style = cell.CellStyle
        
        oRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oRangeAddr.Sheet = oSheet.RangeAddress.Sheet
        
        # Se la cella contiene "VDS_", gestisci il gruppo VDS
        if "VDS_" in cell_value:
            # Prima chiudi eventuale gruppo senza VDS
            if start_group_no_vds is not None:
                oRangeAddr.StartColumn = 0
                oRangeAddr.EndColumn = oSheet.Columns.Count - 1
                oRangeAddr.StartRow = start_group_no_vds
                oRangeAddr.EndRow = row - 1
                
                oSheet.group(oRangeAddr, 1)
                start_group_no_vds = None
            
            # Poi gestisci il gruppo VDS
            if start_group_vds is None:
                start_group_vds = row  # Inizio del blocco VDS
        else:
            # Se la cella NON contiene "VDS_" ma ha lo stile "List-stringa-sin"
            if cell_style == "List-stringa-sin":
                # Prima chiudi eventuale gruppo VDS
                if start_group_vds is not None:
                    oRangeAddr.StartColumn = 0
                    oRangeAddr.EndColumn = oSheet.Columns.Count - 1
                    oRangeAddr.StartRow = start_group_vds
                    oRangeAddr.EndRow = row - 1
                    
                    oSheet.group(oRangeAddr, 1)
                    start_group_vds = None
                
                # Poi gestisci il gruppo senza VDS
                if start_group_no_vds is None:
                    start_group_no_vds = row  # Inizio del blocco senza VDS
            else:
                # Se non è né VDS né cella con stile specifico, chiudi entrambi i gruppi
                if start_group_vds is not None:
                    oRangeAddr.StartColumn = 0
                    oRangeAddr.EndColumn = oSheet.Columns.Count - 1
                    oRangeAddr.StartRow = start_group_vds
                    oRangeAddr.EndRow = row - 1
                    
                    oSheet.group(oRangeAddr, 1)
                    start_group_vds = None
                
                if start_group_no_vds is not None:
                    oRangeAddr.StartColumn = 0
                    oRangeAddr.EndColumn = oSheet.Columns.Count - 1
                    oRangeAddr.StartRow = start_group_no_vds
                    oRangeAddr.EndRow = row - 1
                    
                    oSheet.group(oRangeAddr, 1)
                    start_group_no_vds = None
    
    # Alla fine del ciclo, chiudi eventuali gruppi ancora aperti
    if start_group_vds is not None:
        oRangeAddr.StartColumn = 0
        oRangeAddr.EndColumn = oSheet.Columns.Count - 1
        oRangeAddr.StartRow = start_group_vds
        oRangeAddr.EndRow = lrow
        
        oSheet.group(oRangeAddr, 1)
    
    if start_group_no_vds is not None:
        oRangeAddr.StartColumn = 0
        oRangeAddr.EndColumn = oSheet.Columns.Count - 1
        oRangeAddr.StartRow = start_group_no_vds
        oRangeAddr.EndRow = lrow
        
        oSheet.group(oRangeAddr, 1)

import LeenoContab as LC

def YesNoCancelDialog_DebugNames():
    ctx = LeenoUtils.getComponentContext()
    psm = ctx.ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDlg = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.DlgSiNoCancel?language=Basic&location=application"
    )
    model = oDlg.getModel()
    for name in model.getElementNames():
        DLG.chi(f"Controllo: {  name} " )


def stop_all_scripts_and_close_dialogs():
    """
    Ferma tutti gli script in esecuzione (Basic o Python)
    e chiude tutti i dialoghi UNO aperti.
    """
    try:
        # Ottiene il contesto UNO e il gestore servizi
        ctx = LeenoUtils.getComponentContext()
        smgr = ctx.ServiceManager

        # --- 1️⃣ Interrompe tutti gli script attivi ---
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
        frame = desktop.getCurrentFrame()

        # Equivale a premere “Interrompi esecuzione” in LibreOffice
        dispatcher.executeDispatch(frame, ".uno:Abort", "", 0, ())

        # --- 2️⃣ Chiude tutti i dialoghi aperti ---
        # Recupera il DialogProvider2 per iterare i dialoghi attivi
        dp2 = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider2", ctx)

        # Alcune versioni di LO non permettono di elencarli direttamente,
        # quindi cerchiamo i dialoghi globali noti nel runtime Basic
        bas_libs = smgr.createInstanceWithContext("com.sun.star.script.BasicLibraries", ctx)

        for lib_name in bas_libs.getElementNames():
            lib = bas_libs.getByName(lib_name)
            if hasattr(lib, "getElementNames"):
                for dlg_name in lib.getElementNames():
                    try:
                        dlg = lib.getByName(dlg_name)
                        if hasattr(dlg, "dispose"):
                            dlg.dispose()
                    except Exception:
                        pass

        # --- 3️⃣ Pulisce i residui in memoria (opzionale) ---
        import gc
        gc.collect()

        DLG.chi("✅ Tutti gli script interrotti e i dialoghi chiusi.")

    except Exception as e:
        DLG.chi(f"❌ Errore durante la chiusura: {e}")

########################################################################
########################################################################
########################################################################
def MENU_debug(Title=None, Testo=None):
    DLG.chi('Debug attivato!')
    return
    import app_bridge
    app_bridge.autocad("COMPUTO", "add")
    return
    attiva_autocad()
    return
    Copia_riga_Ent(num_righe=80)
    LeenoUtils.DocumentRefresh(True)


    # DLG.chi(Cancel())
    return

    Dialogs.YesNoCancelDialog(IconType="question", Image=None, Title= 'Debug', Text='Sei sicuro di voler eseguire il debug?')
    # DLG.chi(Dialogs.YesNoCancelDialog(IconType="question", Image=None, Title= 'Debug', Text='Sei sicuro di voler eseguire il debug?'))
    return

    inizializza_elenco()
    # genera_sommario()
    return
    sistema_cose()

    return
    # DLG.chi(Dialogs.YesNoDialog('Debug', 'Sei sicuro di voler eseguire il debug?'))
    YesNoDialog('Debug', 'Sei sicuro di voler eseguire il debug?')
    

    # chiudi_dialoghi()

    # DLG.chi(LeenoSheetUtils.cercaUltimaVoce(oSheet))
    return
    # LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    # oSheet.removeAllManualPageBreaks()
    # SheetUtils.visualizza_PageBreak()
    LC.firme_contabili()
    # LC.insrow()
    return
    chiudi_dialoghi()
    GotoSheet('S2')
    _primaCella(0, 0)
    return

    uRow = SheetUtils.getLastUsedRow(oSheet) + 1
    
    prezzi = [oSheet.getCellByPosition(4, el).String for el in range(1, uRow)]
    
    oSheetEp = oDoc.Sheets.getByName('Elenco Prezzi')
    highlight_color = 111111  # colore RGB valido (#111111)

    for el in prezzi:
        try:
            lrow = SheetUtils.uFindStringCol(el, 0, oSheetEp)
            if lrow >= 0:
                oSheetEp.getCellByPosition(0, lrow).CellBackColor = highlight_color
        except Exception:
            try:
                lrow = SheetUtils.uFindStringCol(el, 4, oSheet)
                if lrow >= 0:
                    oSheet.getCellByPosition(0, lrow).CellBackColor = 222222
            except Exception:
                pass





    # DLG.chi(oSheet.Name)
    return

    LeenoUtils.DocumentRefresh(False)
    return
    MENU_export_selected_range_to_odt()
    return  
    with LeenoUtils.DocumentRefreshContext(False):
        struttura_Registro()

    return
    lrow = LeggiPosizioneCorrente()[1]
    Circoscrive_Analisi(lrow)

    # inizializza_elenco()
    # export_selected_range_to_odt()
    # trova_np()
    # MENU_inserisci_Riga_rossa()
    return
    return
    # applica_barre_dati()
    # return
    LeenoSheetUtils.show_config_snippet()
    return

    LeenoSheetUtils.setLarghezzaColonne(oSheet)
    
########################################################################
# ELENCO DEGLI SCRIPT VISUALIZZATI NEL SELETTORE DI MACRO              #
# g_exportedScripts = (MENU_debug, )
########################################################################
########################################################################
# ... here is the python script code
# this must be added to every script file(the

# name org.openoffice.script.DummyImplementationForPythonScripts should be changed to something
# different(must be unique within an office installation !)
# --- faked component, dummy to allow registration with unopkg, no functionality expected
#  import unohelper
# questo mi consente di inserire i comandi python in Accelerators.xcu
# vedi pag.264 di "Manuel du programmeur oBasic"
# <<< vedi in description.xml
########################################################################
