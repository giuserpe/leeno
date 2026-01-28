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
import ctypes
import uno
import unohelper
import zipfile
import inspect

import tempfile

from pathlib import Path
# from uno import fileUrlToSystemPath, systemPathToFileUrl


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
# import Debug
import LeenoVariante

import LeenoConfig
from LeenoConfig import COLORE_COLONNE_RAFFRONTO, COLORE_GIALLO_VARIANTE, COLORE_ROSA_INPUT,\
      COLORE_VERDE_SPUNTA, COLORE_GRIGIO_INATTIVA, COLORE_BIANCO_SFONDO
cfg = LeenoConfig.Config()

import Dialogs

from undo_utils import with_undo, with_undo_batch, no_undo

# cos'e' il namespace:
# http://www.html.it/articol\i/il-misterioso-mondo-dei-namespaces-1/

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

    if cfg.read('Generale', 'nuova_voce') == 'True':
        oDlg_config.getControl('CheckBox7').State = 1

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

    if oDlg_config.getControl('CheckBox7').State == 1:
        cfg.write('Generale', 'nuova_voce', 'True')
    else:
        cfg.write('Generale', 'nuova_voce', 'False')


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
@LeenoUtils.no_refresh
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
@LeenoUtils.no_refresh
def invia_voce_interno():
    '''
    Invia le voci di Elenco Prezzi verso uno degli altri elaborati.
    Richiede comunque la scelta del DP
    '''
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
        Dialogs.NotifyDialog(IconType="warning",Title='AVVISO!',
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
    return

@with_undo('Invia voce a Variante')
@LeenoUtils.no_refresh
def invia_voce_interno():
    oDoc = LeenoUtils.getDocument()
    oSheetEP = oDoc.CurrentController.ActiveSheet

    elenco = seleziona()
    codici = [oSheetEP.getCellByPosition(0, el).String for el in elenco]
    meta = oSheetEP.getCellRangeByName('C2').String

    if meta == 'VARIANTE':
        # Chiamata alla nuova funzione: crea (se manca) ma NON svuota
        LeenoVariante.generaVariante(oDoc, clear=False)

    elif meta == 'CONTABILITA':
        LeenoContab.attiva_contabilita()
        oDoc.getSheets().getByName('CONTABILITA')

    elif meta == 'COMPUTO':
        GotoSheet('COMPUTO')

    else:
        # Messaggio di errore se C2 non è impostata correttamente
        Dialogs.NotifyDialog(
            IconType="warning",
            Title='AVVISO!',
            Text='Scegli in cella "C2" l\'elaborato di destinazione (COMPUTO, VARIANTE o CONTABILITA).'
        )
        _gotoCella(2, 1) # Torna sulla cella C2
        return

    # 4. Inserimento effettivo delle voci nell'elaborato scelto
    for codice_voce in codici:
        if meta == 'CONTABILITA':
            # Nota: GotoSheet dentro il loop può rallentare, ma garantisce il focus
            GotoSheet('CONTABILITA')
            ins_voce_contab(cod=codice_voce)
        else:
            # Per COMPUTO e VARIANTE la logica di inserimento è identica
            # Assicurati che ins_voce_computo inserisca nel foglio attivo
            GotoSheet(meta)
            LeenoComputo.ins_voce_computo(cod=codice_voce)

    # Pulizia finale della selezione blu
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))


###############################################################################
@LeenoUtils.no_refresh
def MENU_invia_voce():
    stato = cfg.read('Generale', 'pesca_auto')
    cfg.write('Generale', 'pesca_auto', 0)

    invia_voce()

    cfg.write('Generale', 'pesca_auto', stato)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    LeenoUtils.DocumentRefresh(True)


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
        return (analisi)

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
        ddcDoc.CurrentController.setFirstVisibleRow(3)
        _gotoCella(0, 4)

    # DLG.chi(1)

    # partenza
    if oSheet.Name == 'Elenco Prezzi':
        analisi = getAnalisi(oSheet)
        voce_da_inviare = oSheet.getCellByPosition(0, lrow).String
        dccSheet = ddcDoc.getSheets().getByName('Elenco Prezzi')
        # verifica presenza codice in EP
        cerca_in_elenco_prezzi = SheetUtils.uFindString(voce_da_inviare, dccSheet)
        if not cerca_in_elenco_prezzi:
            recupera_voce(voce_da_inviare)
        if nSheetDCC in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
            _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
            dccSheet = ddcDoc.getSheets().getByName(nSheetDCC)
            LeenoUtils.memorizza_posizione()
            if cfg.read('Generale', 'nuova_voce') == 'True':
                MENU_nuova_voce_scelta()
                lrow = LeggiPosizioneCorrente()
                dccSheet.getCellByPosition(lrow[0], lrow[1]).String = voce_da_inviare
                # dccSheet.getCellByPosition(lrow[0], lrow[1]).CellBackColor = COLORE_VERDE_SPUNTA
                _gotoCella(lrow[0]+1, lrow[1]+1)

            else:
                lrow = LeggiPosizioneCorrente()[1]
                LeenoComputo.cambia_articolo(dccSheet, lrow, voce_da_inviare)
                lrow = LeggiPosizioneCorrente()[1]
                # dccSheet.getCellByPosition(1, lrow).CellBackColor = COLORE_VERDE_SPUNTA
            LeenoUtils.ripristina_posizione()
        return

    # partenza
    if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):

        lrow = LeggiPosizioneCorrente()
        dv = LeenoComputo.DatiVoce(oSheet, lrow[1])
        art = dv.art
        ER = dv.ER
        SR = dv.SR

        range_src = f'A{SR+1}:AZ{ER+1}'

        data = oSheet.getCellRangeByName(range_src).FormulaArray

        # oSheet.getCellRangeByPosition(30, SR, 30, ER).CellBackColor = 15757935
        oSheet.getCellByPosition(1, SR +1).CellBackColor = COLORE_VERDE_SPUNTA

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
            dccSheet.getCellByPosition(1, SR + 1).CellBackColor = COLORE_VERDE_SPUNTA
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
            dccSheet.getCellByPosition(1, lrow).CellBackColor = COLORE_VERDE_SPUNTA
            _gotoCella(2, lrow)
        else:
            dccSheet.getCellByPosition(1, lrow + 1).CellBackColor = COLORE_VERDE_SPUNTA
            _gotoCella(2, lrow + 1)
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
        # subprocess.Popen(f'w: && cd {dest} && "W:/programmi/PortableGit/git-bash.exe" && gitk &', shell=True, stdout=subprocess.PIPE)
        subprocess.Popen(f'w: && cd {dest} && gitk &', shell=True, stdout=subprocess.PIPE)
    else:
        comandi = f'cd {dest} && mate-terminal && gitk &'
        if not processo('wish'):
            subprocess.Popen(comandi, shell=True, stdout=subprocess.PIPE)

    return


########################################################################

def cerca_path_valido():
    """
    Cerca il percorso di un editor di codice installato sul sistema.
    Priorità: Antigravity > VS Code
    Supporta Windows, Linux e macOS.

    Returns:
        str: Percorso completo dell'editor trovato

    Raises:
        FileNotFoundError: Se nessun editor viene trovato
    """
    if 'giuserpe' in os.getlogin():
        import platform
        system = platform.system()

        possible_paths = []

        # === ANTIGRAVITY (Priorità 1) ===
        if system == "Windows":
            possible_paths.extend([
                os.path.expanduser("~\\AppData\\Local\\Programs\\Antigravity\\Antigravity.exe"),
                "C:\\Program Files\\Antigravity\\Antigravity.exe",
                "C:\\Program Files (x86)\\Antigravity\\Antigravity.exe",
                "C:\\Users\\TEST\\AppData\\Local\\Programs\\Antigravity\\Antigravity.exe",
            ])
        elif system == "Linux":
            possible_paths.extend([
                "/usr/bin/antigravity",
                "/usr/local/bin/antigravity",
                os.path.expanduser("~/.local/bin/antigravity"),
            ])
        elif system == "Darwin":  # macOS
            possible_paths.extend([
                "/Applications/Antigravity.app/Contents/MacOS/Antigravity",
                os.path.expanduser("~/Applications/Antigravity.app/Contents/MacOS/Antigravity"),
            ])

        # === VS CODE (Fallback) ===
        if system == "Windows":
            possible_paths.extend([
                os.path.expanduser("~\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"),
                "C:\\Program Files\\Microsoft VS Code\\Code.exe",
                "C:\\Program Files (x86)\\Microsoft VS Code\\Code.exe",
                "C:\\Users\\giuserpe\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
                "C:\\Users\\DELL\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            ])
        elif system == "Linux":
            possible_paths.extend([
                "/usr/bin/code",
                "/usr/local/bin/code",
                "/snap/bin/code",
                os.path.expanduser("~/.local/bin/code"),
            ])
        elif system == "Darwin":  # macOS
            possible_paths.extend([
                "/Applications/Visual Studio Code.app/Contents/Resources/app/bin/code",
                "/usr/local/bin/code",
            ])

        editor_path = None
        for path in possible_paths:
            if os.path.exists(path):
                editor_path = path
                break

        if editor_path is None:
            raise FileNotFoundError(
                f"Impossibile trovare un editor (Antigravity o VS Code) su {system}. "
                "Assicurati che almeno uno sia installato."
            )
        return editor_path

def apri_con_editor(full_file_path, line_number):
    """
    Apre un file nell'editor di codice alla riga specificata.
    Riutilizza la finestra già aperta dell'editor se disponibile.

    Args:
        full_file_path (str): Percorso completo del file da aprire
        line_number (int): Numero di riga da visualizzare
    """
    editor_path = cerca_path_valido()

    # Controlla se il file esiste
    if not os.path.exists(full_file_path):
        DLG.chi(f"File non trovato: {full_file_path}")
        return

    # Controlla che il numero di riga sia valido
    if not isinstance(line_number, int) or line_number < 1:
        DLG.chi("Numero di riga non valido. Deve essere un intero maggiore di 0.")
        return

    # Costruisci il comando per aprire il file alla riga specifica
    # --reuse-window: riutilizza la finestra già aperta invece di crearne una nuova
    # --goto: apre il file alla riga:colonna specificata
    comando = f'"{editor_path}" --reuse-window --goto "{full_file_path}:{line_number}"'

    # Apri il file nell'editor
    try:
        subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except Exception as e:
        DLG.chi(f"Errore durante l'apertura del file con l'editor: {e}")


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
    with LeenoUtils.no_refresh_context():
        oDoc = LeenoUtils.getDocument()
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
    LeenoSheetUtils.adattaAltezzaRiga()


########################################################################


def MENU_Inser_SuperCapitolo():
    Ins_Categorie(0)

########################################################################

def MENU_Inser_Capitolo():
    Ins_Categorie(1)


########################################################################


def MENU_Inser_SottoCapitolo():
    Ins_Categorie(2)

########################################################################

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


# def ordina_col(ncol):
#     '''
#     ncol   { integer } : id colonna
#     ordina i dati secondo la colonna con id ncol
#     '''

#     ctx = LeenoUtils.getComponentContext()
#     desktop = LeenoUtils.getDesktop()
#     oFrame = desktop.getCurrentFrame()
#     dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
#         'com.sun.star.frame.DispatchHelper', ctx)
#     oProp = []
#     oProp0 = PropertyValue()
#     oProp0.Name = 'ByRows'
#     oProp0.Value = True
#     oProp1 = PropertyValue()
#     oProp1.Name = 'HasHeader'
#     oProp1.Value = False
#     oProp2 = PropertyValue()
#     oProp2.Name = 'CaseSensitive'
#     oProp2.Value = False
#     oProp3 = PropertyValue()
#     oProp3.Name = 'NaturalSort'
#     oProp3.Value = False
#     oProp4 = PropertyValue()
#     oProp4.Name = 'IncludeAttribs'
#     oProp4.Value = True
#     oProp5 = PropertyValue()
#     oProp5.Name = 'UserDefIndex'
#     oProp5.Value = 0
#     oProp6 = PropertyValue()
#     oProp6.Name = 'Col1'
#     oProp6.Value = ncol
#     oProp7 = PropertyValue()
#     oProp7.Name = 'Ascending1'
#     oProp7.Value = True
#     oProp.append(oProp0)
#     oProp.append(oProp1)
#     oProp.append(oProp2)
#     oProp.append(oProp3)
#     oProp.append(oProp4)
#     oProp.append(oProp5)
#     oProp.append(oProp6)
#     oProp.append(oProp7)
#     properties = tuple(oProp)
#     dispatchHelper.executeDispatch(oFrame, '.uno:DataSort', '', 0, properties)


def ordina_col(ncol):
    '''
    ncol { integer } : indice della colonna (0-based) su cui eseguire l'ordinamento.
    Ordina l'area selezionata o l'area dati corrente in base alla colonna ncol.
    '''
    ctx = LeenoUtils.getComponentContext()
    oFrame = LeenoUtils.getDesktop().getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)

    # Definiamo le proprietà in un dizionario per una gestione più semplice
    sort_settings = {
        'ByRows': True,
        'HasHeader': False,
        'CaseSensitive': False,
        'NaturalSort': False,
        'IncludeAttribs': True,
        'UserDefIndex': 0,
        'Col1': ncol + 1, # Nota: il Dispatcher spesso richiede l'indice 1-based per le colonne
        'Ascending1': True
    }

    # Trasformiamo il dizionario in una tupla di oggetti PropertyValue
    properties = []
    for key, value in sort_settings.items():
        prop = PropertyValue()
        prop.Name = key
        prop.Value = value
        properties.append(prop)

    # Esecuzione del comando di ordinamento
    dispatchHelper.executeDispatch(oFrame, '.uno:DataSort', '', 0, tuple(properties))

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

def _gotoCella(cellRef=0, IDrow=None):
    """
    Muove il cursore nella cella indicata.

    Parametri:
        cellRef {str|int} : indirizzo della cella (es. 'A1') oppure indice di colonna.
        IDrow   {int|None}: indice di riga (se si usano indici numerici)

    Esempi:
        _gotoCella('A1')
        _gotoCella(2, 5)
    """
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # Se viene passato un indirizzo testuale (es. 'B3')
    if isinstance(cellRef, str):
        try:
            oRange = oSheet.getCellRangeByName(cellRef)
            oDoc.CurrentController.select(oRange)
            oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
            return
        except Exception:
            raise ValueError(f"Indirizzo di cella non valido: '{cellRef}'")

    # Se invece vengono passati indici numerici (comportamento originale)
    if isinstance(cellRef, int) and isinstance(IDrow, int):
        try:
            oCell = oSheet.getCellByPosition(cellRef, IDrow)
            oDoc.CurrentController.select(oCell)
            oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
            return
        except Exception:
            raise ValueError(f"Coordinate non valide: col={cellRef}, row={IDrow}")

    raise TypeError("Usa _gotoCella('A1') oppure _gotoCella(col, row)")



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
    LeenoSheetUtils.adattaAltezzaRiga()

########################################################################
def MENU_voce_breve():
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
@LeenoUtils.no_refresh # evita il refresh automatico
def MENU_prefisso_VDS_():
    '''
    Duplica la voce di Elenco Prezzi corrente aggiunge il prefisso 'VDS_'
    e individuandola come Voce Della Sicurezza
    '''
    oDoc = LeenoUtils.getDocument()

    pref = "VDS_"
    # pref = "NP_"

    def vds_ep():
        oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
        lrow = LeggiPosizioneCorrente()[1]
        if pref in  oSheet.getCellByPosition(0, lrow).String:
            #  Dialogs.Info(Title = 'Infomazione', Text = 'Voce della sicurezza già esistente')
            LeenoUtils.DocumentRefresh(True)
            return
        oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, lrow, 9, lrow))
        comando('Copy')
        MENU_nuova_voce_scelta()
        paste_clip(pastevalue = False)
        oSheet.getCellRangeByName("A5").String = pref + oSheet.getCellRangeByName("A5").String
        oSheet.getCellRangeByName("A5").CellBackColor = COLORE_VERDE_SPUNTA

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
                oSheet.getCellByPosition(1, inizio + 1).CellBackColor = COLORE_VERDE_SPUNTA
                if oSheet.Name == 'CONTABILITA':
                    fine -= 1
                _gotoCella(2, fine - 1)

                if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
                    # voce = (num, art, desc, um, quantP, prezzo, importo, sic, mdo)
                    art = LeenoComputo.datiVoceComputo(oSheet, lrow)[1]
                    pesca_cod()
                    vds_ep()
                    pesca_cod()
                    if pref not in art:
                        LeenoComputo.cambia_articolo(oSheet, lrow, pref+art)
                    # DLG.chi(pref)
                    if pref in art:
                        LeenoComputo.cambia_articolo(oSheet, lrow, art.replace(pref, ''))


                elif oSheet.Name == 'Elenco Prezzi':
                    vds_ep()

                    ###
                lrow = LeggiPosizioneCorrente()[1]
                lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 1)
            except Exception:
                pass
        # numera_voci()
    except Exception:
        pass
    _gotoCella(1, lrow+1)

    LeenoUtils.DocumentRefresh(True)

########################################################################

@LeenoUtils.no_refresh # evita il refresh automatico
@with_undo # abilita la funzionalità di undo
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

    # LeenoUtils.DocumentRefresh(False)

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
        # LeenoUtils.DocumentRefresh(True)
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
        # LeenoUtils.DocumentRefresh(True)

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
            # LeenoUtils.DocumentRefresh(True)
            pass
    # LeenoUtils.DocumentRefresh(True)

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

@LeenoUtils.no_refresh # evita il refresh automatico
@with_undo # abilita la funzionalità di undo
def cancella_voci_non_usate():
    '''
    Cancella le voci di prezzo non utilizzate.
    '''
    chiudi_dialoghi()

    if Dialogs.YesNoDialog(
        IconType="question",
        Title='AVVISO!',
        Text='''Questo comando ripulisce l'Elenco Prezzi
dalle voci non utilizzate in nessuno degli altri elaborati.

La procedura potrebbe richiedere del tempo.

Vuoi procedere comunque?'''
    ) == 0:
        return

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    oRange = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    SRep = oRange.StartRow + 1
    ERep = oRange.EndRow

    lista_prezzi = []
    # prende l'elenco dal foglio Elenco Prezzi
    for n in range(SRep, ERep):
        lista_prezzi.append(oSheet.getCellByPosition(0, n).String)

    # case-insensitive
    lista_prezzi_norm = [s.lower() for s in lista_prezzi]

    # attiva la progressbar
    indicator = oDoc.getCurrentController().getStatusIndicator()

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

        except Exception:
            pass

    # normalizza anche questa lista
    lista_norm = [s.lower() for s in lista]

    indicator.start("Eliminazione delle voci in corso...", 5)
    indicator.setValue(2)

    # calcola differenza case-insensitive
    da_cancellare = set(lista_prezzi_norm).difference(set(lista_norm))

    # torna al foglio Elenco Prezzi
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet

    indicator.setValue(3)
    struttura_off()
    struttura_off()
    struttura_off()
    indicator.setValue(4)

    # cancella
    for n in reversed(range(SRep, ERep)):
        cell_0 = oSheet.getCellByPosition(0, n).String
        cell_1 = oSheet.getCellByPosition(1, n).String
        cell_4 = oSheet.getCellByPosition(4, n).String

        # confronto case-insensitive
        if cell_0.lower() in da_cancellare or (cell_0 == '' and cell_1 == '' and cell_4 == ''):
            oSheet.Rows.removeByIndex(n, 1)

    indicator.setValue(5)
    indicator.end()

    _gotoCella(0, 3)
    Dialogs.Info(
        Title='Ricerca conclusa',
        Text=f"Eliminate {len(da_cancellare)} voci dall'elenco prezzi."
    )


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
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    oRange = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    SR = oRange.StartRow + 1
    ER = oRange.EndRow

    # Definiamo il range completo delle righe da modificare
    oRows = oSheet.getCellRangeByPosition(0, SR, 0, ER - 1).Rows

    # Prendiamo la prima riga del range come riferimento per il toggle
    # Invece di una cella fissa (B4), usiamo la riga SR per coerenza
    riga_esempio = oSheet.getRows().getByIndex(SR)

    if not riga_esempio.OptimalHeight:
        # 1. Torna all'altezza automatica
        oRows.OptimalHeight = True
    else:
        # 2. Passa all'altezza fissa (3 righe)
        # DISATTIVIAMO l'automatismo prima di dare l'altezza
        oRows.OptimalHeight = False

        # Calcolo altezza
        altezza_base = oSheet.getCellRangeByName('B4').CharHeight * 64 / 3 * 2
        nr_descrizione = float(cfg.read('Generale', 'altezza_celle'))
        hriga = 100 + altezza_base * nr_descrizione

        # Applichiamo l'altezza numerica
        oRows.Height = hriga

########################################################################

@LeenoUtils.no_refresh
def scelta_viste():
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

    LeenoUtils.memorizza_posizione()

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
        try:
            oDoc.getSheets().getByName('S1')
        except:
            return
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
        if not oDoc.Sheets.hasByName('VARIANTE') and not oDoc.Sheets.hasByName('CONTABILITA'):
            oDialog1.getControl('CommandButton7').setEnable(0)

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
                    0, el, 1, el).CellBackColor = COLORE_COLONNE_RAFFRONTO

        oDoc.CurrentController.select(oSheet.getCellRangeByName('Z2'))
        comando('Copy')
        oDoc.CurrentController.select(
            oSheet.getCellRangeByPosition(25, 3, 25, ER+1))
        paste_format()

        _primaCella()
        oSheet.getCellRangeByPosition(11, 3, 13, ER+1).CellBackColor = COLORE_COLONNE_RAFFRONTO

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
    LeenoUtils.ripristina_posizione()


########################################################################
# @Debug.measure_time()
@LeenoUtils.no_refresh
def genera_sommario():
    '''
    Genera i sommari in Elenco Prezzi
    '''
    struttura_off()
    inizializza_elenco()


    oDoc = LeenoUtils.getDocument()

    sistema_aree()

    # with LeenoUtils.DocumentRefreshContext(False):
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
    # formule = tuple(formule)
    formule = tuple(tuple(riga) for riga in formule)
    oRange.setFormulaArray(formule)
    return

########################################################################
def MENU_riordina_ElencoPrezzi():
    riordina_ElencoPrezzi()
    Menu_adattaAltezzaRiga()

@LeenoUtils.no_refresh
def riordina_ElencoPrezzi():
    """
    Riordina l'Elenco Prezzi in ordine alfabetico dei codici di prezzo,
    ignorando l'eventuale prefisso 'VDS_'.
    """
    # with LeenoUtils.DocumentRefreshContext(False):

    oDoc = LeenoUtils.getDocument()
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

    def ordina_intervallo_ignorando_prefisso(sheet, start_row, end_row, start_col, end_col):
        """
        Ordina un intervallo di celle in base alla prima colonna,
        ignorando il prefisso 'VDS_' nei codici.
        """
        temp_col = end_col + 1
        # Crea colonna ausiliaria con codici senza 'VDS_'
        for r in range(start_row, end_row + 1):
            codice = sheet.getCellByPosition(start_col, r).String
            codice_senza_prefisso = codice[4:] if codice.startswith('VDS_') else codice
            sheet.getCellByPosition(temp_col, r).String = codice_senza_prefisso

        # Ordina in base alla colonna temporanea
        oRange = sheet.getCellRangeByPosition(start_col, start_row, temp_col, end_row)
        SheetUtils.simpleSortColumn(oRange, temp_col - start_col, True)

        # Cancella la colonna temporanea
        sheet.getColumns().removeByIndex(temp_col, 1)

    try:
        costo_elem_row = SheetUtils.uFindStringCol('ELENCO DEI COSTI ELEMENTARI', 1, oSheet)

        if costo_elem_row is None:
            ordina_intervallo_ignorando_prefisso(oSheet, start_row, end_row, start_col, end_col)
        else:
            ordina_intervallo_ignorando_prefisso(oSheet, start_row, costo_elem_row - 1, start_col, end_col)
            ordina_intervallo_ignorando_prefisso(oSheet, costo_elem_row + 1, end_row, start_col, end_col)

    except Exception as e:
        DLG.errore(e)

@LeenoUtils.no_refresh
def riordina_ElencoPrezzi_():
    """
    Riordina l'Elenco Prezzi in ordine alfabetico-numerico naturale dei codici di prezzo,
    ignorando i prefissi di 6 caratteri terminanti con '_' (incluso VDS_).
    """
    # with LeenoUtils.DocumentRefreshContext(False):
    oDoc = LeenoUtils.getDocument()
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
    end_col = 7  # Colonna H (0-indexed: A=0, B=1, ..., H=7)
    end_row = oRangeAddress.EndRow - 1
    if start_row == end_row:
        return

    def rimuovi_prefisso(codice):
        """
        Rimuove il prefisso se formato da esattamente 6 caratteri terminanti con '_'.
        Questo include VDS_, BAS25_, BAS23_, ecc.
        """
        if len(codice) >= 6 and codice[5] == '_':
            return codice[6:]
        return codice

    def chiave_ordinamento_naturale(testo):
        """
        Converte un testo in una chiave per ordinamento naturale.
        I numeri vengono convertiti in tuple (int, str) per ordinarli numericamente.
        """
        def converti(parte):
            if parte.isdigit():
                return (0, int(parte))
            else:
                return (1, parte.lower())

        return [converti(c) for c in re.split('([0-9]+)', testo)]

    def ordina_intervallo_ignorando_prefisso(sheet, start_row, end_row, start_col, end_col):
        """
        Ordina un intervallo di celle in base alla prima colonna,
        ignorando tutti i prefissi di 6 caratteri terminanti con '_'.
        """
        # Raccogli tutte le righe con i loro codici
        righe = []
        for r in range(start_row, end_row + 1):
            codice_originale = sheet.getCellByPosition(start_col, r).String
            codice_senza_prefisso = rimuovi_prefisso(codice_originale)
            chiave = chiave_ordinamento_naturale(codice_senza_prefisso)
            righe.append((chiave, r, codice_originale))

        # Ordina in base alla chiave naturale
        righe.sort(key=lambda x: x[0])

        # Crea una mappa di destinazione -> sorgente
        mappa_riordino = {}
        for nuovo_idx, (_, vecchio_idx, _) in enumerate(righe):
            mappa_riordino[start_row + nuovo_idx] = vecchio_idx

        # Copia i dati riordinati in un buffer temporaneo
        buffer = []
        for dest_row in range(start_row, end_row + 1):
            source_row = mappa_riordino[dest_row]
            riga_dati = []
            for col in range(start_col, end_col + 1):
                cell = sheet.getCellByPosition(col, source_row)
                # Salva tutti i dettagli della cella
                from com.sun.star.table.CellContentType import TEXT, VALUE, FORMULA, EMPTY
                cell_type = cell.Type
                riga_dati.append({
                    'string': cell.String if cell_type == TEXT or cell_type == EMPTY else '',
                    'value': cell.Value if cell_type == VALUE or cell_type == FORMULA else 0,
                    'formula': cell.Formula if cell_type == FORMULA else '',
                    'type': cell_type
                })
            buffer.append(riga_dati)

        # Scrivi il buffer nelle celle preservando i tipi
        for idx, riga_dati in enumerate(buffer):
            dest_row = start_row + idx
            for col_offset, cell_data in enumerate(riga_dati):
                dest_col = start_col + col_offset
                cell = sheet.getCellByPosition(dest_col, dest_row)
                from com.sun.star.table.CellContentType import TEXT, VALUE, FORMULA, EMPTY

                if cell_data['type'] == FORMULA:
                    cell.setFormula(cell_data['formula'])
                elif cell_data['type'] == VALUE:
                    cell.setValue(cell_data['value'])
                elif cell_data['type'] == TEXT:
                    cell.setString(cell_data['string'])
                else:  # EMPTY
                    cell.setString('')

    try:
        costo_elem_row = SheetUtils.uFindStringCol('ELENCO DEI COSTI ELEMENTARI', 1, oSheet)
        if costo_elem_row is None:
            ordina_intervallo_ignorando_prefisso(oSheet, start_row, end_row, start_col, end_col)
        else:
            ordina_intervallo_ignorando_prefisso(oSheet, start_row, costo_elem_row - 1, start_col, end_col)
            ordina_intervallo_ignorando_prefisso(oSheet, costo_elem_row + 1, end_row, start_col, end_col)
    except Exception as e:
        DLG.errore(e)

########################################################################

@LeenoUtils.no_refresh
def MENU_doppioni():
    # Inizializza la progress bar
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    indicator = oDoc.getCurrentController().getStatusIndicator()
    if indicator:
        indicator.start("Elaborazione in corso...", 100)  # 100 = max progresso
    LeenoUtils.memorizza_posizione()
    try:
        # with LeenoUtils.DocumentRefreshContext(False):
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
        sistema_stili()

        # Fase 3: Genera sommario (60%)
        if indicator:
            indicator.Text = "Generazione sommario..."
            indicator.Value = 60

        # oSheet.getCellRangeByName("X1").String = ''
        LeenoSheetUtils.setLarghezzaColonne(oSheet)
        genera_sommario()

        # Fase 4: Riordina EP (80%)
        if indicator:
            indicator.Text = "Riordino Elenco Prezzi..."
            indicator.Value = 80
        riordina_ElencoPrezzi()

        # Fase 5: Sistema stili (100%)
        if indicator:
            indicator.Text = "Applicazione stili..."
            indicator.Value = 100

        LeenoSheetUtils.adattaAltezzaRiga()

    finally:
        if indicator:
            indicator.end()  # Chiude la progress bar
    LeenoUtils.ripristina_posizione()


def EliminaVociDoppieElencoPrezzi():
    """
    Rimuove dall'elenco prezzi:
    1. Righe con formule nella colonna B
    2. Voci duplicate (stessa chiave fino a MAX_COMPARE_COLS, esclusa colonna 1)
       - Per duplicati: mantiene righe con markup (col5) o prima riga
    """
    LeenoUtils.memorizza_posizione()
    # --- CONFIGURAZIONE ---
    MARKUP_COL = 5             # Colonna markup/note
    FORMULA_CHECK_COL = 1      # Colonna B (indice 1)
    MAX_COMPARE_COLS = 5       # Colonne per confronto duplicati
    TOTAL_COLS = 14            # Totale colonne elenco

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    named_range = oDoc.NamedRanges.getByName('elenco_prezzi').ReferredCells.RangeAddress
    start_row, end_row = named_range.StartRow + 1, named_range.EndRow - 1

    if end_row <= start_row:
        return

    # Lettura dati
    data_range = oSheet.getCellRangeByPosition(0, start_row, TOTAL_COLS - 1, end_row)
    full_data = data_range.getDataArray()

    # Filtraggio: escludi righe con formule nella colonna B
    filtered_data = []
    for idx, row in enumerate(full_data):
        cell = oSheet.getCellByPosition(FORMULA_CHECK_COL, start_row + idx)
        # Verifica se la cella contiene una formula
        if cell.getType().value == 'FORMULA':
            continue
        filtered_data.append(row)

    # Eliminazione duplicati
    groups = OrderedDict()
    for row in filtered_data:
        key = tuple(row[i] for i in range(MAX_COMPARE_COLS) if i != 1)
        if key not in groups:
            groups[key] = []
        groups[key].append(row)

    clean_data = []
    for rows in groups.values():
        with_markup = [r for r in rows if r[MARKUP_COL] not in ('', None)]
        if with_markup:
            clean_data.append(with_markup[0])
        else:
            clean_data.append(rows[0])

    # Scrittura risultati
    if len(clean_data) != len(full_data):
        oSheet.getRows().removeByIndex(start_row, end_row - start_row + 1)
        oSheet.getRows().insertByIndex(start_row, len(clean_data))
        output_range = oSheet.getCellRangeByPosition(
            0, start_row,
            TOTAL_COLS - 1,
            start_row + len(clean_data) - 1
        )
        output_range.setDataArray(clean_data)

    LeenoUtils.ripristina_posizione()

########################################################################

def XPWE_export_run():
    '''
    Visualizza il menù export/import XPWE
    '''
    oDoc = LeenoUtils.getDocument()
    psm = LeenoUtils.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    Dialog_XPWE = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.Dialog_XPWE?language=Basic&location=application"
    )
    oSheet = oDoc.CurrentController.ActiveSheet
    # Configurazione iniziale dei controlli del dialogo
    for el in ("COMPUTO", "VARIANTE", "CONTABILITA", "Elenco Prezzi", 'Analisi di Prezzo'):
        try:
            importo = oDoc.getSheets().getByName(el).getCellRangeByName('A2').String
            # Usa elif invece di multipli if
            if el == 'COMPUTO':
                Dialog_XPWE.getControl(el).Label = 'Computo:'.ljust(13) + importo.rjust(15)
            elif el == 'VARIANTE':
                Dialog_XPWE.getControl(el).Label = 'Variante:'.ljust(13) + importo.rjust(15)
            elif el == 'CONTABILITA':
                Dialog_XPWE.getControl(el).Label = 'Contabilità:'.ljust(13) + importo.rjust(15)
            elif el == 'Elenco Prezzi':
                Dialog_XPWE.getControl(el).Label = 'Elenco Prezzi'
            elif el == 'Analisi di Prezzo':
                Dialog_XPWE.getControl(el).Label = 'Analisi di Prezzo'
            Dialog_XPWE.getControl(el).Enable = True
        except Exception:
            Dialog_XPWE.getControl(el).Enable = False
    Dialog_XPWE.Title = 'Esportazione XPWE'
    # Seleziona il foglio corrente se disponibile
    try:
        Dialog_XPWE.getControl(oSheet.Name).State = True
    except Exception:
        pass
    lista = []
    analisi = False
    # Esegue il dialogo e gestisce la risposta
    if Dialog_XPWE.execute() == 1:
        for el in ("Elenco Prezzi", "COMPUTO", "VARIANTE", "CONTABILITA", "Analisi di Prezzo"):
            if Dialog_XPWE.getControl(el).State == 1:
                if el == "Analisi di Prezzo":
                    analisi = True
                    # Non aggiungere "Analisi di Prezzo" alla lista, serve solo per il flag
                else:
                    lista.append(el)
    else:
        # L'utente ha annullato il dialogo
        return
    # Richiede il file di output
    out_file = Dialogs.FileSelect('Salva con nome...', '*.xpwe', 1)
    if out_file == '':
        return
    # Se l'utente ha selezionato "Analisi di Prezzo", assicurati che "Elenco Prezzi" sia nella lista
    if analisi and "Elenco Prezzi" not in lista:
        lista.insert(0, "Elenco Prezzi")
    # Esporta i dati selezionati
    for el in lista:
        XPWE_out(el, out_file, analisi)


# Scrive un file.
@LeenoUtils.no_refresh
def XPWE_out(elaborato, out_file, analisi=False):
    '''
    esporta il documento in formato XPWE
    elaborato { string } : nome del foglio da esportare
    out_file  { string } : nome base del file
    analisi   { bool }   : se True esporta anche l'analisi di prezzo
    il nome file risulterà out_file-elaborato.xpwe
    '''
    XPWE_out_run(elaborato, out_file, analisi=analisi)


def XPWE_out_run(elaborato, out_file, analisi=False):
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
    numera_voci()
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
        # Se analisi=True, raccogli le voci che hanno analisi per esportarle dopo
        ha_analisi = False
        if(oSheet.getCellByPosition(1, n).Type.value == 'FORMULA' and
           oSheet.getCellByPosition(2, n).Type.value == 'FORMULA' and
           analisi==True):
            lista_AP.append(oSheet.getCellByPosition(0, n).String)
            ha_analisi = True

        # ESPORTA TUTTE LE VOCI, ma salta quelle con analisi se analisi=True
        # (saranno esportate dopo con i dettagli completi)
        if ha_analisi:
            continue

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

    # Analisi di prezzo - ESPORTATE SOLO SE analisi=True
    indicator.Value = 5
    if analisi and len(lista_AP) != 0:
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
        Text=f'Esportazione in formato XPWE eseguita con successo sul file:\n\n {LeenoUtils.wrap_path(out_file)}'
'\n\n----\n'
'Il formato XPWE è un formato XML di interscambio per Primus di ACCA. '
'Prima di utilizzare questo file in Primus, assicurarsi che le percentuali'
'di Spese Generali e Utile d\'Impresa siano impostate correttamente,'
'in modo da garantire l\'esatta elaborazione dei dati.')

        # Apri la cartella contenente il file ZIP
        try:
            apri = LeenoUtils.createUnoService("com.sun.star.system.SystemShellExecute")
            zip_url = uno.systemPathToFileUrl(str(out_file.parent))
            apri.execute(zip_url, "", 0)
        except Exception:
            pass

    except IOError:
        Dialogs.Exclamation(Title = 'E R R O R E !',
            Text='''Esportazione non eseguita!
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
    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        if oSheet.Name != 'Analisi di Prezzo':
            return
        # oDoc.enableAutomaticCalculation(False)  # blocco il calcolo automatico
        sStRange = Circoscrive_Analisi(LeggiPosizioneCorrente()[1])
        riga = sStRange.RangeAddress.StartRow + 2

        codice = oSheet.getCellByPosition(0, riga - 1).String

        oSheet = oDoc.Sheets.getByName('Elenco Prezzi')
        oDoc.CurrentController.setActiveSheet(oSheet)

        target_row = 4
        oSheet.getRows().insertByIndex(target_row, 1)

        # Imposta codice e formule
        oSheet.getCellByPosition(0, target_row).String = codice

        formule = {
            1: f"=$'Analisi di Prezzo'.B{riga}",
            2: f"=$'Analisi di Prezzo'.C{riga}",
            3: f"=$'Analisi di Prezzo'.K{riga}",
            4: f"=$'Analisi di Prezzo'.G{riga}",
            5: f"=$'Analisi di Prezzo'.I{riga}",
            6: f"=$'Analisi di Prezzo'.J{riga}",
            11: '=LET(s; SUMIF(AA; A5; BB); IF(s <> 0; s; "--"))',
            12: '=LET(s; SUMIF(varAA; A5; varBB); IFERROR(IF(s; s; "--"); "--"))',
            13: '=LET(s; SUMIF(GG; A5; G1G1); IFERROR(IF(s; s; "--"); "--"))'
        }

        for col, formula in formule.items():
            oSheet.getCellByPosition(col, target_row).Formula = formula
        _gotoCella("A5")
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
        target_row = 5  # Inizia dalla riga 4 in Elenco Prezzi

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
                    '',
                    '',
                    '',
                    '',
                    '',
                    f'=LET(s; SUMIF(AA; A{target_row}; BB); IF(s <> 0; s; "--"))',
                    f'=LET(s; SUMIF(varAA; A{target_row}; varBB); IFERROR(IF(s; s; "--"); "--"))',
                    f'=LET(s; SUMIF(GG; A{target_row}; G1G1); IFERROR(IF(s; s; "--"); "--"))'
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
    LeenoSheetUtils.adattaAltezzaRiga(dest_sheet)

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
                    _gotoCella(1, inizio + 1)
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
                    oProp.Value = COLORE_GRIGIO_INATTIVA
                    properties = (oProp, )
                    dispatchHelper.executeDispatch(oFrame, '.uno:BackgroundColor', '', 0, properties)
                    _gotoCella(1, fine + 3)
                    ###
                lrow = LeggiPosizioneCorrente()[1]
                lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 1)
            except Exception:
                pass
        # numera_voci()
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

            numera_voci()
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

@with_undo() # abilita l'undo per l'intera funzione
@LeenoUtils.no_refresh # evita il flickering del documento durante l'eliminazione
@LeenoUtils.preserva_posizione(step=0)
def MENU_elimina_voce():
    # LeenoUtils.DocumentRefresh(False)
    elimina_voce()
    # LeenoUtils.DocumentRefresh(True)
    oDoc = LeenoUtils.getDocument()
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))

# def elimina_voce(lrow=None):
#     '''
#     @@@ MODIFICA IN CORSO CON 'LeenoSheetUtils.eliminaVoce'
#     Elimina una voce in COMPUTO, VARIANTE, CONTABILITA o Analisi di Prezzo
#     lrow { long }  : numero riga
#     msg  { bit }   : 1 chiedi conferma con messaggio
#                      0 esegui senza conferma
#     '''
#     oDoc = LeenoUtils.getDocument()
#     oSheet = oDoc.CurrentController.ActiveSheet

#     if oSheet.Name == 'Elenco Prezzi':
#         Dialogs.Info(Title = 'Info', Text="""Per eliminare una o più voci dall'Elenco Prezzi
# devi selezionarle ed utilizzare il comando 'Elimina righe' di Calc.""")
#         return

#     if oSheet.Name not in ('COMPUTO', 'CONTABILITA', 'VARIANTE', 'Analisi di Prezzo'):
#         return

#     try:
#         SR = seleziona_voce()[0]
#     except:
#         return
#     ER = seleziona_voce()[1]

#     oDoc.CurrentController.select(oSheet.getCellRangeByPosition(
#         0, SR, 250, ER))
#     if '$C$' in oSheet.getCellByPosition(9, ER).queryDependents(False).AbsoluteName:

#         _gotoCella(9, ER)
#         comando ('ClearArrowDependents')
#         comando ('ShowDependents')
#         oDoc.CurrentController.select(oSheet.getCellRangeByPosition(
#             0, SR, 250, ER))
#         Dialogs.Exclamation(Title = 'ATTENZIONE!',
#             Text="Da questa voce dipende almeno un Vedi Voce.\n\n"
#                     "Cancellazione interrotta per sicurezza.")

#         return

#     oSheet.getRows().removeByIndex(SR, ER - SR + 1)
#     if oSheet.Name != 'Analisi di Prezzo':
#         numera_voci()
#     else:
#         _gotoCella(0, SR+2)

@with_undo("Elimina Voci Selezionate")
def elimina_voce(lrow=None):
    '''
    Elimina una o più voci selezionate in COMPUTO, VARIANTE, CONTABILITA o Analisi di Prezzo.
    Gestisce selezioni estese identificando i blocchi completi.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # --- 1. CONTROLLI PRELIMINARI ---
    if oSheet.Name == 'Elenco Prezzi':
        Dialogs.Info(Title='Info', Text="""Per eliminare una o più voci dall'Elenco Prezzi
devi selezionarle ed utilizzare il comando 'Elimina righe' di Calc.""")
        return

    if oSheet.Name not in ('COMPUTO', 'CONTABILITA', 'VARIANTE', 'Analisi di Prezzo'):
        return

    # --- 2. IDENTIFICAZIONE DEL RANGE DI VOCI ---
    oSel = oDoc.CurrentController.getSelection()
    if not oSel.supportsService("com.sun.star.table.CellRange"):
        return

    try:
        # Identifica l'inizio della prima voce e la fine dell'ultima voce nella selezione
        SR = LeenoComputo.circoscriveVoceComputo(oSheet, oSel.getRangeAddress().StartRow).RangeAddress.StartRow
        ER = LeenoComputo.circoscriveVoceComputo(oSheet, oSel.getRangeAddress().EndRow).RangeAddress.EndRow
    except:
        return

    # Seleziona visivamente l'area che sta per essere eliminata
    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, SR, 250, ER))

    # --- 3. CONTROLLO DIPENDENZE (VEDI VOCE) ---
    # Controlliamo se qualche cella nel range selezionato ha dipendenti critici
    # Verifichiamo la colonna dei totali (J, indice 9) per l'intero intervallo
    area_totali = oSheet.getCellRangeByPosition(9, SR, 9, ER)
    try:
        if '$C$' in area_totali.queryDependents(False).AbsoluteName:
            # Mostra graficamente le frecce delle dipendenze per aiutare l'utente
            _gotoCella(9, ER)
            comando('ClearArrowDependents')
            comando('ShowDependents')

            # Ri-seleziona l'area
            oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, SR, 250, ER))

            Dialogs.Exclamation(Title='ATTENZIONE!',
                Text="In questo blocco sono presenti voci da cui dipende almeno un 'Vedi Voce'.\n\n"
                     "Cancellazione interrotta per sicurezza.")
            return
    except:
        # Se queryDependents fallisce, significa che non ci sono dipendenti
        pass

    # --- 4. ESECUZIONE ELIMINAZIONE ---
    # Chiediamo conferma solo se la selezione coinvolge più di una riga o per sicurezza
    # (Opzionale: puoi aggiungere un Dialogs.Ask qui)

    num_rows = ER - SR + 1
    oSheet.getRows().removeByIndex(SR, num_rows)

    # --- 5. RICALCOLO E POSIZIONAMENTO ---
    if oSheet.Name != 'Analisi di Prezzo':
        numera_voci()
    else:
        # In Analisi di Prezzo, ci posizioniamo dove è avvenuta l'eliminazione
        _gotoCella(0, max(0, SR - 1))

    # Pulizia selezione
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
########################################################################

@with_undo() # abilita l'undo per l'intera funzione
@LeenoUtils.no_refresh # evita il flickering del documento durante l'eliminazione
def MENU_elimina_righe():
    """
    Elimina le righe selezionate (contigue o no) ottimizzando le operazioni UNO.
    Raggruppa i blocchi contigui e li elimina con una sola removeByIndex per blocco.
    """
    # with LeenoUtils.DocumentRefreshContext(False):

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    if oSheet.Name == 'Elenco Prezzi':
        Dialogs.Exclamation(Title = 'Avviso.',
        Text="'Per eliminare voci dall'Elenco Prezzi usa il comando nativo 'Elimina righe'.")
        return

    if oSheet.Name not in ('COMPUTO', 'CONTABILITA', 'VARIANTE', 'Analisi di Prezzo'):
        return

    # --- 1) ACQUISIZIONE SELEZIONE ---
    try:
        ranges = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        ranges = [oDoc.getCurrentSelection().getRangeAddress()]

    righe = set()
    for r in ranges:
        righe.update(range(r.StartRow, r.EndRow + 1))

    # Ordine decrescente per non alterare indici
    righe = sorted(righe, reverse=True)

    if not righe:
        return

    # --- 2) COSTRUZIONE BLOCCHI CONTIGUI ---
    blocchi = []
    start = end = None

    for r in righe:
        if start is None:
            start = end = r
        elif r == end - 1:
            end = r
        else:
            blocchi.append((end, start))
            start = end = r
    if start is not None:
        blocchi.append((end, start))  # ultimo blocco

    rigen = False  # serve per rigenera_parziali

    # Prepara struct per eventuale copia Data_bianca
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr.StartColumn = 1
    oCellRangeAddr.EndColumn = 1

    # --- 3) ELABORAZIONE BLOCCHI ---
    for start, end in blocchi:

        righe_valide = []

        # filtriamo solo le righe che rispettano la logica originale
        for y in range(start, end + 1):

            stile2 = oSheet.getCellByPosition(2, y).CellStyle
            testo8 = oSheet.getCellByPosition(8, y).String

            # Condizioni ORIGINALE (non modificate)
            if stile2 not in (
                'An-lavoraz-generica',
                'An-lavoraz-Cod-sx',
                'comp 1-a',
                'comp 1-a ROSSO',
                'comp sotto centro',
                'EP-mezzo',
                'Livello-0-scritta mini',
                'Livello-1-scritta mini',
                'livello2_'
            ) or \
            'Somma positivi e negativi [' in testo8 or \
            'SOMMANO' in testo8:

                # riga NON eliminabile → ignorata
                continue

            # OK, riga eliminabile
            righe_valide.append(y)

            # check rigenerazione
            if stile2 == 'comp sotto centro':
                rigen = True

            # caso speciale Data_bianca
            if oSheet.getCellByPosition(1, y).CellStyle == 'Data_bianca':
                oCellAddress = oSheet.getCellByPosition(1, y + 1).getCellAddress()
                oCellRangeAddr.StartRow = y
                oCellRangeAddr.EndRow = y
                oSheet.copyRange(oCellAddress, oCellRangeAddr)

            # rinumerazione capitoli (solo se necessario)
            if stile2 in ('Livello-0-scritta mini', 'Livello-1-scritta mini', 'livello2_'):
                Rinumera_TUTTI_Capitoli2(oSheet)

        # Se nel blocco ci sono righe valide → rimuoviamo tutto in una volta
        if righe_valide:
            primo = min(righe_valide)
            ultimo = max(righe_valide)
            count = ultimo - primo + 1
            oSheet.getRows().removeByIndex(primo, count)

    # --- 4) RIGENERA PARZIALI SE NECESSARIO ---
    if rigen:
        rigenera_parziali(False)
    Rinumera_TUTTI_Capitoli2(oSheet)

    # --- 5) Deseleziona tutto ---
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")
    )


########################################################################

def copia_riga_computo(lrow, num_righe=1):
    """
    Inserisce una o più righe di misurazione nel computo.
    """
    # Validazione preventiva
    if lrow is None or lrow < 0:
        return

    if num_righe is None or num_righe < 1:
        return

    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet

        # Controllo stile
        try:
            stile = oSheet.getCellByPosition(1, lrow).CellStyle
        except Exception:
            return  # riga inesistente

        if stile not in ('comp Art-EP', 'comp Art-EP_R', 'Comp-Bianche in mezzo'):
            return

        # Inserimento blocco di righe
        lrow_ins = lrow + 1

        # Controllo che lrow_ins sia nel range
        rowcount = oSheet.Rows.getCount()
        if lrow_ins < 0:
            return
        if lrow_ins > rowcount:
            lrow_ins = rowcount  # si inserisce in fondo

        # === Punto che generava l'errore ===
        oSheet.getRows().insertByIndex(lrow_ins, num_righe)

        # Prepara intervalli da aggiornare in blocco
        for i in range(num_righe):
            r = lrow_ins + i
            oSheet.getCellRangeByPosition(5, r, 7, r).CellStyle = 'comp 1-a'
            oSheet.getCellByPosition(0, r).CellStyle = 'comp 10 s'
            oSheet.getCellByPosition(1, r).CellStyle = 'Comp-Bianche in mezzo'
            oSheet.getCellByPosition(2, r).CellStyle = 'comp 1-a'
            oSheet.getCellRangeByPosition(3, r, 4, r).CellStyle = 'Comp-Bianche in mezzo bordate_R'
            oSheet.getCellByPosition(5, r).CellStyle = 'comp 1-a PU'
            oSheet.getCellByPosition(6, r).CellStyle = 'comp 1-a LUNG'
            oSheet.getCellByPosition(7, r).CellStyle = 'comp 1-a LARG'
            oSheet.getCellByPosition(8, r).CellStyle = 'comp 1-a peso'
            oSheet.getCellByPosition(9, r).CellStyle = 'Blu'
            oSheet.getCellByPosition(9, r).Formula = (
                f'=IF(PRODUCT(E{r+1}:I{r+1})=0;"";PRODUCT(E{r+1}:I{r+1}))'
            )

        _gotoCella(2, lrow_ins)
        return lrow_ins + num_righe - 1


def copia_riga_contab(lrow, num_righe=1):
    """
    Inserisce un blocco di righe di misurazione in contabilità,
    copiando il template in un'unica operazione.
    """
    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        oSheetto = oDoc.getSheets().getByName('S5')

        # template: riga 25 di S5 (0-based)
        template_range = oSheetto.getCellRangeByPosition(0, 24, 42, 24)
        stile = oSheet.getCellByPosition(1, lrow).CellStyle

        # controllo se già presente riga sotto
        if oSheet.getCellByPosition(1, lrow + 1).CellStyle == 'comp sotto Bianche_R':
            return

        # controllo stile valido
        if stile not in ('comp Art-EP_R', 'Data_bianca', 'Comp-Bianche in mezzo_R'):
            return

        # sblocca foglio se protetto
        if oSheet.isProtected():
            oSheet.unprotect("password")

        blocca_data = oSheet.getCellByPosition(1, lrow + 1).CellStyle == 'Data_bianca'

        # inserisci blocco di righe
        oSheet.getRows().insertByIndex(lrow + 1, num_righe)

        # # costruisci range di destinazione
        # target_range = oSheet.getCellRangeByPosition(
        #     0, lrow + 1, 42, lrow + num_righe
        # )

        # copia template su tutto il blocco in un'unica chiamata
        # ripete il contenuto del template su ogni riga
        for i in range(num_righe):
            cell_addr = oSheet.getCellByPosition(0, lrow + 1 + i).getCellAddress()
            oSheet.copyRange(cell_addr, template_range.getRangeAddress())

        # gestione stile speciale 'comp Art-EP_R'
        # if stile == 'comp Art-EP_R':
        #     for i in range(num_righe):
        #         cell = oSheet.getCellByPosition(1, lrow + 1 + i)
        #         cell.String = ""
        #         cell.CellStyle = 'Comp-Bianche in mezzo_R'
        # else:
        #     for i in range(num_righe):
        #         oSheet.getCellByPosition(1, lrow + 1 + i).CellStyle = 'Comp-Bianche in mezzo_R'

        for i in range(num_righe):
            cell = oSheet.getCellByPosition(1, lrow + 1 + i)
            if stile == 'comp Art-EP_R':
                cell.String = ""
            cell.CellStyle = 'Comp-Bianche in mezzo_R'

        if blocca_data:
            oSheet.getCellByPosition(1, lrow + 1).CellStyle = 'Data_bianca'
            oSheet.getCellByPosition(1, lrow + 1).Value = oSheet.getCellByPosition(1, lrow + 2).Value
            oSheet.getCellByPosition(1, lrow + 2).String = ""
            oSheet.getCellByPosition(1, lrow + 2).CellStyle = 'Comp-Bianche in mezzo_R'

        # sposta selezione sull'ultima riga
        _gotoCella(2, lrow + num_righe)

########################################################################

def copia_riga_analisi(lrow, num_righe=1):
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoAnalysis.copiaRigaAnalisi'
    Inserisce una nuova riga di misurazione in analisi di prezzo
    '''
    with LeenoUtils.DocumentRefreshContext(False):
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


@with_undo()  # abilita l'undo per l'intera funzione
def MENU_Copia_riga_Ent():
    '''
    @@ DA DOCUMENTARE
    '''
    Copia_riga_Ent()

# @LeenoUtils.no_refresh
def Copia_riga_Ent(num_righe=None):
    """
    Aggiunge righe di misurazione.
    Se num_righe non è specificato, usa il numero di righe attualmente selezionate.
    Args:
        num_righe (int, optional): Numero di righe da inserire.
                                   Se None, usa la selezione corrente.
    Returns:
        int: Indice dell'ultima riga inserita, o None se nessuna operazione
    """
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nome_sheet = oSheet.Name
    # Determina il numero di righe da inserire dalla selezione
    if num_righe is None:
        try:
            selection = oDoc.CurrentSelection.getRangeAddress()
            sRow = selection.StartRow
            eRow = selection.EndRow + 1
            num_righe = eRow - sRow
        except AttributeError:
            num_righe = 1

    # Validazione parametro
    if not isinstance(num_righe, int) or num_righe < 1:
        DLG.chi(f"Numero righe non valido: {num_righe}. Uso valore 1.")
        num_righe = 1

    # Se le colonne di misura sono nascoste, vengono visualizzate
    col_misura = oSheet.getColumns()
    if not col_misura.getByIndex(5).IsVisible:
        n = SheetUtils.getLastUsedRow(oSheet)
        # Pulisce le formule nelle celle con stile specifico
        for el in range(4, n + 1):
            cell = oSheet.getCellByPosition(2, el)
            if cell.CellStyle == "comp sotto centro":
                cell.Formula = ''
        # Rende visibili le colonne di misura (5, 6, 7)
        for el in range(5, 8):
            col_misura.getByIndex(el).IsVisible = True

    lrow = LeggiPosizioneCorrente()[1]
    lrow_originale = lrow  # Salva il valore originale
    dettaglio_attivo = cfg.read('Generale', 'dettaglio') == '1'

    # Dizionario delle azioni per tipo di foglio
    azioni = {
        'COMPUTO': copia_riga_computo,
        'VARIANTE': copia_riga_computo,
        'CONTABILITA': copia_riga_contab,
        'Analisi di Prezzo': copia_riga_analisi,
    }

    # Esegue l'azione appropriata in base al tipo di foglio
    if nome_sheet in azioni:
        if dettaglio_attivo and nome_sheet in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
            dettaglio_misura_rigo()

        # Chiamata alla funzione con gestione blocco
        result = azioni[nome_sheet](lrow, num_righe)

        # Se la funzione restituisce None, usa la posizione corrente + num_righe
        if result is None:
            # Calcola la nuova posizione basandoti su lrow_originale + num_righe
            lrow = lrow_originale + num_righe - 1
        elif isinstance(result, int):
            lrow = result
        else:
            DLG.chi(f"ERRORE: La funzione {azioni[nome_sheet].__name__} ha restituito tipo non valido: {type(result)}")
            return None

        # Aggiorna altezza ultima riga inserita
        oSheet.getCellRangeByPosition(0, lrow, 48, lrow).Rows.OptimalHeight = True

    elif nome_sheet == "Elenco Prezzi":
        MENU_nuova_voce_scelta()
        return None
    else:
        DLG.chi(f"Tipo di foglio '{nome_sheet}' non supportato")
        return None

    return lrow

########################################################################
def count_clipboard_lines():
    ctx = LeenoUtils.getComponentContext()
    smgr = ctx.getServiceManager()
    clip = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)

    # Ottiene il contenuto della clipboard
    transferable = clip.getContents()

    # Cerca il formato text/plain
    flavors = transferable.getTransferDataFlavors()

    text = None
    for flavor in flavors:
        if "text/plain" in flavor.MimeType:
            data = transferable.getTransferData(flavor)
            # data è già uno str in UTF-16 → usalo direttamente
            text = str(data)
            break

    if text is None:
        return 0

    # Conta le righe
    num_lines = len(text.splitlines())

    # Restituisce il valore
    return num_lines


@with_undo()
def paste_smart():
    """
    Incolla contenuti multi-riga creando automaticamente il numero
    necessario di righe di misurazione, rigenerando i parziali.
    """
    with LeenoUtils.no_refresh_context():
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        lrow1 = LeggiPosizioneCorrente()[1]

        cell_style = oSheet.getCellByPosition(2, lrow1).CellStyle
        cell_8_str = oSheet.getCellByPosition(8, lrow1).String

        # STOP se almeno una delle due condizioni NON è soddisfatta
        # Definisci condizioni chiare con nomi significativi


        if any(["comp 1" in cell_style, "Comp-Bianche in mezzo Descr" in cell_style, "Parziale" in cell_8_str]):

            # with LeenoUtils.DocumentRefreshContext(False):
            LeenoUtils.memorizza_posizione(step=0)
            nr = count_clipboard_lines()
            if nr == 0:
                DLG.chi("Nessun dato di testo rilevato negli appunti.")
                return
            Copia_riga_Ent()   # prima riga sempre
            if nr > 1:
                Copia_riga_Ent(nr - 1)   # aggiungi le restanti
            LeenoUtils.ripristina_posizione()
            lrow = LeggiPosizioneCorrente()[1]
            _gotoCella(2, lrow + 1)
            paste_clip()
            rigenera_parziali(False)
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    # oDoc.CurrentController.select(oSheet.getCellRangeByPosition(1, 0, 1, fine))

    return


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
# @Debug.measure_time(show_popup=True) # Misura il tempo di esecuzione della funzione
@no_undo
@LeenoUtils.no_refresh # evita il flickering del documento durante l'operazione
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
    if oSheet.isProtected():
        Dialogs.NotifyDialog(IconType="warning",Title='AVVISO!',
        Text=f"Il foglio {oSheet.Name} è protetto.\n\nPer poter procedere devi prima sbloccarlo.")
        return
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
                DLG.chi("Controllo atti contabili non eseguito: verificate di non avere atti registrati prima di procedere.")
                # if DLG.DlgSiNo(
                #         "Risulta già registrato un SAL. VUOI PROCEDERE COMUQUE?",
                #         'ATTENZIONE!') == 3:
                #     return
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
                if oSheet.isProtected():
                    GotoSheet('Elenco Prezzi')
                    Dialogs.NotifyDialog(IconType="warning",Title='AVVISO!',
                    Text=f"Il foglio di destinazione {oSheet.Name} è protetto.\n\nPer poter procedere devi prima sbloccarlo.")
                    return
                oSheet.getCellByPosition(1, partenza[1]).String = codice
                _gotoCella(2, partenza[1] + 1)
        except NameError:
            return


########################################################################
@no_undo
@LeenoUtils.no_refresh # evita il flickering del documento durante l'operazione
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
        oSheet.getCellByPosition(2, sopra - 1).CellBackColor = COLORE_VERDE_SPUNTA
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
        oDest.getCellByPosition(2, partenza[1]).CellBackColor = COLORE_VERDE_SPUNTA
        rigenera_voce(partenza[1])

        _gotoCella(2, partenza[1] + 1)

    LeenoUtils.DocumentRefresh(True)

    LeenoSheetUtils.adattaAltezzaRiga(oDoc.CurrentController.ActiveSheet)

########################################################################
def MENU_inverti_segno():
    inverti_segno()

@with_undo('Inverti segno delle misure')
def inverti_segno():
    '''
    Inverte il segno delle formule di quantità e gestisce lo stile ROSSO tramite suffisso.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    selection = oDoc.getCurrentSelection()

    # Estrazione robusta degli indirizzi (gestisce selezioni singole o multiple)
    try:
        ranges = selection.getRangeAddresses()
    except AttributeError:
        ranges = [selection.getRangeAddress()]

    # Creazione lista univoca delle righe selezionate
    lista_righe = set()
    for r in ranges:
        for row in range(r.StartRow, r.EndRow + 1):
            lista_righe.add(row)

    # Elaborazione per fogli COMPUTO e VARIANTE
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        for lrow in sorted(list(lista_righe)):
            cell_desc = oSheet.getCellByPosition(2, lrow)
            style_desc = cell_desc.CellStyle

            # Applichiamo la logica solo se la riga è di misurazione
            if 'comp 1-a' in style_desc:
                row_idx = str(lrow + 1)
                cell_qta = oSheet.getCellByPosition(9, lrow)

                if 'ROSSO' in style_desc:
                    # Invertiamo a POSITIVO
                    cell_qta.Formula = f'=IF(PRODUCT(E{row_idx}:I{row_idx})=0;"";PRODUCT(E{row_idx}:I{row_idx}))'
                    # Rimuoviamo il suffisso ROSSO (tua logica originale)
                    for x in range(2, 10):
                        current_cell = oSheet.getCellByPosition(x, lrow)
                        current_cell.CellStyle = current_cell.CellStyle.split(' ROSSO')[0]
                else:
                    # Invertiamo a NEGATIVO
                    cell_qta.Formula = f'=IF(PRODUCT(E{row_idx}:I{row_idx})=0;"";-PRODUCT(E{row_idx}:I{row_idx}))'
                    # Aggiungiamo il suffisso ROSSO (tua logica originale)
                    for x in range(2, 10):
                        current_cell = oSheet.getCellByPosition(x, lrow)
                        current_cell.CellStyle = current_cell.CellStyle + ' ROSSO'

    # Elaborazione per foglio CONTABILITA
    elif oSheet.Name == 'CONTABILITA':
        for lrow in sorted(list(lista_righe)):
            if 'comp 1-a' in oSheet.getCellByPosition(2, lrow).CellStyle:
                cell_9 = oSheet.getCellByPosition(9, lrow)
                cell_11 = oSheet.getCellByPosition(11, lrow)

                f1 = cell_9.Formula
                f2 = cell_11.Formula
                cell_11.Formula = f1
                cell_9.Formula = f2

                if cell_11.Value > 0:
                    for x in range(2, 12):
                        current_cell = oSheet.getCellByPosition(x, lrow)
                        current_cell.CellStyle = current_cell.CellStyle + ' ROSSO'
                else:
                    for x in range(2, 12):
                        current_cell = oSheet.getCellByPosition(x, lrow)
                        current_cell.CellStyle = current_cell.CellStyle.split(' ROSSO')[0]

########################################################################

def valuta_cella(oCell):
    '''
    Estrae qualsiasi valore da una cella, restituendo una stringa.
    '''
    # Otteniamo il tipo di cella (Enum UNO)
    tipo = oCell.Type.value
    valore = ""

    if tipo == 'FORMULA':
        # Se la formula contiene lettere (es. =A1+B1), prendiamo il risultato calcolato
        if re.search('[a-zA-Z]', oCell.Formula):
            valore = str(oCell.Value)
        else:
            # Se è una costante (es. =10+5), prendiamo solo la parte numerica
            valore = oCell.Formula.split('=')[-1]

    elif tipo == 'VALUE':
        valore = str(oCell.Value)

    elif tipo == 'TEXT':
        valore = oCell.String

    elif tipo == 'EMPTY':
        valore = ''

    # Rimuove spazi bianchi superflui (più efficace di if valore == ' ')
    return valore.strip()

########################################################################
@with_undo
@LeenoUtils.no_refresh
def dettaglio_misura_rigo():
    '''
    Aggiorna il dettaglio delle misure (formule esplicitate) solo per la riga corrente.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # Recuperiamo la riga corrente usando la tua funzione
    # pos = (nCol, nRow, NameSheet)
    pos = LeggiPosizioneCorrente()
    lrow = pos[1]

    # Costanti di LeenO
    COL_DESC = 2
    SEP = ' ►'

    cell_desc = oSheet.getCellByPosition(COL_DESC, lrow)
    desc_orig = cell_desc.String

    # 1. Pulizia: se esiste già un dettaglio, lo rimuoviamo per rigenerarlo
    if SEP in desc_orig:
        desc_orig = desc_orig.split(SEP)[0]
        cell_desc.String = desc_orig

    # 2. Controllo condizioni: stile rigo e voce non azzerata
    if 'comp 1-a' in cell_desc.CellStyle and "*** VOCE AZZERATA ***" not in desc_orig:
        parti = []
        ha_formula = False

        # Scansione colonne misure (5=N.parti, 6=Lung, 7=Larg, 8=Alt/Peso)
        for col in range(5, 9):
            cell_m = oSheet.getCellByPosition(col, lrow)

            if cell_m.Type.value == 'FORMULA':
                ha_formula = True
                # Estraiamo la formula senza '=' e gestiamo l'elevamento a potenza per compatibilità
                f_clean = cell_m.Formula.replace('^', '**').split('=')[-1]
                parti.append(f"({f_clean})")
            else:
                val_m = cell_m.String
                if val_m and val_m != "0":
                    parti.append(val_m)

        # 3. Se abbiamo trovato almeno una formula, generiamo il suffisso
        if ha_formula and parti:
            # Uniamo con '*' e ripuliamo eventuali doppi asterischi generati
            stringa_misure = "*".join(parti).replace('**', '*')

            # Formattazione finale: aggiunta separatore e conversione virgole decimali
            dettaglio = f"{SEP}{stringa_misure}".replace('.', ',')

            # Scrittura nella cella descrizione (se non è essa stessa una formula)
            if cell_desc.Type.value != 'FORMULA':
                cell_desc.String = desc_orig + dettaglio

########################################################################

@LeenoUtils.no_refresh
def dettaglio_misure(bit):
    '''
    Indica il dettaglio delle misure nel rigo di descrizione.
    bit 1: inserisce i dettagli (es. ►(2+2)*5)
    bit 0: cancella i dettagli
    '''
    oDoc = LeenoUtils.getDocument()
    try:
        oSheet = oDoc.CurrentController.ActiveSheet
    except Exception:
        return

    # Usiamo l'area usata per limitare il ciclo
    used_area = SheetUtils.getUsedArea(oSheet)
    ER = used_area.EndRow

    # Separatore utilizzato per identificare il dettaglio
    SEP = ' ►'

    if bit == 1:
        # Inizializzazione indicatore di progresso nella barra di stato
        total_steps = LeenoSheetUtils.cercaUltimaVoce(oSheet)
        indicator = oDoc.CurrentController.getStatusIndicator()
        indicator.start('Generazione dettagli misure...', total_steps)

        for lrow in range(0, ER + 1):
            if lrow % 10 == 0: # Aggiorna l'indicatore ogni 10 righe per non rallentare
                indicator.setValue(lrow)

            # Accesso alla cella descrizione (Colonna 2)
            cell_desc = oSheet.getCellByPosition(2, lrow)
            cell_style = cell_desc.CellStyle
            desc_text = cell_desc.String

            # Filtro: solo righe con stile 'comp 1-a' e non azzerate
            if 'comp 1-a' in cell_style and "*** VOCE AZZERATA ***" not in desc_text:

                # Evitiamo di processare righe che hanno già il dettaglio
                if SEP in desc_text:
                    continue

                parti_formula = []
                ha_formula = False

                # Controlla le colonne delle misure (dalla 5 alla 8)
                for el in range(5, 9):
                    cell_misura = oSheet.getCellByPosition(el, lrow)

                    if cell_misura.Type.value == 'FORMULA':
                        ha_formula = True
                        # Estrae la parte dopo l'uguale e pulisce
                        f_text = cell_misura.Formula.replace('^', '**').split('=')[-1]
                        parti_formula.append(f"({f_text})")
                    else:
                        val_text = cell_misura.String
                        if val_text and val_text != "0":
                            parti_formula.append(val_text)

                if ha_formula and parti_formula:
                    # Costruzione stringa: unisce le parti con '*' e pulisce i doppi asterischi
                    stringa_misure = "*".join(parti_formula).replace('**', '*')
                    # Sostituisce il punto con la virgola per standard italiano
                    stringa_finale = f"{SEP}{stringa_misure}".replace('.', ',')

                    # Scrive solo se non è una formula (per non sovrascrivere calcoli in descrizione)
                    if cell_desc.Type.value != 'FORMULA':
                        cell_desc.String = desc_text + stringa_finale

        indicator.end()

    else:
        # Modalità CANCELLAZIONE: molto più rapida
        for lrow in range(0, ER + 1):
            cell_desc = oSheet.getCellByPosition(2, lrow)
            txt = cell_desc.String
            if SEP in txt:
                cell_desc.String = txt.split(SEP)[0]

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


# def valida_cella(oCell, lista_val, titoloInput='', msgInput='', err=False):
#     '''
#     Validità lista valori
#     Imposta un elenco di valori a cascata, da cui scegliere.
#     oCell       { object }  : cella da validare
#     lista_val   { string }  : lista dei valori in questa forma: '"UNO";"DUE";"TRE"'
#     titoloInput { string }  : titolo del suggerimento che compare passando il cursore sulla cella
#     msgInput    { string }  : suggerimento che compare passando il cursore sulla cella
#     err         { boolean } : permette di abilitare il messaggio di errore per input non validi
#     '''
#     # oDoc = LeenoUtils.getDocument()
#     # oSheet = oDoc.CurrentController.ActiveSheet

#     oTabVal = oCell.getPropertyValue("Validation")
#     oTabVal.setPropertyValue('ConditionOperator', 1)

#     oTabVal.setPropertyValue("ShowInputMessage", True)
#     oTabVal.setPropertyValue("InputTitle", titoloInput)
#     oTabVal.setPropertyValue("InputMessage", msgInput)
#     oTabVal.setPropertyValue("ErrorMessage",
#                              "ERRORE: Questo valore non è consentito.")
#     oTabVal.setPropertyValue("ShowErrorMessage", err)
#     oTabVal.ErrorAlertStyle = uno.Enum(
#         "com.sun.star.sheet.ValidationAlertStyle", "STOP")
#     oTabVal.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")
#     oTabVal.Operator = uno.Enum("com.sun.star.sheet.ConditionOperator",
#                                 "EQUAL")
#     oTabVal.setFormula1(lista_val)
#     oCell.setPropertyValue("Validation", oTabVal)


def valida_cella(oCell, lista_val, titoloInput='', msgInput='', err=False):
    '''
    Imposta un elenco di valori a cascata (Validation LIST) su una cella.

    oCell       {object}  : oggetto cella (es. oSheet.getCellByPosition(0,0))
    lista_val   {string}  : stringa valori separati da punto e virgola: '"A";"B";"C"'
    titoloInput {string}  : titolo del tooltip di aiuto
    msgInput    {string}  : messaggio del tooltip di aiuto
    err         {boolean} : se True, impedisce l'inserimento di valori non in lista
    '''
    # Recuperiamo l'oggetto Validation esistente della cella
    oTabVal = oCell.Validation

    # Configurazione Messaggio di Input (il tooltip che appare al passaggio del mouse)
    oTabVal.ShowInputMessage = True
    oTabVal.InputTitle = titoloInput
    oTabVal.InputMessage = msgInput

    # Configurazione Messaggio di Errore
    oTabVal.ShowErrorMessage = err
    oTabVal.ErrorMessage = "ERRORE: Questo valore non è consentito."
    oTabVal.ErrorAlertStyle = uno.Enum(
        "com.sun.star.sheet.ValidationAlertStyle", "STOP")

    # Definizione del tipo di validazione: LIST
    oTabVal.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")

    # Impostazione della formula (la lista dei valori)
    oTabVal.setFormula1(lista_val)

    # Nota importante: l'oggetto Validation va riassegnato alla cella per rendere effettive le modifiche
    oCell.Validation = oTabVal


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

########################################################################

@LeenoUtils.no_refresh
def delete_cells(direction='up'):
    '''
    Elimina le celle selezionate e sposta le celle adiacenti.
    '''
    # Validazione rapida
    direction = direction.lower()
    if direction not in ('up', 'u', 'left', 'l'):
        raise ValueError("Direzione non valida. Usare 'up'/'u' o 'left'/'l'")

    oDoc = LeenoUtils.getDocument()
    ctx = LeenoUtils.getComponentContext()

    # Mappatura parametri per .uno:DeleteCell
    # In LibreOffice, per questo comando specifico:
    # Flags: 'U' sposta in alto, 'L' sposta a sinistra (a volte espresso come 'V' o 'H')
    # Tuttavia, l'approccio più robusto via Dispatcher è usare i parametri direzionali.

    is_left = direction in ('left', 'l')

    # Nota: Molti dispatcher usano 'Flags' per il tipo di eliminazione
    # e parametri specifici per la direzione.
    args = {
        'Flags': 'A', # Elimina tutto il contenuto
        'MoveMode': 1 if is_left else 0 # In genere 1=Sinistra, 0=Alto in questo contesto
    }

    # Se preferisci usare la logica esatta che avevi impostato con ToRight:
    # args = {'ToRight': is_left}

    properties = LeenoUtils.dictToProperties(args)

    oFrame = oDoc.CurrentController.Frame
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)

    dispatchHelper.executeDispatch(oFrame, ".uno:DeleteCell", "", 0, properties)
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
    Incolla solo il formato della cella (corrisponde al flag 'T' di Paste Special).
    '''
    oDoc = LeenoUtils.getDocument()
    ctx = LeenoUtils.getComponentContext()

    # Prepariamo i parametri per .uno:InsertContents tramite dizionario
    # Il flag 'T' indica 'Styles' (formattazione)
    paste_options = {
        'Flags': 'T',
        'FormulaCommand': 0,
        'SkipEmptyCells': False,
        'Transpose': False,
        'AsLink': False
    }

    # Convertiamo il dizionario in una tupla di PropertyValue usando l'utility dedicata
    properties = LeenoUtils.dictToProperties(paste_options)

    # Esecuzione del comando tramite DispatchHelper
    oFrame = oDoc.CurrentController.Frame
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, '.uno:InsertContents', '', 0, properties)

    # Deseleziona: un metodo pulito è selezionare l'oggetto vuoto come facevi,
    # oppure selezionare la singola cella attiva per rimuovere l'evidenziazione del range incollato.
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))


########################################################################
@LeenoUtils.no_refresh
def MENU_copia_celle_visibili():
    '''
    Copia negli appunti solo le celle visibili della selezione corrente.
    Utilizza un foglio temporaneo per consolidare i dati ed evitare celle vuote intermedie.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    selection = oDoc.CurrentSelection

    # Gestione sicura del range
    if hasattr(selection, "getRangeAddresses"):
        oRangeAddress = selection.getRangeAddresses()[0]
    else:
        oRangeAddress = selection.getRangeAddress()

    # Uso del context manager per disabilitare il refresh ed evitare sfarfallii
    try:
        # Creazione o recupero foglio temporaneo
        if not oDoc.getSheets().hasByName('tmp_clip'):
            oDoc.getSheets().insertNewByName('tmp_clip', oDoc.Sheets.Count)
        tmp = oDoc.getSheets().getByName('tmp_clip')

        # 1. Otteniamo solo le celle visibili dal range originale
        oRange = oSheet.getCellRangeByPosition(
            oRangeAddress.StartColumn, oRangeAddress.StartRow,
            oRangeAddress.EndColumn, oRangeAddress.EndRow
        )
        visible_cells = oRange.queryVisibleCells()

        # 2. Copiamo le celle visibili nel foglio temporaneo
        # Nota: copyRange non supporta queryVisibleCells direttamente,
        # quindi usiamo il metodo delle coordinate filtrate o il foglio d'appoggio.
        dest_cursor_row = 0
        for sub_range in visible_cells:
            addr = sub_range.getRangeAddress()
            rows_count = addr.EndRow - addr.StartRow + 1
            cols_count = addr.EndColumn - addr.StartColumn + 1

            dest_addr = tmp.getCellByPosition(0, dest_cursor_row).getCellAddress()
            tmp.copyRange(dest_addr, addr)
            dest_cursor_row += rows_count

        # 3. Selezioniamo il risultato nel foglio tmp e copiamo
        final_range = tmp.getCellRangeByPosition(0, 0, oRangeAddress.EndColumn - oRangeAddress.StartColumn, dest_cursor_row - 1)
        oDoc.CurrentController.select(final_range)

        # Dispatch del comando Copy tramite le utility
        # ctx = LeenoUtils.getComponentContext()
        # dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
        # dispatchHelper.executeDispatch(oDoc.CurrentController.Frame, ".uno:Copy", "", 0, ())
        comando('Copy')

    finally:
        # Pulizia: rimuove il foglio e torna alla selezione originale
        if oDoc.getSheets().hasByName('tmp_clip'):
            oDoc.getSheets().removeByName('tmp_clip')

        oDoc.CurrentController.setActiveSheet(oSheet)
        oDoc.CurrentController.select(oRange)

########################################################################
def LeggiPosizioneCorrente():
    '''
    Restituisce la tupla (IDcolonna, IDriga, NameSheet) della posizione corrente.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # Sfrutta la funzione già esistente in LeenoUtils per ottenere riga e colonna
    # getCursorPosition restituisce (Row, Column)
    pos = LeenoUtils.getCursorPosition(oDoc)

    if pos is not None:
        nRow, nCol = pos
        return (nCol, nRow, oSheet.Name)

    # Fallback in caso di selezione non valida
    return (None, None, oSheet.Name)

########################################################################
# numera le voci di computo o contabilità
# @Debug.measure_time() # Misura il tempo di esecuzione della funzione
# @LeenoUtils.no_refresh # Disabilita il refresh del documento durante l'esecuzione della funzione
def MENU_numera_voci():
    '''
    Comando di menu per numera_voci()
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    # LeenoSheetUtils.numeraVoci(oSheet, 4, True)
    with LeenoUtils.no_refresh():
        numera_voci()
    Rinumera_TUTTI_Capitoli2(oSheet)


def numera_voci():
    '''
    Rinumera tutte le voci dalla riga 4 alla fine del foglio.
    Utilizza setDataArray per la massima velocità e compatibilità con l'Undo.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # 1. Definiamo l'area di lavoro (dalla riga 4 alla fine dell'area usata)
    # Nota: index 3 corrisponde alla riga 4 di Calc
    first_row = 3
    last_row = SheetUtils.getUsedArea(oSheet).EndRow

    if last_row < first_row:
        return

    # 2. Acquisiamo i dati e gli stili
    # Prendiamo le prime due colonne (A e B)
    oRange = oSheet.getCellRangeByPosition(0, first_row, 1, last_row)
    data = list(oRange.getDataArray())

    n = 1

    # 3. Ciclo di elaborazione in memoria (non tocca la UI)
    for i in range(len(data)):
        current_row_index = first_row + i
        # Dobbiamo comunque controllare lo stile per sapere se numerare
        # ma lo facciamo in lettura, che è molto più veloce della scrittura
        oCellB = oSheet.getCellByPosition(1, current_row_index)
        style = oCellB.CellStyle
        color = oCellB.CellBackColor

        row_list = list(data[i])
        '''
        Logica di numerazione:
        - Se è una voce di computo (style in ('comp Art-EP', 'comp Art-EP_R')),
          allora assegna il numero progressivo n alla colonna A e incrementa n.
        - Altrimenti, imposta la colonna A a stringa vuota.
        '''
        if style in ('comp Art-EP', 'comp Art-EP_R'):
            '''
            Se è una voce di computo, numeriamo
            '''
            # if color == COLORE_GRIGIO_INATTIVA: # Grigio (voce non numerata/computata)
            #     '''
            #     La voce non è numerata/computata
            #     ️ Imposta la colonna A a stringa vuota
            #     '''
            #     row_list[0] = ""
            # else:
            row_list[0] = float(n) # Calc preferisce float per i valori numerici
            n += 1
        else:
            # Se non è una voce di computo, puliamo la cella del numero (opzionale)
            row_list[0] = ""

        data[i] = tuple(row_list)

    # 4. SCRITTURA ATOMICA: Unica operazione nel registro di Undo
    oRange.setDataArray(tuple(data))
########################################################################
@LeenoUtils.no_refresh
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
@LeenoUtils.no_refresh
@with_undo("Inserisci voce elenco prezzi")
def ins_voce_elenco():
    '''
    Inserisce una nuova riga voce in Elenco Prezzi
    '''
    oDoc = LeenoUtils.getDocument()

    oSheet = oDoc.CurrentController.ActiveSheet
    _gotoCella("A5")
    oSheet.getRows().insertByIndex(4, 1)

    lrow = LeggiPosizioneCorrente()[1] -1

    _gotoCella("L5")
    oDoc.CurrentController.select(oSheet.getCellRangeByName('L5:Z5'))
    comando('FillDown')
    comando('FillDown')

    _gotoCella("A5")


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
@LeenoUtils.no_refresh # Disabilita il refresh del documento durante l'esecuzione della funzione
@LeenoUtils.preserva_posizione(step=0)
def rigenera_tutte(arg=None, ):
    '''
    Ripristina le formule in tutto il foglio
    '''
    # with LeenoUtils.DocumentRefreshContext(False):

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
        last_row = LeenoSheetUtils.cercaUltimaVoce(oSheet) +1
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

    LeenoSheetUtils.adattaAltezzaRiga(oSheet)

########################################################################
@LeenoUtils.no_refresh # Decoratore per disabilitare l'aggiornamento del documento
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
# @Debug.measure_time()
def MENU_nuova_voce_scelta():  # assegnato a ctrl-shift-n
    '''
    Contestualizza in ogni tabella l'inserimento delle voci.
    '''
    oDoc = LeenoUtils.getDocument()
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
@LeenoUtils.no_refresh # Disabilita il refresh del documento durante l'esecuzione della funzione
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
    numera_voci()
    if cfg.read('Generale', 'pesca_auto') == '1':
        if arg == 0:
            return
        pesca_cod()


########################################################################
@LeenoUtils.no_refresh
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

@LeenoUtils.no_refresh
def inizializza_elenco():
    '''
    Riscrive le intestazioni di colonna e le formule dei totali in Elenco Prezzi.
    Versione ottimizzata per performance e leggibilità.
    '''
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
        'T2': f'=IF(SUBTOTAL(9;T:T)=0;"--";SUBTOTAL(9;T:T))',
        'U2': f'=IF(SUBTOTAL(9;U:U)=0;"--";SUBTOTAL(9;U:U))',
        'V2': f'=IF(SUBTOTAL(9;V:V)=0;"--";SUBTOTAL(9;V:V))',
        'X2': f'=IF(SUBTOTAL(9;X:X)=0;"--";SUBTOTAL(9;X:X))',
        'Y2': f'=-IF(SUBTOTAL(9;Y:Y)=0;"--";SUBTOTAL(9;Y:Y))',

        # 'T2': f'=IF(SUBTOTAL(9;T3:T{y})=0;"--";SUBTOTAL(9;T3:T{y}))',
        # 'U2': f'=IF(SUBTOTAL(9;U3:U{y})=0;"--";SUBTOTAL(9;U3:U{y}))',
        # 'V2': f'=IF(SUBTOTAL(9;V3:V{y})=0;"--";SUBTOTAL(9;V3:V{y}))',
        # 'X2': f'=IF(SUBTOTAL(9;X3:X{y})=0;"--";SUBTOTAL(9;X3:X{y}))',
        # 'Y2': f'=IF(SUBTOTAL(9;Y3:Y{y})=0;"--";SUBTOTAL(9;Y3:Y{y}))',
    }

    for cell, formula in FORMULE_TOTALI.items():
        oSheet.getCellRangeByName(cell).Formula = formula

    # 7. Righe di totale finali
    TOTALI_FINALI = {
        15: 'TOTALE',
        19: f'=IF(SUBTOTAL(9;T:T)=0;"--";SUBTOTAL(9;T:T))',
        20: f'=IF(SUBTOTAL(9;U:U)=0;"--";SUBTOTAL(9;U:U))',
        21: f'=IF(SUBTOTAL(9;V:V)=0;"--";SUBTOTAL(9;V:V))',
        23: f'=IF(SUBTOTAL(9;X:X)=0;"--";SUBTOTAL(9;X:X))',
        24: f'=IF(SUBTOTAL(9;Y:Y)=0;"--";SUBTOTAL(9;Y:Y))',



        # 19: f'=IF(SUBTOTAL(9;T3:T{y})=0;"--";SUBTOTAL(9;T3:T{y}))',
        # 20: f'=IF(SUBTOTAL(9;U3:U{y})=0;"--";SUBTOTAL(9;U3:U{y}))',
        # 21: f'=IF(SUBTOTAL(9;V3:V{y})=0;"--";SUBTOTAL(9;V3:V{y}))',
        # 23: f'=IF(SUBTOTAL(9;X3:X{y})=0;"--";SUBTOTAL(9;X3:X{y}))',
        # 24: f'=IF(SUBTOTAL(9;Y3:Y{y})=0;"--";SUBTOTAL(9;Y3:Y{y}))',
        # 24: f'=IF(SUBTOTAL(9;Y3:Y{y})=0;"--";SUBTOTAL(9;Y3:Y{y}))',
    }
    oSheet.getCellRangeByName(f'L{y+1}:N{y+1}').merge(True)
    oSheet.getCellRangeByName(f'P{y+1}:R{y+1}').merge(True)

    for col, value in TOTALI_FINALI.items():
        oSheet.getCellByPosition(col, y).Formula = value if isinstance(value, str) and value.startswith('=') else value

    oSheet.getCellRangeByPosition(10, y, 25, y).CellStyle = STILI['contab']

    # 8. Pulizia finale e stili
    # y += 1
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
                col_end, y - 3
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
@LeenoUtils.no_refresh
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
    LeenoUtils.memorizza_posizione()
    MENU_struttura_on()
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    LeenoUtils.ripristina_posizione()



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
@LeenoUtils.no_refresh
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
@LeenoUtils.no_refresh
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
    # with LeenoUtils.DocumentRefreshContext(False):

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
    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


########################################################################

# @Debug.measure_time()
@LeenoUtils.no_refresh
def MENU_importa_stili():
    '''
    Importa tutti gli stili da un documento di riferimento. Se non è
    selezionato, il file di riferimento è il template di leenO.
    '''
    # with LeenoUtils.DocumentRefreshContext(False):

    if Dialogs.YesNoDialog(IconType="question",
        Title='Vuoi sostituire gli stili del documento?',
        Text="""
► Scegli "Sì" per sostituire gli stili del documento selezionando
    un file di riferimento (facoltativo). Se non selezioni alcun
    file, verranno applicati gli stili predefiniti di LeenO.
    ----
    ⚠️  ATTENZIONE: l'applicazione di stili che visualizzano un
    numero diverso di cifre decimali può influire sui risultati
    dei calcoli, in funzione dell'opzione “Precisione come mostrato”.

► Scegli "No" per mantenere gli stili attuali.
"""
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
@with_undo
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
@LeenoUtils.no_refresh
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
# @LeenoUtils.no_refresh # decoratore per disabilitare il refresh automatico
def vedi_voce_xpwe(oSheet, lrow, vRif):
    """
    (riga d'inserimento, riga di riferimento)
    """
    # with LeenoUtils.DocumentRefreshContext(False):
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
@with_undo
def MENU_vedi_voce():
    '''
    Inserisce un riferimento a voce precedente sulla riga corrente.
    '''
    # Usiamo il context manager per velocizzare le operazioni pesanti
    with LeenoUtils.no_refresh_context():
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

            try:
                to_row_index = int(to.split('$')[-1]) - 1
                _gotoCella(2, lrow)

                if to_row_index < lrow:
                    vedi_voce_xpwe(oSheet, lrow, to_row_index)

            except Exception:
                pass
    oSheet.getRows().getByIndex(lrow).OptimalHeight = True
    # Fuori dal context manager il refresh è già attivo,
    # ma se Calc fosse pigro, puoi forzare un aggiornamento finale qui.

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
@LeenoUtils.no_refresh
@with_undo
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

import unohelper
from com.sun.star.awt import XKeyHandler

class LeenoKeyHandler(unohelper.Base, XKeyHandler):
    def __init__(self):
        self.ctrl_pressed = False
        self.shift_pressed = False

    def keyPressed(self, event):
        # Modifiers: 2 = CTRL, 1 = SHIFT
        if event.Modifiers == 2:
            self.ctrl_pressed = True
        elif event.Modifiers == 1:
            self.shift_pressed = True
        return False # Permette a LO di processare il tasto normalmente

    def keyReleased(self, event):
        self.ctrl_pressed = False
        self.shift_pressed = False
        return False

    def disposing(self, event):
        pass

# Istanza globale del gestore
KEY_HANDLER = None

def register_key_handler():
    global KEY_HANDLER
    if KEY_HANDLER is None:
        oDoc = LeenoUtils.getDocument()
        KEY_HANDLER = LeenoKeyHandler()
        oDoc.CurrentController.addKeyHandler(KEY_HANDLER)


@with_undo
def MENU_filtra_codice():
    import sys
    is_ctrl = False
    is_shift = False

    # --- Rilevamento Tasti Modificatori ---
    if sys.platform == 'win32':
        try:
            import ctypes
            # GetAsyncKeyState legge lo stato fisico del tasto in tempo reale
            # 0x10 = Shift, 0x11 = Control
            # Il bit 0x8000 indica se il tasto è attualmente premuto
            is_shift = (ctypes.windll.user32.GetAsyncKeyState(0x10) & 0x8000) != 0
            is_ctrl = (ctypes.windll.user32.GetAsyncKeyState(0x11) & 0x8000) != 0
        except Exception as e:
            # In caso di errore ctypes, procediamo come click normale
            pass
    else:
        # Per Linux/Mac: verifica se il KEY_HANDLER (XKeyHandler) è attivo
        if 'KEY_HANDLER' in globals() and KEY_HANDLER:
            is_ctrl = KEY_HANDLER.ctrl_pressed
            is_shift = KEY_HANDLER.shift_pressed

    # --- Logica di Navigazione ---
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    # Otteniamo la riga corrente una sola volta
    lrow = LeggiPosizioneCorrente()[1]

    if not is_ctrl and not is_shift:
        # Comportamento standard: filtra la riga corrente
        filtra_codice()
        return

    if is_ctrl:
        # VAI ALLA PROSSIMA VOCE (Ctrl + Click)
        target_next = LeenoSheetUtils.prossimaVoce(oSheet, lrow, saltaCat=True)
        _gotoCella(2, target_next)
        filtra_codice()

    elif is_shift:
        # VAI ALLA VOCE PRECEDENTE (Shift + Click)
        # 1. Trova l'inizio della voce in cui ti trovi
        curr_start = LeenoSheetUtils.prossimaVoce(oSheet, lrow, n=0, saltaCat=True)
        search_row = curr_start - 1

        # 2. Risali finché non trovi uno stile di computo o contabilità
        stili_validi = LeenoUtils.getGlobalVar('stili_computo') + LeenoUtils.getGlobalVar('stili_contab')

        found_row = -1
        while search_row >= 0:
            try:
                style = oSheet.getCellByPosition(0, search_row).CellStyle
                if style in stili_validi:
                    found_row = search_row
                    break
            except:
                pass
            search_row -= 1

        if found_row != -1:
            # 3. Vai all'inizio della voce precedente trovata
            target_prev = LeenoSheetUtils.prossimaVoce(oSheet, found_row, n=0, saltaCat=True)
            _gotoCella(2, target_prev)
            filtra_codice()

@LeenoUtils.no_refresh
def filtra_codice(voce=None):
    '''
    Applica un filtro di visualizzazione basato sul raggruppamento (outline).
    Il cursore si posiziona sulla prima occorrenza della voce trovata.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_contab = LeenoUtils.getGlobalVar('stili_contab')
    stili_totali = stili_computo + stili_contab

    # --- 1. IDENTIFICAZIONE DELLA VOCE E DEL FOGLIO ---
    if oSheet.Name == "Elenco Prezzi":
        oCell_C2 = oSheet.getCellRangeByName('C2')
        lrow_ep = LeggiPosizioneCorrente()[1]
        voce = oSheet.getCellByPosition(0, lrow_ep).String

        if oCell_C2.String in ('<DIALOGO>', ''):
            try:
                elaborato = DLG.ScegliElaborato(Titolo='Ricerca di ' + voce)
                GotoSheet(elaborato)
            except Exception:
                return
        else:
            elaborato = oCell_C2.String
            try:
                GotoSheet(elaborato)
            except Exception:
                return

        oSheet = oDoc.CurrentController.ActiveSheet
        _gotoCella(0, 6)
        LeenoSheetUtils.prossimaVoce(oSheet, 6, 1, saltaCat=True)

    # Posizione di partenza per fallback
    lrow_originale = LeggiPosizioneCorrente()[1]

    if not voce:
        try:
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow_originale)
            sopra = sStRange.RangeAddress.StartRow
            voce = oSheet.getCellByPosition(1, sopra + 1).String
        except Exception:
            try:
                if oSheet.getCellByPosition(0, lrow_originale).CellStyle in stili_totali:
                    voce = LeenoComputo.datiVoceComputo(oSheet, lrow_originale)[1][1]
            except:
                pass

    if not voce:
        Dialogs.Exclamation(Title='ATTENZIONE!', Text='Seleziona una voce o una misurazione.')
        return

    # --- 2. APPLICAZIONE DEL FILTRO ---
    oSheet.clearOutline()
    fine = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1

    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet

    n = 0
    closest_row = None
    min_dist = float('inf')

    while n < fine:
        cell_style = oSheet.getCellByPosition(0, n).CellStyle

        if cell_style in ('Comp Start Attributo', 'Comp Start Attributo_R'):
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, n)
            sopra = sStRange.RangeAddress.StartRow
            sotto = sStRange.RangeAddress.EndRow

            codice_corrente = oSheet.getCellByPosition(1, sopra + 1).String

            if codice_corrente != voce:
                oCellRangeAddr.StartRow = sopra
                oCellRangeAddr.EndRow = sotto
                oSheet.group(oCellRangeAddr, 1)
                oSheet.getCellRangeByPosition(0, sopra, 0, sotto).Rows.IsVisible = False
            else:
                oSheet.getCellRangeByPosition(0, sopra, 0, sotto).Rows.IsVisible = True
                oSheet.getCellByPosition(1, sopra + 1).CellBackColor = COLORE_VERDE_SPUNTA

                dist = abs((sopra + 1) - lrow_originale)
                if dist < min_dist:
                    min_dist = dist
                    closest_row = sopra + 1

            n = sotto + 1
        else:
            n += 1

    # --- 3. COLONNE E CHIUSURA ---
    oCellRangeAddr.StartColumn = 29
    oCellRangeAddr.EndColumn = 30
    oSheet.group(oCellRangeAddr, 0)
    oSheet.getCellRangeByPosition(29, 0, 30, 0).Columns.IsVisible = False

    if closest_row is not None:
        _gotoCella(1, closest_row)

########################################################################
# @LeenoUtils.no_refresh
def MENU_struttura_on():
    oDoc = LeenoUtils.getDocument()
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
    oSheet.getColumns().getByIndex(29).IsVisible = False
    oSheet.getColumns().getByIndex(30).IsVisible = False

    # Raggruppa le colonne di misura e mostra
    oCellRangeAddr.StartColumn = 5
    oCellRangeAddr.EndColumn = 8
    oSheet.group(oCellRangeAddr, 0)
    for i in range(5, 9):
        oSheet.getColumns().getByIndex(i).IsVisible = True

    # # attiva la prog
    indicator = oDoc.getCurrentController().getStatusIndicator()
    indicator.start('Creazione vista struttura in corso...', 4)
    '''
    # se color = True allora struct(n, color=color) colora le righe
    '''
    color = False
    color = True
    for n in range(0, 4):
        indicator.Value = n
        struct(n, color = color)
        # color = not color
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
        # LeenoUtils.memorizza_posizione()
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        oSheet.clearOutline()
        # LeenoUtils.ripristina_posizione()


@LeenoUtils.no_refresh # Decoratore per disabilitare l'aggiornamento del documento
def struct(level, vedi = True, color = False):
    '''
    mette in vista struttura secondo categorie
    level { integer } : specifica il livello di categoria
    ### COMPUTO/VARIANTE ###
    0 = super-categoria
    1 = categoria
    2 = sotto-categoria
    3 = intera voce di misurazione

    se color = True allora struct(n, color=color) colora le righe
    se vedi = True allora struct(n, vedi=vedi) mostra le righe
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
            # 'ULTIMUS_2',
            # 'ULTIMUS_3',
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

    colors = LeenoConfig.PASTEL_COLORS

    for i, el in enumerate(lista_cat):
        oCellRangeAddr.StartRow = el[0]
        oCellRangeAddr.EndRow = el[1]
        oSheet.group(oCellRangeAddr, 1)

        # Applica colore alla prima colonna SOLO per le Categorie (livello 1) e se richiesto
        if color and level == 1:
            colore = colors[i % len(colors)]
            # Include l'intestazione (sopra - Dsopra) e arriva fino alla fine del blocco (sotto)
            # Nota: el[0] è 'sopra', el[1] è 'sotto'
            r_start = el[0] - Dsopra
            r_end = el[1]
            try:
                # Colora la prima colonna (A)
                oSheet.getCellRangeByPosition(0, r_start, 0, r_end).CellBackColor = colore
                # Colora l'intestazione (righe r_start) fino alla colonna AE (indice 30)
                oSheet.getCellRangeByPosition(0, r_start, 30, r_start).CellBackColor = colore
            except:
                pass

        if vedi == False:
            oSheet.getCellRangeByPosition(0, el[0], 0, el[1]).Rows.IsVisible = False


########################################################################
def MENU_apri_manuale():
    '''
    Apre il manuale utente di LeenO in formato PDF.

    Utilizza il gestore di sistema predefinito per aprire il file MANUALE_LeenO.pdf
    situato nella directory di installazione di LeenO.

    Returns:
        None
    '''
    apri = LeenoUtils.createUnoService("com.sun.star.system.SystemShellExecute")
    apri.execute(LeenO_path() + '/MANUALE_LeenO.pdf', "", 0)


########################################################################
@LeenoUtils.no_refresh
def autoexec_off():
    '''
    @@ DA DOCUMENTARE
    '''
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
    autoexec_run()

@LeenoUtils.no_refresh
def autoexec_run():
    '''
    questa è richiamata da creaComputo()
    '''
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
        LeenoUtils.memorizza_posizione()

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
            struct(3, color = False)
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
            struct(3, vedi=False, color=False)

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
        LeenoUtils.ripristina_posizione()

def vista_terra_terra():
    vista_configurazione('terra_terra')

def vista_mdo():
    vista_configurazione('mdo')


########################################################################
@LeenoUtils.no_refresh
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
# def inizializza():
#     '''
#     Inserisce tutti i dati e gli stili per preparare il lavoro.
#     lanciata in autoexec()
#     '''
#     oDoc = LeenoUtils.getDocument()

#     #  oDoc.IsUndoEnabled = False
#     oDoc.getSheets().getByName('copyright_LeenO').getCellRangeByName(
#         'A3').String = '# © 2001-2013 Bartolomeo Aimar - © 2014-' + str(
#             datetime.now().year) + ' Giuseppe Vizziello'

#     oUDP = oDoc.getDocumentProperties().getUserDefinedProperties()
#     oSheet = oDoc.getSheets().getByName('S1')

#     oSheet.getCellRangeByName('G219').String = 'Copyright 2014-' + str(datetime.now().year)

#     # allow non-numeric version codes (example: testing)
#     rvc = version_code.read().split('-')[0]
#     if isinstance(rvc, str):
#         oSheet.getCellRangeByName('H194').String = rvc
#     else:
#         oSheet.getCellRangeByName('H194').Value = rvc

#     oSheet.getCellRangeByName('I194').Value = LeenoUtils.getGlobalVar('Lmajor')
#     oSheet.getCellRangeByName('J194').Value = LeenoUtils.getGlobalVar('Lminor')
#     oSheet.getCellRangeByName('H291').Value = oUDP.Versione
#     oSheet.getCellRangeByName('I291').String = oUDP.Versione_LeenO.split('.')[0]
#     oSheet.getCellRangeByName('J291').String = oUDP.Versione_LeenO.split('.')[1]

#     oSheet.getCellRangeByName('H295').String = oUDP.Versione_LeenO.split('.')[0]
#     oSheet.getCellRangeByName('I295').String = oUDP.Versione_LeenO.split('.')[1]
#     oSheet.getCellRangeByName('J295').String = oUDP.Versione_LeenO.split('.')[2]

#     oSheet.getCellRangeByName('K194').String = LeenoUtils.getGlobalVar('Lsubv')
#     oSheet.getCellRangeByName('H296').Value = LeenoUtils.getGlobalVar('Lmajor')
#     oSheet.getCellRangeByName('I296').Value = LeenoUtils.getGlobalVar('Lminor')
#     oSheet.getCellRangeByName('J296').String = LeenoUtils.getGlobalVar('Lsubv')

#     if oDoc.getSheets().hasByName('CONTABILITA'):
#         oSheet.getCellRangeByName('H328').Value = 1
#     else:
#         oSheet.getCellRangeByName('H328').Value = 0

# # inizializza la lista di scelta in elenco Prezzi
#     oCell = oDoc.getSheets().getByName('Elenco Prezzi').getCellRangeByName('C2')
#     valida_cella(oCell,
#                  '"<DIALOGO>";"COMPUTO";"VARIANTE";"CONTABILITA"',
#                  titoloInput='Scegli...',
#                  msgInput='Applica Filtra Codice a...',
#                  err=True)
#     oCell.String = "<DIALOGO>"
#     oCell.CellStyle = 'EP-aS'
#     oCell = oDoc.getSheets().getByName('Elenco Prezzi').getCellRangeByName('C1')
#     oCell.String = "Applica Filtro a:"
#     oCell.CellStyle = 'EP-aS'
# # inizializza la lista di scelta per la copertona cP_Cop
#     oCell = oDoc.getSheets().getByName('cP_Cop').getCellRangeByName('B19')
#     # if oCell.String == '';
#     valida_cella(oCell,
#                  '"ANALISI DI PREZZO";"ELENCO PREZZI";"ELENCO PREZZI E COSTI ELEMENTARI";\
#                  "COMPUTO METRICO";"PERIZIA DI VARIANTE";"LIBRETTO DELLE MISURE";"REGISTRO DI CONTABILITÀ";\
#                  "S.A.L. A TUTTO IL"',
#                  titoloInput='Scegli...',
#                  msgInput='Titolo della copertina...',
#                  err=False)
#     # oCell.String = ""
#     # Indica qual è il Documento Principale
#     ScriviNomeDocumentoPrincipale()
    # nascondi_sheets()

def inizializza():
    '''
    Configura dati, versioni e menu a tendina all'avvio del documento.
    Lanciata solitamente da autoexec().
    '''
    oDoc = LeenoUtils.getDocument()

    # 1. Aggiornamento Copyright (Foglio nascosto e S1)
    current_year = str(datetime.now().year)
    copy_string = f"# © 2001-2013 Bartolomeo Aimar - © 2014-{current_year} Giuseppe Vizziello"

    try:
        oDoc.Sheets.getByName('copyright_LeenO').getCellRangeByName('A3').String = copy_string
        oSheetS1 = oDoc.Sheets.getByName('S1')
        oSheetS1.getCellRangeByName('G219').String = f"Copyright 2014-{current_year}"
    except:
        pass # Evita interruzioni se i fogli di servizio sono mancanti

    # 2. Sincronizzazione Versioni
    oUDP = oDoc.getDocumentProperties().getUserDefinedProperties()

    # Versione LeenO (es. "3.22.1")
    v_parts = oUDP.Versione_LeenO.split('.')
    v_major = v_parts[0] if len(v_parts) > 0 else "0"
    v_minor = v_parts[1] if len(v_parts) > 1 else "0"
    v_sub   = v_parts[2] if len(v_parts) > 2 else "0"

    # Scrittura su S1 (usando i valori globali caricati all'avvio)
    oSheetS1.getCellRangeByName('I194').Value = LeenoUtils.getGlobalVar('Lmajor')
    oSheetS1.getCellRangeByName('J194').Value = LeenoUtils.getGlobalVar('Lminor')
    oSheetS1.getCellRangeByName('K194').String = LeenoUtils.getGlobalVar('Lsubv')

    # Sincronizzazione con proprietà utente
    oSheetS1.getCellRangeByName('H291').Value = oUDP.Versione
    oSheetS1.getCellRangeByName('I291').String = v_major
    oSheetS1.getCellRangeByName('J291').String = v_minor

    # 3. Setup Menu a tendina (Validazione Dati)

    # Validazione Elenco Prezzi
    oSheetEP = oDoc.Sheets.getByName('Elenco Prezzi')
    cell_filtro = oSheetEP.getCellRangeByName('C2')
    # Nota: la funzione valida_cella deve essere definita nel tuo modulo
    valida_cella(cell_filtro,
                 '"<DIALOGO>";"COMPUTO";"VARIANTE";"CONTABILITA"',
                 titoloInput='Scegli...',
                 msgInput='Applica Filtra Codice a...',
                 err=True)
    cell_filtro.String = "<DIALOGO>"
    cell_filtro.CellStyle = 'EP-aS'

    oSheetEP.getCellRangeByName('C1').String = "Applica Filtro a:"

    # Validazione Copertina
    if oDoc.Sheets.hasByName('cP_Cop'):
        cell_cop = oDoc.Sheets.getByName('cP_Cop').getCellRangeByName('B19')
        elenco_titoli = ('"ANALISI DI PREZZO";"ELENCO PREZZI";'
                         '"ELENCO PREZZI E COSTI ELEMENTARI";"COMPUTO METRICO";'
                         '"PERIZIA DI VARIANTE";"LIBRETTO DELLE MISURE";'
                         '"REGISTRO DI CONTABILITÀ";"S.A.L. A TUTTO IL"')

        valida_cella(cell_cop, elenco_titoli,
                     titoloInput='Scegli...',
                     msgInput='Titolo della copertina...',
                     err=False)

    # 4. Finalizzazione
    ScriviNomeDocumentoPrincipale()


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
    # LeenoUtils.DocumentRefresh(True)
    # with LeenoUtils.DocumentRefreshContext(False):

    LeenoUtils.memorizza_posizione()

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
    sString.Text = version_code.read()[6:]

    # sString.Text = (
    #     str(LeenoUtils.getGlobalVar('Lmajor')) + '.' +
    #     str(LeenoUtils.getGlobalVar('Lminor')) + '.' +
    #     LeenoUtils.getGlobalVar('Lsubv'))

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
    LeenoUtils.ripristina_posizione()
    LeenoUtils.DocumentRefresh(True)
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
    """
    Esegue un backup incrementale del documento corrente, mantenendo
    un numero massimo di copie come definito nelle impostazioni.
    """
    oDoc = LeenoUtils.getDocument()

    orig_url = oDoc.getURL()
    if not orig_url:
        return

    # Nome base del file
    base = os.path.basename(orig_url)
    name_without_ext = '.'.join(base.split('.')[:-1])

    # Timestamp compatto: yyyyMMddHHmm
    tempo = datetime.now().strftime("%Y%m%d%H%M")

    # Nome del file di backup
    dest_name = f"{name_without_ext}-{tempo}.ods"

    # Cartella dei backup
    dir_bak_url = os.path.dirname(orig_url) + "/leeno-bk/"
    dir_bak_sys = uno.fileUrlToSystemPath(dir_bak_url)

    # Crea cartella se non esiste
    if not os.path.exists(dir_bak_sys):
        os.makedirs(dir_bak_sys)

    # URL del file di destinazione
    dest_url = dir_bak_url + dest_name

    # Salva il documento
    try:
        oDoc.storeToURL(dest_url, ())
    except Exception as e:
        DLG.chi(f"Errore durante il salvataggio del backup: {e}")
        return

    # Gestione del numero massimo di copie
    max_copies = int(cfg.read("Generale", "copie_backup"))
    prefix = name_without_ext + "-"

    # Lista dei file nella cartella di backup
    files = sorted(
        (f for f in os.listdir(dir_bak_sys) if f.startswith(prefix)),
        reverse=True
    )

    # Elimina le copie in eccesso
    for i, fname in enumerate(files):
        if i >= max_copies:
            try:
                os.remove(os.path.join(dir_bak_sys, fname))
            except Exception as e:
                DLG.chi(f"Impossibile eliminare {fname}: {e}")

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
    ctx = LeenoUtils.getComponentContext()
    smgr = ctx.ServiceManager

    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    frame = desktop.getCurrentFrame()

    # Esegue il comando .uno:FillDown due volte
    dispatcher.executeDispatch(frame, ".uno:FillDown", "", 0, ())
    dispatcher.executeDispatch(frame, ".uno:FillDown", "", 0, ())


########################################################################
########################################################################

@LeenoUtils.no_refresh
def MENU_taglia_x():
    '''
    Taglia il contenuto della selezione (senza formattazione).
    Funziona solo su selezioni singole (limite di LibreOffice).
    '''
    oDoc = LeenoUtils.getDocument()
    selection = oDoc.CurrentSelection

    # Se è una selezione multipla, Calc non permette il Copy,
    # quindi usciamo per evitare errori nel dispatcher.
    if hasattr(selection, "getRangeAddresses"):
        return

    comando('Copy')

    # Flag per pulire i dati mantenendo bordi e stili
    flags = VALUE + DATETIME + STRING + ANNOTATION + FORMULA + OBJECTS + EDITATTR

    # Pulizia sicura del contenuto
    try:
        selection.clearContents(flags)
    except Exception as e:
        DLG.errore(f"Errore durante il taglio: {e}")
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
@LeenoUtils.no_refresh
def calendario_liste():
    '''
    Colora le colonne del sabato e della domenica
    nel foglio LISTA ore del file di computo
    0 = nero
    COLORE_BIANCO_SFONDO = bianco
    12632256 = grigio chiaro
    14277081 = azzurro chiaro
    13434777 = giallo
    8388608 = verde chiaro
    16753920 = arancione
    16711680 = rosso
    255 = blu
    8388736 = verde scuro
    6697728 = marrone
    10066329 = grigio scuro
    '''

    def rgb(r, g, b):
        return 256*256*r + 256*g + b

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



########################################################################
@LeenoUtils.no_refresh
def clean_text(desc):
    """
    Pulisce il testo da caratteri non stampabili e HTML entities.
    CORRETTO: rimosso il bug che inseriva \n tra ogni carattere
    """
    # Rimuove caratteri non stampabili (mantiene \n, \r, \t normali)
    desc = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', desc)

    sostituzioni = {
        # Entità HTML per lettere accentate
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

        # Tab diventa spazio
        '\t': ' ',

        # Correzioni encoding errati
        'Ã¨': 'è',
        'Â°': '°',
        'Ã': 'à',
        'Ó': 'à',
        'Þ': 'é',

        # Rimuovi caratteri speciali
        ' $': '',

        # Entità XML/HTML varie
        '&#x13;': '',
        '&#xD;&#xA;': '\n',  # CRLF Windows
        '&#xA;': '\n',        # LF Unix
        '&apos;': "'",
        '&#x3;&#x1;': '',

        # Correzioni trattini
        '- -': '- ',
        '—': '-',  # em dash
        '–': '-',  # en dash
        '\n- -': '\n-',

        # Pulizia spazi con newline
        '\n \n': '\n',
        '\n ': '\n',

        # RIMOSSO: '': '\n',  <-- QUESTO ERA IL BUG!
        # Sostituire stringa vuota con \n significa sostituire OGNI carattere con \n
    }

    # Esegue tutte le sostituzioni
    for old, new in sostituzioni.items():
        desc = desc.replace(old, new)

    # Rimuove spazi multipli con una singola regex
    desc = re.sub(r' +', ' ', desc)

    # Rimuove righe vuote multiple (max 2 newline consecutivi)
    desc = re.sub(r'\n{3,}', '\n\n', desc)

    # Rimuove righe contenenti solo trattini
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

@LeenoUtils.no_refresh
@with_undo("Pulizia testo selezione")
def sistema_cose():
    '''
    Ripulisce il testo da capoversi, spazi multipli e cattive codifiche.
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
        for el in reversed(range(y[0], y[1] + 1)):
            lista_y.append(el)

    for y in lista_y:
        cell = oSheet.getCellByPosition(lcol, y)
        if cell.Type.value == 'TEXT':
            cell.String = clean_text(cell.String)

    Menu_adattaAltezzaRiga()


########################################################################

@LeenoUtils.no_refresh
def descrizione_in_una_colonna(flag=False):
    '''
    Questa funzione consente di estendere su più colonne o ridurre ad una colonna lo spazio
    occupato dalla descrizione di voce in COMPUTO, VARIANTE e CONTABILITA.

    Args:
        flag (bool, optional): Se True, effettua l'unione delle celle. Se False, annulla l'unione.

    '''
    oDoc = LeenoUtils.getDocument()

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

    Menu_adattaAltezzaRiga()


########################################################################

@LeenoUtils.no_refresh
def MENU_numera_colonna():
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
@LeenoUtils.no_refresh
@with_undo
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
    committente = "\n\nCommittente: " + oDoc.getSheets().getByName('S2').getCellRangeByName("C6").String
    luogo = '\n' + oSheet.Name
    if oSheet.Name == 'COMPUTO':
        luogo = '\n\nComputo Metrico Estimativo'
    elif oSheet.Name == 'VARIANTE':
        luogo = '\n\nPerizia di Variante'

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
            oHeader.LeftText.Text.String = committente
            oHeader.LeftText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            oHeader.LeftText.Text.Text.CharHeight = htxt

            oHeader.CenterText.Text.String = oggetto
            oHeader.CenterText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            oHeader.CenterText.Text.Text.CharHeight = htxt

            oHeader.RightText.Text.String = luogo
            oHeader.RightText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            oHeader.RightText.Text.Text.CharHeight = htxt

            oAktPage.RightPageHeaderContent = oHeader
            # FOOTER
            oFooter = oAktPage.RightPageFooterContent
            oFooter.CenterText.Text.String = ''
            nomefile = oDoc.getURL().replace('%20',' ')
            oFooter.LeftText.Text.String = "\nrealizzato con LeenO: " + os.path.basename(nomefile)
            oFooter.LeftText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            oFooter.LeftText.Text.Text.CharHeight = htxt * 0.5
            oFooter.RightText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            oFooter.RightText.Text.Text.CharHeight = htxt
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
@LeenoUtils.preserva_posizione(step=0)
@LeenoUtils.no_refresh
def fissa():
    '''
    Fissa le righe e le colonne nel foglio attivo,
    evitando che le prime righe rimangano nascoste.
    '''
    oDoc = LeenoUtils.getDocument()
    oDoc.CurrentController.freezeAtPosition(0, 0) # Rimuove eventuali blocchi esistenti

    # Riporta la vista all'inizio per evitare che le righe
    # rimangano intrappolate sopra l'area di blocco
    oDoc.CurrentController.setFirstVisibleColumn(0)
    oDoc.CurrentController.setFirstVisibleRow(0)

    oSheet = oDoc.CurrentController.ActiveSheet

    if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA', 'Elenco Prezzi'):
        # Blocca sopra la riga 3 (quindi righe 0, 1, 2 fissate)
        oDoc.CurrentController.freezeAtPosition(0, 3)
    elif oSheet.Name in ('Analisi di Prezzo'):
        oDoc.CurrentController.freezeAtPosition(0, 2)
    elif oSheet.Name in ('Registro', 'SAL', 'S2'):
        oDoc.CurrentController.freezeAtPosition(0, 1)

########################################################################

@LeenoUtils.no_refresh
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
@LeenoUtils.no_refresh
def trova_np():
    '''
    Raggruppa le righe in modo da rendere evidenti i nuovi prezzi
    e aggiunge il prefisso NPxx_ al loro codice (se confermato dall'utente),
    propagando la modifica in tutti i fogli.
    Se il codice ha già il prefisso VDS_, mantienilo e aggiungi NPxx_ dopo.
    Rileva la presenza di prefissi NPxx_ esistenti e informa l'utente.
    Permette di scegliere se confrontare COMPUTO con VARIANTE o con CONTABILITÀ
    oppure VARIANTE con CONTABILITÀ
    '''
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    lrow = SheetUtils.getUsedArea(oSheet).EndRow
    confronto_nome = DLG.ScegliElaborato(Titolo = "Individuazione Nuovi Prezzi - confronto:", flag = "parallelo")

    if confronto_nome == 'computo_variante':
        confronto_col = 20
    elif confronto_nome == 'computo_contabilità' or confronto_nome == 'variante_contabilità':
        confronto_col = 21
    else:
        # indicator.end()
        return


    # Chiedi all'utente se vuole aggiungere il prefisso NP
    if Dialogs.YesNoDialog(IconType="question",Title='Nuovi Prezzi',
        Text='Vuoi aggiungere il prefisso "NPxx_"\n\nai codici dei nuovi prezzi?') == 1:
        add_prefix = True
    else:
        add_prefix = False

    indicator = oDoc.getCurrentController().getStatusIndicator()
    indicator.start('Elaborazione in corso...', 5)
    indicator.Value = 1

    struttura_off()
    genera_sommario()

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
    indicator.Value = 2

    for el in range(3, lrow):
        # i += 1
        if confronto_nome == 'variante_contabilità':
            val_base = oSheet.getCellByPosition(20, el).Value
        else:
            val_base = oSheet.getCellByPosition(19, el).Value

        val_confronto = oSheet.getCellByPosition(confronto_col, el).Value

        if val_base == 0 and val_confronto > 0:
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
            if confronto_col == 21:
                oSheet.getCellByPosition(confronto_col, el).CellBackColor = 16770000
                cont += val_confronto
            else:
                oSheet.getCellByPosition(confronto_col, el).CellBackColor = 16770000
                var += val_confronto
        else:
            oCellRangeAddr.StartRow = el
            oCellRangeAddr.EndRow = el
            oSheet.group(oCellRangeAddr, 1)
            oSheet.getCellRangeByPosition(0, el, 1, el).Rows.IsVisible = False

    indicator.Value = 3
    if add_prefix and code_mappings:
        for sheet_index, sheet in enumerate(oDoc.Sheets):
            if sheet.Name != oSheet.Name:
                used_range = SheetUtils.getUsedArea(sheet)
                if used_range:
                    # Determina la colonna corretta in base al nome del foglio
                    search_col = 1 if sheet.Name in ["COMPUTO", "VARIANTE", "CONTABILITA"] else 0

                    for row in range(used_range.StartRow, used_range.EndRow + 1):
                        cell_value = sheet.getCellByPosition(search_col, row).String
                        if cell_value in code_mappings:
                            sheet.getCellByPosition(search_col, row).String = code_mappings[cell_value]

    indicator.Value = 4
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
    indicator.Value = 5
    indicator.end()
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

    step = 0
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

@with_undo("Sposta Voci Selezionate")
def sposta_voce(lrow=None, msg=1):
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # 1. Identifica sorgente (Multivocale o Singola)
    oSel = oDoc.CurrentController.getSelection()
    if not oSel.supportsService("com.sun.star.table.CellRange"):
        return

    SR = LeenoComputo.circoscriveVoceComputo(oSheet, oSel.getRangeAddress().StartRow).RangeAddress.StartRow
    ER = LeenoComputo.circoscriveVoceComputo(oSheet, oSel.getRangeAddress().EndRow).RangeAddress.EndRow
    num_rows = ER - SR + 1

    # 2. Identifica destinazione
    to = basic_LeenO('ListenersSelectRange.getRange', "Seleziona destinazione")
    if not to: return

    try:
        to_row = int(to.split('$')[-1]) - 1
    except: return

    dest_row = LeenoSheetUtils.prossimaVoce(oSheet, to_row, 1, saltaCat=False)

    if SR <= dest_row <= ER + 1:
        Dialogs.Exclamation(Title="Errore", Text="Destinazione non valida.")
        return

    # --- OTTIMIZZAZIONE MASSIMA ---
    oDoc.addActionLock()
    oDoc.lockControllers()
    # Sospende il calcolo automatico delle formule (fondamentale in file pesanti)
    # oDoc.calculateAll()
    # oDoc.enableAutomaticCalculation(False)
    with LeenoUtils.DocumentRefreshContext(False):
    # try:
        # Inserimento righe vuote a destinazione
        oSheet.getRows().insertByIndex(dest_row, num_rows)

        # Ricalcolo posizione sorgente se slittata
        actual_SR = SR + num_rows if dest_row <= SR else SR

        # Definiamo gli indirizzi per lo spostamento
        oRangeAddress = oSheet.getCellRangeByPosition(0, actual_SR, 250, actual_SR + num_rows - 1).getRangeAddress()
        oCellAddress = oSheet.getCellByPosition(0, dest_row).getCellAddress()

        # Spostamento fisico dei dati
        oSheet.moveRange(oCellAddress, oRangeAddress)

        # Rimozione vecchie righe
        oSheet.getRows().removeByIndex(actual_SR, num_rows)

    # finally:
        # Riattiviamo il motore di calcolo e il layout
        # oDoc.enableAutomaticCalculation(True)
        # oDoc.unlockControllers()
        # oDoc.removeActionLock()

    # 3. Altezza ottimale (eseguita DOPO lo sblocco del controller)
    # Applichiamo l'altezza ottimale solo alle righe coinvolte nello spostamento
    dest_range = oSheet.getCellRangeByPosition(0, dest_row, 250, dest_row + num_rows - 1)
    dest_range.getRows().OptimalHeight = True

    # Ripristino visivo
    oDoc.CurrentController.setFirstVisibleRow(max(0, dest_row - 8))
    _gotoCella(1, dest_row + 1)

    # Ricalcolo finale necessario dopo aver riabilitato il calcolo automatico
    oDoc.calculateAll()
    numera_voci()

def copia_stili_celle(sheet_src, range_src, sheet_dest, range_dest):
    '''
    Copia gli stili di un range da un foglio all'altro.
    '''
    cell_range_src = sheet_src.getCellRangeByName(range_src)
    cell_range_dest = sheet_dest.getCellRangeByName(range_dest)

    # Ottieni dimensioni range
    rows = cell_range_src.RangeAddress.EndRow - cell_range_src.RangeAddress.StartRow + 1
    cols = cell_range_src.RangeAddress.EndColumn - cell_range_src.RangeAddress.StartColumn + 1

    # attiva la progressbar
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
    return

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
@LeenoUtils.no_refresh
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
@with_undo("Riordina Analisi di Prezzo Alfabetico")
@LeenoUtils.no_refresh
########################################################################
def Main_Riordina_Analisi_Alfabetico():
    chiudi_dialoghi()
    with LeenoUtils.DocumentRefreshContext(False):
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.Sheets.getByName("Analisi di Prezzo")
        lLastUrow = SheetUtils.getLastUsedRow(oSheet) + 1

        # Raccogli tutte le schede con le loro posizioni
        schede = []  # Lista di tuple: (codice, riga_inizio, riga_fine)
        i = 0
        while i <= lLastUrow:
            cell = oSheet.getCellByPosition(0, i)
            # Verifica se è l'inizio di una scheda
            if cell.CellStyle == "An.1v-Att Start":
                inizio = i
                # La cella con il codice è sempre la seconda cella (riga successiva)
                codice = oSheet.getCellByPosition(0, i + 1).String
                # Trova la fine della scheda (cerca "----")
                fine = None
                for x in range(i + 1, lLastUrow + 1):
                    if oSheet.getCellByPosition(0, x).String == "----":
                        fine = x
                        break
                if fine is None:
                    msg = f"Errore: scheda '{codice}' non ha riga di fine '----'"
                    DLG.chi(msg)
                    return
                # Verifica doppioni
                if any(s[0] == codice for s in schede):
                    msg = f"Mi fermo! Il codice:\n\t\t\t\t\t\t{codice}\nè presente più volte. Correggi e ripeti il comando."
                    DLG.chi(msg)
                    return
                schede.append((codice, inizio, fine))
                # Salta alla riga dopo la fine
                i = fine + 1
            else:
                i += 1

        if not schede:
            return

        # Crea una lista ordinata dei codici
        schede_ordinate = sorted(schede, key=lambda x: x[0])

        # Verifica se le schede sono già in ordine
        gia_ordinato = True
        for i in range(len(schede)):
            if schede[i][0] != schede_ordinate[i][0]:
                gia_ordinato = False
                break

        if gia_ordinato:
            Dialogs.Info(Title = 'Informazione', Text = "Le Analisi di Prezzo sono già in ordine alfabetico.")
            return

        struttura_off()
        # Sposta le schede nell'ordine corretto
        lrowDest = schede[0][1]  # Posizione della prima scheda originale

        for codice_target, _, _ in schede_ordinate:
            # Trova la scheda nella posizione attuale
            trovata = False
            inizio = None
            fine = None

            for i in range(lrowDest, SheetUtils.getLastUsedRow(oSheet) + 1):
                cell = oSheet.getCellByPosition(0, i)
                if cell.CellStyle == "An.1v-Att Start":
                    # Verifica che sia la scheda giusta controllando il codice
                    codice_trovato = oSheet.getCellByPosition(0, i + 1).String
                    if codice_trovato == codice_target:
                        inizio = i
                        # Trova la fine
                        for x in range(i + 1, SheetUtils.getLastUsedRow(oSheet) + 1):
                            if oSheet.getCellByPosition(0, x).String == "----":
                                fine = x
                                trovata = True
                                break
                        break

            if not trovata:
                continue

            # Se la scheda è già nella posizione corretta, aggiorna solo la destinazione
            if inizio == lrowDest:
                # Include la riga "Analisi_Sfondo" dopo "----"
                lrowDest = fine + 2
                continue

            # Calcola numero di righe (include "----" ma non "Analisi_Sfondo")
            nrighe = fine - inizio + 1

            # Inserisci spazio per la scheda nella posizione di destinazione
            oSheet.getRows().insertByIndex(lrowDest, nrighe + 1)  # +1 per "Analisi_Sfondo"

            # Aggiorna la posizione originale
            inizio += nrighe + 1
            fine += nrighe + 1

            # Copia la scheda (include "----")
            selezione = oSheet.getCellRangeByPosition(0, inizio, 250, fine)
            oDest = oSheet.getCellByPosition(0, lrowDest).CellAddress
            oSheet.copyRange(oDest, selezione.RangeAddress)

            # Copia la riga "Analisi_Sfondo" se presente
            if fine + 1 <= SheetUtils.getLastUsedRow(oSheet):
                riga_sfondo = oSheet.getCellRangeByPosition(0, fine + 1, 250, fine + 1)
                if oSheet.getCellByPosition(0, fine + 1).CellStyle == "Analisi_Sfondo":
                    oDest_sfondo = oSheet.getCellByPosition(0, lrowDest + nrighe).CellAddress
                    oSheet.copyRange(oDest_sfondo, riga_sfondo.RangeAddress)

            # Cancella la vecchia scheda (include "----" e "Analisi_Sfondo")
            oSheet.getRows().removeByIndex(inizio, nrighe + 1)

            # Aggiorna il punto di inserimento per la prossima scheda
            lrowDest += nrighe + 1
        MENU_struttura_on()
    Menu_adattaAltezzaRiga()


########################################################################

# @LeenoUtils.no_refresh  #questa va in errore
def MENU_export_selected_range_to_odt():
    """
    Esporta l'intervallo di celle selezionato in Calc in un nuovo documento Writer (ODT).
    Solo righe e colonne visibili, tabulazione a destra con puntini e paragrafi giustificati.
    """
    # try:
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
    # tab_position = left_margin + usable_width
    tab_position = page_width - right_margin

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
            if cell_value.startswith("VDS_"):
                cell_value = cell_value[4:]  # elimina i primi 4 caratteri

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
                        cell_value = f"{cell_value} (euro {converted}).\r"

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

    # except Exception as e:
    #     DLG.errore(f"Errore durante l'esportazione:\n{str(e)}")
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
    Ferma eventuali esecuzioni in corso e chiude tutti i dialoghi,
    restituendo 12 se invocata da un tasto Cancel.
    """
    try:
        ctx = LeenoUtils.getComponentContext()
        smgr = ctx.ServiceManager

        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
        frame = desktop.getCurrentFrame()

        # Interruzione "soft"
        dispatcher.executeDispatch(frame, ".uno:Cancel", "", 0, ())

        # Chiude eventuali dialoghi Python
        try:
            if hasattr(DLG, "chiudi_tutti"):
                DLG.chiudi_tutti()
            elif hasattr(DLG, "chi"):
                DLG.chi()
        except Exception:
            pass

        import gc
        gc.collect()

        # DLG.chi("✅ Tutti gli script interrotti e i dialoghi chiusi (Cancel).")
        return 12

    except Exception as e:
        DLG.chi(f"❌ Errore durante la chiusura: {e}")
        return 12
    return 12


########################################################################
########################################################################
########################################################################
def count_clipboard_lines():
        ctx = LeenoUtils.getComponentContext()
        smgr = ctx.getServiceManager()
        clip = smgr.createInstanceWithContext("com.sun.star.datatransfer.clipboard.SystemClipboard", ctx)

        # Ottiene il contenuto della clipboard
        transferable = clip.getContents()

        # Cerca il formato text/plain
        flavors = transferable.getTransferDataFlavors()

        text = None
        for flavor in flavors:
            if "text/plain" in flavor.MimeType:
                data = transferable.getTransferData(flavor)
                # data è già uno str in UTF-16 → usalo direttamente
                text = str(data)
                break

        if text is None:
            return 0

        # Conta le righe
        num_lines = len(text.splitlines())

        # Restituisce il valore
        return num_lines

def ApriFileDaFormula():
    ctx = LeenoUtils.getComponentContext()
    smgr = ctx.getServiceManager()
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # Documento attivo
    doc = LeenoUtils.getDocument()
    # if not doc supportsService("com.sun.star.sheet.SpreadsheetDocument"):
    #     return

    # Cella selezionata
    cell = doc.getCurrentSelection()

    if not cell.supportsService("com.sun.star.sheet.SheetCell"):
        Dialogs.Exclamation(Title= "Seleziona una singola cella con formula.", Text= "Errore")
        return

    formula = cell.getFormula()

    # Cerca un URL del tipo 'file:///...'
    match = re.search(r"'(file:///[^']+)'", formula)

    if not match:
        Dialogs.Exclamation (Title = 'ATTENZIONE!',
            Text="Nessun percorso file trovato nella formula.")
        return

    url = match.group(1)

    # Apre il file
    desktop.loadComponentFromURL(url, "_blank", 0, ())




#########################################################################
#########################################################################
#########################################################################


def _col_letter(col_num):
    """Converte numero colonna in lettera (0->A, 1->B, etc.)"""
    result = ""
    while col_num >= 0:
        result = chr(col_num % 26 + 65) + result
        col_num = col_num // 26 - 1
    return result




def _col_letter(col_num):
    """Converte numero colonna in lettera (0->A, 1->B, etc.)"""
    result = ""
    while col_num >= 0:
        result = chr(col_num % 26 + 65) + result
        col_num = col_num // 26 - 1
    return result

#########################################################################
#########################################################################
#########################################################################
@LeenoUtils.no_refresh
def somma_per_colore_nella_colonna():
    """
    Somma i valori nella colonna della cella attiva filtrando per
    Stile di Cella o Colore di sfondo.
    """
    oDoc = LeenoUtils.getDocument() #
    selection = oDoc.CurrentSelection

    # Recuperiamo i riferimenti della cella attiva
    target_color = selection.CellBackColor
    target_column = selection.RangeAddress.StartColumn
    sheet = oDoc.CurrentController.getActiveSheet()

    # Troviamo l'ultima riga usata per non scansionare l'intero foglio
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    last_row = cursor.RangeAddress.EndRow

    totale_somma = 0.0
    celle_contate = 0

    for riga in range(last_row + 1):
        cell = sheet.getCellByPosition(target_column, riga)

        if cell.CellBackColor == target_color:
            celle_contate += 1

            totale_somma += cell.Value


    # Formattazione italiana: punto per le migliaia, virgola per i decimali
    formattato = "{:,.2f}".format(totale_somma).replace(",", "X").replace(".", ",").replace("X", ".")

    messaggio = (
        f"Colonna: {target_column + 1}\n"
        f"Celle trovate: {celle_contate}\n\n"
        f"Somma totale per il colore selezionato: {formattato}"
    )

    Dialogs.Info(Title="Risultato Calcolo", Text=messaggio)



from Debug import measure_time, mostra_statistiche_performance, pulisci_log_performance, measure_time_simple
# @measure_time(show_popup=True)
@LeenoUtils.no_refresh
def MENU_debug():
    DLG.chi(LeenoConfig.PASTEL_COLORS)

    return


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
