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

from datetime import datetime, date
from xml.etree.ElementTree import Element, SubElement, tostring

import distutils.dir_util

import codecs
import subprocess
# import psutil
import re
import traceback
import threading
import time

import os
import shutil
import sys
import uno

import SheetUtils
import LeenoUtils
import LeenoSheetUtils
import LeenoToolbars as Toolbars
import LeenoFormat
import LeenoComputo

import LeenoConfig
cfg = LeenoConfig.Config()

import Dialogs

# cos'e' il namespace:
# http://www.html.it/articoli/il-misterioso-mondo-dei-namespaces-1/

# from com.sun.star.lang import Locale
from com.sun.star.beans import PropertyValue
# from com.sun.star.table.CellContentType import TEXT, EMPTY, VALUE, FORMULA
from com.sun.star.sheet.CellFlags import \
    VALUE, DATETIME, STRING, ANNOTATION, FORMULA, HARDATTR, OBJECTS, EDITATTR, FORMATTED

from com.sun.star.sheet.GeneralFunction import MAX

from com.sun.star.beans.PropertyAttribute import \
    MAYBEVOID, REMOVEABLE, MAYBEDEFAULT

########################################################################
# https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=27805&p=127383

########################################################################
# IMPORT DEI MODULI SEPARATI DI LEENO
########################################################################
import LeenoDialogs as DLG


def basic_LeenO(funcname, *args):
    '''Richiama funzioni definite in Basic'''

    xCompCont = LeenoUtils.getComponentContext()
    sm = xCompCont.ServiceManager
    mspf = sm.createInstance("com.sun.star.script.provider.MasterScriptProviderFactory")
    scriptPro = mspf.createScriptProvider("")
    Xscript = scriptPro.getScript(
        "vnd.sun.star.script:UltimusFree2." +
        funcname +
        "?language=Basic&location=application")
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
    # ~for nome in ('M1', 'S1', 'S2', 'S4', 'S5', 'Elenco Prezzi', 'COMPUTO'):
    for nome in ('M1', 'S1', 'S2', 'S5', 'Elenco Prezzi', 'COMPUTO'):
        oSheets.remove(nome)
    for nome in oSheets:
        oSheet = oDoc.getSheets().getByName(nome)
        if not oSheet.IsVisible:
            oDlg_config.getControl('CheckBox2').State = 0
            test = 0
            break
        oDlg_config.getControl('CheckBox2').State = 1
        test = 1
    if oDoc.getSheets().getByName("copyright_LeenO").IsVisible:
        oDlg_config.getControl('CheckBox2').State = 1
    if cfg.read('Generale', 'pesca_auto') == '1':
        oDlg_config.getControl('CheckBox1').State = 1  # pesca codice automatico
    if cfg.read('Generale', 'toolbar_contestuali') == '1':
        oDlg_config.getControl('CheckBox6').State = 1

    oSheet = oDoc.getSheets().getByName('S5')
    # descrizione_in_una_colonna
    if not oSheet.getCellRangeByName('C9').IsMerged:
        oDlg_config.getControl('CheckBox5').State = 1
    else:
        oDlg_config.getControl('CheckBox5').State = 0

    #  if conf.read(path_conf, 'Generale', 'descrizione_in_una_colonna') == '1': oDlg_config.getControl('CheckBox5').State = 1

    sString = oDlg_config.getControl('TextField1')
    sString.Text = cfg.read('Generale', 'altezza_celle')

    #  sString = oDlg_config.getControl("ComboBox1")
    #  sString.Text = conf.read(path_conf, 'Generale', 'visualizza') #visualizza all'avvio

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
    sString = oDlg_config.getControl('TextField4')
    sString.Text = oSheet.getCellRangeByName(
        'S1.H335').String  # cont_inizio_voci_abbreviate
    if oDoc.NamedRanges.hasByName("#Lib#1"):
        sString.setEnable(False)
    sString = oDlg_config.getControl('TextField12')
    sString.Text = oSheet.getCellRangeByName(
        'S1.H336').String  # cont_fine_voci_abbreviate
    if oDoc.NamedRanges.hasByName("#Lib#1"):
        sString.setEnable(False)

    if cfg.read('Generale', 'torna_a_ep') == '1':
        oDlg_config.getControl('CheckBox8').State = 1

    # Contabilità abilita
    if oSheet.getCellRangeByName('S1.H328').Value == 1:
        oDlg_config.getControl('CheckBox7').State = 1
    sString = oDlg_config.getControl('TextField13')
    if cfg.read('Contabilità', 'idxsal') == '&273.Dlg_config.TextField13.Text':
        sString.Text = '20'
    else:
        sString.Text = cfg.read('Contabilità', 'idxsal')
        if sString.Text == '':
            sString.Text = '20'
    sString = oDlg_config.getControl('ComboBox3')
    sString.Text = cfg.read('Contabilità', 'ricicla_da')

    sString = oDlg_config.getControl('ComboBox4')
    sString.Text = cfg.read('Generale', 'copie_backup')
    sString = oDlg_config.getControl('TextField5')
    sString.Text = cfg.read('Generale', 'pausa_backup')

    # MOSTRA IL DIALOGO
    oDlg_config.execute()

    if oDlg_config.getControl('CheckBox2').State != test:
        if oDlg_config.getControl('CheckBox2').State == 1:
            show_sheets(True)
        else:
            show_sheets(False)

    if oDlg_config.getControl('CheckBox3').State == 1:
        Toolbars.Switch(False)
    else:
        Toolbars.Switch(True)

    #  conf.write(path_conf, 'Generale', 'visualizza', oDlg_config.getControl('ComboBox1').getText())

    ctx = LeenoUtils.getComponentContext()
    oGSheetSettings = ctx.ServiceManager.createInstanceWithContext("com.sun.star.sheet.GlobalSheetSettings", ctx)
    if oDlg_config.getControl('ComboBox2').getText() == 'IN BASSO':
        cfg.write('Generale', 'movedirection', '0')
        oGSheetSettings.MoveDirection = 0
    else:
        cfg.write('Generale', 'movedirection', '1')
        oGSheetSettings.MoveDirection = 1
    cfg.write('Generale', 'altezza_celle', oDlg_config.getControl('TextField1').getText())

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

    if oDlg_config.getControl('TextField4').getText() != '10000':
        cfg.write('Contabilità', 'cont_inizio_voci_abbreviate', oDlg_config.getControl('TextField4').getText())
    oSheet.getCellRangeByName('S1.H335').Value = float(oDlg_config.getControl('TextField4').getText())

    if oDlg_config.getControl('TextField12').getText() != '10000':
        cfg.write('Contabilità', 'cont_fine_voci_abbreviate', oDlg_config.getControl('TextField12').getText())
    oSheet.getCellRangeByName('S1.H336').Value = float(oDlg_config.getControl('TextField12').getText())
    adatta_altezza_riga()

    cfg.write('Contabilità', 'abilita', str(oDlg_config.getControl('CheckBox7').State))
    cfg.write('Contabilità', 'idxsal', oDlg_config.getControl('TextField13').getText())
    if oDlg_config.getControl('ComboBox3').getText() in ('COMPUTO', '&305.Dlg_config.ComboBox3.Text'):
        cfg.write('Contabilità', 'ricicla_da', 'COMPUTO')
    else:
        cfg.write('Contabilità', 'ricicla_da', 'VARIANTE')
    cfg.write('Generale', 'copie_backup', oDlg_config.getControl('ComboBox4').getText())
    cfg.write('Generale', 'pausa_backup', oDlg_config.getControl('TextField5').getText())
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


class New_file:
    def __init__(self):  # , computo):
        pass

    def computo(arg=1):
        '''arg  { integer } : 1 mostra il dialogo di salvataggio file'''
        desktop = LeenoUtils.getDesktop()
        opz = PropertyValue()
        opz.Name = 'AsTemplate'
        opz.Value = True
        document = desktop.loadComponentFromURL(
            LeenO_path() + '/template/leeno/Computo_LeenO.ots', "_blank", 0,
            (opz, ))
        autoexec()
        if arg == 1:
            DLG.MsgBox(
                '''Prima di procedere è consigliabile salvare il lavoro.
Provvedi subito a dare un nome al file di computo...''',
                'Dai un nome al file...')
            salva_come()
            DlgMain()
        return document

    def usobollo():
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
    New_file.computo()


########################################################################


def MENU_nuovo_usobollo():
    '''Crea un nuovo documento in formato uso bollo.'''
    New_file.usobollo()


########################################################################


def MENU_invia_voce():
    '''
    Invia le voci di computo, elenco prezzi e analisi, con costi elementari,
    dal documento corrente al Documento Principale.
    '''
    # ~refresh(0)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    stili_computo = LeenoUtils.getGlobalVar('stili_computo')

    nSheet = oSheet.Name
    fpartenza = uno.fileUrlToSystemPath(oDoc.getURL())
    if fpartenza == LeenoUtils.getGlobalVar('sUltimus'):
        DLG.MsgBox("Questo file coincide con il Documento Principale (DP).", "Attenzione!")
        return
    elif LeenoUtils.getGlobalVar('sUltimus') == '':
        DLG.MsgBox("E' necessario impostare il Documento Principale (DP).", "Attenzione!")
        return
    nSheetDCC = getDCCSheet()
    lrow = LeggiPosizioneCorrente()[1]

    def getAnalisi(oSheet):
        try:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
        except AttributeError:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
        el_y = list()
        try:
            len(oRangeAddress)
            for el in oRangeAddress:
                el_y.append((el.StartRow, el.EndRow))
        except TypeError:
            el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))
        lista = list()
        for y in el_y:
            for el in range(y[0], y[1] + 1):
                lista.append(el)
        analisi = list()
        for y in lista:
            if oSheet.getCellByPosition(1, y).Type.value == 'FORMULA':
                analisi.append(oSheet.getCellByPosition(0, y).String)
        return (analisi, lista)

    def Circoscrive_Analisi(lrow):
        # oDoc = LeenoUtils.getDocument()
        # oSheet = oDoc.CurrentController.ActiveSheet
        stili_analisi = LeenoUtils.getGlobalVar('stili_analisi')
        if oSheet.getCellByPosition(0, lrow).CellStyle in stili_analisi:
            for el in reversed(range(0, lrow)):
                if oSheet.getCellByPosition(0,
                                            el).CellStyle == 'Analisi_Sfondo':
                    SR = el
                    break
            for el in range(lrow, SheetUtils.getUsedArea(oSheet).EndRow):
                if oSheet.getCellByPosition(
                        0, el).CellStyle == 'An-sfondo-basso Att End':
                    ER = el
                    break
        celle = oSheet.getCellRangeByPosition(0, SR, 250, ER)
        return celle

    # partenza
    if oSheet.Name == 'Elenco Prezzi':
        if oSheet.getCellByPosition(
                0,
                LeggiPosizioneCorrente()[1]).CellStyle not in ('EP-Cs', 'EP-aS'):
            DLG.MsgBox('La posizione di PARTENZA non è corretta.', 'ATTENZIONE!')
            return
        analisi = getAnalisi(oSheet)[0]
        lrow = getAnalisi(oSheet)[1][0]
        LeenoUtils.setGlobalVar('cod', oSheet.getCellByPosition(0, lrow).String)
        lista = getAnalisi(oSheet)[1]

        selezione = list()
        voci = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")
        for y in lista:
            rangen = oSheet.getCellRangeByPosition(0, y, 100, y).RangeAddress
            selezione.append(rangen)
        voci.addRangeAddresses(selezione, True)

        coppia = list()

        if analisi:
            GotoSheet('Analisi di Prezzo')
            oSheet = oDoc.getSheets().getByName('Analisi di Prezzo')

            ranges = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")
            selezione_analisi = list()
            for el in analisi:
                y = SheetUtils.uFindStringCol(el, 0, oSheet)
                sStRange = Circoscrive_Analisi(y)
                SR = sStRange.RangeAddress.StartRow
                ER = sStRange.RangeAddress.EndRow
                coppia.append((SR, ER))
                selezione_analisi.append(sStRange.RangeAddress)
            costi = list()
            for el in coppia:
                for y in range(el[0], el[1]):
                    if oSheet.getCellByPosition(0, y).CellStyle == 'An-lavoraz-Cod-sx' and \
                       oSheet.getCellByPosition(0, y).Type.value != 'EMPTY':
                        costi.append(oSheet.getCellByPosition(0, y).String)
            if len(costi) > 0:
                GotoSheet('Elenco Prezzi')
                oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
                el_y = list()
                for el in costi:
                    el_y.append(SheetUtils.uFindStringCol(el, 0, oSheet))
                for y in el_y:
                    rangen = oSheet.getCellRangeByPosition(0, y, 100,
                                                           y).RangeAddress
                    selezione.append(rangen)
                voci.addRangeAddresses(selezione, True)
        oDoc.CurrentController.select(voci)
        copy_clip()
        oDoc.CurrentController.select(
            oDoc.createInstance(
                "com.sun.star.sheet.SheetCellRanges"))  # unselect
        _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
        ddcDoc = LeenoUtils.getDocument()
        dccSheet = ddcDoc.CurrentController.ActiveSheet
        nome = dccSheet.Name

        if nome in ('Elenco Prezzi'):
            ddcDoc.CurrentController.setActiveSheet(dccSheet)
            _gotoCella(0, 3)
            paste_clip(insCells=1)
            # EliminaVociDoppieElencoPrezzi()
        if nome in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
            dccSheet = ddcDoc.getSheets().getByName('Elenco Prezzi')
            dccSheet.IsVisible = True
            ddcDoc.CurrentController.setActiveSheet(dccSheet)
            _gotoCella(0, 3)
            paste_clip(insCells=1)
            # EliminaVociDoppieElencoPrezzi()
            _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
            ddcDoc = LeenoUtils.getDocument()
            GotoSheet(nome)
            dccSheet = ddcDoc.getSheets().getByName(nome)
            lrow = LeggiPosizioneCorrente()[1]
            if dccSheet.getCellByPosition(0, lrow).CellStyle in ('comp Int_colonna'):
                LeenoComputo.insertVoceComputoGrezza(dccSheet, lrow + 1)
                # @@ PROVVISORIO !!!
                _gotoCella(1, lrow + 1 + 1)

                numera_voci(1)
                lrow = LeggiPosizioneCorrente()[1]
            if dccSheet.getCellByPosition(
                    0, lrow).CellStyle in (stili_computo + ('comp Int_colonna', )):
                if codice_voce(lrow) in ('', 'Cod. Art.?'):
                    codice_voce(lrow, LeenoUtils.getGlobalVar('cod'))
                else:
                    LeenoComputo.ins_voce_computo()
                    GotoSheet(nome)
                    codice_voce(LeggiPosizioneCorrente()[1], LeenoUtils.getGlobalVar('cod'))
                if LeggiPosizioneCorrente()[1] > 20:
                    ddcDoc.CurrentController.setFirstVisibleColumn(0)
                    ddcDoc.CurrentController.setFirstVisibleRow(LeggiPosizioneCorrente()[1] - 5)
            else:
                return
    # partenza
    if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        # sopra = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
        LeenoUtils.setGlobalVar('cod', codice_voce(lrow))
        try:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
        except AttributeError:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
        try:
            SR = oRangeAddress.StartRow
            SR = LeenoComputo.circoscriveVoceComputo(oSheet, SR).RangeAddress.StartRow
        except AttributeError:
            DLG.MsgBox(
                'La selezione delle voci dal COMPUTO di partenza\ndeve essere contigua.',
                'ATTENZIONE!')
            return
        ER = oRangeAddress.EndRow
        ER = LeenoComputo.circoscriveVoceComputo(oSheet, ER).RangeAddress.EndRow
        oDoc.CurrentController.select(
            oSheet.getCellRangeByPosition(0, SR, 100, ER))
        lista = list()
        for el in range(SR, ER + 1):
            if oSheet.getCellByPosition(
                    0, el).CellStyle in ('Comp Start Attributo'):
                lista.append(codice_voce(el))
        # seleziona()
        if nSheetDCC in ('Analisi di Prezzo'):
            DLG.MsgBox('Il foglio di destinazione non è corretto.', 'ATTENZIONE!')
            oDoc.CurrentController.select(
                oDoc.createInstance(
                    "com.sun.star.sheet.SheetCellRanges"))  # unselect
            return
        if nSheetDCC in ('COMPUTO', 'VARIANTE'):
            copy_clip()
            _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
            ddcDoc = LeenoUtils.getDocument()
            dccSheet = ddcDoc.getSheets().getByName(nSheet)
            lrow = LeggiPosizioneCorrente()[1]
            if dccSheet.getCellByPosition(
                    0, lrow).CellStyle in ('comp Int_colonna', ):
                lrow = LeggiPosizioneCorrente()[1] + 1
            elif dccSheet.getCellByPosition(
                    0, lrow).CellStyle not in stili_computo:
                DLG.MsgBox('La posizione di destinazione non è corretta.', 'ATTENZIONE!')
                oDoc.CurrentController.select(
                    oDoc.createInstance(
                        "com.sun.star.sheet.SheetCellRanges"))  # unselect
                return
            else:
                lrow = next_voice(LeggiPosizioneCorrente()[1], 1)
            _gotoCella(0, lrow)
            paste_clip(insCells=1)
            numera_voci(1)
            last = lrow + ER - SR + 1
            while lrow < last:
                rigenera_voce(lrow)
                lrow = next_voice(lrow, 1)
            # torno su partenza per prendere i prezzi
            _gotoDoc(fpartenza)
            oDoc = LeenoUtils.getDocument()
            GotoSheet('Elenco Prezzi')
            oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
            selezione = list()
            ranges = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")
            for el in lista:
                y = SheetUtils.uFindStringCol(el, 0, oSheet)
                rangen = oSheet.getCellRangeByPosition(0, y, 100,
                                                       y).RangeAddress
                selezione.append(rangen)

            ranges.addRangeAddresses(selezione, True)
            oDoc.CurrentController.select(ranges)
            copy_clip()
            #
            _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
            ddcDoc = LeenoUtils.getDocument()
            dccSheet = ddcDoc.getSheets().getByName('Elenco Prezzi')
            GotoSheet('Elenco Prezzi')
            _gotoCella(0, 4)
            paste_clip(insCells=1)
            # EliminaVociDoppieElencoPrezzi()
        if nSheetDCC in ('Elenco Prezzi'):
            DLG.MsgBox("Non è possibile inviare voci da un COMPUTO all'Elenco Prezzi.")
            return
        oDoc.CurrentController.select(
            oDoc.createInstance(
                "com.sun.star.sheet.SheetCellRanges"))  # unselect

    try:
        len(analisi)

        selezione = list()
        lista = list()
        _gotoDoc(fpartenza)
        oDoc = LeenoUtils.getDocument()
        GotoSheet('Analisi di Prezzo')
        ranges = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")
        ranges.addRangeAddresses(selezione_analisi, True)
        oDoc.CurrentController.select(ranges)

        copy_clip()

        _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
        ddcDoc = LeenoUtils.getDocument()
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
    if DLG.DlgSiNo("Ricerco ed elimino le voci di prezzo duplicate?") == 2:
        EliminaVociDoppieElencoPrezzi()
    adatta_altezza_riga('Elenco Prezzi')
    GotoSheet(nSheetDCC)
    # ~refresh(1)


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
        sopra = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
    elif oSheet.Name in ('Analisi di Prezzo'):
        sopra = Circoscrive_Analisi(lrow).RangeAddress.StartRow + 1
    if cod is None:
        return oSheet.getCellByPosition(1, sopra + 1).String
    else:
        oSheet.getCellByPosition(1, sopra + 1).String = cod


# def getVoce(cod=None):
# oDoc = LeenoUtils.getDocument()
# oSheet = oDoc.CurrentController.ActiveSheet
# lrow = LeggiPosizioneCorrente()[1]
# sopra = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
# return oSheet.getCellByPosition(1, sopra+1).String
# def setVoce(cod):
# oDoc = LeenoUtils.getDocument()
# oSheet = oDoc.CurrentController.ActiveSheet
# lrow = LeggiPosizioneCorrente()[1]
# sopra = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
# oSheet.getCellByPosition(1, sopra+1).String = cod
########################################################################


def _gotoDoc(sUrl):
    '''
    sUrl  { string } : nome del file
    porta il focus su di un determinato documento
    '''
    sUrl = uno.systemPathToFileUrl(sUrl)
    if sys.platform == 'linux' or sys.platform == 'darwin':
        target = LeenoUtils.getDesktop().loadComponentFromURL(
            sUrl, "_default", 0, list())
        target.getCurrentController().Frame.ContainerWindow.toFront()
        target.getCurrentController().Frame.activate()
    elif sys.platform == 'win32':
        desktop = LeenoUtils.getDesktop()
        oFocus = uno.createUnoStruct('com.sun.star.awt.FocusEvent')
        target = desktop.loadComponentFromURL(sUrl, "_default", 0, list())
        target.getCurrentController().getFrame().focusGained(oFocus)
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
    return '/'.join(reversed(str(datetime.now()).split(' ')[0].split('-')))


########################################################################


def MENU_copia_sorgente_per_git():
    '''
    fa una copia della directory del codice nel repository locale ed apre una shell per la commit
    '''
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

    make_pack(bar=1)
    oxt_path = uno.fileUrlToSystemPath(LeenO_path())
    if sys.platform == 'linux' or sys.platform == 'darwin':
        dest = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt'
        if not os.path.exists(dest):
            try:
                dest = os.getenv(
                    "HOME") + '/' + src_oxt + '/leeno/src/Ultimus.oxt/'
                os.makedirs(dest)
                os.makedirs(os.getenv("HOME") + '/' + src_oxt + '/leeno/bin/')
                os.makedirs(os.getenv("HOME") + '/' + src_oxt + '/_SRC/OXT')
            except FileExistsError:
                pass

            comandi = 'cd ' + dest + ' && mate-terminal && gitk &'
        else:
            comandi = 'cd /media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt && mate-terminal && gitk &'
        if not processo('wish'):
            subprocess.Popen(comandi, shell=True, stdout=subprocess.PIPE)
    elif sys.platform == 'win32':
        if not os.path.exists('w:/_dwg/ULTIMUSFREE/_SRC/leeno/src/'):
            try:
                os.makedirs(
                    os.getenv("HOMEPATH") + '\\' + src_oxt +
                    '\\leeno\\src\\Ultimus.oxt\\')
            except FileExistsError:
                pass
            dest = os.getenv("HOMEDRIVE") + os.getenv(
                "HOMEPATH") + '\\' + src_oxt + '\\leeno\\src\\Ultimus.oxt\\'
        else:
            dest = 'w:/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt'
        subprocess.Popen(
            'w: && cd w:/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt && "C:/Program Files/Git/git-bash.exe"',
            shell=True,
            stdout=subprocess.PIPE)
    distutils.dir_util.copy_tree(oxt_path, dest)
    return


########################################################################


def MENU_avvia_IDE():
    '''
    Avvia la modifica di pyleeno.py con geany
    '''
    avvia_IDE()


def avvia_IDE():
    '''Avvia la modifica di pyleeno.py con geany'''
    basic_LeenO('file_gest.avvia_IDE')
    oDoc = LeenoUtils.getDocument()
    oLayout = oDoc.CurrentController.getFrame().LayoutManager
    oLayout.showElement(
        "private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV")
    if sys.platform == 'linux' or sys.platform == 'darwin':
        subprocess.Popen('caja ' + LeenO_path(),
                         shell=True,
                         stdout=subprocess.PIPE)
        subprocess.Popen('geany ' + LeenO_path() +
                         '/python/pythonpath/pyleeno.py',
                         shell=True,
                         stdout=subprocess.PIPE)
    elif sys.platform == 'win32':
        subprocess.Popen('explorer.exe ' +
                         uno.fileUrlToSystemPath(LeenO_path()),
                         shell=True,
                         stdout=subprocess.PIPE)
        subprocess.Popen('"C:/Program Files (x86)/Geany/bin/geany.exe" ' +
                         uno.fileUrlToSystemPath(LeenO_path()) +
                         '/python/pythonpath/pyleeno.py',
                         shell=True,
                         stdout=subprocess.PIPE)
    return


########################################################################


def MENU_Inser_SottoCapitolo():
    '''
    @@ DA DOCUMENTARE
    '''
    Inser_SottoCapitolo()


def Inser_SottoCapitolo():
    '''
    @@ DA DOCUMENTARE
    '''
    Ins_Categorie(2)


########################################################################


def Ins_Categorie(n):
    '''
    n    { int } : livello della categoria
    0 = SuperCategoria
    1 = Categoria
    2 = SubCategoria
    '''
    # datarif = datetime.now()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_contab = LeenoUtils.getGlobalVar('stili_contab')
    noVoce = LeenoUtils.getGlobalVar('noVoce')

    row = LeggiPosizioneCorrente()[1]
    if oSheet.getCellByPosition(0,row).CellStyle in stili_computo + stili_contab:
        lrow = next_voice(row, 1)
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
    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400
    if n == 0:
        LeenoSheetUtils.inserSuperCapitolo(oSheet, lrow, sString)
    elif n == 1:
        LeenoSheetUtils.inserCapitolo(oSheet, lrow, sString)
    elif n == 2:
        LeenoSheetUtils.inserSottoCapitolo(oSheet, lrow, sString)

    _gotoCella(2, lrow)
    Rinumera_TUTTI_Capitoli2()
    oDoc.CurrentController.ZoomValue = zoom
    oDoc.CurrentController.setFirstVisibleColumn(0)
    oDoc.CurrentController.setFirstVisibleRow(lrow - 5)
    # MsgBox('eseguita in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!','')


########################################################################


def MENU_Inser_SuperCapitolo():
    '''
    @@ DA DOCUMENTARE
    '''
    Inser_SuperCapitolo()


def Inser_SuperCapitolo():
    '''
    @@ DA DOCUMENTARE
    '''
    Ins_Categorie(0)

########################################################################

def MENU_Inser_Capitolo():
    '''
    @@ DA DOCUMENTARE
    '''
    Inser_Capitolo()


def Inser_Capitolo():
    '''
    @@ DA DOCUMENTARE
    '''
    Ins_Categorie(1)

########################################################################


def MENU_Rinumera_TUTTI_Capitoli2():
    Rinumera_TUTTI_Capitoli2()


def Rinumera_TUTTI_Capitoli2():
    Sincronizza_SottoCap_Tag_Capitolo_Cor()  # sistemo gli idcat voce per voce
    Tutti_Subtotali()  # ricalcola i totali di categorie e subcategorie


def Tutti_Subtotali():
    '''ricalcola i subtotali di categorie e subcategorie'''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
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
    oSheet.getCellByPosition(
        18, 1).Formula = '=SUBTOTAL(9;S4:S' + str(lrow + 1) + ')'
    oSheet.getCellByPosition(
        18, lrow).Formula = '=SUBTOTAL(9;S4:S' + str(lrow + 1) + ')'
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


def Sincronizza_SottoCap_Tag_Capitolo_Cor():
    '''
    lrow    { double } : id della riga di inserimento
    sincronizza il categoria e sottocategorie
    '''
    # datarif = datetime.now()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
#    lrow = LeggiPosizioneCorrente()[1]
    lastRow = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1

    listasbcat = list()
    listacat = list()
    listaspcat = list()
    for lrow in range(0, lastRow):
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

    # MsgBox('Importazione eseguita con successo\n in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!','')


########################################################################


def MENU_join_sheets():
    '''
    unisci fogli
    serve per unire tanti fogli in un unico foglio
    '''
    oDoc = LeenoUtils.getDocument()
    lista_fogli = oDoc.Sheets.ElementNames
    if not oDoc.getSheets().hasByName('unione_fogli'):
        sheet = oDoc.createInstance("com.sun.star.sheet.Spreadsheet")
        unione = oDoc.Sheets.insertByName('unione_fogli', sheet)
        unione = oDoc.getSheets().getByName('unione_fogli')
        for el in lista_fogli:
            oSheet = oDoc.getSheets().getByName(el)
            oRangeAddress = oSheet.getCellRangeByPosition(
                0, 0, (SheetUtils.getUsedArea(oSheet).EndColumn),
                (SheetUtils.getUsedArea(oSheet).EndRow)).getRangeAddress()
            oCellAddress = unione.getCellByPosition(
                0,
                SheetUtils.getUsedArea(unione).EndRow + 1).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
        DLG.MsgBox('Unione dei fogli eseguita.', 'Avviso')
    else:
        unione = oDoc.getSheets().getByName('unione_fogli')
        DLG.MsgBox('Il foglio "unione_fogli" è già esistente, quindi non procedo.', 'Avviso!')
    oDoc.CurrentController.setActiveSheet(unione)


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


def copia_sheet(nSheet, tag):
    '''
    nSheet   { string } : nome sheet
    tag      { string } : stringa di tag
    duplica copia sheet corrente di fianco a destra
    '''
    oDoc = LeenoUtils.getDocument()
    # nSheet = 'COMPUTO'
    oSheet = oDoc.getSheets().getByName(nSheet)
    idSheet = oSheet.RangeAddress.Sheet + 1
    if oDoc.getSheets().hasByName(nSheet + '_' + tag):
        DLG.MsgBox('La tabella di nome ' + nSheet + '_' + tag + 'è già presente.', 'ATTENZIONE! Impossibile procedere.')
        return
    else:
        oDoc.Sheets.copyByName(nSheet, nSheet + '_' + tag, idSheet)
        oSheet = oDoc.getSheets().getByName(nSheet + '_' + tag)
        oDoc.CurrentController.setActiveSheet(oSheet)
        # oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect


########################################################################


def Filtra_computo(nSheet, nCol, sString):
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
                lrow = next_voice(lrow, 0)
        except Exception:
            lrow = next_voice(lrow, 0)
    for lrow in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
        if(oSheet.getCellByPosition(18, lrow).CellStyle == 'Livello-1-scritta mini val' and
           oSheet.getCellByPosition(18, lrow).Value == 0 or
           oSheet.getCellByPosition(18, lrow).CellStyle == 'livello2 scritta mini' and
           oSheet.getCellByPosition(18, lrow).Value == 0):

            oSheet.getRows().removeByIndex(lrow, 1)

    # iCellAttr =(oDoc.createInstance("com.sun.star.sheet.CellFlags.OBJECTS"))
    flags = OBJECTS
    oSheet.getCellRangeByPosition(0, 0, 42, 0).clearContents(
        flags)  # cancello gli oggetti
    oDoc.CurrentController.select(oSheet.getCellByPosition(0, 3))
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect


########################################################################


def Filtra_Computo_Cap():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7, 8).String
    sString = oSheet.getCellByPosition(7, 10).String
    Filtra_computo(nSheet, 31, sString)


########################################################################


def Filtra_Computo_SottCap():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7, 8).String
    sString = oSheet.getCellByPosition(7, 12).String
    Filtra_computo(nSheet, 32, sString)


########################################################################


def Filtra_Computo_A():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7, 8).String
    sString = oSheet.getCellByPosition(7, 14).String
    Filtra_computo(nSheet, 33, sString)


########################################################################


def Filtra_Computo_B():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7, 8).String
    sString = oSheet.getCellByPosition(7, 16).String
    Filtra_computo(nSheet, 34, sString)


########################################################################


def Filtra_Computo_C():  # filtra in base al codice di prezzo
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.getCellByPosition(7, 8).String
    sString = oSheet.getCellByPosition(7, 20).String
    Filtra_computo(nSheet, 1, sString)


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
    # oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet
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
    Sprotegge e riordina tutti fogli del documento.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheets = oDoc.Sheets.ElementNames
    for nome in oSheets:
        oSheet = oDoc.getSheets().getByName(nome)
        oSheet.unprotect('')
    # riordino le sheet
    oDoc.Sheets.moveByName("Elenco Prezzi", 0)
    if oDoc.Sheets.hasByName("Analisi di Prezzo"):
        oDoc.Sheets.moveByName("Analisi di Prezzo", 1)
    oDoc.Sheets.moveByName("COMPUTO", 2)
    if oDoc.Sheets.hasByName("VARIANTE"):
        oDoc.Sheets.moveByName("VARIANTE", 3)
    if oDoc.Sheets.hasByName("CONTABILITA"):
        oDoc.Sheets.moveByName("CONTABILITA", 4)
    if oDoc.Sheets.hasByName("M1"):
        oDoc.Sheets.moveByName("M1", 5)
    oDoc.Sheets.moveByName("S1", 6)
    oDoc.Sheets.moveByName("S2", 7)
    # ~oDoc.Sheets.moveByName("S4", 9)
    if oDoc.Sheets.hasByName("S5"):
        oDoc.Sheets.moveByName("S5", 10)
    if oDoc.Sheets.hasByName("copyright_LeenO"):
        oDoc.Sheets.moveByName("copyright_LeenO", oDoc.Sheets.Count)


########################################################################


def setPreview(arg=0):
    '''
    colore   { integer } : id colore
    attribuisce al foglio corrente un colore a scelta
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet  # se questa dà errore, il preview è già attivo
    adatta_altezza_riga(oSheet.Name)
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    oProp = PropertyValue()
    properties = (oProp, )
    dispatchHelper.executeDispatch(oFrame, '.uno:PrintPreview', '', arg, properties)


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
    # oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet
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


#  ########################################################################


def adatta_altezza_riga(nSheet=None):
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoSheetUtils.adattaAltezzaRiga
    Adatta l'altezza delle righe al contenuto delle celle.

    nSheet   { string } : nSheet della sheet
    imposta l'altezza ottimale delle celle
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if not oDoc.getSheets().hasByName('S1'):
        return
    nSheet = oSheet.Name
    oDoc.getSheets().hasByName(nSheet)
    oSheet.getCellRangeByPosition(
        0, 0,
        SheetUtils.getUsedArea(oSheet).EndColumn,
        SheetUtils.getUsedArea(oSheet).EndRow).Rows.OptimalHeight = True
    if float(loVersion()[:5].replace(
            '.', '')) >= 642:  # DALLA VERSIONE 6.4.2 IL PROBLEMA è RISOLTO
        return
    #  se la versione di LibreOffice è maggiore della 5.2, esegue il comando agendo direttamente sullo stile
    lista_stili = ('comp 1-a', 'Comp-Bianche in mezzo Descr_R',
                   'Comp-Bianche in mezzo Descr', 'EP-a',
                   'Ultimus_centro_bordi_lati')
    # NELLE VERSIONI DA 5.4.2 A 6.4.1
    if float(loVersion()[:5].replace('.', '')) > 520 or float(
            loVersion()[:5].replace('.', '')) < 642:
        # chi(float(loVersion()[:5].replace('.','')))
        for stile_cella in lista_stili:
            try:
                oDoc.StyleFamilies.getByName("CellStyles").getByName(
                    stile_cella).IsTextWrapped = True
            except Exception:
                pass
        #  #if nSheet in('VARIANTE', 'COMPUTO', 'CONTABILITA', 'Richiesta offerta'):
        test = SheetUtils.getUsedArea(oSheet).EndRow + 1
        for y in range(0, test):
            if oSheet.getCellByPosition(2, y).CellStyle in lista_stili:
                oSheet.getCellRangeByPosition(
                    0, y,
                    SheetUtils.getUsedArea(oSheet).EndColumn,
                    y).Rows.OptimalHeight = True
    if oSheet.Name in ('Elenco Prezzi', 'VARIANTE', 'COMPUTO', 'CONTABILITA'):
        oSheet.getCellByPosition(0, 2).Rows.Height = 800
    if nSheet == 'Elenco Prezzi':
        test = SheetUtils.getUsedArea(oSheet).EndRow + 1
        for y in range(0, test):
            oSheet.getCellRangeByPosition(0, y,
                                          SheetUtils.getUsedArea(oSheet).EndColumn,
                                          y).Rows.OptimalHeight = True
    return


########################################################################


def voce_breve():
    '''
    Cambia il numero di caratteri visualizzati per la descrizione voce in COMPUTO,
    CONTABILITA E VARIANTE.
    '''
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.getCellRangeByPosition(
        0, 0,
        SheetUtils.getUsedArea(oSheet).EndColumn,
        SheetUtils.getUsedArea(oSheet).EndRow).Rows.OptimalHeight = True
    if not oDoc.getSheets().hasByName('S1'):
        return
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        oSheet = oDoc.getSheets().getByName('S1')
        if oSheet.getCellRangeByName('S1.H337').Value < 10000:
            cfg.write('Computo', 'inizio_voci_abbreviate', oSheet.getCellRangeByName('S1.H337').String)
            oSheet.getCellRangeByName('S1.H337').Value = 10000
        else:
            oSheet.getCellRangeByName('S1.H337').Value = int(
                cfg.read('Computo', 'inizio_voci_abbreviate'))
        if oSheet.getCellRangeByName('S1.H338').Value < 10000:
            cfg.write('Computo', 'fine_voci_abbreviate', oSheet.getCellRangeByName('S1.H338').String)
            oSheet.getCellRangeByName('S1.H338').Value = 10000
        else:
            oSheet.getCellRangeByName('S1.H338').Value = int(cfg.read('Computo', 'fine_voci_abbreviate'))
        adatta_altezza_riga()

    elif oSheet.Name == 'CONTABILITA':
        oSheet = oDoc.getSheets().getByName('S1')
        if oDoc.NamedRanges.hasByName("#Lib#1"):
            DLG.MsgBox(
                "Risulta già registrato un SAL. NON E' POSSIBILE PROCEDERE.",
                'ATTENZIONE!')
            return
        else:
            if oSheet.getCellRangeByName('S1.H335').Value < 10000:
                cfg.write('Contabilità', 'cont_inizio_voci_abbreviate', oSheet.getCellRangeByName('S1.H335').String)
                oSheet.getCellRangeByName('S1.H335').Value = 10000
            else:
                oSheet.getCellRangeByName('S1.H335').Value = int(cfg.read('Contabilità', 'cont_inizio_voci_abbreviate'))
            if oSheet.getCellRangeByName('S1.H336').Value < 10000:
                cfg.write('Contabilità', 'cont_fine_voci_abbreviate', oSheet.getCellRangeByName('S1.H336').String)
                oSheet.getCellRangeByName('S1.H336').Value = 10000
            else:
                oSheet.getCellRangeByName('S1.H336').Value = int(cfg.read('Contabilità', 'cont_fine_voci_abbreviate'))
            adatta_altezza_riga()


########################################################################


def cancella_voci_non_usate():
    '''
    Cancella le voci di prezzo non utilizzate.
    '''
    chiudi_dialoghi()
    #  oDialogo_attesa = dlg_attesa()
    #  attesa().start() #mostra il dialogo

    if DLG.DlgSiNo(
            '''Questo comando ripulisce l'Elenco Prezzi
dalle voci non utilizzate in nessuno degli altri elaborati.

LA PROCEDURA POTREBBE RICHIEDERE DEL TEMPO.

Vuoi procedere comunque?''', 'AVVISO!') == 3:
        #  oDialogo_attesa.endExecute() #chiude il dialogo
        return
    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(False)
    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400
    oDoc.enableAutomaticCalculation(False)
    oSheet = oDoc.CurrentController.ActiveSheet

    oRange = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    SR = oRange.StartRow + 1
    ER = oRange.EndRow + 1
    lista_prezzi = list()
    for n in range(SR, ER):
        lista_prezzi.append(oSheet.getCellByPosition(0, n).String)
    lista = list()
    for tab in ('COMPUTO', 'Analisi di Prezzo', 'VARIANTE', 'CONTABILITA'):
        try:
            oSheet = oDoc.getSheets().getByName(tab)
            if tab == 'Analisi di Prezzo':
                col = 0
            else:
                col = 1
            for el in lista_prezzi:
                if SheetUtils.uFindStringCol(el, col, oSheet):
                    lista.append(el)
        except Exception:
            pass
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')

    da_cancellare = set(lista_prezzi).difference(set(lista))
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet

    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, SR, 0, ER))
    struttura_off('R')
    struttura_off('R')
    struttura_off('R')
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
    for n in reversed(range(SR, ER)):
        if oSheet.getCellByPosition(0, n).String in da_cancellare:
            oSheet.Rows.removeByIndex(n, 1)
        if(oSheet.getCellByPosition(0, n).String == '' and
           oSheet.getCellByPosition(1, n).String == '' and
           oSheet.getCellByPosition(4, n).String == ''):
            oSheet.Rows.removeByIndex(n, 1)
    oDoc.enableAutomaticCalculation(True)
    oDoc.CurrentController.ZoomValue = zoom
    _gotoCella(0, 3)
    #  oDialogo_attesa.endExecute() #chiude il dialogo


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

    if not oSheet.getCellByPosition(1, 3).Rows.OptimalHeight:
        adatta_altezza_riga()
    else:
        hriga = oSheet.getCellRangeByName(
            'B4').CharHeight * 65 * 2 + 100  # visualizza tre righe
        oSheet.getCellRangeByPosition(0, SR, 0, ER).Rows.Height = hriga


########################################################################


def scelta_viste():
    '''
    Gestisce i dialoghi del menù viste nelle tabelle di Analisi di Prezzo,
    Elenco Prezzi, COMPUTO, VARIANTE, CONTABILITA'
    Genera i raffronti tra COMPUTO e VARIANTE e CONTABILITA'
    '''
    #  refresh(0)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    psm = LeenoUtils.getComponentContext().ServiceManager
    dp = psm.createInstance('com.sun.star.awt.DialogProvider')
    if oSheet.Name in ('VARIANTE', 'COMPUTO'):
        oDialog1 = dp.createDialog(
            'vnd.sun.star.script:UltimusFree2.DialogViste_A?language=Basic&location=application'
        )
        # oDialog1Model = oDialog1.Model
        oDialog1.getControl('Dettaglio').State = cfg.read('Generale', 'dettaglio')
        if oSheet.getColumns().getByIndex(5).Columns.IsVisible:
            oDialog1.getControl('CBMis').State = 1
        if oSheet.getColumns().getByIndex(17).Columns.IsVisible:
            oDialog1.getControl('CBSic').State = 1
        if oSheet.getColumns().getByIndex(28).Columns.IsVisible:
            oDialog1.getControl('CBMat').State = 1
        if oSheet.getColumns().getByIndex(29).Columns.IsVisible:
            oDialog1.getControl('CBMdo').State = 1
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

        oDialog1.execute()

        # il salvataggio anche su leeno.conf serve alla funzione voce_breve()
        if oDialog1.getControl('TextField10').getText() != '10000':
            cfg.write('Computo', 'inizio_voci_abbreviate', oDialog1.getControl('TextField10').getText())
        oDoc.getSheets().getByName('S1').getCellRangeByName('H337').Value = float(oDialog1.getControl('TextField10').getText())

        if oDialog1.getControl('TextField11').getText() != '10000':
            cfg.write('Computo', 'fine_voci_abbreviate', oDialog1.getControl('TextField11').getText())
        oDoc.getSheets().getByName('S1').getCellRangeByName('H338').Value = float(oDialog1.getControl('TextField11').getText())
        #  oDialog1.getControl('CBMdo').State = False
        #  if oSheet.getColumns().getByIndex(29).Columns.IsVisible:
        #  oDialog1.getControl('CBMdo').State = True

        if oDialog1.getControl('OBTerra').State:
            computo_terra_terra()
            oDialog1.getControl('CBSic').State = 0
            oDialog1.getControl('CBMdo').State = 0
            oDialog1.getControl('CBMat').State = 0
            oDialog1.getControl('CBCat').State = 0
            oDialog1.getControl('CBFig').State = 0
            oDialog1.getControl('CBMis').State = 1

        if oDialog1.getControl("CBMis").State == 0:  # misure
            oSheet.getColumns().getByIndex(5).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(6).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(7).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(8).Columns.IsVisible = False
        else:
            oSheet.getColumns().getByIndex(5).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(6).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(7).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(8).Columns.IsVisible = True

        if oDialog1.getControl('CBMdo').State:  # manodopera
            oSheet.getColumns().getByIndex(29).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(30).Columns.IsVisible = True
            oSheet.getColumns().getByIndex(5).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(6).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(7).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(8).Columns.IsVisible = False
            #  adatta_altezza_riga(oSheet)
            oSheet.clearOutline()
            struct(3)
        else:
            oSheet.getColumns().getByIndex(29).Columns.IsVisible = False
            oSheet.getColumns().getByIndex(30).Columns.IsVisible = False

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
    elif oSheet.Name in ('Elenco Prezzi'):
        oCellRangeAddr = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
        oDialog1 = dp.createDialog(
            "vnd.sun.star.script:UltimusFree2.DialogViste_EP?language=Basic&location=application"
        )
        # oDialog1Model = oDialog1.Model

        if oSheet.getColumns().getByIndex(3).Columns.IsVisible:
            oDialog1.getControl('CBSic').State = 1
        if oSheet.getColumns().getByIndex(5).Columns.IsVisible:
            oDialog1.getControl('CBMdo').State = 1
        if not oSheet.getCellByPosition(1, 3).Rows.OptimalHeight:
            oDialog1.getControl('CBDesc').State = 1
        if oSheet.getColumns().getByIndex(7).Columns.IsVisible:
            oDialog1.getControl('CBOrig').State = 1
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

            if oDialog1.getControl("CBDesc").State == 1:  # descrizione
                oSheet.getColumns().getByIndex(3).Columns.IsVisible = False
                oSheet.getCellByPosition(1, 3).Rows.OptimalHeight
                voce_breve_ep()
            #  elif oDialog1.getControl("CBDesc").State == 0: adatta_altezza_riga(oSheet.Name)

            if oDialog1.getControl("CBOrig").State == 0:  # origine
                oSheet.getColumns().getByIndex(7).Columns.IsVisible = False
            else:
                oSheet.getColumns().getByIndex(7).Columns.IsVisible = True

            if oDialog1.getControl("CBSom").State == 1:
                genera_sommario()

            oRangeAddress = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
            SR = oRangeAddress.StartRow + 1
            ER = oRangeAddress.EndRow  # -1

            oSheet.getCellRangeByPosition(11, 0, 26,
                                          0).Columns.IsVisible = True
            oSheet.getCellRangeByPosition(23, SR, 25,
                                          ER).CellStyle = 'EP statistiche'
            oSheet.getCellRangeByPosition(26, SR, 26,
                                          ER).CellStyle = 'EP-mezzo %'
            oSheet.getCellRangeByName('AA2').CellStyle = 'EP-mezzo %'
            formule = list()
            oSheet.getCellByPosition(11, 0).String = 'COMPUTO'
            oSheet.getCellByPosition(15, 0).String = 'VARIANTE'
            oSheet.getCellByPosition(19, 0).String = "CONTABILITA"
            if oDialog1.getControl("ComVar").State:  # Computo - Variante
                genera_sommario()
                oRangeAddress.StartColumn = 19
                oRangeAddress.EndColumn = 22

                oSheet.getCellByPosition(23, 0).String = 'COMPUTO - VARIANTE'
                for n in range(4, LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2):
                    formule.append([
                        '=IF(Q' + str(n) + '-M' + str(n) + '=0;"--";Q' +
                        str(n) + '-M' + str(n) + ')', '=IF(R' + str(n) + '-N' +
                        str(n) + '>0;R' + str(n) + '-N' + str(n) + ';"")',
                        '=IF(R' + str(n) + '-N' + str(n) + '<0;N' + str(n) +
                        '-R' + str(n) + ';"")', '=IFERROR(IFS(AND(N' + str(n) +
                        '>R' + str(n) + ';R' + str(n) + '=0);-1;AND(N' +
                        str(n) + '<R' + str(n) + ';N' + str(n) + '=0);1;N' +
                        str(n) + '=R' + str(n) + ';"--";N' + str(n) + '>R' +
                        str(n) + ';-(N' + str(n) + '-R' + str(n) + ')/N' +
                        str(n) + ';N' + str(n) + '<R' + str(n) + ';-(N' +
                        str(n) + '-R' + str(n) + ')/N' + str(n) + ');"--")'
                    ])
                n += 1
                oSheet.getCellByPosition(
                    26, 1
                ).Formula = '=IFERROR(IFS(AND(N2>R2;R2=0);-1;AND(N2<R2;N2=0);1;N2=R2;"--";N2>R2;-(N2-R2)/N2;N2<R2;-(N2-R2)/N2);"--")'
                oSheet.getCellByPosition(
                    26, ER
                ).Formula = '=IFERROR(IFS(AND(N' + str(n) + '>R' + str(
                    n) + ';R' + str(n) + '=0);-1;AND(N' + str(n) + '<R' + str(
                        n) + ';N' + str(n) + '=0);1;N' + str(n) + '=R' + str(
                            n) + ';"--";N' + str(n) + '>R' + str(
                                n) + ';-(N' + str(n) + '-R' + str(
                                    n) + ')/N' + str(n) + ';N' + str(
                                        n) + '<R' + str(n) + ';-(N' + str(
                                            n) + '-R' + str(n) + ')/N' + str(
                                                n) + ');"--")'
                oRange = oSheet.getCellRangeByPosition(23, 3, 26,
                                                       LeenoSheetUtils.cercaUltimaVoce(oSheet))
                formule = tuple(formule)
                oRange.setFormulaArray(formule)
                ###
                if oRangeAddress.StartColumn != 0:
                    oCellRangeAddr.StartColumn = 18
                    oCellRangeAddr.EndColumn = 21
                    oSheet.group(oCellRangeAddr, 0)
                    oSheet.getCellRangeByPosition(18, 0, 21,
                                                  0).Columns.IsVisible = False

                    oCellRangeAddr.StartColumn = 15
                    oCellRangeAddr.EndColumn = 15
                    oSheet.group(oCellRangeAddr, 0)
                    oSheet.getCellRangeByPosition(15, 0, 15,
                                                  0).Columns.IsVisible = False
                ###

            if oDialog1.getControl("ComCon").State:  # Computo - Contabilità
                genera_sommario()
                oRangeAddress.StartColumn = 15
                oRangeAddress.EndColumn = 18

                oSheet.getCellByPosition(23,
                                         0).String = 'COMPUTO - CONTABILITÀ'
                for n in range(4, LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2):
                    formule.append([
                        '=IF(U' + str(n) + '-M' + str(n) + '=0;"--";U' +
                        str(n) + '-M' + str(n) + ')', '=IF(V' + str(n) + '-N' +
                        str(n) + '>0;V' + str(n) + '-N' + str(n) + ';"")',
                        '=IF(V' + str(n) + '-N' + str(n) + '<0;N' + str(n) +
                        '-V' + str(n) + ';"")', '=IFERROR(IFS(AND(N' + str(n) +
                        '>V' + str(n) + ';V' + str(n) + '=0);-1;AND(N' +
                        str(n) + '<V' + str(n) + ';N' + str(n) + '=0);1;N' +
                        str(n) + '=V' + str(n) + ';"--";N' + str(n) + '>V' +
                        str(n) + ';-(N' + str(n) + '-V' + str(n) + ')/N' +
                        str(n) + ';N' + str(n) + '<V' + str(n) + ';-(N' +
                        str(n) + '-V' + str(n) + ')/N' + str(n) + ');"--")'
                    ])
                n += 1
                #  for el in(1, ER+1):
                oSheet.getCellByPosition(
                    26, 1
                ).Formula = '=IFERROR(IFS(AND(N2>V2;V2=0);-1;AND(N2<V2;N2=0);1;N2=V2;"--";N2>V2;-(N2-V2)/N2;N2<V2;-(N2-V2)/N2);"--")'
                oSheet.getCellByPosition(
                    26, ER
                ).Formula = '=IFERROR(IFS(AND(N' + str(n) + '>V' + str(
                    n) + ';V' + str(n) + '=0);-1;AND(N' + str(n) + '<V' + str(
                        n) + ';N' + str(n) + '=0);1;N' + str(n) + '=V' + str(
                            n) + ';"--";N' + str(n) + '>V' + str(
                                n) + ';-(N' + str(n) + '-V' + str(
                                    n) + ')/N' + str(n) + ';N' + str(
                                        n) + '<V' + str(n) + ';-(N' + str(
                                            n) + '-V' + str(n) + ')/N' + str(
                                                n) + ');"--")'
                oRange = oSheet.getCellRangeByPosition(23, 3, 26,
                                                       LeenoSheetUtils.cercaUltimaVoce(oSheet))
                formule = tuple(formule)
                oRange.setFormulaArray(formule)
                ###
                if oRangeAddress.StartColumn != 0:
                    # evidenzia le quantità eccedenti il VI/I
                    for el in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
                        if oSheet.getCellByPosition(
                                26,
                                el).Value >= 0.2 or oSheet.getCellByPosition(
                                    26, el).String == '20,00%':
                            oSheet.getCellRangeByPosition(
                                0, el, 25, el).CellBackColor = 16777062
                    #  oCellRangeAddr=oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
                    if DLG.DlgSiNo(
                            "Nascondo eventuali voci non ancora contabilizzate?"
                    ) == 2:
                        struttura_off()
                        for el in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
                            if oSheet.getCellByPosition(20, el).Value == 0:
                                oCellRangeAddr.StartRow = el
                                oCellRangeAddr.EndRow = el
                                oSheet.group(oCellRangeAddr, 1)
                                oSheet.getCellRangeByPosition(
                                    0, el, 1, el).Rows.IsVisible = False

                    oCellRangeAddr.StartColumn = 5
                    oCellRangeAddr.EndColumn = 11
                    oSheet.group(oCellRangeAddr, 0)
                    oSheet.getCellRangeByPosition(5, 0, 11,
                                                  0).Columns.IsVisible = False
                    oCellRangeAddr.StartColumn = 15
                    oCellRangeAddr.EndColumn = 19
                    oSheet.group(oCellRangeAddr, 0)
                    oSheet.getCellRangeByPosition(15, 0, 19,
                                                  0).Columns.IsVisible = False
                ###

            if oDialog1.getControl(
                    "VarCon").State:  # Variante - Contabilità
                genera_sommario()

                oRangeAddress.StartColumn = 11
                oRangeAddress.EndColumn = 14

                oSheet.getCellByPosition(23,
                                         0).String = 'VARIANTE - CONTABILITÀ'
                for n in range(4, LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2):
                    formule.append([
                        '=IF(U' + str(n) + '-Q' + str(n) + '=0;"--";U' +
                        str(n) + '-Q' + str(n) + ')', '=IF(V' + str(n) + '-R' +
                        str(n) + '>0;V' + str(n) + '-R' + str(n) + ';"")',
                        '=IF(V' + str(n) + '-R' + str(n) + '<0;R' + str(n) +
                        '-V' + str(n) + ';"")', '=IFERROR(IFS(AND(R' + str(n) +
                        '>V' + str(n) + ';V' + str(n) + '=0);-1;AND(R' +
                        str(n) + '<V' + str(n) + ';R' + str(n) + '=0);1;R' +
                        str(n) + '=V' + str(n) + ';"--";R' + str(n) + '>V' +
                        str(n) + ';-(R' + str(n) + '-V' + str(n) + ')/R' +
                        str(n) + ';R' + str(n) + '<V' + str(n) + ';-(R' +
                        str(n) + '-V' + str(n) + ')/R' + str(n) + ');"--")'
                    ])
                n += 1
                #  for el in(1, ER+1):
                oSheet.getCellByPosition(
                    26, 1
                ).Formula = '=IFERROR(IFS(AND(R2>V2;V2=0);-1;AND(R2<V2;R2=0);1;R2=V2;"--";R2>V2;-(R2-V2)/R2;R2<V2;-(R2-V2)/R2);"--")'
                oSheet.getCellByPosition(
                    26, ER
                ).Formula = '=IFERROR(IFS(AND(R' + str(n) + '>V' + str(
                    n) + ';V' + str(n) + '=0);-1;AND(R' + str(n) + '<V' + str(
                        n) + ';R' + str(n) + '=0);1;R' + str(n) + '=V' + str(
                            n) + ';"--";R' + str(n) + '>V' + str(
                                n) + ';-(R' + str(n) + '-V' + str(
                                    n) + ')/R' + str(n) + ';R' + str(
                                        n) + '<V' + str(n) + ';-(R' + str(
                                            n) + '-V' + str(n) + ')/R' + str(
                                                n) + ');"--")'
                oRange = oSheet.getCellRangeByPosition(23, 3, 26,
                                                       LeenoSheetUtils.cercaUltimaVoce(oSheet))
                formule = tuple(formule)
                oRange.setFormulaArray(formule)
            # operazioni comuni
            for el in (11, 15, 19, 26):
                oSheet.getCellRangeByPosition(
                    el, 3, el, LeenoSheetUtils.cercaUltimaVoce(oSheet)).CellStyle = 'EP-mezzo %'
            for el in (12, 16, 20, 23):
                oSheet.getCellRangeByPosition(
                    el, 3, el,
                    LeenoSheetUtils.cercaUltimaVoce(oSheet)).CellStyle = 'EP statistiche_q'
            for el in (13, 17, 21, 24, 25):
                oSheet.getCellRangeByPosition(
                    el, 3, el,
                    LeenoSheetUtils.cercaUltimaVoce(oSheet)).CellStyle = 'EP statistiche'
            oCellRangeAddr.StartColumn = 3
            oCellRangeAddr.EndColumn = 3
            oSheet.group(oCellRangeAddr, 0)
            oSheet.getCellRangeByPosition(3, 0, 3, 0).Columns.IsVisible = False
            oCellRangeAddr.StartColumn = 5
            oCellRangeAddr.EndColumn = 11
            oSheet.group(oCellRangeAddr, 0)
            oSheet.getCellRangeByPosition(5, 0, 11,
                                          0).Columns.IsVisible = False

            oDoc.CurrentController.select(oSheet.getCellRangeByName('AA2'))
            #  oDoc.CurrentController.select(oDoc.getSheets().getByName('S5').getCellRangeByName('B30'))
            copy_clip()
            oDoc.CurrentController.select(
                oSheet.getCellRangeByPosition(26, 3, 26, ER))
            paste_format()

            if(oDialog1.getControl("ComVar").State or
               oDialog1.getControl("ComCon").State or
               oDialog1.getControl("VarCon").State):
                if DLG.DlgSiNo("Nascondo eventuali righe con scostamento nullo?") == 2:
                    errori = ('#DIV/0!', '--')
                    hide_error(errori, 26)
                    oSheet.group(oRangeAddress, 0)
                    oSheet.getCellRangeByPosition(oRangeAddress.StartColumn, 0,
                                                  oRangeAddress.EndColumn,
                                                  1).Columns.IsVisible = False
            _primaCella()
        else:
            return
    elif oSheet.Name in ('Analisi di Prezzo'):
        oDialog1 = dp.createDialog(
            "vnd.sun.star.script:UltimusFree2.DialogViste_AN?language=Basic&location=application"
        )
        # oDialog1Model = oDialog1.Model
        if not oSheet.getCellByPosition(1, 2).Rows.OptimalHeight:
            oDialog1.getControl("CBDesc").State = 1  # descrizione breve

        oS1 = oDoc.getSheets().getByName('S1')
        sString = oDialog1.getControl('TextField5')
        sString.Text = oS1.getCellRangeByName(
            'S1.H319').Value * 100  # sicurezza
        sString = oDialog1.getControl('TextField6')
        sString.Text = oS1.getCellRangeByName(
            'S1.H320').Value * 100  # spese_generali
        sString = oDialog1.getControl('TextField7')
        sString.Text = oS1.getCellRangeByName(
            'S1.H321').Value * 100  # utile_impresa

        # accorpa_spese_utili
        if oS1.getCellRangeByName('S1.H323').Value == 1:
            oDialog1.getControl('CheckBox4').State = 1
        sString = oDialog1.getControl('TextField8')
        sString.Text = oS1.getCellRangeByName('S1.H324').Value * 100  # sconto
        sString = oDialog1.getControl('TextField9')
        sString.Text = oS1.getCellRangeByName(
            'S1.H326').Value * 100  # maggiorazione

        oDialog1.execute()  # mostra il dialogo

        if(oSheet.getCellByPosition(1, 2).Rows.OptimalHeight and
           oDialog1.getControl("CBDesc").State == 1):  # descrizione breve
            basic_LeenO('Strutture.Tronca_Altezza_Analisi')
        #  elif oDialog1.getControl("CBDesc").State == 0: adatta_altezza_riga(oSheet.Name)

        #  sString.Text =oSheet.getCellRangeByName('S1.H321').Value * 100 #utile_impresa
        oS1.getCellRangeByName('S1.H319').Value = float(
            oDialog1.getControl('TextField5').getText().replace(
                ',', '.')) / 100  # sicurezza
        oS1.getCellRangeByName('S1.H320').Value = float(
            oDialog1.getControl('TextField6').getText().replace(
                ',', '.')) / 100  # spese generali
        oS1.getCellRangeByName('S1.H321').Value = float(
            oDialog1.getControl('TextField7').getText().replace(
                ',', '.')) / 100  # utile_impresa
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

    elif oSheet.Name in ('CONTABILITA', 'Registro', 'SAL'):
        oDialog1 = dp.createDialog(
            "vnd.sun.star.script:UltimusFree2.Dialogviste_N?language=Basic&location=application"
        )
        # oDialog1Model = oDialog1.Model
        oDialog1.getControl('Dettaglio').State = cfg.read('Generale', 'dettaglio')
        oDialog1.execute()
        if oDialog1.getControl('Dettaglio').State == 0:
            cfg.write('Generale', 'dettaglio', '0')
            dettaglio_misure(0)
        else:
            cfg.write('Generale', 'dettaglio', '1')
            dettaglio_misure(0)
            dettaglio_misure(1)
    # adatta_altezza_riga(oSheet.Name)
    EnableAutoCalc()
    # MsgBox('Operazione eseguita con successo!','')


########################################################################


def genera_variante():
    '''
    Genera il foglio di VARIANTE a partire dal COMPUTO
    @@@ MODIFICA IN CORSO CON 'LeenoVariante.generaVariante'
    '''
    chiudi_dialoghi()
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
        setTabColor(16777062)
        oSheet.getCellByPosition(2, 0).String = "VARIANTE"
        oSheet.getCellByPosition(2, 0).CellStyle = "comp Int_colonna"
        oSheet.getCellRangeByName("C1").CellBackColor = 16777062
        oSheet.getCellRangeByPosition(0, 2, 42, 2).CellBackColor = 16777062
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
            adatta_altezza_riga('VARIANTE')
    #  else:
    GotoSheet('VARIANTE')
    ScriviNomeDocumentoPrincipale()
    basic_LeenO("Menu.eventi_assegna")


########################################################################


def genera_sommario():
    '''
    Genera i sommari in Elenco Prezzi
    '''
    #  oDialogo_attesa = dlg_attesa()
    #  attesa().start() #mostra il dialogo
    DisableAutoCalc()
    struttura_off()

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName('COMPUTO')
    lrow = SheetUtils.getUsedArea(oSheet).EndRow
    SheetUtils.NominaArea(oDoc, 'COMPUTO', '$AJ$3:$AJ$' + str(lrow), 'AA')
    SheetUtils.NominaArea(oDoc, 'COMPUTO', '$N$3:$N$' + str(lrow), "BB")
    SheetUtils.NominaArea(oDoc, 'COMPUTO', '$AK$3:$AK$' + str(lrow), "cEuro")

    if oDoc.getSheets().hasByName('VARIANTE'):
        oSheet = oDoc.getSheets().getByName('VARIANTE')
        lrow = SheetUtils.getUsedArea(oSheet).EndRow
        SheetUtils.NominaArea(oDoc, 'VARIANTE', '$AJ$3:$AJ$' + str(lrow), 'varAA')
        SheetUtils.NominaArea(oDoc, 'VARIANTE', '$N$3:$N$' + str(lrow), "varBB")
        SheetUtils.NominaArea(oDoc, 'VARIANTE', '$AK$3:$AK$' + str(lrow), "varEuro")

    if oDoc.getSheets().hasByName('CONTABILITA'):
        oSheet = oDoc.getSheets().getByName('CONTABILITA')
        lrow = SheetUtils.getUsedArea(oSheet).EndRow
        lrow = SheetUtils.getUsedArea(
            oDoc.getSheets().getByName('CONTABILITA')).EndRow
        SheetUtils.NominaArea(oDoc, 'CONTABILITA', '$AJ$3:$AJ$' + str(lrow), 'GG')
        SheetUtils.NominaArea(oDoc, 'CONTABILITA', '$S$3:$S$' + str(lrow), "G1G1")
        SheetUtils.NominaArea(oDoc, 'CONTABILITA', '$AK$3:$AK$' + str(lrow), "conEuro")

    formule = list()
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    for n in range(4, LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2):
        stringa = ([
            '=N' + str(n) + '/$N$2', '=SUMIF(AA;A' + str(n) + ';BB)',
            '=SUMIF(AA;A' + str(n) + ';cEuro)', '', '', '', '', '', '', '', ''
        ])
        if oDoc.getSheets().hasByName('VARIANTE'):
            stringa = ([
                '=N' + str(n) + '/$N$2', '=SUMIF(AA;A' + str(n) + ';BB)',
                '=SUMIF(AA;A' + str(n) + ';cEuro)', '',
                '=R' + str(n) + '/$R$2', '=SUMIF(varAA;A' + str(n) + ';varBB)',
                '=SUMIF(varAA;A' + str(n) + ';varEuro)', '', '', '', ''
            ])
            if oDoc.getSheets().hasByName('CONTABILITA'):
                stringa = ([
                    '=N' + str(n) + '/$N$2', '=SUMIF(AA;A' + str(n) + ';BB)',
                    '=SUMIF(AA;A' + str(n) + ';cEuro)', '',
                    '=R' + str(n) + '/$R$2',
                    '=SUMIF(varAA;A' + str(n) + ';varBB)',
                    '=SUMIF(varAA;A' + str(n) + ';varEuro)', '',
                    '=V' + str(n) + '/$V$2', '=SUMIF(GG;A' + str(n) + ';G1G1)',
                    '=SUMIF(GG;A' + str(n) + ';conEuro)'
                ])
        elif oDoc.getSheets().hasByName('CONTABILITA'):
            stringa = ([
                '=N' + str(n) + '/$N$2', '=SUMIF(AA;A' + str(n) + ';BB)',
                '=SUMIF(AA;A' + str(n) + ';cEuro)', '', '', '', '', '',
                '=V' + str(n) + '/$V$2', '=SUMIF(GG;A' + str(n) + ';G1G1)',
                '=SUMIF(GG;A' + str(n) + ';conEuro)'
            ])
        formule.append(stringa)
    oRange = oSheet.getCellRangeByPosition(11, 3, 21, LeenoSheetUtils.cercaUltimaVoce(oSheet))
    formule = tuple(formule)
    oRange.setFormulaArray(formule)

    EnableAutoCalc()
    adatta_altezza_riga(oSheet.Name)
    #  oDialogo_attesa.endExecute() #chiude il dialogo


########################################################################


def MENU_riordina_ElencoPrezzi():
    '''
    Riordina l'Elenco Prezzi secondo l'ordine alfabetico dei codici di prezzo
    '''
    riordina_ElencoPrezzi()


def riordina_ElencoPrezzi():
    '''
    Riordina l'Elenco Prezzi secondo l'ordine alfabetico dei codici di prezzo
    '''
    chiudi_dialoghi()
    DisableAutoCalc()

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    if SheetUtils.uFindStringCol('Fine elenco', 0, oSheet) is None:
        LeenoSheetUtils.inserisciRigaRossa(oSheet)
    test = str(SheetUtils.uFindStringCol('Fine elenco', 0, oSheet))
    SheetUtils.NominaArea(oDoc, 'Elenco Prezzi', "$A$3:$AF$" + test, 'elenco_prezzi')
    SheetUtils.NominaArea(oDoc, 'Elenco Prezzi', "$A$3:$A$" + test, 'Lista')
    oRangeAddress = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    SC = oRangeAddress.StartColumn
    EC = oRangeAddress.EndColumn
    SR = oRangeAddress.StartRow + 1
    ER = oRangeAddress.EndRow
    if SR == ER:
        return
    oRange = oSheet.getCellRangeByPosition(SC, SR, EC, ER)

    '''
    REPLACED WITH DIRECT SORT WITHOUT USING CurrentController
    So it can be done headless without screen flickering

    oDoc.CurrentController.select(oRange)
    ordina_col(1)
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect
    '''
    SheetUtils.simpleSortColumn(oRange, 0, True)

    EnableAutoCalc()


########################################################################


def MENU_doppioni():
    EliminaVociDoppieElencoPrezzi()


def EliminaVociDoppieElencoPrezzi():
    oDoc = LeenoUtils.getDocument()
    '''
    Cancella eventuali voci che si ripetono in Elenco Prezzi
    '''
    zoom = oDoc.CurrentController.ZoomValue
    DisableAutoCalc()

    if oDoc.getSheets().hasByName('Analisi di Prezzo'):
        lista_tariffe_analisi = list()
        oSheet = oDoc.getSheets().getByName('Analisi di Prezzo')
        for n in range(0, LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1):
            if oSheet.getCellByPosition(0, n).CellStyle == 'An-1_sigla':
                lista_tariffe_analisi.append(
                    oSheet.getCellByPosition(0, n).String)
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')

    SR = 0
    ER = SheetUtils.getUsedArea(oSheet).EndRow

    try:
        lista_tariffe_analisi
        for i in reversed(range(SR, ER)):
            if oSheet.getCellByPosition(0, i).String in lista_tariffe_analisi:
                oSheet.getRows().removeByIndex(i, 1)
    except Exception:
        pass
    oRangeAddress = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    SR = oRangeAddress.StartRow + 1
    ER = oRangeAddress.EndRow - 1
    if ER < SR:
        return
    oRange = oSheet.getCellRangeByPosition(0, SR, 7, ER)
    lista_come_array = tuple(set(oRange.getDataArray()))
    # ~chi ([len(lista_come_array),(lista_come_array)])
    # ~return
    oSheet.getRows().removeByIndex(SR, ER - SR + 1)
    lista_tar = list()
    oSheet.getRows().insertByIndex(SR, len(set(lista_come_array)))
    for el in set(lista_come_array):
        lista_tar.append(el[0])
    colonne_lista = len(lista_come_array[0]
                        )  # numero di colonne necessarie per ospitare i dati
    righe_lista = len(
        lista_come_array)  # numero di righe necessarie per ospitare i dati
    oRange = oSheet.getCellRangeByPosition(
        0,
        3,
        colonne_lista + 0 - 1,  # l'indice parte da 0
        righe_lista + 3 - 1)
    oRange.setDataArray(lista_come_array)
    oSheet.getCellRangeByPosition(0, 3, 0,
                                  righe_lista + 3 - 1).CellStyle = "EP-aS"
    oSheet.getCellRangeByPosition(1, 3, 1,
                                  righe_lista + 3 - 1).CellStyle = "EP-a"
    oSheet.getCellRangeByPosition(2, 3, 7,
                                  righe_lista + 3 - 1).CellStyle = "EP-mezzo"
    oSheet.getCellRangeByPosition(5, 3, 5,
                                  righe_lista + 3 - 1).CellStyle = "EP-mezzo %"
    oSheet.getCellRangeByPosition(8, 3, 9,
                                  righe_lista + 3 - 1).CellStyle = "EP-sfondo"

    oSheet.getCellRangeByPosition(11, 3, 11,
                                  righe_lista + 3 - 1).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition(12, 3, 12, righe_lista + 3 -
                                  1).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition(13, 3, 13, righe_lista + 3 -
                                  1).CellStyle = 'EP statistiche_Contab_q'
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect
    if oDoc.getSheets().hasByName('Analisi di Prezzo'):
        tante_analisi_in_ep()

    EnableAutoCalc()
    oDoc.CurrentController.ZoomValue = zoom
    adatta_altezza_riga(oSheet.Name)
    riordina_ElencoPrezzi()
    if len(set(lista_tar)) != len(set(lista_come_array)):
        DLG.MsgBox(
            'Ci sono ancora 2 o più voci che hanno lo stesso Codice Articolo pur essendo diverse.',
            'C o n t r o l l a!')


########################################################################
# Scrive un file.
def XPWE_out(elaborato, out_file):
    '''
    esporta il documento in formato XPWE

    elaborato { string } : nome del foglio da esportare
    out_file  { string } : nome base del file

    il nome file risulterà out_file-elaborato.xpwe
    '''
    DisableAutoCalc()
    oDoc = LeenoUtils.getDocument()
    oDialogo_attesa = DLG.dlg_attesa('Esportazione di ' + elaborato + ' in corso...')
    DLG.attesa().start()  # mostra il dialogo
    if cfg.read('Generale', 'dettaglio') == '1':
        dettaglio_misure(0)
    numera_voci(1)
    top = Element('PweDocumento')
    #  intestazioni
    CopyRight = SubElement(top, 'CopyRight')
    CopyRight.text = 'Copyright ACCA software S.p.A.'
    TipoDocumento = SubElement(top, 'TipoDocumento')
    TipoDocumento.text = '1'
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
    listaspcat = list()
    PweDGSuperCategorie = SubElement(PweDGCapitoliCategorie,
                                     'PweDGSuperCategorie')
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
    listaCat = list()
    PweDGCategorie = SubElement(PweDGCapitoliCategorie, 'PweDGCategorie')
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
    listasbCat = list()
    PweDGSubCategorie = SubElement(PweDGCapitoliCategorie, 'PweDGSubCategorie')
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

    #  Elenco Prezzi
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    PweElencoPrezzi = SubElement(PweMisurazioni, 'PweElencoPrezzi')
    diz_ep = dict()
    lista_AP = list()
    for n in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
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
            DesEstesa = SubElement(EPItem, 'DesEstesa')
            DesEstesa.text = oSheet.getCellByPosition(1, n).String
            DesRidotta = SubElement(EPItem, 'DesRidotta')
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
            IDSpCap = SubElement(EPItem, 'IDSpCap')
            IDSpCap.text = '0'
            IDCap = SubElement(EPItem, 'IDCap')
            IDCap.text = '0'
            IDSbCap = SubElement(EPItem, 'IDSbCap')
            IDSbCap.text = '0'
            Flags = SubElement(EPItem, 'Flags')
            if oSheet.getCellByPosition(8, n).String == '(AP)':
                Flags.text = '131072'
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
    if len(lista_AP) != 0:
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
                DesEstesa = SubElement(EPItem, 'DesEstesa')
                DesEstesa.text = oSheet.getCellByPosition(1, m).String
                DesRidotta = SubElement(EPItem, 'DesRidotta')
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
                if oSheet.getCellByPosition(10, n).Value == 0.0:
                    IncSIC.text = ''
                else:
                    IncSIC.text = str(oSheet.getCellByPosition(10, n).Value)

                IncMDO = SubElement(EPItem, 'IncMDO')
                if oSheet.getCellByPosition(8, n).Value == 0.0:
                    IncMDO.text = ''
                else:
                    IncMDO.text = str(
                        oSheet.getCellByPosition(5, n).Value * 100)
                k += 1
            except Exception:
                pass

    # COMPUTO/VARIANTE/CONTABILITA
    oSheet = oDoc.getSheets().getByName(elaborato)
    PweVociComputo = SubElement(PweMisurazioni, 'PweVociComputo')
    oDoc.CurrentController.setActiveSheet(oSheet)
    nVCItem = 2
    for n in range(0, LeenoSheetUtils.cercaUltimaVoce(oSheet)):
        if oSheet.getCellByPosition(0,
                                    n).CellStyle in ('Comp Start Attributo',
                                                     'Comp Start Attributo_R'):
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, n)
            sStRange.RangeAddress
            sopra = sStRange.RangeAddress.StartRow
            sotto = sStRange.RangeAddress.EndRow
            if elaborato == 'CONTABILITA':
                sotto -= 1
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
            ##########################
            IDSpCat = SubElement(VCItem, 'IDSpCat')
            IDSpCat.text = str(oSheet.getCellByPosition(31, sotto).String)
            if IDSpCat.text == '':
                IDSpCat.text = '0'
            # #########################
            IDCat = SubElement(VCItem, 'IDCat')
            IDCat.text = str(oSheet.getCellByPosition(32, sotto).String)
            if IDCat.text == '':
                IDCat.text = '0'
            # #########################
            IDSbCat = SubElement(VCItem, 'IDSbCat')
            IDSbCat.text = str(oSheet.getCellByPosition(33, sotto).String)
            if IDSbCat.text == '':
                IDSbCat.text = '0'
            # #########################
            PweVCMisure = SubElement(VCItem, 'PweVCMisure')
            for m in range(sopra + 2, sotto):
                RGItem = SubElement(PweVCMisure, 'RGItem')
                x = 2
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
                if oSheet.getCellByPosition(11, m).Value != 0:
                    Quantita.text = '-' + oSheet.getCellByPosition(11,
                                                                   m).String
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
                elif 'PARTITA IN CONTO PROVVISORIO' in Descrizione.text:
                    Flags.text = '16'
                else:
                    Flags.text = '0'
                # #########################
                if 'DETRAE LA PARTITA IN CONTO PROVVISORIO' in Descrizione.text:
                    Flags.text = '32'
                if '- vedi voce n.' in Descrizione.text:
                    IDVV.text = str(
                        int(
                            Descrizione.text.split('- vedi voce n.')[1].split(
                                ' ')[0]) + 1)
                    Flags.text = '32768'
                    #  PartiUguali.text =''
                    if '-' in Quantita.text or oSheet.getCellByPosition(
                            11, m).Value != 0:
                        Flags.text = '32769'
            n = sotto + 1
    # #########################
    oDialogo_attesa.endExecute()
    # ~out_file = Dialogs.FileSelect('Salva con nome...', '*.xpwe', 1)
    # ~out_file = uno.fileUrlToSystemPath(oDoc.getURL())
    # ~mri (uno.fileUrlToSystemPath(oDoc.getURL()))
    # ~chi(out_file)
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
    try:
        of = codecs.open(out_file, 'w', 'utf-8')
        of.write(riga)
        # ~MsgBox('Esportazione in formato XPWE eseguita con successo\nsul file ' + out_file + '!','Avviso.')
    except Exception:
        DLG.MsgBox(
            'Esportazione non eseguita!\n\nVerifica che il file di destinazione non sia già in uso!',
            'E R R O R E !')

    EnableAutoCalc()


########################################################################
#  def firme_in_calce_run():
def MENU_firme_in_calce():
    oDialogo_attesa = DLG.dlg_attesa(
    )  # avvia il diaolgo di attesa che viene chiuso alla fine con
    '''
    Inserisce(in COMPUTO o VARIANTE) un riepilogo delle categorie
    ed i dati necessari alle firme
    '''
    oDoc = LeenoUtils.getDocument()

    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in ('Analisi di Prezzo', 'Elenco Prezzi'):
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
        oSheet.getCellRangeByPosition(0, lrowF, 100, lrowF + 15 -
                                      1).CellStyle = "Ultimus_centro"
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
        #  consolido il risultato
        oRange = oSheet.getCellByPosition(1, riga_corrente + 3)
        # flags = (oDoc.createInstance('com.sun.star.sheet.CellFlags.FORMULA'))
        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)
        oSheet.getCellRangeByPosition(1, riga_corrente + 3, 1,
                                      riga_corrente + 3).CellStyle = 'ULTIMUS'
        oSheet.getCellByPosition(1,
                                 riga_corrente + 5).Formula = 'Il Progettista'
        oSheet.getCellByPosition(
            1, riga_corrente + 6
        ).Formula = '=CONCATENATE($S2.$C$13)'  # senza concatenate, se la cella di origine è vuota il risultato è '0,00'

    if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CompuM_NoP'):
        zoom = oDoc.CurrentController.ZoomValue
        oDoc.CurrentController.ZoomValue = 400

        DLG.attesa().start()
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
        inizio_gruppo = riga_corrente
        riga_corrente += 1
        for i in range(0, lrowF):
            if oSheet.getCellByPosition(1, i).CellStyle == 'Livello-0-scritta':
                oSheet.getRows().insertByIndex(riga_corrente, 1)
                oSheet.getCellRangeByPosition(
                    0, riga_corrente, 30,
                    riga_corrente).CellStyle = 'ULTIMUS_1'
                oSheet.getCellByPosition(
                    1, riga_corrente).Formula = '=B' + str(i + 1)
                oSheet.getCellByPosition(
                    1, riga_corrente).CellStyle = 'Ultimus_destra_1'
                oSheet.getCellByPosition(
                    2, riga_corrente).Formula = '=C' + str(i + 1)
                oSheet.getCellByPosition(
                    ii, riga_corrente).Formula = '=' + col + str(
                        riga_corrente + 1) + '/' + col + str(lrowF) + '*100'
                oSheet.getCellByPosition(
                    ii, riga_corrente).CellStyle = 'Ultimus %_1'
                oSheet.getCellByPosition(
                    vv, riga_corrente).Formula = '=' + col + str(i + 1)
                oSheet.getCellRangeByPosition(
                    vv, riga_corrente, ae,
                    riga_corrente).CellStyle = 'Ultimus_totali_1'
                oSheet.getCellByPosition(
                    ac, riga_corrente).Formula = '=AC' + str(i + 1)
                oSheet.getCellByPosition(
                    ad, riga_corrente).Formula = '=AD' + str(i + 1) + '*100'
                oSheet.getCellByPosition(
                    ad, riga_corrente).CellStyle = 'Ultimus %_1'
                oSheet.getCellByPosition(
                    ae, riga_corrente).Formula = '=AE' + str(i + 1)
                riga_corrente += 1
            elif oSheet.getCellByPosition(1,
                                          i).CellStyle == 'Livello-1-scritta':
                oSheet.getRows().insertByIndex(riga_corrente, 1)
                oSheet.getCellRangeByPosition(
                    0, riga_corrente, 30,
                    riga_corrente).CellStyle = 'ULTIMUS_2'
                oSheet.getCellByPosition(
                    1, riga_corrente).Formula = '=B' + str(i + 1)
                oSheet.getCellByPosition(
                    1, riga_corrente).CellStyle = 'Ultimus_destra'
                oSheet.getCellByPosition(
                    2, riga_corrente).Formula = '=C' + str(i + 1)
                oSheet.getCellByPosition(
                    ii, riga_corrente).Formula = '=' + col + str(
                        riga_corrente + 1) + '/' + col + str(lrowF) + '*100'
                oSheet.getCellByPosition(ii,
                                         riga_corrente).CellStyle = 'Ultimus %'
                oSheet.getCellByPosition(
                    vv, riga_corrente).Formula = '=' + col + str(i + 1)
                oSheet.getCellByPosition(
                    vv, riga_corrente).CellStyle = 'Ultimus_bordo'
                oSheet.getCellByPosition(
                    ac, riga_corrente).Formula = '=AC' + str(i + 1)
                oSheet.getCellByPosition(
                    ad, riga_corrente).Formula = '=AD' + str(i + 1) + '*100'
                oSheet.getCellByPosition(ad,
                                         riga_corrente).CellStyle = 'Ultimus %'
                oSheet.getCellByPosition(
                    ae, riga_corrente).Formula = '=AE' + str(i + 1)
                riga_corrente += 1
            elif oSheet.getCellByPosition(1, i).CellStyle == 'livello2 valuta':
                oSheet.getRows().insertByIndex(riga_corrente, 1)
                oSheet.getCellRangeByPosition(
                    0, riga_corrente, 30,
                    riga_corrente).CellStyle = 'ULTIMUS_3'
                oSheet.getCellByPosition(
                    1, riga_corrente).Formula = '=B' + str(i + 1)
                oSheet.getCellByPosition(
                    1, riga_corrente).CellStyle = 'Ultimus_destra_3'
                oSheet.getCellByPosition(
                    2, riga_corrente).Formula = '=C' + str(i + 1)
                oSheet.getCellByPosition(
                    ii, riga_corrente).Formula = '=' + col + str(
                        riga_corrente + 1) + '/' + col + str(lrowF) + '*100'
                oSheet.getCellByPosition(
                    ii, riga_corrente).CellStyle = 'Ultimus %_3'
                oSheet.getCellByPosition(
                    vv, riga_corrente).Formula = '=' + col + str(i + 1)
                oSheet.getCellByPosition(vv,
                                         riga_corrente).CellStyle = 'ULTIMUS_3'
                oSheet.getCellByPosition(
                    ac, riga_corrente).Formula = '=AC' + str(i + 1)
                oSheet.getCellByPosition(
                    ad, riga_corrente).Formula = '=AD' + str(i + 1) + '*100'
                oSheet.getCellByPosition(
                    ad, riga_corrente).CellStyle = 'Ultimus %_3'
                oSheet.getCellByPosition(
                    ae, riga_corrente).Formula = '=AE' + str(i + 1)
                riga_corrente += 1
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
        # fine_gruppo = riga_corrente
        #  DATA
        oSheet.getCellByPosition(
            2, riga_corrente +
            3).Formula = '=CONCATENATE("Data, ";TEXT(NOW();"GG/MM/AAAA"))'
        #  consolido il risultato
        oRange = oSheet.getCellByPosition(2, riga_corrente + 3)
        # flags = (oDoc.createInstance('com.sun.star.sheet.CellFlags.FORMULA'))
        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)

        oSheet.getCellByPosition(2,
                                 riga_corrente + 5).Formula = 'Il Progettista'
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
        dispatchHelper.executeDispatch(oFrame, ".uno:InsertRowBreak", "", 0, list())
        oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
        ###
        #  oSheet.getCellByPosition(lrowF,0).Rows.IsManualPageBreak = True
    oDialogo_attesa.endExecute()
    oDoc.CurrentController.ZoomValue = zoom


########################################################################
def next_voice(lrow, n=1):
    # ~def debug (arg=None, n=1):
    '''
    lrow { double }   : riga di riferimento
    n    { integer }  : se 0 sposta prima della voce corrente
                        se 1 sposta dopo della voce corrente
    sposta il cursore prima o dopo la voce corrente restituendo un idrow
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_contab = LeenoUtils.getGlobalVar('stili_contab')
    noVoce = LeenoUtils.getGlobalVar('noVoce')

    # ~lrow = LeggiPosizioneCorrente()[1]
    if lrow == 0:
        while oSheet.getCellByPosition(0, lrow).CellStyle not in stili_computo + stili_contab:
            lrow += 1
        return lrow
    fine = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1
    # la parte che segue sposta il focus dopo della voce corrente (ad esempio sul titolo di categoria)
    if lrow >= fine:
        return lrow
    if oSheet.getCellByPosition(0, lrow).CellStyle in stili_computo + stili_contab:
        if n == 0:
            sopra = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
            lrow = sopra
        elif n == 1:
            sotto = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow
            lrow = sotto + 1
    elif oSheet.getCellByPosition(
            0, lrow).CellStyle in ('Ultimus_centro_bordi_lati', ):
        for y in range(lrow, SheetUtils.getUsedArea(oSheet).EndRow + 1):
            if oSheet.getCellByPosition(0, y).CellStyle != 'Ultimus_centro_bordi_lati':
                lrow = y
                break
    elif oSheet.getCellByPosition(0, lrow).CellStyle in noVoce:
        # ~while oSheet.getCellByPosition(0, lrow).CellStyle in noVoce:
        lrow += 1
    else:
        return
    return lrow
    # la parte che segue sposta il focus all'effettivo inizio della voce successiva
    # ~fine = LeenoSheetUtils.cercaUltimaVoce(oSheet)+1
    # ~if lrow <= 1: lrow = 2
    # ~if lrow >= fine or oSheet.getCellByPosition(0, lrow).CellStyle in('Comp TOTALI'): return lrow
    # ~if oSheet.getCellByPosition(0, lrow).CellStyle in stili_computo + stili_contab:
    # ~if n==0:
    # ~sopra = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
    # ~lrow = sopra
    # ~elif n==1:
    # ~sotto = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow
    # ~lrow = sotto+1
    # ~elif oSheet.getCellByPosition(0, lrow).CellStyle in ('Ultimus_centro_bordi_lati',):
    # ~for y in range(lrow, SheetUtils.getUsedArea(oSheet).EndRow+1):
    # ~if oSheet.getCellByPosition(0, y).CellStyle != 'Ultimus_centro_bordi_lati':
    # ~lrow = y
    # ~break
    # ~while oSheet.getCellByPosition(0, lrow).CellStyle in noVoce:
    # ~lrow +=1
    # ~return lrow


########################################################################
def cancella_analisi_da_ep():
    '''
    cancella le voci in Elenco Prezzi che derivano da analisi
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
    lista_an = list()
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
    @@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()
    try:
        oSheet = oDoc.CurrentController.ActiveSheet
        if oSheet.Name != 'Analisi di Prezzo':
            return
        oDoc.enableAutomaticCalculation(False)  # blocco il calcolo automatico
        sStRange = Circoscrive_Analisi(LeggiPosizioneCorrente()[1])
        riga = sStRange.RangeAddress.StartRow + 2

        codice = oSheet.getCellByPosition(0, riga).String

        oSheet = oDoc.Sheets.getByName('Elenco Prezzi')
        oDoc.CurrentController.setActiveSheet(oSheet)

        oSheet.getRows().insertByIndex(3, 1)

        oSheet.getCellByPosition(0, 3).CellStyle = 'EP-aS'
        oSheet.getCellByPosition(1, 3).CellStyle = 'EP-a'
        oSheet.getCellRangeByPosition(2, 3, 8, 3).CellStyle = 'EP-mezzo'
        oSheet.getCellByPosition(5, 3).CellStyle = 'EP-mezzo %'
        oSheet.getCellByPosition(9, 3).CellStyle = 'EP-sfondo'
        oSheet.getCellByPosition(10, 3).CellStyle = 'Default'
        oSheet.getCellByPosition(11, 3).CellStyle = 'EP-mezzo %'
        oSheet.getCellByPosition(12, 3).CellStyle = 'EP statistiche_q'
        oSheet.getCellByPosition(13, 3).CellStyle = 'EP statistiche_Contab_q'

        oSheet.getCellByPosition(0, 3).String = codice

        oSheet.getCellByPosition(
            1, 3).Formula = "=$'Analisi di Prezzo'.B" + str(riga + 1)
        oSheet.getCellByPosition(
            2, 3).Formula = "=$'Analisi di Prezzo'.C" + str(riga + 1)
        oSheet.getCellByPosition(
            3, 3).Formula = "=$'Analisi di Prezzo'.K" + str(riga + 1)
        oSheet.getCellByPosition(
            4, 3).Formula = "=$'Analisi di Prezzo'.G" + str(riga + 1)
        oSheet.getCellByPosition(
            5, 3).Formula = "=$'Analisi di Prezzo'.I" + str(riga + 1)
        oSheet.getCellByPosition(
            6, 3).Formula = "=$'Analisi di Prezzo'.J" + str(riga + 1)
        oSheet.getCellByPosition(
            7, 3).Formula = "=$'Analisi di Prezzo'.A" + str(riga + 1)
        oSheet.getCellByPosition(8, 3).String = "(AP)"
        oSheet.getCellByPosition(11, 3).Formula = "=N4/$N$2"
        oSheet.getCellByPosition(12, 3).Formula = "=SUMIF(AA;A4;BB)"
        oSheet.getCellByPosition(13, 3).Formula = "=SUMIF(AA;A4;cEuro)"
        oDoc.enableAutomaticCalculation(True)  # sblocco il calcolo automatico
        _gotoCella(1, 3)
    except Exception:
        oDoc.enableAutomaticCalculation(True)


########################################################################
def tante_analisi_in_ep():
    '''
    Trasferisce le analisi all'Elenco Prezzi.
    '''
    chiudi_dialoghi()
    DisableAutoCalc()

    oDoc = LeenoUtils.getDocument()
    lista_analisi = list()
    oSheet = oDoc.getSheets().getByName('Analisi di prezzo')
    SheetUtils.NominaArea(oDoc, 'Analisi di Prezzo',
                  '$A$3:$K$' + str(SheetUtils.getUsedArea(oSheet).EndRow), 'analisi')
    voce = list()
    idx = 4
    for n in range(0, LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1):
        if oSheet.getCellByPosition(
                0, n
        ).CellStyle == 'An-1_sigla' and oSheet.getCellByPosition(
                1, n
        ).String != '<<<Scrivi la descrizione della nuova voce da analizzare   ':
            voce = (
                oSheet.getCellByPosition(0, n).String,
                "=$'Analisi di Prezzo'.B" + str(n + 1),
                "=$'Analisi di Prezzo'.C" + str(n + 1),
                "=$'Analisi di Prezzo'.K" + str(n + 1),
                "=$'Analisi di Prezzo'.G" + str(n + 1),
                "=$'Analisi di Prezzo'.I" + str(n + 1),
                "=$'Analisi di Prezzo'.J" + str(n + 1),
                "=$'Analisi di Prezzo'.A" + str(n + 1),
                "(AP)",
                '',
                '',
                "=N" + str(idx) + "/$N$2",
                "=SUMIF(AA;A" + str(idx) + ";BB)",
                "=SUMIF(AA;A" + str(idx) + ";cEuro)",
            )
            lista_analisi.append(voce)
            idx += 1
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    if len(lista_analisi) != 0:
        oSheet.getRows().insertByIndex(3, len(lista_analisi))
    else:
        return
    oRange = oSheet.getCellRangeByPosition(0, 3, 13,
                                           3 + len(lista_analisi) - 1)
    lista_come_array = tuple(lista_analisi)
    oRange.setDataArray(
        lista_come_array
    )  # setFormulaArray() sarebbe meglio, ma mi fa storie sul codice articolo
    for y in range(3, 3 + len(lista_analisi)):
        for x in range(
                1, len(lista_analisi[0])
        ):  # evito il codice articolo, altrimenti me lo converte in numero
            oSheet.getCellByPosition(x, y).Formula = oSheet.getCellByPosition(
                x, y).String
    oSheet.getCellRangeByPosition(0, 3, 0, 3 + len(lista_analisi) -
                                  1).CellStyle = 'EP-aS'
    oSheet.getCellRangeByPosition(1, 3, 1, 3 + len(lista_analisi) -
                                  1).CellStyle = 'EP-a'
    oSheet.getCellRangeByPosition(2, 3, 8, 3 + len(lista_analisi) -
                                  1).CellStyle = 'EP-mezzo'
    oSheet.getCellRangeByPosition(5, 3, 5, 3 + len(lista_analisi) -
                                  1).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition(9, 3, 9, 3 + len(lista_analisi) -
                                  1).CellStyle = 'EP-sfondo'
    oSheet.getCellRangeByPosition(10, 3, 10, 3 + len(lista_analisi) -
                                  1).CellStyle = 'Default'
    oSheet.getCellRangeByPosition(11, 3, 11, 3 + len(lista_analisi) -
                                  1).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition(12, 3, 12, 3 + len(lista_analisi) -
                                  1).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition(13, 3, 13, 3 + len(lista_analisi) -
                                  1).CellStyle = 'EP statistiche_Contab_q'

    EnableAutoCalc()
    GotoSheet('Elenco Prezzi')
    #  MsgBox('Trasferite ' + str(len(lista_analisi)) + ' analisi di prezzo in Elenco Prezzi.', 'Avviso')


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
            #  chi(oSheet.getCellByPosition(0, el).CellStyle)
            if oSheet.getCellByPosition(0, el).CellStyle == 'Analisi_Sfondo':
                SR = el
                break
        for el in range(lrow, SheetUtils.getUsedArea(oSheet).EndRow):
            if oSheet.getCellByPosition(
                    0, el).CellStyle == 'An-sfondo-basso Att End':
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
    DisableAutoCalc()

    try:
        oDoc = LeenoUtils.getDocument()
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
            fini = list()
            for x in range(sRow, eRow):
                if oSheet.getCellByPosition(
                        0, x).CellStyle == 'Comp End Attributo':
                    fini.append(x)
                elif oSheet.getCellByPosition(
                        0, x).CellStyle == 'Comp End Attributo_R':
                    fini.append(x - 2)
        idx = 0
        for lrow in fini:
            lrow += idx
            try:
                sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
                sStRange.RangeAddress
                inizio = sStRange.RangeAddress.StartRow
                fine = sStRange.RangeAddress.EndRow
                if oSheet.Name == 'CONTABILITA':
                    fine -= 1
                _gotoCella(2, fine - 1)
                if oSheet.getCellByPosition(
                        2, fine - 1).String == '*** VOCE AZZERATA ***':
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
                    oSheet.getCellByPosition(
                        2, fine).String = '*** VOCE AZZERATA ***'
                    if oSheet.Name == 'CONTABILITA':
                        oSheet.getCellByPosition(
                            5, fine).Formula = '=SUBTOTAL(9;J' + str(
                                inizio + 1) + ':J' + str(
                                    fine + 1) + ')-SUBTOTAL(9;L' + str(
                                        inizio) + ':L' + str(fine) + ')'
                        inverti_segno()
                        # ~oSheet.getCellByPosition(9, fine).Formula = '=-SUM(J' + str(inizio+1) + ':J' + str(fine) + ')'
                        # ~oSheet.getCellByPosition(11, fine).Formula = '=-SUM(L' + str(inizio+1) + ':L' + str(fine) + ')'
                    else:
                        oSheet.getCellByPosition(
                            5, fine).Formula = '=-SUBTOTAL(9;J' + str(
                                inizio + 1) + ':J' + str(fine) + ')'
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
                lrow = next_voice(lrow, 1)
            except Exception:
                pass
        # ~numera_voci(1)
    except Exception:
        pass

    EnableAutoCalc()


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
            for lrow in reversed(range(0, ER)):
                if oSheet.getCellByPosition(
                        2, lrow).String == '*** VOCE AZZERATA ***':
                    elimina_voce(lrow=lrow, msg=0)
            numera_voci(1)
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
            oSheet.clearOutline()
            ER = SheetUtils.getUsedArea(oSheet).EndRow
            for lrow in reversed(range(0, ER)):
                if oSheet.getCellByPosition(
                        2, lrow).String == '*** VOCE AZZERATA ***':
                    raggruppa_righe_voce(lrow, 1)
    except Exception:
        return


########################################################################
def seleziona(lrow=None):
    #  def debug(lrow=None):
    '''
    Seleziona voci intere
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in ('Elenco Prezzi'):
        return
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
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
            DLG.MsgBox('La selezione deve essere contigua.', 'ATTENZIONE!')
            return 0
        if lrow is not None:
            ER = oRangeAddress.EndRow
            ER = LeenoComputo.circoscriveVoceComputo(oSheet, ER).RangeAddress.EndRow
        else:
            ER = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow
    if oSheet.Name == 'Analisi di Prezzo':
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
            DLG.MsgBox('La selezione deve essere contigua.', 'ATTENZIONE!')
            return 0
        if lrow is not None:
            ER = oRangeAddress.EndRow
            ER = LeenoComputo.circoscriveVoceComputo(oSheet, ER).RangeAddress.EndRow
        else:
            ER = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow
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
            DLG.MsgBox('La selezione deve essere contigua.', 'ATTENZIONE!')
            return 0
        if lrow is not None:
            ER = oRangeAddress.EndRow
            ER = LeenoComputo.circoscriveVoceComputo(oSheet, ER).RangeAddress.EndRow
        else:
            ER = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow
    return oDoc.CurrentController.select(
        oSheet.getCellRangeByPosition(0, SR, 50, ER))


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
    sStRange.RangeAddress
    SR = sStRange.RangeAddress.StartRow
    ER = sStRange.RangeAddress.EndRow
    # ~ oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, SR, 250, ER))
    return (SR, ER)


########################################################################


def MENU_elimina_voce():
    elimina_voce()


def elimina_voce(lrow=None, msg=1):
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoSheetUtils.eliminaVoce'
    Elimina una voce in COMPUTO, VARIANTE, CONTABILITA o Analisi di Prezzo
    lrow { long }  : numero riga
    msg  { bit }   : 1 chiedi conferma con messaggio
                     0 esegui senza conferma
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    SR = seleziona_voce()[0]
    ER = seleziona_voce()[1]
    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(
        0, SR, 250, ER))
    if msg == 1:
        if DLG.DlgSiNo(
                """OPERAZIONE NON ANNULLABILE!

Stai per eliminare la voce selezionata.
Vuoi Procedere?
 """, 'AVVISO!') == 2:
            delete('R')
            # ~ oSheet.getRows().removeByIndex(SR, ER-SR+1)
            numera_voci(0)
        else:
            return
    elif msg == 0:
        delete('R')
        # ~ oSheet.getRows().removeByIndex(SR, ER-SR+1)
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))


########################################################################
def copia_riga_computo(lrow):
    # ~def debug(lrow):
    '''
    Inserisce una nuova riga di misurazione nel computo
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    # ~lrow = LeggiPosizioneCorrente()[1]
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
        _gotoCella(2, lrow)
        # ~oDoc.CurrentController.select(oSheet.getCellByPosition(2, lrow))
        # ~oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))


def copia_riga_contab(lrow):
    '''
    Inserisce una nuova riga di misurazione in contabilità
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #  lrow = LeggiPosizioneCorrente()[1]
    stile = oSheet.getCellByPosition(1, lrow).CellStyle
    if oSheet.getCellByPosition(1,
                                lrow + 1).CellStyle == 'comp sotto Bianche_R':
        return
    if stile in ('comp Art-EP_R', 'Data_bianca', 'Comp-Bianche in mezzo_R'):
        lrow = lrow + 1  # PER INSERIMENTO SOTTO RIGA CORRENTE
        oSheet.getRows().insertByIndex(lrow, 1)
        # imposto gli stili
        oSheet.getCellByPosition(1, lrow).CellStyle = 'Comp-Bianche in mezzo_R'
        oSheet.getCellByPosition(2, lrow).CellStyle = 'comp 1-a'
        oSheet.getCellByPosition(5, lrow).CellStyle = 'comp 1-a PU'
        oSheet.getCellByPosition(6, lrow).CellStyle = 'comp 1-a LUNG'
        oSheet.getCellByPosition(7, lrow).CellStyle = 'comp 1-a LARG'
        oSheet.getCellByPosition(8, lrow).CellStyle = 'comp 1-a peso'
        oSheet.getCellRangeByPosition(
            11, lrow, 23, lrow).CellStyle = 'Comp-Bianche in mezzo_R'
        oSheet.getCellByPosition(8, lrow).CellStyle = 'comp 1-a peso'
        oSheet.getCellRangeByPosition(9, lrow, 11, lrow).CellStyle = 'Blu'
        # ci metto le formule
        oSheet.getCellByPosition(
            9, lrow).Formula = '=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(
                lrow + 1) + ')<=0;"";PRODUCT(E' + str(lrow +
                                                      1) + ':I' + str(lrow +
                                                                      1) + '))'
        # ~ oSheet.getCellByPosition(11, lrow).Formula = '=IF(PRODUCT(E' +
        # str(lrow+1) + ':I' + str(lrow+1) + ')>=0;"";PRODUCT(E' +
        # str(lrow+1) + ':I' + str(lrow+1) + ')*-1)'
        # preserva la data di misura
        if oSheet.getCellByPosition(1, lrow + 1).CellStyle == 'Data_bianca':
            oRangeAddress = oSheet.getCellByPosition(1, lrow +
                                                     1).getRangeAddress()
            oCellAddress = oSheet.getCellByPosition(1, lrow).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
            oSheet.getCellByPosition(1, lrow + 1).String = ""
            oSheet.getCellByPosition(1, lrow +
                                     1).CellStyle = 'Comp-Bianche in mezzo_R'
        _gotoCella(2, lrow)


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
        oSheet.getCellByPosition(0, lrow).String = 'Cod. Art.?'
    _gotoCella(1, lrow)

########################################################################


def MENU_Copia_riga_Ent():
    '''
    @@ DA DOCUMENTARE
    '''
    Copia_riga_Ent()


def Copia_riga_Ent(arg=None):
    '''
    @@ DA DOCUMENTARE
    '''
    # A ggiungi Componente - capisce su quale tipologia di tabelle è
    # ~datarif = datetime.now()
    #  refresh(0)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]
    nome_sheet = oSheet.Name
    if nome_sheet in ('COMPUTO', 'VARIANTE'):
        if cfg.read('Generale', 'dettaglio') == '1':
            dettaglio_misura_rigo()
        copia_riga_computo(lrow)
    elif nome_sheet == 'CONTABILITA':
        if cfg.read('Generale', 'dettaglio') == '1':
            dettaglio_misura_rigo()
        copia_riga_contab(lrow)
    elif nome_sheet == 'Analisi di Prezzo':
        copia_riga_analisi(lrow)
    #  refresh(1)
    # ~MsgBox('eseguita in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!','')


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
    Controlla che non ci siano atti contabili registrati e dà il consenso a procedere.
    '''
    partenza = LeenoUtils.getGlobalVar('partenza')
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in ('CONTABILITA'):
        partenza = cerca_partenza()
        DLG.chi(partenza[2])
        DLG.chi(LeenoUtils.getGlobalVar('sblocca_computo'))
        if LeenoUtils.getGlobalVar('sblocca_computo') == 1:
            pass
        else:
            if partenza[2] == '':
                pass
            if partenza[2] == '#reg':
                if DLG.DlgSiNo(
                        """Lavorando in questo punto del foglio,
comprometterai la validità degli atti contabili già emessi.

Vuoi procedere?

SCEGLIENDO SI' SARAI COSTRETTO A RIGENERARLI!""", 'Voce già registrata!') == 3:
                    pass
                else:
                    LeenoUtils.setGlobalVar('sblocca_computo', 1)
        DLG.chi(LeenoUtils.getGlobalVar('sblocca_computo'))


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
    if oSheet.Name in ('COMPUTO', 'CONTABILITA', 'VARIANTE',
                       'Analisi di Prezzo'):
        if oSheet.Name == 'Analisi di Prezzo':
            if oSheet.getCellByPosition(
                    0, lrow).CellStyle in ('An-lavoraz-Cod-sx', 'An-1_sigla'):
                codice_da_cercare = oSheet.getCellByPosition(0, lrow).String
            else:
                return
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
        # ~_gotoCella(oCell[0], oCell[1])
        oDoc.CurrentController.select(
            oSheet.getCellRangeByPosition(oCell[0], oCell[1], 30, oCell[1]))

########################################################################


def MENU_pesca_cod():
    '''
    @@ DA DOCUMENTARE
    '''
    pesca_cod()


def pesca_cod():
    #  def debug():
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
        partenza = cerca_partenza()
        cerca_in_elenco()
        GotoSheet('Elenco Prezzi')
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
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        if oDoc.NamedRanges.hasByName("#Lib#1"):
            if LeenoUtils.getGlobalVar('sblocca_computo') == 0:
                if DLG.DlgSiNo(
                        "Risulta già registrato un SAL. VUOI PROCEDERE COMUQUE?",
                        'ATTENZIONE!') == 3:
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

        stili_contab = LeenoUtils.getGlobalVar('stili_contab')

        if oSheet.getCellByPosition(0, lrow).CellStyle not in stili_contab + (
                'comp Int_colonna_R_prima', ):
            return
        ins_voce_contab(arg=0)
        partenza = cerca_partenza()
        GotoSheet(cfg.read('Contabilità', 'ricicla_da'))
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        lrow = LeggiPosizioneCorrente()[1]
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        sopra = sStRange.RangeAddress.StartRow + 2
        sotto = sStRange.RangeAddress.EndRow - 1

        oSrc = oSheet.getCellRangeByPosition(2, sopra, 8,
                                             sotto).getRangeAddress()
        partenza = LeenoUtils.getGlobalVar('partenza')
        if partenza is None:
            return
        oDest = oDoc.getSheets().getByName('CONTABILITA')
        oCellAddress = oDest.getCellByPosition(2, partenza[1] + 1).getCellAddress()
        GotoSheet('CONTABILITA')
        for n in range(sopra, sotto):
            copia_riga_contab(partenza[1])
            #  if oDest.getCellByPosition(2, n).CellStyle == 'comp 1-a ROSSO':
            #  chi(n -3 + partenza[1])
            #  chi(partenza)
            #  LeenoSheetUtils.invertiUnSegno(oDest, n - 2 + partenza[1])
        oDest.copyRange(oCellAddress, oSrc)
        oDest.getCellByPosition(1, partenza[1]).String = oSheet.getCellByPosition(1, sopra - 1).String
        parziale_verifica()
        start = LeggiPosizioneCorrente()[1]
        end = sotto - sopra + start + 1
        for n in range(start, end):
            #  chi(n)
            if oDest.getCellByPosition(2, n).CellStyle == 'comp 1-a ROSSO':

                LeenoSheetUtils.invertiUnSegno(oDest, n)
        _gotoCella(2, partenza[1] + 1)

def MENU_inverti_segno():
    inverti_segno()


def inverti_segno():
    '''
    Inverte il segno delle formule di quantità nei righi di misurazione selezionati.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lista = list()
    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
    el_y = list()
    try:
        len(oRangeAddress)
        for el in oRangeAddress:
            el_y.append((el.StartRow, el.EndRow))
    except TypeError:
        el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))
    for y in el_y:
        for el in range(y[0], y[1] + 1):
            lista.append(el)
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        for lrow in lista:
            if 'comp 1-a' in oSheet.getCellByPosition(2, lrow).CellStyle:
                if 'ROSSO' in oSheet.getCellByPosition(2, lrow).CellStyle:
                    oSheet.getCellByPosition(
                        9, lrow
                    ).Formula = '=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(
                        lrow + 1) + ')=0;"";PRODUCT(E' + str(
                            lrow + 1) + ':I' + str(lrow +
                                                   1) + '))'  # se VediVoce
                    # ~ if oSheet.getCellByPosition(4, lrow).Type.value != 'EMPTY':
                    # ~ oSheet.getCellByPosition(9, lrow).Formula=
                    # '=IF(PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + ')=0;
                    # "";PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + '))' # se VediVoce
                    # ~ else:
                    # ~ oSheet.getCellByPosition(9, lrow).Formula=
                    # '=IF(PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) +
                    # ')=0;"";PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + '))'
                    for x in range(2, 9):
                        oSheet.getCellByPosition(
                            x, lrow).CellStyle = oSheet.getCellByPosition(
                                x, lrow).CellStyle.split(' ROSSO')[0]
                else:
                    oSheet.getCellByPosition(
                        9, lrow
                    ).Formula = '=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(
                        lrow + 1) + ')=0;"";-PRODUCT(E' + str(
                            lrow + 1) + ':I' + str(lrow +
                                                   1) + '))'  # se VediVoce
                    # ~ if oSheet.getCellByPosition(4, lrow).Type.value != 'EMPTY':
                    # ~ oSheet.getCellByPosition(9, lrow).Formula =
                    # '=IF(PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) +
                    # ')=0;"";-PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + '))' # se VediVoce
                    # ~ else:
                    # ~ oSheet.getCellByPosition(9, lrow).Formula =
                    # '=IF(PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) +
                    # ')=0;"";-PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + '))'
                    for x in range(2, 9):
                        oSheet.getCellByPosition(
                            x, lrow).CellStyle = oSheet.getCellByPosition(
                                x, lrow).CellStyle + ' ROSSO'
    if oSheet.Name in ('CONTABILITA'):
        for lrow in lista:
            if 'comp 1-a' in oSheet.getCellByPosition(2, lrow).CellStyle:
                formula1 = oSheet.getCellByPosition(9, lrow).Formula
                formula2 = oSheet.getCellByPosition(11, lrow).Formula
                oSheet.getCellByPosition(11, lrow).Formula = formula1
                oSheet.getCellByPosition(9, lrow).Formula = formula2
                if oSheet.getCellByPosition(11, lrow).Value > 0:
                    for x in range(2, 12):
                        oSheet.getCellByPosition(
                            x, lrow).CellStyle = oSheet.getCellByPosition(
                                x, lrow).CellStyle + ' ROSSO'
                else:
                    for x in range(2, 12):
                        oSheet.getCellByPosition(
                            x, lrow).CellStyle = oSheet.getCellByPosition(
                                x, lrow).CellStyle.split(' ROSSO')[0]


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
    # ~def debug():
    '''
    Indica il dettaglio delle misure nel rigo di descrizione quando
    incontra delle formule nei valori immessi.
    bit { integer }  : 1 inserisce i dettagli
                       0 cancella i dettagli
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]
    if ' >(' in oSheet.getCellByPosition(2, lrow).String:
        oSheet.getCellByPosition(2, lrow).String = oSheet.getCellByPosition(
            2, lrow).String.split(' >(')[0]
    if oSheet.getCellByPosition(2, lrow).CellStyle in (
            'comp 1-a'
    ) and "*** VOCE AZZERATA ***" not in oSheet.getCellByPosition(2,
                                                                  lrow).String:
        for el in range(5, 9):
            if oSheet.getCellByPosition(el, lrow).Type.value == 'FORMULA':
                stringa = ''
                break
            else:
                stringa = None

        if stringa == '':
            for el in range(5, 9):
                # test = '>('
                if oSheet.getCellByPosition(el, lrow).Type.value == 'FORMULA':
                    if '$' not in oSheet.getCellByPosition(el, lrow).Formula:
                        try:
                            eval(
                                oSheet.getCellByPosition(
                                    el, lrow).Formula.split('=')[1].replace(
                                        '^', '**'))
                            # ~eval(oSheet.getCellByPosition(el, lrow).Formula.split('=')[1])
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
            stringa = ' >(' + stringa + ')'
            if oSheet.getCellByPosition(2, lrow).Type.value != 'FORMULA':
                oSheet.getCellByPosition(
                    2, lrow).String = oSheet.getCellByPosition(
                        2, lrow).String + stringa.replace('.', ',')


########################################################################
def dettaglio_misure(bit):
    '''
    Indica il dettaglio delle misure nel rigo di descrizione quando
    incontra delle formule nei valori immessi.
    bit { integer }  : 1 inserisce i dettagli
                       0 cancella i dettagli
    '''
    oDoc = LeenoUtils.getDocument()
    try:
        oSheet = oDoc.CurrentController.ActiveSheet
    except Exception:
        return
    ER = SheetUtils.getUsedArea(oSheet).EndRow
    if bit == 1:
        for lrow in range(0, ER):
            if oSheet.getCellByPosition(2, lrow).CellStyle in (
                    'comp 1-a'
            ) and "*** VOCE AZZERATA ***" not in oSheet.getCellByPosition(
                    2, lrow).String:
                for el in range(5, 9):
                    if oSheet.getCellByPosition(el, lrow).Type.value == 'FORMULA':
                        stringa = ''
                        break
                    else:
                        stringa = None
                if stringa == '':
                    for el in range(5, 9):
                        # test = '>('
                        if oSheet.getCellByPosition(
                                el, lrow).Type.value == 'FORMULA':
                            if '$' not in oSheet.getCellByPosition(
                                    el, lrow).Formula:
                                try:
                                    eval(
                                        oSheet.getCellByPosition(
                                            el, lrow).Formula.split('=')
                                        [1].replace('^', '**'))
                                    # ~eval(oSheet.getCellByPosition(el, lrow).Formula.split('=')[1])
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
                    stringa = ' >(' + stringa + ')'
                    if oSheet.getCellByPosition(2,
                                                lrow).Type.value != 'FORMULA':
                        oSheet.getCellByPosition(
                            2, lrow).String = oSheet.getCellByPosition(
                                2, lrow).String + stringa.replace('.', ',')
    else:
        for lrow in range(0, ER):
            if ' >(' in oSheet.getCellByPosition(2, lrow).String:
                oSheet.getCellByPosition(
                    2, lrow).String = oSheet.getCellByPosition(
                        2, lrow).String.split(' >(')[0]
    return


########################################################################
def debug_validation():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #  mri(oDoc.CurrentSelection.Validation)

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
def debugclip():
    #  mri(LeenoUtils.getComponentContext())
    # sText = 'sticazzi'
    # create SystemClipboard instance
    oClip = LeenoUtils.createUnoService(
        "com.sun.star.datatransfer.clipboard.SystemClipboard")
    # oClipContents = oClip.getContents()
    # flavors = oClipContents.getTransferDataFlavors()
    DLG.mri(oClip)
    #  for i in flavors:
    #  aDataFlavor = flavors(i)
    #  chi(aDataFlavor)

    return
    #  createUnoService =(LeenoUtils.getComponentContext().getServiceManager().createInstance)
    #  oTR = createUnoListener("Tr_", "com.sun.star.datatransfer.XTransferable")
    oClip.setContents(oTR, None)
    # sTxtCString = sText
    oClip.flushClipboard()


########################################################################
def rimuovi_area_di_stampa():
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, ".uno:DeletePrintArea", "", 0,
                                   list())


########################################################################
def visualizza_PageBreak():
    '''
    @@ DA DOCUMENTARE
    '''
    # oDoc = LeenoUtils.getDocument()
    #  oSheet = oDoc.CurrentController.ActiveSheet
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    oProp = PropertyValue()
    oProp.Name = 'PagebreakMode'
    oProp.Value = True
    properties = (oProp, )

    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, ".uno:PagebreakMode", "", 0, properties)


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
def copy_clip():
    #  oDoc = LeenoUtils.getDocument()
    #  oSheet = oDoc.CurrentController.ActiveSheet
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()

    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, ".uno:Copy", "", 0, list())


########################################################################
def paste_clip(arg=None, insCells=0):
    oDoc = LeenoUtils.getDocument()
    #  oSheet = oDoc.CurrentController.ActiveSheet
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    oProp = []
    oProp0 = PropertyValue()
    oProp0.Name = 'Flags'
    oProp0.Value = 'A'
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
    # insert mode ON
    if insCells == 1:
        oProp5 = PropertyValue()
        oProp5.Name = 'MoveMode'
        oProp5.Value = 0
        oProp.append(oProp5)
    properties = tuple(oProp)

    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, '.uno:InsertContents', '', 0, properties)
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect


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
    righe = list()
    colonne = list()
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
    dispatchHelper.executeDispatch(oFrame, ".uno:Copy", "", 0, list())
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
    numera_voci(1)


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
            # ~if oSheet.getCellByPosition(1,row).CellBackColor == 15066597:
            # ~oSheet.getCellByPosition(0,row).String = ''
            # ~elif oSheet.getCellByPosition(1,row).CellStyle in('comp Art-EP', 'comp Art-EP_R'):
            # ~oSheet.getCellByPosition(0,row).Value = n
            # ~n = n+1
            if oSheet.getCellByPosition(1, row).CellStyle in ('comp Art-EP','comp Art-EP_R'):
                oSheet.getCellByPosition(0, row).Value = n
                n = n + 1
            # ~oSheet.getCellByPosition(0,row).Value = n
            # ~n = n+1


########################################################################
def DisableAutoCalc():
    '''
    Disabilita il refresh per accelerare le procedure
    '''
    oDoc = LeenoUtils.getDocument()
    # blocco il calcolo automatico
    oDoc.enableAutomaticCalculation(False)


def EnableAutoCalc():
    '''
    Riabilita il refresh
    '''
    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(True)

########################################################################
def richiesta_offerta():
    #  def debug():
    '''Crea la Lista Lavorazioni e Forniture dall'Elenco Prezzi,
per la formulazione dell'offerta'''
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    GotoSheet('Elenco Prezzi')
    genera_sommario()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        oDoc.Sheets.copyByName(oSheet.Name, 'Elenco Prezzi', 5)
    except Exception:
        pass
    nSheet = oDoc.getSheets().getByIndex(5).Name
    GotoSheet(nSheet)
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.Name = 'Richiesta offerta'
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
    for x in range(3, SheetUtils.getUsedArea(oSheet).EndRow - 1):
        formule.append([
            '=IF(E' + str(x + 1) + '<>"";D' + str(x + 1) + '*E' + str(x + 1) +
            ';""'
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
    oSheet.Columns.insertByIndex(0, 1)

    oSrc = oSheet.getCellRangeByPosition(1, 0, 1, fine).RangeAddress
    oDest = oSheet.getCellByPosition(0, 0).CellAddress
    oSheet.copyRange(oDest, oSrc)
    oSheet.getCellByPosition(0, 2).String = "N."
    for x in range(3, fine - 1):
        oSheet.getCellByPosition(0, x).Value = x - 2
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
    oSheet.Rows.removeByIndex(fine - 1, 1)
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
    adatta_altezza_riga(nSheet)
    pagestyle.RightPageHeaderContent = oHContent
    _gotoCella(0, 1)
    return


########################################################################
def ins_voce_elenco():
    '''
    Inserisce una nuova riga voce in Elenco Prezzi
    '''
    oDoc = LeenoUtils.getDocument()
    DisableAutoCalc()

    oSheet = oDoc.CurrentController.ActiveSheet
    _gotoCella(0, 3)
    oSheet.getRows().insertByIndex(3, 1)

    oSheet.getCellByPosition(0, 3).CellStyle = "EP-aS"
    oSheet.getCellByPosition(1, 3).CellStyle = "EP-a"
    oSheet.getCellRangeByPosition(2, 3, 7, 3).CellStyle = "EP-mezzo"
    oSheet.getCellRangeByPosition(8, 3, 9, 3).CellStyle = "EP-sfondo"
    for el in (5, 11, 15, 19, 26):
        oSheet.getCellByPosition(el, 3).CellStyle = "EP-mezzo %"

    for el in (12, 16, 20, 21):  # (12, 16, 20):
        oSheet.getCellByPosition(el, 3).CellStyle = 'EP statistiche_q'

    for el in (13, 17, 23, 24, 25):  # (12, 16, 20):
        oSheet.getCellByPosition(el, 3).CellStyle = 'EP statistiche'

    oSheet.getCellRangeByPosition(0, 3, 26, 3).clearContents(HARDATTR)
    oSheet.getCellByPosition(11,
                             3).Formula = '=IF(ISERROR(N4/$N$2);"--";N4/$N$2)'
    #  oSheet.getCellByPosition(11, 3).Formula = '=N4/$N$2'
    oSheet.getCellByPosition(12, 3).Formula = '=SUMIF(AA;A4;BB)'
    oSheet.getCellByPosition(13, 3).Formula = '=SUMIF(AA;A4;cEuro)'

    # copio le formule dalla riga sotto
    oRangeAddress = oSheet.getCellRangeByPosition(15, 4, 26,
                                                  4).getRangeAddress()
    oCellAddress = oSheet.getCellByPosition(15, 3).getCellAddress()
    oSheet.copyRange(oCellAddress, oRangeAddress)
    oCell = oSheet.getCellByPosition(2, 3)
    valida_cella(
        oCell,
        '"cad";"corpo";"dm";"dm²";"dm³";"kg";"lt";"m";"m²";"m³";"q";"t";"',
        titoloInput='Scegli...',
        msgInput='Unità di misura')

    EnableAutoCalc()


########################################################################


########################################################################
def rigenera_voce(lrow=None):
    '''
    Ripristina/ricalcola le formule di descrizione e somma di una voce.
    in COMPUTO, VARIANTE e CONTABILITA
    '''
    lrow = LeggiPosizioneCorrente()[1]
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    except Exception:
        return
    sopra = sStRange.RangeAddress.StartRow
    sotto = sStRange.RangeAddress.EndRow
    for n in range(sopra + 2, sotto):
        if oSheet.Name in ('COMPUTO', 'VARIANTE'):
            oSheet.getCellByPosition(
                2, sopra + 1
            ).Formula = '=IF(LEN(VLOOKUP(B' + str(
                sopra + 2
            ) + ';elenco_prezzi;2;FALSE()))<($S1.$H$337+$S1.$H$338);VLOOKUP(B' + str(
                sopra + 2
            ) + ';elenco_prezzi;2;FALSE());CONCATENATE(LEFT(VLOOKUP(B' + str(
                sopra + 2
            ) + ';elenco_prezzi;2;FALSE());$S1.$H$337);" [...] ";RIGHT(VLOOKUP(B' + str(
                sopra + 2) + ';elenco_prezzi;2;FALSE());$S1.$H$338)))'
            oSheet.getCellByPosition(
                8, sotto).Formula = '=CONCATENATE("SOMMANO [";VLOOKUP(B' + str(
                    sopra + 2) + ';elenco_prezzi;3;FALSE());"]")'
            oSheet.getCellByPosition(
                9,
                sotto).Formula = '=SUBTOTAL(9;J' + str(sopra +
                                                       2) + ':J' + str(sotto +
                                                                       1) + ')'
            oSheet.getCellByPosition(11, sotto).Formula = '=VLOOKUP(B' + str(
                sopra + 2) + ';elenco_prezzi;5;FALSE())'
            oSheet.getCellByPosition(13, sotto).Formula = '=J' + str(sotto + 1)
            oSheet.getCellByPosition(
                17,
                sotto).Formula = '=AB' + str(sotto + 1) + '*J' + str(sotto + 1)
            #  oSheet.getCellByPosition(18, sotto).Formula = '=J'+ str(sotto+1) +'*L'+ str(sotto+1)
            oSheet.getCellByPosition(
                18, sotto).Formula = '=IF(VLOOKUP(B' + str(
                    sopra + 2) + ';elenco_prezzi;3;FALSE())="%";J' + str(
                        sotto + 1) + '*L' + str(sotto + 1) + '/100;J' + str(
                            sotto + 1) + '*L' + str(sotto + 1) + ')'
            oSheet.getCellByPosition(27, sotto).Formula = '=VLOOKUP(B' + str(
                sopra + 2) + ';elenco_prezzi;4;FALSE())'
            oSheet.getCellByPosition(
                28,
                sotto).Formula = '=S' + str(sotto + 1) + '-AE' + str(sotto + 1)
            oSheet.getCellByPosition(29, sotto).Formula = '=VLOOKUP(B' + str(
                sopra + 2) + ';elenco_prezzi;6;FALSE())'
            oSheet.getCellByPosition(
                30, sotto
            ).Formula = '=IF(AD' + str(sotto + 1) + '<>""; PRODUCT(AD' + str(
                sotto + 1) + '*S' + str(sotto + 1) + '))'
            oSheet.getCellByPosition(35, sotto).Formula = '=B' + str(sopra + 2)
            oSheet.getCellByPosition(
                36, sotto
            ).Formula = '=IF(ISERROR(S' + str(sotto + 1) + ');"";IF(S' + str(
                sotto + 1) + '<>"";S' + str(sotto + 1) + ';""))'
            if 'comp 1-a' in (
                    oSheet.getCellByPosition(2, n).CellStyle
            ):  # and oSheet.getCellByPosition(9, n).Type.value != 'FORMULA':
                oSheet.getCellByPosition(9, n).Formula = '=IF(PRODUCT(E' + str(
                    n + 1) + ':I' + str(n + 1) + ')=0;"";PRODUCT(E' + str(
                        n + 1) + ':I' + str(n + 1) + '))'

        if oSheet.Name in ('CONTABILITA'):
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


########################################################################
def sistema_stili(lrow=None):
    '''
    Ripristina stili di cella per una singola voce.
    in COMPUTO, VARIANTE e CONTABILITA
    '''
    lrow = LeggiPosizioneCorrente()[1]
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    except Exception:
        return
    sopra = sStRange.RangeAddress.StartRow
    sotto = sStRange.RangeAddress.EndRow
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        for x in range(sopra + 1, sotto - 1):
            if 'comp 1-a' in oSheet.getCellByPosition(2, x).CellStyle:
                oSheet.getCellByPosition(9, x).CellStyle = 'Blu'
                if oSheet.getCellByPosition(9, x).Value < 0:
                    oSheet.getCellByPosition(2, x).CellStyle = 'comp 1-a ROSSO'
                    oSheet.getCellByPosition(5,
                                             x).CellStyle = 'comp 1-a PU ROSSO'
                    oSheet.getCellByPosition(
                        6, x).CellStyle = 'comp 1-a LUNG ROSSO'
                    oSheet.getCellByPosition(
                        7, x).CellStyle = 'comp 1-a LARG ROSSO'
                    oSheet.getCellByPosition(
                        8, x).CellStyle = 'comp 1-a peso ROSSO'
                else:
                    oSheet.getCellByPosition(2, x).CellStyle = 'comp 1-a'
                    oSheet.getCellByPosition(5, x).CellStyle = 'comp 1-a PU'
                    oSheet.getCellByPosition(6, x).CellStyle = 'comp 1-a LUNG'
                    oSheet.getCellByPosition(7, x).CellStyle = 'comp 1-a LARG'
                    oSheet.getCellByPosition(8, x).CellStyle = 'comp 1-a peso'

    if oSheet.Name in ('CONTABILITA'):
        oSheet.getCellByPosition(9, sopra + 1).CellStyle = 'vuote2'
        oSheet.getCellByPosition(11, sopra +
                                 1).CellStyle = 'Comp-Bianche in mezzo_R'
        oSheet.getCellByPosition(9, sotto -
                                 1).CellStyle = 'Comp-Variante num sotto'
        oSheet.getCellByPosition(9,
                                 sotto).CellStyle = 'Comp-Variante num sotto'
        oSheet.getCellByPosition(13, sotto).CellStyle = 'comp sotto Unitario'
        oSheet.getCellByPosition(15,
                                 sotto).CellStyle = 'comp sotto Euro Originale'
        oSheet.getCellByPosition(17,
                                 sotto).CellStyle = 'comp sotto Euro Originale'
        oSheet.getCellByPosition(11, sotto -
                                 1).CellStyle = 'Comp-Variante num sotto ROSSO'
        oSheet.getCellByPosition(11, sotto).CellStyle = 'comp sotto centro_R'
        oSheet.getCellByPosition(28, sotto).CellStyle = 'Comp-sotto euri'
        for x in range(sopra + 1, sotto):
            oSheet.getCellByPosition(11, x).CellStyle = 'Blu ROSSO'
            if 'comp 1-a' in oSheet.getCellByPosition(2, x).CellStyle:
                oSheet.getCellByPosition(9, x).CellStyle = 'Blu'
            elif oSheet.getCellByPosition(
                    2, x).CellStyle == 'comp sotto centro':  # parziale
                oSheet.getCellByPosition(
                    9, x).CellStyle = 'Comp-Variante num sotto'

        for x in range(sopra + 2, sotto - 1):
            test = 0
            for y in range(2, 8):
                if oSheet.getCellByPosition(y, x).String != '':
                    test = 1
                    break
            rosso = 0
            for y in range(2, 8):
                if 'ROSSO' in oSheet.getCellByPosition(y, x).CellStyle:
                    rosso = 1
                    break
            if str(test) + str(rosso) == '10':
                oSheet.getCellByPosition(9, x).Formula = '=IF(PRODUCT(E' + str(
                    x + 1) + ':I' + str(x + 1) + ')=0;"";PRODUCT(E' + str(
                        x + 1) + ':I' + str(x + 1) + '))'
                oSheet.getCellByPosition(2, x).CellStyle = 'comp 1-a'
                oSheet.getCellByPosition(5, x).CellStyle = 'comp 1-a PU'
                oSheet.getCellByPosition(6, x).CellStyle = 'comp 1-a LUNG'
                oSheet.getCellByPosition(7, x).CellStyle = 'comp 1-a LARG'
                oSheet.getCellByPosition(8, x).CellStyle = 'comp 1-a peso'
                oSheet.getCellByPosition(11, x).String = ''
            if str(test) + str(rosso) == '11':
                oSheet.getCellByPosition(
                    11, x).Formula = '=IF(PRODUCT(E' + str(x + 1) + ':I' + str(
                        x + 1) + ')=0;"";PRODUCT(E' + str(x + 1) + ':I' + str(
                            x + 1) + '))'
                oSheet.getCellByPosition(2, x).CellStyle = 'comp 1-a ROSSO'
                oSheet.getCellByPosition(5, x).CellStyle = 'comp 1-a PU ROSSO'
                oSheet.getCellByPosition(6,
                                         x).CellStyle = 'comp 1-a LUNG ROSSO'
                oSheet.getCellByPosition(7,
                                         x).CellStyle = 'comp 1-a LARG ROSSO'
                oSheet.getCellByPosition(8,
                                         x).CellStyle = 'comp 1-a peso ROSSO'
                oSheet.getCellByPosition(9, x).String = ''


########################################################################
def rigenera_tutte(arg=None, ):
    '''
    Ripristina le formule in tutto il foglio
    '''
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    DisableAutoCalc()

    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400
    nome = oSheet.Name
    oDialogo_attesa = DLG.dlg_attesa('Rigenerazione delle formule in ' + oSheet.Name + '...')
    DLG.attesa().start()  # mostra il dialogo
    # ~oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, 0, 30, 0))
    if nome in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        try:
            oSheet = oDoc.Sheets.getByName(nome)
            row = next_voice(0, 1)
            last = LeenoSheetUtils.cercaUltimaVoce(oSheet)
            while row < last:
                oDoc.CurrentController.select(
                    oSheet.getCellRangeByPosition(0, row, 30, row))
                rigenera_voce(row)
                row = next_voice(row, 1)
            Rinumera_TUTTI_Capitoli2()
        except Exception:
            pass
    oDoc.CurrentController.ZoomValue = zoom
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect
    oDialogo_attesa.endExecute()

    EnableAutoCalc()


########################################################################
def MENU_nuova_voce_scelta():  # assegnato a ctrl-shift-n
    '''
    Contestualizza in ogni tabella l'inserimento delle voci.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        LeenoComputo.ins_voce_computo()
    elif oSheet.Name == 'Analisi di Prezzo':
        inizializza_analisi()
    elif oSheet.Name == 'CONTABILITA':
        ins_voce_contab()
    elif oSheet.Name == 'Elenco Prezzi':
        ins_voce_elenco()


# nuova_voce_contab  ##################################################
def ins_voce_contab(lrow=0, arg=1):
    '''
    Inserisce una nuova voce in CONTABILITA.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    stili_contab = LeenoUtils.getGlobalVar('stili_contab')
    if lrow == 0:
        lrow = LeggiPosizioneCorrente()[1]
    # nome = oSheet.Name
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
    if stile == 'comp Int_colonna_R_prima':
        lrow += 1
    elif stile == 'Ultimus_centro_bordi_lati':
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
    elif stile in stili_contab:
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        nSal = int(
            oSheet.getCellByPosition(23,
                                     sStRange.RangeAddress.StartRow + 1).Value)
        lrow = next_voice(lrow)
    else:
        return
    oSheetto = oDoc.getSheets().getByName('S5')
    oRangeAddress = oSheetto.getCellRangeByPosition(0, 22, 48,
                                                    26).getRangeAddress()
    oCellAddress = oSheet.getCellByPosition(0, lrow).getCellAddress()
    oSheet.getRows().insertByIndex(lrow, 5)  # inserisco le righe
    oSheet.copyRange(oCellAddress, oRangeAddress)
    oSheet.getCellRangeByPosition(0, lrow, 48,
                                  lrow + 5).Rows.OptimalHeight = True
    _gotoCella(1, lrow + 1)

    #  if(oSheet.getCellByPosition(0,lrow).queryIntersection(oSheet.getCellRangeByName('#Lib#'+str(nSal)).getRangeAddress())):
    #  chi('appartiene')
    #  else:
    #  chi('nooooo')
    #  return

    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    sopra = sStRange.RangeAddress.StartRow
    for n in reversed(range(0, sopra)):
        if oSheet.getCellByPosition(
                1, n).CellStyle == 'Ultimus_centro_bordi_lati':
            break
        if oSheet.getCellByPosition(1, n).CellStyle == 'Data_bianca':
            data = oSheet.getCellByPosition(1, n).Value
            break
    try:
        oSheet.getCellByPosition(1, sopra + 2).Value = data
    except Exception:
        oSheet.getCellByPosition(1, sopra +
                                 2).Value = date.today().toordinal() - 693594
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

    if oDoc.NamedRanges.hasByName('#Lib#' + str(nSal)):
        if lrow - 1 == oSheet.getCellRangeByName(
                '#Lib#' + str(nSal)).getRangeAddress().EndRow:
            nSal += 1

    oSheet.getCellByPosition(23, sopra + 1).Value = nSal
    oSheet.getCellByPosition(23, sopra + 1).CellStyle = 'Sal'

    oSheet.getCellByPosition(35, sopra + 4).Formula = '=B' + str(sopra + 2)
    oSheet.getCellByPosition(
        36, sopra +
        4).Formula = '=IF(ISERROR(P' + str(sopra + 5) + ');"";IF(P' + str(
            sopra + 5) + '<>"";P' + str(sopra + 5) + ';""))'
    oSheet.getCellByPosition(36, sopra + 4).CellStyle = "comp -controolo"
    numera_voci(0)
    if cfg.read('Generale', 'pesca_auto') == '1':
        if arg == 0:
            return
        pesca_cod()


########################################################################
# CONTABILITA ## CONTABILITA ## CONTABILITA ## CONTABILITA ## CONTABILITA #
def attiva_contabilita():
    '''Se presenti, attiva e visualizza le tabelle di contabilità'''
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    if oDoc.Sheets.hasByName('S1'):
        oDoc.Sheets.getByName('S1').getCellByPosition(7, 327).Value = 1
        if oDoc.Sheets.hasByName('CONTABILITA'):
            for el in ('Registro', 'SAL', 'CONTABILITA'):
                if oDoc.Sheets.hasByName(el):
                    GotoSheet(el)
        else:
            oDoc.Sheets.insertNewByName('CONTABILITA', 5)
            GotoSheet('CONTABILITA')
            svuota_contabilita()
            ins_voce_contab()
            #  set_larghezza_colonne()
        GotoSheet('CONTABILITA')
    ScriviNomeDocumentoPrincipale()
    basic_LeenO("Menu.eventi_assegna")


########################################################################
def svuota_contabilita():
    '''Ricrea il foglio di contabilità partendo da zero.'''
    oDoc = LeenoUtils.getDocument()
    for n in range(1, 20):
        if oDoc.NamedRanges.hasByName('#Lib#' + str(n)):
            oDoc.NamedRanges.removeByName('#Lib#' + str(n))
            oDoc.NamedRanges.removeByName('#SAL#' + str(n))
            oDoc.NamedRanges.removeByName('#Reg#' + str(n))
    for el in ('Registro', 'SAL', 'CONTABILITA'):
        if oDoc.Sheets.hasByName(el):
            oDoc.Sheets.removeByName(el)

    oDoc.Sheets.insertNewByName('CONTABILITA', 3)
    oSheet = oDoc.Sheets.getByName('CONTABILITA')

    GotoSheet('CONTABILITA')
    setTabColor(16757935)
    oSheet.getCellRangeByName('C1').String = 'CONTABILITA'
    oSheet.getCellRangeByName('C1').CellStyle = 'comp Int_colonna'
    oSheet.getCellRangeByName('C1').CellBackColor = 16757935
    oSheet.getCellByPosition(0, 2).String = 'N.'
    oSheet.getCellByPosition(1, 2).String = 'Articolo\nData'
    oSheet.getCellByPosition(2, 2).String = 'LAVORAZIONI\nO PROVVISTE'
    oSheet.getCellByPosition(5, 2).String = 'P.U.\nCoeff.'
    oSheet.getCellByPosition(6, 2).String = 'Lung.'
    oSheet.getCellByPosition(7, 2).String = 'Larg.'
    oSheet.getCellByPosition(8, 2).String = 'Alt.\nPeso'
    oSheet.getCellByPosition(9, 2).String = 'Quantità\nPositive'
    oSheet.getCellByPosition(11, 2).String = 'Quantità\nNegative'
    oSheet.getCellByPosition(13, 2).String = 'Prezzo\nunitario'
    oSheet.getCellByPosition(15, 2).String = 'Importi'
    oSheet.getCellByPosition(16, 2).String = 'Incidenza\nsul totale'
    oSheet.getCellByPosition(17, 2).String = 'Sicurezza\ninclusa'
    oSheet.getCellByPosition(18, 2).String = 'importo totale\nsenza errori'
    oSheet.getCellByPosition(19, 2).String = 'Lib.\nN.'
    oSheet.getCellByPosition(20, 2).String = 'Lib.\nP.'
    oSheet.getCellByPosition(22, 2).String = 'flag'
    oSheet.getCellByPosition(23, 2).String = 'SAL\nN.'
    oSheet.getCellByPosition(25, 2).String = 'Importi\nSAL parziali'
    oSheet.getCellByPosition(27, 2).String = 'Sicurezza\nunitaria'
    oSheet.getCellByPosition(28, 2).String = 'Materiali\ne Noli €'
    oSheet.getCellByPosition(29, 2).String = 'Incidenza\nMdO %'
    oSheet.getCellByPosition(30, 2).String = 'Importo\nMdO'
    oSheet.getCellByPosition(31, 2).String = 'Super Cat'
    oSheet.getCellByPosition(32, 2).String = 'Cat'
    oSheet.getCellByPosition(33, 2).String = 'Sub Cat'
    #  oSheet.getCellByPosition(34,2).String = 'tag B'sub Scrivi_header_moduli
    #  oSheet.getCellByPosition(35,2).String = 'tag C'
    oSheet.getCellByPosition(36, 2).String = 'Importi\nsenza errori'
    oSheet.getCellByPosition(0, 2).Rows.Height = 800
    #  colore colonne riga di intestazione
    oSheet.getCellRangeByPosition(0, 2, 36, 2).CellStyle = 'comp Int_colonna_R'
    oSheet.getCellByPosition(0, 2).CellStyle = 'comp Int_colonna_R_prima'
    oSheet.getCellByPosition(18, 2).CellStyle = 'COnt_noP'
    oSheet.getCellRangeByPosition(0, 0, 0, 3).Rows.OptimalHeight = True
    #  riga di controllo importo
    oSheet.getCellRangeByPosition(0, 1, 36, 1).CellStyle = 'comp In testa'
    oSheet.getCellByPosition(2, 1).String = 'QUESTA RIGA NON VIENE STAMPATA'
    oSheet.getCellRangeByPosition(0, 1, 1, 1).merge(True)
    oSheet.getCellByPosition(13, 1).String = 'TOTALE:'
    oSheet.getCellByPosition(20, 1).String = 'SAL SUCCESSIVO:'

    oSheet.getCellByPosition(25, 1).Formula = '=$P$2-SUBTOTAL(9;$P$2:$P$2)'

    oSheet.getCellByPosition(15,
                             1).Formula = '=SUBTOTAL(9;P3:P4)'  # importo lavori
    oSheet.getCellByPosition(0, 1).Formula = '=AK2'  # importo lavori
    oSheet.getCellByPosition(
        17, 1).Formula = '=SUBTOTAL(9;R3:R4)'  # importo sicurezza

    oSheet.getCellByPosition(
        28, 1).Formula = '=SUBTOTAL(9;AC3:AC4)'  # importo materiali
    oSheet.getCellByPosition(29,
                             1).Formula = '=AE2/Z2'  # Incidenza manodopera %
    oSheet.getCellByPosition(29, 1).CellStyle = 'Comp TOTALI %'
    oSheet.getCellByPosition(
        30, 1).Formula = '=SUBTOTAL(9;AE3:AE4)'  # importo manodopera
    oSheet.getCellByPosition(
        36, 1).Formula = '=SUBTOTAL(9;AK3:AK4)'  # importo certo

    #  rem riga del totale
    oSheet.getCellByPosition(2, 3).String = 'T O T A L E'
    oSheet.getCellByPosition(15,
                             3).Formula = '=SUBTOTAL(9;P3:P4)'  # importo lavori
    oSheet.getCellByPosition(
        17, 3).Formula = '=SUBTOTAL(9;R3:R4)'  # importo sicurezza
    oSheet.getCellByPosition(
        30, 3).Formula = '=SUBTOTAL(9;AE3:AE4)'  # importo manodopera
    oSheet.getCellRangeByPosition(0, 3, 36, 3).CellStyle = 'Comp TOTALI'
    #  rem riga rossa
    oSheet.getCellByPosition(0, 4).String = 'Fine Computo'
    oSheet.getCellRangeByPosition(0, 4, 36, 4).CellStyle = 'Riga_rossa_Chiudi'
    _gotoCella(0, 2)
    set_larghezza_colonne()


########################################################################
def partita(testo):
    '''
    Aggiunge/detrae rigo di PARTITA PROVVISORIA
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name != "CONTABILITA":
        return
    x = LeggiPosizioneCorrente()[1]
    if oSheet.getCellByPosition(0, x).CellStyle == 'comp 10 s_R':
        if oSheet.getCellByPosition(2, x).Type.value != 'EMPTY':
            Copia_riga_Ent()
            x += 1
        oSheet.getCellByPosition(2, x).String = testo
        oSheet.getCellRangeByPosition(2, x, 8, x).CellBackColor = 16777113
        _gotoCella(5, x)


def MENU_partita_aggiungi():
    '''
    @@ DA DOCUMENTARE
    '''
    partita('PARTITA PROVVISORIA')


def MENU_partita_detrai():
    '''
    @@ DA DOCUMENTARE
    '''
    partita('SI DETRAE PARTITA PROVVISORIA')


########################################################################
def genera_libretto():
    '''
    @@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()
    #  mri(oDoc.StyleFamilies.getByName("CellStyles").getByName('comp 1-a PU'))
    #  return
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name != 'CONTABILITA':
        return
    numera_voci()
    oRanges = oDoc.NamedRanges
    nSal = 1
    if oRanges.hasByName("#Lib#1"):
        nSal = 20
    else:
        try:
            nSal = int(cfg.read('Contabilità', 'idxsal'))
        except ValueError:
            nSal = 20
    while nSal > 0:
        if oRanges.hasByName("#Lib#" + str(nSal)):
            break
        nSal = nSal - 1
    #  Recupero la prima riga non registrata
    oSheetCont = oDoc.Sheets.getByName('CONTABILITA')
    if nSal >= 1:
        oNamedRange = oRanges.getByName("#Lib#" +
                                        str(nSal)).ReferredCells.RangeAddress
        frow = oNamedRange.StartRow
        lrow = oNamedRange.EndRow
        daVoce = oNamedRange.EndRow + 2
        #  recupero l'ultimo numero di pagina usato (serve in caso di libretto unico)
        #  oSheetCont = oDoc.Sheets.getByName('CONTABILITA')
        # old_nPage = int(oSheetCont.getCellByPosition(20, lrow).Value)
        daVoce = int(oSheetCont.getCellByPosition(0, daVoce).Value)
        if daVoce == 0:
            DLG.MsgBox('Non ci sono voci di misurazione da registrare.', 'ATTENZIONE!')
            return
        oCell = oSheetCont.getCellRangeByPosition(0, frow, 25, lrow)
        #  'Raggruppa_righe
        oCell.Rows.IsVisible = False
    else:
        daVoce = 1
    #############
    # PRIMA RIGA
    daVoce = InputBox(str(daVoce), "Registra voci Libretto da n.")
    try:
        lrow = int(SheetUtils.uFindStringCol(daVoce, 0, oSheet))
    except TypeError:
        return
    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    primariga = sStRange.RangeAddress.StartRow
    #############
    #  ULTIMA RIGA
    oCellRange = oSheetCont.getCellRangeByPosition(
        0, 3, 0,
        SheetUtils.getUsedArea(oSheetCont).EndRow - 2)
    aVoce = int(oCellRange.computeFunction(MAX))
    aVoce = InputBox(str(aVoce), "A voce n.:")
    try:
        lrow = int(SheetUtils.uFindStringCol(aVoce, 0, oSheet))
    except TypeError:
        return
    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    # ultimariga = sStRange.RangeAddress.EndRow

    lrowDown = SheetUtils.uFindStringCol("T O T A L E", 2, oSheetCont)
    oCell = oSheetCont.getCellRangeByPosition(19, primariga, 25, lrowDown)
    oDoc.CurrentController.select(oCell)
    rimuovi_area_di_stampa()
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect
    oSheetCont.removeAllManualPageBreaks()
    visualizza_PageBreak()

    #  nomearea="#Lib#"+nsal
    #  rem annota importo parziale SAL


#  'Print ultimariga & " - " & "SAL n." & nSal
#  oSheetCont.getCellByPosition(25, ultimariga-1).string = "SAL n." & nSal
#  oSheetCont.getCellByPosition(25, ultimariga).formula = "=SUBTOTAL(9;P" & primariga+1 & ":P" & ultimariga+1 & ")"
#  oSheetCont.getCellByPosition(25, ultimariga).cellstyle = "comp sotto Euro 3_R"
#  rem immetti le firme
#  inizioFirme = ultimariga+1
#  firme (ultimariga+1)' riga di inserimento
#  fineFirme = ultimariga+10 rem 10 è un numero deciso nella sub FIRME
#  '    Print finefirme
#  rem definisci il range del #Lib#
#  area="$A$" & primariga+1 & ":$AJ$"&fineFirme+1
#  'Print area
#  ScriptPy("pyleeno.py","NominaArea", "CONTABILITA", area , nomearea)

#  oSheetCont.getCellRangeByPosition (0,inizioFirme,32,finefirme).CellStyle = "Ultimus_centro_bordi_lati"
#  oNamedRange=oRanges.getByName("#Lib#" & nSal).referredCells
#  '    ThisComponent.CurrentController.Select(oNamedRange)
#  rem ----------------------------------------------------------------------
#  With oNamedRange.RangeAddress
#  daRiga = .StartRow
#  aRiga = .EndRow
#  daColonna = .StartColumn
#  aColonna = .EndColumn
#  End With
#  rem set area di stampa
#  Dim selLib(0) as new com.sun.star.table.CellRangeAddress
#  selLib(0).StartColumn = daColonna
#  selLib(0).StartRow = daRiga
#  selLib(0).EndColumn = 11
#  selLib(0).EndRow = aRiga
#  rem set intestazione area di stampa
#  oTitles = createUnoStruct("com.sun.star.table.CellRangeAddress")
#  oTitles.startRow = 2' headstart - 1
#  oTitles.EndRow = 2 'headend - 1
#  oTitles.startColumn = 0
#  oTitles.EndColumn = 11
#  oSheetCont.setTitleRows(oTitles)
#  oSheetCont.setPrintareas(selLib())
#  oSheetCont.setPrintTitleRows(true)
#  rem ----------------------------------------------------------------------
#  rem sbianco i dati e l'intestazione
#  ThisComponent.CurrentController.Select(oSheetCont.getCellRangeByPosition(0, daRiga, 11, fineFirme))
#  ScriptPy("pyleeno.py","adatta_altezza_riga")
#  Sbianca_celle
#  ThisComponent.CurrentController.Select(oSheetCont.getCellRangeByPosition(0, 2, 11, 2))
#  Sbianca_celle
#  ThisComponent.currentController.removeRangeSelectionListener(oRangeSelectionListener) 'deseleziona
#  rem ----------------------------------------------------------------------
#  ThisComponent.CurrentController.setFirstVisibleRow (fineFirme-3) 'solo debug
#  i=1
#  Do While oSheetCont.getCellByPosition(1,fineFirme).rows.IsStartOfNewPage = False
#  '    oSheetCont.getCellByPosition(2 ,fineFirme).setstring("Sto sistemando il Libretto...")
#  insRows (fineFirme,1) 'insertByIndex non funziona
#  If i=3 Then
#  oSheetCont.getCellByPosition(2, fineFirme).setstring("====================")
#  daqui=fineFirme
#  End If
#  fineFirme = fineFirme+1
#  i=i+1
#  Loop
#  oSheetCont.rows.removeByIndex (fineFirme-1, 1)

#  rem ----------------------------------------------------------------------
#  rem cancella l'ultima riga
#  fineFirme = fineFirme-1
#  If daqui<>0 then
#  ThisComponent.CurrentController.Select(oSheetCont.getCellByPosition(2, daqui))
#  copy_clip
#  ThisComponent.CurrentController.Select(oSheetCont.getCellRangeByPosition(2, daqui, 2, finefirme-1))
#  'Print finefirme-1
#  paste_clip
#  End If

#  rem ----------------------------------------------------------------------
#  rem definisci il range del #Lib#
#  area="$A$" & primariga+1 & ":$AJ$"&fineFirme+1
#  ScriptPy("pyleeno.py","NominaArea", "CONTABILITA", area , nomearea)

#  rem raggruppo
#  oCell = oSheetCont.getCellRangeByPosition(0,primariga,25,finefirme)
#  oSheetCont.getCellRangeByPosition (0,inizioFirme,25,finefirme).CellStyle = "Ultimus_centro_bordi_lati"
#  oCell = oSheetCont.getCellRangeByPosition (0,finefirme,35,finefirme)
#  oSheetCont.getCellByPosition(2 , inizioFirme+1).CellStyle = "ULTIMUS" 'stile data
#  rem recupero la data
#  datafirma = Right (oSheetCont.getCellByPosition(2 , inizioFirme+1).value, 10)
#  ThisComponent.CurrentController.Select(oCell)
#  bordo_sotto
#  ThisComponent.currentController.removeRangeSelectionListener(oRangeSelectionListener)
#  rem ----------------------------------------------------------------------
#  rem QUESTA DEVE DIVENTARE UN'OPZIONE A SCELTA DELL'UTENTE
#  rem in caso di libretto unico questo If è da attivare
#  rem in modo che la numerazione delle pagine non ricominci da capo
#  '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#  '        If nSal > 1 Then
#  '            nLib = 1
#  '            inumPag = 1 + old_nPage 'SE IL LIBRETTO è UNICO
#  '        End If
#  '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#  rem ----------------------------------------------------------------------
#  nLib = nSal
#  'End If
#  '#########################################################################
#  rem COMPILO LA SITUAZIONE CONTABILE IN "S2" 1di2
#  oS2 = ThisComponent.Sheets.getByName("S2")
#  rem TROVO LA POSIZIONE DEL TITOLO
#  oEnd=SheetUtils.uFindString("SITUAZIONE CONTABILE", oS2)
#  xS2=oEnd.RangeAddress.EndRow        'riga
#  yS2=oEnd.RangeAddress.EndColumn    'colonna

#  oS2.getCellByPosition(yS2+nSal,xS2+1).value = nSal 'numero sal
#  '    Print "datafirma " & datafirma
#  oS2.getCellByPosition(yS2+nSal,xS2+2).value = datafirma 'data
#  oS2.getCellByPosition(yS2+nSal,xS2+24).value = aVoce ' ultima voce libretto
#  oS2.getCellByPosition(yS2+nSal,xS2+25).value = inumPag ' ultima pagina libretto
#  ThisComponent.currentController.removeRangeSelectionListener(oRangeSelectionListener)
#  '#########################################################################

#  'BARRA_chiudi
#  Barra_Apri_Chiudi_5("                         Sto elaborando il Libetto delle Misure...", 75)

#  Dim IntPag()
#  IntPag() = oSheetCont.RowPageBreaks
#  'GoTo togo:
#  rem ----------------------------------------------------------------------
#  rem col seguente ciclo FOR inserisco i dati ma non il numero di pagina
#  For i = primariga to fineFirme
#  IF     oSheetCont.getCellByPosition( 1 , i).cellstyle = "comp Art-EP_R" then
#  if primariga=0 then
#  sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, i)
#  With sStRange.RangeAddress
#  primariga =.StartRow
#  End With
#  end If
#  oSheetCont.getCellByPosition(19, i).value= nLib '1 ' numero libretto
#  oSheetCont.getCellByPosition(22, i).string= "#reg" ' flag registrato
#  oSheetCont.getCellByPosition(23, i).value= nSal ' numero SAL
#  oSheetCont.getCellByPosition(23, i).cellstyle = "Sal"
#  '        else
#  '        oSheetCont.getCellByPosition( 1 , i).Rows.Height = 500 'altezza riga
#  end if
#  Next
#  inumPag = 0'+ old_nPage 'SE IL LIBRETTO è UNICO
#  rem ----------------------------------------------------------------------
#  rem il ciclo For che segue è preso da https://forum.openoffice.org/it/forum/viewtopic.php?f=26&t=6014
#  '    Dim IntPag()
#  '    IntPag() = oSheetCont.RowPageBreaks
#  rem ----------------------------------------------------------------------
#  togo:
#  rem ----------------------------------------------------------------------
#  rem col seguente ciclo FOR inserisco solo il numero di pagina
#  rem inserendo qui anche il resto dei dati ho numeri di pagina un po' ballerini
#  For i = primariga to fineFirme
#  For n = LBound(IntPag) To UBound(IntPag)
#  if i < IntPag(n).Position Then
#  if oSheetCont.getCellByPosition( 1 , i).cellstyle = "comp Art-EP_R" then
#  inumPag = n ' + old_nPage 'SE IL LIBRETTO è UNICO
#  oSheetCont.getCellByPosition(20, i).value = inumPag 'numero Pagina
#  '    oSheetCont.getCellByPosition(19, i).value= nLib '1 ' numero libretto
#  '    oSheetCont.getCellByPosition(22, i).string= "#reg" ' flag registrato
#  '    oSheetCont.getCellByPosition(23, i).value= nSal ' numero SAL
#  '    oSheetCont.getCellByPosition(23, i).cellstyle = "Sal"
#  Exit For
#  end If
#  end if
#  Next n
#  Next i
#  rem ----------------------------------------------------------------------
#  rem annoto ultimo numero di pagina
#  oSheetCont.getCellByPosition(20 , fineFirme).value = UBound(IntPag)'inumPag
#  oSheetCont.getCellByPosition(20 , fineFirme).CellStyle = "num centro"
#  rem ----------------------------------------------------------------------
#  'fissa (0,idxrow+1)

#  ThisComponent.currentController.removeRangeSelectionListener(oRangeSelectionListener)
#  rem ----------------------------------------------------------------------
#  rem inserisco la prima riga GIALLA del LIBRETTO
#  oNamedRange=oRanges.getByName(nomearea).referredCells
#  ins = oNamedRange.RangeAddress.StartRow
#  insRows (ins, 1) 'insertByIndex non funziona
#  oSheetCont.getCellRangeByPosition (0,ins,25,ins).CellStyle = "uuuuu" '"Ultimus_Bordo_sotto"
#  fissa (0, ins + 1)
#  rem ----------------------------------------------------------------------
#  rem ci metto un po' di informazioni
#  oSheetCont.getCellByPosition(2,ins).string = "segue Libretto delle Misure n." & nSal & " - " & davoce & "÷" & avoce
#  oSheetCont.getCellByPosition(20,ins).value =  UBound(IntPag) 'ultimo numero pagina
#  oSheetCont.getCellByPosition(19, ins).value= nLib '1 ' numero libretto
#  oSheetCont.getCellByPosition(23, ins).value= nSal ' numero SAL
#  oSheetCont.getCellByPosition(25, ins).formula = "=SUBTOTAL(9;$P$" & primariga+1 & ":$P$" & ultimariga+2 & ")"
#  oSheetCont.getCellByPosition(25, ins).cellstyle = "comp sotto Euro 3_R"
#  rem ----------------------------------------------------------------------
#  rem annoto il sal corrente sulla riga di intestazione
#  ins =SheetUtils.uFindString("LAVORAZIONI"+ chr(10) + "O PROVVISTE", oSheetCont).RangeAddress.EndRow
#  oSheetCont.getCellByPosition(25,ins).value = nSal
#  oSheetCont.getCellByPosition(25, ins).cellstyle = "Menu_sfondo _input_grasBig"
#  '    oSheetCont.getCellByPosition(25, ins-1).formula = "=SUBTOTAL(9;$P$" & primariga+1 & ":$P$" & ultimariga+2 & ")"
#  oSheetCont.getCellByPosition(25, ins-1).formula = "=$P$2-SUBTOTAL(9;$P$" & IdxRow & ":$P$" & ultimariga+2 & ")"
#  rem ----------------------------------------------------------------------
#  rem fisso la riga alla prima voce
#  For i= 2 To 100
#  If oSheetCont.getCellByPosition(1,i).CellStyle = "Comp-Bianche sopra_R" Then
#  Exit For
#  EndIf

#  Next
#  'fissa (0,idxRow)
#  ThisComponent.CurrentController().freezeAtPosition(0,idxRow+1)
#  Ripristina_statusLine 'Barra_chiudi_sempre_4
#  Protezione_area ("CONTABILITA",nomearea)
#  Struttura_Contab ("#Lib#")
#  RiDefinisci_Area_Elenco_prezzi ' non capisco come mai l'area di elenco_prezzi viene cambiata
#  Genera_REGISTRO
#  end Sub


########################################################################
def genera_atti_contabili():
    '''
    @@ DA DOCUMENTARE
    '''
    # ~def debug():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name != "CONTABILITA":
        return
    if DLG.DlgSiNo(
            '''Prima di procedere è consigliabile salvare il lavoro.
Puoi continuare, ma a tuo rischio!
Se decidi di continuare, devi attendere il messaggio di procedura completata senza interferire con mouse e/o tastiera.
Procedo senza salvare?''', 'Avviso') == 3:
        return
    #  genera_libretto()
    DLG.MsgBox(
        '''La generazione degli allegati contabili è stata completata.
Grazie per l'attesa.''', 'Voci registrate!')


# FINE_CONTABILITA ## FINE_CONTABILITA ## FINE_CONTABILITA ## FINE_CONTABILITA
########################################################################
def inizializza_elenco():
    '''
    Riscrive le intestazioni di colonna e le formule dei totali in Elenco Prezzi.
    '''
    chiudi_dialoghi()
    DisableAutoCalc()

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.Sheets.getByName('Elenco Prezzi')

    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400
    #  oDialogo_attesa = dlg_attesa()
    #  attesa().start() #mostra il dialogo

    struttura_off()
    oCellRangeAddr = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    # SR = oCellRangeAddr.StartRow
    ER = oCellRangeAddr.EndRow
    # SC = oCellRangeAddr.StartColumn
    EC = oCellRangeAddr.EndColumn
    oSheet.getCellRangeByPosition(11, 3, EC, ER -
                                  1).clearContents(STRING + VALUE + FORMULA)

    oDoc.CurrentController.freezeAtPosition(0, 3)
    oSheet.getCellRangeByPosition(0, 0, 100, 0).CellStyle = "Default"
    #  riscrivo le intestazioni di colonna
    set_larghezza_colonne()
    oSheet.getCellRangeByName('L1').String = 'COMPUTO'
    oSheet.getCellRangeByName('P1').String = 'VARIANTE'
    oSheet.getCellRangeByName('T1').String = 'CONTABILITA'
    oSheet.getCellRangeByName('B2').String = 'QUESTA RIGA NON VIENE STAMPATA'
    oSheet.getCellRangeByName(
        "'Elenco Prezzi'.A2:AA2").CellStyle = "comp In testa"
    oSheet.getCellRangeByName("'Elenco Prezzi'.AA2").CellStyle = 'EP-mezzo %'

    oSheet.getCellRangeByName("'Elenco Prezzi'.A3:AA3").CellStyle = "EP-a -Top"
    oSheet.getCellRangeByName('A3').String = 'Codice\nArticolo'
    oSheet.getCellRangeByName(
        'B3').String = 'DESCRIZIONE DEI LAVORI\nE DELLE SOMMINISTRAZIONI'
    oSheet.getCellRangeByName('C3').String = 'Unità\ndi misura'
    oSheet.getCellRangeByName('D3').String = 'Sicurezza\ninclusa'
    oSheet.getCellRangeByName('E3').String = 'Prezzo\nunitario'
    oSheet.getCellRangeByName('F3').String = 'Incidenza\nMdO'
    oSheet.getCellRangeByName('G3').String = 'Importo\nMdO'
    oSheet.getCellRangeByName('H3').String = 'Codice di origine'
    oSheet.getCellRangeByName('L3').String = 'Inc. % \nComputo'
    oSheet.getCellRangeByName('M3').String = 'Quantità\nComputo'
    oSheet.getCellRangeByName('N3').String = 'Importi\nComputo'
    oSheet.getCellRangeByName('L3:N3').CellBackColor = 16762855
    oSheet.getCellRangeByName('P3').String = 'Inc. % \nVariante'
    oSheet.getCellRangeByName('Q3').String = 'Quantità\nVariante'
    oSheet.getCellRangeByName('R3').String = 'Importi\nVariante'
    oSheet.getCellRangeByName('P3:R3').CellBackColor = 16777062
    oSheet.getCellRangeByName('T3').String = 'Inc. % \nContabilità'
    oSheet.getCellRangeByName('U3').String = 'Quantità\nContabilità'
    oSheet.getCellRangeByName('V3').String = 'Importi\nContabilità'
    oSheet.getCellRangeByName('T3:V3').CellBackColor = 16757935
    oSheet.getCellRangeByName('X3').String = 'Quantità\nvariaz.'
    oSheet.getCellRangeByName('Y3').String = 'IMPORTI\nin più'
    oSheet.getCellRangeByName('Z3').String = 'IMPORTI\nin meno'
    oSheet.getCellRangeByName('AA3').String = 'VAR. %'
    oSheet.getCellRangeByName('I1:J1').Columns.IsVisible = False

    y = SheetUtils.uFindStringCol('Fine elenco', 0, oSheet) + 1
    oSheet.getCellRangeByName('N2').Formula = '=SUBTOTAL(9;N3:N' + str(y) + ')'
    oSheet.getCellRangeByName('R2').Formula = '=SUBTOTAL(9;R3:R' + str(y) + ')'
    oSheet.getCellRangeByName('V2').Formula = '=SUBTOTAL(9;V3:V' + str(y) + ')'
    oSheet.getCellRangeByName('Y2').Formula = '=SUBTOTAL(9;Y3:Y' + str(y) + ')'
    oSheet.getCellRangeByName('Z2').Formula = '=SUBTOTAL(9;Z3:Z' + str(y) + ')'
    #   riga di totale importo COMPUTO
    y -= 1
    oSheet.getCellByPosition(12, y).String = 'TOTALE'
    oSheet.getCellByPosition(13, y).Formula = '=SUBTOTAL(9;N3:N' + str(y) + ')'
    #  riga di totale importo CONTABILITA'
    oSheet.getCellByPosition(16, y).String = 'TOTALE'
    oSheet.getCellByPosition(17, y).Formula = '=SUBTOTAL(9;R3:R' + str(y) + ')'
    #  rem	riga di totale importo VARIANTE
    oSheet.getCellByPosition(20, y).String = 'TOTALE'
    oSheet.getCellByPosition(21, y).Formula = '=SUBTOTAL(9;V3:V' + str(y) + ')'
    #  rem	riga di totale importo PARALLELO
    oSheet.getCellByPosition(23, y).String = 'TOTALE'
    oSheet.getCellByPosition(24, y).Formula = '=SUBTOTAL(9;Y3:Y' + str(y) + ')'
    oSheet.getCellByPosition(25, y).Formula = '=SUBTOTAL(9;Z3:Z' + str(y) + ')'
    oSheet.getCellRangeByPosition(10, y, 26,
                                  y).CellStyle = 'EP statistiche_Contab'

    y += 1
    #  oSheet.getCellRangeByName('C2').String = 'prezzi'
    #  oSheet.getCellRangeByName('E2').Formula = '=COUNT(E3:E' + str(y) +')'
    oSheet.getCellRangeByName('K2:K' + str(y)).CellStyle = 'Default'
    oSheet.getCellRangeByName('O2:O' + str(y)).CellStyle = 'Default'
    oSheet.getCellRangeByName('S2:S' + str(y)).CellStyle = 'Default'
    oSheet.getCellRangeByName('W2:W' + str(y)).CellStyle = 'Default'
    oSheet.getCellRangeByPosition(3, 3, 250, y + 10).clearContents(HARDATTR)
    #  riga_rossa()
    #  oDialogo_attesa.endExecute()
    oDoc.CurrentController.ZoomValue = zoom

    EnableAutoCalc()
    #  MsgBox('Rigenerazione del foglio eseguita!','')


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
        17, 1).Formula = '=SUBTOTAL(9;R3:R' + str(lRowE + 1) + ')'  # sicurezza
    oSheet.getCellByPosition(
        18,
        1).Formula = '=SUBTOTAL(9;S3:S' + str(lRowE + 1) + ')'  # importo lavori
    oSheet.getCellByPosition(0, 1).Formula = '=AK2'

    oSheet.getCellByPosition(
        28,
        1).Formula = '=SUBTOTAL(9;AC3:AC' + str(lRowE +
                                                1) + ')'  # importo materiali

    oSheet.getCellByPosition(29,
                             1).Formula = '=AE2/S2'  # Incidenza manodopera %
    oSheet.getCellByPosition(29, 1).CellStyle = "Comp TOTALI %"
    oSheet.getCellByPosition(
        30,
        1).Formula = '=SUBTOTAL(9;AE3:AE' + str(lRowE +
                                                1) + ')'  # importo manodopera
    oSheet.getCellByPosition(36, 1).Formula = '=SUBTOTAL(9;AK3:AK' + str(
        lRowE + 1) + ')'  # totale computo sole voci senza errori

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
    oSheet.getCellByPosition(36, 2).String = 'importo totale\nsenza errori'
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
        lRowE).Formula = "=SUBTOTAL(9;R3:R" + str(lRowE +
                                                  1) + ")"  # importo sicurezza
    oSheet.getCellByPosition(
        18, lRowE).Formula = "=SUBTOTAL(9;S3:S" + str(lRowE +
                                                      1) + ")"  # importo lavori
    oSheet.getCellByPosition(29, lRowE).Formula = "=AE" + str(
        lRowE + 1) + "/S" + str(lRowE + 1) + ""  # Incidenza manodopera %
    oSheet.getCellByPosition(30, lRowE).Formula = "=SUBTOTAL(9;AE3:AE" + str(
        lRowE + 1) + ")"  # importo manodopera
    oSheet.getCellByPosition(36, lRowE).Formula = "=SUBTOTAL(9;AK3:AK" + str(
        lRowE + 1) + ")"  # totale computo sole voci senza errori
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
    set_larghezza_colonne()
    setTabColor(16762855)


########################################################################
def inizializza_analisi():
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoAnalysis.inizializzaAnalisi'
    Se non presente, crea il foglio 'Analisi di Prezzo' ed inserisce la prima scheda
    '''
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
        set_larghezza_colonne()
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
    basic_LeenO("Menu.eventi_assegna")
    LeenoSheetUtils.inserisciRigaRossa(oSheet)
    ScriviNomeDocumentoPrincipale()


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

########################################################################
def struct_colore(level):
    '''
    Mette in vista struttura secondo il colore
    level { integer } : specifica il livello di categoria
    '''
    oDoc = LeenoUtils.getDocument()
    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400
    oSheet = oDoc.CurrentController.ActiveSheet
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    hriga = oSheet.getCellRangeByName('B4').CharHeight * 65
    #  giallo(16777072,16777120,16777168)
    #  verde(9502608,13696976,15794160)
    #  viola(12632319,13684991,15790335)
    col0 = 16724787  # riga_rossa
    col1 = 16777072
    col2 = 16777120
    col3 = 16777168
    # attribuisce i colori
    for y in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
        if oSheet.getCellByPosition(0, y).String == '':
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

        for n in (3, 7):
            oCellRangeAddr.StartColumn = n
            oCellRangeAddr.EndColumn = n
            oSheet.group(oCellRangeAddr, 0)
            oSheet.getCellRangeByPosition(n, 0, n, 0).Columns.IsVisible = False
    test = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2
    lista = list()
    for n in range(0, test):
        if oSheet.getCellByPosition(0, n).CellBackColor == colore:
            oSheet.getCellByPosition(0, n).Rows.Height = hriga
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
        oSheet.getCellRangeByPosition(0, el[0], 0,
                                      el[1]).Rows.IsVisible = False
    oDoc.CurrentController.ZoomValue = zoom
    return


########################################################################
def struttura_Elenco():
    '''
    Dà una tonalità di colore, diverso dal colore dello stile cella, alle righe
    che non hanno il prezzo, come i titoli di capitolo e sottocapitolo.
    '''
    chiudi_dialoghi()
    if DLG.DlgSiNo(
            '''Adesso puoi dare ai titoli di capitolo e sottocapitolo
una tonalità di colore che ne facilita la leggibilità, ma
il risultato finale dipende dalla struttura dei codici di voce.

QUESTA OPERAZIONE RICHIEDE DEL TEMPO E
LibreOffice POTREBBE SEMBRARE BLOCCATO!

Vuoi procedere con la creazione della struttura dei capitoli?''',
            'Avviso') == 3:
        return
    # ~ riordina_ElencoPrezzi()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    struct_colore(0)  # attribuisce i colori
    struct_colore(1)
    struct_colore(2)
    return


########################################################################
def ns_ins(filename=None):
    '''
    Se assente, inserisce il namespace nel file XML.
    '''
    f = codecs.open(filename, 'r', 'utf-8')
    out_file = '.'.join(filename.split('.')[:-1]) + '.bak'
    of = codecs.open(out_file, 'w', 'utf-8')

    for row in f:
        nrow = row.replace(
            '<PRT:Prezzario>',
            '<PRT:Prezzario xmlns="http://www.regione.toscana.it/Prezzario" xmlns:PRT="http://www.regione.toscana.it/Prezzario/Prezzario.xsd">'
        )
        of.write(nrow)
    f.close()
    of.close()
    shutil.move(out_file, filename)

########################################################################
# XML_toscana_import moved to LeenoImport.py
########################################################################

########################################################################
# MENU_fuf moved to LeenoImport.py
########################################################################

########################################################################
# XML_import_ep moved to LeenoImport.py
########################################################################

########################################################################
# XML_import_multi moved to LeenoImport.py
########################################################################

########################################################################
# importa_listino_leeno moved to LeenoImport.py
########################################################################


def colora_vecchio_elenco():
    '''
    @@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400
    #  giallo(16777072,16777120,16777168)
    #  verde(9502608,13696976,15794160)
    #  viola(12632319,13684991,15790335)
    col1 = 16777072
    col2 = 16777120
    col3 = 16777168
    inizio = SheetUtils.uFindStringCol('COMPLETO', 4, oSheet) + 1
    fine = SheetUtils.getUsedArea(oSheet).EndRow + 1
    for el in range(inizio, fine):
        if len(oSheet.getCellByPosition(2, el).String.split('.')) == 1:
            oSheet.getCellByPosition(2, el).CellBackColor = col1
        if len(oSheet.getCellByPosition(2, el).String.split('.')) == 2:
            oSheet.getCellByPosition(2, el).CellBackColor = col2
        if len(oSheet.getCellByPosition(2, el).String.split('.')) == 3:
            oSheet.getCellByPosition(2, el).CellBackColor = col3
    oDoc.CurrentController.ZoomValue = zoom


########################################################################
def importa_stili(ctx):
    '''
    Importa tutti gli stili da un documento di riferimento. Se non è
    selezionato, il file di riferimento è il template di leenO.
    '''
    if DLG.DlgSiNo(
            '''Questa operazione sovrascriverà gli stili
del documento attivo, se già presenti!

Se non scegli un file di riferimento, saranno
importati gli stili di default di LeenO.

Vuoi continuare?''', 'Importa Stili in blocco?') == 3:
        return
    filename = Dialogs.FileSelect('Scegli il file di riferimento...', '*.ods')
    if filename is None:
        #  desktop = LeenoUtils.getDesktop()
        filename = LeenO_path() + '/template/leeno/Computo_LeenO.ots'
    else:
        filename = uno.systemPathToFileUrl(filename)
    oDoc = LeenoUtils.getDocument()
    oDoc.getStyleFamilies().loadStylesFromURL(filename, list())
    for el in oDoc.Sheets.ElementNames:
        oDoc.CurrentController.setActiveSheet(oDoc.getSheets().getByName(el))
        adatta_altezza_riga(el)
    try:
        GotoSheet('Elenco Prezzi')
    except Exception:
        pass


########################################################################
def MENU_parziale():
    '''
    Inserisce una riga con l'indicazione della somma parziale.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]
    if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        parziale_core(lrow)
        parziale_verifica()


###
def parziale_core(lrow):
    '''
    lrow    { double } : id della riga di inserimento
    '''

    if lrow == 0:
        return
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    sopra = sStRange.RangeAddress.StartRow
    # sotto = sStRange.RangeAddress.EndRow
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        if(oSheet.getCellByPosition(0, lrow).CellStyle == 'comp 10 s' and
           oSheet.getCellByPosition(1, lrow).CellStyle == 'Comp-Bianche in mezzo' and
           oSheet.getCellByPosition(2, lrow).CellStyle == 'comp 1-a' or
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

        if(oSheet.getCellByPosition(0, lrow).CellStyle == "comp 10 s_R" and
           oSheet.getCellByPosition(1, lrow).CellStyle == "Comp-Bianche in mezzo_R" and
           oSheet.getCellByPosition(2, lrow).CellStyle == "comp 1-a" or
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


###
def parziale_verifica():
    '''
    Controlla l'esattezza del calcolo del parziale quanto le righe di
    misura vengono aggiunte o cancellate.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]
    #  if oSheet.Name in('COMPUTO','VARIANTE', 'CONTABILITA'):
    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    sopra = sStRange.RangeAddress.StartRow + 2
    sotto = sStRange.RangeAddress.EndRow
    for n in range(sopra, sotto):
        if 'Parziale [' in (oSheet.getCellByPosition(8, n).String):
            parziale_core(n)


########################################################################
def vedi_voce_xpwe(lrow, vRif, flags=''):
    """(riga d'inserimento, riga di riferimento)"""
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, vRif)
    # sStRange.RangeAddress
    idv = sStRange.RangeAddress.StartRow + 1
    sotto = sStRange.RangeAddress.EndRow
    art = 'B$' + str(idv + 1)
    idvoce = 'A$' + str(idv + 1)
    des = 'C$' + str(idv + 1)
    quantity = 'J$' + str(sotto + 1)
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
    oSheet.getCellByPosition(4, lrow).Formula = '=' + quantity
    oSheet.getCellByPosition(
        9, lrow).Formula = '=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(
            lrow + 1) + ')=0;"";PRODUCT(E' + str(lrow +
                                                 1) + ':I' + str(lrow +
                                                                 1) + '))'
    if flags in ('32769', '32801'):  # 32768
        inverti_segno()
        oSheet.getCellRangeByPosition(2, lrow, 10, lrow).CharColor = 16724787


########################################################################
def MENU_vedi_voce():
    '''
    Inserisce un riferimento a voce precedente sulla riga corrente.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]
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
            to = '$' + oSheet.Name + '.$C$' + str(SheetUtils.uFindStringCol(
                to, 0, oSheet))
        try:
            to = int(to.split('$')[-1]) - 1
        except ValueError:
            return
        _gotoCella(2, lrow)
        # focus = oDoc.CurrentController.getFirstVisibleRow
        if to < lrow:
            vedi_voce_xpwe(
                lrow,
                to,
            )


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
def MENU_converti_stringhe():
    '''
    Converte in numeri le stinghe o viceversa.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    # ctx = LeenoUtils.getComponentContext()
    # desktop = LeenoUtils.getDesktop()
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
    for y in range(sCol, eCol + 1):
        for x in range(sRow, eRow + 1):
            try:
                if oSheet.getCellByPosition(y, x).Type.value == 'TEXT':
                    oSheet.getCellByPosition(y, x).Value = float(
                        oSheet.getCellByPosition(y,
                                                 x).String.replace(',', '.'))
                else:
                    oSheet.getCellByPosition(
                        y, x).String = oSheet.getCellByPosition(y, x).String
            except Exception:
                pass
    return


########################################################################
def ssUltimus():
    '''
    Scrive la variabile globale che individua il Documento Principale (DCC)
    che è il file a cui giungono le voci di prezzo inviate da altri file
    '''
    oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('M1'):
        return
    try:
        LeenoUtils.getGlobalVar('oDlgMain').endExecute()
    except NameError:
        pass
    if len(oDoc.getURL()) == 0:
        DLG.MsgBox(
            '''Prima di procedere, devi salvare il lavoro!
Provvedi subito a dare un nome al file di computo...''',
            'Dai un nome al file...')
        salva_come()
        autoexec()
    try:
        LeenoUtils.setGlobalVar('sUltimus', uno.fileUrlToSystemPath(oDoc.getURL()))
    except Exception:
        pass
    DlgMain()
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
    filtra_codice()


def filtra_codice(voce=None):
    '''
    Applica un filtro di visualizzazione sulla base del codice di voce selezionata.
    Lanciando il comando da Elenco Prezzi, il comportamento è regolato dal valore presente nella cella 'C2'
    '''
    DisableAutoCalc()

    oDoc = LeenoUtils.getDocument()
    # ~zoom = oDoc.CurrentController.ZoomValue
    # ~oDoc.CurrentController.ZoomValue = 400
    oSheet = oDoc.CurrentController.ActiveSheet

    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_contab = LeenoUtils.getGlobalVar('stili_contab')

    if oSheet.Name == "Elenco Prezzi":
        oCell = oSheet.getCellRangeByName('C2')
        voce = oDoc.Sheets.getByName('Elenco Prezzi').getCellByPosition(
            0,
            LeggiPosizioneCorrente()[1]).String
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
        next_voice(LeggiPosizioneCorrente()[1], 1)
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
        DLG.MsgBox('Devi prima selezionare una voce di misurazione.', 'Avviso!')
        return
    fine = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1
    lista_pt = list()
    _gotoCella(0, 0)
    for n in range(0, fine):
        if oSheet.getCellByPosition(0,
                                    n).CellStyle in ('Comp Start Attributo',
                                                     'Comp Start Attributo_R'):
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, n)
            sopra = sStRange.RangeAddress.StartRow
            sotto = sStRange.RangeAddress.EndRow
            if oSheet.getCellByPosition(1, sopra + 1).String != voce:
                lista_pt.append((sopra, sotto))
                #  lista_pt.append((sopra+2, sotto-1))
    for el in lista_pt:
        oCellRangeAddr.StartRow = el[0]
        oCellRangeAddr.EndRow = el[1]
        oSheet.group(oCellRangeAddr, 1)
        oSheet.getCellRangeByPosition(0, el[0], 0,
                                      el[1]).Rows.IsVisible = False
    _gotoCella(0, lrow)

    EnableAutoCalc()
    # ~oDoc.CurrentController.ZoomValue = zoom
    # ~MsgBox('Filtro attivato in base al codice!','Codice voce: ' + voce)


def struttura_ComputoM():
    '''
    @@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    Rinumera_TUTTI_Capitoli2()
    struct(0)
    struct(1)
    struct(2)
    struct(3)


def struttura_Analisi():
    '''
    @@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    struct(4)


def MENU_struttura_off():
    '''
    Cancella la vista in struttura
    '''
    struttura_off()


def struttura_off():
    '''
    Cancella la vista in struttura
    '''
    oDoc = LeenoUtils.getDocument()
    lrow = LeggiPosizioneCorrente()[1]
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    oDoc.CurrentController.setFirstVisibleColumn(0)
    oDoc.CurrentController.setFirstVisibleRow(lrow - 4)


def struct(level):
    ''' mette in vista struttura secondo categorie
    level { integer } : specifica il livello di categoria
    ### COMPUTO/VARIANTE ###
    0 = super-categoria
    1 = categoria
    2 = sotto-categoria
    3 = intera voce di misurazione
    ### ANALISI ###
    4 = simile all'elenco prezzi
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet

    if level == 0:
        stile = 'Livello-0-scritta'
        myrange = (
            'Livello-0-scritta',
            'Comp TOTALI',
        )
        Dsopra = 1
        Dsotto = 1
    elif level == 1:
        stile = 'Livello-1-scritta'
        myrange = (
            'Livello-1-scritta',
            'Livello-0-scritta',
            'Comp TOTALI',
        )
        Dsopra = 1
        Dsotto = 1
    elif level == 2:
        stile = 'livello2 valuta'
        myrange = (
            'livello2 valuta',
            'Livello-1-scritta',
            'Livello-0-scritta',
            'Comp TOTALI',
        )
        Dsopra = 1
        Dsotto = 1
    elif level == 3:
        stile = 'Comp Start Attributo'
        myrange = (
            'Comp End Attributo',
            'Comp TOTALI',
        )
        Dsopra = 2
        Dsotto = 1

    elif level == 4:  # Analisi di Prezzo
        stile = 'An-1_sigla'
        myrange = (
            'An.1v-Att Start',
            'Analisi_Sfondo',
        )
        Dsopra = 1
        Dsotto = -1
        # ~for n in(3, 5, 7):
        # ~oCellRangeAddr.StartColumn = n
        # ~oCellRangeAddr.EndColumn = n
        # ~oSheet.group(oCellRangeAddr,0)
        # ~oSheet.getCellRangeByPosition(n, 0, n, 0).Columns.IsVisible=False

    test = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2
    lista_cat = list()
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
    Toolbars.Switch(True)
    oDoc = LeenoUtils.getDocument()
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
            # ~datarif = datetime.now()
            minuti = 60 * int(cfg.read('Generale', 'pausa_backup'))
            time.sleep(minuti)
            bak()
            # ~MsgBox('eseguita in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!','')


def autorun():
    '''
    @@ DA DOCUMENTARE
    '''
    #  global utsave
    utsave = trun()
    utsave._stop()
    utsave.start()


########################################################################
def autoexec():
    '''
    questa è richiamata da New_File()
    '''
    inizializza()

    # rinvia a autoexec in basic
    basic_LeenO('_variabili.autoexec')
    bak0()
    autorun()
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
    oDoc.enableAutomaticCalculation(True)
    oLayout = oDoc.CurrentController.getFrame().LayoutManager
    oLayout.hideElement(
        "private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV")
    #  RegularExpressions Wildcards are mutually exclusive, only one can have the value TRUE.
    #  If both are set to TRUE via API calls then the last one set takes precedence.
    try:
        oDoc.Wildcards = False
    except Exception:
        pass
    oDoc.RegularExpressions = False
    oDoc.CalcAsShown = True  # precisione come mostrato
    adegua_tmpl()  # esegue degli aggiustamenti del template
    Toolbars.Vedi()
    ScriviNomeDocumentoPrincipale()
    #  if len(oDoc.getURL()) != 0:
    # scegli cosa visualizzare all'avvio:
    #  vedi = conf.read(path_conf, 'Generale', 'visualizza')
    #  if vedi == 'Menù Principale':
    #  DlgMain()
    #  elif vedi == 'Dati Generali':
    #  vai_a_variabili()
    #  elif vedi in('Elenco Prezzi', 'COMPUTO'):
    #  GotoSheet(vedi)


#
########################################################################
def computo_terra_terra():
    '''
    Settaggio base di configurazione colonne in COMPUTO e VARIANTE
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.getCellRangeByPosition(33, 0, 1023, 0).Columns.IsVisible = False
    set_larghezza_colonne()


########################################################################
def viste_nuove(sValori):
    '''
    sValori { string } : una stringa di configurazione della visibilità colonne
    permette di visualizzare/nascondere un set di colonne
    T = visualizza
    F = nasconde
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    n = 0
    for el in sValori:
        if el == 'T':
            oSheet.getCellByPosition(n, 2).Columns.IsVisible = True
        elif el == 'F':
            oSheet.getCellByPosition(n, 2).Columns.IsVisible = False
        n += 1


########################################################################
def set_larghezza_colonne():
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoSheetUtils.setLarghezzaColonne'
    regola la larghezza delle colonne a seconda della sheet
    '''
    oDoc = LeenoUtils.getDocument()
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
        oSheet.getCellRangeByPosition(
            13, 0, 1023, 0).Columns.Width = 1900  # larghezza colonne importi
        oSheet.getCellRangeByPosition(
            19, 0, 23, 0).Columns.Width = 1000  # larghezza colonne importi
        oSheet.getCellRangeByPosition(
            51, 0, 1023, 0).Columns.IsVisible = False  # nascondi colonne
        oSheet.getColumns().getByName('A').Columns.Width = 600
        oSheet.getColumns().getByName('B').Columns.Width = 1500
        oSheet.getColumns().getByName('C').Columns.Width = 6300  # 7800
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
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        oSheet.getCellRangeByPosition(
            5, 0, 8, 0).Columns.IsVisible = True  # mostra colonne
        oSheet.getColumns().getByName('A').Columns.Width = 600
        oSheet.getColumns().getByName('B').Columns.Width = 1500
        oSheet.getColumns().getByName('C').Columns.Width = 6300  # 7800
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
def elimina_stili_cella():
    '''
    @@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()
    stili = oDoc.StyleFamilies.getByName('CellStyles').getElementNames()
    for el in stili:
        if not oDoc.StyleFamilies.getByName('CellStyles').getByName(el).isInUse():
            oDoc.StyleFamilies.getByName('CellStyles').removeByName(el)


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

# inizializza la lista di scelta per in elenco Prezzi
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
    # Indica qual è il Documento Principale
    ScriviNomeDocumentoPrincipale()
    nascondi_sheets()


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
    '''
    DisableAutoCalc()

    oDoc = LeenoUtils.getDocument()
    # LE VARIABILI NUOVE VANNO AGGIUNTE IN config_default()
    # cambiare stile http://bit.ly/2cDcCJI
    ver_tmpl = oDoc.getDocumentProperties().getUserDefinedProperties().Versione
    if ver_tmpl > 200:
        basic_LeenO('_variabili.autoexec')  # rinvia a autoexec in basic
    adegua_a = 216  # VERSIONE CORRENTE
    if ver_tmpl < adegua_a:
        if DLG.DlgSiNo(
                '''Vuoi procedere con l'adeguamento di questo file
alla versione di LeenO installata?

In caso affermativo dovrai attendere il completamento
dell'operazione che terminerà con un avviso.
''', "Richiesta") != 2:
            DLG.MsgBox(
                '''Non avendo effettuato l'adeguamento del file alla versione di LeenO installata, potresti avere dei malfunzionamenti!''',
                'Avviso!')
            return
        sproteggi_sheet_TUTTE()
        oDialogo_attesa = DLG.dlg_attesa(
            "Adeguamento file alla versione di LeenO installata...")
        zoom = oDoc.CurrentController.ZoomValue
        oDoc.CurrentController.ZoomValue = 400
        DLG.attesa().start()  # mostra il dialogo
        #  adeguo gli stili secondo il template corrente
        # stili = oDoc.StyleFamilies.getByName('CellStyles').getElementNames()
        # diz_stili = dict()
        ############
        # aggiungi stili di cella
        for el in ('comp 1-a PU', 'comp 1-a LUNG', 'comp 1-a LARG',
                   'comp 1-a peso', 'comp 1-a', 'Blu',
                   'Comp-Variante num sotto'):
            oStileCella = oDoc.createInstance("com.sun.star.style.CellStyle")
            if not oDoc.StyleFamilies.getByName('CellStyles').hasByName(el):
                oDoc.StyleFamilies.getByName('CellStyles').insertByName(
                    el, oStileCella)
                oStileCella.ParentStyle = 'comp 1-a'
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
#  sUrl = LeenO_path()+'/template/leeno/Computo_LeenO.ots'
#  styles = oDoc.getStyleFamilies()
#  styles.loadStylesFromURL(sUrl, list())
############
        oSheet = oDoc.getSheets().getByName('S1')
        oSheet.getCellRangeByName('S1.H291').Value = \
            oDoc.getDocumentProperties().getUserDefinedProperties().Versione = 216
        for el in oDoc.Sheets.ElementNames:
            oDoc.getSheets().getByName(el).IsVisible = True
            oDoc.CurrentController.setActiveSheet(oDoc.getSheets().getByName(el))
            # ~adatta_altezza_riga(el)
            oDoc.getSheets().getByName(el).IsVisible = False
        # dal template 212
        flags = VALUE + DATETIME + STRING + ANNOTATION + FORMULA + OBJECTS + EDITATTR  # FORMATTED + HARDATTR
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
                    lrow = next_voice(lrow, 1)
                    lrow += 1
                # ~ rigenera_tutte() affido la rigenerazione delle formule al menu Viste
                # 214 aggiorna stili di cella per ogni colonna
                test = SheetUtils.getUsedArea(oSheet).EndRow + 1
                for y in range(0, test):
                    # aggiorna formula vedi voce #214
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
                        vedi_voce_xpwe(y, vRif)
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
        oSheet.getCellRangeByName('L25').CellStyle = 'Blu ROSSO'
        oSheet.getCellRangeByName('J26').Formula = '=IF(SUBTOTAL(9;J24:J26)<0;"";SUBTOTAL(9;J24:J26))'
        oSheet.getCellRangeByName('L26').Formula = '=IF(SUBTOTAL(9;L24:L26)<0;"";SUBTOTAL(9;L24:L26))'
        oSheet.getCellRangeByName('L26').CellStyle = 'Comp-Variante num sotto ROSSO'

        # CONTABILITA CONTABILITA CONTABILITA CONTABILITA CONTABILITA
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
            # ~ rigenera_tutte() affido la rigenerazione delle formule al menu Viste
            lrow = 4
            while lrow < n:
                oDoc.CurrentController.select(oSheet.getCellByPosition(
                    0, lrow))
                sistema_stili()
                lrow = next_voice(lrow, 1)
                lrow += 1
        for el in oDoc.Sheets.ElementNames:
            oDoc.CurrentController.setActiveSheet(
                oDoc.getSheets().getByName(el))
            adatta_altezza_riga(el)
        oDoc.CurrentController.ZoomValue = zoom
        GotoSheet('COMPUTO')
        oDialogo_attesa.endExecute()  # chiude il dialogo
        mostra_fogli_principali()
        DLG.MsgBox("Adeguamento del file completato con successo.", "Avviso")

    EnableAutoCalc()


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
    for el in ("COMPUTO", "VARIANTE", "CONTABILITA"):
        try:
            importo = oDoc.getSheets().getByName(el).getCellRangeByName(
                'A2').String
            if el == 'COMPUTO':
                Dialog_XPWE.getControl(el).Label = 'Computo: €: ' + importo
            if el == 'VARIANTE':
                Dialog_XPWE.getControl(el).Label = 'Variante: €: ' + importo
            if el == 'CONTABILITA':
                Dialog_XPWE.getControl(el).Label = 'Contabilità: €: ' + importo
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
    # ~systemPathToFileUrl
    lista = list()
    #  Dialog_XPWE.execute()
    # ~try:
    # ~Dialog_XPWE.execute()
    # ~except Exception:
    # ~pass
    if Dialog_XPWE.execute() == 1:
        for el in ("COMPUTO", "VARIANTE", "CONTABILITA"):
            if Dialog_XPWE.getControl(el).State == 1:
                lista.append(el)
    out_file = Dialogs.FileSelect('Salva con nome...', '*.xpwe', 1)
    testo = '\n'
    for el in lista:
        XPWE_out(el, out_file)
        testo = testo + '● ' + out_file + '-' + el + '.xpwe\n\n'
    DLG.MsgBox('Esportazione in formato XPWE eseguita con successo su:\n' + testo, 'Avviso.')


########################################################################
def chiudi_dialoghi(event=None):
    '''
    @@ DA DOCUMENTARE
    '''
    return
    if event:
        event.Source.Context.endExecute()
    return


########################################################################
def ScriviNomeDocumentoPrincipale():
    '''
    Indica qual è il Documento Principale
    '''
    oDoc = LeenoUtils.getDocument()
    try:
        if LeenoUtils.getGlobalVar('sUltimus') == uno.fileUrlToSystemPath(oDoc.getURL()):
            return
    except Exception:
        return  # file senza nome
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
            if LeenoUtils.getGlobalVar('sUltimus') == uno.fileUrlToSystemPath(oDoc.getURL()):
                Dialogs.Exclamation(Title="BAH", Text="Inatteso!!!!")
                # 13434777 giallo
                oSheet.getCellRangeByName("A1:AT1").CellBackColor = 16773632
                oSheet.getCellRangeByName(d[el]).String = 'DP: Questo documento'
        except Exception:
            pass


####
def DlgMain():
    '''
    Visualizza il menù principale dialog_fil
    '''

    oDoc = LeenoUtils.getDocument()
    psm = LeenoUtils.getComponentContext().ServiceManager
    oSheet = oDoc.CurrentController.ActiveSheet
    if not oDoc.getSheets().hasByName('S2'):
        Toolbars.AllOff()
        if(len(oDoc.getURL()) == 0 and
           SheetUtils.getUsedArea(oSheet).EndColumn == 0 and
           SheetUtils.getUsedArea(oSheet).EndRow == 0):
            oDoc.close(True)
        New_file.computo()
    Toolbars.Vedi()
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
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
    sString.Text = version_code.read()[:-9]

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
    #  sString = oDlgMain.getControl("ComboBox1")
    #  sString.Text = conf.read(path_conf, 'Generale', 'visualizza')
    oDlgMain.getControl('CheckBox1').State = int(
        cfg.read('Generale', 'dialogo'))
    #  _gotoCella(x, y)
    oDlgMain.execute()
    sString = oDlgMain.getControl("Label_DDC").Text
    if oDlgMain.getControl('CheckBox1').State == 1:
        cfg.write('Generale', 'dialogo', '1')
        #  sString = oDlgMain.getControl("ComboBox1")
        #  conf.write(path_conf, 'Generale', 'visualizza', sString.getText())
    else:
        cfg.write('Generale', 'dialogo', '0')
        #  conf.write(path_conf, 'Generale', 'visualizza', 'Senza Menù')
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
        except Exception:
            pass
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
    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400
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
    oDoc.CurrentController.ZoomValue = zoom


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
    # filename = '.'.join(os.path.basename(orig).split('.')[0:-1]) + '-'
    if len(orig) == 0:
        return
    if not os.path.exists(uno.fileUrlToSystemPath(dir_bak)):
        os.makedirs(uno.fileUrlToSystemPath(dir_bak))
    orig = uno.fileUrlToSystemPath(orig)
    dest = uno.fileUrlToSystemPath(dest)
    if os.path.exists(uno.fileUrlToSystemPath(dir_bak) + dest):
        shutil.copyfile(
            uno.fileUrlToSystemPath(dir_bak) + dest,
            uno.fileUrlToSystemPath(dir_bak) + dest + '.old')
    shutil.copyfile(orig, uno.fileUrlToSystemPath(dir_bak) + dest)


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
    oDoc.storeToURL(dir_bak + dest, list())
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
        Ldev = str (int(f.readline().split('-')[0].split('.')[-1]) + 1)
        tempo = ''.join(''.join(''.join(str(datetime.now()).split('.')[0].split(' ')).split('-')).split(':'))
        of = open(code_file, 'w')

        new = (
            str(LeenoUtils.getGlobalVar('Lmajor')) + '.' +
            str(LeenoUtils.getGlobalVar('Lminor')) + '.' +
            LeenoUtils.getGlobalVar('Lsubv').split('.')[0] + '.' +
            Ldev + '-TESTING-' +
            tempo[:-6])
        of.write(new)
        of.close()
        return new


########################################################################
def MENU_grid_switch():
    '''Mostra / nasconde griglia'''
    oDoc = LeenoUtils.getDocument()
    oDoc.CurrentController.ShowGrid = not oDoc.CurrentController.ShowGrid


def MENU_make_pack():
    '''
    @@ DA DOCUMENTARE
    '''
    make_pack()


def make_pack(bar=0):
    '''
    bar { integer } : toolbar 0=spenta 1=accesa
    Pacchettizza l'estensione in duplice copia: LeenO.oxt e LeenO-yyyymmddhhmm.oxt
    in una directory precisa(per ora - da parametrizzare)
    '''
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
    if bar == 0:
        oDoc = LeenoUtils.getDocument()
        Toolbars.AllOff()
        oLayout = oDoc.CurrentController.getFrame().LayoutManager
        oLayout.hideElement(
            "private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV")
    oxt_path = uno.fileUrlToSystemPath(LeenO_path())
    if sys.platform == 'linux' or sys.platform == 'darwin':
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
    elif sys.platform == 'win32':
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
    shutil.make_archive(nomeZip2, 'zip', oxt_path)
    shutil.move(nomeZip2 + '.zip', nomeZip2)
    #~ shutil.copyfile(nomeZip2, nomeZip)


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
    oDialog1.execute()


########################################################################
def donazioni():
    '''
    @@ DA DOCUMENTARE
    '''
    apri = LeenoUtils.createUnoService("com.sun.star.system.SystemShellExecute")
    apri.execute("https://leeno.org/donazioni/", "", 0)


########################################################################
#  class firme_in_calce_th(threading.Thread):
#  def __init__(self):
#  threading.Thread.__init__(self)
#  def run(self):
#  firme_in_calce_run()
#  def firme_in_calce():
#  firme_in_calce_th().start()
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
        oDialogo_attesa = DLG.dlg_attesa()
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
            return
        descrizione = InputBox(t='inserisci una descrizione per la nuova riga')
        DLG.attesa().start()  # mostra il dialogo
        zoom = oDoc.CurrentController.ZoomValue
        oDoc.CurrentController.ZoomValue = 400
        i = 0
        while (i < SheetUtils.getUsedArea(oSheet).EndRow):

            if oSheet.getCellByPosition(2, i).CellStyle == 'comp 1-a':
                sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, i)
                qui = sStRange.RangeAddress.StartRow + 1

                i = sotto = sStRange.RangeAddress.EndRow + 3
                oDoc.CurrentController.select(oSheet.getCellByPosition(2, qui))
                Copia_riga_Ent()
                oSheet.getCellByPosition(2, qui + 1).String = descrizione
                next_voice(sotto)

                oDoc.CurrentController.select(oSheet.getCellByPosition(2, i))
            i += 1
        oDialogo_attesa.endExecute()  # chiude il dialogo
        oDoc.CurrentController.ZoomValue = zoom


def MENU_inserisci_nuova_riga_con_descrizione():
    '''
    inserisce, all'inizio di ogni voce di computo o variante,
    una nuova riga con una descrizione a scelta
    '''
    inserisci_nuova_riga_con_descrizione_th().start()


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
    copy_clip()
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
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()

    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
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
    # oRange = oSheet.getCellRangeByPosition(sCol, sRow, eCol, eRow)
    flags = VALUE + DATETIME + STRING + ANNOTATION + FORMULA + OBJECTS + EDITATTR  # FORMATTED + HARDATTR
    oSheet.getCellRangeByPosition(sCol, sRow, eCol, eRow).clearContents(flags)


########################################################################
def calendario_mensile():
    '''
    Colora le colonne del sabato e della domenica, oltre le festività,
    nel file ../PRIVATO/LeenO/extra/calendario.ods che potrei implementare
    in LeenO per la gestione delle ore in economia o del diagramma di Gantt.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName('elenco festività')
    oRangeAddress = oDoc.NamedRanges.feste.ReferredCells.RangeAddress
    SR = oRangeAddress.StartRow
    ER = oRangeAddress.EndRow
    lFeste = list()
    for x in range(SR, ER):
        if oSheet.getCellByPosition(0, x).Value != 0:
            lFeste.append(oSheet.getCellByPosition(0, x).String)
    oSheet = oDoc.getSheets().getByName('CALENDARIO')
    test = SheetUtils.getUsedArea(oSheet).EndColumn + 1
    slist = list()
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
def sistema_cose():
    '''
    @@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lcol = LeggiPosizioneCorrente()[0]
    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
    el_y = list()
    # el_x = list()
    lista_y = list()
    try:
        len(oRangeAddress)
        for el in oRangeAddress:
            el_y.append((el.StartRow, el.EndRow))
    except TypeError:
        el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))
    for y in el_y:
        for el in range(y[0], y[1] + 1):
            lista_y.append(el)
    for y in lista_y:
        oDoc.CurrentController.select(oSheet.getCellByPosition(lcol, y))
        if oDoc.getCurrentSelection().Type.value == 'TEXT':
            testo = oDoc.getCurrentSelection().String.replace(
                '\t', ' ').replace('\n', ' ').replace('Ã¨', 'è').replace(
                    'Â°', '°').replace('Ã', 'à').replace(' $', '')
            while '  ' in testo:
                testo = testo.replace('  ', ' ')
            oDoc.getCurrentSelection().String = testo.strip().strip().strip()


########
def debug_link():
    '''
    @@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()
    window = oDoc.getCurrentController().getFrame().getContainerWindow()
    ctx = LeenoUtils.getComponentContext()

    def create(name):
        return ctx.getServiceManager().createInstanceWithContext(name, ctx)

    toolkit = create("com.sun.star.awt.Toolkit")
    msgbox = toolkit.createMessageBox(window, 0, 1, "Message", 'foo')
    link = create("com.sun.star.awt.UnoControlFixedHyperlink")
    link_model = create("com.sun.star.awt.UnoControlFixedHyperlinkModel")
    link.setModel(link_model)
    link.createPeer(toolkit, msgbox)
    link.setPosSize(35, 8, 100, 15, 15)
    link.setText("Canale Telegram")
    link.setURL("https://t.me/leeno_computometrico")
    link.setVisible(True)
    msgbox.execute()
    msgbox.dispose()


########################################################################
def descrizione_in_una_colonna(flag=False):
    '''
    Consente di estendere su più colonne o ridurre ad una colonna lo spazio
    occupato dalla descrizione di voce in COMPUTO, VARIANTE e CONTABILITA.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet = oDoc.getSheets().getByName('S5')
    oSheet.getCellRangeByName('C9:I9').merge(flag)
    oSheet.getCellRangeByName('C10:I10').merge(flag)
    oSheet = oDoc.getSheets().getByName('COMPUTO')
    for y in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
        if oSheet.getCellByPosition(
                2, y).CellStyle in ('Comp-Bianche sopraS',
                                    'Comp-Bianche in mezzo Descr'):
            oSheet.getCellRangeByPosition(2, y, 8, y).merge(flag)
    if oDoc.getSheets().hasByName('VARIANTE'):
        oSheet = oDoc.getSheets().getByName('VARIANTE')
        for y in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
            if oSheet.getCellByPosition(
                    2, y).CellStyle in ('Comp-Bianche sopraS',
                                        'Comp-Bianche in mezzo Descr'):
                oSheet.getCellRangeByPosition(2, y, 8, y).merge(flag)
    if oDoc.getSheets().hasByName('CONTABILITA'):
        if oDoc.NamedRanges.hasByName("#Lib#1"):
            DLG.MsgBox(
                "Risulta già registrato un SAL. NON E' POSSIBILE PROCEDERE.",
                'ATTENZIONE!')
            return
        oSheet = oDoc.getSheets().getByName('S5')
        oSheet.getCellRangeByName('C23').merge(flag)
        oSheet.getCellRangeByName('C24').merge(flag)
        oSheet = oDoc.getSheets().getByName('CONTABILITA')
        for y in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
            if oSheet.getCellByPosition(
                    2, y).CellStyle in ('Comp-Bianche sopra_R',
                                        'Comp-Bianche in mezzo Descr_R'):
                oSheet.getCellRangeByPosition(2, y, 8, y).merge(flag)
    else:
        oSheet = oDoc.getSheets().getByName('S5')
        oSheet.getCellRangeByName('C23:I23').merge(flag)
        oSheet.getCellRangeByName('C24:I24').merge(flag)
    return


########################################################################
def numera_colonna():
    '''Inserisce l'indice di colonna nelle prime 100 colonne del rigo selezionato
Associato a Ctrl+Shift+C'''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]
    for x in range(0, 50):
        if oSheet.getCellByPosition(x, lrow).Type.value == 'EMPTY':
            oSheet.getCellByPosition(x, lrow).Formula = '=CELL("col")-1'
            oSheet.getCellByPosition(x, lrow).HoriJustify = 'CENTER'
        elif oSheet.getCellByPosition(x, lrow).Formula == '=CELL("col")-1':
            oSheet.getCellByPosition(x, lrow).String = ''
            oSheet.getCellByPosition(x, lrow).HoriJustify = 'STANDARD'


########################################################################
def subst_str():
    '''
    Sostituisce stringhe di testi nel foglio corrente
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    ReplaceDescriptor = oSheet.createReplaceDescriptor()
    ReplaceDescriptor.SearchString = "str1"
    ReplaceDescriptor.ReplaceString = "str2"
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
def MENU_sistema_pagine():
    '''
    Configura intestazioni e pie' di pagina degli stili di stampa
    e propone un'anteprima
    '''
    basic_LeenO('Finale.Set_Area_Stampa_N')
    basic_LeenO('Finale.Pulisci_Tabella_Tutta')
    oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('M1'):
        return
    #  committente = oDoc.NamedRanges.Super_ego_8.ReferredCells.String
    # committente = oDoc.getSheets().getByName('S2').getCellRangeByName("C6").String  # committente
    # luogo = oDoc.getSheets().getByName('S2').getCellRangeByName("C4").String

    # try:
    #    oSheet = oDoc.CurrentController.ActiveSheet
    # except Exception:
    #    pass
    #  su_dx = oDoc.NamedRanges.Bozza_8.ReferredCells.String

    # cancella stili di pagina #######################################
    # ~stili_pagina = list()
    # ~fine = oDoc.StyleFamilies.getByName('PageStyles').Count
    # ~for n in range(0, fine):
    # ~oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByIndex(n)
    # ~stili_pagina.append(oAktPage.DisplayName)
    # ~for el in stili_pagina:
    # ~if el not in ('PageStyle_Analisi di Prezzo', 'Page_Style_COPERTINE',
    # 'Page_Style_Libretto_Misure2', 'PageStyle_REGISTRO_A4', 'PageStyle_COMPUTO_A4',
    # 'PageStyle_Elenco Prezzi'):
    # ~oDoc.StyleFamilies.getByName('PageStyles').removeByName(el)
    # ~return
    # cancella stili di pagina #######################################
    stili = {
        'cP_Cop': 'Page_Style_COPERTINE',
        'COMPUTO': 'PageStyle_COMPUTO_A4',
        'VARIANTE': 'PageStyle_COMPUTO_A4',
        'Elenco Prezzi': 'PageStyle_Elenco Prezzi',
        'Analisi di Prezzo': 'PageStyle_Analisi di Prezzo',
        'CONTABILITA': 'Page_Style_Libretto_Misure2',
        'Registro': 'PageStyle_REGISTRO_A4',
        'SAL': 'PageStyle_REGISTRO_A4',
    }
    for el in stili.keys():
        try:
            oDoc.getSheets().getByName(el).PageStyle = stili[el]
        except Exception:
            pass
    ###
    #  oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByName('PageStyle_COMPUTO_A4')
    #  mri(oAktPage)
    #  return
    ###
    if cfg.read('Generale', 'dettaglio') == '1':
        dettaglio_misure(0)
        dettaglio_misure(1)
    else:
        dettaglio_misure(0)
    for n in range(0, oDoc.StyleFamilies.getByName('PageStyles').Count):
        oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByIndex(n)
        # ~chi((n , oAktPage.DisplayName))
        oAktPage.HeaderIsOn = True
        oAktPage.FooterIsOn = True

        if oAktPage.DisplayName == 'Page_Style_COPERTINE':
            oAktPage.HeaderIsOn = False
            oAktPage.FooterIsOn = False
            # Adatto lo zoom alla larghezza pagina
            oAktPage.PageScale = 0
            oAktPage.CenterHorizontally = True
            oAktPage.ScaleToPagesX = 1
            oAktPage.ScaleToPagesY = 0
        if oAktPage.DisplayName in ('PageStyle_Analisi di Prezzo',
                                    'PageStyle_COMPUTO_A4',
                                    'PageStyle_Elenco Prezzi'):
            # htxt = 8.0
            # if oAktPage.DisplayName in ('PageStyle_Analisi di Prezzo'):
            #     htxt = 10.0
            bordo = oAktPage.TopBorder
            bordo.LineWidth = 0
            bordo.OuterLineWidth = 0
            oAktPage.TopBorder = bordo

            bordo = oAktPage.BottomBorder
            bordo.LineWidth = 0
            bordo.OuterLineWidth = 0
            oAktPage.BottomBorder = bordo

            bordo = oAktPage.LeftBorder
            bordo.LineWidth = 0
            bordo.OuterLineWidth = 0
            oAktPage.LeftBorder = bordo
            # bordo lato destro attivo in attesa di LibreOffice 6.2
            #  bordo = oAktPage.RightBorder
            #  bordo.Color = 0
            #  bordo.LineWidth = 2
            #  bordo.OuterLineWidth = 2
            #  oAktPage.RightBorder = bordo
            # Adatto lo zoom alla larghezza pagina
            oAktPage.PageScale = 0
            oAktPage.ScaleToPagesX = 1
            oAktPage.ScaleToPagesY = 0

            # ~HEADER
            oHeader = oAktPage.RightPageHeaderContent
            # ~oAktPage.PageScale = 95
            # oHLText = oHeader.LeftText.Text.String = committente
            # oHRText = oHeader.LeftText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            # oHRText = oHeader.LeftText.Text.Text.CharHeight = htxt  # / 100 * oAktPage.PageScale
            # oHRText = oHeader.RightText.Text.String = luogo
            # oHRText = oHeader.RightText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            # oHRText = oHeader.RightText.Text.Text.CharHeight = htxt  # / 100 * oAktPage.PageScale

            oAktPage.RightPageHeaderContent = oHeader
            # ~FOOTER
            oFooter = oAktPage.RightPageFooterContent
            # oHLText = oFooter.CenterText.Text.String = ''
            # oHLText = oFooter.LeftText.Text.String = "realizzato con LeenO.org\n" + os.path.basename(oDoc.getURL())
            # oHRText = oFooter.LeftText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            # oHRText = oFooter.LeftText.Text.Text.CharHeight = htxt  # / 100 * oAktPage.PageScale
            # oHRText = oFooter.RightText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            # oHRText = oFooter.RightText.Text.Text.CharHeight = htxt  # / 100 * oAktPage.PageScale
            # oHRText = oFooter.RightText.Text.String = '#/##'
            oAktPage.RightPageFooterContent = oFooter

        if oAktPage.DisplayName == 'Page_Style_Libretto_Misure2':
            # ~HEADER
            oHeader = oAktPage.RightPageHeaderContent
            # oHLText = oHeader.LeftText.Text.String = committente + '\nLibretto delle misure n.'
            # oHRText = oHeader.RightText.Text.String = luogo
            oAktPage.RightPageHeaderContent = oHeader
            # ~FOOTER
            oFooter = oAktPage.RightPageFooterContent
            # oHLText = oFooter.CenterText.Text.String = "L'IMPRESA					IL DIRETTORE DEI LAVORI"
            # oHLText = oFooter.LeftText.Text.String = "realizzato con LeenO.org\n" + os.path.basename(
            #    oDoc.getURL() + '\n\n\n')
            oAktPage.RightPageFooterContent = oFooter

        if oAktPage.DisplayName == 'PageStyle_REGISTRO_A4':
            # ~HEADER
            oHeader = oAktPage.RightPageHeaderContent
            # oHLText = oHeader.LeftText.Text.String = committente + '\nRegistro di contabilità n.'
            # oHRText = oHeader.RightText.Text.String = luogo
            oAktPage.RightPageHeaderContent = oHeader
            # ~FOOTER
            oFooter = oAktPage.RightPageFooterContent
            # oHLText = oFooter.CenterText.Text.String = ''
            # oHLText = oFooter.LeftText.Text.String = "realizzato con LeenO.org\n" + os.path.basename(
            #    oDoc.getURL() + '\n\n\n')
            oAktPage.RightPageFooterContent = oFooter
    try:
        if oDoc.CurrentController.ActiveSheet.Name in ('COMPUTO', 'VARIANTE',
                                                       'CONTABILITA',
                                                       'Elenco Prezzi'):
            _gotoCella(0, 3)
        if oDoc.CurrentController.ActiveSheet.Name in ('Analisi di Prezzo'):
            _gotoCella(0, 2)
        setPreview(1)
    except Exception:
        pass
        # bordo lato destro in attesa di LibreOffice 6.2
        # bordo = oAktPage.RightBorder
        # bordo.LineWidth = 0
        # bordo.OuterLineWidth = 0
        # oAktPage.RightBorder = bordo
    return


########################################################################
def fissa():
    '''
    @@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lcol = LeggiPosizioneCorrente()[0]
    lrow = LeggiPosizioneCorrente()[1]
    if oSheet.Name in (
            'COMPUTO',
            'VARIANTE',
            'CONTABILITA',
    ):
        oDoc.CurrentController.freezeAtPosition(0, 3)
    elif oSheet.Name in ('Elenco Prezzi'):
        #  _gotoCella(0, 3)
        oDoc.CurrentController.freezeAtPosition(0, 3)
    elif oSheet.Name in ('Analisi di Prezzo'):
        oDoc.CurrentController.freezeAtPosition(0, 2)
    _gotoCella(lcol, lrow)


########################################################################
def debug_errore():
    '''
    @@ DA DOCUMENTARE
    '''
    #  sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    #  return

    try:
        # sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        LeenoComputo.circoscriveVoceComputo(oSheet, lrow)

    except Exception as e:
        #  MsgBox ("CSV Import failure exception " + str(type(e)) +
        #  ". Messaggio: " + str(e) + " args " + str(e.args) +
        #  traceback.format_exc());
        DLG.MsgBox("Eccezione " + str(type(e)) + "\nMessaggio: " + str(e.args) + '\n' + traceback.format_exc())


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
        lista = list()
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
        DLG.MsgBox('Non ci sono voci di prezzo ricorrenti.', 'Informazione')
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
# def debug_():
#     '''
#     @@ DA DOCUMENTARE
#     '''
#     oDoc = LeenoUtils.getDocument()
#     oSheet = oDoc.CurrentController.ActiveSheet
#     if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
#         try:
#             sRow = oDoc.getCurrentSelection().getRangeAddresses()[0].StartRow
#             eRow = oDoc.getCurrentSelection().getRangeAddresses()[0].EndRow
#         except Exception:
#             sRow = oDoc.getCurrentSelection().getRangeAddress().StartRow
#             eRow = oDoc.getCurrentSelection().getRangeAddress().EndRow
#     DLG.chi((sRow, eRow))


# def debug_():
#     refresh(0)
#     oDoc = LeenoUtils.getDocument()
#     oSheet = oDoc.CurrentController.ActiveSheet
#     try:
#         oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
#     except AttributeError:
#         oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
#     el_y = list()
#     try:
#         len(oRangeAddress)
#         for el in oRangeAddress:
#             el_y.append((el.StartRow, el.EndRow))
#     except TypeError:
#         el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))
#     lista = list()
#     for y in el_y:
#         for el in range(y[0], y[1] + 1):
#             lista.append(el)
#     for el in lista:
#         oSheet.getCellByPosition(
#             7, el).Formula = '=' + oSheet.getCellByPosition(
#                 6, el).Formula + '*' + oSheet.getCellByPosition(7, el).Formula
#         oSheet.getCellByPosition(6, el).String = ''
#     refresh(1)


########################################################################
def trova_np():
    '''
    Raggruppa le righe in modo da rendere evidenti i nuovi prezzi
    '''
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    DisableAutoCalc()

    struttura_off()
    oCellRangeAddr = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    for el in range(3, SheetUtils.getUsedArea(oSheet).EndRow):
        if oSheet.getCellByPosition(
                12,
                el).Value == 0 and oSheet.getCellByPosition(20, el).Value > 0:
            pass
        else:
            oCellRangeAddr.StartRow = el
            oCellRangeAddr.EndRow = el
            oSheet.group(oCellRangeAddr, 1)
            oSheet.getCellRangeByPosition(0, el, 1, el).Rows.IsVisible = False

    EnableAutoCalc()


def debug_syspath():
    '''
    @@ DA DOCUMENTARE
    '''
    # ~pydevd.settrace()
    # pathsstring = "paths \n"
    somestring = ''
    for i in sys.path:
        somestring = somestring + i + "\n"
    DLG.chi(somestring)


def debug_progressbar():
    '''
    @@ DA DOCUMENTARE
    '''
    try:
        oDoc = LeenoUtils.getDocument()
        # set up Status Indicator
        oCntl = oDoc.getCurrentController()
        oFrame = oCntl.getFrame()
        oSI = oFrame.createStatusIndicator()
        oEnd = 100
        oSI.reset()  # Reset : NG
        oSI.start('Excuting', oEnd)  # Start : NG

        for i in range(1, 11):
            oSI.setText('Processing: ' + str(i))
            oSI.setValue(20 * i)
            time.sleep(0.1)

        oSI.setText('Finished')
        oDisp = 'Success'
    except Exception as er:
        oDisp = ''
        oDisp = str(traceback.format_exc()) + '\n' + str(er)
    finally:
        DLG.MsgBox(oDisp)
        oSI.end()


########################################################################
def elimina_voci_doppie():
    '''
    @@ DA DOCUMENTARE
    '''
    DLG.chi('prova')
    # elimina voci doppie hard - grezza e lenta, ma efficace
    oDoc = LeenoUtils.getDocument()
    GotoSheet('Elenco Prezzi')
    oSheet = oDoc.CurrentController.ActiveSheet
    riordina_ElencoPrezzi()
    fine = SheetUtils.getUsedArea(oSheet).EndRow + 1

    oSheet.getCellByPosition(30, 3).Formula = '=IF(A4=A3;1;0)'
    oDoc.CurrentController.select(oSheet.getCellByPosition(30, 3))
    copy_clip()
    oDoc.CurrentController.select(
        oSheet.getCellRangeByPosition(30, 3, 30, fine))
    paste_clip(insCells=1)
    # ~oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
    # ~for i in range (3, fine):
    # ~oSheet.getCellByPosition(30, i).Formula = '=IF(A' + str(i+1) + '=A' + str(i) + ';1;0)'
    # ~return
    for i in reversed(range(0, fine)):
        if oSheet.getCellByPosition(30, i).Value == 1:
            _gotoCella(30, i)
            oSheet.getRows().removeByIndex(i, 1)
    oSheet.getCellRangeByPosition(30, 3, 30, fine).clearContents(FORMULA)


########################################################################
def hl():
    '''
    Sostituisce hiperlink alla stringa nella colonna B, se questa è un
    indirizzo di file o cartella
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    for el in reversed(range(0, SheetUtils.getUsedArea(oSheet).EndRow)):
        try:
            if oSheet.getCellByPosition(1, el).String[1] == ':':
                stringa = '=HYPERLINK("' + oSheet.getCellByPosition(
                    1, el).String + '";"LINK")'
                oSheet.getCellByPosition(1, el).Formula = stringa
        except Exception:
            pass


########################################################################
def filtro_descrizione():
    '''
    Raggruppa e nasconde tutte le voci di misura in cui non compare
    la stringa cercata.
    '''
    struttura_off()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.getCellRangeByPosition(2, 0, 2, 1048575).clearContents(HARDATTR)

    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    fine = SheetUtils.getUsedArea(oSheet).EndRow + 1
    el_y = list()
    if oDoc.getCurrentSelection().CellStyle == 'comp 1-a':
        testo = oDoc.getCurrentSelection().String
    else:
        testo = ''
    descrizione = InputBox(
        testo, t='Inserisci la descrizione da cercare o OK per conferma.')
    if descrizione in (None, ''):
        return
    y = 4
    while y < fine:
        test = SheetUtils.uFindStringCol(descrizione, 2, oSheet, y)
        if test is not None:
            y = test
            oSheet.getCellByPosition(2, y).CellBackColor = 15757935
            el_y.append(seleziona_voce(y))
            try:
                y = next_voice(seleziona_voce(y)[1])
            except TypeError:
                DLG.MsgBox(
                    '''Questo comando non produce risultato se il cursore
è oltre la riga rossa di Fine Computo.''', 'ATTENZIONE!')
                return
            y += 1
        y += 1
    if len(el_y) == 0:
        DLG.MsgBox('''Testo non trovato.''', 'ATTENZIONE!')
        return
    lista_y = list()
    lista_y.append(2)
    for el in el_y:
        y = el[0]
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


########################################################################
# sardegna_2019 moved to LeenoImport.py
########################################################################

########################################################################
# basilicata_2020 moved to LeenoImport.py
########################################################################

########################################################################
# Piemonte_2019 moved to LeenoImport.py
########################################################################


########################################################################
# def debug_():
    # '''cambio data contabilità'''
    # oDoc = LeenoUtils.getDocument()
    # #  mri(oDoc)
    # oSheet = oDoc.CurrentController.ActiveSheet
    # DLG.chi(oDoc.getCurrentSelection().CellBackColor)
    # # ~ return
    # fine = SheetUtils.getUsedArea(oSheet).EndRow + 1
    # for i in range(0, fine):
    #     if oSheet.getCellByPosition(1, i).String == 'Data_bianca':
    #         oSheet.getCellByPosition(1, i).Value = 43861.0


########################################################################
def errore():
    '''
    @@ DA DOCUMENTARE
    '''
    DLG.MsgBox(traceback.format_exc())

#~from collections import OrderedDict

def MENU_debug():
    '''
    Utile per testare comandi dalla toolbar DEV
    '''
    #~DLG.chi('poipo')
    #~return
    sistema_cose()
    return
#~ def split_chunks(l, n):
    """
       Splits list l into n chunks with approximately equals sum of values
       see  http://stackoverflow.com/questions/6855394/splitting-list-in-chunks-of-balanced-weight
    """
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lr = SheetUtils.getLastUsedRow() + 1
    l = list()
    d = list()
    for i in range(1, lr):
        l.append (oSheet.getCellByPosition(1, i).Value)

        a = oSheet.getCellByPosition(1, i).Value
        b = oSheet.getCellByPosition(0, i).AbsoluteName
        d.append ([a, b])

    n = 6
    result = [[] for i in range(n)]
    sums   = {i:0 for i in range(n)}
    c = 0
    for e in l:
        for i in sums:
            if c == sums[i]:
                result[i].append(e)
                break
        sums[i] += e
        c = min(sums.values())
    c = 5
    for el in result:
        n = 1
        for i in el:
            oSheet.getCellByPosition(c, n).Value = i
            m = 0
            for x in d:
                if x[0] == i:
                    formula = x[1]
                    d.pop(m)
                    break
                m +=1
            oSheet.getCellByPosition(c+1, n).Formula = '=' + formula
            n +=1
        c +=2
    return result

########################################################################
# ELENCO DEGLI SCRIPT VISUALIZZATI NEL SELETTORE DI MACRO              #
g_exportedScripts = donazioni
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
