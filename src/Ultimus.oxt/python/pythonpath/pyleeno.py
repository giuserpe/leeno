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

    # funzioni per misurare la velocità dalle macro
    # ~datarif = datetime.now()
    # ~DLG.chi('eseguita in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!')


from datetime import datetime, date
from xml.etree.ElementTree import Element, SubElement, tostring

# import distutils.dir_util

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
import LeenoContab
import LeenoAnalysis
import LeenoDialogs as DLG
import PersistUtils as PU
import LeenoEvents
import LeenoBasicBridge

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
    if oDoc.NamedRanges.hasByName("_Lib_1"):
        sString.setEnable(False)
    sString = oDlg_config.getControl('TextField12')
    sString.Text = oSheet.getCellRangeByName(
        'S1.H336').String  # cont_fine_voci_abbreviate
    if oDoc.NamedRanges.hasByName("_Lib_1"):
        sString.setEnable(False)

    if cfg.read('Generale', 'torna_a_ep') == '1':
        oDlg_config.getControl('CheckBox8').State = 1

    sString = oDlg_config.getControl('ComboBox4')
    sString.Text = cfg.read('Generale', 'copie_backup')
    if int(cfg.read('Generale', 'copie_backup')) != 0:
        sString = oDlg_config.getControl('ComboBox5')
        sString.Text = cfg.read('Generale', 'pausa_backup')
    # ~else: 
        # ~oDlg_config.getControl('ComboBox5').setEnable(False)
        # ~oDlg_config.execute()
        # ~DLG.chi(oDlg_config.getControl('ComboBox5'))


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
        cfg.write('Contabilita', 'cont_inizio_voci_abbreviate', oDlg_config.getControl('TextField4').getText())
    oSheet.getCellRangeByName('S1.H335').Value = float(oDlg_config.getControl('TextField4').getText())

    if oDlg_config.getControl('TextField12').getText() != '10000':
        cfg.write('Contabilita', 'cont_fine_voci_abbreviate', oDlg_config.getControl('TextField12').getText())
    oSheet.getCellRangeByName('S1.H336').Value = float(oDlg_config.getControl('TextField12').getText())
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)

    cfg.write('Generale', 'copie_backup', oDlg_config.getControl('ComboBox4').getText())
    cfg.write('Generale', 'pausa_backup', oDlg_config.getControl('ComboBox5').getText())
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
    document = desktop.loadComponentFromURL(
        LeenO_path() + '/template/leeno/Computo_LeenO.ots', "_blank", 0,
        (opz, ))
    autoexec()
    if arg == 1:
        Dialogs.Exclamation(Title = 'ATTENZIONE!',
        Text='''
Prima di procedere è
meglio dare un nome al file.
Lavorando su un file senza nome
potresti avere dei malfunzionamenti.\n
''')
        # ~DLG.MsgBox(
            # ~"Prima di procedere è consigliabile salvare il lavoro.\n"
            # ~"Provvedi subito a dare un nome al file di computo...",
            # ~"Dai un nome al file...")
        salva_come()
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
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # ~zoom = oDoc.CurrentController.ZoomValue
    # ~oDoc.CurrentController.ZoomValue = 400

    elenco = seleziona()
    codici = []
    for el in elenco:
        cod = oSheet.getCellByPosition(0, el).String
        codici.append(cod)
    dest = oSheet.getCellRangeByName('C2').String
    
    if dest == 'VARIANTE':
        genera_variante()
    elif dest == 'CONTABILITA':
        LeenoContab.attiva_contabilita()
        # ~ins_voce_contab()
    elif dest == 'COMPUTO':
        GotoSheet(dest)
    else:
        Dialogs.Exclamation(Title='AVVISO!',
    Text='''Per procedere devi prima scegliere,
dalla cella "C2", l'elaborato a cui
inviare le voci di prezzo selezionate.

Se l'elaborato è già esistente,
assicurati di aver scelto anche
la posizione di destinazione.''')
        _gotoCella(2, 1)
        return
    oSheet = oDoc.getSheets().getByName(dest)
    for el in codici:
        if oSheet.Name == 'CONTABILITA':
            GotoSheet(dest)
            ins_voce_contab(cod=el)
        else:
            LeenoComputo.ins_voce_computo(cod=el)
        lrow = SheetUtils.getLastUsedRow(oSheet)
    # ~oDoc.CurrentController.ZoomValue = zoom
    return

def MENU_invia_voce():
    '''
    Invia le voci di computo, elenco prezzi e analisi, con costi elementari,
    dal documento corrente al Documento Principale.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_cat = LeenoUtils.getGlobalVar('stili_cat')

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
            Dialogs.Exclamation(Title = 'ATTENZIONE!',
            Text='La posizione di PARTENZA non è corretta.')
            # ~DLG.MsgBox('La posizione di PARTENZA non è corretta.', 'ATTENZIONE!')
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
        comando('Copy')
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

            # ~if dccSheet.getCellByPosition(0, lrow).CellStyle in stili_cat:
                # ~DLG.chi(dccSheet.getCellByPosition(0, lrow).CellStyle)
                # ~lrow += 1

            if dccSheet.getCellByPosition(0, lrow).CellStyle in ('comp Int_colonna'):
                LeenoComputo.insertVoceComputoGrezza(dccSheet, lrow + 1)
                # @@ PROVVISORIO !!!
                _gotoCella(1, lrow + 2)
                numera_voci(1)
                lrow = LeggiPosizioneCorrente()[1]
            if dccSheet.getCellByPosition(
                 0, lrow).CellStyle in (stili_cat + stili_computo + ('comp Int_colonna', )):
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
        LeenoUtils.setGlobalVar('cod', codice_voce(lrow))
        try:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
        except AttributeError:
            oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
        try:
            SR = oRangeAddress.StartRow
            SR = LeenoComputo.circoscriveVoceComputo(oSheet, SR).RangeAddress.StartRow
        except AttributeError:
            Dialogs.Exclamation(Title = 'ATTENZIONE!',
            Text='''La selezione delle voci dal COMPUTO
di partenza deve essere contigua.''')
            # ~DLG.MsgBox(
                # ~'La selezione delle voci dal COMPUTO di partenza\ndeve essere contigua.',
                # ~'ATTENZIONE!')
            return
        ER = oRangeAddress.EndRow
        ER = LeenoComputo.circoscriveVoceComputo(oSheet, ER).RangeAddress.EndRow
        oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, SR, 100, ER))

        oSheet.getCellRangeByPosition(45, SR, 45, ER).CellBackColor = 15757935

        lista = list()
        for el in range(SR, ER + 1):
            if oSheet.getCellByPosition(0, el).CellStyle in ('Comp Start Attributo'):
                lista.append(codice_voce(el))
        # seleziona()
        if nSheetDCC in ('Analisi di Prezzo'):
            # ~DLG.MsgBox('Il foglio di destinazione non è corretto.', 'ATTENZIONE!')
            Dialogs.Exclamation(Title = 'ATTENZIONE!',
            Text='Il foglio di destinazione non è corretto.')
            oDoc.CurrentController.select(
                oDoc.createInstance(
                    "com.sun.star.sheet.SheetCellRanges"))  # unselect
            return
        if nSheetDCC in ('COMPUTO', 'VARIANTE'):
            comando('Copy')
            # arrivo
            _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
            ddcDoc = LeenoUtils.getDocument()
            dccSheet = ddcDoc.getSheets().getByName(nSheet)
            lrow = LeggiPosizioneCorrente()[1]
            if dccSheet.getCellByPosition(0, lrow).CellStyle in ('comp Int_colonna', ):
                lrow = LeggiPosizioneCorrente()[1] + 1
            elif dccSheet.getCellByPosition(0, lrow).CellStyle not in stili_computo + stili_cat:
                # ~DLG.MsgBox('La posizione di destinazione non è corretta.', 'ATTENZIONE!')
                Dialogs.Exclamation(Title = 'ATTENZIONE!',
                Text='La posizione di destinazione non è corretta.')
                # unselect
                oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
                return
            else:
                lrow = LeenoSheetUtils.prossimaVoce(dccSheet, LeggiPosizioneCorrente()[1], 1)
            _gotoCella(0, lrow)
            paste_clip(insCells=1)
            numera_voci(1)
            last = lrow + ER - SR + 1
            while lrow < last:
                rigenera_voce(lrow)
                lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 1)
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
            comando('Copy')
            #
            _gotoDoc(LeenoUtils.getGlobalVar('sUltimus'))
            ddcDoc = LeenoUtils.getDocument()
            dccSheet = ddcDoc.getSheets().getByName('Elenco Prezzi')
            GotoSheet('Elenco Prezzi')
            _gotoCella(0, 3)
            paste_clip(insCells=1)
            # EliminaVociDoppieElencoPrezzi()
        if nSheetDCC in ('Elenco Prezzi'):
            # ~DLG.MsgBox("Non è possibile inviare voci da un COMPUTO all'Elenco Prezzi.")
            Dialogs.Exclamation(Title = 'ATTENZIONE!',
            Text="Non è possibile inviare voci da un COMPUTO all'Elenco Prezzi.")
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

        comando('Copy')

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
    # ~if DLG.DlgSiNo("Ricerco ed elimino le voci di prezzo duplicate?") == 2:
        # ~EliminaVociDoppieElencoPrezzi()
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    # ~LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    GotoSheet(nSheetDCC)
    if nSheetDCC in ('COMPUTO', 'VARIANTE'):
        lrow = LeggiPosizioneCorrente()[1]
        _gotoCella(2, lrow + 1)
    oSheet = oDoc.getSheets().getByName(nSheetDCC)
    # ~LeenoSheetUtils.adattaAltezzaRiga(oSheet)


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
            dest = 'w:\\_dwg\\ULTIMUSFREE\\_SRC\\leeno\\src\\Ultimus.oxt'
        subprocess.Popen(
            'w: && cd w:/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt && "C:/Program Files/Git/git-bash.exe"',
            shell=True,
            stdout=subprocess.PIPE)
    return


########################################################################


def MENU_avvia_IDE():
    '''
    Avvia la modifica di pyleeno.py con geany
    '''
    avvia_IDE()


def avvia_IDE():
    '''Avvia la modifica di pyleeno.py con geany o eric6'''
    basic_LeenO('file_gest.avvia_IDE')
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

    if sys.platform == 'linux' or sys.platform == 'darwin':

        dest = '/media/giuserpe/PRIVATO/LeenO/_SRC/leeno/src/Ultimus.oxt/python/pythonpath'
        if not os.path.exists(dest):
            try:
                dest = os.getenv(
                    "HOME") + '/' + src_oxt + '/leeno/src/Ultimus.oxt/'
                os.makedirs(dest)
                os.makedirs(os.getenv("HOME") + '/' + src_oxt + '/leeno/bin/')
                os.makedirs(os.getenv("HOME") + '/' + src_oxt + '/_SRC/OXT')
            except FileExistsError:
                pass

        subprocess.Popen('caja '+
                         # ~dest,
                         uno.fileUrlToSystemPath(LeenO_path()),
                         shell=True,
                         stdout=subprocess.PIPE)
        subprocess.Popen('geany ' + dest + '/pyleeno.py',
        # ~ subprocess.Popen('eric ' + dest + '/pyleeno.py',
                         shell=True,
                         stdout=subprocess.PIPE)
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
            dest = 'w:\\_dwg\\ULTIMUSFREE\\_SRC\\leeno\\src\\Ultimus.oxt'
        # ~subprocess.Popen('explorer.exe ' +
                        # dest,
                         # ~uno.fileUrlToSystemPath(LeenO_path()),
                         # ~shell=True,
                         # ~stdout=subprocess.PIPE)
        subprocess.Popen('"C:/Program Files/Geany/bin/geany.exe" ' +
                         dest +
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
    LeenoUtils.DocumentRefresh(False)
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
    # ~zoom = oDoc.CurrentController.ZoomValue
    # ~oDoc.CurrentController.ZoomValue = 400
    if n == 0:
        LeenoSheetUtils.inserSuperCapitolo(oSheet, lrow, sString)
    elif n == 1:
        LeenoSheetUtils.inserCapitolo(oSheet, lrow, sString)
    elif n == 2:
        LeenoSheetUtils.inserSottoCapitolo(oSheet, lrow, sString)

    _gotoCella(2, lrow)
    Rinumera_TUTTI_Capitoli2(oSheet)
    # ~oDoc.CurrentController.ZoomValue = zoom
    oDoc.enableAutomaticCalculation(True)
    oDoc.CurrentController.setFirstVisibleColumn(0)
    oDoc.CurrentController.setFirstVisibleRow(lrow - 5)
    # MsgBox('eseguita in ' + str((datetime.now() - datarif).total_seconds()) + ' secondi!','')
    LeenoUtils.DocumentRefresh(True)


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
    oSheet.getCellByPosition(
        18, 1).Formula = '=SUBTOTAL(9;S4:S' + str(lrow + 1) + ')'
    oSheet.getCellByPosition(
        18, lrow).Formula = '=SUBTOTAL(9;S4:S' + str(lrow + 1) + ')'

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


# ~def Filtra_Computo_Cap():
    # ~oDoc = LeenoUtils.getDocument()
    # ~oSheet = oDoc.CurrentController.ActiveSheet
    # ~nSheet = oSheet.getCellByPosition(7, 8).String
    # ~sString = oSheet.getCellByPosition(7, 10).String
    # ~Filtra_computo(nSheet, 31, sString)


########################################################################


# ~def Filtra_Computo_SottCap():
    # ~oDoc = LeenoUtils.getDocument()
    # ~oSheet = oDoc.CurrentController.ActiveSheet
    # ~nSheet = oSheet.getCellByPosition(7, 8).String
    # ~sString = oSheet.getCellByPosition(7, 12).String
    # ~Filtra_computo(nSheet, 32, sString)


########################################################################


# ~def Filtra_Computo_A():
    # ~oDoc = LeenoUtils.getDocument()
    # ~oSheet = oDoc.CurrentController.ActiveSheet
    # ~nSheet = oSheet.getCellByPosition(7, 8).String
    # ~sString = oSheet.getCellByPosition(7, 14).String
    # ~Filtra_computo(nSheet, 33, sString)


########################################################################


# ~def Filtra_Computo_B():
    # ~oDoc = LeenoUtils.getDocument()
    # ~oSheet = oDoc.CurrentController.ActiveSheet
    # ~nSheet = oSheet.getCellByPosition(7, 8).String
    # ~sString = oSheet.getCellByPosition(7, 16).String
    # ~Filtra_computo(nSheet, 34, sString)


########################################################################


# ~def Filtra_Computo_C():  # filtra in base al codice di prezzo
    # ~oDoc = LeenoUtils.getDocument()
    # ~oSheet = oDoc.CurrentController.ActiveSheet
    # ~nSheet = oSheet.getCellByPosition(7, 8).String
    # ~sString = oSheet.getCellByPosition(7, 20).String
    # ~Filtra_computo(nSheet, 1, sString)


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
    # ~LeenoSheetUtils.adattaAltezzaRiga(oSheet)
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

########################################################################
# adatta_altezza_riga moved to LeenoSheetUtils.py as adattaAltezzaRiga
########################################################################

def Menu_adattaAltezzaRiga():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)

def voce_breve():
    '''
    Cambia il numero di caratteri visualizzati per la descrizione voce in COMPUTO,
    CONTABILITA E VARIANTE.
    '''
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    # ~oSheet.getCellRangeByPosition(
        # ~0, 0,
        # ~SheetUtils.getUsedArea(oSheet).EndColumn,
        # ~SheetUtils.getUsedArea(oSheet).EndRow).Rows.OptimalHeight = True
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

    elif oSheet.Name == 'CONTABILITA':
        oSheet = oDoc.getSheets().getByName('S1')
        if oDoc.NamedRanges.hasByName("_Lib_1"):
            Dialogs.Exclamation(Title = 'ATTENZIONE!',
            Text='''Risulta già registrato un SAL, quindi
    NON E' POSSIBILE PROCEDERE.''')
            # ~DLG.MsgBox(
                # ~"Risulta già registrato un SAL. NON E' POSSIBILE PROCEDERE.",
                # ~'ATTENZIONE!')
            return
        else:
            if oSheet.getCellRangeByName('S1.H335').Value < 10000:
                cfg.write('Contabilita', 'cont_inizio_voci_abbreviate', oSheet.getCellRangeByName('S1.H335').String)
                oSheet.getCellRangeByName('S1.H335').Value = 10000
            else:
                oSheet.getCellRangeByName('S1.H335').Value = int(cfg.read('Contabilita', 'cont_inizio_voci_abbreviate'))
            if oSheet.getCellRangeByName('S1.H336').Value < 10000:
                cfg.write('Contabilita', 'cont_fine_voci_abbreviate', oSheet.getCellRangeByName('S1.H336').String)
                oSheet.getCellRangeByName('S1.H336').Value = 10000
            else:
                oSheet.getCellRangeByName('S1.H336').Value = int(cfg.read('Contabilita', 'cont_fine_voci_abbreviate'))
    Menu_adattaAltezzaRiga()


########################################################################


def cancella_voci_non_usate():
    '''
    Cancella le voci di prezzo non utilizzate.
    '''
    chiudi_dialoghi()


    if Dialogs.YesNoDialog(Title='AVVISO!',
    Text='''Questo comando ripulisce l'Elenco Prezzi
dalle voci non utilizzate in nessuno degli altri elaborati.

LA PROCEDURA POTREBBE RICHIEDERE DEL TEMPO.

Vuoi procedere comunque?''') == 0:
        return
    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(False)
    oSheet = oDoc.CurrentController.ActiveSheet

    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400

    oRange = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    SR = oRange.StartRow + 1
    ER = oRange.EndRow + 1
    lista_prezzi = list()
    for n in range(SR, ER):
        lista_prezzi.append(oSheet.getCellByPosition(0, n).String)
    lista = list()
    # attiva la progressbar
    progress = Dialogs.Progress(Title='Ricerca delle voci da eliminare in corso...', Text="Lettura dati")
    n = 0
    progress.setLimits(n, len(lista_prezzi))
    progress.show()
    progress.setValue(1)
    for tab in ('COMPUTO', 'Analisi di Prezzo', 'VARIANTE', 'CONTABILITA'):
        try:
            oSheet = oDoc.getSheets().getByName(tab)
            if tab == 'Analisi di Prezzo':
                col = 0
            else:
                col = 1
            for el in lista_prezzi:
                n += 1
                progress.setValue(n)
                if SheetUtils.uFindStringCol(el, col, oSheet):
                    lista.append(el)
        except Exception:
            pass
    progress.setLimits(0, 5)
    progress.setValue(2)
    da_cancellare = set(lista_prezzi).difference(set(lista))
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    progress.setValue(3)
    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, SR, 0, ER))
    struttura_off()
    struttura_off()
    struttura_off()
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")) #'unselect
    progress.setValue(4)

    # ~for el in da_cancellare:
        # ~oCellRangeAddr.StartRow = el
        # ~oCellRangeAddr.EndRow = el
        # ~oSheet.group(oCellRangeAddr, 1)
        # ~oSheet.getCellRangeByPosition(0, el, 0,
                                  # ~el).Rows.IsVisible = False

    for n in reversed(range(SR, ER)):
        if oSheet.getCellByPosition(0, n).String in da_cancellare:
            oSheet.Rows.removeByIndex(n, 1)
        if(oSheet.getCellByPosition(0, n).String == '' and
           oSheet.getCellByPosition(1, n).String == '' and
           oSheet.getCellByPosition(4, n).String == ''):
            oSheet.Rows.removeByIndex(n, 1)

    progress.setValue(5)
    oDoc.enableAutomaticCalculation(True)
    progress.hide()
    _gotoCella(0, 3)
    oDoc.CurrentController.ZoomValue = zoom
    Dialogs.Info(Title = 'Ricerca conclusa', Text='Eliminate ' + str(len(da_cancellare)) + " voci dall'elenco prezzi.")


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
        LeenoSheetUtils.adattaAltezzaRiga(oSheet)
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
    oDoc = LeenoUtils.getDocument()
    LeenoUtils.DocumentRefresh(False)
    oSheet = oDoc.CurrentController.ActiveSheet
    psm = LeenoUtils.getComponentContext().ServiceManager
    dp = psm.createInstance('com.sun.star.awt.DialogProvider')
    global oDialog1
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
            vista_terra_terra()
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
            #  LeenoSheetUtils.adattaAltezzaRiga(oSheet)
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
            #  elif oDialog1.getControl("CBDesc").State == 0: LeenoSheetUtils.adattaAltezzaRiga(oSheet)

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
            comando('Copy')
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
        #  elif oDialog1.getControl("CBDesc").State == 0: LeenoSheetUtils.adattaAltezzaRiga(oSheet)

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
        if oSheet.getColumns().getByIndex(19).Columns.IsVisible:
            oDialog1.getControl('vista_pre').State = 1
        else:
            oDialog1.getControl('vista_sem').State = 1
        sString = oDialog1.getControl('TextField3')
        sString.Text = oDoc.getSheets().getByName('S1').getCellRangeByName(
            'H335').Value  # cont_inizio_voci_abbreviate
        sString = oDialog1.getControl('TextField2')
        sString.Text = oDoc.getSheets().getByName('S1').getCellRangeByName(
            'H336').Value  # cont_fine_voci_abbreviate

        # Contabilità abilita
        # ~if oSheet.getCellRangeByName('S1.H328').Value == 1:
            # ~oDialog1.getControl('CheckBox7').State = 1
        sString = oDialog1.getControl('TextField13')
        if cfg.read('Contabilita', 'idxsal') == '&273.Dlg_config.TextField13.Text':
            sString.Text = '20'
        else:
            sString.Text = cfg.read('Contabilita', 'idxsal')
            if sString.Text == '':
                sString.Text = '20'
        sString = oDialog1.getControl('ComboBox3')
        sString.Text = cfg.read('Contabilita', 'ricicla_da')

        # oDialog1Model = oDialog1.Model
        oDialog1.getControl('Dettaglio').State = cfg.read('Generale', 'dettaglio')
        oDialog1.execute()

        if oDialog1.getControl('vista_pre').State:
            LeenoSheetUtils.setLarghezzaColonne(oSheet)
        if oDialog1.getControl('vista_sem').State:
            LeenoSheetUtils.setVisibilitaColonne(oSheet, 'TTTFFTTTTTFTFTFTFFFFFFFFFFFFFFFFFFFFF')
        # ~vista_terra_terra()

        # il salvataggio anche su leeno.conf serve alla funzione voce_breve()
        if oDialog1.getControl('TextField3').getText() != '10000':
            cfg.write('Contabilita', 'cont_inizio_voci_abbreviate', oDialog1.getControl('TextField3').getText())
        oDoc.getSheets().getByName('S1').getCellRangeByName('H335').Value = float(oDialog1.getControl('TextField3').getText())

        if oDialog1.getControl('TextField2').getText() != '10000':
            cfg.write('Contabilita', 'cont_fine_voci_abbreviate', oDialog1.getControl('TextField2').getText())
        oDoc.getSheets().getByName('S1').getCellRangeByName('H336').Value = float(oDialog1.getControl('TextField2').getText())

        cfg.write('Contabilita', 'idxsal', oDialog1.getControl('TextField13').getText())
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
    # LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    LeenoUtils.DocumentRefresh(True)
    # ~oDoc.enableAutomaticCalculation(True)
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
            oSheet = oDoc.getSheets().getByName('VARIANTE')
            LeenoSheetUtils.adattaAltezzaRiga(oSheet)
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
    oDoc.enableAutomaticCalculation(False)
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

    # attiva la progressbar
    progress = Dialogs.Progress(Title='Generazione dei sommari in corso...', Text="Lettura dati")
    progress.setLimits(0, LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2)
    progress.setValue(0)
    progress.show()

    for n in range(4, LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2):
        progress.setValue(n)
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

    oDoc.enableAutomaticCalculation(True)
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    progress.hide()

########################################################################


def MENU_riordina_ElencoPrezzi():
    '''
    Riordina l'Elenco Prezzi secondo l'ordine alfabetico dei codici di prezzo
    '''
    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(False)
    riordina_ElencoPrezzi(oDoc)
    oDoc.enableAutomaticCalculation(True)


def riordina_ElencoPrezzi(oDoc):
    '''
    Riordina l'Elenco Prezzi secondo l'ordine alfabetico dei codici di prezzo
    '''
    #chiudi_dialoghi()
    oDoc.enableAutomaticCalculation(False)
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    if SheetUtils.uFindStringCol('Fine elenco', 0, oSheet) is None:
        LeenoSheetUtils.inserisciRigaRossa(oSheet)
    test = str(SheetUtils.uFindStringCol('Fine elenco', 0, oSheet) +1)
    SheetUtils.NominaArea(oDoc, 'Elenco Prezzi', "$A$3:$AF$" + test, 'elenco_prezzi')
    SheetUtils.NominaArea(oDoc, 'Elenco Prezzi', "$A$3:$A$" + test, 'Lista')
    oRangeAddress = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    SR = oRangeAddress.StartRow + 1
    SC = 0 #oRangeAddress.StartColumn
    EC = oRangeAddress.EndColumn
    SR = oRangeAddress.StartRow + 1
    ER = oRangeAddress.EndRow -1
    if SR == ER:
        return

    oRange = oSheet.getCellRangeByPosition(SC, SR, EC, ER)
    SheetUtils.simpleSortColumn(oRange, 0, True)

    oDoc.enableAutomaticCalculation(True)


########################################################################


def MENU_doppioni():
    # ~EliminaVociDoppieElencoPrezzi()
    elimina_voci_doppie()


def EliminaVociDoppieElencoPrezzi():
    oDoc = LeenoUtils.getDocument()
    '''
    Cancella eventuali voci che si ripetono in Elenco Prezzi
    '''
    oDoc.enableAutomaticCalculation(False)
    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400
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
    # ~oDoc.CurrentController.select(
        # ~oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect
    if oDoc.getSheets().hasByName('Analisi di Prezzo'):
        tante_analisi_in_ep()

    oDoc.enableAutomaticCalculation(True)
    oDoc.CurrentController.ZoomValue = zoom
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    riordina_ElencoPrezzi(oDoc)
    if len(set(lista_tar)) != len(set(lista_come_array)):
        Dialogs.Exclamation(Title = 'ATTENZIONE!',
            Text='''Ci sono ancora 2 o più voci che hanno
lo stesso Codice Articolo pur essendo diverse.''')
        # ~DLG.MsgBox(
            # ~'Ci sono ancora 2 o più voci che hanno lo stesso Codice Articolo pur essendo diverse.',
            # ~'C o n t r o l l a!')


########################################################################
# Scrive un file.
def XPWE_out(elaborato, out_file):
    '''
    esporta il documento in formato XPWE

    elaborato { string } : nome del foglio da esportare
    out_file  { string } : nome base del file

    il nome file risulterà out_file-elaborato.xpwe
    '''
    LeenoUtils.DocumentRefresh(False)
    # attiva la progressbar
    progress = Dialogs.Progress(Title='Esportazione di ' + elaborato + ' in corso...', Text="Lettura dati")
    progress.setLimits(0, 7)
    progress.setValue(0)
    progress.show()

    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(False)
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
    progress.setValue(1)
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
    progress.setValue(2)
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
    progress.setValue(3)
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
    progress.setValue(4)
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
    progress.setValue(5)
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
    Rinumera_TUTTI_Capitoli2(oSheet)
    nVCItem = 2
    progress.setValue(6)
    progress.setLimits(0, LeenoSheetUtils.cercaUltimaVoce(oSheet))
    for n in range(0, LeenoSheetUtils.cercaUltimaVoce(oSheet)):
        progress.setValue(n)
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
                        # quando vedi_voce guarda ad un valore negativo
                        try:
                            if test:
                                Flags.text = '32768'
                        except:
                            pass
            n = sotto + 1
    # #########################
    # ~out_file = Dialogs.FileSelect('Salva con nome...', '*.xpwe', 1)
    # ~out_file = uno.fileUrlToSystemPath(oDoc.getURL())
    # ~DLG.mri (uno.fileUrlToSystemPath(oDoc.getURL()))
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
    progress.hide()
    try:
        of = codecs.open(out_file, 'w', 'utf-8')
        of.write(riga)
        # ~MsgBox('Esportazione in formato XPWE eseguita con successo\nsul file ' + out_file + '!','Avviso.')
    except Exception:
        Dialogs.Exclamation(Title = 'E R R O R E !',
            Text='''               Esportazione non eseguita!
Verifica che il file di destinazione non sia già in uso!''')
        # ~DLG.MsgBox(
            # ~'Esportazione non eseguita!\n\nVerifica che il file di destinazione non sia già in uso!',
            # ~'E R R O R E !')

    LeenoUtils.DocumentRefresh(True)


########################################################################
def MENU_firme_in_calce(lrowF=None):
    '''
    Inserisce(in COMPUTO o VARIANTE) un riepilogo delle categorie
    ed i dati necessari alle firme
    '''

    LeenoUtils.DocumentRefresh(False)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet_S2 = oDoc.getSheets().getByName('S2')

    datafirme = oSheet_S2.getCellRangeByName('$S2.C4').String

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
        # ~DLG.chi(datafirme)
        oSheet.getCellByPosition(2 , riga_corrente).Formula = (
            '=CONCATENATE("' + datafirme + '";TEXT(NOW();"GG/mm/aaaa"))')
        oSheet.getCellByPosition(2 , riga_corrente + 2).Formula = (
            "L'Impresa esecutrice\n(" + oSheet_S2.getCellByPosition(
                2, 16).String + ")")
        oSheet.getCellByPosition(2 , riga_corrente + 6).Formula = (
            "Il Direttore dei Lavori\n(" + oSheet_S2.getCellByPosition(
                2, 15).String + ")")
# ~rem CONSOLIDA LA DATA
        oRange = oSheet.getCellRangeByPosition (2, riga_corrente, 40, riga_corrente)
        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)
    if oSheet.Name in ("Registro", "SAL"):
        if lrowF == None:
            lrowF = SheetUtils.getLastUsedRow(oSheet)

        oSheet.getRows().insertByIndex(lrowF, 13)
        riga_corrente = lrowF + 1
        oSheet.getCellByPosition(1 , riga_corrente).Formula = '=CONCATENATE("' + datafirme + '";TEXT(NOW();"GG/mm/aaaa"))'
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
        '=CONCATENATE("In data ";TEXT(NOW();"DD/MM/YYYY");" è stato emesso il CERTIFICATO DI PAGAMENTO n.' + str(nSal) + ' per un importo di €")')
        oRange = oSheet.getCellRangeByPosition (1, riga_corrente + 10, 40, riga_corrente + 10)

        aSaveData = oRange.getDataArray()
        oRange.setDataArray(aSaveData)

        oSheet.getCellByPosition(1 , riga_corrente + 12).Formula = (
            "Il Direttore dei Lavori\n(" + oSheet_S2.getCellRangeByName(
                '$S2.C16').String + ")")
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
        oSheet.getCellRangeByPosition(0, lrowF, 100, lrowF + 15 -
                                      1).CellStyle = "Ultimus_centro"
        oSheet.getCellRangeByPosition(0, lrowF + 15 - 1, 100, lrowF + 15 -
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
        inizio_gruppo = riga_corrente
        riga_corrente += 1

    # attiva la progressbar
        progress = Dialogs.Progress(Title='Esecuzione in corso...', Text="Composizione del riepilogo strutturale.")
        i = 0
        progress.setLimits(0, LeenoSheetUtils.cercaUltimaVoce(oSheet))
        progress.setValue(i)
        progress.show()
        for i in range(0, lrowF):
            progress.setValue(i)

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
        progress.hide()
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

        #  oSheet.getCellByPosition(lrowF,0).Rows.IsManualPageBreak = True
    LeenoUtils.DocumentRefresh(True)


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
        pass
    oDoc.enableAutomaticCalculation(True)


########################################################################
def tante_analisi_in_ep():
    '''
    Trasferisce le analisi all'Elenco Prezzi.
    '''
    chiudi_dialoghi()

    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(False)
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

    oDoc.enableAutomaticCalculation(True)
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
            fini = list()
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
        # ~numera_voci(1)
    except Exception:
        pass
    _gotoCella(0, fine)
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
            progress = Dialogs.Progress(Title='Esecuzione in corso...', Text="Cancellazione voci azzerate")
            n = 0
            progress.setLimits(0, LeenoSheetUtils.cercaUltimaVoce(oSheet))
            progress.setValue(n)
            progress.show()
            for lrow in reversed(range(0, ER)):
                n += 1
                progress.setValue(n)
                # ~if oSheet.getCellByPosition(
                        # ~2, lrow).String == '*** VOCE AZZERATA ***':
                if '*** VOCE AZZERATA ***' in oSheet.getCellByPosition(2, lrow).String:
                    # ~elimina_voce(lrow=lrow, msg=0)
                    LeenoSheetUtils.eliminaVoce(oSheet, lrow)

            numera_voci(1)
            progress.hide()
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
            # ~DLG.MsgBox('La selezione deve essere contigua.', 'ATTENZIONE!')
            Dialogs.Exclamation(Title = 'ATTENZIONE!',
            Text='''La selezione deve essere contigua.''')
            return 0
        if lrow is not None:
            ER = oRangeAddress.EndRow
            ER = LeenoComputo.circoscriveVoceComputo(oSheet, ER).RangeAddress.EndRow
        else:
            ER = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow
        lista_y = [SR, ER]
    # ~if oSheet.Name == 'Analisi di Prezzo':
        # ~try:
            # ~oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
        # ~except AttributeError:
            # ~oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
        # ~try:
            # ~if lrow is not None:
                # ~SR = oRangeAddress.StartRow
                # ~SR = LeenoComputo.circoscriveVoceComputo(oSheet, SR).RangeAddress.StartRow
            # ~else:
                # ~SR = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
        # ~except AttributeError:
            # ~DLG.MsgBox('La selezione deve essere contigua.', 'ATTENZIONE!')
            # ~return 0
        # ~if lrow is not None:
            # ~ER = oRangeAddress.EndRow
            # ~ER = LeenoComputo.circoscriveVoceComputo(oSheet, ER).RangeAddress.EndRow
        # ~else:
            # ~ER = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow
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
            # ~DLG.MsgBox('La selezione deve essere contigua.', 'ATTENZIONE!')
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

    if oSheet.Name not in ('COMPUTO', 'CONTABILITA', 'VARIANTE'):
        return

    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
    el_y = list()
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
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    rigen = False
    for y in reversed(lista_y):
        if oSheet.getCellByPosition(2, y).CellStyle not in ('An-lavoraz-generica',
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
    if rigen == True:
        rigenera_parziali(False)
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
    oDoc.enableAutomaticCalculation(True)

########################################################################
def copia_riga_computo(lrow):
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
    # vado alla vecchia maniera ## copio il range di righe computo da S5 ##
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheetto = oDoc.getSheets().getByName('S5')
    oRangeAddress = oSheetto.getCellRangeByPosition(0, 24, 42, 24).getRangeAddress()
    #  lrow = LeggiPosizioneCorrente()[1]
    stile = oSheet.getCellByPosition(1, lrow).CellStyle
    if oSheet.getCellByPosition(1,
                                lrow + 1).CellStyle == 'comp sotto Bianche_R':
        return
    if stile in ('comp Art-EP_R', 'Data_bianca', 'Comp-Bianche in mezzo_R'):

        lrow = lrow + 1  # PER INSERIMENTO SOTTO RIGA CORRENTE

        oCellAddress = oSheet.getCellByPosition(0, lrow).getCellAddress()
        oSheet.getRows().insertByIndex(lrow, 1)
        oSheet.copyRange(oCellAddress, oRangeAddress)
        if stile in ('comp Art-EP_R'):
            oRangeAddress = oSheet.getCellByPosition(1, lrow +
                                                     1).getRangeAddress()
            oCellAddress = oSheet.getCellByPosition(1, lrow).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
            oSheet.getCellByPosition(1, lrow + 1).String = ""
            oSheet.getCellByPosition(1, lrow + 1
                                ).CellStyle = 'Comp-Bianche in mezzo_R'
        else:
            oSheet.getCellByPosition(1, lrow).CellStyle = 'Comp-Bianche in mezzo_R'
    _gotoCella(2, lrow)
    return


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
    Aggiunge riga di misurazione
    '''
    LeenoUtils.DocumentRefresh(False)
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
    LeenoUtils.DocumentRefresh(True)


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
        # ~DLG.chi(partenza[2])
        # ~DLG.chi(LeenoUtils.getGlobalVar('sblocca_computo'))
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

SCEGLIENDO SI' SARAI COSTRETTO A RIGENERARLI!""") == 0:
                    pass
                else:
                    LeenoUtils.setGlobalVar('sblocca_computo', 1)
        # ~DLG.chi(LeenoUtils.getGlobalVar('sblocca_computo'))


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
        if oDoc.NamedRanges.hasByName("_Lib_1"):
            if LeenoUtils.getGlobalVar('sblocca_computo') == 0:
                if DLG.DlgSiNo(
                        "Risulta già registrato un SAL. VUOI PROCEDERE COMUQUE?",
                        'ATTENZIONE!') == 3:
                    return
                if Dialogs.YesNoDialog(Title='ATTENZIONE!',
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
    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(False)
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
        GotoSheet(cfg.read('Contabilita', 'ricicla_da'))
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        lrow = LeggiPosizioneCorrente()[1]
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        sopra = sStRange.RangeAddress.StartRow + 2
        sotto = sStRange.RangeAddress.EndRow - 1

        oSrc = oSheet.getCellRangeByPosition(2, sopra, 8,
                                             sotto).getRangeAddress()
        oSheet.getCellByPosition(2, sopra - 1).CellBackColor = 13500076
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
        oDest.getCellByPosition(1, partenza[1]).String = oSheet.getCellByPosition(1, sopra - 1).String
        oDest.getCellByPosition(2, partenza[1]).CellBackColor = 13500076
        rigenera_voce(partenza[1])
        # ~rigenera_parziali(False)
        _gotoCella(2, partenza[1] + 1)
    oDoc.enableAutomaticCalculation(True)

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
    el_y = list()
    try:
        len(oRangeAddress)
        for el in oRangeAddress:
            el_y.append((el.StartRow, el.EndRow))
    except TypeError:
        el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))

    # estrate tutte le righe incluse nel o nei range(s)
    # e le inserisce in una lista di righe
    lista = list()
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
    oDoc = LeenoUtils.getDocument()
    try:
        oSheet = oDoc.CurrentController.ActiveSheet
    except Exception:
        return
    ER = SheetUtils.getUsedArea(oSheet).EndRow

    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400
    # attiva la progressbar
    progress = Dialogs.Progress(Title='Rigenerazione in corso...', Text="Lettura dati")
    progress.setLimits(0, LeenoSheetUtils.cercaUltimaVoce(oSheet))
    progress.setValue(0)
    progress.show()

    if bit == 1:
        for lrow in range(0, ER):
            progress.setValue(lrow)
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
                    stringa = ' ►' + stringa #+ ')'
                    if oSheet.getCellByPosition(2,
                                                lrow).Type.value != 'FORMULA':
                        oSheet.getCellByPosition(
                            2, lrow).String = oSheet.getCellByPosition(
                                2, lrow).String + stringa.replace('.', ',')
    else:
        for lrow in range(0, ER):
            progress.setValue(lrow)
            if ' ►' in oSheet.getCellByPosition(2, lrow).String:
                oSheet.getCellByPosition(
                    2, lrow).String = oSheet.getCellByPosition(
                        2, lrow).String.split(' ►')[0]

    oDoc.CurrentController.ZoomValue = zoom
    progress.hide()
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
    '''
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, ".uno:" + cmd, "", 0,
                                   list())


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


def richiesta_offerta():
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
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    pagestyle.RightPageHeaderContent = oHContent
    _gotoCella(0, 1)
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
        '"cad";"corpo";"dm";"dm²";"dm³";"kg";"lt";"m";"m²";"m³";"q";"t";""',
        titoloInput='Scegli...',
        msgInput='Unità di misura')
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
        oSheet.getCellByPosition(
            1, sopra + 1
        ).CellStyle = 'comp Art-EP_R'
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
                formula = (['=IF(PRODUCT(E' + str(n + 1) + ':I' +
                    str(n + 1) + ')=0;"";-PRODUCT(E' + str(n + 1) +
                    ':I' + str(n + 1) + '))'])
            else:
                formula = (['=IF(PRODUCT(E' + str(n + 1) + ':I' +
                    str(n + 1) + ')=0;"";PRODUCT(E' + str(n + 1) +
                    ':I' + str(n + 1) + '))'])
            if oSheet.getCellByPosition(4, n).Value < 0:
                formula = (['=IF(PRODUCT(E' + str(n + 1) + ':I' +
                    str(n + 1) + ')=0;"";PRODUCT(E' + str(n + 1) +
                    ':I' + str(n + 1) + '))'])
            formule.append(formula)

        oRange = oSheet.getCellRangeByPosition(9, sopra + 2, 9, sotto - 1)
        formule = tuple(formule)
        # ~oDoc.CurrentController.select(oRange)
        # ~DLG.chi(formule)
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

        oRange = oSheet.getCellRangeByPosition(9, sopra + 2, 11, sotto - 2)
        formule = tuple(formule)
        oRange.setFormulaArray(formule)
#    progress.hide()



########################################################################
def rigenera_tutte(arg=None, ):
    '''
    Ripristina le formule in tutto il foglio
    '''
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(False)

    riordina_ElencoPrezzi(oDoc)


    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400

    oSheet = oDoc.CurrentController.ActiveSheet
    nome = oSheet.Name
    stili_cat = LeenoUtils.getGlobalVar('stili_cat')


    # attiva la progressbar
    progress = Dialogs.Progress(Title='Rigenerazione di ' + nome + ' in corso...', Text="Lettura dati")
    progress.setLimits(0, LeenoSheetUtils.cercaUltimaVoce(oSheet))
    progress.setValue(0)
    progress.show()
    if nome in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        try:
            oSheet = oDoc.Sheets.getByName(nome)
            row = LeenoSheetUtils.prossimaVoce(oSheet, 0, 1, True)
            oDoc.CurrentController.select(oSheet.getCellByPosition(0, row))
            last = LeenoSheetUtils.cercaUltimaVoce(oSheet)
            while row < last:
                progress.setValue(row)
                rigenera_voce(row)
                # ~sistema_stili(row)
                row = LeenoSheetUtils.prossimaVoce(oSheet, row, 1, True)
        except Exception:
            pass
    rigenera_parziali(True)
    Rinumera_TUTTI_Capitoli2(oSheet)
    numera_voci()
    fissa()
    progress.hide()
    oDoc.enableAutomaticCalculation(True)
    comando("CalculateHard")
    oDoc.CurrentController.ZoomValue = zoom


########################################################################
def sistema_stili(lrow=None):
    '''
    Ripristina stili di cella per una singola voce.
    in COMPUTO, VARIANTE e CONTABILITA
    '''
    # ~lrow = LeggiPosizioneCorrente()[1]
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
            # ~oSheet.getCellByPosition(11, x).CellStyle = 'Blu ROSSO'
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
        for y in range (2, 8):
            rosso = 0
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
def rigenera_parziali (arg=False):
    '''
    arg { boolean }: Se False rigenera solo voce corrente
    Rigenera i parziali di tutte le voci
    '''
    oDoc = LeenoUtils.getDocument()
    # ~oDoc.enableAutomaticCalculation(False)
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
    progress = Dialogs.Progress(Title='Rigenerazione in corso...', Text="Parziali")
    progress.setLimits(0, sotto - sopra)
    n = 0
    progress.setValue(n)
    progress.show()
    if lrow == True:
        sopra = lrow 
    for i in range(sopra, sotto):
        n += 1
        progress.setValue(n)
        if 'Parziale [' in oSheet.getCellByPosition(8, i).Formula:
            parziale_core(oSheet, i)
    # ~oDoc.enableAutomaticCalculation(True)
    LeenoUtils.DocumentRefresh(True)
    progress.hide()
    return


########################################################################
def MENU_nuova_voce_scelta():  # assegnato a ctrl-shift-n
    '''
    Contestualizza in ogni tabella l'inserimento delle voci.
    '''
    oDoc = LeenoUtils.getDocument()
    LeenoUtils.DocumentRefresh(False)
    # ~oDoc.enableAutomaticCalculation(False)
    oSheet = oDoc.CurrentController.ActiveSheet
#    lrow = LeggiPosizioneCorrente()[1]

    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        LeenoComputo.ins_voce_computo()
    elif oSheet.Name == 'Analisi di Prezzo':
        inizializza_analisi()
    elif oSheet.Name == 'CONTABILITA':
        # ~LeenoContab.insertVoceContabilita(oSheet, lrow)  <<< non va
        ins_voce_contab()
    elif oSheet.Name == 'Elenco Prezzi':
        ins_voce_elenco()
    LeenoUtils.DocumentRefresh(True)
    # ~oDoc.enableAutomaticCalculation(True)

# nuova_voce_contab  ##################################################
def ins_voce_contab(lrow=0, arg=1, cod=None):
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoContab.insertVoceContabilita
    Inserisce una nuova voce in CONTABILITA.
    '''
    oDoc = LeenoUtils.getDocument()
    # ~oSheet = oDoc.CurrentController.ActiveSheet
    oSheet = oDoc.Sheets.getByName('CONTABILITA')

    stili_contab = LeenoUtils.getGlobalVar('stili_contab')
    stili_cat = LeenoUtils.getGlobalVar('stili_cat')

    if lrow == 0:
        lrow = LeggiPosizioneCorrente()[1]
        if oSheet.getCellByPosition(0, lrow + 1).CellStyle == 'uuuuu':
            return
        # ~else:
            # ~lrow += 1
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
    # ~if stile in stili_cat:
        # ~lrow += 1
        # ~stile = oSheet.getCellByPosition(0, lrow).CellStyle
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
    # ~ if stile == 'comp Int_colonna_R_prima':
        # ~ lrow += 1

    oSheetto = oDoc.getSheets().getByName('S5')
    oRangeAddress = oSheetto.getCellRangeByPosition(0, 22, 48, 26).getRangeAddress()
    oCellAddress = oSheet.getCellByPosition(0, lrow).getCellAddress()
    oSheet.getRows().insertByIndex(lrow, 5)  # inserisco le righe
    oSheet.copyRange(oCellAddress, oRangeAddress)
    # ~oSheet.getCellRangeByPosition(0, lrow, 48, lrow + 5).Rows.OptimalHeight = True
    _gotoCella(1, lrow + 1)

    #  if(oSheet.getCellByPosition(0,lrow).queryIntersection(oSheet.getCellRangeByName('_Lib_'+str(nSal)).getRangeAddress())):
    #  chi('appartiene')
    #  else:
    #  chi('nooooo')
    #  return

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


def inizializza_elenco():
    '''
    Riscrive le intestazioni di colonna e le formule dei totali in Elenco Prezzi.
    '''
    chiudi_dialoghi()

    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(False)
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
    LeenoSheetUtils.setLarghezzaColonne(oSheet)
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
    oSheet.getCellRangeByName('N2').Formula = '=SUBTOTAL(9;N:N)'
    oSheet.getCellRangeByName('R2').Formula = '=SUBTOTAL(9;R:R)'
    oSheet.getCellRangeByName('V2').Formula = '=SUBTOTAL(9;V:V)'
    oSheet.getCellRangeByName('Y2').Formula = '=SUBTOTAL(9;Y:Y)'
    oSheet.getCellRangeByName('Z2').Formula = '=SUBTOTAL(9;Z:Z)'
    #   riga di totale importo COMPUTO
    y -= 1
    oSheet.getCellByPosition(12, y).String = 'TOTALE'
    oSheet.getCellByPosition(13, y).Formula = '=SUBTOTAL(9;N3:N' + str(y) + ')'
    #  riga di totale importo CONTABILITA'
    oSheet.getCellByPosition(16, y).String = 'TOTALE'
    oSheet.getCellByPosition(17, y).Formula = '=SUBTOTAL(9;R3:R' + str(y) + ')'
    #  rem    riga di totale importo VARIANTE
    oSheet.getCellByPosition(20, y).String = 'TOTALE'
    oSheet.getCellByPosition(21, y).Formula = '=SUBTOTAL(9;V3:V' + str(y) + ')'
    #  rem    riga di totale importo PARALLELO
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

    oDoc.enableAutomaticCalculation(True)
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
    LeenoSheetUtils.setLarghezzaColonne(oSheet)
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
        LeenoSheetUtils.setLarghezzaColonne(oSheet)
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
    LeenoEvents.assegna()
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
    col4 = 12632319
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

        for n in (3, 7):
            oCellRangeAddr.StartColumn = n
            oCellRangeAddr.EndColumn = n
            oSheet.group(oCellRangeAddr, 0)
            oSheet.getCellRangeByPosition(n, 0, n, 0).Columns.IsVisible = False
    test = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2
    # attiva la progressbar
    progress = Dialogs.Progress(Title='Rigenerazione in corso...', Text="Lettura dati")
    progress.setLimits(0, test)
    progress.setValue(0)
    progress.show()
    lista = list()
    x = 0
    for n in range(0, test):
        x += 1
        progress.setValue(x)
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
    progress.hide()
    return


########################################################################
def struttura_Elenco():
    '''
    Dà una tonalità di colore, diverso dal colore dello stile cella, alle righe
    che non hanno il prezzo, come i titoli di capitolo e sottocapitolo.
    '''
    chiudi_dialoghi()

    if Dialogs.YesNoDialog(Title='AVVISO!',
    Text='''Adesso puoi dare ai titoli di capitolo e sottocapitolo
una tonalità di colore che ne facilita la leggibilità, ma
il risultato finale dipende dalla struttura dei codici di voce.

L'OPERAZIONE POTREBBE RICHIEDERE DEL TEMPO E
LibreOffice POTREBBE SEMBRARE BLOCCATO!

Vuoi procedere comunque?''') == 0:
        return

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    struct_colore(0)  # attribuisce i colori
    struct_colore(1)
    struct_colore(2)
    struct_colore(3)
    return


########################################################################
# ns_ins moved to LeenoImport_XmlToscana.py
########################################################################

########################################################################
# XML_toscana_import moved to LeenoImport_XmlToscana.py
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
    oDoc.CurrentController.ZoomValue = zoom


########################################################################
def MENU_importa_stili():
    '''
    Importa tutti gli stili da un documento di riferimento. Se non è
    selezionato, il file di riferimento è il template di leenO.
    '''

    if Dialogs.YesNoDialog(Title='Importa Stili in blocco?',
    Text='''Questa operazione sovrascriverà gli stili
del documento attivo, se già presenti!

Se non scegli un file di riferimento, saranno
importati gli stili di default di LeenO.

Vuoi continuare?''') == 0:
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
        oSheet = oDoc.getSheets().getByName(el)
        LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    try:
        GotoSheet('Elenco Prezzi')
    except Exception:
        pass


########################################################################
def MENU_parziale():
    '''
    Inserisce una riga con l'indicazione della somma parziale.
    '''
    LeenoUtils.DocumentRefresh(False)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]
    if oSheet.getCellByPosition(1, lrow-1).CellStyle in ('comp Art-EP_R') or \
        lrow == 0:
        return
    if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        parziale_core(oSheet, lrow)
        rigenera_parziali(False)
    LeenoUtils.DocumentRefresh(True)


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
        # ~if oSheet.Name == 'CONTABILITA':
            # ~oSheet.getCellByPosition(11, lrow).Formula = (
                # ~'=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(lrow + 1) +
                # ~')>=0;"";PRODUCT(E' + str(lrow + 1) + ':I' +
                # ~str(lrow + 1) + ')*-1)')
            # ~oSheet.getCellByPosition(9, lrow).String = ''
        # ~if oSheet.Name in ('COMPUTO', 'VARIANTE'):
            # ~oSheet.getCellByPosition(9, lrow).Formula = (
                # ~'=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(lrow + 1) +
                # ~')=0;"";PRODUCT(E' + str(lrow + 1) + ':I' + str(lrow + 1) + '))')
        for x in range(2, 12):
            oSheet.getCellByPosition(x, lrow).CellStyle = (
            oSheet.getCellByPosition(x, lrow).CellStyle + ' ROSSO')
        return '-'

########################################################################
def MENU_vedi_voce():
    '''
    Inserisce un riferimento a voce precedente sulla riga corrente.
    '''
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
            to = int(to.split('$')[-1]) - 1
        except ValueError:
            return
        _gotoCella(2, lrow)
        # focus = oDoc.CurrentController.getFirstVisibleRow
        if to < lrow:
            vedi_voce_xpwe(oSheet, lrow, to)


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
    oDoc = LeenoUtils.getDocument()
    # ~oDoc.enableAutomaticCalculation(False)
    LeenoUtils.DocumentRefresh(False)

    oSheet = oDoc.CurrentController.ActiveSheet

    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_contab = LeenoUtils.getGlobalVar('stili_contab')

    if oSheet.Name == "Elenco Prezzi":
        oCell = oSheet.getCellRangeByName('C2')
        voce = oDoc.Sheets.getByName('Elenco Prezzi').getCellByPosition(
            0, LeggiPosizioneCorrente()[1]).String

        # colora la descrizione scelta
        # ~oDoc.Sheets.getByName('Elenco Prezzi').getCellByPosition(
            # ~1, LeggiPosizioneCorrente()[1]).CellBackColor = 16777120

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
        # ~DLG.MsgBox('Devi prima selezionare una voce di misurazione.', 'Avviso!')
        Dialogs.Exclamation(Title = 'ATTENZIONE!',
        Text='''Devi prima selezionare una voce di misurazione.''')
        return
    fine = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1

    progress = Dialogs.Progress(Title='Applicazione filtro in corso...', Text="Lettura dati")
    progress.setLimits(0, fine)
    progress.setValue(0)
    progress.show()

    qui = None
    lista_pt = list()
    _gotoCella(0, 0)
    for n in range(0, fine):
        progress.setValue(n)
        if oSheet.getCellByPosition(0,
                                    n).CellStyle in ('Comp Start Attributo',
                                                     'Comp Start Attributo_R'):
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, n)
            sopra = sStRange.RangeAddress.StartRow
            sotto = sStRange.RangeAddress.EndRow
            if oSheet.getCellByPosition(1, sopra + 1).String != voce:
                lista_pt.append((sopra, sotto))
            else:

                # ~# colora lo sfondo della voce filtrata
                # ~oSheet.getCellRangeByPosition(0, sopra, 40, sotto).CellBackColor = 16777120

                if qui == None:
                    qui = sopra + 1
    progress.setValue(fine)
    for el in lista_pt:
        oCellRangeAddr.StartRow = el[0]
        oCellRangeAddr.EndRow = el[1]
        oSheet.group(oCellRangeAddr, 1)
        oSheet.getCellRangeByPosition(0, el[0], 0,
                                      el[1]).Rows.IsVisible = False
    try:
        _gotoCella(0, qui)
    except:
        struttura_off()
        progress.hide()
        GotoSheet("Elenco Prezzi")
        Dialogs.Exclamation(Title = 'Ricerca conclusa', Text='Nessuna corrispondenza trovata')
    # ~oDoc.enableAutomaticCalculation(True)
    LeenoUtils.DocumentRefresh(True)
    progress.hide()

########################################################################

def MENU_struttura_on():
    oDoc = LeenoUtils.getDocument()
    LeenoUtils.DocumentRefresh(False)
    oSheet = oDoc.CurrentController.ActiveSheet

    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        struttura_ComputoM()
    elif oSheet.Name == 'Elenco Prezzi':
        struttura_Elenco()
    elif oSheet.Name == 'Analisi di Prezzo':
        struttura_Analisi()
    elif oSheet.Name in ('CONTABILITA', 'Registro', 'SAL'):
        LeenoContab.struttura_CONTAB()
    LeenoUtils.DocumentRefresh(True)

def struttura_ComputoM():
    '''
    @@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    # ~# attiva la progressbar
    progress = Dialogs.Progress(Title='Creazione vista struttura in corso...', Text="Lettura dati")
    progress.setLimits(0, 5)
    progress.setValue(0)
    progress.show()
    Rinumera_TUTTI_Capitoli2(oSheet)
    progress.setValue(1)
    struct(0)
    progress.setValue(2)
    struct(1)
    progress.setValue(3)
    struct(2)
    progress.setValue(4)
    struct(3)
    progress.setValue(5)
    progress.hide()


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
    x = LeggiPosizioneCorrente()[0]
    lrow = LeggiPosizioneCorrente()[1]
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.clearOutline()
    oDoc.CurrentController.setFirstVisibleColumn(0)
    oDoc.CurrentController.setFirstVisibleRow(lrow - 4)
    _gotoCella( x, lrow)


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
    # ~global utsave
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


########################################################################
def autoexec():
    '''
    questa è richiamata da creaComputo()
    '''
    LeenoEvents.pulisci()
    inizializza()
    LeenoEvents.assegna()
    SheetUtils.FixNamedArea()
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
    LeenoUtils.DocumentRefresh(False)
    #  RegularExpressions Wildcards are mutually exclusive, only one can have the value TRUE.
    #  If both are set to TRUE via API calls then the last one set takes precedence.
    try:
        oDoc.Wildcards = False
    except Exception:
        pass
    oDoc.RegularExpressions = False
    oDoc.CalcAsShown = True  # precisione come mostrato
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
    LeenoUtils.DocumentRefresh(True)
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
def vista_terra_terra():
    '''
    Settaggio base di configurazione colonne in COMPUTO e VARIANTE
    '''
    oDoc = LeenoUtils.getDocument()
    LeenoUtils.DocumentRefresh(False)

    oSheet = oDoc.CurrentController.ActiveSheet

    # raggruppo le colonne
    ncol = oSheet.getColumns().getCount()
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        col = 46
    elif oSheet.Name == 'CONTABILITA':
        col = 39
    oCellRangeAddr.StartColumn = col
    oCellRangeAddr.EndColumn = ncol
    oSheet.ungroup(oCellRangeAddr, 0)
    oSheet.group(oCellRangeAddr, 0)
    
    oSheet.getCellRangeByPosition(col, 0, ncol -1, 0).Columns.IsVisible = False
    LeenoSheetUtils.setLarghezzaColonne(oSheet)
    LeenoUtils.DocumentRefresh(True)


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
        unione = oDoc.Sheets.insertByName('stili', sheet)
        oSheet = oDoc.Sheets.getByName("stili")
    GotoSheet("stili")
    # attiva la progressbar
    progress = Dialogs.Progress(Title='Stili cella', Text="Scrittura in corso...")
    progress.setLimits(0, len(sty))
    progress.show()
    i = 0
    sty = sorted(sty)
    for el in sty:
        oSheet.getCellByPosition( 0, i).String = el
        oSheet.getCellByPosition( 1, i).CellStyle = el
        oSheet.getCellByPosition( 3, i).CellStyle = el
        oSheet.getCellByPosition( 1, i).Value = 2000.00
        oSheet.getCellByPosition( 3, i).String = "LeenO"
        i += 2
        progress.setValue(i)
    progress.hide()


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
    # ~if oCell.String == '';
    valida_cella(oCell,
                 '"ANALISI DI PREZZO";"ELENCO PREZZI";"ELENCO PREZZI E COSTI ELEMENTARI";\
                 "COMPUTO METRICO";"PERIZIA DI VARIANTE";"LIBRETTO DELLE MISURE";"REGISTRO DI CONTABILITÀ";\
                 "S.A.L. A TUTTO IL"',
                 titoloInput='Scegli...',
                 msgInput='Titolo della copertina...',
                 err=False)
    # ~oCell.String = ""
    # Indica qual è il Documento Principale
    ScriviNomeDocumentoPrincipale()
    # ~nascondi_sheets()


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

    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(False)
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
            Dialogs.Info(Title = 'Avviso!',
                         Text='''Non avendo effettuato l'adeguamento del file alla versione di LeenO installata, potresti avere dei malfunzionamenti!''')

            return
        sproteggi_sheet_TUTTE()
        if oDoc.getSheets().hasByName('S4'):
            oDoc.Sheets.removeByName('S4')
        zoom = oDoc.CurrentController.ZoomValue
        oDoc.CurrentController.ZoomValue = 400
        # attiva la progressbar
        progress = Dialogs.Progress(Title='Adeguamento del lavoro in corso...',
                                    Text="Lettura dati")
        progress.setLimits(0, 10)
        progress.setValue(0)
        progress.show()

        ############
        # aggiungi stili di cella
        progress.setValue(1)
        for el in ('comp 1-a PU', 'comp 1-a LUNG', 'comp 1-a LARG',
                   'comp 1-a peso', 'comp 1-a', 'Blu',
                   'Comp-Variante num sotto'):
            oStileCella = oDoc.createInstance("com.sun.star.style.CellStyle")
            if not oDoc.StyleFamilies.getByName('CellStyles').hasByName(el):
                oDoc.StyleFamilies.getByName('CellStyles').insertByName(
                    el, oStileCella)
                oStileCella.ParentStyle = 'comp 1-a'
        progress.setValue(2)
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
        progress.setValue(3)
        oSheet = oDoc.getSheets().getByName('S1')
        oSheet.getCellRangeByName('S1.H291').Value = \
            oDoc.getDocumentProperties().getUserDefinedProperties().Versione = adegua_a
        for el in oDoc.Sheets.ElementNames:
            oDoc.getSheets().getByName(el).IsVisible = True
            oDoc.CurrentController.setActiveSheet(oDoc.getSheets().getByName(el))
            oDoc.getSheets().getByName(el).IsVisible = False
        # dal template 212
        flags = VALUE + DATETIME + STRING + ANNOTATION + FORMULA + OBJECTS + EDITATTR  # FORMATTED + HARDATTR
        progress.setValue(4)
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
        progress.setValue(5)
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
        progress.setValue(6)
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
        progress.setValue(7)
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
            # ~rigenera_tutte() affido la rigenerazione delle formule al menu Viste
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
        oDoc.CurrentController.ZoomValue = zoom
#        oDialogo_attesa.endExecute()  # chiude il dialogo
        progress.setValue(8)
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
        progress.hide()
        Dialogs.Info(Title = 'Avviso', Text='Adeguamento del file completato con successo.')
    oDoc.enableAutomaticCalculation(True)


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
    if out_file == '':
        return
    testo = '\n'
    for el in lista:
        XPWE_out(el, out_file)
        testo = testo + '● ' + out_file + '-' + el + '.xpwe\n\n'
    # ~DLG.MsgBox('Esportazione in formato XPWE eseguita con successo su:\n' + testo, 'Avviso.')
    Dialogs.Info(Title = 'Avviso.',
    Text='Esportazione in formato XPWE eseguita con successo su:\n' + testo)


########################################################################
def chiudi_dialoghi(event=None):
    '''
    @@ DA DOCUMENTARE
    '''
    try:
        oDialog1.endExecute()
    except:
        pass
    # ~return
    # ~if event:
        # ~event.Source.Context.endExecute()
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
    # ~DLG.chi(list(formulas.values()))
    # ~return

    lista = list(formulas.values())

    # ~oDlgPDF.getControl("ComboBox1").Text = "Prova"
    oDlgPDF.getControl("ComboBox1").addItems(lista, 1)

    oDlgPDF.execute()


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
        creaComputo()
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
    sString.Text = version_code.read()#[:-9]
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
    LeenoEvents.assegna()
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
            else:
                oSheet.getCellRangeByName(
                    "A1:AT1").clearContents(HARDATTR)
                oSheet.getCellRangeByName(
                    d[el]).String = 'DP:'

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
    description_upd() # aggiorna description.xml - da disattivare prima del rilascio
    if bar == 0:
        oDoc = LeenoUtils.getDocument()
        Toolbars.AllOff()
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
    # ~shutil.copyfile(nomeZip2, nomeZip)


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
        progress = Dialogs.Progress(Title='Esportazione di ' + oSheet.Name + ' in corso...', Text="Lettura dati")
        progress.setLimits(0, SheetUtils.getUsedArea(oSheet).EndRow)
        progress.setValue(0)

        if oSheet.Name not in ('COMPUTO', 'VARIANTE'):
            return
        descrizione = InputBox(t='inserisci una descrizione per la nuova riga')
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
    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400

    LeenoUtils.DocumentRefresh(False)

    oSheet = oDoc.CurrentController.ActiveSheet
    lcol = LeggiPosizioneCorrente()[0]
    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
    el_y = list()
    lista_y = list()
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
        oDoc.CurrentController.select(oSheet.getCellByPosition(lcol, y))
        if oDoc.getCurrentSelection().Type.value == 'TEXT':
            testo = oDoc.getCurrentSelection().String.replace(
                '\t', ' ').replace('Ã¨', 'è').replace(
                'Â°', '°').replace('Ã', 'à').replace(
                ' $', '')
            while '  ' in testo:
                testo = testo.replace('  ', ' ')
            while '\n\n' in testo:
                testo = testo.replace('\n\n', '\n')
            oDoc.getCurrentSelection().String = testo.strip().strip().strip()
    LeenoUtils.DocumentRefresh(True)
    oDoc.CurrentController.ZoomValue = zoom


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
        if oDoc.NamedRanges.hasByName("_Lib_1"):
            Dialogs.Exclamation(Title = 'ATTENZIONE!',
            Text="Risulta già registrato un SAL. NON E' POSSIBILE PROCEDERE.")
            # ~DLG.MsgBox(
                # ~"Risulta già registrato un SAL. NON E' POSSIBILE PROCEDERE.",
                # ~'ATTENZIONE!')
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
def MENU_numera_colonna():
    '''
    Comando di menu per numera_colonna()
    '''
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
            oSheet.getCellByPosition(x, lrow).Value = larg
            # ~oSheet.getCellByPosition(x, lrow).Formula = '=CELL("col")-1'
            oSheet.getCellByPosition(x, lrow).HoriJustify = 'CENTER'
        elif oSheet.getCellByPosition(x, lrow).Formula == '=CELL("col")-1':
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
    corrente del foglio cP_Cop
    '''
    oDoc = LeenoUtils.getDocument()
    nome = oDoc.CurrentController.ActiveSheet.Name
    lista_fogli = oDoc.Sheets.ElementNames
    for el in lista_fogli:
        if el not in (nome, 'cP_Cop'):
            oSheet = oDoc.getSheets().getByName(el)
            iSheet = oSheet.RangeAddress.Sheet
            oStampa = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
            oStampa.Sheet = iSheet
            oSheet.setPrintAreas(())
    return


########################################################################


def set_area_stampa():
    ''' Imposta area di stampa il relazione all'elaborato da produrre'''
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
    if oSheet.Name in ('Analisi di Prezzo'):
        EC = 7
        SR = 1
        ER -= 1
        oSheet.setPrintTitleRows(False)
# imposta area di stampa
    oStampa = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oStampa.Sheet = iSheet
    oStampa.StartColumn = 0
    oStampa.StartRow = SR
    oStampa.EndColumn = EC
    oStampa.EndRow = ER
    oSheet.setPrintAreas((oStampa,))
    return


########################################################################


def MENU_sistema_pagine():
    '''
    Configura intestazioni e pie' di pagina degli stili di stampa
    e propone un'anteprima di stampa
    '''
    oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('M1'):
        return

    set_area_stampa()

    if Dialogs.YesNoDialog(Title='AVVISO!',
    Text='''Vuoi attribuire il colore bianco allo sfondo delle celle?
Le formattazioni dirette impostate durante il lavoro andranno perse.

Per ripristinare i colori, tipici dei fogli di LeenO, basterà selezionare
le celle ed usare "CTRL+M".

Procedo cambiando i colori?''') == 1:
        LeenoSheetUtils.SbiancaCellePrintArea()

    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.removeAllManualPageBreaks()
    SheetUtils.visualizza_PageBreak()

    #  committente = oDoc.NamedRanges.Super_ego_8.ReferredCells.String
    oggetto = oDoc.getSheets().getByName('S2').getCellRangeByName("C3").String + '\n\n'
    committente = "\nCommittente: " + oDoc.getSheets().getByName('S2').getCellRangeByName("C6").String
    luogo = '\n' + oSheet.Name
    if oSheet.Name == 'COMPUTO' and oSheet.getColumns().getByName("AD").Columns.Width > 10:
        luogo = luogo + ' - Incidenza MdO'
    # ~luogo = '\nLocalità: ' + oDoc.getSheets().getByName('S2').getCellRangeByName("C4").String

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
        # ~'Analisi di Prezzo': 'PageStyle_Analisi di Prezzo',
        'Analisi di Prezzo': 'PageStyle_COMPUTO_A4',
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
        # ~chi((n , oAktPage.DisplayName))
        oAktPage.HeaderIsOn = True
        oAktPage.FooterIsOn = True
        
        oAktPage.TopMargin = 1500
        oAktPage.BottomMargin = 800
        oAktPage.LeftMargin = 1500
        oAktPage.RightMargin = 1000
        
        oAktPage.FooterLeftMargin = 0
        oAktPage.FooterRightMargin = 0
        oAktPage.HeaderLeftMargin = 0
        oAktPage.HeaderRightMargin = 0

        oAktPage.HeaderBodyDistance = 0
        oAktPage.FooterBodyDistance = 0

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
            htxt = 8.0 / 100 * oAktPage.PageScale
            if oSheet.Name == 'Analisi di Prezzo':
                htxt = 9.0 / 100 * oAktPage.PageScale
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

            bordo = oAktPage.RightBorder
            bordo.LineWidth = 0
            bordo.OuterLineWidth = 0
            oAktPage.RightBorder = bordo

            bordo = oAktPage.LeftBorder
            bordo.LineWidth = 0
            bordo.OuterLineWidth = 0
            oAktPage.LeftBorder = bordo
            # Adatto lo zoom alla larghezza pagina
            oAktPage.PageScale = 0
            oAktPage.ScaleToPagesX = 1
            oAktPage.ScaleToPagesY = 0

            # ~HEADER
            oHeader = oAktPage.RightPageHeaderContent
            # ~oAktPage.PageScale = 95
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
            # ~FOOTER
            oFooter = oAktPage.RightPageFooterContent
            oHLText = oFooter.CenterText.Text.String = ''
            nomefile = oDoc.getURL().replace('%20',' ')
            oHLText = oFooter.LeftText.Text.String = "\nrealizzato con LeenO\n" + os.path.basename(nomefile)
            oHLText = oFooter.LeftText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            oHLText = oFooter.LeftText.Text.Text.CharHeight = htxt * 0.70
            oHLText = oFooter.RightText.Text.Text.CharFontName = 'Liberation Sans Narrow'
            oHLText = oFooter.RightText.Text.Text.CharHeight = htxt
            # ~oHLText = oFooter.RightText.Text.String = '#/##'
            oAktPage.RightPageFooterContent = oFooter

        if oAktPage.DisplayName == 'Page_Style_Libretto_Misure2':
            # ~HEADER
            oHeader = oAktPage.RightPageHeaderContent
            # oHLText = oHeader.LeftText.Text.String = committente + '\nLibretto delle misure n.'
            # oHRText = oHeader.RightText.Text.String = luogo
            oAktPage.RightPageHeaderContent = oHeader
            # ~FOOTER
            oFooter = oAktPage.RightPageFooterContent
            # oHLText = oFooter.CenterText.Text.String = "L'IMPRESA                    IL DIRETTORE DEI LAVORI"
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
    LeenoAnalysis.MENU_impagina_analisi()
    last = SheetUtils.getUsedArea(oSheet).EndRow
    # ~oSheet.getCellRangeByPosition(1, 0, 41, last).Rows.OptimalHeight = True
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
            'Elenco Prezzi'
    ):
        oDoc.CurrentController.freezeAtPosition(0, 3)
    elif oSheet.Name in ('CONTABILITA'):
        oDoc.CurrentController.freezeAtPosition(0, 3)
    elif oSheet.Name in ('Analisi di Prezzo'):
        oDoc.CurrentController.freezeAtPosition(0, 2)
    elif oSheet.Name in ('Registro', 'SAL'):
        oDoc.CurrentController.freezeAtPosition(0, 1)
    # ~_gotoCella(lcol, lrow)


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
        # ~DLG.MsgBox('Non ci sono voci di prezzo ricorrenti.', 'Informazione')
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



########################################################################
def trova_np():
    '''
    Raggruppa le righe in modo da rendere evidenti i nuovi prezzi
    '''
    chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oDoc.enableAutomaticCalculation(False)

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

    oDoc.enableAutomaticCalculation(True)


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


########################################################################
def elimina_voci_doppie():
    '''
    @@ DA DOCUMENTARE
    '''
    # elimina voci doppie hard - grezza e lenta, ma efficace
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    riordina_ElencoPrezzi(oDoc)
    fine = SheetUtils.getUsedArea(oSheet).EndRow + 1

    # attiva la progressbar
    progress = Dialogs.Progress(Title='Ricerca delle voci da eliminare in corso...', Text="Lettura dati")
    progress.setLimits(0, fine)
    progress.setValue(0)
    progress.show()

    oSheet.getCellByPosition(30, 3).Formula = '=IF(A4=A3;1;0)'
    oDoc.CurrentController.select(oSheet.getCellByPosition(30, 3))
    comando('Copy')
    oDoc.CurrentController.select(
        oSheet.getCellRangeByPosition(30, 3, 30, fine))
    paste_clip(insCells=1)

    for i in reversed(range(0, fine)):
        progress.setValue(i)
        if oSheet.getCellByPosition(30, i).Value == 1:
            _gotoCella(30, i)
            oSheet.getRows().removeByIndex(i, 1)
    oSheet.getCellRangeByPosition(30, 3, 30, fine).clearContents(FORMULA)
    _gotoCella(0, 3)
    progress.hide()


########################################################################
def MENU_hl():
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
                    1, el).String + '";">>>")'
                oSheet.getCellByPosition(1, el).Formula = stringa
        except Exception:
            pass


########################################################################
def MENU_filtro_descrizione():
    '''
    Raggruppa e nasconde tutte le voci di misura in cui non compare
    la stringa cercata.
    '''
    LeenoUtils.DocumentRefresh(False)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = iSheet
    fine = SheetUtils.getUsedArea(oSheet).EndRow + 1
    el_y = list()
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
        return

    struttura_off()
    oSheet.getCellRangeByPosition(2, 0, 2, 1048575).clearContents(HARDATTR)

    y = 4
    progress = Dialogs.Progress(Title='Applicazione filtro in corso...', Text="Lettura dati")
    progress.setLimits(0, fine)
    progress.setValue(0)
    progress.show()
    lRow = SheetUtils.sStrColtoList(descrizione, 2, oSheet, y)
    if len(lRow) == 0:
        progress.hide()
        # ~DLG.MsgBox('''Testo non trovato.''', 'ATTENZIONE!')
        Dialogs.Exclamation (Title = 'ATTENZIONE!',
        Text="Testo non trovato.")

        return
    el_y = list()
    for y in lRow:
        progress.setValue(y)
        oSheet.getCellByPosition(2, y).CellBackColor = 15757935
        el_y.append(seleziona_voce(y))
    lista_y = list()
    lista_y.append(2)
    for el in el_y:
        y = el[0]
        progress.setValue(y)
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
    progress.hide()
    LeenoUtils.DocumentRefresh(True)

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
    # #  DLG.mri(oDoc)
    # oSheet = oDoc.CurrentController.ActiveSheet
    # DLG.chi(oDoc.getCurrentSelection().CellBackColor)
    # # ~return
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

# ~from collections import OrderedDict
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
    el_y = list()
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
    somma = list()
    for y in lista_y:
        somma.append(oSheet.getCellByPosition(lcol, y).Value)
    DLG.chi(sum(somma))


########################################################################

def calendario():
    #calendario
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    x = LeggiPosizioneCorrente()[0]
    y = LeggiPosizioneCorrente()[1]
    testo = Dialogs.pickDate()
    lst = str(testo).split('-')
    try:
        testo = lst[2] + '/' + lst[1] + '/' + lst[0]
        oSheet.getCellByPosition(x, y).String = testo
    except:
        pass
    return

########################################################################
import itertools
import operator
import functools
import LeenoImport as LI

def MENU_debug():
    # ~import LeenoPdf
    # ~LeenoPdf.MENU_Pdf()
    # ~sistema_cose()
    # ~MENU_nasconde_voci_azzerate()
    # ~oDoc = LeenoUtils.getDocument()
    # ~oSheet = oDoc.CurrentController.ActiveSheet
    # ~lrow = LeggiPosizioneCorrente()[1]

    # ~raggruppa_righe_voce(lrow, 1)
    return
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lr = SheetUtils.getLastUsedRow(oSheet) + 1
    for el in reversed(range (1, lr)):
        if oSheet.getCellByPosition(2, el).CellStyle == 'comp 1-a' and \
            "'" in oSheet.getCellByPosition(2, el).Formula:
            ff = oSheet.getCellByPosition(2, el).Formula.split("'")
            oSheet.getCellByPosition(2, el).Formula = ff[0] + ff[-1][1:]

    return
    # ~LeenoSheetUtils.setAdatta()
    # ~sistema_cose()
    # ~return
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    usedArea = SheetUtils.getUsedArea(oSheet)
    # ~oSheet.getCellRangeByPosition(0, 0, usedArea.EndColumn, usedArea.EndRow).Rows.OptimalHeight = False
    oSheet.getCellRangeByPosition(0, 0, 1023, 1048575).Rows.OptimalHeight = False
    oSheet.getCellRangeByPosition(0, 0, 1023, 1048575).Rows.Height = 1576
    DLG.mri(oSheet.getCellRangeByPosition(0, 0, usedArea.EndColumn, usedArea.EndRow).Rows)
    return
    lr = SheetUtils.getLastUsedRow(oSheet) + 1
    for el in reversed(range (1, lr)):
        if oSheet.getCellByPosition(2, el).CellStyle == 'comp 1-a' and \
            oSheet.getCellByPosition(2, el).String == '' and \
            oSheet.getCellByPosition(9, el).String == '':
            oSheet.getRows().removeByIndex(el, 1)
        elif oSheet.getCellByPosition(2, el).Type.value == 'TEXT':
            oSheet.getCellByPosition(2, el).String = '- ' + oSheet.getCellByPosition(2, el).String
    return

def MENU_debug():

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheets = oDoc.Sheets.ElementNames
    set_area_stampa()
    orig = oDoc.getURL()

    dest = '.'.join(os.path.basename(orig).split('.')[0:-1]) + '.pdf'
    orig = uno.fileUrlToSystemPath(orig)
    dir_bak = os.path.dirname(oDoc.getURL())
    # ~DelPrintArea()
    oDoc.storeToURL(dir_bak + '/' + dest, list())

    # ~DLG.chi(dir_bak)
    return


def MENU_debug():
    '''

    '''
    DelPrintArea()
    oDoc = LeenoUtils.getDocument()
    # ~oProp = []
    # ~oProp0 = PropertyValue()
    # ~oProp0.Name = 'Overwrite'
    # ~oProp0.Value = True
    # ~oProp1 = PropertyValue()
    # ~oProp1.Name = 'FilterName'
    # ~oProp1.Value = 'calc_pdf_Export'
    # ~oProp.append(oProp0)
    # ~oProp.append(oProp1)
    # ~properties = tuple(oProp)
    # ~sUrl = "file:///W:/test.pdf"
    # ~oDoc.storeToURL(sUrl, properties)

    # ~'crea proprietà e valori in filterData, che verranno passati a filterProps
    filterData = []
    filterData0 = PropertyValue()
    filterData0.Name = "Selection"
    filterData0.Value = oDoc.CurrentController.ActiveSheet
    filterData1 = PropertyValue()
    filterData1.Name = "IsAddStream"
    filterData1.Value = True
    filterData.append(filterData0)
    filterData.append(filterData1)

    # ~'crea proprietà e valori in filterProps, che verranno passati alla funzione di esportazione storeToURL
    filterProps = []
    filterProps0 = PropertyValue()
    filterProps0.Name = "FilterName"
    filterProps0.Value = "calc_pdf_Export"
    filterProps1 = PropertyValue()
    filterProps1.Name = "FilterData"
    filterProps1.Value = tuple(filterData)
    filterProps.append(filterProps0)
    filterProps.append(filterProps1)
    
    properties = tuple(filterProps)

    sUrl = "file:///W:/test.pdf"
    oDoc.storeToURL(sUrl, properties)

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
            # ~Dialogs.VSizer(Items=[
                # ~Dialogs.FixedText(Text='Elenco_prova'),
                # ~Dialogs.Spacer(),
                # ~Dialogs.Edit(Id='npElencoPrezzi', Align=1, FixedHeight=hItems, FixedWidth=nWidth),
                # ~Dialogs.Spacer(),
                # ~Dialogs.Edit(Id='npComputoMetrico', Align=1, FixedHeight=hItems, FixedWidth=nWidth),
                # ~Dialogs.Spacer(),
                # ~Dialogs.Edit(Id='npCostiManodopera', Align=1, FixedHeight=hItems, FixedWidth=nWidth),
                # ~Dialogs.Spacer(),
                # ~Dialogs.Edit(Id='npQuadroEconomico', Align=1, FixedHeight=hItems, FixedWidth=nWidth),
            # ~]),
            Dialogs.Spacer(),
            Dialogs.VSizer(Items=[
                Dialogs.FixedText(Text='Oggetto'),
                Dialogs.Spacer(),
                Dialogs.ListBox(List=oSheets, FixedHeight=hItems * 1, FixedWidth=nWidth * 6),
                # ~Dialogs.Spacer(),
                # ~Dialogs.CheckBox(Id="cbElencoPrezzi", Label="Elenco prezzi", FixedHeight=hItems),
                # ~Dialogs.Spacer(),
                # ~Dialogs.CheckBox(Id="cbComputoMetrico", Label="Computo metrico", FixedHeight=hItems),
                # ~Dialogs.Spacer(),
                # ~Dialogs.CheckBox(Id="cbCostiManodopera", Label="Costi manodopera", FixedHeight=hItems),
                # ~Dialogs.Spacer(),
                # ~Dialogs.CheckBox(Id="cbQuadroEconomico", Label="Quadro economico", FixedHeight=hItems),
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
    
def MENU_debug():
    
    # ~DlgPDF()
    # ~return
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    DLG.chi(len(oSheet.RowPageBreaks))
    return
    # ~testo = oSheet.getCellByPosition(0, 0).String
    # ~txt = " ".join(testo.split())
    # ~oSheet.getCellByPosition(0, 1).String = txt
    # ~DLG.chi(txt)
    import LeenoSettings
    LeenoSettings.MENU_PrintSettings()
    # ~LeenoSettings.MENU_JobSettings()
    return
    import LeenoPdf
    # ~LeenoPdf.MENU_Pdf()
    # ~return
    
    oDoc = LeenoUtils.getDocument()
    es = LeenoPdf.loadExportSettings(oDoc)
    
    # ~DLG.chi(es)
    # ~return

    # ~dlg = PdfDlg()
    dlg = LeenoPdf.PdfDialog()
    dlg.setData(es)

    # se premuto "annulla" non fa nulla
    if dlg.run() < 0:
        return

    es = dlg.getData(_EXPORTSETTINGSITEMS)
    storeExportSettings(oDoc, es)

    # estrae la path
    # ~destFolder = dlg['pathEdit'].getPath()
    destFolder = 'W:\\_dwg\\ULTIMUSFREE\\_SRC'
    
    # ~import LeenoDialogs as DLG
    # ~DLG.chi(destFolder)
    # ~return

    # controlla se selezionato elenco prezzi
    if dlg['cbElencoPrezzi'].getState():
        PdfElencoPrezzi(destFolder, es['npElencoPrezzi'])

    # controlla se selezionato computo metrico
    if dlg['cbComputoMetrico'].getState():
        PdfComputoMetrico(destFolder, es['npComputoMetrico'])
    return

    oDoc = LeenoUtils.getDocument()

    oSheets = list(oDoc.getSheets().getElementNames())
    # ~DLG.chi(oSheets)
    # ~DLG.chi(oSheets)
    # ~nWidth, hItems = Dialogs.getEditBox('g')

    # ~Dialogs.FolderSelect()
    # ~Dialogs.ListBox(Id=None, List=oSheets, Current=None)
    # ~nWidth, hItems = Dialogs.getEditBox('aa')
    Dialogs.ListBox.setList(self, oSheets)
    return
    
    return
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]
    rigenera_voce(lrow)
    lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 1)

    # ~dispatchHelper.executeDispatch(oFrame, '.uno:DataSort', '', 0, properties)

    # ~For Each oSh In oSheets
        # ~If oSh.Name <> "cP_Cop" and oSh.Name <> oActiveSheet Then ' and oSh.Name <> "copyright_LeenO" Then
        # ~p = 0

        # ~'    ThisComponent.CurrentController.Select(ThisComponent.Sheets.GetByName(oSh.Name).getCellByPosition(0,0))
        # ~'    oSh.IsVisible = False
        # ~Else

            # ~Set_Area_Stampa_N("NO_messaggio")
            # ~If     oSh.Name = oActiveSheet Then
                # ~ThisComponent.CurrentController.Select(oSh.getCellRangeByposition(0,0,getLastUsedCol(oSh),getLastUsedRow(oSh)))
                # ~if msgbox (CHR$(10) &"Preferisci nascondere i colori?",36, "") = 6 Then ScriptPy("LeenoSheetUtils.py","SbiancaCellePrintArea")
                # ~unSelect 'unselect ranges 
            # ~Else
            # ~End If
        # ~End If
    # ~Next
# ~'parametri di esportazione
    # ~dim dispatcher as Object
    # ~dim document as Object
    # ~document   = ThisComponent.CurrentController.Frame
    # ~dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    
    # ~rem ----------------------------------------------------------------------
def stampa_PDF():
    DelPrintArea()
    set_area_stampa()
    tempo = ''.join(''.join(''.join(
        str(datetime.now()).split('.')[0].split(' ')).split('-')).split(
            ':'))[:12]
    oDoc = LeenoUtils.getDocument()
    orig = oDoc.getURL()
    dest = orig.split('.')[0] + '-' + tempo + '.pdf'
    ods2pdf(oDoc, dest)
    # ~DLG.chi(dest)
    # ~rem ----------------------------------------------------------------------
import LeenoUtils
import LeenoEvents

import LeenoImport

def MENU_debug():
    LeenoUtils.DocumentRefresh(True)
    return
    oDoc = LeenoUtils.getDocument()
    # ~
    DLG.chi(oDoc.isAutomaticCalculationEnabled())
    return
    lrow = LeggiPosizioneCorrente()[1]

    oSheet = oDoc.CurrentController.ActiveSheet

    DLG.chi(oSheet.getCellRangeByName("A1").CellBackColor) 
    return
    sistema_cose()
    return
    oDoc = LeenoUtils.getDocument()
    oDoc.enableAutomaticCalculation(True)
    oDoc.unlockControllers()
    oDoc.calculateAll()
    oDoc.removeActionLock()
    return
    oDoc = LeenoUtils.getDocument()
    oRange = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    SR = oRange.StartRow + 1
    ER = oRange.EndRow
    oSheet = oDoc.CurrentController.ActiveSheet

    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(1, SR, 1, ER -1))
    return
    LeenoUtils.DocumentRefresh(False)
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = SheetUtils.getUsedArea(oSheet).EndRow + 1
    lCol = SheetUtils.getUsedArea(oSheet).EndColumn 
    for y in reversed(range(1, lrow)):
        if oSheet.getCellByPosition(1, y).String ==  "CAM":
            oSheet.getCellByPosition(2, y).String = "CAM - " + oSheet.getCellByPosition(2, y).String
            # ~oSheet.getRows().removeByIndex(y, 1)

    LeenoUtils.DocumentRefresh(True)


    # ~ LeenoSheetUtils.elimina_righe_vuote()
    # ~SheetUtils.MENU_unisci_fogli()
    # ~DLG.chi(loVersion())
    # ~LeenoEvents.assegna()

    return
    # ~vista_terra_terra()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # ~DLG.mri(oSheet.getColumns().getCount())
    DLG.mri(oSheet)
    return
    import LeenoPdf
    # ~LeenoPdf.MENU_Pdf()

    dlg = LeenoPdf.PdfDialog()

    return
    # ~LeenoSheetUtils.elimina_righe_vuote()
    # ~sistema_cose()
    # ~LI.MENU_emilia_romagna()
    # ~return

# ~'ACCODA PIù FILE DI CALC IN UNO SOLO
	# ~Dim DocName as object, DocUlr as string, dummy()
	# ~Doc = ThisComponent
	# ~Sheet = Doc.Sheets(0) 
	# ~sPath ="W:/_dwg/ULTIMUSFREE/elenchi/Piemonte/2022_luglio/"  ' cartella con i documenti da copiare (non ci deve essere il file destinazione con la macro
	# ~sFileName = Dir(sPath & "*.ods", 0)
# ~'	Barra_Apri_Chiudi_5(".......................Sto lavorando su "& sFileName, 0)
	# ~Do While (sFileName <> "")
		# ~c = Sheet.createCursor
		# ~c.gotoEndOfUsedArea(false)
		# ~LastRow = c.RangeAddress.EndRow + 1
		# ~DocUrl = ConvertToURL(sPath & sFileName)
# ~'on error goto errore
		# ~DocName = StarDesktop.loadComponentFromURL (DocUrl, "_blank",0, Dummy() )
		# ~Sheet1 = DocName.Sheets(0) ' questo indica l'index del foglio da copiare
		# ~c = Sheet1.createCursor
		# ~c.gotoEndOfUsedArea(false)
		# ~LastRow1 = c.RangeAddress.EndRow
	# ~'	oStart=uFindString("ATTENZIONE!", Sheet1)
	# ~'	Srow=oStart.RangeAddress.EndRow+1
	# ~Srow = 2
		# ~Range = Sheet1.getCellRangeByPosition(0, Srow,  12, LastRow1).getDataArray '(1^ colonna, 1^ riga, 10^ colonna, ultima riga)
		# ~DocName.dispose
		# ~dRange  = Sheet.getCellRangeByPosition(0, LastRow, 12, LastRow1 + LastRow-Srow)
		# ~dRange.setDataArray(Range)
		# ~sFileName = Dir()
	# ~Loop
	# ~print "fatto!"
	# ~errore:
# ~End Sub

########################################################################
# ELENCO DEGLI SCRIPT VISUALIZZATI NEL SELETTORE DI MACRO              #
# ~g_exportedScripts = donazioni
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
