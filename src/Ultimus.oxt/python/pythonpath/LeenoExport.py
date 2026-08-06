# -*- Mode: Python; coding: utf-8; indent-tabs-mode: nil; tab-width: 4 -*-
########################################################################
# LeenO - Libreria Esportazione
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

import os
import sys
import threading
import codecs
from xml.etree.ElementTree import Element, SubElement, tostring

# pyrefly: ignore [missing-import]
import uno
# pyrefly: ignore [missing-import]
import unohelper
# pyrefly: ignore [missing-import]
from com.sun.star.beans import PropertyValue

# Local imports
import LeenoUtils
import SheetUtils
import LeenoSheetUtils
import LeenoComputo
import LeenoFormat
import LeenoGlobals
import LeenoConfig
import LeenoDialogs as DLG
import Dialogs

cfg = LeenoConfig.Config()


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
    for el in ("COMPUTO", "VARIANTE", "CONTABILITA"):#, "Elenco Prezzi", 'Analisi di Prezzo'):
        try:
            importo = oDoc.getSheets().getByName(el).getCellRangeByName('A2').String
            # Usa elif invece di multipli if
            if el == 'COMPUTO':
                Dialog_XPWE.getControl(el).Label = 'Computo:'.ljust(13) + importo.rjust(15)
            elif el == 'VARIANTE':
                Dialog_XPWE.getControl(el).Label = 'Variante:'.ljust(13) + importo.rjust(15)
            elif el == 'CONTABILITA':
                Dialog_XPWE.getControl(el).Label = 'Contabilità:'.ljust(13) + importo.rjust(15)
            # elif el == 'Elenco Prezzi':
            #     Dialog_XPWE.getControl(el).Label = 'Elenco Prezzi'
            # elif el == 'Analisi di Prezzo':
            #     Dialog_XPWE.getControl(el).Label = 'Analisi di Prezzo'
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
    # analisi = False
    # Esegue il dialogo e gestisce la risposta
    if Dialog_XPWE.execute() == 1:
        # for el in ("Elenco Prezzi", "COMPUTO", "VARIANTE", "CONTABILITA", "Analisi di Prezzo"):
        for el in ("COMPUTO", "VARIANTE", "CONTABILITA"):
            if Dialog_XPWE.getControl(el).State == 1:
                # if el == "Analisi di Prezzo":
                #     analisi = True
                #     # Non aggiungere "Analisi di Prezzo" alla lista, serve solo per il flag
                # else:
                lista.append(el)
    else:
        # L'utente ha annullato il dialogo
        return
    # Richiede il file di output
    out_file = Dialogs.FileSelect('Salva con nome...', '*.xpwe', 1)
    if out_file == '':
        return
    # Se l'utente ha selezionato "Analisi di Prezzo", assicurati che "Elenco Prezzi" sia nella lista
    # if analisi and "Elenco Prezzi" not in lista:
    #     lista.insert(0, "Elenco Prezzi")
    # Esporta i dati selezionati
    for el in lista:
        # XPWE_out(el, out_file, analisi)
        XPWE_out(el, out_file)






# Scrive un file.
@LeenoUtils.no_refresh
def XPWE_out(elaborato, out_file):
    '''
    esporta il documento in formato XPWE
    elaborato { string } : nome del foglio da esportare
    out_file  { string } : nome base del file
    analisi   { bool }   : se True esporta anche l'analisi di prezzo
    il nome file risulterà out_file-elaborato.xpwe
    '''
    XPWE_out_run(elaborato, out_file)






def XPWE_out_run(elaborato, out_file):
    '''
    esporta il documento in formato XPWE
    elaborato { string } : nome del foglio da esportare
    out_file  { string } : nome base del file
    il nome file risulterà out_file-elaborato.xpwe
    '''
    from pyleeno import dettaglio_misure, numera_voci, Rinumera_TUTTI_Capitoli2, valuta_cella, oggi

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
        if Dialogs.YesNoCancelDialog(
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
       str(LeenoGlobals.getGlobalVar('Lmajor')) + '.' +
       str(LeenoGlobals.getGlobalVar('Lminor')) + '.' +
       LeenoGlobals.getGlobalVar('Lsubv')
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
    if indicator:
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
    if indicator:
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
    if indicator:
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
    if indicator:
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
        # Raccogli le voci che hanno analisi per esportarle dopo con i dettagli completi
        ha_analisi = False
        if(oSheet.getCellByPosition(1, n).Type.value == 'FORMULA' and
           oSheet.getCellByPosition(2, n).Type.value == 'FORMULA'):
            lista_AP.append(oSheet.getCellByPosition(0, n).String)
            ha_analisi = True

        # Salta le voci con analisi: saranno esportate dopo con i dettagli
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
        s_prezzo = valuta_cella(oSheet.getCellByPosition(4, n)).replace(',', '.')
        try:
            Prezzo1.text = str(float(s_prezzo))
        except ValueError:
            Prezzo1.text = '0'
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
        # if oSheet.getCellByPosition(8, n).String == '(AP)':
        if oSheet.getCellByPosition(1, n).Type.value == 'FORMULA':
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
    if indicator:
        indicator.Value = 5
    try:
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
                s_prezzo = valuta_cella(oSheet.getCellByPosition(6, m)).replace(',', '.')
                try:
                    Prezzo1.text = str(float(s_prezzo))
                except ValueError:
                    Prezzo1.text = '0'
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
                        s_prezzo = valuta_cella(oSheet.getCellByPosition(4, x)).replace(',', '.')
                        try:
                            Prezzo.text = str(float(s_prezzo))
                        except ValueError:
                            Prezzo.text = '0'
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
    except Exception:
        pass
    # if elaborato == 'Elenco_Prezzi':
    #     pass
    # else:
    # COMPUTO/VARIANTE/CONTABILITA
    oSheet = oDoc.getSheets().getByName(elaborato)
    PweVociComputo = SubElement(PweMisurazioni, 'PweVociComputo')
    oDoc.CurrentController.setActiveSheet(oSheet)
    Rinumera_TUTTI_Capitoli2(oSheet)
    nVCItem = 2
    if indicator:
        indicator.Value = 6
        indicator.start(f'Esportazione di {elaborato} in corso...', LeenoSheetUtils.cercaUltimaVoce(oSheet))  # max progresso
    for n in range(0, LeenoSheetUtils.cercaUltimaVoce(oSheet)):
        if indicator:
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
        Text=f'Esportazione in formato XPWE eseguita con successo sul file:\n\n {LeenoUtils.wrap_path(out_file, max_len=60)}'
'\n\n----\n'
'XPWE è un formato XML di interscambio per Primus di ACCA.\n'
'Prima di utilizzare questo file in Primus, assicurarsi che le percentuali \
di Spese Generali e Utile d\'Impresa siano impostate correttamente, \
in modo da garantire l\'esatta elaborazione dei dati.')

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



def clean_markdown(val):
    """
    Pulisce e formatta una stringa per l'inclusione in una cella di tabella Markdown.
    Sostituisce i ritorni a capo con tag <br> e scherma il carattere pipe '|'.
    """
    if val is None:
        return ""
    val_str = str(val)
    # Rimpiazza i ritorni a capo con un break HTML per preservare la struttura di riga
    val_str = val_str.replace('\r\n', ' <br> ').replace('\n', ' <br> ').replace('\r', ' <br> ')
    # Escapa il carattere pipe '|' per non interrompere le colonne Markdown
    val_str = val_str.replace('|', '\\|')
    return val_str.strip()






def MENU_esporta_markdown():
    """
    Esporta l'area selezionata di una tabella in formato Markdown.
    La macro recupera l'intervallo selezionato dal controller attivo di Calc,
    filtra per esportare solo le righe e colonne visibili, chiede conferma
    per l'intestazione e salva il file risultante tramite FileSelect.
    """
    oDoc = LeenoUtils.getDocument()
    oController = oDoc.getCurrentController()
    oSheet = oController.getActiveSheet()
    oSel = oController.getSelection()

    if oSel is None:
        Dialogs.Exclamation(Title='Avviso!', Text='Nessuna selezione attiva.')
        return

    # Gestione delle selezioni multiple disgiunte
    if oSel.supportsService("com.sun.star.sheet.SheetCellRanges"):
        count = oSel.getCount()
        if count == 0:
            Dialogs.Exclamation(Title='Avviso!', Text='Nessuna selezione attiva.')
            return
        # Avverte l'utente ed esporta il primo intervallo
        Dialogs.Info(Title='Avviso!', Text='La selezione contiene più intervalli disgiunti. Verrà esportato solo il primo intervallo.')
        oRange = oSel.getByIndex(0)
    elif oSel.supportsService("com.sun.star.sheet.SheetCellRange"):
        oRange = oSel
    else:
        Dialogs.Exclamation(Title='Avviso!', Text='La selezione corrente non è un intervallo di celle valido.')
        return

    range_addr = oRange.getRangeAddress()
    start_col = range_addr.StartColumn
    end_col = range_addr.EndColumn
    start_row = range_addr.StartRow
    end_row = range_addr.EndRow

    num_rows = end_row - start_row + 1
    num_cols = end_col - start_col + 1

    if num_rows == 0 or num_cols == 0:
        Dialogs.Exclamation(Title='Avviso!', Text='L\'intervallo selezionato è vuoto.')
        return

    # Filtra solo colonne visibili
    visible_cols = []
    for c in range(num_cols):
        abs_col = start_col + c
        if oSheet.getColumns().getByIndex(abs_col).IsVisible:
            visible_cols.append(c)

    # Filtra solo righe visibili
    visible_rows = []
    for r in range(num_rows):
        abs_row = start_row + r
        if oSheet.getRows().getByIndex(abs_row).IsVisible:
            visible_rows.append(r)

    if not visible_cols or not visible_rows:
        Dialogs.Exclamation(Title='Avviso!', Text='La selezione non contiene celle visibili.')
        return

    # Chiede all'utente se impostare la prima riga visibile come intestazione
    usa_prima_riga_come_intestazione = False
    if len(visible_rows) >= 2:
        res = Dialogs.YesNoDialog(
            IconType="question",
            Title="Esporta in Markdown",
            Text="Vuoi utilizzare la prima riga visibile selezionata come intestazione della tabella?"
        )
        if res == 1:
            usa_prima_riga_come_intestazione = True

    # Lettura delle celle con indicatore di progresso nativo se la selezione è grande
    mostra_progresso = len(visible_rows) > 100
    if mostra_progresso:
        indicator = oController.getStatusIndicator()
        indicator.start("Esportazione Markdown...", len(visible_rows))

    data = []
    for idx, r in enumerate(visible_rows):
        if mostra_progresso:
            indicator.setValue(idx)
        row_data = []
        for c in visible_cols:
            cell = oRange.getCellByPosition(c, r)
            row_data.append(cell.String)
        data.append(row_data)

    if mostra_progresso:
        indicator.end()

    # Generazione della struttura della tabella Markdown
    if usa_prima_riga_come_intestazione:
        headers = [clean_markdown(val) for val in data[0]]
        rows = [[clean_markdown(val) for val in row] for row in data[1:]]
    else:
        headers = [f"Colonna {i+1}" for i in range(len(visible_cols))]
        rows = [[clean_markdown(val) for val in row] for row in data]

    # Composizione delle righe in formato Markdown
    header_str = "| " + " | ".join(headers) + " |"
    delimiter_str = "| " + " | ".join(["---"] * len(headers)) + " |"
    body_strs = []
    for row in rows:
        body_strs.append("| " + " | ".join(row) + " |")

    markdown_content = "\n".join([header_str, delimiter_str] + body_strs) + "\n"

    # Ottieni il nome del file corrente come suggerimento predefinito
    doc_url = oDoc.getURL()
    default_filename = ""
    if doc_url:
        system_path = uno.fileUrlToSystemPath(doc_url)
        base_name = os.path.basename(system_path)
        name_without_ext = os.path.splitext(base_name)[0]
        default_filename = name_without_ext + ".md"
    elif hasattr(oDoc, 'Title') and oDoc.Title:
        name_without_ext = os.path.splitext(oDoc.Title)[0]
        default_filename = name_without_ext + ".md"
    else:
        default_filename = "tabella.md"

    # Selezione del percorso e salvataggio
    out_file = Dialogs.FileSelect('Salva tabella Markdown con nome...', '*.md', 1, defaultName=default_filename)
    if not out_file:
        return

    # Garantisce che il file esportato abbia l'estensione .md
    if not out_file.lower().endswith('.md'):
        out_file += '.md'

    try:
        with open(out_file, 'w', encoding='utf-8', newline='') as f:
            f.write(markdown_content)
        Dialogs.Info(
            Title='Esportazione completata',
            Text=f'Il file è stato esportato correttamente in:\n\n{out_file}'
        )
    except Exception as e:
        Dialogs.Exclamation(
            Title='Errore!',
            Text=f'Impossibile scrivere il file:\n{str(e)}'
        )




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
        Text=f"File creato con successo:\n{output_path}"
        )





# ###############################################################


@LeenoUtils.preserve_clipboard
def MENU_SheetToDoc():
    '''
    Copia il foglio corrente in un nuovo documento.
    '''
    oDoc = LeenoUtils.getDocument()
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    oProp = []
    oProp0 = PropertyValue()
    oProp0.Name = 'DocName'
    oProp0.Value = ''
    oProp1 = PropertyValue()
    oProp1.Name = 'Index'
    oProp1.Value = 32767
    oProp2 = PropertyValue()
    oProp2.Name = 'Copy'
    oProp2.Value = True
    oProp.append(oProp0)
    oProp.append(oProp1)
    oProp.append(oProp2)
    properties = tuple(oProp)
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, '.uno:Move', '', 0, properties)
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect
    oDoc = LeenoUtils.getDocument()

    oSheet = oDoc.CurrentController.ActiveSheet

    if "COMPUTO" in oSheet.Name or "VARIANTE" in oSheet.Name:
        import pyleeno as PL
        oDoc.CurrentController.select(oSheet.getCellRangeByName('A1:I1048576'))
        PL.comando('Copy')
        #oDoc.CurrentController.select(oCell)
        PL.paste_clip(insCells=0, pastevalue=True)
        oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselec
    oDoc.CurrentController.ZoomValue = 100
    return
