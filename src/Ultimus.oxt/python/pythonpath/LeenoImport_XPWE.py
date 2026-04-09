"""
Importazione computo/variante/contabilità/prezzario
dal formato XPWE
"""
import logging
import re
import os

from xml.etree.ElementTree import ElementTree, ParseError
from com.sun.star.table import CellRangeAddress

import LeenoUtils
import LeenoFormat

import pyleeno as PL
import LeenoDialogs as DLG

import SheetUtils
import LeenoConfig
import LeenoSheetUtils
import LeenoAnalysis
import LeenoComputo
import LeenoVariante
import LeenoContab

import Dialogs


def get_xml_text(element, path, default=''):
    ''' Estrattore sicuro di testo da un elemento XML '''
    try:
        found = element.find(path)
        if found is not None:
            return found.text or default
        return default
    except Exception:
        return default


def leggiAnagraficaGenerale(dati):
    ''' legge i dati anagrafici generali '''
    datiAnagrafici = {}
    if dati is None:
        return {k: '' for k in ('comune', 'oggetto', 'committente', 'impresa')}

    try:
        # PweDatiGenerali (solitamente ha un figlio che contiene i dati)
        child = list(dati)[0]
        # Spesso c'è un ulteriore livello (es. DatiGenerali)
        content = list(child)[0] if list(child) else child

        datiAnagrafici['comune'] = get_xml_text(content, 'Comune')
        datiAnagrafici['oggetto'] = get_xml_text(content, 'Oggetto')
        datiAnagrafici['committente'] = get_xml_text(content, 'Committente')
        datiAnagrafici['impresa'] = get_xml_text(content, 'Impresa')
    except Exception:
        datiAnagrafici['comune'] = ''
        datiAnagrafici['oggetto'] = ''
        datiAnagrafici['committente'] = ''
        datiAnagrafici['impresa'] = ''

    return datiAnagrafici


def leggiSuperCapitoli(CapCat):
    ''' legge SuperCapitoli '''
    listaSuperCapitoli = []
    found = CapCat.find('PweDGSuperCapitoli')
    if found is not None:
        for elem in list(found):
            listaSuperCapitoli.append({
                'id_sc': elem.get('ID'),
                'codice': get_xml_text(elem, 'Codice'),
                'dessintetica': get_xml_text(elem, 'DesSintetica'),
                'percentuale': get_xml_text(elem, 'Percentuale')
            })
    return listaSuperCapitoli


def leggiCapitoli(CapCat):
    ''' legge Capitoli '''
    listaCapitoli = []
    found = CapCat.find('PweDGCapitoli')
    if found is not None:
        for elem in list(found):
            dessintetica = get_xml_text(elem, 'DesSintetica')
            if dessintetica == "Nuova voce":
                dessintetica = get_xml_text(elem, 'DesEstesa')

            listaCapitoli.append({
                'id_sc': elem.get('ID'),
                'codice': get_xml_text(elem, 'Codice'),
                'dessintetica': dessintetica,
                'percentuale': get_xml_text(elem, 'Percentuale')
            })
    return listaCapitoli


def leggiSottoCapitoli(CapCat):
    ''' legge SottoCapitoli '''
    listaSottoCapitoli = []
    found = CapCat.find('PweDGSubCapitoli')
    if found is not None:
        for elem in list(found):
            dessintetica = get_xml_text(elem, 'DesSintetica')
            if dessintetica == "Nuova voce":
                dessintetica = get_xml_text(elem, 'DesEstesa')

            listaSottoCapitoli.append({
                'id_sc': elem.get('ID'),
                'codice': get_xml_text(elem, 'Codice'),
                'dessintetica': dessintetica,
                'percentuale': get_xml_text(elem, 'Percentuale')
            })
    return listaSottoCapitoli


def leggiSuperCategorie(CapCat):
    ''' legge SuperCategorie '''
    listaSuperCategorie = []
    found = CapCat.find('PweDGSuperCategorie')
    if found is not None:
        for elem in list(found):
            listaSuperCategorie.append({
                'id_sc': elem.get('ID'),
                'codice': get_xml_text(elem, 'Codice'),
                'dessintetica': get_xml_text(elem, 'DesSintetica'),
                'percentuale': get_xml_text(elem, 'Percentuale')
            })
    return listaSuperCategorie


def leggiCategorie(CapCat):
    ''' legge Categorie '''
    listaCategorie = []
    found = CapCat.find('PweDGCategorie')
    if found is not None:
        for elem in list(found):
            listaCategorie.append({
                'id_sc': elem.get('ID'),
                'codice': get_xml_text(elem, 'Codice'),
                'dessintetica': get_xml_text(elem, 'DesSintetica'),
                'percentuale': get_xml_text(elem, 'Percentuale')
            })
    return listaCategorie


def leggiSottoCategorie(CapCat):
    ''' legge SottoCategorie '''
    listaSottoCategorie = []
    found = CapCat.find('PweDGSubCategorie')
    if found is not None:
        for elem in list(found):
            listaSottoCategorie.append({
                'id_sc': elem.get('ID'),
                'codice': get_xml_text(elem, 'Codice'),
                'dessintetica': get_xml_text(elem, 'DesSintetica'),
                'percentuale': get_xml_text(elem, 'Percentuale')
            })
    return listaSottoCategorie


def leggiCapitoliCategorie(dati):
    ''' legge capitoli e categorie '''

    res = {
        'SuperCapitoli': [],
        'Capitoli': [],
        'SottoCapitoli': [],
        'SuperCategorie': [],
        'Categorie': [],
        'SottoCategorie': []
    }

    # PweDGCapitoliCategorie
    try:
        CapCat = dati.find('PweDGCapitoliCategorie')
    except AttributeError:
        return res

    # legge SuperCapitoli
    res['SuperCapitoli'] = leggiSuperCapitoli(CapCat)
    # legge Capitoli
    res['Capitoli'] = leggiCapitoli(CapCat)
    # legge SottoCapitoli
    res['SottoCapitoli'] = leggiSottoCapitoli(CapCat)
    # legge SuperCategorie
    res['SuperCategorie'] = leggiSuperCategorie(CapCat)
    # legge Categorie
    res['Categorie'] = leggiCategorie(CapCat)
    # legge SottoCategorie
    res['SottoCategorie'] = leggiSottoCategorie(CapCat)

    return res


def leggiDatiGeneraliAnalisi(dati):
    ''' legge i dati generali di analisi '''
    speseGenerali = 0
    utiliImpresa = 0
    oneriAccessoriSicurezza = 0

    try:
        found_moduli = dati.find('PweDGModuli')
        if found_moduli is not None:
            PweDGAnalisi = list(found_moduli)[0]

            def parse_percent(tag):
                text = get_xml_text(PweDGAnalisi, tag, '0')
                try:
                    return float(text.replace(',', '.')) / 100
                except (ValueError, TypeError):
                    return 0.0

            speseGenerali = parse_percent('SpeseGenerali')
            utiliImpresa = parse_percent('UtiliImpresa')
            oneriAccessoriSicurezza = parse_percent('OneriAccessoriSc')
    except Exception:
        pass

    return {
        'SpeseGenerali': speseGenerali,
        'UtiliImpresa': utiliImpresa,
        'OneriAccessoriSicurezza': oneriAccessoriSicurezza
    }


def leggiApprossimazioni(dati):
    ''' legge le impostazioni di approssimazione numerica '''
    res = {}
    try:
        found_config = dati.find('PweDGConfigurazione')
        if found_config is None:
            return res

        PweDGConfigNumeri = list(found_config)[0]

        def parse_decimal_digits(tag):
            text = get_xml_text(PweDGConfigNumeri, tag)
            if not text:
                return None
            try:
                # XPWE format often uses "0.00|0" or similar
                return int(text.split('.')[-1].split('|')[0])
            except (ValueError, IndexError):
                return None

        for field in ('PartiUguali', 'Larghezza', 'Lunghezza', 'HPeso', 'Quantita', 'Prezzi', 'PrezziTotale'):
            val = parse_decimal_digits(field)
            if val is not None:
                res[field] = val
    except Exception:
        pass

    return res


def compilaAnagraficaGenerale(oDoc, datiAnagrafici):
    ''' compilo Anagrafica generale '''

    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = datiAnagrafici['oggetto']
    oSheet.getCellByPosition(2, 3).String = datiAnagrafici['comune']
    oSheet.getCellByPosition(2, 5).String = datiAnagrafici['committente']
    oSheet.getCellByPosition(3, 16).String = datiAnagrafici['impresa']


def compilaDatiGeneraliAnalisi(oDoc, datiGeneraliAnalisi):
    ''' compila i dati generali dell'analisi, ovvero le percentuali relative '''

    oSheet = oDoc.getSheets().getByName('S1')

    oSheet.getCellByPosition(7, 318).Value = datiGeneraliAnalisi['OneriAccessoriSicurezza']
    oSheet.getCellByPosition(7, 319).Value = datiGeneraliAnalisi['SpeseGenerali']
    oSheet.getCellByPosition(7, 320).Value = datiGeneraliAnalisi['UtiliImpresa']


def compilaApprossimazioni(oDoc, approssimazioni):
    ''' compila le approssimazioni modificando la precisione nelle relative celle '''

    if 'PartiUguali' in approssimazioni:
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a PU', approssimazioni['PartiUguali'])
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a PU ROSSO', approssimazioni['PartiUguali'])
    if 'Lunghezza' in approssimazioni:
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a LUNG', approssimazioni['Lunghezza'])
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a LUNG ROSSO', approssimazioni['Lunghezza'])
    if 'Larghezza' in approssimazioni:
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a LARG', approssimazioni['Larghezza'])
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a LARG ROSSO', approssimazioni['Larghezza'])
    if 'HPeso' in approssimazioni:
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a peso', approssimazioni['HPeso'])
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a peso ROSSO', approssimazioni['HPeso'])
    if 'Quantita' in approssimazioni:
        for el in ('Comp-Variante num sotto', 'Comp-Variante num sotto ROSSO', 'An-lavoraz-input', 'Blu', 'Blu ROSSO'):
            LeenoFormat.setCellStyleDecimalPlaces(el, approssimazioni['Quantita'])
    if 'Prezzi' in approssimazioni:
        for el in ('comp sotto Unitario', 'An-lavoraz-generica'):
            LeenoFormat.setCellStyleDecimalPlaces(el, approssimazioni['Prezzi'])
    if 'PrezziTotale' in approssimazioni:
        for el in ('comp sotto Euro Originale', 'Livello-0-scritta mini val',
                   'Livello-1-scritta mini val', 'livello2 scritta mini',
                   'Comp TOTALI', 'Ultimus_totali_1', 'Ultimus_bordo',
                   'ULTIMUS_3', 'Ultimus_Bordo_sotto',
                   'Comp-Variante num sotto', 'An-valuta-dx', 'An-1v-dx',
                   'An-lavoraz-generica', 'An-lavoraz-Utili-num sin'):
            LeenoFormat.setCellStyleDecimalPlaces(el, approssimazioni['PrezziTotale'])


def leggiElencoPrezzi(misurazioni):
    ''' legge l'elenco prezzi e le eventuali voci di analisi prezzi '''
    dizionarioArticoli = {}
    listaArticoli = []
    listaAnalisi = []
    listaTariffeAnalisi = []

    if misurazioni is None or len(list(misurazioni)) == 0:
        return {
            'DizionarioArticoli': {},
            'ListaArticoli': [],
            'ListaAnalisi': [],
            'ListaTariffeAnalisi': []
        }

    PweElencoPrezzi = list(misurazioni)[0]
    epitems = PweElencoPrezzi.findall('EPItem')

    for elem in epitems:
        id_ep = elem.get('ID')
        if not id_ep: continue

        dizionarioArticolo = {}

        tipoep = get_xml_text(elem, 'TipoEP', '0')
        tariffa = get_xml_text(elem, 'Tariffa')

        # Voce Della Sicurezza
        flags = get_xml_text(elem, 'Flags')
        if flags == '134217728':
            tariffa = "VDS_" + tariffa

        articolo = get_xml_text(elem, 'Articolo')
        desridotta = get_xml_text(elem, 'DesRidotta')
        destestesa = get_xml_text(elem, 'DesEstesa')
        desbreve = get_xml_text(elem, 'DesBreve')
        if desbreve:
            desridotta = desbreve

        unmisura = get_xml_text(elem, 'UnMisura')

        def parse_price(tag):
            text = get_xml_text(elem, tag)
            if not text or text == '0':
                return ''
            try:
                return float(text.replace(',', '.'))
            except (ValueError, TypeError):
                return ''

        prezzo1 = parse_price('Prezzo1')
        prezzo2 = get_xml_text(elem, 'Prezzo2', '0')
        prezzo3 = get_xml_text(elem, 'Prezzo3', '0')
        prezzo4 = get_xml_text(elem, 'Prezzo4', '0')
        prezzo5 = get_xml_text(elem, 'Prezzo5', '0')

        idspcap = get_xml_text(elem, 'IDSpCap')
        idcap = get_xml_text(elem, 'IDCap')
        data = get_xml_text(elem, 'Data')
        adrinternet = get_xml_text(elem, 'AdrInternet')

        def parse_incidenza(tag):
            text = get_xml_text(elem, tag)
            if not text: return ''
            try:
                val = float(text.replace(',', '.'))
                return val / 100 if val != 0 else ''
            except (ValueError, TypeError):
                return ''

        IncSIC = parse_incidenza('IncSIC')
        IncMDO = parse_incidenza('IncMDO')
        IncMAT = parse_incidenza('IncMAT')
        IncATTR = parse_incidenza('IncATTR')

        dizionarioArticolo.update({
            'tipoep': tipoep, 'tariffa': tariffa, 'articolo': articolo,
            'desridotta': desridotta, 'destestesa': destestesa, 'desbreve': desbreve,
            'unmisura': unmisura, 'prezzo1': prezzo1, 'prezzo2': prezzo2,
            'prezzo3': prezzo3, 'prezzo4': prezzo4, 'prezzo5': prezzo5,
            'idspcap': idspcap, 'idcap': idcap, 'flags': flags,
            'data': data, 'adrinternet': adrinternet,
            'IncSIC': IncSIC, 'IncMDO': IncMDO, 'IncMAT': IncMAT, 'IncATTR': IncATTR
        })

        dizionarioArticoli[id_ep] = dizionarioArticolo

        # leggo analisi di prezzo
        pweepanalisi = elem.find('PweEPAnalisi')
        if pweepanalisi is not None:
            PweEPAR = pweepanalisi.find('PweEPAR')
            if PweEPAR is not None:
                EPARItem = PweEPAR.findall('EPARItem')
                analisi = []
                for el in EPARItem:
                    id_ep_an = get_xml_text(el, 'IDEP')
                    an_des = get_xml_text(el, 'Descrizione')
                    an_um = get_xml_text(el, 'Misura')
                    an_qt = get_xml_text(el, 'Qt').replace(' ', '')
                    an_pr = get_xml_text(el, 'Prezzo').replace(' ', '')
                    analisi.append((id_ep_an, an_des, an_um, an_qt, an_pr))

                listaAnalisi.append([tariffa, destestesa, unmisura, analisi, prezzo1])
                listaTariffeAnalisi.append(tariffa)
            else:
                listaArticoli.append((tariffa, destestesa, unmisura, IncSIC, prezzo1, IncMDO, IncMAT, IncATTR))
        else:
            listaArticoli.append((tariffa, destestesa, unmisura, IncSIC, prezzo1, IncMDO, IncMAT, IncATTR))

    return {
        'DizionarioArticoli': dizionarioArticoli,
        'ListaArticoli': listaArticoli,
        'ListaAnalisi': listaAnalisi,
        'ListaTariffeAnalisi': listaTariffeAnalisi
    }

def leggiMisurazioni(misurazioni, ordina):
    """leggo voci di misurazione e righe"""
    if not misurazioni or len(list(misurazioni)) < 2:
        return []

    def clean_value(elem, tag, pattern):
        value = get_xml_text(elem, tag)
        if not value:
            return value
        try:
            value = re.sub(pattern, r'\1*\2', value)
        except Exception:
            pass
        if value is not None:
            if '  ' in value or value == '0.00':
                return None
        return value

    listaMisure = []
    try:
        items = list(misurazioni)
        PweVociComputo = items[1]
        vcitems = PweVociComputo.findall('VCItem')
        prova_l = []
        pattern = r'(\d+)(\()'

        for elem in vcitems:
            id_vc = elem.get('ID')
            id_ep = get_xml_text(elem, 'IDEP')
            quantita_voce = get_xml_text(elem, 'Quantita')
            datamis = get_xml_text(elem, 'DataMis')
            flags_voce = get_xml_text(elem, 'Flags')
            idspcat = get_xml_text(elem, 'IDSpCat')
            idcat = get_xml_text(elem, 'IDCat')
            idsbcat = get_xml_text(elem, 'IDSbCat')

            righi_found = list(elem)[-1] if list(elem) else None
            righi_mis = righi_found.findall('RGItem') if righi_found is not None else []
            lista_righe = []

            for el in righi_mis:
                riga_misura = (
                    get_xml_text(el, 'Descrizione'),
                    '', '',
                    clean_value(el, 'PartiUguali', pattern),
                    clean_value(el, 'Lunghezza', pattern),
                    clean_value(el, 'Larghezza', pattern),
                    clean_value(el, 'HPeso', pattern),
                    get_xml_text(el, 'Quantita'),
                    get_xml_text(el, 'Flags'),
                    get_xml_text(el, 'IDVV')
                )
                lista_righe.append(riga_misura)

            diz_misura = {
                'id_vc': id_vc, 'id_ep': id_ep, 'quantita': quantita_voce,
                'datamis': datamis, 'flags': flags_voce, 'idspcat': idspcat,
                'idcat': idcat, 'idsbcat': idsbcat, 'lista_rig': lista_righe
            }
            new_id = f"{PL.strall(idspcat)}.{PL.strall(idcat)}.{PL.strall(idsbcat)}"
            prova_l.append((new_id, diz_misura))
            listaMisure.append(diz_misura)

        if listaMisure and ordina:
            riordinate = sorted(prova_l, key=lambda el: el[0])
            listaMisure = [el[1] for el in riordinate]

    except Exception:
        Dialogs.Exclamation(
            Title="Attenzione",
            Text="Nel file scelto non risultano esserci voci di misurazione,\n"
                 "perciò saranno importate le sole voci di Elenco Prezzi."
        )

    return listaMisure

################################################
################################################
################################################

def estraiDatiCapitoliCategorie(capitoliCategorie, catName):
    resList = []
    for el in capitoliCategorie[catName]:
        tariffa = el.get('codice')
        if tariffa is not None:
            destestesa = el.get('dessintetica')
            titolo = (tariffa, destestesa, '', '', '', '', '')
            resList.append(titolo)
    return tuple(resList)


def riempiBloccoElencoPrezzi(oSheet, dati, col, indicator=None, case_sensitive=False):
    # 1. Recupera tutti i codici esistenti nel foglio (colonna 0, dalla riga 3 in poi)
    existing_codes = set()
    max_row = oSheet.getRows().getCount()

    if max_row > 3:
        # Legge i codici in batch per ottimizzazione (evita timeout su fogli grandi)
        chunk_size = 1000  # Adjust based on performance
        for start_row in range(3, max_row, chunk_size):
            end_row = min(start_row + chunk_size, max_row)
            codici_esistenti = oSheet.getCellRangeByPosition(0, start_row, 0, end_row - 1).getDataArray()
            for codice in codici_esistenti:
                if codice and codice[0]:
                    code = str(codice[0]).strip()
                    if not case_sensitive:
                        code = code.lower()
                    existing_codes.add(code)

    # 2. Filtra i nuovi dati: rimuove duplicati interni + codici già esistenti
    nuovi_dati = []
    seen_new_codes = set()

    for riga in dati:
        if not riga or not riga[0]:  # Skip righe vuote
            continue

        codice = str(riga[0]).strip()
        if not case_sensitive:
            codice = codice.lower()

        # Controlla sia nei codici esistenti che nei nuovi già processati
        if codice not in existing_codes and codice not in seen_new_codes:
            nuovi_dati.append(riga)
            seen_new_codes.add(codice)

    if not nuovi_dati:
        return

    # 3. Inserimento dati (ottimizzato per grandi blocchi)
    righe_totali = len(nuovi_dati)
    colonne = len(nuovi_dati[0])
    sRow = 4
    oSheet.getRows().insertByIndex(sRow, righe_totali)

    # Inserimento a step (es. 100 righe alla volta)
    step = 100
    riga = 0
    while riga < righe_totali:
        sliced = nuovi_dati[riga:riga + step]
        num = len(sliced)
        oRange = oSheet.getCellRangeByPosition(0, sRow + riga, colonne - 1, sRow + riga + num - 1)
        oRange.setDataArray(sliced)
        # PL.stileCelleElencoPrezzi(oSheet, sRow + riga, sRow + riga + num - 1, col)
        riga += num

def compilaElencoPrezzi(oDoc, capitoliCategorie, elencoPrezzi, indicator=None):
    ''' compila l'elenco prezzi '''

    # per prima cosa estrae le liste di EP, capitoli, eccetera
    # e le converte in array

    # articoli
    arrayArticoli = tuple(elencoPrezzi['ListaArticoli'])
    righeArticoli = len(arrayArticoli)
    # SuperCapitoli
    arraySuperCapitoli = estraiDatiCapitoliCategorie(capitoliCategorie, 'SuperCapitoli')
    righeSuperCapitoli = len(arraySuperCapitoli)
    # capitoli
    arrayCapitoli = estraiDatiCapitoliCategorie(capitoliCategorie, 'Capitoli')
    righeCapitoli = len(arrayCapitoli)
    # SottoCapitoli
    arraySottoCapitoli = estraiDatiCapitoliCategorie(capitoliCategorie, 'SottoCapitoli')
    righeSottoCapitoli = len(arraySottoCapitoli)

    # numero totale di righe da inserire
    righeTotali = righeArticoli + righeSuperCapitoli + righeCapitoli + righeSottoCapitoli

    # inizializza la progressbar
    # progress.setLimits(0, righeTotali)
    # progress.setValue(0)

    # compilo Elenco Prezzi
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')

    riempiBloccoElencoPrezzi(oSheet, arrayArticoli, None, indicator=indicator)
    '''
    aggiungo i capitoli alla lista delle voci
     giallo(16777072,16777120,16777168)
     verde(9502608,13696976,15794160)
     viola(12632319,13684991,15790335)
    SUPERCAPITOLI
    '''
    # SuperCapitoli
    if righeSuperCapitoli:
        riempiBloccoElencoPrezzi(oSheet, arraySuperCapitoli, 16777072, indicator=indicator)

    # Capitoli
    if righeCapitoli:
        riempiBloccoElencoPrezzi(oSheet, arrayCapitoli, 16777120, indicator=indicator)

    # SottoCapitoli
    if righeSottoCapitoli:
        riempiBloccoElencoPrezzi(oSheet, arraySottoCapitoli, 16777168, indicator=indicator)

    PL.riordina_ElencoPrezzi()


def rimuoviAnalisiVuote(oSheet):
    """Scansiona ed elimina tutte le schede di analisi segnaposto o vuote."""
    import SheetUtils
    import LeenoAnalysis
    last_row = SheetUtils.getLastUsedRow(oSheet)
    # Scansiona all'indietro per poter eliminare in sicurezza
    for n in reversed(range(0, last_row + 1)):
        # Identifica la riga dei dati dallo stile della colonna B (An-1-descr_)
        if oSheet.getCellByPosition(1, n).CellStyle == 'An-1-descr_':
            cod = oSheet.getCellByPosition(0, n).String.strip()
            des = oSheet.getCellByPosition(1, n).String.strip()
            um  = oSheet.getCellByPosition(2, n).String.strip()

            # Criteri di identificazione segnaposto (tutte le condizioni devono essere vere)
            if cod == 'AP' and des.startswith('<<<Scrivi la descrizione') and um == 'U.M. ?':
                # Chiamiamo circoscriveAnalisi su n per determinare SR ed ER
                try:
                    sStRange = LeenoAnalysis.circoscriveAnalisi(oSheet, n)
                    SR = sStRange.RangeAddress.StartRow
                    ER = sStRange.RangeAddress.EndRow
                    oSheet.getRows().removeByIndex(SR, ER - SR + 1)
                except Exception:
                    pass

def compilaAnalisiPrezzi(oDoc, elencoPrezzi, indicator=None):
    ''' Compilo Analisi di prezzo '''
    numAnalisi = len(elencoPrezzi['ListaAnalisi'])
    if numAnalisi == 0:
        return

    # Se non c'è un indicatore passato, ne crea uno locale per fallback (ma non dovrebbe accadere)
    if not indicator:
        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start("Elaborazione Analisi dei Prezzi in corso...", numAnalisi)

    # Ottieni il foglio delle analisi (solo setup, senza inserire schede)
    oSheet, _ = LeenoAnalysis.inizializza_analisi(oDoc, nuovaScheda=False)
    existing_codes = set()
    max_row = oSheet.getRows().getCount()

    if max_row > 0:
        # Legge i codici esistenti (prima colonna)
        codici_esistenti = oSheet.getCellRangeByPosition(0, 0, 0, max_row - 1).getDataArray()
        existing_codes = {str(codice[0]).strip().lower() for codice in codici_esistenti if codice and codice[0]}

    # inizializza l'analisi dei prezzi
    oSheet, startRow = LeenoAnalysis.inizializza_analisi(oDoc, nuovaScheda=True)

    # compila le voci dell'analisi
    for i, el in enumerate(elencoPrezzi['ListaAnalisi']):
        # Scaling progress: 60% -> 80%
        if indicator:
            indicator.Value = 60 + int((i / numAnalisi) * 20)
            indicator.Text = f"Compilazione analisi: {el[0]}"

        codice = str(el[0]).strip().lower()
        if codice in existing_codes:
            continue

        prezzo_finale = el[-1]
        sStRange = LeenoAnalysis.circoscriveAnalisi(oSheet, startRow)
        lrow = sStRange.RangeAddress.StartRow + 2
        oSheet.getCellByPosition(0, lrow).String = el[0]
        oSheet.getCellByPosition(1, lrow).String = el[1]
        oSheet.getCellByPosition(2, lrow).String = el[2]

        y = 0
        n = lrow + 2
        for x in el[3]:
            # el[3][y] è lo stesso di x (tupla: id, descrizione, u.m., quantità, prezzo)
            if x[1] in ('MANODOPERA', 'MATERIALI', 'NOLI', 'TRASPORTI', 'ALTRE FORNITURE E PRESTAZIONI', 'overflow'):
                if x[1] != 'overflow':
                    n = SheetUtils.uFindStringCol(x[1], 1, oSheet, lrow)
            else:
                LeenoAnalysis.copiaRigaAnalisi(oSheet, n)
                # Cerca la tariffa nel dizionario articoli se presente
                art_diz = elencoPrezzi['DizionarioArticoli'].get(x[0])
                if art_diz is not None:
                    oSheet.getCellByPosition(0, n).String = art_diz.get('tariffa', '')
                else:
                    # Per gli inserimenti liberi
                    oSheet.getCellByPosition(0, n).String = ''
                    try:
                        oSheet.getCellByPosition(1, n).String = x[1]
                    except:
                        oSheet.getCellByPosition(1, n).String = ''
                    oSheet.getCellByPosition(2, n).String = x[2]

                    # Gestione quantità e prezzo
                    try:
                        qt = str(x[3]).replace(',', '.')
                        oSheet.getCellByPosition(3, n).Value = float(qt) if qt else 0
                    except:
                        oSheet.getCellByPosition(3, n).Value = 0

                    try:
                        pr = str(x[4]).replace(',', '.')
                        oSheet.getCellByPosition(4, n).Value = float(pr) if pr else 0
                    except:
                        oSheet.getCellByPosition(4, n).Value = 0

                # Se non è una categoria predefinita, gestisce la quantità forzata o formula
                if x[1] not in ('MANODOPERA', 'MATERIALI', 'NOLI', 'TRASPORTI', 'ALTRE FORNITURE E PRESTAZIONI', 'overflow'):
                    if x[3] == '':
                        oSheet.getCellByPosition(3, n).Value = 0
                    else:
                        try:
                            float(str(x[3]).replace(',', '.'))
                            oSheet.getCellByPosition(3, n).Value = float(str(x[3]).replace(',', '.'))
                        except Exception:
                            oSheet.getCellByPosition(3, n).Formula = '=' + str(x[3])

            y += 1
            n += 1

        # Pulizia righe vuote e segnaposto (Cod. Art.?)
        sStRange = LeenoAnalysis.circoscriveAnalisi(oSheet, lrow)
        sStart = sStRange.RangeAddress.StartRow
        sEnd = sStRange.RangeAddress.EndRow
        for m in reversed(range(sStart, sEnd)):
            # Se la riga è un segnaposto rimasto dal template, la elimina o la svuota
            cell_val = oSheet.getCellByPosition(0, m).String
            if cell_val == 'Cod. Art.?':
                # Se è la riga di una lavorazione, la elimina se sopra c'è lo stile corretto
                if oSheet.getCellByPosition(0, m - 1).CellStyle == 'An-lavoraz-Cod-sx':
                    oSheet.getRows().removeByIndex(m, 1)
                else:
                    oSheet.getCellByPosition(0, m).String = ''

        sStRange = LeenoAnalysis.circoscriveAnalisi(oSheet, lrow)
        startRow = sStRange.RangeAddress.StartRow
        oSheet, startRow = LeenoAnalysis.inizializza_analisi(oDoc, nuovaScheda=True)

    # Pulizia globale delle schede vuote/segnaposto
    rimuoviAnalisiVuote(oSheet)

    # Assicura che ci sia una riga rossa di chiusura corretta
    LeenoSheetUtils.inserisciRigaRossa(oSheet)
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    PL.tante_analisi_in_ep()



def compilaComputo(oDoc, elaborato, capitoliCategorie, elencoPrezzi, listaMisure, indicator=None):
    ''' compila il computo '''
    from datetime import datetime, date

    LeenoUtils.DocumentRefresh(False)

    # --- Creazione/attivazione del foglio ---
    if elaborato == 'VARIANTE':
        if oDoc.getSheets().hasByName('VARIANTE'):
            oSheet = LeenoVariante.generaVariante(oDoc, False)
        else:
            oSheet = LeenoVariante.generaVariante(oDoc, True)
            oSheet.getRows().removeByIndex(2, 4)
    elif elaborato == 'CONTABILITA':
        oSheet = LeenoContab.generaContabilita(oDoc)
    else:
        oSheet = oDoc.getSheets().getByName(elaborato)

    # --- Rimozione della riga vuota iniziale ---
    if oSheet.getCellByPosition(1, 4).String == 'Cod. Art.?':
        oSheet.getRows().removeByIndex(3, 5 if elaborato == 'CONTABILITA' else 4)

    # --- Setup cell range ---
    oCellRangeAddr = CellRangeAddress()
    oCellRangeAddr.Sheet = oSheet.RangeAddress.Sheet

    # --- Mappe e variabili di controllo ---
    mappaVociRighe = {}        # id voce computo → riga foglio
    numeroVoce = 1             # numerazione progressiva voci

    def get_map(key):
        return {str(el.get('id_sc')): el.get('dessintetica', '') for el in capitoliCategorie.get(key, [])}

    mapSuperCategorie = get_map('SuperCategorie')
    mapCategorie = get_map('Categorie')
    mapSottoCategorie = get_map('SottoCategorie')

    testspcat = '0'            # per evitare duplicati supercategoria
    testcat = '0'              # per evitare duplicati categoria
    testsbcat = '0'            # per evitare duplicati sottocategoria

    # Se non c'è un indicatore passato, ne crea uno locale per fallback
    if not indicator:
        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start(f'Compilazione {elaborato}...', len(listaMisure))

    # -------------------------------------------------------------------------
    # Funzioni interne di utilità
    # -------------------------------------------------------------------------

    def insert_categoria_if_needed(idcorr, idtest, reset_test, inser_func, mapping):
        """Inserisce capitolo/sottocapitolo solo se cambia l'id."""
        nonlocal lrow
        if not idcorr or idcorr == '0':
            return

        try:
            if idcorr != idtest:
                reset_test[0] = idcorr
                des = mapping.get(str(idcorr))
                if des:
                    inser_func(oSheet, lrow, des)
                    lrow += 1
        except Exception:
            pass

    def set_num_or_formula(col, row, value):
        """Scrive un numero o formula mantenendo la logica originale."""
        cell = oSheet.getCellByPosition(col, row)
        if value is None or str(value).strip() == '':
            return
        v_strip = str(value).strip()
        try:
            # tenta di inserire come valore numerico
            cell.Value = float(v_strip.replace(',', '.'))
        except (ValueError, TypeError):
            # se non è numerico, controlla se assomiglia a una formula
            if any(c in v_strip for c in '+-*/()'):
                cell.Formula = ('=' + v_strip).replace('=-', '=')
            elif col in (5, 6, 7, 8):
                pass
            else:
                cell.String = v_strip

    def handle_negatives(startRow, vedi_neg=False, current_mis=None):
        """
        Gestione del segno negativo per i componenti della quantità.
        """
        try:
            if current_mis and (len(current_mis) > 7 and '-' in str(current_mis[7])) or vedi_neg:
                # Elabora colonne da 4 (E) a 8 (I) per renderle positive
                for x in range(4, 9):
                    cell = oSheet.getCellByPosition(x, startRow)
                    formula = cell.Formula
                    if formula and str(formula).startswith('='):
                        if not formula.startswith('=ABS('):
                            cell.Formula = '=ABS(' + formula[1:] + ')'
                    elif cell.Value < 0:
                        cell.Value = abs(cell.Value)

                # Inverte una sola volta
                LeenoSheetUtils.invertiUnSegno(oSheet, startRow)
        except Exception:
            pass

    # -------------------------------------------------------------------------
    # CICLO PRINCIPALE
    # -------------------------------------------------------------------------

    total_misure = len(listaMisure)
    for i, el in enumerate(listaMisure):
        # Scaling progress: 80% -> 100%
        if indicator:
            # indicator.Value = 80 + int((i / total_misure) * 20)
            indicator.Value = int((i / total_misure) * 100)
            # indicator.Text = f"Compilazione {elaborato}: {el.get('id_ep')}"
            indicator.Text = f"Compilazione {elaborato}: {i+1}/{total_misure}"

        datamis = el.get('datamis')
        idspcat = el.get('idspcat')
        idcat = el.get('idcat')
        idsbcat = el.get('idsbcat')

        # trova la prima riga libera
        lrow = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1

        # --- Categorie ---
        insert_categoria_if_needed(idspcat, testspcat, [idspcat, None], LeenoSheetUtils.inserSuperCapitolo, mapSuperCategorie)
        testspcat = idspcat
        insert_categoria_if_needed(idcat, testcat, [idcat, None], LeenoSheetUtils.inserCapitolo, mapCategorie)
        testcat = idcat
        insert_categoria_if_needed(idsbcat, testsbcat, [idsbcat, None], LeenoSheetUtils.inserSottoCapitolo, mapSottoCategorie)
        testsbcat = idsbcat

        # --- Inserimento voce ---
        if elaborato == 'CONTABILITA':
            LeenoContab.insertVoceContabilita(oSheet, lrow)
        else:
            LeenoComputo.insertVoceComputoGrezza(oSheet, lrow)

        ID = el.get('id_ep')
        try:
            tariffa = elencoPrezzi['DizionarioArticoli'].get(ID).get('tariffa')
            oSheet.getCellByPosition(1, lrow + 1).String = tariffa
        except:
            pass

        idVoceComputo = el.get('id_vc')
        mappaVociRighe[idVoceComputo] = lrow + 1
        oSheet.getCellByPosition(0, lrow + 1).String = str(numeroVoce)
        numeroVoce += 1

        # --- Misurazioni ---
        startRow = lrow + 2
        lista_righe = el.get('lista_rig')
        nrighe = len(lista_righe)

        if nrighe > 0:
            endRow = startRow + nrighe
            if nrighe > 1:
                oSheet.getRows().insertByIndex(startRow + 1, nrighe - 1)

            oRangeAddress = oSheet.getCellRangeByPosition(0, startRow, 250, startRow).getRangeAddress()
            for n in range(startRow + 1, endRow):
                oSheet.copyRange(oSheet.getCellByPosition(0, n).getCellAddress(), oRangeAddress)
                if elaborato == 'CONTABILITA':
                    c = oSheet.getCellByPosition(1, n)
                    c.String = ''
                    c.CellStyle = 'Comp-Bianche in mezzo_R'

            # Data contabilita
            if elaborato == 'CONTABILITA' and datamis:
                cdata = oSheet.getCellByPosition(1, startRow)
                d_parts = datamis.split('/')
                if len(d_parts) == 3:
                    try:
                        cdata.FormulaLocal = '=DATA(' + d_parts[2] + ';' + d_parts[1] + ';' + d_parts[0] + ')'
                    except:
                        cdata.Formula = '=DATE(' + d_parts[2] + ',' + d_parts[1] + ',' + d_parts[0] + ')'
                    cdata.Value = cdata.Value

            # Popola righe
            for mis in lista_righe:
                descrizione = (mis[0].strip() if mis[0] else '')
                oSheet.getCellByPosition(2, startRow).String = descrizione

                set_num_or_formula(5, startRow, mis[3])  # parti uguali
                set_num_or_formula(6, startRow, mis[4])  # lunghezza
                set_num_or_formula(7, startRow, mis[5])  # larghezza
                set_num_or_formula(8, startRow, mis[6])  # HPESO

                if mis[8] == '2':
                    PL.parziale_core(oSheet, startRow)
                    if elaborato != 'CONTABILITA':
                        oSheet.getRows().removeByIndex(startRow + 1, 1)

                is_vedi_neg = False
                if mis[9] != '-2':
                    vedi = mappaVociRighe.get(mis[9])
                    if vedi:
                        try:
                            test = PL.vedi_voce_xpwe(oSheet, startRow, vedi)
                            if test == '-':
                                is_vedi_neg = True
                        except Exception:
                            pass

                handle_negatives(startRow, vedi_neg=is_vedi_neg, current_mis=mis)
                startRow += 1

    # Finalizzazione
    LeenoSheetUtils.numeraVoci(oSheet, 0, True)
    try:
        PL.Rinumera_TUTTI_Capitoli2(oSheet)
    except:
        pass
    PL.fissa()

def MENU_XPWE_import(filename = None):
    # with LeenoUtils.DocumentRefreshContext(False):
    XPWE_import(filename = None)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)

@LeenoUtils.no_refresh
def XPWE_import(filename = None):
    '''
    Importazione dati dal formato XPWE
    '''
    oDoc = LeenoUtils.getDocument()
    isLeenoDoc = LeenoUtils.isLeenoDocument()
    if isLeenoDoc == False:
        PL.creaComputo(0)
        isLeenoDoc = LeenoUtils.isLeenoDocument()

    # legge i totali dal documento
    if isLeenoDoc:
        oDoc = LeenoUtils.getDocument()
        vals = []
        for el in ("COMPUTO", "VARIANTE", "CONTABILITA"):
            try:
                vals.append(oDoc.getSheets().getByName(el).getCellRangeByName('A2').Value)
            except Exception:
                vals.append(None)
    else:
        vals = [None, None, None]

    # sceglie il tipo di dati da importare
    ordina = LeenoConfig.Config().read('Importazione', 'ordina_computo') == '1'
    elabdest = DLG.ScegliElabDest(
        Title="Importa dal formato XPWE",
        AskTarget=isLeenoDoc,
        AskSort=True,
        Sort=ordina,
        ValComputo=vals[0],
        ValVariante=vals[1],
        ValContabilita=vals[2]
    )
    if elabdest is None:
        return

    elaborato = elabdest['elaborato']
    destinazione = elabdest['destinazione']
    ordina = elabdest['ordina']
    LeenoConfig.Config().write('Importazione', 'ordina_computo', '1' if ordina else '0')

    if elaborato in ('Elenco', 'CONTABILITA'):
        ordina = False

    if filename == None:
        filename = Dialogs.FileSelect('Scegli il file XPWE da importare...', '*.xpwe')
    if filename in ('Cancel', '', None):
        return

    indicator = None
    try:
        # effettua il parsing del file XML
        tree = ElementTree()
        try:
            tree.parse(filename)
        except ParseError:
            PL.clean_text_file(filename)
            tree.parse(filename)
        except PermissionError:
            Dialogs.Exclamation(Title="Errore",
                                Text="Impossibile leggere il file\n"
                                     "Accertati che il nome del file sia corretto.")
            return
        except Exception as e:
            Dialogs.Exclamation(Title="Errore",
                                Text=f"Errore nella lettura del file XML:\n{str(e)}\n\n"
                                     "Accertati che il file sia in formato XPWE valido.")
            return

        # ottieni l'item root
        root = tree.getroot()

        # attiva la progressbar
        indicator = oDoc.getCurrentController().getStatusIndicator()
        if indicator:
            indicator.start("Elaborazione in corso...", 100)
            indicator.Text = "Importazione file XPWE in corso..."
            indicator.Value = 20

        # va alla sezione dei dati generali
        dati = root.find('PweDatiGenerali')
        if dati == None:
            dati = list(root)[0].find('PweDatiGenerali')

        # legge i dati anagrafici generali
        if indicator:
            indicator.Text = "Lettura dati..."
            indicator.Value = 25
        datiAnagrafici = leggiAnagraficaGenerale(dati)

        # legge capitoli e categorie
        if indicator:
            indicator.Text = "Lettura Capitoli e Categorie..."
            indicator.Value = 30
        capitoliCategorie = leggiCapitoliCategorie(dati)

        # legge i dati generali per l'analisi
        if indicator:
            indicator.Text = "Lettura dati generali per analisi..."
            indicator.Value = 40
        datiGeneraliAnalisi = leggiDatiGeneraliAnalisi(dati)

        # legge le approssimazioni
        approssimazioni = leggiApprossimazioni(dati)

        misurazioni = root.find('PweMisurazioni')
        if misurazioni == None:
            misurazioni = list(root)[0].find('PweMisurazioni')

        # legge l'elenco prezzi
        if indicator:
            indicator.Text = "Lettura Elenco Prezzi..."
            indicator.Value = 45
        elencoPrezzi = leggiElencoPrezzi(misurazioni)

        # legge le misurazioni
        if elaborato != 'Elenco':
            if indicator:
                indicator.Text = "Lettura Misurazioni..."
                indicator.Value = 50
            listaMisure = leggiMisurazioni(misurazioni, ordina)
        else:
            listaMisure = []

        # SCRITTURA COMPUTO
        if destinazione == 'NUOVO':
            oDoc = PL.creaComputo(0)
            indicator.end()
            indicator = oDoc.getCurrentController().getStatusIndicator()
            indicator.start("Scrittura dati nel nuovo documento...", 100)

        # Progress focus shift to COMPILATION (50% -> 100%)
        if indicator:
            indicator.Text = "Compilazione dati generali..."
            indicator.Value = 55

        compilaDatiGeneraliAnalisi(oDoc, datiGeneraliAnalisi)
        compilaApprossimazioni(oDoc, approssimazioni)
        compilaAnagraficaGenerale(oDoc, datiAnagrafici)

        # compilo Elenco Prezzi
        if indicator:
            indicator.Text = "Compilazione Elenco Prezzi..."
            indicator.Value = 60
        compilaElencoPrezzi(oDoc, capitoliCategorie, elencoPrezzi, indicator=indicator)

        oSheet_ep = oDoc.getSheets().getByName('Elenco Prezzi')
        oSheet_ep.getCellRangeByName('E2').Formula = '=COUNT(E:E) & " prezzi"'

        # Compilo Analisi di prezzo (60% -> 80% internally)
        compilaAnalisiPrezzi(oDoc, elencoPrezzi, indicator=indicator)

        if len(listaMisure) == 0:
            if indicator:
                indicator.Value = 100
            Dialogs.Info(Title="Importazione completata",
                         Text="Importate n." +
                                str(len(elencoPrezzi['ListaArticoli'])) +
                                " voci dall'elenco prezzi\ndel file: " + filename)
            oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
            oDoc.CurrentController.setActiveSheet(oSheet)
            return

        # compila il computo (80% -> 100% internally)
        compilaComputo(oDoc, elaborato, capitoliCategorie, elencoPrezzi, listaMisure, indicator=indicator)

        PL.GotoSheet(elaborato)
        oSheet_dest = oDoc.getSheets().getByName(elaborato)
        PL.Rinumera_TUTTI_Capitoli2(oSheet_dest)

        # salva il file
        if len(oDoc.getURL()) == 0:
            dest = filename[0:-5] + '.ods'
            PL.salva_come(dest)

        PL.inizializza_computo()
        Dialogs.Ok(Text=f'Importazione di {len(listaMisure)} voci di {elaborato} eseguita con successo!')
        if 'giuserpe' not in os.getlogin():
            PL.dlg_donazioni()

    except Exception as e:
        Dialogs.Exclamation(Title="Errore Critico",
                            Text=f"L'importazione si è interrotta a causa di un errore imprevisto ({type(e).__name__}):\n{str(e)}")
        logging.error(f"Errore XPWE_import: {str(e)}", exc_info=True)
    finally:
        if indicator:
            indicator.end()
