#!/usr/bin/env python3
# -*- Mode: Python; coding: utf-8; indent-tabs-mode: nil; tab-width: 4 -*-
########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
# Import degli appunti di cantiere (JSON schema v1) nel foglio GIORNALE.
# Schema: documentazione/schemi/giornale_appunti.schema.json
########################################################################
import datetime
import json
import os
import re
import zipfile

import uno

import LeenoGiornale
import LeenoUtils
import SheetUtils

SCHEMA_VERSION = '1'
ORIGINE = 'appunti-mobile'

# chiave JSON -> etichetta in colonna A del foglio GIORNALE.
# Deve coincidere con x-leeno-etichetta nello schema JSON.
ETICHETTE = {
    'meteo': 'Meteo:',
    'presenti': 'Presenti/intervenuti:',
    'annotazioni': 'Annotazioni, attività svolte:',
    'operai': 'Qualifica e n. operai:',
    'attrezzature': 'Attrezzature impiegate:',
    'provviste': 'Provviste:',
    'rifiuti': 'Rifiuto di materiali e/o manufatti:',
    'disposizioni': 'Disposizioni e ordini di servizio del R.U.P. e del D.L.:',
    'relazione_rup': 'Relazione indirizzata al R.U.P.:',
    'verbali': 'Verbali di accertamento e prove:',
    'contestazioni': 'Contestazioni, sospensioni e riprese lavori:',
    'varianti': 'Varianti disposte, modifiche e/o aggiunte prezzi:',
    'infortuni': 'Evento infortunistico:',
    'osservazioni': 'Osservazioni, prescrizioni, avvertenze della D.L.:',
}
_CHIAVI_PER_ETICHETTA = {v: k for k, v in ETICHETTE.items()}


def _valida_pacchetto(dati):
    '''
    Valida la struttura già deserializzata (da .json o da dentro uno .zip).
    Restituisce {datetime.date: {chiave: testo}}. Solleva ValueError.
    '''
    if not isinstance(dati, dict):
        raise ValueError('Il file non contiene un oggetto JSON.')
    if dati.get('schema_version') != SCHEMA_VERSION:
        raise ValueError(
            f"Versione dello schema non supportata: {dati.get('schema_version')!r} "
            f"(attesa {SCHEMA_VERSION!r}).")
    if dati.get('origine') != ORIGINE:
        raise ValueError('Il file non è un export degli appunti di cantiere per LeenO.')
    giornate = dati.get('giornate')
    if not isinstance(giornate, list):
        raise ValueError("Manca l'elenco 'giornate'.")

    risultato = {}
    for n, g in enumerate(giornate, 1):
        if not isinstance(g, dict) or not isinstance(g.get('campi'), dict):
            raise ValueError(f'Giornata n. {n}: struttura non valida.')
        try:
            data = datetime.date.fromisoformat(str(g.get('data')))
        except ValueError:
            raise ValueError(f"Giornata n. {n}: data non valida ({g.get('data')!r}).")
        if data in risultato:
            raise ValueError(f'Data duplicata nel file: {data.isoformat()}.')
        campi = {}
        for chiave, testo in g['campi'].items():
            if chiave not in ETICHETTE:
                continue  # chiave sconosciuta: ignorata
            if not isinstance(testo, str):
                raise ValueError(f'Giornata {data.isoformat()}: il campo {chiave!r} non è testo.')
            if testo.strip():
                campi[chiave] = testo
        risultato[data] = campi
    return risultato


def leggi_appunti(percorso):
    '''
    Legge e valida un file .json semplice.
    Restituisce la lista [(datetime.date, {chiave: testo})] ordinata per data.
    Solleva ValueError con un messaggio leggibile se il file non è valido.
    '''
    try:
        with open(percorso, encoding='utf-8-sig') as f:
            dati = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        raise ValueError(f'File non leggibile: {e}')
    return sorted(_valida_pacchetto(dati).items())


_RE_CARTELLA_FOTO = re.compile(r'^(\d{4})(\d{2})(\d{2})/([^/]+)$')


def leggi_pacchetto(percorso):
    '''
    Legge un file .json o .zip esportato dal Brogliaccio.
    Restituisce (giornate, foto_per_giorno):
      giornate      : lista [(datetime.date, {chiave: testo})], come leggi_appunti()
      foto_per_giorno : {datetime.date: [(nome_file, contenuto_bytes), ...]},
                        vuoto per un .json semplice o uno .zip senza cartelle foto.
    Solleva ValueError con un messaggio leggibile se il file non è valido.
    '''
    if not percorso.lower().endswith('.zip'):
        return leggi_appunti(percorso), {}

    try:
        z = zipfile.ZipFile(percorso)
    except (OSError, zipfile.BadZipFile) as e:
        raise ValueError(f'File compresso non leggibile: {e}')

    nomi_json = [n for n in z.namelist() if n.endswith('.json') and '/' not in n]
    if len(nomi_json) != 1:
        raise ValueError('Il file compresso deve contenere un solo file .json alla radice.')
    try:
        dati = json.loads(z.read(nomi_json[0]).decode('utf-8-sig'))
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        raise ValueError(f'File non leggibile: {e}')
    giornate = sorted(_valida_pacchetto(dati).items())

    foto_per_giorno = {}
    for nome in z.namelist():
        m = _RE_CARTELLA_FOTO.match(nome)
        if not m:
            continue
        data = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        foto_per_giorno.setdefault(data, []).append((m.group(4), z.read(nome)))
    return giornate, foto_per_giorno


def _estrai_foto(oDoc, foto_per_giorno):
    '''
    Scrive le foto su disco accanto al documento, in una sottocartella
    "foto_giornale/<AAAAMMGG>/". Se un file con lo stesso nome e contenuto
    esiste già, non lo riscrive (import ripetuto = nessun duplicato); se il
    nome esiste con contenuto diverso, usa un nome libero senza sovrascrivere.
    Restituisce {datetime.date: percorso_cartella_relativo} solo per le
    giornate per cui è stata effettivamente scritta almeno una foto.
    '''
    url = oDoc.getURL()
    if not url:
        raise ValueError('Salva il documento prima di importare le foto.')
    cartella_doc = os.path.dirname(uno.fileUrlToSystemPath(url))
    risultato = {}
    for data, foto in foto_per_giorno.items():
        sottocartella = data.strftime('%Y%m%d')
        cartella = os.path.join(cartella_doc, 'foto_giornale', sottocartella)
        os.makedirs(cartella, exist_ok=True)
        for nome, contenuto in foto:
            destino = os.path.join(cartella, nome)
            if os.path.exists(destino):
                with open(destino, 'rb') as f:
                    if f.read() == contenuto:
                        continue  # stessa foto già estratta in un import precedente
                base, ext = os.path.splitext(nome)
                n = 2
                while os.path.exists(destino):
                    destino = os.path.join(cartella, f'{base}_{n}{ext}')
                    n += 1
            with open(destino, 'wb') as f:
                f.write(contenuto)
        risultato[data] = 'foto_giornale/' + sottocartella
    return risultato


def _righe_giorni(oSheet):
    '''
    Restituisce (colonna A come lista di stringhe, indici delle righe 'Data:').
    Lettura in blocco di tutta la colonna.
    '''
    ultima = SheetUtils.getLastUsedRow(oSheet)
    colonna = [str(r[0]) for r in oSheet.getCellRangeByPosition(0, 0, 0, ultima).getDataArray()]
    return colonna, [i for i, s in enumerate(colonna) if s.startswith('Data:')]


def _stringa_data(oSheet, data, riga_scratch):
    '''
    Formatta la data come fa nuovo_giorno(): usa il formato della cella in
    colonna B di una riga 'Data:' esistente, che viene poi riportata vuota.
    '''
    oCella = oSheet.getCellByPosition(1, riga_scratch)
    oCella.Value = (data - datetime.date(1899, 12, 30)).days
    testo = oCella.String
    oCella.String = ''
    if not testo or testo.strip().isdigit():
        raise ValueError('Formato data non disponibile nel foglio GIORNALE.')
    return testo


def _confini_blocco(colonna, righe_data, riga_data):
    '''Ultima riga (inclusa) del blocco che inizia a riga_data.'''
    fine = len(colonna) - 1
    for r in righe_data:
        if r > riga_data:
            fine = r - 1
            break
    return fine


def _scrivi_campi(oSheet, riga_data, campi, mancanti):
    '''
    Scrive i campi nel blocco che inizia a riga_data: il testo va nella riga
    sotto l'etichetta. I campi non presenti nel blocco (giornate create con un
    template precedente) vengono registrati in 'mancanti', mai persi in silenzio.
    '''
    colonna, righe_data = _righe_giorni(oSheet)
    fine = _confini_blocco(colonna, righe_data, riga_data)
    posizioni = {}
    for i in range(riga_data, fine + 1):
        chiave = _CHIAVI_PER_ETICHETTA.get(colonna[i])
        if chiave:
            posizioni[chiave] = i
    for chiave, testo in campi.items():
        riga = posizioni.get(chiave)
        if riga is None or riga + 1 > fine or colonna[riga + 1] in _CHIAVI_PER_ETICHETTA:
            mancanti.add(chiave)
            continue
        oSheet.getCellByPosition(0, riga + 1).String = testo


TESTO_LINK_FOTO = 'Apri \u2197'


def _assicura_link_foto(oSheet, riga_data, cartella_relativa):
    '''
    Inserisce (o aggiorna, se già presente da un import precedente) una riga
    con un collegamento alla cartella delle foto del giorno, subito dopo il
    testo di "Annotazioni, attività svolte:". Non fa nulla se quell'etichetta
    non è nel blocco (template precedente senza quel campo).
    '''
    colonna, righe_data = _righe_giorni(oSheet)
    fine = _confini_blocco(colonna, righe_data, riga_data)
    riga_etichetta = None
    for i in range(riga_data, fine + 1):
        if colonna[i] == ETICHETTE['annotazioni']:
            riga_etichetta = i
            break
    if riga_etichetta is None or riga_etichetta + 1 > fine:
        return  # campo assente nel blocco: nessun posto sensato dove mettere il link

    riga_testo = riga_etichetta + 1
    riga_link = riga_testo + 1
    formula = f'=HYPERLINK("{cartella_relativa}/";"{TESTO_LINK_FOTO}")'
    esiste_gia = (riga_link <= fine and str(oSheet.getCellByPosition(0, riga_link).getFormula()).startswith('=HYPERLINK('))
    if not esiste_gia:
        oSheet.getRows().insertByIndex(riga_link, 1)
    oSheet.getCellByPosition(0, riga_link).setFormula(formula)


def importa_giornate(oDoc, giornate, conferma, foto_per_giorno=None):
    '''
    giornate        : lista [(date, campi)] da leggi_appunti()/leggi_pacchetto()
    conferma        : funzione conferma(stringa_data) -> 'si' | 'no' | 'annulla',
                      chiamata solo per le date già presenti in GIORNALE.
    foto_per_giorno : {date: percorso_cartella_relativo}, da _estrai_foto(); opzionale.

    Prima si raccolgono tutte le risposte (dialoghi), poi si scrive nel foglio:
    nessun dialogo modale tra una scrittura e l'altra.
    Restituisce un dizionario con l'esito.
    '''
    foto_per_giorno = foto_per_giorno or {}
    esito = {'annullato': False, 'importate': 0, 'sovrascritte': 0,
             'saltate': 0, 'mancanti': set()}
    if not giornate:
        return esito
    oSheet = oDoc.getSheets().getByName('GIORNALE')
    colonna, righe_data = _righe_giorni(oSheet)
    riutilizza = None
    if righe_data:
        esistenti = {colonna[i]: i for i in righe_data}
    else:
        # giornale nuovo: si crea il primo blocco, che serve anche da cella di
        # formato per le date e viene riusato per la prima giornata importata
        LeenoGiornale.nuovo_giorno()
        colonna, righe_data = _righe_giorni(oSheet)
        riutilizza = righe_data[-1]
        esistenti = {}

    piano = []
    for data, campi in giornate:
        stringa = _stringa_data(oSheet, data, righe_data[-1])
        etichetta = 'Data: ' + stringa
        if etichetta in esistenti:
            risposta = conferma(stringa)
            if risposta == 'annulla':
                esito['annullato'] = True
                return esito
            if risposta != 'si':
                piano.append(('salta', data, etichetta, campi))
                continue
            piano.append(('sovrascrivi', data, etichetta, campi))
        else:
            piano.append(('nuova', data, etichetta, campi))

    for azione, data, etichetta, campi in piano:
        if azione == 'salta':
            esito['saltate'] += 1
            continue
        if azione == 'nuova':
            if riutilizza is None:
                LeenoGiornale.nuovo_giorno()
            _, righe_data = _righe_giorni(oSheet)
            riga = righe_data[-1]
            riutilizza = None
            oSheet.getCellByPosition(0, riga).String = etichetta
            esito['importate'] += 1
        else:
            _, righe_data = _righe_giorni(oSheet)
            riga = next(i for i in righe_data
                        if str(oSheet.getCellByPosition(0, i).String) == etichetta)
            esito['sovrascritte'] += 1
        _scrivi_campi(oSheet, riga, campi, esito['mancanti'])
        cartella = foto_per_giorno.get(data)
        if cartella:
            _assicura_link_foto(oSheet, riga, cartella)
    return esito


def MENU_importa_appunti():
    '''
    Importa nel Giornale Lavori aperto gli appunti di cantiere (file JSON).
    Per ogni giornata già presente chiede se sovrascrivere i campi compilati.
    '''
    import Dialogs  # import locale: evita la catena circolare con pyleeno

    oDoc = LeenoUtils.getDocument()
    if oDoc is None or not (oDoc.getSheets().hasByName('GIORNALE')
                            and oDoc.getSheets().hasByName('GIORNALE_BIANCO')):
        Dialogs.Exclamation(
            Title='Importa appunti di cantiere',
            Text='Apri un Giornale Lavori di LeenO e riprova.')
        return
    percorso = Dialogs.FileSelect('Importa appunti di cantiere...', '*.json;*.zip', 0)
    if not percorso:
        return
    try:
        giornate, foto_per_giorno = leggi_pacchetto(percorso)
    except ValueError as e:
        Dialogs.Exclamation(Title='Importa appunti di cantiere', Text=str(e))
        return

    try:
        cartelle = _estrai_foto(oDoc, foto_per_giorno) if foto_per_giorno else {}
    except ValueError as e:
        Dialogs.Exclamation(Title='Importa appunti di cantiere', Text=str(e))
        return

    def conferma(stringa):
        r = Dialogs.YesNoCancelDialog(
            Title='Giornata già presente',
            Text=f'La giornata {stringa} esiste già nel giornale.\n\n'
                 'Sì: sovrascrive i campi compilati negli appunti\n'
                 'No: salta questa giornata\n'
                 'Annulla: interrompe l\'import')
        return {1: 'si', 0: 'no'}.get(r, 'annulla')

    try:
        esito = importa_giornate(oDoc, giornate, conferma, cartelle)
    except ValueError as e:
        Dialogs.Exclamation(Title='Importa appunti di cantiere', Text=str(e))
        return
    if esito['annullato']:
        Dialogs.Info(Title='Importa appunti di cantiere', Text='Import annullato: nessuna modifica.')
        return
    testo = (f"Giornate aggiunte: {esito['importate']}\n"
             f"Giornate sovrascritte: {esito['sovrascritte']}\n"
             f"Giornate saltate: {esito['saltate']}")
    if cartelle:
        testo += f"\n\nFoto estratte in: foto_giornale/ (accanto al documento)"
    if esito['mancanti']:
        testo += ('\n\nCampi non scritti perché assenti nel giornale '
                  '(template precedente): ' + ', '.join(sorted(esito['mancanti'])))
    Dialogs.Info(Title='Importa appunti di cantiere', Text=testo)
