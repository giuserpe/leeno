"""
    LeenO - Indici di costo TOL (ISTAT)

    Recupero degli Indici di costo per Tipologie Omogenee di Lavorazioni (TOL)
    dal servizio SDMX ISTAT, per il calcolo dello stato di avanzamento lavori
    secondo l'art. 60 commi 4 e 4-quater del D.Lgs. 36/2023.

    Dataflow: IT1,145_362_DCSC_INDICITOL_1,1.0
    Endpoint: https://esploradati.istat.it/SDMXWS/rest/data/

    Chiave SDMX (ordine fisso da DataStructure DCSC_INDICITOL):
        FREQ.REF_AREA.DATA_TYPE.ADJUSTMENT.HOM_TYPE_WORK
        M   .IT      .INDTOL_CST_2022.N   .<codice TOL>

    Le prime quattro dimensioni sono costanti per questo dataflow (verificato
    su estrazione 2022-01/2026-05: REF_AREA=IT, DATA_TYPE=INDTOL_CST_2022,
    ADJUSTMENT=N). L'unica dimensione variabile e' HOM_TYPE_WORK, il codice
    della categoria TOL (vedi TOL_CODICI sotto, da codelist CL_COSTO_GRUPCATEG).

    Limite del servizio: 5 richieste al minuto per IP, blocco di 1-2 giorni
    in caso di superamento. Per questo motivo ogni chiamata e' mediata da
    una cache locale su file (.config/leeno/cache_indici_tol.json, stessa
    cartella usata da LeenoConfig per leeno.conf) e da un throttle minimo
    tra le richieste effettive al servizio.
"""

import os
import json
import time
import urllib.request
import urllib.error
from datetime import datetime, timedelta

import LeenoUtils
import LeenoDialogs as DLG
import Dialogs

# https://esploradati.istat.it/databrowser/#/it/dw/categories/IT1,Z0400PRI,1.0/DCSC_INDICITOL/IT1,145_362_DCSC_INDICITOL_1,1.0
DATAFLOW = 'IT1,145_362_DCSC_INDICITOL_1,1.0'
BASE_URL = 'https://esploradati.istat.it/SDMXWS/rest/data/'

# Dimensioni costanti della chiave (verificate su estrazione dataflow.xml/all.csv)
_FREQ = 'M'
_REF_AREA = 'IT'
_DATA_TYPE = 'INDTOL_CST_2022'
_ADJUSTMENT = 'N'

# CL_COSTO_GRUPCATEG - codici HOM_TYPE_WORK per le 20 Tipologie Omogenee di Lavorazioni
TOL_CODICI = {
    1: ('BUIL_ART_NOHEPROT', "Opere edili su edifici e manufatti non soggetti a tutela dei beni culturali"),
    2: ('BUIL_ART_HEPROT', "Opere edili su edifici e manufatti soggetti a tutela dei beni culturali"),
    3: ('ARC_RES_HE', "Scavi archeologici, restauri specialistici di beni culturali e di interesse storico"),
    4: ('EARTDEM_ENV_NAT', "Movimento terra, demolizioni, opere di protezione ambientale, ingegneria naturalistica e opere a verde"),
    5: ('PAV_BITCONG', "Pavimentazioni in conglomerato bituminoso"),
    6: ('STSTRU_ENGWOART', "Strutture, opere di ingegneria e manufatti in acciaio"),
    7: ('REINFCONC_ENGWOART', "Strutture, opere di ingegneria e manufatti in calcestruzzo armato, anche prefabbricato"),
    8: ('WOSTR_ENGWOART', "Strutture, opere di ingegneria e manufatti in legno"),
    9: ('TUN_UND_TRAD', "Gallerie e opere d'arte nel sottosuolo realizzate con metodo tradizionale"),
    10: ('TUN_UND_MECH', "Gallerie e opere d'arte nel sottosuolo realizzate con metodo meccanizzato"),
    11: ('AC_GAS_IRR_SEW', "Acquedotti, gasdotti, opere di irrigazione e fognature"),
    12: ('MAR_DRED_RIV_PROT', "Opere marittime e lavori di dragaggio, opere fluviali e di difesa del suolo"),
    13: ('PLNT_PTD_EEL_HMV_TRPULI', "Impianti per produzione/trasformazione/distribuzione energia elettrica AT/MT, trazione elettrica e illuminazione pubblica"),
    14: ('SYS_EL_TEC_RA_TEL_INT', "Impianti elettrici, tecnologici, radiotelefonici e antintrusione"),
    15: ('SYS_MEC_TER_AIR_WAT_SAN_CON', "Impianti meccanici, termici, di condizionamento, idrico sanitari e trasportatori"),
    16: ('PLNT_WAT_PUR_WSTTRE', "Impianti di potabilizzazione e depurazione"),
    17: ('SYS_SUG_TRASAF_TEL', "Impianti di segnalamento, sicurezza del traffico e telecomunicazioni"),
    18: ('RAIL_INFR', "Armamento ferroviario"),
    19: ('SPEC_FOUN_GEOL_GEOTH', "Opere di fondazione speciale, indagini geologiche e geotecniche"),
    20: ('DELIV_WASTE_FAC_DIS_REC', "Conferimento rifiuti a impianto di smaltimento o recupero"),
}
_TOL_BY_CODE = {codice: numero for numero, (codice, _desc) in TOL_CODICI.items()}

_CACHE_PATH = None
_MIN_INTERVAL_SEC = 13.0  # 5 richieste/minuto = 1 ogni 12s, margine di sicurezza
_last_request_ts = 0.0
_CACHE_TTL = timedelta(hours=24)


# ---------------------------------------------------------------------------
# Cache locale su file
# ---------------------------------------------------------------------------

def _cache_path():
    global _CACHE_PATH
    if _CACHE_PATH is None:
        if os.name == 'nt':
            base = os.path.join(os.getenv('APPDATA', ''), '.config', 'leeno')
        else:
            base = os.path.join(os.getenv('HOME', ''), '.config', 'leeno')
        if not os.path.exists(base):
            os.makedirs(base)
        _CACHE_PATH = os.path.join(base, 'cache_indici_tol.json')
    return _CACHE_PATH


def _load_cache():
    path = _cache_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def _save_cache(cache):
    with open(_cache_path(), 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Accesso al servizio SDMX ISTAT
# ---------------------------------------------------------------------------

def _rispetta_rate_limit():
    global _last_request_ts
    trascorso = time.time() - _last_request_ts
    if trascorso < _MIN_INTERVAL_SEC:
        time.sleep(_MIN_INTERVAL_SEC - trascorso)
    _last_request_ts = time.time()


def _codice_sdmx(tol):
    """
    tol { int|str } : numero TOL (1-20) oppure codice SDMX testuale
                       (es. 'AC_GAS_IRR_SEW').
    Ritorna sempre il codice SDMX testuale (HOM_TYPE_WORK).
    """
    if isinstance(tol, int) or (isinstance(tol, str) and tol.strip().isdigit()):
        numero = int(tol)
        if numero not in TOL_CODICI:
            raise ValueError("TOL %d inesistente (intervallo valido 1-20)" % numero)
        return TOL_CODICI[numero][0]

    codice = str(tol).strip()
    if codice not in _TOL_BY_CODE:
        raise ValueError("Codice TOL sconosciuto: %r" % tol)
    return codice


def _scarica_serie(codice_sdmx):
    """
    Scarica dal servizio SDMX ISTAT la serie storica mensile (CSV) per
    la categoria TOL richiesta. Ritorna una lista di tuple
    (periodo 'AAAA-MM', valore float) ordinata per periodo.
    """
    chiave = '%s.%s.%s.%s.%s' % (
        _FREQ, _REF_AREA, _DATA_TYPE, _ADJUSTMENT, codice_sdmx
    )
    url = '%s%s/%s?format=csv' % (BASE_URL, DATAFLOW, chiave)

    _rispetta_rate_limit()

    richiesta = urllib.request.Request(url, headers={'User-Agent': 'LeenO-pyleeno'})
    try:
        with urllib.request.urlopen(richiesta, timeout=20) as risposta:
            testo = risposta.read().decode('utf-8-sig')
    except urllib.error.HTTPError as e:
        raise RuntimeError("ISTAT SDMX ha risposto %s per la chiave %s" % (e.code, chiave)) from e
    except urllib.error.URLError as e:
        raise RuntimeError("Impossibile raggiungere il servizio ISTAT SDMX: %s" % e.reason) from e

    righe = [r for r in testo.splitlines() if r.strip()]
    if len(righe) < 2:
        raise RuntimeError("Risposta ISTAT vuota per la chiave %s" % chiave)

    intestazione = righe[0].split(',')
    try:
        idx_periodo = intestazione.index('TIME_PERIOD')
        idx_valore = intestazione.index('OBS_VALUE')
    except ValueError as e:
        raise RuntimeError(
            "Formato CSV ISTAT inatteso (colonne mancanti): %s" % intestazione
        ) from e

    serie = []
    for riga in righe[1:]:
        campi = riga.split(',')
        periodo = campi[idx_periodo]
        grezzo = campi[idx_valore]
        if not grezzo:
            continue
        try:
            serie.append((periodo, float(grezzo)))
        except ValueError:
            continue

    serie.sort(key=lambda t: t[0])
    return serie


def get_indice_tol(tol, periodo=None, forza_aggiornamento=False):
    """
    tol      { int|str } : numero TOL (1-20) o codice SDMX testuale
    periodo  { str }      : 'AAAA-MM'; se None, ultimo periodo disponibile
    forza_aggiornamento { bool } : ignora la cache e ricontatta ISTAT

    Ritorna una tupla (periodo_effettivo, valore).
    Solleva RuntimeError/ValueError in caso di errore o dato non trovato.
    """
    codice_sdmx = _codice_sdmx(tol)
    cache = _load_cache()
    voce = cache.get(codice_sdmx, {})

    scaduta = True
    if voce.get('scaricato_il'):
        eta = datetime.now() - datetime.fromisoformat(voce['scaricato_il'])
        scaduta = eta > _CACHE_TTL

    if forza_aggiornamento or scaduta or not voce.get('serie'):
        serie = _scarica_serie(codice_sdmx)
        voce = {
            'serie': serie,
            'scaricato_il': datetime.now().isoformat(),
        }
        cache[codice_sdmx] = voce
        _save_cache(cache)

    serie = voce['serie']
    if not serie:
        raise RuntimeError("Nessun dato disponibile per TOL %s" % codice_sdmx)

    if periodo is None:
        ultimo = serie[-1]
        return (ultimo[0], ultimo[1])

    for p, v in serie:
        if p == periodo:
            return (p, v)

    raise RuntimeError(
        "Periodo %s non disponibile per TOL %s (ultimo disponibile: %s)"
        % (periodo, codice_sdmx, serie[-1][0])
    )


# ---------------------------------------------------------------------------
# Punto di ingresso da menu/toolbar
# ---------------------------------------------------------------------------

def MENU_leeno_aggiorna_indice_tol():
    """
    Punto di ingresso da menu/toolbar.

    Legge il numero TOL (1-20) e il periodo (AAAA-MM, opzionale) dalle
    due celle immediatamente a sinistra della cella attiva; scrive il
    valore dell'indice nella cella attiva.

    Layout atteso sulla riga (relativo alla cella attiva C):
        [colonna C-2: n. TOL]   [colonna C-1: periodo AAAA-MM o vuota]   [C: valore indice]

    Se il periodo e' vuoto, viene usato l'ultimo disponibile e il periodo
    effettivo viene mostrato in un messaggio (per evitare letture
    silenziosamente disallineate rispetto al periodo contrattuale).
    """
    oDoc = LeenoUtils.resolve_document()
    if oDoc is None:
        Dialogs.messageBox(
            text="Documento non disponibile.",
            title="Indice TOL",
            msg_type=Dialogs.ERRORBOX,
        )
        return

    selezione = oDoc.CurrentSelection
    try:
        indirizzo = selezione.CellAddress
    except AttributeError:
        Dialogs.messageBox(
            text="Seleziona una singola cella, non un intervallo.",
            title="Indice TOL",
            msg_type=Dialogs.WARNINGBOX,
        )
        return

    foglio = oDoc.Sheets.getByIndex(indirizzo.Sheet)
    colonna = indirizzo.Column
    riga = indirizzo.Row

    if colonna < 2:
        Dialogs.messageBox(
            text="Servono almeno due colonne a sinistra della cella attiva "
                 "(n. TOL e periodo).",
            title="Indice TOL",
            msg_type=Dialogs.WARNINGBOX,
        )
        return

    cella_tol = foglio.getCellByPosition(colonna - 2, riga)
    cella_periodo = foglio.getCellByPosition(colonna - 1, riga)
    cella_valore = foglio.getCellByPosition(colonna, riga)

    tol_grezzo = cella_tol.getString().strip()
    periodo_grezzo = cella_periodo.getString().strip() or None

    if not tol_grezzo:
        Dialogs.messageBox(
            text="Indica il numero TOL (1-20) nella cella a sinistra.",
            title="Indice TOL",
            msg_type=Dialogs.WARNINGBOX,
        )
        return

    try:
        periodo_effettivo, valore = get_indice_tol(tol_grezzo, periodo_grezzo)
    except (ValueError, RuntimeError) as e:
        Dialogs.messageBox(
            text=str(e),
            title="Indice TOL - errore",
            msg_type=Dialogs.ERRORBOX,
        )
        return

    cella_valore.setValue(valore)

    if periodo_grezzo is None:
        Dialogs.messageBox(
            text="Periodo non indicato: applicato l'ultimo disponibile (%s)."
                 % periodo_effettivo,
            title="Indice TOL",
            msg_type=Dialogs.INFOBOX,
        )
