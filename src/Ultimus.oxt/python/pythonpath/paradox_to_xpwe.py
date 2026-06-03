"""
paradox_to_xpwe.py — Converte archivi Paradox di PriMus anni '90 in XPWE 5.05

Uso:
    python3 paradox_to_xpwe.py <cartella_con_file_paradox> [output.xpwe]

Oppure come modulo:
    from paradox_to_xpwe import paradox_to_xpwe
    paradox_to_xpwe('/path/alla/cartella', '/path/output.xpwe')

Struttura attesa nella cartella:
    _E*.DB / _E*.MB   — Elenco Prezzi (tabella + blob memo)
    _V*.DB / _V*.MB   — Voci Computo  (tabella + blob memo)
    __*.DB / __*.MB   — Dati testata  (tabella + blob memo)

Nota:
    I prezzi sono in LIRE (file anni '90). Dividere per 1936.27 per convertire
    in euro. Le quantità del computo risultano zero: il formato numerico del
    campo quantità non è stato decodificato (formato Currency proprietario).
"""

from __future__ import annotations
import os
import re
import struct
import sys
import datetime
from typing import Any


# ===========================================================================
# LOGGING
# ===========================================================================

LOG_FILE: str = os.path.join(os.path.expanduser('~'), 'paradox_to_xpwe.log')


def _log(*args, sep=' ') -> None:
    msg = sep.join(str(a) for a in args)
    ts  = datetime.datetime.now().strftime('%H:%M:%S.%f')[:-3]
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as fh:
            fh.write(f'[{ts}] {msg}\n')
    except OSError:
        pass


def log_reset() -> None:
    try:
        with open(LOG_FILE, 'w', encoding='utf-8') as fh:
            fh.write(f'# paradox_to_xpwe log — {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n')
    except OSError:
        pass


# ===========================================================================
# LETTURA PARADOX
# ===========================================================================

def _read_alpha(rec: bytes, off: int, size: int) -> str:
    """Legge un campo Alpha Paradox (null-terminated, CP1252, filtra ctrl chars)."""
    raw = rec[off:off+size].rstrip(b'\x00').decode('cp1252', errors='replace').strip()
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', raw)


def _read_number(rec: bytes, off: int) -> float:
    """
    Legge un campo Number Paradox (8 byte, big-endian IEEE 754 con XOR).
    Se il MSB >= 0x80: numero positivo, XOR solo il bit 7 del primo byte.
    Se il MSB <  0x80: numero negativo, XOR tutti i byte con 0xFF.
    """
    raw = bytearray(rec[off:off+8])
    if not any(raw):
        return 0.0
    if raw[0] & 0x80:
        raw[0] ^= 0x80
        try:
            return struct.unpack('>d', bytes(raw))[0]
        except Exception:
            return 0.0
    else:
        for i in range(8):
            raw[i] ^= 0xFF
        try:
            return -struct.unpack('>d', bytes(raw))[0]
        except Exception:
            return 0.0


def _read_long(rec: bytes, off: int) -> int:
    """Legge un campo Long/AutoInc Paradox (4 byte, big-endian XOR 0x80000000)."""
    v = struct.unpack_from('>I', rec, off)[0]
    return (v ^ 0x80000000) if v else 0


def _read_memo(data_mb: bytes, raw10: bytes) -> str:
    """
    Legge un campo FmtMemo Paradox dal file .MB.
    raw10: 10 byte del campo nel record DB.
      byte 0-3: offset assoluto nel file MB (LE u32)
      byte 4-5: modifier (LE u16)
      byte 6-7: lunghezza blob (LE u16)
      byte 8-9: riservato
    """
    if len(raw10) < 10:
        return ''
    blob_off = struct.unpack_from('<I', raw10, 0)[0]
    blob_len = struct.unpack_from('<H', raw10, 6)[0]
    if not blob_off or not blob_len:
        return ''
    if blob_off + blob_len > len(data_mb):
        return ''
    raw = data_mb[blob_off:blob_off+blob_len]
    txt = raw.decode('cp1252', errors='replace').strip()
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', txt)


def _iter_blocks(data: bytes, rec_size: int, hdr_size: int, block_size: int):
    """
    Itera su tutti i record Paradox di un file .DB.
    Ogni blocco ha 4 byte di header ([next_block u16][n_recs u8][?])
    seguiti dai record consecutivi.
    """
    recs_per_block = (block_size - 4) // rec_size
    n_blocks = (len(data) - hdr_size) // block_size
    for blk in range(n_blocks):
        blk_off = hdr_size + blk * block_size
        for r in range(recs_per_block):
            rec_off = blk_off + 4 + r * rec_size
            if rec_off + rec_size > len(data):
                break
            yield data[rec_off:rec_off+rec_size]


def _pdx_params(data: bytes) -> tuple[int, int, int]:
    """
    Restituisce (rec_size, hdr_size, block_size) dal file .DB Paradox.
    block_size = ((max_table_size + 1) * 512) arrotondato al record_size.
    """
    rec_size  = struct.unpack_from('<H', data, 0x00)[0]
    hdr_size  = struct.unpack_from('<H', data, 0x02)[0]
    max_tbl   = data[0x05]
    # block_size: usa la formula standard, poi verifica con file_size
    block_size = (max_tbl + 1) * 512
    # Normalizza: deve essere multiplo di rec_size + 4
    if block_size < rec_size + 4:
        block_size = rec_size + 4 + rec_size  # minimo 2 record
    return rec_size, hdr_size, block_size


# ===========================================================================
# PARSER TABELLA EP  (_E*.DB)
# ===========================================================================

# Layout campo _E (rec_size=486, skip 2 byte iniziali Paradox):
#   s+ 0  Alpha(20)   tariffa
#   s+20  AutoInc(4)  id_autoinc
#   s+24  Alpha(20)   articolo
#   s+44  Alpha(120)  des_ridotta
#   s+164 FmtMemo(10) des_estesa
#   s+174 FmtMemo(10) memo2
#   s+184 FmtMemo(10) memo3
#   s+194 Number(8)   prezzo1
#   s+202 Number(8)   prezzo2
#   s+210 Number(8)   prezzo3
#   s+218 Number(8)   prezzo4
#   s+226 Number(8)   prezzo5
#   s+234 Number(8)   ?
#   s+242 Alpha(6)    um
#   s+248 Alpha(20)   ?
#   s+268 Short(2)*4  id_spcap, id_cap, id_sbcap, ?
#   s+276 Long(4)     ?
#   s+280 Short(2)    ?
#   s+282 Number(8)   inc_mdo
#   s+290 Number(8)   inc_sic
#   ...

_EP_SKIP = 2  # byte interni Paradox da saltare all'inizio di ogni record


def _parse_ep_record(rec: bytes, data_mb: bytes) -> dict[str, Any] | None:
    s = _EP_SKIP
    trf = _read_alpha(rec, s+0, 20)
    # Filtra record vuoti o con tariffa non valida
    if not trf or len(trf) < 2:
        return None
    if not re.match(r'^[\w\.\-\/\s]+$', trf):
        return None

    des_rid = _read_alpha(rec, s+44, 120)
    des_est = _read_memo(data_mb, rec[s+164:s+174]) or des_rid

    return {
        'tariffa':     trf,
        'articolo':    _read_alpha(rec, s+24, 20),
        'des_ridotta': des_rid,
        'des_estesa':  des_est,
        'um':          _read_alpha(rec, s+242, 6),
        'prezzo':      _read_number(rec, s+194),
        'prezzo2':     _read_number(rec, s+202),
        'prezzo3':     _read_number(rec, s+210),
        'prezzo4':     _read_number(rec, s+218),
        'prezzo5':     _read_number(rec, s+226),
        'inc_mdo':     _read_number(rec, s+282) if s+290 <= len(rec) else 0.0,
        'inc_sic':     _read_number(rec, s+290) if s+298 <= len(rec) else 0.0,
    }


def read_ep_table(db_path: str) -> list[dict[str, Any]]:
    """Legge la tabella elenco prezzi (_E*.DB) e restituisce lista voci."""
    data    = open(db_path, 'rb').read()
    mb_path = db_path.replace('.DB', '.MB')
    data_mb = open(mb_path, 'rb').read() if os.path.exists(mb_path) else b''
    rec_size, hdr_size, block_size = _pdx_params(data)
    _log(f'EP: {db_path}  rec_size={rec_size}  hdr_size={hdr_size}  block_size={block_size}')

    by_trf: dict[str, dict] = {}
    for rec in _iter_blocks(data, rec_size, hdr_size, block_size):
        row = _parse_ep_record(rec, data_mb)
        if not row:
            continue
        trf = row['tariffa']
        # Tieni la versione con descrizione più lunga
        if trf not in by_trf or len(row['des_estesa']) > len(by_trf[trf]['des_estesa']):
            by_trf[trf] = row

    # Ri-numera con ID sequenziali ordinati per tariffa
    ep_list: list[dict] = []
    for i, ep in enumerate(sorted(by_trf.values(), key=lambda e: e['tariffa']), 1):
        ep.update({
            'id':           i,
            'des_breve':    '',
            'prezzo_netto': 0.0,
            'inc_mat':      0.0,
            'inc_attr':     0.0,
            'ribassabile':  True,
            'id_spcap':     0,
            'id_cap':       0,
            'id_sbcap':     0,
            'flags':        0,
            'data':         '',
            'adr_internet': '',
            'tag_bim':      '',
        })
        ep_list.append(ep)

    _log(f'EP: {len(ep_list)} voci estratte (dedup su tariffa)')
    return ep_list


# ===========================================================================
# PARSER TABELLA VC  (_V*.DB)
# ===========================================================================

# Layout campo _V (rec_size=116, skip 2 byte iniziali Paradox):
#   s+ 0  Long(4)    id_vc
#   s+ 4  Long(4)    ref_ep (indice ordinale nella tabella EP, 1-based + offset)
#   s+ 8  Number(8)  quantità (spesso vuoto in file molto vecchi)
#   s+16  Number(8)  importo  (spesso vuoto)
#   ...

_VC_SKIP = 2


def read_vc_table(db_path: str, ep_list: list[dict]) -> list[dict[str, Any]]:
    """
    Legge la tabella voci computo (_V*.DB).
    Mappa i riferimenti EP tramite indice ordinale (con auto-rilevamento offset).
    """
    data    = open(db_path, 'rb').read()
    mb_path = db_path.replace('.DB', '.MB')
    data_mb = open(mb_path, 'rb').read() if os.path.exists(mb_path) else b''
    rec_size, hdr_size, block_size = _pdx_params(data)
    _log(f'VC: {db_path}  rec_size={rec_size}  hdr_size={hdr_size}  block_size={block_size}')

    # Leggi ordine di inserimento degli EP nel file (diverso dall'ordine per tariffa)
    ep_db_path  = db_path.replace('_V', '_E')
    ep_data     = open(ep_db_path, 'rb').read()
    ep_mb_path  = ep_db_path.replace('.DB', '.MB')
    ep_data_mb  = open(ep_mb_path, 'rb').read() if os.path.exists(ep_mb_path) else b''
    ep_rec_size, ep_hdr_size, ep_block_size = _pdx_params(ep_data)

    trf_to_id = {ep['tariffa']: ep['id'] for ep in ep_list}
    ep_ordinal: dict[int, int] = {}  # ordine lettura -> ep_id
    idx = 0
    for rec in _iter_blocks(ep_data, ep_rec_size, ep_hdr_size, ep_block_size):
        row = _parse_ep_record(rec, ep_data_mb)
        if not row:
            continue
        idx += 1
        ep_id = trf_to_id.get(row['tariffa'], 0)
        if ep_id:
            ep_ordinal[idx] = ep_id

    # Raccoglie tutti i ref dal file VC
    raw_refs: list[tuple[int, int]] = []  # (id_vc, ref)
    for rec in _iter_blocks(data, rec_size, hdr_size, block_size):
        s = _VC_SKIP
        id_vc = _read_long(rec, s+0)
        ref   = _read_long(rec, s+4)
        if not id_vc:
            continue
        qt  = _read_number(rec, s+8)
        imp = _read_number(rec, s+16)
        raw_refs.append((id_vc, ref, qt, imp))

    # Auto-rileva offset (i ref partono tipicamente da 7 = primo EP reale)
    valid_refs = [ref for _, ref, _, _ in raw_refs if 0 < ref <= len(ep_ordinal) + 20]
    offset = min(valid_refs) - 1 if valid_refs else 0
    _log(f'VC: offset rilevato={offset}  record totali={len(raw_refs)}')

    vc_list: list[dict] = []
    for i, (id_vc, ref, qt, imp) in enumerate(raw_refs, 1):
        ordinal = ref - offset
        ep_id   = ep_ordinal.get(ordinal, 0)
        vc_list.append({
            'id':       i,
            'id_ep':    ep_id,
            'quantita': qt,
            'importo':  imp,
            'data_mis': '',
            'id_spcal': 0,
            'id_cat':   0,
            'id_sbcat': 0,
            'flags':    0,
            'misure':   [],
            'cod_wbs':  '',
        })

    mapped = sum(1 for v in vc_list if v['id_ep'] > 0)
    _log(f'VC: {len(vc_list)} voci  mappate={mapped}')
    return vc_list


# ===========================================================================
# PARSER TESTATA  (__*.DB)
# ===========================================================================

def read_header_table(db_path: str) -> dict[str, str]:
    """
    Legge la tabella di testata (__*.DB).
    Restituisce dict con info del progetto (oggetto, committente, ecc.).
    """
    if not os.path.exists(db_path):
        return {}
    data    = open(db_path, 'rb').read()
    mb_path = db_path.replace('.DB', '.MB')
    data_mb = open(mb_path, 'rb').read() if os.path.exists(mb_path) else b''
    rec_size, hdr_size, block_size = _pdx_params(data)

    info: dict[str, str] = {}
    for rec in _iter_blocks(data, rec_size, hdr_size, block_size):
        s = _EP_SKIP
        # Cerca stringhe significative nei primi campi Alpha
        candidates = []
        for off in range(0, min(rec_size - 2, 300), 20):
            t = _read_alpha(rec, s+off, 20)
            if len(t) > 4 and t.replace(' ', '').isalnum():
                candidates.append(t)
        if candidates:
            if 'oggetto' not in info:
                info['oggetto'] = candidates[0]
            break
    return info


# ===========================================================================
# FUNZIONE PRINCIPALE
# ===========================================================================

def paradox_to_xpwe(folder: str, output_xpwe: str | None = None,
                    lire_to_euro: bool = False) -> str:
    """
    Converte un archivio Paradox PriMus in file XPWE 5.05.

    Parametri:
        folder       : cartella contenente i file .DB / .MB
        output_xpwe  : percorso del file XPWE da generare
                       (default: stessa cartella, nome dedotto dai file)
        lire_to_euro : se True, divide tutti i prezzi per 1936.27

    Restituisce il percorso del file XPWE generato.
    """
    import dcf_parser

    log_reset()
    _log(f'paradox_to_xpwe: folder={folder!r}')

    # Trova i file .DB nella cartella
    files = os.listdir(folder)

    def find_db(prefix: str) -> str | None:
        for f in files:
            if f.upper().startswith(prefix.upper()) and f.upper().endswith('.DB'):
                return os.path.join(folder, f)
        return None

    db_e  = find_db('_E')
    db_v  = find_db('_V')
    db_hdr = find_db('__')

    if not db_e:
        raise FileNotFoundError(f'File _E*.DB non trovato in {folder!r}')

    # Nome progetto dal nome del file
    base_name = os.path.splitext(os.path.basename(db_e))[0]
    # Rimuovi prefisso _E e normalizza
    project_name = re.sub(r'^_E\d*[-_]?', '', base_name).replace('-', ' ').replace('_', ' ')

    if output_xpwe is None:
        output_xpwe = os.path.join(folder, base_name + '.xpwe')

    # Leggi tabelle
    _log(f'Lettura EP da {db_e}')
    ep_list = read_ep_table(db_e)

    vc_list: list[dict] = []
    if db_v and os.path.exists(db_v):
        _log(f'Lettura VC da {db_v}')
        vc_list = read_vc_table(db_v, ep_list)
    else:
        _log('File _V*.DB non trovato — computo vuoto')

    info: dict[str, str] = {}
    if db_hdr and os.path.exists(db_hdr):
        info = read_header_table(db_hdr)
    info.setdefault('oggetto',     project_name)
    info.setdefault('committente', '')
    info.setdefault('comune',      '')
    info.setdefault('provincia',   '')
    info.setdefault('impresa',     '')
    info.setdefault('operatore',   '')
    info['perc_prezzi'] = 0
    info['parte_opera'] = ''

    # Conversione lire -> euro
    if lire_to_euro:
        CAMBIO = 1936.27
        for ep in ep_list:
            ep['prezzo']  = round(ep['prezzo']  / CAMBIO, 2)
            ep['prezzo2'] = round(ep['prezzo2'] / CAMBIO, 2)
        for vc in vc_list:
            if vc['importo']:
                vc['importo'] = round(vc['importo'] / CAMBIO, 2)
        _log(f'Conversione lire->euro applicata (cambio {CAMBIO})')

    doc: dict[str, Any] = {
        'formato':          'paradox',
        'info':             info,
        'quadro_economico': [],
        'super_categorie':  {},
        'categorie':        {},
        'sotto_categorie':  {},
        'super_capitoli':   {},
        'capitoli':         {},
        'sotto_capitoli':   {},
        'elenco_prezzi':    ep_list,
        'computo':          vc_list,
        'strutture_stampa': [],
        '_ep_by_id':        {ep['id']: ep for ep in ep_list},
    }

    dcf_parser.generate_xpwe(doc, output_xpwe)
    _log(f'Scritto: {output_xpwe!r}')
    _log(f'  EP={len(ep_list)}  VC={len(vc_list)}')
    return output_xpwe


# ===========================================================================
# ENTRY POINT
# ===========================================================================

def select_folder_and_convert(lire_to_euro: bool = False) -> 'str | None':
    """
    Apre un dialog per selezionare la cartella Paradox e avvia la conversione.
    Da usare in LeenO / pyleeno.py.

    Prova nell'ordine:
      1. Dialogs.FolderSelect()  — se disponibile in LeenO
      2. Dialogs.FileSelect()    — seleziona un file .DB nella cartella,
                                   poi usa la directory del file
      3. FolderPicker UNO nativo — fallback senza Dialogs

    Restituisce il percorso del file XPWE generato, o None se annullato.
    """
    import Dialogs
    fpath = Dialogs.FileSelect('Seleziona un file .DB della cartella Paradox...', '*.DB')
    if not fpath:
        return None
    folder = os.path.dirname(fpath)

    if not folder:
        _log('select_folder_and_convert: nessuna cartella selezionata')
        return None

    _log(f'select_folder_and_convert: cartella={folder!r}  lire_to_euro={lire_to_euro}')
    try:
        xpwe_path = paradox_to_xpwe(folder, lire_to_euro=lire_to_euro)
    except FileNotFoundError as e:
        _log(f'Errore: {e}')
        try:
            import Dialogs as DLG
            DLG.chi(f'Nessun file Paradox trovato nella cartella selezionata.\n{e}')
        except Exception:
            pass
        return None
    except Exception as e:
        _log(f'Errore conversione: {e}')
        try:
            import Dialogs as DLG
            DLG.chi(f'Errore durante la conversione:\n{e}')
        except Exception:
            pass
        return None

    try:
        from LeenoImport_XPWE import XPWE_import
        _log(f'XPWE_import: {xpwe_path!r}')
        XPWE_import(xpwe_path)
    except Exception as e:
        _log(f'Errore XPWE_import: {e}')
        try:
            import Dialogs as DLG
            DLG.chi(f"Errore durante l'importazione XPWE:\n{e}")
        except Exception:
            pass
        return None

    return xpwe_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    folder  = sys.argv[1]
    outfile = sys.argv[2] if len(sys.argv) > 2 else None
    euro    = '--euro' in sys.argv

    try:
        result = paradox_to_xpwe(folder, outfile, lire_to_euro=euro)
        print(f'Generato: {result}')
        print(f'Log: {LOG_FILE}')
    except Exception as e:
        print(f'Errore: {e}')
        sys.exit(1)