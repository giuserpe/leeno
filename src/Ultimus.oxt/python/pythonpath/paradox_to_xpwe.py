########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

"""
paradox_to_xpwe.py — Converte archivi Paradox di PriMus anni '90 in XPWE 5.05


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
    '''Scrive un messaggio con timestamp nel file di log.'''
    msg = sep.join(str(a) for a in args)
    ts  = datetime.datetime.now().strftime('%H:%M:%S.%f')[:-3]
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as fh:
            fh.write(f'[{ts}] {msg}\n')
    except OSError:
        pass


def log_reset() -> None:
    '''Reinizializza il file di log con una nuova intestazione.'''
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


def _parse_mb_file(data_mb: bytes) -> dict[tuple[int, int], bytes]:
    """
    Analizza un file memo Paradox .MB indicizzato a blocchi.
    Restituisce un dizionario con chiave (block_index, entry_id) e valore bytes del testo.
    """
    if len(data_mb) < 4096:
        return {}
    blk_size = data_mb[12] * 256 if len(data_mb) > 13 and data_mb[12] > 0 else 4096
    if blk_size <= 0:
        blk_size = 4096

    memos: dict[tuple[int, int], bytes] = {}
    n_blocks = len(data_mb) // blk_size
    for b in range(1, n_blocks):
        blk_off = b * blk_size
        max_idx_off = min(blk_off + 336, len(data_mb))
        for p in range(blk_off + 16, max_idx_off - 4):
            chunk_off = data_mb[p]
            chunk_cnt = data_mb[p+1]
            entry_id  = data_mb[p+2] + (data_mb[p+3] << 8)
            if 0 < entry_id < 2000 and 0 < chunk_off < 256 and 0 < chunk_cnt < 256:
                text_off = blk_off + chunk_off * 16
                text_end = text_off + chunk_cnt * 16
                if text_end <= len(data_mb):
                    cand = data_mb[text_off:text_end]
                    if any(32 <= c <= 126 for c in cand[:10]):
                        memos[(b, entry_id)] = cand
    return memos


def _read_memo(data_mb: bytes, raw10: bytes,
               mb_index: dict[tuple[int, int], bytes] | None = None) -> str:
    """
    Legge un campo FmtMemo Paradox dal file .MB.
    raw10: 10 byte del campo nel record DB.
    Supporta sia gli archivi indicizzati per blocchi sia i blob con offset diretto.
    """
    if len(raw10) < 10:
        return ''

    # Prova lettura tramite indice di blocco
    if mb_index:
        b_idx = raw10[1] // 16 if raw10[1] else 0
        eid   = raw10[8] + (raw10[9] << 8)
        mlen  = raw10[4] + (raw10[5] << 8)
        if (b_idx, eid) in mb_index:
            raw = mb_index[(b_idx, eid)]
            txt = raw[:mlen].decode('cp1252', errors='replace').strip() if mlen else (
                raw.split(b'\x00')[0].decode('cp1252', errors='replace').strip()
            )
            return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', txt)

    # Fallback su offset diretto nel file MB
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
    Determina la dimensione dell'header di blocco (4 byte in Paradox 3.5 o 6 byte in Paradox 4.0+)
    ed estrae i record.
    """
    blk_hdr_size = 6 if hdr_size >= 2048 else 4
    recs_per_block = (block_size - blk_hdr_size) // rec_size
    n_blocks = (len(data) - hdr_size) // block_size
    for blk in range(n_blocks):
        blk_off = hdr_size + blk * block_size
        for r in range(recs_per_block):
            rec_off = blk_off + blk_hdr_size + r * rec_size
            if rec_off + rec_size > len(data):
                break
            yield data[rec_off:rec_off+rec_size]


def _pdx_params(data: bytes) -> tuple[int, int, int]:
    """
    Restituisce (rec_size, hdr_size, block_size) dal file .DB Paradox.
    Calcola block_size basandosi sui blocchi totali nel file o sul codice max_table_size.
    """
    rec_size = struct.unpack_from('<H', data, 0x00)[0]
    hdr_size = struct.unpack_from('<H', data, 0x02)[0]
    max_tbl  = data[0x05]
    total_blocks = struct.unpack_from('<H', data, 0x0C)[0] if len(data) >= 0x0E else 0

    if total_blocks > 0 and len(data) > hdr_size and (len(data) - hdr_size) % total_blocks == 0:
        block_size = (len(data) - hdr_size) // total_blocks
    else:
        block_size = (max_tbl + 1) * 512

    if block_size < rec_size + 4:
        block_size = rec_size + 4 + rec_size
    return rec_size, hdr_size, block_size


def _get_pdx_field_offsets(data: bytes, hdr_size: int) -> dict[str, tuple[int, int, int]]:
    """
    Estrae dinamicamente lo schema dei campi Paradox dall'intestazione del file .DB.
    Restituisce un dizionario: nome_campo.lower() -> (offset_nel_record, lunghezza, tipo).
    """
    if len(data) < 0x80:
        return {}
    num_fields = data[0x21] + (data[0x22] << 8)
    if num_fields <= 0 or num_fields > 255:
        return {}

    fspecs = []
    curr = 0x78
    if curr + num_fields * 2 > len(data):
        return {}
    for _ in range(num_fields):
        fspecs.append((data[curr], data[curr+1]))
        curr += 2

    pos = curr
    names: list[str] = []
    while pos < hdr_size - 10:
        cur = pos
        curr_names = []
        for _ in range(num_fields):
            end = data.find(b'\x00', cur, min(cur + 50, hdr_size))
            if end == -1 or end == cur:
                break
            name_bytes = data[cur:end]
            if not name_bytes.replace(b'_', b'').isalnum():
                break
            curr_names.append(name_bytes.decode('latin1', errors='replace'))
            cur = end + 1
        if len(curr_names) == num_fields:
            names = curr_names
            break
        pos += 1

    if len(names) != num_fields:
        return {}

    offsets: dict[str, tuple[int, int, int]] = {}
    cum_off = 0  # I record reali iniziano direttamente con i campi dello schema
    for (ftype, flen), name in zip(fspecs, names):
        offsets[name.lower()] = (cum_off, flen, ftype)
        cum_off += flen
    return offsets


# ===========================================================================
# PARSER TABELLA EP  (_E*.DB)
# ===========================================================================

_EP_SKIP = 2  # byte interni Paradox da saltare all'inizio di ogni record


def _parse_ep_record(rec: bytes, data_mb: bytes,
                     schema: dict[str, tuple[int, int, int]] | None = None,
                     mb_index: dict[tuple[int, int], bytes] | None = None) -> dict[str, Any] | None:
    '''Decodifica un record binario dell'elenco prezzi ed estrae i relativi campi.'''
    if schema and 'tariffa' in schema:
        off_trf, len_trf, _ = schema['tariffa']
        trf = _read_alpha(rec, off_trf, len_trf)
        if not trf or len(trf) < 2 or not re.match(r'^[\w\.\-\/\s]+$', trf):
            return None

        off_art, len_art, _ = schema.get('articolo', (off_trf + len_trf + 4, 20, 1))
        art = _read_alpha(rec, off_art, len_art)

        off_rid, len_rid, _ = schema.get('desridotta', (off_art + len_art, 120, 1))
        des_rid = _read_alpha(rec, off_rid, len_rid)

        des_est = des_rid
        if 'desestesa' in schema:
            off_est, len_est, _ = schema['desestesa']
            memo_txt = _read_memo(data_mb, rec[off_est:off_est+len_est], mb_index=mb_index)
            if memo_txt:
                des_est = memo_txt

        um = ''
        if 'unmisura' in schema:
            off_um, len_um, ftype_um = schema['unmisura']
            if ftype_um in (12, 13):
                um = _read_memo(data_mb, rec[off_um:off_um+len_um], mb_index=mb_index)
            else:
                um = _read_alpha(rec, off_um, len_um)

        prezzo = _read_number(rec, schema['prezzo1'][0]) if 'prezzo1' in schema else 0.0
        prezzo2 = _read_number(rec, schema['prezzo2'][0]) if 'prezzo2' in schema else 0.0
        prezzo3 = _read_number(rec, schema['prezzo3'][0]) if 'prezzo3' in schema else 0.0
        prezzo4 = _read_number(rec, schema['prezzo4'][0]) if 'prezzo4' in schema else 0.0
        prezzo5 = _read_number(rec, schema['prezzo5'][0]) if 'prezzo5' in schema else 0.0
        inc_mdo = _read_number(rec, schema['incmdo'][0]) if 'incmdo' in schema else 0.0
        inc_sic = _read_number(rec, schema['incsic'][0]) if 'incsic' in schema else 0.0

        cuep = _read_long(rec, schema['cuep'][0]) if 'cuep' in schema else 0
        return {
            'tariffa':     trf,
            'cuep':        cuep,
            'articolo':    art,
            'des_ridotta': des_rid,
            'des_estesa':  des_est,
            'um':          um,
            'prezzo':      prezzo,
            'prezzo2':     prezzo2,
            'prezzo3':     prezzo3,
            'prezzo4':     prezzo4,
            'prezzo5':     prezzo5,
            'inc_mdo':     inc_mdo,
            'inc_sic':     inc_sic,
        }

    # Layout di fallback per tabelle con schema statico / test sintetici
    s = _EP_SKIP
    trf = _read_alpha(rec, s+0, 20)
    if not trf or len(trf) < 2:
        return None
    if not re.match(r'^[\w\.\-\/\s]+$', trf):
        return None

    des_rid = _read_alpha(rec, s+44, 120)
    des_est = _read_memo(data_mb, rec[s+164:s+174], mb_index=mb_index) or des_rid

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
    mb_index = _parse_mb_file(data_mb)
    rec_size, hdr_size, block_size = _pdx_params(data)
    schema = _get_pdx_field_offsets(data, hdr_size)
    _log(f'EP: {db_path}  rec_size={rec_size}  hdr_size={hdr_size}  block_size={block_size}  fields={len(schema)}')

    by_trf: dict[str, dict] = {}
    for rec in _iter_blocks(data, rec_size, hdr_size, block_size):
        row = _parse_ep_record(rec, data_mb, schema=schema, mb_index=mb_index)
        if not row:
            continue
        trf = row['tariffa']
        # Tieni la versione con descrizione più completa
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

_VC_SKIP = 2


def read_vc_table(db_path: str, ep_list: list[dict]) -> list[dict[str, Any]]:
    """
    Legge la tabella voci computo (_V*.DB o _B*.DB).
    Mappa i riferimenti EP tramite identificativo relazionale CuEP (o ordine ordinale di fallback).
    """
    data    = open(db_path, 'rb').read()
    mb_path = db_path.replace('.DB', '.MB')
    data_mb = open(mb_path, 'rb').read() if os.path.exists(mb_path) else b''
    mb_index = _parse_mb_file(data_mb)
    rec_size, hdr_size, block_size = _pdx_params(data)
    schema_v = _get_pdx_field_offsets(data, hdr_size)
    _log(f'VC: {db_path}  rec_size={rec_size}  hdr_size={hdr_size}  block_size={block_size}  fields={len(schema_v)}')

    # Mappa CuEP -> ep['id'] per join relazionale diretto
    cuep_to_id: dict[int, int] = {ep['cuep']: ep['id'] for ep in ep_list if ep.get('cuep')}

    # Fallback ordinale per tabelle senza CuEP / test sintetici
    ep_ordinal: dict[int, int] = {}
    if not cuep_to_id:
        ep_db_path = None
        for prefix in ('_V', '_B'):
            if prefix in os.path.basename(db_path):
                ep_db_path = os.path.join(os.path.dirname(db_path), os.path.basename(db_path).replace(prefix, '_E', 1))
                break
        if ep_db_path and os.path.exists(ep_db_path):
            ep_data     = open(ep_db_path, 'rb').read()
            ep_mb_path  = ep_db_path.replace('.DB', '.MB')
            ep_data_mb  = open(ep_mb_path, 'rb').read() if os.path.exists(ep_mb_path) else b''
            ep_mb_index = _parse_mb_file(ep_data_mb)
            ep_rec_size, ep_hdr_size, ep_block_size = _pdx_params(ep_data)
            ep_schema   = _get_pdx_field_offsets(ep_data, ep_hdr_size)
            trf_to_id   = {ep['tariffa']: ep['id'] for ep in ep_list}
            idx = 0
            for rec in _iter_blocks(ep_data, ep_rec_size, ep_hdr_size, ep_block_size):
                row = _parse_ep_record(rec, ep_data_mb, schema=ep_schema, mb_index=ep_mb_index)
                if not row:
                    continue
                idx += 1
                ep_id = trf_to_id.get(row['tariffa'], 0)
                if ep_id:
                    ep_ordinal[idx] = ep_id

    vc_list: list[dict] = []
    if schema_v and 'cuvc' in schema_v and 'cuep' in schema_v:
        off_vc   = schema_v['cuvc'][0]
        off_ep   = schema_v['cuep'][0]
        off_qt   = schema_v.get('quantita', (12, 8, 6))[0]
        off_mis  = schema_v.get('misure', (48, 10, 13))[0] if 'misure' in schema_v else None
        len_mis  = schema_v.get('misure', (48, 10, 13))[1] if 'misure' in schema_v else 10
        off_sp   = schema_v.get('supercategoria', (38, 2, 3))[0] if 'supercategoria' in schema_v else None
        off_cat  = schema_v.get('categoria', (40, 2, 3))[0] if 'categoria' in schema_v else None
        off_sb   = schema_v.get('subcategoria', (42, 2, 3))[0] if 'subcategoria' in schema_v else None
        off_dt   = schema_v.get('datamis', (20, 4, 2))[0] if 'datamis' in schema_v else None

        for rec in _iter_blocks(data, rec_size, hdr_size, block_size):
            id_vc = _read_long(rec, off_vc)
            ref_ep = _read_long(rec, off_ep)
            if not id_vc or not ref_ep:
                continue

            ep_id = cuep_to_id.get(ref_ep, 0)
            if not ep_id:
                continue

            qt = _read_number(rec, off_qt)
            misure_txt = _read_memo(data_mb, rec[off_mis:off_mis+len_mis], mb_index=mb_index) if off_mis else ''

            spcat = (struct.unpack_from('>H', rec, off_sp)[0] ^ 0x8000) if off_sp and off_sp + 2 <= len(rec) else 0
            cat   = (struct.unpack_from('>H', rec, off_cat)[0] ^ 0x8000) if off_cat and off_cat + 2 <= len(rec) else 0
            sbcat = (struct.unpack_from('>H', rec, off_sb)[0] ^ 0x8000) if off_sb and off_sb + 2 <= len(rec) else 0

            misure: list[dict[str, Any]] = []
            if misure_txt:
                for line in misure_txt.splitlines():
                    l_str = line.strip()
                    if not l_str:
                        continue
                    misure.append({
                        'id':          len(misure) + 1,
                        'rif_voce':    -2,
                        'descrizione': l_str,
                        'pu':          '',
                        'lu':          '',
                        'la':          '',
                        'hp':          '',
                        'qt':          qt if not misure else 0.0,
                        'flags':       0,
                    })

            if not misure:
                misure.append({
                    'id':          1,
                    'rif_voce':    -2,
                    'descrizione': '',
                    'pu':          '',
                    'lu':          '',
                    'la':          '',
                    'hp':          '',
                    'qt':          qt,
                    'flags':       0,
                })

            vc_list.append({
                'id':       len(vc_list) + 1,
                'id_ep':    ep_id,
                'quantita': qt,
                'importo':  0.0,
                'data_mis': '',
                'id_spcal': spcat,
                'id_cat':   cat,
                'id_sbcat': sbcat,
                'flags':    0,
                'misure':   misure,
                'cod_wbs':  '',
            })
    else:
        # Percorso di compatibilità per tabelle senza schema (es. test sintetici)
        raw_refs: list[tuple[int, int, float, float]] = []
        for rec in _iter_blocks(data, rec_size, hdr_size, block_size):
            s = _VC_SKIP
            id_vc = _read_long(rec, s+0)
            ref   = _read_long(rec, s+4)
            if not id_vc:
                continue
            qt  = _read_number(rec, s+8)
            imp = _read_number(rec, s+16)
            raw_refs.append((id_vc, ref, qt, imp))

        valid_refs = [ref for _, ref, _, _ in raw_refs if 0 < ref <= len(ep_ordinal) + 20]
        offset = min(valid_refs) - 1 if valid_refs else 0

        for (id_vc, ref, qt, imp) in raw_refs:
            ordinal = ref - offset
            ep_id   = ep_ordinal.get(ordinal, 0)
            if ep_id > 0:
                vc_list.append({
                    'id':       len(vc_list) + 1,
                    'id_ep':    ep_id,
                    'quantita': qt,
                    'importo':  imp,
                    'data_mis': '',
                    'id_spcal': 0,
                    'id_cat':   0,
                    'id_sbcat': 0,
                    'flags':    0,
                    'misure':   [{
                        'id':          1,
                        'rif_voce':    -2,
                        'descrizione': '',
                        'pu':          '',
                        'lu':          '',
                        'la':          '',
                        'hp':          '',
                        'qt':          qt,
                        'flags':       0,
                    }],
                    'cod_wbs':  '',
                })

    mapped = len(vc_list)
    _log(f'VC: {mapped} voci computo valide estratte da {db_path}')
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
    rec_size, hdr_size, block_size = _pdx_params(data)

    info: dict[str, str] = {}
    for rec in _iter_blocks(data, rec_size, hdr_size, block_size):
        s = _EP_SKIP
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
        '''Cerca un file con estensione DB corrispondente al prefisso specificato.'''
        for f in files:
            if f.upper().startswith(prefix.upper()) and f.upper().endswith('.DB'):
                return os.path.join(folder, f)
        return None

    db_e   = find_db('_E')
    db_v   = find_db('_V')
    db_b   = find_db('_B')
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
    # In PriMus: _B = Computo Base, _V = Computo Variante
    # Controlla entrambe le tabelle raccogliendo le voci computo valide
    for cand_db in (db_b, db_v):
        if cand_db and os.path.exists(cand_db):
            _log(f'Lettura computo da {cand_db}')
            items = read_vc_table(cand_db, ep_list)
            if items:
                vc_list.extend(items)

    # Riassegna identificatori sequenziali alle voci di computo
    for i, vc in enumerate(vc_list, 1):
        vc['id'] = i

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
