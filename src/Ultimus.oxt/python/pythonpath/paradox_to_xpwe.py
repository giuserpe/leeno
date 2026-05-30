"""
paradox_to_xpwe.py  —  Convertitore archivi Primus (Paradox) → XPWE 5.05

Converte i file .DB / .MB del vecchio PriMus DOS/Win 3.x nel formato standard
XPWE 5.05 (XML), direttamente importabile in LeenO via LeenoImport_XPWE.

Dipendenze: solo libreria standard Python 3.

Utilizzo standalone:
    python paradox_to_xpwe.py

Utilizzo da LeenO / pyleeno.py:
    import paradox_to_xpwe
    paradox_to_xpwe.import_paradox_db()   # apre dialogo file e importa

Formato interno dei .DB (Paradox custom di PriMus):
    Record size: 546 byte
    Block size:  2048 byte
    Ogni blocco inizia con 6 byte di intestazione (prev, next, n_recs) 
    poi n_recs * 546 byte di dati.

    Offsets dei campi rilevanti per record:
        0..19   Articolo       (Alpha 20)
        20      Flag record    (0x80 = valido)
        24..43  Tariffa        (Alpha 20)
        44..163 DesRidotta     (Alpha 120, troncata con ' ...' se >120 car)
        164..167 MB_OFFSET     (uint32 LE = offset diretto nel file .MB)
        168..175 MB_SIZE       (uint32 LE = dimensione blob in .MB)
        194..201 Prezzo1       (float64 big-endian con MSB invertito)
        202..209 Prezzo2
        210..217 Prezzo3
        218..225 Prezzo4
        226..233 Prezzo5
        234..241 QtyProgetto   (quantità totale progetto)
        242..249 PrezzoNetto   (prezzo netto)
        258..259 SuperCapitolo (int16 BE con MSB flip)
        260..261 Capitolo
        262..263 SubCapitolo
        264..265 Flags
        268..271 Long Flags (int32 BE con MSB flip)

    Formato blob nel .MB:
        mb_offset + 0 ..15 : UM (Unità di Misura, 16 byte null-padded)
        mb_offset + 16 ... : DesEstesa (stringa null-terminated)
"""

import os
import re
import struct
import math
import xml.etree.ElementTree as ET
import xml.dom.minidom
import LeenoDialogs as DLG
import Dialogs

# ---------------------------------------------------------------------------
# Decodifica numeri nel formato Paradox (Big-Endian, bit di segno invertito)
# ---------------------------------------------------------------------------

def _px_float(raw8: bytes) -> float:
    """Decodifica un double a 64 bit nel formato proprietario Paradox."""
    if raw8 == b'\x00' * 8:
        return 0.0
    if raw8[0] & 0x80:
        # numero positivo: flip solo il bit di segno (MSB)
        b = bytes([raw8[0] ^ 0x80]) + raw8[1:]
    else:
        # numero negativo: flip tutti i bit
        b = bytes([x ^ 0xFF for x in raw8])
    try:
        f = struct.unpack('>d', b)[0]
    except struct.error:
        return 0.0
    if math.isnan(f) or math.isinf(f):
        return 0.0
    return f


def _px_short(raw2: bytes) -> int:
    """Decodifica un intero a 16 bit nel formato Paradox (BE, MSB flip)."""
    if raw2 == b'\x00\x00':
        return 0
    b = bytes([raw2[0] ^ 0x80]) + raw2[1:]
    return struct.unpack('>h', b)[0]


def _px_long(raw4: bytes) -> int:
    """Decodifica un intero a 32 bit nel formato Paradox (BE, MSB flip)."""
    if raw4 == b'\x00' * 4:
        return 0
    b = bytes([raw4[0] ^ 0x80]) + raw4[3:0:-1]  # big-endian, flip MSB
    return struct.unpack('>i', b)[0]


def _sanitize(text: str) -> str:
    """Rimuove caratteri di controllo non validi in XML (tranne \t \n \r)."""
    if not text:
        return ''
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)


# ---------------------------------------------------------------------------
# Parser del file .MB  (descrizioni estese e unità di misura)
# ---------------------------------------------------------------------------

def _mb_block_size(mb_data: bytes) -> int:
    """
    Legge la dimensione del blocco dal header del file .MB (Paradox).
    Offset 6 nell'header del .MB: uint16 LE = block size in pagine da 1024 byte.
    Restituisce la dimensione in byte del blocco, default 512.
    """
    if len(mb_data) < 8:
        return 512
    try:
        blk_pages = struct.unpack_from('<H', mb_data, 6)[0]
        if blk_pages > 0:
            return blk_pages * 1024
    except Exception:
        pass
    return 512


def _read_mb_blob(mb_data: bytes, block_num: int, mb_size: int = 0) -> tuple[str, str]:
    """
    Legge un blob dal file .MB.
    In Paradox, il campo nel .DB è un *numero di blocco* (non un byte offset diretto).
    Il byte offset reale è: block_num * mb_block_size.
    Struttura blob: 16 byte UM (null-padded) + desc (null-terminated).
    Prova in ordine: block-based (standard Paradox), poi offset diretto come fallback.
    Restituisce (um, des_estesa).
    """
    if not mb_data or block_num <= 0:
        return '', ''

    blksz = _mb_block_size(mb_data)
    # Candidati: block-based (standard Paradox) e offset diretto (fallback)
    candidates = [block_num * blksz, block_num]

    for start in candidates:
        if start <= 0 or start + 17 > len(mb_data):
            continue
        um_raw = mb_data[start: start + 16]
        um = um_raw.rstrip(b'\x00').decode('windows-1252', errors='ignore').strip()
        desc_start = start + 16
        # Usa mb_size per delimitare il blob se disponibile e ragionevole
        if mb_size > 16:
            desc_end = min(start + mb_size, len(mb_data))
            raw_desc = mb_data[desc_start:desc_end]
            null_pos = raw_desc.find(b'\x00')
            if null_pos >= 0:
                raw_desc = raw_desc[:null_pos]
        else:
            null_pos = mb_data.find(b'\x00', desc_start)
            desc_end = null_pos if 0 < null_pos < desc_start + 8192 else min(desc_start + 8192, len(mb_data))
            raw_desc = mb_data[desc_start:desc_end]
        desc = raw_desc.decode('windows-1252', errors='ignore').strip()
        # Scarta risultati chiaramente non validi (testo non stampabile > 20%)
        if desc:
            printable = sum(1 for c in desc if c.isprintable() or c in '\n\r\t')
            if printable >= len(desc) * 0.8:
                return um, desc

    return '', ''


# ---------------------------------------------------------------------------
# Parser del file .DB  (record Elenco Prezzi)
# ---------------------------------------------------------------------------

# Offset dei campi nel record da 546 byte (rilevati empiricamente)
_OFF_ARTICOLO   = 0    # Alpha 20
_OFF_FLAG       = 20   # 1 byte: 0x80 = record valido
_OFF_TARIFFA    = 24   # Alpha 20
_OFF_DES_RID    = 44   # Alpha 120  (troncata)
_OFF_MB_OFFSET  = 164  # uint32 LE = numero di blocco nel .MB (Paradox block_num)
_OFF_MB_SIZE    = 168  # uint32 LE = dimensione blob .MB (byte)
_OFF_PREZZO1    = 194  # float64 Paradox (MSB flip)
_OFF_PREZZO2    = 202
_OFF_PREZZO3    = 210
_OFF_PREZZO4    = 218
_OFF_PREZZO5    = 226
_OFF_QTY        = 234  # quantità progetto
_OFF_PREZZO_NET = 242  # prezzo netto
_OFF_SPCAP      = 258  # int16 Paradox = ID SuperCapitolo
_OFF_CAP        = 260
_OFF_SBCAP      = 262
_OFF_FLAGS      = 264  # int16 flags
_OFF_LONG_FLAGS = 268  # int32 flags esteso

_BLOCK_HEADER   = 6    # prev(2) + next(2) + n_recs(2)
_REC_SIZE       = 546
_BLOCK_SIZE     = 2048
_HDR_SIZE       = 4096


def parse_ep_db(db_path: str, mb_path: str) -> list[dict]:
    """
    Parsa il file .DB dell'Elenco Prezzi di PriMus e restituisce
    una lista di dizionari con i dati di ogni voce EP.
    """
    with open(db_path, 'rb') as f:
        db = f.read()

    mb = b''
    if mb_path and os.path.exists(mb_path):
        with open(mb_path, 'rb') as f:
            mb = f.read()

    # Verifica magic / sanity check header
    rec_size  = struct.unpack_from('<H', db, 0)[0]
    hdr_size  = struct.unpack_from('<H', db, 2)[0]
    block_sz  = db[5] * 0x400

    if rec_size == 0:
        raise ValueError("File .DB non valido o vuoto (rec_size=0)")

    records = []
    seen_tariffe = set()  # deduplicazione
    block_offset = hdr_size

    while block_offset + _BLOCK_HEADER + rec_size <= len(db):
        val = struct.unpack_from('<h', db, block_offset + 4)[0]
        n_in_blk = val // rec_size

        if n_in_blk <= 0 or n_in_blk > block_sz // rec_size:
            block_offset += block_sz
            continue

        for r in range(n_in_blk):
            roff = block_offset + _BLOCK_HEADER + r * rec_size
            if roff + rec_size > len(db):
                break
            rec = db[roff: roff + rec_size]

            # Controlla flag validità
            if rec[_OFF_FLAG] != 0x80:
                continue

            # Leggi campi stringa
            articolo = rec[_OFF_ARTICOLO: _OFF_ARTICOLO + 20].rstrip(b'\x00').decode('windows-1252', errors='ignore').strip()
            tariffa  = rec[_OFF_TARIFFA:  _OFF_TARIFFA  + 20].rstrip(b'\x00').decode('windows-1252', errors='ignore').strip()
            des_rid  = rec[_OFF_DES_RID:  _OFF_DES_RID  + 120].rstrip(b'\x00').decode('windows-1252', errors='ignore').strip()

            # Salta record senza tariffa e duplicati
            if not tariffa:
                continue
            if tariffa in seen_tariffe:
                continue
            seen_tariffe.add(tariffa)

            # Leggi puntatore blob MB (numero di blocco Paradox) e dimensione
            mb_block_num = struct.unpack_from('<I', rec, _OFF_MB_OFFSET)[0]
            mb_size      = struct.unpack_from('<I', rec, _OFF_MB_SIZE)[0]
            um, des_estesa = _read_mb_blob(mb, mb_block_num, mb_size)

            # Se des_estesa è vuota usa des_rid
            if not des_estesa:
                des_estesa = des_rid
            # Se des_rid è troncata ma des_estesa è completa, aggiorna anche des_rid
            if des_rid.endswith(' ...') or (des_estesa and len(des_estesa) > len(des_rid)):
                des_rid_clean = des_estesa[:120].rstrip()
            else:
                des_rid_clean = des_rid

            # Leggi prezzi e altre info numeriche
            prezzo1    = _px_float(rec[_OFF_PREZZO1:   _OFF_PREZZO1   + 8])
            prezzo2    = _px_float(rec[_OFF_PREZZO2:   _OFF_PREZZO2   + 8])
            prezzo3    = _px_float(rec[_OFF_PREZZO3:   _OFF_PREZZO3   + 8])
            prezzo4    = _px_float(rec[_OFF_PREZZO4:   _OFF_PREZZO4   + 8])
            prezzo5    = _px_float(rec[_OFF_PREZZO5:   _OFF_PREZZO5   + 8])
            prezzo_net = _px_float(rec[_OFF_PREZZO_NET:_OFF_PREZZO_NET + 8])
            id_spcap   = _px_short(rec[_OFF_SPCAP: _OFF_SPCAP + 2])
            id_cap     = _px_short(rec[_OFF_CAP:   _OFF_CAP   + 2])
            id_sbcap   = _px_short(rec[_OFF_SBCAP: _OFF_SBCAP + 2])
            flags      = struct.unpack_from('<I', rec, _OFF_LONG_FLAGS)[0]

            records.append({
                'id':           len(records) + 1,
                'articolo':     articolo,
                'tariffa':      tariffa,
                'des_ridotta':  des_rid_clean,
                'des_estesa':   des_estesa,
                'um':           um,
                'prezzo1':      prezzo1,
                'prezzo2':      prezzo2,
                'prezzo3':      prezzo3,
                'prezzo4':      prezzo4,
                'prezzo5':      prezzo5,
                'prezzo_netto': prezzo_net,
                'id_spcap':     max(0, id_spcap),
                'id_cap':       max(0, id_cap),
                'id_sbcap':     max(0, id_sbcap),
                'flags':        flags,
            })

        block_offset += block_sz

    return records


# ---------------------------------------------------------------------------
# Generatore XML XPWE 5.05
# ---------------------------------------------------------------------------

def _sub(parent: ET.Element, tag: str, text: str = '') -> ET.Element:
    el = ET.SubElement(parent, tag)
    el.text = text
    return el


def build_xpwe(db_path: str, mb_path: str, output_path: str) -> str:
    """
    Legge db_path (Elenco Prezzi Paradox) e genera un file XPWE 5.05.
    Restituisce il percorso del file generato.
    """
    ep_records = parse_ep_db(db_path, mb_path)

    # --- Root ---
    root = ET.Element('PweDocumento')
    _sub(root, 'CopyRight',  'Copyright ACCA software S.p.A.')
    _sub(root, 'TipoDocumento', '1')     # 1 = Progetto
    _sub(root, 'TipoFormato',  'XMLPwe')
    _sub(root, 'Versione',     '5.05')
    _sub(root, 'SourceVersione', '1.0')
    _sub(root, 'SourceNome',     'paradox_to_xpwe')
    _sub(root, 'FileNameDocumento', os.path.basename(db_path))

    # --- Dati Generali (vuoti, non presenti nel .DB EP) ---
    dg = ET.SubElement(root, 'PweDatiGenerali')
    dgp = ET.SubElement(dg, 'PweDGProgetto')
    ET.SubElement(ET.SubElement(dgp, 'PweDGDatiGenerali'), 'Oggetto')

    # --- Misurazioni → Elenco Prezzi ---
    mis = ET.SubElement(root, 'PweMisurazioni')
    ep_root = ET.SubElement(mis, 'PweElencoPrezzi')

    for ep in ep_records:
        item = ET.SubElement(ep_root, 'EPItem', ID=str(ep['id']))
        _sub(item, 'TipoEP',      '0')
        _sub(item, 'Tariffa',     _sanitize(ep['tariffa']))
        _sub(item, 'Articolo',    _sanitize(ep['articolo']))
        _sub(item, 'DesRidotta',  _sanitize(ep['des_ridotta']))
        _sub(item, 'DesEstesa',   _sanitize(ep['des_estesa']))
        _sub(item, 'UnMisura',    _sanitize(ep['um']))
        _sub(item, 'Prezzo1',     f"{ep['prezzo1']:.4f}")
        _sub(item, 'Prezzo2',     f"{ep['prezzo2']:.4f}")
        _sub(item, 'Prezzo3',     f"{ep['prezzo3']:.4f}")
        _sub(item, 'Prezzo4',     f"{ep['prezzo4']:.4f}")
        _sub(item, 'Prezzo5',     f"{ep['prezzo5']:.4f}")
        _sub(item, 'IDSpCap',     str(ep['id_spcap']))
        _sub(item, 'IDCap',       str(ep['id_cap']))
        _sub(item, 'IDSbCap',     str(ep['id_sbcap']))
        _sub(item, 'Flags',       str(ep['flags']))

    # --- PweVociComputo (vuoto — no file _V nel sample) ---
    ET.SubElement(mis, 'PweVociComputo')

    # --- Serializza XML con pretty-print ---
    xml_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=False)
    dom = xml.dom.minidom.parseString(xml_bytes)
    pretty = dom.toprettyxml(indent='  ', encoding='utf-8').decode('utf-8')
    # Rimuovi la dichiarazione duplicata generata da toprettyxml
    lines = pretty.splitlines()
    if lines and lines[0].startswith('<?xml'):
        lines = lines[1:]

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('<?xml version="1.0" encoding="utf-8"?>\n')
        f.write('<!-- Generato da paradox_to_xpwe.py  — Protocollo XPWE 5.05 -->\n')
        f.write('\n'.join(lines))

    Dialogs.Info(Title='Info', Text=f"XPWE generato: {output_path}")
    Dialogs.Info(Title='Info', Text=f"  Voci EP estratte: {len(ep_records)}")
    return output_path


# ---------------------------------------------------------------------------
# Interfaccia LeenO  (chiamata da pyleeno.py)
# ---------------------------------------------------------------------------

def import_paradox_db():
    """
    Apre una finestra di selezione file e importa il file .DB Paradox
    in un Computo LeenO, convertendolo prima in XPWE.
    Da chiamare da pyleeno.py / LeenO.
    """
    db_path = Dialogs.FileSelect('Seleziona Elenco Prezzi Paradox (*.DB)...', '*.DB', 0)
    if not db_path:
        return

    mb_path   = db_path[:-3] + '.MB' if db_path.upper().endswith('.DB') else db_path + '.MB'
    xpwe_path = db_path + '.xpwe'

    try:
        build_xpwe(db_path, mb_path, xpwe_path)
    except Exception as e:
        DLG.chi(f"Errore durante l'estrazione Paradox → XPWE:\n{e}")
        return

    if not os.path.exists(xpwe_path):
        DLG.chi('Errore: file XPWE non generato.\n'
                'Verificare i permessi sulla cartella di destinazione.')
        return

    try:
        # pyrefly: ignore [missing-import]
        from LeenoImport_XPWE import XPWE_import
        XPWE_import(xpwe_path)
    except Exception as e:
        DLG.chi(f"Errore durante l'importazione XPWE:\n{e}")


# ---------------------------------------------------------------------------
# Entry point standalone
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    BASE = r'w:\_dwg\ULTIMUSFREE\_SRC\leeno\@Specifiche_XPWE\paradox'
    db_file  = os.path.join(BASE, '_E107-ristrutturazione-coperture-COMPUTO.DB')
    mb_file  = os.path.join(BASE, '_E107-ristrutturazione-coperture-COMPUTO.MB')
    out_file = os.path.join(BASE, 'estratto_XPWE.xpwe')

    if os.path.exists(db_file):
        records = parse_ep_db(db_file, mb_file)
        print(f"\n{'Articolo':<8} {'Tariffa':<18} {'UM':<8} {'Prezzo1':>10}  Descrizione")
        print('-' * 90)
        for r in records:
            print(f"{r['articolo']:<8} {r['tariffa']:<18} {r['um']:<8} "
                  f"{r['prezzo1']:>10.4f}  {r['des_ridotta'][:40]}")
        build_xpwe(db_file, mb_file, out_file)
    else:
        print(f"File DB non trovato: {db_file}")
