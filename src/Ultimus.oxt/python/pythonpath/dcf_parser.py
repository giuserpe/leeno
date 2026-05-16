"""
dcf_parser.py — Parser per file PriMus/ACCA
Supporta due formati:
  • .dcf  — formato binario SFS (interno a PriMus), con XML embedded compresso
  • .xpwe — formato standard XPWE 5.05 (pubblico, XML puro)

Dipendenze: solo libreria standard (xml.etree.ElementTree, zlib, re)

Uso:
    from dcf_parser import parse_dcf, parse_xpwe, parse_auto, computo_con_descrizioni

    doc = parse_dcf('menzella.dcf')     # formato binario SFS
    doc = parse_xpwe('miofile.xpwe')    # formato standard XPWE
    doc = parse_auto('file.dcf')        # rileva automaticamente

    # Struttura doc identica per entrambi i formati:
    doc['info']              -> dict dati generali progetto
    doc['quadro_economico']  -> list voci QE  (solo .dcf)
    doc['super_categorie']   -> dict id->{nome, importo, cod}
    doc['categorie']         -> dict id->{nome, importo, cod}
    doc['sotto_categorie']   -> dict id->{nome, importo, cod}
    doc['super_capitoli']    -> dict id->{nome, importo}
    doc['capitoli']          -> dict id->{nome, importo}
    doc['sotto_capitoli']    -> dict id->{nome, importo}
    doc['elenco_prezzi']     -> list voci EP (deduplicate per tariffa)
    doc['computo']           -> list voci VC con misure nested
    doc['strutture_stampa']  -> list template stampa  (solo .dcf)
    doc['_ep_by_id']         -> dict id->voce EP (lookup rapido)
"""

from __future__ import annotations
import re
import zlib
import html as _html_module
import xml.etree.ElementTree as ET
import os
import datetime
from typing import Any

import sys
import os

# === Miglioramento per menu contestuale ===
if len(sys.argv) > 1:
    path_arg = sys.argv[1]
    if not os.path.isabs(path_arg):
        # Se viene passato solo il nome file, cerca nella stessa cartella dello script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        path_arg = os.path.join(script_dir, path_arg)
    # Normalizza il percorso
    if os.path.exists(path_arg):
        sys.argv[1] = os.path.normpath(path_arg)
    else:
        print(f"ERRORE: File non trovato: {path_arg}")
        input("Premi INVIO per chiudere...")
        sys.exit(1)


# ===========================================================================
# LOGGING SU FILE  (sostituto di print() per uso con LibreOffice / DLG.chi)
# ===========================================================================

LOG_FILE: str = os.path.join(os.path.expanduser('~'), 'dcf_parser.log')
_LOG_ENABLED: bool = True


def _log(*args, sep=' ') -> None:
    if not _LOG_ENABLED:
        return
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
            fh.write(f'# dcf_parser log — {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n')
    except OSError:
        pass


# ===========================================================================
# STATUS INDICATOR  (progressbar nativa LibreOffice Calc)
# ===========================================================================
#
# Uso da LeenO / pyleeno.py:
#
#   import dcf_parser
#   indicator = oDoc.getCurrentController().getStatusIndicator()
#   dcf_parser.set_indicator(indicator)
#   doc = dcf_parser.parse_dcf(percorso)
#   dcf_parser.set_indicator(None)          # rilascia al termine

_indicator = None   # com.sun.star.task.XStatusIndicator


def set_indicator(indicator) -> None:
    global _indicator
    _indicator = indicator


def _ind_start(msg: str, max_val: int = 100) -> None:
    if _indicator is None:
        return
    try:
        _indicator.start(msg, max_val)
    except Exception:
        pass


def _ind_update(msg: str, value: int) -> None:
    if _indicator is None:
        return
    try:
        _indicator.setValue(value)
        _indicator.setText(msg)
    except Exception:
        pass


def _ind_end() -> None:
    if _indicator is None:
        return
    try:
        _indicator.end()
    except Exception:
        pass


# ===========================================================================
# HELPERS COMUNI
# ===========================================================================

def _float(v) -> float:
    if v is None:
        return 0.0
    try:
        return float(str(v).replace(' ', '').replace(',', '.'))
    except (ValueError, TypeError):
        return 0.0


def _int(v) -> int:
    if v is None:
        return 0
    try:
        return int(str(v).split('.')[0])
    except (ValueError, TypeError):
        return 0


def _unescape(s: str) -> str:
    return _html_module.unescape(s or '')


# ===========================================================================
# FORMATO SFS (.dcf)  —  estrazione binaria + XML embedded compresso
# ===========================================================================

_HDR = 0x1c8

# Tag XML validi attesi in un file PriMus — filtra tag spazzatura
# (es. 'p', 'u' o caratteri non-ASCII da decompressione parziale errata)
_VALID_TAGS = {
    'CollectionEP', 'CollectionVC', 'CollectionST', 'CollectionGP',
    'CollectionOD', 'DatiGenerali', 'InfoDoc', 'InfoMRR',
    'OpzioniInterfaccia',
}


def _sfs_extract_xml(dcf_path: str) -> dict[str, str]:
    """
    Legge un file .dcf PriMus (formato SFS) e restituisce un dict nome->testo XML.

    Usa due strategie in cascata:
    1. AACS-driven: cerca i marker 'AACS' nel file e decomprime i dati che seguono.
       Questo funziona per tutti i formati PriMus indipendentemente dalla slot_size.
    2. Slot-driven (fallback): prova slot_size 0x1000 e 0x800 se la scansione AACS
       non produce blocchi sufficienti.

    I file contabilita (TpDoc=2) hanno blocchi dati crittografati:
    solo InfoDoc e OpzioniInterfaccia risultano leggibili.
    """
    with open(dcf_path, 'rb') as f:
        data = f.read()
    if data[:8] != b'AAMVHFSS':
        raise ValueError(f"Non e' un file SFS valido (magic: {data[:8]!r})")

    zlib_magics = {b'\x78\x9c', b'\x78\xda', b'\x78\x01', b'\x78\x5e'}
    file_size   = len(data)
    _log(f'SFS scan: file={dcf_path!r}  size={file_size} bytes')

    _VALID_TAGS = {
        'CollectionEP', 'CollectionVC', 'CollectionST', 'CollectionGP',
        'CollectionOD', 'DatiGenerali', 'InfoDoc', 'InfoMRR',
        'OpzioniInterfaccia',
    }

    def _decode(raw: bytes) -> str:
        try:
            return raw.decode('utf-8')
        except UnicodeDecodeError:
            return raw.decode('windows-1252', errors='replace')

    def _decompress(payload: bytes) -> bytes:
        """Prova zlib diretto, poi deflate raw."""
        try:
            return zlib.decompress(payload)
        except zlib.error:
            pass
        d = zlib.decompressobj(wbits=-15)
        dec = b''
        for chunk in [payload[2:][i:i+8192] for i in range(0, len(payload[2:]), 8192)]:
            try:
                dec += d.decompress(chunk)
            except zlib.error:
                break
        try:
            dec += d.flush()
        except zlib.error:
            pass
        return dec

    def _tag_from_text(txt: str):
        m = re.search(r'<([A-Za-z]\w*)', txt)
        if not m:
            return None
        c = m.group(1)
        if c in _VALID_TAGS:
            return c
        if c.isascii() and len(c) >= 4 and c[0].isupper():
            return c
        return None

    def _merge(blocks: dict, tag: str, txt: str):
        if tag not in blocks or len(txt) > len(blocks[tag]):
            blocks[tag] = txt

    # ------------------------------------------------------------------
    # STRATEGIA 1: AACS-driven
    # Ogni blocco SFS inizia con 0x138 byte di header, poi 'AACS' (4 byte),
    # poi 0x8c byte di header AACS, poi i dati zlib.
    # Distanza da inizio 'AACS' al magic zlib = 0x90 byte.
    # ------------------------------------------------------------------
    AACS_TO_ZLIB = 0x90
    blocks_aacs: dict[str, str] = {}

    aacs_offsets = []
    p = 0
    while True:
        pos = data.find(b'AACS', p)
        if pos < 0:
            break
        aacs_offsets.append(pos)
        p = pos + 1

    _log(f'  AACS markers trovati: {len(aacs_offsets)}')
    _ind_start('Lettura DCF...', 100)

    for idx, pos in enumerate(aacs_offsets):
        pct = idx * 100 // max(len(aacs_offsets), 1)
        _ind_update(f'AACS {idx+1}/{len(aacs_offsets)}', pct)

        zlib_start = pos + AACS_TO_ZLIB
        if zlib_start + 2 > file_size:
            continue
        if data[zlib_start:zlib_start+2] not in zlib_magics:
            continue

        # Payload: da zlib_start fino al prossimo AACS marker (escluso header)
        next_pos = aacs_offsets[idx+1] if idx+1 < len(aacs_offsets) else file_size
        payload  = data[zlib_start:next_pos]

        dec = _decompress(payload)
        if len(dec) < 10:
            continue

        # Rilevamento DRM ACCA: blocco cifrato inizia con 4 byte lunghezza
        # + "Copyright ACCA software S.p.A." invece di XML
        _CR = b'Copyright ACCA software S.p.A.'
        if len(dec) >= 34 and dec[4:4+len(_CR)] == _CR:
            _log(f'  AACS@0x{pos:06X}  DRM_CIFRATO  dec={len(dec)}')
            blocks_aacs.setdefault('_drm_count', 0)
            blocks_aacs['_drm_count'] = blocks_aacs.get('_drm_count', 0) + 1
            continue

        txt = _decode(dec)
        tag = _tag_from_text(txt)
        if not tag:
            _log(f'  AACS@0x{pos:06X}  SCARTATO  dec={len(dec)}')
            continue

        _log(f'  AACS@0x{pos:06X}  tag={tag}  dec={len(dec)}')
        _merge(blocks_aacs, tag, txt)

    _log(f'  AACS-driven: {list(blocks_aacs.keys())}')

    # Rileva DRM: se molti blocchi sono cifrati e nessun XML trovato -> DRM ACCA
    drm_count = blocks_aacs.pop('_drm_count', 0)
    if drm_count >= 3 and not blocks_aacs:
        _log(f'  RILEVATO DRM ACCA ({drm_count} blocchi cifrati) — dati non estraibili')
        return {'_drm': 'ACCA', '_drm_count': drm_count}

    # Se la strategia AACS ha prodotto almeno i blocchi fondamentali, usala
    has_core = bool(blocks_aacs.get('CollectionEP') or blocks_aacs.get('DatiGenerali'))
    if has_core:
        _log('SFS: usando risultati AACS-driven')
        _log(f'SFS: blocchi finali={list(blocks_aacs.keys())}')
        return blocks_aacs

    # ------------------------------------------------------------------
    # STRATEGIA 2: Slot-driven (fallback)
    # ------------------------------------------------------------------
    _log('  AACS-driven insufficiente, fallback slot-driven...')
    best_blocks: dict[str, str] = {}
    best_count  = 0

    for pass_num, slot_size in enumerate((0x1000, 0x800), 1):
        n_slots = file_size // slot_size
        blocks_pass: dict[str, str] = {}
        found: list[str] = []
        base_pct = (pass_num - 1) * 50

        _ind_update(f'Pass {pass_num}/2 — slot 0x{slot_size:04X}...', base_pct)
        last_pct = -1

        for slot in range(n_slots):
            pct = slot * 50 // max(n_slots, 1)
            if pct != last_pct:
                _ind_update(f'Pass {pass_num}/2 — {slot}/{n_slots} slot', base_pct + pct)
                last_pct = pct

            offset = slot * slot_size + _HDR
            if offset >= file_size:
                break
            payload = data[offset: slot * slot_size + slot_size]
            if not payload or payload[:2] not in zlib_magics:
                continue

            dec = _decompress(payload)
            if len(dec) < 10:
                continue

            txt = _decode(dec)
            tag = _tag_from_text(txt)
            if not tag:
                _log(f'  slot_size=0x{slot_size:04X}  slot={slot:4d}  SCARTATO')
                continue

            _log(f'  slot_size=0x{slot_size:04X}  slot={slot:4d}  '
                 f'offset=0x{slot*slot_size:08X}  tag={tag}  dec={len(dec)}')
            found.append(tag)
            _merge(blocks_pass, tag, txt)

        _log(f'  slot_size=0x{slot_size:04X}: {len(found)} blocchi: {found}')
        if len(found) > best_count:
            best_count  = len(found)
            best_blocks = blocks_pass

    _log(f'SFS: blocchi finali={list(best_blocks.keys())}')
    return best_blocks



# ---------------------------------------------------------------------------
# Parser XML interno .dcf  (dati come attributi XML)
# ---------------------------------------------------------------------------

def _attr_parse(seg: str) -> dict[str, str]:
    return {k: _unescape(v) for k, v in re.findall(r'(\w+)="([^"]*)"', seg)}


def _iter_attr_tags(xml: str, tag: str):
    for m in re.finditer(rf'<{tag}\s[^<]+?/?>', xml, re.DOTALL):
        yield _attr_parse(m.group(0))


def _dcf_parse_dati_generali(xml: str) -> dict[str, Any]:
    result: dict[str, Any] = {}

    m = re.search(r'<DGDatiGenerali\s([^>]+)>', xml)
    if m:
        a = _attr_parse(m.group(0))
        result['info'] = {
            'oggetto':     a.get('Ogg', ''),
            'committente': a.get('Comm', ''),
            'comune':      a.get('Cmn', ''),
            'provincia':   a.get('Prvnc', ''),
            'impresa':     a.get('Imprs', ''),
            'operatore':   a.get('Oprtr', ''),
        }

    qe_items = []
    for a in _iter_attr_tags(xml, 'ItemQE'):
        imprt = a.get('Imprt', '').strip()
        qe_items.append({
            'descrizione': a.get('Des', ''),
            'importo':     _float(imprt) if imprt else None,
            'formula':     a.get('Frml', ''),
            'tipo':        _int(a.get('TpImprt', '0')),
        })
    result['quadro_economico'] = qe_items

    for prefix, key in [
        ('SpCap', 'super_capitoli'), ('Cap', 'capitoli'), ('SbCap', 'sotto_capitoli'),
        ('SpCat', 'super_categorie'), ('Cat', 'categorie'), ('SbCat', 'sotto_categorie'),
    ]:
        items: dict[int, dict] = {}
        for a in _iter_attr_tags(xml, f'Item{prefix}'):
            id_ = _int(a.get('Id', '0'))
            nome = a.get('DesSnt', '')
            if nome and nome not in ('<nessuna>', 'overflow'):
                items[id_] = {
                    'nome':    nome,
                    'importo': _float(a.get('MDOImp', '0')),
                    'cod':     a.get('Cod', ''),
                }
        result[key] = items

    return result


def _dcf_parse_ep(xml: str) -> tuple[list[dict], dict[int, dict]]:
    """
    Parser EP da .dcf.
    PriMus duplica ogni ItemEP per ogni capitolo/categoria.
    Deduplica per tariffa tenendo la DesEst piu' lunga.
    """
    segments = re.split(r'(?=<ItemEP\s)', xml)
    by_id:  dict[int, dict] = {}
    by_trf: dict[str, dict] = {}

    for seg in segments:
        if not seg.strip().startswith('<ItemEP'):
            continue
        a = _attr_parse(seg)
        if not a.get('IdEP'):
            continue
        try:
            id_ep = int(a['IdEP'])
        except ValueError:
            continue
        trf  = a.get('Trf', '')
        dest = a.get('DesEst', '')
        prev = by_id.get(id_ep)
        if prev is None or len(dest) > len(prev.get('DesEst', '')):
            by_id[id_ep] = a
        if trf:
            prev_trf = by_trf.get(trf)
            if prev_trf is None or len(dest) > len(prev_trf.get('DesEst', '')):
                by_trf[trf] = a

    def _build(a: dict) -> dict[str, Any]:
        dest = _unescape(a.get('DesEst', ''))
        rid  = _unescape(a.get('DesRid', ''))
        brv  = _unescape(a.get('DesBrv', ''))
        if not dest or ' ... ' in dest:
            if rid and ' ... ' not in rid:
                dest = rid
            elif brv and ' ... ' not in brv:
                dest = brv
        return {
            'id':           _int(a.get('IdEP', '0')),
            'tariffa':      a.get('Trf', ''),
            'articolo':     a.get('Art', ''),
            'des_ridotta':  rid,
            'des_breve':    brv,
            'des_estesa':   dest,
            'um':           a.get('UnMsr', ''),
            'prezzo':       _float(a.get('Prz1', '0')),
            'prezzo_netto': _float(a.get('PrNetto', '0')),
            'inc_mdo':      _float(a.get('MDOInc', '0')),
            'inc_sic':      _float(a.get('SICInc', '0')),
            'ribassabile':  a.get('Rbs', '0') != '1',
            'id_spcap':     _int(a.get('IdSpCap', '0')),
            'id_cap':       _int(a.get('IdCap', '0')),
            'id_sbcap':     _int(a.get('IdSbCap', '0')),
            'flags':        _int(a.get('Fgs', '0')),
        }

    ep_list  = [_build(a) for a in sorted(by_trf.values(), key=lambda x: x.get('Trf', ''))]
    ep_by_id = {id_ep: _build(a) for id_ep, a in by_id.items()}
    return ep_list, ep_by_id


def _dcf_parse_vc(xml: str) -> list[dict[str, Any]]:
    items = []
    for m in re.finditer(r'<ItemVC\s[^>]+>.*?</ItemVC>', xml, re.DOTALL):
        seg = m.group(0)
        a = _attr_parse(seg)
        if not a.get('IdVC'):
            continue
        misure = []
        for rm in _iter_attr_tags(seg, 'ItemRM'):
            misure.append({
                'id':          _int(rm.get('IdRM', '0')),
                'descrizione': rm.get('Des', ''),
                'pu':          _float(rm.get('Pu', '0')),
                'lu':          _float(rm.get('Lu', '0')),
                'la':          _float(rm.get('La', '0')),
                'hp':          _float(rm.get('Hp', '0')),
                'qt':          _float(rm.get('Qt', '0')),
                'rif_voce':    _int(rm.get('IdVV', '0')),
            })
        items.append({
            'id':       _int(a.get('IdVC', '0')),
            'id_ep':    _int(a.get('IdEP', '0')),
            'quantita': _float(a.get('Qt', '0')),
            'importo':  _float(a.get('Imprt', '0')),
            'data_mis': a.get('DataMis', ''),
            'id_spcal': _int(a.get('IdSpCat', '0')),
            'id_cat':   _int(a.get('IdCat', '0')),
            'id_sbcat': _int(a.get('IdSbCat', '0')),
            'flags':    _int(a.get('Fgs', '0')),
            'misure':   misure,
        })
    # Self-closing senza misure
    for m in re.finditer(r'<ItemVC\s[^>]+/>', xml):
        a = _attr_parse(m.group(0))
        if not a.get('IdVC'):
            continue
        items.append({
            'id':       _int(a.get('IdVC', '0')),
            'id_ep':    _int(a.get('IdEP', '0')),
            'quantita': _float(a.get('Qt', '0')),
            'importo':  _float(a.get('Imprt', '0')),
            'data_mis': a.get('DataMis', ''),
            'id_spcal': _int(a.get('IdSpCat', '0')),
            'id_cat':   _int(a.get('IdCat', '0')),
            'id_sbcat': _int(a.get('IdSbCat', '0')),
            'flags':    _int(a.get('Fgs', '0')),
            'misure':   [],
        })
    items.sort(key=lambda x: x['id'])
    return items


def _dcf_parse_st(xml: str) -> list[dict[str, Any]]:
    tp_map = {1: 'radice', 2: 'analisi_prezzi', 3: 'fabbisogno',
              4: 'computo', 5: 'elenco_prezzi', 6: 'SAL',
              7: 'relazione', 8: 'cronoprogramma'}
    items = []
    for a in _iter_attr_tags(xml, 'ItemST'):
        tp = _int(a.get('Tp', '0'))
        items.append({
            'id':      _int(a.get('IdST', '0')),
            'id_prnt': _int(a.get('IdPrnt', '0')),
            'tipo':    tp_map.get(tp, f'tipo_{tp}'),
            'titolo':  a.get('Titolo', ''),
        })
    return items


# ===========================================================================
# FORMATO XPWE (.xpwe)  —  XML puro secondo protocollo XPWE 5.05
# ===========================================================================

def _xpwe_parse_categorie(root: ET.Element) -> dict[str, Any]:
    result: dict[str, Any] = {
        'super_capitoli':  {},
        'capitoli':        {},
        'sotto_capitoli':  {},
        'super_categorie': {},
        'categorie':       {},
        'sotto_categorie': {},
    }
    cc = root.find('.//PweDGCapitoliCategorie')
    if cc is None:
        return result
    for xpath, key in [
        ('PweDGSuperCapitoli/DGSuperCapitoliItem', 'super_capitoli'),
        ('PweDGCapitoli/DGCapitoliItem',            'capitoli'),
        ('PweDGSubCapitoli/DGSubCapitoliItem',       'sotto_capitoli'),
        ('PweDGSuperCategorie/DGSuperCategorieItem', 'super_categorie'),
        ('PweDGCategorie/DGCategorieItem',           'categorie'),
        ('PweDGSubCategorie/DGSubCategorieItem',     'sotto_categorie'),
    ]:
        for el in cc.findall(xpath):
            id_ = _int(el.get('ID', '0'))
            nome = (el.findtext('DesSintetica') or '').strip()
            if nome:
                result[key][id_] = {
                    'nome':    nome,
                    'des_est': (el.findtext('DesEstesa') or '').strip(),
                    'importo': 0.0,
                    'cod':     (el.findtext('Codice') or '').strip(),
                }
    return result


def _xpwe_parse_ep(root: ET.Element) -> tuple[list[dict], dict[int, dict]]:
    ep_list: list[dict] = []
    ep_by_id: dict[int, dict] = {}
    for el in root.findall('.//PweElencoPrezzi/EPItem'):
        id_ep = _int(el.get('ID', '0'))
        dest  = (el.findtext('DesEstesa') or '').strip()
        rid   = (el.findtext('DesRidotta') or '').strip()
        brv   = (el.findtext('DesBreve') or '').strip()
        desc_best = dest or rid or brv
        item = {
            'id':           id_ep,
            'tariffa':      (el.findtext('Tariffa') or '').strip(),
            'articolo':     (el.findtext('Articolo') or '').strip(),
            'des_ridotta':  rid,
            'des_breve':    brv,
            'des_estesa':   desc_best,
            'um':           (el.findtext('UnMisura') or '').strip(),
            'prezzo':       _float(el.findtext('Prezzo1')),
            'prezzo_netto': 0.0,
            'inc_mdo':      _float(el.findtext('IncMDO')),
            'inc_sic':      _float(el.findtext('IncSIC')),
            'inc_mat':      _float(el.findtext('IncMAT')),
            'inc_attr':     _float(el.findtext('IncATTR')),
            'ribassabile':  True,
            'id_spcap':     _int(el.findtext('IDSpCap')),
            'id_cap':       _int(el.findtext('IDCap')),
            'id_sbcap':     _int(el.findtext('IDSbCap')),
            'flags':        _int(el.findtext('Flags')),
            'tag_bim':      (el.findtext('TagBIM') or '').strip(),
        }
        ep_list.append(item)
        ep_by_id[id_ep] = item
    return ep_list, ep_by_id


def _xpwe_parse_vc(root: ET.Element) -> list[dict[str, Any]]:
    items = []
    for vc_el in root.findall('.//PweVociComputo/VCItem'):
        id_vc = _int(vc_el.get('ID', '0'))
        misure = []
        for rg in vc_el.findall('.//PweVCMisure/RGItem'):
            misure.append({
                'id':          _int(rg.get('ID', '0')),
                'descrizione': (rg.findtext('Descrizione') or '').strip(),
                'pu':          _float(rg.findtext('PartiUguali')),
                'lu':          _float(rg.findtext('Lunghezza')),
                'la':          _float(rg.findtext('Larghezza')),
                'hp':          _float(rg.findtext('HPeso')),
                'qt':          _float(rg.findtext('Quantita')),
                'rif_voce':    _int(rg.findtext('IDVV')),
                'flags':       _int(rg.findtext('Flags')),
            })
        items.append({
            'id':       id_vc,
            'id_ep':    _int(vc_el.findtext('IDEP')),
            'quantita': _float(vc_el.findtext('Quantita')),
            'importo':  0.0,
            'data_mis': (vc_el.findtext('DataMis') or '').strip(),
            'id_spcal': _int(vc_el.findtext('IDSpCat')),
            'id_cat':   _int(vc_el.findtext('IDCat')),
            'id_sbcat': _int(vc_el.findtext('IDSbCat')),
            'cod_wbs':  (vc_el.findtext('CodiceWBS') or '').strip(),
            'flags':    _int(vc_el.findtext('Flags')),
            'misure':   misure,
        })
    return items


def _xpwe_parse_dg(root: ET.Element) -> dict[str, Any]:
    dg = root.find('.//PweDGProgetto/PweDGDatiGenerali')
    if dg is None:
        return {}
    return {
        'oggetto':     (dg.findtext('Oggetto') or '').strip(),
        'committente': (dg.findtext('Committente') or '').strip(),
        'comune':      (dg.findtext('Comune') or '').strip(),
        'provincia':   (dg.findtext('Provincia') or '').strip(),
        'impresa':     (dg.findtext('Impresa') or '').strip(),
        'operatore':   '',
        'perc_prezzi': _float(dg.findtext('PercPrezzi')),
        'parte_opera': (dg.findtext('ParteOpera') or '').strip(),
    }


# ===========================================================================
# API PUBBLICA
# ===========================================================================

def parse_dcf(dcf_path: str) -> dict[str, Any]:
    """
    Parsa un file .dcf PriMus (formato binario SFS).

    NOTA: I file .dcf di tipo "Contabilità" (TpDoc=2) hanno i blocchi dati
    crittografati e non sono estraibili. In quel caso restituisce un dizionario
    con doc['_encrypted']=True. Usare parse_xpwe() sull'export .xpwe.
    """
    blocks = _sfs_extract_xml(dcf_path)

    _log(f'parse_dcf: blocchi estratti = {list(blocks.keys())}')
    for tag, txt in blocks.items():
        if isinstance(txt, str):
            _log(f'  {tag}: {len(txt)} chars')

    # Rileva DRM ACCA (dati XML cifrati con protezione proprietaria)
    if blocks.get('_drm') == 'ACCA':
        _log(f'parse_dcf: DRM ACCA rilevato — {blocks.get("_drm_count",0)} blocchi cifrati')
        _ind_end()
        return {
            'formato':          'dcf',
            '_drm':             True,
            '_tipo_documento':  'drm_acca',
            'info':             {},
            'quadro_economico': [],
            'super_categorie':  {},
            'categorie':        {},
            'sotto_categorie':  {},
            'super_capitoli':   {},
            'capitoli':         {},
            'sotto_capitoli':   {},
            'elenco_prezzi':    [],
            'computo':          [],
            'strutture_stampa': [],
            '_ep_by_id':        {},
            '_raw_xml':         blocks,
        }

    info_doc = blocks.get('InfoDoc', '')
    tp_doc_match = re.search(r'TpDoc="(\d+)"', info_doc)
    tp_doc = int(tp_doc_match.group(1)) if tp_doc_match else 0
    has_data = bool(blocks.get('DatiGenerali') or blocks.get('CollectionEP'))
    _log(f'parse_dcf: TpDoc={tp_doc}  has_data={has_data}')

    if tp_doc == 2 and not has_data:
        _ind_end()
        return {
            'formato':          'dcf',
            '_encrypted':       True,
            '_tipo_documento':  'contabilita',
            '_info_doc':        info_doc,
            'info':             {},
            'quadro_economico': [],
            'super_categorie':  {},
            'categorie':        {},
            'sotto_categorie':  {},
            'super_capitoli':   {},
            'capitoli':         {},
            'sotto_capitoli':   {},
            'elenco_prezzi':    [],
            'computo':          [],
            'strutture_stampa': [],
            '_ep_by_id':        {},
            '_raw_xml':         blocks,
        }

    dg = _dcf_parse_dati_generali(blocks.get('DatiGenerali', ''))
    ep_list, ep_by_id_map = _dcf_parse_ep(blocks.get('CollectionEP', ''))

    # Rileva DCF con dati XML insufficienti (PriMus non ha esportato l'XML embedded)
    # In questo caso il DCF esiste ma i blocchi sono quasi vuoti (< 1KB)
    ep_xml_size = len(blocks.get('CollectionEP', ''))
    dg_xml_size = len(blocks.get('DatiGenerali', ''))
    if ep_xml_size < 500 and dg_xml_size < 500:
        _log(f'parse_dcf: ATTENZIONE dati XML insufficienti '
             f'(CollectionEP={ep_xml_size} chars, DatiGenerali={dg_xml_size} chars)')
        _log('parse_dcf: il DCF non contiene XML embedded leggibile.')
        _log('parse_dcf: Soluzione -> esportare il file XPWE da PriMus e usare parse_xpwe()')
        _ind_end()
        return {
            'formato':          'dcf',
            '_no_xml':          True,
            '_ep_xml_size':     ep_xml_size,
            '_dg_xml_size':     dg_xml_size,
            'info':             {},
            'quadro_economico': [],
            'super_categorie':  {},
            'categorie':        {},
            'sotto_categorie':  {},
            'super_capitoli':   {},
            'capitoli':         {},
            'sotto_capitoli':   {},
            'elenco_prezzi':    [],
            'computo':          [],
            'strutture_stampa': [],
            '_ep_by_id':        {},
            '_raw_xml':         blocks,
        }

    doc: dict[str, Any] = {
        'formato':          'dcf',
        'info':             dg.get('info', {}),
        'quadro_economico': dg.get('quadro_economico', []),
        'super_categorie':  dg.get('super_categorie', {}),
        'categorie':        dg.get('categorie', {}),
        'sotto_categorie':  dg.get('sotto_categorie', {}),
        'super_capitoli':   dg.get('super_capitoli', {}),
        'capitoli':         dg.get('capitoli', {}),
        'sotto_capitoli':   dg.get('sotto_capitoli', {}),
        'elenco_prezzi':    ep_list,
        'computo':          _dcf_parse_vc(blocks.get('CollectionVC', '')),
        'strutture_stampa': _dcf_parse_st(blocks.get('CollectionST', '')),
        '_ep_by_id':        ep_by_id_map,
        '_raw_xml':         blocks,
    }
    _ind_end()
    return doc


def analyze_dcf(dcf_path: str) -> dict[str, Any]:
    """
    Analizza un file .dcf e restituisce gli stessi dati di parse_dcf
    con l'aggiunta della chiave '_blocchi' (lista dei tag XML trovati).
    Usata da LeenO per diagnostica e import guidato.
    """
    doc = parse_dcf(dcf_path)
    doc['_blocchi'] = list(doc.get('_raw_xml', {}).keys())
    return doc


def parse_xpwe(xpwe_path: str) -> dict[str, Any]:
    """Parsa un file .xpwe (formato standard XPWE 5.05 — XML puro)."""
    with open(xpwe_path, 'rb') as f:
        raw = f.read()

    enc_match = re.search(rb'encoding=["\']([^"\']+)["\']', raw[:200])
    enc = enc_match.group(1).decode('ascii') if enc_match else 'utf-8'
    try:
        xml_text = raw.decode(enc)
    except (UnicodeDecodeError, LookupError):
        xml_text = raw.decode('utf-8', errors='replace')

    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as e:
        raise ValueError(f'File XPWE non valido: {e}')

    cats = _xpwe_parse_categorie(root)
    ep_list, ep_by_id_map = _xpwe_parse_ep(root)

    doc: dict[str, Any] = {
        'formato':          'xpwe',
        'info':             _xpwe_parse_dg(root),
        'quadro_economico': [],
        'super_categorie':  cats['super_categorie'],
        'categorie':        cats['categorie'],
        'sotto_categorie':  cats['sotto_categorie'],
        'super_capitoli':   cats['super_capitoli'],
        'capitoli':         cats['capitoli'],
        'sotto_capitoli':   cats['sotto_capitoli'],
        'elenco_prezzi':    ep_list,
        'computo':          _xpwe_parse_vc(root),
        'strutture_stampa': [],
        '_ep_by_id':        ep_by_id_map,
        '_raw_xml':         {},
    }
    return doc


def parse_auto(path: str) -> dict[str, Any]:
    """Rileva automaticamente il formato e chiama parse_dcf o parse_xpwe."""
    ext = path.lower().rsplit('.', 1)[-1]
    if ext == 'dcf':
        return parse_dcf(path)
    if ext in ('xpwe', 'xml'):
        with open(path, 'rb') as f:
            header = f.read(512)
        if b'PweDocumento' in header or b'EPItem' in header or b'VCItem' in header:
            return parse_xpwe(path)
    with open(path, 'rb') as f:
        magic = f.read(8)
    if magic == b'AAMVHFSS':
        return parse_dcf(path)
    return parse_xpwe(path)


# ===========================================================================
# FUNZIONI DI UTILITA'
# ===========================================================================

def ep_by_id(doc: dict) -> dict[int, dict]:
    """Restituisce la mappa id->voce EP."""
    return doc.get('_ep_by_id', {ep['id']: ep for ep in doc['elenco_prezzi']})


def computo_con_descrizioni(doc: dict) -> list[dict]:
    """
    Arricchisce ogni voce del computo con EP e categorie.
    Restituisce list di dict flat pronti per export / LeenO.
    """
    ep_idx = ep_by_id(doc)
    cats   = doc.get('categorie', {})
    spcats = doc.get('super_categorie', {})

    rows = []
    for vc in doc['computo']:
        ep = ep_idx.get(vc['id_ep'], {})
        cat_nome   = cats.get(vc['id_cat'], {}).get('nome', '')
        spcat_nome = spcats.get(vc['id_spcal'], {}).get('nome', '')
        cat_display = cat_nome or spcat_nome
        importo = vc.get('importo') or round(vc['quantita'] * ep.get('prezzo', 0.0), 2)
        rows.append({
            'id_vc':       vc['id'],
            'id_ep':       vc['id_ep'],
            'tariffa':     ep.get('tariffa', ''),
            'super_cat':   spcat_nome,
            'categoria':   cat_display,
            'descrizione': ep.get('des_estesa', '') or ep.get('des_ridotta', ''),
            'um':          ep.get('um', ''),
            'prezzo_unit': ep.get('prezzo', 0.0),
            'quantita':    vc['quantita'],
            'importo':     round(importo, 2),
            'data_mis':    vc.get('data_mis', ''),
            'n_misure':    len(vc.get('misure', [])),
            'misure':      vc.get('misure', []),
        })
    return rows


# ===========================================================================
# CLI
# ===========================================================================

if __name__ == '__main__':
    import sys
    import json
    import tkinter as tk
    from tkinter import scrolledtext

    def mostra_risultati(titolo: str, testo: str) -> None:
        """Mostra i risultati in una finestra con testo selezionabile."""
        win = tk.Tk()
        win.title(titolo)
        win.attributes('-topmost', True)
        win.resizable(True, True)
        txt = scrolledtext.ScrolledText(win, wrap=tk.WORD, width=90, height=30,
                                        font=('Consolas', 9))
        txt.insert(tk.END, testo)
        txt.config(state=tk.DISABLED)
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        btn = tk.Button(win, text='Chiudi', command=win.destroy, width=10)
        btn.pack(pady=(0, 8))
        win.mainloop()

    if len(sys.argv) < 2:
        mostra_risultati('dcf_parser',
                         'Uso: python dcf_parser.py <file.dcf|file.xpwe> [--json]')
        sys.exit(1)

    for arg in sys.argv[1:]:
        if arg.startswith('--log='):
            LOG_FILE = arg[6:]

    log_reset()

    path = sys.argv[1]
    _log(f'parse_auto({path!r})')

    try:
        doc = parse_auto(path)
    except Exception as e:
        mostra_risultati('dcf_parser — Errore', f'Errore durante il parsing:\n\n{e}')
        sys.exit(1)

    if '--json' in sys.argv:
        out = {k: v for k, v in doc.items() if not k.startswith('_')}
        sys.stdout.write(json.dumps(out, ensure_ascii=False, indent=2))
        sys.exit(0)

    # Costruisci il testo di riepilogo
    info = doc['info']
    fmt  = doc['formato'].upper()
    righe = []
    righe.append('=' * 62)
    righe.append(f'  [{fmt}]  {os.path.basename(path)}')
    righe.append('=' * 62)
    righe.append(f'  Oggetto      : {info.get("oggetto", "N/D")}')
    righe.append(f'  Committente  : {info.get("committente", "")}')
    righe.append(f'  Comune       : {info.get("comune", "")}')
    righe.append(f'  Importo      : EUR {info.get("importo_lavori", 0):,.2f}')
    righe.append('')
    righe.append(f'  Elenco prezzi : {len(doc["elenco_prezzi"])} voci')
    righe.append(f'  Computo       : {len(doc["computo"])} voci')
    righe.append(f'  Categorie     : {len(doc["categorie"])}')
    if doc.get('strutture_stampa'):
        righe.append(f'  Strutture ST  : {len(doc["strutture_stampa"])}')

    if doc.get('_no_xml'):
        righe.append('')
        righe.append('  ATTENZIONE: dati XML insufficienti.')
        righe.append('  Soluzione: esportare il file XPWE da PriMus.')
    else:
        righe.append('')
        righe.append('Categorie:')
        for id_, cat in sorted(doc['super_categorie'].items()):
            imp = f'  -> EUR {cat["importo"]:,.2f}' if cat['importo'] else ''
            righe.append(f'  [{id_}] {cat["nome"]}{imp}')
        for id_, cat in sorted(doc['categorie'].items()):
            imp = f'  -> EUR {cat["importo"]:,.2f}' if cat['importo'] else ''
            righe.append(f'    [{id_}] {cat["nome"]}{imp}')

        righe.append('')
        righe.append('Elenco prezzi (prime 8):')
        for ep in doc['elenco_prezzi'][:8]:
            righe.append(f'  {ep["tariffa"]:20s} {ep["um"]:5s} '
                         f'EUR{ep["prezzo"]:>10.2f}  {ep["des_estesa"][:60]}')
        if len(doc['elenco_prezzi']) > 8:
            righe.append(f'  ... e altre {len(doc["elenco_prezzi"])-8} voci')

        righe.append('')
        righe.append('Computo (prime 5):')
        for r in computo_con_descrizioni(doc)[:5]:
            righe.append(f'  [{r["categoria"][:18]:18s}] {r["tariffa"]:18s} '
                         f'{r["um"]:4s} Qt={r["quantita"]:8.3f}  '
                         f'EUR{r["importo"]:>10.2f}  {r["descrizione"][:40]}')

    righe.append('')
    righe.append(f'Log: {LOG_FILE}')

    testo = '\n'.join(righe)
    _log(testo)
    mostra_risultati(f'dcf_parser — {os.path.basename(path)}', testo)
