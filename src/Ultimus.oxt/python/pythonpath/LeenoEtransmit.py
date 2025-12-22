########################################################################
# e-Transmit per LibreOffice Calc
########################################################################
from pydoc import doc
import Dialogs
import LeenoUtils
# import uno
import zipfile
import tempfile
import shutil
import re
from pathlib import Path
from uno import fileUrlToSystemPath, systemPathToFileUrl
from com.sun.star.beans import PropertyValue
# --------------------------------------------------
# Utility: normalizzazione link cross-platform
# --------------------------------------------------
def normalize_link_target(target):
    target = target.strip()
    # URL file:///...
    if target.startswith("file:///"):
        return Path(fileUrlToSystemPath(target))
    # POSIX absolute (/home/... o /Volumes/...)
    if target.startswith("/"):
        return Path(target)
    # Windows absolute (C:\...)
    if re.match(r"^[A-Za-z]:[\\/]", target):
        return Path(target)
    return None

# --------------------------------------------------
# Verifica se un path è dentro una cartella leeno-bk
# --------------------------------------------------
def is_in_leeno_bk(path):
    """Controlla se il path si trova dentro una cartella leeno-bk"""
    try:
        for parent in path.parents:
            if parent.name == "leeno-bk":
                return True
        return path.name == "leeno-bk"
    except Exception:
        return False

# --------------------------------------------------
# Estrazione link dal documento
# --------------------------------------------------
def extract_links(doc):
    links = set()
    for sheet in doc.Sheets:
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        rows = cursor.RangeAddress.EndRow + 1
        cols = cursor.RangeAddress.EndColumn + 1
        for r in range(rows):
            for c in range(cols):
                cell = sheet.getCellByPosition(c, r)
                # HyperLinkURL
                try:
                    psi = cell.getPropertySetInfo()
                    if psi.hasPropertyByName("HyperLinkURL"):
                        url = cell.getPropertyValue("HyperLinkURL")
                        if url:
                            p = normalize_link_target(url)
                            if p:
                                links.add(p)
                except Exception:
                    pass
                # Formula HYPERLINK("...")
                try:
                    formula = cell.Formula
                    if formula and formula.upper().startswith("=HYPERLINK"):
                        m = re.search(r'HYPERLINK\("([^"]+)"', formula, re.IGNORECASE)
                        if m:
                            p = normalize_link_target(m.group(1))
                            if p:
                                links.add(p)
                except Exception:
                    pass
    return links

def highlight_missing_links(doc):
    """
    Colora di rosso (#FF0000) le celle che contengono link a file mancanti.
    Funziona sia su HyperLinkURL sia su formule HYPERLINK.
    """
    RED = 0xFF0000
    for sheet in doc.Sheets:
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        rows = cursor.RangeAddress.EndRow + 1
        cols = cursor.RangeAddress.EndColumn + 1
        for r in range(rows):
            for c in range(cols):
                cell = sheet.getCellByPosition(c, r)
                missing = False
                # --- HyperLinkURL
                try:
                    psi = cell.getPropertySetInfo()
                    if psi.hasPropertyByName("HyperLinkURL"):
                        url = cell.getPropertyValue("HyperLinkURL")
                        p = normalize_link_target(url)
                        if p and not p.exists():
                            missing = True
                except Exception:
                    pass
                # --- Formula HYPERLINK
                try:
                    formula = cell.Formula
                    if formula and formula.upper().startswith("=HYPERLINK"):
                        m = re.search(r'HYPERLINK\("([^"]+)"', formula, re.IGNORECASE)
                        if m:
                            p = normalize_link_target(m.group(1))
                            if p and not p.exists():
                                missing = True
                except Exception:
                    pass
                # --- Evidenzia se manca
                if missing:
                    cell.CellBackColor = RED
    return

# --------------------------------------------------
# Comando principale e-Transmit
# --------------------------------------------------
def e_transmit_calc(*args):
    doc = LeenoUtils.getDocument()
    if not doc.hasLocation():
        Dialogs.Info(Title="Errore", Text="Salvare il file prima di creare il pacchetto.")
        return
    # attiva la progressbar
    indicator = doc.getCurrentController().getStatusIndicator()
    indicator.start("Creazione pacchetto e-Transmit...", 100)
    ods_path = Path(fileUrlToSystemPath(doc.URL))
    zip_path = ods_path.with_name(f"{ods_path.stem}_trasmissione.zip")
    indicator.setValue(10)
    linked_files = extract_links(doc)
    # --- cartella temporanea
    tmpdir = Path(tempfile.mkdtemp())
    allegati_dir = tmpdir / "allegati"
    allegati_dir.mkdir()
    ods_copy = tmpdir / ods_path.name
    shutil.copy2(ods_path, ods_copy)
    # --- copia file linkati, esclusi cartelle e file in leeno-bk
    indicator.setValue(20)
    link_map = {}
    skipped_dirs = []
    skipped_leeno_bk = []
    missing = []

    for f in linked_files:
        # Salta file dentro cartelle leeno-bk
        if is_in_leeno_bk(f):
            skipped_leeno_bk.append(f)
            continue

        if not f.exists():
            missing.append(f)
            continue
        if f.is_dir():
            skipped_dirs.append(f)
            continue
        if not f.is_file():
            continue
        dest = allegati_dir / f.name
        shutil.copy2(f, dest)
        link_map[f] = Path("allegati") / f.name

    # --- apri la copia ODS e riscrivi link relativi
    indicator.setValue(50)
    desktop = LeenoUtils.getDesktop()
    ods_url = systemPathToFileUrl(str(ods_copy))
    props = (
        PropertyValue("Hidden", 0, True, 0),
        PropertyValue("ReadOnly", 0, False, 0),
    )
    doc2 = desktop.loadComponentFromURL(ods_url, "_blank", 0, props)
    try:
        for sheet in doc2.Sheets:
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(False)
            rows = cursor.RangeAddress.EndRow + 1
            cols = cursor.RangeAddress.EndColumn + 1
            for r in range(rows):
                for c in range(cols):
                    cell = sheet.getCellByPosition(c, r)
                    # HyperLinkURL
                    try:
                        psi = cell.getPropertySetInfo()
                        if psi.hasPropertyByName("HyperLinkURL"):
                            url = cell.getPropertyValue("HyperLinkURL")
                            if url:
                                p = normalize_link_target(url)
                                if p in link_map:
                                    cell.HyperLinkURL = link_map[p].as_posix()
                    except Exception:
                        pass
                    # Formula HYPERLINK
                    try:
                        formula = cell.Formula
                        if formula and formula.upper().startswith("=HYPERLINK"):
                            m = re.search(r'HYPERLINK\("([^"]+)"', formula, re.IGNORECASE)
                            if m:
                                p = normalize_link_target(m.group(1))
                                if p in link_map:
                                    cell.Formula = f'=HYPERLINK("{link_map[p].as_posix()}";"{cell.String}")'
                    except Exception:
                        pass
        doc2.store()
    finally:
        doc2.close(True)

    # --- crea ZIP
    indicator.setValue(80)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in tmpdir.rglob("*"):
            if p.is_file():
                zf.write(p, p.relative_to(tmpdir))

    # --- crea manifest con cartelle escluse e file mancanti
    indicator.setValue(90)
    manifest_content = ""
    if skipped_dirs or skipped_leeno_bk or missing:
        log = []
        # if skipped_leeno_bk:
        #     log.append("FILE IN CARTELLE LEENO-BK (ESCLUSI):\n")
        #     log.extend(str(p) for p in skipped_leeno_bk)
        #     log.append("")
        # if skipped_dirs:
        #     log.append("CARTELLE NON INCLUSE:\n")
        #     log.extend(str(p) for p in skipped_dirs)
        #     log.append("")
        if missing:
            log.append("FILE NON TROVATI:")
            log.extend('- ' + str(p) for p in missing)
        manifest_content = "\n".join(log)
        manifest_path = tmpdir / "manifest.txt"
        manifest_path.write_text(manifest_content, encoding="utf-8")
        # Aggiungi il manifest allo ZIP
        with zipfile.ZipFile(zip_path, "a", zipfile.ZIP_DEFLATED) as zf:
            zf.write(manifest_path, "manifest.txt")

    shutil.rmtree(tmpdir)
    indicator.setValue(100)
    indicator.end()
    highlight_missing_links(doc)

    # --- Messaggio finale con report
    default_text = "----\nTutti i file linkati sono stati inclusi correttamente."
    manifest_content = manifest_content + "\n----\nLe celle con link a file mancanti sono state evidenziate in rosso." if manifest_content else ""
    final_message = f"Pacchetto creato:\n{zip_path}\n\n{manifest_content or default_text}"
    Dialogs.Info(Title="eTransmit completato", Text=final_message)

    # Apri la cartella contenente il file ZIP
    try:
        apri = LeenoUtils.createUnoService("com.sun.star.system.SystemShellExecute")
        zip_url = systemPathToFileUrl(str(zip_path.parent))
        apri.execute(zip_url, "", 0)
    except Exception:
        pass
    return
