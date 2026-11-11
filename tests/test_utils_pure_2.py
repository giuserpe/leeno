import ast
import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

BASE_DIR = Path(__file__).resolve().parent.parent
LEENO_EXPORT_PATH = BASE_DIR / "src/Ultimus.oxt/python/pythonpath/LeenoExport.py"
DIALOGS_PATH = BASE_DIR / "src/Ultimus.oxt/python/pythonpath/Dialogs.py"
DEBUG_PATH = BASE_DIR / "src/Ultimus.oxt/python/pythonpath/Debug.py"
LEENO_ISTAT_PATH = BASE_DIR / "src/Ultimus.oxt/python/pythonpath/LeenoISTAT.py"
DISPATCHER_PATH = BASE_DIR / "src/Ultimus.oxt/python/LeenoDispatcher.py"
LEENO_IMPORT_PATH = BASE_DIR / "src/Ultimus.oxt/python/pythonpath/LeenoImport.py"
LEENO_UTILS_PATH = BASE_DIR / "src/Ultimus.oxt/python/pythonpath/LeenoUtils.py"


def _carica_utils_date():
    import calendar
    import datetime
    from datetime import date
    tree = ast.parse(LEENO_UTILS_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.Assign):
            for t in n.targets:
                if isinstance(t, ast.Name) and t.id in ("DAYNAMES", "MONTHNAMES"):
                    nodi.append(n)
        elif isinstance(n, ast.FunctionDef) and n.name in ("daysInMonth", "firstWeekDay", "date2String", "string2Date"):
            nodi.append(n)
    ns = {
        "calendar": calendar,
        "datetime": datetime,
        "date": date,
    }
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(LEENO_UTILS_PATH), "exec"), ns)
    return ns


def _carica_export():
    tree = ast.parse(LEENO_EXPORT_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name in ("clean_markdown", "split_markdown_table"):
            nodi.append(n)
    ns = {}
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(LEENO_EXPORT_PATH), "exec"), ns)
    return ns


def _carica_dialogs():
    tree = ast.parse(DIALOGS_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name in ("shortenPath", "_mapBasicYesNo", "storeLastPath"):
            nodi.append(n)
    ns = {
        "os": os,
        "YES": 2,
        "NO": 3,
        "getEditBox": None,
        "Config": None,
    }
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(DIALOGS_PATH), "exec"), ns)
    return ns


def _carica_debug():
    tree = ast.parse(DEBUG_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name == "_format_time":
            nodi.append(n)
    ns = {}
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(DEBUG_PATH), "exec"), ns)
    return ns


def _carica_istat():
    tree = ast.parse(LEENO_ISTAT_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id in ("TOL_CODICI", "_TOL_BY_CODE") for t in n.targets
        ):
            nodi.append(n)
        elif isinstance(n, ast.FunctionDef) and n.name == "_codice_sdmx":
            nodi.append(n)
    ns = {}
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(LEENO_ISTAT_PATH), "exec"), ns)
    return ns


def _carica_dispatcher():
    tree = ast.parse(DISPATCHER_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name in ("cerca_path_valido", "fixPythonPath"):
            nodi.append(n)
    ns = {
        "os": os,
        "sys": sys,
        "inspect": __import__("inspect"),
    }
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(DISPATCHER_PATH), "exec"), ns)
    return ns


def _carica_import():
    tree = ast.parse(LEENO_IMPORT_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name == "_readXmlFile":
            nodi.append(n)
    ns = {}
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(LEENO_IMPORT_PATH), "exec"), ns)
    return ns


class TestUtilsPure2(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.export_ns = _carica_export()
        cls.dialogs_ns = _carica_dialogs()
        cls.debug_ns = _carica_debug()
        cls.istat_ns = _carica_istat()
        cls.dispatcher_ns = _carica_dispatcher()
        cls.import_ns = _carica_import()

    # -------------------------------------------------------------------------
    # LeenoExport.py: clean_markdown, split_markdown_table
    # -------------------------------------------------------------------------
    def test_clean_markdown(self):
        clean_markdown = self.export_ns["clean_markdown"]

        # Valore None restituisce stringa vuota
        self.assertEqual(clean_markdown(None), "")

        # Valori non stringa (es. interi, float) vengono convertiti
        self.assertEqual(clean_markdown(123), "123")
        self.assertEqual(clean_markdown(45.6), "45.6")

        # Pipe '|' viene escapato come '\|'
        self.assertEqual(clean_markdown("Colonna 1 | Colonna 2"), "Colonna 1 \\| Colonna 2")

        # Spazi iniziali e finali vengono rimossi tramite strip()
        self.assertEqual(clean_markdown("  testo con spazi   "), "testo con spazi")

        # Nuove righe (\r\n, \n, \r) vengono sostituite da ' <br> ' e pipe escapato
        self.assertEqual(clean_markdown("Riga 1\r\nRiga 2\nRiga 3\rRiga 4"), "Riga 1 <br> Riga 2 <br> Riga 3 <br> Riga 4")
        self.assertEqual(clean_markdown("Test | con | pipe\ne a capo"), "Test \\| con \\| pipe <br> e a capo")

    def test_split_markdown_table_sotto_limite(self):
        split_markdown_table = self.export_ns["split_markdown_table"]

        header = "| Codice | Descrizione | U.M. | Prezzo |"
        delim = "|---|---|---|---|"
        body = [
            "| E.01 | Scavo di sbancamento | m3 | 12.50 |",
            "| E.02 | Calcestruzzo magrone | m3 | 85.00 |",
        ]
        parti = split_markdown_table(header, delim, body)
        self.assertEqual(len(parti), 1)
        atteso = ("| Codice | Descrizione | U.M. | Prezzo |\n"
                  "|---|---|---|---|\n"
                  "| E.01 | Scavo di sbancamento | m3 | 12.50 |\n"
                  "| E.02 | Calcestruzzo magrone | m3 | 85.00 |\n").encode("utf-8")
        self.assertEqual(parti[0], atteso)

    def test_split_markdown_table_senza_body(self):
        split_markdown_table = self.export_ns["split_markdown_table"]

        header = "| Col 1 | Col 2 |"
        delim = "|---|---|"
        parti = split_markdown_table(header, delim, [])
        self.assertEqual(len(parti), 1)
        atteso = "| Col 1 | Col 2 |\n|---|---|\n".encode("utf-8")
        self.assertEqual(parti[0], atteso)

    def test_split_markdown_table_superamento_limite(self):
        split_markdown_table = self.export_ns["split_markdown_table"]

        header = "| A | B |"
        delim = "|---|---|"
        body = [
            "| riga 1 | val |",
            "| riga 2 | val |",
            "| riga 3 | val |",
        ]
        # Con limiti ridotti per testare lo splittamento
        # header_prefix e' di 18 byte. Con limit_3mb=50 e limit_2mb=40, ciascuna riga body e' ~17 byte
        parti = split_markdown_table(header, delim, body, limit_3mb=50, limit_2mb=40)
        self.assertEqual(len(parti), 3)

        prefisso = b"| A | B |\n|---|---|\n"
        self.assertEqual(parti[0], prefisso + b"| riga 1 | val |\n")
        self.assertEqual(parti[1], prefisso + b"| riga 2 | val |\n")
        self.assertEqual(parti[2], prefisso + b"| riga 3 | val |\n")

    def test_split_markdown_table_riga_singola_eccedente(self):
        split_markdown_table = self.export_ns["split_markdown_table"]

        header = "| H |"
        delim = "|---|"
        riga_gigante = "| " + ("x" * 60) + " |"
        riga_normale = "| norm |"
        body = [riga_gigante, riga_normale]

        # limit_2mb=30: riga_gigante eccede 30 byte da sola
        parti = split_markdown_table(header, delim, body, limit_3mb=50, limit_2mb=30)
        self.assertEqual(len(parti), 2)
        prefisso = b"| H |\n|---|\n"
        self.assertEqual(parti[0], prefisso + riga_gigante.encode("utf-8") + b"\n")
        self.assertEqual(parti[1], prefisso + riga_normale.encode("utf-8") + b"\n")

    def test_split_markdown_table_celle_vuote_e_barre(self):
        split_markdown_table = self.export_ns["split_markdown_table"]

        header = "| Desc | Info | Extra |"
        delim = "|---|---|---|"
        body = [
            "| | | |",
            "| Test con \\| barra | | val |",
        ]
        parti = split_markdown_table(header, delim, body)
        self.assertEqual(len(parti), 1)
        testo = parti[0].decode("utf-8")
        self.assertIn("| | | |", testo)
        self.assertIn("| Test con \\| barra | | val |", testo)

    # -------------------------------------------------------------------------
    # Dialogs.py: shortenPath, _mapBasicYesNo, storeLastPath
    # -------------------------------------------------------------------------
    def test_shorten_path(self):
        shortenPath = self.dialogs_ns["shortenPath"]

        # Se pth e' None ritorna None
        self.assertIsNone(shortenPath(None, 100))

        # Mock di getEditBox: larghezza proporzionale a len(testo) * 10
        mock_get_box = lambda s: (len(s) * 10, 20)
        self.dialogs_ns["getEditBox"] = mock_get_box

        # Path che rientra nella larghezza (len=11, width_needed=110 <= 200)
        self.assertEqual(shortenPath("/short/path", 200), "/short/path")

        # Path lungo che necessita di accorciamento
        # con width=150, max caratteri ammessi = 15. "...ome/file.txt" = 15 caratteri -> 150 pixel
        risultato = shortenPath("/very/long/path/to/some/file.txt", 150)
        self.assertEqual(risultato, "...ome/file.txt")

        # Larghezza insufficiente anche per accorciare: ritorna '...'
        self.assertEqual(shortenPath("/very/long/path/to/some/file.txt", 20), "...")

    def test_map_basic_yes_no(self):
        _mapBasicYesNo = self.dialogs_ns["_mapBasicYesNo"]
        YES = self.dialogs_ns["YES"]
        NO = self.dialogs_ns["NO"]

        # Costante YES -> 1
        self.assertEqual(_mapBasicYesNo(YES), 1)
        # Costante NO -> 0
        self.assertEqual(_mapBasicYesNo(NO), 0)
        # Altro valore restituito inalterato
        self.assertEqual(_mapBasicYesNo(4), 4)
        self.assertEqual(_mapBasicYesNo(None), None)

    def test_store_last_path(self):
        storeLastPath = self.dialogs_ns["storeLastPath"]

        salvati = {}
        class MockConfig:
            def write(self, section, key, val):
                salvati[(section, key)] = val

        self.dialogs_ns["Config"] = MockConfig

        # Passando il percorso di un file, estrae la cartella padre e garantisce separatore finale
        storeLastPath("/home/user/documenti/computo.ods")
        self.assertEqual(salvati.get(("Generale", "ultimo_percorso")), "/home/user/documenti/")

    # -------------------------------------------------------------------------
    # Debug.py: _format_time
    # -------------------------------------------------------------------------
    def test_format_time(self):
        _format_time = self.debug_ns["_format_time"]

        # Zero
        self.assertEqual(_format_time(0), "0 μs")

        # Microsecondi (< 0.001 s)
        self.assertEqual(_format_time(0.000005), "5 μs")
        self.assertEqual(_format_time(0.0005), "500 μs")

        # Valore limite 0.001 s (passa a millisecondi)
        self.assertEqual(_format_time(0.001), "1.00 ms")
        self.assertEqual(_format_time(0.5), "500.00 ms")
        self.assertEqual(_format_time(0.999), "999.00 ms")

        # Valore limite 1.0 s (passa a secondi)
        self.assertEqual(_format_time(1.0), "1.00 s")
        self.assertEqual(_format_time(15.25), "15.25 s")
        self.assertEqual(_format_time(59.99), "59.99 s")

        # Valore limite 60.0 s (passa a minuti e secondi)
        self.assertEqual(_format_time(60.0), "1m 0.00s")
        self.assertEqual(_format_time(125.5), "2m 5.50s")

        # Valori >= 3600 secondi (1 ora):
        self.assertEqual(_format_time(3600.0), "1h 0m 0.00s")
        self.assertEqual(_format_time(3665.0), "1h 1m 5.00s")
        self.assertEqual(_format_time(7325.25), "2h 2m 5.25s")

    # -------------------------------------------------------------------------
    # LeenoISTAT.py: _codice_sdmx
    # -------------------------------------------------------------------------
    def test_codice_sdmx(self):
        _codice_sdmx = self.istat_ns["_codice_sdmx"]

        # Numeri interi 1-20
        self.assertEqual(_codice_sdmx(1), "BUIL_ART_NOHEPROT")
        self.assertEqual(_codice_sdmx(11), "AC_GAS_IRR_SEW")
        self.assertEqual(_codice_sdmx(20), "DELIV_WASTE_FAC_DIS_REC")

        # Stringhe numeriche con o senza spazi
        self.assertEqual(_codice_sdmx("1"), "BUIL_ART_NOHEPROT")
        self.assertEqual(_codice_sdmx(" 11 "), "AC_GAS_IRR_SEW")

        # Codici SDMX testuali validi
        self.assertEqual(_codice_sdmx("AC_GAS_IRR_SEW"), "AC_GAS_IRR_SEW")
        self.assertEqual(_codice_sdmx(" BUIL_ART_NOHEPROT "), "BUIL_ART_NOHEPROT")

        # Fuori intervallo 1-20 solleva ValueError
        with self.assertRaises(ValueError) as ctx:
            _codice_sdmx(0)
        self.assertIn("TOL 0 inesistente", str(ctx.exception))

        with self.assertRaises(ValueError) as ctx:
            _codice_sdmx(21)
        self.assertIn("TOL 21 inesistente", str(ctx.exception))

        # Codice alfabetico sconosciuto solleva ValueError
        with self.assertRaises(ValueError) as ctx:
            _codice_sdmx("CODICE_INESISTENTE")
        self.assertIn("Codice TOL sconosciuto", str(ctx.exception))

    # -------------------------------------------------------------------------
    # LeenoDispatcher.py: cerca_path_valido, fixPythonPath
    # -------------------------------------------------------------------------
    def test_cerca_path_valido_trovato(self):
        cerca_path_valido = self.dispatcher_ns["cerca_path_valido"]

        with tempfile.TemporaryDirectory() as td:
            finto_antigravity = os.path.join(td, "antigravity")
            with open(finto_antigravity, "w") as f:
                f.write("#!/bin/sh\n")

            with patch("os.getlogin", return_value="giuserpe"), \
                 patch("platform.system", return_value="Linux"), \
                 patch("os.path.exists", side_effect=lambda p: p == "/usr/bin/antigravity"):
                res = cerca_path_valido()
                self.assertEqual(res, "/usr/bin/antigravity")

            with patch("os.getlogin", return_value="giuserpe"), \
                 patch("platform.system", return_value="Linux"), \
                 patch("os.path.exists", side_effect=lambda p: p == "/usr/bin/code"):
                res_code = cerca_path_valido()
                self.assertEqual(res_code, "/usr/bin/code")

    def test_cerca_path_valido_non_trovato(self):
        cerca_path_valido = self.dispatcher_ns["cerca_path_valido"]

        with patch("os.getlogin", return_value="giuserpe"), \
             patch("platform.system", return_value="Linux"), \
             patch("os.path.exists", return_value=False):
            with self.assertRaises(FileNotFoundError) as ctx:
                cerca_path_valido()
            self.assertIn("Impossibile trovare un editor", str(ctx.exception))

    def test_cerca_path_valido_login_non_giuserpe(self):
        cerca_path_valido = self.dispatcher_ns["cerca_path_valido"]

        # BUG NOTO: cerca_path_valido ha una condizione 'if giuserpe in os.getlogin():' che racchiude
        # l'intera logica di ricerca. Se il login non contiene 'giuserpe', la funzione esce senza
        # eseguire il blocco e ritorna implicitamente None invece di cercare l'editor o sollevare FileNotFoundError.
        with patch("os.getlogin", return_value="mario"):
            risultato = cerca_path_valido()
            self.assertIsNone(risultato)

    def test_fix_python_path(self):
        fixPythonPath = self.dispatcher_ns["fixPythonPath"]

        old_path = list(sys.path)
        try:
            # Rimuove temporaneamente voci che potrebbero coincidere
            fixPythonPath()
            # Verifica che inspect.currentframe() determini il path e lo aggiunga a sys.path
            # In particolare, la cartella 'pythonpath' associata
            self.assertTrue(any(p.endswith("pythonpath") for p in sys.path))
        finally:
            sys.path[:] = old_path

    # -------------------------------------------------------------------------
    # LeenoImport.py: _readXmlFile
    # -------------------------------------------------------------------------
    def test_read_xml_file(self):
        _readXmlFile = self.import_ns["_readXmlFile"]

        with tempfile.TemporaryDirectory() as td:
            # 1. UTF-8 standard
            p_utf8 = os.path.join(td, "doc_utf8.xml")
            with open(p_utf8, "wb") as f:
                f.write("<?xml version=\"1.0\"?><testo>caffè</testo>".encode("utf-8"))
            self.assertEqual(_readXmlFile(p_utf8), "<?xml version=\"1.0\"?><testo>caffè</testo>")

            # 2. UTF-8 con BOM
            p_utf8_bom = os.path.join(td, "doc_utf8_bom.xml")
            with open(p_utf8_bom, "wb") as f:
                f.write(b"\xef\xbb\xbf" + "<?xml version=\"1.0\"?><testo>caffè bom</testo>".encode("utf-8"))
            self.assertEqual(_readXmlFile(p_utf8_bom), "<?xml version=\"1.0\"?><testo>caffè bom</testo>")

            # 3. UTF-16 con BOM (Little Endian)
            p_utf16_le = os.path.join(td, "doc_utf16_le.xml")
            with open(p_utf16_le, "wb") as f:
                f.write("<?xml version=\"1.0\"?><testo>caffè utf16</testo>".encode("utf-16"))
            # La funzione decodifica utf-16-le mantenendo il carattere BOM \ufeff all'inizio
            self.assertEqual(_readXmlFile(p_utf16_le), "\ufeff<?xml version=\"1.0\"?><testo>caffè utf16</testo>")

            # 4. UTF-16 Big Endian con BOM
            p_utf16_be = os.path.join(td, "doc_utf16_be.xml")
            with open(p_utf16_be, "wb") as f:
                f.write(b"\xfe\xff" + "<?xml version=\"1.0\"?><testo>caffè be</testo>".encode("utf-16-be"))
            self.assertEqual(_readXmlFile(p_utf16_be), "\ufeff<?xml version=\"1.0\"?><testo>caffè be</testo>")

            # 5. File vuoto
            p_empty = os.path.join(td, "doc_empty.xml")
            with open(p_empty, "wb") as f:
                f.write(b"")
            self.assertEqual(_readXmlFile(p_empty), "")

            # 6. File ZIP contenente XML
            p_zip = os.path.join(td, "archivio.zip")
            with zipfile.ZipFile(p_zip, "w") as zf:
                zf.writestr("dati.xml", "<?xml version=\"1.0\"?><testo>da zip</testo>".encode("utf-8"))
            self.assertEqual(_readXmlFile(p_zip), "<?xml version=\"1.0\"?><testo>da zip</testo>")

            # 7. Codifica cp1252 / Windows-1252 senza BOM: nessun carattere perso
            p_cp1252 = os.path.join(td, "doc_cp1252.xml")
            with open(p_cp1252, "wb") as f:
                f.write("<?xml version=\"1.0\"?><testo>caff\u00e8 \u00b0C</testo>".encode("cp1252"))
            self.assertEqual(_readXmlFile(p_cp1252), "<?xml version=\"1.0\"?><testo>caff\u00e8 \u00b0C</testo>")

            # 8. Codifica dichiarata nell'intestazione (ISO-8859-1) senza BOM
            p_iso = os.path.join(td, "doc_iso.xml")
            with open(p_iso, "wb") as f:
                f.write("<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?><testo>perch\u00e9</testo>".encode("iso-8859-1"))
            self.assertEqual(
                _readXmlFile(p_iso),
                "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?><testo>perch\u00e9</testo>")

            # 9. UTF-8 valido senza BOM: invariato
            p_u8 = os.path.join(td, "doc_utf8.xml")
            with open(p_u8, "wb") as f:
                f.write("<?xml version=\"1.0\"?><testo>caff\u00e8</testo>".encode("utf-8"))
            self.assertEqual(_readXmlFile(p_u8), "<?xml version=\"1.0\"?><testo>caff\u00e8</testo>")

    def test_date_utils_leeno_utils(self):
        """Testa daysInMonth, firstWeekDay, date2String e string2Date estratti da LeenoUtils."""
        import datetime
        mod = _carica_utils_date()
        daysInMonth = mod["daysInMonth"]
        firstWeekDay = mod["firstWeekDay"]
        date2String = mod["date2String"]
        string2Date = mod["string2Date"]

        # daysInMonth per anni bisestili e mesi standard
        d_feb_leap = datetime.date(2024, 2, 10)
        d_feb_non_leap = datetime.date(2023, 2, 10)
        d_apr = datetime.date(2023, 4, 15)
        d_dec = datetime.date(2023, 12, 1)

        self.assertEqual(daysInMonth(d_feb_leap), 29)
        self.assertEqual(daysInMonth(d_feb_non_leap), 28)
        self.assertEqual(daysInMonth(d_apr), 30)
        self.assertEqual(daysInMonth(d_dec), 31)

        # firstWeekDay: restituisce il weekday del primo giorno del mese (Lunedì=0, Domenica=6)
        # 1 Novembre 2023 era Mercoledì -> 2
        d_nov = datetime.date(2023, 11, 15)
        self.assertEqual(firstWeekDay(d_nov), 2)

        # date2String formati 0, 1, 2, 3
        d_test = datetime.date(2023, 5, 4)
        self.assertEqual(date2String(d_test, fmt=0), "4 Maggio 2023")
        self.assertEqual(date2String(d_test, fmt=1), "4/05/2023")
        self.assertEqual(date2String(d_test, fmt=2), "4-05-2023")
        self.assertEqual(date2String(d_test, fmt=3), "4.05.2023")

        # string2Date parsing con diversi separatori
        self.assertEqual(string2Date("04/05/2023"), datetime.date(2023, 5, 4))
        self.assertEqual(string2Date("04-05-2023"), datetime.date(2023, 5, 4))
        self.assertEqual(string2Date("04.05.2023"), datetime.date(2023, 5, 4))

        # string2Date senza separatori validi fallback a today
        self.assertEqual(string2Date("invalid_date"), datetime.date.today())

        # string2Date formato malformato con separatore solleva eccezione
        with self.assertRaises(Exception):
            string2Date("04/05")


if __name__ == "__main__":
    unittest.main()
