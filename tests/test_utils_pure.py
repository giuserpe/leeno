import ast
import calendar
import datetime
import textwrap
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

UTILS_PATH = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoUtils.py"
CONTAB_PATH = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoContab.py"
GLOBALS_PATH = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoGlobals.py"


def _carica_funzioni_utils():
    tree = ast.parse(UTILS_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name in (
            "daysInMonth",
            "firstWeekDay",
            "date2String",
            "string2Date",
            "countPdfPages",
            "replacePatternWithField",
            "wrap_text",
            "wrap_path",
        ):
            nodi.append(n)
        elif isinstance(n, ast.Assign):
            for t in n.targets:
                if isinstance(t, ast.Name) and t.id in ("DAYNAMES", "MONTHNAMES"):
                    nodi.append(n)
    ns = {
        "calendar": calendar,
        "date": datetime.date,
        "textwrap": textwrap,
        "PyPDF2": MagicMock(),
    }
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(UTILS_PATH), "exec"), ns)
    return ns


def _carica_funzioni_contab():
    tree = ast.parse(CONTAB_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name == "_col_letter":
            nodi.append(n)
    ns = {}
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(CONTAB_PATH), "exec"), ns)
    return ns


def _carica_funzioni_globals():
    tree = ast.parse(GLOBALS_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name in (
            "getGlobalVar",
            "setGlobalVar",
            "initGlobalVars",
        ):
            nodi.append(n)
    ns = {"sys": __import__("sys")}
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(GLOBALS_PATH), "exec"), ns)
    return ns


class MockTextCursor:
    def __init__(self, text):
        self.text = text
        self.pos = 0
        self.length = 0

    def collapseToStart(self):
        self.length = 0

    def goRight(self, count, expand):
        if expand:
            self.length += count
        else:
            self.pos += count

    @property
    def String(self):
        return self.text.String[self.pos : self.pos + self.length]

    @String.setter
    def String(self, val):
        self.text.String = (
            self.text.String[: self.pos] + val + self.text.String[self.pos + self.length :]
        )
        self.length = len(val)


class MockText:
    def __init__(self, text):
        self.String = text
        self.inserted = []

    def createTextCursor(self):
        return MockTextCursor(self)

    def insertTextContent(self, cursor, field, absorb):
        self.inserted.append((cursor.pos, field))


class TestUtilsPure(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.utils_ns = _carica_funzioni_utils()
        cls.contab_ns = _carica_funzioni_contab()
        cls.globals_ns = _carica_funzioni_globals()

    def test_days_in_month(self):
        daysInMonth = self.utils_ns["daysInMonth"]
        # Anno bisestile
        self.assertEqual(daysInMonth(datetime.date(2024, 2, 1)), 29)
        # Anno non bisestile
        self.assertEqual(daysInMonth(datetime.date(2023, 2, 1)), 28)
        # Mese da 30 giorni (Aprile, Giugno, Settembre, Novembre)
        self.assertEqual(daysInMonth(datetime.date(2023, 4, 1)), 30)
        self.assertEqual(daysInMonth(datetime.date(2023, 6, 1)), 30)
        # Mese da 31 giorni (Gennaio, Marzo, Maggio, Luglio, Agosto, Ottobre, Dicembre)
        self.assertEqual(daysInMonth(datetime.date(2023, 1, 1)), 31)
        self.assertEqual(daysInMonth(datetime.date(2023, 7, 1)), 31)

    def test_first_week_day(self):
        firstWeekDay = self.utils_ns["firstWeekDay"]
        # 1 Maggio 2023 era Lunedi (0)
        self.assertEqual(firstWeekDay(datetime.date(2023, 5, 15)), 0)
        # 1 Gennaio 2023 era Domenica (6)
        self.assertEqual(firstWeekDay(datetime.date(2023, 1, 10)), 6)

    def test_date2string(self):
        date2String = self.utils_ns["date2String"]
        d = datetime.date(2020, 2, 25)
        self.assertEqual(date2String(d, fmt=0), "25 Febbraio 2020")
        self.assertEqual(date2String(d, fmt=1), "25/02/2020")
        self.assertEqual(date2String(d, fmt=2), "25-02-2020")
        self.assertEqual(date2String(d, fmt=3), "25.02.2020")
        # Default fmt=0
        self.assertEqual(date2String(d), "25 Febbraio 2020")

    def test_string2date(self):
        string2Date = self.utils_ns["string2Date"]
        # Separatore punto
        self.assertEqual(string2Date("25.02.2020"), datetime.date(2020, 2, 25))
        # Separatore slash
        self.assertEqual(string2Date("25/02/2020"), datetime.date(2020, 2, 25))
        # Separatore trattino
        self.assertEqual(string2Date("25-02-2020"), datetime.date(2020, 2, 25))

        # Senza separatori riconosciuti ritorna date.today()
        self.assertEqual(string2Date("nodatehere"), datetime.date.today())

        # Formato non valido con separatore (troppe o poche parti solleva Exception generica)
        with self.assertRaises(Exception):
            string2Date("25.02")

        # Formato con valori non numerici solleva ValueError
        with self.assertRaises(ValueError):
            string2Date("aa.bb.cccc")

    def test_wrap_text(self):
        wrap_text = self.utils_ns["wrap_text"]
        testo = "uno due tre quattro cinque sei sette otto nove dieci"
        risultato = wrap_text(testo, width=20)
        righe = risultato.split("\n")
        for riga in righe:
            self.assertLessEqual(len(riga), 20)
        self.assertEqual(risultato, "uno due tre quattro\ncinque sei sette\notto nove dieci")

    def test_wrap_path(self):
        wrap_path = self.utils_ns["wrap_path"]
        path = "C:\\cartella1\\cartella2\\cartella3\\sottocartella\\file.txt"
        risultato = wrap_path(path, max_len=30)
        self.assertIn("\\\n", risultato)
        parti = risultato.split("\\\n")
        self.assertEqual(len(parti), 3)
        self.assertEqual(parti[0], "C:\\cartella1\\cartella2")
        self.assertEqual(parti[1], "cartella3\\sottocartella")
        self.assertEqual(parti[2], "file.txt")

    def test_count_pdf_pages(self):
        countPdfPages = self.utils_ns["countPdfPages"]
        mock_reader = MagicMock()
        mock_reader.getNumPages.return_value = 5
        self.utils_ns["PyPDF2"].PdfFileReader.return_value = mock_reader

        with patch("builtins.open", MagicMock()):
            pagine = countPdfPages("fittizio.pdf")
            self.assertEqual(pagine, 5)

    def test_replace_pattern_with_field(self):
        replacePatternWithField = self.utils_ns["replacePatternWithField"]
        doc_text = MockText("Documento del [DATA] approvato il [DATA].")
        sostituito = replacePatternWithField(doc_text, "[DATA]", "OGGETTO_CAMPO")
        self.assertTrue(sostituito)
        self.assertEqual(doc_text.String, "Documento del  approvato il .")
        self.assertEqual(len(doc_text.inserted), 2)
        self.assertEqual(doc_text.inserted[0], (14, "OGGETTO_CAMPO"))
        self.assertEqual(doc_text.inserted[1], (28, "OGGETTO_CAMPO"))

        # Nessuna occorrenza trovata
        non_trovato = replacePatternWithField(doc_text, "[INESISTENTE]", "OGGETTO_CAMPO")
        self.assertFalse(non_trovato)

    def test_col_letter(self):
        _col_letter = self.contab_ns["_col_letter"]
        self.assertEqual(_col_letter(0), "A")
        self.assertEqual(_col_letter(25), "Z")
        self.assertEqual(_col_letter(26), "AA")
        self.assertEqual(_col_letter(51), "AZ")
        self.assertEqual(_col_letter(52), "BA")
        self.assertEqual(_col_letter(701), "ZZ")
        self.assertEqual(_col_letter(702), "AAA")

    def test_leeno_globals(self):
        getGlobalVar = self.globals_ns["getGlobalVar"]
        setGlobalVar = self.globals_ns["setGlobalVar"]
        initGlobalVars = self.globals_ns["initGlobalVars"]

        # Inizializzazione
        initGlobalVars({"TEST_VAR_A": 100, "TEST_VAR_B": "valore"})
        self.assertEqual(getGlobalVar("TEST_VAR_A"), 100)
        self.assertEqual(getGlobalVar("TEST_VAR_B"), "valore")
        self.assertIsNone(getGlobalVar("TEST_VAR_NON_PRESENTE"))

        # Modifica tramite setGlobalVar
        setGlobalVar("TEST_VAR_A", 200)
        self.assertEqual(getGlobalVar("TEST_VAR_A"), 200)

        # Sovrascrittura tramite initGlobalVars
        initGlobalVars({"TEST_VAR_A": 300, "TEST_VAR_C": True})
        self.assertEqual(getGlobalVar("TEST_VAR_A"), 300)
        self.assertEqual(getGlobalVar("TEST_VAR_C"), True)


if __name__ == "__main__":
    unittest.main()
