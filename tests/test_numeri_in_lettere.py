import ast
import unittest
from pathlib import Path

UTILS_PATH = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoUtils.py"
CONTAB_PATH = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoContab.py"


def _carica_funzioni_utils():
    """Estrae int_to_italian e convert_number_string da LeenoUtils senza importare il modulo."""
    tree = ast.parse(UTILS_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name in ("int_to_italian", "under_thousand", "convert_number_string"):
            nodi.append(n)
    ns = {}
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(UTILS_PATH), "exec"), ns)
    return ns


def _carica_funzioni_contab():
    """Estrae _intero_in_lettere e numero_in_lettere_euro da LeenoContab senza importare il modulo."""
    tree = ast.parse(CONTAB_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name in ("_intero_in_lettere", "numero_in_lettere_euro"):
            nodi.append(n)
    ns = {}
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(CONTAB_PATH), "exec"), ns)
    return ns


class TestNumeriInLettere(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.utils_ns = _carica_funzioni_utils()
        cls.contab_ns = _carica_funzioni_contab()

    def test_int_to_italian_valori_base(self):
        int_to_italian = self.utils_ns["int_to_italian"]
        self.assertEqual(int_to_italian(0), "zero")
        self.assertEqual(int_to_italian(1), "uno")
        self.assertEqual(int_to_italian(11), "undici")
        self.assertEqual(int_to_italian(16), "sedici")
        self.assertEqual(int_to_italian(20), "venti")
        self.assertEqual(int_to_italian(21), "ventuno")
        self.assertEqual(int_to_italian(28), "ventotto")
        self.assertEqual(int_to_italian(100), "cento")
        self.assertEqual(int_to_italian(101), "centouno")
        self.assertEqual(int_to_italian(180), "centottanta")
        self.assertEqual(int_to_italian(999), "novecentonovantanove")
        self.assertEqual(int_to_italian(1000), "mille")
        self.assertEqual(int_to_italian(1001), "milleuno")
        self.assertEqual(int_to_italian(2000), "duemila")
        self.assertEqual(int_to_italian(10000), "diecimila")
        self.assertEqual(int_to_italian(1000000), "unmilione")
        self.assertEqual(int_to_italian(2000000), "duemilioni")
        self.assertEqual(int_to_italian(1000000000), "unmiliardo")
        self.assertEqual(
            int_to_italian(1234567890),
            "unmiliardoduecentotrentaquattromilionicinquecentosessantasettemilaottocentonovanta",
        )

    def test_convert_number_string(self):
        convert_number_string = self.utils_ns["convert_number_string"]
        self.assertEqual(convert_number_string("123"), "centoventitre")
        self.assertEqual(convert_number_string("123.45"), "centoventitre/45")
        self.assertEqual(convert_number_string("123,45"), "centoventitre/45")
        self.assertEqual(convert_number_string("-123"), "menocentoventitre")
        self.assertEqual(convert_number_string("-123.45"), "menocentoventitre/45")
        self.assertEqual(convert_number_string("0"), "zero")
        self.assertEqual(convert_number_string("0.0"), "zero/0")
        with self.assertRaises(ValueError):
            convert_number_string("")
        with self.assertRaises(ValueError):
            convert_number_string("   ")
        with self.assertRaises(ValueError):
            convert_number_string("abc")

    def test_intero_in_lettere(self):
        _intero_in_lettere = self.contab_ns["_intero_in_lettere"]
        self.assertEqual(_intero_in_lettere(0), "zero")
        self.assertEqual(_intero_in_lettere(1), "uno")
        self.assertEqual(_intero_in_lettere(11), "undici")
        self.assertEqual(_intero_in_lettere(16), "sedici")
        self.assertEqual(_intero_in_lettere(20), "venti")
        self.assertEqual(_intero_in_lettere(21), "ventuno")
        self.assertEqual(_intero_in_lettere(28), "ventotto")
        self.assertEqual(_intero_in_lettere(100), "cento")
        self.assertEqual(_intero_in_lettere(101), "centouno")

        # BUG NOTO: divergenza su 180 tra LeenoUtils.int_to_italian e LeenoContab._intero_in_lettere.
        # LeenoUtils.int_to_italian(180) restituisce 'centottanta' con elisione della vocale 'o',
        # mentre LeenoContab._intero_in_lettere(180) restituisce 'centoottanta' senza elisione.
        self.assertEqual(_intero_in_lettere(180), "centoottanta")
        self.assertNotEqual(self.utils_ns["int_to_italian"](180), _intero_in_lettere(180))

        self.assertEqual(_intero_in_lettere(999), "novecentonovantanove")
        self.assertEqual(_intero_in_lettere(1000), "mille")
        self.assertEqual(_intero_in_lettere(1001), "milleuno")
        self.assertEqual(_intero_in_lettere(2000), "duemila")
        self.assertEqual(_intero_in_lettere(10000), "diecimila")
        self.assertEqual(_intero_in_lettere(1000000), "unmilione")
        self.assertEqual(_intero_in_lettere(2000000), "duemilioni")
        self.assertEqual(_intero_in_lettere(1000000000), "unmiliardo")
        self.assertEqual(
            _intero_in_lettere(1234567890),
            "unmiliardoduecentotrentaquattromilionicinquecentosessantasettemilaottocentonovanta",
        )

    def test_numero_in_lettere_euro(self):
        numero_in_lettere_euro = self.contab_ns["numero_in_lettere_euro"]
        # Esempio dalla docstring di numero_in_lettere_euro:
        self.assertEqual(numero_in_lettere_euro(1234.56), "milleduecentotrentaquattro/56")
        # 0 centesimi
        self.assertEqual(numero_in_lettere_euro(100.00), "cento/00")
        # 5 centesimi
        self.assertEqual(numero_in_lettere_euro(100.05), "cento/05")
        # Arrotondamento per difetto
        self.assertEqual(numero_in_lettere_euro(100.004), "cento/00")
        # Arrotondamento per eccesso
        self.assertEqual(numero_in_lettere_euro(100.006), "cento/01")
        # Valore negativo
        self.assertEqual(numero_in_lettere_euro(-1234.56), "meno milleduecentotrentaquattro/56")
        # Zero euro
        self.assertEqual(numero_in_lettere_euro(0), "zero/00")


if __name__ == "__main__":
    unittest.main()
