# -*- coding: utf-8 -*-
"""
Test di SheetUtils.leggi_dato_etichettato (lettura "etichetta in B, valore in C" del foglio S2).

- unitari, senza UNO: equivalenza con il comportamento storico di
  LeenoContab._leggi_dato_anagrafico su griglie casuali, semantica dell'occorrenza,
  ricerca in colonna A, gestione degli errori;
- delega: _leggi_dato_anagrafico usa il nuovo lettore;
- documento reale (opzionale, LEENO_TEST_UNO=1): confronto storico/nuovo su tutte le
  etichette del foglio S2 di tests/data/LeenO_Esempio.ods e valori attesi noti.
"""
import ast
import random
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

try:
    from tests import uno_harness
    from tests.estrai_sorgente import FoglioFinto, carica
except ImportError:  # esecuzione con la cartella tests nel percorso di ricerca
    import uno_harness
    from estrai_sorgente import FoglioFinto, carica

ROOT = Path(__file__).resolve().parent.parent
PYTHONPATH_DIR = ROOT / "src/Ultimus.oxt/python/pythonpath"
ESEMPIO = ROOT / "tests/data/LeenO_Esempio.ods"


def _leggi_dato_storico(oSheet, label):
    """Comportamento di _leggi_dato_anagrafico prima dello spostamento in SheetUtils."""
    try:
        label_up = label.upper()
        oCell = oSheet.getCellByPosition(0, 0)
        oCursor = oSheet.createCursorByRange(oCell)
        oCursor.gotoEndOfUsedArea(True)
        max_row = oCursor.RangeAddress.EndRow

        for col in (1, 0):
            for r in range(max_row + 1):
                cell_txt = oSheet.getCellByPosition(col, r).String.upper()
                if label_up in cell_txt:
                    val = oSheet.getCellByPosition(2, r).String
                    if val:
                        return val.strip()
    except Exception:
        pass
    return ""


def _carica(log_errore=None):
    """Compila leggi_dato_etichettato da SheetUtils senza importare il modulo (richiede uno)."""
    ns = {"LeenoUtils": SimpleNamespace(log_errore=log_errore or MagicMock())}
    carica(PYTHONPATH_DIR / "SheetUtils.py", ("leggi_dato_etichettato",), ns=ns)
    return ns["leggi_dato_etichettato"]


class TestLetturaEtichette(unittest.TestCase):
    def setUp(self):
        self.leggi = _carica()

    def test_equivalente_al_comportamento_storico_su_griglie_casuali(self):
        rng = random.Random(20261123)
        parole = ["OGGETTO:", "Località:", "Importo contrattuale:", "Direttore Lavori:",
                  "Appaltatore:", "Committente", "", "altro", "garanzia:"]
        valori = ["", "", "x", " Comune di Matera ", "€ 1.000,00", "  "]
        for _ in range(300):
            celle = {}
            for r in range(rng.randint(1, 25)):
                for col in (0, 1):
                    if rng.random() < 0.6:
                        celle[(col, r)] = rng.choice(parole)
                if rng.random() < 0.7:
                    celle[(2, r)] = rng.choice(valori)
            foglio = FoglioFinto(celle)
            for label in parole + ["MATERA", "dir", "inesistente"]:
                with self.subTest(label=label):
                    self.assertEqual(self.leggi(foglio, label), _leggi_dato_storico(foglio, label))

    def test_ignora_maiuscole_e_toglie_gli_spazi(self):
        foglio = FoglioFinto({(1, 0): "Località:", (2, 0): "  Comune di Matera  "})
        self.assertEqual(self.leggi(foglio, "LOCALITÀ"), "Comune di Matera")

    def test_cerca_in_colonna_a_se_manca_in_b(self):
        foglio = FoglioFinto({(0, 3): "Appaltatore", (2, 3): "Ditta Costruzioni"})
        self.assertEqual(self.leggi(foglio, "appaltatore"), "Ditta Costruzioni")

    def test_senza_occorrenza_salta_i_valori_vuoti(self):
        foglio = FoglioFinto({(1, 0): "Importo contrattuale:", (2, 0): "",
                              (1, 5): "Importo contrattuale:", (2, 5): "€ 7"})
        self.assertEqual(self.leggi(foglio, "Importo contrattuale"), "€ 7")

    def test_occorrenza_seleziona_la_riga_anche_se_vuota(self):
        foglio = FoglioFinto({
            (1, 1): "Importo contrattuale:", (2, 1): "€ 1",
            (1, 5): "Importo contrattuale:", (2, 5): "",
            (1, 9): "Importo contrattuale:", (2, 9): "€ 3",
        })
        letti = [self.leggi(foglio, "Importo contrattuale", n) for n in (1, 2, 3, 4)]
        self.assertEqual(letti, ["€ 1", "", "€ 3", ""])

    def test_etichetta_assente_o_foglio_vuoto(self):
        self.assertEqual(self.leggi(FoglioFinto({}), "OGGETTO"), "")
        self.assertEqual(self.leggi(FoglioFinto({(1, 0): "a", (2, 0): "b"}), "OGGETTO"), "")

    def test_errore_di_lettura_ritorna_vuoto_e_viene_registrato(self):
        log = MagicMock()
        leggi = _carica(log)
        guasto = MagicMock()
        guasto.getCellByPosition.side_effect = RuntimeError("cella non leggibile")
        self.assertEqual(leggi(guasto, "OGGETTO"), "")
        log.assert_called_once()
        self.assertIn("OGGETTO", log.call_args[0][0])


class TestDelega(unittest.TestCase):
    def test_leggi_dato_anagrafico_usa_il_lettore_di_sheetutils(self):
        albero = ast.parse((PYTHONPATH_DIR / "LeenoContab.py").read_text(encoding="utf-8"))
        nodo = [n for n in albero.body if isinstance(n, ast.FunctionDef) and n.name == "_leggi_dato_anagrafico"]
        self.assertEqual(len(nodo), 1)
        self.assertIn("SheetUtils.leggi_dato_etichettato(", ast.unparse(nodo[0]))
        # nessuna seconda implementazione del ciclo di ricerca
        self.assertNotIn("gotoEndOfUsedArea", ast.unparse(nodo[0]))


@unittest.skipUnless(uno_harness.uno_disponibile(ESEMPIO),
                     "richiede LEENO_TEST_UNO=1, uno, soffice e LeenO_Esempio.ods")
class TestFoglioS2Reale(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.sessione = uno_harness.SessioneUno()
        cls.addClassCleanup(cls.sessione.chiudi)
        cls.sessione.avvia()
        cls.doc = cls.sessione.apri_copia(ESEMPIO, "s2")
        cls.s2 = cls.doc.Sheets.getByName("S2")
        cls.leggi = staticmethod(_carica())

    def test_tutte_le_etichette_di_s2_come_prima(self):
        etichette = set()
        for r in range(0, 120):
            for col in (0, 1):
                testo = self.s2.getCellByPosition(col, r).String.strip()
                if testo:
                    etichette.add(testo)
        self.assertGreater(len(etichette), 20)
        for label in sorted(etichette):
            with self.subTest(label=label):
                self.assertEqual(self.leggi(self.s2, label), _leggi_dato_storico(self.s2, label))

    def test_etichette_usate_dal_codice(self):
        for label in ("OGGETTO", "Committente", "Stazione Appaltante", "Località", "Appaltatore",
                      "Direttore Lavori", "Direttore dei Lavori", "garanzia:", "Impresa", "Ente"):
            with self.subTest(label=label):
                self.assertEqual(self.leggi(self.s2, label), _leggi_dato_storico(self.s2, label))

    def test_valori_noti_dell_esempio(self):
        self.assertEqual(self.leggi(self.s2, "Località"), "Comune di Matera")
        self.assertEqual(self.leggi(self.s2, "Appaltatore"), "Ditta Costruzioni")

    def test_importo_contrattuale_per_occorrenza(self):
        primo = self.leggi(self.s2, "Importo contrattuale", 1)
        secondo = self.leggi(self.s2, "Importo contrattuale", 2)
        self.assertTrue(primo.startswith("€ 695.837"), primo)
        self.assertEqual(secondo, "")
        self.assertEqual(self.leggi(self.s2, "Importo contrattuale", 4), "")


if __name__ == "__main__":
    unittest.main()
