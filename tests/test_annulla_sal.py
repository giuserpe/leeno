import ast
import unittest
from pathlib import Path
from unittest.mock import MagicMock

SRC = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoContab.py"


def _funzione(nome):
    tree = ast.parse(SRC.read_text(encoding="utf-8"))
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name == nome:
            return n
    raise AssertionError("funzione non trovata: " + nome)


def _carica_core(registro, uf_find=None):
    """Compila _annulla_ultimo_sal_core con dipendenze simulate (nessun UNO richiesto)."""
    leeno_utils = MagicMock()
    leeno_utils.log_errore.side_effect = lambda contesto, exc=None: registro.append(contesto)
    ns = {
        "LeenoUtils": leeno_utils,
        "SheetUtils": MagicMock(),
        "PL": MagicMock(),
        "LeenoSheetUtils": MagicMock(),
        "VALUE": 1, "STRING": 4, "FORMULA": 16, "DATETIME": 2, "HARDATTR": 32,
        "ultimo_sal": lambda: ["1"],
        "GeneraCdP": lambda *a, **k: None,
    }
    if uf_find is not None:
        ns["SheetUtils"].uFindString.side_effect = uf_find
    mod = ast.Module(body=[_funzione("_annulla_ultimo_sal_core")], type_ignores=[])
    exec(compile(mod, str(SRC), "exec"), ns)
    return ns["_annulla_ultimo_sal_core"]


class TestAnnullaSalStruttura(unittest.TestCase):
    def test_nessun_except_silenzioso(self):
        f = _funzione("_annulla_ultimo_sal_core")
        for n in ast.walk(f):
            if isinstance(n, ast.ExceptHandler):
                solo_pass = all(isinstance(b, ast.Pass) for b in n.body)
                self.assertFalse(solo_pass, "except silenzioso alla riga %d" % n.lineno)

    def test_ogni_handler_registra_su_file(self):
        f = _funzione("_annulla_ultimo_sal_core")
        for n in ast.walk(f):
            if isinstance(n, ast.ExceptHandler):
                self.assertIn("log_errore", ast.unparse(n), "handler senza log alla riga %d" % n.lineno)

    def test_chiamanti_usano_il_risultato(self):
        for nome in ("MENU_AnnullaAttiContabili", "MENU_AnnullaTuttiAttiContabili"):
            codice = ast.unparse(_funzione(nome))
            self.assertIn("_annulla_ultimo_sal_core", codice)
            self.assertIn("Exclamation", codice)


class TestAnnullaSalComportamento(unittest.TestCase):
    def test_senza_sal_restituisce_lista_vuota(self):
        core = _carica_core([])
        self.assertEqual(core(MagicMock(), listaSal=[]), [])

    def test_fallimento_libretto_viene_segnalato(self):
        registro = []
        core = _carica_core(registro)
        oDoc = MagicMock()
        oDoc.NamedRanges.getByName.side_effect = KeyError("_Lib_1")
        errori = core(oDoc, listaSal=["1"])
        self.assertEqual(errori, ["eliminazione del libretto in CONTABILITA"])
        self.assertIn("annulla SAL: eliminazione libretto", registro)

    def test_errore_pulizia_cdp_solo_registrato(self):
        registro = []

        def boom(*a, **k):
            raise RuntimeError("cella non trovata")

        core = _carica_core(registro, uf_find=boom)
        errori = core(MagicMock(), listaSal=["1", "2"])
        # la pulizia del CdP e la pulizia di S2 usano uFindString: la prima non e' bloccante
        self.assertNotIn("pulizia voci riepilogo CdP", " ".join(errori))
        self.assertIn("annulla SAL: pulizia voci riepilogo CdP", registro)
        self.assertIn("pulizia della situazione contabile in S2", errori)


if __name__ == "__main__":
    unittest.main()
