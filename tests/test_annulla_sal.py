import ast
import unittest
import unittest.mock
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
        "_rimuovi_blocco_cdp": MagicMock(),
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

    def test_non_svuota_celle_cdp_cercando_etichette_dall_alto(self):
        """La vecchia pulizia colpiva il primo blocco CdP e toglieva il nome area prima della rimozione."""
        codice = ast.unparse(_funzione("_annulla_ultimo_sal_core"))
        self.assertNotIn("anchors_cdp", codice)
        self.assertNotIn("removeByName(nome_cdp)", codice)

    def test_sal_maggiore_di_uno_rimuove_il_blocco_cdp_annullato(self):
        core = _carica_core([])
        rimuovi = MagicMock()
        core.__globals__["_rimuovi_blocco_cdp"] = rimuovi
        core(MagicMock(), listaSal=["1", "2", "3"])
        rimuovi.assert_any_call(unittest.mock.ANY, "3")

    def test_ultimo_sal_elimina_il_foglio_cdp_senza_rimuovere_blocchi(self):
        core = _carica_core([])
        rimuovi = MagicMock()
        core.__globals__["_rimuovi_blocco_cdp"] = rimuovi
        oDoc = MagicMock()
        oDoc.Sheets.hasByName.return_value = True
        core(oDoc, listaSal=["1"])
        oDoc.Sheets.removeByName.assert_any_call("CdP")
        # il nome area del foglio eliminato non resta orfano
        oDoc.NamedRanges.removeByName.assert_any_call("_CdP_1")
        rimuovi.assert_not_called()

    def test_rigenerazione_sostituisce_il_blocco_dell_ultimo_cdp_rimasto(self):
        core = _carica_core([])
        chiamate = []
        core.__globals__["ultimo_sal"] = lambda: ["1", "2"]
        core.__globals__["_rimuovi_blocco_cdp"] = lambda d, n: chiamate.append(("rimuovi", str(n)))
        core.__globals__["GeneraCdP"] = lambda d, **k: chiamate.append(("genera", k["nSal"]))
        core(MagicMock(), listaSal=["1", "2", "3"])
        # prima rimuove il blocco n.3 annullato e il vecchio n.2, poi rigenera il n.2
        self.assertEqual(chiamate, [("rimuovi", "3"), ("rimuovi", "2"), ("genera", 2)])

    def test_rigenerazione_del_cdp_1_non_rimuove_blocchi(self):
        core = _carica_core([])
        chiamate = []
        core.__globals__["ultimo_sal"] = lambda: ["1"]
        core.__globals__["_rimuovi_blocco_cdp"] = lambda d, n: chiamate.append(("rimuovi", str(n)))
        core.__globals__["GeneraCdP"] = lambda d, **k: chiamate.append(("genera", k["nSal"]))
        core(MagicMock(), listaSal=["1", "2"])
        self.assertEqual(chiamate, [("rimuovi", "2"), ("genera", 1)])


class TestRimuoviBloccoCdP(unittest.TestCase):
    def _fn(self):
        ns = {}
        mod = ast.Module(body=[_funzione("_rimuovi_blocco_cdp")], type_ignores=[])
        exec(compile(mod, str(SRC), "exec"), ns)
        return ns["_rimuovi_blocco_cdp"]

    def test_rimuove_intestazione_righe_e_nome_area(self):
        oDoc = MagicMock()
        oDoc.Sheets.hasByName.return_value = True
        oDoc.NamedRanges.hasByName.return_value = True
        addr = oDoc.NamedRanges.getByName.return_value.ReferredCells.RangeAddress
        addr.StartRow, addr.EndRow = 95, 139
        oSheet = oDoc.Sheets.getByName.return_value

        self.assertTrue(self._fn()(oDoc, 3))

        oDoc.NamedRanges.getByName.assert_called_with("_CdP_3")
        # intestazione a StartRow - 1: righe 94..139 = 46 righe
        oSheet.Rows.removeByIndex.assert_called_once_with(94, 46)
        oDoc.NamedRanges.removeByName.assert_called_once_with("_CdP_3")

    def test_senza_nome_area_non_tocca_il_foglio(self):
        oDoc = MagicMock()
        oDoc.Sheets.hasByName.return_value = True
        oDoc.NamedRanges.hasByName.return_value = False

        self.assertFalse(self._fn()(oDoc, 3))
        oDoc.Sheets.getByName.return_value.Rows.removeByIndex.assert_not_called()


if __name__ == "__main__":
    unittest.main()
