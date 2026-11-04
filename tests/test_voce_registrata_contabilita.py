import ast
import unittest
from pathlib import Path
from unittest.mock import MagicMock

SRC_CONTAB = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoContab.py"
SRC_PYLEENO = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/pyleeno.py"


def _estrai_funzione(file_path, nome):
    tree = ast.parse(file_path.read_text(encoding="utf-8"))
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name == nome:
            mod = ast.Module(body=[n], type_ignores=[])
            return compile(mod, filename=str(file_path), mode="exec")
    raise AssertionError(f"Funzione {nome} non trovata in {file_path}")


class TestVoceRegistrataContabilita(unittest.TestCase):

    def setUp(self):
        self.mock_dialogs = MagicMock()
        self.mock_sheet_utils = MagicMock()
        self.mock_computo = MagicMock()
        self.mock_globals = MagicMock()
        self.mock_utils = MagicMock()
        self.mock_utils.no_refresh = lambda fn: fn
        self.mock_pl = MagicMock()

        self.ns_contab = {
            "Dialogs": self.mock_dialogs,
            "LeenoSheetUtils": self.mock_sheet_utils,
            "LeenoComputo": self.mock_computo,
            "LeenoGlobals": self.mock_globals,
            "LeenoUtils": self.mock_utils,
            "PL": self.mock_pl,
        }

        # Carica is_voce_registrata, sbloccaContabilita, partita
        code_is_reg = _estrai_funzione(SRC_CONTAB, "is_voce_registrata")
        exec(code_is_reg, self.ns_contab)
        self.is_voce_registrata = self.ns_contab["is_voce_registrata"]

        code_sblocca = _estrai_funzione(SRC_CONTAB, "sbloccaContabilita")
        exec(code_sblocca, self.ns_contab)
        self.sbloccaContabilita = self.ns_contab["sbloccaContabilita"]

        code_partita = _estrai_funzione(SRC_CONTAB, "partita")
        exec(code_partita, self.ns_contab)
        self.partita = self.ns_contab["partita"]

    def test_is_voce_registrata_flag_partenza(self):
        oSheet = MagicMock()
        oSheet.Name = "CONTABILITA"
        self.mock_sheet_utils.cercaPartenza.return_value = ("CONTABILITA", 10, "#reg")

        self.assertTrue(self.is_voce_registrata(oSheet, 10))

    def test_is_voce_registrata_scansione_colonna_w(self):
        oSheet = MagicMock()
        oSheet.Name = "CONTABILITA"
        self.mock_sheet_utils.cercaPartenza.return_value = ("CONTABILITA", 10, "")

        sStRange = MagicMock()
        sStRange.RangeAddress.StartRow = 5
        sStRange.RangeAddress.EndRow = 8
        self.mock_computo.circoscriveVoceComputo.return_value = sStRange

        def get_cell(col, row):
            cell = MagicMock()
            cell.String = "#reg" if (col == 22 and row == 6) else ""
            return cell

        oSheet.getCellByPosition.side_effect = get_cell
        self.assertTrue(self.is_voce_registrata(oSheet, 7))

    def test_is_voce_registrata_non_registrata(self):
        oSheet = MagicMock()
        oSheet.Name = "CONTABILITA"
        self.mock_sheet_utils.cercaPartenza.return_value = ("CONTABILITA", 10, "")
        sStRange = MagicMock()
        sStRange.RangeAddress.StartRow = 5
        sStRange.RangeAddress.EndRow = 8
        self.mock_computo.circoscriveVoceComputo.return_value = sStRange

        def get_cell_empty(col, row):
            cell = MagicMock()
            cell.String = ""
            return cell

        oSheet.getCellByPosition.side_effect = get_cell_empty
        self.assertFalse(self.is_voce_registrata(oSheet, 7))

    def test_is_voce_registrata_foglio_non_contabilita(self):
        oSheet = MagicMock()
        oSheet.Name = "COMPUTO"
        self.assertFalse(self.is_voce_registrata(oSheet, 10))

    def test_sblocca_contabilita_blocca_con_alert(self):
        oSheet = MagicMock()
        oSheet.Name = "CONTABILITA"
        self.mock_sheet_utils.cercaPartenza.return_value = ("CONTABILITA", 10, "#reg")

        res = self.sbloccaContabilita(oSheet, 10)
        self.assertFalse(res)
        self.mock_dialogs.Exclamation.assert_called_once()
        self.assertIn("#reg", self.mock_dialogs.Exclamation.call_args[1]["Text"])

    def test_sblocca_contabilita_consente_se_non_registrata(self):
        oSheet = MagicMock()
        oSheet.Name = "CONTABILITA"
        self.mock_sheet_utils.cercaPartenza.return_value = ("CONTABILITA", 10, "")
        self.mock_computo.circoscriveVoceComputo.return_value = None

        res = self.sbloccaContabilita(oSheet, 10)
        self.assertTrue(res)
        self.mock_dialogs.Exclamation.assert_not_called()

    def test_partita_blocca_se_voce_registrata(self):
        oSheet = MagicMock()
        oSheet.Name = "CONTABILITA"
        self.mock_utils.getDocument.return_value.CurrentController.ActiveSheet = oSheet
        self.mock_pl.LeggiPosizioneCorrente.return_value = (0, 10)
        self.mock_sheet_utils.cercaPartenza.return_value = ("CONTABILITA", 10, "#reg")

        self.partita("PARTITA PROVVISORIA")
        self.mock_dialogs.Exclamation.assert_called_once()
        self.mock_pl.Copia_riga_Ent.assert_not_called()

    def test_sblocca_cont_pyleeno_blocca_con_alert(self):
        ns_pyleeno = {
            "Dialogs": self.mock_dialogs,
            "LeenoGlobals": self.mock_globals,
            "LeenoUtils": self.mock_utils,
            "cerca_partenza": lambda: ("CONTABILITA", 10, "#reg"),
        }
        code_sblocca_cont = _estrai_funzione(SRC_PYLEENO, "sblocca_cont")
        exec(code_sblocca_cont, ns_pyleeno)
        sblocca_cont = ns_pyleeno["sblocca_cont"]

        oSheet = MagicMock()
        oSheet.Name = "CONTABILITA"
        self.mock_utils.getDocument.return_value.CurrentController.ActiveSheet = oSheet

        sblocca_cont()
        self.mock_dialogs.Exclamation.assert_called_once()
        self.mock_globals.setGlobalVar.assert_called_with("sblocca_computo", 0)


if __name__ == "__main__":
    unittest.main()
