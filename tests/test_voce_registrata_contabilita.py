import ast
import unittest
from pathlib import Path
from unittest.mock import MagicMock

SRC_CONTAB = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoContab.py"
SRC_PYLEENO = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/pyleeno.py"
SRC_SHEETUTILS = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoSheetUtils.py"


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

    def test_seleziona_voce_non_bloccata_da_voce_registrata(self):
        """Verifica che seleziona_voce() restituisca inizio/fine voce anche se registrata (#reg) senza bloccare."""
        ns_pyleeno = {
            "Dialogs": self.mock_dialogs,
            "LeenoComputo": self.mock_computo,
            "LeenoUtils": self.mock_utils,
            "LeggiPosizioneCorrente": lambda: (0, 10),
        }
        code_seleziona_voce = _estrai_funzione(SRC_PYLEENO, "seleziona_voce")
        exec(code_seleziona_voce, ns_pyleeno)
        seleziona_voce = ns_pyleeno["seleziona_voce"]

        oSheet = MagicMock()
        oSheet.Name = "CONTABILITA"
        self.mock_utils.getDocument.return_value.CurrentController.ActiveSheet = oSheet

        sStRange = MagicMock()
        sStRange.RangeAddress.StartRow = 10
        sStRange.RangeAddress.EndRow = 14
        self.mock_computo.circoscriveVoceComputo.return_value = sStRange

        res = seleziona_voce(row=10)
        self.assertEqual(res, (10, 14))
        self.mock_dialogs.Exclamation.assert_not_called()

    def test_riepilogo_quantita_non_bloccata_da_voci_registrate(self):
        """Verifica che riepilogo_quantita operi su CONTABILITA anche in presenza di voci registrate."""
        ns_sheetutils = {
            "PL": self.mock_pl,
            "LeenoUtils": self.mock_utils,
            "with_undo": lambda fn: fn,
            "prossimaVoce": MagicMock(),
            "cercaUltimaVoce": MagicMock(),
            "cancella_riepilogo_quantita": MagicMock(),
        }
        # Mock release_ram e no_refresh
        self.mock_utils.release_ram = lambda fn: fn
        self.mock_utils.no_refresh = lambda fn: fn

        code_riepilogo = _estrai_funzione(SRC_SHEETUTILS, "riepilogo_quantita")
        exec(code_riepilogo, ns_sheetutils)
        riepilogo_quantita = ns_sheetutils["riepilogo_quantita"]

        self.mock_pl.GetModifiers.return_value = (False, False)
        oSheet = MagicMock()
        oSheet.Name = "CONTABILITA"
        self.mock_utils.getDocument.return_value.CurrentController.ActiveSheet = oSheet

        # Simula una voce che ha seleziona_voce = (10, 14) e last = 20
        ns_sheetutils["cercaUltimaVoce"].return_value = 20
        ns_sheetutils["prossimaVoce"].side_effect = [10, 25]  # prima voce 10, poi esce dal loop (> 20)
        self.mock_pl.seleziona_voce.return_value = (10, 14)

        cell_mock = MagicMock()
        oSheet.getCellByPosition.return_value = cell_mock

        riepilogo_quantita()

        # Verifica che siano state scritte le intestazioni su AM3:AO3
        oSheet.getCellRangeByName.assert_any_call("AM3")
        oSheet.getCellRangeByName.assert_any_call("AN3")
        oSheet.getCellRangeByName.assert_any_call("AO3")
        # Nessun dialogo di allerta deve essere stato invocato
        self.mock_dialogs.Exclamation.assert_not_called()
        # Le celle di colonna AM (38), AN (39), AO (40) per ER=14 sono state valorizzate
        oSheet.getCellByPosition.assert_any_call(38, 14)
        oSheet.getCellByPosition.assert_any_call(39, 14)
        oSheet.getCellByPosition.assert_any_call(40, 14)


if __name__ == "__main__":
    unittest.main()
