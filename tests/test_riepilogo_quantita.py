import ast
import unittest
from pathlib import Path
from unittest.mock import MagicMock

SRC_SHEETUTILS = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoSheetUtils.py"


def _estrai_funzione(file_path, nome):
    tree = ast.parse(file_path.read_text(encoding="utf-8"))
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name == nome:
            mod = ast.Module(body=[n], type_ignores=[])
            return compile(mod, filename=str(file_path), mode="exec")
    raise AssertionError(f"Funzione {nome} non trovata in {file_path}")


class TestRiepilogoQuantita(unittest.TestCase):

    def setUp(self):
        self.mock_uno = MagicMock()
        self.mock_pl = MagicMock()
        self.mock_utils = MagicMock()
        self.mock_sheet_utils = MagicMock()
        self.mock_utils.no_refresh = lambda fn: fn
        self.mock_utils.release_ram = lambda fn: fn
        self.mock_pl.GetModifiers.return_value = (False, False)

        def mock_create_uno_struct(struct_name):
            obj = MagicMock()
            obj._struct_name = struct_name
            return obj

        self.mock_uno.createUnoStruct.side_effect = mock_create_uno_struct

        self.ns = {
            "uno": self.mock_uno,
            "PL": self.mock_pl,
            "LeenoUtils": self.mock_utils,
            "SheetUtils": self.mock_sheet_utils,
            "with_undo": lambda fn: fn,
            "prossimaVoce": MagicMock(),
            "cercaUltimaVoce": MagicMock(),
            "cancella_riepilogo_quantita": MagicMock(),
            "_riepilogo_quantita_attivo": MagicMock(return_value=False),
            "VALUE": 1,
            "DATETIME": 2,
            "STRING": 4,
            "FORMULA": 16,
        }
        code = _estrai_funzione(SRC_SHEETUTILS, "riepilogo_quantita")
        exec(code, self.ns)
        self.riepilogo_quantita = self.ns["riepilogo_quantita"]

    def test_contabilita_area_vuota_non_inserisce_celle(self):
        oSheet = MagicMock()
        oSheet.Name = "CONTABILITA"
        self.mock_utils.getDocument.return_value.CurrentController.ActiveSheet = oSheet
        self.mock_sheet_utils.uFindStringCol.return_value = 50

        # Simula range queryContentCells vuoto (count = 0)
        area_mock = MagicMock()
        content_cells_mock = MagicMock()
        content_cells_mock.getCount.return_value = 0
        area_mock.queryContentCells.return_value = content_cells_mock
        oSheet.getCellRangeByPosition.return_value = area_mock

        # Voci nel foglio
        self.ns["cercaUltimaVoce"].return_value = 50
        self.ns["prossimaVoce"].side_effect = [10, 60]
        self.mock_pl.seleziona_voce.return_value = (10, 14)

        self.riepilogo_quantita()

        # Verifica che non sia stato chiamato insertCells
        oSheet.insertCells.assert_not_called()
        self.mock_sheet_utils.uFindStringCol.assert_called_with('T O T A L E', 2, oSheet, start=2, equal=1, up=True)

    def test_contabilita_area_con_dati_sposta_celle_a_destra(self):
        oSheet = MagicMock()
        oSheet.Name = "CONTABILITA"
        oSheet.RangeAddress.Sheet = 1
        self.mock_utils.getDocument.return_value.CurrentController.ActiveSheet = oSheet
        self.mock_sheet_utils.uFindStringCol.return_value = 45

        # Simula range con queryContentCells con count > 0
        area_mock = MagicMock()
        content_cells_mock = MagicMock()
        content_cells_mock.getCount.return_value = 3
        area_mock.queryContentCells.return_value = content_cells_mock
        oSheet.getCellRangeByPosition.return_value = area_mock

        # Voci nel foglio
        self.ns["cercaUltimaVoce"].return_value = 45
        self.ns["prossimaVoce"].side_effect = [10, 60]
        self.mock_pl.seleziona_voce.return_value = (10, 14)

        self.riepilogo_quantita()

        # Verifica che sia stato chiamato insertCells con le coordinate 38-40 e riga_totale=45
        oSheet.insertCells.assert_called_once()
        addr, mode = oSheet.insertCells.call_args[0]
        self.assertEqual(addr.Sheet, 1)
        self.assertEqual(addr.StartColumn, 38)
        self.assertEqual(addr.EndColumn, 40)
        self.assertEqual(addr.StartRow, 2)
        self.assertEqual(addr.EndRow, 45)

    def test_computo_area_con_dati_fallback_dataarray(self):
        oSheet = MagicMock()
        oSheet.Name = "COMPUTO"
        oSheet.RangeAddress.Sheet = 0
        self.mock_utils.getDocument.return_value.CurrentController.ActiveSheet = oSheet
        self.mock_sheet_utils.uFindStringCol.return_value = None  # test fallback su cercaUltimaVoce
        self.ns["cercaUltimaVoce"].return_value = 30

        # Simula queryContentCells che solleva eccezione e fallback su getDataArray
        area_mock = MagicMock()
        area_mock.queryContentCells.side_effect = Exception("Not supported")
        area_mock.getDataArray.return_value = (("", "", ""), ("", "DATO", ""), ("", "", ""))
        oSheet.getCellRangeByPosition.return_value = area_mock

        self.ns["prossimaVoce"].side_effect = [5, 40]
        self.mock_pl.seleziona_voce.return_value = (5, 8)

        self.riepilogo_quantita()

        oSheet.insertCells.assert_called_once()
        addr, mode = oSheet.insertCells.call_args[0]
        self.assertEqual(addr.Sheet, 0)
        self.assertEqual(addr.StartColumn, 44)
        self.assertEqual(addr.EndColumn, 46)
        self.assertEqual(addr.StartRow, 2)
        self.assertEqual(addr.EndRow, 31)  # cercaUltimaVoce + 1 = 31

    def test_se_riepilogo_gia_attivo_non_sposta_celle(self):
        oSheet = MagicMock()
        oSheet.Name = "CONTABILITA"
        self.mock_utils.getDocument.return_value.CurrentController.ActiveSheet = oSheet
        self.ns["_riepilogo_quantita_attivo"].return_value = True

        self.ns["cercaUltimaVoce"].return_value = 20
        self.ns["prossimaVoce"].side_effect = [5, 30]
        self.mock_pl.seleziona_voce.return_value = (5, 8)

        self.riepilogo_quantita()

        oSheet.insertCells.assert_not_called()


if __name__ == "__main__":
    unittest.main()
