import ast
import unittest
from pathlib import Path
from unittest.mock import MagicMock

SRC_SHEET_UTILS = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoSheetUtils.py"
SRC_PYLEENO = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/pyleeno.py"
SRC_CONTAB = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoContab.py"


def _estrai_funzione(file_path, nome):
    tree = ast.parse(file_path.read_text(encoding="utf-8"))
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name == nome:
            mod = ast.Module(body=[n], type_ignores=[])
            return compile(mod, filename=str(file_path), mode="exec")
    raise AssertionError(f"Funzione {nome} non trovata in {file_path}")


class MockCellRangeAddress:
    def __init__(self, sheet=0, start_col=0, start_row=0, end_col=0, end_row=0):
        self.Sheet = sheet
        self.StartColumn = start_col
        self.StartRow = start_row
        self.EndColumn = end_col
        self.EndRow = end_row


class MockColumns:
    def __init__(self):
        self.IsVisible = True
        self.Width = 1000
        self.OptimalWidth = False


class MockCellRange:
    def __init__(self, sheet, start_col, start_row, end_col, end_row):
        self.Sheet = sheet
        self.StartColumn = start_col
        self.StartRow = start_row
        self.EndColumn = end_col
        self.EndRow = end_row
        self.Columns = MockColumns()


class MockSheet:
    def __init__(self, name="CONTABILITA", sheet_idx=2):
        self.Name = name
        self.RangeAddress = MockCellRangeAddress(sheet=sheet_idx)
        self.grouped_calls = []
        self._ranges = {}
        self._cols = {}
        self.Columns = {}
        for col_name in ['A', 'B', 'C', 'F', 'G', 'H', 'I', 'J', 'L', 'N', 'P', 'T', 'U', 'W', 'X', 'Z', 'AC', 'AD', 'AE', 'AX', 'AY']:
            self.Columns[col_name] = MockColumns()

    def group(self, addr, orientation):
        self.grouped_calls.append((addr.StartColumn, addr.EndColumn, addr.StartRow, addr.EndRow, orientation))

    def ungroup(self, addr, orientation):
        # rimuove i gruppi interamente contenuti nell'intervallo, come LibreOffice
        self.grouped_calls = [
            g for g in self.grouped_calls
            if not (addr.StartColumn <= g[0] and g[1] <= addr.EndColumn)
        ]

    def getColumns(self):
        return self

    def getByIndex(self, idx):
        if idx not in self._cols:
            self._cols[idx] = MockColumns()
        return self._cols[idx]

    def getCellRangeByPosition(self, start_col, start_row, end_col, end_row):
        key = (start_col, start_row, end_col, end_row)
        if key not in self._ranges:
            self._ranges[key] = MockCellRange(self, start_col, start_row, end_col, end_row)
        return self._ranges[key]

    def clearOutline(self):
        self.grouped_calls.clear()


class TestRaggruppaColonneContabilita(unittest.TestCase):

    def setUp(self):
        self.mock_uno = MagicMock()
        def side_effect_createUnoStruct(type_name):
            if type_name == 'com.sun.star.table.CellRangeAddress':
                return MockCellRangeAddress()
            return MagicMock()
        self.mock_uno.createUnoStruct.side_effect = side_effect_createUnoStruct

        self.mock_utils = MagicMock()

        # Configurazione namespace per LeenoSheetUtils
        self.ns_sheet_utils = {
            "uno": self.mock_uno,
            "LeenoUtils": self.mock_utils,
            "SheetUtils": MagicMock(),
        }
        for nome in ("raggruppa_colonne", "raggruppa_colonne_contabilita"):
            exec(_estrai_funzione(SRC_SHEET_UTILS, nome), self.ns_sheet_utils)
        self.raggruppa_colonne_contabilita = self.ns_sheet_utils["raggruppa_colonne_contabilita"]

        # Configurazione namespace per pyleeno
        self.ns_pyleeno = {
            "uno": self.mock_uno,
            "LeenoUtils": self.mock_utils,
            "LeenoSheetUtils": MagicMock(),
        }
        self.ns_pyleeno["LeenoSheetUtils"].raggruppa_colonne_contabilita = self.raggruppa_colonne_contabilita
        code_pyleeno_fn = _estrai_funzione(SRC_PYLEENO, "raggruppa_colonne_contabilita")
        exec(code_pyleeno_fn, self.ns_pyleeno)
        self.pyleeno_raggruppa = self.ns_pyleeno["raggruppa_colonne_contabilita"]

    def test_raggruppa_colonne_contabilita_funzione_base(self):
        sheet_contab = MockSheet("CONTABILITA", 2)
        self.raggruppa_colonne_contabilita(sheet_contab)

        # Un solo gruppo per intervallo, orientamento 0 (colonne); Y:AA restano fuori
        self.assertEqual(sorted(sheet_contab.grouped_calls), [
            (19, 23, 0, 0, 0), (27, 28, 0, 0, 0), (29, 30, 0, 0, 0), (31, 37, 0, 0, 0)])

        # Gruppi chiusi
        for start, end in ((19, 23), (27, 28), (29, 30), (31, 37)):
            self.assertFalse(sheet_contab.getCellRangeByPosition(start, 0, end, 0).Columns.IsVisible)

        # K, M, Y, Z, AA sempre a larghezza 0
        for col in (10, 12, 24, 25, 26):
            self.assertFalse(sheet_contab.getByIndex(col).IsVisible)

    def test_raggruppa_colonne_contabilita_idempotente(self):
        sheet_contab = MockSheet("CONTABILITA", 2)
        # vecchio gruppo unico T:AC, anche duplicato
        for _ in range(2):
            sheet_contab.group(MockCellRangeAddress(2, 19, 0, 28, 0), 0)
        for _ in range(3):
            self.raggruppa_colonne_contabilita(sheet_contab)
        self.assertEqual(sorted(sheet_contab.grouped_calls), [
            (19, 23, 0, 0, 0), (27, 28, 0, 0, 0), (29, 30, 0, 0, 0), (31, 37, 0, 0, 0)])

    def test_ripristina_vista_mantiene_gruppi_aperti(self):
        sheet_contab = MockSheet("CONTABILITA", 2)
        # T:X aperto dall'utente, gli altri gruppi chiusi
        for start, end in ((27, 28), (29, 30), (31, 37)):
            sheet_contab.getCellRangeByPosition(start, 0, end, 0)
            for col in range(start, end + 1):
                sheet_contab.getByIndex(col).IsVisible = False
        sheet_contab.getByIndex(19).IsVisible = True
        ns = {"LeenoUtils": self.mock_utils, "setLarghezzaColonne": MagicMock()}
        exec("GRUPPI_CONTABILITA = ((19, 23), (27, 28), (29, 30), (31, 37))", ns)
        exec("MARCATORE_GRUPPI_CONTABILITA = 'x'", ns)
        for nome in ("_gruppi_contabilita_normalizzati", "gruppi_contabilita_aperti", "ripristina_vista_contabilita"):
            exec(_estrai_funzione(SRC_SHEET_UTILS, nome), ns)
        ns["ripristina_vista_contabilita"](sheet_contab)
        ns["setLarghezzaColonne"].assert_called_once_with(sheet_contab, altezza_righe=False)
        self.assertTrue(sheet_contab.getCellRangeByPosition(19, 0, 23, 0).Columns.IsVisible)
        self.assertFalse(sheet_contab.getByIndex(27).IsVisible)
        self.assertFalse(sheet_contab.getByIndex(29).IsVisible)

        # stato letto prima e passato esplicitamente: vale quello passato
        sheet_contab.getByIndex(19).IsVisible = False
        ns["ripristina_vista_contabilita"](sheet_contab, [(19, 23)])
        self.assertTrue(sheet_contab.getCellRangeByPosition(19, 0, 23, 0).Columns.IsVisible)

    def test_raggruppa_colonne_contabilita_ignora_altri_fogli(self):
        sheet_computo = MockSheet("COMPUTO", 0)
        self.raggruppa_colonne_contabilita(sheet_computo)
        self.assertEqual(len(sheet_computo.grouped_calls), 0)

        self.raggruppa_colonne_contabilita(None)

    def test_raggruppa_colonne_contabilita_reexport_pyleeno(self):
        sheet_contab = MockSheet("CONTABILITA", 2)
        self.pyleeno_raggruppa(sheet_contab)
        self.assertEqual(len(sheet_contab.grouped_calls), 4)

    def test_struttura_off_integrazione_contabilita(self):
        sheet_contab = MockSheet("CONTABILITA", 2)
        mock_doc = MagicMock()
        mock_doc.CurrentController.ActiveSheet = sheet_contab

        ns = {
            "uno": self.mock_uno,
            "LeenoUtils": self.mock_utils,
            "LeenoSheetUtils": MagicMock(),
            "HARDATTR": 1,
        }
        ns["LeenoUtils"].getDocument.return_value = mock_doc
        ns["LeenoUtils"].DocumentRefreshContext.return_value.__enter__ = MagicMock()
        ns["LeenoUtils"].DocumentRefreshContext.return_value.__exit__ = MagicMock()
        ns["LeenoSheetUtils"].cercaUltimaVoce.return_value = 1
        ns["LeenoSheetUtils"].raggruppa_colonne_contabilita = MagicMock()

        code_struttura_off = _estrai_funzione(SRC_PYLEENO, "struttura_off")
        exec(code_struttura_off, ns)
        ns["struttura_off"]()

        ns["LeenoSheetUtils"].raggruppa_colonne_contabilita.assert_called_once_with(sheet_contab)

    def test_struttura_contab_integrazione(self):
        sheet_contab = MockSheet("CONTABILITA", 2)
        mock_doc = MagicMock()
        mock_doc.CurrentController.ActiveSheet = sheet_contab
        mock_doc.NamedRanges.hasByName.return_value = True

        ns = {
            "uno": self.mock_uno,
            "LeenoUtils": self.mock_utils,
            "LeenoSheetUtils": MagicMock(),
            "PL": MagicMock(),
        }
        ns["LeenoUtils"].getDocument.return_value = mock_doc
        ns["PL"].GetModifiers.return_value = (False, False)
        ns["LeenoSheetUtils"].raggruppa_colonne_contabilita = MagicMock()

        code_struttura_contab = _estrai_funzione(SRC_CONTAB, "struttura_CONTAB")
        exec(code_struttura_contab, ns)
        ns["struttura_CONTAB"]()

        ns["LeenoSheetUtils"].raggruppa_colonne_contabilita.assert_called_with(sheet_contab)

    def test_documento_non_normalizzato_parte_chiuso(self):
        sheet_contab = MockSheet("CONTABILITA", 2)
        sheet_contab.getByIndex(29).IsVisible = True  # AD:AE visibili nel vecchio template
        utils = MagicMock()
        info = utils.getDocument.return_value.getDocumentProperties.return_value \
            .getUserDefinedProperties.return_value.getPropertySetInfo.return_value
        info.hasPropertyByName.return_value = False
        ns = {"LeenoUtils": utils}
        exec("GRUPPI_CONTABILITA = ((19, 23), (27, 28), (29, 30), (31, 37))", ns)
        exec("MARCATORE_GRUPPI_CONTABILITA = 'x'", ns)
        for nome in ("_gruppi_contabilita_normalizzati", "gruppi_contabilita_aperti"):
            exec(_estrai_funzione(SRC_SHEET_UTILS, nome), ns)
        self.assertEqual(ns["gruppi_contabilita_aperti"](sheet_contab), [])



if __name__ == '__main__':
    unittest.main()
