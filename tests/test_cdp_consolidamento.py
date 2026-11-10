# Copyright (C) Giuseppe Vizziello - supporto@leeno.org. Licenza LGPL 2.1 (vedi LICENSE nella radice del repository).
"""Test di unità per la gestione consolidamento, raggruppamento e vista del CdP."""

import ast
import unittest
from pathlib import Path
from unittest.mock import MagicMock

REPO_ROOT = Path(__file__).resolve().parent.parent
LEENO_CONTAB_PATH = REPO_ROOT / "src" / "Ultimus.oxt" / "python" / "pythonpath" / "LeenoContab.py"


def _estrai_funzione(nome_funzione):
    """Estrae l'AST di una funzione da LeenoContab.py."""
    tree = ast.parse(LEENO_CONTAB_PATH.read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == nome_funzione:
            return node
    raise AssertionError(f"Funzione {nome_funzione} non trovata in {LEENO_CONTAB_PATH}")


class TestMostraSalCdP(unittest.TestCase):
    def test_mostra_sal_include_cdp(self):
        """Verifica che in mostra_sal la tupla ('CdP', '_CdP_', 5) sia inclusa tra gli atti gestiti."""
        fn_ast = _estrai_funzione("mostra_sal")
        codice = ast.unparse(fn_ast)
        self.assertIn("('CdP', '_CdP_', 5)", codice)


class TestAnnullaUltimoSalCdP(unittest.TestCase):
    def test_annulla_sal_rimuove_righe_cdp_se_multi_sal(self):
        """Verifica che _annulla_ultimo_sal_core su SAL > 1 rimuova le righe e il NamedRange di CdP."""
        fn_ast = _estrai_funzione("_annulla_ultimo_sal_core")
        codice = ast.unparse(fn_ast)
        self.assertIn("nome_cdp_del = '_CdP_' + nSal_str", codice)
        self.assertIn("oSheet_C.Rows.removeByIndex", codice)
        self.assertIn("oRanges.removeByName(nome_cdp_del)", codice)


class TestGeneraCdPComportamento(unittest.TestCase):
    def setUp(self):
        leeno_utils = MagicMock()
        sheet_utils = MagicMock()
        leeno_sheet_utils = MagicMock()
        pl_mock = MagicMock()

        self.ns = {
            "LeenoUtils": leeno_utils,
            "SheetUtils": sheet_utils,
            "LeenoSheetUtils": leeno_sheet_utils,
            "PL": pl_mock,
            "uno": MagicMock(),
            "DocUtils": MagicMock(),
            "Dialogs": MagicMock(),
            "DLG": MagicMock(),
            "numero_in_lettere_euro": lambda v: "zero",
            "_leggi_iva_da_S2": lambda s: 0.22,
            "_leggi_anticipo_da_S2": lambda s: 0.0,
            "_leggi_infortuni_da_S2": lambda s: 0.005,
            "_leggi_dato_anagrafico": lambda s, l: "Dato",
            "_col_letter": lambda c: chr(65 + c),
            "date": MagicMock(),
        }
        self.ns["date"].today.return_value.strftime.return_value = "01/01/2026"

        mod = ast.Module(body=[
            _estrai_funzione("setup_foglio_CdP"),
            _estrai_funzione("GeneraCdP")
        ], type_ignores=[])
        exec(compile(mod, str(LEENO_CONTAB_PATH), "exec"), self.ns)

    def test_genera_cdp_sal_maggiore_di_uno_consolida_e_titolo(self):
        """Verifica che GeneraCdP su nSal=2 consolidi il CdP 1 e crei il titolo 'segue CdP' con stile 'uuuuu'."""
        GeneraCdP = self.ns["GeneraCdP"]

        oDoc = MagicMock()
        oCdP = MagicMock()
        oS2 = MagicMock()

        # Sheets
        def get_sheet(name):
            if name == 'CdP': return oCdP
            if name == 'S2': return oS2
            raise KeyError(name)

        oDoc.getSheets().getByName.side_effect = get_sheet
        oDoc.getSheets().hasByName.side_effect = lambda n: n in ('CdP', 'S2')
        oDoc.Sheets.getByName.side_effect = get_sheet

        # Cells map per coordinate
        cells = {}

        def get_cell(col, row):
            key = (col, row)
            if key not in cells:
                c = MagicMock()
                c.Value = 1234.56
                c.String = "1234.56"
                c.Type.value = "FLOAT"
                cells[key] = c
            return cells[key]

        oCdP.getCellByPosition.side_effect = get_cell
        oS2.getCellByPosition.side_effect = get_cell

        # NamedRanges
        oPrevRange = MagicMock()
        oPrevRange.ReferredCells.RangeAddress.StartColumn = 0
        oPrevRange.ReferredCells.RangeAddress.StartRow = 0
        oPrevRange.ReferredCells.RangeAddress.EndColumn = 5
        oPrevRange.ReferredCells.RangeAddress.EndRow = 45

        prev_cell_range = MagicMock()
        prev_cell_range.getDataArray.return_value = (("val1", "val2"),)
        oCdP.getCellRangeByPosition.return_value = prev_cell_range

        def has_nr(name):
            return name in ("_CdP_1",)

        def get_nr(name):
            if name == "_CdP_1":
                return oPrevRange
            raise KeyError(name)

        oDoc.NamedRanges.hasByName.side_effect = has_nr
        oDoc.NamedRanges.getByName.side_effect = get_nr

        # S2 markers
        self.ns["SheetUtils"].uFindString.return_value = (5, 10)  # yS2, xS2
        self.ns["SheetUtils"].getLastUsedRow.return_value = 45

        dati = (2, 1, 30, 0, 0, [], 0, 0, [])
        GeneraCdP(oDoc, dati=dati)

        # 1. Verifica consolidamento CdP 1: setDataArray chiamato con i dati del getDataArray
        prev_cell_range.setDataArray.assert_called_with((("val1", "val2"),))

        # 2. Verifica inserimento raggruppamento e titolo "segue CdP n.2 - 1÷30" con stile "uuuuu"
        # insRow era prev_addr.EndRow + 1 = 46 (47 righe inserite per il nuovo blocco)
        oCdP.getRows().insertByIndex.assert_any_call(46, 47)
        self.assertEqual(cells[(1, 46)].String, "segue CdP n.2 - 1÷30")

        # 3. Verifica NominaArea per il nuovo CdP (_CdP_2)
        self.ns["SheetUtils"].NominaArea.assert_called()
        args, _ = self.ns["SheetUtils"].NominaArea.call_args
        self.assertEqual(args[0], oDoc)
        self.assertEqual(args[1], "CdP")
        self.assertEqual(args[3], "_CdP_2")


if __name__ == "__main__":
    unittest.main()
