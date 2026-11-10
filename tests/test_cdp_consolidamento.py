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
            "CDP_RIGHE_BLOCCO": 45,
            "VALUE": 1, "DATETIME": 2, "STRING": 4, "FORMULA": 16,
        }
        self.ns["date"].today.return_value.strftime.return_value = "01/01/2026"

        mod = ast.Module(body=[
            _estrai_funzione("_copia_blocco_template_cdp"),
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
        # insRow era prev_addr.EndRow + 1 = 46 (46 righe inserite: intestazione, riga vuota, righe 1..44 del template)
        oCdP.getRows().insertByIndex.assert_any_call(46, 46)
        self.assertEqual(cells[(1, 46)].String, "segue CdP n.2 - 1÷30")

        # Il layout del nuovo blocco viene replicato dal template: intestazione e corpo
        self.assertEqual(oCdP.copyRange.call_count, 2)

        # 3. Verifica NominaArea per il nuovo CdP (_CdP_2)
        self.ns["SheetUtils"].NominaArea.assert_called()
        args, _ = self.ns["SheetUtils"].NominaArea.call_args
        self.assertEqual(args[0], oDoc)
        self.assertEqual(args[1], "CdP")
        self.assertEqual(args[3], "_CdP_2")


class TestCopiaBloccoTemplateCdP(unittest.TestCase):
    """La copia del blocco CdP deve venire dal template, non da coordinate scritte a mano."""

    def setUp(self):
        self.leeno_utils = MagicMock()
        self.doc_utils = MagicMock()
        self.ns = {
            "LeenoUtils": self.leeno_utils,
            "DocUtils": self.doc_utils,
            "PL": MagicMock(),
            "uno": MagicMock(),
            "CDP_RIGHE_BLOCCO": 45,
        }
        mod = ast.Module(body=[_estrai_funzione("_copia_blocco_template_cdp")], type_ignores=[])
        exec(compile(mod, str(LEENO_CONTAB_PATH), "exec"), self.ns)
        self.copia = self.ns["_copia_blocco_template_cdp"]

    def _documenti(self):
        oDoc = MagicMock()
        oCdP = MagicMock()
        oTmp = MagicMock()
        oTmp.getName.return_value = "CdP_tmp"
        oDoc.getSheets().importSheet.return_value = 7
        oDoc.getSheets().getByIndex.return_value = oTmp
        oDoc.getSheets().hasByName.return_value = True
        oTemplate = MagicMock()
        self.doc_utils.loadDocument.return_value = oTemplate
        return oDoc, oCdP, oTmp, oTemplate

    def test_copia_righe_0_44_del_template_e_pulisce(self):
        oDoc, oCdP, oTmp, oTemplate = self._documenti()

        self.assertTrue(self.copia(oDoc, oCdP, 92))

        # Sorgente: riga 0 (intestazione) e righe 1..44, colonne A..F del foglio importato dal template
        self.assertEqual([c.args for c in oTmp.getCellRangeByPosition.call_args_list],
                         [(0, 0, 5, 0), (0, 1, 5, 44)])
        # Destinazione: intestazione nella prima riga del blocco, corpo dopo una riga vuota
        self.assertEqual([c.args for c in oCdP.getCellByPosition.call_args_list],
                         [(0, 92), (0, 94)])
        self.assertEqual(oCdP.copyRange.call_count, 2)
        # Foglio temporaneo rimosso e template chiuso
        oDoc.getSheets().removeByName.assert_called_once_with("CdP_tmp")
        oTemplate.close.assert_called_once_with(True)

    def test_altezze_righe_replicate_dal_template(self):
        oDoc, oCdP, oTmp, _ = self._documenti()

        self.copia(oDoc, oCdP, 92)

        richieste = [c.args[0] for c in oCdP.getRows().getByIndex.call_args_list]
        # 45 righe del template, con la riga vuota (93) saltata
        self.assertEqual(richieste, [92] + list(range(94, 94 + 44)))

    def test_errore_di_copia_ritorna_false_e_pulisce(self):
        oDoc, oCdP, _, oTemplate = self._documenti()
        oCdP.copyRange.side_effect = RuntimeError("copia fallita")

        self.assertFalse(self.copia(oDoc, oCdP, 92))

        self.leeno_utils.log_errore.assert_called()
        oDoc.getSheets().removeByName.assert_called_once_with("CdP_tmp")
        oTemplate.close.assert_called_once_with(True)

    def test_template_non_caricabile_ritorna_false(self):
        oDoc, oCdP, _, _ = self._documenti()
        self.doc_utils.loadDocument.return_value = None

        self.assertFalse(self.copia(oDoc, oCdP, 92))
        oCdP.copyRange.assert_not_called()


if __name__ == "__main__":
    unittest.main()
