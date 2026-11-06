# Copyright (C) Giuseppe Vizziello - supporto@leeno.org. Licenza LGPL 2.1 (vedi LICENSE nella radice del repository).
"""Test per la rimozione dell'area F2:L40 in copyright_LeenO quando G5 != 'NB'."""

import ast
import unittest
from pathlib import Path
from unittest.mock import MagicMock

PYLEENO_PATH = Path(__file__).resolve().parent.parent / "src" / "Ultimus.oxt" / "python" / "pythonpath" / "pyleeno.py"


def _estrai_funzione(nome):
    tree = ast.parse(PYLEENO_PATH.read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == nome:
            return node
    raise AssertionError(f"Funzione non trovata: {nome}")


class TestAdeguamentoCopyright(unittest.TestCase):
    def setUp(self):
        self.node = _estrai_funzione("adegua_tmpl")
        self.log_errori = []

        self.mock_leeno_utils = MagicMock()
        self.mock_leeno_utils.log_errore.side_effect = lambda ctx, exc=None: self.log_errori.append((ctx, exc))

        self.mock_uno = MagicMock()
        self.mock_uno.Enum.side_effect = lambda enum_type, val: f"{enum_type}.{val}"

        self.ns = {
            "LeenoUtils": self.mock_leeno_utils,
            "basic_LeenO": MagicMock(),
            "Dialogs": MagicMock(),
            "uno": self.mock_uno,
        }

        mod = ast.Module(body=[self.node], type_ignores=[])
        exec(compile(mod, str(PYLEENO_PATH), "exec"), self.ns)
        self.adegua_tmpl = self.ns["adegua_tmpl"]

    def test_rimuove_area_se_g5_diverso_da_nb(self):
        oDoc = MagicMock()
        oDoc.getDocumentProperties().getUserDefinedProperties().Versione = 219

        oSheet_copy = MagicMock()
        oSheet_copy.getCellRangeByName.side_effect = lambda name: (
            MagicMock(String="Crediti vecchi") if name == "G5"
            else MagicMock(RangeAddress="STRUCT_F2_L40") if name == "F2:L40"
            else MagicMock()
        )

        oDoc.Sheets.hasByName.side_effect = lambda name: name == "copyright_LeenO"
        oDoc.Sheets.getByName.side_effect = lambda name: oSheet_copy if name == "copyright_LeenO" else None

        self.mock_leeno_utils.getDocument.return_value = oDoc

        self.adegua_tmpl()

        oSheet_copy.removeRange.assert_called_once_with(
            "STRUCT_F2_L40",
            "com.sun.star.sheet.CellDeleteMode.UP"
        )
        self.assertEqual(len(self.log_errori), 0)

    def test_non_rimuove_area_se_g5_uguale_a_nb(self):
        oDoc = MagicMock()
        oDoc.getDocumentProperties().getUserDefinedProperties().Versione = 219

        oSheet_copy = MagicMock()
        oSheet_copy.getCellRangeByName.side_effect = lambda name: (
            MagicMock(String="NB") if name == "G5"
            else MagicMock(RangeAddress="STRUCT_F2_L40") if name == "F2:L40"
            else MagicMock()
        )

        oDoc.Sheets.hasByName.side_effect = lambda name: name == "copyright_LeenO"
        oDoc.Sheets.getByName.side_effect = lambda name: oSheet_copy if name == "copyright_LeenO" else None

        self.mock_leeno_utils.getDocument.return_value = oDoc

        self.adegua_tmpl()

        oSheet_copy.removeRange.assert_not_called()
        self.assertEqual(len(self.log_errori), 0)

    def test_nessun_errore_se_foglio_copyright_assente(self):
        oDoc = MagicMock()
        oDoc.getDocumentProperties().getUserDefinedProperties().Versione = 219
        oDoc.Sheets.hasByName.return_value = False

        self.mock_leeno_utils.getDocument.return_value = oDoc

        self.adegua_tmpl()
        self.assertEqual(len(self.log_errori), 0)


if __name__ == "__main__":
    unittest.main()

