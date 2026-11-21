# Copyright (C) Giuseppe Vizziello - supporto@leeno.org. Licenza LGPL 2.1 (vedi LICENSE nella radice del repository).
"""Test per la funzione s1_cella in pyleeno: ricerca riga in S1 in base a etichetta in colonna G."""

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


class TestS1Cella(unittest.TestCase):
    def setUp(self):
        self.node = _estrai_funzione("s1_cella")
        self.mock_leeno_utils = MagicMock()
        self.mock_sheet_utils = MagicMock()
        self.ns = {
            "LeenoUtils": self.mock_leeno_utils,
            "SheetUtils": self.mock_sheet_utils,
        }
        mod = ast.Module(body=[self.node], type_ignores=[])
        exec(compile(mod, str(PYLEENO_PATH), "exec"), self.ns)
        self.s1_cella = self.ns["s1_cella"]

    def _crea_mock_s1(self, tabella_g_h_i):
        """
        tabella_g_h_i: dizionario {row_idx: (val_G, val_H, val_I)}
        """
        max_r = max(tabella_g_h_i.keys()) if tabella_g_h_i else 0
        self.mock_sheet_utils.getLastUsedRow.return_value = max_r

        # Costruisce la matrice per getDataArray
        num_rows = max(max_r + 1, 351)
        matrix = []
        for r in range(num_rows):
            if r in tabella_g_h_i:
                g, h, i = tabella_g_h_i[r]
                matrix.append((g, h, i))
            else:
                matrix.append(("", "", ""))

        data_array = tuple(matrix)

        oS1 = MagicMock()
        mock_range = MagicMock()
        mock_range.getDataArray.return_value = data_array
        oS1.getCellRangeByPosition.return_value = mock_range

        celle_pos = {}
        def _get_cell_by_pos(col, row):
            if (col, row) not in celle_pos:
                val = ""
                if row in tabella_g_h_i:
                    g, h, i = tabella_g_h_i[row]
                    val = h if col == 7 else (g if col == 6 else i)
                c = MagicMock()
                c.Value = float(val) if str(val).replace('.', '', 1).isdigit() else 0.0
                c.String = str(val)
                celle_pos[(col, row)] = c
            return celle_pos[(col, row)]

        oS1.getCellByPosition.side_effect = _get_cell_by_pos
        return oS1

    def test_trova_chiavi_standard_template(self):
        tabella = {
            318: ("Oneri_di_Sicurezza", "0.00", ""),
            319: ("spese_generali", "0.15", ""),
            320: ("utili d'impresa", "0.10", ""),
            322: ("Spese_Utili_Accorpati", "0", ""),
            323: ("Sconto", "-0.10", ""),
            325: ("maggiorazione", "0.10", ""),
            334: ("[Contabilità] Descrizioni abbreviate: PRIMI caratteri della voce", "100", "desc inizio"),
            335: ("Descrizioni abbreviate: primi caratteri della voce", "120", "[Contabilità] Descrizioni abbreviate: ULTIMI caratteri della voce"),
            336: ("Descrizioni abbreviate: primi caratteri della voce", "100", "desc comp inizio"),
            337: ("Descrizioni abbreviate: ultimi caratteri della voce", "120", "desc comp fine"),
        }
        oS1 = self._crea_mock_s1(tabella)

        c_cont_in = self.s1_cella(oS1, "cont_inizio", "H335")
        self.assertEqual(c_cont_in.Value, 100.0)
        oS1.getCellByPosition.assert_any_call(7, 334)

        c_cont_fi = self.s1_cella(oS1, "cont_fine", "H336")
        self.assertEqual(c_cont_fi.Value, 120.0)
        oS1.getCellByPosition.assert_any_call(7, 335)

        c_comp_in = self.s1_cella(oS1, "comp_inizio", "H337")
        self.assertEqual(c_comp_in.Value, 100.0)
        oS1.getCellByPosition.assert_any_call(7, 336)

        c_comp_fi = self.s1_cella(oS1, "comp_fine", "H338")
        self.assertEqual(c_comp_fi.Value, 120.0)
        oS1.getCellByPosition.assert_any_call(7, 337)

        c_spese = self.s1_cella(oS1, "spese_generali", "H320")
        self.assertEqual(c_spese.Value, 0.15)
        oS1.getCellByPosition.assert_any_call(7, 319)

    def test_trova_con_righe_spostate(self):
        # Sposta tutte le righe di 20 posizioni verso il basso
        delta = 20
        tabella = {
            319 + delta: ("spese_generali", "0.18", ""),
            334 + delta: ("[Contabilità] Descrizioni abbreviate: PRIMI caratteri della voce", "80", ""),
            335 + delta: ("[Contabilità] Descrizioni abbreviate: ULTIMI caratteri della voce", "90", ""),
            336 + delta: ("[Computo e Variante] Descrizioni abbreviate: primi caratteri della voce", "110", ""),
            337 + delta: ("[Computo e Variante] Descrizioni abbreviate: ultimi caratteri della voce", "130", ""),
        }
        oS1 = self._crea_mock_s1(tabella)

        c_spese = self.s1_cella(oS1, "spese_generali", "H320")
        self.assertEqual(c_spese.Value, 0.18)
        oS1.getCellByPosition.assert_any_call(7, 319 + delta)

        c_cont_in = self.s1_cella(oS1, "cont_inizio", "H335")
        self.assertEqual(c_cont_in.Value, 80.0)
        oS1.getCellByPosition.assert_any_call(7, 334 + delta)

        c_cont_fi = self.s1_cella(oS1, "cont_fine", "H336")
        self.assertEqual(c_cont_fi.Value, 90.0)
        oS1.getCellByPosition.assert_any_call(7, 335 + delta)

        c_comp_in = self.s1_cella(oS1, "comp_inizio", "H337")
        self.assertEqual(c_comp_in.Value, 110.0)
        oS1.getCellByPosition.assert_any_call(7, 336 + delta)

        c_comp_fi = self.s1_cella(oS1, "comp_fine", "H338")
        self.assertEqual(c_comp_fi.Value, 130.0)
        oS1.getCellByPosition.assert_any_call(7, 337 + delta)

    def test_fallback_quando_etichetta_assente(self):
        oS1 = self._crea_mock_s1({})
        mock_fallback = MagicMock()
        mock_fallback.Value = 999.0
        oS1.getCellRangeByName.return_value = mock_fallback

        c = self.s1_cella(oS1, "chiave_inesistente", "H337")
        self.assertEqual(c, mock_fallback)
        oS1.getCellRangeByName.assert_called_with("S1.H337")


if __name__ == "__main__":
    unittest.main()
