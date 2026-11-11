# Copyright (C) Giuseppe Vizziello - supporto@leeno.org. Licenza LGPL 2.1 (vedi LICENSE nella radice del repository).
"""Test di caratterizzazione per funzioni pure e isolate in LeenoSheetUtils."""

import ast
import unittest
from itertools import groupby
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
LEENO_SHEET_UTILS_PATH = BASE_DIR / "src/Ultimus.oxt/python/pythonpath/LeenoSheetUtils.py"


def _carica_sheet_utils_pure():
    tree = ast.parse(LEENO_SHEET_UTILS_PATH.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name in ("optimize_visibility_string", "find_hidden_ranges"):
            nodi.append(n)
        elif isinstance(n, ast.FunctionDef) and n.name == "getColumnWidths":
            # Estrae la funzione interna col_index_to_letter
            for sub_n in n.body:
                if isinstance(sub_n, ast.FunctionDef) and sub_n.name == "col_index_to_letter":
                    nodi.append(sub_n)

    ns = {
        "groupby": groupby,
    }
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(LEENO_SHEET_UTILS_PATH), "exec"), ns)
    return ns


class TestSheetUtilsPure(unittest.TestCase):
    def setUp(self):
        self.utils = _carica_sheet_utils_pure()

    def test_col_index_to_letter_base(self):
        col_index_to_letter = self.utils["col_index_to_letter"]
        self.assertEqual(col_index_to_letter(0), "A")
        self.assertEqual(col_index_to_letter(1), "B")
        self.assertEqual(col_index_to_letter(25), "Z")

    def test_col_index_to_letter_doppia_lettera(self):
        col_index_to_letter = self.utils["col_index_to_letter"]
        self.assertEqual(col_index_to_letter(26), "AA")
        self.assertEqual(col_index_to_letter(27), "AB")
        self.assertEqual(col_index_to_letter(51), "AZ")
        self.assertEqual(col_index_to_letter(52), "BA")
        self.assertEqual(col_index_to_letter(701), "ZZ")

    def test_col_index_to_letter_tripla_lettera(self):
        col_index_to_letter = self.utils["col_index_to_letter"]
        self.assertEqual(col_index_to_letter(702), "AAA")

    def test_optimize_visibility_string_corta(self):
        opt = self.utils["optimize_visibility_string"]
        # Per ripetizioni <= 5, non usa moltiplicazione
        self.assertEqual(opt("TTT"), "TTT")
        self.assertEqual(opt("TTTTT"), "TTTTT")
        self.assertEqual(opt("TTTFF"), "TTT+FF")

    def test_optimize_visibility_string_lunga(self):
        opt = self.utils["optimize_visibility_string"]
        # Per ripetizioni > 5, compatta con 'K'*count
        self.assertEqual(opt("TTTTTT"), "'T'*6")
        self.assertEqual(opt("TTT" + "F" * 10), "TTT+'F'*10")
        self.assertEqual(opt("T" * 7 + "F" * 8 + "T" * 2), "'T'*7+'F'*8+TT")

    def test_find_hidden_ranges_nessun_nascosto(self):
        find_hidden = self.utils["find_hidden_ranges"]
        self.assertEqual(find_hidden(["T", "T", "T"]), [])
        self.assertEqual(find_hidden([]), [])

    def test_find_hidden_ranges_tutti_nascosti(self):
        find_hidden = self.utils["find_hidden_ranges"]
        self.assertEqual(find_hidden(["F", "F", "F", "F"]), [(0, 3)])

    def test_find_hidden_ranges_misti(self):
        find_hidden = self.utils["find_hidden_ranges"]
        # F a inizio, centro e fine
        vis = ["F", "F", "T", "F", "F", "F", "T", "T", "F"]
        ranges = find_hidden(vis)
        self.assertEqual(ranges, [(0, 1), (3, 5), (8, 8)])


if __name__ == "__main__":
    unittest.main()
