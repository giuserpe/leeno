# -*- coding: utf-8 -*-
"""
Test di _riempi_pagina (LeenoContab) con un foglio simulato: nessun UNO richiesto.

Regressione: le righe alte (descrizioni lunghe) lasciano un vuoto a fondo pagina
perché Calc non spezza mai una riga. La stima per modulo sull'altezza totale lo
ignorava e il filler traboccava su una pagina con sole righe tratteggiate.
"""
import ast
import unittest
from pathlib import Path
from unittest.mock import MagicMock

SRC = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoContab.py"

H_FILLER = 452


class FakeRow:
    def __init__(self, height, nuova_pagina=False, visibile=True):
        self.Height = height
        self.IsStartOfNewPage = nuova_pagina
        self.IsVisible = visibile


class FakeRows:
    def __init__(self, righe):
        self.righe = righe

    def getByIndex(self, i):
        return self.righe[i]

    def insertByIndex(self, i, n):
        for _ in range(n):
            self.righe.insert(i, FakeRow(H_FILLER))

    def removeByIndex(self, i, n):
        del self.righe[i:i + n]


class FakeSheet:
    PageStyle = "stile"

    def __init__(self, altezze, manuali=()):
        self._rows = FakeRows([FakeRow(h, r in manuali) for r, h in enumerate(altezze)])

    def getRows(self):
        return self._rows

    def getPrintTitleRows(self):
        return False

    def getCellRangeByPosition(self, *a):
        return MagicMock()

    def getCellByPosition(self, *a):
        return MagicMock()


def _riempi_pagina():
    tree = ast.parse(SRC.read_text(encoding="utf-8"))
    nodo = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "_riempi_pagina")
    leeno_utils = MagicMock()
    leeno_utils.getDocument.side_effect = RuntimeError("nessun documento")  # usa h_pagina
    ns = {"PL": MagicMock(), "LeenoUtils": leeno_utils}
    exec(compile(ast.Module(body=[nodo], type_ignores=[]), str(SRC), "exec"), ns)
    return ns["_riempi_pagina"]


class TestRiempiPagina(unittest.TestCase):
    def test_righe_alte_non_fanno_traboccare_il_filler(self):
        # pagina da 10000: 3 righe da 3000 per pagina, la quarta va nella seconda.
        # 5 righe: seconda pagina occupata per 6000, restano 4000 = 8 righe filler.
        # Il vecchio calcolo per modulo (15000 % 10000 = 5000) ne metteva 11.
        sheet = FakeSheet([3000] * 5)
        n = _riempi_pagina()(sheet, 5, h_pagina=10000, margine=0, inizio_blocco=0)
        self.assertEqual(n, (10000 - 6000) // H_FILLER)

    def test_parte_dal_salto_pagina_manuale(self):
        # salto manuale alla riga 2: pagina occupata solo dalle righe 2 e 3
        sheet = FakeSheet([3000] * 4, manuali={2})
        n = _riempi_pagina()(sheet, 4, h_pagina=10000, margine=0)
        self.assertEqual(n, (10000 - 6000) // H_FILLER)

    def test_inizio_blocco_ignora_righe_precedenti(self):
        sheet = FakeSheet([9000, 9000, 2000, 2000])
        n = _riempi_pagina()(sheet, 4, h_pagina=10000, margine=0, inizio_blocco=2)
        self.assertEqual(n, (10000 - 4000) // H_FILLER)

    def test_pagina_piena_non_inserisce_filler(self):
        sheet = FakeSheet([9900])
        self.assertEqual(_riempi_pagina()(sheet, 1, h_pagina=10000, margine=0, inizio_blocco=0), 0)
        self.assertEqual(len(sheet.getRows().righe), 1)


if __name__ == "__main__":
    unittest.main()
