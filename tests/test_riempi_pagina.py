# -*- coding: utf-8 -*-
"""
Test di _riempi_pagina (LeenoContab) con un foglio simulato: nessun UNO richiesto.

Regressioni coperte:
- le righe alte (descrizioni lunghe) lasciano un vuoto a fondo pagina perche' Calc non
  spezza mai una riga: una stima per modulo sull'altezza totale faceva traboccare il
  filler su una pagina con sole righe tratteggiate;
- durante la generazione il ricalcolo delle altezze e' disattivato e le righe a altezza
  automatica hanno ancora l'altezza predefinita;
- la scala di stampa ("adatta alla larghezza") cambia quante righe entrano in una pagina.
"""
import ast
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

SRC = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoContab.py"

H_FILLER = 452


class FakeDoc:
    def __init__(self, adjust=False, stile=None):
        self.IsAdjustHeightEnabled = adjust
        self._stile = stile

    @property
    def StyleFamilies(self):
        if self._stile is None:
            raise RuntimeError("nessuno stile pagina")
        return SimpleNamespace(getByName=lambda _n: SimpleNamespace(getByName=lambda _s: self._stile))


class FakeRow:
    def __init__(self, sheet, height, nuova_pagina=False, ottimale=True):
        self.sheet = sheet
        self.Height = height
        self.IsStartOfNewPage = nuova_pagina
        self.IsVisible = True
        self._ottimale = ottimale

    @property
    def OptimalHeight(self):
        return self._ottimale

    @OptimalHeight.setter
    def OptimalHeight(self, valore):
        self._ottimale = valore
        # con il ricalcolo altezze disattivato l'assegnazione non ha effetto
        if valore and self.sheet.doc.IsAdjustHeightEnabled:
            self.Height = self.sheet.reali.get(id(self), self.Height)


class FakeRows:
    def __init__(self, sheet):
        self.sheet = sheet
        self.righe = []

    def getByIndex(self, i):
        return self.righe[i]

    def insertByIndex(self, i, n):
        for _ in range(n):
            self.righe.insert(i, FakeRow(self.sheet, H_FILLER))

    def removeByIndex(self, i, n):
        del self.righe[i:i + n]


class FakeColumn:
    IsVisible = True

    def __init__(self, width):
        self.Width = width


class FakeSheet:
    PageStyle = "stile"

    def __init__(self, altezze, manuali=(), stale=None, doc=None, larghezze=()):
        """altezze: altezze reali; stale: altezze visibili prima del ricalcolo."""
        self.doc = doc or FakeDoc()
        self.reali = {}
        self._rows = FakeRows(self)
        iniziali = stale if stale is not None else altezze
        for r, (h, reale) in enumerate(zip(iniziali, altezze)):
            riga = FakeRow(self, h, r in manuali)
            self.reali[id(riga)] = reale
            self._rows.righe.append(riga)
        self._cols = [FakeColumn(w) for w in larghezze]
        self.Columns = SimpleNamespace(getByIndex=lambda i: self._cols[i])

    def getRows(self):
        return self._rows

    def getPrintTitleRows(self):
        return False

    def getPrintAreas(self):
        return ()

    def getCellRangeByPosition(self, *a):
        return MagicMock()

    def getCellByPosition(self, *a):
        return MagicMock()


def _riempi_pagina(doc=None):
    tree = ast.parse(SRC.read_text(encoding="utf-8"))
    nodo = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "_riempi_pagina")
    leeno_utils = MagicMock()
    if doc is None:
        leeno_utils.getDocument.side_effect = RuntimeError("nessun documento")  # usa h_pagina
    else:
        leeno_utils.getDocument.return_value = doc
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

    def test_altezze_non_ricalcolate_durante_la_generazione(self):
        giusto = (10000 - 6000) // H_FILLER
        # senza ricalcolo la misura e' sbagliata (sovrastima lo spazio libero)
        doc = FakeDoc(adjust=False)
        sheet = FakeSheet([3000] * 5, stale=[H_FILLER] * 5, doc=doc)
        n_errato = _riempi_pagina(doc)(sheet, 5, h_pagina=10000, margine=0, inizio_blocco=0)
        self.assertGreater(n_errato, giusto)
        # con ottimizza_altezze le righe vengono ricalcolate e lo stato ripristinato
        doc = FakeDoc(adjust=False)
        sheet = FakeSheet([3000] * 5, stale=[H_FILLER] * 5, doc=doc)
        n = _riempi_pagina(doc)(sheet, 5, h_pagina=10000, margine=0, inizio_blocco=0,
                                ottimizza_altezze=True)
        self.assertEqual(n, giusto)
        self.assertFalse(doc.IsAdjustHeightEnabled)

    def test_ricalcolo_altezze_non_tocca_righe_a_altezza_fissa(self):
        doc = FakeDoc(adjust=False)
        sheet = FakeSheet([3000] * 5, stale=[H_FILLER] * 5, doc=doc)
        fissa = sheet.getRows().getByIndex(0)
        fissa._ottimale = False
        _riempi_pagina(doc)(sheet, 5, h_pagina=10000, margine=0, inizio_blocco=0, ottimizza_altezze=True)
        self.assertEqual(fissa.Height, H_FILLER)

    def test_con_ricalcolo_attivo_le_altezze_non_vengono_forzate(self):
        # ricalcolo gia' attivo: le altezze sono aggiornate, non si forza nulla
        doc = FakeDoc(adjust=True)
        sheet = FakeSheet([3000] * 5, doc=doc)
        riga = sheet.getRows().getByIndex(0)
        riga.Height = 2500  # altezza impostata a mano ma con flag automatico
        _riempi_pagina(doc)(sheet, 5, h_pagina=10000, margine=0, inizio_blocco=0, ottimizza_altezze=True)
        self.assertEqual(riga.Height, 2500)
        self.assertTrue(doc.IsAdjustHeightEnabled)

    def test_scala_di_stampa_aumenta_la_capacita_della_pagina(self):
        # colonne larghe il doppio dell'area utile: stampa al 50%, in una pagina
        # entrano 20000 unita' di foglio invece di 10000
        stile = SimpleNamespace(Height=12000, TopMargin=1000, BottomMargin=1000,
                                HeaderIsOn=False, FooterIsOn=False,
                                ScaleToPagesX=1, ScaleToPages=0, PageScale=100,
                                Width=12000, LeftMargin=1000, RightMargin=1000)
        doc = FakeDoc(adjust=True, stile=stile)
        sheet = FakeSheet([3000] * 5, doc=doc, larghezze=[10000, 10000])
        n = _riempi_pagina(doc)(sheet, 5, last_col=1, margine=0, inizio_blocco=0)
        self.assertEqual(n, (20000 - 15000) // H_FILLER)


if __name__ == "__main__":
    unittest.main()
