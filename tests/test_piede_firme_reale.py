# -*- coding: utf-8 -*-
"""
Prova su PDF reale della riga delle firme nel piè di pagina del Libretto.

Scrive nello stile del Libretto di tests/data/LeenO_Esempio.ods il piè di pagina costruito con
LeenoSettings.piede_con_firme (e i campi numerici di SheetUtils.paginationFields), esporta in PDF
un SAL del Libretto e misura con pdftotext le coordinate delle parole: le firme devono stare sotto
l'ultima riga di testo senza sovrapporsi, e il numero di pagine non deve cambiare.
Gira solo con LEENO_TEST_UNO=1, LibreOffice e poppler (pdftotext).
"""
import re
import shutil
import subprocess
import unittest
from pathlib import Path
from types import SimpleNamespace

try:
    from tests import uno_harness
    from tests.estrai_sorgente import carica
except ImportError:  # esecuzione con la cartella tests nel percorso di ricerca
    import uno_harness
    from estrai_sorgente import carica

ROOT = Path(__file__).resolve().parent.parent
PYTHONPATH_DIR = ROOT / "src/Ultimus.oxt/python/pythonpath"
ESEMPIO = ROOT / "tests/data/LeenO_Esempio.ods"
STILE = "Page_Style_Libretto_Misure2"


def _disponibile():
    return uno_harness.uno_disponibile(ESEMPIO) and shutil.which("pdftotext") is not None


def _parole(pdf, pagina=1):
    """{parola: [(yMin, yMax)]} dalla pagina `pagina` del PDF."""
    out = subprocess.run(["pdftotext", "-bbox", "-f", str(pagina), "-l", str(pagina), pdf, "-"],
                         capture_output=True, text=True).stdout
    res = {}
    for _x1, y1, _x2, y2, parola in re.findall(
            r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]+)</word>', out):
        res.setdefault(parola, []).append((float(y1), float(y2)))
    return res


@unittest.skipUnless(_disponibile(), "richiede LEENO_TEST_UNO=1, uno, soffice, LeenO_Esempio.ods e pdftotext")
class TestPiedeFirmeReale(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import uno
        cls.uno = uno
        cls.sessione = uno_harness.SessioneUno()
        cls.addClassCleanup(cls.sessione.chiudi)
        cls.sessione.avvia()
        ns = carica(PYTHONPATH_DIR / "LeenoSettings.py", ("piede_con_firme", "senza_firme", "ha_firme"),
                    ("_FIRME_TESTO", "_SEGNAPOSTO_FIRME", "_PARTI_PIEDE"))
        cls.piede_con_firme = staticmethod(ns["piede_con_firme"])
        cls.firme = ns["_FIRME_TESTO"]
        utils = carica(PYTHONPATH_DIR / "LeenoUtils.py", ("replacePatternWithField",))
        sheet = carica(PYTHONPATH_DIR / "SheetUtils.py", ("paginationFields",),
                       ns={"LeenoUtils": SimpleNamespace(**utils)})
        cls.paginationFields = staticmethod(sheet["paginationFields"])
        cls.contatore = 0

    def _esporta(self, sx=None, centro=None, dx=None, con_firme=True):
        """Esporta il Libretto del SAL 1 con il piè di pagina dato; ritorna il percorso del PDF."""
        uno = self.uno
        type(self).contatore += 1
        doc = self.sessione.apri_copia(ESEMPIO, "firme%d" % self.contatore)
        area = doc.NamedRanges.getByName("_Lib_1").ReferredCells.RangeAddress
        area.EndColumn = 11
        stile = doc.StyleFamilies.getByName("PageStyles").getByName(STILE)
        if sx is not None:
            testi = {"ppSx": sx, "ppCenter": centro or "", "ppDx": dx or ""}
            if con_firme:
                testi = self.piede_con_firme(testi)
            contenuto = stile.RightPageFooterContent
            contenuto.LeftText.String = testi["ppSx"]
            contenuto.CenterText.String = testi["ppCenter"]
            contenuto.RightText.String = testi["ppDx"]
            for parte in (contenuto.LeftText, contenuto.CenterText, contenuto.RightText):
                self.paginationFields(doc, parte)
            stile.RightPageFooterContent = contenuto
        foglio = doc.Sheets.getByName("CONTABILITA")
        foglio.PrintAreas = (area,)
        intervallo = foglio.getCellRangeByPosition(area.StartColumn, area.StartRow, area.EndColumn, area.EndRow)
        intervallo.Rows.IsVisible = True
        filtro = uno.Any("[]com.sun.star.beans.PropertyValue",
                         (uno.createUnoStruct("com.sun.star.beans.PropertyValue", "Selection", 0, intervallo, 0),))
        props = (uno.createUnoStruct("com.sun.star.beans.PropertyValue", "FilterName", 0, "calc_pdf_Export", 0),
                 uno.createUnoStruct("com.sun.star.beans.PropertyValue", "FilterData", 0, filtro, 0))
        pdf = str(Path(self.sessione.tmp) / ("firme%d.pdf" % self.contatore))
        doc.storeToURL(uno.systemPathToFileUrl(pdf), props)
        return pdf

    @staticmethod
    def _pagine(pdf):
        out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
        return int(re.search(r"Pages:\s+(\d+)", out).group(1))

    def _distanza(self, pdf, marcatori):
        """Distanza (pt) tra la parte bassa dell'ultima riga di testo e la parte alta delle firme."""
        parole = _parole(pdf)
        self.assertIn("DIRETTORE", parole, "firme non trovate nel PDF")
        basso_testo = max(y2 for m in marcatori for _y1, y2 in parole[m])
        alto_firme = min(y1 for y1, _y2 in parole["DIRETTORE"])
        return alto_firme - basso_testo

    def test_una_riga_di_piede(self):
        pdf = self._esporta("realizzato LUNO", "[FIRME]", "Pagina [PAGINA] di [PAGINE] DUNO")
        self.assertGreaterEqual(self._distanza(pdf, ["LUNO", "DUNO"]), -0.1)

    def test_il_segnaposto_firme_in_sinistra_funziona_allo_stesso_modo(self):
        pdf = self._esporta("realizzato LUNO [FIRME]", "", "Pagina [PAGINA] di [PAGINE] DUNO")
        self.assertGreaterEqual(self._distanza(pdf, ["LUNO", "DUNO"]), -0.1)
        testo = subprocess.run(["pdftotext", "-f", "1", "-l", "1", pdf, "-"], capture_output=True, text=True).stdout
        self.assertNotIn("[FIRME]", testo)

    def test_il_numero_di_pagina_diventa_un_campo(self):
        pdf = self._esporta("realizzato LUNO", "[FIRME]", "Pagina [PAGINA] di [PAGINE] DUNO")
        testo = subprocess.run(["pdftotext", "-f", "1", "-l", "1", pdf, "-"], capture_output=True, text=True).stdout
        self.assertRegex(testo, r"Pagina 1 di \d+")
        self.assertNotIn("[PAGINA]", testo)

    def test_sinistra_su_due_righe(self):
        pdf = self._esporta("riga uno\nriga due LDUE", "[FIRME]", "Pagina [PAGINA] di [PAGINE] DUNO")
        self.assertGreaterEqual(self._distanza(pdf, ["LDUE", "DUNO"]), -0.1)

    def test_destra_su_tre_righe_e_centro_con_testo(self):
        pdf = self._esporta("sinistra LUNO", "centro CUNO [FIRME]", "dx 1\ndx 2\ndx 3 DTRE")
        self.assertGreaterEqual(self._distanza(pdf, ["LUNO", "CUNO", "DTRE"]), -0.1)

    def test_senza_la_correzione_le_firme_si_sovrapporrebbero(self):
        # controllo della misura: con la sola riga vuota davanti alle firme c'e' sovrapposizione
        pdf = self._esporta("realizzato LUNO", "\n" + self.firme, "Pagina [PAGINA] di [PAGINE] DUNO",
                            con_firme=False)
        self.assertLess(self._distanza(pdf, ["LUNO", "DUNO"]), -1)

    def test_il_numero_di_pagine_non_cambia(self):
        originale = self._esporta()
        con_firme = self._esporta("realizzato LUNO", "[FIRME]", "Pagina [PAGINA] di [PAGINE] DUNO")
        self.assertEqual(self._pagine(con_firme), self._pagine(originale))


if __name__ == "__main__":
    unittest.main()
