# -*- coding: utf-8 -*-
"""
Test del dialogo di esportazione PDF e dell'aspetto delle celle nel PDF, senza UNO:
- DlgEsportazionePdf.xdl e impostazioni di LeenoPdf coerenti;
- SbiancaCellePrintArea su un foglio finto.
"""
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

try:
    from tests.estrai_sorgente import carica
except ImportError:  # esecuzione con la cartella tests nel percorso di ricerca
    from estrai_sorgente import carica

ROOT = Path(__file__).resolve().parent.parent
PYTHONPATH_DIR = ROOT / "src/Ultimus.oxt/python/pythonpath"
XDL = ROOT / "src/Ultimus.oxt/UltimusFree2/DlgEsportazionePdf.xdl"
DLG = "{http://openoffice.org/2000/dialog}"


def _controlli():
    radice = ET.parse(XDL).getroot()
    return [(e.get(DLG + "id"), int(e.get(DLG + "tab-index")), e) for e in radice.iter()
            if e.get(DLG + "tab-index") is not None]


class TestDialogoEsportazione(unittest.TestCase):
    def setUp(self):
        self.pdf = carica(PYTHONPATH_DIR / "LeenoPdf.py", (),
                          ("_EXPORTSETTINGSITEMS", "_EXPORTCBDEFAULT"))

    def test_ogni_impostazione_ha_il_suo_controllo(self):
        ids = {i for i, _, _ in _controlli()}
        self.assertFalse(set(self.pdf["_EXPORTSETTINGSITEMS"]) - ids)

    def test_ordine_di_tabulazione_senza_duplicati(self):
        indici = [t for _, t, _ in _controlli()]
        self.assertEqual(sorted(indici), list(range(len(indici))))

    def test_fondo_bianco_attivo_per_impostazione_predefinita(self):
        self.assertIn("cbFondoBianco", self.pdf["_EXPORTSETTINGSITEMS"])
        self.assertTrue(self.pdf["_EXPORTCBDEFAULT"]["cbFondoBianco"])
        casella = {i: e for i, _, e in _controlli()}["cbFondoBianco"]
        self.assertEqual(casella.get(DLG + "checked"), "true")

    def test_configurazione_predefinita(self):
        sorgente = (PYTHONPATH_DIR / "LeenoConfig.py").read_text(encoding="utf-8")
        self.assertIn("('ImpostazioniExport', 'cbFondoBianco', '(bool)True')", sorgente)

    def test_il_dialogo_non_ha_controlli_sovrapposti_al_pulsante_ok(self):
        c = {i: e for i, _, e in _controlli()}
        fondo = int(c["lblCartella"].get(DLG + "top"))
        casella = c["cbFondoBianco"]
        self.assertLess(int(casella.get(DLG + "top")) + int(casella.get(DLG + "height")), fondo)
        altezza = int(ET.parse(XDL).getroot().get(DLG + "height"))
        self.assertLess(int(c["btnOk"].get(DLG + "top")) + int(c["btnOk"].get(DLG + "height")), altezza)


class TestDialogoPiuVecchioDelCodice(unittest.TestCase):
    """Dialogo installato senza la casella del fondo bianco: niente errore, valore predefinito."""

    def setUp(self):
        self.ns = {"LeenoUtils": MagicMock(), "Dialogs": MagicMock(), "os": __import__("os"),
                   "_SfogliaCartella": MagicMock()}
        carica(PYTHONPATH_DIR / "LeenoPdf.py", ("PdfDialog", "_leggiImpostazioniExport"),
               ("_EXPORTSETTINGSITEMS", "_EXPORTSETTINGSCHECKS", "_EXPORTCBDEFAULT", "_EXPORTNPDEFAULT",
                "_DLG_ESPORTAZIONE_PDF", "_DOCUMENTI_SAL"), self.ns)

    def _dialogo(self):
        oDlg = MagicMock()
        oModel = oDlg.getModel.return_value
        oModel.hasByName.side_effect = lambda nome: nome != "cbFondoBianco"
        self.ns["LeenoUtils"].getComponentContext.return_value.ServiceManager.createInstance \
            .return_value.createDialog.return_value = oDlg
        return oDlg, oModel

    def test_apertura_senza_la_casella(self):
        oDlg, oModel = self._dialogo()
        self.ns["PdfDialog"]({}, ())
        richiesti = [c.args[0] for c in oModel.getByName.call_args_list]
        self.assertNotIn("cbFondoBianco", richiesti)

    def test_lettura_senza_la_casella_usa_il_predefinito(self):
        oDlg, oModel = self._dialogo()
        res = self.ns["_leggiImpostazioniExport"](oDlg, {})
        self.assertIs(res["cbFondoBianco"], True)
        res = self.ns["_leggiImpostazioniExport"](oDlg, {"cbFondoBianco": False})
        self.assertIs(res["cbFondoBianco"], False)


class TestCopiaDiStampa(unittest.TestCase):
    def test_il_documento_copia_non_viene_mostrato(self):
        sorgente = (PYTHONPATH_DIR / "SheetUtils.py").read_text(encoding="utf-8")
        inizio = sorgente.index("def pdfExport")
        corpo = sorgente[inizio:sorgente.index("\ndef ", inizio + 10)]
        self.assertIn("createSheetDocument(Hidden=True)", corpo)
        self.assertNotIn("Hidden=False", corpo)


class _StileFinto:
    """Stile di pagina con proprieta' in un dizionario; alcune di sola lettura."""
    _SOLA_LETTURA = ("DisplayName",)

    def __init__(self, props):
        object.__setattr__(self, "p", dict(props))
        object.__setattr__(self, "PropertySetInfo",
                           SimpleNamespace(Properties=[SimpleNamespace(Name=k) for k in props]))

    def getPropertyValue(self, nome):
        return self.p[nome]

    def setPropertyValue(self, nome, valore):
        if nome in self._SOLA_LETTURA:
            raise RuntimeError(nome)
        self.p[nome] = valore

    def __getattr__(self, nome):
        return self.p[nome]

    def __setattr__(self, nome, valore):
        self.p[nome] = valore


class _FamigliaFinta:
    def __init__(self, stili):
        self.stili = dict(stili)

    def getByName(self, nome):
        return self.stili[nome]

    def hasByName(self, nome):
        return nome in self.stili

    def insertByName(self, nome, stile):
        self.stili[nome] = stile


class TestStileOrizzontale(unittest.TestCase):
    def _doc(self):
        origine = _StileFinto({"DisplayName": "Elenco", "Width": 21000, "Height": 29700, "IsLandscape": False,
                               "LeftMargin": 1000, "PageScale": 100})
        famiglia = _FamigliaFinta({"PageStyle_Elenco Prezzi": origine})
        doc = MagicMock()
        doc.StyleFamilies.getByName.return_value = famiglia
        doc.createInstance.side_effect = lambda servizio: _StileFinto({})
        return doc, famiglia, origine

    def _copia(self, doc, nome="PageStyle_Raffronto_A4"):
        ns = carica(PYTHONPATH_DIR / "SheetUtils.py", ("copiaStilePaginaOrizzontale",))
        return ns["copiaStilePaginaOrizzontale"](doc, "PageStyle_Elenco Prezzi", nome)

    def test_copia_in_orizzontale_scambiando_larghezza_e_altezza(self):
        doc, famiglia, origine = self._doc()
        nuovo = self._copia(doc)
        self.assertIs(famiglia.getByName("PageStyle_Raffronto_A4"), nuovo)
        self.assertEqual((nuovo.Width, nuovo.Height, nuovo.IsLandscape), (29700, 21000, True))
        self.assertEqual((nuovo.LeftMargin, nuovo.PageScale), (1000, 100))

    def test_lo_stile_di_origine_non_cambia(self):
        doc, famiglia, origine = self._doc()
        self._copia(doc)
        self.assertEqual((origine.Width, origine.Height, origine.IsLandscape), (21000, 29700, False))

    def test_proprieta_di_sola_lettura_non_bloccano(self):
        doc, _, _ = self._doc()
        self.assertNotIn("DisplayName", self._copia(doc).p)

    def test_secondo_passaggio_aggiorna_lo_stile_esistente(self):
        doc, famiglia, _ = self._doc()
        primo = self._copia(doc)
        secondo = self._copia(doc)
        self.assertIs(primo, secondo)
        self.assertEqual(doc.createInstance.call_count, 1)


class _Intervallo:
    def __init__(self):
        self.CellBackColor = None


class _FoglioFinto:
    def __init__(self, aree, stili):
        self._aree, self._stili, self.intervalli = aree, stili, {}

    def getPrintAreas(self):
        return self._aree

    def getCellRangeByPosition(self, c1, r1, c2, r2):
        return self.intervalli.setdefault((c1, r1, c2, r2), _Intervallo())

    def getCellByPosition(self, col, riga):
        return SimpleNamespace(CellStyle=self._stili.get(riga, "Normale"))


class TestSbiancaCelle(unittest.TestCase):
    def _funzione(self):
        utils = MagicMock()
        ns = {"LeenoUtils": utils, "LeenoGlobals": MagicMock(), "COLORE_BIANCO_SFONDO": 0xFFFFFF}
        ns["LeenoGlobals"].getGlobalVar.return_value = ["Cat"]
        carica(PYTHONPATH_DIR / "LeenoSheetUtils.py", ("SbiancaCellePrintArea",), ns=ns)
        return ns["SbiancaCellePrintArea"], utils

    def test_foglio_dato_bianco_e_grigio_sulle_categorie(self):
        sbianca, utils = self._funzione()
        area = SimpleNamespace(StartColumn=0, StartRow=2, EndColumn=5, EndRow=6)
        foglio = _FoglioFinto([area], {4: "Cat"})
        sbianca(foglio)
        utils.getDocument.assert_not_called()
        self.assertEqual(foglio.intervalli[(0, 2, 5, 6)].CellBackColor, 0xFFFFFF)
        self.assertEqual(foglio.intervalli[(0, 4, 40, 4)].CellBackColor, int("eeeeee", 16))
        self.assertNotIn((0, 3, 40, 3), foglio.intervalli)

    def test_senza_area_di_stampa_non_fa_nulla(self):
        sbianca, _ = self._funzione()
        foglio = _FoglioFinto([], {})
        sbianca(foglio)
        self.assertEqual(foglio.intervalli, {})

    def test_senza_foglio_usa_quello_attivo(self):
        sbianca, utils = self._funzione()
        area = SimpleNamespace(StartColumn=0, StartRow=0, EndColumn=1, EndRow=1)
        attivo = _FoglioFinto([area], {})
        utils.getDocument.return_value.CurrentController.ActiveSheet = attivo
        sbianca()
        self.assertEqual(attivo.intervalli[(0, 0, 1, 1)].CellBackColor, 0xFFFFFF)


if __name__ == "__main__":
    unittest.main()
