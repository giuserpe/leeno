# -*- coding: utf-8 -*-
"""
Test di intestazione, piè di pagina e copertina nell'esportazione PDF (LeenoPdf):
testi per documento, firme del Libretto, numero del SAL, aree dei SAL e chiamate a pdfExport.
Senza UNO: LibreOffice e i moduli di LeenO sono sostituiti da oggetti finti; il lettore di S2
e le funzioni dei segnaposto sono quelli veri, compilati dai sorgenti.
"""
import json
import os
import re
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

try:
    from tests.estrai_sorgente import FoglioFinto, carica
except ImportError:  # esecuzione con la cartella tests nel percorso di ricerca
    from estrai_sorgente import FoglioFinto, carica

ROOT = Path(__file__).resolve().parent.parent
PYTHONPATH_DIR = ROOT / "src/Ultimus.oxt/python/pythonpath"
FIRME = "L'IMPRESA\t\t\t\t\t\tIL DIRETTORE DEI LAVORI    "

S2 = {
    (1, 2): "OGGETTO:", (2, 2): "Lavori di prova",
    (1, 5): "Stazione Appaltante \u2013 Committente:", (2, 5): "Comune di Matera",
}


def _impostazioni(ps):
    """LeenoSettings finto: funzioni vere dei segnaposto e dei documenti, loadPrintSettings fisso."""
    lettore = {"LeenoUtils": SimpleNamespace(log_errore=MagicMock())}
    carica(PYTHONPATH_DIR / "SheetUtils.py", ("leggi_dato_etichettato",), ns=lettore)
    ns = carica(
        PYTHONPATH_DIR / "LeenoSettings.py",
        ("nome_file_documento", "valori_segnaposto", "sostituisci_segnaposto", "testi_predefiniti",
         "_personalizzazioni", "impostazioni_documento", "senza_firme", "ha_firme", "piede_con_firme",
         "titolo_documento"),
        ("_SEGNAPOSTI_S2", "_SEGNAPOSTI_LAVORO", "_FIRME_TESTO", "_SEGNAPOSTO_FIRME", "_PARTI_PIEDE",
         "_DOCUMENTI_STAMPA", "_DOCUMENTI_SENZA_TESTATA", "_DOCUMENTI_CON_FIRME", "_TESTI_STAMPA",
         "_TESTI_PRECEDENTI"),
        {"os": os, "re": re, "json": json,
         "SheetUtils": SimpleNamespace(leggi_dato_etichettato=lettore["leggi_dato_etichettato"]),
         "loadJobSettings": lambda _doc: {}})
    ns["loadPrintSettings"] = lambda _doc: (dict(ps), [])
    return SimpleNamespace(**{k: v for k, v in ns.items() if not k.startswith("__")})


def _documento():
    s2 = FoglioFinto(S2)
    return SimpleNamespace(
        Sheets=SimpleNamespace(getByName=lambda nome: s2 if nome == "S2" else MagicMock(name=nome)),
        getURL=lambda: "file:///W:/lavori/Prova%20Esempio.ods")


def _carica_pdf(ps=None, **extra):
    aree = carica(PYTHONPATH_DIR / "LeenoContab.py", costanti=("AREE_SAL",))["AREE_SAL"]
    ns = {
        "os": os,
        "LeenoSettings": _impostazioni(ps or {}),
        "LeenoContab": SimpleNamespace(AREE_SAL=aree),
        "LeenoUtils": SimpleNamespace(getDocument=MagicMock()),
        "SheetUtils": MagicMock(),
        "DocUtils": MagicMock(),
    }
    ns.update(extra)
    return carica(
        PYTHONPATH_DIR / "LeenoPdf.py",
        ("prepareCover", "prepareHeaderFooter", "_pagineFoglio", "_esportaPdfFoglio",
         "PdfElencoPrezzi", "PdfComputoMetrico", "PdfAttoSal"),
        ("_DOCUMENTI_SAL",), ns)


class TestTestata(unittest.TestCase):
    def _testata(self, chiave, docSubst, ps=None):
        ns = _carica_pdf(ps)
        return ns["prepareHeaderFooter"](_documento(), docSubst, chiave)

    def test_computo_come_la_stampa_a_schermo(self):
        t = self._testata("COMPUTO", {"[OGGETTO]": "Computo Metrico", "[NUMERO_DOCUMENTO]": "1"})
        self.assertEqual(t, {
            "intSx": "Committente: Comune di Matera", "intCenter": "Lavori di prova", "intDx": "Computo Metrico",
            "ppSx": "realizzato con LeenO: Prova Esempio.ods", "ppCenter": "",
            "ppDx": "Pagina [PAGINA] di [PAGINE]"})

    def test_libretto_ha_le_firme_da_sole_in_fondo_e_il_numero_del_sal(self):
        t = self._testata("LIBRETTO", {"[OGGETTO]": "Libretto delle Misure", "[SAL]": "2"})
        self.assertEqual(t["ppCenter"], "\n" + FIRME)
        self.assertEqual(t["intDx"], "Libretto delle Misure n. 2")

    def test_le_firme_non_si_sovrappongono_a_realizzato_con_e_alle_pagine(self):
        t = self._testata("LIBRETTO", {"[OGGETTO]": "Libretto delle Misure", "[SAL]": "2"})
        # tutte e tre le parti hanno due righe: testo sulla prima, firme sulla seconda del centro
        self.assertEqual(t["ppSx"], "realizzato con LeenO: Prova Esempio.ods\n")
        self.assertEqual(t["ppDx"], "Pagina [PAGINA] di [PAGINE]\n")
        self.assertEqual(t["ppCenter"].split("\n"), ["", FIRME])

    def test_registro_senza_firme_salvo_scelta(self):
        docSubst = {"[OGGETTO]": "Registro di Contabilità", "[SAL]": "3"}
        self.assertEqual(self._testata("REGISTRO", docSubst)["ppCenter"], "")
        ps = {"documenti": json.dumps({"REGISTRO": {"ppCenter": "[FIRME]"}})}
        t = self._testata("REGISTRO", docSubst, ps)
        self.assertEqual((t["ppCenter"], t["intDx"]), ("\n" + FIRME, "Registro di Contabilità n. 3"))

    def test_libretto_senza_firme_se_l_utente_le_toglie(self):
        ps = {"documenti": json.dumps({"LIBRETTO": {"ppCenter": ""}})}
        t = self._testata("LIBRETTO", {"[OGGETTO]": "Libretto delle Misure"}, ps)
        self.assertEqual(t["ppCenter"], "")

    def test_firme_sotto_un_piede_su_piu_righe(self):
        ps = {"documenti": json.dumps({"LIBRETTO": {"ppSx": "riga 1\nriga 2 [FILE]", "ppCenter": "Centro [FIRME]"}})}
        t = self._testata("LIBRETTO", {"[OGGETTO]": "Libretto delle Misure"}, ps)
        self.assertEqual(t["ppCenter"], "Centro\n\n" + FIRME)
        self.assertEqual(t["ppCenter"].split("\n")[-1], FIRME)
        self.assertEqual(len(t["ppSx"].split("\n")), 3)
        self.assertEqual(len(t["ppDx"].split("\n")), 3)

    def test_firme_da_sole_se_il_piede_e_vuoto(self):
        ps = {"documenti": json.dumps({"LIBRETTO": {"ppSx": "", "ppDx": ""}})}  # centro: [FIRME]
        t = self._testata("LIBRETTO", {"[OGGETTO]": "Libretto delle Misure"}, ps)
        self.assertEqual((t["ppSx"], t["ppCenter"], t["ppDx"]), ("", FIRME, ""))

    def test_certificato_di_pagamento_senza_testata(self):
        t = self._testata("CDP", {"[OGGETTO]": "Certificato di Pagamento"}, {"intSx": "Mio ente"})
        self.assertEqual(set(t.values()), {""})
        self.assertEqual(len(t), 6)

    def test_personalizzazione_del_documento(self):
        ps = {"documenti": json.dumps({"ELENCO": {"ppSx": "Tavola [NUMERO_DOCUMENTO] - [OGGETTO]"}})}
        t = self._testata("ELENCO", {"[OGGETTO]": "Elenco Prezzi", "[NUMERO_DOCUMENTO]": "2"}, ps)
        self.assertEqual(t["ppSx"], "Tavola 2 - Elenco Prezzi")

    def test_stringhe_comuni_della_versione_precedente_restano_valide(self):
        ps = {"intSx": "Ente: [COMMITTENTE]", "ppDx": "Pag. [PAGINA]"}
        t = self._testata("COMPUTO", {"[OGGETTO]": "Computo Metrico"}, ps)
        self.assertEqual((t["intSx"], t["ppDx"]), ("Ente: Comune di Matera", "Pag. [PAGINA]"))

    def test_firme_scritte_in_un_documento_senza_firme_predefinite(self):
        ps = {"documenti": json.dumps({"COMPUTO": {"ppSx": "realizzato con LeenO: [FILE] [FIRME]"}})}
        t = self._testata("COMPUTO", {"[OGGETTO]": "Computo Metrico"}, ps)
        self.assertEqual(t["ppSx"], "realizzato con LeenO: Prova Esempio.ods\n")
        self.assertEqual(t["ppCenter"], "\n" + FIRME)

    def test_il_segnaposto_firme_non_resta_nel_testo(self):
        for chiave in ("COMPUTO", "VARIANTE", "ELENCO", "ANALISI", "LIBRETTO", "REGISTRO", "SAL", "CDP"):
            with self.subTest(chiave=chiave):
                t = self._testata(chiave, {"[OGGETTO]": "x"})
                self.assertNotIn("[FIRME]", "".join(t.values()))

    def test_pagina_e_pagine_restano_per_i_campi_numerici(self):
        t = self._testata("COMPUTO", {"[OGGETTO]": "x", "[PAGINE]": "9"})
        self.assertIn("[PAGINA]", t["ppDx"])
        self.assertEqual(t["ppDx"], "Pagina [PAGINA] di [PAGINE]")


class TestEsportazione(unittest.TestCase):
    def setUp(self):
        self.doc = _documento()
        self.foglio = SimpleNamespace(RowPageBreaks=[SimpleNamespace(Position=p) for p in (2, 50, 100, 150, 200)])
        self.doc.Sheets = SimpleNamespace(getByName=lambda nome: self.foglio if nome != "S2" else FoglioFinto(S2))
        self.prepareCover = MagicMock()
        self.sheet_utils = MagicMock()
        self.ns = _carica_pdf(SheetUtils=self.sheet_utils)
        self.ns["prepareCover"] = self.prepareCover  # sostituisce la funzione vera compilata dal sorgente
        self.ns["LeenoUtils"].getDocument.return_value = self.doc

    def _argomenti_pdfexport(self):
        (args, _kw), = self.sheet_utils.pdfExport.call_args_list
        return args

    def test_elenco_prezzi(self):
        self.ns["PdfElencoPrezzi"](os.path.join("cartella", ""), 3)
        oDoc, fogli, dest, testata, cover, aree = self._argomenti_pdfexport()
        self.assertEqual(dest, os.path.join("cartella", "3-ElencoPrezzi.pdf"))
        self.assertEqual(fogli, [self.foglio])
        self.assertEqual(testata["intDx"], "Elenco Prezzi")
        self.assertIsNone(aree)

    def test_computo_metrico(self):
        self.ns["PdfComputoMetrico"]("c", 1)
        testata = self._argomenti_pdfexport()[3]
        self.assertEqual(testata["intDx"], "Computo Metrico")
        self.assertEqual(testata["intCenter"], "Lavori di prova")

    def test_pagine_nella_copertina_dopo_la_testata(self):
        self.ns["PdfComputoMetrico"]("c", 1)
        cover = self._argomenti_pdfexport()[4]
        cover("documento", "nuovo")
        (args, _), = self.prepareCover.call_args_list
        self.assertEqual(args[:2], ("documento", "nuovo"))
        self.assertEqual(args[2]["[PAGINE]"], 4)
        self.assertEqual(args[2]["[OGGETTO]"], "Computo Metrico")
        self.assertEqual(args[2]["[NUMERO_DOCUMENTO]"], "1")

    def _atto(self, tipo, nSal, aree_presenti):
        area = SimpleNamespace(StartRow=10, EndRow=149, StartColumn=0, EndColumn=3)
        nomi = {n: SimpleNamespace(ReferredCells=SimpleNamespace(RangeAddress=area)) for n in aree_presenti}
        self.doc.NamedRanges = SimpleNamespace(hasByName=lambda n: n in nomi, getByName=lambda n: nomi[n])
        return self.ns["PdfAttoSal"]("c", 5, tipo, nSal, self.doc), area

    def test_libretto_del_sal(self):
        esito, area = self._atto("Libretto", 2, ["_Lib_2"])
        self.assertTrue(esito)
        oDoc, fogli, dest, testata, cover, aree = self._argomenti_pdfexport()
        self.assertEqual(dest, os.path.join("c", "5-LibrettoMisure_SAL2.pdf"))
        self.assertEqual(testata["ppCenter"], "\n" + FIRME)
        self.assertEqual(testata["intDx"], "Libretto delle Misure n. 2")
        self.assertEqual(aree, {"CONTABILITA": area})
        self.assertEqual(area.EndColumn, 11)

    def test_altri_atti_contabili(self):
        for tipo, prefisso, file, titolo, colonna in (
                ("Registro", "_Reg_", "RegistroContabilita_SAL4.pdf", "Registro di Contabilità n. 4", 9),
                ("SAL", "_SAL_", "SAL_SAL4.pdf", "Stato di Avanzamento Lavori n. 4", 5),
                ("CdP", "_CdP_", "CertificatoPagamento_SAL4.pdf", "", 5)):
            with self.subTest(tipo=tipo):
                self.sheet_utils.pdfExport.reset_mock()
                esito, area = self._atto(tipo, 4, [prefisso + "4"])
                self.assertTrue(esito)
                args = self._argomenti_pdfexport()
                self.assertEqual(args[2], os.path.join("c", "5-" + file))
                self.assertEqual(args[3]["intDx"], titolo)
                self.assertEqual(area.EndColumn, colonna)

    def test_area_del_sal_mancante(self):
        esito, _ = self._atto("Libretto", 9, ["_Lib_1"])
        self.assertFalse(esito)
        self.sheet_utils.pdfExport.assert_not_called()


class TestCopertina(unittest.TestCase):
    def test_i_segnaposto_della_copertina_usano_s2_e_il_documento(self):
        sheet_utils = MagicMock()
        doc_utils = MagicMock()
        cover_doc = MagicMock()
        cover_doc.Sheets.__contains__.return_value = True
        cover_doc.Sheets.__getitem__.return_value = SimpleNamespace(PageStyle="", PrintAreas=())
        doc_utils.loadDocument.return_value = cover_doc
        ns = _carica_pdf({"fileCopertine": "c.ods", "copertina": "Copertina 1"},
                         SheetUtils=sheet_utils, DocUtils=doc_utils)
        nDoc = MagicMock()
        nDoc.Sheets.Count = 1
        docSubst = {"[OGGETTO]": "Libretto delle Misure", "[NUMERO_DOCUMENTO]": "5", "[PAGINE]": 7, "[SAL]": "2"}
        self.assertTrue(ns["prepareCover"](_documento(), nDoc, docSubst))
        (_, sostituzioni), _ = sheet_utils.replaceText.call_args
        self.assertEqual(sostituzioni["[PROGETTO]"], "Lavori di prova")
        self.assertEqual(sostituzioni["[COMMITTENTE]"], "Comune di Matera")
        self.assertEqual(sostituzioni["[FILE]"], "Prova Esempio.ods")
        self.assertEqual((sostituzioni["[OGGETTO]"], sostituzioni["[PAGINE]"], sostituzioni["[SAL]"]),
                         ("Libretto delle Misure", 7, "2"))

    def test_senza_copertina_non_fa_nulla(self):
        ns = _carica_pdf({"fileCopertine": "", "copertina": ""})
        self.assertFalse(ns["prepareCover"](_documento(), MagicMock(), {}))


if __name__ == "__main__":
    unittest.main()
