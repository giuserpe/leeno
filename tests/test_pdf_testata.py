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
         "titolo_documento", "titolo_a_schermo", "testata_a_schermo"),
        ("_SEGNAPOSTI_S2", "_SEGNAPOSTI_LAVORO", "_FIRME_TESTO", "_SEGNAPOSTO_FIRME", "_PARTI_PIEDE",
         "_DOCUMENTI_STAMPA", "_DOCUMENTI_SENZA_TESTATA", "_DOCUMENTI_CON_FIRME", "_TESTI_STAMPA",
         "_TESTI_PRECEDENTI", "_ATTI_CONTABILI", "_FOGLIO_DEL_DOCUMENTO",
         "STILE_PAGINA_RAFFRONTO"),
        {"os": os, "re": re, "json": json,
         "SheetUtils": SimpleNamespace(leggi_dato_etichettato=lettore["leggi_dato_etichettato"]),
         "loadJobSettings": lambda _doc: {}})
    ns["loadPrintSettings"] = lambda _doc: (dict(ps), [])
    return SimpleNamespace(**{k: v for k, v in ns.items() if not k.startswith("__")})


def _documento():
    s2 = FoglioFinto(S2)
    return SimpleNamespace(
        Sheets=SimpleNamespace(getByName=lambda nome: s2 if nome == "S2" else MagicMock(name=nome),
                               hasByName=lambda nome: nome == "S2"),
        getURL=lambda: "file:///W:/lavori/Prova%20Esempio.ods")


def _carica_pdf(ps=None, **extra):
    aree = carica(PYTHONPATH_DIR / "LeenoContab.py", costanti=("AREE_SAL",))["AREE_SAL"]
    ns = {
        "os": os,
        "LeenoSettings": _impostazioni(ps or {}),
        "LeenoContab": SimpleNamespace(AREE_SAL=aree),
        "LeenoUtils": SimpleNamespace(getDocument=MagicMock()),
        "LeenoPagina": MagicMock(),
        "LeenoSheetUtils": MagicMock(),
        "SheetUtils": MagicMock(),
        "DocUtils": MagicMock(),
    }
    ns.update(extra)
    return carica(
        PYTHONPATH_DIR / "LeenoPdf.py",
        ("prepareCover", "_pagineFoglio", "_esportaPdfFoglio", "_areaElenco",
         "PdfElencoPrezzi", "PdfQuadroRaffronto", "PdfComputoMetrico", "PdfAttoSal"),
        ("_DOCUMENTI_SAL",), ns)


def _area(colonna_iniziale, riga_iniziale, colonna_finale, riga_finale):
    return SimpleNamespace(StartColumn=colonna_iniziale, StartRow=riga_iniziale,
                           EndColumn=colonna_finale, EndRow=riga_finale)


class TestTestata(unittest.TestCase):
    """Le stringhe che il motore (LeenoPagina) scrive negli stili, con le funzioni vere di LeenoSettings."""

    def _testata(self, chiave, docSubst, ps=None):
        ls = _impostazioni(ps or {})
        doc = _documento()
        valori = ls.valori_segnaposto(
            doc, documento=docSubst.get("[OGGETTO]", ""), nSal=docSubst.get("[SAL]") or None,
            nomefile=ls.nome_file_documento(doc),
            extra={k: v for k, v in docSubst.items() if k == "[NUMERO_DOCUMENTO]"})
        testi = ls.impostazioni_documento(ls.loadPrintSettings(doc)[0], chiave)
        return ls.testata_a_schermo(testi, valori, chiave)

    def test_computo_come_la_stampa_a_schermo(self):
        t = self._testata("COMPUTO", {"[OGGETTO]": "Computo Metrico", "[NUMERO_DOCUMENTO]": "1"})
        self.assertEqual(t, {
            "intSx": "\n\nCommittente: Comune di Matera", "intCenter": "Lavori di prova\n\n",
            "intDx": "\n\nComputo Metrico",
            "ppSx": "\nrealizzato con LeenO: Prova Esempio.ods", "ppCenter": "",
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
        self.assertEqual(t["ppSx"], "\nTavola 2 - Elenco Prezzi")

    def test_stringhe_comuni_della_versione_precedente_restano_valide(self):
        ps = {"intSx": "Ente: [COMMITTENTE]", "ppDx": "Pag. [PAGINA]"}
        t = self._testata("COMPUTO", {"[OGGETTO]": "Computo Metrico"}, ps)
        self.assertEqual((t["intSx"], t["ppDx"]), ("\n\nEnte: Comune di Matera", "Pag. [PAGINA]"))

    def test_firme_scritte_in_un_documento_senza_firme_predefinite(self):
        ps = {"documenti": json.dumps({"COMPUTO": {"ppSx": "realizzato con LeenO: [FILE] [FIRME]"}})}
        t = self._testata("COMPUTO", {"[OGGETTO]": "Computo Metrico"}, ps)
        self.assertEqual(t["ppSx"], "\nrealizzato con LeenO: Prova Esempio.ods\n")
        # tre righe a sinistra (la prima vuota dell'impaginazione): le firme sulla terza del centro
        self.assertEqual(t["ppCenter"], "\n\n" + FIRME)

    def test_titoli_degli_altri_atti_contabili(self):
        for chiave, oggetto, atteso in (
                ("REGISTRO", "Registro di Contabilità", "Registro di Contabilità n. 4"),
                ("SAL", "Stato di Avanzamento Lavori", "Stato di Avanzamento Lavori n. 4")):
            with self.subTest(chiave=chiave):
                t = self._testata(chiave, {"[OGGETTO]": oggetto, "[SAL]": "4"})
                self.assertEqual(t["intDx"], atteso)

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
        self.range_chiesti = []
        self.foglio = SimpleNamespace(
            RowPageBreaks=[SimpleNamespace(Position=p) for p in (2, 50, 100, 150, 200)],
            getCellRangeByPosition=lambda *a: (self.range_chiesti.append(a)
                                               or SimpleNamespace(RangeAddress=_area(*a))))
        self.doc.Sheets = SimpleNamespace(getByName=lambda nome: self.foglio if nome != "S2" else FoglioFinto(S2),
                                          hasByName=lambda nome: nome == "S2")
        self.prepareCover = MagicMock()
        self.sheet_utils = MagicMock()
        self.sheet_utils.getLastUsedRow.return_value = 120
        self.ns = _carica_pdf(SheetUtils=self.sheet_utils)
        self.ns["prepareCover"] = self.prepareCover  # sostituisce la funzione vera compilata dal sorgente
        self.ns["LeenoUtils"].getDocument.return_value = self.doc
        self.pagina = self.ns["LeenoPagina"]
        self.sbianca = self.ns["LeenoSheetUtils"].SbiancaCellePrintArea

    def _argomenti_pdfexport(self):
        (args, _kw), = self.sheet_utils.pdfExport.call_args_list
        return args

    def _applica_stili(self, nome_stile, foglio):
        """Esegue la funzione degli stili passata a pdfExport su una copia finta e ritorna (stile, copia)."""
        applicaStili = self._argomenti_pdfexport()[3]
        stile = SimpleNamespace(DisplayName=nome_stile)
        copia = MagicMock()
        copia.StyleFamilies.getByName.return_value.getByName.return_value = stile
        applicaStili(copia, [nome_stile])
        (args, _), = self.pagina.applica_stile_pagina.call_args_list
        self.assertEqual(args[0], self.doc)
        self.assertIs(args[1], stile)
        self.assertEqual(args[2], foglio)
        self.assertIs(args[4], copia)
        return args[3]

    def test_elenco_prezzi(self):
        self.ns["PdfElencoPrezzi"](os.path.join("cartella", ""), 3)
        oDoc, fogli, dest, applicaStili, cover, aree = self._argomenti_pdfexport()
        self.assertEqual(dest, os.path.join("cartella", "3-ElencoPrezzi.pdf"))
        self.assertEqual(fogli, [self.foglio])
        # colonne fino alla J (le nascoste non si stampano); righe come set_area_stampa: da 2 a ultima - 1
        self.assertEqual(aree, {"Elenco Prezzi": _area(0, 2, 9, 119)})
        self.pagina.contesto_stampa.assert_called_once_with(
            self.doc, None, {"[NUMERO_DOCUMENTO]": "3"}, "Elenco Prezzi")
        self._applica_stili("PageStyle_Elenco Prezzi", "Elenco Prezzi")

    def test_computo_metrico(self):
        self.ns["PdfComputoMetrico"]("c", 1)
        self.pagina.contesto_stampa.assert_called_once_with(self.doc, None, {"[NUMERO_DOCUMENTO]": "1"}, None)
        ctx = self._applica_stili("PageStyle_COMPUTO_A4", "COMPUTO")
        self.assertIs(ctx, self.pagina.contesto_stampa.return_value)

    def _prepara_copia(self):
        preparaCopia = self._argomenti_pdfexport()[3]
        copia = MagicMock()
        preparaCopia(copia, ["PageStyle_COMPUTO_A4"])
        return copia

    def test_fondo_bianco_sulla_copia_per_il_foglio_esportato(self):
        self.ns["PdfComputoMetrico"]("c", 1)
        copia = self._prepara_copia()
        copia.Sheets.getByName.assert_called_once_with("COMPUTO")
        self.sbianca.assert_called_once_with(copia.Sheets.getByName.return_value)

    def test_fondo_bianco_predefinito_attivo_e_disattivabile(self):
        self.ns["PdfElencoPrezzi"]("c", 1, False)
        self._prepara_copia()
        self.sbianca.assert_not_called()
        self.sheet_utils.pdfExport.reset_mock()
        self.ns["PdfElencoPrezzi"]("c", 1)
        self._prepara_copia()
        self.sbianca.assert_called_once()

    def test_fondo_bianco_negli_atti_contabili(self):
        self._atto("SAL", 3, ["_SAL_3"])
        self.ns["PdfAttoSal"]("c", 5, "SAL", 3, self.doc, False)
        self.sbianca.assert_not_called()

    def test_quadro_di_raffronto_a4_orizzontale_fino_alla_colonna_z(self):
        self.ns["PdfQuadroRaffronto"](os.path.join("cartella", ""), 4)
        oDoc, fogli, dest, preparaCopia, cover, aree = self._argomenti_pdfexport()
        self.assertEqual(dest, os.path.join("cartella", "4-QuadroRaffronto.pdf"))
        self.assertEqual(aree, {"Elenco Prezzi": _area(0, 2, 25, 120)})
        self.pagina.contesto_stampa.assert_called_once_with(
            self.doc, None, {"[NUMERO_DOCUMENTO]": "4"}, "Quadro di Raffronto")
        copia = MagicMock()
        foglio_copia = copia.Sheets.getByName.return_value
        foglio_copia.PageStyle = "PageStyle_Elenco Prezzi"
        preparaCopia(copia, ["PageStyle_Elenco Prezzi"])
        # nuovo stile orizzontale ricavato da quello dell'Elenco e assegnato al foglio della copia
        self.sheet_utils.copiaStilePaginaOrizzontale.assert_called_once_with(
            copia, "PageStyle_Elenco Prezzi", "PageStyle_Raffronto_A4")
        self.assertEqual(foglio_copia.PageStyle, "PageStyle_Raffronto_A4")
        # il motore lavora solo sullo stile orizzontale, con il foglio Elenco Prezzi
        (args, _), = self.pagina.applica_stile_pagina.call_args_list
        self.assertEqual(args[2], "Elenco Prezzi")
        self.assertIs(args[4], copia)

    def test_elenco_prezzi_resta_verticale(self):
        self.ns["PdfElencoPrezzi"]("c", 1)
        self._argomenti_pdfexport()[3](MagicMock(), ["PageStyle_Elenco Prezzi"])
        self.sheet_utils.copiaStilePaginaOrizzontale.assert_not_called()

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
        oDoc, fogli, dest, applicaStili, cover, aree = self._argomenti_pdfexport()
        self.assertEqual(dest, os.path.join("c", "5-LibrettoMisure_SAL2.pdf"))
        self.assertEqual(aree, {"CONTABILITA": area})
        self.assertEqual(area.EndColumn, 11)
        self.pagina.contesto_stampa.assert_called_once_with(self.doc, 2, {"[NUMERO_DOCUMENTO]": "5"}, None)
        self._applica_stili("Page_Style_Libretto_Misure2", "CONTABILITA")

    def test_altri_atti_contabili(self):
        for tipo, prefisso, file, colonna in (
                ("Registro", "_Reg_", "RegistroContabilita_SAL4.pdf", 9),
                ("SAL", "_SAL_", "SAL_SAL4.pdf", 5),
                ("CdP", "_CdP_", "CertificatoPagamento_SAL4.pdf", 5)):
            with self.subTest(tipo=tipo):
                self.sheet_utils.pdfExport.reset_mock()
                self.pagina.contesto_stampa.reset_mock()
                esito, area = self._atto(tipo, 4, [prefisso + "4"])
                self.assertTrue(esito)
                args = self._argomenti_pdfexport()
                self.assertEqual(args[2], os.path.join("c", "5-" + file))
                self.assertEqual(self.pagina.contesto_stampa.call_args[0][1], 4)
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
