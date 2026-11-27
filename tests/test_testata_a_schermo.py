# -*- coding: utf-8 -*-
"""
Test dell'intestazione e del piè di pagina a schermo (LeenoSettings, usate da MENU_sistema_pagine).

- unitari, senza UNO: documento dello stile, titoli a schermo, disposizione delle stringhe;
- documento reale (opzionale, LEENO_TEST_UNO=1): confronto tra il vecchio codice di
  MENU_sistema_pagine (copiato qui come riferimento) e le nuove funzioni su LeenO_Esempio.ods,
  campi numerici di pagina e prova su PDF delle firme in uno stile del Computo.
"""
import json
import re
import shutil
import subprocess
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

try:
    from tests import uno_harness
    from tests.estrai_sorgente import carica
except ImportError:  # esecuzione con la cartella tests nel percorso di ricerca
    import uno_harness
    from estrai_sorgente import carica

ROOT = Path(__file__).resolve().parent.parent
PYTHONPATH_DIR = ROOT / "src/Ultimus.oxt/python/pythonpath"
ESEMPIO = ROOT / "tests/data/LeenO_Esempio.ods"
PYLEENO = PYTHONPATH_DIR / "pyleeno.py"

FUNZIONI = ("testi_predefiniti", "_personalizzazioni", "impostazioni_documento", "senza_firme", "ha_firme",
            "piede_con_firme", "sostituisci_segnaposto", "titolo_documento", "documento_dello_stile",
            "dimensione_testata", "titolo_a_schermo", "testata_a_schermo", "scrivi_testata_stile")
COSTANTI = ("_DOCUMENTI_STAMPA", "_DOCUMENTI_SENZA_TESTATA", "_TESTI_STAMPA", "_TESTI_PRECEDENTI",
            "_FIRME_TESTO", "_SEGNAPOSTO_FIRME", "_PARTI_PIEDE", "STILI_TESTATA", "STILI_ATTI_CONTABILI",
            "_ATTI_CONTABILI", "_FONT_TESTATA",
            "_HTXT_PREDEFINITO", "_HTXT_ANALISI", "_FOGLIO_DEL_DOCUMENTO")
FIRME = "L'IMPRESA\t\t\t\t\t\tIL DIRETTORE DEI LAVORI    "


def _carica(**extra):
    ns = {"json": json, "re": re, "SheetUtils": MagicMock()}
    ns.update(extra)
    return carica(PYTHONPATH_DIR / "LeenoSettings.py", FUNZIONI, COSTANTI, ns)


def _colonna(visibile):
    return SimpleNamespace(IsVisible=visibile, Columns=SimpleNamespace(IsVisible=visibile))


def _foglio(nome, nascoste=()):
    colonne = {i: _colonna(i not in nascoste) for i in range(40)}
    return SimpleNamespace(Name=nome, getColumns=lambda: SimpleNamespace(getByIndex=lambda i: colonne[i]))


def _documento(*fogli):
    per_nome = {f.Name: f for f in fogli}
    return SimpleNamespace(Sheets=SimpleNamespace(hasByName=lambda n: n in per_nome,
                                                  getByName=lambda n: per_nome[n]))


class TestDocumentoDelloStile(unittest.TestCase):
    def setUp(self):
        self.ns = _carica()
        self.doc = self.ns["documento_dello_stile"]

    def test_elenco_prezzi_e_sempre_elenco(self):
        for foglio in ("COMPUTO", "Elenco Prezzi", "CdP", "Analisi di Prezzo"):
            self.assertEqual(self.doc("PageStyle_Elenco Prezzi", foglio), "ELENCO")

    def test_stile_condiviso_dal_computo_dalla_variante_e_dall_analisi(self):
        self.assertEqual(self.doc("PageStyle_COMPUTO_A4", "COMPUTO"), "COMPUTO")
        self.assertEqual(self.doc("PageStyle_COMPUTO_A4", "VARIANTE"), "VARIANTE")
        self.assertEqual(self.doc("PageStyle_COMPUTO_A4", "Analisi di Prezzo"), "ANALISI")

    def test_per_ogni_altro_foglio_vale_il_computo(self):
        for foglio in ("CdP", "S2", "QUADRO ECONOMICO", "CONTABILITA", "Elenco Prezzi", ""):
            self.assertEqual(self.doc("PageStyle_COMPUTO_A4", foglio), "COMPUTO", foglio)

    def test_stile_analisi(self):
        self.assertEqual(self.doc("PageStyle_Analisi di Prezzo", "COMPUTO"), "ANALISI")

    def test_stili_degli_atti_contabili_non_dipendono_dal_foglio_attivo(self):
        for stile, chiave in (("Page_Style_Libretto_Misure2", "LIBRETTO"), ("PageStyle_REGISTRO_A4", "REGISTRO"),
                              ("PageStyle_SAL_A4", "SAL")):
            for foglio in ("COMPUTO", "CONTABILITA", "Registro", "S2"):
                self.assertEqual(self.doc(stile, foglio), chiave, (stile, foglio))
        self.assertEqual(self.ns["STILI_ATTI_CONTABILI"], {"Page_Style_Libretto_Misure2", "PageStyle_REGISTRO_A4",
                                                           "PageStyle_SAL_A4"})

    def test_altri_stili_senza_testata(self):
        for nome in ("Default", "Page_Style_COPERTINE", "PageStyle_CdP", ""):
            self.assertIsNone(self.doc(nome, "COMPUTO"), nome)

    def test_ogni_stile_con_testata_ha_un_documento(self):
        for stile in self.ns["STILI_TESTATA"] | self.ns["STILI_ATTI_CONTABILI"]:
            self.assertIn(self.doc(stile, "COMPUTO"), [c for c, _ in self.ns["_DOCUMENTI_STAMPA"]])

    def test_dimensione_del_testo(self):
        self.assertEqual(self.ns["dimensione_testata"]("ANALISI"), 9.0)
        for chiave in ("COMPUTO", "VARIANTE", "ELENCO", "LIBRETTO"):
            self.assertEqual(self.ns["dimensione_testata"](chiave), 8.0)


class TestTitoloAschermo(unittest.TestCase):
    def setUp(self):
        self.titolo = _carica()["titolo_a_schermo"]

    # la colonna 30 (AE) mostra l'incidenza della manodopera: visibile = vista "Incidenza MdO"
    def test_computo(self):
        self.assertEqual(self.titolo("COMPUTO", _documento(_foglio("COMPUTO", nascoste=(30,)))),
                         "Computo Metrico Estimativo")
        self.assertEqual(self.titolo("COMPUTO", _documento(_foglio("COMPUTO"))), "Incidenza MdO")

    def test_variante(self):
        self.assertEqual(self.titolo("VARIANTE", _documento(_foglio("VARIANTE", nascoste=(30,)))),
                         "Perizia di Variante")
        self.assertEqual(self.titolo("VARIANTE", _documento(_foglio("VARIANTE"))), "Incidenza MdO (Variante)")

    def test_elenco_prezzi(self):
        self.assertEqual(self.titolo("ELENCO", _documento(_foglio("Elenco Prezzi"))), "Elenco Prezzi")
        for nascosta in (11, 12, 13):
            doc = _documento(_foglio("Elenco Prezzi", nascoste=(nascosta,)))
            self.assertEqual(self.titolo("ELENCO", doc), "Quadro di Raffronto", nascosta)

    def test_analisi_usa_il_nome_del_foglio(self):
        self.assertEqual(self.titolo("ANALISI", _documento(_foglio("Analisi di Prezzo"))), "Analisi di Prezzo")

    def test_il_foglio_del_documento_non_e_quello_attivo(self):
        # Elenco Prezzi con le colonne nascoste si vede anche se il foglio attivo e' un altro
        doc = _documento(_foglio("COMPUTO", nascoste=(30,)), _foglio("Elenco Prezzi", nascoste=(12,)))
        self.assertEqual(self.titolo("ELENCO", doc), "Quadro di Raffronto")
        self.assertEqual(self.titolo("COMPUTO", doc), "Computo Metrico Estimativo")

    def test_foglio_assente_ripiega_sul_titolo_base(self):
        self.assertEqual(self.titolo("VARIANTE", _documento(_foglio("COMPUTO"))), "Variante")
        self.assertEqual(self.titolo("ELENCO", _documento()), "Elenco Prezzi")
        self.assertEqual(self.titolo("LIBRETTO", _documento(_foglio("COMPUTO"))), "Libretto delle Misure")


class TestTestataAschermo(unittest.TestCase):
    VALORI = {"[COMMITTENTE]": "Comune di Matera", "[PROGETTO]": "Lavori di prova",
              "[OGGETTO]": "Computo Metrico Estimativo", "[FILE]": "lavori.ods", "[SAL]": ""}

    def setUp(self):
        self.ns = _carica()
        self.testata = self.ns["testata_a_schermo"]
        self.pred = self.ns["testi_predefiniti"]

    def test_computo_riproduce_il_testo_di_prima(self):
        t = self.testata(self.pred("COMPUTO"), self.VALORI, "COMPUTO")
        self.assertEqual(t["intSx"], "\n\nCommittente: Comune di Matera")
        self.assertEqual(t["intCenter"], "Lavori di prova\n\n")
        self.assertEqual(t["intDx"], "\n\nComputo Metrico Estimativo")
        self.assertEqual(t["ppSx"], "\nrealizzato con LeenO: lavori.ods")
        self.assertEqual(t["ppCenter"], "")
        self.assertEqual(t["ppDx"], "Pagina [PAGINA] di [PAGINE]")

    def test_analisi_con_una_sola_riga_vuota_a_destra(self):
        valori = dict(self.VALORI, **{"[OGGETTO]": "Analisi di Prezzo"})
        t = self.testata(self.pred("ANALISI"), valori, "ANALISI")
        self.assertEqual(t["intDx"], "\nAnalisi di Prezzo")
        self.assertEqual(t["intSx"], "\n\nCommittente: Comune di Matera")

    def test_atti_contabili_senza_righe_vuote_per_non_alzare_intestazione_e_piede(self):
        valori = dict(self.VALORI, **{"[OGGETTO]": "Libretto delle Misure", "[SAL]": "3"})
        for chiave in ("LIBRETTO", "REGISTRO", "SAL"):
            t = self.testata(self.pred(chiave), valori, chiave)
            with self.subTest(chiave=chiave):
                self.assertEqual(t["intSx"], "Committente: Comune di Matera")
                self.assertEqual(t["intCenter"], "Lavori di prova")
                self.assertEqual(t["intDx"], "Libretto delle Misure n. 3")
                for parte in ("intSx", "intCenter", "intDx"):
                    self.assertNotIn("\n", t[parte], parte)

    def test_libretto_ha_le_firme_su_due_righe_per_parte(self):
        valori = dict(self.VALORI, **{"[OGGETTO]": "Libretto delle Misure", "[SAL]": "3"})
        t = self.testata(self.pred("LIBRETTO"), valori, "LIBRETTO")
        self.assertEqual(t["ppSx"], "realizzato con LeenO: lavori.ods\n")
        self.assertEqual(t["ppCenter"], "\n" + FIRME)
        self.assertEqual(t["ppDx"], "Pagina [PAGINA] di [PAGINE]\n")

    def test_registro_e_sal_senza_firme_e_con_il_piede_su_una_riga(self):
        valori = dict(self.VALORI, **{"[SAL]": "3"})
        for chiave in ("REGISTRO", "SAL"):
            t = self.testata(self.pred(chiave), valori, chiave)
            self.assertEqual((t["ppSx"], t["ppCenter"], t["ppDx"]),
                             ("realizzato con LeenO: lavori.ods", "", "Pagina [PAGINA] di [PAGINE]"), chiave)

    def test_i_testi_personalizzati_seguono_le_stesse_regole(self):
        testi = dict(self.pred("ELENCO"), intSx="Ente: [COMMITTENTE]", ppDx="Foglio [PAGINA]")
        t = self.testata(testi, self.VALORI, "ELENCO")
        self.assertEqual((t["intSx"], t["ppDx"]), ("\n\nEnte: Comune di Matera", "Foglio [PAGINA]"))

    def test_senza_dati_in_s2_i_campi_sono_vuoti_come_prima(self):
        valori = dict(self.VALORI, **{"[COMMITTENTE]": "", "[PROGETTO]": ""})
        t = self.testata(self.pred("COMPUTO"), valori, "COMPUTO")
        self.assertEqual((t["intSx"], t["intCenter"]), ("\n\nCommittente: ", "\n\n"))

    def test_con_le_firme_le_tre_parti_del_piede_hanno_lo_stesso_numero_di_righe(self):
        testi = dict(self.pred("COMPUTO"), ppCenter="[FIRME]")
        t = self.testata(testi, self.VALORI, "COMPUTO")
        righe = {len(t[k].split("\n")) for k in ("ppSx", "ppCenter", "ppDx")}
        self.assertEqual(righe, {3})
        self.assertEqual(t["ppCenter"].split("\n")[-1], FIRME)
        self.assertEqual(t["ppSx"].split("\n")[1], "realizzato con LeenO: lavori.ods")
        self.assertEqual(t["ppDx"].split("\n")[0], "Pagina [PAGINA] di [PAGINE]")
        self.assertNotIn("[FIRME]", "".join(t.values()))

    def test_le_firme_non_toccano_l_intestazione(self):
        senza = self.testata(self.pred("COMPUTO"), self.VALORI, "COMPUTO")
        con = self.testata(dict(self.pred("COMPUTO"), ppCenter="[FIRME]"), self.VALORI, "COMPUTO")
        for k in ("intSx", "intCenter", "intDx"):
            self.assertEqual(con[k], senza[k], k)

    def test_non_modifica_i_dizionari_ricevuti(self):
        testi = self.pred("COMPUTO")
        copia, valori = dict(testi), dict(self.VALORI)
        self.testata(testi, self.VALORI, "COMPUTO")
        self.assertEqual((testi, self.VALORI), (copia, valori))


class TestSorgenteDiSistemaPagine(unittest.TestCase):
    def test_usa_le_funzioni_condivise_e_non_ha_piu_i_testi_fissi(self):
        sorgente = PYLEENO.read_text(encoding="utf-8")
        inizio = sorgente.index("def MENU_sistema_pagine")
        fine = sorgente.index("\ndef ", inizio + 10)
        corpo = sorgente[inizio:fine]
        for atteso in ("LeenoPagina.contesto_stampa(", "LeenoPagina.applica_stile_pagina(",
                       "LeenoContab.ultimo_sal()"):
            self.assertIn(atteso, corpo)
        motore = (PYTHONPATH_DIR / "LeenoPagina.py").read_text(encoding="utf-8")
        for atteso in ("LS.documento_dello_stile(", "LS.impostazioni_documento(", "LS.valori_segnaposto(",
                       "LS.titolo_a_schermo(", "LS.testata_a_schermo(", "LS.scrivi_testata_stile(",
                       "LS.STILI_TESTATA", "LS.STILI_ATTI_CONTABILI",
                       "LS.set_page_margins(", "LS.set_page_borders("):
            self.assertIn(atteso, motore)
        for vecchio in ("realizzato con LeenO", "Committente: ", "getCellRangeByName(\"C3\")",
                        "Incidenza MdO", "Liberation Sans Narrow"):
            self.assertNotIn(vecchio, corpo)
            self.assertNotIn(vecchio, motore)


# --------------------------------------------------------------------------------------
# Documento reale
# --------------------------------------------------------------------------------------
def _riferimento_precedente(oDoc, foglio_attivo, nomefile):
    """Copia fedele del codice di MENU_sistema_pagine prima del 3c, per il foglio attivo dato."""
    oSheet = oDoc.Sheets.getByName(foglio_attivo)
    oggetto = oDoc.getSheets().getByName('S2').getCellRangeByName("C3").String + '\n\n'
    committente = "\n\nCommittente: " + oDoc.getSheets().getByName('S2').getCellRangeByName("C6").String
    luogo = '\n' + oSheet.Name
    if oSheet.Name == 'COMPUTO':
        if oSheet.getColumns().getByIndex(30).Columns.IsVisible == True:
            luogo = '\n\nIncidenza MdO'
        else:
            luogo = '\n\nComputo Metrico Estimativo'
    elif oSheet.Name == 'VARIANTE':
        if oSheet.getColumns().getByIndex(30).Columns.IsVisible == True:
            luogo = '\n\nIncidenza MdO (Variante)'
        else:
            luogo = '\n\nPerizia di Variante'
    elif oSheet.Name == 'Elenco Prezzi':
        if oSheet.getColumns().getByIndex(11).IsVisible == False or \
                oSheet.getColumns().getByIndex(12).IsVisible == False or \
                oSheet.getColumns().getByIndex(13).IsVisible == False:
            luogo = '\n\nQuadro di Raffronto'
        else:
            luogo = '\n\nElenco Prezzi'
    FONT = 'Liberation Sans Narrow'
    HTXT_DEFAULT = 8.0
    HTXT_ANALISI = 9.0
    STILI_CON_HEADER = frozenset(('PageStyle_Analisi di Prezzo', 'PageStyle_COMPUTO_A4', 'PageStyle_Elenco Prezzi'))
    oPageStyles = oDoc.StyleFamilies.getByName('PageStyles')
    for n in range(oPageStyles.Count):
        oAktPage = oPageStyles.getByIndex(n)
        if oAktPage.DisplayName in STILI_CON_HEADER:
            htxt = HTXT_ANALISI if oSheet.Name == 'Analisi di Prezzo' else HTXT_DEFAULT
            oHeader = oAktPage.RightPageHeaderContent
            oHeader.LeftText.Text.String = committente
            oHeader.LeftText.Text.Text.CharFontName = FONT
            oHeader.LeftText.Text.Text.CharHeight = htxt
            oHeader.CenterText.Text.String = oggetto
            oHeader.CenterText.Text.Text.CharFontName = FONT
            oHeader.CenterText.Text.Text.CharHeight = htxt
            oHeader.RightText.Text.String = luogo
            oHeader.RightText.Text.Text.CharFontName = FONT
            oHeader.RightText.Text.Text.CharHeight = htxt
            oAktPage.RightPageHeaderContent = oHeader
            oFooter = oAktPage.RightPageFooterContent
            oFooter.CenterText.Text.String = ''
            oFooter.LeftText.Text.String = "\nrealizzato con LeenO: " + nomefile
            oFooter.LeftText.Text.Text.CharFontName = FONT
            oFooter.LeftText.Text.Text.CharHeight = htxt * 0.5
            oFooter.RightText.Text.Text.CharFontName = FONT
            oFooter.RightText.Text.Text.CharHeight = htxt
            oAktPage.RightPageFooterContent = oFooter


def _paragrafi(oTesto):
    """[(testo, altezza, font)] dei paragrafi di una parte di intestazione o piè di pagina."""
    en = oTesto.createEnumeration()
    out = []
    while en.hasMoreElements():
        p = en.nextElement()
        out.append((p.String, p.CharHeight, p.CharFontName))
    return out


def _parti(oAktPage):
    """Paragrafi delle tre parti dell'intestazione e del piè di pagina dello stile."""
    intestazione = oAktPage.RightPageHeaderContent
    piede = oAktPage.RightPageFooterContent
    return {"intSx": _paragrafi(intestazione.LeftText), "intCenter": _paragrafi(intestazione.CenterText),
            "intDx": _paragrafi(intestazione.RightText), "ppSx": _paragrafi(piede.LeftText),
            "ppCenter": _paragrafi(piede.CenterText), "ppDx": _paragrafi(piede.RightText)}


def _disponibile():
    return uno_harness.uno_disponibile(ESEMPIO)


@unittest.skipUnless(_disponibile(), "richiede LEENO_TEST_UNO=1, uno, soffice e LeenO_Esempio.ods")
class TestTestataReale(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import uno
        cls.uno = uno
        cls.sessione = uno_harness.SessioneUno()
        cls.addClassCleanup(cls.sessione.chiudi)
        cls.sessione.avvia()
        utils = carica(PYTHONPATH_DIR / "LeenoUtils.py", ("replacePatternWithField",))
        sheet = carica(PYTHONPATH_DIR / "SheetUtils.py", ("paginationFields",),
                       ns={"LeenoUtils": SimpleNamespace(**utils)})
        cls.ns = _carica(SheetUtils=SimpleNamespace(paginationFields=sheet["paginationFields"]))
        cls.contatore = 0

    def _doc(self, **s2):
        type(self).contatore += 1
        doc = self.sessione.apri_copia(ESEMPIO, "testata%d" % self.contatore)
        sheet = doc.Sheets.getByName("S2")
        sheet.getCellRangeByName("C3").String = s2.get("c3", "Lavori di prova")
        sheet.getCellRangeByName("C6").String = s2.get("c6", "Comune di Matera")
        return doc

    def _applica(self, doc, foglio_attivo, ps=None, nomefile="lavori.ods", n_sal="5", atti=False):
        """
        Applica le nuove funzioni come fa MENU_sistema_pagine, per il foglio attivo dato.
        atti=True include gli stili degli atti contabili (Libretto, Registro, SAL).
        """
        ns = self.ns
        ps = ps or {}
        stili = doc.StyleFamilies.getByName("PageStyles")
        for n in range(stili.Count):
            pagina = stili.getByIndex(n)
            nome = pagina.DisplayName
            if nome not in ns["STILI_TESTATA"] and not (atti and nome in ns["STILI_ATTI_CONTABILI"]):
                continue
            chiave = ns["documento_dello_stile"](nome, foglio_attivo)
            testi = ns["impostazioni_documento"](ps, chiave)
            valori = self._valori(doc, ns["titolo_a_schermo"](chiave, doc), nomefile, n_sal)
            testata = ns["testata_a_schermo"](testi, valori, chiave)
            ns["scrivi_testata_stile"](doc, pagina, testata, ns["dimensione_testata"](chiave), ns["ha_firme"](testi))

    def _valori(self, doc, titolo, nomefile, n_sal=""):
        s2 = doc.Sheets.getByName("S2")
        return {"[COMMITTENTE]": s2.getCellRangeByName("C6").String, "[PROGETTO]": s2.getCellRangeByName("C3").String,
                "[OGGETTO]": titolo, "[FILE]": nomefile, "[SAL]": n_sal}

    def test_computo_come_il_vecchio_codice(self):
        for visibile in (False, True):  # False: "Computo Metrico Estimativo", True: "Incidenza MdO"
            with self.subTest(incidenza_visibile=visibile):
                self._confronta("COMPUTO", "PageStyle_COMPUTO_A4", colonna=("COMPUTO", 30, visibile))

    def test_analisi_di_prezzo_come_il_vecchio_codice(self):
        self._confronta("Analisi di Prezzo", "PageStyle_COMPUTO_A4")

    def test_elenco_prezzi_come_il_vecchio_codice(self):
        for visibile in (True, False):  # False: "Quadro di Raffronto"
            with self.subTest(colonna_visibile=visibile):
                self._confronta("Elenco Prezzi", "PageStyle_Elenco Prezzi", colonna=("Elenco Prezzi", 12, visibile))

    def _confronta(self, foglio_attivo, stile, colonna=None):
        vecchio, nuovo = self._doc(), self._doc()
        if colonna:
            for d in (vecchio, nuovo):
                d.Sheets.getByName(colonna[0]).getColumns().getByIndex(colonna[1]).IsVisible = colonna[2]
        _riferimento_precedente(vecchio, foglio_attivo, "lavori.ods")
        self._applica(nuovo, foglio_attivo)
        pv = _parti(vecchio.StyleFamilies.getByName("PageStyles").getByName(stile))
        pn = _parti(nuovo.StyleFamilies.getByName("PageStyles").getByName(stile))
        for parte in ("intSx", "intCenter", "intDx", "ppSx", "ppCenter"):
            self.assertEqual(pn[parte], pv[parte], "%s (%s)" % (parte, foglio_attivo))
        # il piè di pagina di destra prima restava il testo del template; ora e' "Pagina N di M"
        self.assertEqual(len(pn["ppDx"]), 1)

    def test_gli_atti_contabili_ricevono_la_testata_del_loro_documento(self):
        doc = self._doc()
        self._applica(doc, "COMPUTO", atti=True, n_sal="5")
        stili = doc.StyleFamilies.getByName("PageStyles")
        attesi = (("Page_Style_Libretto_Misure2", "Libretto delle Misure", True),
                  ("PageStyle_REGISTRO_A4", "Registro di Contabilità", False),
                  ("PageStyle_SAL_A4", "Stato di Avanzamento Lavori", False))
        for stile, titolo, firme in attesi:
            with self.subTest(stile=stile):
                pagina = stili.getByName(stile)
                parti = _parti(pagina)
                self.assertEqual([p[0] for p in parti["intSx"]], ["Committente: Comune di Matera"])
                self.assertEqual([p[0] for p in parti["intCenter"]], ["Lavori di prova"])
                self.assertEqual([p[0] for p in parti["intDx"]], [titolo + " n. 5"])
                self.assertEqual([p[0] for p in parti["ppCenter"]], ["", FIRME] if firme else [""])
                self.assertEqual(len(parti["ppSx"]), 2 if firme else 1)
                self.assertEqual(len(parti["ppDx"]), 2 if firme else 1)

    def test_senza_atti_gli_stili_contabili_restano_come_sono(self):
        doc = self._doc()
        prima = _parti(doc.StyleFamilies.getByName("PageStyles").getByName("Page_Style_Libretto_Misure2"))
        self._applica(doc, "COMPUTO")
        dopo = _parti(doc.StyleFamilies.getByName("PageStyles").getByName("Page_Style_Libretto_Misure2"))
        self.assertEqual(dopo, prima)

    def test_la_paginazione_degli_atti_gia_emessi_non_peggiora(self):
        """
        Le interruzioni di pagina di ogni area di SAL (Libretto, Registro, SAL) non devono
        anticipare: una pagina che contiene meno righe di prima sfalserebbe i riempimenti.
        """
        prima, dopo = self._doc(), self._doc()
        self._applica(dopo, "COMPUTO", atti=True, n_sal="5")
        aree = (("CONTABILITA", "_Lib_", 11), ("Registro", "_Reg_", 9), ("SAL", "_SAL_", 5))
        controllate = 0
        for foglio, prefisso, colonna in aree:
            nomi = [n for n in prima.NamedRanges.ElementNames if n.startswith(prefisso)]
            for nome in nomi:
                interruzioni = []
                for doc in (prima, dopo):
                    area = doc.NamedRanges.getByName(nome).ReferredCells.RangeAddress
                    area.EndColumn = colonna
                    sheet = doc.Sheets.getByName(foglio)
                    sheet.PrintAreas = (area,)
                    sheet.getCellRangeByPosition(area.StartColumn, area.StartRow, area.EndColumn,
                                                 area.EndRow).Rows.IsVisible = True
                    interruzioni.append([b.Position for b in sheet.RowPageBreaks
                                         if area.StartRow <= b.Position <= area.EndRow + 1])
                a, d = interruzioni
                with self.subTest(area=nome):
                    self.assertEqual(len(d), len(a), "numero di pagine")
                    for prima_pos, dopo_pos in zip(a, d):
                        self.assertGreaterEqual(dopo_pos, prima_pos, "pagina con meno righe di prima")
                controllate += 1
        self.assertGreaterEqual(controllate, 10)

    def test_il_numero_di_pagina_diventa_un_campo_nello_stile(self):
        doc = self._doc()
        self._applica(doc, "COMPUTO")
        pagina = doc.StyleFamilies.getByName("PageStyles").getByName("PageStyle_COMPUTO_A4")
        piede = pagina.RightPageFooterContent
        testo = piede.RightText.String
        self.assertRegex(testo, r"Pagina \d+ di \d+")
        self.assertNotIn("[PAGINA]", testo)

    def test_i_testi_personalizzati_arrivano_nello_stile(self):
        doc = self._doc()
        ps = {"documenti": json.dumps({"COMPUTO": {"intSx": "Ente: [COMMITTENTE]", "ppSx": "File [FILE]"}})}
        self._applica(doc, "COMPUTO", ps)
        pagina = doc.StyleFamilies.getByName("PageStyles").getByName("PageStyle_COMPUTO_A4")
        # il contenuto va tenuto in una variabile: vedi LESSONS_UNO_QUIRKS
        intestazione = pagina.RightPageHeaderContent
        piede = pagina.RightPageFooterContent
        # String trasforma gli "a capo" in spazi: si leggono i paragrafi
        self.assertEqual([p[0] for p in _paragrafi(intestazione.LeftText)], ["", "", "Ente: Comune di Matera"])
        self.assertEqual([p[0] for p in _paragrafi(piede.LeftText)], ["", "File lavori.ods"])

    def test_con_le_firme_il_piede_ha_tre_righe_per_parte_e_le_firme_sono_l_ultima(self):
        doc = self._doc()
        ps = {"documenti": json.dumps({"COMPUTO": {"ppCenter": "[FIRME]"}})}
        self._applica(doc, "COMPUTO", ps)
        parti = _parti(doc.StyleFamilies.getByName("PageStyles").getByName("PageStyle_COMPUTO_A4"))
        for nome in ("ppSx", "ppCenter", "ppDx"):
            self.assertEqual(len(parti[nome]), 3, nome)
        self.assertEqual(parti["ppCenter"][-1][0], FIRME)
        self.assertEqual([p[0] for p in parti["ppSx"]], ["", "realizzato con LeenO: lavori.ods", ""])

    @staticmethod
    def _parole(pdf):
        out = subprocess.run(["pdftotext", "-bbox", "-f", "1", "-l", "1", pdf, "-"],
                             capture_output=True, text=True).stdout
        res = {}
        for _x1, y1, _x2, y2, parola in re.findall(
                r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]+)</word>', out):
            res.setdefault(parola, []).append((float(y1), float(y2)))
        return res

    def _pdf(self, doc, nome):
        uno = self.uno
        foglio = doc.Sheets.getByName("COMPUTO")
        intervallo = foglio.getCellRangeByPosition(0, 0, 10, 60)
        filtro = uno.Any("[]com.sun.star.beans.PropertyValue",
                         (uno.createUnoStruct("com.sun.star.beans.PropertyValue", "Selection", 0, intervallo, 0),))
        props = (uno.createUnoStruct("com.sun.star.beans.PropertyValue", "FilterName", 0, "calc_pdf_Export", 0),
                 uno.createUnoStruct("com.sun.star.beans.PropertyValue", "FilterData", 0, filtro, 0))
        percorso = str(Path(self.sessione.tmp) / (nome + ".pdf"))
        doc.storeToURL(uno.systemPathToFileUrl(percorso), props)
        return self._parole(percorso)

    @unittest.skipUnless(shutil.which("pdftotext") is not None, "richiede pdftotext")
    def test_posizione_e_dimensione_come_il_vecchio_codice(self):
        vecchio, nuovo = self._doc(), self._doc()
        _riferimento_precedente(vecchio, "COMPUTO", "lavori.ods")
        self._applica(nuovo, "COMPUTO")
        pv, pn = self._pdf(vecchio, "vecchio"), self._pdf(nuovo, "nuovo")
        for parola in ("Committente:", "realizzato", "Lavori", "Matera"):
            self.assertIn(parola, pn, parola)
            self.assertAlmostEqual(pn[parola][0][0], pv[parola][0][0], delta=0.2, msg=parola)
            self.assertAlmostEqual(pn[parola][0][1], pv[parola][0][1], delta=0.2, msg=parola)

    @unittest.skipUnless(shutil.which("pdftotext") is not None, "richiede pdftotext")
    def test_dimensioni_a_schermo_senza_firme(self):
        doc = self._doc()
        self._applica(doc, "COMPUTO")
        parole = self._pdf(doc, "senza")
        altezza = lambda w: parole[w][0][1] - parole[w][0][0]
        self.assertAlmostEqual(altezza("Committente:"), altezza("Pagina"), delta=0.1)  # 8 pt
        self.assertAlmostEqual(altezza("realizzato"), altezza("Pagina") / 2, delta=0.1)  # 4 pt: meta'

    @unittest.skipUnless(shutil.which("pdftotext") is not None, "richiede pdftotext")
    def test_con_le_firme_il_piede_ha_una_sola_dimensione_e_le_firme_sono_sotto(self):
        doc = self._doc()
        ps = {"documenti": json.dumps({"COMPUTO": {"ppCenter": "[FIRME]"}})}
        self._applica(doc, "COMPUTO", ps)
        parole = self._pdf(doc, "firme")
        altezza = lambda w: parole[w][0][1] - parole[w][0][0]
        self.assertAlmostEqual(altezza("realizzato"), altezza("Pagina"), delta=0.1)
        self.assertAlmostEqual(altezza("DIRETTORE"), altezza("Pagina"), delta=0.1)
        # le firme stanno una riga sotto "realizzato con" e sotto "Pagina N di M"
        for riferimento in ("realizzato", "Pagina"):
            self.assertGreaterEqual(parole["DIRETTORE"][0][0] - parole[riferimento][0][0], 0.8 * altezza("Pagina"))


if __name__ == "__main__":
    unittest.main()
