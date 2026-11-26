# -*- coding: utf-8 -*-
"""
Test dei segnaposto di intestazione e piè di pagina (LeenoSettings).

- unitari, senza UNO: S2 prevale sulle impostazioni di lavoro, ripiego, segnaposto
  calcolati ([OGGETTO] = titolo del documento, [SAL], [FILE]), sostituzione in un solo passaggio,
  equivalenza con la sostituzione storica di prepareHeaderFooter;
- documento reale (opzionale, LEENO_TEST_UNO=1): valori letti dal vero S2 di
  tests/data/LeenO_Esempio.ods e testo delle firme uguale a quello del template.
"""
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

try:
    from tests import uno_harness
    from tests.estrai_sorgente import FoglioFinto, carica
except ImportError:  # esecuzione con la cartella tests nel percorso di ricerca
    import uno_harness
    from estrai_sorgente import FoglioFinto, carica

ROOT = Path(__file__).resolve().parent.parent
PYTHONPATH_DIR = ROOT / "src/Ultimus.oxt/python/pythonpath"
ESEMPIO = ROOT / "tests/data/LeenO_Esempio.ods"

FUNZIONI = ("nome_file_documento", "valori_segnaposto", "sostituisci_segnaposto")
COSTANTI = ("_SEGNAPOSTI_S2", "_SEGNAPOSTI_LAVORO", "_FIRME_TESTO")
FIRME_TEMPLATE = "L'IMPRESA\t\t\t\t\t\tIL DIRETTORE DEI LAVORI    "


def _carica_impostazioni():
    """Funzioni dei segnaposto con il vero lettore di S2 (compilato da SheetUtils)."""
    ns_lettore = {"LeenoUtils": SimpleNamespace(log_errore=MagicMock())}
    carica(PYTHONPATH_DIR / "SheetUtils.py", ("leggi_dato_etichettato",), ns=ns_lettore)
    import os
    import re
    ns = {
        "os": os,
        "re": re,
        "SheetUtils": SimpleNamespace(leggi_dato_etichettato=ns_lettore["leggi_dato_etichettato"]),
        "loadJobSettings": lambda _doc: {},
    }
    return carica(PYTHONPATH_DIR / "LeenoSettings.py", FUNZIONI, COSTANTI, ns)


def _documento_finto(celle_s2):
    s2 = FoglioFinto(celle_s2)
    return SimpleNamespace(
        Sheets=SimpleNamespace(getByName=lambda nome: s2 if nome == "S2" else None),
        getURL=lambda: "")


def _sostituisci_storico(testo, valori):
    """Sostituzione di prepareHeaderFooter prima dell'introduzione di sostituisci_segnaposto."""
    for chiave, valore in valori.items():
        while chiave in testo:
            testo = testo.replace(chiave, valore)
    return testo


S2_ESEMPIO = {
    (1, 2): "OGGETTO:", (2, 2): "Lavori di prova",
    (1, 5): "Stazione Appaltante \u2013 Committente:", (2, 5): "Comune di Matera",
    (1, 8): "Progettista:", (2, 8): "arch. Rossi",
    (1, 9): "Direttore Lavori:", (2, 9): "ing. Bianchi",
    (1, 10): "Responsabile del Progetto:", (2, 10): "geom. Verdi",
}


class TestValoriSegnaposto(unittest.TestCase):
    def setUp(self):
        self.ns = _carica_impostazioni()
        self.valori = self.ns["valori_segnaposto"]
        self.doc = _documento_finto(S2_ESEMPIO)

    def test_i_dati_di_s2_vengono_letti_per_etichetta(self):
        v = self.valori(self.doc, lavoro={})
        self.assertEqual(v["[PROGETTO]"], "Lavori di prova")
        self.assertEqual(v["[COMMITTENTE]"], "Comune di Matera")
        self.assertEqual(v["[STAZIONE_APPALTANTE]"], "Comune di Matera")
        self.assertEqual(v["[PROGETTISTA]"], "arch. Rossi")
        self.assertEqual(v["[DIRETTORE_LAVORI]"], "ing. Bianchi")
        self.assertEqual(v["[RUP]"], "geom. Verdi")

    def test_s2_prevale_sulle_impostazioni_di_lavoro(self):
        lavoro = {"committente": "Altro Ente", "progettista": "Altro Progettista"}
        v = self.valori(self.doc, lavoro=lavoro)
        self.assertEqual(v["[COMMITTENTE]"], "Comune di Matera")
        self.assertEqual(v["[PROGETTISTA]"], "arch. Rossi")

    def test_ripiego_sulle_impostazioni_di_lavoro_se_s2_e_vuoto(self):
        doc = _documento_finto({(1, 2): "OGGETTO:", (1, 5): "Committente:"})
        v = self.valori(doc, lavoro={"committente": "Altro Ente", "rup": "dott. Neri"})
        self.assertEqual(v["[COMMITTENTE]"], "Altro Ente")
        self.assertEqual(v["[RUP]"], "dott. Neri")

    def test_progetto_da_s2_con_ripiego_su_lavoro(self):
        vuoto = _documento_finto({(1, 2): "OGGETTO:"})
        self.assertEqual(self.valori(vuoto, lavoro={"progetto": "Progetto X"})["[PROGETTO]"], "Progetto X")
        self.assertEqual(self.valori(self.doc, lavoro={"progetto": "Progetto X"})["[PROGETTO]"], "Lavori di prova")

    def test_oggetto_e_il_titolo_del_documento_e_non_legge_s2(self):
        self.assertEqual(self.valori(self.doc, lavoro={})["[OGGETTO]"], "")
        v = self.valori(self.doc, documento="Computo Metrico", lavoro={})
        self.assertEqual(v["[OGGETTO]"], "Computo Metrico")
        self.assertNotIn("[DOCUMENTO]", v)

    def test_etichetta_alternativa_del_direttore_dei_lavori(self):
        doc = _documento_finto({(1, 4): "Direttore dei Lavori:", (2, 4): "ing. Gialli"})
        self.assertEqual(self.valori(doc, lavoro={})["[DIRETTORE_LAVORI]"], "ing. Gialli")

    def test_segnaposto_solo_da_impostazioni_di_lavoro(self):
        lavoro = {"data": "01/02/2026", "revisione": 3, "dataRevisione": None}
        v = self.valori(self.doc, lavoro=lavoro)
        self.assertEqual((v["[DATA]"], v["[REVISIONE]"], v["[DATA_REVISIONE]"]), ("01/02/2026", "3", ""))

    def test_documento_sal_e_file(self):
        v = self.valori(self.doc, documento="Libretto delle Misure", nSal=2, nomefile="lavori.ods", lavoro={})
        self.assertEqual(v["[OGGETTO]"], "Libretto delle Misure")
        self.assertEqual(v["[SAL]"], "2")
        self.assertEqual(v["[FILE]"], "lavori.ods")

    def test_firme_non_ha_un_valore_ma_e_gestito_da_piede_con_firme(self):
        # [FIRME] non si sostituisce con un testo: LeenoSettings.piede_con_firme lo trasforma
        # nella riga delle firme da sola in fondo al piè di pagina
        self.assertNotIn("[FIRME]", self.valori(self.doc, lavoro={}))

    def test_sal_assente_e_vuoto(self):
        self.assertEqual(self.valori(self.doc, lavoro={})["[SAL]"], "")

    def test_extra_prevale_sui_valori_calcolati(self):
        v = self.valori(self.doc, documento="A", extra={"[OGGETTO]": "B", "[NUMERO_DOCUMENTO]": "5"}, lavoro={})
        self.assertEqual((v["[OGGETTO]"], v["[NUMERO_DOCUMENTO]"]), ("B", "5"))

    def test_pagina_e_pagine_non_sono_valori(self):
        v = self.valori(self.doc, lavoro={})
        self.assertNotIn("[PAGINA]", v)
        self.assertNotIn("[PAGINE]", v)

    def test_insieme_dei_segnaposto(self):
        atteso = {"[COMMITTENTE]", "[STAZIONE_APPALTANTE]", "[OGGETTO]", "[RUP]", "[PROGETTISTA]",
                  "[DIRETTORE_LAVORI]", "[PROGETTO]", "[DATA]", "[REVISIONE]", "[DATA_REVISIONE]",
                  "[SAL]", "[FILE]"}
        self.assertEqual(set(self.valori(self.doc, lavoro={})), atteso)

    def test_lavoro_caricato_da_loadjobsettings_se_non_passato(self):
        chiamate = []
        self.ns["loadJobSettings"] = lambda doc: chiamate.append(doc) or {"rup": "dott. Neri"}
        doc = _documento_finto({})
        self.assertEqual(self.valori(doc)["[RUP]"], "dott. Neri")
        self.assertEqual(chiamate, [doc])


class TestSostituisci(unittest.TestCase):
    def setUp(self):
        self.ns = _carica_impostazioni()
        self.sostituisci = self.ns["sostituisci_segnaposto"]

    def test_sostituisce_piu_segnaposto(self):
        valori = {"[COMMITTENTE]": "Comune", "[SAL]": "2"}
        self.assertEqual(self.sostituisci("Committente: [COMMITTENTE], n. [SAL] ([SAL])", valori),
                         "Committente: Comune, n. 2 (2)")

    def test_pagina_pagine_e_sconosciuti_restano(self):
        testo = "Pagina [PAGINA] di [PAGINE] [IGNOTO]"
        self.assertEqual(self.sostituisci(testo, {"[SAL]": "2"}), testo)

    def test_un_solo_passaggio_senza_ricorsione(self):
        # il valore contiene un segnaposto: non va espanso ne' deve bloccarsi
        valori = {"[OGGETTO]": "lavori [OGGETTO]", "[SAL]": "[OGGETTO]"}
        self.assertEqual(self.sostituisci("[OGGETTO] / [SAL]", valori), "lavori [OGGETTO] / [OGGETTO]")

    def test_firme_resta_invariato_per_piede_con_firme(self):
        self.assertEqual(self.sostituisci("a [FIRME] [SAL]", {"[SAL]": "2"}), "a [FIRME] 2")

    def test_valore_vuoto_cancella_il_segnaposto(self):
        self.assertEqual(self.sostituisci("a[SAL]b", {"[SAL]": ""}), "ab")

    def test_testo_vuoto_e_firme_con_tabulazioni(self):
        self.assertEqual(self.sostituisci("", {"[SAL]": "2"}), "")
        self.assertEqual(self.sostituisci("[TESTO]", {"[TESTO]": FIRME_TEMPLATE}), FIRME_TEMPLATE)

    def test_equivale_alla_sostituzione_storica(self):
        valori = {"[COMMITTENTE]": "Comune di Matera", "[PROGETTO]": "Progetto X", "[OGGETTO]": "Libretto",
                  "[DATA]": "", "[RUP]": "geom. Verdi"}
        for testo in ("[COMMITTENTE]", "Pagina [PAGINA] di [PAGINE]", "[OGGETTO] - [PROGETTO] - [DATA]",
                      "[RUP][RUP]", "senza segnaposto", "", "[COMMITTENTE]\n[PROGETTO]"):
            with self.subTest(testo=testo):
                self.assertEqual(self.sostituisci(testo, valori), _sostituisci_storico(testo, valori))


class TestNomeFile(unittest.TestCase):
    def setUp(self):
        self.nome = _carica_impostazioni()["nome_file_documento"]

    def test_nome_senza_percorso_e_spazi_decodificati(self):
        doc = SimpleNamespace(getURL=lambda: "file:///W:/lavori/Mio%20Computo.ods")
        self.assertEqual(self.nome(doc), "Mio Computo.ods")

    def test_documento_non_salvato(self):
        self.assertEqual(self.nome(SimpleNamespace(getURL=lambda: "")), "")


@unittest.skipUnless(uno_harness.uno_disponibile(ESEMPIO),
                     "richiede LEENO_TEST_UNO=1, uno, soffice e LeenO_Esempio.ods")
class TestDocumentoReale(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.sessione = uno_harness.SessioneUno()
        cls.addClassCleanup(cls.sessione.chiudi)
        cls.sessione.avvia()
        cls.doc = cls.sessione.apri_copia(ESEMPIO, "segnaposti")
        cls.s2 = cls.doc.Sheets.getByName("S2")
        cls.s2.getCellRangeByName("C3").String = "Lavori di prova"
        cls.ns = _carica_impostazioni()

    def _valore_riga(self, inizio):
        """Valore in C della prima riga con colonna B che inizia per `inizio` (lettura indipendente)."""
        for r in range(0, 80):
            if self.s2.getCellByPosition(1, r).String.strip().lower().startswith(inizio.lower()):
                val = self.s2.getCellByPosition(2, r).String.strip()
                if val:
                    return val
        return ""

    def test_valori_dal_vero_s2(self):
        v = self.ns["valori_segnaposto"](self.doc, "Libretto delle Misure", 2, "x.ods", lavoro={})
        self.assertEqual(v["[PROGETTO]"], "Lavori di prova")
        self.assertEqual(v["[OGGETTO]"], "Libretto delle Misure")
        for segnaposto, inizio in (("[PROGETTISTA]", "Progettista"), ("[DIRETTORE_LAVORI]", "Direttore Lavori"),
                                   ("[RUP]", "Responsabile del Progetto")):
            with self.subTest(segnaposto=segnaposto):
                atteso = self._valore_riga(inizio)
                self.assertNotEqual(atteso, "", inizio)
                self.assertEqual(v[segnaposto], atteso)
        self.assertEqual(v["[SAL]"], "2")

    def test_testo_delle_firme_uguale_al_template(self):
        stili = self.doc.StyleFamilies.getByName("PageStyles")
        # il contenuto del piè di pagina va tenuto in una variabile: in una catena di attributi
        # UNO rilascia l'oggetto temporaneo prima di leggere CenterText
        contenuto = stili.getByName("Page_Style_Libretto_Misure2").RightPageFooterContent
        self.assertEqual(self.ns["_FIRME_TESTO"], contenuto.CenterText.String)

    def test_intestazione_completa_con_il_dato_reale(self):
        valori = self.ns["valori_segnaposto"](self.doc, "Computo Metrico", lavoro={})
        testo = self.ns["sostituisci_segnaposto"]("[PROGETTO] ([OGGETTO]) Pagina [PAGINA]", valori)
        self.assertEqual(testo, "Lavori di prova (Computo Metrico) Pagina [PAGINA]")


if __name__ == "__main__":
    unittest.main()
