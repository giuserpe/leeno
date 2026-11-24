# -*- coding: utf-8 -*-
"""
Test delle impostazioni di intestazione e piè di pagina per documento (LeenoSettings):
predefiniti, personalizzazioni salvate come JSON, compatibilita' con le sei stringhe
comuni della versione precedente, salvataggio delle sole differenze.
"""
import json
import unittest
from pathlib import Path

try:
    from tests.estrai_sorgente import carica
except ImportError:  # esecuzione con la cartella tests nel percorso di ricerca
    from estrai_sorgente import carica

ROOT = Path(__file__).resolve().parent.parent
SETTINGS = ROOT / "src/Ultimus.oxt/python/pythonpath/LeenoSettings.py"
CONFIG = ROOT / "src/Ultimus.oxt/python/pythonpath/LeenoConfig.py"

FUNZIONI = ("testi_predefiniti", "_testi_personalizzati", "impostazioni_documento",
            "aggiorna_impostazioni_documenti")
COSTANTI = ("_DOCUMENTI_STAMPA", "_DOCUMENTI_SENZA_TESTATA", "_DOCUMENTI_CON_FIRME",
            "_TESTI_STAMPA", "_TESTI_PRECEDENTI", "_PRINTSETTINGSITEMS")

TESTI = ("intSx", "intCenter", "intDx", "ppSx", "ppCenter", "ppDx")


def _carica():
    return carica(SETTINGS, FUNZIONI, COSTANTI, {"json": json})


class TestPredefiniti(unittest.TestCase):
    def setUp(self):
        self.ns = _carica()
        self.chiavi = [c for c, _ in self.ns["_DOCUMENTI_STAMPA"]]

    def test_elenco_dei_documenti(self):
        self.assertEqual(self.ns["_DOCUMENTI_STAMPA"], (
            ("COMPUTO", "Computo Metrico"), ("VARIANTE", "Variante"), ("ELENCO", "Elenco Prezzi"),
            ("ANALISI", "Analisi di Prezzo"), ("LIBRETTO", "Libretto delle Misure"),
            ("REGISTRO", "Registro di Contabilità"), ("SAL", "Stato di Avanzamento Lavori"),
            ("CDP", "Certificato di Pagamento")))

    def test_ogni_documento_ha_le_sei_stringhe(self):
        for chiave in self.chiavi:
            with self.subTest(chiave=chiave):
                self.assertEqual(set(self.ns["testi_predefiniti"](chiave)), set(TESTI))

    def test_documenti_ordinari_come_la_stampa_a_schermo(self):
        for chiave in ("COMPUTO", "VARIANTE", "ELENCO", "ANALISI"):
            with self.subTest(chiave=chiave):
                self.assertEqual(self.ns["testi_predefiniti"](chiave), {
                    "intSx": "Committente: [COMMITTENTE]", "intCenter": "[PROGETTO]", "intDx": "[OGGETTO]",
                    "ppSx": "realizzato con LeenO: [FILE]", "ppCenter": "",
                    "ppDx": "Pagina [PAGINA] di [PAGINE]"})

    def test_atti_contabili_riportano_il_numero_del_sal(self):
        for chiave in ("LIBRETTO", "REGISTRO", "SAL"):
            with self.subTest(chiave=chiave):
                self.assertEqual(self.ns["testi_predefiniti"](chiave)["intDx"], "[OGGETTO] n. [SAL]")

    def test_firme_solo_nel_libretto_per_impostazione_predefinita(self):
        self.assertEqual(self.ns["testi_predefiniti"]("LIBRETTO")["ppCenter"], "[FIRME]")
        self.assertEqual(self.ns["testi_predefiniti"]("REGISTRO")["ppCenter"], "")
        self.assertEqual(self.ns["_DOCUMENTI_CON_FIRME"], ("LIBRETTO", "REGISTRO"))

    def test_certificato_di_pagamento_senza_testata(self):
        self.assertEqual(self.ns["testi_predefiniti"]("CDP"), dict.fromkeys(TESTI, ""))
        self.assertEqual(self.ns["_DOCUMENTI_SENZA_TESTATA"], ("CDP",))

    def test_copia_indipendente(self):
        testi = self.ns["testi_predefiniti"]("COMPUTO")
        testi["intSx"] = "modificato"
        self.assertEqual(self.ns["testi_predefiniti"]("COMPUTO")["intSx"], "Committente: [COMMITTENTE]")

    def test_documenti_e_stringhe_comuni_fanno_parte_delle_impostazioni(self):
        self.assertIn("documenti", self.ns["_PRINTSETTINGSITEMS"])
        self.assertEqual(set(self.ns["_TESTI_PRECEDENTI"]), set(TESTI))


class TestLetturaImpostazioni(unittest.TestCase):
    def setUp(self):
        self.ns = _carica()
        self.leggi = self.ns["impostazioni_documento"]
        self.pred = self.ns["testi_predefiniti"]

    def test_senza_dati_salvati_valgono_i_predefiniti(self):
        for chiave, _ in self.ns["_DOCUMENTI_STAMPA"]:
            self.assertEqual(self.leggi({}, chiave), self.pred(chiave))

    def test_personalizzazione_parziale_del_documento(self):
        ps = {"documenti": json.dumps({"LIBRETTO": {"intSx": "Ente: [COMMITTENTE]"}})}
        testi = self.leggi(ps, "LIBRETTO")
        self.assertEqual(testi["intSx"], "Ente: [COMMITTENTE]")
        self.assertEqual(testi["ppCenter"], "[FIRME]")
        self.assertEqual(self.leggi(ps, "COMPUTO"), self.pred("COMPUTO"))

    def test_dati_non_validi_sono_ignorati(self):
        for dati in ("non json", "[1, 2]", "null", '{"LIBRETTO": 5}', '{"LIBRETTO": {"intSx": 7, "xx": "a"}}', None, ""):
            with self.subTest(dati=dati):
                self.assertEqual(self.leggi({"documenti": dati}, "LIBRETTO"), self.pred("LIBRETTO"))

    def test_stringhe_comuni_modificate_valgono_per_tutti_i_documenti(self):
        ps = {"intSx": "Mio ente", "ppDx": "Pag. [PAGINA]"}
        for chiave in ("COMPUTO", "LIBRETTO", "SAL"):
            with self.subTest(chiave=chiave):
                testi = self.leggi(ps, chiave)
                self.assertEqual((testi["intSx"], testi["ppDx"]), ("Mio ente", "Pag. [PAGINA]"))
                self.assertEqual(testi["intCenter"], self.pred(chiave)["intCenter"])

    def test_stringhe_comuni_ai_vecchi_valori_non_cambiano_i_predefiniti(self):
        ps = dict(self.ns["_TESTI_PRECEDENTI"])
        for chiave, _ in self.ns["_DOCUMENTI_STAMPA"]:
            self.assertEqual(self.leggi(ps, chiave), self.pred(chiave))

    def test_personalizzazione_del_documento_prevale_sulle_stringhe_comuni(self):
        ps = {"intSx": "Mio ente", "documenti": json.dumps({"COMPUTO": {"intSx": "Solo computo"}})}
        self.assertEqual(self.leggi(ps, "COMPUTO")["intSx"], "Solo computo")
        self.assertEqual(self.leggi(ps, "VARIANTE")["intSx"], "Mio ente")

    def test_cdp_resta_vuoto(self):
        ps = {"intSx": "Mio ente", "documenti": json.dumps({"CDP": {"intSx": "x"}})}
        self.assertEqual(self.leggi(ps, "CDP"), dict.fromkeys(TESTI, ""))


class TestSalvataggio(unittest.TestCase):
    def setUp(self):
        self.ns = _carica()
        self.leggi = self.ns["impostazioni_documento"]
        self.pred = self.ns["testi_predefiniti"]
        self.aggiorna = self.ns["aggiorna_impostazioni_documenti"]
        self.chiavi = [c for c, _ in self.ns["_DOCUMENTI_STAMPA"]]

    def _tutti(self, ps):
        return {c: self.leggi(ps, c) for c in self.chiavi}

    def test_salva_solo_le_differenze(self):
        valori = self._tutti({})
        valori["LIBRETTO"]["intSx"] = "Ente"
        ps = self.aggiorna({"fileCopertine": "a.ods"}, valori)
        self.assertEqual(json.loads(ps["documenti"]), {"LIBRETTO": {"intSx": "Ente"}})
        self.assertEqual(ps["fileCopertine"], "a.ods")

    def test_documento_riportato_ai_predefiniti_non_resta_salvato(self):
        ps = {"documenti": json.dumps({"SAL": {"intCenter": "x"}})}
        valori = self._tutti(ps)
        valori["SAL"] = self.pred("SAL")
        self.assertEqual(json.loads(self.aggiorna(ps, valori)["documenti"]), {})

    def test_non_modifica_il_dizionario_ricevuto(self):
        ps = {"intSx": "Mio ente"}
        copia = dict(ps)
        self.aggiorna(ps, self._tutti(ps))
        self.assertEqual(ps, copia)

    def test_andata_e_ritorno_con_stringhe_comuni_della_versione_precedente(self):
        ps = {"intSx": "Mio ente", "ppDx": "Pag. [PAGINA]", "documenti": "{}"}
        prima = self._tutti(ps)
        salvato = self.aggiorna(ps, prima)
        self.assertEqual(self._tutti(salvato), prima)
        for k, v in self.ns["_TESTI_PRECEDENTI"].items():
            self.assertEqual(salvato[k], v)

    def test_dopo_il_salvataggio_ripristinare_un_documento_vale_i_predefiniti(self):
        ps = {"intSx": "Mio ente"}
        valori = self._tutti(ps)
        valori["COMPUTO"] = self.pred("COMPUTO")
        salvato = self.aggiorna(ps, valori)
        self.assertEqual(self.leggi(salvato, "COMPUTO"), self.pred("COMPUTO"))
        self.assertEqual(self.leggi(salvato, "VARIANTE")["intSx"], "Mio ente")

    def test_cdp_non_viene_salvato(self):
        valori = self._tutti({})
        valori["CDP"]["intSx"] = "x"
        self.assertEqual(json.loads(self.aggiorna({}, valori)["documenti"]), {})

    def test_testi_con_caratteri_speciali_resistono(self):
        valori = self._tutti({})
        valori["ELENCO"]["ppSx"] = 'Così: "[FILE]" \\ à ù\t[SAL]'
        salvato = self.aggiorna({}, valori)
        self.assertEqual(self.leggi(salvato, "ELENCO")["ppSx"], valori["ELENCO"]["ppSx"])

    def test_json_stabile(self):
        valori = self._tutti({})
        valori["SAL"]["intSx"] = "a"
        valori["COMPUTO"]["intSx"] = "b"
        self.assertEqual(self.aggiorna({}, valori)["documenti"], self.aggiorna({}, valori)["documenti"])
        self.assertEqual(list(json.loads(self.aggiorna({}, valori)["documenti"])), ["COMPUTO", "SAL"])


class TestConfigPredefinita(unittest.TestCase):
    def test_chiave_documenti_presente_nei_default_di_config(self):
        self.assertIn("('ImpostazioniStampa', 'documenti', '(str){}')", CONFIG.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
