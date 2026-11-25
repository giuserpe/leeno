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

FUNZIONI = ("testi_predefiniti", "firme_predefinite", "_personalizzazioni", "_testi_effettivi",
            "impostazioni_documento", "firme_documento", "piede_con_firme",
            "aggiorna_impostazioni_documenti")
COSTANTI = ("_DOCUMENTI_STAMPA", "_DOCUMENTI_SENZA_TESTATA", "_DOCUMENTI_CON_FIRME",
            "_TESTI_STAMPA", "_TESTI_PRECEDENTI", "_PRINTSETTINGSITEMS", "_FIRME_TESTO")

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
        self.assertEqual([c for c in self.chiavi if self.ns["firme_predefinite"](c)], ["LIBRETTO"])
        self.assertEqual(self.ns["_DOCUMENTI_CON_FIRME"], ("LIBRETTO", "REGISTRO"))

    def test_le_firme_non_sono_un_testo_del_piede(self):
        for chiave in self.chiavi:
            self.assertNotIn("FIRME", " ".join(self.ns["testi_predefiniti"](chiave).values()))

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
        self.assertEqual(testi["ppCenter"], "")
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


class TestFirme(unittest.TestCase):
    def setUp(self):
        self.ns = _carica()
        self.firme = self.ns["firme_documento"]

    def test_predefinite(self):
        self.assertTrue(self.firme({}, "LIBRETTO"))
        for chiave in ("REGISTRO", "COMPUTO", "VARIANTE", "ELENCO", "ANALISI", "SAL", "CDP"):
            self.assertFalse(self.firme({}, chiave), chiave)

    def test_scelta_dell_utente(self):
        ps = {"documenti": json.dumps({"REGISTRO": {"firme": True}, "LIBRETTO": {"firme": False}})}
        self.assertTrue(self.firme(ps, "REGISTRO"))
        self.assertFalse(self.firme(ps, "LIBRETTO"))

    def test_documenti_senza_firme_restano_senza(self):
        ps = {"documenti": json.dumps({"COMPUTO": {"firme": True}, "CDP": {"firme": True}})}
        self.assertFalse(self.firme(ps, "COMPUTO"))
        self.assertFalse(self.firme(ps, "CDP"))

    def test_valori_non_booleani_sono_ignorati(self):
        for valore in ("true", 1, None, [True]):
            with self.subTest(valore=valore):
                ps = {"documenti": json.dumps({"REGISTRO": {"firme": valore}, "LIBRETTO": {"firme": valore}})}
                self.assertFalse(self.firme(ps, "REGISTRO"))
                self.assertTrue(self.firme(ps, "LIBRETTO"))

    def test_firme_scritte_al_centro_da_una_versione_precedente(self):
        ps = {"documenti": json.dumps({"REGISTRO": {"ppCenter": "[FIRME]"}})}
        self.assertTrue(self.firme(ps, "REGISTRO"))
        self.assertEqual(self.ns["impostazioni_documento"](ps, "REGISTRO")["ppCenter"], "")

    def test_la_scelta_esplicita_prevale_su_firme_al_centro_precedenti(self):
        ps = {"documenti": json.dumps({"REGISTRO": {"ppCenter": "[FIRME]", "firme": False}})}
        self.assertFalse(self.firme(ps, "REGISTRO"))


class TestRigaFirme(unittest.TestCase):
    def setUp(self):
        self.ns = _carica()
        self.con_firme = self.ns["piede_con_firme"]
        self.firme = self.ns["_FIRME_TESTO"]

    @staticmethod
    def _righe(testo):
        return testo.count("\n") + 1

    def test_testo_delle_firme_come_nel_template(self):
        self.assertEqual(self.firme, "L'IMPRESA\t\t\t\t\t\tIL DIRETTORE DEI LAVORI    ")

    def test_firme_sotto_una_riga_di_piede(self):
        piede = self.con_firme({"ppSx": "realizzato con LeenO: a.ods", "ppCenter": "", "ppDx": "Pagina 1 di 2"})
        self.assertEqual(piede, {"ppSx": "realizzato con LeenO: a.ods\n", "ppCenter": "\n" + self.firme,
                                 "ppDx": "Pagina 1 di 2\n"})

    def test_firme_dopo_il_centro_esistente(self):
        piede = self.con_firme({"ppSx": "sx", "ppCenter": "centro", "ppDx": "dx"})
        self.assertEqual(piede["ppCenter"], "centro\n" + self.firme)
        self.assertEqual((piede["ppSx"], piede["ppDx"]), ("sx\n", "dx\n"))

    def test_firme_dopo_tutte_le_righe_delle_tre_parti(self):
        piede = self.con_firme({"ppSx": "sx 1\nsx 2", "ppCenter": "c", "ppDx": "dx 1\ndx 2\ndx 3"})
        self.assertEqual(piede["ppSx"], "sx 1\nsx 2\n\n")
        self.assertEqual(piede["ppCenter"], "c\n\n\n" + self.firme)
        self.assertEqual(piede["ppDx"], "dx 1\ndx 2\ndx 3\n")
        # la riga delle firme e' la quarta, sotto le tre righe piu' lunghe
        self.assertEqual(piede["ppCenter"].split("\n").index(self.firme), 3)

    def test_firme_da_sole_se_il_piede_e_vuoto(self):
        piede = self.con_firme({"ppSx": "", "ppCenter": " ", "ppDx": ""})
        self.assertEqual(piede, {"ppSx": "", "ppCenter": self.firme, "ppDx": ""})
        self.assertEqual(self.con_firme({})["ppCenter"], self.firme)

    def test_le_tre_parti_hanno_sempre_lo_stesso_numero_di_righe(self):
        # le parti sono centrate sulla stessa linea: con righe uguali il testo resta in alto
        valori = ("", "a", "a\nb", "a\nb\nc", "a\n", " ")
        for sx in valori:
            for centro in valori:
                for dx in valori:
                    testi = {"ppSx": sx, "ppCenter": centro, "ppDx": dx}
                    with self.subTest(testi=testi):
                        piede = self.con_firme(testi)
                        if not any(v.strip() for v in valori_di(testi)):
                            self.assertEqual(piede["ppCenter"], self.firme)
                            continue
                        righe = {self._righe(v) for v in piede.values()}
                        self.assertEqual(len(righe), 1)
                        massimo = max(self._righe(v) for v in testi.values() if v.strip())
                        self.assertEqual(righe.pop(), massimo + 1)

    def test_le_firme_sono_l_ultima_riga_e_il_testo_resta_al_suo_posto(self):
        valori = ("", "a", "a\nb", "a\nb\nc")
        for sx in valori:
            for centro in valori:
                for dx in valori:
                    testi = {"ppSx": sx, "ppCenter": centro, "ppDx": dx}
                    if not any(testi.values()):
                        continue
                    with self.subTest(testi=testi):
                        piede = self.con_firme(testi)
                        self.assertEqual(piede["ppCenter"].split("\n")[-1], self.firme)
                        self.assertNotIn(self.firme, piede["ppSx"] + piede["ppDx"])
                        for k in ("ppSx", "ppCenter", "ppDx"):
                            self.assertTrue(piede[k].startswith(testi[k]), k)
                        massimo = max(self._righe(v) for v in testi.values() if v.strip())
                        # nessuna riga di testo sta alla riga delle firme o sotto
                        self.assertEqual(piede["ppCenter"].split("\n").index(self.firme), massimo)

    def test_non_modifica_il_dizionario_ricevuto(self):
        testi = {"ppSx": "a", "ppCenter": "", "ppDx": "c"}
        copia = dict(testi)
        self.con_firme(testi)
        self.assertEqual(testi, copia)


def valori_di(testi):
    return list(testi.values())


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

    def test_salva_solo_la_scelta_delle_firme_diversa_dal_predefinito(self):
        valori = self._tutti({})
        valori["LIBRETTO"]["firme"] = False
        valori["REGISTRO"]["firme"] = True
        valori["SAL"]["firme"] = True  # documento senza firme: ignorato
        salvato = self.aggiorna({}, valori)
        self.assertEqual(json.loads(salvato["documenti"]),
                         {"LIBRETTO": {"firme": False}, "REGISTRO": {"firme": True}})
        self.assertFalse(self.ns["firme_documento"](salvato, "LIBRETTO"))
        self.assertTrue(self.ns["firme_documento"](salvato, "REGISTRO"))

    def test_firme_ai_valori_predefiniti_non_si_salvano(self):
        valori = self._tutti({})
        valori["LIBRETTO"]["firme"] = True
        valori["REGISTRO"]["firme"] = False
        self.assertEqual(json.loads(self.aggiorna({}, valori)["documenti"]), {})

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
