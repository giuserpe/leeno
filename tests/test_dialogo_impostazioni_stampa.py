# -*- coding: utf-8 -*-
"""
Test del dialogo Impostazioni stampa / PDF con un set di stringhe per documento.

- unitari, senza UNO: logica del controllo (cambio documento, ripristino, firme, documento
  senza testata) su un modello finto i cui nomi devono esistere nell'xdl reale;
- documento reale (opzionale, LEENO_TEST_UNO=1): dialogo vero creato dall'xdl senza schermo,
  con ascoltatori e salvataggio completo. La riga delle firme e' un'opzione per documento
  (casella), non un testo del piè di pagina.
"""
import json
import re
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
SETTINGS = ROOT / "src/Ultimus.oxt/python/pythonpath/LeenoSettings.py"
XDL = ROOT / "src/Ultimus.oxt/UltimusFree2/DlgImpostazioniStampa.xdl"
ESEMPIO = ROOT / "tests/data/LeenO_Esempio.ods"

TESTI = ("intSx", "intCenter", "intDx", "ppSx", "ppCenter", "ppDx")
FUNZIONI = ("testi_predefiniti", "_personalizzazioni", "impostazioni_documento", "senza_firme", "ha_firme",
            "aggiorna_impostazioni_documenti", "documento_stampa_iniziale", "PrintSettingsDialog",
            "_leggiImpostazioniStampa")
CLASSI = ("_ControlloImpostazioniStampa", "_AscoltatoreImpostazioniStampa")
COSTANTI = ("_DOCUMENTI_STAMPA", "_DOCUMENTI_SENZA_TESTATA", "_DOCUMENTI_CON_FIRME", "_TESTI_STAMPA",
            "_TESTI_PRECEDENTI", "_DOCSTRINGS", "_LEGENDA_SEGNAPOSTI", "_NOTA_SENZA_TESTATA",
            "_DOCUMENTO_DEL_FOGLIO", "_DLG_IMPOSTAZIONI_STAMPA", "_SEGNAPOSTO_FIRME", "_PARTI_PIEDE")


def _ids_xdl():
    return set(re.findall(r'dlg:id="([^"]+)"', XDL.read_text(encoding="utf-8")))


def _carica(**extra):
    class _Base:
        pass

    class _XAction:
        pass

    class _XItem:
        pass

    class _XText:
        pass

    ns = {"json": json, "unohelper": SimpleNamespace(Base=_Base),
          "XActionListener": _XAction, "XItemListener": _XItem, "XTextListener": _XText}
    ns.update(extra)
    return carica(SETTINGS, FUNZIONI, COSTANTI, ns, CLASSI)


class ModelloFinto:
    """Modello del dialogo: i nomi ammessi sono quelli dell'xdl reale."""

    def __init__(self):
        self._ids = _ids_xdl()
        self._oggetti = {}

    def getByName(self, nome):
        if nome not in self._ids:
            raise KeyError("controllo inesistente nell'xdl: " + nome)
        if nome not in self._oggetti:
            self._oggetti[nome] = SimpleNamespace(Text="", Enabled=True, State=0, Label="")
        return self._oggetti[nome]


class TestLogicaDialogo(unittest.TestCase):
    def setUp(self):
        self.ns = _carica()
        self.modello = ModelloFinto()
        self.pred = self.ns["testi_predefiniti"]

    def _controllo(self, ps=None, chiave="COMPUTO"):
        return self.ns["_ControlloImpostazioniStampa"](self.modello, ps or {}, chiave)

    def _mostrato(self):
        return {k: self.modello.getByName(k).Text for k in TESTI}

    def test_i_nomi_usati_esistono_nell_xdl(self):
        ids = _ids_xdl()
        attesi = {"docSel", "btnPredefiniti", "cbFirme", "lblSegnaposto", "fileCopertine", "copertina",
                  "btnSfoglia", "btnOk", "btnAnnulla", "imgAnteprima"}
        attesi |= set(TESTI) | {"lbl_" + k for k in TESTI}
        self.assertEqual(attesi - ids, set())

    def test_all_apertura_mostra_i_predefiniti_del_documento(self):
        self._controllo()
        self.assertEqual(self._mostrato(), self.pred("COMPUTO"))
        self.assertFalse(self.modello.getByName("cbFirme").Enabled)
        self.assertEqual(self.modello.getByName("cbFirme").State, 0)
        self.assertEqual(self.modello.getByName("lblSegnaposto").Label, self.ns["_LEGENDA_SEGNAPOSTI"])

    def test_libretto_ha_la_riga_delle_firme(self):
        self._controllo(chiave="LIBRETTO")
        self.assertEqual(self.modello.getByName("ppCenter").Text, "[FIRME]")
        self.assertTrue(self.modello.getByName("cbFirme").Enabled)
        self.assertEqual(self.modello.getByName("cbFirme").State, 1)

    def test_registro_firme_facoltative_spente(self):
        self._controllo(chiave="REGISTRO")
        self.assertTrue(self.modello.getByName("cbFirme").Enabled)
        self.assertEqual(self.modello.getByName("cbFirme").State, 0)

    def test_le_modifiche_restano_al_cambio_documento(self):
        c = self._controllo()
        self.modello.getByName("intSx").Text = "Ente di prova"
        c.cambia_documento(4)  # Libretto delle Misure
        self.assertEqual(self.modello.getByName("intSx").Text, self.pred("LIBRETTO")["intSx"])
        c.cambia_documento(0)
        self.assertEqual(self.modello.getByName("intSx").Text, "Ente di prova")

    def test_i_documenti_non_si_contaminano(self):
        c = self._controllo()
        self.modello.getByName("ppSx").Text = "solo computo"
        risultato = c.risultato()
        self.assertEqual(risultato["COMPUTO"]["ppSx"], "solo computo")
        for chiave in ("VARIANTE", "ELENCO", "ANALISI", "LIBRETTO", "REGISTRO", "SAL"):
            self.assertEqual(risultato[chiave], self.pred(chiave))

    def test_ripristina_solo_il_documento_mostrato(self):
        c = self._controllo()
        self.modello.getByName("intSx").Text = "x"
        c.cambia_documento(1)
        self.modello.getByName("intSx").Text = "y"
        c.ripristina()
        self.assertEqual(self._mostrato(), self.pred("VARIANTE"))
        self.assertEqual(c.risultato()["VARIANTE"], self.pred("VARIANTE"))
        self.assertEqual(c.risultato()["COMPUTO"]["intSx"], "x")

    def test_certificato_di_pagamento_senza_testata(self):
        c = self._controllo()
        c.cambia_documento(7)
        for k in TESTI:
            self.assertFalse(self.modello.getByName(k).Enabled, k)
            self.assertFalse(self.modello.getByName("lbl_" + k).Enabled, k)
        self.assertFalse(self.modello.getByName("btnPredefiniti").Enabled)
        self.assertFalse(self.modello.getByName("cbFirme").Enabled)
        self.assertEqual(self.modello.getByName("lblSegnaposto").Label, self.ns["_NOTA_SENZA_TESTATA"])
        c.cambia_documento(0)
        for k in TESTI:
            self.assertTrue(self.modello.getByName(k).Enabled, k)
        self.assertTrue(self.modello.getByName("btnPredefiniti").Enabled)
        self.assertEqual(self.modello.getByName("lblSegnaposto").Label, self.ns["_LEGENDA_SEGNAPOSTI"])

    def test_certificato_di_pagamento_ignora_il_testo_dei_campi_disabilitati(self):
        c = self._controllo()
        self.modello.getByName("intSx").Text = "residuo"
        c.cambia_documento(7)
        self.assertEqual(c.risultato()["CDP"], dict.fromkeys(TESTI, ""))
        c.ripristina()  # non fa nulla
        self.assertEqual(c.risultato()["CDP"], dict.fromkeys(TESTI, ""))

    def test_casella_firme_inserisce_e_toglie_il_segnaposto(self):
        c = self._controllo(chiave="REGISTRO")
        self.assertEqual(self.modello.getByName("cbFirme").State, 0)
        self.modello.getByName("cbFirme").State = 1
        c.firme_cambiate()
        self.assertEqual(self.modello.getByName("ppCenter").Text, "[FIRME]")
        self.modello.getByName("cbFirme").State = 0
        c.firme_cambiate()
        self.assertEqual(self.modello.getByName("ppCenter").Text, "")

    def test_la_casella_aggiunge_il_segnaposto_dopo_il_testo_del_centro(self):
        c = self._controllo(chiave="REGISTRO")
        self.modello.getByName("ppCenter").Text = "Centro"
        self.modello.getByName("cbFirme").State = 1
        c.firme_cambiate()
        self.assertEqual(self.modello.getByName("ppCenter").Text, "Centro [FIRME]")
        self.modello.getByName("cbFirme").State = 0
        c.firme_cambiate()
        self.assertEqual(self.modello.getByName("ppCenter").Text, "Centro")

    def test_togliere_le_firme_le_toglie_da_tutte_le_parti(self):
        c = self._controllo(chiave="LIBRETTO")
        self.modello.getByName("ppSx").Text = "sx [FIRME]"
        self.modello.getByName("ppDx").Text = "[FIRME] dx"
        self.modello.getByName("cbFirme").State = 0
        c.firme_cambiate()
        self.assertEqual([self.modello.getByName(k).Text for k in ("ppSx", "ppCenter", "ppDx")],
                         ["sx", "", " dx"])

    def test_casella_se_il_segnaposto_c_e_gia_non_lo_raddoppia(self):
        c = self._controllo(chiave="REGISTRO")
        self.modello.getByName("ppSx").Text = "x [FIRME]"
        self.modello.getByName("cbFirme").State = 1
        c.firme_cambiate()
        self.assertEqual(self.modello.getByName("ppCenter").Text, "")

    def test_scrivere_il_segnaposto_a_mano_spunta_la_casella(self):
        c = self._controllo(chiave="REGISTRO")
        self.modello.getByName("ppDx").Text = "Pagina [PAGINA] [FIRME]"
        c.piede_cambiato()
        self.assertEqual(self.modello.getByName("cbFirme").State, 1)
        self.modello.getByName("ppDx").Text = "Pagina [PAGINA]"
        c.piede_cambiato()
        self.assertEqual(self.modello.getByName("cbFirme").State, 0)

    def test_casella_inattiva_per_i_documenti_senza_firme(self):
        c = self._controllo(chiave="COMPUTO")
        self.assertFalse(self.modello.getByName("cbFirme").Enabled)
        self.modello.getByName("cbFirme").State = 1
        c.firme_cambiate()
        self.assertEqual(self.modello.getByName("ppCenter").Text, "")

    def test_il_segnaposto_si_puo_scrivere_in_qualsiasi_documento(self):
        c = self._controllo(chiave="COMPUTO")
        self.modello.getByName("ppCenter").Text = "[FIRME]"
        self.assertEqual(c.risultato()["COMPUTO"]["ppCenter"], "[FIRME]")

    def test_ripristina_riporta_anche_le_firme_ai_predefiniti(self):
        c = self._controllo(chiave="LIBRETTO")
        self.modello.getByName("ppCenter").Text = ""
        c.piede_cambiato()
        self.assertEqual(self.modello.getByName("cbFirme").State, 0)
        c.ripristina()
        self.assertEqual(self.modello.getByName("cbFirme").State, 1)
        self.assertEqual(self.modello.getByName("ppCenter").Text, "[FIRME]")

    def test_la_casella_segue_il_documento_mostrato(self):
        c = self._controllo(chiave="REGISTRO")
        self.assertEqual(self.modello.getByName("cbFirme").State, 0)
        c.cambia_documento(4)  # Libretto: firme predefinite
        self.assertEqual(self.modello.getByName("cbFirme").State, 1)
        c.cambia_documento(5)
        self.assertEqual(self.modello.getByName("cbFirme").State, 0)

    def test_il_risultato_salva_il_segnaposto(self):
        c = self._controllo(chiave="REGISTRO")
        self.modello.getByName("cbFirme").State = 1
        c.firme_cambiate()
        salvato = self.ns["aggiorna_impostazioni_documenti"]({}, c.risultato())
        self.assertEqual(json.loads(salvato["documenti"]), {"REGISTRO": {"ppCenter": "[FIRME]"}})

    def test_indice_documento_non_valido_e_ignorato(self):
        c = self._controllo()
        c.cambia_documento(-1)
        c.cambia_documento(99)
        self.assertEqual(c.corrente, "COMPUTO")

    def test_parte_dalle_impostazioni_salvate(self):
        ps = {"documenti": json.dumps({"LIBRETTO": {"intSx": "Ente"}})}
        c = self._controllo(ps, "LIBRETTO")
        self.assertEqual(self.modello.getByName("intSx").Text, "Ente")
        self.assertEqual(c.risultato()["LIBRETTO"]["ppCenter"], "[FIRME]")

    def test_il_risultato_si_salva_con_le_sole_differenze(self):
        c = self._controllo()
        self.modello.getByName("intDx").Text = "Dx personalizzato"
        salvato = self.ns["aggiorna_impostazioni_documenti"]({}, c.risultato())
        self.assertEqual(json.loads(salvato["documenti"]), {"COMPUTO": {"intDx": "Dx personalizzato"}})

    def test_ps_non_viene_modificato(self):
        ps = {"intSx": "Mio ente"}
        copia = dict(ps)
        self._controllo(ps)
        self.assertEqual(ps, copia)


class TestDocumentoIniziale(unittest.TestCase):
    def setUp(self):
        self.iniziale = _carica()["documento_stampa_iniziale"]

    @staticmethod
    def _doc(nome):
        return SimpleNamespace(CurrentController=SimpleNamespace(ActiveSheet=SimpleNamespace(Name=nome)))

    def test_foglio_attivo_e_documento(self):
        for foglio, chiave in (("COMPUTO", "COMPUTO"), ("VARIANTE", "VARIANTE"), ("Elenco Prezzi", "ELENCO"),
                               ("Analisi di Prezzo", "ANALISI"), ("CONTABILITA", "LIBRETTO"),
                               ("Registro", "REGISTRO"), ("SAL", "SAL"), ("CdP", "CDP")):
            with self.subTest(foglio=foglio):
                self.assertEqual(self.iniziale(self._doc(foglio)), chiave)

    def test_foglio_sconosciuto_o_errore(self):
        self.assertEqual(self.iniziale(self._doc("S2")), "COMPUTO")
        self.assertEqual(self.iniziale(SimpleNamespace()), "COMPUTO")


@unittest.skipUnless(uno_harness.uno_disponibile(ESEMPIO),
                     "richiede LEENO_TEST_UNO=1, uno, soffice e LeenO_Esempio.ods")
class TestDialogoReale(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import uno
        import unohelper
        # uno.getClass non passa da sys.modules, dove altri test mettono finti per com.sun.star
        XActionListener = uno.getClass("com.sun.star.awt.XActionListener")
        XItemListener = uno.getClass("com.sun.star.awt.XItemListener")
        XTextListener = uno.getClass("com.sun.star.awt.XTextListener")
        cls.sessione = uno_harness.SessioneUno()
        cls.addClassCleanup(cls.sessione.chiudi)
        cls.sessione.avvia()
        ctx = uno_harness.connetti()
        cls.errori = []
        leeno_utils = SimpleNamespace(
            getComponentContext=lambda: ctx,
            LeenO_path=lambda: str(ROOT / "src/Ultimus.oxt"),
            log_errore=lambda contesto, exc=None, percorso=None: cls.errori.append((contesto, exc)))
        cls.ns = _carica(uno=uno, unohelper=unohelper, XActionListener=XActionListener,
                         XItemListener=XItemListener,
                         XTextListener=XTextListener, LeenoUtils=leeno_utils, Dialogs=SimpleNamespace(FileSelect=lambda **kw: None))
        cls.url = uno.systemPathToFileUrl(str(XDL))

    def setUp(self):
        self.errori.clear()

    def _dialogo(self, ps=None, chiave="COMPUTO", covers=()):
        oDlg, ascolt, controllo = self.ns["PrintSettingsDialog"](ps or {}, covers, chiave, url=self.url,
                                                                  immagini=False)
        self.addCleanup(oDlg.dispose)
        return oDlg, ascolt, controllo

    def _evento(self, oDlg, nome, **kw):
        return SimpleNamespace(Source=oDlg.getControl(nome), **kw)

    def test_il_dialogo_si_crea_con_tutti_i_documenti(self):
        oDlg, _, _ = self._dialogo(chiave="LIBRETTO")
        m = oDlg.getModel()
        self.assertEqual(tuple(m.getByName("docSel").StringItemList),
                         tuple(t for _, t in self.ns["_DOCUMENTI_STAMPA"]))
        self.assertEqual(m.getByName("docSel").SelectedItems, (4,))
        self.assertEqual(m.getByName("ppCenter").Text, "[FIRME]")
        self.assertEqual(m.getByName("cbFirme").State, 1)
        self.assertEqual(m.getByName("lblSegnaposto").Label, self.ns["_LEGENDA_SEGNAPOSTI"])
        self.assertIn("[FIRME]", m.getByName("lblSegnaposto").Label)
        self.assertIn("[FIRME]", tuple(m.getByName("ppSx").StringItemList))
        self.assertIn("[SAL]", tuple(m.getByName("ppSx").StringItemList))

    def test_cambio_documento_dal_selettore(self):
        oDlg, ascolt, _ = self._dialogo()
        m = oDlg.getModel()
        m.getByName("intSx").Text = "Ente"
        ascolt.itemStateChanged(self._evento(oDlg, "docSel", Selected=5))  # Registro
        self.assertEqual(m.getByName("intDx").Text, "[OGGETTO] n. [SAL]")
        ascolt.itemStateChanged(self._evento(oDlg, "docSel", Selected=0))
        self.assertEqual(m.getByName("intSx").Text, "Ente")
        self.assertEqual(self.errori, [])

    def test_pulsante_ripristina_testi_predefiniti(self):
        oDlg, ascolt, _ = self._dialogo()
        m = oDlg.getModel()
        m.getByName("ppDx").Text = "x"
        ascolt.actionPerformed(self._evento(oDlg, "btnPredefiniti"))
        self.assertEqual(m.getByName("ppDx").Text, "Pagina [PAGINA] di [PAGINE]")
        self.assertEqual(self.errori, [])

    def test_casella_firme_con_il_dialogo_reale(self):
        oDlg, ascolt, _ = self._dialogo(chiave="REGISTRO")
        m = oDlg.getModel()
        m.getByName("cbFirme").State = 1
        ascolt.itemStateChanged(self._evento(oDlg, "cbFirme"))
        self.assertEqual(m.getByName("ppCenter").Text, "[FIRME]")
        m.getByName("ppCenter").Text = "altro"
        ascolt.textChanged(self._evento(oDlg, "ppCenter"))
        self.assertEqual(m.getByName("cbFirme").State, 0)
        self.assertEqual(self.errori, [])

    def test_certificato_di_pagamento_disabilita_i_campi(self):
        oDlg, ascolt, _ = self._dialogo()
        m = oDlg.getModel()
        ascolt.itemStateChanged(self._evento(oDlg, "docSel", Selected=7))
        self.assertFalse(m.getByName("intSx").Enabled)
        self.assertFalse(m.getByName("btnPredefiniti").Enabled)
        self.assertEqual(m.getByName("lblSegnaposto").Label, self.ns["_NOTA_SENZA_TESTATA"])

    def test_salvataggio_completo(self):
        oDlg, ascolt, controllo = self._dialogo(ps={"fileCopertine": "", "copertina": ""}, covers=("A", "B"))
        m = oDlg.getModel()
        m.getByName("intSx").Text = "Ente di prova"
        ascolt.itemStateChanged(self._evento(oDlg, "docSel", Selected=4))
        m.getByName("ppSx").Text = "Libretto [FILE]"
        oDlg.getControl("copertina").selectItem("B", True)
        ps = self.ns["_leggiImpostazioniStampa"](oDlg, {"fileCopertine": "x.ods"}, controllo)
        self.assertEqual(ps["copertina"], "B")
        m.getByName("cbFirme").State = 0
        ascolt.itemStateChanged(self._evento(oDlg, "cbFirme"))
        ps = self.ns["_leggiImpostazioniStampa"](oDlg, {"fileCopertine": "x.ods"}, controllo)
        self.assertEqual(json.loads(ps["documenti"]),
                         {"COMPUTO": {"intSx": "Ente di prova"},
                          "LIBRETTO": {"ppSx": "Libretto [FILE]", "ppCenter": ""}})
        for k, v in self.ns["_TESTI_PRECEDENTI"].items():
            self.assertEqual(ps[k], v)

    def test_documento_iniziale_non_valido_ripiega_su_computo(self):
        oDlg, _, controllo = self._dialogo(chiave="INESISTENTE")
        self.assertEqual(controllo.corrente, "COMPUTO")
        self.assertEqual(oDlg.getModel().getByName("docSel").SelectedItems, (0,))


if __name__ == "__main__":
    unittest.main()
