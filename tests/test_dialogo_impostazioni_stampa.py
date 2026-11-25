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
FUNZIONI = ("testi_predefiniti", "firme_predefinite", "_personalizzazioni", "_testi_effettivi",
            "impostazioni_documento", "firme_documento", "aggiorna_impostazioni_documenti",
            "documento_stampa_iniziale", "PrintSettingsDialog", "_leggiImpostazioniStampa")
CLASSI = ("_ControlloImpostazioniStampa", "_AscoltatoreImpostazioniStampa")
COSTANTI = ("_DOCUMENTI_STAMPA", "_DOCUMENTI_SENZA_TESTATA", "_DOCUMENTI_CON_FIRME", "_TESTI_STAMPA",
            "_TESTI_PRECEDENTI", "_DOCSTRINGS", "_LEGENDA_SEGNAPOSTI", "_NOTA_SENZA_TESTATA",
            "_DOCUMENTO_DEL_FOGLIO", "_DLG_IMPOSTAZIONI_STAMPA")


def _ids_xdl():
    return set(re.findall(r'dlg:id="([^"]+)"', XDL.read_text(encoding="utf-8")))


def _carica(**extra):
    class _Base:
        pass

    class _XAction:
        pass

    class _XItem:
        pass

    ns = {"json": json, "unohelper": SimpleNamespace(Base=_Base),
          "XActionListener": _XAction, "XItemListener": _XItem}
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
        self.assertEqual(self.modello.getByName("ppCenter").Text, "")  # le firme non sono un testo del centro
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
            self.assertEqual(risultato[chiave], dict(self.pred(chiave), firme=(chiave == "LIBRETTO")))

    def test_ripristina_solo_il_documento_mostrato(self):
        c = self._controllo()
        self.modello.getByName("intSx").Text = "x"
        c.cambia_documento(1)
        self.modello.getByName("intSx").Text = "y"
        c.ripristina()
        self.assertEqual(self._mostrato(), self.pred("VARIANTE"))
        self.assertEqual(c.risultato()["VARIANTE"], dict(self.pred("VARIANTE"), firme=False))
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
        self.assertEqual(c.risultato()["CDP"], dict(dict.fromkeys(TESTI, ""), firme=False))
        c.ripristina()  # non fa nulla
        self.assertEqual(c.risultato()["CDP"], dict(dict.fromkeys(TESTI, ""), firme=False))

    def test_la_casella_delle_firme_e_per_documento(self):
        c = self._controllo(chiave="REGISTRO")
        self.modello.getByName("cbFirme").State = 1
        c.cambia_documento(4)  # Libretto: firme predefinite
        self.assertEqual(self.modello.getByName("cbFirme").State, 1)
        self.modello.getByName("cbFirme").State = 0
        c.cambia_documento(5)  # torna al Registro
        self.assertEqual(self.modello.getByName("cbFirme").State, 1)
        risultato = c.risultato()
        self.assertTrue(risultato["REGISTRO"]["firme"])
        self.assertFalse(risultato["LIBRETTO"]["firme"])

    def test_documenti_senza_firme_non_le_salvano(self):
        c = self._controllo(chiave="COMPUTO")
        self.assertFalse(self.modello.getByName("cbFirme").Enabled)
        self.modello.getByName("cbFirme").State = 1  # controllo disabilitato: valore residuo
        self.assertFalse(c.risultato()["COMPUTO"]["firme"])

    def test_ripristina_riporta_anche_le_firme_ai_predefiniti(self):
        c = self._controllo(chiave="LIBRETTO")
        self.modello.getByName("cbFirme").State = 0
        c.ripristina()
        self.assertEqual(self.modello.getByName("cbFirme").State, 1)
        self.assertTrue(c.risultato()["LIBRETTO"]["firme"])

    def test_le_firme_non_toccano_il_testo_del_centro(self):
        c = self._controllo(chiave="REGISTRO")
        self.modello.getByName("ppCenter").Text = "Centro"
        self.modello.getByName("cbFirme").State = 1
        risultato = c.risultato()["REGISTRO"]
        self.assertEqual((risultato["ppCenter"], risultato["firme"]), ("Centro", True))

    def test_il_risultato_salva_la_scelta_delle_firme(self):
        c = self._controllo(chiave="REGISTRO")
        self.modello.getByName("cbFirme").State = 1
        salvato = self.ns["aggiorna_impostazioni_documenti"]({}, c.risultato())
        self.assertEqual(json.loads(salvato["documenti"]), {"REGISTRO": {"firme": True}})
        self.assertTrue(self.ns["firme_documento"](salvato, "REGISTRO"))

    def test_indice_documento_non_valido_e_ignorato(self):
        c = self._controllo()
        c.cambia_documento(-1)
        c.cambia_documento(99)
        self.assertEqual(c.corrente, "COMPUTO")

    def test_parte_dalle_impostazioni_salvate(self):
        ps = {"documenti": json.dumps({"LIBRETTO": {"intSx": "Ente"}})}
        c = self._controllo(ps, "LIBRETTO")
        self.assertEqual(self.modello.getByName("intSx").Text, "Ente")
        self.assertTrue(c.risultato()["LIBRETTO"]["firme"])

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
                         LeenoUtils=leeno_utils, Dialogs=SimpleNamespace(FileSelect=lambda **kw: None))
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
        self.assertEqual(m.getByName("ppCenter").Text, "")
        self.assertEqual(m.getByName("cbFirme").State, 1)
        self.assertEqual(m.getByName("lblSegnaposto").Label, self.ns["_LEGENDA_SEGNAPOSTI"])
        self.assertNotIn("[FIRME]", tuple(m.getByName("ppSx").StringItemList))
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
        ascolt.itemStateChanged(self._evento(oDlg, "docSel", Selected=4))
        self.assertEqual(m.getByName("cbFirme").State, 1)
        m.getByName("cbFirme").State = 0
        ascolt.itemStateChanged(self._evento(oDlg, "docSel", Selected=5))
        self.assertEqual(m.getByName("cbFirme").State, 1)
        self.assertEqual(m.getByName("ppCenter").Text, "")
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
        ps = self.ns["_leggiImpostazioniStampa"](oDlg, {"fileCopertine": "x.ods"}, controllo)
        self.assertEqual(json.loads(ps["documenti"]),
                         {"COMPUTO": {"intSx": "Ente di prova"},
                          "LIBRETTO": {"ppSx": "Libretto [FILE]", "firme": False}})
        for k, v in self.ns["_TESTI_PRECEDENTI"].items():
            self.assertEqual(ps[k], v)

    def test_documento_iniziale_non_valido_ripiega_su_computo(self):
        oDlg, _, controllo = self._dialogo(chiave="INESISTENTE")
        self.assertEqual(controllo.corrente, "COMPUTO")
        self.assertEqual(oDlg.getModel().getByName("docSel").SelectedItems, (0,))


if __name__ == "__main__":
    unittest.main()
