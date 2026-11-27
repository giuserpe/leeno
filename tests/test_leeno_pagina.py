# -*- coding: utf-8 -*-
"""
Test del motore di impaginazione (LeenoPagina), senza UNO: stili di pagina e LeenoSettings finti.
Verificano quali operazioni riceve ogni tipo di stile e che la testata di un documento sia
calcolata una volta sola.
"""
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

try:
    from tests.estrai_sorgente import carica
except ImportError:  # esecuzione con la cartella tests nel percorso di ricerca
    from estrai_sorgente import carica

ROOT = Path(__file__).resolve().parent.parent
PAGINA = ROOT / "src/Ultimus.oxt/python/pythonpath/LeenoPagina.py"

CDP = "PageStyle_CdP_A4"
COMPUTO = "PageStyle_COMPUTO_A4"
SAL = "PageStyle_SAL_A4"


def _motore():
    ls = MagicMock()
    ls.STILE_PAGINA_CDP = CDP
    ls.STILI_TESTATA = frozenset((COMPUTO, "PageStyle_Elenco Prezzi", "PageStyle_Analisi di Prezzo"))
    ls.STILI_ATTI_CONTABILI = frozenset((SAL, "PageStyle_REGISTRO_A4", "Page_Style_Libretto_Misure2"))
    ls.loadPrintSettings.return_value = ({"x": 1}, None)
    ls.nome_file_documento.return_value = "prova.ods"
    ls.documento_dello_stile.side_effect = lambda nome, foglio: (
        "VARIANTE" if nome == COMPUTO and foglio == "VARIANTE" else
        "COMPUTO" if nome == COMPUTO else "SAL")
    ls.titolo_a_schermo.side_effect = lambda chiave, oDoc: "Titolo " + chiave
    ls.testata_a_schermo.side_effect = lambda testi, valori, chiave: {"chiave": chiave}
    ls.dimensione_testata.return_value = 8.0
    ls.ha_firme.return_value = False
    ns = carica(PAGINA, ("contesto_stampa", "applica_stile_pagina"), ("_STILE_COPERTINA",), {"LS": ls})
    return ns, ls


def _stile(nome):
    return SimpleNamespace(DisplayName=nome, HeaderIsOn=None, FooterIsOn=None)


class TestContestoStampa(unittest.TestCase):
    def test_raccoglie_i_dati_una_volta(self):
        ns, ls = _motore()
        ctx = ns["contesto_stampa"]("doc", "3")
        self.assertEqual(ctx["printSettings"], {"x": 1})
        self.assertEqual(ctx["nomefile"], "prova.ods")
        self.assertEqual(ctx["n_sal"], "3")
        self.assertEqual(ctx["testate"], {})

    def test_n_sal_predefinito_assente(self):
        ns, ls = _motore()
        self.assertIsNone(ns["contesto_stampa"]("doc")["n_sal"])


class TestApplicaStilePagina(unittest.TestCase):
    def setUp(self):
        self.ns, self.ls = _motore()
        self.ctx = self.ns["contesto_stampa"]("doc", "2")

    def _applica(self, nome, foglio="COMPUTO"):
        stile = _stile(nome)
        self.ns["applica_stile_pagina"]("doc", stile, foglio, self.ctx)
        return stile

    def test_copertina_e_cdp_senza_intestazione_ne_testata(self):
        for nome in ("Page_Style_COPERTINE", CDP):
            self.ls.reset_mock()
            stile = self._applica(nome)
            self.assertFalse(stile.HeaderIsOn)
            self.assertFalse(stile.FooterIsOn)
            self.ls.set_page_margins.assert_called_once_with(stile)
            self.ls.scrivi_testata_stile.assert_not_called()

    def test_stile_estraneo_solo_geometria(self):
        stile = self._applica("Default")
        self.assertTrue(stile.HeaderIsOn)
        self.assertTrue(stile.FooterIsOn)
        self.ls.set_page_margins.assert_called_once_with(stile)
        self.ls.scrivi_testata_stile.assert_not_called()
        self.ls.set_page_borders.assert_not_called()

    def test_stile_del_computo_scrive_testata_e_bordi(self):
        stile = self._applica(COMPUTO)
        self.ls.set_page_borders.assert_called_once_with(stile)
        self.ls.scrivi_testata_stile.assert_called_once_with(
            "doc", stile, {"chiave": "COMPUTO"}, 8.0, False)

    def test_atto_contabile_senza_bordi(self):
        stile = self._applica(SAL)
        self.ls.set_page_borders.assert_not_called()
        self.ls.scrivi_testata_stile.assert_called_once_with(
            "doc", stile, {"chiave": "SAL"}, 8.0, False)

    def test_foglio_attivo_sceglie_il_documento_dello_stile_condiviso(self):
        self._applica(COMPUTO, foglio="VARIANTE")
        args = self.ls.scrivi_testata_stile.call_args[0]
        self.assertEqual(args[2], {"chiave": "VARIANTE"})

    def test_testata_calcolata_una_volta_per_documento(self):
        self._applica(COMPUTO)
        self._applica(COMPUTO)
        self.assertEqual(self.ls.testata_a_schermo.call_count, 1)
        self.assertEqual(self.ls.scrivi_testata_stile.call_count, 2)

    def test_valori_segnaposto_con_titolo_sal_e_nomefile(self):
        self._applica(SAL)
        self.ls.valori_segnaposto.assert_called_once_with(
            "doc", documento="Titolo SAL", nSal="2", nomefile="prova.ods")


if __name__ == "__main__":
    unittest.main()
