import ast
import unittest
from pathlib import Path
from unittest.mock import MagicMock

PYTHONPATH_DIR = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath"
PYLEENO = PYTHONPATH_DIR / "pyleeno.py"
SETTINGS = PYTHONPATH_DIR / "LeenoSettings.py"


def _funzione(percorso, nome):
    tree = ast.parse(percorso.read_text(encoding="utf-8"))
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name == nome:
            n.decorator_list = []
            return n
    raise AssertionError("funzione non trovata: " + nome)


def _compila(percorso, nome, ns):
    mod = ast.Module(body=[_funzione(percorso, nome)], type_ignores=[])
    exec(compile(mod, str(percorso), "exec"), ns)
    return ns[nome]


class TestAreaStampaCdP(unittest.TestCase):
    def _esegui(self, nomi, aree):
        """aree: {nome_area: (StartRow, EndRow)}"""
        oDoc = MagicMock()
        oSheet = MagicMock()
        oSheet.Name = "CdP"
        oSheet.RangeAddress.Sheet = 5
        oSheet.getPrintAreas.return_value = []
        oDoc.CurrentController.ActiveSheet = oSheet
        oDoc.NamedRanges.ElementNames = nomi

        def per_nome(nome):
            r = MagicMock()
            r.ReferredCells.RangeAddress.StartRow = aree[nome][0]
            r.ReferredCells.RangeAddress.EndRow = aree[nome][1]
            r.ReferredCells.RangeAddress.EndColumn = 5
            return r

        oDoc.NamedRanges.getByName.side_effect = per_nome
        leeno_utils = MagicMock()
        leeno_utils.getDocument.return_value = oDoc
        sheet_utils = MagicMock()
        sheet_utils.getLastUsedRow.return_value = 200
        sheet_utils.getLastUsedColumn.return_value = 9
        uno = MagicMock()
        uno.createUnoStruct.side_effect = lambda _: MagicMock()
        ns = {"LeenoUtils": leeno_utils, "SheetUtils": sheet_utils, "uno": uno, "LS": MagicMock()}
        _compila(PYLEENO, "set_area_stampa", ns)
        ns["set_area_stampa"]()
        return oSheet, leeno_utils

    def test_area_di_stampa_coincide_con_l_area_dati_dell_ultimo_cdp(self):
        oSheet, _ = self._esegui(
            ["_CdP_1", "_CdP_2", "_CdP_10", "_CdP_x", "_Lib_1"],
            {"_CdP_1": (1, 46), "_CdP_2": (48, 92), "_CdP_10": (500, 545)},
        )
        area = oSheet.setPrintAreas.call_args.args[0][0]
        self.assertEqual(area.StartRow, 500)
        self.assertEqual(area.EndRow, 545)
        self.assertEqual((area.StartColumn, area.EndColumn), (0, 5))

    def test_senza_nomi_area_ripiega_sull_area_usata_e_registra_l_errore(self):
        oSheet, leeno_utils = self._esegui(["_Lib_1"], {})
        area = oSheet.setPrintAreas.call_args.args[0][0]
        self.assertEqual((area.StartRow, area.EndRow), (0, 200))
        leeno_utils.log_errore.assert_called_once()

    def test_nessuna_riga_da_ripetere(self):
        oSheet, _ = self._esegui(["_CdP_1"], {"_CdP_1": (1, 46)})
        oSheet.setTitleRows.assert_not_called()
        oSheet.setPrintTitleRows.assert_called_once_with(False)


class TestStilePaginaCdP(unittest.TestCase):
    def _crea(self, ha_foglio, ha_stile):
        oDoc = MagicMock()
        oDoc.getSheets().hasByName.return_value = ha_foglio
        oPageStyles = oDoc.StyleFamilies.getByName.return_value
        oPageStyles.hasByName.return_value = ha_stile
        ns = {"STILE_PAGINA_CDP": "PageStyle_CdP_A4"}
        fn = _compila(SETTINGS, "_crea_stile_pagina_cdp", ns)
        fn(oDoc)
        return oDoc, oPageStyles

    def test_crea_lo_stile_se_c_e_il_foglio_cdp(self):
        oDoc, oPageStyles = self._crea(True, False)
        oPageStyles.insertByName.assert_called_once()
        self.assertEqual(oPageStyles.insertByName.call_args.args[0], "PageStyle_CdP_A4")
        stile = oPageStyles.getByName.return_value
        self.assertFalse(stile.HeaderIsOn)
        self.assertFalse(stile.FooterIsOn)
        self.assertEqual((stile.Width, stile.Height), (21000, 29700))

    def test_non_crea_lo_stile_senza_foglio_cdp(self):
        _, oPageStyles = self._crea(False, False)
        oPageStyles.insertByName.assert_not_called()

    def test_non_ricrea_lo_stile_se_esiste(self):
        _, oPageStyles = self._crea(True, True)
        oPageStyles.insertByName.assert_not_called()

    def test_stile_associato_al_foglio_cdp(self):
        codice = SETTINGS.read_text(encoding="utf-8")
        self.assertIn("'CdP': STILE_PAGINA_CDP", codice)

    def test_sistema_pagine_non_riattiva_intestazione_del_cdp(self):
        fn = _funzione(PYLEENO, "MENU_sistema_pagine")
        esclusione = None
        for n in ast.walk(fn):
            if isinstance(n, ast.Assign) and any(
                    isinstance(t, ast.Name) and t.id == "senza_intestazione" for t in n.targets):
                esclusione = ast.unparse(n.value)
        self.assertIsNotNone(esclusione)
        self.assertIn("LS.STILE_PAGINA_CDP", esclusione)
        attivazioni = [ast.unparse(n.value) for n in ast.walk(fn)
                       if isinstance(n, ast.Assign)
                       and any(isinstance(t, ast.Attribute) and t.attr in ("HeaderIsOn", "FooterIsOn")
                               for t in n.targets)]
        self.assertEqual(len(attivazioni), 2)
        for a in attivazioni:
            self.assertIn("senza_intestazione", a)


if __name__ == "__main__":
    unittest.main()
