import ast
import os
import unittest
from pathlib import Path
from unittest.mock import MagicMock

SRC_EXPORT = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoExport.py"


def _estrai_funzione(file_path, nome):
    tree = ast.parse(file_path.read_text(encoding="utf-8"))
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name == nome:
            mod = ast.Module(body=[n], type_ignores=[])
            return compile(mod, filename=str(file_path), mode="exec")
    raise AssertionError(f"Funzione {nome} non trovata in {file_path}")


class TestExportSelectedRangeToOdt(unittest.TestCase):

    def setUp(self):
        self.mock_leeno_utils = MagicMock()
        self.mock_dlg = MagicMock()
        self.mock_dialogs = MagicMock()
        self.mock_uno = MagicMock()
        self.mock_property_value = MagicMock()

        self.ns = {
            "LeenoUtils": self.mock_leeno_utils,
            "DLG": self.mock_dlg,
            "Dialogs": self.mock_dialogs,
            "uno": self.mock_uno,
            "PropertyValue": self.mock_property_value,
            "os": os,
        }

        code = _estrai_funzione(SRC_EXPORT, "MENU_export_selected_range_to_odt")
        exec(code, self.ns)
        self.func = self.ns["MENU_export_selected_range_to_odt"]

    def test_selezione_non_valida_mostra_chi(self):
        oDoc = MagicMock()
        selection = MagicMock()
        selection.supportsService.return_value = False
        oDoc.getCurrentSelection.return_value = selection
        self.mock_leeno_utils.getDocument.return_value = oDoc

        self.func()

        self.mock_dlg.chi.assert_called_once_with("Seleziona un range di celle prima di eseguire la macro!")
        self.mock_dialogs.FileSelect.assert_not_called()

    def test_default_name_da_doc_url(self):
        oDoc = MagicMock()
        selection = MagicMock()
        selection.supportsService.return_value = True
        oDoc.getCurrentSelection.return_value = selection
        oDoc.getURL.return_value = "file:///home/user/documenti/computo_test.ods"
        self.mock_leeno_utils.getDocument.return_value = oDoc

        self.mock_uno.fileUrlToSystemPath.return_value = "/home/user/documenti/computo_test.ods"
        self.mock_dialogs.FileSelect.return_value = ""  # Utente annulla

        self.func()

        self.mock_dialogs.FileSelect.assert_called_once_with(
            'Salva con nome...',
            '*.odt',
            1,
            startPath='/home/user/documenti',
            defaultName='computo_test.odt'
        )
        self.mock_leeno_utils.getDesktop.assert_not_called()

    def test_default_name_da_title_se_url_vuoto(self):
        oDoc = MagicMock()
        selection = MagicMock()
        selection.supportsService.return_value = True
        oDoc.getCurrentSelection.return_value = selection
        oDoc.getURL.return_value = ""
        oDoc.Title = "Progetto_Villa.ods"
        self.mock_leeno_utils.getDocument.return_value = oDoc

        self.mock_dialogs.FileSelect.return_value = ""  # Utente annulla

        self.func()

        self.mock_dialogs.FileSelect.assert_called_once_with(
            'Salva con nome...',
            '*.odt',
            1,
            startPath=None,
            defaultName='Progetto_Villa.odt'
        )

    def test_default_name_fallback_documento(self):
        oDoc = MagicMock(spec=["getCurrentSelection"])
        selection = MagicMock()
        selection.supportsService.return_value = True
        oDoc.getCurrentSelection.return_value = selection
        self.mock_leeno_utils.getDocument.return_value = oDoc

        self.mock_dialogs.FileSelect.return_value = ""  # Utente annulla

        self.func()

        self.mock_dialogs.FileSelect.assert_called_once_with(
            'Salva con nome...',
            '*.odt',
            1,
            startPath=None,
            defaultName='documento.odt'
        )

    def test_salvataggio_con_estensione_odt_e_storeAsURL(self):
        oDoc = MagicMock()
        selection = MagicMock()
        selection.supportsService.return_value = True
        oDoc.getCurrentSelection.return_value = selection
        oDoc.getURL.return_value = "file:///tmp/lavoro.ods"
        self.mock_leeno_utils.getDocument.return_value = oDoc

        self.mock_uno.fileUrlToSystemPath.return_value = "/tmp/lavoro.ods"
        self.mock_dialogs.FileSelect.return_value = "/tmp/esportato"  # Senza estensione

        mock_desktop = MagicMock()
        mock_writer_doc = MagicMock()
        mock_desktop.loadComponentFromURL.return_value = mock_writer_doc
        self.mock_leeno_utils.getDesktop.return_value = mock_desktop

        mock_text = MagicMock()
        mock_writer_doc.Text = mock_text

        # Mock rows and cols
        rows = MagicMock()
        rows.getCount.return_value = 0
        selection.getRows.return_value = rows
        cols = MagicMock()
        cols.getCount.return_value = 0
        selection.getColumns.return_value = cols

        self.mock_uno.systemPathToFileUrl.return_value = "file:///tmp/esportato.odt"

        self.func()

        self.mock_uno.systemPathToFileUrl.assert_called_with("/tmp/esportato.odt")
        mock_writer_doc.storeAsURL.assert_called_once()
        mock_writer_doc.storeToURL.assert_not_called()
        self.mock_dialogs.Info.assert_called_once_with(
            Title='Informazione',
            Text='File creato con successo:\n/tmp/esportato.odt'
        )


if __name__ == "__main__":
    unittest.main()
