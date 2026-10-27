import datetime
import json
import os
import shutil
import tempfile
import unittest
from unittest.mock import MagicMock
import sys

# Mock UNO modules if not present
if 'uno' not in sys.modules:
    sys.modules['uno'] = MagicMock()
if 'unohelper' not in sys.modules:
    sys.modules['unohelper'] = MagicMock()

com = MagicMock()
sys.modules['com'] = com
sys.modules['com.sun'] = com.sun
sys.modules['com.sun.star'] = com.sun.star
sys.modules['com.sun.star.beans'] = com.sun.star.beans
sys.modules['com.sun.star.xml'] = com.sun.star.xml
sys.modules['com.sun.star.util'] = com.sun.star.util
sys.modules['com.sun.star.sheet'] = com.sun.star.sheet

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src', 'Ultimus.oxt', 'python', 'pythonpath'))

import LeenoGiornaleImport


class MockCell:
    def __init__(self, value=0.0, string="", formula=""):
        self.Value = value
        self.String = string
        self._formula = formula

    def getFormula(self):
        return self._formula

    def setFormula(self, formula):
        self._formula = formula


class MockRows:
    def __init__(self, sheet):
        self._sheet = sheet

    def insertByIndex(self, index, count):
        # Shift rows down
        new_cells = {}
        for (col, r), cell in self._sheet._cells.items():
            if r >= index:
                new_cells[(col, r + count)] = cell
            else:
                new_cells[(col, r)] = cell
        self._sheet._cells = new_cells


class MockSheet:
    def __init__(self):
        self._cells = {}
        self.Rows = MockRows(self)

    def getRows(self):
        return self.Rows

    def getCellByPosition(self, col, row):
        if (col, row) not in self._cells:
            self._cells[(col, row)] = MockCell()
        return self._cells[(col, row)]

    def getCellRangeByPosition(self, start_col, start_row, end_col, end_row):
        mock_range = MagicMock()
        # Return data array for column A
        data = []
        max_r = max([r for (_, r) in self._cells.keys()], default=0)
        for r in range(start_row, max_r + 1):
            row_data = []
            for c in range(start_col, end_col + 1):
                cell = self.getCellByPosition(c, r)
                row_data.append(cell.String or cell.getFormula())
            data.append(row_data)
        mock_range.getDataArray.return_value = data
        return mock_range


class TestGiornaleImport(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.doc_dir = os.path.join(self.temp_dir, 'doc')
        os.makedirs(self.doc_dir, exist_ok=True)
        self.doc_path = os.path.join(self.doc_dir, 'giornale.ods')
        with open(self.doc_path, 'w') as f:
            f.write('dummy')

        self.mock_doc = MagicMock()
        self.mock_doc.getURL.return_value = 'file://' + self.doc_path

        # Configure uno.fileUrlToSystemPath to handle system paths
        LeenoGiornaleImport.uno.fileUrlToSystemPath.side_effect = lambda url: url.replace('file://', '')

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_estrai_foto_json_con_foto_in_percorso_dir(self):
        # Setup JSON path and local photos in candidate folder
        json_dir = os.path.join(self.temp_dir, 'export')
        os.makedirs(json_dir, exist_ok=True)
        json_path = os.path.join(json_dir, 'cantiere.json')
        with open(json_path, 'w') as f:
            f.write('{}')

        photo_folder = os.path.join(json_dir, '20260930')
        os.makedirs(photo_folder, exist_ok=True)
        photo_file = os.path.join(photo_folder, 'foto1.jpg')
        with open(photo_file, 'wb') as f:
            f.write(b'JPEG_DATA_123')

        giornate = [(datetime.date(2026, 9, 30), {'annotazioni': 'Lavori di scavo'})]

        cartelle = LeenoGiornaleImport._estrai_foto(self.mock_doc, foto_per_giorno={}, percorso=json_path, giornate=giornate)

        self.assertIn(datetime.date(2026, 9, 30), cartelle)
        self.assertEqual(cartelle[datetime.date(2026, 9, 30)], 'FOTO/20260930')

        # Check file was copied to doc_dir/FOTO/20260930/foto1.jpg
        copied_file = os.path.join(self.doc_dir, 'FOTO', '20260930', 'foto1.jpg')
        self.assertTrue(os.path.exists(copied_file))
        with open(copied_file, 'rb') as f:
            self.assertEqual(f.read(), b'JPEG_DATA_123')

    def test_estrai_foto_gia_presenti_in_doc_dir(self):
        # Setup photos already present in doc_dir/FOTO/20260930/
        doc_photo_dir = os.path.join(self.doc_dir, 'FOTO', '20260930')
        os.makedirs(doc_photo_dir, exist_ok=True)
        with open(os.path.join(doc_photo_dir, 'foto_esistente.jpg'), 'wb') as f:
            f.write(b'EXISTING_PHOTO')

        json_path = os.path.join(self.temp_dir, 'cantiere.json')
        giornate = [(datetime.date(2026, 9, 30), {'annotazioni': 'Verifica'})]

        cartelle = LeenoGiornaleImport._estrai_foto(self.mock_doc, foto_per_giorno={}, percorso=json_path, giornate=giornate)

        self.assertIn(datetime.date(2026, 9, 30), cartelle)
        self.assertEqual(cartelle[datetime.date(2026, 9, 30)], 'FOTO/20260930')

    def test_estrai_foto_nessuna_foto(self):
        json_path = os.path.join(self.temp_dir, 'cantiere.json')
        giornate = [(datetime.date(2026, 9, 30), {'annotazioni': 'Note'})]

        cartelle = LeenoGiornaleImport._estrai_foto(self.mock_doc, foto_per_giorno={}, percorso=json_path, giornate=giornate)

        self.assertEqual(cartelle, {})

    def test_estrai_foto_documento_non_salvato_con_foto(self):
        mock_doc_unsaved = MagicMock()
        mock_doc_unsaved.getURL.return_value = ''

        json_dir = os.path.join(self.temp_dir, 'export')
        os.makedirs(json_dir, exist_ok=True)
        json_path = os.path.join(json_dir, 'cantiere.json')
        photo_folder = os.path.join(json_dir, '20260930')
        os.makedirs(photo_folder, exist_ok=True)
        with open(os.path.join(photo_folder, 'foto1.jpg'), 'wb') as f:
            f.write(b'DATA')

        giornate = [(datetime.date(2026, 9, 30), {'annotazioni': 'Test'})]

        with self.assertRaises(ValueError) as cm:
            LeenoGiornaleImport._estrai_foto(mock_doc_unsaved, foto_per_giorno={}, percorso=json_path, giornate=giornate)
        self.assertIn('Salva il documento prima di importare le foto', str(cm.exception))

    def test_estrai_foto_zip_input(self):
        foto_per_giorno = {
            datetime.date(2026, 9, 30): [('zip_photo.jpg', b'ZIP_PHOTO_BYTES')]
        }
        cartelle = LeenoGiornaleImport._estrai_foto(self.mock_doc, foto_per_giorno=foto_per_giorno)

        self.assertIn(datetime.date(2026, 9, 30), cartelle)
        self.assertEqual(cartelle[datetime.date(2026, 9, 30)], 'FOTO/20260930')

        copied_file = os.path.join(self.doc_dir, 'FOTO', '20260930', 'zip_photo.jpg')
        self.assertTrue(os.path.exists(copied_file))
        with open(copied_file, 'rb') as f:
            self.assertEqual(f.read(), b'ZIP_PHOTO_BYTES')

    def test_estrai_foto_deduplicazione_e_rinominazione(self):
        # Existing file with different content
        dest_dir = os.path.join(self.doc_dir, 'FOTO', '20260930')
        os.makedirs(dest_dir, exist_ok=True)
        with open(os.path.join(dest_dir, 'foto1.jpg'), 'wb') as f:
            f.write(b'OLD_CONTENT')

        # New photo with same name 'foto1.jpg' but new content
        foto_per_giorno = {
            datetime.date(2026, 9, 30): [('foto1.jpg', b'NEW_CONTENT')]
        }

        cartelle = LeenoGiornaleImport._estrai_foto(self.mock_doc, foto_per_giorno=foto_per_giorno)
        self.assertEqual(cartelle[datetime.date(2026, 9, 30)], 'FOTO/20260930')

        # 'foto1.jpg' remains OLD_CONTENT
        with open(os.path.join(dest_dir, 'foto1.jpg'), 'rb') as f:
            self.assertEqual(f.read(), b'OLD_CONTENT')

        # 'foto1_2.jpg' created with NEW_CONTENT
        self.assertTrue(os.path.exists(os.path.join(dest_dir, 'foto1_2.jpg')))
        with open(os.path.join(dest_dir, 'foto1_2.jpg'), 'rb') as f:
            self.assertEqual(f.read(), b'NEW_CONTENT')

    def test_assicura_link_foto_inserimento_e_aggiornamento(self):
        sheet = MockSheet()
        sheet.getCellByPosition(0, 0).String = 'Data: 30/09/2026'
        sheet.getCellByPosition(0, 1).String = 'Meteo:'
        sheet.getCellByPosition(0, 2).String = 'Sole'
        sheet.getCellByPosition(0, 3).String = 'Annotazioni, attività svolte:'
        sheet.getCellByPosition(0, 4).String = 'Sopralluogo effettuato'
        sheet.getCellByPosition(0, 5).String = 'Qualifica e n. operai:'

        # Call _assicura_link_foto at riga_data = 0
        LeenoGiornaleImport._assicura_link_foto(sheet, 0, 'FOTO/20260930')

        # Check hyperlink inserted at row 5 (riga_testo + 1)
        cell_link = sheet.getCellByPosition(0, 5)
        expected_formula = '=HYPERLINK("FOTO/20260930/";"Apri ↗")'
        self.assertEqual(cell_link.getFormula(), expected_formula)
        self.assertEqual(sheet.getCellByPosition(0, 6).String, 'Qualifica e n. operai:')

        # Call again (re-import)
        LeenoGiornaleImport._assicura_link_foto(sheet, 0, 'FOTO/20260930')

        # Should update in-place without inserting extra rows
        self.assertEqual(sheet.getCellByPosition(0, 5).getFormula(), expected_formula)
        self.assertEqual(sheet.getCellByPosition(0, 6).String, 'Qualifica e n. operai:')


if __name__ == '__main__':
    unittest.main()
