import unittest
import sys
import os
from unittest.mock import MagicMock

# Add module paths
pythonpath_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'python', 'pythonpath')
if pythonpath_dir not in sys.path:
    sys.path.insert(0, pythonpath_dir)

# Mock com.sun.star hierarchy before imports
com = MagicMock()
sys.modules['com'] = com
sys.modules['com.sun'] = com.sun
sys.modules['com.sun.star'] = com.sun.star
sys.modules['com.sun.star.table'] = com.sun.star.table
sys.modules['com.sun.star.sheet'] = com.sun.star.sheet
sys.modules['com.sun.star.sheet.GeneralFunction'] = MagicMock()
sys.modules['com.sun.star.sheet.CellFlags'] = MagicMock()
sys.modules['com.sun.star.beans'] = com.sun.star.beans
sys.modules['com.sun.star.lang'] = com.sun.star.lang

class MockCell:
    def __init__(self, string_val="", num_val=0.0, cell_style=""):
        self.String = string_val
        self.Value = num_val
        self.CellStyle = cell_style
        self.Formula = ""
        self.Type = MagicMock()
        self.Type.value = "STRING" if string_val else ("VALUE" if num_val != 0.0 else "EMPTY")

class MockRange:
    def __init__(self, start_row, end_row):
        self.RangeAddress = MagicMock()
        self.RangeAddress.StartRow = start_row
        self.RangeAddress.EndRow = end_row
        self.CellBackColor = None

class MockSheet:
    def __init__(self, name="CONTABILITA"):
        self.Name = name
        self.cells = {}
        self.ranges = {}
        self._protected = False
        self.Rows = MagicMock()
        self.Rows.Count = 1000

    def isProtected(self):
        return self._protected

    def unprotect(self, pwd=""):
        self._protected = False

    def protect(self, pwd=""):
        self._protected = True

    def getCellByPosition(self, col, row):
        key = (col, row)
        if key not in self.cells:
            self.cells[key] = MockCell()
        return self.cells[key]

    def getCellRangeByPosition(self, c1, r1, c2, r2):
        key = (c1, r1, c2, r2)
        if key not in self.ranges:
            self.ranges[key] = MockRange(r1, r2)
        return self.ranges[key]


# Set up mocks in sys.modules
mock_mods = [
    'uno', 'unohelper',
    'LeenoUtils', 'SheetUtils', 'LeenoSheetUtils', 'LeenoComputo',
    'LeenoGlobals', 'LeenoSettings', 'LeenoDispatcher', 'Dialogs', 'LeenoDialogs', 'pyleeno', 'Calendario'
]
for m in mock_mods:
    if m not in sys.modules:
        sys.modules[m] = MagicMock()

import LeenoContab


class TestPartiteProvvisorie(unittest.TestCase):
    def setUp(self):
        self.sheet = MockSheet('CONTABILITA')
        self.doc = MagicMock()
        self.doc.getSheets().getByName.return_value = self.sheet
        self.doc.Sheets.getByName.return_value = self.sheet

        sys.modules['LeenoUtils'].getDocument.return_value = self.doc
        LeenoContab.LeenoUtils.getDocument.return_value = self.doc
        sys.modules['LeenoGlobals'].getGlobalVar.return_value = 1
        sys.modules['LeenoSheetUtils'].cercaPartenza.return_value = (0, 0, '')

    def test_no_suspended_partite(self):
        sys.modules['LeenoSheetUtils'].cercaUltimaVoce.return_value = 0
        LeenoContab.LeenoSheetUtils.cercaUltimaVoce.return_value = 0
        count = LeenoContab.annulla_partite_provvisorie_sospese()
        self.assertEqual(count, 0)

    def test_fully_offset_partita(self):
        # Setup item 1 (provisional) at rows 0-4
        self.sheet.getCellByPosition(0, 0).CellStyle = 'Comp Start Attributo'
        self.sheet.getCellByPosition(0, 1).String = '1'
        self.sheet.getCellByPosition(1, 1).String = 'NP.01'
        self.sheet.getCellByPosition(2, 2).String = 'PARTITA PROVVISORIA'
        self.sheet.getCellByPosition(9, 2).Value = 100.0
        self.sheet.getCellByPosition(0, 4).CellStyle = 'Comp End Attributo'
        self.sheet.getCellByPosition(9, 4).Value = 100.0

        # Setup item 2 (storno) at rows 5-9
        self.sheet.getCellByPosition(0, 5).CellStyle = 'Comp Start Attributo'
        self.sheet.getCellByPosition(0, 6).String = '2'
        self.sheet.getCellByPosition(1, 6).String = 'NP.01'
        self.sheet.getCellByPosition(2, 7).String = 'DETRAE PARTITA PROVVISORIA'
        self.sheet.getCellByPosition(2, 8).String = '- vedi voce n.1 - art. NP.01'
        self.sheet.getCellByPosition(9, 8).Value = -100.0
        self.sheet.getCellByPosition(0, 9).CellStyle = 'Comp End Attributo'
        self.sheet.getCellByPosition(9, 9).Value = -100.0

        sys.modules['LeenoSheetUtils'].cercaUltimaVoce.return_value = 9
        LeenoContab.LeenoSheetUtils.cercaUltimaVoce.return_value = 9

        def mock_circoscrive(sheet, row):
            if row <= 4:
                return MockRange(0, 4)
            elif row <= 9:
                return MockRange(5, 9)
            return None

        sys.modules['LeenoComputo'].circoscriveVoceComputo.side_effect = mock_circoscrive
        LeenoContab.LeenoComputo.circoscriveVoceComputo.side_effect = mock_circoscrive

        count = LeenoContab.annulla_partite_provvisorie_sospese()
        self.assertEqual(count, 0)

    def test_suspended_partita_created(self):
        mock_insert = MagicMock()
        LeenoContab.insertVoceContabilita = mock_insert

        # Setup item 1 (provisional) at rows 0-4
        self.sheet.getCellByPosition(0, 0).CellStyle = 'Comp Start Attributo'
        self.sheet.getCellByPosition(0, 1).String = '1'
        self.sheet.getCellByPosition(1, 1).String = 'NP.01'
        self.sheet.getCellByPosition(2, 2).String = 'PARTITA PROVVISORIA'
        self.sheet.getCellByPosition(9, 2).Value = 100.0
        self.sheet.getCellByPosition(0, 4).CellStyle = 'Comp End Attributo'
        self.sheet.getCellByPosition(9, 4).Value = 100.0

        last_row_state = [4]

        def mock_cerca_ultima(sheet):
            return last_row_state[0]

        def mock_insert_impl(lrow=0, arg=1, cod=None):
            last_row_state[0] = 9

        mock_insert.side_effect = mock_insert_impl
        sys.modules['LeenoSheetUtils'].cercaUltimaVoce.side_effect = mock_cerca_ultima
        LeenoContab.LeenoSheetUtils.cercaUltimaVoce.side_effect = mock_cerca_ultima

        def mock_circoscrive(sheet, row):
            if row <= 4:
                return MockRange(0, 4)
            else:
                return MockRange(5, 9)

        sys.modules['LeenoComputo'].circoscriveVoceComputo.side_effect = mock_circoscrive
        LeenoContab.LeenoComputo.circoscriveVoceComputo.side_effect = mock_circoscrive

        count = LeenoContab.annulla_partite_provvisorie_sospese()
        self.assertEqual(count, 1)

        # Verify that insertVoceContabilita was called
        mock_insert.assert_called_with(lrow=4, arg=1, cod='NP.01')

        # Verify new item measure rows
        # Row 1 of measures (r1 = 5 + 2 = 7)
        self.assertEqual(self.sheet.getCellByPosition(2, 7).String, 'DETRAE PARTITA PROVVISORIA')
        # Verify invertiUnSegno was invoked for row r2 = 8
        sys.modules['LeenoSheetUtils'].invertiUnSegno.assert_called_with(self.sheet, 8)
        # Verify adattaAltezzaRiga was called for newly inserted entry
        sys.modules['LeenoSheetUtils'].adattaAltezzaRiga.assert_called_with(self.sheet, all=False, lrow=5)

    def test_partially_offset_partita(self):
        mock_insert = MagicMock()
        LeenoContab.insertVoceContabilita = mock_insert

        # Setup item 1 (provisional Q=100)
        self.sheet.getCellByPosition(0, 0).CellStyle = 'Comp Start Attributo'
        self.sheet.getCellByPosition(0, 1).String = '1'
        self.sheet.getCellByPosition(1, 1).String = 'NP.01'
        self.sheet.getCellByPosition(2, 2).String = 'PARTITA PROVVISORIA'
        self.sheet.getCellByPosition(9, 2).Value = 100.0
        self.sheet.getCellByPosition(0, 4).CellStyle = 'Comp End Attributo'
        self.sheet.getCellByPosition(9, 4).Value = 100.0

        # Setup item 2 (storno Q=-60)
        self.sheet.getCellByPosition(0, 5).CellStyle = 'Comp Start Attributo'
        self.sheet.getCellByPosition(0, 6).String = '2'
        self.sheet.getCellByPosition(1, 6).String = 'NP.01'
        self.sheet.getCellByPosition(2, 7).String = 'DETRAE PARTITA PROVVISORIA'
        self.sheet.getCellByPosition(2, 8).String = '- vedi voce n.1 - art. NP.01'
        self.sheet.getCellByPosition(9, 8).Value = -60.0
        self.sheet.getCellByPosition(0, 9).CellStyle = 'Comp End Attributo'
        self.sheet.getCellByPosition(9, 9).Value = -60.0

        last_row_state = [9]

        def mock_cerca_ultima(sheet):
            return last_row_state[0]

        def mock_insert_impl(lrow=0, arg=1, cod=None):
            last_row_state[0] = 14

        mock_insert.side_effect = mock_insert_impl
        sys.modules['LeenoSheetUtils'].cercaUltimaVoce.side_effect = mock_cerca_ultima
        LeenoContab.LeenoSheetUtils.cercaUltimaVoce.side_effect = mock_cerca_ultima

        def mock_circoscrive(sheet, row):
            if row <= 4:
                return MockRange(0, 4)
            elif row <= 9:
                return MockRange(5, 9)
            else:
                return MockRange(10, 14)

        sys.modules['LeenoComputo'].circoscriveVoceComputo.side_effect = mock_circoscrive
        LeenoContab.LeenoComputo.circoscriveVoceComputo.side_effect = mock_circoscrive

        count = LeenoContab.annulla_partite_provvisorie_sospese()
        self.assertEqual(count, 1)

        self.assertEqual(self.sheet.getCellByPosition(4, 13).Value, 40.0)
        sys.modules['LeenoSheetUtils'].invertiUnSegno.assert_called_with(self.sheet, 13)

    def test_multiple_suspended_same_code_grouped(self):
        mock_insert = MagicMock()
        LeenoContab.insertVoceContabilita = mock_insert

        # Setup item 1 (provisional Q=100, cod=NP.01)
        self.sheet.getCellByPosition(0, 0).CellStyle = 'Comp Start Attributo'
        self.sheet.getCellByPosition(0, 1).String = '1'
        self.sheet.getCellByPosition(1, 1).String = 'NP.01'
        self.sheet.getCellByPosition(2, 2).String = 'PARTITA PROVVISORIA 1'
        self.sheet.getCellByPosition(9, 2).Value = 100.0
        self.sheet.getCellByPosition(0, 4).CellStyle = 'Comp End Attributo'
        self.sheet.getCellByPosition(9, 4).Value = 100.0

        # Setup item 2 (provisional Q=50, cod=NP.01)
        self.sheet.getCellByPosition(0, 5).CellStyle = 'Comp Start Attributo'
        self.sheet.getCellByPosition(0, 6).String = '2'
        self.sheet.getCellByPosition(1, 6).String = 'NP.01'
        self.sheet.getCellByPosition(2, 7).String = 'PARTITA PROVVISORIA 2'
        self.sheet.getCellByPosition(9, 7).Value = 50.0
        self.sheet.getCellByPosition(0, 9).CellStyle = 'Comp End Attributo'
        self.sheet.getCellByPosition(9, 9).Value = 50.0

        last_row_state = [9]

        def mock_cerca_ultima(sheet):
            return last_row_state[0]

        def mock_insert_impl(lrow=0, arg=1, cod=None):
            last_row_state[0] = 15

        mock_insert.side_effect = mock_insert_impl
        sys.modules['LeenoSheetUtils'].cercaUltimaVoce.side_effect = mock_cerca_ultima
        LeenoContab.LeenoSheetUtils.cercaUltimaVoce.side_effect = mock_cerca_ultima

        def mock_circoscrive(sheet, row):
            if row <= 4:
                return MockRange(0, 4)
            elif row <= 9:
                return MockRange(5, 9)
            else:
                return MockRange(10, 15)

        sys.modules['LeenoComputo'].circoscriveVoceComputo.side_effect = mock_circoscrive
        LeenoContab.LeenoComputo.circoscriveVoceComputo.side_effect = mock_circoscrive

        count = LeenoContab.annulla_partite_provvisorie_sospese()
        self.assertEqual(count, 2)

        # Only ONE insertVoceContabilita call for NP.01
        mock_insert.assert_called_once_with(lrow=9, arg=1, cod='NP.01')


if __name__ == '__main__':
    unittest.main()
