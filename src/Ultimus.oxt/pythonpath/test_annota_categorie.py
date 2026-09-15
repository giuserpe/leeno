import unittest
from unittest.mock import MagicMock, patch
import sys
import os

class MockBase: pass
class MockXStatusListener: pass
class MockXRangeSelectionListener: pass
class MockXJobExecutor: pass

mock_unohelper = MagicMock()
mock_unohelper.Base = MockBase

com = MagicMock()
com.sun.star.frame.XStatusListener = MockXStatusListener
com.sun.star.sheet.XRangeSelectionListener = MockXRangeSelectionListener
com.sun.star.task.XJobExecutor = MockXJobExecutor

sys.modules['uno'] = MagicMock()
sys.modules['unohelper'] = mock_unohelper
sys.modules['com'] = com
sys.modules['com.sun'] = com.sun
sys.modules['com.sun.star'] = com.sun.star
sys.modules['com.sun.star.table'] = com.sun.star.table
sys.modules['com.sun.star.frame'] = com.sun.star.frame
sys.modules['com.sun.star.sheet'] = com.sun.star.sheet
sys.modules['com.sun.star.sheet.GeneralFunction'] = MagicMock()
sys.modules['com.sun.star.sheet.CellFlags'] = MagicMock()
sys.modules['com.sun.star.beans'] = com.sun.star.beans
sys.modules['com.sun.star.container'] = com.sun.star.container
sys.modules['com.sun.star.util'] = com.sun.star.util
sys.modules['com.sun.star.text'] = com.sun.star.text
sys.modules['com.sun.star.awt'] = com.sun.star.awt
sys.modules['com.sun.star.lang'] = com.sun.star.lang
sys.modules['com.sun.star.style'] = com.sun.star.style
sys.modules['com.sun.star.xml'] = com.sun.star.xml
sys.modules['com.sun.star.task'] = com.sun.star.task

# Set up mocks in sys.modules before imports to avoid circular import issues
mock_mods = [
    'LeenoUtils', 'SheetUtils', 'LeenoSheetUtils',
    'LeenoGlobals', 'LeenoSettings', 'LeenoDispatcher', 'Dialogs', 'LeenoDialogs', 'pyleeno', 'Calendario', 'LeenoConfig', 'Debug', 'undo_utils'
]
for m in mock_mods:
    if m not in sys.modules:
        sys.modules[m] = MagicMock()

def noop_decorator(*args, **kwargs):
    if len(args) == 1 and callable(args[0]):
        return args[0]
    def wrapper(func):
        return func
    return wrapper

sys.modules['LeenoUtils'].no_refresh = noop_decorator
sys.modules['undo_utils'].with_undo = noop_decorator

# Add pythonpath and python directories to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../python/pythonpath')))
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../python')))

class MockCell:
    def __init__(self, cell_style='Default', value=0, string=''):
        self.CellStyle = cell_style
        self.Value = value
        self.String = string

class MockSheet:
    def __init__(self, name, rows_data):
        self.Name = name
        self.rows_data = rows_data  # dict: (col, row) -> MockCell
        self.Rows = MagicMock()
        self.Rows.Count = 1000

    def getCellByPosition(self, col, row):
        if (col, row) not in self.rows_data:
            self.rows_data[(col, row)] = MockCell()
        return self.rows_data[(col, row)]

class MockRange:
    def __init__(self, start_row, end_row):
        self.RangeAddress = MagicMock()
        self.RangeAddress.StartRow = start_row
        self.RangeAddress.EndRow = end_row

class TestAnnotaCategorie(unittest.TestCase):

    def test_annota_categorie_voci(self):
        # Setup mock sheet data with Super Categoria, Categoria, Sotto Categoria and an item
        rows_data = {}

        # Row 10: Super Categoria "2" - "OPERE EDILI"
        rows_data[(0, 10)] = MockCell(cell_style='Livello-0-scritta')
        rows_data[(1, 10)] = MockCell(cell_style='Livello-0-scritta', string='2')
        rows_data[(2, 10)] = MockCell(cell_style='Livello-0-scritta', string='OPERE EDILI')

        # Row 15: Categoria "2.4" - "Rifacimenti"
        rows_data[(0, 15)] = MockCell(cell_style='Livello-1-scritta')
        rows_data[(1, 15)] = MockCell(cell_style='Livello-1-scritta', string='2.4')
        rows_data[(2, 15)] = MockCell(cell_style='Livello-1-scritta', string='Rifacimenti')

        # Row 20: Sotto Categoria "2.4.1" - "pavimento e rivestimenti"
        rows_data[(0, 20)] = MockCell(cell_style='livello2 valuta')
        rows_data[(1, 20)] = MockCell(cell_style='livello2 valuta', string='2.4.1')
        rows_data[(2, 20)] = MockCell(cell_style='livello2 valuta', string='pavimento e rivestimenti')

        # Item 1: rows 25 to 28
        # Row 25: Comp Start Attributo (start_row)
        rows_data[(0, 25)] = MockCell(cell_style='Comp Start Attributo')
        # Row 26: comp progress
        rows_data[(0, 26)] = MockCell(cell_style='comp progress')
        rows_data[(1, 26)] = MockCell(cell_style='comp progress', string='E.01.01')
        rows_data[(2, 26)] = MockCell(cell_style='comp progress', string='Descrizione articolo 1')
        # Row 28: Comp End Attributo (end_row)
        rows_data[(0, 28)] = MockCell(cell_style='Comp End Attributo')

        sheet = MockSheet('COMPUTO', rows_data)

        sys.modules['LeenoGlobals'].getGlobalVar.return_value = ['Comp Start Attributo', 'Comp End Attributo']

        # Import LeenoComputo module
        import LeenoComputo

        # Patch circoscriveVoceComputo and cercaUltimaVoce
        with patch.object(LeenoComputo, 'circoscriveVoceComputo', side_effect=lambda s, r: MockRange(25, 28) if r == 25 else None), \
             patch('LeenoSheetUtils.cercaUltimaVoce', return_value=30), \
             patch('LeenoUtils.getDocument') as mock_get_doc:

            mock_doc = MagicMock()
            mock_doc.CurrentController.ActiveSheet = sheet
            mock_get_doc.return_value = mock_doc

            sys.modules['LeenoConfig'].Config.return_value.read.return_value = 'True'

            LeenoComputo.annota_categorie_voci(sheet)

            # Check that start_row (row 25) was annotated:
            # Column C (index 2) should contain "OPERE EDILI | Rifacimenti | pavimento e rivestimenti"
            # Column B (index 1) should contain "2.4.1"
            self.assertEqual(sheet.getCellByPosition(2, 25).String, "OPERE EDILI | Rifacimenti | pavimento e rivestimenti")
            self.assertEqual(sheet.getCellByPosition(1, 25).String, "2.4.1")

    def test_annota_categorie_voci_disabled(self):
        rows_data = {}
        rows_data[(0, 25)] = MockCell(cell_style='Comp Start Attributo')
        sheet = MockSheet('COMPUTO', rows_data)

        import LeenoComputo

        with patch.object(LeenoComputo, 'circoscriveVoceComputo', side_effect=lambda s, r: MockRange(25, 28) if r == 25 else None), \
             patch('LeenoSheetUtils.cercaUltimaVoce', return_value=30), \
             patch('LeenoUtils.getDocument') as mock_get_doc:

            mock_doc = MagicMock()
            mock_doc.CurrentController.ActiveSheet = sheet
            mock_get_doc.return_value = mock_doc

            sys.modules['LeenoConfig'].Config.return_value.read.return_value = 'False'

            LeenoComputo.annota_categorie_voci(sheet)

            # Check that start_row was NOT annotated
            self.assertEqual(sheet.getCellByPosition(2, 25).String, "")


if __name__ == '__main__':
    unittest.main()
