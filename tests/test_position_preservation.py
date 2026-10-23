import unittest
from unittest.mock import MagicMock, patch
import sys
import os
from pathlib import Path

# Add pythonpath to sys.path for importing LeenoUtils / LeenoGlobals
pythonpath = str(Path(__file__).parent.parent / 'src' / 'Ultimus.oxt' / 'python' / 'pythonpath')
if pythonpath not in sys.path:
    sys.path.insert(0, pythonpath)

# Mock com, uno, and unohelper if not present
if 'com' not in sys.modules:
    mock_com = MagicMock()
    sys.modules['com'] = mock_com
    sys.modules['com.sun'] = MagicMock()
    sys.modules['com.sun.star'] = MagicMock()
    sys.modules['com.sun.star.beans'] = MagicMock()

if 'uno' not in sys.modules:
    mock_uno = MagicMock()
    sys.modules['uno'] = mock_uno

if 'unohelper' not in sys.modules:
    mock_unohelper = MagicMock()
    class MockBase:
        pass
    mock_unohelper.Base = MockBase
    sys.modules['unohelper'] = mock_unohelper

# Mock dialog/module dependencies to prevent circular import chain during isolated unit testing
if 'LeenoDialogs' not in sys.modules:
    sys.modules['LeenoDialogs'] = MagicMock()
if 'Dialogs' not in sys.modules:
    sys.modules['Dialogs'] = MagicMock()
if 'pyleeno' not in sys.modules:
    sys.modules['pyleeno'] = MagicMock()

import LeenoGlobals
import LeenoUtils


class TestPositionPreservation(unittest.TestCase):

    def setUp(self):
        LeenoGlobals.setGlobalVar('ultima_posizione', None)

    def test_ripristina_posizione_no_empty_sheet_cell_ranges_select(self):
        """Test that ripristina_posizione restores cell position and does NOT select an empty SheetCellRanges."""
        mock_doc = MagicMock()
        mock_controller = MagicMock()
        mock_sheet = MagicMock()
        mock_sheets = MagicMock()
        mock_cell = MagicMock()

        mock_doc.getCurrentController.return_value = mock_controller
        mock_doc.getSheets.return_value = mock_sheets
        mock_sheets.getByIndex.return_value = mock_sheet
        mock_sheet.getCellByPosition.return_value = mock_cell

        # Position on sheet 2 (CONTABILITA), col 1, row 5
        LeenoGlobals.setGlobalVar('ultima_posizione', {
            'type': 'cell',
            'sheet': 2,
            'col': 1,
            'row': 5
        })

        with patch('LeenoUtils.getDocument', return_value=mock_doc):
            LeenoUtils.ripristina_posizione()

        # Check that sheet 2 was set active and cell selected
        mock_sheets.getByIndex.assert_called_with(2)
        mock_controller.setActiveSheet.assert_called_with(mock_sheet)
        mock_sheet.getCellByPosition.assert_called_with(1, 5)
        mock_controller.select.assert_called_with(mock_cell)

        # Ensure createInstance("com.sun.star.sheet.SheetCellRanges") was NOT called or selected
        mock_doc.createInstance.assert_not_called()

    def test_preserva_posizione_decorator(self):
        """Test that @preserva_posizione decorator preserves active sheet and selection."""
        mock_doc = MagicMock()
        mock_controller = MagicMock()
        mock_active_sheet = MagicMock()
        mock_selection = MagicMock()
        mock_cell_addr = MagicMock()
        mock_cell = MagicMock()

        mock_cell_addr.Sheet = 2
        mock_cell_addr.Column = 0
        mock_cell_addr.Row = 10
        mock_selection.supportsService.side_effect = lambda service: service == "com.sun.star.sheet.SheetCell"
        mock_selection.getCellAddress.return_value = mock_cell_addr

        mock_controller.getActiveSheet.return_value = mock_active_sheet
        mock_controller.getSelection.return_value = mock_selection
        mock_doc.getCurrentController.return_value = mock_controller
        mock_active_sheet.getCellByPosition.return_value = mock_cell

        @LeenoUtils.preserva_posizione()
        def dummy_function():
            return "ok"

        with patch('LeenoUtils.getDocument', return_value=mock_doc), \
             patch('LeenoUtils.getComponentContext', return_value=MagicMock()):
            res = dummy_function()

        self.assertEqual(res, "ok")
        mock_controller.setActiveSheet.assert_called_with(mock_active_sheet)
        mock_active_sheet.getCellByPosition.assert_called_with(0, 10)
        mock_controller.select.assert_called_with(mock_cell)


if __name__ == '__main__':
    unittest.main()
