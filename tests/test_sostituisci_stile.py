import unittest
from unittest.mock import MagicMock
import sys
import os

# Add pythonpath to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src/Ultimus.oxt/python/pythonpath')))

# Mock UNO modules to avoid import errors when running outside LibreOffice
com_mock = MagicMock()
sys.modules['com'] = com_mock
sys.modules['com.sun'] = com_mock
sys.modules['com.sun.star'] = com_mock
sys.modules['com.sun.star.beans'] = com_mock
sys.modules['uno'] = MagicMock()
sys.modules['unohelper'] = MagicMock()
sys.modules['LeenoUtils'] = MagicMock()

import LeenoFormat


class TestSostituisciStileColonna(unittest.TestCase):

    def test_col_to_index(self):
        self.assertEqual(LeenoFormat.col_to_index("A"), 0)
        self.assertEqual(LeenoFormat.col_to_index("C"), 2)
        self.assertEqual(LeenoFormat.col_to_index("Z"), 25)
        self.assertEqual(LeenoFormat.col_to_index("AA"), 26)
        self.assertEqual(LeenoFormat.col_to_index("AB"), 27)
        self.assertEqual(LeenoFormat.col_to_index("c"), 2)
        self.assertEqual(LeenoFormat.col_to_index(5), 5)
        self.assertEqual(LeenoFormat.col_to_index("2"), 2)

    def test_sostituisci_stile_colonna_by_sheet_name(self):
        mock_doc = MagicMock()
        mock_sheet = MagicMock()
        mock_col_range = MagicMock()
        mock_search_desc = MagicMock()

        mock_doc.getSheets().getByName.return_value = mock_sheet
        mock_sheet.getCellRangeByPosition.return_value = mock_col_range
        mock_col_range.createSearchDescriptor.return_value = mock_search_desc

        mock_item1 = MagicMock()
        mock_item1.getRangeAddress.return_value.StartRow = 5
        mock_item1.getRangeAddress.return_value.EndRow = 5
        mock_item1.getRangeAddress.return_value.StartColumn = 2
        mock_item1.getRangeAddress.return_value.EndColumn = 2

        mock_found = MagicMock()
        mock_found.Count = 1
        mock_found.getByIndex.return_value = mock_item1
        mock_col_range.findAll.return_value = mock_found

        count = LeenoFormat.sostituisci_stile_colonna(
            nome_foglio="CONTABILITA",
            colonna="C",
            stile_origine="Comp-Bianche sopra_R",
            stile_destinazione="Comp-Bianche sopraS",
            oDoc=mock_doc
        )

        mock_doc.getSheets().getByName.assert_called_with("CONTABILITA")
        # Colonna C is index 2
        mock_sheet.getCellRangeByPosition.assert_called()
        args = mock_sheet.getCellRangeByPosition.call_args[0]
        self.assertEqual(args[0], 2)
        self.assertEqual(args[2], 2)
        self.assertTrue(mock_search_desc.SearchStyles)
        self.assertEqual(mock_search_desc.SearchString, "Comp-Bianche sopra_R")
        self.assertEqual(mock_item1.CellStyle, "Comp-Bianche sopraS")
        self.assertEqual(count, 1)

    def test_sostituisci_stile_colonna_active_sheet_fallback(self):
        mock_doc = MagicMock()
        mock_active_sheet = MagicMock()
        mock_col_range = MagicMock()
        mock_search_desc = MagicMock()

        mock_doc.CurrentController.ActiveSheet = mock_active_sheet
        mock_active_sheet.getCellRangeByPosition.return_value = mock_col_range
        mock_col_range.createSearchDescriptor.return_value = mock_search_desc

        mock_item1 = MagicMock()
        mock_item1.getRangeAddress.return_value.StartRow = 2
        mock_item1.getRangeAddress.return_value.EndRow = 4
        mock_item1.getRangeAddress.return_value.StartColumn = 2
        mock_item1.getRangeAddress.return_value.EndColumn = 2

        mock_found = MagicMock()
        mock_found.Count = 1
        mock_found.getByIndex.return_value = mock_item1
        mock_col_range.findAll.return_value = mock_found

        count = LeenoFormat.sostituisci_stile_colonna(
            nome_foglio=None,
            colonna="C",
            stile_origine="Comp-Bianche sopra_R",
            stile_destinazione="Comp-Bianche sopraS",
            oDoc=mock_doc
        )

        self.assertEqual(mock_item1.CellStyle, "Comp-Bianche sopraS")
        self.assertEqual(count, 3)

    def test_sostituisci_stile_colonna_single_cell_found(self):
        mock_doc = MagicMock()
        mock_sheet = MagicMock()
        mock_col_range = MagicMock()
        mock_search_desc = MagicMock()

        mock_doc.getSheets().getByName.return_value = mock_sheet
        mock_sheet.getCellRangeByPosition.return_value = mock_col_range
        mock_col_range.createSearchDescriptor.return_value = mock_search_desc

        # Single cell returned directly (no Count attribute)
        mock_found = MagicMock(spec=['CellStyle', 'getRangeAddress'])
        del mock_found.Count
        mock_found.getRangeAddress.return_value.StartRow = 3
        mock_found.getRangeAddress.return_value.EndRow = 3
        mock_found.getRangeAddress.return_value.StartColumn = 2
        mock_found.getRangeAddress.return_value.EndColumn = 2

        mock_col_range.findAll.return_value = mock_found

        count = LeenoFormat.sostituisci_stile_colonna(
            nome_foglio=mock_sheet,
            colonna="C",
            stile_origine="OldStyle",
            stile_destinazione="NewStyle",
            oDoc=mock_doc
        )

        self.assertEqual(mock_found.CellStyle, "NewStyle")
        self.assertEqual(count, 1)


if __name__ == '__main__':
    unittest.main()
