import unittest
import sys
from pathlib import Path

# Add src path for module imports
src_path = Path(__file__).resolve().parent.parent / 'src' / 'Ultimus.oxt' / 'python' / 'pythonpath'
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))


class MockCell:
    def __init__(self, value=0.0, string="", cell_style="", cell_back_color=0):
        self.Value = value
        self.String = string
        self.CellStyle = cell_style
        self.CellBackColor = cell_back_color
        self.HARDATTR = 1

    def clearContents(self, flags):
        self.CellBackColor = 0


class MockRangeAddress:
    def __init__(self, sheet=0, start_col=0, start_row=0, end_col=0, end_row=0):
        self.Sheet = sheet
        self.StartColumn = start_col
        self.StartRow = start_row
        self.EndColumn = end_col
        self.EndRow = end_row


class MockRows:
    def __init__(self, count):
        self.Count = count
        self.IsVisible = True


class MockColumns:
    def __init__(self, count):
        self.Count = count
        self.IsVisible = True


class MockCellRange:
    def __init__(self, sheet, start_col, start_row, end_col, end_row):
        self._sheet = sheet
        self.StartColumn = start_col
        self.StartRow = start_row
        self.EndColumn = end_col
        self.EndRow = end_row
        self.RangeAddress = MockRangeAddress(sheet.RangeAddress.Sheet, start_col, start_row, end_col, end_row)
        self.Rows = MockRows(end_row - start_row + 1)
        self.Columns = MockColumns(end_col - start_col + 1)

    def getCellByPosition(self, rel_col, rel_row):
        abs_col = self.StartColumn + rel_col
        abs_row = self.StartRow + rel_row
        return self._sheet.getCellByPosition(abs_col, abs_row)

    def getDataArray(self):
        res = []
        for r in range(self.StartRow, self.EndRow + 1):
            row_vals = []
            for c in range(self.StartColumn, self.EndColumn + 1):
                cell = self._sheet.getCellByPosition(c, r)
                val = cell.String if cell.String != "" else cell.Value
                row_vals.append(val)
            res.append(tuple(row_vals))
        return tuple(res)

    def clearContents(self, flags):
        for r in range(self.StartRow, self.EndRow + 1):
            for c in range(self.StartColumn, self.EndColumn + 1):
                self._sheet.getCellByPosition(c, r).clearContents(flags)


class MockSheet:
    def __init__(self, name, sheet_idx=0, rows_count=1000):
        self.Name = name
        self.RangeAddress = MockRangeAddress(sheet=sheet_idx, start_col=0, start_row=0, end_col=50, end_row=rows_count)
        self.Rows = MockRows(rows_count)
        self._cells = {}
        self.grouped_ranges = []

    def set_cell(self, col, row, value=0.0, string="", cell_style=""):
        self._cells[(col, row)] = MockCell(value=value, string=string, cell_style=cell_style)

    def getCellByPosition(self, col, row):
        if (col, row) not in self._cells:
            self._cells[(col, row)] = MockCell()
        return self._cells[(col, row)]

    def getCellRangeByPosition(self, col1, row1, col2, row2):
        return MockCellRange(self, col1, row1, col2, row2)

    def clearOutline(self):
        self.grouped_ranges.clear()

    def group(self, cell_range_addr, flag):
        self.grouped_ranges.append((cell_range_addr.StartRow, cell_range_addr.EndRow, flag))


class TestFiltraCodice(unittest.TestCase):

    def test_filtra_codice_logic_structures(self):
        sheet = MockSheet("COMPUTO", sheet_idx=0)
        # Construct item 1 (Code A01)
        sheet.set_cell(0, 5, cell_style='Comp Start Attributo')
        sheet.set_cell(1, 6, string='A01')
        sheet.set_cell(0, 7, cell_style='Comp End Attributo')

        # Construct item 2 (Code A02)
        sheet.set_cell(0, 8, cell_style='Comp Start Attributo')
        sheet.set_cell(1, 9, string='A02')
        sheet.set_cell(0, 10, cell_style='Comp End Attributo')

        fine = 11
        rng_col_a = sheet.getCellRangeByPosition(0, 0, 0, fine)
        col_a_styles = [rng_col_a.getCellByPosition(0, i).CellStyle for i in range(fine + 1)]
        rng_col_b_data = sheet.getCellRangeByPosition(1, 0, 1, fine).getDataArray()

        self.assertEqual(col_a_styles[5], 'Comp Start Attributo')
        self.assertEqual(col_a_styles[7], 'Comp End Attributo')
        self.assertEqual(rng_col_b_data[6][0], 'A01')
        self.assertEqual(rng_col_b_data[9][0], 'A02')

    def test_shift_click_backward_search(self):
        sheet = MockSheet("COMPUTO", sheet_idx=0)
        sheet.set_cell(0, 10, cell_style='comp progress')
        sheet.set_cell(0, 20, cell_style='comp progress')

        stili_validi = {'comp progress', 'comp 10 s'}
        search_row = 25

        col_a_range = sheet.getCellRangeByPosition(0, 0, 0, search_row)
        col_a_styles = [col_a_range.getCellByPosition(0, i).CellStyle for i in range(search_row + 1)]
        found_row = -1
        for idx in range(search_row, -1, -1):
            if col_a_styles[idx] in stili_validi:
                found_row = idx
                break

        self.assertEqual(found_row, 20)


if __name__ == '__main__':
    unittest.main()
