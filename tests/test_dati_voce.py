import unittest
import ast
from pathlib import Path


class MockCell:
    def __init__(self, value=0.0, string="", cell_style=""):
        self.Value = value
        self.String = string
        self.CellStyle = cell_style


class MockRangeAddress:
    def __init__(self, start_row, end_row):
        self.StartRow = start_row
        self.EndRow = end_row


class MockCellRange:
    def __init__(self, start_row, end_row):
        self.RangeAddress = MockRangeAddress(start_row, end_row)


class MockRows:
    def __init__(self, count):
        self.Count = count


class MockSheet:
    def __init__(self, name, rows_count=100):
        self.Name = name
        self.Rows = MockRows(rows_count)
        self._cells = {}

    def set_cell(self, col, row, value=0.0, string="", cell_style=""):
        self._cells[(col, row)] = MockCell(value=value, string=string, cell_style=cell_style)

    def getCellByPosition(self, col, row):
        return self._cells.get((col, row), MockCell())

    def getCellRangeByPosition(self, start_col, start_row, end_col, end_row):
        return MockCellRange(start_row, end_row)


class TestDatiVoce(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        computo_path = Path(__file__).parent.parent / 'src' / 'Ultimus.oxt' / 'python' / 'pythonpath' / 'LeenoComputo.py'
        with open(computo_path, 'r', encoding='utf-8') as f:
            tree = ast.parse(f.read())

        target_node = None
        for node in tree.body:
            if isinstance(node, ast.ClassDef) and node.name == 'DatiVoce':
                target_node = node
                break

        if target_node is None:
            raise RuntimeError("class DatiVoce not found in LeenoComputo.py")

        module_ast = ast.Module(body=[target_node], type_ignores=[])
        code = compile(module_ast, filename=str(computo_path), mode='exec')
        namespace = {}
        exec(code, namespace)
        cls.DatiVoce = namespace['DatiVoce']

    def test_dati_voce_contabilita(self):
        sheet = MockSheet('CONTABILITA', 50)
        # Setup una voce in CONTABILITA da riga 5 a riga 9
        sr, er = 5, 9
        sheet.set_cell(0, 5, cell_style='Comp Start Attributo_R')
        sheet.set_cell(0, 6, string='1', cell_style='comp 10 s_R')
        sheet.set_cell(1, 6, string='01.A02.B03', cell_style='comp Art-EP_R')
        sheet.set_cell(2, 6, string='Lavori di scavo', cell_style='comp 10 s_R')
        sheet.set_cell(19, 6, value=1.0)  # Nlib
        sheet.set_cell(20, 6, value=1.0)  # Plib
        sheet.set_cell(23, 6, value=1.0)  # nSal

        sheet.set_cell(0, 7, cell_style='comp 10 s_R')
        sheet.set_cell(1, 7, string='15/03/2026', cell_style='Data_bianca')

        sheet.set_cell(0, 8, cell_style='comp 10 s_R')

        sheet.set_cell(0, 9, cell_style='Comp End Attributo_R')
        sheet.set_cell(8, 9, string='SOMMANO [m3]')
        sheet.set_cell(9, 9, value=125.5)  # quantP
        sheet.set_cell(13, 9, value=45.0)  # prezzo
        sheet.set_cell(15, 9, value=5647.5)  # importo
        sheet.set_cell(17, 9, value=10.0)  # sic
        sheet.set_cell(30, 9, value=25.0)  # mdo

        dv = self.DatiVoce(sheet, 6)

        self.assertEqual(dv.SR, 5)
        self.assertEqual(dv.ER, 9)
        self.assertEqual(dv.num, '1')
        self.assertEqual(dv.art, '01.A02.B03')
        self.assertEqual(dv.desc, 'Lavori di scavo')
        self.assertIsNotNone(dv.voce)
        self.assertEqual(len(dv.voce), 9)
        self.assertEqual(dv.voce[0], '1')
        self.assertEqual(dv.voce[1], '01.A02.B03')
        self.assertEqual(dv.voce[2], 'Lavori di scavo')
        self.assertEqual(dv.voce[3], 'm3')
        self.assertEqual(dv.voce[4], 125.5)
        self.assertEqual(dv.voce[5], 45.0)
        self.assertEqual(dv.voce[6], 5647.5)
        self.assertEqual(dv.voce[7], 10.0)
        self.assertEqual(dv.voce[8], 25.0)

        self.assertEqual(dv.SAL[1], '01.A02.B03')
        self.assertIn('01.A02.B03', dv.REG[0])

    def test_dati_voce_computo(self):
        sheet = MockSheet('COMPUTO', 50)
        sr, er = 2, 5
        sheet.set_cell(0, 2, cell_style='Comp Start Attributo')
        sheet.set_cell(0, 3, string='1', cell_style='comp progress')
        sheet.set_cell(1, 3, string='02.B03.C04', cell_style='comp Art-EP')
        sheet.set_cell(2, 3, string='Calcestruzzo armato', cell_style='comp 10 s')

        sheet.set_cell(0, 4, cell_style='comp 10 s')

        sheet.set_cell(0, 5, cell_style='Comp End Attributo')
        sheet.set_cell(8, 5, string='SOMMANO [m3]')
        sheet.set_cell(9, 5, value=50.0)
        sheet.set_cell(11, 5, value=110.0)
        sheet.set_cell(18, 5, value=5500.0)
        sheet.set_cell(17, 5, value=5.0)
        sheet.set_cell(30, 5, value=15.0)

        dv = self.DatiVoce(sheet, 3)

        self.assertEqual(dv.SR, 2)
        self.assertEqual(dv.ER, 5)
        self.assertEqual(dv.num, '1')
        self.assertEqual(dv.art, '02.B03.C04')
        self.assertEqual(dv.desc, 'Calcestruzzo armato')
        self.assertEqual(dv.voce[1], '02.B03.C04')


if __name__ == '__main__':
    unittest.main()
