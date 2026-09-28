import unittest
import ast
import re
import os
from pathlib import Path


class TestHyperlinkClassification(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        pyleeno_path = Path(__file__).parent.parent / 'src' / 'Ultimus.oxt' / 'python' / 'pythonpath' / 'pyleeno.py'
        with open(pyleeno_path, 'r', encoding='utf-8') as f:
            tree = ast.parse(f.read())

        target_node = None
        for node in tree.body:
            if isinstance(node, ast.FunctionDef) and node.name == 'classify_hyperlink_target':
                target_node = node
                break

        if target_node is None:
            raise RuntimeError("function classify_hyperlink_target not found in pyleeno.py")

        module_ast = ast.Module(body=[target_node], type_ignores=[])
        code = compile(module_ast, filename=str(pyleeno_path), mode='exec')
        namespace = {'re': re, 'os': os}
        exec(code, namespace)
        cls.classify = staticmethod(namespace['classify_hyperlink_target'])

    def test_email_addresses(self):
        cases = [
            ('mario.rossi@example.com', ('mailto:mario.rossi@example.com', '@>>')),
            ('mailto:info@leeno.org', ('mailto:info@leeno.org', '@>>')),
            ('user.name+tag@sub.domain.co.it', ('mailto:user.name+tag@sub.domain.co.it', '@>>')),
        ]
        for input_str, expected in cases:
            with self.subTest(input_str=input_str):
                self.assertEqual(self.classify(input_str), expected)

    def test_web_addresses(self):
        cases = [
            ('https://leeno.org', ('https://leeno.org', 'Apri ↗')),
            ('http://www.google.com/search?q=test', ('http://www.google.com/search?q=test', 'Apri ↗')),
            ('www.leeno.org', ('http://www.leeno.org', 'Apri ↗')),
            ('ftps://files.example.org', ('ftps://files.example.org', 'Apri ↗')),
            ('t.me/leeno_computometrico', ('https://t.me/leeno_computometrico', 'Apri ↗')),
        ]
        for input_str, expected in cases:
            with self.subTest(input_str=input_str):
                self.assertEqual(self.classify(input_str), expected)

    def test_file_and_folder_paths(self):
        cases = [
            ('C:\\Users\\Public\\Documents\\progetto.pdf', ('C:\\Users\\Public\\Documents\\progetto.pdf', 'Apri ↗')),
            ('D:/lavori/computo.ods', ('D:/lavori/computo.ods', 'Apri ↗')),
            ('C:\\ ', ('C:\\', 'Apri ↗')),
            ('\\\\NAS\\condivisa\\progetto', ('\\\\NAS\\condivisa\\progetto', 'Apri ↗')),
            ('/home/user/documenti/disegno.dwg', ('/home/user/documenti/disegno.dwg', 'Apri ↗')),
            ('/tmp/foto.png', ('/tmp/foto.png', 'Apri ↗')),
            ('file:///C:/docs/file.pdf', ('file:///C:/docs/file.pdf', 'Apri ↗')),
            ('allegati/relazione.docx', ('allegati/relazione.docx', 'Apri ↗')),
            ('disegni\\tavola1.dwg', ('disegni\\tavola1.dwg', 'Apri ↗')),
            ('schema_elettrico.pdf', ('schema_elettrico.pdf', 'Apri ↗')),
            ('computo_metrico.ods', ('computo_metrico.ods', 'Apri ↗')),
            ('foto_cantiere.jpg', ('foto_cantiere.jpg', 'Apri ↗')),
            ('modello.dwg', ('modello.dwg', 'Apri ↗')),
            ('documenti/', ('documenti/', 'Apri ↗')),
            ('C:\\Lavori\\', ('C:\\Lavori\\', 'Apri ↗')),
        ]
        for input_str, expected in cases:
            with self.subTest(input_str=input_str):
                self.assertEqual(self.classify(input_str), expected)

    def test_negative_cases(self):
        negative_cases = [
            'Meteo:',
            'Presenti/intervenuti:',
            'Rifiuto di materiali e/o manufatti:',
            'Varianti disposte, modifiche e/o aggiunte prezzi:',
            'Descrizione dei lavori di scavo',
            'Nota: verificare con la direzione lavori',
            'Ore 10:30',
            '12/05/2024',
            'Prezzo unitario: € 15,00',
            '',
            '  ',
            '=HYPERLINK("http://example.com";"Apri ↗")',
        ]
        for input_str in negative_cases:
            with self.subTest(input_str=input_str):
                self.assertIsNone(self.classify(input_str))


if __name__ == '__main__':
    unittest.main()
