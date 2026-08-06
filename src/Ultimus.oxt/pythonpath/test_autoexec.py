import unittest
import re

def rimuovi_S2_da_codice_basic(code):
    sub_pattern = re.compile(r'(Sub\s+Controlla_Esistenza_LibUltimus.*?End\s+Sub)', re.DOTALL | re.IGNORECASE)
    match = sub_pattern.search(code)
    if not match:
        return code, False

    sub_body = match.group(1)
    array_pattern = re.compile(r'(Array\s*\(\s*([^)]*)\s*\))', re.IGNORECASE | re.DOTALL)
    array_match = array_pattern.search(sub_body)
    if not array_match:
        return code, False

    full_array_expr = array_match.group(1)
    array_contents = array_match.group(2)
    elements = array_contents.split(',')
    target_found = False
    new_elements = []
    for el in elements:
        cleaned = el.strip()
        if cleaned.startswith('&quot;') and cleaned.endswith('&quot;'):
            val = cleaned[6:-6]
        elif cleaned.startswith('&apos;') and cleaned.endswith('&apos;'):
            val = cleaned[6:-6]
        else:
            val = cleaned.strip('\"\'')
        if val.upper() == 'S2':
            target_found = True
        else:
            new_elements.append(el)
    if target_found:
        new_array_contents = ','.join(new_elements)
        new_array_contents = re.sub(r',\s*,', ',', new_array_contents)
        new_array_contents = new_array_contents.strip().strip(',')
        new_array_expr = f'Array({new_array_contents})'
        new_sub_body = sub_body.replace(full_array_expr, new_array_expr)
        new_code = code.replace(match.group(1), new_sub_body)
        return new_code, True
    return code, False

class TestAutoexecS2Removal(unittest.TestCase):
    def test_removal_with_quotes(self):
        code = 'Sub Controlla_Esistenza_LibUltimus\nFor Each el In Array("S1", "S2", "S3")\nEnd Sub'
        new_code, changed = rimuovi_S2_da_codice_basic(code)
        self.assertTrue(changed)
        self.assertNotIn('"S2"', new_code)
        self.assertIn('"S1"', new_code)
        self.assertIn('"S3"', new_code)

    def test_removal_with_xml_entities(self):
        code = 'Sub Controlla_Esistenza_LibUltimus\nFor Each el In Array("S1", &quot;S2&quot;, "S3")\nEnd Sub'
        new_code, changed = rimuovi_S2_da_codice_basic(code)
        self.assertTrue(changed)
        self.assertNotIn('S2', new_code)
        self.assertIn('"S1"', new_code)
        self.assertIn('"S3"', new_code)

    def test_no_removal_if_not_present(self):
        code = 'Sub Controlla_Esistenza_LibUltimus\nFor Each el In Array("S1", "S3")\nEnd Sub'
        new_code, changed = rimuovi_S2_da_codice_basic(code)
        self.assertFalse(changed)
        self.assertEqual(code, new_code)

    def test_case_insensitivity(self):
        code = 'SUB CONTROLLA_ESISTENZA_LIBULTIMUS\nFor Each el In Array("S1", "s2", "S3")\nEND SUB'
        new_code, changed = rimuovi_S2_da_codice_basic(code)
        self.assertTrue(changed)
        self.assertNotIn('s2', new_code)
        self.assertNotIn('S2', new_code)

# --- Mocking per i test di esportazione Markdown ---
import sys
import os
from unittest.mock import MagicMock

# Aggiunge il path della cartella corrente a sys.path per consentire l'importazione
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Salva i moduli originali se presenti per non interferire con altri contesti
original_modules = {}
mock_modules = [
    'uno', 'unohelper', 'com', 'com.sun.star', 'com.sun.star.beans',
    'LeenoUtils', 'SheetUtils', 'LeenoSheetUtils', 'LeenoComputo',
    'LeenoFormat', 'LeenoGlobals', 'LeenoConfig', 'LeenoDialogs',
    'Dialogs', 'pyleeno', 'Debug'
]

for mod in mock_modules:
    if mod in sys.modules:
        original_modules[mod] = sys.modules[mod]
    sys.modules[mod] = MagicMock()


class TestMarkdownExportSplitting(unittest.TestCase):
    def test_no_split_under_3mb(self):
        from LeenoExport import split_markdown_table
        header = "| H1 | H2 |"
        delimiter = "| --- | --- |"
        rows = ["| R1 | R2 |", "| R3 | R4 |"]
        parts = split_markdown_table(header, delimiter, rows, limit_3mb=100, limit_2mb=50)
        self.assertEqual(len(parts), 1)
        expected = ("| H1 | H2 |\n| --- | --- |\n| R1 | R2 |\n| R3 | R4 |\n").encode('utf-8')
        self.assertEqual(parts[0], expected)

    def test_split_over_3mb_into_2mb_parts(self):
        from LeenoExport import split_markdown_table
        header = "| H1 | H2 |"  # 11 chars
        delimiter = "| --- | --- |"  # 13 chars
        # Prefisso: 11 + 1 + 13 + 1 = 26 byte
        rows = [
            "| R1 | R2 |",  # 11 chars + 1 newline = 12 byte
            "| R3 | R4 |",  # 12 byte
            "| R5 | R6 |",  # 12 byte
        ]
        # Con limite 3mb = 40, limite 2mb = 45.
        # Dimensione totale: 26 + 36 = 62 byte (> 40), quindi viene suddiviso.
        # Parte 1: prefisso (26) + riga 1 (12) = 38 <= 45. La riga 2 sforerebbe (38 + 12 = 50 > 45).
        # Parte 1 ha solo riga 1.
        # Parte 2: prefisso (26) + riga 2 (12) = 38 <= 45. La riga 3 sforerebbe (50 > 45).
        # Parte 2 ha solo riga 2.
        # Parte 3: prefisso (26) + riga 3 (12) = 38 <= 45.
        # Parte 3 ha solo riga 3.
        parts = split_markdown_table(header, delimiter, rows, limit_3mb=40, limit_2mb=45)
        self.assertEqual(len(parts), 3)
        self.assertEqual(parts[0], ("| H1 | H2 |\n| --- | --- |\n| R1 | R2 |\n").encode('utf-8'))
        self.assertEqual(parts[1], ("| H1 | H2 |\n| --- | --- |\n| R3 | R4 |\n").encode('utf-8'))
        self.assertEqual(parts[2], ("| H1 | H2 |\n| --- | --- |\n| R5 | R6 |\n").encode('utf-8'))

    def test_single_row_exceeding_limit(self):
        # Caso in cui una singola riga supera da sola il limite di 2Mb
        from LeenoExport import split_markdown_table
        header = "| H1 |"  # 6 chars
        delimiter = "| --- |"  # 7 chars
        # Prefisso: 6 + 1 + 7 + 1 = 15 byte
        rows = [
            "| R1_molto_lunga |",  # 17 chars + 1 newline = 18 byte
        ]
        # limite 2mb = 30 (prefisso + riga = 15 + 18 = 33 > 30)
        parts = split_markdown_table(header, delimiter, rows, limit_3mb=20, limit_2mb=30)
        self.assertEqual(len(parts), 1)
        self.assertEqual(parts[0], ("| H1 |\n| --- |\n| R1_molto_lunga |\n").encode('utf-8'))

if __name__ == '__main__':
    unittest.main()
