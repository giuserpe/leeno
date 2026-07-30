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

if __name__ == '__main__':
    unittest.main()
