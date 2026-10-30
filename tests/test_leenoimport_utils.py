########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

import os
import sys
import unittest
from unittest import mock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src', 'Ultimus.oxt', 'python', 'pythonpath'))

# Setup mocks before importing LeenoImport
MOCK_MODULES = [
    'uno',
    'unohelper',
    'com',
    'com.sun',
    'com.sun.star',
    'com.sun.star.beans',
    'LeenoUtils',
    'pyleeno',
    'LeenoDialogs',
    'SheetUtils',
    'LeenoSheetUtils',
    'Dialogs',
    'DocUtils',
    'LeenoImport_XmlSix',
    'LeenoImport_XmlToscana',
    'LeenoImport_XmlSardegna',
    'LeenoImport_XmlLiguria',
    'LeenoImport_XmlVeneto',
    'LeenoImport_XmlBasilicata',
    'LeenoImport_XmlLombardia',
]


class TestLeenoImportUtils(unittest.TestCase):

    def setUp(self):
        self.mocks = {m: mock.MagicMock() for m in MOCK_MODULES}
        # Assign distinct mock functions for parser submodules
        self.mocks['LeenoImport_XmlSix'].parseXML = mock.MagicMock(name='Six_parseXML')
        self.mocks['LeenoImport_XmlToscana'].parseXML = mock.MagicMock(name='Toscana_parseXML')
        self.mocks['LeenoImport_XmlSardegna'].parseXML = mock.MagicMock(name='Sardegna_parseXML')
        self.mocks['LeenoImport_XmlLiguria'].parseXML = mock.MagicMock(name='Liguria_parseXML')
        self.mocks['LeenoImport_XmlVeneto'].parseXML = mock.MagicMock(name='Veneto_parseXML')
        self.mocks['LeenoImport_XmlBasilicata'].parseXML = mock.MagicMock(name='Basilicata_parseXML')
        self.mocks['LeenoImport_XmlLombardia'].parseXML = mock.MagicMock(name='Lombardia_parseXML')
        self.mocks['LeenoImport_XmlLombardia'].parseXML1 = mock.MagicMock(name='Lombardia_parseXML1')

        self.patcher = mock.patch.dict(sys.modules, self.mocks)
        self.patcher.start()

        import LeenoImport
        self.LeenoImport = LeenoImport

    def tearDown(self):
        self.patcher.stop()
        sys.modules.pop('LeenoImport', None)

    # --- Tests for fixParagraphSize ---

    def test_fixParagraphSize_riga_singola(self):
        txt = 'Descrizione su una sola riga'
        res = self.LeenoImport.fixParagraphSize(txt)
        self.assertEqual(res, txt)

    def test_fixParagraphSize_due_righe_sotto_soglia(self):
        txt = 'Prima riga\nSeconda riga'
        res = self.LeenoImport.fixParagraphSize(txt)
        # minLen = 130; len('Prima riga') + len('Seconda riga') = 10 + 12 = 22
        # dl = 130 - 22 = 108 spazi aggiunti a spl0
        spl = res.split('\n')
        self.assertEqual(len(spl), 2)
        self.assertEqual(spl[0], 'Prima riga' + ' ' * 108)
        self.assertEqual(spl[1], 'Seconda riga')

    def test_fixParagraphSize_due_righe_sopra_soglia(self):
        spl0 = 'a' * 100
        spl1 = 'b' * 40
        txt = spl0 + '\n' + spl1
        res = self.LeenoImport.fixParagraphSize(txt)
        self.assertEqual(res, txt)

    def test_fixParagraphSize_piu_di_due_righe(self):
        txt_sotto = 'Riga 1\nRiga 2\nRiga 3\nRiga 4'
        res_sotto = self.LeenoImport.fixParagraphSize(txt_sotto)
        spl_sotto = res_sotto.split('\n')
        self.assertEqual(len(spl_sotto), 4)
        self.assertTrue(spl_sotto[0].startswith('Riga 1'))
        self.assertEqual(spl_sotto[1], 'Riga 2')
        self.assertEqual(spl_sotto[2], 'Riga 3')
        self.assertEqual(spl_sotto[3], 'Riga 4')

        txt_sopra = 'a' * 100 + '\n' + 'b' * 40 + '\nRiga 3\nRiga 4'
        res_sopra = self.LeenoImport.fixParagraphSize(txt_sopra)
        self.assertEqual(res_sopra, txt_sopra)

    # --- Tests for stripXMLNamespaces ---

    def test_stripXMLNamespaces_con_namespace(self):
        xml_data = '<ns:root xmlns:ns="http://example.com/ns"><ns:child>test</ns:child></ns:root>'
        root = self.LeenoImport.stripXMLNamespaces(xml_data)
        self.assertEqual(root.tag, 'root')
        self.assertEqual(root[0].tag, 'child')
        self.assertEqual(root[0].text, 'test')

    def test_stripXMLNamespaces_senza_namespace(self):
        xml_data = '<root><child>test</child></root>'
        root = self.LeenoImport.stripXMLNamespaces(xml_data)
        self.assertEqual(root.tag, 'root')
        self.assertEqual(root[0].tag, 'child')
        self.assertEqual(root[0].text, 'test')

    # --- Tests for isBasilicataXml ---

    def test_isBasilicataXml_formato_storico(self):
        xml_text = '<pdf>Prezzario_Regione_Basilicata_2020.pdf</pdf>'
        self.assertTrue(self.LeenoImport.isBasilicataXml(xml_text))

    def test_isBasilicataXml_formato_2026(self):
        xml_text = '<Prezzario><anno>2026</anno><capitoli></capitoli><sottovoci></sottovoci></Prezzario>'
        self.assertTrue(self.LeenoImport.isBasilicataXml(xml_text))

    def test_isBasilicataXml_stringa_non_basilicata(self):
        xml_text = '<Prezzario><anno>2026</anno></Prezzario>'
        self.assertFalse(self.LeenoImport.isBasilicataXml(xml_text))

    def test_isBasilicataXml_caso_parziale_senza_sottovoci(self):
        xml_text = '<Prezzario><anno>2026</anno><capitoli></capitoli></Prezzario>'
        self.assertFalse(self.LeenoImport.isBasilicataXml(xml_text))

    # --- Tests for findXmlParser ---

    def test_findXmlParser_tutti_i_pattern(self):
        test_cases = [
            ('xmlns="six.xsd"', self.mocks['LeenoImport_XmlSix'].parseXML),
            ('autore="Regione Toscana"', self.mocks['LeenoImport_XmlToscana'].parseXML),
            ('autore="Regione Calabria"', self.mocks['LeenoImport_XmlToscana'].parseXML),
            ('autore="Regione Campania"', self.mocks['LeenoImport_XmlToscana'].parseXML),
            ('autore="Regione Piemonte"', self.mocks['LeenoImport_XmlToscana'].parseXML),
            ('autore="Regione Sardegna"', self.mocks['LeenoImport_XmlSardegna'].parseXML),
            ('xsi:noNamespaceSchemaLocation="SAR24', self.mocks['LeenoImport_XmlSardegna'].parseXML),
            ('autore="Regione Liguria"', self.mocks['LeenoImport_XmlLiguria'].parseXML),
            ('rks=', self.mocks['LeenoImport_XmlVeneto'].parseXML),
            ('settore cod="VEN26-01"', self.mocks['LeenoImport_XmlVeneto'].parseXML),
            ('<pdf>Prezzario_Regione_Basilicata', self.mocks['LeenoImport_XmlBasilicata'].parseXML),
            ('<autore>Regione Lombardia', self.mocks['LeenoImport_XmlLombardia'].parseXML),
            ('<autore>LOM', self.mocks['LeenoImport_XmlLombardia'].parseXML),
            ('xsi:noNamespaceSchemaLocation="Parte', self.mocks['LeenoImport_XmlLombardia'].parseXML1),
            ('urn:schemas-microsoft-com:officedata', self.mocks['LeenoImport_XmlLombardia'].parseXML1),
        ]

        for pattern, expected_parser in test_cases:
            with self.subTest(pattern=pattern):
                parser = self.LeenoImport.findXmlParser(pattern)
                self.assertEqual(parser, expected_parser)

    def test_findXmlParser_testo_sconosciuto(self):
        parser = self.LeenoImport.findXmlParser('testo completamente sconosciuto')
        self.assertIsNone(parser)


if __name__ == '__main__':
    unittest.main()
