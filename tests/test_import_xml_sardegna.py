########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

import os
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src', 'Ultimus.oxt', 'python', 'pythonpath')))


class TestImportXmlSardegna(unittest.TestCase):

    def setUp(self):
        self.mock_pyleeno = MagicMock()
        self.mock_pyleeno.clean_text.side_effect = lambda x: x.strip() if isinstance(x, str) else x

        self.mock_modules = {
            'uno': MagicMock(),
            'unohelper': MagicMock(),
            'com': MagicMock(),
            'com.sun.star': MagicMock(),
            'com.sun.star.beans': MagicMock(),
            'com.sun.star.sheet': MagicMock(),
            'Dialogs': MagicMock(),
            'LeenoDialogs': MagicMock(),
            'pyleeno': self.mock_pyleeno,
            'SheetUtils': MagicMock(),
            'LeenoSheetUtils': MagicMock(),
            'DocUtils': MagicMock(),
            'LeenoUtils': MagicMock(),
        }
        self.patcher = patch.dict(sys.modules, self.mock_modules)
        self.patcher.start()

        import LeenoImport_XmlSardegna
        self.parser = LeenoImport_XmlSardegna

    def tearDown(self):
        self.patcher.stop()
        for mod in ['LeenoImport_XmlSardegna', 'LeenoImport_XmlToscana']:
            if mod in sys.modules:
                del sys.modules[mod]

    def test_parseXML_sar24_standard(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <dataroot>
            <SAR24_prezzario_articoli>
                <cod>SAR24_01.02.003</cod>
                <famiglia>Opere Generali</famiglia>
                <capitolo>Scavi e Demolizioni</capitolo>
                <sottocapitolo>Scavi a sezione aperta</sottocapitolo>
                <descrizione>Scavo con escavatore su terreno ordinario</descrizione>
                <um>m3</um>
                <prezzo_complessivo>15,50</prezzo_complessivo>
                <sicurezza>1,25</sicurezza>
            </SAR24_prezzario_articoli>
            <SAR24_prezzario_articoli>
                <cod>SAR24_01.02.004</cod>
                <famiglia>Opere Generali</famiglia>
                <capitolo>Scavi e Demolizioni</capitolo>
                <sottocapitolo>Scavi a sezione aperta</sottocapitolo>
                <descrizione>Scavi a sezione aperta - con mezzo meccanico compatto</descrizione>
                <um>m3</um>
                <prezzo>22,00</prezzo>
                <sicurezza></sicurezza>
            </SAR24_prezzario_articoli>
        </dataroot>
        """
        res = self.parser.parseXML(xml_data, "Titolo Default")

        self.assertEqual(res['titolo'], "Elenco prezzi - Regione Sardegna - anno 2024")
        self.assertIn("SAR24_01", res['superCategorie'])
        self.assertEqual(res['superCategorie']["SAR24_01"], "Opere Generali - Scavi e Demolizioni")

        self.assertIn("SAR24_01.02", res['categorie'])
        self.assertEqual(res['categorie']["SAR24_01.02"], "Scavi a sezione aperta")

        self.assertEqual(len(res['articoli']), 2)
        art1 = res['articoli']["SAR24_01.02.003"]
        self.assertEqual(art1['codice'], "SAR24_01.02.003")
        self.assertEqual(art1['um'], "m3")
        self.assertEqual(art1['prezzo'], 15.5)
        self.assertEqual(art1['sicurezza'], 1.25)
        self.assertIn("Scavo con escavatore su terreno ordinario", art1['desc'])

        art2 = res['articoli']["SAR24_01.02.004"]
        self.assertEqual(art2['prezzo'], 22.0)
        self.assertEqual(art2['sicurezza'], '')

    def test_parseXML_sar24_empty_cod_ignored(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <dataroot>
            <SAR24_prezzario_articoli>
                <cod></cod>
                <descrizione>Articolo senza codice</descrizione>
            </SAR24_prezzario_articoli>
        </dataroot>
        """
        res = self.parser.parseXML(xml_data, "Titolo Default")
        self.assertEqual(len(res['articoli']), 0)

    def test_parseXML_fallback_toscana(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <Prezzario>
            <Livello1><Codice>01</Codice><Descrizione>Opere Edili</Descrizione></Livello1>
            <Livello2><Codice>01.01</Codice><Descrizione>Demolizioni</Descrizione></Livello2>
            <Articolo>
                <Codice>01.01.001</Codice>
                <DescrizioneBreve>Demolizione muratura</DescrizioneBreve>
                <UnitaMisura>m3</UnitaMisura>
                <Prezzo>50,00</Prezzo>
                <Sicurezza>5,00</Sicurezza>
                <Manodopera>20,00</Manodopera>
            </Articolo>
        </Prezzario>
        """
        with patch('LeenoImport_XmlToscana.parseXML') as mock_toscana:
            mock_toscana.return_value = {
                'titolo': 'Prezzario Toscana Fallback',
                'superCategorie': {},
                'categorie': {},
                'articoli': {}
            }
            res = self.parser.parseXML(xml_data, "Titolo Toscana")
            mock_toscana.assert_called_once_with(xml_data, "Titolo Toscana")
            self.assertEqual(res['titolo'], 'Prezzario Toscana Fallback')


if __name__ == '__main__':
    unittest.main()
