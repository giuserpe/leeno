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


class TestImportXmlLombardia(unittest.TestCase):

    def setUp(self):
        self.mock_pyleeno = MagicMock()
        self.mock_pyleeno.clean_text.side_effect = lambda x: x

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

        import LeenoImport
        import LeenoImport_XmlLombardia
        self.LeenoImport = LeenoImport
        self.LeenoImport_XmlLombardia = LeenoImport_XmlLombardia

    def tearDown(self):
        self.patcher.stop()

    def test_parseXML1_struttura_sintetica(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <Parte3 titolo="Prezzario Lombardia Parte3.xml">
            <Voce>
                <Codice>1.1 - Cemento e Calcestruzzo</Codice>
                <Declaratoria>Opere in calcestruzzo armato e non</Declaratoria>
            </Voce>
            <Voce>
                <Codice>1.1.1</Codice>
                <Declaratoria>Calcestruzzo strutturale C25/30</Declaratoria>
                <UM>m3</UM>
                <Prezzo>120,50 €</Prezzo>
                <Rapporto_RU>18,00 €</Rapporto_RU>
            </Voce>
            <Voce>
                <Codice>1.1.2 - Sottovoce B</Codice>
                <Declaratoria>in opera per solai</Declaratoria>
                <UM>m3</UM>
                <Prezzo>1.350,00 €</Prezzo>
            </Voce>
        </Parte3>"""

        res = self.LeenoImport_XmlLombardia.parseXML1(xml_data)

        self.assertEqual(res['titolo'], "Prezzario Lombardia Parte3")

        self.assertIn('1.1.1', res['articoli'])
        art1 = res['articoli']['1.1.1']
        self.assertEqual(art1['codice'], '1.1.1')
        self.assertEqual(art1['desc'], 'Opere in calcestruzzo armato e non\nCalcestruzzo strutturale C25/30')
        self.assertEqual(art1['um'], 'm3')
        self.assertEqual(art1['prezzo'], 120.50)

        self.assertIn('1.1.2', res['articoli'])
        art2 = res['articoli']['1.1.2']
        self.assertEqual(art2['codice'], '1.1.2')
        self.assertEqual(art2['desc'], 'Opere in calcestruzzo armato e non\nin opera per solai')
        self.assertEqual(art2['prezzo'], 1350.00)

    def test_parseXML_bug_noto_ciclo_voci_header(self):
        # BUG NOTO: in parseXML, `voci = root.find('voci')` e `for voce in voci:` itera su tutti i figli
        # del nodo <voci>, incluso il sotto-nodo <voci> che contiene le informazioni di intestazione (<riferimenti_voce>).
        # Poiche' la prima iterazione tenta di accedere a `list(voce)[1]` ed estrarre attributi dell'articolo
        # (come udm_voce / unita_misura_voce), se l'intestazione non possiede tali attributi solleva KeyError o IndexError.
        xml_data = """<root>
            <voci>
                <voci>
                    <riferimenti_voce>
                        <autore>LOM</autore>
                        <anno>2025</anno>
                    </riferimenti_voce>
                </voci>
            </voci>
        </root>"""

        with self.assertRaises((IndexError, KeyError)):
            self.LeenoImport_XmlLombardia.parseXML(xml_data)


if __name__ == '__main__':
    unittest.main()
