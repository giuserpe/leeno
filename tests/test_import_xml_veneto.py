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


class TestImportXmlVeneto(unittest.TestCase):

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
        import LeenoImport_XmlVeneto
        self.LeenoImport = LeenoImport
        self.LeenoImport_XmlVeneto = LeenoImport_XmlVeneto

    def tearDown(self):
        self.patcher.stop()

    def test_parseXML_struttura_sintetica(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <root desc="Prezzario Veneto 2025">
            <settore cod="S1" desc="Opere Edili">
                <capitolo cod="S1.C1" desc="Scavi e Rilevanti">
                    <paragrafo cod="S1.C1.P1">
                        <sint>Scavi generali</sint>
                        <estesa>Scavo di sbancamento in terreno di qualsiasi natura</estesa>
                        <prezzi>
                            <prezzo cod="ART01" umi="m3" val="12.50" man="25">Scavo di sbancamento in terreno di qualsiasi natura</prezzo>
                            <prezzo cod="ART02" umi="m3" val="18.00" man="30">Scavo a sezione obbligata</prezzo>
                        </prezzi>
                    </paragrafo>
                </capitolo>
            </settore>
        </root>"""

        res = self.LeenoImport_XmlVeneto.parseXML(xml_data)

        self.assertEqual(res['titolo'], 'Prezzario Veneto 2025')
        self.assertEqual(res['superCategorie'], {'S1': 'Opere Edili'})
        self.assertEqual(res['categorie'], {'S1.C1': 'Scavi e Rilevanti', 'S1.C1.P1': 'Scavi generali'})

        self.assertIn('ART01', res['articoli'])
        art1 = res['articoli']['ART01']
        self.assertEqual(art1['codice'], 'ART01')
        self.assertEqual(art1['desc'], 'Scavo di sbancamento in terreno di qualsiasi natura')
        self.assertEqual(art1['um'], 'm3')
        self.assertEqual(art1['prezzo'], 12.50)
        self.assertEqual(art1['mdo'], 0.25)
        self.assertEqual(art1['sicurezza'], '')

        self.assertIn('ART02', res['articoli'])
        art2 = res['articoli']['ART02']
        self.assertEqual(art2['codice'], 'ART02')
        self.assertEqual(art2['desc'], 'Scavo di sbancamento in terreno di qualsiasi natura\n- Scavo a sezione obbligata')

    def test_parseXML_casi_limite_e_caratteri_speciali(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <PRT:root xmlns:PRT="http://example.com/schema" desc="Prezzario Veneto - Note &amp; Accentati">
            <settore cod="S1" desc="Opere Speciali e Strutture">
                <capitolo cod="S1.C1" desc="Calcestruzzo ed Acciaio">
                    <paragrafo cod="S1.C1.P1">
                        <sint>Getto in opera</sint>
                        <estesa>Getto di calcestruzzo con pompa&#10;in opera per fondazioni</estesa>
                        <prezzi>
                            <prezzo cod="ART_ACC" umi="m3" val="105.50" man="15">Getto di calcestruzzo con pompa in opera per fondazioni</prezzo>
                        </prezzi>
                    </paragrafo>
                </capitolo>
            </settore>
        </PRT:root>"""

        res = self.LeenoImport_XmlVeneto.parseXML(xml_data)
        self.assertEqual(res['titolo'], 'Prezzario Veneto - Note & Accentati')
        self.assertIn('ART_ACC', res['articoli'])
        art = res['articoli']['ART_ACC']
        self.assertEqual(art['prezzo'], 105.50)
        self.assertEqual(art['mdo'], 0.15)

    def test_parseXML_codice_duplicato_sovrascrive(self):
        xml_data = """<root desc="Test Duplicati">
            <settore cod="S1" desc="Settore 1">
                <capitolo cod="S1.C1" desc="Capitolo 1">
                    <paragrafo cod="S1.C1.P1">
                        <sint>Sintesi</sint>
                        <estesa>Voce Base</estesa>
                        <prezzi>
                            <prezzo cod="DUP01" umi="m" val="10.00" man="10">Voce Base 1</prezzo>
                            <prezzo cod="DUP01" umi="m" val="20.00" man="15">Voce Base 2</prezzo>
                        </prezzi>
                    </paragrafo>
                </capitolo>
            </settore>
        </root>"""

        res = self.LeenoImport_XmlVeneto.parseXML(xml_data)
        self.assertEqual(len(res['articoli']), 1)
        self.assertEqual(res['articoli']['DUP01']['prezzo'], 20.00)

    def test_parseXML_bug_noto_mdo_persistenza(self):
        # BUG NOTO: se la tipologia e' 'Manodopera', mdo e' 1 per tutti i prezzi del paragrafo.
        # Inoltre, la variabile mdo non viene resettata tra i prezzi in prezzi[0],
        # quindi il secondo articolo eredita la percentuale man del primo se non la specifica.
        xml_data = """<root desc="Test MDO Bug">
            <settore cod="S1" desc="S1">
                <capitolo cod="S1.C1" desc="C1">
                    <paragrafo cod="S1.C1.P1">
                        <sint>P1</sint>
                        <estesa>Desc</estesa>
                        <prezzi>
                            <prezzo cod="M1" umi="h" val="30.00" man="50">M1</prezzo>
                            <prezzo cod="M2" umi="h" val="40.00">M2</prezzo>
                        </prezzi>
                    </paragrafo>
                </capitolo>
            </settore>
        </root>"""

        res = self.LeenoImport_XmlVeneto.parseXML(xml_data)
        self.assertEqual(res['articoli']['M1']['mdo'], 0.50)
        # BUG NOTO: M2 non ha attribute 'man', ma poiche' mdo non viene azzerato ad ogni ciclo del loop prezzi[0],
        # M2 eredita 0.50 da M1 invece di fallback a 0 o stringa vuota.
        self.assertEqual(res['articoli']['M2']['mdo'], 0.50)


if __name__ == '__main__':
    unittest.main()
