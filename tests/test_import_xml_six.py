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


class TestImportXmlSix(unittest.TestCase):

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

        import LeenoImport_XmlSix
        self.LeenoImport_XmlSix = LeenoImport_XmlSix

    def tearDown(self):
        self.patcher.stop()

    def test_parseXML_struttura_sintetica_singola_lingua(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <SIX:prezzarioRoot xmlns:SIX="http://www.six.it/six">
            <prezzario przId="PREZZARIO 2024">
                <przDescrizione lingua="it" breve="Prezzario Provinciale Bolzano 2024" />
                <unitaDiMisura unitaDiMisuraId="1" simbolo="m3" />
                <listaQuotazione listaQuotazioneId="1" lqtId="2024" />
                <prodotto prdId="01.01">
                    <prdDescrizione lingua="it" breve="Opere di scavo" />
                </prodotto>
                <prodotto prdId="01.01.01" unitaDiMisuraId="1" onereSicurezza="2.0">
                    <prdQuotazione listaQuotazioneId="1" valore="28.50" />
                    <incidenzaManodopera>35</incidenzaManodopera>
                    <prdDescrizione lingua="it" breve="Scavo a sezione aperta" estesa="in terreno di qualsiasi natura" />
                </prodotto>
            </prezzario>
        </SIX:prezzarioRoot>"""

        res = self.LeenoImport_XmlSix.parseXML(xml_data, "Default")

        self.assertEqual(res['titolo'], "Prezzario Provinciale Bolzano 2024")
        self.assertIn('01.01', res['articoli'])
        self.assertIn('01.01.01', res['articoli'])

        art_base = res['articoli']['01.01']
        self.assertEqual(art_base['codice'], '01.01')
        self.assertEqual(art_base['desc'], 'Opere di scavo')
        self.assertEqual(art_base['prezzo'], '')

        art_figlio = res['articoli']['01.01.01']
        self.assertEqual(art_figlio['codice'], '01.01.01')
        self.assertEqual(art_figlio['desc'], 'Opere di scavo\nScavo a sezione aperta\nin terreno di qualsiasi natura')
        self.assertEqual(art_figlio['um'], 'm3')
        self.assertEqual(art_figlio['prezzo'], 28.50)
        self.assertEqual(art_figlio['mdo'], 0.35)
        self.assertEqual(art_figlio['sicurezza'], 0.57)

    def test_parseXML_multilingua_scelta_lingua_specifica(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <SIX:prezzarioRoot xmlns:SIX="http://www.six.it/six">
            <prezzario przId="PREZZARIO MULTI">
                <przDescrizione lingua="it" breve="Prezzario Italiano" />
                <przDescrizione lingua="de" breve="Preisliste Deutsch" />
                <unitaDiMisura unitaDiMisuraId="1" simbolo="m2" />
                <prodotto prdId="02.01.01" unitaDiMisuraId="1">
                    <prdQuotazione valore="42.00" />
                    <incidenzaManodopera>20</incidenzaManodopera>
                    <prdDescrizione lingua="it" breve="Pavimentazione interna" />
                    <prdDescrizione lingua="de" breve="Bodenbelag innen" />
                </prodotto>
            </prezzario>
        </SIX:prezzarioRoot>"""

        self.mock_modules['Dialogs'].MultiButton.return_value = 'de'
        res = self.LeenoImport_XmlSix.parseXML(xml_data, "Default")

        self.assertEqual(res['titolo'], "Preisliste Deutsch")
        self.assertEqual(res['articoli']['02.01.01']['desc'], "Bodenbelag innen")
        self.assertEqual(res['articoli']['02.01.01']['prezzo'], 42.00)

    def test_parseXML_multilingua_scelta_tutte(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <SIX:prezzarioRoot xmlns:SIX="http://www.six.it/six">
            <prezzario przId="PREZZARIO MULTI">
                <przDescrizione lingua="it" breve="Prezzario Italiano" />
                <przDescrizione lingua="de" breve="Preisliste Deutsch" />
                <unitaDiMisura unitaDiMisuraId="1" simbolo="m2" />
                <prodotto prdId="02.01.01" unitaDiMisuraId="1">
                    <prdQuotazione valore="42.00" />
                    <prdDescrizione lingua="it" breve="Pavimentazione interna" />
                    <prdDescrizione lingua="de" breve="Bodenbelag innen" />
                </prodotto>
            </prezzario>
        </SIX:prezzarioRoot>"""

        self.mock_modules['Dialogs'].MultiButton.return_value = 'tutte'
        res = self.LeenoImport_XmlSix.parseXML(xml_data, "Default")

        self.assertEqual(res['titolo'], "Prezzario Italiano\nPreisliste Deutsch")
        self.assertEqual(res['articoli']['02.01.01']['desc'], "Pavimentazione interna\n§\nBodenbelag innen")

    def test_parseXML_multilingua_annulla(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <SIX:prezzarioRoot xmlns:SIX="http://www.six.it/six">
            <prezzario przId="PREZZARIO MULTI">
                <przDescrizione lingua="it" breve="Prezzario Italiano" />
                <przDescrizione lingua="de" breve="Preisliste Deutsch" />
            </prezzario>
        </SIX:prezzarioRoot>"""

        self.mock_modules['Dialogs'].MultiButton.return_value = 'annulla'
        res = self.LeenoImport_XmlSix.parseXML(xml_data, "Default")
        self.assertIsNone(res)


if __name__ == '__main__':
    unittest.main()
