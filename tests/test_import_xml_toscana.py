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


class TestImportXmlToscanaESardegna(unittest.TestCase):

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
        import LeenoImport_XmlToscana
        import LeenoImport_XmlSardegna
        self.LeenoImport = LeenoImport
        self.LeenoImport_XmlToscana = LeenoImport_XmlToscana
        self.LeenoImport_XmlSardegna = LeenoImport_XmlSardegna

    def tearDown(self):
        self.patcher.stop()

    def test_toscana_parseXML_struttura_sintetica(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <EASY:Prezzario xmlns:EASY="https://prezzariollpp.regione.toscana.it/prezzario.xsd">
            <intestazione>
                <dettaglio anno="2025" area="Toscana Centro" />
                <copyright tipo="CC-BY" descrizione="Regione Toscana" />
            </intestazione>
            <Contenuto>
                <Articolo codice="01.01.001">
                    <tipo>Opere Civili</tipo>
                    <capitolo>Movimento terra</capitolo>
                    <voce>Scavo di sbancamento</voce>
                    <articolo>in terreno vegetale</articolo>
                    <um>m3</um>
                    <prezzo>14.20</prezzo>
                    <Analisi>
                        <onerisicurezza valore="0.50" />
                        <incidenzamanodopera percentuale="22.5" />
                    </Analisi>
                </Articolo>
                <Articolo codice="01.01.002">
                    <livello1>Opere Civili</livello1>
                    <livello2>Movimento terra</livello2>
                    <livello3>Scavo a sezione obbligata</livello3>
                    <livello4>Scavo a sezione obbligata in roccia compatta con mezzi meccanici</livello4>
                    <um>m3</um>
                    <prezzo>48.00</prezzo>
                </Articolo>
            </Contenuto>
        </EASY:Prezzario>"""

        res = self.LeenoImport_XmlToscana.parseXML(xml_data)

        self.assertEqual(res['titolo'], "Elenco prezzi - Toscana Centro - anno 2025\nCopyright: CC-BY - Regione Toscana")
        self.assertEqual(res['superCategorie'], {'01': 'Opere Civili'})
        self.assertEqual(res['categorie'], {'01.01': 'Movimento terra'})

        self.assertIn('01.01.001', res['articoli'])
        art1 = res['articoli']['01.01.001']
        self.assertEqual(art1['codice'], '01.01.001')
        self.assertEqual(art1['desc'], 'Scavo di sbancamento\nin terreno vegetale')
        self.assertEqual(art1['um'], 'm3')
        self.assertEqual(art1['prezzo'], 14.20)
        self.assertEqual(art1['mdo'], 0.225)
        self.assertEqual(art1['sicurezza'], 0.50)

        self.assertIn('01.01.002', res['articoli'])
        art2 = res['articoli']['01.01.002']
        self.assertEqual(art2['desc'], 'Scavo a sezione obbligata in roccia compatta con mezzi meccanici')
        self.assertEqual(art2['prezzo'], 48.00)

    def test_toscana_parseXML_prezzo_formattazione_e_mancanza_namespace(self):
        xml_data = """<PRT:Prezzario>
            <intestazione>
                <dettaglio anno="2025" area="Firenze" />
                <copyright tipo="CC" descrizione="Uso pubblico" />
            </intestazione>
            <Contenuto>
                <Articolo codice="02.01.010">
                    <tipo>Strutture</tipo>
                    <capitolo>Calcestruzzi</capitolo>
                    <voce>Calcestruzzo A</voce>
                    <articolo>Calcestruzzo A in opera</articolo>
                    <um>m3</um>
                    <prezzo>1.150,50</prezzo>
                </Articolo>
            </Contenuto>
        </PRT:Prezzario>"""

        res = self.LeenoImport_XmlToscana.parseXML(xml_data)
        self.assertIn('02.01.010', res['articoli'])
        art = res['articoli']['02.01.010']
        self.assertEqual(art['prezzo'], 1150.50)

    def test_sardegna_parseXML_formato_SAR24(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <dataroot>
            <SAR24_prezzario_articoli>
                <cod>01.01.001</cod>
                <famiglia>Opere Generali</famiglia>
                <capitolo>Scavi</capitolo>
                <sottocapitolo>Scavo di sbancamento</sottocapitolo>
                <descrizione>Scavo di sbancamento in terreno vegetale</descrizione>
                <um>m3</um>
                <prezzo_complessivo>16,50</prezzo_complessivo>
                <sicurezza>1,20</sicurezza>
            </SAR24_prezzario_articoli>
            <SAR24_prezzario_articoli>
                <cod>01.01.002</cod>
                <famiglia>Opere Generali</famiglia>
                <capitolo>Scavi</capitolo>
                <sottocapitolo>Rinterro</sottocapitolo>
                <descrizione>Rinterro con materiale di scavo</descrizione>
                <um>m3</um>
                <prezzo>8,40</prezzo>
            </SAR24_prezzario_articoli>
        </dataroot>"""

        res = self.LeenoImport_XmlSardegna.parseXML(xml_data, defaultTitle="Sardegna 2024")

        self.assertEqual(res['titolo'], "Elenco prezzi - Regione Sardegna - anno 2024")
        self.assertEqual(res['superCategorie'], {'01': 'Opere Generali - Scavi'})
        self.assertEqual(res['categorie'], {'01.01': 'Scavo di sbancamento'})

        self.assertIn('01.01.001', res['articoli'])
        art1 = res['articoli']['01.01.001']
        self.assertEqual(art1['codice'], '01.01.001')
        self.assertEqual(art1['desc'], 'Scavo di sbancamento in terreno vegetale')
        self.assertEqual(art1['um'], 'm3')
        self.assertEqual(art1['prezzo'], 16.50)
        self.assertEqual(art1['sicurezza'], 1.20)

        self.assertIn('01.01.002', res['articoli'])
        art2 = res['articoli']['01.01.002']
        self.assertEqual(art2['desc'], 'Rinterro con materiale di scavo')
        self.assertEqual(art2['prezzo'], 8.40)

    def test_sardegna_parseXML_fallback_toscana(self):
        xml_data = """<EASY:Prezzario xmlns:EASY="https://prezzario.regione.sardegna.it/prezzario/prezzario.xsd">
            <intestazione>
                <dettaglio anno="2025" area="Sardegna" />
                <copyright tipo="CC" descrizione="Regione Autonoma della Sardegna" />
            </intestazione>
            <Contenuto>
                <Articolo codice="05.01.001">
                    <tipo>Strade</tipo>
                    <capitolo>Rilevati</capitolo>
                    <voce>Rilevato rilevante</voce>
                    <articolo>misto granulare</articolo>
                    <um>m3</um>
                    <prezzo>22.00</prezzo>
                </Articolo>
            </Contenuto>
        </EASY:Prezzario>"""

        res = self.LeenoImport_XmlSardegna.parseXML(xml_data, defaultTitle="Fallback Title")
        self.assertEqual(res['titolo'], "Elenco prezzi - Sardegna - anno 2025\nCopyright: CC - Regione Autonoma della Sardegna")
        self.assertIn('05.01.001', res['articoli'])


if __name__ == '__main__':
    unittest.main()
