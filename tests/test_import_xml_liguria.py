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


class TestImportXmlLiguria(unittest.TestCase):

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
        import LeenoImport_XmlLiguria
        self.LeenoImport = LeenoImport
        self.LeenoImport_XmlLiguria = LeenoImport_XmlLiguria

    def tearDown(self):
        self.patcher.stop()

    def test_parseXML_struttura_sintetica(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <Prezzario>
            <intestazione autore="Regione Liguria">
                <dettaglio anno="2025" area="Genova" />
                <copyright tipo="CC-BY" descrizione="Licenza aperta" />
            </intestazione>
            <Contenuto>
                <Articolo codice="01.01.001">
                    <tipo>Opere Edili</tipo>
                    <capitolo>Scavi</capitolo>
                    <voce>Scavo di sbancamento</voce>
                    <articolo>terreno di qualsiasi natura</articolo>
                    <um>metro cubo (m3)</um>
                    <prezzo valore="15.50" />
                    <mo>20</mo>
                    <sicurezza>1.50</sicurezza>
                </Articolo>
                <Articolo codice="01.01.002">
                    <tipo>Opere Edili</tipo>
                    <capitolo>Scavi</capitolo>
                    <voce>Scavo a sezione</voce>
                    <articolo>in roccia</articolo>
                    <um>m3</um>
                    <prezzo valore="45.00" />
                    <mo>15</mo>
                    <sicurezza>0</sicurezza>
                </Articolo>
            </Contenuto>
        </Prezzario>"""

        res = self.LeenoImport_XmlLiguria.parseXML(xml_data)

        self.assertEqual(res['titolo'], "Elenco prezzi - Genova - anno 2025\nCopyright: CC-BY - Licenza aperta")
        self.assertEqual(res['superCategorie'], {'01': 'Opere Edili'})
        self.assertEqual(res['categorie'], {'01.01': 'Scavi'})

        self.assertIn('01.01.001', res['articoli'])
        art1 = res['articoli']['01.01.001']
        self.assertEqual(art1['codice'], '01.01.001')
        self.assertEqual(art1['desc'], 'Scavo di sbancamento\n- terreno di qualsiasi natura')
        self.assertEqual(art1['um'], 'm3')
        self.assertEqual(art1['prezzo'], 15.50)
        self.assertEqual(art1['mdo'], 0.20)
        self.assertEqual(art1['sicurezza'], 1.50)

        self.assertIn('01.01.002', res['articoli'])
        art2 = res['articoli']['01.01.002']
        self.assertEqual(art2['sicurezza'], '')

    def test_parseXML_casi_limite_prezzo_formattazione(self):
        # Prezzo con separatore di migliaia (es. 1.250.00 -> 1250.00)
        xml_data = """<Prezzario>
            <intestazione autore="Regione Liguria">
                <dettaglio anno="2025" area="Savona" />
                <copyright tipo="CC" descrizione="Uso pubblico" />
            </intestazione>
            <Contenuto>
                <Articolo codice="02.01.010">
                    <tipo>Strutture</tipo>
                    <capitolo>Calcestruzzo</capitolo>
                    <voce>Fornitura e posa&#10;in opera</voce>
                    <articolo>calcestruzzo Rkc 30</articolo>
                    <um>m3</um>
                    <prezzo valore="1.250.50" />
                    <mo>10</mo>
                    <sicurezza>25.00</sicurezza>
                </Articolo>
            </Contenuto>
        </Prezzario>"""

        res = self.LeenoImport_XmlLiguria.parseXML(xml_data)
        art = res['articoli']['02.01.010']
        self.assertEqual(art['prezzo'], 1250.50)
        self.assertIn('Fornitura e posa\nin opera', art['desc'])

    def test_parseXML_codice_senza_sottolivelli(self):
        xml_data = """<Prezzario>
            <intestazione autore="Regione Liguria">
                <dettaglio anno="2025" area="Imperia" />
                <copyright tipo="CC" descrizione="Uso pubblico" />
            </intestazione>
            <Contenuto>
                <Articolo codice="03">
                    <tipo>Noleggi</tipo>
                    <capitolo>Mezzi</capitolo>
                    <voce>Escavatore</voce>
                    <articolo>orario</articolo>
                    <um>h</um>
                    <prezzo valore="65.00" />
                    <mo>100</mo>
                    <sicurezza>0</sicurezza>
                </Articolo>
            </Contenuto>
        </Prezzario>"""

        res = self.LeenoImport_XmlLiguria.parseXML(xml_data)
        self.assertEqual(res['superCategorie'], {'03': 'Noleggi'})
        self.assertEqual(res['categorie'], {'03': 'Mezzi'})
        self.assertIn('03', res['articoli'])

    def test_parseXML_bug_noto_nodo_mo_e_sicurezza_mancanti(self):
        # BUG NOTO: se i nodi <mo> o <sicurezza> sono assenti nell'XML dell'articolo,
        # articolo.find('mo') e' None e articolo.find('mo').text solleva AttributeError.
        xml_data = """<Prezzario>
            <intestazione autore="Regione Liguria">
                <dettaglio anno="2025" area="La Spezia" />
                <copyright tipo="CC" descrizione="Uso pubblico" />
            </intestazione>
            <Contenuto>
                <Articolo codice="04.01.001">
                    <tipo>Sicurezza</tipo>
                    <capitolo>Apprestamenti</capitolo>
                    <voce>Ponteggio</voce>
                    <articolo>fisso</articolo>
                    <um>m2</um>
                    <prezzo valore="12.00" />
                </Articolo>
            </Contenuto>
        </Prezzario>"""

        with self.assertRaises(AttributeError):
            self.LeenoImport_XmlLiguria.parseXML(xml_data)


if __name__ == '__main__':
    unittest.main()
