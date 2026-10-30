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


class TestImportXmlBasilicata(unittest.TestCase):

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
        import LeenoImport_XmlBasilicata
        self.LeenoImport = LeenoImport
        self.LeenoImport_XmlBasilicata = LeenoImport_XmlBasilicata

    def tearDown(self):
        self.patcher.stop()

    def test_parseXML_formato_storico_pdf(self):
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <Prezzario>
            <pdf>Prezzario_Regione_Basilicata_2025.pdf</pdf>
            <capitoli>
                <Capitolo>
                    <codice>01</codice>
                    <descrizione>Opere edili</descrizione>
                    <categorie>
                        <Categoria>
                            <codice>01</codice>
                            <descrizione>Demolizioni</descrizione>
                            <voci>
                                <Voce>
                                    <codice>001</codice>
                                    <descrizione>Demolizione di muratura</descrizione>
                                    <sottovoci>
                                        <Sottovoce>
                                            <codice>a</codice>
                                            <descrizione>in mattoni pieni</descrizione>
                                            <unitaMisura><codice>m3</codice></unitaMisura>
                                            <prezzo>45.50</prezzo>
                                            <manodopera>30.0</manodopera>
                                        </Sottovoce>
                                    </sottovoci>
                                </Voce>
                            </voci>
                        </Categoria>
                    </categorie>
                </Capitolo>
            </capitoli>
        </Prezzario>"""

        res = self.LeenoImport_XmlBasilicata.parseXML(xml_data)

        self.assertEqual(res['titolo'], 'Prezzario Regione Basilicata 2025')
        self.assertEqual(res['superCategorie'], {'01': 'Opere edili'})
        self.assertEqual(res['categorie'], {'01.01': 'Demolizioni'})

        self.assertIn('01.01.001.a', res['articoli'])
        art = res['articoli']['01.01.001.a']
        self.assertEqual(art['codice'], '01.01.001.a')
        self.assertEqual(art['desc'], 'Demolizione di muratura\n- in mattoni pieni')
        self.assertEqual(art['um'], 'm3')
        self.assertEqual(art['prezzo'], 45.50)
        self.assertEqual(art['mdo'], 0.30)
        self.assertEqual(art['sicurezza'], '')

    def test_parseXML_formato_2026_anno_e_voce_diretta(self):
        # Formato 2026 con <anno> e voce senza sottovoci (articolo diretto)
        xml_data = """<?xml version="1.0" encoding="utf-8"?>
        <Prezzario>
            <anno>2026</anno>
            <capitoli>
                <Capitolo>
                    <codice>02</codice>
                    <descrizione>Infrastrutture</descrizione>
                    <categorie>
                        <Categoria>
                            <codice>05</codice>
                            <descrizione>Scavi e tombinature</descrizione>
                            <voci>
                                <Voce>
                                    <codice>010</codice>
                                    <descrizione>Scavo a sezione obbligata in terreno sciolto</descrizione>
                                    <unitaMisura><codice>m3</codice></unitaMisura>
                                    <prezzo>18.20</prezzo>
                                    <manodopera>15.0</manodopera>
                                </Voce>
                            </voci>
                        </Categoria>
                    </categorie>
                </Capitolo>
            </capitoli>
        </Prezzario>"""

        res = self.LeenoImport_XmlBasilicata.parseXML(xml_data)

        self.assertEqual(res['titolo'], 'Elenco prezzi - Regione Basilicata - anno 2026')
        self.assertEqual(res['superCategorie'], {'02': 'Infrastrutture'})
        self.assertEqual(res['categorie'], {'02.05': 'Scavi e tombinature'})

        self.assertIn('02.05.010', res['articoli'])
        art = res['articoli']['02.05.010']
        self.assertEqual(art['codice'], '02.05.010')
        self.assertEqual(art['desc'], 'Scavo a sezione obbligata in terreno sciolto')
        self.assertEqual(art['um'], 'm3')
        self.assertEqual(art['prezzo'], 18.20)
        self.assertEqual(art['mdo'], 0.15)

    def test_parseXML_casi_limite_e_speciali(self):
        # Default title fallback e gestione eccezione descrizione/manodopera
        xml_data = """<Prezzario>
            <capitoli>
                <Capitolo>
                    <codice>03</codice>
                    <descrizione>Impianti</descrizione>
                    <categorie>
                        <Categoria>
                            <codice>01</codice>
                            <descrizione>Elettrici</descrizione>
                            <voci>
                                <Voce>
                                    <codice>001</codice>
                                    <descrizione>Cavo unipolare con isolamento speciale</descrizione>
                                    <sottovoci>
                                        <Sottovoce>
                                            <codice>a</codice>
                                            <unitaMisura><codice>m</codice></unitaMisura>
                                            <prezzo>3.50</prezzo>
                                        </Sottovoce>
                                    </sottovoci>
                                </Voce>
                            </voci>
                        </Categoria>
                    </categorie>
                </Capitolo>
            </capitoli>
        </Prezzario>"""

        res = self.LeenoImport_XmlBasilicata.parseXML(xml_data, defaultTitle="Prezzario Titolo Default")
        self.assertEqual(res['titolo'], "Prezzario Titolo Default")
        self.assertIn('03.01.001.a', res['articoli'])
        art = res['articoli']['03.01.001.a']
        self.assertEqual(art['desc'], 'Cavo unipolare con isolamento speciale')
        self.assertEqual(art['mdo'], '')


if __name__ == '__main__':
    unittest.main()
