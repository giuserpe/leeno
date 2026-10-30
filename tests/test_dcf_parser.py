########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src', 'Ultimus.oxt', 'python', 'pythonpath'))

import dcf_parser


class TestDcfParserCharacterization(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_parse_xpwe_e_generate_xpwe_roundtrip(self):
        doc_data = {
            'info': {
                'oggetto': 'Lavori test',
                'committente': 'Comune Alfa',
                'comune': 'Citta Beta',
                'provincia': 'XX',
                'impresa': 'Edile Gamma',
                'perc_prezzi': 0.0,
                'parte_opera': 'Opere edili',
            },
            'super_categorie': {1: {'nome': 'SuperCat 1', 'des_est': '', 'data_init': '', 'durata': '', 'cod_fase': '', 'percent': 0.0, 'cod': '01'}},
            'categorie': {1: {'nome': 'Cat 1', 'des_est': '', 'data_init': '', 'durata': '', 'cod_fase': '', 'percent': 0.0, 'cod': '01.01'}},
            'sotto_categorie': {},
            'super_capitoli': {},
            'capitoli': {},
            'sotto_capitoli': {},
            'elenco_prezzi': [
                {
                    'id': 1001,
                    'tariffa': 'EP.01',
                    'articolo': 'ART01',
                    'des_ridotta': 'Descrizione ridotta',
                    'des_estesa': 'Descrizione estesa completa',
                    'des_breve': 'Desc breve',
                    'um': 'm2',
                    'prezzo': 12.50,
                    'prezzo2': 0.0,
                    'prezzo3': 0.0,
                    'prezzo4': 0.0,
                    'prezzo5': 0.0,
                    'inc_mdo': 0.2,
                    'inc_sic': 0.05,
                    'inc_mat': 0.0,
                    'inc_attr': 0.0,
                    'id_spcap': 0,
                    'id_cap': 0,
                    'id_sbcap': 0,
                    'flags': 0,
                    'data': '01/01/2025',
                    'adr_internet': '',
                    'tag_bim': '',
                }
            ],
            'computo': [
                {
                    'id': 1,
                    'id_ep': 1001,
                    'quantita': 10.0,
                    'data_mis': '01/01/2025',
                    'id_spcal': 1,
                    'id_cat': 1,
                    'id_sbcat': 0,
                    'cod_wbs': '',
                    'flags': 0,
                    'misure': [
                        {
                            'id': 1,
                            'descrizione': 'Misura test',
                            'pu': '2',
                            'lu': '5',
                            'la': '',
                            'hp': '',
                            'qt': 10.0,
                            'rif_voce': -2,
                            'flags': 0,
                        }
                    ],
                }
            ],
        }

        xpwe_path = os.path.join(self.temp_dir.name, 'test_output.xpwe')
        dcf_parser.generate_xpwe(doc_data, xpwe_path)
        self.assertTrue(os.path.exists(xpwe_path))

        parsed_doc = dcf_parser.parse_xpwe(xpwe_path)

        self.assertEqual(parsed_doc['formato'], 'xpwe')
        self.assertEqual(parsed_doc['info']['oggetto'], 'Lavori test')
        self.assertEqual(parsed_doc['info']['committente'], 'Comune Alfa')

        self.assertIn(1001, parsed_doc['_ep_by_id'])
        ep_parsed = parsed_doc['_ep_by_id'][1001]
        self.assertEqual(ep_parsed['tariffa'], 'EP.01')
        self.assertEqual(ep_parsed['des_estesa'], 'Descrizione estesa completa')
        self.assertEqual(ep_parsed['um'], 'm2')
        self.assertEqual(ep_parsed['prezzo'], 12.50)

        self.assertEqual(len(parsed_doc['computo']), 1)
        vc_parsed = parsed_doc['computo'][0]
        self.assertEqual(vc_parsed['id'], 1)
        self.assertEqual(vc_parsed['id_ep'], 1001)
        self.assertEqual(vc_parsed['quantita'], 10.0)
        self.assertEqual(len(vc_parsed['misure']), 1)
        self.assertEqual(vc_parsed['misure'][0]['descrizione'], 'Misura test')

    def test_parse_auto_detection(self):
        sample_xpwe = """<?xml version="1.0" encoding="utf-8"?>
        <PweDocumento>
            <PweDatiGenerali>
                <PweDGProgetto>
                    <PweDGDatiGenerali>
                        <Oggetto>Auto Detection Test</Oggetto>
                    </PweDGDatiGenerali>
                </PweDGProgetto>
            </PweDatiGenerali>
        </PweDocumento>"""

        xpwe_file = os.path.join(self.temp_dir.name, 'sample.xpwe')
        with open(xpwe_file, 'w', encoding='utf-8') as f:
            f.write(sample_xpwe)

        doc = dcf_parser.parse_auto(xpwe_file)
        self.assertEqual(doc['formato'], 'xpwe')
        self.assertEqual(doc['info']['oggetto'], 'Auto Detection Test')

        sfs_invalid = os.path.join(self.temp_dir.name, 'invalid.dcf')
        with open(sfs_invalid, 'wb') as f:
            f.write(b'NOT_AAMV')

        with self.assertRaises(ValueError) as cm:
            dcf_parser.parse_auto(sfs_invalid)
        self.assertIn("Non e' un file SFS valido", str(cm.exception))

    def test_computo_con_descrizioni(self):
        ep_item = {
            'id': 100,
            'tariffa': 'E.01',
            'des_estesa': 'Scavo di sbancamento in terreno di qualsiasi natura',
            'des_ridotta': 'Scavo sbancamento',
            'um': 'm3',
            'prezzo': 15.00,
        }
        doc = {
            'elenco_prezzi': [ep_item],
            'categorie': {1: {'nome': 'Scavi'}},
            'super_categorie': {10: {'nome': 'Opere Generali'}},
            '_ep_by_id': {100: ep_item},
            'computo': [
                {
                    'id': 1,
                    'id_ep': 100,
                    'id_cat': 1,
                    'id_spcal': 10,
                    'quantita': 20.0,
                    'data_mis': '10/01/2025',
                    'misure': [{'id': 1, 'descrizione': 'Fossa A', 'qt': 20.0}],
                }
            ],
        }

        flat_rows = dcf_parser.computo_con_descrizioni(doc)
        self.assertEqual(len(flat_rows), 1)
        row = flat_rows[0]

        self.assertEqual(row['id_vc'], 1)
        self.assertEqual(row['tariffa'], 'E.01')
        self.assertEqual(row['categoria'], 'Scavi')
        self.assertEqual(row['super_cat'], 'Opere Generali')
        self.assertEqual(row['um'], 'm3')
        self.assertEqual(row['prezzo_unit'], 15.00)
        self.assertEqual(row['quantita'], 20.0)
        self.assertEqual(row['importo'], 300.00)
        self.assertEqual(row['n_misure'], 1)


if __name__ == '__main__':
    unittest.main()
