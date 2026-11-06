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

    def test_riferimento_dcf_xml_well_formed(self):
        import xml.etree.ElementTree as ET
        dcf_path = os.path.join(os.path.dirname(__file__), 'data', 'riferimento.dcf')
        if not os.path.exists(dcf_path):
            self.skipTest("tests/data/riferimento.dcf non trovato")

        doc = dcf_parser.parse_dcf(dcf_path)
        raw_xml = doc.get('_raw_xml', {})
        self.assertTrue(raw_xml, "Nessun blocco XML grezzo estratto")

        expected_blocks = [
            'InfoDoc', 'DatiGenerali', 'CollectionEP',
            'CollectionVC', 'CollectionOD', 'CollectionST',
            'OpzioniInterfaccia',
        ]
        for tag in expected_blocks:
            self.assertIn(tag, raw_xml, f"Blocco {tag} mancante in _raw_xml")
            xml_text = raw_xml[tag]
            try:
                root = ET.fromstring(xml_text)
            except ET.ParseError as err:
                self.fail(f"Il blocco XML {tag} non e' ben formato: {err}")
            self.assertEqual(root.tag, tag)

    def test_riferimento_dcf_confronto_con_xpwe(self):
        dcf_path = os.path.join(os.path.dirname(__file__), 'data', 'riferimento.dcf')
        xpwe_path = os.path.join(os.path.dirname(__file__), 'data', 'riferimento.xpwe')
        if not os.path.exists(dcf_path) or not os.path.exists(xpwe_path):
            self.skipTest("riferimento.dcf o riferimento.xpwe non trovati")

        doc_dcf = dcf_parser.parse_dcf(dcf_path)
        doc_xpwe = dcf_parser.parse_xpwe(xpwe_path)

        # 1. Elenco Prezzi: 201 voci attese (id, tariffa, um, prezzi 1-5, incidenze)
        self.assertEqual(len(doc_dcf['elenco_prezzi']), 201)
        self.assertEqual(len(doc_xpwe['elenco_prezzi']), 201)
        self.assertEqual(len(doc_dcf['_ep_by_id']), 201)
        self.assertEqual(len(doc_xpwe['_ep_by_id']), 201)

        # Differenze note tra DCF e XPWE per Elenco Prezzi:
        # - IdEP=2: presente solo nel DCF come 'Voce riservata!!!' interna a PriMus,
        #   omessa dall'export XPWE standard e scartata dall'elenco prezzi.
        # - Data: l'XPWE esporta la data standard se omessa, mentre il DCF usa '30/12/1899' di default.
        for idep, x_ep in doc_xpwe['_ep_by_id'].items():
            self.assertIn(idep, doc_dcf['_ep_by_id'], f"IdEP {idep} presente in XPWE ma mancante in DCF")
            d_ep = doc_dcf['_ep_by_id'][idep]
            self.assertEqual(d_ep['id'], x_ep['id'])
            self.assertEqual(d_ep['tariffa'], x_ep['tariffa'])
            self.assertEqual(d_ep['um'], x_ep['um'])
            self.assertAlmostEqual(d_ep['prezzo'], x_ep['prezzo'], places=4)
            self.assertAlmostEqual(d_ep['prezzo2'], x_ep['prezzo2'], places=4)
            self.assertAlmostEqual(d_ep['prezzo3'], x_ep['prezzo3'], places=4)
            self.assertAlmostEqual(d_ep['prezzo4'], x_ep['prezzo4'], places=4)
            self.assertAlmostEqual(d_ep['prezzo5'], x_ep['prezzo5'], places=4)
            self.assertAlmostEqual(d_ep['inc_mdo'], x_ep['inc_mdo'], places=4)
            self.assertAlmostEqual(d_ep['inc_sic'], x_ep['inc_sic'], places=4)
            self.assertAlmostEqual(d_ep['inc_mat'], x_ep['inc_mat'], places=4)
            self.assertAlmostEqual(d_ep['inc_attr'], x_ep['inc_attr'], places=4)

        # 2. Voci di computo e misurazioni: 96 VC, 2896 righi di misura
        self.assertEqual(len(doc_dcf['computo']), 96)
        self.assertEqual(len(doc_xpwe['computo']), 96)

        total_dcf_misure = sum(len(vc['misure']) for vc in doc_dcf['computo'])
        total_xpwe_misure = sum(len(vc['misure']) for vc in doc_xpwe['computo'])
        self.assertEqual(total_dcf_misure, 2896)
        self.assertEqual(total_xpwe_misure, 2896)

        # Confronto per ogni VC: id, id_ep, quantita e misurazioni
        for d_vc, x_vc in zip(doc_dcf['computo'], doc_xpwe['computo']):
            self.assertEqual(d_vc['id'], x_vc['id'])
            self.assertEqual(d_vc['id_ep'], x_vc['id_ep'])
            self.assertAlmostEqual(d_vc['quantita'], x_vc['quantita'], places=4)
            self.assertEqual(len(d_vc['misure']), len(x_vc['misure']))
            for d_rm, x_rm in zip(d_vc['misure'], x_vc['misure']):
                self.assertEqual(d_rm['id'], x_rm['id'])
                self.assertEqual(d_rm['pu'], x_rm['pu'])
                self.assertEqual(d_rm['lu'], x_rm['lu'])
                self.assertEqual(d_rm['la'], x_rm['la'])
                self.assertEqual(d_rm['hp'], x_rm['hp'])
                self.assertAlmostEqual(d_rm['qt'], x_rm['qt'], places=4)
                self.assertEqual(d_rm['descrizione'], x_rm['descrizione'])

        # 3. Categorie e Capitoli
        # Differenze note tra DCF e XPWE:
        # - data_init: in XPWE e' valorizzata (es. '18/09/2025' o '30/12/1899'),
        #   mentre in DGDatiGenerali/ItemSpCat del DCF l'attributo DtInz e' vuoto.
        self.assertEqual(len(doc_dcf['super_categorie']), 3)
        self.assertEqual(len(doc_xpwe['super_categorie']), 3)
        for sc_id, x_sc in doc_xpwe['super_categorie'].items():
            self.assertIn(sc_id, doc_dcf['super_categorie'])
            d_sc = doc_dcf['super_categorie'][sc_id]
            self.assertEqual(d_sc['nome'], x_sc['nome'])
            self.assertEqual(d_sc['cod'], x_sc['cod'])
            self.assertEqual(d_sc['durata'], x_sc['durata'])
            self.assertAlmostEqual(d_sc['percent'], x_sc['percent'], places=4)

        self.assertEqual(doc_dcf['categorie'], {})
        self.assertEqual(doc_xpwe['categorie'], {})
        self.assertEqual(doc_dcf['sotto_categorie'], {})
        self.assertEqual(doc_xpwe['sotto_categorie'], {})
        self.assertEqual(doc_dcf['super_capitoli'], {})
        self.assertEqual(doc_xpwe['super_capitoli'], {})
        self.assertEqual(doc_dcf['capitoli'], {})
        self.assertEqual(doc_xpwe['capitoli'], {})
        self.assertEqual(doc_dcf['sotto_capitoli'], {})
        self.assertEqual(doc_xpwe['sotto_capitoli'], {})

        # 4. Analisi di prezzo: 21 schede in CollectionAN / PweEPAR
        dcf_an = [ep for ep in doc_dcf['elenco_prezzi'] if ep.get('analisi')]
        xpwe_an = [ep for ep in doc_xpwe['elenco_prezzi'] if ep.get('analisi')]
        self.assertEqual(len(dcf_an), 21)
        self.assertEqual(len(xpwe_an), 21)
        plenum_d = next(ep for ep in dcf_an if ep['tariffa'] == 'NP_Plenum')
        plenum_x = next(ep for ep in xpwe_an if ep['tariffa'] == 'NP_Plenum')
        self.assertEqual(len(plenum_d['analisi']), 8)
        self.assertEqual(len(plenum_x['analisi']), 8)
        self.assertEqual(plenum_d['analisi'][0]['des'], 'MANODOPERA')
        self.assertEqual(plenum_d['analisi'][1]['id_ep'], 152)
        self.assertEqual(plenum_x['analisi'][1]['id_ep'], 152)

    def test_riferimento_dcf_non_segnala_corruzione_falsa(self):
        dcf_path = os.path.join(os.path.dirname(__file__), 'data', 'riferimento.dcf')
        if not os.path.exists(dcf_path):
            self.skipTest("tests/data/riferimento.dcf non trovato")
        msgs = []
        orig = dcf_parser._user_msg
        dcf_parser._user_msg = lambda m: msgs.append(m)
        try:
            doc = dcf_parser.parse_dcf(dcf_path)
        finally:
            dcf_parser._user_msg = orig
        self.assertEqual(msgs, [])
        self.assertEqual(len(doc['elenco_prezzi']), 201)

        xpwe_path = os.path.join(self.temp_dir.name, 'from_dcf.xpwe')
        dcf_parser.generate_xpwe(doc, xpwe_path)
        parsed = dcf_parser.parse_xpwe(xpwe_path)
        an = [ep for ep in parsed['elenco_prezzi'] if ep.get('analisi')]
        self.assertEqual(len(an), 21)
        plenum = next(ep for ep in an if ep['tariffa'] == 'NP_Plenum')
        self.assertEqual(len(plenum['analisi']), 8)
        self.assertEqual(plenum['analisi'][1]['id_ep'], 152)


if __name__ == '__main__':
    unittest.main()
