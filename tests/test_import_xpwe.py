########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

import os
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src', 'Ultimus.oxt', 'python', 'pythonpath')))


def fake_strall(el, n=3, pos=0):
    el = str(el)
    if pos == 0:
        return n * '0' + el
    else:
        return el + n * '0'


class TestImportXpwe(unittest.TestCase):

    def setUp(self):
        self.mock_pyleeno = MagicMock()
        self.mock_pyleeno.clean_text.side_effect = lambda x: x
        self.mock_pyleeno.strall.side_effect = fake_strall

        self.mock_dialogs = MagicMock()

        self.mock_modules = {
            'uno': MagicMock(),
            'unohelper': MagicMock(),
            'com': MagicMock(),
            'com.sun.star': MagicMock(),
            'com.sun.star.table': MagicMock(),
            'Dialogs': self.mock_dialogs,
            'LeenoDialogs': MagicMock(),
            'pyleeno': self.mock_pyleeno,
            'SheetUtils': MagicMock(),
            'LeenoSheetUtils': MagicMock(),
            'DocUtils': MagicMock(),
            'LeenoUtils': MagicMock(),
            'LeenoFormat': MagicMock(),
            'LeenoConfig': MagicMock(),
            'LeenoAnalysis': MagicMock(),
            'LeenoComputo': MagicMock(),
            'LeenoVariante': MagicMock(),
            'LeenoContab': MagicMock(),
        }
        self.patcher = patch.dict(sys.modules, self.mock_modules)
        self.patcher.start()

        import LeenoImport_XPWE
        self.IMP = LeenoImport_XPWE

    def tearDown(self):
        self.patcher.stop()
        if 'LeenoImport_XPWE' in sys.modules:
            del sys.modules['LeenoImport_XPWE']

    # -------------------------------------------------------------------------
    # 1. get_xml_text
    # -------------------------------------------------------------------------
    def test_get_xml_text_elemento_presente(self):
        root = ET.fromstring('<root><nome>Progetto Alfa</nome></root>')
        res = self.IMP.get_xml_text(root, 'nome')
        self.assertEqual(res, 'Progetto Alfa')

    def test_get_xml_text_elemento_assente(self):
        root = ET.fromstring('<root><nome>Progetto Alfa</nome></root>')
        res_default_vuoto = self.IMP.get_xml_text(root, 'descrizione')
        self.assertEqual(res_default_vuoto, '')

        res_default_custom = self.IMP.get_xml_text(root, 'descrizione', default='Non specificato')
        self.assertEqual(res_default_custom, 'Non specificato')

    def test_get_xml_text_elemento_vuoto(self):
        root = ET.fromstring('<root><vuoto></vuoto><spazi>   </spazi></root>')
        # Tag vuoto senza testo: found.text e' None -> restituisce default
        res_vuoto = self.IMP.get_xml_text(root, 'vuoto', default='N/D')
        self.assertEqual(res_vuoto, 'N/D')

        # Tag con soli spazi: restituisce gli spazi
        res_spazi = self.IMP.get_xml_text(root, 'spazi')
        self.assertEqual(res_spazi, '   ')

    def test_get_xml_text_percorso_annidato(self):
        root = ET.fromstring('<root><livello1><livello2><chiave>Valore Annidato</chiave></livello2></livello1></root>')
        res = self.IMP.get_xml_text(root, 'livello1/livello2/chiave')
        self.assertEqual(res, 'Valore Annidato')

    def test_get_xml_text_input_non_valido(self):
        # Passando None o tipo non XML, l'eccezione viene catturata e restituisce default
        res = self.IMP.get_xml_text(None, 'qualsiasi', default='Fallback')
        self.assertEqual(res, 'Fallback')

    # -------------------------------------------------------------------------
    # 2. leggiAnagraficaGenerale e leggiDatiGeneraliAnalisi
    # -------------------------------------------------------------------------
    def test_leggiAnagraficaGenerale_completa(self):
        xml_data = """<PweDatiGenerali>
            <PweDGProgetto>
                <PweDGDatiGenerali>
                    <Comune>Torino</Comune>
                    <Oggetto>Ristrutturazione Edificio Scolastico</Oggetto>
                    <Committente>Citta di Torino</Committente>
                    <Impresa>Costruzioni Generali Spa</Impresa>
                </PweDGDatiGenerali>
            </PweDGProgetto>
        </PweDatiGenerali>"""
        dati = ET.fromstring(xml_data)
        anag = self.IMP.leggiAnagraficaGenerale(dati)
        self.assertEqual(anag['comune'], 'Torino')
        self.assertEqual(anag['oggetto'], 'Ristrutturazione Edificio Scolastico')
        self.assertEqual(anag['committente'], 'Citta di Torino')
        self.assertEqual(anag['impresa'], 'Costruzioni Generali Spa')

    def test_leggiAnagraficaGenerale_vuota_o_none(self):
        anag_none = self.IMP.leggiAnagraficaGenerale(None)
        self.assertEqual(anag_none, {'comune': '', 'oggetto': '', 'committente': '', 'impresa': ''})

        dati_vuoti = ET.fromstring('<PweDatiGenerali></PweDatiGenerali>')
        anag_vuota = self.IMP.leggiAnagraficaGenerale(dati_vuoti)
        self.assertEqual(anag_vuota, {'comune': '', 'oggetto': '', 'committente': '', 'impresa': ''})

    def test_leggiDatiGeneraliAnalisi_completa(self):
        xml_data = """<PweDatiGenerali>
            <PweDGModuli>
                <PweDGAnalisi>
                    <SpeseGenerali>15,00</SpeseGenerali>
                    <UtiliImpresa>10.50</UtiliImpresa>
                    <OneriAccessoriSc>2,5</OneriAccessoriSc>
                </PweDGAnalisi>
            </PweDGModuli>
        </PweDatiGenerali>"""
        dati = ET.fromstring(xml_data)
        res = self.IMP.leggiDatiGeneraliAnalisi(dati)
        self.assertAlmostEqual(res['SpeseGenerali'], 0.15)
        self.assertAlmostEqual(res['UtiliImpresa'], 0.105)
        self.assertAlmostEqual(res['OneriAccessoriSicurezza'], 0.025)

    def test_leggiDatiGeneraliAnalisi_none_o_mancante(self):
        res_none = self.IMP.leggiDatiGeneraliAnalisi(None)
        self.assertEqual(res_none, {'SpeseGenerali': 0, 'UtiliImpresa': 0, 'OneriAccessoriSicurezza': 0})

        dati_vuoti = ET.fromstring('<PweDatiGenerali></PweDatiGenerali>')
        res_vuoto = self.IMP.leggiDatiGeneraliAnalisi(dati_vuoti)
        self.assertEqual(res_vuoto, {'SpeseGenerali': 0, 'UtiliImpresa': 0, 'OneriAccessoriSicurezza': 0})

    # -------------------------------------------------------------------------
    # 3. Capitoli e Categorie
    # -------------------------------------------------------------------------
    def test_leggi_gerarchie_capitoli_e_categorie(self):
        xml_data = """<PweDGCapitoliCategorie>
            <PweDGSuperCapitoli>
                <SCItem ID="1">
                    <Codice>SC.01</Codice>
                    <DesSintetica>Super Capitolo Edile</DesSintetica>
                    <Percentuale>60.0</Percentuale>
                </SCItem>
            </PweDGSuperCapitoli>
            <PweDGCapitoli>
                <CItem ID="10">
                    <Codice>CAP.01</Codice>
                    <DesSintetica>Opere di Fondazione</DesSintetica>
                    <Percentuale>25.0</Percentuale>
                </CItem>
                <CItem ID="11">
                    <Codice>CAP.02</Codice>
                    <DesSintetica>Nuova voce</DesSintetica>
                    <DesEstesa>Descrizione capitolo ricavata da estesa</DesEstesa>
                    <Percentuale>15.0</Percentuale>
                </CItem>
            </PweDGCapitoli>
            <PweDGSubCapitoli>
                <SBCItem ID="100">
                    <Codice>SB.01</Codice>
                    <DesSintetica>Pali e micropali</DesSintetica>
                    <Percentuale>10.0</Percentuale>
                </SBCItem>
                <SBCItem ID="101">
                    <Codice>SB.02</Codice>
                    <DesSintetica>Nuova voce</DesSintetica>
                    <DesEstesa>Sottocapitolo con descrizione estesa</DesEstesa>
                    <Percentuale>5.0</Percentuale>
                </SBCItem>
            </PweDGSubCapitoli>
            <PweDGSuperCategorie>
                <SCatItem ID="2">
                    <Codice>CAT.S1</Codice>
                    <DesSintetica>Super Categoria Strutture</DesSintetica>
                    <Percentuale>40.0</Percentuale>
                </SCatItem>
            </PweDGSuperCategorie>
            <PweDGCategorie>
                <CatItem ID="20">
                    <Codice>CAT.01</Codice>
                    <DesSintetica>Cemento Armato</DesSintetica>
                    <Percentuale>30.0</Percentuale>
                </CatItem>
            </PweDGCategorie>
            <PweDGSubCategorie>
                <SBCatItem ID="200">
                    <Codice>SUBCAT.01</Codice>
                    <DesSintetica>Travi e Pilastri</DesSintetica>
                    <Percentuale>10.0</Percentuale>
                </SBCatItem>
            </PweDGSubCategorie>
        </PweDGCapitoliCategorie>"""
        capcat = ET.fromstring(xml_data)

        # SuperCapitoli
        sc = self.IMP.leggiSuperCapitoli(capcat)
        self.assertEqual(len(sc), 1)
        self.assertEqual(sc[0]['id_sc'], '1')
        self.assertEqual(sc[0]['codice'], 'SC.01')
        self.assertEqual(sc[0]['dessintetica'], 'Super Capitolo Edile')
        self.assertEqual(sc[0]['percentuale'], '60.0')

        # Capitoli: verifica sostituzione "Nuova voce" con "DesEstesa"
        c = self.IMP.leggiCapitoli(capcat)
        self.assertEqual(len(c), 2)
        self.assertEqual(c[0]['dessintetica'], 'Opere di Fondazione')
        self.assertEqual(c[1]['dessintetica'], 'Descrizione capitolo ricavata da estesa')

        # SottoCapitoli: verifica sostituzione "Nuova voce" con "DesEstesa"
        sbc = self.IMP.leggiSottoCapitoli(capcat)
        self.assertEqual(len(sbc), 2)
        self.assertEqual(sbc[0]['dessintetica'], 'Pali e micropali')
        self.assertEqual(sbc[1]['dessintetica'], 'Sottocapitolo con descrizione estesa')

        # SuperCategorie
        scat = self.IMP.leggiSuperCategorie(capcat)
        self.assertEqual(len(scat), 1)
        self.assertEqual(scat[0]['dessintetica'], 'Super Categoria Strutture')

        # Categorie
        cat = self.IMP.leggiCategorie(capcat)
        self.assertEqual(len(cat), 1)
        self.assertEqual(cat[0]['dessintetica'], 'Cemento Armato')

        # SottoCategorie
        sbcat = self.IMP.leggiSottoCategorie(capcat)
        self.assertEqual(len(sbcat), 1)
        self.assertEqual(sbcat[0]['dessintetica'], 'Travi e Pilastri')

    def test_leggiCapitoliCategorie_aggregatore_e_casi_limite(self):
        # Caso dati vuoti / None
        res_none = self.IMP.leggiCapitoliCategorie(None)
        self.assertEqual(res_none['SuperCapitoli'], [])
        self.assertEqual(res_none['Capitoli'], [])
        self.assertEqual(res_none['SottoCapitoli'], [])
        self.assertEqual(res_none['SuperCategorie'], [])
        self.assertEqual(res_none['Categorie'], [])
        self.assertEqual(res_none['SottoCategorie'], [])

        # Caso PweDatiGenerali senza nodo PweDGCapitoliCategorie
        dati_senza_capcat = ET.fromstring('<PweDatiGenerali><Altro>1</Altro></PweDatiGenerali>')
        res_senza = self.IMP.leggiCapitoliCategorie(dati_senza_capcat)
        self.assertEqual(res_senza['Capitoli'], [])

        # Caso con sezioni vuote
        dati_vuoti = ET.fromstring('<PweDatiGenerali><PweDGCapitoliCategorie></PweDGCapitoliCategorie></PweDatiGenerali>')
        res_vuoti = self.IMP.leggiCapitoliCategorie(dati_vuoti)
        self.assertEqual(res_vuoti['SuperCapitoli'], [])

    def test_leggiCapitoliCategorie_codici_duplicati_e_ordine(self):
        # I codici duplicati vengono entrambi aggiunti preservando l'ordine
        xml_data = """<PweDGCapitoliCategorie>
            <PweDGSuperCapitoli>
                <SCItem ID="1"><Codice>DUP.01</Codice><DesSintetica>Primo</DesSintetica><Percentuale>1</Percentuale></SCItem>
                <SCItem ID="2"><Codice>DUP.01</Codice><DesSintetica>Secondo</DesSintetica><Percentuale>2</Percentuale></SCItem>
                <SCItem ID="3"><Codice>DUP.02</Codice><DesSintetica>Terzo</DesSintetica><Percentuale>3</Percentuale></SCItem>
            </PweDGSuperCapitoli>
        </PweDGCapitoliCategorie>"""
        capcat = ET.fromstring(xml_data)
        sc = self.IMP.leggiSuperCapitoli(capcat)
        self.assertEqual(len(sc), 3)
        self.assertEqual([x['id_sc'] for x in sc], ['1', '2', '3'])
        self.assertEqual([x['dessintetica'] for x in sc], ['Primo', 'Secondo', 'Terzo'])

    # -------------------------------------------------------------------------
    # 4. leggiElencoPrezzi
    # -------------------------------------------------------------------------
    def test_leggiElencoPrezzi_struttura_completa(self):
        xml_data = """<PweMisurazioni>
            <PweElencoPrezzi>
                <EPItem ID="101">
                    <TipoEP>1</TipoEP>
                    <Tariffa>EP.01.001</Tariffa>
                    <Articolo>ART_SCAVO</Articolo>
                    <DesRidotta>Scavo a sezione aperta</DesRidotta>
                    <DesEstesa>Scavo di sbancamento in roccia compatta&#10;seconda riga con caratteri accentati: perimetro, quantita, societa</DesEstesa>
                    <DesBreve>Scavo breve</DesBreve>
                    <UnMisura>m3</UnMisura>
                    <Prezzo1>28,75</Prezzo1>
                    <Prezzo2>30.00</Prezzo2>
                    <Prezzo3>32.50</Prezzo3>
                    <Prezzo4>35.00</Prezzo4>
                    <Prezzo5>40.00</Prezzo5>
                    <IDSpCap>1</IDSpCap>
                    <IDCap>10</IDCap>
                    <Data>15/01/2025</Data>
                    <AdrInternet>https://example.com</AdrInternet>
                    <IncMDO>35,5</IncMDO>
                    <IncSIC>4,5</IncSIC>
                    <IncMAT>50,0</IncMAT>
                    <IncATTR>10,0</IncATTR>
                </EPItem>
                <EPItem ID="102">
                    <Tariffa>SIC.01</Tariffa>
                    <Articolo>ART_SIC</Articolo>
                    <DesEstesa>Parapetto provvisorio</DesEstesa>
                    <Flags>134217728</Flags>
                    <UnMisura>m</UnMisura>
                    <Prezzo1>15.00</Prezzo1>
                </EPItem>
                <EPItem ID="103">
                    <Tariffa>NO_PRZ.01</Tariffa>
                    <DesEstesa>Articolo senza prezzo</DesEstesa>
                    <UnMisura>cad</UnMisura>
                    <Prezzo1>0</Prezzo1>
                    <IncMDO>0</IncMDO>
                </EPItem>
            </PweElencoPrezzi>
        </PweMisurazioni>"""
        mis = ET.fromstring(xml_data)
        res = self.IMP.leggiElencoPrezzi(mis)

        self.assertEqual(len(res['DizionarioArticoli']), 3)
        self.assertEqual(len(res['ListaArticoli']), 3)
        self.assertEqual(len(res['ListaAnalisi']), 0)

        # Articolo 101
        art1 = res['DizionarioArticoli']['101']
        self.assertEqual(art1['tipoep'], '1')
        self.assertEqual(art1['tariffa'], 'EP.01.001')
        # DesBreve sovrascrive DesRidotta se presente
        self.assertEqual(art1['desridotta'], 'Scavo breve')
        self.assertEqual(art1['desbreve'], 'Scavo breve')
        self.assertIn('\nseconda riga con caratteri accentati: perimetro, quantita, societa', art1['destestesa'])
        self.assertEqual(art1['prezzo1'], 28.75)
        self.assertEqual(art1['prezzo2'], '30.00')
        self.assertEqual(art1['prezzo5'], '40.00')
        self.assertAlmostEqual(art1['IncMDO'], 0.355)
        self.assertAlmostEqual(art1['IncSIC'], 0.045)
        self.assertAlmostEqual(art1['IncMAT'], 0.50)
        self.assertAlmostEqual(art1['IncATTR'], 0.10)

        # In ListaArticoli per articolo 101: tupla (tariffa, destestesa, unmisura, IncSIC, prezzo1, IncMDO, IncMAT, IncATTR)
        self.assertEqual(res['ListaArticoli'][0], (
            'EP.01.001',
            art1['destestesa'],
            'm3',
            0.045,
            28.75,
            0.355,
            0.50,
            0.10
        ))

        # Articolo 102: Voce Della Sicurezza con Flags 134217728 -> prefisso VDS_
        art2 = res['DizionarioArticoli']['102']
        self.assertEqual(art2['tariffa'], 'VDS_SIC.01')
        self.assertEqual(res['ListaArticoli'][1][0], 'VDS_SIC.01')

        # Articolo 103: prezzo zero e incidenza zero diventano stringa vuota
        art3 = res['DizionarioArticoli']['103']
        self.assertEqual(art3['prezzo1'], '')
        self.assertEqual(art3['IncMDO'], '')

    def test_leggiElencoPrezzi_con_analisi_prezzo(self):
        xml_data = """<PweMisurazioni>
            <PweElencoPrezzi>
                <EPItem ID="200">
                    <Tariffa>AP.01</Tariffa>
                    <Articolo>ANALISI_1</Articolo>
                    <DesEstesa>Calcestruzzo strutturale Rck 30 in opera</DesEstesa>
                    <UnMisura>m3</UnMisura>
                    <Prezzo1>145,00</Prezzo1>
                    <PweEPAnalisi>
                        <PweEPAR>
                            <EPARItem>
                                <IDEP>10</IDEP>
                                <Descrizione>Calcestruzzo preconfezionato</Descrizione>
                                <Misura>m3</Misura>
                                <Qt> 1,05 </Qt>
                                <Prezzo> 95,00 </Prezzo>
                            </EPARItem>
                            <EPARItem>
                                <IDEP>11</IDEP>
                                <Descrizione>Operaio specializzato</Descrizione>
                                <Misura>h</Misura>
                                <Qt> 1,50 </Qt>
                                <Prezzo> 32,00 </Prezzo>
                            </EPARItem>
                        </PweEPAR>
                    </PweEPAnalisi>
                </EPItem>
            </PweElencoPrezzi>
        </PweMisurazioni>"""
        mis = ET.fromstring(xml_data)
        res = self.IMP.leggiElencoPrezzi(mis)

        self.assertEqual(len(res['ListaArticoli']), 0)
        self.assertEqual(len(res['ListaAnalisi']), 1)
        self.assertEqual(res['ListaTariffeAnalisi'], ['AP.01'])

        analisi = res['ListaAnalisi'][0]
        self.assertEqual(analisi[0], 'AP.01')
        self.assertEqual(analisi[1], 'Calcestruzzo strutturale Rck 30 in opera')
        self.assertEqual(analisi[2], 'm3')
        self.assertEqual(analisi[4], 145.0)

        righe_an = analisi[3]
        self.assertEqual(len(righe_an), 2)
        # La tupla interna e' (id_ep_an, an_des, an_um, an_qt, an_pr) con spazi rimossi da Qt e Prezzo
        self.assertEqual(righe_an[0], ('10', 'Calcestruzzo preconfezionato', 'm3', '1,05', '95,00'))
        self.assertEqual(righe_an[1], ('11', 'Operaio specializzato', 'h', '1,50', '32,00'))

    def test_leggiElencoPrezzi_vuoto_o_none(self):
        res_none = self.IMP.leggiElencoPrezzi(None)
        self.assertEqual(res_none['DizionarioArticoli'], {})
        self.assertEqual(res_none['ListaArticoli'], [])
        self.assertEqual(res_none['ListaAnalisi'], [])
        self.assertEqual(res_none['ListaTariffeAnalisi'], [])

        mis_vuoto = ET.fromstring('<PweMisurazioni></PweMisurazioni>')
        res_vuoto = self.IMP.leggiElencoPrezzi(mis_vuoto)
        self.assertEqual(res_vuoto['ListaArticoli'], [])

    # -------------------------------------------------------------------------
    # 5. leggiApprossimazioni
    # -------------------------------------------------------------------------
    def test_leggiApprossimazioni_standard(self):
        xml_data = """<PweDatiGenerali>
            <PweDGConfigurazione>
                <PweDGConfigNumeri>
                    <PartiUguali>9.2|0</PartiUguali>
                    <Lunghezza>9.3|0</Lunghezza>
                    <Larghezza>9.2|0</Larghezza>
                    <HPeso>9.4|0</HPeso>
                    <Quantita>10.3|1</Quantita>
                    <Prezzi>10.2|1</Prezzi>
                    <PrezziTotale>14.2|1</PrezziTotale>
                </PweDGConfigNumeri>
            </PweDGConfigurazione>
        </PweDatiGenerali>"""
        dati = ET.fromstring(xml_data)
        appr = self.IMP.leggiApprossimazioni(dati)
        self.assertEqual(appr['PartiUguali'], 2)
        self.assertEqual(appr['Lunghezza'], 3)
        self.assertEqual(appr['Larghezza'], 2)
        self.assertEqual(appr['HPeso'], 4)
        self.assertEqual(appr['Quantita'], 3)
        self.assertEqual(appr['Prezzi'], 2)
        self.assertEqual(appr['PrezziTotale'], 2)

    def test_leggiApprossimazioni_mancante_o_anomalo(self):
        self.assertEqual(self.IMP.leggiApprossimazioni(None), {})

        dati_senza_cfg = ET.fromstring('<PweDatiGenerali></PweDatiGenerali>')
        self.assertEqual(self.IMP.leggiApprossimazioni(dati_senza_cfg), {})

        # Formato senza punto o non numerico: il campo non viene inserito nel dizionario
        dati_err = ET.fromstring("""<PweDatiGenerali>
            <PweDGConfigurazione>
                <PweDGConfigNumeri>
                    <PartiUguali>NON_NUMERICO</PartiUguali>
                    <Lunghezza>9.abc|0</Lunghezza>
                    <Larghezza>9.1|0</Larghezza>
                </PweDGConfigNumeri>
            </PweDGConfigurazione>
        </PweDatiGenerali>""")
        appr_err = self.IMP.leggiApprossimazioni(dati_err)
        self.assertNotIn('PartiUguali', appr_err)
        self.assertNotIn('Lunghezza', appr_err)
        self.assertEqual(appr_err.get('Larghezza'), 1)

    # -------------------------------------------------------------------------
    # 6. leggiMisurazioni
    # -------------------------------------------------------------------------
    def test_leggiMisurazioni_singola_e_piu_righe_no_ordina(self):
        xml_data = """<PweMisurazioni>
            <PweElencoPrezzi></PweElencoPrezzi>
            <PweVociComputo>
                <VCItem ID="1">
                    <IDEP>101</IDEP>
                    <Quantita>25.00</Quantita>
                    <DataMis>20/02/2025</DataMis>
                    <Flags>0</Flags>
                    <IDSpCat>1</IDSpCat>
                    <IDCat>2</IDCat>
                    <IDSbCat>3</IDSbCat>
                    <PweRigheMisura>
                        <RGItem>
                            <Descrizione>Parete Est</Descrizione>
                            <PartiUguali>2</PartiUguali>
                            <Lunghezza>5.00</Lunghezza>
                            <Larghezza>2.50</Larghezza>
                            <HPeso></HPeso>
                            <Quantita>25.00</Quantita>
                            <Flags>0</Flags>
                            <IDVV>-2</IDVV>
                        </RGItem>
                    </PweRigheMisura>
                </VCItem>
                <VCItem ID="2">
                    <IDEP>102</IDEP>
                    <Quantita>50.00</Quantita>
                    <DataMis>21/02/2025</DataMis>
                    <Flags>0</Flags>
                    <IDSpCat>1</IDSpCat>
                    <IDCat>1</IDCat>
                    <IDSbCat>1</IDSbCat>
                    <PweRigheMisura>
                        <RGItem>
                            <Descrizione>Piano Terra</Descrizione>
                            <PartiUguali>1</PartiUguali>
                            <Lunghezza>10.00</Lunghezza>
                            <Larghezza>2.50</Larghezza>
                            <HPeso>1.00</HPeso>
                            <Quantita>25.00</Quantita>
                            <Flags>0</Flags>
                            <IDVV>-2</IDVV>
                        </RGItem>
                        <RGItem>
                            <Descrizione>Piano Primo</Descrizione>
                            <PartiUguali>1</PartiUguali>
                            <Lunghezza>10.00</Lunghezza>
                            <Larghezza>2.50</Larghezza>
                            <HPeso>1.00</HPeso>
                            <Quantita>25.00</Quantita>
                            <Flags>0</Flags>
                            <IDVV>-2</IDVV>
                        </RGItem>
                    </PweRigheMisura>
                </VCItem>
            </PweVociComputo>
        </PweMisurazioni>"""
        mis = ET.fromstring(xml_data)
        res = self.IMP.leggiMisurazioni(mis, ordina=False)

        self.assertEqual(len(res), 2)
        # Ordine preservato come nel documento originale: VCItem 1 poi VCItem 2
        self.assertEqual(res[0]['id_vc'], '1')
        self.assertEqual(res[0]['id_ep'], '101')
        self.assertEqual(len(res[0]['lista_rig']), 1)
        riga1 = res[0]['lista_rig'][0]
        self.assertEqual(riga1[0], 'Parete Est')
        self.assertEqual(riga1[3], '2')
        self.assertEqual(riga1[4], '5.00')
        self.assertEqual(riga1[5], '2.50')
        self.assertEqual(riga1[6], '')
        self.assertEqual(riga1[7], '25.00')
        self.assertEqual(riga1[8], '0')
        self.assertEqual(riga1[9], '-2')

        self.assertEqual(res[1]['id_vc'], '2')
        self.assertEqual(len(res[1]['lista_rig']), 2)
        self.assertEqual(res[1]['lista_rig'][0][0], 'Piano Terra')
        self.assertEqual(res[1]['lista_rig'][1][0], 'Piano Primo')

    def test_leggiMisurazioni_pulizia_valori_formule_e_parziali_negativi(self):
        # clean_value sostituisce r'(\d+)(\()' con r'\1*\2' (es. 2(3) -> 2*(3))
        # e mappa '0.00' oppure stringhe con doppi spazi in None
        xml_data = """<PweMisurazioni>
            <PweElencoPrezzi></PweElencoPrezzi>
            <PweVociComputo>
                <VCItem ID="10">
                    <IDEP>105</IDEP>
                    <Quantita>-15.00</Quantita>
                    <PweRigheMisura>
                        <RGItem>
                            <Descrizione>Detrazione vano porta</Descrizione>
                            <PartiUguali>2(3)</PartiUguali>
                            <Lunghezza>2.10</Lunghezza>
                            <Larghezza>0.90</Larghezza>
                            <HPeso>0.00</HPeso>
                            <Quantita>-15.00</Quantita>
                            <Flags>2</Flags>
                            <IDVV>-2</IDVV>
                        </RGItem>
                        <RGItem>
                            <Descrizione>Commento informativo</Descrizione>
                            <PartiUguali>spazi  doppi</PartiUguali>
                            <Lunghezza></Lunghezza>
                            <Larghezza></Larghezza>
                            <HPeso></HPeso>
                            <Quantita></Quantita>
                            <Flags>1</Flags>
                            <IDVV>-2</IDVV>
                        </RGItem>
                    </PweRigheMisura>
                </VCItem>
            </PweVociComputo>
        </PweMisurazioni>"""
        mis = ET.fromstring(xml_data)
        res = self.IMP.leggiMisurazioni(mis, ordina=False)

        self.assertEqual(len(res), 1)
        righe = res[0]['lista_rig']
        self.assertEqual(len(righe), 2)

        riga_neg = righe[0]
        self.assertEqual(riga_neg[0], 'Detrazione vano porta')
        self.assertEqual(riga_neg[3], '2*(3)')  # Formula normalizzata con moltiplicazione
        self.assertEqual(riga_neg[4], '2.10')
        self.assertEqual(riga_neg[5], '0.90')
        self.assertIsNone(riga_neg[6])          # '0.00' viene convertito in None da clean_value
        self.assertEqual(riga_neg[7], '-15.00')
        self.assertEqual(riga_neg[8], '2')      # Flag parziale

        riga_comm = righe[1]
        self.assertEqual(riga_comm[0], 'Commento informativo')
        self.assertIsNone(riga_comm[3])         # Stringa con doppi spazi ('spazi  doppi') diventa None
        self.assertEqual(riga_comm[8], '1')      # Flag commento

    def test_leggiMisurazioni_voce_senza_misure(self):
        xml_data = """<PweMisurazioni>
            <PweElencoPrezzi></PweElencoPrezzi>
            <PweVociComputo>
                <VCItem ID="99">
                    <IDEP>106</IDEP>
                    <Quantita>0</Quantita>
                    <DataMis>01/01/2025</DataMis>
                </VCItem>
            </PweVociComputo>
        </PweMisurazioni>"""
        mis = ET.fromstring(xml_data)
        res = self.IMP.leggiMisurazioni(mis, ordina=False)
        self.assertEqual(len(res), 1)
        self.assertEqual(res[0]['id_vc'], '99')
        self.assertEqual(res[0]['lista_rig'], [])

    def test_leggiMisurazioni_con_ordina_true(self):
        # Con ordina=True, le voci vengono ordinate in base a f"{PL.strall(idspcat)}.{PL.strall(idcat)}.{PL.strall(idsbcat)}"
        # VCItem 1 ha categorie 2.1.1
        # VCItem 2 ha categorie 1.2.1
        # VCItem 3 ha categorie 1.1.1
        # Con fake_strall(pos=0):
        # VCItem 1 -> 0002.0001.0001
        # VCItem 2 -> 0001.0002.0001
        # VCItem 3 -> 0001.0001.0001
        # Ordine atteso: VCItem 3, poi VCItem 2, poi VCItem 1
        xml_data = """<PweMisurazioni>
            <PweElencoPrezzi></PweElencoPrezzi>
            <PweVociComputo>
                <VCItem ID="1">
                    <IDEP>101</IDEP>
                    <IDSpCat>2</IDSpCat>
                    <IDCat>1</IDCat>
                    <IDSbCat>1</IDSbCat>
                </VCItem>
                <VCItem ID="2">
                    <IDEP>102</IDEP>
                    <IDSpCat>1</IDSpCat>
                    <IDCat>2</IDCat>
                    <IDSbCat>1</IDSbCat>
                </VCItem>
                <VCItem ID="3">
                    <IDEP>103</IDEP>
                    <IDSpCat>1</IDSpCat>
                    <IDCat>1</IDCat>
                    <IDSbCat>1</IDSbCat>
                </VCItem>
            </PweVociComputo>
        </PweMisurazioni>"""
        mis = ET.fromstring(xml_data)

        res_no_ord = self.IMP.leggiMisurazioni(mis, ordina=False)
        self.assertEqual([x['id_vc'] for x in res_no_ord], ['1', '2', '3'])

        res_ord = self.IMP.leggiMisurazioni(mis, ordina=True)
        self.assertEqual([x['id_vc'] for x in res_ord], ['3', '2', '1'])

    def test_leggiMisurazioni_input_invalido_o_troppo_corto(self):
        # Se misurazioni e' None o ha meno di 2 figli (manca PweVociComputo), restituisce lista vuota
        self.assertEqual(self.IMP.leggiMisurazioni(None, ordina=False), [])
        self.assertEqual(self.IMP.leggiMisurazioni(ET.fromstring('<PweMisurazioni></PweMisurazioni>'), ordina=False), [])
        mis_un_figlio = ET.fromstring('<PweMisurazioni><PweElencoPrezzi></PweElencoPrezzi></PweMisurazioni>')
        self.assertEqual(self.IMP.leggiMisurazioni(mis_un_figlio, ordina=False), [])

    def test_leggiMisurazioni_eccezione_mostra_dialogo(self):
        # Quando si verifica un'eccezione durante il processing delle voci (es. PL.strall fallisce),
        # leggiMisurazioni cattura l'eccezione nel blocco except, invoca Dialogs.Exclamation
        # e restituisce la lista delle misure vuota.
        xml_data = """<PweMisurazioni>
            <PweElencoPrezzi></PweElencoPrezzi>
            <PweVociComputo>
                <VCItem ID="1">
                    <IDEP>101</IDEP>
                    <IDSpCat>1</IDSpCat>
                    <IDCat>1</IDCat>
                    <IDSbCat>1</IDSbCat>
                </VCItem>
            </PweVociComputo>
        </PweMisurazioni>"""
        mis = ET.fromstring(xml_data)
        self.IMP.PL.strall.side_effect = TypeError('Errore tipo simulato in strall')
        res = self.IMP.leggiMisurazioni(mis, ordina=True)
        self.assertEqual(res, [])
        self.mock_dialogs.Exclamation.assert_called_once()
        self.IMP.PL.strall.side_effect = fake_strall

    # -------------------------------------------------------------------------
    # 7. Test Roundtrip con dcf_parser.generate_xpwe
    # -------------------------------------------------------------------------
    def test_roundtrip_generate_xpwe_e_leggi_funzioni(self):
        import dcf_parser

        with tempfile.TemporaryDirectory() as td:
            xpwe_file = os.path.join(td, 'roundtrip_test.xpwe')
            doc_data = {
                'info': {
                    'oggetto': 'Lavori di Manutenzione Idraulica',
                    'committente': 'Consorzio Bonifica',
                    'comune': 'Matera',
                    'provincia': 'MT',
                    'impresa': 'Idro Costruzioni Srl',
                    'perc_prezzi': 0.0,
                    'parte_opera': 'Opere Idrauliche',
                },
                'super_categorie': {1: {'nome': 'Super Cat Idraulica', 'des_est': '', 'data_init': '', 'durata': '', 'cod_fase': '', 'percent': 0.0, 'cod': '01'}},
                'categorie': {1: {'nome': 'Canalizzazioni', 'des_est': '', 'data_init': '', 'durata': '', 'cod_fase': '', 'percent': 0.0, 'cod': '01.01'}},
                'sotto_categorie': {1: {'nome': 'Tubazioni PEAD', 'des_est': '', 'data_init': '', 'durata': '', 'cod_fase': '', 'percent': 0.0, 'cod': '01.01.01'}},
                'super_capitoli': {1: {'nome': 'Super Cap Opere', 'des_est': '', 'data_init': '', 'durata': '', 'cod_fase': '', 'percent': 0.0, 'cod': 'SC.01'}},
                'capitoli': {1: {'nome': 'Scavi e Posa', 'des_est': '', 'data_init': '', 'durata': '', 'cod_fase': '', 'percent': 0.0, 'cod': 'C.01'}},
                'sotto_capitoli': {1: {'nome': 'Posa Condotte', 'des_est': '', 'data_init': '', 'durata': '', 'cod_fase': '', 'percent': 0.0, 'cod': 'SCAP.01'}},
                'elenco_prezzi': [
                    {
                        'id': 500,
                        'tariffa': 'IDR.01',
                        'articolo': 'ART_TUB',
                        'des_ridotta': 'Tubo PEAD DN 200',
                        'des_estesa': 'Fornitura e posa di tubo in polietilene PEAD ad alta densita DN 200 mm',
                        'des_breve': '',
                        'um': 'm',
                        'prezzo': 35.80,
                        'prezzo2': 40.0,
                        'prezzo3': 0.0,
                        'prezzo4': 0.0,
                        'prezzo5': 0.0,
                        'inc_mdo': 0.20,
                        'inc_sic': 0.03,
                        'inc_mat': 0.0,
                        'inc_attr': 0.0,
                        'id_spcap': 1,
                        'id_cap': 1,
                        'id_sbcap': 1,
                        'flags': 0,
                        'data': '10/02/2025',
                        'adr_internet': '',
                        'tag_bim': '',
                    }
                ],
                'computo': [
                    {
                        'id': 1,
                        'id_ep': 500,
                        'quantita': 80.0,
                        'data_mis': '10/02/2025',
                        'id_spcal': 1,
                        'id_cat': 1,
                        'id_sbcat': 1,
                        'cod_wbs': '',
                        'flags': 0,
                        'misure': [
                            {
                                'id': 1,
                                'descrizione': 'Tratto A-B',
                                'pu': '1',
                                'lu': '80',
                                'la': '',
                                'hp': '',
                                'qt': 80.0,
                                'rif_voce': -2,
                                'flags': 0,
                            }
                        ],
                    }
                ],
            }

            dcf_parser.generate_xpwe(doc_data, xpwe_file)
            self.assertTrue(os.path.exists(xpwe_file))

            tree = ET.parse(xpwe_file)
            root = tree.getroot()
            dg = root.find('PweDatiGenerali')
            mis = root.find('PweMisurazioni')

            # 1. Anagrafica
            anag = self.IMP.leggiAnagraficaGenerale(dg)
            self.assertEqual(anag['oggetto'], 'Lavori di Manutenzione Idraulica')
            self.assertEqual(anag['committente'], 'Consorzio Bonifica')
            self.assertEqual(anag['comune'], 'Matera')
            self.assertEqual(anag['impresa'], 'Idro Costruzioni Srl')

            # 2. Capitoli e Categorie
            capcat = self.IMP.leggiCapitoliCategorie(dg)
            self.assertEqual(len(capcat['SuperCapitoli']), 1)
            self.assertEqual(capcat['SuperCapitoli'][0]['dessintetica'], 'Super Cap Opere')
            self.assertEqual(len(capcat['Capitoli']), 1)
            self.assertEqual(capcat['Capitoli'][0]['dessintetica'], 'Scavi e Posa')
            self.assertEqual(len(capcat['SottoCapitoli']), 1)
            self.assertEqual(capcat['SottoCapitoli'][0]['dessintetica'], 'Posa Condotte')
            self.assertEqual(len(capcat['SuperCategorie']), 1)
            self.assertEqual(capcat['SuperCategorie'][0]['dessintetica'], 'Super Cat Idraulica')
            self.assertEqual(len(capcat['Categorie']), 1)
            self.assertEqual(capcat['Categorie'][0]['dessintetica'], 'Canalizzazioni')
            self.assertEqual(len(capcat['SottoCategorie']), 1)
            self.assertEqual(capcat['SottoCategorie'][0]['dessintetica'], 'Tubazioni PEAD')

            # 3. Elenco Prezzi
            # BUG NOTO: in leggiElencoPrezzi, Prezzo1 viene convertito con parse_price (float),
            # mentre Prezzo2..Prezzo5 vengono letti come stringa grezza tramite get_xml_text.
            # dcf_parser.generate_xpwe scrive '40' se float 40.0 ha parte decimale zero (f"{p:g}"),
            # quindi art['prezzo2'] vale la stringa '40' invece del float 40.0.
            ep = self.IMP.leggiElencoPrezzi(mis)
            self.assertEqual(len(ep['ListaArticoli']), 1)
            art = ep['DizionarioArticoli']['500']
            self.assertEqual(art['tariffa'], 'IDR.01')
            self.assertEqual(art['unmisura'], 'm')
            self.assertEqual(art['prezzo1'], 35.80)
            self.assertEqual(art['prezzo2'], '40')

            # 4. Misurazioni
            misure = self.IMP.leggiMisurazioni(mis, ordina=False)
            self.assertEqual(len(misure), 1)
            self.assertEqual(misure[0]['id_vc'], '1')
            self.assertEqual(misure[0]['id_ep'], '500')
            self.assertEqual(len(misure[0]['lista_rig']), 1)
            self.assertEqual(misure[0]['lista_rig'][0][0], 'Tratto A-B')
            self.assertEqual(misure[0]['lista_rig'][0][4], '80')


if __name__ == '__main__':
    unittest.main()
