########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################
'''
Test della logica pura di LeenoCdPBis (taglio del prefisso, normalizzazione,
similarità, classificazione, confronto, report). Nessun UNO: il modulo viene
caricato con i moduli UNO sostituiti da mock solo per la durata dei test.
Fixture sintetiche: codici, descrizioni e prezzi sono inventati.
'''
import ast
import csv
import os
import sys
from unittest.mock import MagicMock, patch

import pytest

PYTHONPATH = os.path.abspath(os.path.join(
    os.path.dirname(__file__), '..', 'src', 'Ultimus.oxt', 'python', 'pythonpath'))
sys.path.insert(0, PYTHONPATH)

SOGLIA_VERDE = 0.95
SOGLIA_GIALLO = 0.80

MODULI_MOCK = (
    'uno', 'unohelper', 'Dialogs', 'LeenoUtils', 'LeenoConfig',
    'LeenoComputo', 'LeenoAnalysis', 'SheetUtils', 'DocUtils')


@pytest.fixture(scope='module')
def cdp():
    mock = {nome: MagicMock() for nome in MODULI_MOCK}
    with patch.dict(sys.modules, mock):
        sys.modules.pop('LeenoCdPBis', None)
        import LeenoCdPBis
        yield LeenoCdPBis
        sys.modules.pop('LeenoCdPBis', None)


def riga_ep(codice, desc, um, prezzo):
    '''Riga A:J di Elenco Prezzi, come restituita da getDataArray.'''
    return (codice, desc, um, '', prezzo, '', '', '', '', '')


# ---------------------------------------------------------------- prefisso

@pytest.mark.parametrize('codice, atteso', [
    ('BAS25_B.11.014.02', 'B.11.014.02'),
    ('NP10_BAS25_B.05.029.01', 'B.05.029.01'),
    ('VDS_NP14_BAS23_S.01.021.01', 'S.01.021.01'),
    ('PUG23_INF.002.094#', 'INF.002.094#'),
    ('TOS_01.A02.003', '01.A02.003'),
    ('LAZ24_A1.2', 'A1.2'),
    ('000', '000'),
    ('Fine elenco', 'Fine elenco'),
    ('  BAS25_X.1  ', 'X.1'),
    ('ABC_', 'ABC_'),
    ('', ''),
    (None, ''),
    (0.0, '0'),
])
def test_taglia_prefisso(cdp, codice, atteso):
    assert cdp.taglia_prefisso(codice) == atteso


def test_taglia_prefisso_nessun_prefisso_cablato(cdp):
    for emittente in ('BAS', 'PUG', 'TOS', 'LAZ', 'NP', 'VDS', 'QQQ99'):
        assert cdp.taglia_prefisso(f'{emittente}_Z.9') == 'Z.9'


# ----------------------------------------------------------- normalizzazione

def test_normalizza_testo_maiuscole_spazi_punteggiatura(cdp):
    a = 'Fornitura, e posa  in opera.\nSpessore 5 mm; (colore: grigio)'
    b = 'FORNITURA E POSA IN OPERA SPESSORE 5 MM COLORE GRIGIO'
    assert cdp.normalizza_testo(a) == cdp.normalizza_testo(b)


def test_normalizza_testo_vuoto_e_none(cdp):
    assert cdp.normalizza_testo('') == ''
    assert cdp.normalizza_testo(None) == ''
    assert cdp.normalizza_testo(' .,; ') == ''


def test_normalizza_testo_distingue_i_numeri(cdp):
    assert cdp.normalizza_testo('spessore 15 cm') != cdp.normalizza_testo('spessore 1,5 cm')


@pytest.mark.parametrize('a, b', [
    ('mq', 'MQ'), ('mq', 'm²'), ('m2', 'mq'), ('m 2', 'MQ'), ('m.q.', 'mq'),
    ('mc', 'm³'), ('m3', 'MC'), ('cad', 'CAD'), ('%', '%'),
])
def test_normalizza_um_equivalenti(cdp, a, b):
    assert cdp.normalizza_um(a) == cdp.normalizza_um(b)


@pytest.mark.parametrize('a, b', [
    ('ml', 'm'), ('mq', 'mc'), ('kg', 'q'), ('mq', 'cad'), ('m2', 'm3'),
])
def test_normalizza_um_diverse(cdp, a, b):
    assert cdp.normalizza_um(a) != cdp.normalizza_um(b)


# ------------------------------------------------------------- similarità

def test_similarita_identica_dopo_normalizzazione(cdp):
    assert cdp.similarita('Posa di  tubo, DN 100.', 'POSA DI TUBO DN 100') == 1.0


def test_similarita_vuoti(cdp):
    assert cdp.similarita('', '') == 1.0
    assert cdp.similarita('', 'tubo') == 0.0
    assert cdp.similarita('tubo', None) == 0.0


def test_similarita_piccola_differenza_resta_alta(cdp):
    a = 'Fornitura e posa in opera di tubazione in polietilene PN 10 diametro 110 mm'
    b = 'Fornitura e posa in opera di tubazione in polietilene PN 16 diametro 110 mm'
    assert 0.95 <= cdp.similarita(a, b) < 1.0


def test_similarita_testi_diversi_bassa(cdp):
    s = cdp.similarita('Scavo a sezione obbligata in terreno ordinario',
                       'Intonaco civile per interni su pareti verticali')
    assert s < 0.5


def test_similarita_ignora_il_prezzo(cdp):
    # il prezzo non e' un argomento: stessa descrizione, prezzi diversi
    assert cdp.similarita('Opera X', 'Opera X') == 1.0


# ---------------------------------------------------------- classificazione

@pytest.mark.parametrize('sim, atteso', [
    (1.0, 'VERDE'), (0.95, 'VERDE'), (0.9499, 'GIALLO'), (0.80, 'GIALLO'),
    (0.7999, 'ROSSO'), (0.0, 'ROSSO'),
])
def test_classifica_soglie(cdp, sim, atteso):
    assert cdp.classifica(sim, True, SOGLIA_VERDE, SOGLIA_GIALLO) == atteso


def test_classifica_um_diversa_e_sempre_rosso(cdp):
    assert cdp.classifica(1.0, False, SOGLIA_VERDE, SOGLIA_GIALLO) == 'ROSSO'
    assert cdp.classifica(0.96, False, SOGLIA_VERDE, SOGLIA_GIALLO) == 'ROSSO'


def test_classifica_tolleranza_sul_confine(cdp):
    assert cdp.classifica(0.95 - 1e-12, True, SOGLIA_VERDE, SOGLIA_GIALLO) == 'VERDE'


# --------------------------------------------------------------- indici

def test_indice_codici_prima_occorrenza_e_maiuscolo(cdp):
    righe = [riga_ep('', '', '', ''), riga_ep('abc_1', 'x', 'mq', 1.0),
             riga_ep('ABC_1', 'y', 'mq', 2.0)]
    assert cdp.indice_codici(righe) == {'ABC_1': 1}


def test_indice_suffissi_ambigui(cdp):
    righe = [riga_ep('AAA_X.1', 'a', 'mq', 1.0), riga_ep('BBB_X.1', 'b', 'mq', 1.0),
             riga_ep('AAA_X.2', 'c', 'mq', 1.0)]
    indice, ambigui = cdp.indice_suffissi(righe)
    assert ambigui == {'X.1'}
    assert indice['X.2'] == 2


# ---------------------------------------------------------------- confronto

def voce(cdp, codice, riga=5, foglio='COMPUTO', numero='1'):
    return cdp.Voce(foglio, riga, codice, numero)


def confronta(cdp, voci, doc, prez, esclusi=frozenset()):
    return cdp.confronta_voci(voci, doc, prez, SOGLIA_VERDE, SOGLIA_GIALLO, esclusi)


def test_confronto_voce_gia_allineata_viene_saltata(cdp):
    prez = [riga_ep('NEW25_A.1', 'Tubo', 'm', 10.0)]
    doc = [riga_ep('NEW25_A.1', 'Tubo', 'm', 9.0)]
    ris = confronta(cdp, [voce(cdp, 'NEW25_A.1')], doc, prez)
    assert ris.esiti == []
    assert len(ris.allineate) == 1


def test_confronto_codice_non_trovato(cdp):
    prez = [riga_ep('NEW25_A.1', 'Tubo', 'm', 10.0)]
    doc = [riga_ep('OLD24_B.9', 'Altro', 'm', 5.0)]
    ris = confronta(cdp, [voce(cdp, 'OLD24_B.9')], doc, prez)
    assert ris.esiti == []
    assert [v.codice for v in ris.non_trovate] == ['OLD24_B.9']


def test_confronto_righe_senza_prefisso_restano_invariate(cdp):
    prez = [riga_ep('NEW25_A.1', 'Tubo', 'm', 10.0),
            riga_ep('Fine elenco', 'Non cancellare', '', '')]
    doc = [riga_ep('000', 'arrotondamenti', '€', 1.0)]
    ris = confronta(cdp, [voce(cdp, '000')], doc, prez)
    assert ris.esiti == []
    assert [v.codice for v in ris.non_trovate] == ['000']


def test_confronto_prezzario_senza_prefisso_stesso_codice_e_allineato(cdp):
    prez = [riga_ep('A.01.001', 'Tubo', 'm', 10.0)]
    doc = [riga_ep('A.01.001', 'Tubo', 'm', 9.0)]
    ris = confronta(cdp, [voce(cdp, 'A.01.001')], doc, prez)
    assert len(ris.allineate) == 1 and ris.esiti == []


def test_confronto_verde_descrizione_identica(cdp):
    prez = [riga_ep('NEW25_A.1', 'Posa di tubo, DN 100.', 'm', 12.0)]
    doc = [riga_ep('OLD24_A.1', 'POSA DI TUBO DN 100', 'm', 10.0)]
    ris = confronta(cdp, [voce(cdp, 'OLD24_A.1')], doc, prez)
    e = ris.esiti[0]
    assert (e.colore, e.similarita, e.codice_nuovo) == ('VERDE', 1.0, 'NEW25_A.1')
    assert (e.prezzo_vecchio, e.prezzo_nuovo) == (10.0, 12.0)


def test_confronto_il_prezzo_non_cambia_il_colore(cdp):
    prez = [riga_ep('NEW25_A.1', 'Tubo DN 100', 'm', 1000.0)]
    doc = [riga_ep('OLD24_A.1', 'Tubo DN 100', 'm', 1.0)]
    assert confronta(cdp, [voce(cdp, 'OLD24_A.1')], doc, prez).esiti[0].colore == 'VERDE'


def test_confronto_giallo(cdp):
    vecchia = 'Fornitura e posa in opera di tubazione in polietilene diametro 110 mm completa di pezzi speciali'
    nuova = 'Fornitura e posa in opera di tubazione in polietilene diametro 110 mm completa'
    prez = [riga_ep('NEW25_A.1', nuova, 'm', 12.0)]
    doc = [riga_ep('OLD24_A.1', vecchia, 'm', 10.0)]
    e = confronta(cdp, [voce(cdp, 'OLD24_A.1')], doc, prez).esiti[0]
    assert e.colore == 'GIALLO'
    assert 0.80 <= e.similarita < 0.95


def test_confronto_rosso_per_descrizione_diversa(cdp):
    prez = [riga_ep('NEW25_A.1', 'Intonaco civile per interni', 'mq', 12.0)]
    doc = [riga_ep('OLD24_A.1', 'Scavo a sezione obbligata', 'mq', 10.0)]
    assert confronta(cdp, [voce(cdp, 'OLD24_A.1')], doc, prez).esiti[0].colore == 'ROSSO'


def test_confronto_rosso_per_um_diversa_anche_con_testo_identico(cdp):
    prez = [riga_ep('NEW25_A.1', 'Tubo DN 100', 'ml', 12.0)]
    doc = [riga_ep('OLD24_A.1', 'Tubo DN 100', 'm', 10.0)]
    e = confronta(cdp, [voce(cdp, 'OLD24_A.1')], doc, prez).esiti[0]
    assert (e.colore, e.similarita) == ('ROSSO', 1.0)
    assert 'unità di misura' in e.nota


def test_confronto_um_mq_e_m2_sono_uguali(cdp):
    prez = [riga_ep('NEW25_A.1', 'Pavimento', 'm²', 12.0)]
    doc = [riga_ep('OLD24_A.1', 'Pavimento', 'mq', 10.0)]
    assert confronta(cdp, [voce(cdp, 'OLD24_A.1')], doc, prez).esiti[0].colore == 'VERDE'


def test_confronto_nuovo_prezzo_np_non_coincidente_e_rosso(cdp):
    prez = [riga_ep('BAS25_B.05.029.01', 'Sovrapprezzo per operazioni speciali', '%', 20.0)]
    doc = [riga_ep('NP10_BAS25_B.05.029.01', 'Lavorazione del tutto diversa', '%', 20.0)]
    e = confronta(cdp, [voce(cdp, 'NP10_BAS25_B.05.029.01')], doc, prez).esiti[0]
    assert e.colore == 'ROSSO'
    assert e.codice_nuovo == 'BAS25_B.05.029.01'


def test_confronto_codice_vecchio_assente_da_elenco_prezzi(cdp):
    prez = [riga_ep('NEW25_A.1', 'Tubo', 'm', 12.0)]
    e = confronta(cdp, [voce(cdp, 'OLD24_A.1')], [], prez).esiti[0]
    assert e.colore == 'ROSSO'
    assert 'assente' in e.nota


def test_confronto_suffisso_ambiguo_non_sostituisce(cdp):
    prez = [riga_ep('AAA_X.1', 'a', 'm', 1.0), riga_ep('BBB_X.1', 'a', 'm', 1.0)]
    doc = [riga_ep('OLD24_X.1', 'a', 'm', 1.0)]
    ris = confronta(cdp, [voce(cdp, 'OLD24_X.1')], doc, prez)
    assert ris.esiti == [] and len(ris.ambigue) == 1


def test_confronto_esclusi(cdp):
    prez = [riga_ep('NEW25_A.1', 'Tubo', 'm', 12.0)]
    doc = [riga_ep('OLD24_A.1', 'Tubo', 'm', 10.0)]
    ris = confronta(cdp, [voce(cdp, 'OLD24_A.1')], doc, prez, esclusi={'OLD24_A.1'})
    assert ris.esiti == [] and len(ris.escluse) == 1


def test_confronto_match_case_insensitive(cdp):
    prez = [riga_ep('NEW25_b.1', 'Tubo', 'm', 12.0)]
    doc = [riga_ep('old24_B.1', 'Tubo', 'm', 10.0)]
    assert len(confronta(cdp, [voce(cdp, 'old24_B.1')], doc, prez).esiti) == 1


def test_confronto_stessa_voce_su_fogli_diversi_resta_distinta(cdp):
    prez = [riga_ep('NEW25_A.1', 'Tubo', 'm', 12.0)]
    doc = [riga_ep('OLD24_A.1', 'Tubo', 'm', 10.0)]
    voci = [voce(cdp, 'OLD24_A.1', 5, 'COMPUTO'), voce(cdp, 'OLD24_A.1', 8, 'CONTABILITA')]
    ris = confronta(cdp, voci, doc, prez)
    assert [(e.voce.foglio, e.voce.riga) for e in ris.esiti] == [('COMPUTO', 5), ('CONTABILITA', 8)]


def test_riepilogo_conteggi(cdp):
    prez = [riga_ep('NEW25_A.1', 'Tubo', 'm', 12.0), riga_ep('NEW25_A.2', 'Scavo', 'mc', 1.0)]
    doc = [riga_ep('OLD24_A.1', 'Tubo', 'm', 10.0), riga_ep('OLD24_A.2', 'Altro', 'mc', 1.0)]
    voci = [voce(cdp, 'OLD24_A.1'), voce(cdp, 'OLD24_A.2'), voce(cdp, 'ZZZ_Q.1'),
            voce(cdp, 'NEW25_A.1')]
    c = cdp.riepilogo_conteggi(confronta(cdp, voci, doc, prez))
    assert (c['VERDE'], c['GIALLO'], c['ROSSO']) == (1, 0, 1)
    assert (c['allineate'], c['non_trovate'], c['ambigue'], c['escluse']) == (1, 1, 0, 0)


# ------------------------------------------------- dialogo e applicazione

def test_ordina_per_colore_stabile_e_preselezione(cdp):
    prez = [riga_ep('N_A.1', 'Tubo', 'm', 1.0), riga_ep('N_A.2', 'Scavo profondo', 'm', 1.0),
            riga_ep('N_A.3', 'Pavimento', 'mq', 1.0)]
    doc = [riga_ep('O_A.1', 'Tubo', 'm', 1.0), riga_ep('O_A.2', 'Intonaco', 'm', 1.0),
           riga_ep('O_A.3', 'Pavimento', 'mq', 1.0)]
    voci = [voce(cdp, 'O_A.1', 1), voce(cdp, 'O_A.2', 2), voce(cdp, 'O_A.3', 3)]
    esiti = cdp.ordina_per_colore(confronta(cdp, voci, doc, prez).esiti)
    assert [e.colore for e in esiti] == ['ROSSO', 'VERDE', 'VERDE']
    assert [e.voce.riga for e in esiti] == [2, 1, 3]
    assert cdp.preselezione(esiti) == [False, True, True]


def test_righe_tabella_formato(cdp):
    prez = [riga_ep('N_A.1', 'Tubo nuovo', 'm', 1234.5)]
    doc = [riga_ep('O_A.1', 'Tubo vecchio', 'm', 9.0)]
    esiti = confronta(cdp, [voce(cdp, 'O_A.1', 7, 'COMPUTO', '12')], doc, prez).esiti
    riga = cdp.righe_tabella(esiti, [True])[0]
    assert riga[0] == '[x]'
    assert riga[1] == 'COMPUTO 12  O_A.1'
    assert riga[2] in ('VERDE', 'GIALLO', 'ROSSO')
    assert riga[6:] == ('9,00', '1.234,50')
    assert cdp.righe_tabella(esiti, [False])[0][0] == '[ ]'


def test_righe_da_aggiungere_una_volta_e_senza_duplicati(cdp):
    prez = [riga_ep('N_A.1', 'Tubo', 'm', 1.0), riga_ep('N_A.2', 'Scavo', 'mc', 1.0)]
    doc = [riga_ep('O_A.1', 'Tubo', 'm', 1.0), riga_ep('O_A.2', 'Scavo', 'mc', 1.0),
           riga_ep('n_a.2', 'Scavo', 'mc', 1.0)]
    voci = [voce(cdp, 'O_A.1', 1), voce(cdp, 'O_A.1', 2, 'CONTABILITA'), voce(cdp, 'O_A.2', 3)]
    esiti = confronta(cdp, voci, doc, prez).esiti
    nuove = cdp.righe_da_aggiungere(esiti, prez, doc)
    # N_A.1 una sola volta; N_A.2 esiste gia' (anche con maiuscole diverse)
    assert [r[0] for r in nuove] == ['N_A.1']


def test_righe_da_aggiungere_solo_voci_scelte(cdp):
    prez = [riga_ep('N_A.1', 'Tubo', 'm', 1.0)]
    doc = [riga_ep('O_A.1', 'Tubo', 'm', 1.0)]
    esiti = confronta(cdp, [voce(cdp, 'O_A.1')], doc, prez).esiti
    assert cdp.righe_da_aggiungere([], prez, doc) == []
    assert len(cdp.righe_da_aggiungere(esiti, prez, doc)) == 1


# ------------------------------------------------------------------ report

def test_formatta_prezzo(cdp):
    assert cdp.formatta_prezzo(1234.5) == '1.234,50'
    assert cdp.formatta_prezzo(0) == '0,00'
    assert cdp.formatta_prezzo(3) == '3,00'
    assert cdp.formatta_prezzo('') == ''
    assert cdp.formatta_prezzo(None) == ''


def test_una_riga_tronca_e_compatta(cdp):
    assert cdp.una_riga('a\n  b\tc') == 'a b c'
    lunga = 'x' * 200
    assert len(cdp.una_riga(lunga, 50)) == 50
    assert cdp.una_riga(lunga, 50).endswith('...')


def test_righe_report_solo_divergenze(cdp):
    prez = [riga_ep('N_A.1', 'Tubo', 'm', 12.0), riga_ep('N_A.2', 'Intonaco', 'mq', 5.0),
            riga_ep('N_A.3', 'Pavimento', 'ml', 5.0)]
    doc = [riga_ep('O_A.1', 'Tubo', 'm', 10.0), riga_ep('O_A.2', 'Scavo', 'mq', 4.0),
           riga_ep('O_A.3', 'Pavimento', 'mq', 4.0)]
    voci = [voce(cdp, 'O_A.1', 1), voce(cdp, 'O_A.2', 2), voce(cdp, 'O_A.3', 3),
            voce(cdp, 'X_Z.9', 4)]
    ris = confronta(cdp, voci, doc, prez)
    righe = cdp.righe_report(ris, {('COMPUTO', 2)})
    tipi = [(r[0], r[4], r[14]) for r in righe]
    # O_A.1 e' identica: non compare. O_A.2 (testo) e O_A.3 (um) si.
    assert tipi == [('DIVERSA', 'O_A.2', 'SI'), ('DIVERSA', 'O_A.3', 'NO'),
                    ('NON_TROVATA', 'X_Z.9', 'NO')]
    assert righe[0][2] == 3  # riga 1-based


def test_scrivi_report_roundtrip(cdp, tmp_path):
    prez = [riga_ep('N_A.1', 'Tubo;\nnuovo "x"', 'm', 12.0)]
    doc = [riga_ep('O_A.1', 'Tubo vecchio', 'm', 10.0)]
    ris = confronta(cdp, [voce(cdp, 'O_A.1')], doc, prez)
    percorso = tmp_path / 'report.csv'
    cdp.scrivi_report(str(percorso), cdp.righe_report(ris, set()))
    with open(percorso, encoding='utf-8-sig', newline='') as f:
        letto = list(csv.reader(f, delimiter=';'))
    assert tuple(letto[0]) == cdp.INTESTAZIONE_REPORT
    assert len(letto) == 2
    assert letto[1][5] == 'N_A.1'
    assert letto[1][11] == 'Tubo; nuovo "x"'
    assert open(percorso, 'rb').read().startswith(b'\xef\xbb\xbf')


def test_nome_report(cdp):
    assert cdp.nome_report('CdP', '20261112_101500') == 'CdP_CdPBis_20261112_101500.csv'
    assert cdp.nome_report('', '20261112_101500') == 'documento_CdPBis_20261112_101500.csv'


# ------------------------------------------------------- sfondi registrati

def test_codifica_decodifica_sfondo_roundtrip(cdp):
    for sfondo in (('DEFAULT_VALUE', 16772089, False), ('DIRECT_VALUE', -5127, True),
                   ('DIRECT_VALUE', 14942166, False)):
        assert cdp.decodifica_sfondo(cdp.codifica_sfondo(*sfondo)) == sfondo


@pytest.mark.parametrize('testo', ['', 'abc', 'a|b|c', 'X|1', None, 'X|1|0|9'])
def test_decodifica_sfondo_non_valida(cdp, testo):
    assert cdp.decodifica_sfondo(testo) is None


def test_unisci_registri_non_sovrascrive_il_primo(cdp):
    esistente = {'standard': {'COMPUTO': {'comp Art-EP_R': 'DEFAULT_VALUE|1|0'}},
                 'eccezioni': {'COMPUTO': {'5': ['A.1', 'X|9|0']}}}
    nuovo = {'standard': {'COMPUTO': {'comp Art-EP_R': 'DIRECT_VALUE|2|0', 'altro': 'X|3|1'},
                          'CONTABILITA': {'comp Art-EP_R': 'DIRECT_VALUE|-1|1'}},
             'eccezioni': {'COMPUTO': {'5': ['B.2', 'Y|8|0'], '6': ['C.3', 'Z|7|1']}}}
    unito = cdp.unisci_registri(esistente, nuovo)
    assert unito['standard']['COMPUTO']['comp Art-EP_R'] == 'DEFAULT_VALUE|1|0'
    assert unito['standard']['COMPUTO']['altro'] == 'X|3|1'
    assert unito['standard']['CONTABILITA']['comp Art-EP_R'] == 'DIRECT_VALUE|-1|1'
    assert unito['eccezioni']['COMPUTO']['5'] == ['A.1', 'X|9|0']
    assert unito['eccezioni']['COMPUTO']['6'] == ['C.3', 'Z|7|1']
    assert esistente['eccezioni']['COMPUTO'] == {'5': ['A.1', 'X|9|0']}


# ------------------------------------------- regole di progetto sul modulo

def _albero():
    percorso = os.path.join(PYTHONPATH, 'LeenoCdPBis.py')
    with open(percorso, encoding='utf-8') as f:
        return ast.parse(f.read())


def test_nessun_effetto_collaterale_a_livello_di_modulo():
    consentiti = (ast.Import, ast.ImportFrom, ast.FunctionDef, ast.Assign)
    for nodo in _albero().body:
        if isinstance(nodo, ast.Expr) and isinstance(nodo.value, ast.Constant):
            continue  # docstring
        assert isinstance(nodo, consentiti), f'riga {nodo.lineno}: {type(nodo).__name__}'
        if isinstance(nodo, ast.Assign):
            for sotto in ast.walk(nodo.value):
                if isinstance(sotto, ast.Call):
                    assert getattr(sotto.func, 'id', '') in ('namedtuple', 'Voce'), \
                        f'riga {nodo.lineno}: chiamata a livello di modulo'


def test_nessun_print_ne_dlg_chi_ne_except_silenzioso():
    albero = _albero()
    for nodo in ast.walk(albero):
        if isinstance(nodo, ast.Call) and isinstance(nodo.func, ast.Name):
            assert nodo.func.id != 'print', f'riga {nodo.lineno}: print()'
        if isinstance(nodo, ast.Attribute) and nodo.attr == 'chi':
            raise AssertionError(f'riga {nodo.lineno}: DLG.chi()')
        if isinstance(nodo, ast.ExceptHandler):
            corpo_pass = all(isinstance(b, ast.Pass) for b in nodo.body)
            assert not corpo_pass, f'riga {nodo.lineno}: except con solo pass'
            if nodo.name is not None:
                assert nodo.name == '_exc', f'riga {nodo.lineno}: usare as _exc'


def test_intestazione_lgpl():
    percorso = os.path.join(PYTHONPATH, 'LeenoCdPBis.py')
    with open(percorso, encoding='utf-8') as f:
        testa = f.read(400)
    assert 'Giuseppe Vizziello - supporto@leeno.org' in testa
    assert 'LGPL' in testa
