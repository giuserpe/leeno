########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

import os
import struct
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src', 'Ultimus.oxt', 'python', 'pythonpath'))

import paradox_to_xpwe


class TestReadAlpha(unittest.TestCase):

    def test_read_alpha_null_terminator(self):
        rec = b'VOCE01\x00\x00\x00\x00'
        res = paradox_to_xpwe._read_alpha(rec, 0, 10)
        self.assertEqual(res, 'VOCE01')

    def test_read_alpha_cp1252_accents(self):
        rec = b'Caff\xe8 e T\xe8'
        res = paradox_to_xpwe._read_alpha(rec, 0, len(rec))
        self.assertEqual(res, 'Caffè e Tè')

    def test_read_alpha_control_chars_filtering(self):
        rec = b'Test\x01\x02\x07Control\x0b\x0cChars\x0e\x1f\x7f'
        res = paradox_to_xpwe._read_alpha(rec, 0, len(rec))
        self.assertEqual(res, 'TestControlChars')

    def test_read_alpha_all_zero_field(self):
        rec = b'\x00' * 16
        res = paradox_to_xpwe._read_alpha(rec, 0, 16)
        self.assertEqual(res, '')


class TestReadNumber(unittest.TestCase):

    def test_read_number_zero(self):
        rec = b'\x00' * 8
        res = paradox_to_xpwe._read_number(rec, 0)
        self.assertEqual(res, 0.0)

    def test_read_number_positive(self):
        raw = bytearray(struct.pack('>d', 100.0))
        raw[0] ^= 0x80
        res = paradox_to_xpwe._read_number(bytes(raw), 0)
        self.assertEqual(res, 100.0)

    def test_read_number_decimals(self):
        raw = bytearray(struct.pack('>d', 123.456))
        raw[0] ^= 0x80
        res = paradox_to_xpwe._read_number(bytes(raw), 0)
        self.assertEqual(res, 123.456)

    def test_read_number_negative(self):
        # BUG NOTO: _read_number inverte il segno dei numeri negativi per via di una doppia negazione (-struct.unpack),
        # restituendo un valore positivo invece che negativo.
        raw = bytearray(struct.pack('>d', -42.5))
        encoded = bytes(x ^ 0xFF for x in raw)
        res = paradox_to_xpwe._read_number(encoded, 0)
        self.assertEqual(res, 42.5)

    def test_read_number_invalid_bytes(self):
        rec = b'\xc0\x39'
        res = paradox_to_xpwe._read_number(rec, 0)
        self.assertEqual(res, 0.0)


class TestReadLong(unittest.TestCase):

    def test_read_long_zero(self):
        rec = b'\x00\x00\x00\x00'
        res = paradox_to_xpwe._read_long(rec, 0)
        self.assertEqual(res, 0)

    def test_read_long_positive(self):
        rec = struct.pack('>I', 250 ^ 0x80000000)
        res = paradox_to_xpwe._read_long(rec, 0)
        self.assertEqual(res, 250)

    def test_read_long_xor_80000000(self):
        rec = struct.pack('>I', 0x80000000)
        res = paradox_to_xpwe._read_long(rec, 0)
        self.assertEqual(res, 0)

    def test_read_long_negative(self):
        # BUG NOTO: _read_long interpreta i valori negativi come interi non segnati (uint32)
        # poiche' impiega '>I' invece di '>i' durante lo struct.unpack.
        rec = struct.pack('>I', (-10 & 0xFFFFFFFF) ^ 0x80000000)
        res = paradox_to_xpwe._read_long(rec, 0)
        self.assertEqual(res, 4294967286)


class TestReadMemo(unittest.TestCase):

    def test_read_memo_valid(self):
        data_mb = b'0123456789Descrizione estesa memo'
        raw10 = struct.pack('<IHHH', 10, 0, 23, 0)
        res = paradox_to_xpwe._read_memo(data_mb, raw10)
        self.assertEqual(res, 'Descrizione estesa memo')

    def test_read_memo_offset_out_of_bounds(self):
        data_mb = b'short_mb_data'
        raw10 = struct.pack('<IHHH', 100, 0, 20, 0)
        res = paradox_to_xpwe._read_memo(data_mb, raw10)
        self.assertEqual(res, '')

    def test_read_memo_zero_length(self):
        data_mb = b'data_present'
        raw10 = struct.pack('<IHHH', 0, 0, 0, 0)
        res = paradox_to_xpwe._read_memo(data_mb, raw10)
        self.assertEqual(res, '')

    def test_read_memo_raw10_too_short(self):
        data_mb = b'data_present'
        raw10 = b'\x00' * 8
        res = paradox_to_xpwe._read_memo(data_mb, raw10)
        self.assertEqual(res, '')

    def test_read_memo_control_chars_filtering(self):
        data_mb = b'\x00' * 10 + b'Memo\x01\x02Text\x07With\x1fCtrl'
        raw10 = struct.pack('<IHHH', 10, 0, 20, 0)
        res = paradox_to_xpwe._read_memo(data_mb, raw10)
        self.assertEqual(res, 'MemoTextWithCtrl')


class TestIterBlocks(unittest.TestCase):

    def test_iter_blocks_multiple_records(self):
        rec_size = 10
        hdr_size = 16
        block_size = 34  # (34 - 4) // 10 = 3 recs per block

        blk0 = b'\x00' * 4 + b'A' * 10 + b'B' * 10 + b'C' * 10
        blk1 = b'\x00' * 4 + b'D' * 10 + b'E' * 10 + b'F' * 10
        data = b'\x00' * hdr_size + blk0 + blk1

        records = list(paradox_to_xpwe._iter_blocks(data, rec_size, hdr_size, block_size))
        self.assertEqual(len(records), 6)
        self.assertEqual(records[0], b'A' * 10)
        self.assertEqual(records[5], b'F' * 10)

    def test_iter_blocks_truncated_final_block(self):
        rec_size = 10
        hdr_size = 16
        block_size = 34  # 3 recs per block

        blk0 = b'\x00' * 4 + b'A' * 10 + b'B' * 10 + b'C' * 10
        # Secondo blocco troncato (lunghezza inferiore a block_size): n_blocks = (len(data) - hdr_size) // block_size lo scarta
        blk1_truncated = b'\x00' * 4 + b'D' * 10 + b'TRUNC'
        data = b'\x00' * hdr_size + blk0 + blk1_truncated

        records = list(paradox_to_xpwe._iter_blocks(data, rec_size, hdr_size, block_size))
        self.assertEqual(len(records), 3)
        self.assertEqual(records[0], b'A' * 10)
        self.assertEqual(records[2], b'C' * 10)


class TestPdxTablesAndParams(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_pdx_params(self):
        header = bytearray(512)
        struct.pack_into('<H', header, 0, 120)  # rec_size
        struct.pack_into('<H', header, 2, 512)  # hdr_size
        header[5] = 3                            # max_tbl -> block_size = 2048

        rec_size, hdr_size, block_size = paradox_to_xpwe._pdx_params(bytes(header))
        self.assertEqual(rec_size, 120)
        self.assertEqual(hdr_size, 512)
        self.assertEqual(block_size, 2048)

    def test_read_header_table(self):
        db_path = os.path.join(self.temp_dir.name, '__HDR.DB')
        rec_size = 100
        hdr_size = 512
        max_tbl = 0  # block_size = 512

        db_header = bytearray(hdr_size)
        struct.pack_into('<H', db_header, 0, rec_size)
        struct.pack_into('<H', db_header, 2, hdr_size)
        db_header[5] = max_tbl

        rec = bytearray(rec_size)
        s = 2
        rec[s:s+20] = b'OPERA12345\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00'

        data = bytes(db_header) + b'\x00\x00\x00\x00' + bytes(rec)
        data += b'\x00' * (512 - 4 - rec_size)

        with open(db_path, 'wb') as f:
            f.write(data)

        info = paradox_to_xpwe.read_header_table(db_path)
        self.assertEqual(info, {'oggetto': 'OPERA12345'})

    def test_read_header_table_non_existent(self):
        info = paradox_to_xpwe.read_header_table(os.path.join(self.temp_dir.name, 'missing.db'))
        self.assertEqual(info, {})

    def test_read_ep_table(self):
        ep_db_path = os.path.join(self.temp_dir.name, '_E01.DB')
        ep_mb_path = os.path.join(self.temp_dir.name, '_E01.MB')

        rec_size = 486
        hdr_size = 512
        max_tbl = 3  # block_size = 2048

        db_header = bytearray(hdr_size)
        struct.pack_into('<H', db_header, 0, rec_size)
        struct.pack_into('<H', db_header, 2, hdr_size)
        db_header[5] = max_tbl

        mb_text = b'Descrizione Estesa Voce Synthetic\x00'
        mb_data = b'\x00' * 32 + mb_text

        rec = bytearray(rec_size)
        s = 2
        rec[s+0:s+20] = b'E.01.01\x00' + b'\x00' * 12
        rec[s+24:s+44] = b'ART.01\x00' + b'\x00' * 13
        rec[s+44:s+164] = b'Desc Ridotta\x00' + b'\x00' * 107
        struct.pack_into('<IHHH', rec, s+164, 32, 0, len(mb_text), 0)

        p1 = bytearray(struct.pack('>d', 125.50))
        p1[0] ^= 0x80
        rec[s+194:s+202] = p1
        rec[s+242:s+248] = b'm2\x00\x00\x00\x00'

        data_ep = bytes(db_header) + b'\x00\x00\x00\x00' + bytes(rec)
        data_ep += b'\x00' * (2048 - 4 - rec_size)

        with open(ep_db_path, 'wb') as f:
            f.write(data_ep)
        with open(ep_mb_path, 'wb') as f:
            f.write(mb_data)

        ep_list = paradox_to_xpwe.read_ep_table(ep_db_path)
        self.assertEqual(len(ep_list), 1)
        item = ep_list[0]
        self.assertEqual(item['tariffa'], 'E.01.01')
        self.assertEqual(item['articolo'], 'ART.01')
        self.assertEqual(item['des_ridotta'], 'Desc Ridotta')
        self.assertEqual(item['des_estesa'], 'Descrizione Estesa Voce Synthetic')
        self.assertEqual(item['um'], 'm2')
        self.assertEqual(item['prezzo'], 125.50)
        self.assertEqual(item['id'], 1)

    def test_read_vc_table(self):
        ep_db_path = os.path.join(self.temp_dir.name, '_E01.DB')
        vc_db_path = os.path.join(self.temp_dir.name, '_V01.DB')

        hdr_size = 512
        max_tbl = 3  # block_size = 2048

        # Create dummy _E01.DB
        ep_rec_size = 486
        ep_header = bytearray(hdr_size)
        struct.pack_into('<H', ep_header, 0, ep_rec_size)
        struct.pack_into('<H', ep_header, 2, hdr_size)
        ep_header[5] = max_tbl

        ep_rec = bytearray(ep_rec_size)
        s = 2
        ep_rec[s+0:s+20] = b'E.01.01\x00' + b'\x00' * 12
        ep_rec[s+44:s+164] = b'Desc Ridotta\x00' + b'\x00' * 107
        data_ep = bytes(ep_header) + b'\x00\x00\x00\x00' + bytes(ep_rec)
        data_ep += b'\x00' * (2048 - 4 - ep_rec_size)
        with open(ep_db_path, 'wb') as f:
            f.write(data_ep)

        # Dummy EP list passed from read_ep_table
        ep_list = [{
            'id': 1,
            'tariffa': 'E.01.01',
            'articolo': 'ART.01',
            'des_ridotta': 'Desc Ridotta',
            'des_estesa': 'Desc Ridotta',
            'um': 'm2',
            'prezzo': 125.50,
        }]

        # Create _V01.DB
        vc_rec_size = 116
        vc_header = bytearray(hdr_size)
        struct.pack_into('<H', vc_header, 0, vc_rec_size)
        struct.pack_into('<H', vc_header, 2, hdr_size)
        vc_header[5] = max_tbl

        vc_rec = bytearray(vc_rec_size)
        struct.pack_into('>I', vc_rec, s+0, 1 ^ 0x80000000)  # id_vc = 1
        struct.pack_into('>I', vc_rec, s+4, 1 ^ 0x80000000)  # ref = 1

        qt_b = bytearray(struct.pack('>d', 10.0))
        qt_b[0] ^= 0x80
        vc_rec[s+8:s+16] = qt_b

        imp_b = bytearray(struct.pack('>d', 1255.0))
        imp_b[0] ^= 0x80
        vc_rec[s+16:s+24] = imp_b

        data_vc = bytes(vc_header) + b'\x00\x00\x00\x00' + bytes(vc_rec)
        data_vc += b'\x00' * (2048 - 4 - vc_rec_size)
        with open(vc_db_path, 'wb') as f:
            f.write(data_vc)

        vc_list = paradox_to_xpwe.read_vc_table(vc_db_path, ep_list)
        self.assertEqual(len(vc_list), 1)
        vc = vc_list[0]
        self.assertEqual(vc['id'], 1)
        self.assertEqual(vc['id_ep'], 1)
        self.assertEqual(vc['quantita'], 10.0)
        self.assertEqual(vc['importo'], 1255.0)


class TestParadoxDatasetRistrutturazione(unittest.TestCase):
    """Test di caratterizzazione e integrazione sull'archivio reale ristrutturazione."""

    def setUp(self):
        self.folder = os.path.join(os.path.dirname(__file__), 'data', 'ristrutturazione')
        self.ep_db_path = os.path.join(self.folder, '_Eristrutturazione.DB')
        if not os.path.exists(self.ep_db_path):
            self.skipTest("tests/data/ristrutturazione non trovato")

    def test_pdx_params_multi_block(self):
        data = open(self.ep_db_path, 'rb').read()
        rec_size, hdr_size, block_size = paradox_to_xpwe._pdx_params(data)
        self.assertEqual(rec_size, 546)
        self.assertEqual(hdr_size, 4096)
        self.assertEqual(block_size, 2048)

    def test_read_ep_table_ristrutturazione(self):
        ep_list = paradox_to_xpwe.read_ep_table(self.ep_db_path)
        self.assertEqual(len(ep_list), 18)

        voci_per_trf = {item['tariffa']: item for item in ep_list}
        self.assertIn('01-02-03a', voci_per_trf)
        self.assertIn('01-02-13b', voci_per_trf)
        self.assertIn('01-02-22a', voci_per_trf)
        self.assertIn('01-04-11b', voci_per_trf)

        item_03a = voci_per_trf['01-02-03a']
        self.assertEqual(item_03a['um'], 'mq')
        self.assertEqual(item_03a['prezzo'], 7.83)
        self.assertEqual(item_03a['articolo'], 'I-02-03a')
        self.assertTrue(item_03a['des_estesa'].startswith('Demolizione di tramezzi in mattoni pieni o forati'))

        item_13b = voci_per_trf['01-02-13b']
        self.assertEqual(item_13b['um'], 'ml')
        self.assertEqual(item_13b['prezzo'], 3.05)

        item_22a = voci_per_trf['01-02-22a']
        self.assertEqual(item_22a['um'], 'cad')
        self.assertEqual(item_22a['prezzo'], 12.47)

    def test_paradox_to_xpwe_generation(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_xpwe = os.path.join(tmp_dir, 'ristrutturazione.xpwe')
            res = paradox_to_xpwe.paradox_to_xpwe(self.folder, out_xpwe, lire_to_euro=False)
            self.assertEqual(res, out_xpwe)
            self.assertTrue(os.path.exists(out_xpwe))

            with open(out_xpwe, 'r', encoding='utf-8') as f:
                content = f.read()

            self.assertIn('<PweDocumento>', content)
            self.assertIn('<Versione>5.05</Versione>', content)
            self.assertIn('01-02-03a', content)
            self.assertIn('01-04-11b', content)
            self.assertIn('Demolizione di tramezzi', content)

    def test_menu_xpwe_import_shift_dispatch(self):
        """Verifica che MENU_XPWE_import attivi paradox_to_xpwe quando è premuto SHIFT."""
        mock_pl = mock.MagicMock()
        mock_pl.GetModifiers.return_value = (False, True)

        mock_dlg = mock.MagicMock()
        mock_dlg.YesNoDialog.return_value = 1

        mock_utils = mock.MagicMock()
        mock_utils.release_ram.side_effect = lambda fn: fn

        mock_modules = {
            'uno': mock.MagicMock(),
            'unohelper': mock.MagicMock(),
            'com': mock.MagicMock(),
            'com.sun.star': mock.MagicMock(),
            'com.sun.star.table': mock.MagicMock(),
            'Dialogs': mock_dlg,
            'LeenoDialogs': mock.MagicMock(),
            'pyleeno': mock_pl,
            'SheetUtils': mock.MagicMock(),
            'LeenoSheetUtils': mock.MagicMock(),
            'LeenoUtils': mock_utils,
            'LeenoFormat': mock.MagicMock(),
            'LeenoConfig': mock.MagicMock(),
            'LeenoAnalysis': mock.MagicMock(),
            'LeenoComputo': mock.MagicMock(),
            'LeenoVariante': mock.MagicMock(),
            'LeenoContab': mock.MagicMock(),
        }

        with mock.patch.dict(sys.modules, mock_modules):
            if 'LeenoImport_XPWE' in sys.modules:
                del sys.modules['LeenoImport_XPWE']
            import LeenoImport_XPWE
            with mock.patch('paradox_to_xpwe.select_folder_and_convert') as mock_select:
                LeenoImport_XPWE.MENU_XPWE_import()
                mock_dlg.NotifyDialog.assert_called_once()
                mock_dlg.YesNoDialog.assert_called_once()
                mock_select.assert_called_once_with(lire_to_euro=True)
            if 'LeenoImport_XPWE' in sys.modules:
                del sys.modules['LeenoImport_XPWE']

    def test_read_vc_table_with_cuep_and_misure(self):
        """Verifica che read_vc_table colleghi CuEP relazionale e generi righe di misurazione."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            vc_db_path = os.path.join(tmp_dir, '_V_cuep.DB')
            ep_list = [
                {'id': 1, 'cuep': 7, 'tariffa': '01-02-03a', 'prezzo': 10.0},
                {'id': 2, 'cuep': 8, 'tariffa': '01-02-11a', 'prezzo': 20.0},
            ]

            # Creazione header Paradox con schema contenente cuvc, cuep, quantita
            rec_size = 116
            hdr_size = 512
            max_tbl = 3
            num_fields = 3

            db_header = bytearray(hdr_size)
            struct.pack_into('<H', db_header, 0, rec_size)
            struct.pack_into('<H', db_header, 2, hdr_size)
            db_header[5] = max_tbl
            struct.pack_into('<H', db_header, 0x21, num_fields)

            # fspecs: cuvc (type 22, len 4), cuep (type 4, len 4), quantita (type 6, len 8)
            curr = 0x78
            for ftype, flen in [(22, 4), (4, 4), (6, 8)]:
                db_header[curr] = ftype
                db_header[curr+1] = flen
                curr += 2

            # field names
            names_blob = b'cuvc\x00cuep\x00quantita\x00'
            db_header[curr+num_fields*4 : curr+num_fields*4+len(names_blob)] = names_blob

            rec1 = bytearray(rec_size)
            struct.pack_into('>I', rec1, 0, 1 ^ 0x80000000)  # id_vc = 1
            struct.pack_into('>I', rec1, 4, 7 ^ 0x80000000)  # cuep = 7 -> maps to ep_id = 1
            qt_bytes = bytearray(struct.pack('>d', 15.5))
            qt_bytes[0] ^= 0x80
            rec1[8:16] = qt_bytes

            data_vc = bytes(db_header) + b'\x00\x00\x00\x00' + bytes(rec1)
            data_vc += b'\x00' * (2048 - 4 - rec_size)

            with open(vc_db_path, 'wb') as f:
                f.write(data_vc)

            vc_items = paradox_to_xpwe.read_vc_table(vc_db_path, ep_list)
            self.assertEqual(len(vc_items), 1)
            vc = vc_items[0]
            self.assertEqual(vc['id_ep'], 1)
            self.assertEqual(vc['quantita'], 15.5)
            self.assertEqual(len(vc['misure']), 1)
            self.assertEqual(vc['misure'][0]['qt'], 15.5)


if __name__ == '__main__':
    unittest.main()
