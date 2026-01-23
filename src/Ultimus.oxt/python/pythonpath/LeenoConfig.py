'''
Gestione delle impostazioni di LeenO
'''
import sys
import os
import configparser

from os.path import expanduser
from datetime import date

import LeenoUtils
import PersistUtils

COLORE_COLONNE_RAFFRONTO = 16645616
COLORE_ROSA_INPUT = 16773375
COLORE_GIALLO_VARIANTE = 16757935
COLORE_VERDE_SPUNTA = 14942166
COLORE_GRIGIO_INATTIVA = 15066597
COLORE_BIANCO_SFONDO = 16777215


class Borg:
    '''
    Singleton/BorgSingleton.py
    Alex Martelli's 'Borg'
    '''

    _shared_state = {}

    def __init__(self):
        self.__dict__ = self._shared_state


class Config(Borg):
    '''
    classe contenente la configurazione di LeenO
    la classe è un singleton - anche creando vari
    oggetti tutti contengono gli stessi dati
    '''
    def __init__(self):
        Borg.__init__(self)

        if self._shared_state == {}:

            # the path of config file
            if sys.platform == 'win32':
                self._path = os.getenv("APPDATA") + '/.config/leeno/leeno.conf'
            else:
                self._path = os.getenv("HOME") + '/.config/leeno/leeno.conf'

            # just in case... crea la cartella radice dove piazzare
            # il file di configurazione
            try:
                os.makedirs(os.path.dirname(self._path))
            except FileExistsError:
                pass

            #self._parser = configparser.ConfigParser()

            self._parser = configparser.RawConfigParser()
            self._parser.optionxform = lambda option: option

            # load values from file, if exist
            self._load()

            # fill with default values items that are missing
            self._initDefaults()

    def _initDefaults(self):
        '''
        default configuration parameters, if not changed by user
        '''
        parametri = (
            ('Zoom', 'fattore', '100'),
            ('Zoom', 'fattore_ottimale', '81'),
            ('Zoom', 'fullscreen', '0'),

            ('Generale', 'dialogo', '1'),
            #  ('Generale', 'visualizza', 'Menù Principale'),
            ('Generale', 'altezza_celle', '1.25'),
            ('Generale', 'pesca_auto', '1'),
            ('Generale', 'movedirection', '1'),
            ('Generale', 'descrizione_in_una_colonna', '0'),
            ('Generale', 'toolbar_contestuali', '1'),
            ('Generale', 'vedi_voce_breve', '50'),
            ('Generale', 'dettaglio', '1'),
            ('Generale', 'torna_a_ep', '1'),
            ('Generale', 'copie_backup', '5'),
            ('Generale', 'pausa_backup', '15'),
            ('Generale', 'conta_usi', '0'),
            ('Generale', 'ultimo_percorso', expanduser("~")),
            ('Generale', 'precisione_come_mostrato', '(bool)True'),
            ('Generale', 'nuova_voce', 'True'),

            #  ('Computo', 'riga_bianca_categorie', '1'),
            #  ('Computo', 'voci_senza_numerazione', '0'),
            ('Computo', 'inizio_voci_abbreviate', '100'),
            ('Computo', 'fine_voci_abbreviate', '120'),
            ('Computo', 'costo_medio_mdo', ''),
            ('Computo', 'addetti_mdo', ''),

            ('Contabilita', 'cont_inizio_voci_abbreviate', '100'),
            ('Contabilita', 'cont_fine_voci_abbreviate', '120'),
            ('Contabilita', 'abilitaconfigparser', '0'),
            ('Contabilita', 'idxsal', '20'),
            ('Contabilita', 'ricicla_da', 'COMPUTO'),

            ('Analisi', 'sicurezza', '15,00%'),
            ('Analisi', 'spese_generali', '15,00%'),
            ('Analisi', 'utile_impresa', '10,00%'),

            ('Importazione', 'ordina_computo', '1'),

            ('Lavoro', 'committente', '(str)'),
            ('Lavoro', 'stazioneAppaltante', '(str)'),
            ('Lavoro', 'progetto', '(str)'),
            ('Lavoro', 'rup', '(str)'),
            ('Lavoro', 'progettista', '(str)'),
            ('Lavoro', 'data', '(date)' + LeenoUtils.date2String(date.today(), 1)),
            ('Lavoro', 'revisione', '(str)'),
            ('Lavoro', 'dataRevisione', '(date)' + LeenoUtils.date2String(date.today(), 1)),

            ('ImpostazioniStampa', 'fileCopertine', '(str)'),
            ('ImpostazioniStampa', 'copertina', '(str)'),
            ('ImpostazioniStampa', 'intSx', '(str)[COMMITTENTE]'),
            ('ImpostazioniStampa', 'intCenter', '(str)'),
            ('ImpostazioniStampa', 'intDx', '(str)[PROGETTO]'),
            ('ImpostazioniStampa', 'ppSx', '(str)[OGGETTO]'),
            ('ImpostazioniStampa', 'ppCenter', '(str)'),
            ('ImpostazioniStampa', 'ppDx', '(str)Pagina [PAGINA] di [PAGINE]'),

            ('ImpostazioniExport', 'npElencoPrezzi', '(str)1'),
            ('ImpostazioniExport', 'npComputoMetrico', '(str)2'),
            ('ImpostazioniExport', 'npCostiManodopera', '(str)3'),
            ('ImpostazioniExport', 'npQuadroEconomico', '(str)4'),
            ('ImpostazioniExport', 'cbElencoPrezzi', '(bool)True'),
            ('ImpostazioniExport', 'cbComputoMetrico', '(bool)True'),
            ('ImpostazioniExport', 'cbCostiManodopera', '(bool)True'),
            ('ImpostazioniExport', 'cbQuadroEconomico', '(bool)True'),

        )

        for param in parametri:
            try:
                self._parser.get(param[0], param[1])
            except Exception:
                if not self._parser.has_section(param[0]):
                    self._parser.add_section(param[0])
                self._parser.set(param[0], param[1], param[2])

    def _load(self):
        '''
        load configuration data from disk
        '''
        try:
            self._parser.read(self._path)
        except Exception:
            os.remove(self._path)

    def _store(self):
        '''
        store configuration data to disk
        '''
        fp = open(self._path, 'w')
        self._parser.write(fp)
        fp.close()

    def read(self, section, option, convert=False):
        '''
        read an option from config
        if convert is True, do the string->value conversion
        (for latter, the string must have the correct format)
        '''
        try:
            return self._parser.get(section, option)
        except Exception:
            return None

    def readBlock(self, section, convert=False):
        '''
        read a block of options from config given the section name
        if convert is True, do the string->value conversion
        (for latter, the strings must have the correct format)
        '''
        options = self._parser.options(section)
        res = {}
        for option in options:
            val = self._parser.get(section, option)
            if val is None:
                continue
            if convert:
                try:
                    val = PersistUtils.string2var(val)
                    if val is not None:
                        res[option] = val
                except Exception:
                    pass
            else:
                res[option] = val
        return res

    def write(self, section, option, val):
        '''
        write an option to config
        '''
        if not self._parser.has_section(section):
            self._parser.add_section(section)
        self._parser.set(section, option, val)

        # we must store on each write, because we can't know
        # when LeenoConfig gets destroyed with python
        self._store()

    def writeBlock(self, section, valDict, convert=False):
        '''
        write a block of options to config given the section name
        if convert is True, do the value->string conversion
        '''
        if not self._parser.has_section(section):
            self._parser.add_section(section)
        for key, val in valDict.items():
            if convert:
                val = PersistUtils.var2string(val)
                if val is None:
                    continue
            self._parser.set(section, key, val)
        self._store()
