'''
Gestione delle impostazioni di LeenO
'''
import sys
import os
import configparser


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

            self._parser = configparser.ConfigParser()

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

            #  ('Computo', 'riga_bianca_categorie', '1'),
            #  ('Computo', 'voci_senza_numerazione', '0'),
            ('Computo', 'inizio_voci_abbreviate', '100'),
            ('Computo', 'fine_voci_abbreviate', '120'),
            ('Contabilità', 'cont_inizio_voci_abbreviate', '100'),
            ('Contabilità', 'cont_fine_voci_abbreviate', '120'),
            ('Contabilità', 'abilitaconfigparser', '0'),
            ('Contabilità', 'idxsal', '20'),
            ('Contabilità', 'ricicla_da', 'COMPUTO'))

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
        self._parser.read(self._path)

    def _store(self):
        '''
        store configuration data to disk
        '''
        fp = open(self._path, 'w')
        self._parser.write(fp)
        fp.close()

    def read(self, section, option):
        '''
        read an option from config
        '''
        try:
            return self._parser.get(section, option)
        except Exception:
            return None

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
