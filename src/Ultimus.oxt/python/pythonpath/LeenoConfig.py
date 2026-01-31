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

PASTEL_COLORS = [
    16764108,  # Pastel Red
    16777164,  # Pastel Yellow
    13434828,  # Pastel Green
    13421823,  # Pastel Blue
    14737632,  # Pastel Gray

    16751052,  # Pastel Pink
    16770790,  # Pastel Orange
    13430526,  # Pastel Mint
    13421772,  # Pastel Cyan
    15132390,  # Pastel Lavender

    15527148,  # Pastel Beige
    14540253,  # Pastel Olive
    14079702,  # Pastel Teal
    15128749,  # Pastel Purple
    15987699   # Pastel Light Gray
]


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

            ('Generale', 'colorazione_categorie', 'Nessuno'),

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

            # In LeenoConfig.py -> _initDefaults()
            ('DecimaliStili', 'parti_uguali', '2'),
            ('DecimaliStili', 'lunghezza', '2'),
            ('DecimaliStili', 'larghezza', '2'),
            ('DecimaliStili', 'pesi', '3'),
            ('DecimaliStili', 'quantità', '2'),
            ('DecimaliStili', 'sommano', '2'),
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

#################################################################################

def MENU_leeno_conf():
    '''
    Visualizza il menù di configurazione
    '''
    cfg = Config()
    oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('S1'):
        Toolbars.AllOff()
        return
    psm = LeenoUtils.getServiceManager()
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDlg_config = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.Dlg_config?language=Basic&location=application"
    )

    oSheets = list(oDoc.getSheets().getElementNames())
    for nome in ('M1', 'S1', 'S2', 'S5', 'Elenco Prezzi', 'COMPUTO'):
        if nome in oSheets:
            oSheets.remove(nome)

    for nome in oSheets:
        oSheet = oDoc.getSheets().getByName(nome)
        if not oSheet.IsVisible:
            oDlg_config.getControl('CheckBox2').State = 0
            test = 0
            break
        oDlg_config.getControl('CheckBox2').State = 1
        test = 1

    if oDoc.getSheets().getByName("copyright_LeenO").IsVisible:
        oDlg_config.getControl('CheckBox2').State = 1

    # precisione come mostrato
    if cfg.read('Generale', 'precisione_come_mostrato') == 'True':
        oDoc.CalcAsShown = True
        oDlg_config.getControl('CheckBox4').State = 1

    if cfg.read('Generale', 'pesca_auto') == '1':
        oDlg_config.getControl('CheckBox1').State = 1

    if cfg.read('Generale', 'toolbar_contestuali') == '1':
        oDlg_config.getControl('CheckBox6').State = 1

    # --- NUOVA GESTIONE COLORI CATEGORIA (ComboBox3) ---
    sComboColori = oDlg_config.getControl('ComboBox3')
    val_colori = cfg.read('Generale', 'colorazione_categorie')
    # Default a "Nessuno" se non presente nel file conf
    sComboColori.Text = val_colori if val_colori else "Nessuno"
    # ----------------------------------------------------

    oSheet = oDoc.getSheets().getByName('S5')
    if not oSheet.getCellRangeByName('C9').IsMerged:
        oDlg_config.getControl('CheckBox5').State = 1
    else:
        oDlg_config.getControl('CheckBox5').State = 0

    oDlg_config.getControl('ComboBox6').Text = cfg.read('Generale', 'altezza_celle')

    sString = oDlg_config.getControl("ComboBox2")
    if cfg.read('Generale', 'movedirection') == '1':
        sString.Text = 'A DESTRA'
    elif cfg.read('Generale', 'movedirection') == '0':
        sString.Text = 'IN BASSO'

    oSheet = oDoc.getSheets().getByName('S1')

    # fullscreen
    oLayout = oDoc.CurrentController.getFrame().LayoutManager
    if not oLayout.isElementVisible('private:resource/toolbar/standardbar'):
        oDlg_config.getControl('CheckBox3').State = 1

    sString = oDlg_config.getControl('TextField14')
    sString.Text = oSheet.getCellRangeByName('S1.H334').String

    if oDoc.NamedRanges.hasByName("_Lib_1"):
        sString.setEnable(False)

    if cfg.read('Generale', 'torna_a_ep') == '1':
        oDlg_config.getControl('CheckBox8').State = 1

    if cfg.read('Generale', 'nuova_voce') == 'True':
        oDlg_config.getControl('CheckBox7').State = 1

    oDlg_config.getControl('ComboBox4').Text = cfg.read('Generale', 'copie_backup')
    if int(cfg.read('Generale', 'copie_backup')) != 0:
        oDlg_config.getControl('ComboBox5').Text = cfg.read('Generale', 'pausa_backup')

    # --- MOSTRA IL DIALOGO ---
    if oDlg_config.execute() != 1: # 1 solitamente è il tasto OK/Esegui
        return

    # --- SALVATAGGIO IMPOSTAZIONI ---
    import pyleeno as PL
    if oDlg_config.getControl('CheckBox2').State != test:
        PL.show_sheets(True if oDlg_config.getControl('CheckBox2').State == 1 else False)

    if oDlg_config.getControl('ComboBox1').getText() == 'Chiaro':
        PL.nuove_icone(True)
    elif oDlg_config.getControl('ComboBox1').getText() == 'Scuro':
        PL.nuove_icone(False)

    import LeenoToolbars as Toolbars
    Toolbars.Switch(False if oDlg_config.getControl('CheckBox3').State == 1 else True)

    if oDlg_config.getControl('CheckBox4').State == 1:
        cfg.write('Generale', 'precisione_come_mostrato', 'True')
        oDoc.CalcAsShown = True
    else:
        cfg.write('Generale', 'precisione_come_mostrato', 'False')
        oDoc.CalcAsShown = False

    cfg.write('Generale', 'nuova_voce', 'True' if oDlg_config.getControl('CheckBox7').State == 1 else 'False')

    ctx = LeenoUtils.getComponentContext()
    oGSheetSettings = ctx.ServiceManager.createInstanceWithContext("com.sun.star.sheet.GlobalSheetSettings", ctx)
    if oDlg_config.getControl('ComboBox2').getText() == 'IN BASSO':
        cfg.write('Generale', 'movedirection', '0')
        oGSheetSettings.MoveDirection = 0
    else:
        cfg.write('Generale', 'movedirection', '1')
        oGSheetSettings.MoveDirection = 1

    # --- SALVATAGGIO PREFERENZA COLORI ---
    cfg.write('Generale', 'colorazione_categorie', oDlg_config.getControl('ComboBox3').getText())
    # --------------------------------------

    cfg.write('Generale', 'altezza_celle', oDlg_config.getControl('ComboBox6').getText())
    cfg.write('Generale', 'pesca_auto', str(oDlg_config.getControl('CheckBox1').State))
    cfg.write('Generale', 'descrizione_in_una_colonna', str(oDlg_config.getControl('CheckBox5').State))
    cfg.write('Generale', 'toolbar_contestuali', str(oDlg_config.getControl('CheckBox6').State))

    Toolbars.Vedi()
    import pyleeno as PL
    PL.descrizione_in_una_colonna(False if oDlg_config.getControl('CheckBox5').State == 1 else True)

    cfg.write('Generale', 'torna_a_ep', str(oDlg_config.getControl('CheckBox8').State))

    if oDlg_config.getControl('TextField14').getText() != '10000':
        cfg.write('Generale', 'vedi_voce_breve', oDlg_config.getControl('TextField14').getText())
    oSheet.getCellRangeByName('S1.H334').Value = float(oDlg_config.getControl('TextField14').getText())

    import LeenoSheetUtils
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)

    cfg.write('Generale', 'copie_backup', oDlg_config.getControl('ComboBox4').getText())
    cfg.write('Generale', 'pausa_backup', oDlg_config.getControl('ComboBox5').getText())

    PL.autorun()
