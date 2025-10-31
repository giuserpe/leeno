'''
Funzioni di utilità per la manipolazione dei fogli
relativamente alle funzionalità specifiche di LeenO
'''
import uno
from com.sun.star.sheet.CellFlags import HARDATTR, EDITATTR, FORMATTED
from com.sun.star.beans import PropertyValue

import pyleeno as PL
import LeenoUtils
import SheetUtils
import LeenoSheetUtils
import LeenoAnalysis
import LeenoComputo
import Dialogs
import LeenoDialogs as DLG


def ScriviNomeDocumentoPrincipaleInFoglio(oSheet):
    '''
    Indica qual è il Documento Principale
    nell'apposita area del foglio corrente
    '''
    # legge il percorso del documento principale
    sUltimus = LeenoUtils.getGlobalVar('sUltimus')

    # dal foglio risale al documento proprietario
    oDoc = SheetUtils.getDocumentFromSheet(oSheet)

    # se si sta lavorando sul Documento Principale, non fa nulla
    try:
        if sUltimus == uno.fileUrlToSystemPath(oDoc.getURL()):
            return
    except Exception:
        # file senza nome
        return

    d = {
        'COMPUTO': 'F1',
        'VARIANTE': 'F1',
        'Elenco Prezzi': 'A1',
        'CONTABILITA': 'F1',
        'Analisi di Prezzo': 'A1'
    }
    cell = d.get(oSheet.Name)
    if cell is None:
        return

    oSheet.getCellRangeByName(cell).String = 'DP: ' + sUltimus
    oSheet.getCellRangeByName("A1:AT1").clearContents(EDITATTR + FORMATTED + HARDATTR)

# ###############################################################

def SbiancaCellePrintArea():
    '''
    area
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    oPrintArea = oSheet.getPrintAreas()

    oSheet.getCellRangeByPosition(
        oPrintArea[0].StartColumn, oPrintArea[0].StartRow,
        oPrintArea[0].EndColumn, oPrintArea[0].EndRow
        ).CellBackColor = 16777215 #sbianca

    stili_cat = LeenoUtils.getGlobalVar('stili_cat')
    for y in range(0, oPrintArea[0].EndRow):
        if oSheet.getCellByPosition(0, y).CellStyle in stili_cat:
            # conserva il colore di sfondo delle categorie
            # ~ oSheet.getCellRangeByPosition(0, y, 40, y).clearContents(HARDATTR)
            # attribuisce il grigio
            oSheet.getCellRangeByPosition(0, y, 40, y).CellBackColor = int('eeeeee', 16)
    return

########################################################################

def DelPrintSheetArea ():
    '''
    Cancella area di stampa del foglio corrente
    '''
    LeenoUtils.DocumentRefresh(True)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.setPrintAreas(())
    return

# ###############################################################

def setVisibilitaColonne(oSheet, sValori):
    '''
    sValori { string } : una stringa di configurazione della visibilità colonne
    permette di visualizzare/nascondere un set di colonne
    T = visualizza
    F = nasconde
    '''
    n = 0
    for el in sValori:
        oSheet.getCellByPosition(n, 2).Columns.IsVisible = True if el == 'T' else False
        n += 1

# ###############################################################

def getColumnWidths(oSheet):
    '''
    Legge tutte le impostazioni delle colonne 
    Restituisce: dict con widths, hidden_ranges, group_levels, visible_cols, freeze
    '''
    def col_index_to_letter(n):
        name = ''
        while n >= 0:
            name = chr(ord('A') + (n % 26)) + name
            n = n // 26 - 1
            if n < 0:
                break
        return name or 'A'

    config = {
        'widths': {},
        'hidden_ranges': [],
        'group_levels': {},
        'visible_cols': '',
        'freeze': None
    }

    try:
        # 1. Rileva colonne/righe bloccate
        try:
            freeze_cols = oSheet.FreezeColumns
            freeze_rows = oSheet.FreezeRows
            if freeze_cols > 0 or freeze_rows > 0:
                config['freeze'] = (freeze_cols, freeze_rows)
        except Exception as e:
            print(f"Errore freeze: {str(e)}")

        # 2. Analizza tutte le colonne
        cols = oSheet.getColumns()
        col_count = 26
        visible_cols = []
        
        for col_idx in range(col_count):
            try:
                col = cols.getByIndex(col_idx)
                col_letter = col_index_to_letter(col_idx)
                
                # Larghezza colonna
                width = col.Width
                config['widths'][col_letter] = width
                
                # Visibilità colonna (invertita in Calc)
                visible_cols.append('T' if col.IsVisible else 'F')
                
                # Livelli di raggruppamento
                outline_level = col.getPropertyValue("OutlineLevel")
                if outline_level > 0:
                    config['group_levels'][col_letter] = outline_level
                    
            except Exception as e:
                print(f"Errore colonna {col_idx}: {str(e)}")
                continue

        # 3. Elabora i risultati
        config['visible_cols'] = optimize_visibility_string(''.join(visible_cols))
        config['hidden_ranges'] = find_hidden_ranges(visible_cols)
        
    except Exception as e:
        print(f"Errore generale: {str(e)}")
    
    return config

def optimize_visibility_string(s):
    '''Compatta la stringa di visibilità per colonne consecutive'''
    from itertools import groupby
    result = []
    for k, g in groupby(s):
        count = len(list(g))
        if count > 5:
            result.append(f"'{k}'*{count}")
        else:
            result.append(k * count)
    return '+'.join(result)

def find_hidden_ranges(visible_list):
    '''Identifica range di colonne nascoste in formato (start, end)'''
    ranges = []
    start = None
    
    for i, visible in enumerate(visible_list):
        if visible == 'F':
            if start is None:
                start = i
        elif start is not None:
            ranges.append((start, i-1))
            start = None
    
    if start is not None:
        ranges.append((start, len(visible_list)-1))
    
    return [r for r in ranges if r[1] >= r[0]]

def generate_config_snippet(oSheet):
    '''Genera la configurazione pronta per l'uso'''
    config = getColumnWidths(oSheet)
    
    if not any([config['widths'], config['freeze'], config['group_levels'], 'F' in config['visible_cols']]):
        return "Nessuna impostazione personalizzata trovata"

    snippet = [f"'{oSheet.getName()}': {{"]
    
    # Widths
    if config['widths']:
        snippet.append("'widths': {")
        for col in sorted(config['widths'], key=lambda x: (len(x), x)):
            snippet.append(f"'{col}': {config['widths'][col]},")
        snippet.append("},")
    
    # Visibilità
    # if 'F' in config['visible_cols']:
    #     snippet.append(f"'visible_cols': \"{config['visible_cols']}\",")
    
    # Freeze
    if config['freeze']:
        snippet.append(f"'freeze': {config['freeze']},")
    
    # Raggruppamenti
    if config['group_levels']:
        snippet.append(f"'group_levels': {config['group_levels']},")
    
    # Nascoste
    if config['hidden_ranges']:
        snippet.append(f"'hidden_ranges': {config['hidden_ranges']},")
    
    snippet.append("}")
    return ' '.join(snippet)

def show_config_snippet():
    '''Mostra la configurazione in una dialog'''
    try:
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        
        config = generate_config_snippet(oSheet)
        
        if "Nessuna impostazione" in config:
            DLG.chi("Nessuna impostazione personalizzata trovata")
        else:
            DLG.chi(
                "=== CONFIGURAZIONE COLONNE ===\n"
                "Copia questo codice:\n\n" +
                config +
                "\n\nSostituisci nel dizionario sheet_configs"
            )
            
    except Exception as e:
        DLG.chi(f"ERRORE: {str(e)}\nAssicurati che:\n1. Un foglio sia aperto\n2. Le macro siano abilitate")
# Esempio di sheet_configs risultante
'''
sheet_configs = {
    'Elenco Prezzi': {
        'widths': {'A': 1535, 'B': 11407, ...},
        'visible_cols': "T'*12+'F'*188+'T'*200",
        'hidden_ranges': [(12, 199)],
        'group_levels': {'C': 1, 'D': 2},
        'freeze': (0, 3)
    }
}
'''



def setLarghezzaColonne(oSheet):
    '''
    Regola la larghezza delle colonne e altre proprietà visive in base al foglio specifico
    Args:
        oSheet: L'oggetto foglio di lavoro su cui operare
    '''
    # Configurazioni predefinite per tutti i fogli
    SHEET_CONFIGS = {
        'Analisi di Prezzo': {
            'widths': {
                'A': 1600, 'B': 11000, 'C': 1500, 'D': 1500, 'E': 1500,
                'F': 1500, 'G': 1500, 'H': 2000, 'I': 1900, 'J': 1900, 'K': 1900
            },
            'freeze': (0, 2),
            'visible_cols': None
        },
        'CONTABILITA': {
            'widths': {
                'A': 600, 'B': 1500, 'C': 8700, 'F': 1300, 'G': 1300,
                'H': 1300, 'I': 1300, 'J': 1700, 'L': 1700, 'N': 1900,
                'P': 1900, 'T': 1000, 'U': 1000, 'W': 1000, 'X': 1000,
                'Z': 1900, 'AC': 1700, 'AD': 1800, 'AE': 1800,
                'AX': 1900, 'AY': 1900
            },
            'freeze': (0, 3),
            'visible_cols': 'TTTFFTTTTTFTFTFTFTFTTFTTFTFTTTTFFFFFF',
            'range_widths': [
                (13, 1023, 1900),
                (19, 23, 1000)
            ],
            'hidden_ranges': [
                (51, 1023)
            ]
        },
        'COMPUTO': {
            'widths': {
                'A': 600, 'B': 1500, 'C': 6300, 'F': 1500, 'G': 1300,
                'H': 1300, 'I': 1300, 'J': 1700, 'L': 1700, 'S': 1950,
                'AC': 1700, 'AD': 1800, 'AE': 1800
            },
            'freeze': (0, 3),
            'visible_cols': 'TTTFFTTTTTFTFFFFFFTFFFFFFFFFFFFFFFFFFFFFFFFFTT',
            'visible_ranges': [
                (5, 8)
            ]
        },
        'VARIANTE': {
            'widths': {
                'A': 600, 'B': 1500, 'C': 6300, 'F': 1500, 'G': 1300,
                'H': 1300, 'I': 1300, 'J': 1700, 'L': 1700, 'S': 1950,
                'AC': 1700, 'AD': 1800, 'AE': 1800
            },
            'freeze': (0, 3),
            'visible_cols': 'TTTFFTTTTTFTFFFFFFTFFFFFFFFFFFFFFFFFFFFFFFFFTT',
            'visible_ranges': [
                (5, 8)
            ]
        },
        'Elenco Prezzi': {
            'default': {
                'widths': { 'A': 1600, 'B': 9999, 'C': 1499, 'D': 1499,
                    'E': 1600, 'F': 1499, 'G': 1499, 'H': 1600, 'I': 1199,
                    'J': 1199, 'K': 101, 'L': 1600, 'M': 1600, 'N': 1600,
                    'O': 101, 'P': 1600, 'Q': 1600, 'R': 1600, 'S': 101,
                    'T': 2000, 'U': 2000, 'V': 2000, 'W': 101, 'X': 2000,
                    'Y': 2000, 'Z': 1500,
                },
                'hidden_ranges': [
                    (6, 9), (3, 3)
                ],
                'visible_ranges': [
                    (6, 9), (3, 3)
                ]
            },
            'COMPUTO_VARIANTE': {
                'widths': {
                    'A': 1600, 'B': 9999, 'C': 1499, 'D': 1499, 'E': 1600,
                    'F': 1499, 'G': 1499, 'H': 1600, 'I': 1199, 'J': 1199,
                    'K': 101, 'L': 1600, 'M': 1600, 'N': 1600, 'O': 101,
                    'P': 1600, 'Q': 1600, 'R': 1600, 'S': 101, 'T': 2000,
                    'U': 2000, 'V': 2000, 'W': 101, 'X': 2000, 'Y': 2000,
                    'Z': 1500
                },
                'hidden_ranges': [
                    (3, 3), (5, 9), (13, 13), (16, 17), (21, 21)
                ]
            },
            'COMPUTO_CONTABILITÀ': {
                'widths': {
                    'A': 1600, 'B': 9999, 'C': 1499, 'D': 1499, 'E': 1600,
                    'F': 1499, 'G': 1499, 'H': 1600, 'I': 1199, 'J': 1199,
                    'K': 101, 'L': 1600, 'M': 1600, 'N': 1600, 'O': 101,
                    'P': 1600, 'Q': 1600, 'R': 1600, 'S': 101, 'T': 2000,
                    'U': 2000, 'V': 2000, 'W': 101, 'X': 2000, 'Y': 2000,
                    'Z': 1500,
                },
                'hidden_ranges': [
                    (3, 3), (5, 9), (12, 12), (15, 15), (17, 17), (20, 20)
                ],
            },
            'VARIANTE_CONTABILITÀ': {
                'widths': {
                    'A': 1600, 'B': 9999, 'C': 1499, 'D': 1499, 'E': 1600,
                    'F': 1499, 'G': 1499, 'H': 1600, 'I': 1199, 'J': 1199,
                    'K': 101, 'L': 1600, 'M': 1600, 'N': 1600, 'O': 101,
                    'P': 1600, 'Q': 1600, 'R': 1600, 'S': 101, 'T': 2000,
                    'U': 2000, 'V': 2000, 'W': 101, 'X': 2000, 'Y': 2000,
                    'Z': 1500,
                },
                'hidden_ranges': [
                    (3, 3), (6, 9), (11, 11), (15, 16), (19, 19)
                ],
            },
        }
    }
    with LeenoUtils.DocumentRefreshContext(False):
        memorizza_posizione()

        # Gestione speciale per Elenco Prezzi
        if oSheet.Name == 'Elenco Prezzi':
            oSheet.clearOutline()
            x1_value = oSheet.getCellRangeByName("X1").String

            if str(x1_value).strip() == 'COMPUTO_VARIANTE':
                variant = 'COMPUTO_VARIANTE'

            elif str(x1_value).strip() == 'COMPUTO_CONTABILITÀ':
                variant = 'COMPUTO_CONTABILITÀ'

            elif str(x1_value).strip() == 'VARIANTE_CONTABILITÀ':
                variant = 'VARIANTE_CONTABILITÀ'

            else:
                variant = 'default'

            config = SHEET_CONFIGS['Elenco Prezzi'].get(variant)

        else:
            config = SHEET_CONFIGS.get(oSheet.Name)

        if not config:
            ripristina_posizione()
            return

        # Applicazione delle configurazioni
        # try:
            # Impostazione larghezze colonne
        for col, width in config.get('widths', {}).items():
            oSheet.Columns[col].Width = width

        # Gestione proprietà speciali
        for col in config.get('optimal_widths', []):
            oSheet.Columns[col].OptimalWidth = True

        for start, end, width in config.get('range_widths', []):
            oSheet.getCellRangeByPosition(start, 0, end, 0).Columns.Width = width

        # Gestione visibilità colonne
        iSheet = oSheet.RangeAddress.Sheet
        oCellRangeAddr = uno.createUnoStruct(
            'com.sun.star.table.CellRangeAddress')
        oCellRangeAddr.Sheet = iSheet
        for start, end in config.get('hidden_ranges', []):
            oCellRangeAddr.StartColumn = start
            oCellRangeAddr.EndColumn = end
            # oSheet.ungroup(oCellRangeAddr, 0)
            # PL.struttura_off()
            oSheet.group(oCellRangeAddr, 0)
            oSheet.getCellRangeByPosition(start, 0, end, 0).Columns.IsVisible = False

        for start, end in config.get('visible_ranges', []):
            oSheet.getCellRangeByPosition(start, 0, end, 0).Columns.IsVisible = True

        if 'visible_cols' in config and config['visible_cols']:
            setVisibilitaColonne(oSheet, config['visible_cols'])

        if 'freeze' in config:
            SheetUtils.freezeRowCol(oSheet, *config['freeze'])

        adattaAltezzaRiga(oSheet)
        # finally:
        ripristina_posizione()
# ###############################################################
def rRow(oSheet):
    '''
    Restituisce la posizione della riga rossa
    '''
    nRow = SheetUtils.getLastUsedRow(oSheet) +10
    for n in reversed(range(0, nRow)):
        if oSheet.getCellByPosition(
                0,
                n).CellStyle == 'Riga_rossa_Chiudi':
            return n

def cercaUltimaVoce(oSheet):
    # if oSheet.Name == 'Elenco Prezzi':
    #     nRow = SheetUtils.getLastUsedRow(oSheet)
    #     return nRow

    # DLG.chi(nRow)
    nRow = SheetUtils.getLastUsedRow(oSheet) +1
    if nRow == 0:
        return 0
    for n in reversed(range(0, nRow)):
        # if oSheet.getCellByPosition(0, n).CellStyle in('Comp TOTALI'):
        if oSheet.getCellByPosition(
                0,
                n).CellStyle in ('EP-aS', 'EP-Cs', 'An-sfondo-basso Att End',
                                 'Comp End Attributo', 'Comp End Attributo_R',
                                 'comp Int_colonna',
                                 'comp Int_colonna_R_prima',
                                 'Livello-0-scritta', 'Livello-1-scritta',
                                 'livello2 valuta',):
            break
    if n == 0:
        n = nRow
    return n


# ###############################################################


def cercaPartenza(oSheet, lrow):
    '''
    oSheet      foglio corrente
    lrow        riga corrente nel foglio
    Ritorna il nome del foglio [0] e l'id della riga di codice prezzo componente [1]
    il flag '#reg' solo per la contabilità.
    partenza = (nome_foglio, id_rcodice, flag_contabilità)
    '''
    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_contab = LeenoUtils.getGlobalVar('stili_contab')

    # COMPUTO, VARIANTE
    if oSheet.getCellByPosition(0, lrow).CellStyle in stili_computo:
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        partenza = (oSheet.Name, sStRange.RangeAddress.StartRow + 1)

    # CONTABILITA
    elif oSheet.getCellByPosition(0, lrow).CellStyle in stili_contab:
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)

        try:
            partenza = (oSheet.Name, sStRange.RangeAddress.StartRow + 1,
                        oSheet.getCellByPosition(22,
                        sStRange.RangeAddress.StartRow + 1).String)
        except:
            lrow = 3
            partenza = (oSheet.Name, lrow, '')

    # ANALISI o riga totale
    elif oSheet.getCellByPosition(0, lrow).CellStyle in ('An-lavoraz-Cod-sx', 'Comp TOTALI'):
        partenza = (oSheet.Name, lrow)

    # nulla di quanto sopra
    else:
        partenza = (oSheet.Name, lrow, '')

    return partenza


# ###############################################################


def selezionaVoce(oSheet, lrow):
    '''
    Restituisce inizio e fine riga di una voce in COMPUTO, VARIANTE,
    CONTABILITA o Analisi di Prezzo
    lrow { long }  : numero riga all'interno della voce
    '''
    if oSheet.Name in ('Elenco Prezzi'):
        return lrow, lrow

    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    elif oSheet.Name == 'Analisi di Prezzo':
        sStRange = LeenoAnalysis.circoscriveAnalisi(oSheet, lrow)
    ###
    elif oSheet.Name == 'CONTABILITA':
        partenza = cercaPartenza(oSheet, lrow)
        if partenza[2] == '#reg':
            PL.sblocca_cont()
            if LeenoUtils.getGlobalVar('sblocca_computo') == 0:
                return
            pass
        else:
            pass
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    else:
        raise

    SR = sStRange.RangeAddress.StartRow
    ER = sStRange.RangeAddress.EndRow
    # ~ oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, SR, 250, ER))
    return SR, ER


# ###############################################################


def prossimaVoce(oSheet, lrow, n=1, saltaCat=True):
    """
    Sposta il cursore prima o dopo la voce corrente restituendo l'ID riga.
    
    Args:
        oSheet (object): Foglio di lavoro
        lrow (int): Riga di riferimento
        n (int, optional): 
            0 = sposta prima della voce corrente
            1 = sposta dopo della voce corrente (default)
        saltaCat (bool, optional): Se True salta le categorie
    
    Returns:
        int: Nuova posizione di riga
    """
    # Precaricamento stili (più efficiente)
    STILI_CAT = set(LeenoUtils.getGlobalVar('stili_cat'))
    STILI_COMPUTO = set(LeenoUtils.getGlobalVar('stili_computo'))
    STILI_CONTAB = set(LeenoUtils.getGlobalVar('stili_contab'))
    NO_VOCE = set(LeenoUtils.getGlobalVar('noVoce'))
    STILI_VALIDI = STILI_COMPUTO | STILI_CONTAB
    
    # Stili da saltare (insieme per ricerca veloce)
    STILI_DA_SALTARE = {
        'uuuuu', 'Ultimus_centro_bordi_lati',
        'comp Int_colonna', 'ULTIMUS', 
        'ULTIMUS_1', 'ULTIMUS_2', 'ULTIMUS_3'
    }

    # Ottieni stile corrente (una sola chiamata)
    stile_corrente = oSheet.getCellByPosition(0, lrow).CellStyle

    # Gestione casi particolari
    if stile_corrente in STILI_CAT:
        lrow += 1
    elif stile_corrente in STILI_CONTAB:
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        nSal = int(oSheet.getCellByPosition(23, sStRange.RangeAddress.StartRow + 1).Value)

    # Caso riga iniziale
    if lrow == 0:
        while stile_corrente not in STILI_VALIDI:
            lrow += 1
            stile_corrente = oSheet.getCellByPosition(0, lrow).CellStyle
        return lrow

    # Trova fine documento
    fine_doc = cercaUltimaVoce(oSheet) + 1
    if lrow >= fine_doc:
        return lrow

    # Logica principale di spostamento
    if saltaCat and stile_corrente in STILI_CAT:
        return lrow + 1

    if stile_corrente in STILI_VALIDI:
        voce_range = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        if n == 0:
            lrow = voce_range.RangeAddress.StartRow
        elif n == 1:
            lrow = voce_range.RangeAddress.EndRow + 1

    # Salta righe con stili particolari
    while stile_corrente in STILI_DA_SALTARE:
        lrow += 1
        stile_corrente = oSheet.getCellByPosition(0, lrow).CellStyle

    return lrow
# ###############################################################

def eliminaVoce(oSheet, lrow):
    '''
    usata in PL.MENU_elimina_voci_azzerate()

    Elimina una voce in COMPUTO, VARIANTE, CONTABILITA o Analisi di Prezzo
    lrow { long }  : numero riga
    '''
    voce = selezionaVoce(oSheet, lrow)
    SR = voce[0]
    ER = voce[1]

    oSheet.getRows().removeByIndex(SR, ER - SR + 1)

def elimina_voce(lrow=None, msg=1):
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoSheetUtils.eliminaVoce'
    Elimina una voce in COMPUTO, VARIANTE, CONTABILITA o Analisi di Prezzo
    lrow { long }  : numero riga
    msg  { bit }   : 1 chiedi conferma con messaggio
                     0 esegui senza conferma
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    if oSheet.Name == 'Elenco Prezzi':
        Dialogs.Info(Title = 'Info', Text="""Per eliminare una o più voci dall'Elenco Prezzi
devi selezionarle ed utilizzare il comando 'Elimina righe' di Calc.""")
        return

    if oSheet.Name not in ('COMPUTO', 'CONTABILITA', 'VARIANTE', 'Analisi di Prezzo'):
        return

    try:
        SR = PL.seleziona_voce()[0]
    except:
        return
    ER = PL.seleziona_voce()[1]
    if msg == 1:
        oDoc.CurrentController.select(oSheet.getCellRangeByPosition(
            0, SR, 250, ER))
        if '$C$' in oSheet.getCellByPosition(9, ER).queryDependents(False).AbsoluteName:
            undo = 1
            PL._gotoCella(9, ER)
            PL.comando ('ClearArrowDependents')
            PL.comando ('ShowDependents')
            oDoc.CurrentController.select(oSheet.getCellRangeByPosition(
                0, SR, 250, ER))
            messaggio= """
Da questa voce dipende almeno un Vedi Voce.
VUOI PROCEDERE UGUALMENTE?"""
        else:
            messaggio = """OPERAZIONE NON ANNULLABILE!\n
Stai per eliminare la voce selezionata.
            Voi Procedere?\n"""
        # ~return
        if Dialogs.YesNoDialog(IconType="warning",Title='*** A T T E N Z I O N E ! ***',
            Text= messaggio) == 1:
            try:
                undo
                PL.comando ('Undo')
            except:
                pass
            oSheet.getRows().removeByIndex(SR, ER - SR + 1)
            PL._gotoCella(0, SR+1)
        else:
            PL.comando ('Undo')
            oDoc.CurrentController.select(oSheet.getCellRangeByPosition(
                0, SR, 250, ER))
            return
    elif msg == 0:
        oSheet.getRows().removeByIndex(SR, ER - SR + 1)
    if oSheet.Name != 'Analisi di Prezzo':
        PL.numera_voci(0)
    else:
        PL._gotoCella(0, SR+2)
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))

# ###############################################################

def inserisciRigaRossa(oSheet):
    '''
    Inserisce la riga rossa di chiusura degli elaborati nel foglio specificato
    Questa riga è un riferimento per varie operazioni
    Errore se il foglio non è un foglio di LeenO
    '''
    lrow = 0
    nome = oSheet.Name
    if nome in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        lrow = cercaUltimaVoce(oSheet) + 2
        for n in range(lrow, SheetUtils.getLastUsedRow(oSheet) + 2):
            if oSheet.getCellByPosition(0, n).CellStyle == 'Riga_rossa_Chiudi':
                return
        oSheet.getRows().insertByIndex(lrow, 1)
        oSheet.getCellByPosition(0, lrow).String = 'Fine Computo'
        oSheet.getCellRangeByPosition(0, lrow, 34, lrow).CellStyle = 'Riga_rossa_Chiudi'
        oSheet.getCellByPosition(2, lrow
        ).String = 'Questa riga NON deve essere cancellata, MAI!!!(ma può rimanere tranquillamente NASCOSTA!)'
    elif nome == 'Analisi di Prezzo':
        lrow = cercaUltimaVoce(oSheet) + 2
        oSheet.getCellByPosition(0, lrow).String = 'Fine ANALISI'
        oSheet.getCellRangeByPosition(0, lrow, 10, lrow).CellStyle = 'Riga_rossa_Chiudi'
        oSheet.getCellByPosition(1, lrow
        ).String = 'Questa riga NON deve essere cancellata, MAI!!!(ma può rimanere tranquillamente NASCOSTA!)'
        oDoc = LeenoUtils.getDocument()
        # ~oSheet = oDoc.getSheets().getByName('Analisi di prezzo')
        SheetUtils.NominaArea(oDoc, 'Analisi di Prezzo',
                      '$A$3:$K$' + str(SheetUtils.getUsedArea(oSheet).EndRow), 'analisi')
    elif nome == 'Elenco Prezzi':
        lrow = cercaUltimaVoce(oSheet)
        if oSheet.getCellByPosition(0, lrow).CellStyle != 'Riga_rossa_Chiudi':
            lrow += 1

        oSheet.getCellByPosition(0, lrow).String = 'Fine elenco'
        oSheet.getCellByPosition(1, lrow
        ).String = 'Questa riga NON deve essere cancellata, MAI!!!(ma può rimanere tranquillamente NASCOSTA!)'

        oSheet.getCellRangeByPosition(0, lrow, 9, lrow).CellStyle = 'Riga_rossa_Chiudi'
        # oSheet.getCellRangeByPosition(11, lrow, 25, lrow).CellStyle = 'EP statistiche_q'
        # oSheet.getCellRangeByPosition(25, lrow, 25, lrow).CellStyle = 'EP-mezzo %'

        # oSheet.getCellByPosition(19, lrow).Formula = '=SUBTOTAL(9;T:T)'
        # oSheet.getCellByPosition(21, lrow).Formula = '=SUBTOTAL(9;U:U)'
        # oSheet.getCellByPosition(21, lrow).Formula = '=SUBTOTAL(9;V:V)'
        # oSheet.getCellByPosition(23, lrow).Formula = '=SUBTOTAL(9;X:X)'
        # oSheet.getCellByPosition(24, lrow).Formula = '=SUBTOTAL(9;Y:Y)'
        # oSheet.getCellByPosition(25, lrow).Formula = '=Z2'


# ###############################################################
from com.sun.star.beans import PropertyValue

def setAdatta():
    # ~da sistemare
    '''
    altezza   { integer } : altezza
    fissa il valore dell'altezza ottimale
    '''
    # oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    oProp = PropertyValue()
    oProp.Name = 'aExtraHeight'
    oProp.Value = 10
    properties = (oProp, )
    dispatchHelper.executeDispatch(oFrame, '.uno:SetOptimalRowHeight', '', 0,
                                   properties)

# def adattaAltezzaRiga(oSheet=False):
#     '''
#     Adatta l'altezza delle righe al contenuto delle celle.
#     imposta l'altezza ottimale delle celle
#     usata in PL.Menu_adattaAltezzaRiga()
#     '''
#     LeenoUtils.DocumentRefresh(True)
#     # qui il refresh manda in freeze
#     # ~LeenoUtils.DocumentRefresh(False)

#     oDoc = LeenoUtils.getDocument()
#     if not oSheet:
#         oSheet = oDoc.CurrentController.ActiveSheet
#     usedArea = SheetUtils.getUsedArea(oSheet)
#     # ~oSheet.getCellRangeByPosition(0, 0, usedArea.EndColumn, usedArea.EndRow).Rows.OptimalHeight = True
#     oSheet.Rows.OptimalHeight = True
#     if oSheet.Name in ('Elenco Prezzi', 'VARIANTE', 'COMPUTO', 'CONTABILITA'):
#         oSheet.getCellByPosition(0, 2).Rows.Height = 800
#     # DALLA VERSIONE 6.4.2 IL PROBLEMA è RISOLTO
#     # DALLA VERSIONE 7 IL PROBLEMA è PRESENTE
#     if float(PL.loVersion()[:5].replace('.', '')) >= 642:
#         return

#     # se la versione di LibreOffice è maggiore della 5.2
#     # esegue il comando agendo direttamente sullo stile
#     lista_stili = ('comp 1-a', 'Comp-Bianche in mezzo Descr_R',
#                    'Comp-Bianche in mezzo Descr', 'EP-a',
#                    'Ultimus_centro_bordi_lati')
#     # NELLE VERSIONI DA 5.4.2 A 6.4.1
#     if 520 < float(PL.loVersion()[:5].replace('.', '')) < 642:
#         for stile_cella in lista_stili:
#             try:
#                 oDoc.StyleFamilies.getByName("CellStyles").getByName(stile_cella).IsTextWrapped = True
#             except Exception:
#                 pass

#         test = usedArea.EndRow + 1

#         for y in range(0, test):
#             if oSheet.getCellByPosition(2, y).CellStyle in lista_stili:
#                 oSheet.getCellRangeByPosition(0, y, usedArea.EndColumn, y).Rows.OptimalHeight = True

#     if oSheet.Name == 'Elenco Prezzi':
#         test = usedArea.EndRow + 1
#         for y in range(0, test):
#             oSheet.getCellRangeByPosition(0, y, usedArea.EndColumn, y).Rows.OptimalHeight = True
#     return

def adattaAltezzaRiga(oSheet=False):
    """
    Adatta l'altezza delle righe al contenuto delle celle in modo ottimizzato.
    Versione bilanciata tra velocità e manutenibilità.
    """
    # Configurazioni (modificabili)
    memorizza_posizione()
    with LeenoUtils.DocumentRefreshContext(True):
        STILI_CELLA = {
            'comp 1-a', 
            'Comp-Bianche in mezzo Descr_R',
            'Comp-Bianche in mezzo Descr', 
            'EP-a',
            'Ultimus_centro_bordi_lati'
        }
        FOGLI_SPECIALI = {'Elenco Prezzi', 'VARIANTE', 'COMPUTO', 'CONTABILITA'}
        RIGA_SPECIALE = 2
        ALTEZZA_SPECIALE = 1050

        try:
            # --- INIZIALIZZAZIONE VELOCE ---
            LeenoUtils.DocumentRefresh(True)
            oDoc = LeenoUtils.getDocument()
            oSheet = oSheet or oDoc.CurrentController.ActiveSheet
            usedArea = SheetUtils.getUsedArea(oSheet)
            versione_lo = float(PL.loVersion()[:5].replace('.', ''))  # Chiamata UNICA

            # --- OPERAZIONE PRINCIPALE (velocizzata) ---
            oSheet.Rows.OptimalHeight = True  # Applica a tutto il foglio in un colpo solo

            # --- CASI SPECIALI (ottimizzati) ---
            if oSheet.Name in FOGLI_SPECIALI:
                # Imposta altezza fissa per riga speciale
                oSheet.getCellByPosition(0, RIGA_SPECIALE).Rows.Height = ALTEZZA_SPECIALE

                # Ottimizzazione per 'Elenco Prezzi': evita loop se non necessario
                if oSheet.Name == 'Elenco Prezzi' and usedArea.EndRow > 0:
                    oSheet.Rows.OptimalHeight = True  # Già fatto sopra, ma ripetuto per sicurezza

            # --- GESTIONE VERSIONI LO (5.4.2 - 6.4.1) ---
            if 520 < versione_lo < 642:
                cell_styles = oDoc.StyleFamilies.getByName("CellStyles")  # Prende gli stili UNA volta
                for stile in STILI_CELLA:
                    try:
                        cell_styles.getByName(stile).IsTextWrapped = True
                    except Exception:
                        continue  # Ignora stili mancanti senza log (più veloce)

                # Ottimizzazione: usa 'getCellRangeByPosition' solo per righe con stili speciali
                for y in range(0, usedArea.EndRow + 1):
                    if oSheet.getCellByPosition(2, y).CellStyle in STILI_CELLA:
                        oSheet.getRows().getByIndex(y).OptimalHeight = True  # Più veloce di getCellRangeByPosition

        except Exception as e:
            print(f"Errore in adattaAltezzaRiga: {str(e)}")  # Log essenziale
            raise  # Rilancia per gestione esterna
    ripristina_posizione()
# ###############################################################


def inserSuperCapitolo(oSheet, lrow, sTesto='Super Categoria'):
    '''
    lrow    { double } : id della riga di inserimento
    sTesto  { string } : titolo della categoria
    '''
    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
    if not sTesto:
        sTesto ='senza_titolo'
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Default':
        # se oltre la riga rossa
        lrow -= 2
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Riga_rossa_Chiudi':
        # se riga rossa
        lrow -= 1

    oSheet.getRows().insertByIndex(lrow, 1)
    oSheet.getCellByPosition(2, lrow).String = sTesto

    # inserisco i valori e le formule
    oSheet.getCellRangeByPosition(0, lrow, 36, lrow).CellStyle = 'Livello-0-scritta'
    oSheet.getCellRangeByPosition(2, lrow, 17, lrow).CellStyle = 'Livello-0-scritta mini'
    oSheet.getCellRangeByPosition( 18, lrow, 18, lrow).CellStyle = 'Livello-0-scritta mini val'
    oSheet.getCellRangeByPosition(24, lrow, 24, lrow).CellStyle = 'Livello-0-scritta mini %'
    oSheet.getCellRangeByPosition(29, lrow, 29, lrow).CellStyle = 'Livello-0-scritta mini %'
    oSheet.getCellRangeByPosition(30, lrow, 30, lrow).CellStyle = 'Livello-0-scritta mini val'
    oSheet.getCellRangeByPosition(2, lrow, 11, lrow).merge(True)

    # rinumero e ricalcolo
    # ocellBaseA = oSheet.getCellByPosition(1, lrow)
    # ocellBaseR = oSheet.getCellByPosition(31, lrow)
    lrowProvv = lrow - 1
    while oSheet.getCellByPosition(31, lrowProvv).CellStyle != 'Livello-0-scritta':
        if lrowProvv > 4:
            lrowProvv -= 1
        else:
            break
    oSheet.getCellByPosition(31, lrow).Value = oSheet.getCellByPosition(1, lrowProvv).Value + 1


# ###############################################################


def inserCapitolo(oSheet, lrow, sTesto='Categoria'):
    '''
    lrow    { double } : id della riga di inserimento
    sTesto  { string } : titolo della categoria
    '''
    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
    if not sTesto:
        sTesto ='senza_titolo'

    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Default':
        # se oltre la riga rossa
        lrow -= 2
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Riga_rossa_Chiudi':
        # se riga rossa
        lrow -= 1
    oSheet.getRows().insertByIndex(lrow, 1)
    oSheet.getCellByPosition(2, lrow).String = sTesto

    # inserisco i valori e le formule
    oSheet.getCellRangeByPosition(0, lrow, 36, lrow).CellStyle = 'Livello-1-scritta'
    oSheet.getCellRangeByPosition(2, lrow, 17, lrow).CellStyle = 'Livello-1-scritta mini'
    oSheet.getCellRangeByPosition(18, lrow, 18, lrow).CellStyle = 'Livello-1-scritta mini val'
    oSheet.getCellRangeByPosition(24, lrow, 24, lrow).CellStyle = 'Livello-1-scritta mini %'
    oSheet.getCellRangeByPosition(29, lrow, 29, lrow).CellStyle = 'Livello-1-scritta mini %'
    oSheet.getCellRangeByPosition(30, lrow, 30, lrow).CellStyle = 'Livello-1-scritta mini val'
    oSheet.getCellRangeByPosition(2, lrow, 11, lrow).merge(True)

    # rinumero e ricalcolo
    # ocellBaseA = oSheet.getCellByPosition(1, lrow)
    # ocellBaseR = oSheet.getCellByPosition(31, lrow)
    lrowProvv = lrow - 1
    while oSheet.getCellByPosition(31, lrowProvv).CellStyle != 'Livello-1-scritta':
        if lrowProvv > 4:
            lrowProvv -= 1
        else:
            break
    oSheet.getCellByPosition(31, lrow).Value = oSheet.getCellByPosition(1, lrowProvv).Value + 1


# ###############################################################


def inserSottoCapitolo(oSheet, lrow, sTesto):
    '''
    lrow    { double } : id della riga di inserimento
    sTesto  { string } : titolo della sottocategoria
    '''
    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
    if not sTesto:
        sTesto ='senza_titolo'

    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Default':
        # se oltre la riga rossa
        lrow -= 2
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Riga_rossa_Chiudi':
        # se riga rossa
        lrow -= 1

    oSheet.getRows().insertByIndex(lrow, 1)
    oSheet.getCellByPosition(2, lrow).String = sTesto

    # inserisco i valori e le formule
    oSheet.getCellRangeByPosition(0, lrow, 36,lrow).CellStyle = 'livello2 valuta'
    oSheet.getCellRangeByPosition(2, lrow, 17, lrow).CellStyle = 'livello2_'
    oSheet.getCellRangeByPosition(18, lrow, 18, lrow).CellStyle = 'livello2 scritta mini'
    oSheet.getCellRangeByPosition(24, lrow, 24, lrow).CellStyle = 'livello2 valuta mini %'
    oSheet.getCellRangeByPosition(29, lrow, 29, lrow).CellStyle = 'livello2 valuta mini %'
    oSheet.getCellRangeByPosition(30, lrow, 30, lrow).CellStyle = 'livello2 valuta mini'
    oSheet.getCellRangeByPosition(31, lrow, 33, lrow).CellStyle = 'livello2_'
    oSheet.getCellRangeByPosition(2, lrow, 11, lrow).merge(True)

    # oSheet.getCellByPosition(1, lrow).Formula = '=AF' + str(lrow+1) + '''&"."&''' + 'AG' + str(lrow+1)
    # rinumero e ricalcolo
    # ocellBaseA = oSheet.getCellByPosition(1, lrow)
    # ocellBaseR = oSheet.getCellByPosition(31, lrow)

    lrowProvv = lrow - 1
    while oSheet.getCellByPosition(32, lrowProvv).CellStyle != 'livello2 valuta':
        if lrowProvv > 4:
            lrowProvv -= 1
        else:
            break
    oSheet.getCellByPosition(
        32, lrow).Value = oSheet.getCellByPosition(1, lrowProvv).Value + 1
    lrowProvv = lrow - 1
    while oSheet.getCellByPosition(31, lrowProvv).CellStyle != 'Livello-1-scritta':
        if lrowProvv > 4:
            lrowProvv -= 1
        else:
            break
    oSheet.getCellByPosition(31, lrow).Value = oSheet.getCellByPosition(1, lrowProvv).Value
    # SubSum_Cap(lrow)


# ###############################################################


def invertiUnSegno(oSheet, lrow):
    '''
    Inverte il segno delle formule di quantità nel rigo di misurazione lrow.
    lrow    { int }  : riga di riferimento
    usata con XPWE_it
    '''
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        if 'comp 1-a' in oSheet.getCellByPosition(2, lrow).CellStyle:
            if 'ROSSO' in oSheet.getCellByPosition(2, lrow).CellStyle:
                oSheet.getCellByPosition(
                    9, lrow
                ).Formula = '=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(
                    lrow + 1) + ')=0;"";PRODUCT(E' + str(
                        lrow + 1) + ':I' + str(lrow + 1) + '))'
                          # se VediVoce
                # ~ if oSheet.getCellByPosition(4, lrow).Type.value != 'EMPTY':
                # ~ oSheet.getCellByPosition(9, lrow).Formula='=IF(PRODUCT(E' +
                # str(lrow+1) + ':I' + str(lrow+1) + ')=0;"";PRODUCT(E' +
                # str(lrow+1) + ':I' + str(lrow+1) + '))' # se VediVoce
                # ~ else:
                # ~ oSheet.getCellByPosition(9, lrow).Formula=
                # '=IF(PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) +
                # ')=0;"";PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + '))'
                for x in range(2, 10):
                    oSheet.getCellByPosition(
                        x, lrow).CellStyle = oSheet.getCellByPosition(
                            x, lrow).CellStyle.split(' ROSSO')[0]
            else:
                oSheet.getCellByPosition(
                    9, lrow
                ).Formula = '=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(
                    lrow + 1) + ')=0;"";-PRODUCT(E' + str(
                        lrow + 1) + ':I' + str(lrow + 1) + '))'  # se VediVoce
                # ~ if oSheet.getCellByPosition(4, lrow).Type.value != 'EMPTY':
                # ~ oSheet.getCellByPosition(9, lrow).Formula =
                # '=IF(PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + ')=0;
                # "";-PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + '))' # se VediVoce
                # ~ else:
                # ~ oSheet.getCellByPosition(9, lrow).Formula =
                # '=IF(PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + ')=0;
                # "";-PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + '))'
                for x in range(2, 10):
                    oSheet.getCellByPosition(
                        x, lrow).CellStyle = oSheet.getCellByPosition(
                            x, lrow).CellStyle + ' ROSSO'
    if oSheet.Name in ('CONTABILITA'):
        formula1 = oSheet.getCellByPosition(9, lrow).Formula
        formula2 = oSheet.getCellByPosition(11, lrow).Formula
        oSheet.getCellByPosition(11, lrow).Formula = formula1
        oSheet.getCellByPosition(9, lrow).Formula = formula2
        if oSheet.getCellByPosition(11, lrow).String != '':
            for x in range(2, 12):
                oSheet.getCellByPosition(
                    x, lrow).CellStyle = oSheet.getCellByPosition(
                        x, lrow).CellStyle + ' ROSSO'
        else:
            for x in range(2, 12):
                oSheet.getCellByPosition(
                    x, lrow).CellStyle = oSheet.getCellByPosition(
                        x, lrow).CellStyle.split(' ROSSO')[0]


# ###############################################################

def numeraVoci(oSheet, lrow, tutte):
    '''
    tutte { boolean }  : True  rinumera tutto
                       False rinumera dalla voce corrente in giù
    '''
    #qui il refresh è inutile
    lastRow = SheetUtils.getUsedArea(oSheet).EndRow + 1
    n = 1

    if not tutte:
        for x in reversed(range(0, lrow)):
            if(
               oSheet.getCellByPosition(1, x).CellStyle in ('comp Art-EP', 'comp Art-EP_R') and
               oSheet.getCellByPosition(1, x).CellBackColor != 15066597):
                n = oSheet.getCellByPosition(0, x).Value + 1
                break
        for row in range(lrow, lastRow):
            if oSheet.getCellByPosition(1, row).CellBackColor == 15066597:
                oSheet.getCellByPosition(0, row).String = ''
            elif oSheet.getCellByPosition(1,row).CellStyle in ('comp Art-EP', 'comp Art-EP_R'):
                oSheet.getCellByPosition(0, row).Value = n
                n += 1
    else:
        for row in range(0, lastRow):
            if oSheet.getCellByPosition(1, row).CellStyle in ('comp Art-EP','comp Art-EP_R'):
                oSheet.getCellByPosition(0, row).Value = n
                n += 1


# ###############################################################


def MENU_elimina_righe_vuote():
    '''Elimina le righe vuote negli elaborati di COMPUTO, VARIANTE o CONTABILITA'''
    with LeenoUtils.DocumentRefreshContext(False):
        LeenoSheetUtils.memorizza_posizione()
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        valid_sheets = ('COMPUTO', 'VARIANTE', 'CONTABILITA')

        if oSheet.Name not in valid_sheets:
            Dialogs.Exclamation(Title='Avviso!', Text=f'Puoi usare questo comando solo nelle tabelle {", ".join(valid_sheets)}.')
            return

        confirmation_text = f'Stai per eliminare tutte le righe\ndi misura vuote nell\'elaborato {oSheet.Name}.\n\nVuoi procedere?'
        if Dialogs.YesNoDialog(IconType="question", Title='ATTENZIONE!', Text=confirmation_text) == 0:
            return

        # lrow_c = PL.LeggiPosizioneCorrente()[1]
        sString = 'T O T A L E' if oSheet.Name == 'CONTABILITA' else 'TOTALI COMPUTO'
        lrow = SheetUtils.uFindStringCol(sString, 2, oSheet, start=2, equal=1, up=True) or SheetUtils.getLastUsedRow(oSheet)

        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start("", lrow)  # 100 = max progresso
        indicator.Text = 'Eliminazione delle righe vuote in corso...'

        for y in reversed(range(0, lrow)):
            indicator.Value = y

            if oSheet.getCellByPosition(0, y).CellStyle not in ('Comp Start Attributo', 'Comp Start Attributo_R'):
                row_has_data = any(oSheet.getCellByPosition(x, y).Type.value != 'EMPTY' for x in range(0, 8 + 1))

            if not row_has_data:
                oSheet.getRows().removeByIndex(y, 1)

        indicator.end()
        LeenoUtils.DocumentRefresh(True)

        lrow_ = SheetUtils.uFindStringCol(sString, 2, oSheet, start=2, equal=1, up=True) or SheetUtils.getLastUsedRow(oSheet)
        # PL._gotoCella(1, 4)
        # Dialogs.Info(Title='Ricerca conclusa', Text=f'Eliminate {lrow - lrow_} righe vuote.')
        LeenoSheetUtils.ripristina_posizione()


# ###############################################################


def MENU_SheetToDoc():
    '''
    Copia il foglio corrente in un nuovo documento.
    '''
    oDoc = LeenoUtils.getDocument()
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    oProp = []
    oProp0 = PropertyValue()
    oProp0.Name = 'DocName'
    oProp0.Value = ''
    oProp1 = PropertyValue()
    oProp1.Name = 'Index'
    oProp1.Value = 32767
    oProp2 = PropertyValue()
    oProp2.Name = 'Copy'
    oProp2.Value = True
    oProp.append(oProp0)
    oProp.append(oProp1)
    oProp.append(oProp2)
    properties = tuple(oProp)
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, '.uno:Move', '', 0, properties)
    oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect
    oDoc = LeenoUtils.getDocument()

    oSheet = oDoc.CurrentController.ActiveSheet

    if "COMPUTO" in oSheet.Name or "VARIANTE" in oSheet.Name:
        oDoc.CurrentController.select(oSheet.getCellRangeByName('A1:I1048576'))
        PL.comando('Copy')
        #oDoc.CurrentController.select(oCell)
        PL.paste_clip(insCells=0, pastevalue=True)
        oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselec
    oDoc.CurrentController.ZoomValue = 100
    return

# ###############################################################


def memorizza_posizione(step=0):
    """Memorizza la posizione corrente del cursore, con incremento opzionale della riga"""
    ctx = LeenoUtils.getComponentContext()
    doc = LeenoUtils.getDocument()
    controller = doc.getCurrentController()
    
    # Ottieni la selezione corrente
    selection = controller.getSelection()
    
    # Gestione per diversi tipi di selezione
    if selection.supportsService("com.sun.star.sheet.SheetCell"):
        # Singola cella
        cell_addr = selection.getCellAddress()
        pos_data = {
            'type': 'cell',
            'sheet': cell_addr.Sheet,
            'col': cell_addr.Column,
            'row': cell_addr.Row + step  # incremento opzionale
        }
    elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):
        # Range di celle
        range_addr = selection.getRangeAddress()
        pos_data = {
            'type': 'range',
            'sheet': range_addr.Sheet,
            'col': range_addr.StartColumn,
            'row': range_addr.StartRow + step,      # incremento opzionale
            'end_col': range_addr.EndColumn,
            'end_row': range_addr.EndRow + step     # incremento opzionale
        }
    else:
        DLG.chi("Tipo di selezione non supportato")
        return
    
    # Memorizza i dati
    LeenoUtils.setGlobalVar('ultima_posizione', pos_data)

    # DLG.chi(f"Posizione salvata: Foglio {pos_data['sheet']}, Riga {pos_data['row']}, Col {pos_data['col']}")

def ripristina_posizione():
    """Ripristina la posizione memorizzata"""
    pos_data = LeenoUtils.getGlobalVar('ultima_posizione')
    if not pos_data:
        DLG.chi("Nessuna posizione memorizzata trovata")
        return
    
    doc = LeenoUtils.getDocument()
    controller = doc.getCurrentController()
    sheets = doc.getSheets()
    
    try:
        sheet = sheets.getByIndex(pos_data['sheet'])
        
        if pos_data['type'] == 'cell':
            # Ripristina singola cella
            cell = sheet.getCellByPosition(pos_data['col'], pos_data['row'])
            controller.select(cell)
        else:
            # Ripristina range di celle
            cell_range = sheet.getCellRangeByPosition(
                pos_data['col'], pos_data['row'],
                pos_data['end_col'], pos_data['end_row']
            )
            controller.select(cell_range)
            
    except Exception as e:
        DLG.chi(f"Errore nel ripristino: {str(e)}")
    doc.CurrentController.select(doc.createInstance("com.sun.star.sheet.SheetCellRanges"))
