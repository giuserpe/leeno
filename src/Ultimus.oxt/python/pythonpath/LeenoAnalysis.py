'''
Funzioni relative alla gestione delle analisi di prezzi
'''

import os
import LeenoUtils
import LeenoGlobals
import SheetUtils
import LeenoSheetUtils
import LeenoEvents
import DocUtils
import Dialogs
import pyleeno as PL

import LeenoDialogs as DLG
from undo_utils import with_undo

@with_undo("Inserimento Oneri Sicurezza")
def Inserisci_Utili():
    oDoc = LeenoUtils.getDocument()
    oSheets = oDoc.getSheets()
    oRanges = oDoc.NamedRanges

    oSheetAP = oSheets.getByName("Analisi di Prezzo")

    # 1. Gestione del Range Nominato "oneri_sicurezza"
    if not oRanges.hasByName("oneri_sicurezza"):
        oCellAddress = oSheetAP.getCellRangeByName("B10").getCellAddress()
        oRanges.addNewByName("oneri_sicurezza", "$S5.$B$93:$P$93", oCellAddress, 0)

    import pyleeno as PL
    lrow = PL.LeggiPosizioneCorrente()[1]
    sStRange = circoscriveAnalisi(oSheetAP, lrow)
    srow = sStRange.RangeAddress.StartRow
    endRow = sStRange.RangeAddress.EndRow

    target_row = -1
    for i in range(srow, endRow):
        cell_a = oSheetAP.getCellByPosition(0, i).String
        cell_d = oSheetAP.getCellByPosition(3, i).String

        if "sicurezza" in cell_d:
            # DLG.chi("La riga degli oneri per la sicurezza è già inserita!")
            return

        if cell_a == "I":
            target_row = i
            break

    # 3. Esecuzione
    oSheetAP.getRows().insertByIndex(target_row, 1)

    oNamedRange = oRanges.getByName("oneri_sicurezza")
    oRangeAddress = oNamedRange.ReferredCells.getRangeAddress()
    oDestAddress = oSheetAP.getCellByPosition(0, target_row).getCellAddress()

    oSheetAP.copyRange(oDestAddress, oRangeAddress)

    # Seleziona la cella di descrizione appena inserita
    oDoc.CurrentController.select(oSheetAP.getCellByPosition(4, target_row))


def calcola_codice_successivo(codice_corrente):
    import re
    if not codice_corrente:
        return None
    codice_corrente = codice_corrente.strip()
    match = re.match(r"^(.*?)(\d+)$", codice_corrente)
    if match:
        prefix, num_str = match.groups()
        num_len = len(num_str)
        next_num = int(num_str) + 1
        return f"{prefix}{next_num:0{num_len}d}"
    else:
        return codice_corrente + "1"


def inizializza_analisi(oDoc=None, nuovaScheda=False):
    '''
    Prepara il foglio 'Analisi di Prezzo' copiandolo dal template master.
    Se nuovaScheda=True, inserisce anche la prima scheda vuota e sposta il cursore.
    Ritorna l'oggetto oSheet e la riga da cui iniziare la compilazione.
    '''
    import pyleeno as PL
    PL.chiudi_dialoghi()
    if oDoc is None:
        oDoc = LeenoUtils.getDocument()

    # Trova il codice corrente se siamo su Analisi di Prezzo prima di ogni modifica
    codice_corrente = None
    try:
        active_sheet = oDoc.CurrentController.ActiveSheet
        if active_sheet and active_sheet.Name == 'Analisi di Prezzo':
            lrow_curr = PL.LeggiPosizioneCorrente()[1]
            try:
                sStRange = circoscriveAnalisi(active_sheet, lrow_curr)
                if sStRange:
                    codice_corrente = active_sheet.getCellByPosition(0, sStRange.RangeAddress.StartRow + 2).String
            except Exception:
                pass
            
            # Fallback: scansiona all'indietro per l'ultima scheda se non trovato o vuoto
            if not codice_corrente or codice_corrente.strip() == "":
                urow = SheetUtils.getLastUsedRow(active_sheet)
                for row_idx in reversed(range(2, urow + 1)):
                    try:
                        if active_sheet.getCellByPosition(0, row_idx).CellStyle == 'An.1v-Att Start':
                            temp_code = active_sheet.getCellByPosition(0, row_idx + 1).String
                            if temp_code and temp_code.strip() != "":
                                codice_corrente = temp_code
                                break
                    except Exception:
                        continue
    except Exception:
        pass

    if not oDoc.getSheets().hasByName('Analisi di Prezzo'):
        # Costruisce il percorso del template
        template_path = os.path.join(LeenoGlobals.dest(), 'template', 'leeno', 'Computo_LeenO.ods')

        # Carica il template in modalità nascosta
        oTemplate = DocUtils.loadDocument(template_path, Hidden=True)
        if not oTemplate:
            Dialogs.Exclamation(
                Title='Analisi di Prezzo',
                Text=f'Impossibile caricare il template:\n{template_path}'
            )
            # Fallback estremo: crea foglio vuoto
            oDoc.getSheets().insertNewByName('Analisi di Prezzo', 1)
        else:
            try:
                # Determina la posizione di inserimento: dopo Elenco Prezzi o in posizione 1
                pos = 1
                if oDoc.getSheets().hasByName('Elenco Prezzi'):
                    pos = oDoc.getSheets().getByName('Elenco Prezzi').getRangeAddress().Sheet + 1
                
                # Importa il foglio Analisi di Prezzo dal template
                oDoc.getSheets().importSheet(oTemplate, 'Analisi di Prezzo', pos)
            except Exception as e:
                Dialogs.Exclamation(
                    Title='Analisi di Prezzo',
                    Text=f'Errore durante l\'importazione dal template:\n{str(e)}'
                )
                if not oDoc.getSheets().hasByName('Analisi di Prezzo'):
                    oDoc.getSheets().insertNewByName('Analisi di Prezzo', 1)
            finally:
                oTemplate.close(True)

        oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
        oSheet.TabColor = 12189608
        
        # Inizializza l'area nominale del blocco nel nuovo foglio
        SheetUtils.NominaArea(oDoc, 'S5', '$B$108:$P$133', 'blocco_analisi')
        
        LeenoEvents.assegna()
        LeenoSheetUtils.ScriviNomeDocumentoPrincipaleInFoglio(oSheet)

        # l'inizio del foglio per l'inserimento dati è riga 2 (dopo i primi due header)
        startRow = 2
        
        # Prepariamo gli indirizzi per l'eventuale inserimento scheda
        oRangeAddress = oDoc.NamedRanges.getByName('blocco_analisi').ReferredCells.RangeAddress
        oCellAddress = oSheet.getCellByPosition(0, startRow).getCellAddress()

    else:
        oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
        oSheet.IsVisible = True

        # Cerchiamo la riga di chiusura dell'ultima analisi presente
        lrow = LeenoSheetUtils.cercaUltimaVoce(oSheet) - 5
        urow = SheetUtils.getLastUsedRow(oSheet)
        
        # Inizializziamo n alla fine dell'area usata come fallback
        n = urow if urow >= 2 else 1
        
        # Scansioniamo solo indici validi (>= 2 per saltare gli header)
        for n_scan in range(max(2, lrow), urow + 1):
            try:
                if oSheet.getCellByPosition(0, n_scan).CellStyle == 'An-sfondo-basso Att End':
                    n = n_scan
                    break
            except Exception:
                continue
        
        oRangeAddress = oDoc.NamedRanges.getByName('blocco_analisi').ReferredCells.RangeAddress
        # la riga dalla quale l'eventuale nuova scheda deve partire è n + 2
        startRow = n + 2
        
        # Verifica finale di validità per startRow
        if startRow < 2:
            startRow = 2
            
        oCellAddress = oSheet.getCellByPosition(0, startRow).getCellAddress()

    # Assicura che il foglio sia visibile
    oDoc.CurrentController.setActiveSheet(oSheet)

    # Esegue l'inserimento solo se richiesto
    if nuovaScheda:
        # Se il foglio non è nuovo (o se vogliamo comunque inserire), aggiungiamo spazio
        if startRow > 2 or oSheet.getCellByPosition(0, startRow).Type.value != 'EMPTY':
            oSheet.getRows().insertByIndex(startRow, 26)
            
        oSheet.copyRange(oCellAddress, oRangeAddress)
        LeenoSheetUtils.inserisciRigaRossa(oSheet)
        
        # Riga finale per la scrittura dati
        startRow = startRow + 1
        if codice_corrente:
            next_code = calcola_codice_successivo(codice_corrente)
            if next_code:
                oSheet.getCellByPosition(0, startRow).String = next_code
        
        # Spostiamo il cursore sulla cella descrizione (colonna E = 4)
        oCell = oSheet.getCellByPosition(0, startRow)
        oDoc.CurrentController.setActiveSheet(oSheet)
        oDoc.CurrentController.select(oCell)

    return oSheet, startRow


def circoscriveAnalisi(oSheet, lrow):
    '''
    Circoscrive una voce di analisi partendo dalla posizione corrente del cursore

    Args:
        oSheet (object): Foglio di lavoro
        lrow (int): Riga di riferimento per la selezione dell'intera voce

    Returns:
        object: Intervallo di celle che rappresenta l'analisi
    '''
    # Pre-carica gli stili necessari
    stili_analisi = LeenoGlobals.getGlobalVar('stili_analisi')
    cell_style = oSheet.getCellByPosition(0, lrow).CellStyle

    # Variabili per i limiti
    start_row = 0
    end_row = SheetUtils.getLastUsedRow(oSheet)

    # Trova inizio analisi (cerca all'indietro)
    if cell_style in stili_analisi:
        # Ottimizzazione: usa xrange in Python 2 o range in Python 3
        for row in reversed(range(lrow)):
            if oSheet.getCellByPosition(0, row).CellStyle == 'Analisi_Sfondo':
                start_row = row
                break

        # Trova fine analisi (cerca in avanti)
        for row in range(lrow, end_row + 1):
            if oSheet.getCellByPosition(0, row).CellStyle == 'An-sfondo-basso Att End':
                end_row = row
                break

    # Restituisci l'intervallo trovato (250 colonne è un valore arbitrario)
    return oSheet.getCellRangeByPosition(0, start_row, 250, end_row)

def copiaRigaAnalisi(oSheet, lrow):
    '''
    Inserisce una nuova riga di misurazione in analisi di prezzo
    '''
    stile = oSheet.getCellByPosition(0, lrow).CellStyle
    if stile in ('An-lavoraz-desc', 'An-lavoraz-Cod-sx'):
        lrow = lrow + 1
        oSheet.getRows().insertByIndex(lrow, 1)
        # imposto gli stili
        oSheet.getCellByPosition(0, lrow).CellStyle = 'An-lavoraz-Cod-sx'
        oSheet.getCellRangeByPosition(1, lrow, 5, lrow).CellStyle = 'An-lavoraz-generica'
        oSheet.getCellByPosition(3, lrow).CellStyle = 'An-lavoraz-input'
        oSheet.getCellByPosition(6, lrow).CellStyle = 'An-senza'
        oSheet.getCellByPosition(7, lrow).CellStyle = 'An-senza-DX'
        # ci metto le formule
        #  oDoc.enableAutomaticCalculation(False)
        oSheet.getCellByPosition(1, lrow).Formula = (
           '=IF(A' + str(lrow + 1) +
           '="";"";CONCATENATE("  ";VLOOKUP(A' +
           str(lrow + 1) + ';elenco_prezzi;2;FALSE());' '))')
        oSheet.getCellByPosition(2, lrow).Formula = (
           '=IF(A' + str(lrow + 1) + '="";"";"  "&VLOOKUP(A' +
           str(lrow + 1) + ';elenco_prezzi;3;FALSE())&" ")')
        oSheet.getCellByPosition(3, lrow).Value = 0
        oSheet.getCellByPosition(4,lrow).Formula = (
           '=IF(A' + str(lrow + 1) + '="";0;VLOOKUP(A' +
           str(lrow + 1) + ';elenco_prezzi;5;FALSE()))')
        oSheet.getCellByPosition(5, lrow).Formula = (
           '=D' + str(lrow + 1) + '*E' + str(lrow + 1))
        oSheet.getCellByPosition(8, lrow).Formula = (
           '=IF(A' + str(lrow + 1) + '="";"";IF(VLOOKUP(A' +
           str(lrow + 1) + ';elenco_prezzi;6;FALSE())="";"";(VLOOKUP(A' +
           str(lrow + 1) + ';elenco_prezzi;6;FALSE()))))')
        oSheet.getCellByPosition(9, lrow).Formula = (
           '=IF(I' + str(lrow + 1) + '="";"";I' +
           str(lrow + 1) + '*F' + str(lrow + 1) + ')')
        if oSheet.getCellByPosition(1, lrow - 1).CellStyle == 'An-lavoraz-dx-senza-bordi':
            oRangeAddress = oSheet.getCellByPosition(0, lrow + 1).getRangeAddress()
            oCellAddress = oSheet.getCellByPosition(0, lrow).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
        oSheet.getCellByPosition(0, lrow).String = 'Cod. Art.?'


def MENU_impagina_analisi():
    '''
    PL.set_area_stampa()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name != 'Analisi di Prezzo':
        return
    lr = SheetUtils.getLastUsedRow(oSheet) + 1
    oSheet.removeAllManualPageBreaks()
    for el in range (1, lr):
        if oSheet.getCellByPosition(0, el).String == '----':
            if oSheet.getCellByPosition(0, el + 2).CellStyle != 'Ultimus_centro':
                oSheet.getCellByPosition(0, el + 2).Rows.IsStartOfNewPage = True
    '''

    PL.set_area_stampa()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # Esegui solo se il foglio attivo è 'Analisi di Prezzo'
    if oSheet.Name != 'Analisi di Prezzo':
        return

    # Calcola l'ultima riga utilizzata
    last_row = SheetUtils.getLastUsedRow(oSheet) + 1

    # Rimuovi tutte le interruzioni di pagina manuali
    oSheet.removeAllManualPageBreaks()

    # Imposta una nuova interruzione di pagina dopo ogni sezione individuata
    for row in range(1, last_row):
        cell = oSheet.getCellByPosition(0, row)
        if cell.String == '----':
            next_cell = oSheet.getCellByPosition(0, row + 2)
            if next_cell.CellStyle != 'Ultimus_centro':
                next_cell.Rows.IsStartOfNewPage = True

########################################################################

def Main_Riordina_Analisi_Alfabetico():
    _Riordina_Analisi_Alfabetico()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.Sheets.getByName("Analisi di Prezzo")
    import LeenoSheetUtils
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    return

@with_undo("Riordina Analisi di Prezzo Alfabetico")
@LeenoUtils.no_refresh
def _Riordina_Analisi_Alfabetico():
    PL.chiudi_dialoghi()

    # with LeenoUtils.DocumentRefreshContext(False):
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.Sheets.getByName("Analisi di Prezzo")
    lLastUrow = SheetUtils.getLastUsedRow(oSheet)

    # 1. Raccogli tutte le schede con le loro posizioni attuali
    schede = []  # Lista di tuple: (codice, riga_inizio, riga_fine)
    i = 0
    while i <= lLastUrow:
        cell = oSheet.getCellByPosition(0, i)
        if cell.CellStyle == "An.1v-Att Start":
            inizio = i
            codice = oSheet.getCellByPosition(0, i + 1).String
            fine = None
            for x in range(i + 1, lLastUrow + 1):
                if oSheet.getCellByPosition(0, x).CellStyle == "Analisi_Sfondo":
                    fine = x
                    break
            if fine is None:
                msg = f"Errore: scheda '{codice}' (riga {inizio + 1}) non ha riga di fine 'Analisi_Sfondo'"
                DLG.chi(msg)
                return
            if any(s[0] == codice for s in schede):
                msg = f"Mi fermo! Il codice:\n\t\t\t\t\t\t{codice}\nè presente più volte. Correggi e ripeti il comando."
                DLG.chi(msg)
                return
            schede.append((codice, inizio, fine))
            i = fine + 1
        else:
            i += 1

    if not schede:
        return

    # 2. Ordina alfabeticamente A-Z
    schede_ordinate = sorted(schede, key=lambda x: x[0])

    PL.struttura_off()

    # 3. Sposta ogni scheda.
    # FIX 1: target_index = riga della prima scheda (non hardcoded 1)
    #         così le intestazioni globali sopra rimangono intatte.
    target_index = schede[0][1]

    # FIX 2: processa in ordine INVERSO (Z→A) inserendo sempre in testa:
    #         il risultato finale sarà A-Z.
    for codice_target, _, _ in reversed(schede_ordinate):
        # Ricerca la posizione attuale della scheda (cambia ad ogni iterazione)
        trovata = False
        inizio = None
        fine = None
        lscanLimit = SheetUtils.getLastUsedRow(oSheet)

        for i in range(target_index, lscanLimit + 1):
            cell = oSheet.getCellByPosition(0, i)
            if cell.CellStyle == "An.1v-Att Start":
                if oSheet.getCellByPosition(0, i + 1).String == codice_target:
                    inizio = i
                    for x in range(i + 1, lscanLimit + 1):
                        if oSheet.getCellByPosition(0, x).CellStyle == "Analisi_Sfondo":
                            fine = x
                            trovata = True
                            break
                    break

        if not trovata:
            continue

        nrighe = fine - inizio + 1

        # Inserisce spazio esattamente uguale alla scheda da spostare
        oSheet.getRows().insertByIndex(target_index, nrighe)

        # La scheda originale è slittata in basso di nrighe
        original_inizio = inizio + nrighe
        original_fine = fine + nrighe

        # Copia il blocco nella nuova posizione
        selezione = oSheet.getCellRangeByPosition(0, original_inizio, 250, original_fine)
        oDest = oSheet.getCellByPosition(0, target_index).CellAddress
        oSheet.copyRange(oDest, selezione.RangeAddress)

        # Rimuove il blocco originale (ora vuoto)
        oSheet.getRows().removeByIndex(original_inizio, nrighe)

    PL.MENU_struttura_on()


def Main_Rinumera_Analisi_Selezionate():
    _Rinumera_Analisi_Selezionate()


@with_undo("Rinumera Analisi di Prezzo Selezionate")
@LeenoUtils.no_refresh
def _Rinumera_Analisi_Selezionate():
    '''
    Rinumera SOLO le schede di 'Analisi di Prezzo' intersecate dalla selezione
    corrente, chiedendo il primo indice della nuova numerazione, e propaga la
    sostituzione dei vecchi codici in 'Elenco Prezzi' (colonna A) e in tutti i
    fogli il cui nome contiene 'COMPUTO', 'VARIANTE' o 'CONTABILITA' (colonna B).

    Esempio: fornendo 'ciccillo_66' come primo codice ('ciccillo_' prefisso, 66 ID),
    le schede selezionate diventano ciccillo_66, ciccillo_67, ..., ciccillo_##
    (il numero di cifre dell'ID segue quello inserito, es. '066' -> 3 cifre).
    '''
    import re

    PL.chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()

    if not oDoc.Sheets.hasByName("Analisi di Prezzo"):
        DLG.chi("Il foglio 'Analisi di Prezzo' non esiste in questo documento.")
        return

    oSheetAP = oDoc.Sheets.getByName("Analisi di Prezzo")

    if oDoc.CurrentController.ActiveSheet.Name != "Analisi di Prezzo":
        DLG.chi("Attivare il foglio 'Analisi di Prezzo' prima di lanciare il comando.")
        return

    # 1. Righe coperte dalla selezione corrente (gestisce anche selezioni multiple)
    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
    except AttributeError:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
    el_y = []
    try:
        len(oRangeAddress)
        for el in oRangeAddress:
            el_y.append((el.StartRow, el.EndRow))
    except TypeError:
        el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))

    righe_selezionate = set()
    for start_row, end_row in el_y:
        righe_selezionate.update(range(start_row, end_row + 1))

    if not righe_selezionate:
        DLG.chi("Nessuna selezione valida sul foglio 'Analisi di Prezzo'.")
        return

    # 2. Individua tutte le schede del foglio (stesso algoritmo di _Riordina_Analisi_Alfabetico)
    lLastUrow = SheetUtils.getLastUsedRow(oSheetAP)
    schede = []  # Lista di tuple: (codice, riga_inizio, riga_fine)
    i = 0
    while i <= lLastUrow:
        cell = oSheetAP.getCellByPosition(0, i)
        if cell.CellStyle == "An.1v-Att Start":
            inizio = i
            codice = oSheetAP.getCellByPosition(0, i + 1).String
            fine = None
            for x in range(i + 1, lLastUrow + 1):
                if oSheetAP.getCellByPosition(0, x).CellStyle == "Analisi_Sfondo":
                    fine = x
                    break
            if fine is None:
                msg = f"Errore: scheda '{codice}' (riga {inizio + 1}) non ha riga di fine 'Analisi_Sfondo'"
                DLG.chi(msg)
                return
            schede.append((codice, inizio, fine))
            i = fine + 1
        else:
            i += 1

    if not schede:
        DLG.chi("Nessuna scheda di analisi trovata nel foglio 'Analisi di Prezzo'.")
        return

    # 3. Filtra solo le schede intersecate dalla selezione, in ordine di riga (dall'alto in basso)
    schede_selezionate = [
        s for s in schede if righe_selezionate.intersection(range(s[1], s[2] + 1))
    ]

    if not schede_selezionate:
        DLG.chi("Seleziona almeno una scheda di analisi (o una sua parte) prima di lanciare il comando.")
        return

    # 4. Chiede il primo codice della nuova numerazione (prefisso + ID), es. 'ciccillo_66'
    codice_riferimento = (schede_selezionate[0][0] or "").strip()
    sPrimoCodice = PL.InputBox(
        sCella=codice_riferimento,
        t=f"Primo codice della nuova numerazione, es. 'ciccillo_66' "
          f"({len(schede_selezionate)} scheda/e selezionate)"
    )
    if not sPrimoCodice or not sPrimoCodice.strip():
        return
    sPrimoCodice = sPrimoCodice.strip()

    # 5. Il prefisso e il numero di cifre dell'ID sono ricavati dal codice appena inserito
    match = re.match(r"^(.*?)(\d+)$", sPrimoCodice)
    if not match:
        DLG.chi(
            "Codice non valido: deve terminare con un numero (es. 'ciccillo_66'), "
            "in modo da poter ricavare prefisso e ID di partenza."
        )
        return
    prefix = match.group(1)
    width = len(match.group(2))
    primo_indice = int(match.group(2))

    # 6. Genera la mappa vecchio_codice -> nuovo_codice, in ordine di riga
    nuovi_codici = [
        f"{prefix}{primo_indice + idx:0{width}d}"
        for idx in range(len(schede_selezionate))
    ]
    mappa = {s[0]: n for s, n in zip(schede_selezionate, nuovi_codici)}

    # 7. Verifica che i nuovi codici non collidano con codici di schede NON coinvolte
    codici_non_toccati = {c for c, _, _ in schede if c not in mappa}
    collisioni = codici_non_toccati.intersection(nuovi_codici)
    if collisioni:
        msg = ("Mi fermo! I seguenti nuovi codici coinciderebbero con codici già "
               "presenti nel foglio:\n\t\t\t\t\t\t" + ", ".join(sorted(collisioni)))
        DLG.chi(msg)
        return
    if len(set(nuovi_codici)) != len(nuovi_codici):
        DLG.chi("Errore interno: la nuova numerazione genera codici duplicati.")
        return

    # 8. Conferma prima di un'operazione che tocca più fogli
    riepilogo = "\n".join(
        f"{s[0]} -> {n}" for s, n in zip(schede_selezionate, nuovi_codici)
    )
    if not Dialogs.YesNoDialog(
        Title="Rinumera Analisi di Prezzo",
        Text=f"Confermi la rinumerazione di {len(schede_selezionate)} scheda/e "
             f"e la propagazione in Elenco Prezzi e nei fogli "
             f"COMPUTO/VARIANTE/CONTABILITA?\n\n{riepilogo}"
    ):
        return

    # 9. Applica i nuovi codici nel foglio 'Analisi di Prezzo'
    for (codice, inizio, fine), nuovo_codice in zip(schede_selezionate, nuovi_codici):
        oSheetAP.getCellByPosition(0, inizio + 1).String = nuovo_codice

    # 10. Propaga la sostituzione dei codici in tutti gli altri fogli:
    #     - 'Elenco Prezzi' (nome esatto): codice in colonna A, come in 'Analisi di Prezzo'
    #     - fogli il cui nome contiene 'COMPUTO', 'VARIANTE' o 'CONTABILITA': codice in colonna B
    nomi_target = ("COMPUTO", "VARIANTE", "CONTABILITA")
    oSheets = oDoc.Sheets
    fogli_aggiornati = 0
    for nome_foglio in oSheets.ElementNames:
        if nome_foglio == "Analisi di Prezzo":
            continue
        if nome_foglio == "Elenco Prezzi":
            colonna_codice = 0
        elif any(t in nome_foglio.upper() for t in nomi_target):
            colonna_codice = 1
        else:
            continue
        oSheet = oSheets.getByName(nome_foglio)
        urow = SheetUtils.getLastUsedRow(oSheet)
        toccato = False
        for r in range(urow + 1):
            oCell = oSheet.getCellByPosition(colonna_codice, r)
            valore = oCell.String
            if valore in mappa:
                oCell.String = mappa[valore]
                toccato = True
        if toccato:
            fogli_aggiornati += 1

    DLG.chi(
        f"Rinumerazione completata: {len(schede_selezionate)} scheda/e aggiornate, "
        f"propagata su {fogli_aggiornati} foglio/i."
    )
