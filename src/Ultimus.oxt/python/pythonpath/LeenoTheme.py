import uno
import LeenoUtils
import LeenoConfig
import LeenoDialogs as DLG

from com.sun.star.beans import PropertyValue
from contextlib import ExitStack

def scegliere_colore_picker(ctx, colore_iniziale):
    """
    Apre il dialogo ColorPicker di LibreOffice per la scelta di un colore.
    Utilizza XPropertyAccess per compatibilità con versioni recenti (25.8+).
    """
    try:
        smgr = ctx.ServiceManager
        picker = smgr.createInstanceWithContext("com.sun.star.ui.dialogs.ColorPicker", ctx)

        if hasattr(picker, "initialize"):
            picker.initialize(())

        # Utilizzo della utility interna di LeenO per le proprietà
        props = LeenoUtils.dictToProperties({"Color": int(colore_iniziale)})
        picker.setPropertyValues(props)

        if picker.execute() == 1:
            # Estrazione sicura del valore dalle proprietà restituite
            for p in picker.getPropertyValues():
                if p.Name == "Color":
                    return int(p.Value)
    except Exception as e:
        DLG.errore(f"Errore nel selettore colori: {str(e)}")
    return None


from undo_utils import with_undo, with_undo_batch, no_undo
@with_undo
@LeenoUtils.no_refresh
def applica_nuovo_colore_tematico():
    """
    Identifica il colore della cella attiva e lo sostituisce in tutti gli stili
    del documento, gestendo la protezione di ogni singolo foglio.
    """
    doc = LeenoUtils.getDocument()
    ctx = LeenoUtils.getComponentContext()

    # 1. Rilevazione colore sorgente
    oSel = doc.CurrentController.getSelection()
    oCell = oSel.getCellByPosition(0, 0) if hasattr(oSel, "getCellByPosition") else oSel

    try:
        colore_vecchio = oCell.CellBackColor
        if colore_vecchio == -1:
            DLG.chi("La cella selezionata non ha un colore di sfondo.")
            return
    except:
        return

    # 2. Scelta nuovo colore
    nuovo_colore = scegliere_colore_picker(ctx, colore_vecchio)
    # nuovo_colore = 16773375
    # DLG.chi(f"applicato il colore: {nuovo_colore}", OFF = True)

    if nuovo_colore is None or int(nuovo_colore) == int(colore_vecchio):
        return

    # 3. Sostituzione massiva con gestione protezione
    count = 0
    try:
        # ExitStack permette di gestire N fogli protetti contemporaneamente
        with ExitStack() as stack:
            fogli = doc.getSheets()
            for i in range(fogli.Count):
                # Utilizziamo il Context Manager nativo di LeenO
                stack.enter_context(LeenoUtils.ProtezioneFoglioContext(fogli.getByIndex(i)))

            # Modifica degli stili di cella
            cell_styles = doc.StyleFamilies.getByName("CellStyles")
            for sName in cell_styles.getElementNames():
                oStyle = cell_styles.getByName(sName)
                try:
                    if int(oStyle.CellBackColor) == int(colore_vecchio):
                        oStyle.CellBackColor = int(nuovo_colore)
                        count += 1
                except:
                    continue # Salta stili di sistema Read-Only

        # DLG.chi(f"Aggiornamento completato!\n\nStili modificati: {count}")
    except Exception as e:
        DLG.errore(f"Errore durante l'applicazione del tema: {str(e)}")

######################################################################
def trova_colore_cella():
    oDoc = LeenoUtils.getDocument()
    active_cell = oDoc.CurrentSelection
    DLG.chi(active_cell.CellBackColor)
    return

########################################################################

@LeenoUtils.no_refresh
def catalogo_stili_cella():
    '''
    Apre un nuovo foglio e vi inserisce tutti gli stili di cella
    con relativo esempio
    '''
    oDoc = LeenoUtils.getDocument()
    sty = oDoc.StyleFamilies.getByName("CellStyles").getElementNames()
    if oDoc.Sheets.hasByName("stili"):
        oSheet = oDoc.Sheets.getByName("stili")
    else:
        sheet = oDoc.createInstance("com.sun.star.sheet.Spreadsheet")
        oDoc.Sheets.insertByName('stili', sheet)
        oSheet = oDoc.Sheets.getByName("stili")
    from pyleeno import GotoSheet
    GotoSheet("stili")
    # attiva la progressbar
    indicator = oDoc.getCurrentController().getStatusIndicator()
    indicator.start('Creazione catalogo stili di cella in corso...', len(sty))
    i = 0
    sty = sorted(sty)
    for el in sty:
        oSheet.getCellByPosition( 0, i).String = el
        oSheet.getCellByPosition( 1, i).CellStyle = el
        oSheet.getCellByPosition( 3, i).CellStyle = el
        oSheet.getCellByPosition( 1, i).Value = -2000
        oSheet.getCellByPosition( 3, i).String = "LeenO"
        i += 1
        indicator.setValue(i)
    indicator.end()


@LeenoUtils.no_refresh
def elimina_stili_cella():
    '''
    Elimina gli stili di cella non utilizzati.
    '''
    oDoc = LeenoUtils.getDocument()
    stili = oDoc.StyleFamilies.getByName('CellStyles').getElementNames()

    # Crea una lista di stili non utilizzati
    stili_da_elim = [el for el in stili if not oDoc.StyleFamilies.getByName('CellStyles').getByName(el).isInUse()]
    #  stili_da_elim = stili # RIMUOVI TUTTI!!!

    # Rimuovi gli stili non utilizzati
    n = 0
    for el in stili_da_elim:
        oDoc.StyleFamilies.getByName('CellStyles').removeByName(el)
        n += 1
    import Dialogs
    Dialogs.Exclamation(Title = 'ATTENZIONE!', Text=f'Eliminati {n} stili di cella!')


def elenca_stili_foglio():
    '''
    Restituisce l'elenco di tutti gli stili di cella applicati alle celle nel foglio corrente.
    '''
    try:
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet

        # Set per tenere traccia degli stili unici applicati
        stili_applicati = set()

        # Ottieni l'area utilizzata nel foglio
        import SheetUtils
        area_utilizzata = SheetUtils.getUsedArea(oSheet)
        row = area_utilizzata.EndRow
        col = area_utilizzata.EndColumn

        # Itera sulle celle dell'area utilizzata
        for riga in range(row + 1):  # Includi l'ultima riga
            for colonna in range(col + 1):  # Includi l'ultima colonna
                cella = oSheet.getCellByPosition(colonna, riga)
                stile = cella.CellStyle
                if stile:  # Controlla se lo stile non è vuoto
                    stili_applicati.add(stile)

        # Converti il set in una lista
        lista_stili_applicati = list(stili_applicati)

        # Mostra o restituisci la lista degli stili applicati
        #  DLG.chi(f'Stili di cella applicati: {", ".join(lista_stili_applicati)}')
        return lista_stili_applicati

    except Exception as e:
        DLG.errore(e)
        return []


def elimina_stile():
    '''
    Elimina lo stile della cella selezionata.
    '''
    stili_utili = elenca_stili_foglio().append('Default')
    try:
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.CurrentController.ActiveSheet
        selezione = oDoc.getCurrentSelection()
        stile = selezione.CellStyle
        DLG.chi(stile)
        selezione.CellStyle = 'Default' # Assegna lo stile predefinito alla cella
        # Rimuovi lo stile
        if stile not in stili_utili:
            oDoc.StyleFamilies.getByName('CellStyles').removeByName(stile)
    except Exception as e:
        DLG.errore(e)
        pass

    # Ottieni la posizione attuale del cursore
    cella_corrente = selezione.getCellAddress()
    nuova_riga = cella_corrente.Row + 1  # Sposta di una riga in basso

    # Assicurati di non uscire dall'intervallo delle righe del foglio
    if nuova_riga < oSheet.Rows.Count:
        # Sposta il cursore alla cella nella stessa colonna ma una riga sotto
        oDoc.CurrentController.select(oSheet.getCellByPosition(cella_corrente.Column, nuova_riga))

########################################################################################



import LeenoUtils
import LeenoConfig
import LeenoDialogs as DLG
from contextlib import ExitStack

def recupera_decimali_da_stile(nome_stile):
    """Analizza il formato numerico di uno stile e restituisce il numero di decimali."""
    try:
        doc = LeenoUtils.getDocument()
        cell_styles = doc.StyleFamilies.getByName("CellStyles")
        if not cell_styles.hasByName(nome_stile):
            return None

        oStyle = cell_styles.getByName(nome_stile)
        num_formats = doc.getNumberFormats()
        format_entry = num_formats.getByKey(oStyle.NumberFormat)
        format_str = format_entry.FormatString

        if "," in format_str:
            # Estrae la parte decimale dopo la virgola
            return len(format_str.split(",")[1].replace(";", "").replace(")", "").strip())
        return 0
    except:
        return None

def dialogo_gestione_decimali():
    """Dialogo IDE con vincolo: Quantità e Sommano condividono lo stesso valore."""
    ctx = LeenoUtils.getComponentContext()
    smgr = ctx.ServiceManager
    dp = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)

    try:
        oDlg = dp.createDialog("vnd.sun.star.script:UltimusFree2.DlgDecimali?language=Basic&location=application")
    except:
        return

    cfg = LeenoConfig.Config()
    conf_values = cfg.readBlock('DecimaliStili')

    # Mappatura: 'chiave': (NomeComboBox, StileRiferimento)
    # Nota: cbBLU comanda sia quantità che sommano
    mappa_controlli = {
        'parti_uguali': ("cbPU", "comp 1-a PU"),
        'lunghezza': ("cbLUNG", "comp 1-a LUNG"),
        'larghezza': ("cbLARG", "comp 1-a LARG"),
        'pesi': ("cbPESO", "comp 1-a peso"),
        'quantità': ("cbBLU", "Blu")
    }

    lista_decimali = ("0", "1", "2", "3", "4", "5", "6")

    for key, (ctrl_name, style_name) in mappa_controlli.items():
        ctrl = oDlg.getControl(ctrl_name)
        if ctrl:
            ctrl.addItems(lista_decimali, 0)
            val_doc = recupera_decimali_da_stile(style_name)
            ctrl.Text = str(val_doc) if val_doc is not None else str(conf_values.get(key, "2"))

    # Gestione specifica per cbSOMMANO (se vuoi comunque popolarlo o nasconderlo)
    cbSommano = oDlg.getControl("cbSOMMANO")
    if cbSommano:
        cbSommano.addItems(lista_decimali, 0)
        val_sommano = recupera_decimali_da_stile("Comp-Variante num sotto")
        cbSommano.Text = str(val_sommano) if val_sommano is not None else str(conf_values.get('sommano', "2"))

    if oDlg.execute() == 1:
        nuovi_valori = {}
        for key, (ctrl_name, _) in mappa_controlli.items():
            nuovi_valori[key] = oDlg.getControl(ctrl_name).Text

        # VINCOLO: Il valore di 'sommano' viene forzato a quello di 'quantità' (cbBLU)
        nuovi_valori['sommano'] = nuovi_valori['quantità']

        cfg.writeBlock('DecimaliStili', nuovi_valori)
        aggiorna_decimali_documento(nuovi_valori)

    oDlg.dispose()

@LeenoUtils.no_refresh
def aggiorna_decimali_documento(mappa_decimali):
    """Applica i formati numerici garantendo l'uguaglianza tra Quantità e Sommano."""
    doc = LeenoUtils.getDocument()
    num_formats = doc.getNumberFormats()
    locale = doc.CharLocale
    cell_styles = doc.StyleFamilies.getByName("CellStyles")

    config_stili = {
        'parti_uguali': ["comp 1-a PU", "comp 1-a PU ROSSO"],
        'lunghezza': ["comp 1-a LUNG", "comp 1-a LUNG ROSSO"],
        'larghezza': ["comp 1-a LARG", "comp 1-a LARG ROSSO"],
        'pesi': ["comp 1-a peso", "comp 1-a peso ROSSO"],
        'quantità': ["Blu", "Blu ROSSO"],
        'sommano': ["Comp-Variante num sotto", "Comp-Variante num sotto ROSSO"]
    }

    with ExitStack() as stack:
        for sheet in doc.getSheets():
            stack.enter_context(LeenoUtils.ProtezioneFoglioContext(sheet))

        for cat, n in mappa_decimali.items():
            if cat in config_stili:
                n_int = int(n) if str(n).isdigit() else 2
                fmt = "#.##0" + ("," + "0" * n_int if n_int > 0 else "")

                key = num_formats.queryKey(fmt, locale, True)
                if key == -1:
                    key = num_formats.addNew(fmt, locale)

                for nome_stile in config_stili[cat]:
                    if cell_styles.hasByName(nome_stile):
                        cell_styles.getByName(nome_stile).NumberFormat = key
