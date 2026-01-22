import uno
import LeenoUtils
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
