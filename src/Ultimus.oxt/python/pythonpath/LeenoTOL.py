"""
    LeenO - Aggregazione per Tipologie Omogenee di Lavorazioni (TOL)

    Classifica le voci del foglio 'Elenco Prezzi' per TOL (colonna AA),
    aggrega gli importi di Contabilita' (colonna V, base del SAL) per
    ciascuna TOL e produce un foglio di riepilogo con incidenze percentuali
    e indice sintetico, ai sensi dell'art. 60 D.Lgs. 36/2023 e Allegato II.2-bis.

    Colonne di Elenco Prezzi coinvolte:
        AA  - numero TOL (1-20) assegnato alla voce (validazione dati a discesa)
        V   - Importi Contabilita' (base per il calcolo del SAL)

    La colonna J ('Codice di origine') resta riservata all'uso del
    maintainer per l'annotazione della TOL riportata nei prezzari
    istituzionali: questo modulo non la legge ne' la scrive.

    Il foglio di riepilogo (SHEET_RIEPILOGO) viene creato se assente e
    ricalcolato ad ogni chiamata di MENU_leeno_aggiorna_riepilogo_tol().
"""

import LeenoUtils
import LeenoISTAT
import SheetUtils
import LeenoDialogs as DLG
import Dialogs


SHEET_ELENCO_PREZZI = 'Elenco Prezzi'
SHEET_RIEPILOGO = 'Riepilogo TOL'

COL_TOL = 26          # AA (0-based: A=0 ... Z=25, AA=26)
COL_IMPORTO_CONTAB = 21  # V  (0-based: V e' la 22a lettera -> indice 21)

RIGA_DATI_INIZIO = 2  # 0-based: i dati di Elenco Prezzi iniziano a riga 3 (indice 2)


# ---------------------------------------------------------------------------
# Aggregazione
# ---------------------------------------------------------------------------

def _get_ep_last_row(oSheet):
    """
    Stessa logica di LeenoNamedAreas._get_ep_last_row: cerca la riga
    sentinella 'Fine elenco', altrimenti usa l'ultima riga usata.

    ATTENZIONE: la riga restituita e' la riga sentinella stessa, che nel
    template contiene anche il totale ('TOTALE') della colonna V su tutto
    l'elenco. LeenoNamedAreas la include nel named range con un +1
    esplicito (serve per la stampa); qui va invece ESCLUSA dal ciclo di
    aggregazione, altrimenti il suo importo (= somma di tutto) finisce
    nel bucket 'non classificato' duplicando l'intero totale classificato.
    """
    row = SheetUtils.uFindStringCol('Fine elenco', 0, oSheet)
    if row is None:
        row = SheetUtils.getUsedArea(oSheet).EndRow
    return row


def aggrega_importi_per_tol(oDoc):
    """
    Scandisce 'Elenco Prezzi' e somma la colonna V (Importi Contabilita')
    per ciascun numero TOL trovato in colonna AA.

    Ritorna un dizionario {numero_tol (int): importo_totale (float)}.
    Le righe con AA vuota, non numerica, o fuori intervallo 1-20 vengono
    ignorate e conteggiate a parte in chiave None (per riscontro rapido
    di voci non ancora classificate).
    """
    if not oDoc.getSheets().hasByName(SHEET_ELENCO_PREZZI):
        raise RuntimeError("Foglio '%s' non trovato." % SHEET_ELENCO_PREZZI)

    oSheet = oDoc.getSheets().getByName(SHEET_ELENCO_PREZZI)
    riga_fine_elenco = _get_ep_last_row(oSheet)  # riga sentinella/TOTALE, ESCLUSA dal ciclo

    totali = {n: 0.0 for n in LeenoISTAT.TOL_CODICI}
    totali[None] = 0.0  # voci non classificate, con importo <> 0

    for riga in range(RIGA_DATI_INIZIO, riga_fine_elenco):
        cella_tol = oSheet.getCellByPosition(COL_TOL, riga)
        cella_importo = oSheet.getCellByPosition(COL_IMPORTO_CONTAB, riga)

        # difesa aggiuntiva: se per qualche motivo la sentinella non e' alla
        # riga attesa (es. rilevata con offset diverso), non farla comunque
        # confluire nel bucket 'non classificato'
        if oSheet.getCellByPosition(0, riga).getString().strip() == 'Fine elenco':
            continue

        importo = cella_importo.getValue()
        if importo == 0:
            continue

        testo_tol = cella_tol.getString().strip()
        numero_tol = None
        if testo_tol.isdigit():
            n = int(testo_tol)
            if n in LeenoISTAT.TOL_CODICI:
                numero_tol = n

        totali[numero_tol] = totali.get(numero_tol, 0.0) + importo

    return totali


def calcola_riepilogo(oDoc, periodo=None):
    """
    Ritorna una lista di dizionari, uno per TOL, con:
        numero, codice_sdmx, descrizione, importo, incidenza,
        periodo_indice, valore_indice
    e in coda l'indice sintetico complessivo (chiave 'indice_sintetico'
    nel dizionario di ritorno principale, insieme a 'importo_totale' e
    'importo_non_classificato').

    Gli indici ISTAT vengono letti via LeenoISTAT.get_indice_tol solo per
    le TOL con importo diverso da zero, per non superare inutilmente il
    rate limit del servizio.
    """
    totali = aggrega_importi_per_tol(oDoc)
    importo_non_classificato = totali.pop(None)
    importo_totale = sum(totali.values())

    righe = []
    indice_sintetico = 0.0

    for numero in sorted(totali):
        importo = totali[numero]
        codice_sdmx, descrizione = LeenoISTAT.TOL_CODICI[numero]
        incidenza = (importo / importo_totale) if importo_totale else 0.0

        periodo_indice = None
        valore_indice = None
        if importo != 0:
            try:
                periodo_indice, valore_indice = LeenoISTAT.get_indice_tol(numero, periodo)
                indice_sintetico += incidenza * valore_indice
            except (ValueError, RuntimeError) as e:
                # non interrompe il riepilogo per una singola TOL non raggiungibile:
                # la riga resta con indice vuoto, visibile nel foglio come anomalia
                DLG.chi("Indice TOL %d non disponibile: %s" % (numero, e))

        righe.append({
            'numero': numero,
            'codice_sdmx': codice_sdmx,
            'descrizione': descrizione,
            'importo': importo,
            'incidenza': incidenza,
            'periodo_indice': periodo_indice,
            'valore_indice': valore_indice,
        })

    return {
        'righe': righe,
        'importo_totale': importo_totale,
        'importo_non_classificato': importo_non_classificato,
        'indice_sintetico': indice_sintetico,
    }


# ---------------------------------------------------------------------------
# Foglio di riepilogo
# ---------------------------------------------------------------------------

_INTESTAZIONI = [
    'TOL', 'Descrizione', 'Importo Contabilita\'', 'Incidenza %',
    'Periodo indice', 'Indice ISTAT',
]


def _assicura_foglio_riepilogo(oDoc):
    oSheets = oDoc.getSheets()
    if not oSheets.hasByName(SHEET_RIEPILOGO):
        oSheets.insertNewByName(SHEET_RIEPILOGO, oSheets.getCount())
    return oSheets.getByName(SHEET_RIEPILOGO)


def scrivi_foglio_riepilogo_tol(oDoc, periodo=None):
    """
    Ricalcola e riscrive per intero il foglio 'Riepilogo TOL' a partire
    dai dati correnti di 'Elenco Prezzi'.
    Ritorna il dizionario prodotto da calcola_riepilogo().
    """
    riepilogo = calcola_riepilogo(oDoc, periodo)
    oSheet = _assicura_foglio_riepilogo(oDoc)

    for col, titolo in enumerate(_INTESTAZIONI):
        oSheet.getCellByPosition(col, 0).setString(titolo)

    riga_corrente = 1
    for r in riepilogo['righe']:
        oSheet.getCellByPosition(0, riga_corrente).setValue(r['numero'])
        oSheet.getCellByPosition(1, riga_corrente).setString(r['descrizione'])
        oSheet.getCellByPosition(2, riga_corrente).setValue(r['importo'])
        oSheet.getCellByPosition(3, riga_corrente).setValue(r['incidenza'])
        if r['periodo_indice']:
            oSheet.getCellByPosition(4, riga_corrente).setString(r['periodo_indice'])
        if r['valore_indice'] is not None:
            oSheet.getCellByPosition(5, riga_corrente).setValue(r['valore_indice'])
        riga_corrente += 1

    riga_corrente += 1  # riga vuota di separazione
    oSheet.getCellByPosition(0, riga_corrente).setString('Importo totale classificato')
    oSheet.getCellByPosition(2, riga_corrente).setValue(riepilogo['importo_totale'])
    riga_corrente += 1
    oSheet.getCellByPosition(0, riga_corrente).setString('Importo non classificato (TOL mancante)')
    oSheet.getCellByPosition(2, riga_corrente).setValue(riepilogo['importo_non_classificato'])
    riga_corrente += 1
    oSheet.getCellByPosition(0, riga_corrente).setString('Indice sintetico')
    oSheet.getCellByPosition(2, riga_corrente).setValue(riepilogo['indice_sintetico'])

    return riepilogo


# ---------------------------------------------------------------------------
# Punto di ingresso da menu/toolbar
# ---------------------------------------------------------------------------

def MENU_leeno_aggiorna_riepilogo_tol():
    """
    Ricalcola il riepilogo TOL sul documento corrente e lo scrive nel
    foglio dedicato 'Riepilogo TOL' (creato automaticamente se assente).
    Il periodo usato per gli indici ISTAT e' sempre l'ultimo disponibile;
    per un periodo specifico usare scrivi_foglio_riepilogo_tol(oDoc, periodo)
    direttamente.
    """
    oDoc = LeenoUtils.resolve_document()
    if oDoc is None:
        Dialogs.messageBox(
            text="Documento non disponibile.",
            title="Riepilogo TOL",
            msg_type=Dialogs.ERRORBOX,
        )
        return

    try:
        riepilogo = scrivi_foglio_riepilogo_tol(oDoc)
    except RuntimeError as e:
        Dialogs.messageBox(
            text=str(e),
            title="Riepilogo TOL - errore",
            msg_type=Dialogs.ERRORBOX,
        )
        return

    if riepilogo['importo_non_classificato'] != 0:
        Dialogs.messageBox(
            text="Attenzione: %.2f euro di importo in Contabilita' non sono "
                 "associati a nessuna TOL (colonna AA vuota o non valida)."
                 % riepilogo['importo_non_classificato'],
            title="Riepilogo TOL",
            msg_type=Dialogs.WARNINGBOX,
        )
