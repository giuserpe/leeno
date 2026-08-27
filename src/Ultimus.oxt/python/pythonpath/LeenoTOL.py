"""
    LeenO - Aggregazione per Tipologie Omogenee di Lavorazioni (TOL)

    Classifica le voci del foglio 'Elenco Prezzi' per TOL (colonna J),
    aggrega gli importi di Contabilita' (colonna V, base del SAL) per
    ciascuna TOL e produce un foglio di riepilogo con incidenze percentuali
    e indice sintetico, ai sensi dell'art. 60 D.Lgs. 36/2023 e Allegato II.2-bis.

    Colonne di Elenco Prezzi coinvolte:
        J   - numero TOL (1-20) assegnato alla voce (validazione dati a discesa)
        V   - Importi Contabilita' (base per il calcolo del SAL)

    Il foglio di riepilogo (SHEET_RIEPILOGO) viene creato se assente e
    ricalcolato ad ogni chiamata di MENU_leeno_aggiorna_riepilogo_tol().
"""

import datetime

import LeenoUtils
import LeenoISTAT
import LeenoContab
import SheetUtils
import LeenoDialogs as DLG
import Dialogs


SHEET_ELENCO_PREZZI = 'Elenco Prezzi'
SHEET_S2 = 'S2'
SHEET_RIEPILOGO = 'Riepilogo TOL'

COL_TOL = 9            # J  (0-based: A=0 ... J=9)
COL_IMPORTO_CONTAB = 21  # V  (0-based: V e' la 22a lettera -> indice 21)

RIGA_DATI_INIZIO = 3  # 0-based: intestazioni a riga 3 (1-based, indice 2), dati da riga 4 (indice 3) - confermato dal maintainer

# Cella con l'importo dell'ultimo SAL calcolato (confermato dal maintainer:
# equivalente al valore in colonna F sulla riga "T O T A L E €" del foglio SAL)
CELLA_SAL = (COL_IMPORTO_CONTAB, 1)  # V2, 0-based (colonna V, riga 2 -> indice 1)

# Etichette possibili del campo "tempo zero" in S2, provate in ordine
ETICHETTE_AGGIUDICAZIONE = (
    'Data di aggiudicazione', 'Data aggiudicazione', 'Aggiudicazione',
)

SOGLIA_ALEA = 0.03
QUOTA_COMPENSAZIONE = 0.90


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
    per ciascun numero TOL trovato in colonna J.

    Ritorna un dizionario {numero_tol (int): importo_totale (float)}.
    Le righe con J vuota, non numerica, o fuori intervallo 1-20 vengono
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
        if testo_tol:
            # Il numero TOL puo' arrivare come intero puro ("4") o come
            # numero formattato con decimali ("4,0" / "4.0") a seconda
            # del formato cella scelto per la colonna J: si normalizza la
            # virgola decimale e si accetta solo se il valore e' un intero
            # esatto nell'intervallo 1-20, per non classificare per errore
            # un valore tipo "4,3" come TOL 4.
            normalizzato = testo_tol.replace(',', '.')
            try:
                valore = float(normalizzato)
            except ValueError:
                valore = None
            if valore is not None:
                n = int(round(valore))
                if abs(valore - n) < 1e-9 and n in LeenoISTAT.TOL_CODICI:
                    numero_tol = n

        totali[numero_tol] = totali.get(numero_tol, 0.0) + importo

    return totali


def leggi_mese_aggiudicazione(oDoc):
    """
    Legge la data di aggiudicazione dal foglio S2 (prova le etichette in
    ETICHETTE_AGGIUDICAZIONE in ordine) e la converte nel periodo 'AAAA-MM'
    da usare come I0 per tutte le TOL.

    Solleva RuntimeError se il campo non e' presente o non e' una data
    nel formato atteso (%d/%m/%Y, la stessa convenzione usata altrove nel
    foglio SAL).
    """
    if not oDoc.getSheets().hasByName(SHEET_S2):
        raise RuntimeError("Foglio '%s' non trovato." % SHEET_S2)

    oS2 = oDoc.getSheets().getByName(SHEET_S2)

    grezzo = ''
    for etichetta in ETICHETTE_AGGIUDICAZIONE:
        valore = LeenoContab._leggi_dato_anagrafico(oS2, etichetta)
        # scarta i segnaposto del template ('\u25ba\u25ba\u25ba', '--', '...')
        # per un campo non ancora compilato: una data vera contiene sempre
        # almeno una cifra, un segnaposto no.
        if valore and any(c.isdigit() for c in valore):
            grezzo = valore
            break

    if not grezzo:
        raise RuntimeError(
            "Data di aggiudicazione assente o non compilata in S2 "
            "(etichette provate: %s)" % ', '.join(ETICHETTE_AGGIUDICAZIONE)
        )

    try:
        data = datetime.datetime.strptime(grezzo.strip(), '%d/%m/%Y')
    except ValueError as e:
        raise RuntimeError(
            "Data di aggiudicazione '%s' non nel formato atteso (GG/MM/AAAA)" % grezzo
        ) from e

    return '%04d-%02d' % (data.year, data.month)


def calcola_riepilogo(oDoc, periodo_sal=None, mese_aggiudicazione=None):
    """
    Calcola il riepilogo TOL con il metodo dell'indice sintetico ponderato
    (Tabella B, Allegato II.2-bis D.Lgs. 36/2023):

      peso_TOL      = importo_TOL_cumulato / importo_totale_cumulato
      I0_TOL        = indice ISTAT della TOL al mese di aggiudicazione
      It_TOL        = indice ISTAT della TOL al mese del SAL (periodo_sal)
      indice_ribasato_TOL = It_TOL / I0_TOL * 100
      indice_sintetico    = somma_TOL( peso_TOL * indice_ribasato_TOL )
      coeff_revisione     = (indice_sintetico - 100) / 100
      eccedenza           = max(coeff_revisione - 3%, 0)
      Krev                = eccedenza * 90%
      SAL_revisionale     = importo_SAL * Krev

    mese_aggiudicazione: se None, viene letto da S2 (vedi leggi_mese_aggiudicazione).
    periodo_sal: 'AAAA-MM'; se None, ultimo periodo ISTAT disponibile.

    Ritorna un dizionario con 'righe' (una per TOL con importo <> 0) e i
    totali/coefficienti aggregati.
    """
    if mese_aggiudicazione is None:
        mese_aggiudicazione = leggi_mese_aggiudicazione(oDoc)

    totali = aggrega_importi_per_tol(oDoc)
    importo_non_classificato = totali.pop(None)
    importo_totale = sum(totali.values())

    righe = []
    indice_sintetico = 0.0

    for numero in sorted(totali):
        importo = totali[numero]
        codice_sdmx, descrizione = LeenoISTAT.TOL_CODICI[numero]
        peso = (importo / importo_totale) if importo_totale else 0.0

        riga = {
            'numero': numero,
            'codice_sdmx': codice_sdmx,
            'descrizione': descrizione,
            'importo': importo,
            'peso': peso,
            'periodo_i0': None,
            'indice_i0': None,
            'periodo_it': None,
            'indice_it': None,
            'indice_ribasato': None,
            'contributo_ponderato': None,
        }

        if importo != 0:
            try:
                periodo_i0, indice_i0 = LeenoISTAT.get_indice_tol(numero, mese_aggiudicazione)
                periodo_it, indice_it = LeenoISTAT.get_indice_tol(numero, periodo_sal)
                indice_ribasato = (indice_it / indice_i0) * 100.0
                contributo = peso * indice_ribasato

                riga.update({
                    'periodo_i0': periodo_i0,
                    'indice_i0': indice_i0,
                    'periodo_it': periodo_it,
                    'indice_it': indice_it,
                    'indice_ribasato': indice_ribasato,
                    'contributo_ponderato': contributo,
                })
                indice_sintetico += contributo
            except (ValueError, RuntimeError) as e:
                # non interrompe il riepilogo per una singola TOL non raggiungibile:
                # la riga resta con indici vuoti, visibile nel foglio come anomalia
                DLG.chi("Indice TOL %d non disponibile: %s" % (numero, e))

        righe.append(riga)

    coeff_revisione = (indice_sintetico - 100.0) / 100.0
    eccedenza = max(coeff_revisione - SOGLIA_ALEA, 0.0)
    krev = eccedenza * QUOTA_COMPENSAZIONE

    return {
        'righe': righe,
        'importo_totale': importo_totale,
        'importo_non_classificato': importo_non_classificato,
        'mese_aggiudicazione': mese_aggiudicazione,
        'indice_sintetico': indice_sintetico,
        'coeff_revisione': coeff_revisione,
        'eccedenza': eccedenza,
        'krev': krev,
    }


def leggi_importo_sal(oDoc):
    """
    Legge l'importo dell'ultimo SAL calcolato dalla cella V2 di
    'Elenco Prezzi' (confermato equivalente al valore in colonna F sulla
    riga "T O T A L E €" del foglio SAL).
    """
    if not oDoc.getSheets().hasByName(SHEET_ELENCO_PREZZI):
        raise RuntimeError("Foglio '%s' non trovato." % SHEET_ELENCO_PREZZI)
    oSheet = oDoc.getSheets().getByName(SHEET_ELENCO_PREZZI)
    return oSheet.getCellByPosition(*CELLA_SAL).getValue()


def calcola_sal_revisionale(oDoc, periodo_sal=None, mese_aggiudicazione=None):
    """
    Combina calcola_riepilogo() con l'importo del SAL (cella V2 di
    Elenco Prezzi) per ottenere l'importo della revisione prezzi da
    liquidare con questo SAL.

    Aggiunge al dizionario di calcola_riepilogo() le chiavi
    'importo_sal' e 'sal_revisionale'.
    """
    riepilogo = calcola_riepilogo(oDoc, periodo_sal, mese_aggiudicazione)
    importo_sal = leggi_importo_sal(oDoc)
    riepilogo['importo_sal'] = importo_sal
    riepilogo['sal_revisionale'] = importo_sal * riepilogo['krev']
    return riepilogo


# ---------------------------------------------------------------------------
# Foglio di riepilogo
# ---------------------------------------------------------------------------

_INTESTAZIONI = [
    'TOL', 'Descrizione', 'Importo Contabilita\'', 'Peso %',
    'Periodo I0', 'Indice I0', 'Periodo It', 'Indice It',
    'Indice ribasato', 'Contributo ponderato',
]

# Ruolo di stile da applicare per ciascuna colonna del blocco righe TOL
# (indice colonna 0-based -> ruolo in _preleva_stili_riferimento)
_STILE_PER_COLONNA = {
    2: 'valuta',       # Importo Contabilita'
    3: 'percentuale',  # Peso %
    4: 'testo',        # Periodo I0
    5: 'numero',       # Indice I0
    6: 'testo',        # Periodo It
    7: 'numero',       # Indice It
    8: 'numero',       # Indice ribasato
    9: 'numero',       # Contributo ponderato
}


def _preleva_stili_riferimento(oSheetEP):
    """
    Legge a runtime i nomi di stile gia' in uso in 'Elenco Prezzi', per
    applicarli al foglio 'Riepilogo TOL'. Non vengono MAI hardcoded i nomi
    di stile (spesso anonimi, es. 'ce27') perche' non sono stabili tra
    installazioni/versioni del template: si copiano dal documento aperto,
    dove gli stili di cella sono comunque condivisi tra tutti i fogli
    dello stesso documento.

    Ritorna un dizionario {ruolo: nome_stile}; un ruolo mancante (colonna
    di riferimento non leggibile) resta assente dal dizionario invece di
    interrompere la scrittura del foglio.
    """
    riga_intestazione = max(RIGA_DATI_INIZIO - 1, 0)
    riga_dati = RIGA_DATI_INIZIO

    sorgenti = {
        'intestazione': (0, riga_intestazione),                 # A, riga header
        'testo': (1, riga_dati),                                # B, Descrizione
        'valuta': (COL_IMPORTO_CONTAB, riga_dati),              # V, Importi Contabilita'
        'numero': (4, riga_dati),                               # E, Prezzo unitario
        'percentuale': (5, riga_dati),                          # F, Incidenza MdO
    }

    stili = {}
    for ruolo, (col, riga) in sorgenti.items():
        try:
            stili[ruolo] = oSheetEP.getCellByPosition(col, riga).CellStyle
        except Exception as e:
            DLG.chi("Stile di riferimento '%s' non leggibile: %s" % (ruolo, e))
    return stili


def _applica_stile(cella, stili, ruolo):
    nome = stili.get(ruolo)
    if not nome:
        return
    try:
        cella.CellStyle = nome
    except Exception as e:
        DLG.chi("Impossibile applicare lo stile '%s' (%s): %s" % (ruolo, nome, e))


def _assicura_foglio_riepilogo(oDoc):
    oSheets = oDoc.getSheets()
    if not oSheets.hasByName(SHEET_RIEPILOGO):
        oSheets.insertNewByName(SHEET_RIEPILOGO, oSheets.getCount())
    return oSheets.getByName(SHEET_RIEPILOGO)


def scrivi_foglio_riepilogo_tol(oDoc, periodo_sal=None, mese_aggiudicazione=None):
    """
    Ricalcola e riscrive per intero il foglio 'Riepilogo TOL' a partire
    dai dati correnti di 'Elenco Prezzi' e dal SAL corrente (cella V2).
    Ritorna il dizionario prodotto da calcola_sal_revisionale().
    """
    riepilogo = calcola_sal_revisionale(oDoc, periodo_sal, mese_aggiudicazione)
    oSheet = _assicura_foglio_riepilogo(oDoc)

    oSheetEP = oDoc.getSheets().getByName(SHEET_ELENCO_PREZZI)
    stili = _preleva_stili_riferimento(oSheetEP)

    for col, titolo in enumerate(_INTESTAZIONI):
        cella = oSheet.getCellByPosition(col, 0)
        cella.setString(titolo)
        _applica_stile(cella, stili, 'intestazione')

    riga_corrente = 1
    for r in riepilogo['righe']:
        cella_numero = oSheet.getCellByPosition(0, riga_corrente)
        cella_numero.setValue(r['numero'])
        _applica_stile(cella_numero, stili, 'numero')

        cella_desc = oSheet.getCellByPosition(1, riga_corrente)
        cella_desc.setString(r['descrizione'])
        _applica_stile(cella_desc, stili, 'testo')

        cella_importo = oSheet.getCellByPosition(2, riga_corrente)
        cella_importo.setValue(r['importo'])
        _applica_stile(cella_importo, stili, 'valuta')

        cella_peso = oSheet.getCellByPosition(3, riga_corrente)
        cella_peso.setValue(r['peso'])
        _applica_stile(cella_peso, stili, 'percentuale')

        if r['periodo_i0']:
            oSheet.getCellByPosition(4, riga_corrente).setString(r['periodo_i0'])
        if r['indice_i0'] is not None:
            oSheet.getCellByPosition(5, riga_corrente).setValue(r['indice_i0'])
        if r['periodo_it']:
            oSheet.getCellByPosition(6, riga_corrente).setString(r['periodo_it'])
        if r['indice_it'] is not None:
            oSheet.getCellByPosition(7, riga_corrente).setValue(r['indice_it'])
        if r['indice_ribasato'] is not None:
            oSheet.getCellByPosition(8, riga_corrente).setValue(r['indice_ribasato'])
        if r['contributo_ponderato'] is not None:
            oSheet.getCellByPosition(9, riga_corrente).setValue(r['contributo_ponderato'])

        for col, ruolo in _STILE_PER_COLONNA.items():
            if col in (2, 3):
                continue  # gia' applicati sopra (importo, peso)
            _applica_stile(oSheet.getCellByPosition(col, riga_corrente), stili, ruolo)

        riga_corrente += 1

    riga_corrente += 1  # riga vuota di separazione
    etichette_totali = [
        ('Importo totale classificato', riepilogo['importo_totale'], 'valuta'),
        ("Importo non classificato (TOL mancante)", riepilogo['importo_non_classificato'], 'valuta'),
        ('Mese di aggiudicazione (I0)', None, 'testo'),
        ('Indice sintetico', riepilogo['indice_sintetico'], 'numero'),
        ('Coefficiente di revisione', riepilogo['coeff_revisione'], 'percentuale'),
        ("Eccedenza oltre soglia 3%", riepilogo['eccedenza'], 'percentuale'),
        ('Krev (90% eccedenza)', riepilogo['krev'], 'percentuale'),
        ('Importo SAL', riepilogo['importo_sal'], 'valuta'),
        ('SAL REVISIONALE', riepilogo['sal_revisionale'], 'valuta'),
    ]
    for etichetta, valore, ruolo_valore in etichette_totali:
        cella_etichetta = oSheet.getCellByPosition(0, riga_corrente)
        cella_etichetta.setString(etichetta)
        _applica_stile(cella_etichetta, stili, 'testo')

        cella_valore = oSheet.getCellByPosition(2, riga_corrente)
        if etichetta.startswith('Mese di aggiudicazione'):
            cella_valore.setString(riepilogo['mese_aggiudicazione'])
        else:
            cella_valore.setValue(valore)
        _applica_stile(cella_valore, stili, ruolo_valore)

        riga_corrente += 1

    return riepilogo


# ---------------------------------------------------------------------------
# Punto di ingresso da menu/toolbar
# ---------------------------------------------------------------------------

def MENU_leeno_aggiorna_riepilogo_tol():
    """
    Ricalcola il riepilogo TOL e il SAL revisionale sul documento corrente
    e li scrive nel foglio dedicato 'Riepilogo TOL' (creato automaticamente
    se assente).

    Il mese di aggiudicazione (I0) viene letto da S2; il periodo del SAL
    (It) e' sempre l'ultimo indice ISTAT disponibile. Per un periodo
    specifico usare scrivi_foglio_riepilogo_tol(oDoc, periodo_sal, ...)
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
                 "associati a nessuna TOL (colonna J vuota o non valida)."
                 % riepilogo['importo_non_classificato'],
            title="Riepilogo TOL",
            msg_type=Dialogs.WARNINGBOX,
        )
