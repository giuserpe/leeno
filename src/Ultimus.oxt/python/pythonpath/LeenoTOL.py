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

# Stile di sfondo di Elenco Prezzi da applicare per intero alle colonne
# A, E, G (TOL, Periodo I0, Periodo It) - nome dato esplicitamente dal
# maintainer, non derivato a runtime come gli altri stili.
_STILE_SFONDO_EP = 'EP-sfondo'
_COLONNE_SFONDO_EP = (0, 4, 6)  # A, E, G

# Larghezze colonna (1/100 mm, stessa unita' di TableColumns.Width)
_LARGHEZZE_COLONNE = {
    0: 1150,   # A
    1: 13660,  # B
    2: 1815,   # C
    3: 1432,   # D
    4: 1432,   # E
    5: 1432,   # F
    6: 1432,   # G
    7: 1432,   # H
    8: 1432,   # I
    9: 1432,   # J
}


def _applica_sfondo_ep(oSheet, riga_tol_inizio_1based, riga_tol_fine_1based):
    """
    Applica lo stile 'EP-sfondo' alle colonne A, E, G limitatamente alle
    righe TOL (quelle con dati reali in quelle colonne) - non oltre,
    non un buffer arbitrario: il range e' legato esplicitamente a
    RIGA_TOL_INIZIO/RIGA_TOL_FINE, che riflettono il numero vero di TOL
    codificate (LeenoISTAT.TOL_CODICI), non un valore fisso che smetterebbe
    di essere corretto se quel numero cambiasse.
    """
    for col in _COLONNE_SFONDO_EP:
        try:
            oRange = oSheet.getCellRangeByPosition(
                col, riga_tol_inizio_1based - 1, col, riga_tol_fine_1based - 1
            )
            oRange.CellStyle = _STILE_SFONDO_EP
        except Exception as e:
            DLG.chi(
                "Impossibile applicare lo stile '%s' alla colonna %d: %s"
                % (_STILE_SFONDO_EP, col, e)
            )


def _applica_larghezze_colonne(oSheet):
    for col, larghezza in _LARGHEZZE_COLONNE.items():
        try:
            oSheet.getColumns().getByIndex(col).Width = larghezza
        except Exception as e:
            DLG.chi("Impossibile impostare la larghezza della colonna %d: %s" % (col, e))


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

    A differenza della versione precedente, i passaggi di puro calcolo
    interno al foglio (peso, indice ribasato, contributo ponderato,
    indice sintetico, coefficiente di revisione, eccedenza, Krev, SAL
    revisionale) sono scritti come FORMULE Calc, non come valori statici:
    aprendo una cella se ne vede la provenienza esatta, e modificando a
    mano un I0/It il foglio si ricalcola da solo.

    Restano valori (non formule), perche' provengono da fonti esterne al
    foglio che una formula Calc non puo' calcolare da sola:
      - Importo per TOL (colonna C): aggregazione robusta da Python
        (gestisce segnaposto/decimali in colonna J e la riga sentinella
        'Fine elenco' - vedi aggrega_importi_per_tol)
      - Periodo I0/It e Indice I0/It (colonne E,F,G,H): da LeenoISTAT
      - Importo non classificato e Mese di aggiudicazione: da Python

    Soglia alea (3%) e Quota compensazione (90%) sono scritte come celle
    modificabili: se il foglio esiste gia' e contengono un valore diverso
    da zero, quel valore viene preservato invece di essere sovrascritto
    dal default, cosi' un'eventuale modifica manuale (es. per la Tabella C
    alternativa) sopravvive ai ricalcoli successivi.

    Ritorna il dizionario prodotto da calcola_sal_revisionale() (calcolato
    comunque in Python, usato per il messaggio di avviso e come riscontro
    indipendente rispetto a quanto il foglio ricalcola da solo).
    """
    riepilogo = calcola_sal_revisionale(oDoc, periodo_sal, mese_aggiudicazione)
    oSheet = _assicura_foglio_riepilogo(oDoc)

    oSheetEP = oDoc.getSheets().getByName(SHEET_ELENCO_PREZZI)
    stili = _preleva_stili_riferimento(oSheetEP)

    # Righe (1-based, come compaiono nell'interfaccia di Calc)
    RIGA_HEADER = 1
    RIGA_TOL_INIZIO = 2
    RIGA_TOL_FINE = RIGA_TOL_INIZIO + len(riepilogo['righe']) - 1
    RIGA_TOT_CLASSIFICATO = RIGA_TOL_FINE + 2
    RIGA_NON_CLASSIFICATO = RIGA_TOT_CLASSIFICATO + 1
    RIGA_MESE_AGG = RIGA_NON_CLASSIFICATO + 1
    RIGA_SOGLIA = RIGA_MESE_AGG + 1
    RIGA_QUOTA = RIGA_SOGLIA + 1
    RIGA_INDICE_SINT = RIGA_QUOTA + 1
    RIGA_COEFF = RIGA_INDICE_SINT + 1
    RIGA_ECCEDENZA = RIGA_COEFF + 1
    RIGA_KREV = RIGA_ECCEDENZA + 1
    RIGA_IMPORTO_SAL = RIGA_KREV + 1
    RIGA_SAL_REV = RIGA_IMPORTO_SAL + 1

    for col, titolo in enumerate(_INTESTAZIONI):
        cella = oSheet.getCellByPosition(col, RIGA_HEADER - 1)
        cella.setString(titolo)
        _applica_stile(cella, stili, 'intestazione')

    riga_corrente = RIGA_TOL_INIZIO
    for r in riepilogo['righe']:
        r_ = riga_corrente  # alias per leggibilita' nelle formule sotto

        cella_numero = oSheet.getCellByPosition(0, r_ - 1)
        cella_numero.setValue(r['numero'])
        _applica_stile(cella_numero, stili, 'numero')

        cella_desc = oSheet.getCellByPosition(1, r_ - 1)
        cella_desc.setString(r['descrizione'])
        _applica_stile(cella_desc, stili, 'testo')

        cella_importo = oSheet.getCellByPosition(2, r_ - 1)
        cella_importo.setValue(r['importo'])
        _applica_stile(cella_importo, stili, 'valuta')

        cella_peso = oSheet.getCellByPosition(3, r_ - 1)
        cella_peso.setFormula(
            '=IF($C$%d=0;0;C%d/$C$%d)' % (RIGA_TOT_CLASSIFICATO, r_, RIGA_TOT_CLASSIFICATO)
        )
        _applica_stile(cella_peso, stili, 'percentuale')

        if r['periodo_i0']:
            oSheet.getCellByPosition(4, r_ - 1).setString(r['periodo_i0'])
        if r['indice_i0'] is not None:
            oSheet.getCellByPosition(5, r_ - 1).setValue(r['indice_i0'])
        if r['periodo_it']:
            oSheet.getCellByPosition(6, r_ - 1).setString(r['periodo_it'])
        if r['indice_it'] is not None:
            oSheet.getCellByPosition(7, r_ - 1).setValue(r['indice_it'])

        cella_ribasato = oSheet.getCellByPosition(8, r_ - 1)
        cella_ribasato.setFormula('=IF(OR(F%d="";F%d=0);"";H%d/F%d*100)' % (r_, r_, r_, r_))

        cella_contributo = oSheet.getCellByPosition(9, r_ - 1)
        cella_contributo.setFormula('=IF(I%d="";"";D%d*I%d)' % (r_, r_, r_))

        for col, ruolo in _STILE_PER_COLONNA.items():
            if col in (2, 3):
                continue  # gia' applicati sopra (importo, peso)
            _applica_stile(oSheet.getCellByPosition(col, r_ - 1), stili, ruolo)

        riga_corrente += 1

    def _scrivi_totale(riga, etichetta, ruolo_valore):
        # colonna A svuotata esplicitamente: prima di questa modifica
        # l'etichetta veniva scritta li'; un ricalcolo su un foglio gia'
        # popolato dalla versione precedente lascerebbe altrimenti il
        # testo vecchio in A oltre a quello nuovo in B.
        oSheet.getCellByPosition(0, riga - 1).setString('')

        cella_etichetta = oSheet.getCellByPosition(1, riga - 1)
        cella_etichetta.setString(etichetta)
        _applica_stile(cella_etichetta, stili, 'testo')
        cella_valore = oSheet.getCellByPosition(2, riga - 1)
        _applica_stile(cella_valore, stili, ruolo_valore)
        return cella_valore

    _scrivi_totale(RIGA_TOT_CLASSIFICATO, 'Importo totale classificato', 'valuta') \
        .setFormula('=SUM(C%d:C%d)' % (RIGA_TOL_INIZIO, RIGA_TOL_FINE))

    _scrivi_totale(RIGA_NON_CLASSIFICATO, "Importo non classificato (TOL mancante)", 'valuta') \
        .setValue(riepilogo['importo_non_classificato'])

    _scrivi_totale(RIGA_MESE_AGG, 'Mese di aggiudicazione (I0)', 'testo') \
        .setString(riepilogo['mese_aggiudicazione'])

    cella_soglia = _scrivi_totale(RIGA_SOGLIA, 'Soglia alea (modificabile)', 'percentuale')
    valore_soglia_esistente = cella_soglia.getValue()
    cella_soglia.setValue(valore_soglia_esistente if valore_soglia_esistente else SOGLIA_ALEA)

    cella_quota = _scrivi_totale(RIGA_QUOTA, 'Quota compensazione (modificabile)', 'percentuale')
    valore_quota_esistente = cella_quota.getValue()
    cella_quota.setValue(valore_quota_esistente if valore_quota_esistente else QUOTA_COMPENSAZIONE)

    _scrivi_totale(RIGA_INDICE_SINT, 'Indice sintetico', 'numero') \
        .setFormula('=SUM(J%d:J%d)' % (RIGA_TOL_INIZIO, RIGA_TOL_FINE))

    _scrivi_totale(RIGA_COEFF, 'Coefficiente di revisione', 'percentuale') \
        .setFormula('=(C%d-100)/100' % RIGA_INDICE_SINT)

    _scrivi_totale(RIGA_ECCEDENZA, "Eccedenza oltre soglia", 'percentuale') \
        .setFormula('=MAX(C%d-C%d;0)' % (RIGA_COEFF, RIGA_SOGLIA))

    _scrivi_totale(RIGA_KREV, 'Krev (eccedenza * quota compensazione)', 'percentuale') \
        .setFormula('=C%d*C%d' % (RIGA_ECCEDENZA, RIGA_QUOTA))

    _scrivi_totale(RIGA_IMPORTO_SAL, 'Importo SAL', 'valuta') \
        .setFormula("='%s'.V2" % SHEET_ELENCO_PREZZI)

    _scrivi_totale(RIGA_SAL_REV, 'SAL REVISIONALE', 'valuta') \
        .setFormula('=C%d*C%d' % (RIGA_IMPORTO_SAL, RIGA_KREV))

    _applica_sfondo_ep(oSheet, RIGA_TOL_INIZIO, RIGA_TOL_FINE)
    _applica_larghezze_colonne(oSheet)

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
