
def circoscriveVoceComputo(oSheet, lrow):
    '''
    lrow    { int }  : riga di riferimento per
                        la selezione dell'intera voce

    Circoscrive una voce di COMPUTO, VARIANTE o CONTABILITÀ
    partendo dalla posizione corrente del cursore
    '''
    #  lrow = LeggiPosizioneCorrente()[1]
    #  if oSheet.Name in('VARIANTE', 'COMPUTO','CONTABILITA'):
    if oSheet.getCellByPosition(
            0,
            lrow).CellStyle in ('comp progress', 'comp 10 s',
                                'Comp Start Attributo', 'Comp End Attributo',
                                'Comp Start Attributo_R', 'comp 10 s_R',
                                'Comp End Attributo_R', 'Livello-0-scritta',
                                'Livello-1-scritta', 'livello2 valuta'):
        y = lrow
        while oSheet.getCellByPosition(
                0, y).CellStyle not in ('Comp End Attributo',
                                        'Comp End Attributo_R'):
            y += 1
        lrowE = y
        y = lrow
        while oSheet.getCellByPosition(
                0, y).CellStyle not in ('Comp Start Attributo',
                                        'Comp Start Attributo_R'):
            y -= 1
        lrowS = y
    celle = oSheet.getCellRangeByPosition(0, lrowS, 250, lrowE)
    return celle

