import LeenoUtils

def MENU_XML_toscana_import():
    '''
    Importazione di un prezzario XML della regione Toscana
    in tabella Elenco Prezzi del template COMPUTO.
    '''
    oDoc = LeenoUtils.getDocument()

    DLG.MsgBox('Questa operazione potrebbe richiedere del tempo.', 'Avviso')
    PL.New_file.computo(0)

    try:
        filename = Dialogs.FileSelect('Scegli il file XML Toscana da importare', '*.xml')
        oDialogo_attesa = DLG.dlg_attesa()

        # mostra il dialogo
        DLG.attesa().start()
        if filename is None:
            return
    except Exception:
        ImportErrorDlg("Errore di importazione")
        return

    if not oDoc.getSheets().hasByName('COMPUTO'):
        if (len(oDoc.getURL()) == 0 and
                SheetUtils.getLastUsedColumn(oDoc.CurrentController.ActiveSheet) == 0 and
                SheetUtils.getLastUsedRow(oDoc.CurrentController.ActiveSheet) == 0):
            oDoc.close(True)

    # effettua il parsing del file XML
    tree = ElementTree()

    try:
        tree.parse(filename)
    except Exception:
        PL.ns_ins(filename)
        tree.parse(filename)
    # ~except Exception as e:
        # ~MsgBox ("Eccezione " + str(type(e)) +
        # ~"\nMessaggio: " + str(e.args) + '\n' +
        # ~traceback.format_exc());
        # ~return

    root = tree.getroot()
    iterator = tree.getiterator()

    PRT = '{' + str(iterator[0].getchildren()[0]).split('}')[0].split('{')[-1] + '}'  # xmlns
    # nome del prezzario
    intestazione = root.find(PRT + 'intestazione')
    titolo = ('Prezzario ' +
              intestazione.get('autore') +
              ' - ' +
              intestazione[0].get('area') +
              ' ' +
              intestazione[0].get('anno'))

    licenza = (intestazione[1].get('descrizione').split(':')[0] +
               ' ' +
               intestazione[1].get('tipo'))

    titolo = (titolo +
              '\nCopyright: ' +
              licenza +
              '\n\nhttp://prezzariollpp.regione.toscana.it')

    # Contenuto = root.find(PRT+'Contenuto')

    voci = root.getchildren()[1]

    tipo_lista = list()
    cap_lista = list()
    lista_articoli = list()
    lista_cap = list()
    lista_subcap = list()
    for el in voci:
        if el.tag == PRT + 'Articolo':
            codice = el.get('codice')
            codicesp = codice.split('.')

        voce = el.getchildren()[2].text
        articolo = el.getchildren()[3].text

        if articolo is None:
            desc_voce = voce
        else:
            desc_voce = voce + ' ' + articolo
        udm = el.getchildren()[4].text

        try:
            sic = float(el.getchildren()[-1][-4].get('valore'))
        except IndexError:
            sic = ''

        try:
            prezzo = float(el.getchildren()[5].text)
        except Exception:
            prezzo = float(el.getchildren()[5].text.split('.')[0] +
                           el.getchildren()[5].text.split('.')[1] +
                           '.' +
                           el.getchildren()[5].text.split('.')[2])

        try:
            mdo = float(el.getchildren()[-1][-1].get('percentuale')) / 100
            mdoE = mdo * prezzo
        except IndexError:
            mdo = ''
            mdoE = ''

        if codicesp[0] not in tipo_lista:
            tipo_lista.append(codicesp[0])
            cap = (codicesp[0], el.getchildren()[0].text, '', '', '', '', '')
            lista_cap.append(cap)
        if codicesp[0] + '.' + codicesp[1] not in cap_lista:
            cap_lista.append(codicesp[0] + '.' + codicesp[1])
            cap = (codicesp[0] +
                   '.' +
                   codicesp[1], el.getchildren()[1].text, '', '', '', '', '', '')

            lista_subcap.append(cap)
        voceel = (codice, desc_voce, udm, sic, prezzo, mdo, mdoE)
        lista_articoli.append(voceel)

    # compilo ##########################################################
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = titolo
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    flags = (VALUE + DATETIME + STRING + ANNOTATION +
             FORMULA + OBJECTS + EDITATTR)  # FORMATTED + HARDATTR
    oSheet.getCellRangeByName('D1:V1').clearContents(flags)
    oDoc.getSheets().getByName('COMPUTO').IsVisible = False
    oSheet.getCellByPosition(1, 0).String = titolo
    oSheet.getCellByPosition(2, 0).String = '''ATTENZIONE!
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.

N.B.: Si rimanda ad una attenta lettura delle note informative disponibili \
sul sito istituzionale ufficiale di riferimento prima di accedere al prezzario.'''

    oSheet.getCellByPosition(1, 0).CellStyle = 'EP-mezzo'
    n = 0

    for el in (lista_articoli, lista_cap, lista_subcap):
        oSheet.getRows().insertByIndex(4, len(el))
        lista_come_array = tuple(el)
        # Parametrizzo il range di celle a seconda della dimensione della lista
        # scarto_colonne = 0  # numero colonne da saltare a partire da sinistra
        # scarto_righe = 4  # numero righe da saltare a partire dall'alto
        colonne_lista = len(lista_come_array[1])  # numero di colonne necessarie per ospitare i dati
        righe_lista = len(lista_come_array)  # numero di righe necessarie per ospitare i dati
        oRange = oSheet.getCellRangeByPosition(0, 4, colonne_lista + 0 - 1, righe_lista + 4 - 1)
        oRange.setDataArray(lista_come_array)
        # ~ oSheet.getRows().removeByIndex(3, 1)
        oDoc.CurrentController.setActiveSheet(oSheet)

        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellStyle = "EP-aS"
        oSheet.getCellRangeByPosition(1, 3, 1, righe_lista + 3 - 1).CellStyle = "EP-a"
        oSheet.getCellRangeByPosition(2, 3, 7, righe_lista + 3 - 1).CellStyle = "EP-mezzo"
        oSheet.getCellRangeByPosition(5, 3, 5, righe_lista + 3 - 1).CellStyle = "EP-mezzo %"
        oSheet.getCellRangeByPosition(8, 3, 9, righe_lista + 3 - 1).CellStyle = "EP-sfondo"
        oSheet.getCellRangeByPosition(11, 3, 11, righe_lista + 3 - 1).CellStyle = 'EP-mezzo %'
        oSheet.getCellRangeByPosition(12, 3, 12, righe_lista + 3 - 1).CellStyle = 'EP statistiche_q'
        oSheet.getCellRangeByPosition(13, 3, 13, righe_lista + 3 - 1).CellStyle = 'EP statistiche'
        if n == 1:
            oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellBackColor = 16777120
        elif n == 2:
            oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellBackColor = 16777168
        n += 1

    Toolbars.Vedi()

    oDialogo_attesa.endExecute()
    PL.struttura_Elenco()
    oSheet.getCellRangeByName('F2').String = 'prezzi'
    oSheet.getCellRangeByName('E2').Formula = ('=COUNT(E3:E' + str(SheetUtils.getLastUsedRow(oSheet) + 1) +
                                               ')')
    dest = filename[0:-4] + '.ods'
    PL.salva_come(dest)
    DLG.MsgBox('''
Importazione eseguita con successo!

ATTENZIONE:
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.

N.B.: Si rimanda ad una attenta lettura delle note informative disponibili sul sito istituzionale ufficiale prima di accedere al Prezzario.

    ''', 'ATTENZIONE!')
########################################################################
