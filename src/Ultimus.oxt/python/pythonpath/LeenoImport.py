"""
    LeenO - modulo di importazione prezzari
"""
import logging
import threading

from xml.etree.ElementTree import ElementTree

import uno

from com.sun.star.sheet.CellFlags import (VALUE, DATETIME, STRING,
                                          ANNOTATION, FORMULA,
                                          OBJECTS, EDITATTR)

from LeenoUtils import getDocument
import pyleeno as PL
import LeenoDialogs as DLG
import LeenoToolbars as Toolbars
from LeenoConfig import Config

import Dialogs


def ImportErrorDlg(msg):
    """ Generico dialogo di errore di importazione con messaggio
        DA FARE
    """
    print("Import error:", msg)


########################################################################
def MENU_importa_listino_leeno():
    '''
    @@ DA DOCUMENTARE
    '''
    importa_listino_leeno_th().start()


class importa_listino_leeno_th(threading.Thread):
    '''
    @@ DA DOCUMENTARE
    '''
    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
        importa_listino_leeno_run()


###
def importa_listino_leeno_run():
    '''
    Esegue la conversione di un listino (formato LeenO) in template Computo
    '''
    oDoc = getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #  giallo(16777072,16777120,16777168)
    #  verde(9502608,13696976,15794160)
    #  viola(12632319,13684991,15790335)
    lista_articoli = list()
    nome = oSheet.getCellByPosition(2, 0).String
    test = PL.uFindStringCol('ATTENZIONE!', 5, oSheet) + 1
    assembla = DLG.DlgSiNo(
        '''Il riconoscimento di descrizioni e sottodescrizioni
dipende dalla colorazione di sfondo delle righe.

Nel caso in cui questa fosse alterata, il risultato finale
della conversione potrebbe essere inatteso.

Considera anche la possibilità di recuperare il formato XML(SIX)
di questo prezzario dal sito ufficiale dell'ente che lo rilascia.

Vuoi assemblare descrizioni e sottodescrizioni?''', 'Richiesta')
    oDialogo_attesa = DLG.dlg_attesa()
    DLG.attesa().start()  # mostra il dialogo

    if assembla == 2:
        PL.colora_vecchio_elenco()
    orig = oDoc.getURL()
    dest0 = orig[0:-4] + '_new.ods'

    orig = uno.fileUrlToSystemPath(PL.LeenO_path() + '/template/leeno/Computo_LeenO.ots')
    dest = uno.fileUrlToSystemPath(dest0)

    PL.shutil.copyfile(orig, dest)
    madre = ''
    for el in range(test, PL.getLastUsedCell(oSheet).EndRow + 1):
        tariffa = oSheet.getCellByPosition(2, el).String
        descrizione = oSheet.getCellByPosition(4, el).String
        um = oSheet.getCellByPosition(6, el).String
        sic = oSheet.getCellByPosition(11, el).String
        prezzo = oSheet.getCellByPosition(7, el).String
        mdo_p = oSheet.getCellByPosition(8, el).String
        mdo = oSheet.getCellByPosition(9, el).String
        if oSheet.getCellByPosition(2,
                                    el).CellBackColor in (16777072, 16777120,
                                                          9502608, 13696976,
                                                          12632319, 13684991):
            articolo = (
                tariffa,
                descrizione,
                um,
                sic,
                prezzo,
                mdo_p,
                mdo,
            )
        elif oSheet.getCellByPosition(2,
                                      el).CellBackColor in (16777168, 15794160,
                                                            15790335):
            if assembla == 2:
                madre = descrizione
            articolo = (
                tariffa,
                descrizione,
                um,
                sic,
                prezzo,
                mdo_p,
                mdo,
            )
        else:
            if madre == '':
                descrizione = oSheet.getCellByPosition(4, el).String
            else:
                descrizione = madre + ' \n- ' + oSheet.getCellByPosition(
                    4, el).String
            articolo = (
                tariffa,
                descrizione,
                um,
                sic,
                prezzo,
                mdo_p,
                mdo,
            )
        lista_articoli.append(articolo)
    oDialogo_attesa.endExecute()
    PL._gotoDoc(dest)  # vado sul nuovo file
    # compilo la tabella ###################################################
    oDoc = getDocument()
    oDialogo_attesa = DLG.dlg_attesa()
    DLG.attesa().start()  # mostra il dialogo

    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = nome
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oSheet.getCellByPosition(1, 1).String = nome

    oSheet.getRows().insertByIndex(4, len(lista_articoli))
    lista_come_array = tuple(lista_articoli)
    # Parametrizzo il range di celle a seconda della dimensione della lista
    colonne_lista = len(lista_come_array[1]
                        )  # numero di colonne necessarie per ospitare i dati
    righe_lista = len(
        lista_come_array)  # numero di righe necessarie per ospitare i dati
    oRange = oSheet.getCellRangeByPosition(
        0,
        4,
        colonne_lista - 1,  # l'indice parte da 0
        righe_lista + 4 - 1)
    oRange.setDataArray(lista_come_array)
    oSheet.getRows().removeByIndex(3, 1)
    oDoc.CurrentController.setActiveSheet(oSheet)
    oDialogo_attesa.endExecute()
    procedo = DLG.DlgSiNo(
        '''Vuoi mettere in ordine la visualizzazione del prezzario?

Le righe senza prezzo avranno una tonalità di sfondo
diversa dalle altre e potranno essere facilmente nascoste.

Questa operazione potrebbe richiedere del tempo.''', 'Richiesta...')
    if procedo == 2:
        DLG.attesa().start()  # mostra il dialogo
        #  struttura_Elenco()
        oDialogo_attesa.endExecute()
    DLG.MsgBox('Conversione eseguita con successo!', '')
    PL.autoexec()


########################################################################
# ~class XPWE_import_th(threading.Thread):
# ~def __init__(self):
# ~threading.Thread.__init__(self)
# ~def run(self):
# ~XPWE_import_run()
def MENU_XPWE_import():
    '''
    Visualizza il menù Esporta XPWE
    '''
    XPWE_in(PL.scegli_elaborato('Importa dal formato XPWE'))


def XPWE_in(arg):
    '''
    @@ DA DOCUMENTARE
    '''
    oDoc = getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    PL.refresh(0)
    oDialogo_attesa = DLG.dlg_attesa('Caricamento dei dati...')
    if not oDoc.getSheets().hasByName('S2'):
        DLG.MsgBox(
            'Puoi usare questo comando da un file di computo nuovo o già esistente.',
            'ATTENZIONE!')
        return
    DLG.MsgBox("Il contenuto dell'archivio XPWE sarà aggiunto a questo file come " + arg + ".", 'Avviso!')
    PL._gotoSheet('COMPUTO')
    oDoc.CurrentController.select(oDoc.getSheets().hasByName(
        'COMPUTO'))  # per evitare che lo script parta da un altro documento
    filename = DLG.filedia('Scegli il file XPWE da importare...', '*.xpwe')  # *.xpwe')
    # xml auto indent: http://www.freeformatter.com/xml-formatter.html
    # inizializzazione delle variabili
    # datarif = datetime.now()
    lista_articoli = list(
    )  # lista in cui memorizzare gli articoli da importare
    diz_ep = dict()  # array per le voci di elenco prezzi
    # effettua il parsing del file XML
    tree = ElementTree()
    if filename == 'Cancel' or filename == '':
        return
    try:
        tree.parse(filename)
    except TypeError:
        return
    except PermissionError:
        DLG.MsgBox('Accertati che il nome del file sia corretto.', 'ATTENZIONE! Impossibile procedere.')
        return
    # ottieni l'item root
    root = tree.getroot()
    logging.debug(list(root))
    # effettua il parsing di tutti gli elementi dell'albero XML
    # iterator = tree.getiterator()
    # if root.find('FileNameDocumento'):
    #     nome_file = root.find('FileNameDocumento').text
    # else:
    #     nome_file = "nome_file"
    # ##
    dati = root.find('PweDatiGenerali')
    DatiGenerali = dati.getchildren()[0][0]
    # percprezzi = DatiGenerali[0].text
    comune = DatiGenerali[1].text
    # provincia = DatiGenerali[2].text
    oggetto = DatiGenerali[3].text
    committente = DatiGenerali[4].text
    impresa = DatiGenerali[5].text
    # parteopera = DatiGenerali[6].text
    # PweDGCapitoliCategorie
    try:
        CapCat = dati.find('PweDGCapitoliCategorie')
        # PweDGSuperCapitoli
        lista_supcap = list()
        if CapCat.find('PweDGSuperCapitoli'):
            PweDGSuperCapitoli = CapCat.find(
                'PweDGSuperCapitoli').getchildren()
            for elem in PweDGSuperCapitoli:
                id_sc = elem.get('ID')
                codice = elem.find('Codice').text
                try:
                    codice = elem.find('Codice').text
                except AttributeError:
                    codice = ''
                dessintetica = elem.find('DesSintetica').text
                percentuale = elem.find('Percentuale').text
                diz = dict()
                diz['id_sc'] = id_sc
                diz['codice'] = codice
                diz['dessintetica'] = dessintetica
                diz['percentuale'] = percentuale
                lista_supcap.append(diz)
        # PweDGCapitoli
        lista_cap = list()
        if CapCat.find('PweDGCapitoli'):
            PweDGCapitoli = CapCat.find('PweDGCapitoli').getchildren()
            for elem in PweDGCapitoli:
                id_sc = elem.get('ID')
                codice = elem.find('Codice').text
                try:
                    codice = elem.find('Codice').text
                except AttributeError:
                    codice = ''
                dessintetica = elem.find('DesSintetica').text
                percentuale = elem.find('Percentuale').text
                diz = dict()
                diz['id_sc'] = id_sc
                diz['codice'] = codice
                diz['dessintetica'] = dessintetica
                diz['percentuale'] = percentuale
                lista_cap.append(diz)
        # PweDGSubCapitoli
        lista_subcap = list()
        if CapCat.find('PweDGSubCapitoli'):
            PweDGSubCapitoli = CapCat.find('PweDGSubCapitoli').getchildren()
            for elem in PweDGSubCapitoli:
                id_sc = elem.get('ID')
                codice = elem.find('Codice').text
                try:
                    codice = elem.find('Codice').text
                except AttributeError:
                    codice = ''
                dessintetica = elem.find('DesSintetica').text
                percentuale = elem.find('Percentuale').text
                diz = dict()
                diz['id_sc'] = id_sc
                diz['codice'] = codice
                diz['dessintetica'] = dessintetica
                diz['percentuale'] = percentuale
                lista_subcap.append(diz)
        # PweDGSuperCategorie
        lista_supcat = list()
        if CapCat.find('PweDGSuperCategorie'):
            PweDGSuperCategorie = CapCat.find(
                'PweDGSuperCategorie').getchildren()
            for elem in PweDGSuperCategorie:
                id_sc = elem.get('ID')
                dessintetica = elem.find('DesSintetica').text
                try:
                    percentuale = elem.find('Percentuale').text
                except AttributeError:
                    percentuale = '0'
                supcat = (id_sc, dessintetica, percentuale)
                lista_supcat.append(supcat)
        # PweDGCategorie
        lista_cat = list()
        if CapCat.find('PweDGCategorie'):
            PweDGCategorie = CapCat.find('PweDGCategorie').getchildren()
            for elem in PweDGCategorie:
                id_sc = elem.get('ID')
                dessintetica = elem.find('DesSintetica').text
                try:
                    percentuale = elem.find('Percentuale').text
                except AttributeError:
                    percentuale = '0'
                cat = (id_sc, dessintetica, percentuale)
                lista_cat.append(cat)
        # PweDGSubCategorie
        lista_subcat = list()
        if CapCat.find('PweDGSubCategorie'):
            PweDGSubCategorie = CapCat.find('PweDGSubCategorie').getchildren()
            for elem in PweDGSubCategorie:
                id_sc = elem.get('ID')
                dessintetica = elem.find('DesSintetica').text
                try:
                    percentuale = elem.find('Percentuale').text
                except AttributeError:
                    percentuale = '0'
                subcat = (id_sc, dessintetica, percentuale)
                lista_subcat.append(subcat)
    except AttributeError:
        pass
    # PweDGWBS
    # try:
    #    PweDGWBS = dati.find('PweDGWBS')
    #    pass
    # except AttributeError:
    #    pass

    # PweDGAnalisi
    PweDGAnalisi = dati.find('PweDGModuli').getchildren()[0]
    # speseutili = PweDGAnalisi.find('SpeseUtili').text
    spesegenerali = PweDGAnalisi.find('SpeseGenerali').text
    utiliimpresa = PweDGAnalisi.find('UtiliImpresa').text
    oneriaccessorisc = PweDGAnalisi.find('OneriAccessoriSc').text
    # confquantita = PweDGAnalisi.find('ConfQuantita').text
    oSheet = oDoc.getSheets().getByName('S1')

    try:
        oSheet.getCellByPosition(7, 318).Value = float(oneriaccessorisc) / 100
    except Exception:
        pass
    try:
        oSheet.getCellByPosition(7, 319).Value = float(spesegenerali) / 100
    except Exception:
        pass
    try:
        oSheet.getCellByPosition(7, 320).Value = float(utiliimpresa) / 100
    except Exception:
        pass
    # imposto le approssimazioni
    try:
        PweDGConfigNumeri = dati.find('PweDGConfigurazione').getchildren()[0]
        # Divisa = PweDGConfigNumeri.find('Divisa').text
        # ConversioniIN = PweDGConfigNumeri.find('ConversioniIN').text
        # FattoreConversione = PweDGConfigNumeri.find('FattoreConversione').text
        # Cambio = PweDGConfigNumeri.find('Cambio').text
        PartiUguali = PweDGConfigNumeri.find('PartiUguali').text.split('.')[-1].split('|')[0]
        Larghezza = PweDGConfigNumeri.find('Larghezza').text.split('.')[-1].split('|')[0]
        Lunghezza = PweDGConfigNumeri.find('Lunghezza').text.split('.')[-1].split('|')[0]
        HPeso = PweDGConfigNumeri.find('HPeso').text.split('.')[-1].split('|')[0]
        Quantita = PweDGConfigNumeri.find('Quantita').text.split('.')[-1].split('|')[0]
        Prezzi = PweDGConfigNumeri.find('Prezzi').text.split('.')[-1].split('|')[0]
        PrezziTotale = PweDGConfigNumeri.find('PrezziTotale').text.split('.')[-1].split('|')[0]
        # ConvPrezzi = PweDGConfigNumeri.find('ConvPrezzi').text.split('.')[-1].split('|')[0]
        # ConvPrezziTotale = PweDGConfigNumeri.find('ConvPrezziTotale').text.split('.')[-1].split('|')[0]
        # IncidenzaPercentuale = PweDGConfigNumeri.find('IncidenzaPercentuale').text.split('.')[-1].split('|')[0]
        # Aliquote = PweDGConfigNumeri.find('Aliquote').text.split('.')[-1].split('|')[0]
        PL.dec_pl('comp 1-a PU', int(PartiUguali))
        PL.dec_pl('comp 1-a LUNG', int(Lunghezza))
        PL.dec_pl('comp 1-a LARG', int(Larghezza))
        PL.dec_pl('comp 1-a peso', int(HPeso))
        for el in ('Comp-Variante num sotto', 'An-lavoraz-input', 'Blu'):
            PL.dec_pl(el, int(Quantita))
        for el in ('comp sotto Unitario', 'An-lavoraz-generica'):
            PL.dec_pl(el, int(Prezzi))
        for el in ('comp sotto Euro Originale', 'Livello-0-scritta mini val',
                   'Livello-1-scritta mini val', 'livello2 scritta mini',
                   'Comp TOTALI', 'Ultimus_totali_1', 'Ultimus_bordo',
                   'ULTIMUS_3', 'Ultimus_Bordo_sotto',
                   'Comp-Variante num sotto', 'An-valuta-dx', 'An-1v-dx',
                   'An-lavoraz-generica', 'An-lavoraz-Utili-num sin'):
            PL.dec_pl(el, int(PrezziTotale))
    except IndexError:
        pass
    #
    misurazioni = root.find('PweMisurazioni')
    PweElencoPrezzi = misurazioni.getchildren()[0]

    # leggo l'elenco prezzi
    epitems = PweElencoPrezzi.findall('EPItem')
    dict_articoli = dict()
    lista_articoli = list()
    lista_analisi = list()
    lista_tariffe_analisi = list()
    for elem in epitems:
        id_ep = elem.get('ID')
        diz_ep = dict()
        tipoep = elem.find('TipoEP').text
        if elem.find('Tariffa').text is not None:
            tariffa = elem.find('Tariffa').text
        else:
            tariffa = ''
        articolo = elem.find('Articolo').text
        desridotta = elem.find('DesRidotta').text
        destestesa = elem.find('DesEstesa').text  # .strip()
        try:
            desridotta = elem.find('DesBreve').text
        except AttributeError:
            pass
        try:
            desbreve = elem.find('DesBreve').text
        except AttributeError:
            desbreve = ''

        if elem.find('UnMisura').text is not None:
            unmisura = elem.find('UnMisura').text
        else:
            unmisura = ''
        if elem.find('Prezzo1').text == '0' or elem.find(
                'Prezzo1').text is None:
            prezzo1 = ''
        else:
            prezzo1 = float(elem.find('Prezzo1').text)
        prezzo2 = elem.find('Prezzo2').text
        prezzo3 = elem.find('Prezzo3').text
        prezzo4 = elem.find('Prezzo4').text
        prezzo5 = elem.find('Prezzo5').text
        try:
            idspcap = elem.find('IDSpCap').text
        except AttributeError:
            idspcap = ''
        try:
            idcap = elem.find('IDCap').text
        except AttributeError:
            idcap = ''

        # try:
        #    idsbcap = elem.find('IDSbCap').text
        # except AttributeError:
        #    idsbcap = ''

        try:
            flags = elem.find('Flags').text
        except AttributeError:
            flags = ''
        try:
            data = elem.find('Data').text
        except AttributeError:
            data = ''
        IncSIC = ''
        IncMDO = ''
        IncMAT = ''
        IncATTR = ''
        try:
            if float(elem.find('IncSIC').text) != 0:
                IncSIC = float(elem.find('IncSIC').text) / 100
        except Exception:  # AttributeError TypeError:
            pass
        try:
            if float(elem.find('IncMDO').text) != 0:
                IncMDO = float(elem.find('IncMDO').text) / 100
        except Exception:  # AttributeError TypeError:
            pass
        try:
            if float(elem.find('IncMAT').text) != 0:
                IncMAT = float(elem.find('IncMAT').text) / 100
        except Exception:  # AttributeError TypeError:
            pass
        try:
            if float(elem.find('IncATTR').text) != 0:
                IncATTR = float(elem.find('IncATTR').text) / 100
        except Exception:  # AttributeError TypeError:
            pass
        try:
            adrinternet = elem.find('AdrInternet').text
        except AttributeError:
            adrinternet = ''
        if elem.find('PweEPAnalisi').text is None:
            pweepanalisi = ''
        else:
            pweepanalisi = elem.find('PweEPAnalisi').text
        #  chi(pweepanalisi)
        diz_ep['tipoep'] = tipoep
        diz_ep['tariffa'] = tariffa
        diz_ep['articolo'] = articolo
        diz_ep['desridotta'] = desridotta
        diz_ep['destestesa'] = destestesa
        diz_ep['desridotta'] = desridotta
        diz_ep['desbreve'] = desbreve
        diz_ep['unmisura'] = unmisura
        diz_ep['prezzo1'] = prezzo1
        diz_ep['prezzo2'] = prezzo2
        diz_ep['prezzo3'] = prezzo3
        diz_ep['prezzo4'] = prezzo4
        diz_ep['prezzo5'] = prezzo5
        diz_ep['idspcap'] = idspcap
        diz_ep['idcap'] = idcap
        diz_ep['flags'] = flags
        diz_ep['data'] = data
        diz_ep['adrinternet'] = adrinternet
        #  diz_ep['pweepanalisi'] = pweepanalisi
        diz_ep['IncSIC'] = IncSIC
        diz_ep['IncMDO'] = IncMDO
        diz_ep['IncMAT'] = IncMDO
        diz_ep['IncATTR'] = IncMDO

        dict_articoli[id_ep] = diz_ep
        articolo_modificato = (
            tariffa,
            destestesa,
            unmisura,
            IncSIC,
            # ~float(prezzo1),
            prezzo1,
            IncMDO,
            IncMAT,
            IncATTR)
        lista_articoli.append(articolo_modificato)
        # leggo analisi di prezzo
        pweepanalisi = elem.find('PweEPAnalisi')
        PweEPAR = pweepanalisi.find('PweEPAR')
        if PweEPAR is not None:
            EPARItem = PweEPAR.findall('EPARItem')
            analisi = list()
            for el in EPARItem:
                # id_an = el.get('ID')
                # an_tipo = el.find('Tipo').text
                id_ep = el.find('IDEP').text
                an_des = el.find('Descrizione').text
                an_um = el.find('Misura').text
                if an_um is None:
                    an_um = ''
                try:
                    an_qt = el.find('Qt').text.replace(' ', '')
                except Exception:
                    an_qt = ''
                try:
                    an_pr = el.find('Prezzo').text.replace(' ', '')
                except Exception:
                    an_pr = ''
                # an_fld = el.find('FieldCTL').text
                an_rigo = (id_ep, an_des, an_um, an_qt, an_pr)
                analisi.append(an_rigo)
            lista_analisi.append(
                [tariffa, destestesa, unmisura, analisi, prezzo1])
            lista_tariffe_analisi.append(tariffa)
    # leggo voci di misurazione e righe ####################################
    lista_misure = list()
    try:
        PweVociComputo = misurazioni.getchildren()[1]
        vcitems = PweVociComputo.findall('VCItem')
        prova_l = list()
        for elem in vcitems:
            diz_misura = dict()
            id_vc = elem.get('ID')
            id_ep = elem.find('IDEP').text
            quantita = elem.find('Quantita').text
            try:
                datamis = elem.find('DataMis').text
            except AttributeError:
                datamis = ''
            try:
                flags = elem.find('Flags').text
            except AttributeError:
                flags = ''
            try:
                idspcat = elem.find('IDSpCat').text
            except AttributeError:
                idspcat = ''
            try:
                idcat = elem.find('IDCat').text
            except AttributeError:
                idcat = ''
            try:
                idsbcat = elem.find('IDSbCat').text
            except AttributeError:
                idsbcat = ''
            # try:
            #    CodiceWBS = elem.find('CodiceWBS').text
            # except AttributeError:
            #    CodiceWBS = ''

            righi_mis = elem.getchildren()[-1].findall('RGItem')
            # lista_rig = list()
            riga_misura = ()
            lista_righe = list()  # []
            new_id_l = list()

            for el in righi_mis:
                # rgitem = el.get('ID')
                idvv = el.find('IDVV').text
                if el.find('Descrizione').text is not None:
                    descrizione = el.find('Descrizione').text
                else:
                    descrizione = ''
                partiuguali = el.find('PartiUguali').text
                lunghezza = el.find('Lunghezza').text
                larghezza = el.find('Larghezza').text
                hpeso = el.find('HPeso').text
                quantita = el.find('Quantita').text
                flags = el.find('Flags').text
                riga_misura = (
                    descrizione,
                    '',
                    '',
                    partiuguali,
                    lunghezza,
                    larghezza,
                    hpeso,
                    quantita,
                    flags,
                    idvv,
                )
                mia = []
                mia.append(riga_misura[0])
                for el in riga_misura[1:]:
                    if el is None:
                        el = ''
                    else:
                        try:
                            el = float(el)
                        except ValueError:
                            if el != '':
                                el = '=' + el.replace('.', ',')
                    mia.append(el)
                lista_righe.append(riga_misura)
            diz_misura['id_vc'] = id_vc
            diz_misura['id_ep'] = id_ep
            diz_misura['quantita'] = quantita
            diz_misura['datamis'] = datamis
            diz_misura['flags'] = flags
            diz_misura['idspcat'] = idspcat
            diz_misura['idcat'] = idcat
            diz_misura['idsbcat'] = idsbcat
            diz_misura['lista_rig'] = lista_righe

            new_id = PL.strall(idspcat) + '.' + PL.strall(idcat) + '.' + PL.strall(idsbcat)
            new_id_l = (new_id, diz_misura)
            prova_l.append(new_id_l)
            lista_misure.append(diz_misura)
    except IndexError:
        DLG.MsgBox(
            """Nel file scelto non risultano esserci voci di misurazione,
perciò saranno importate le sole voci di Elenco Prezzi.

Si tenga conto che:
- sarà importato solo il "Prezzo 1" dell'elenco;
- a seconda della versione, il formato XPWE potrebbe
  non conservare alcuni dati come le incidenze di
  sicurezza e di manodopera!""", 'ATTENZIONE!')
    if len(lista_misure) != 0 and arg not in ('Elenco', 'CONTABILITA'):
        if DLG.DlgSiNo(
                """Vuoi tentare un riordino delle voci secondo la struttura delle Categorie?

    Scegliendo Sì, nel caso in cui il file di origine risulti particolarmente disordinato, riceverai un messaggio che ti indica come intervenire.

    Se il risultato finale non dovesse andar bene, puoi ripetere l'importazione senza il riordino delle voci rispondendo No a questa domanda.""",
                "Richiesta") == 2:
            riordine = sorted(prova_l, key=lambda el: el[0])
            lista_misure = list()
            for el in riordine:
                lista_misure.append(el[1])
    DLG.attesa().start()
    ###
    # compilo Anagrafica generale ##########################################
    #  New_file.computo()
    # compilo Anagrafica generale ##########################################
    oSheet = oDoc.getSheets().getByName('S2')
    if oggetto is not None:
        oSheet.getCellByPosition(2, 2).String = oggetto
    if comune is not None:
        oSheet.getCellByPosition(2, 3).String = comune
    if committente is not None:
        oSheet.getCellByPosition(2, 5).String = committente
    if impresa is not None:
        oSheet.getCellByPosition(3, 16).String = impresa
###
    zoom = oDoc.CurrentController.ZoomValue
    oDoc.CurrentController.ZoomValue = 400

    # compilo Elenco Prezzi ################################################
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    # Siccome setDataArray pretende una tupla(array 1D) o una tupla di tuple(array 2D)
    # trasformo la lista_articoli da una lista di tuple a una tupla di tuple
    lista_come_array = tuple(lista_articoli)
    # Parametrizzo il range di celle a seconda della dimensione della lista
    colonne_lista = len(lista_come_array[0]
                        )  # numero di colonne necessarie per ospitare i dati
    righe_lista = len(
        lista_come_array)  # numero di righe necessarie per ospitare i dati

    oSheet.getRows().insertByIndex(3, righe_lista)
    oRange = oSheet.getCellRangeByPosition(
        0,
        3,
        colonne_lista - 1,  # l'indice parte da 0
        righe_lista + 3 - 1)
    oRange.setDataArray(lista_come_array)
    lrow = PL.getLastUsedCell(oSheet).EndRow - 1
    oSheet.getCellRangeByPosition(0, 3, 0, lrow).CellStyle = "EP-aS"
    oSheet.getCellRangeByPosition(1, 3, 1, lrow).CellStyle = "EP-a"
    oSheet.getCellRangeByPosition(2, 3, 7, lrow).CellStyle = "EP-mezzo"
    oSheet.getCellRangeByPosition(5, 3, 5, lrow).CellStyle = "EP-mezzo %"
    oSheet.getCellRangeByPosition(8, 3, 9, lrow).CellStyle = "EP-sfondo"

    oSheet.getCellRangeByPosition(11, 3, 11, lrow).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition(12, 3, 12,
                                  lrow).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition(13, 3, 13,
                                  lrow).CellStyle = 'EP statistiche_Contab_q'
    # aggiungo i capitoli alla lista delle voci ############################
    #  giallo(16777072,16777120,16777168)
    #  verde(9502608,13696976,15794160)
    #  viola(12632319,13684991,15790335)
    #  col1 = 16777072
    #  col2 = 16777120
    #  col3 = 16777168
    #  capitoli = list()
    # SUPERCAPITOLI
    try:
        for el in lista_supcap:
            tariffa = el.get('codice')
            if tariffa is not None:
                destestesa = el.get('dessintetica')
                titolo = (tariffa, destestesa, '', '', '', '', '')
                capitoli.append(titolo)
        lista_come_array = tuple(capitoli)
        colonne_lista = len(
            lista_come_array[0]
        )  # numero di colonne necessarie per ospitare i dati
        righe_lista = len(lista_come_array)  # numero di righe necessarie per ospitare i dati

        oSheet.getRows().insertByIndex(3, righe_lista)
        oRange = oSheet.getCellRangeByPosition(
            # l'indice parte da 0
            0, 3, colonne_lista - 1, righe_lista + 3 - 1)
        oRange.setDataArray(lista_come_array)
        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellStyle = "EP-aS"
        oSheet.getCellRangeByPosition(1, 3, 1, righe_lista + 3 - 1).CellStyle = "EP-a"
        oSheet.getCellRangeByPosition(2, 3, 7, righe_lista + 3 - 1).CellStyle = "EP-mezzo"
        oSheet.getCellRangeByPosition(5, 3, 5, righe_lista + 3 - 1).CellStyle = "EP-mezzo %"
        oSheet.getCellRangeByPosition(8, 3, 9, righe_lista + 3 - 1).CellStyle = "EP-sfondo"

        oSheet.getCellRangeByPosition(11, 3, 11, righe_lista + 3 - 1).CellStyle = 'EP-mezzo %'
        oSheet.getCellRangeByPosition(12, 3, 12, righe_lista + 3 - 1).CellStyle = 'EP statistiche_q'
        oSheet.getCellRangeByPosition(13, 3, 13, righe_lista + 3 - 1).CellStyle = 'EP statistiche_Contab_q'
        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellBackColor = col1

    except Exception:
        pass
    # CAPITOLI
    capitoli = list()
    try:
        for el in lista_cap:  # + lista_subcap:
            tariffa = el.get('codice')
            if tariffa is not None:
                destestesa = el.get('dessintetica')
                titolo = (tariffa, destestesa, '', '', '', '', '')
                capitoli.append(titolo)
        lista_come_array = tuple(capitoli)
        colonne_lista = len(
            lista_come_array[0]
        )  # numero di colonne necessarie per ospitare i dati
        righe_lista = len(
            lista_come_array)  # numero di righe necessarie per ospitare i dati

        oSheet.getRows().insertByIndex(3, righe_lista)
        oRange = oSheet.getCellRangeByPosition(
            0,
            3,
            colonne_lista - 1,  # l'indice parte da 0
            righe_lista + 3 - 1)
        oRange.setDataArray(lista_come_array)
        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellStyle = "EP-aS"
        oSheet.getCellRangeByPosition(1, 3, 1, righe_lista + 3 - 1).CellStyle = "EP-a"
        oSheet.getCellRangeByPosition(2, 3, 7, righe_lista + 3 - 1).CellStyle = "EP-mezzo"
        oSheet.getCellRangeByPosition(5, 3, 5, righe_lista + 3 - 1).CellStyle = "EP-mezzo %"
        oSheet.getCellRangeByPosition(8, 3, 9, righe_lista + 3 - 1).CellStyle = "EP-sfondo"

        oSheet.getCellRangeByPosition(11, 3, 11, righe_lista + 3 - 1).CellStyle = 'EP-mezzo %'
        oSheet.getCellRangeByPosition(12, 3, 12, righe_lista + 3 - 1).CellStyle = 'EP statistiche_q'
        oSheet.getCellRangeByPosition(13, 3, 13, righe_lista + 3 - 1).CellStyle = 'EP statistiche_Contab_q'
        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellBackColor = col2
    except Exception:
        pass
    # SUBCAPITOLI
    capitoli = list()
    try:
        for el in lista_subcap:
            tariffa = el.get('codice')
            if tariffa is not None:
                destestesa = el.get('dessintetica')
                titolo = (tariffa, destestesa, '', '', '', '', '')
                capitoli.append(titolo)
        lista_come_array = tuple(capitoli)

        # numero di colonne necessarie per ospitare i dati
        colonne_lista = len(lista_come_array[0])

        # numero di righe necessarie per ospitare i dati
        righe_lista = len(lista_come_array)

        oSheet.getRows().insertByIndex(4, righe_lista)
        oRange = oSheet.getCellRangeByPosition(0, 3, colonne_lista - 1, righe_lista + 3 - 1)
        oRange.setDataArray(lista_come_array)
        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellStyle = "EP-aS"
        oSheet.getCellRangeByPosition(1, 3, 1, righe_lista + 3 - 1).CellStyle = "EP-a"
        oSheet.getCellRangeByPosition(2, 3, 7, righe_lista + 3 - 1).CellStyle = "EP-mezzo"
        oSheet.getCellRangeByPosition(5, 3, 5, righe_lista + 3 - 1).CellStyle = "EP-mezzo %"
        oSheet.getCellRangeByPosition(8, 3, 9, righe_lista + 3 - 1).CellStyle = "EP-sfondo"

        oSheet.getCellRangeByPosition(11, 3, 11, righe_lista + 3 - 1).CellStyle = 'EP-mezzo %'
        oSheet.getCellRangeByPosition(12, 3, 12, righe_lista + 3 - 1).CellStyle = 'EP statistiche_q'
        oSheet.getCellRangeByPosition(13, 3, 13, righe_lista + 3 - 1).CellStyle = 'EP statistiche_Contab_q'
        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellBackColor = col3
    except Exception:
        pass
    for el in (11, 15, 19, 26):
        oSheet.getCellRangeByPosition(el, 3, el, PL.ultima_voce(oSheet)).CellStyle = 'EP-mezzo %'
    for el in (12, 16, 20, 23):
        oSheet.getCellRangeByPosition(el, 3, el, PL.ultima_voce(oSheet)).CellStyle = 'EP statistiche_q'
    for el in (13, 17, 21, 24, 25):
        oSheet.getCellRangeByPosition(el, 3, el, PL.ultima_voce(oSheet)).CellStyle = 'EP statistiche'
    #  adatta_altezza_riga('Elenco Prezzi')
    PL.riordina_ElencoPrezzi()
    #  struttura_Elenco()

    # elimino le voci che hanno analisi
    for i in reversed(range(3, PL.getLastUsedCell(oSheet).EndRow)):
        if oSheet.getCellByPosition(0, i).String in lista_tariffe_analisi:
            oSheet.getRows().removeByIndex(i, 1)

    # Compilo Analisi di prezzo
    if len(lista_analisi) != 0:
        PL.inizializza_analisi()
        oSheet = oDoc.getSheets().getByName('Analisi di Prezzo')
        for el in lista_analisi:
            prezzo_finale = el[-1]
            sStRange = PL.Circoscrive_Analisi(PL.Range2Cell()[1])
            lrow = sStRange.RangeAddress.StartRow + 2
            oSheet.getCellByPosition(0, lrow).String = el[0]
            oSheet.getCellByPosition(1, lrow).String = el[1]
            oSheet.getCellByPosition(2, lrow).String = el[2]
            y = 0
            n = lrow + 2
            for x in el[3]:
                if el[3][y][1] in ('MANODOPERA', 'MATERIALI', 'NOLI',
                                   'TRASPORTI',
                                   'ALTRE FORNITURE E PRESTAZIONI',
                                   'overflow'):
                    if el[3][y][1] != 'overflow':
                        n = PL.uFindStringCol(el[3][y][1], 1, oSheet, lrow)
                else:
                    PL.copia_riga_analisi(n)
                    if dict_articoli.get(el[3][y][0]) is not None:
                        oSheet.getCellByPosition(
                            0, n).String = dict_articoli.get(
                                el[3][y][0]).get('tariffa')
                    # per gli inserimenti liberi (L)
                    else:
                        oSheet.getCellByPosition(0, n).String = ''
                        oSheet.getCellByPosition(1, n).String = x[1]
                        oSheet.getCellByPosition(2, n).String = x[2]
                        try:
                            float(x[3].replace(',', '.'))
                            oSheet.getCellByPosition(3, n).Value = float(
                                x[3].replace(',', '.'))
                        except Exception:
                            oSheet.getCellByPosition(3, n).Value = 0
                        oSheet.getCellByPosition(4,
                                                 n).Value = float(x[4].replace(
                                                     ',', '.'))
                    if el[3][y][1] not in ('MANODOPERA', 'MATERIALI', 'NOLI',
                                           'TRASPORTI',
                                           'ALTRE FORNITURE E PRESTAZIONI',
                                           'overflow'):
                        if el[3][y][3] == '':
                            oSheet.getCellByPosition(3, n).Value = 0
                        else:
                            try:
                                float(el[3][y][3])
                                oSheet.getCellByPosition(3,
                                                         n).Value = el[3][y][3]
                            except Exception:
                                oSheet.getCellByPosition(
                                    3, n).Formula = '=' + el[3][y][3]
                y += 1
                n += 1
            sStRange = PL.Circoscrive_Analisi(lrow)
            SR = sStRange.RangeAddress.StartRow
            ER = sStRange.RangeAddress.EndRow
            for m in reversed(range(SR, ER)):
                if oSheet.getCellByPosition(
                        0,
                        m).String == 'Cod. Art.?' and oSheet.getCellByPosition(
                            0, m - 1).CellStyle == 'An-lavoraz-Cod-sx':
                    oSheet.getRows().removeByIndex(m, 1)
                if oSheet.getCellByPosition(0, m).String == 'Cod. Art.?':
                    oSheet.getCellByPosition(0, m).String = ''
            if oSheet.getCellByPosition(6, sStRange.RangeAddress.StartRow +
                                        2).Value != prezzo_finale:
                oSheet.getCellByPosition(6, sStRange.RangeAddress.StartRow +
                                         2).Value = prezzo_finale
            PL.inizializza_analisi()
        PL.elimina_voce(PL.ultima_voce(oSheet), 0)
        PL.tante_analisi_in_ep()
    if len(lista_misure) == 0:
        #  MsgBox('Importazione eseguita con successo in ' +
        # str((datetime.now() - datarif).total_seconds()) +
        # ' secondi!        \n\nImporto € ' +
        # oSheet.getCellByPosition(0, 1).String ,'')
        DLG.MsgBox(
            "Importate n." + str(len(lista_articoli)) +
            " voci dall'elenco prezzi\ndel file: " + filename, 'Avviso')
        oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
        oDoc.CurrentController.setActiveSheet(oSheet)
        oDoc.CurrentController.ZoomValue = zoom
        PL.refresh(1)
        oDialogo_attesa.endExecute()
        return
    PL.refresh(0)
    PL.doppioni()
    if arg == 'Elenco':
        oDoc.CurrentController.ZoomValue = zoom
        PL.refresh(1)
        oDialogo_attesa.endExecute()
        return
# Inserisco i dati nel COMPUTO #########################################
    if arg == 'VARIANTE':
        PL.genera_variante()
    elif arg == 'CONTABILITA':
        PL.attiva_contabilita()
    oSheet = oDoc.getSheets().getByName(arg)
    if oSheet.getCellByPosition(1, 4).String == 'Cod. Art.?':
        if arg == 'CONTABILITA':
            oSheet.getRows().removeByIndex(3, 5)
        else:
            oSheet.getRows().removeByIndex(3, 4)
    oDoc.CurrentController.select(oSheet)
    # iSheet_num = oSheet.RangeAddress.Sheet
    ###
    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = oSheet.RangeAddress.Sheet  # recupero l'index del foglio
    diz_vv = dict()

    testspcat = '0'
    testcat = '0'
    testsbcat = '0'
    x = 1
    for el in lista_misure:
        datamis = el.get('datamis')
        idspcat = el.get('idspcat')
        idcat = el.get('idcat')
        idsbcat = el.get('idsbcat')

        lrow = PL.ultima_voce(oSheet) + 1
        #  inserisco le categorie
        try:
            if idspcat != testspcat:
                testspcat = idspcat
                testcat = '0'
                PL.Inser_SuperCapitolo_arg(lrow, lista_supcat[eval(idspcat) - 1][1])
                lrow += 1
        except UnboundLocalError:
            pass
        try:
            if idcat != testcat:
                testcat = idcat
                testsbcat = '0'
                PL.Inser_Capitolo_arg(lrow, lista_cat[eval(idcat) - 1][1])
                lrow += 1
        except UnboundLocalError:
            pass
        try:
            if idsbcat != testsbcat:
                testsbcat = idsbcat
                PL.Inser_SottoCapitolo_arg(lrow, lista_subcat[eval(idsbcat) - 1][1])
        except UnboundLocalError:
            pass
        lrow = PL.ultima_voce(oSheet) + 1
        if arg == 'CONTABILITA':
            PL.ins_voce_contab(lrow=PL.ultima_voce(oSheet) + 1, arg=0)
        else:
            PL.ins_voce_computo_grezza(lrow)
        ID = el.get('id_ep')
        id_vc = el.get('id_vc')

        try:
            oSheet.getCellByPosition(
                1, lrow + 1).String = dict_articoli.get(ID).get('tariffa')
        except Exception:
            pass
        diz_vv[id_vc] = lrow + 1
        oSheet.getCellByPosition(0, lrow + 1).String = str(x)
        x = x + 1
        SC = 2
        SR = lrow + 2 + 1
        nrighe = len(el.get('lista_rig')) - 1

        if nrighe > -1:
            EC = SC + len(el.get('lista_rig')[0])
            ER = SR + nrighe

            if nrighe > 0:
                oSheet.getRows().insertByIndex(SR, nrighe)

            oRangeAddress = oSheet.getCellRangeByPosition(
                0, SR - 1, 250, SR - 1).getRangeAddress()

            for n in range(SR, SR + nrighe):
                oCellAddress = oSheet.getCellByPosition(0, n).getCellAddress()
                oSheet.copyRange(oCellAddress, oRangeAddress)
                if arg == 'CONTABILITA':
                    oSheet.getCellByPosition(1, n).String = ''
                    oSheet.getCellByPosition(
                        1, n).CellStyle = 'Comp-Bianche in mezzo_R'

            oCellRangeAddr.StartColumn = SC
            oCellRangeAddr.StartRow = SR
            oCellRangeAddr.EndColumn = EC
            oCellRangeAddr.EndRow = ER

            ###
            # INSERISCO PRIMA SOLO LE RIGHE SE NO MI FA CASINO
            SR = SR - 1
            if arg == 'CONTABILITA':
                oSheet.getCellByPosition(
                    1, SR).Formula = '=DATE(' + datamis.split(
                        '/')[2] + ';' + datamis.split(
                            '/')[1] + ';' + datamis.split('/')[0] + ')'
                oSheet.getCellByPosition(1,
                                         SR).Value = oSheet.getCellByPosition(
                                             1, SR).Value
            for mis in el.get('lista_rig'):
                if mis[0] is not None:  # descrizione
                    descrizione = mis[0].strip()
                    oSheet.getCellByPosition(2, SR).String = descrizione
                else:
                    descrizione = ''

                if mis[3] is not None:  # parti uguali
                    try:
                        oSheet.getCellByPosition(5, SR).Value = float(
                            mis[3].replace(',', '.'))
                    except ValueError:
                        oSheet.getCellByPosition(
                            5, SR).Formula = '=' + str(mis[3]).split('=')[
                                -1]  # tolgo evenutali '=' in eccesso
                if mis[4] is not None:  # lunghezza
                    try:
                        oSheet.getCellByPosition(6, SR).Value = float(
                            mis[4].replace(',', '.'))
                    except ValueError:
                        oSheet.getCellByPosition(
                            6, SR).Formula = '=' + str(mis[4]).split('=')[
                                -1]  # tolgo evenutali '=' in eccesso
                if mis[5] is not None:  # larghezza
                    try:
                        oSheet.getCellByPosition(7, SR).Value = float(
                            mis[5].replace(',', '.'))
                    except ValueError:
                        oSheet.getCellByPosition(
                            7, SR).Formula = '=' + str(mis[5]).split('=')[
                                -1]  # tolgo evenutali '=' in eccesso
                if mis[6] is not None:  # HPESO
                    try:
                        oSheet.getCellByPosition(8, SR).Value = float(
                            mis[6].replace(',', '.'))

                    except Exception:
                        oSheet.getCellByPosition(
                            8, SR).Formula = '=' + str(mis[6]).split('=')[
                                -1]  # tolgo evenutali '=' in eccesso
                if mis[8] == '2':
                    PL.parziale_core(SR)
                    oSheet.getRows().removeByIndex(SR + 1, 1)
                    descrizione = ''

                if mis[9] != '-2':
                    vedi = diz_vv.get(mis[9])
                    try:
                        PL.vedi_voce_xpwe(SR, vedi, mis[8])
                    except Exception:
                        DLG.MsgBox(
                            """Il file di origine è particolarmente disordinato.
Riordinando il computo trovo riferimenti a voci non ancora inserite.

Al termine dell'importazione controlla la voce con tariffa """ +
                            dict_articoli.get(ID).get('tariffa') +
                            """\nella riga n.""" + str(lrow + 2) +
                            """ del foglio, evidenziata qui a sinistra.""",
                            'Attenzione!')
                        oSheet.getCellByPosition(
                            44,
                            SR).String = dict_articoli.get(ID).get('tariffa')
                try:
                    mis[7]
                    if '-' in mis[7]:
                        for x in range(5, 9):
                            try:
                                if oSheet.getCellByPosition(x, SR).Value != 0:
                                    oSheet.getCellByPosition(
                                        x, SR).Value = abs(
                                            oSheet.getCellByPosition(x,
                                                                     SR).Value)
                            except Exception:
                                pass
                        PL.inverti_un_segno(SR)

                    # if oSheet.getCellByPosition(5, SR).Type.value == 'FORMULA':
                    #    va = oSheet.getCellByPosition(5, SR).Formula
                    # else:
                    #    va = oSheet.getCellByPosition(5, SR).Value

                    # if oSheet.getCellByPosition(6, SR).Type.value == 'FORMULA':
                    #    vb = oSheet.getCellByPosition(6, SR).Formula
                    # else:
                    #    vb = oSheet.getCellByPosition(6, SR).Value

                    # if oSheet.getCellByPosition(7, SR).Type.value == 'FORMULA':
                    #    vc = oSheet.getCellByPosition(7, SR).Formula
                    # else:
                    #    vc = oSheet.getCellByPosition(7, SR).Value

                    # if oSheet.getCellByPosition(8, SR).Type.value == 'FORMULA':
                    #    vd = oSheet.getCellByPosition(8, SR).Formula
                    # else:
                    #    vd = oSheet.getCellByPosition(8, SR).Value

                    # if mis[3] is None:
                    #    va = ''
                    # else:
                    #    if '^' in mis[3]:
                    #        va = eval(mis[3].replace('^', '**'))
                    #    else:
                    #        va = eval(mis[3])

                except Exception:
                    pass
                SR = SR + 1
    PL.numera_voci()
    try:
        PL.Rinumera_TUTTI_Capitoli2()
    except Exception:
        pass
    oDoc.CurrentController.ZoomValue = zoom
    PL.refresh(1)
    #  MsgBox('Importazione eseguita con successo in ' +
    # str((datetime.now() - datarif).total_seconds()) +
    # ' secondi!        \n\nImporto € ' +
    # oSheet.getCellByPosition(0, 1).String ,'')
    oDialogo_attesa.endExecute()
    PL._gotoSheet(arg)
    #  if uFindStringCol('Riepilogo strutturale delle Categorie', 2, oSheet) !='None':
    #  firme_in_calce()
    PL.adatta_altezza_riga()
    DLG.MsgBox('Importazione di\n\n' + arg + '\n\neseguita con successo!', '')

########################################################################


def MENU_XML_toscana_import():
    '''
    Importazione di un prezzario XML della regione Toscana
    in tabella Elenco Prezzi del template COMPUTO.
    '''
    oDoc = getDocument()

    DLG.MsgBox('Questa operazione potrebbe richiedere del tempo.', 'Avviso')
    PL.New_file.computo(0)

    try:
        filename = DLG.filedia('Scegli il file XML Toscana da importare', '*.xml')
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
                PL.getLastUsedCell(oDoc.CurrentController.ActiveSheet).EndColumn == 0 and
                PL.getLastUsedCell(oDoc.CurrentController.ActiveSheet).EndRow == 0):
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
    oDoc = getDocument()
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
    # ~ set_larghezza_colonne()
    Toolbars.Vedi()
    # ~ adatta_altezza_riga('Elenco Prezzi')
    # ~ riordina_ElencoPrezzi()
    oDialogo_attesa.endExecute()
    PL.struttura_Elenco()
    oSheet.getCellRangeByName('F2').String = 'prezzi'
    oSheet.getCellRangeByName('E2').Formula = ('=COUNT(E3:E' + str(PL.getLastUsedCell(oSheet).EndRow + 1) +
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


def MENU_sardegna_2019():
    '''
    @@@ DA DOCUMENTARE
    '''
    oDoc = getDocument()

    try:
        oDoc.getSheets().insertNewByName('nuova_tabella', 2)
    except Exception:
        pass

    oSheet0 = oDoc.getSheets().getByName('Worksheet')
    oSheet1 = oDoc.getSheets().getByName('nuova_tabella')
    # fine = PL.getLastUsedCell(oSheet0).EndRow + 1
    n = 1
    test1 = test2 = test3 = test4 = 1
    for i in range(2, 50):
        cod = oSheet0.getCellByPosition(0, i).String
        cods = cod.split('.')
        # ~ chi(cod)
        cod0 = cods[0]
        if test1 == 1:
            cod1 = cods[1]
            # ~ test1 =1
        if test2 == 1:
            cod2 = cods[2]
            # ~ test2 =1
        # if test3 == 1:
        #    cod3 = cods[3]
        # ~ test3 =1
        cap1 = oSheet0.getCellByPosition(1, i).String
        cap2 = oSheet0.getCellByPosition(2, i).String
        cap3 = oSheet0.getCellByPosition(3, i).String
        des = oSheet0.getCellByPosition(4, i).String
        um = oSheet0.getCellByPosition(5, i).String
        sic = oSheet0.getCellByPosition(10, i).Value
        prz = oSheet0.getCellByPosition(7, i).Value
        mdo = oSheet0.getCellByPosition(13, i).Value

        if test1 == 1:
            oSheet1.getCellByPosition(0, n).String = cod0
            oSheet1.getCellByPosition(1, n).String = cap1
            test1 = 0
        elif test2 == 1:
            n += 1
            oSheet1.getCellByPosition(0, n).String = cod0 + '.' + cod1
            oSheet1.getCellByPosition(1, n).String = cap2
            test2 = 0
        elif test3 == 1:
            n += 1
            oSheet1.getCellByPosition(0, n).String = cod0 + '.' + cod1 + '.' + cod2
            oSheet1.getCellByPosition(1, n).String = cap3
            test3 = 0
        elif test4 == 1:
            n += 1
            oSheet1.getCellByPosition(0, n).String = cod
            oSheet1.getCellByPosition(1, n).String = des
            oSheet1.getCellByPosition(2, n).String = um
            oSheet1.getCellByPosition(3, n).String = sic
            oSheet1.getCellByPosition(4, n).String = prz
            oSheet1.getCellByPosition(5, n).String = mdo
            # ~ n += 1

########################################################################


def MENU_basilicata_2020():
    '''
    Adatta la struttura del prezzario rilasciato dalla regione Basilicata
    partendo dalle colonne: CODICE	DESCRIZIONE	U. MISURA	PREZZO	MANODOPERA
    Il risultato ottenuto va inserito in Elenco Prezzi.
    '''
    oDoc = getDocument()
    for el in ('CAPITOLI', 'CATEGORIE', 'VOCI'):
        oSheet = oDoc.getSheets().getByName(el)
        oSheet.getRows().removeByIndex(0, 1)
    oSheet = oDoc.getSheets().getByName('CATEGORIE')
    PL._gotoSheet('CATEGORIE')
    fine = PL.getLastUsedCell(oSheet).EndRow + 1
    for i in range(0, fine):
        oSheet.getCellByPosition(1, i).String = (
            oSheet.getCellByPosition(0, i).String +
            "." +
            oSheet.getCellByPosition(1, i).String)

    oSheet.getColumns().removeByIndex(0, 1)
    oSheet = oDoc.getSheets().getByName('VOCI')
    PL._gotoSheet('VOCI')
    oSheet.getColumns().removeByIndex(0, 3)
    oSheet = oDoc.getSheets().getByName('SOTTOVOCI')
    PL._gotoSheet('SOTTOVOCI')
    oSheet.getColumns().removeByIndex(0, 4)
    PL.join_sheets()
    oSheet = oDoc.getSheets().getByName('unione_fogli')
    PL._gotoSheet('unione_fogli')
    oSheet.getRows().removeByIndex(0, 1)
    PL.ordina_col(1)
    fine = PL.getLastUsedCell(oSheet).EndRow + 1
    for i in range(0, fine):
        if len(oSheet.getCellByPosition(0, i).String.split('.')) == 3:
            madre = oSheet.getCellByPosition(1, i).String
        elif len(oSheet.getCellByPosition(0, i).String.split('.')) == 4:
            if oSheet.getCellByPosition(1, i).String != '':
                oSheet.getCellByPosition(1, i).String = (
                    madre +
                    "\n- " +
                    oSheet.getCellByPosition(1, i).String)
            else:
                oSheet.getCellByPosition(1, i).String = madre
            oSheet.getCellByPosition(4, i).Value = oSheet.getCellByPosition(4, i).Value / 100
    for i in reversed(range(0, fine)):
        if len(oSheet.getCellByPosition(0, i).String.split('.')) == 3:
            oSheet.getRows().removeByIndex(i, 1)
    oSheet.getRows().removeByIndex(0, 1)
    oSheet.getColumns().insertByIndex(3, 1)

########################################################################


def MENU_Piemonte_2019():
    '''
    Adatta la struttura del prezzario rilasciato dalla regione Piemonte
    partendo dalle colonne: Sez.	Codice	Descrizione	U.M.	Euro	Manod. lorda	% Manod.	Note
    Il risultato ottenuto va inserito in Elenco Prezzi.
    '''
    oDoc = getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    fine = PL.getLastUsedCell(oSheet).EndRow + 1
    elenco = list()
    for i in range(0, fine):
        if len(oSheet.getCellByPosition(1, i).String.split('.')) <= 2:
            cod = oSheet.getCellByPosition(1, i).String
            des = oSheet.getCellByPosition(2, i).String.replace('\n\n', '\n')
            um = ''
            eur = ''
            mdol = ''
            mdo = ''
            if oSheet.getCellByPosition(7, i).String != '':
                des = des + '\n(' + oSheet.getCellByPosition(7, i).String + ')'
            elenco.append((cod, des, um, '', eur, mdo, mdol))

        if len(oSheet.getCellByPosition(1, i).String.split('.')) == 3:
            cod = oSheet.getCellByPosition(1, i).String
            des = oSheet.getCellByPosition(2, i).String.replace(' \n\n', '')
            madre = des
            um = ''
            eur = ''
            mdol = ''
            mdo = ''
            if oSheet.getCellByPosition(7, i).String != '':
                des = des + '\n(' + oSheet.getCellByPosition(7, i).String + ')'
            # ~elenco.append ((cod, des, um, '', eur, mdo, mdol))
        if len(oSheet.getCellByPosition(1, i).String.split('.')) == 4:
            cod = oSheet.getCellByPosition(1, i).String
            des = madre
            if oSheet.getCellByPosition(2, i).String != '...':
                des = madre + '\n- ' + oSheet.getCellByPosition(2, i).String.replace('\n\n', '')
            um = oSheet.getCellByPosition(3, i).String
            eur = ''
            if oSheet.getCellByPosition(4, i).Value != 0:
                eur = oSheet.getCellByPosition(4, i).Value
            mdol = ''
            if oSheet.getCellByPosition(5, i).Value != 0:
                mdol = oSheet.getCellByPosition(5, i).Value
            mdo = ''
            if oSheet.getCellByPosition(6, i).Value != 0:
                mdo = oSheet.getCellByPosition(6, i).Value
            # ~note= oSheet.getCellByPosition(7, i).String
            elenco.append((cod, des, um, '', eur, mdo, mdol))

    try:
        oDoc.getSheets().insertNewByName('nuova_tabella', 2)
    except Exception:
        pass

    PL._gotoSheet('nuova_tabella')
    oSheet = oDoc.getSheets().getByName('nuova_tabella')
    elenco = tuple(elenco)
    oRange = oSheet.getCellRangeByPosition(0,
                                           0,
                                           # l'indice parte da 0
                                           len(elenco[0]) - 1,
                                           len(elenco) - 1)
    oRange.setDataArray(elenco)
