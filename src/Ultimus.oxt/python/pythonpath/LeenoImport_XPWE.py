u"""
Importazione computo/variante/contabilità/prezzario
dal formato XPWE
"""
import logging

from xml.etree.ElementTree import ElementTree, ParseError

import uno

import LeenoUtils
import LeenoFormat

import pyleeno as PL
import LeenoDialogs as DLG

import SheetUtils
import LeenoConfig
import LeenoSheetUtils
import LeenoAnalysis

import Dialogs


def leggiAnagraficaGenerale(dati):
    ''' legge i dati anagrafici generali '''

    datiAnagrafici = {}
    DatiGenerali = dati.getchildren()[0][0]

    datiAnagrafici['comune'] = DatiGenerali[1].text or ''
    datiAnagrafici['oggetto'] = DatiGenerali[3].text or ''
    datiAnagrafici['committente'] = DatiGenerali[4].text or ''
    datiAnagrafici['impresa'] = DatiGenerali[5].text or ''
    '''
    datiAnagrafici['percprezzi'] = DatiGenerali[0].text or ''
    datiAnagrafici['provincia'] = DatiGenerali[2].text or ''
    datiAnagrafici['parteopera'] = DatiGenerali[6].text or ''
    '''
    return datiAnagrafici


def leggiSuperCapitoli(CapCat):
    ''' legge SuperCapitoli '''

    # PweDGSuperCapitoli
    listaSuperCapitoli = []
    if CapCat.find('PweDGSuperCapitoli'):
        PweDGSuperCapitoli = CapCat.find('PweDGSuperCapitoli').getchildren()
        for elem in PweDGSuperCapitoli:
            id_sc = elem.get('ID')
            codice = elem.find('Codice').text
            try:
                codice = elem.find('Codice').text
            except AttributeError:
                codice = ''
            dessintetica = elem.find('DesSintetica').text
            percentuale = elem.find('Percentuale').text
            diz = {}
            diz['id_sc'] = id_sc
            diz['codice'] = codice
            diz['dessintetica'] = dessintetica
            diz['percentuale'] = percentuale
            listaSuperCapitoli.append(diz)
    return listaSuperCapitoli


def leggiCapitoli(CapCat):
    ''' legge Capitoli '''

    # PweDGCapitoli
    listaCapitoli = []
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
            diz = {}
            diz['id_sc'] = id_sc
            diz['codice'] = codice
            diz['dessintetica'] = dessintetica
            diz['percentuale'] = percentuale
            listaCapitoli.append(diz)
    return listaCapitoli


def leggiSottoCapitoli(CapCat):
    ''' legge SottoCapitoli '''

    # PweDGSubCapitoli
    listaSottoCapitoli = []
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
            diz = {}
            diz['id_sc'] = id_sc
            diz['codice'] = codice
            diz['dessintetica'] = dessintetica
            diz['percentuale'] = percentuale
            listaSottoCapitoli.append(diz)
    return listaSottoCapitoli


def leggiSuperCategorie(CapCat):
    ''' legge SuperCategorie '''

    # PweDGSuperCategorie
    listaSuperCategorie = []
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
            listaSuperCategorie.append(supcat)
    return listaSuperCategorie


def leggiCategorie(CapCat):
    ''' legge Categorie '''

    # PweDGCategorie
    listaCategorie = []
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
            listaCategorie.append(cat)
    return listaCategorie


def leggiSottoCategorie(CapCat):
    ''' legge SottoCategorie '''

    # PweDGSubCategorie
    listaSottoCategorie = []
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
            listaSottoCategorie.append(subcat)
    return listaSottoCategorie


def leggiCapitoliCategorie(dati):
    ''' legge capitoli e categorie '''

    res = {
        'SuperCapitoli': [],
        'Capitoli': [],
        'SottoCapitoli': [],
        'SuperCategorie': [],
        'Categorie': [],
        'SottoCategorie': []
    }

    # PweDGCapitoliCategorie
    try:
        CapCat = dati.find('PweDGCapitoliCategorie')
    except AttributeError:
        return res

    # legge SuperCapitoli
    res['SuperCapitoli'] = leggiSuperCapitoli(CapCat)
    # legge Capitoli
    res['Capitoli'] = leggiCapitoli(CapCat)
    # legge SottoCapitoli
    res['SottoCapitoli'] = leggiSottoCapitoli(CapCat)
    # legge SuperCategorie
    res['SuperCategorie'] = leggiSuperCategorie(CapCat)
    # legge Categorie
    res['Categorie'] = leggiCategorie(CapCat)
    # legge SottoCategorie
    res['SottoCategorie'] = leggiSottoCategorie(CapCat)

    return res


def leggiDatiGeneraliAnalisi(dati):
    ''' legge i dati generali di analisi '''

    PweDGAnalisi = dati.find('PweDGModuli').getchildren()[0]

    try:
        speseGenerali = float(PweDGAnalisi.find('SpeseGenerali').text) / 100
    except ValueError:
        speseGenerali = 0

    try:
        utiliImpresa = float(PweDGAnalisi.find('UtiliImpresa').text) / 100
    except ValueError:
        utiliImpresa = 0

    try:
        oneriAccessoriSicurezza = float(PweDGAnalisi.find('OneriAccessoriSc').text) / 100
    except ValueError:
        oneriAccessoriSicurezza = 0
    '''
    speseutili = PweDGAnalisi.find('SpeseUtili').text
    confquantita = PweDGAnalisi.find('ConfQuantita').text
    '''
    return {
        'SpeseGenerali': speseGenerali,
        'UtiliImpresa': utiliImpresa,
        'OneriAccessoriSicurezza': oneriAccessoriSicurezza
    }


def leggiApprossimazioni(dati):
    ''' legge le impostazioni di approssimazione numerica '''

    PweDGConfigNumeri = dati.find('PweDGConfigurazione')
    if PweDGConfigNumeri is None:
        return {}
    PweDGConfigNumeri = PweDGConfigNumeri.getchildren()[0]
    res = {}

    partiUguali = PweDGConfigNumeri.find('PartiUguali')
    if partiUguali:
        res['PartiUguali'] = int(partiUguali.text.split('.')[-1].split('|')[0])

    larghezza = PweDGConfigNumeri.find('Larghezza')
    if larghezza:
        res['Larghezza'] = int(larghezza.text.split('.')[-1].split('|')[0])

    lunghezza = PweDGConfigNumeri.find('Lunghezza')
    if lunghezza:
        res['Lunghezza'] = int(lunghezza.text.split('.')[-1].split('|')[0])

    hPeso = PweDGConfigNumeri.find('HPeso')
    if hPeso:
        res['HPeso'] = int(hPeso.text.split('.')[-1].split('|')[0])

    quantita = PweDGConfigNumeri.find('Quantita')
    if quantita:
        res['Quantita'] = int(quantita.text.split('.')[-1].split('|')[0])

    prezzi = PweDGConfigNumeri.find('Prezzi')
    if prezzi:
        res['Prezzi'] = int(prezzi.text.split('.')[-1].split('|')[0])

    prezziTotale = PweDGConfigNumeri.find('PrezziTotale')
    if prezziTotale:
        res['PrezziTotale'] = int(prezziTotale.text.split('.')[-1].split('|')[0])
    '''
    Divisa = PweDGConfigNumeri.find('Divisa').text
    ConversioniIN = PweDGConfigNumeri.find('ConversioniIN').text
    FattoreConversione = PweDGConfigNumeri.find('FattoreConversione').text
    Cambio = PweDGConfigNumeri.find('Cambio').text
    ConvPrezzi = PweDGConfigNumeri.find('ConvPrezzi').text.split('.')[-1].split('|')[0]
    ConvPrezziTotale = PweDGConfigNumeri.find('ConvPrezziTotale').text.split('.')[-1].split('|')[0]
    IncidenzaPercentuale = PweDGConfigNumeri.find('IncidenzaPercentuale').text.split('.')[-1].split('|')[0]
    Aliquote = PweDGConfigNumeri.find('Aliquote').text.split('.')[-1].split('|')[0]
    '''
    return res


def compilaAnagraficaGenerale(oDoc, datiAnagrafici):
    ''' compilo Anagrafica generale '''

    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = datiAnagrafici['oggetto']
    oSheet.getCellByPosition(2, 3).String = datiAnagrafici['comune']
    oSheet.getCellByPosition(2, 5).String = datiAnagrafici['committente']
    oSheet.getCellByPosition(3, 16).String = datiAnagrafici['impresa']


def compilaDatiGeneraliAnalisi(oDoc, datiGeneraliAnalisi):
    ''' compila i dati generali dell'analisi, ovvero le percentuali relative '''

    oSheet = oDoc.getSheets().getByName('S1')

    oSheet.getCellByPosition(7, 318).Value = datiGeneraliAnalisi['OneriAccessoriSicurezza']
    oSheet.getCellByPosition(7, 319).Value = datiGeneraliAnalisi['SpeseGenerali']
    oSheet.getCellByPosition(7, 320).Value = datiGeneraliAnalisi['UtiliImpresa']


def compilaApprossimazioni(oDoc, approssimazioni):
    ''' compila le approssimazioni modificando la precisione nelle relative celle '''

    if 'PartiUguali' in approssimazioni:
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a PU', approssimazioni['PartiUguali'])
    if 'Lunghezza' in approssimazioni:
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a LUNG', approssimazioni['Lunghezza'])
    if 'Larghezza' in approssimazioni:
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a LARG', approssimazioni['Larghezza'])
    if 'HPeso' in approssimazioni:
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a peso', approssimazioni['HPeso'])
    if 'Quantita' in approssimazioni:
        for el in ('Comp-Variante num sotto', 'An-lavoraz-input', 'Blu'):
            LeenoFormat.setCellStyleDecimalPlaces(el, approssimazioni['Quantita'])
    if 'Prezzi' in approssimazioni:
        for el in ('comp sotto Unitario', 'An-lavoraz-generica'):
            LeenoFormat.setCellStyleDecimalPlaces(el, approssimazioni['Prezzi'])
    if 'PrezziTotale' in approssimazioni:
        for el in ('comp sotto Euro Originale', 'Livello-0-scritta mini val',
                   'Livello-1-scritta mini val', 'livello2 scritta mini',
                   'Comp TOTALI', 'Ultimus_totali_1', 'Ultimus_bordo',
                   'ULTIMUS_3', 'Ultimus_Bordo_sotto',
                   'Comp-Variante num sotto', 'An-valuta-dx', 'An-1v-dx',
                   'An-lavoraz-generica', 'An-lavoraz-Utili-num sin'):
            LeenoFormat.setCellStyleDecimalPlaces(el, approssimazioni['PrezziTotale'])


def leggiElencoPrezzi(misurazioni):
    ''' legge l'elenco prezzi e le eventuali voci di analisi prezzi '''

    dizionarioArticoli = {}
    listaArticoli = []
    listaAnalisi = []
    listaTariffeAnalisi = []

    PweElencoPrezzi = misurazioni.getchildren()[0]

    # leggo l'elenco prezzi
    epitems = PweElencoPrezzi.findall('EPItem')

    for elem in epitems:
        id_ep = elem.get('ID')
        dizionarioArticolo = {}
        tipoep = elem.find('TipoEP').text
        tariffa = elem.find('Tariffa').text or ''
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

        unmisura = elem.find('UnMisura').text or ''
        if elem.find('Prezzo1').text == '0' or elem.find('Prezzo1').text is None:
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
        '''
        try:
           idsbcap = elem.find('IDSbCap').text
        except AttributeError:
           idsbcap = ''
        '''
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

        dizionarioArticolo['tipoep'] = tipoep
        dizionarioArticolo['tariffa'] = tariffa
        dizionarioArticolo['articolo'] = articolo
        dizionarioArticolo['desridotta'] = desridotta
        dizionarioArticolo['destestesa'] = destestesa
        dizionarioArticolo['desridotta'] = desridotta
        dizionarioArticolo['desbreve'] = desbreve
        dizionarioArticolo['unmisura'] = unmisura
        dizionarioArticolo['prezzo1'] = prezzo1
        dizionarioArticolo['prezzo2'] = prezzo2
        dizionarioArticolo['prezzo3'] = prezzo3
        dizionarioArticolo['prezzo4'] = prezzo4
        dizionarioArticolo['prezzo5'] = prezzo5
        dizionarioArticolo['idspcap'] = idspcap
        dizionarioArticolo['idcap'] = idcap
        dizionarioArticolo['flags'] = flags
        dizionarioArticolo['data'] = data
        dizionarioArticolo['adrinternet'] = adrinternet

        dizionarioArticolo['IncSIC'] = IncSIC
        dizionarioArticolo['IncMDO'] = IncMDO
        dizionarioArticolo['IncMAT'] = IncMDO
        dizionarioArticolo['IncATTR'] = IncMDO
        '''
        dizionarioArticolo['pweepanalisi'] = pweepanalisi
        '''
        # il dizionario articolo lo tengo completo di voci
        # con analisi di prezzi; la lista articoli, per contro,
        # visto che è usata SOLO per riempire il foglio EP,
        # viene direttamente purgata dalle voci con analisi
        dizionarioArticoli[id_ep] = dizionarioArticolo

        # leggo analisi di prezzo

        pweepanalisi = elem.find('PweEPAnalisi')
        PweEPAR = pweepanalisi.find('PweEPAR')
        if PweEPAR is not None:
            EPARItem = PweEPAR.findall('EPARItem')
            analisi = []
            for el in EPARItem:
                '''
                id_an = el.get('ID')
                an_tipo = el.find('Tipo').text
                '''
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
                '''
                an_fld = el.find('FieldCTL').text
                '''
                an_rigo = (id_ep, an_des, an_um, an_qt, an_pr)
                analisi.append(an_rigo)
            listaAnalisi.append([tariffa, destestesa, unmisura, analisi, prezzo1])
            listaTariffeAnalisi.append(tariffa)
        else:
            # analisi non presente, includo il prezzo nell'elenco
            articolo_modificato = (
                tariffa,
                destestesa,
                unmisura,
                IncSIC,
                prezzo1,
                IncMDO,
                IncMAT,
                IncATTR)
            listaArticoli.append(articolo_modificato)
    return {
        'DizionarioArticoli': dizionarioArticoli,
        'ListaArticoli': listaArticoli,
        'ListaAnalisi': listaAnalisi,
        'ListaTariffeAnalisi': listaTariffeAnalisi
    }


def leggiMisurazioni(misurazioni, ordina):
    ''' leggo voci di misurazione e righe '''

    listaMisure = []
    try:
        PweVociComputo = misurazioni.getchildren()[1]
        vcitems = PweVociComputo.findall('VCItem')
        prova_l = []
        for elem in vcitems:
            diz_misura = {}
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
            '''
            try:
               CodiceWBS = elem.find('CodiceWBS').text
            except AttributeError:
               CodiceWBS = ''
            '''
            righi_mis = elem.getchildren()[-1].findall('RGItem')
            riga_misura = []
            lista_righe = []
            new_id_l = []

            for el in righi_mis:
                '''
                rgitem = el.get('ID')
                '''
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
            listaMisure.append(diz_misura)

        # se richiesto ordina le misure
        if len(listaMisure) != 0 and ordina:
            riordine = sorted(prova_l, key=lambda el: el[0])
            listaMisure = []
            for el in riordine:
                listaMisure.append(el[1])

    except IndexError:
        DLG.MsgBox(
            "Nel file scelto non risultano esserci voci di misurazione,\n"
            "perciò saranno importate le sole voci di Elenco Prezzi.\n\n"
            "Si tenga conto che:\n"
            "- sarà importato solo il 'Prezzo 1' dell'elenco;\n"
            "- a seconda della versione, il formato XPWE potrebbe\n"
            "  non conservare alcuni dati come le incidenze di\n"
            "  sicurezza e di manodopera!",
            'ATTENZIONE!'
        )

    return listaMisure


def stileCelleElencoPrezzi(oSheet, startRow, endRow, color=None):
    oSheet.getCellRangeByPosition(0, startRow, 0, endRow).CellStyle = "EP-aS"
    oSheet.getCellRangeByPosition(1, startRow, 1, endRow).CellStyle = "EP-a"
    oSheet.getCellRangeByPosition(2, startRow, 4, endRow).CellStyle = "EP-mezzo"
    oSheet.getCellRangeByPosition(5, startRow, 5, endRow).CellStyle = "EP-mezzo %"
    oSheet.getCellRangeByPosition(6, startRow, 7, endRow).CellStyle = "EP-mezzo"
    oSheet.getCellRangeByPosition(8, startRow, 9, endRow).CellStyle = "EP-sfondo"
    oSheet.getCellRangeByPosition(11, startRow, 11, endRow).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition(15, startRow, 15, endRow).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition(19, startRow, 19, endRow).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition(26, startRow, 26, endRow).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition(12, startRow, 12, endRow).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition(16, startRow, 16, endRow).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition(20, startRow, 20, endRow).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition(23, startRow, 23, endRow).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition(13, startRow, 13, endRow).CellStyle = 'EP statistiche'
    oSheet.getCellRangeByPosition(17, startRow, 17, endRow).CellStyle = 'EP statistiche'
    oSheet.getCellRangeByPosition(21, startRow, 21, endRow).CellStyle = 'EP statistiche'
    oSheet.getCellRangeByPosition(24, startRow, 25, endRow).CellStyle = 'EP statistiche'
    if color is not None:
        oSheet.getCellRangeByPosition(0, startRow, 0, endRow).CellBackColor = color

def estraiDatiCapitoliCategorie(capitoliCategorie, catName):
    resList = []
    for el in capitoliCategorie[catName]:
        tariffa = el.get('codice')
        if tariffa is not None:
            destestesa = el.get('dessintetica')
            titolo = (tariffa, destestesa, '', '', '', '', '')
            resList.append(titolo)
    return tuple(resList)

def riempiBloccoElencoPrezzi(oSheet, dati, col, progress):

    progStart = progress.getValue()
    righe = len(dati)
    colonne = len(dati[0])

    # i dati partono dalla riga 3 (quarta, in effetti)
    oSheet.getRows().insertByIndex(3, righe)

    riga = 0
    step = 100
    while riga < righe:
        sliced = dati[riga:riga + step]
        num = len(sliced)
        oRange = oSheet.getCellRangeByPosition(
            0,
            3 + riga,
            colonne - 1,
            3 + riga + num - 1)
        oRange.setDataArray(sliced)

        # modifica lo stile del gruppo di celle
        stileCelleElencoPrezzi(oSheet, 3 + riga, 3 + riga + num - 1, col)

        riga = riga + num
        progress.setValue(riga + progStart)


def compilaElencoPrezzi(oDoc, capitoliCategorie, elencoPrezzi, progress):
    ''' compila l'elenco prezzi '''

    # per prima cosa estrae le liste di EP, capitoli, eccetera
    # e le converte in array

    # articoli
    arrayArticoli = tuple(elencoPrezzi['ListaArticoli'])
    righeArticoli = len(arrayArticoli)
    # SuperCapitoli
    arraySuperCapitoli = estraiDatiCapitoliCategorie(capitoliCategorie, 'SuperCapitoli')
    righeSuperCapitoli = len(arraySuperCapitoli)
    # capitoli
    arrayCapitoli = estraiDatiCapitoliCategorie(capitoliCategorie, 'Capitoli')
    righeCapitoli = len(arrayCapitoli)
    # SottoCapitoli
    arraySottoCapitoli = estraiDatiCapitoliCategorie(capitoliCategorie, 'SottoCapitoli')
    righeSottoCapitoli = len(arraySottoCapitoli)

    # numero totale di righe da inserire
    righeTotali = righeArticoli + righeSuperCapitoli + righeCapitoli + righeSottoCapitoli

    # inizializza la progressbar
    progress.setLimits(0, righeTotali)
    progress.setText("Scrittura elenco prezzi")
    progress.setValue(0)

    # compilo Elenco Prezzi
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')

    riempiBloccoElencoPrezzi(oSheet, arrayArticoli, None, progress)
    '''
    aggiungo i capitoli alla lista delle voci
     giallo(16777072,16777120,16777168)
     verde(9502608,13696976,15794160)
     viola(12632319,13684991,15790335)
    SUPERCAPITOLI
    '''
    # SuperCapitoli
    progress.setText("Scrittura supercapitoli")
    if righeSuperCapitoli:
        riempiBloccoElencoPrezzi(oSheet, arraySuperCapitoli, 16777072, progress)

    # Capitoli
    progress.setText("Scrittura capitoli")
    if righeCapitoli:
        riempiBloccoElencoPrezzi(oSheet, arrayCapitoli, 16777120, progress)

    # SottoCapitoli
    progress.setText("Scrittura sottocapitoli")
    if righeSottoCapitoli:
        riempiBloccoElencoPrezzi(oSheet, arraySottoCapitoli, 16777168, progress)

    progress.setText("Ordinamento elenco prezzi")
    PL.riordina_ElencoPrezzi()

    # elimino le voci che hanno analisi
    '''
    ELIMINAZIONE SPOSTATA A MONTE, NELL'ARRAY
    PRIMA DELL'INSERIMENTO NEL FOGLIO

    print("Eliminazione voci doppie elenco prezzi")
    progress.setText("Eliminazione voci doppie elenco prezzi")
    for i in reversed(range(3, SheetUtils.getLastUsedRow(oSheet))):
        if oSheet.getCellByPosition(0, i).String in elencoPrezzi['ListaTariffeAnalisi']:
            oSheet.getRows().removeByIndex(i, 1)
    '''
    print("Fine scrittura elenco prezzi")
    progress.setText("Fine scrittura elenco prezzi")


def compilaAnalisiPrezzi(oDoc, elencoPrezzi):
    ''' Compilo Analisi di prezzo '''
    if len(elencoPrezzi['ListaAnalisi']) != 0:
        oSheet, startRow = LeenoAnalysis.inizializzaAnalisi(oDoc)
        for el in elencoPrezzi['ListaAnalisi']:
            prezzo_finale = el[-1]
            sStRange = LeenoAnalysis.circoscriveAnalisi(oSheet, startRow)
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
                        n = SheetUtils.uFindStringCol(el[3][y][1], 1, oSheet, lrow)
                else:
                    LeenoAnalysis.copiaRigaAnalisi(oSheet, n)
                    if elencoPrezzi['DizionarioArticoli'].get(el[3][y][0]) is not None:
                        oSheet.getCellByPosition(
                            0, n).String = elencoPrezzi['DizionarioArticoli'].get(
                                el[3][y][0]).get('tariffa')
                    # per gli inserimenti liberi (L)
                    else:
                        oSheet.getCellByPosition(0, n).String = ''
                        oSheet.getCellByPosition(1, n).String = x[1]
                        oSheet.getCellByPosition(2, n).String = x[2]
                        try:
                            float(x[3].replace(',', '.'))
                            oSheet.getCellByPosition(3, n).Value = float(x[3].replace(',', '.'))
                        except Exception:
                            oSheet.getCellByPosition(3, n).Value = 0
                        oSheet.getCellByPosition(4, n).Value = float(x[4].replace(',', '.'))
                    if el[3][y][1] not in ('MANODOPERA', 'MATERIALI', 'NOLI',
                                           'TRASPORTI',
                                           'ALTRE FORNITURE E PRESTAZIONI',
                                           'overflow'):
                        if el[3][y][3] == '':
                            oSheet.getCellByPosition(3, n).Value = 0
                        else:
                            try:
                                float(el[3][y][3])
                                oSheet.getCellByPosition(3, n).Value = el[3][y][3]
                            except Exception:
                                oSheet.getCellByPosition(3, n).Formula = '=' + el[3][y][3]
                y += 1
                n += 1
            sStRange = LeenoAnalysis.circoscriveAnalisi(oSheet, lrow)
            SR = sStRange.RangeAddress.StartRow
            ER = sStRange.RangeAddress.EndRow
            for m in reversed(range(SR, ER)):
                if(oSheet.getCellByPosition(0, m).String == 'Cod. Art.?' and
                   oSheet.getCellByPosition(0, m - 1).CellStyle == 'An-lavoraz-Cod-sx'):
                    oSheet.getRows().removeByIndex(m, 1)
                if oSheet.getCellByPosition(0, m).String == 'Cod. Art.?':
                    oSheet.getCellByPosition(0, m).String = ''
            if oSheet.getCellByPosition(6, sStRange.RangeAddress.StartRow + 2).Value != prezzo_finale:
                oSheet.getCellByPosition(6, sStRange.RangeAddress.StartRow + 2).Value = prezzo_finale
            oSheet, startRow = LeenoAnalysis.inizializzaAnalisi(oDoc)

        LeenoSheetUtils.eliminaVoce(oSheet, LeenoSheetUtils.cercaUltimaVoce(oSheet))
        print("Tante analisi")


def compilaComputo(oDoc, elaborato, capitoliCategorie, elencoPrezzi, listaMisure):
    ''' compila il computo '''

    if elaborato == 'VARIANTE':
        PL.genera_variante()
    elif elaborato == 'CONTABILITA':
        PL.attiva_contabilita()
    oSheet = oDoc.getSheets().getByName(elaborato)
    if oSheet.getCellByPosition(1, 4).String == 'Cod. Art.?':
        if elaborato == 'CONTABILITA':
            oSheet.getRows().removeByIndex(3, 5)
        else:
            oSheet.getRows().removeByIndex(3, 4)
    oDoc.CurrentController.select(oSheet)

    oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oCellRangeAddr.Sheet = oSheet.RangeAddress.Sheet  # recupero l'index del foglio
    diz_vv = {}

    testspcat = '0'
    testcat = '0'
    testsbcat = '0'
    x = 1
    for el in listaMisure:
        datamis = el.get('datamis')
        idspcat = el.get('idspcat')
        idcat = el.get('idcat')
        idsbcat = el.get('idsbcat')

        lrow = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1
        #  inserisco le categorie
        try:
            if idspcat != testspcat:
                testspcat = idspcat
                testcat = '0'
                PL.Inser_SuperCapitolo_arg(lrow, capitoliCategorie['SuperCategorie'][eval(idspcat) - 1][1])
                lrow += 1
        except UnboundLocalError:
            pass
        try:
            if idcat != testcat:
                testcat = idcat
                testsbcat = '0'
                PL.Inser_Capitolo_arg(lrow, capitoliCategorie['Categorie'][eval(idcat) - 1][1])
                lrow += 1
        except UnboundLocalError:
            pass
        try:
            if idsbcat != testsbcat:
                testsbcat = idsbcat
                PL.Inser_SottoCapitolo_arg(lrow, capitoliCategorie['SottoCategorie'][eval(idsbcat) - 1][1])
        except UnboundLocalError:
            pass
        lrow = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1
        if elaborato == 'CONTABILITA':
            PL.ins_voce_contab(lrow=LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1, elaborato=0)
        else:
            PL.ins_voce_computo_grezza(lrow)
        ID = el.get('id_ep')
        id_vc = el.get('id_vc')

        try:
            oSheet.getCellByPosition(
                1, lrow + 1).String = elencoPrezzi['DizionarioArticoli'].get(ID).get('tariffa')
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

            oRangeAddress = oSheet.getCellRangeByPosition(0, SR - 1, 250, SR - 1).getRangeAddress()

            for n in range(SR, SR + nrighe):
                oCellAddress = oSheet.getCellByPosition(0, n).getCellAddress()
                oSheet.copyRange(oCellAddress, oRangeAddress)
                if elaborato == 'CONTABILITA':
                    oSheet.getCellByPosition(1, n).String = ''
                    oSheet.getCellByPosition(1, n).CellStyle = 'Comp-Bianche in mezzo_R'

            oCellRangeAddr.StartColumn = SC
            oCellRangeAddr.StartRow = SR
            oCellRangeAddr.EndColumn = EC
            oCellRangeAddr.EndRow = ER

            # INSERISCO PRIMA SOLO LE RIGHE SE NO MI FA CASINO
            SR = SR - 1
            if elaborato == 'CONTABILITA':
                oSheet.getCellByPosition(1, SR).Formula = (
                    '=DATE(' + datamis.split('/')[2] +
                    ';' + datamis.split('/')[1] + ';' +
                    datamis.split('/')[0] + ')'
                )
                oSheet.getCellByPosition(1, SR).Value = oSheet.getCellByPosition(1, SR).Value
            for mis in el.get('lista_rig'):
                # descrizione
                if mis[0] is not None:
                    descrizione = mis[0].strip()
                    oSheet.getCellByPosition(2, SR).String = descrizione
                else:
                    descrizione = ''
                # parti uguali
                if mis[3] is not None:
                    try:
                        oSheet.getCellByPosition(5, SR).Value = float(mis[3].replace(',', '.'))
                    except ValueError:
                        # tolgo evenutali '=' in eccesso
                        oSheet.getCellByPosition(5, SR).Formula = '=' + str(mis[3]).split('=')[-1]
                # lunghezza
                if mis[4] is not None:
                    try:
                        oSheet.getCellByPosition(6, SR).Value = float(mis[4].replace(',', '.'))
                    except ValueError:
                        # tolgo evenutali '=' in eccesso
                        oSheet.getCellByPosition(6, SR).Formula = '=' + str(mis[4]).split('=')[-1]
                # larghezza
                if mis[5] is not None:
                    try:
                        oSheet.getCellByPosition(7, SR).Value = float(mis[5].replace(',', '.'))
                    except ValueError:
                        # tolgo evenutali '=' in eccesso
                        oSheet.getCellByPosition(7, SR).Formula = '=' + str(mis[5]).split('=')[-1]
                # HPESO
                if mis[6] is not None:
                    try:
                        oSheet.getCellByPosition(8, SR).Value = float(mis[6].replace(',', '.'))
                    except Exception:
                        # tolgo evenutali '=' in eccesso
                        oSheet.getCellByPosition(8, SR).Formula = '=' + str(mis[6]).split('=')[-1]
                if mis[8] == '2':
                    PL.parziale_core(SR)
                    oSheet.getRows().removeByIndex(SR + 1, 1)
                    descrizione = ''
                if mis[9] != '-2':
                    vedi = diz_vv.get(mis[9])
                    try:
                        PL.vedi_voce_xpwe(SR, vedi, mis[8])
                    except Exception:
                        Dialogs.Exclamation(Title="Attenzione",
                                            Text="Il file di origine è particolarmente disordinato.\n"
                                                 "Riordinando il computo trovo riferimenti a voci "
                                                 "non ancora inserite.\n\n"
                                                 "Al termine dell'importazione controlla la voce con tariffa " +
                                                 elencoPrezzi['DizionarioArticoli'].get(ID).get('tariffa') +
                                                 "\nella riga n." + str(lrow + 2) +
                                                 " del foglio, evidenziata qui a sinistra.")
                        oSheet.getCellByPosition(44, SR).String = (
                            elencoPrezzi['DizionarioArticoli'].get(ID).get('tariffa'))
                try:
                    mis[7]
                    if '-' in mis[7]:
                        for x in range(5, 9):
                            try:
                                if oSheet.getCellByPosition(x, SR).Value != 0:
                                    oSheet.getCellByPosition(x, SR).Value = abs(oSheet.getCellByPosition(x, SR).Value)
                            except Exception:
                                pass
                        PL.inverti_un_segno(SR)
                except Exception:
                    pass
                SR = SR + 1

    PL.numera_voci()

    try:
        PL.Rinumera_TUTTI_Capitoli2()
    except Exception:
        pass


def MENU_XPWE_import():
    '''
    Importazione dati dal formato XPWE
    '''
    isLeenoDoc = LeenoUtils.isLeenoDocument()

    # legge i totali dal documento
    if isLeenoDoc:
        oDoc = LeenoUtils.getDocument()
        vals = []
        for el in ("COMPUTO", "VARIANTE", "CONTABILITA"):
            try:
                vals.append(oDoc.getSheets().getByName(el).getCellRangeByName('A2').Value)
            except Exception:
                vals.append(None)
    else:
        vals = [None, None, None]

    # sceglie il tipo di dati da importare
    ordina = LeenoConfig.Config().read('Importazione', 'ordina_computo') == '1'
    elabdest = DLG.ScegliElabDest(
        Title="Importa dal formato XPWE",
        AskTarget=isLeenoDoc,
        AskSort=True,
        Sort=ordina,
        ValComputo=vals[0],
        ValVariante=vals[1],
        ValContabilita=vals[2]
    )
    # controlla se si è annullato il comando
    if elabdest is None:
        return
    elaborato = elabdest['elaborato']
    destinazione = elabdest['destinazione']
    ordina = elabdest['ordina']
    LeenoConfig.Config().write('Importazione', 'ordina_computo', '1' if ordina else '0')

    if elaborato in ('Elenco', 'CONTABILITA'):
        ordina = False

    filename = Dialogs.FileSelect('Scegli il file XPWE da importare...', '*.xpwe')  # *.xpwe')
    if filename == 'Cancel' or filename == '':
        return

    # effettua il parsing del file XML
    tree = ElementTree()
    try:
        tree.parse(filename)
    except ParseError:
        Dialogs.Exclamation(Title="Errore nel file",
                            Text="È stato riscontrato un errore nel contenuto del file\n"
                                 "Accertati il file sia in formato XPWE.")
        return
    except PermissionError:
        Dialogs.Exclamation(Title="Errore",
                            Text="Impossibile leggere il file\n"
                                 "Accertati che il nome del file sia corretto.")
        return

    # ottieni l'item root
    root = tree.getroot()
    logging.debug(list(root))

    # attiva la progressbar
    progress = Dialogs.Progress(Title="Importazione file XPWE in corso", Text="Lettura dati")
    progress.show()

    # ########################################################################################
    # LETTURA DATI

    # va alla sezione dei dati generali
    dati = root.find('PweDatiGenerali')

    # legge i dati anagrafici generali
    datiAnagrafici = leggiAnagraficaGenerale(dati)

    # legge capitoli e categorie
    capitoliCategorie = leggiCapitoliCategorie(dati)

    # legge i dati generali per l'analisi
    datiGeneraliAnalisi = leggiDatiGeneraliAnalisi(dati)

    # legge le approssimazioni
    approssimazioni = leggiApprossimazioni(dati)

    misurazioni = root.find('PweMisurazioni')

    # legge l'elenco prezzi
    elencoPrezzi = leggiElencoPrezzi(misurazioni)

    # legge le misurazioni
    if elaborato != 'Elenco':
        listaMisure = leggiMisurazioni(misurazioni, ordina)
    else:
        listaMisure = []

    # ########################################################################################
    # SCRITTURA COMPUTO

    # se la destinazione è un nuovo documento, crealo
    if destinazione == 'NUOVO':
        PL.New_file.computo(0)
        oDoc = LeenoUtils.getDocument()

    # occorre ricreare di nuovo la progressbar, in modo che sia
    # agganciata al nuovo documento
    progress.hide();
    progress = Dialogs.Progress(Title="Importazione file XPWE in corso", Text="Scrittura computo")
    progress.show()

    # disattiva l'output a video
    LeenoUtils.DisableDocumentRefresh(oDoc)

    # compila i dati generali per l'analisi
    print("compilaDatiGeneraliAnalisi")
    progress.setText("Compilazione dati generali di analisi")
    compilaDatiGeneraliAnalisi(oDoc, datiGeneraliAnalisi)

    # compila le approssimazioni
    print("compilaApprossimazioni")
    progress.setText("Compilazione approssimazioni")
    compilaApprossimazioni(oDoc, approssimazioni)

    # compilo Anagrafica generale
    print("compilaAnagraficaGenerale")
    progress.setText("Compilazione anagrafica generale")
    compilaAnagraficaGenerale(oDoc, datiAnagrafici)

    # compilo Elenco Prezzi
    print("compilaElencoPrezzi")
    compilaElencoPrezzi(oDoc, capitoliCategorie, elencoPrezzi, progress)

    # Compilo Analisi di prezzo
    print("compilaAnalisiPrezzi")
    progress.setText("Compilazione analisi prezzi")
    compilaAnalisiPrezzi(oDoc, elencoPrezzi)

    # elimina doppioni nell'elenco prezzi
    print("EliminaVociDoppieElencoPrezzi")
    progress.setText("Eliminazione voci doppie elenco prezzi")
    PL.EliminaVociDoppieElencoPrezzi()

    # se non ci sono misurazioni di computo, finisce qui
    if len(listaMisure) == 0:
        progress.hide()

        Dialogs.Info(Title="Importazione completata",
                     Text="Importate n." +
                           str(len(elencoPrezzi['ListaArticoli'])) +
                           " voci dall'elenco prezzi\ndel file: " + filename)
        oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
        oDoc.CurrentController.setActiveSheet(oSheet)

        # riattiva l'output a video
        LeenoUtils.EnableDocumentRefresh(oDoc)
        return

    # compila il computo
    print("compilaComputo")
    progress.setText(f'Compilazione {elaborato}')
    compilaComputo(oDoc, elaborato, capitoliCategorie, elencoPrezzi, listaMisure)

    # riattiva l'output a video
    LeenoUtils.EnableDocumentRefresh(oDoc)

    PL.GotoSheet(elaborato)
    progress.setText("Adattamento altezze righe")
    PL.adatta_altezza_riga()

    progress.setText("Fine")
    progress.hide()
    Dialogs.Ok(Text='Importazione di\n\n' + elaborato + '\n\neseguita con successo!')
