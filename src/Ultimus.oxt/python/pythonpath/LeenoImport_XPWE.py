u"""
Importazione computo/variante/contabilità/prezzario
dal formato XPWE
"""
import logging
import re
import os

from xml.etree.ElementTree import ElementTree, ParseError
from com.sun.star.table import CellRangeAddress

import LeenoUtils
import LeenoFormat

import pyleeno as PL
import LeenoDialogs as DLG

import SheetUtils
import LeenoConfig
import LeenoSheetUtils
import LeenoAnalysis
import LeenoComputo
import LeenoVariante
import LeenoContab

import Dialogs


def leggiAnagraficaGenerale(dati):
    ''' legge i dati anagrafici generali '''

    datiAnagrafici = {}
    try:
        DatiGenerali = list(dati)[0][0]

        datiAnagrafici['comune'] = DatiGenerali[1].text or ''
        datiAnagrafici['oggetto'] = DatiGenerali[3].text or ''
        datiAnagrafici['committente'] = DatiGenerali[4].text or ''
        datiAnagrafici['impresa'] = DatiGenerali[5].text or ''
        '''
        datiAnagrafici['percprezzi'] = DatiGenerali[0].text or ''
        datiAnagrafici['provincia'] = DatiGenerali[2].text or ''
        datiAnagrafici['parteopera'] = DatiGenerali[6].text or ''
        '''
    except:
        datiAnagrafici['comune'] = ''
        datiAnagrafici['oggetto'] = ''
        datiAnagrafici['committente'] = ''
        datiAnagrafici['impresa'] = ''

    return datiAnagrafici


def leggiSuperCapitoli(CapCat):
    ''' legge SuperCapitoli '''

    # PweDGSuperCapitoli
    listaSuperCapitoli = []
    if CapCat.find('PweDGSuperCapitoli'):
        PweDGSuperCapitoli = list(CapCat.find('PweDGSuperCapitoli'))
        for elem in PweDGSuperCapitoli:
            id_sc = elem.get('ID')
            #~ codice = elem.find('Codice').text
            try:
                codice = elem.find('Codice').text
            except AttributeError:
                codice = ''
            dessintetica = elem.find('DesSintetica').text
            try:
                percentuale = elem.find('Percentuale').text
            except:
                percentuale = ''
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
        PweDGCapitoli = list(CapCat.find('PweDGCapitoli'))
        for elem in PweDGCapitoli:
            id_sc = elem.get('ID')
            #~ codice = elem.find('Codice').text
            try:
                codice = elem.find('Codice').text
            except AttributeError:
                codice = ''
            dessintetica = elem.find('DesSintetica').text
            if dessintetica == "Nuova voce":
                dessintetica = elem.find('DesEstesa').text
            try:
                percentuale = elem.find('Percentuale').text
            except:
                percentuale = ''
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
        PweDGSubCapitoli = list(CapCat.find('PweDGSubCapitoli'))
        for elem in PweDGSubCapitoli:
            id_sc = elem.get('ID')
            try:
                codice = elem.find('Codice').text
            except AttributeError:
                codice = ''
            try:
                codice = elem.find('Codice').text
            except AttributeError:
                codice = ''
            dessintetica = elem.find('DesSintetica').text
            if dessintetica == "Nuova voce":
                dessintetica = elem.find('DesEstesa').text
            try:
                percentuale = elem.find('Percentuale').text
            except:
                percentuale = ''
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
        PweDGSuperCategorie = list(CapCat.find('PweDGSuperCategorie'))
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
        PweDGCategorie = list(CapCat.find('PweDGCategorie'))
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
        PweDGSubCategorie = list(CapCat.find('PweDGSubCategorie'))
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

    try:
        PweDGAnalisi = list(dati.find('PweDGModuli'))[0]

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
    # ~except AttributeError:
    except:
        speseGenerali = 0
        utiliImpresa = 0
        oneriAccessoriSicurezza = 0

    return {
        'SpeseGenerali': speseGenerali,
        'UtiliImpresa': utiliImpresa,
        'OneriAccessoriSicurezza': oneriAccessoriSicurezza
    }


def leggiApprossimazioni(dati):
    ''' legge le impostazioni di approssimazione numerica '''

    try:
        PweDGConfigNumeri = dati.find('PweDGConfigurazione')
    except AttributeError:
        PweDGConfigNumeri = None

    if PweDGConfigNumeri is None:
        return {}
    PweDGConfigNumeri = list(PweDGConfigNumeri)[0]
    res = {}

    partiUguali = PweDGConfigNumeri.find('PartiUguali')
    try:
        res['PartiUguali'] = int(partiUguali.text.split('.')[-1].split('|')[0])
    except:
        pass
    larghezza = PweDGConfigNumeri.find('Larghezza')

    try:
        res['Larghezza'] = int(larghezza.text.split('.')[-1].split('|')[0])
    except:
        pass

    lunghezza = PweDGConfigNumeri.find('Lunghezza')
    try:
        res['Lunghezza'] = int(lunghezza.text.split('.')[-1].split('|')[0])
    except:
        pass

    hPeso = PweDGConfigNumeri.find('HPeso')
    try:
        res['HPeso'] = int(hPeso.text.split('.')[-1].split('|')[0])
    except:
        pass

    quantita = PweDGConfigNumeri.find('Quantita')
    try:
        res['Quantita'] = int(quantita.text.split('.')[-1].split('|')[0])
    except:
        pass

    prezzi = PweDGConfigNumeri.find('Prezzi')
    try:
        res['Prezzi'] = int(prezzi.text.split('.')[-1].split('|')[0])
    except:
        pass

    prezziTotale = PweDGConfigNumeri.find('PrezziTotale')
    try:
        res['PrezziTotale'] = int(prezziTotale.text.split('.')[-1].split('|')[0])
    except:
        pass
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
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a PU ROSSO', approssimazioni['PartiUguali'])
    if 'Lunghezza' in approssimazioni:
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a LUNG', approssimazioni['Lunghezza'])
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a LUNG ROSSO', approssimazioni['Lunghezza'])
    if 'Larghezza' in approssimazioni:
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a LARG', approssimazioni['Larghezza'])
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a LARG ROSSO', approssimazioni['Larghezza'])
    if 'HPeso' in approssimazioni:
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a peso', approssimazioni['HPeso'])
        LeenoFormat.setCellStyleDecimalPlaces('comp 1-a peso ROSSO', approssimazioni['HPeso'])
    if 'Quantita' in approssimazioni:
        for el in ('Comp-Variante num sotto', 'Comp-Variante num sotto ROSSO', 'An-lavoraz-input', 'Blu', 'Blu ROSSO'):
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

    PweElencoPrezzi = list(misurazioni)[0]

    # leggo l'elenco prezzi
    epitems = PweElencoPrezzi.findall('EPItem')

    for elem in epitems:
        id_ep = elem.get('ID')
        dizionarioArticolo = {}
        try:
            tipoep = elem.find('TipoEP').text
        except:
            tipoep = '0'
        tariffa = elem.find('Tariffa').text or ''
        # Voce Della Sicurezza
        if elem.find('Flags').text == '134217728':
            tariffa = "VDS_" + tariffa
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
        try:
            prezzo2 = elem.find('Prezzo2').text
        except:
            prezzo2 = '0'
        try:
            prezzo3 = elem.find('Prezzo3').text
        except:
            prezzo3 = '0'
        try:
            prezzo4 = elem.find('Prezzo4').text
        except:
            prezzo4 = '0'
        try:
            prezzo5 = elem.find('Prezzo5').text
        except:
            prezzo5 = '0'
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
        try:
            pweepanalisi = elem.find('PweEPAnalisi').text
        except:
            pweepanalisi = ''

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
        # try:
        #     spese = int(pweepanalisi.find('Spese').text)
        # except:
        #     spese = None

        try:
            PweEPAR = pweepanalisi.find('PweEPAR')
        except:
            PweEPAR = None
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
                except:
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


################################################
################################################
################################################

def leggiMisurazioni(misurazioni, ordina):
    """leggo voci di misurazione e righe"""

    def text_or_empty(elem, tag):
        nodo = elem.find(tag)
        return nodo.text if nodo is not None else ''

    def clean_value(elem, tag, pattern):
        value = text_or_empty(elem, tag)

        if not isinstance(value, (str, bytes)):
            return value

        # applica sostituzione
        try:
            value = re.sub(pattern, r'\1*\2', value)
        except Exception:
            pass

        # conversione None secondo logica originale
        if value is not None:
            if '  ' in value or value == '0.00':
                return None

        return value

    if not misurazioni:
        return

    listaMisure = []

    try:
        items = list(misurazioni)
        PweVociComputo = items[1]  # mantengo la logica originale
        vcitems = PweVociComputo.findall('VCItem')

        prova_l = []
        pattern = r'(\d+)(\()'

        for elem in vcitems:
            diz_misura = {}
            id_vc = elem.get('ID')
            id_ep = text_or_empty(elem, 'IDEP')
            quantita_voce = text_or_empty(elem, 'Quantita')
            datamis = text_or_empty(elem, 'DataMis')
            flags_voce = text_or_empty(elem, 'Flags')
            idspcat = text_or_empty(elem, 'IDSpCat')
            idcat = text_or_empty(elem, 'IDCat')
            idsbcat = text_or_empty(elem, 'IDSbCat')

            # ultime righe
            righi_mis = list(elem)[-1].findall('RGItem')
            lista_righe = []

            for el in righi_mis:
                idvv = text_or_empty(el, 'IDVV')
                descrizione = text_or_empty(el, 'Descrizione')

                # pulizia campi numerici
                partiuguali = clean_value(el, 'PartiUguali', pattern)
                lunghezza = clean_value(el, 'Lunghezza', pattern)
                larghezza = clean_value(el, 'Larghezza', pattern)
                hpeso = clean_value(el, 'HPeso', pattern)

                quantita_riga = text_or_empty(el, 'Quantita')
                flags_riga = text_or_empty(el, 'Flags')

                riga_misura = (
                    descrizione,
                    '',
                    '',
                    partiuguali,
                    lunghezza,
                    larghezza,
                    hpeso,
                    quantita_riga,
                    flags_riga,
                    idvv,
                )

                lista_righe.append(riga_misura)

            # popolazione dizionario misura
            diz_misura['id_vc'] = id_vc
            diz_misura['id_ep'] = id_ep
            diz_misura['quantita'] = quantita_voce
            diz_misura['datamis'] = datamis
            diz_misura['flags'] = flags_voce
            diz_misura['idspcat'] = idspcat
            diz_misura['idcat'] = idcat
            diz_misura['idsbcat'] = idsbcat
            diz_misura['lista_rig'] = lista_righe

            new_id = (
                PL.strall(idspcat)
                + '.'
                + PL.strall(idcat)
                + '.'
                + PL.strall(idsbcat)
            )

            prova_l.append((new_id, diz_misura))
            listaMisure.append(diz_misura)

        # ordinamento (logica invariata)
        if len(listaMisure) != 0 and ordina:
            riordinate = sorted(prova_l, key=lambda el: el[0])
            listaMisure = [el[1] for el in riordinate]

    except IndexError:
        Dialogs.Exclamation(
            Title="Attenzione",
            Text="Nel file scelto non risultano esserci voci di misurazione,\n"
                 "perciò saranno importate le sole voci di Elenco Prezzi.\n\n"
                 "Si tenga conto che:\n"
                 "- sarà importato solo il 'Prezzo 1' dell'elenco;\n"
                 "- a seconda della versione, il formato XPWE potrebbe\n"
                 "  non conservare alcuni dati come le incidenze di\n"
                 "  sicurezza e di manodopera!"
        )

    return listaMisure

################################################
################################################
################################################

def estraiDatiCapitoliCategorie(capitoliCategorie, catName):
    resList = []
    for el in capitoliCategorie[catName]:
        tariffa = el.get('codice')
        if tariffa is not None:
            destestesa = el.get('dessintetica')
            titolo = (tariffa, destestesa, '', '', '', '', '')
            resList.append(titolo)
    return tuple(resList)


def riempiBloccoElencoPrezzi(oSheet, dati, col, progress=None, case_sensitive=False):
    # 1. Recupera tutti i codici esistenti nel foglio (colonna 0, dalla riga 3 in poi)
    existing_codes = set()
    max_row = oSheet.getRows().getCount()

    if max_row > 3:
        # Legge i codici in batch per ottimizzazione (evita timeout su fogli grandi)
        chunk_size = 1000  # Adjust based on performance
        for start_row in range(3, max_row, chunk_size):
            end_row = min(start_row + chunk_size, max_row)
            codici_esistenti = oSheet.getCellRangeByPosition(0, start_row, 0, end_row - 1).getDataArray()
            for codice in codici_esistenti:
                if codice and codice[0]:
                    code = str(codice[0]).strip()
                    if not case_sensitive:
                        code = code.lower()
                    existing_codes.add(code)

    # 2. Filtra i nuovi dati: rimuove duplicati interni + codici già esistenti
    nuovi_dati = []
    seen_new_codes = set()

    for riga in dati:
        if not riga or not riga[0]:  # Skip righe vuote
            continue

        codice = str(riga[0]).strip()
        if not case_sensitive:
            codice = codice.lower()

        # Controlla sia nei codici esistenti che nei nuovi già processati
        if codice not in existing_codes and codice not in seen_new_codes:
            nuovi_dati.append(riga)
            seen_new_codes.add(codice)

    if not nuovi_dati:
        return

    # 3. Inserimento dati (ottimizzato per grandi blocchi)
    righe_totali = len(nuovi_dati)
    colonne = len(nuovi_dati[0])
    sRow = 4
    oSheet.getRows().insertByIndex(sRow, righe_totali)

    # Inserimento a step (es. 100 righe alla volta)
    step = 100
    riga = 0
    while riga < righe_totali:
        sliced = nuovi_dati[riga:riga + step]
        num = len(sliced)
        oRange = oSheet.getCellRangeByPosition(0, sRow + riga, colonne - 1, sRow + riga + num - 1)
        oRange.setDataArray(sliced)
        # PL.stileCelleElencoPrezzi(oSheet, sRow + riga, sRow + riga + num - 1, col)
        riga += num

def compilaElencoPrezzi(oDoc, capitoliCategorie, elencoPrezzi, progress = None):
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
    # progress.setLimits(0, righeTotali)
    # progress.setValue(0)

    # compilo Elenco Prezzi
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')

    riempiBloccoElencoPrezzi(oSheet, arrayArticoli, None, progress = None)
    '''
    aggiungo i capitoli alla lista delle voci
     giallo(16777072,16777120,16777168)
     verde(9502608,13696976,15794160)
     viola(12632319,13684991,15790335)
    SUPERCAPITOLI
    '''
    # SuperCapitoli
    if righeSuperCapitoli:
        riempiBloccoElencoPrezzi(oSheet, arraySuperCapitoli, 16777072, progress = None)

    # Capitoli
    if righeCapitoli:
        riempiBloccoElencoPrezzi(oSheet, arrayCapitoli, 16777120, progress = None)

    # SottoCapitoli
    if righeSottoCapitoli:
        riempiBloccoElencoPrezzi(oSheet, arraySottoCapitoli, 16777168, progress = None)

    PL.riordina_ElencoPrezzi()


def compilaAnalisiPrezzi(oDoc, elencoPrezzi, progress):
    ''' Compilo Analisi di prezzo '''
    numAnalisi = len(elencoPrezzi['ListaAnalisi'])
    if numAnalisi == 0:
        return
    # inizializza la progressbar
    indicator = oDoc.getCurrentController().getStatusIndicator()
    indicator.start("Elaborazione Analisi dei Prezzi in corso...", numAnalisi)

    # Ottieni tutti i codici già presenti nel foglio (colonna 0)
    oSheet, startRow = LeenoAnalysis.inizializzaAnalisi(oDoc)
    existing_codes = set()
    max_row = oSheet.getRows().getCount()

    if max_row > 0:
        # Legge i codici esistenti (prima colonna)
        codici_esistenti = oSheet.getCellRangeByPosition(0, 0, 0, max_row - 1).getDataArray()
        existing_codes = {str(codice[0]).strip().lower() for codice in codici_esistenti if codice and codice[0]}

    val = 0
    skipped = 0

    # inizializza l'analisi dei prezzi
    oSheet, startRow = LeenoAnalysis.inizializzaAnalisi(oDoc)

    # compila le voci dell'analisi
    for el in elencoPrezzi['ListaAnalisi']:
        codice = str(el[0]).strip().lower()

        # Skip se il codice esiste già
        if codice in existing_codes:
            skipped += 1
            val += 1
            indicator.Value = val
            continue

        prezzo_finale = el[-1]

        # circoscrive la voce di analisi corrente
        sStRange = LeenoAnalysis.circoscriveAnalisi(oSheet, startRow)

        lrow = sStRange.RangeAddress.StartRow + 2
        oSheet.getCellByPosition(0, lrow).String = el[0]
        oSheet.getCellByPosition(1, lrow).String = el[1]
        oSheet.getCellByPosition(2, lrow).String = el[2]
        y = 0
        n = lrow + 2
        for x in el[3]:
            if el[3][y][1] in (
                'MANODOPERA', 'MATERIALI', 'NOLI', 'TRASPORTI',
                'ALTRE FORNITURE E PRESTAZIONI', 'overflow'):
                if el[3][y][1] != 'overflow':
                    n = SheetUtils.uFindStringCol(el[3][y][1], 1, oSheet, lrow)
            else:
                LeenoAnalysis.copiaRigaAnalisi(oSheet, n)
                if elencoPrezzi['DizionarioArticoli'].get(el[3][y][0]) is not None:
                    oSheet.getCellByPosition(0, n).String = (
                        elencoPrezzi['DizionarioArticoli'].get(el[3][y][0]).get('tariffa'))
                # per gli inserimenti liberi (L)
                else:
                    oSheet.getCellByPosition(0, n).String = ''
                    try:
                        oSheet.getCellByPosition(1, n).String = x[1]
                    except:
                        oSheet.getCellByPosition(1, n).String = ''
                    oSheet.getCellByPosition(2, n).String = x[2]
                    try:
                        float(x[3].replace(',', '.'))
                        oSheet.getCellByPosition(3, n).Value = float(x[3].replace(',', '.'))
                    except Exception:
                        oSheet.getCellByPosition(3, n).Value = 0
                    try:
                        float(x[4].replace(',', '.'))
                        oSheet.getCellByPosition(4, n).Value = float(x[4].replace(',', '.'))
                    except:
                        oSheet.getCellByPosition(4, n).Value = 0
                if el[3][y][1] not in (
                    'MANODOPERA', 'MATERIALI', 'NOLI', 'TRASPORTI',
                    'ALTRE FORNITURE E PRESTAZIONI', 'overflow'):
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
        startRow = sStRange.RangeAddress.StartRow
        endRow = sStRange.RangeAddress.EndRow
        for m in reversed(range(startRow, endRow)):
            if(oSheet.getCellByPosition(0, m).String == 'Cod. Art.?' and
                oSheet.getCellByPosition(0, m - 1).CellStyle == 'An-lavoraz-Cod-sx'):
                oSheet.getRows().removeByIndex(m, 1)
            if oSheet.getCellByPosition(0, m).String == 'Cod. Art.?':
                oSheet.getCellByPosition(0, m).String = ''
        # ~ if oSheet.getCellByPosition(6, startRow + 2).Value != prezzo_finale:
            # ~ oSheet.getCellByPosition(6, startRow + 2).Value = prezzo_finale
        oSheet, startRow = LeenoAnalysis.inizializzaAnalisi(oDoc)

        # aggiorna la progressbar
        val += 1
        indicator.Value = val
    indicator.end()

    # siccome viene inserita una voce PRIMA di iniziare la compilazione
    # occorre eliminare l'ultima voce che risulta vuota
    LeenoSheetUtils.eliminaVoce(oSheet, LeenoSheetUtils.cercaUltimaVoce(oSheet))
    PL.tante_analisi_in_ep()


def compilaAnalisiPrezzi_(oDoc, elencoPrezzi, progress):
    ''' Compila Analisi di prezzo, saltando voci già esistenti '''
    numAnalisi = len(elencoPrezzi['ListaAnalisi'])
    if numAnalisi == 0:
        return

    # Inizializza progressbar
    indicator = oDoc.getCurrentController().getStatusIndicator()
    indicator.start("Elaborazione Analisi dei Prezzi in corso...", numAnalisi)

    # Ottieni tutti i codici già presenti nel foglio (colonna 0)
    oSheet, startRow = LeenoAnalysis.inizializzaAnalisi(oDoc)
    existing_codes = set()
    max_row = oSheet.getRows().getCount()

    if max_row > 0:
        # Legge i codici esistenti (prima colonna)
        codici_esistenti = oSheet.getCellRangeByPosition(0, 0, 0, max_row - 1).getDataArray()
        existing_codes = {str(codice[0]).strip().lower() for codice in codici_esistenti if codice and codice[0]}

    val = 0
    skipped = 0

    for el in elencoPrezzi['ListaAnalisi']:
        codice = str(el[0]).strip().lower()

        # Skip se il codice esiste già
        if codice in existing_codes:
            # logger.info(f"Saltata analisi '{el[1]}' (codice '{el[0]}' già presente)")
            skipped += 1
            val += 1
            indicator.Value = val
            continue

        prezzo_finale = el[-1]
        sStRange = LeenoAnalysis.circoscriveAnalisi(oSheet, startRow)
        lrow = sStRange.RangeAddress.StartRow + 2

        # Compila i dati (come nel tuo codice originale)
        oSheet.getCellByPosition(0, lrow).String = el[0]
        oSheet.getCellByPosition(1, lrow).String = el[1]
        oSheet.getCellByPosition(2, lrow).String = el[2]

        y = 0
        n = lrow + 2
        for x in el[3]:
            if el[3][y][1] in ('MANODOPERA', 'MATERIALI', 'NOLI', 'TRASPORTI', 'ALTRE FORNITURE E PRESTAZIONI', 'overflow'):
                if el[3][y][1] != 'overflow':
                    n = SheetUtils.uFindStringCol(el[3][y][1], 1, oSheet, lrow)
            else:
                LeenoAnalysis.copiaRigaAnalisi(oSheet, n)
                if elencoPrezzi['DizionarioArticoli'].get(el[3][y][0]) is not None:
                    oSheet.getCellByPosition(0, n).String = elencoPrezzi['DizionarioArticoli'].get(el[3][y][0]).get('tariffa')
                else:
                    oSheet.getCellByPosition(0, n).String = ''
                    try:
                        oSheet.getCellByPosition(1, n).String = x[1]
                    except:
                        oSheet.getCellByPosition(1, n).String = ''
                    oSheet.getCellByPosition(2, n).String = x[2]
                    try:
                        float(x[3].replace(',', '.'))
                        oSheet.getCellByPosition(3, n).Value = float(x[3].replace(',', '.'))
                    except Exception:
                        oSheet.getCellByPosition(3, n).Value = 0
                    try:
                        float(x[4].replace(',', '.'))
                        oSheet.getCellByPosition(4, n).Value = float(x[4].replace(',', '.'))
                    except:
                        oSheet.getCellByPosition(4, n).Value = 0
            y += 1
            n += 1

        # Pulizia righe vuote (come nel tuo codice originale)
        sStRange = LeenoAnalysis.circoscriveAnalisi(oSheet, lrow)
        startRow = sStRange.RangeAddress.StartRow
        endRow = sStRange.RangeAddress.EndRow
        for m in reversed(range(startRow, endRow)):
            if(oSheet.getCellByPosition(0, m).String == 'Cod. Art.?' and
               oSheet.getCellByPosition(0, m - 1).CellStyle == 'An-lavoraz-Cod-sx'):
                oSheet.getRows().removeByIndex(m, 1)
            if oSheet.getCellByPosition(0, m).String == 'Cod. Art.?':
                oSheet.getCellByPosition(0, m).String = ''

        oSheet, startRow = LeenoAnalysis.inizializzaAnalisi(oDoc)
        val += 1
        indicator.Value = val

    # Elimina l'ultima voce vuota (come nel tuo codice originale)
    LeenoSheetUtils.eliminaVoce(oSheet, LeenoSheetUtils.cercaUltimaVoce(oSheet))
    indicator.end()

    # Dialogs.Info(
    #     Title="Elaborazione completata",
    #     Text=f"Analisi di prezzo elaborate: {numAnalisi - skipped}\n"
    #          f"Analisi saltate: {skipped}\n"
    #          "Controlla il foglio 'Analisi Prezzi' per i risultati."
    # )

    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    PL.tante_analisi_in_ep()


def compilaComputo(oDoc, elaborato, capitoliCategorie, elencoPrezzi, listaMisure):
    ''' compila il computo '''
    from datetime import datetime, date

    LeenoUtils.DocumentRefresh(False)

    # --- Creazione/attivazione del foglio ---
    if elaborato == 'VARIANTE':
        if oDoc.getSheets().hasByName('VARIANTE'):
            oSheet = LeenoVariante.generaVariante(oDoc, False)
        else:
            oSheet = LeenoVariante.generaVariante(oDoc, True)
            oSheet.getRows().removeByIndex(2, 4)
    elif elaborato == 'CONTABILITA':
        oSheet = LeenoContab.generaContabilita(oDoc)
    else:
        oSheet = oDoc.getSheets().getByName(elaborato)

    # --- Rimozione della riga vuota iniziale ---
    if oSheet.getCellByPosition(1, 4).String == 'Cod. Art.?':
        oSheet.getRows().removeByIndex(3, 5 if elaborato == 'CONTABILITA' else 4)

    # --- Setup cell range ---
    oCellRangeAddr = CellRangeAddress()
    oCellRangeAddr.Sheet = oSheet.RangeAddress.Sheet

    # --- Mappe e variabili di controllo ---
    mappaVociRighe = {}        # id voce computo → riga foglio
    numeroVoce = 1             # numerazione progressiva voci

    testspcat = '0'            # per evitare duplicati supercategoria
    testcat = '0'              # per evitare duplicati categoria
    testsbcat = '0'            # per evitare duplicati sottocategoria

    # --- Progress bar ---
    indicator = oDoc.getCurrentController().getStatusIndicator()
    indicator.start(f'Compilazione {elaborato}...', len(listaMisure))

    # -------------------------------------------------------------------------
    # Funzioni interne di utilità (NON alterano la logica)
    # -------------------------------------------------------------------------

    def insert_categoria_if_needed(idcorr, idtest, reset_test, inser_func, dict_list):
        """Inserisce capitolo/sottocapitolo solo se cambia l'id."""
        nonlocal lrow
        try:
            if idcorr != idtest:
                reset_test[0] = idcorr
                inser_func(oSheet, lrow, dict_list[eval(idcorr) - 1][1])
                lrow += 1
        except UnboundLocalError:
            pass

    def set_num_or_formula(col, row, value):
        """Scrive un numero o formula mantenendo la logica originale."""
        cell = oSheet.getCellByPosition(col, row)
        if value is None:
            return
        try:
            cell.Value = float(value.replace(',', '.'))
        except Exception:
            cell.Formula = ('=' + str(value).strip()).replace('=-', '=')

    def handle_negatives(startRow):
        """Gestione del segno negativo (logica invariata)."""
        try:
            if '-' in mis[7]:
                for x in range(5, 9):
                    try:
                        if oSheet.getCellByPosition(x, startRow).Value != 0:
                            val = abs(oSheet.getCellByPosition(x, startRow).Value)
                            oSheet.getCellByPosition(x, startRow).Value = val
                    except Exception:
                        pass
                LeenoSheetUtils.invertiUnSegno(oSheet, startRow)

                if elaborato == 'CONTABILITA':
                    if test == '-':
                        LeenoSheetUtils.invertiUnSegno(oSheet, startRow)
        except Exception:
            pass

    # -------------------------------------------------------------------------
    # CICLO PRINCIPALE
    # -------------------------------------------------------------------------

    val = 0
    for el in listaMisure:
        indicator.Value = val
        val += 1

        datamis = el.get('datamis')
        idspcat = el.get('idspcat')
        idcat = el.get('idcat')
        idsbcat = el.get('idsbcat')

        # trova la prima riga libera
        lrow = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1

        # --- Supercategoria ---
        insert_categoria_if_needed(
            idspcat, testspcat, [idspcat, None],
            LeenoSheetUtils.inserSuperCapitolo,
            capitoliCategorie['SuperCategorie']
        )
        testspcat = idspcat

        # --- Categoria ---
        insert_categoria_if_needed(
            idcat, testcat, [idcat, None],
            LeenoSheetUtils.inserCapitolo,
            capitoliCategorie['Categorie']
        )
        testcat = idcat

        # --- Sottocategoria ---
        insert_categoria_if_needed(
            idsbcat, testsbcat, [idsbcat, None],
            LeenoSheetUtils.inserSottoCapitolo,
            capitoliCategorie['SottoCategorie']
        )
        testsbcat = idsbcat

        # --- Inserimento voce ---
        if elaborato == 'CONTABILITA':
            LeenoContab.insertVoceContabilita(oSheet, lrow)
        else:
            LeenoComputo.insertVoceComputoGrezza(oSheet, lrow)

        ID = el.get('id_ep')

        # --- Inserisce tariffa ---
        try:
            tariffa = elencoPrezzi['DizionarioArticoli'].get(ID).get('tariffa')
            oSheet.getCellByPosition(1, lrow + 1).String = tariffa
        except Exception:
            pass

        # --- Mappa id voce XPWE → riga foglio ---
        idVoceComputo = el.get('id_vc')
        mappaVociRighe[idVoceComputo] = lrow + 1

        # --- Numerazione voce ---
        oSheet.getCellByPosition(0, lrow + 1).String = str(numeroVoce)
        numeroVoce += 1

        # --- Misurazioni ---
        startRow = lrow + 2
        lista_righe = el.get('lista_rig')
        nrighe = len(lista_righe)

        if nrighe > 0:
            endRow = startRow + nrighe

            # inserisce righe aggiuntive
            if nrighe > 1:
                oSheet.getRows().insertByIndex(startRow + 1, nrighe - 1)

            # copia layout prima riga
            oRangeAddress = oSheet.getCellRangeByPosition(0, startRow, 250, startRow).getRangeAddress()
            for n in range(startRow + 1, endRow):
                oSheet.copyRange(oSheet.getCellByPosition(0, n).getCellAddress(), oRangeAddress)
                if elaborato == 'CONTABILITA':
                    c = oSheet.getCellByPosition(1, n)
                    c.String = ''
                    c.CellStyle = 'Comp-Bianche in mezzo_R'

            # --- Data contabilita (logica originale) ---
            if elaborato == 'CONTABILITA':
                cdata = oSheet.getCellByPosition(1, startRow)
                cdata.Formula = (
                    '=DATE(' + datamis.split('/')[2] +
                    ';' + datamis.split('/')[1] + ';' +
                    ';' + datamis.split('/')[0] + ')'
                )
                cdata.Value = cdata.Value

            # --- Popola righe ---
            for mis in lista_righe:
                descrizione = (mis[0].strip() if mis[0] else '')
                oSheet.getCellByPosition(2, startRow).String = descrizione

                set_num_or_formula(5, startRow, mis[3])  # parti uguali
                set_num_or_formula(6, startRow, mis[4])  # lunghezza
                set_num_or_formula(7, startRow, mis[5])  # larghezza
                set_num_or_formula(8, startRow, mis[6])  # HPESO

                # riga parziale
                if mis[8] == '2':
                    PL.parziale_core(oSheet, startRow)
                    if elaborato != 'CONTABILITA':
                        oSheet.getRows().removeByIndex(startRow + 1, 1)
                    descrizione = ''

                # VEDI VOCE
                test = ''
                if mis[9] != '-2':
                    vedi = mappaVociRighe.get(mis[9])
                    try:
                        test = PL.vedi_voce_xpwe(oSheet, startRow, vedi)
                    except Exception:
                        Dialogs.Exclamation(
                            Title="Attenzione",
                            Text="Il file di origine è disordinato.\n"
                                 "Riordinando il computo trovo riferimenti a voci "
                                 "non ancora inserite.\n\n"
                                 "Controlla la voce con tariffa "
                                 + elencoPrezzi['DizionarioArticoli'].get(ID).get('tariffa')
                                 + "\nalla riga n." + str(lrow + 2)
                        )
                        oSheet.getCellByPosition(44, startRow).String = elencoPrezzi['DizionarioArticoli'].get(ID).get('tariffa')

                # segni negativi
                handle_negatives(startRow)

                startRow += 1

    # -------------------------------------------------------------------------
    # Finalizzazione
    # -------------------------------------------------------------------------

    indicator.end()
    LeenoSheetUtils.numeraVoci(oSheet, 0, True)

    try:
        PL.Rinumera_TUTTI_Capitoli2(oSheet)
    except Exception:
        pass

    PL.fissa()

def MENU_XPWE_import(filename = None):
    # with LeenoUtils.DocumentRefreshContext(False):
    XPWE_import(filename = None)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)

@LeenoUtils.no_refresh
def XPWE_import(filename = None):
    '''
    Importazione dati dal formato XPWE
    '''
    oDoc = LeenoUtils.getDocument()
    isLeenoDoc = LeenoUtils.isLeenoDocument()
    if isLeenoDoc == False:
        PL.creaComputo(0)
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

    if filename == None:
        filename = Dialogs.FileSelect('Scegli il file XPWE da importare...', '*.xpwe')  # *.xpwe')
    if filename in ('Cancel', '', None):
        return

    # effettua il parsing del file XML
    tree = ElementTree()
    # DLG.chi(tree.parse(filename))
    try:
        tree.parse(filename)
    except ParseError:
        PL.clean_text_file(filename)
        tree.parse(filename)
    except PermissionError:
        Dialogs.Exclamation(Title="Errore",
                            Text="Impossibile leggere il file\n"
                                 "Accertati che il nome del file sia corretto.")
    except Exception as e:
        Dialogs.Exclamation(Title="Errore",
                            Text="Errore generico nella lettura del file\n"
                                 "Accertati che il file sia in formato XPWE.")
        DLG.errore(str(e))
        return

    # ottieni l'item root
    root = tree.getroot()
    logging.debug(list(root))

    # attiva la progressbar
    indicator = oDoc.getCurrentController().getStatusIndicator()
    if indicator:
        indicator.start("Elaborazione in corso...", 100)  # 100 = max progresso


    if indicator:
        indicator.Text = "Importazione file XPWE in corso..."
        indicator.Value = 20

    # ########################################################################################
    # LETTURA DATI

    # va alla sezione dei dati generali
    dati = root.find('PweDatiGenerali')
    if dati == None:
        dati = list(root)[0].find('PweDatiGenerali')

    # legge i dati anagrafici generali
    datiAnagrafici = leggiAnagraficaGenerale(dati)
    if indicator:
        indicator.Text = "Lettura dati..."
        indicator.Value = 35

    # legge capitoli e categorie
    capitoliCategorie = leggiCapitoliCategorie(dati)
    if indicator:
        indicator.Text = "Lettura Capitoli e Categorie..."
        indicator.Value = 50

    # legge i dati generali per l'analisi
    datiGeneraliAnalisi = leggiDatiGeneraliAnalisi(dati)
    if indicator:
        indicator.Text = "Lettura dati generali per analisi..."
        indicator.Value = 65

    # legge le approssimazioni
    approssimazioni = leggiApprossimazioni(dati)

    misurazioni = root.find('PweMisurazioni')
    if misurazioni == None:
        misurazioni = list(root)[0].find('PweMisurazioni')

    # legge l'elenco prezzi
    try:
        elencoPrezzi = leggiElencoPrezzi(misurazioni)
    except Exception as e:
        # ~ DLG.chi(f"Errore: {e}")
        return
    if indicator:
        indicator.Text = "Lettura Elenco Prezzi..."
        indicator.Value = 70

    # legge le misurazioni
    if elaborato != 'Elenco':
        listaMisure = leggiMisurazioni(misurazioni, ordina)
    else:
        listaMisure = []
    if indicator:
        indicator.Text = "Lettura Misurazioni..."
        indicator.Value = 100

    # ########################################################################################
    # SCRITTURA COMPUTO

    # se la destinazione è un nuovo documento, crealo
    if destinazione == 'NUOVO':
        oDoc = PL.creaComputo(0)

    # occorre ricreare di nuovo la progressbar, in modo che sia
    # agganciata al nuovo documento

    if indicator:
        indicator.Text = "Importazione file XPWE in corso..."
        indicator.Value = 15

    # disattiva l'output a video
    # LeenoUtils.DocumentRefresh(False)

    # compila i dati generali per l'analisi
    if indicator:
        indicator.Text = "Compilazione dati generali di analisi..."
        indicator.Value = 35
    compilaDatiGeneraliAnalisi(oDoc, datiGeneraliAnalisi)

    # compila le approssimazioni
    if indicator:
        indicator.Text = "Compilazione approssimazioni..."
        indicator.Value = 50
    compilaApprossimazioni(oDoc, approssimazioni)

    # compilo Anagrafica generale
    indicator.Text = "Compilazione anagrafica generale..."
    indicator.Value = 60

    compilaAnagraficaGenerale(oDoc, datiAnagrafici)

    # compilo Elenco Prezzi
    indicator.Text = "Compilazione anagrafica generale..."
    indicator.Value = 85
    compilaElencoPrezzi(oDoc, capitoliCategorie, elencoPrezzi, progress = None)
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oSheet.getCellRangeByName('E2').Formula = '=COUNT(E:E) & " prezzi"'

    # Compilo Analisi di prezzo
    indicator.Text = "Compilazione analisi prezzi..."
    indicator.Value = 85
    compilaAnalisiPrezzi(oDoc, elencoPrezzi, progress = None)

    # se non ci sono misurazioni di computo, finisce qui
    if len(listaMisure) == 0:
    #     progress.hide()

        Dialogs.Info(Title="Importazione completata",
                     Text="Importate n." +
                           str(len(elencoPrezzi['ListaArticoli'])) +
                           " voci dall'elenco prezzi\ndel file: " + filename)
        oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
        oDoc.CurrentController.setActiveSheet(oSheet)

        return
    # compila il computo
    compilaComputo(oDoc, elaborato, capitoliCategorie, elencoPrezzi, listaMisure)

    oSheet = oDoc.getSheets().getByName(elaborato)

    PL.GotoSheet(elaborato)

    indicator.end()
    PL.Rinumera_TUTTI_Capitoli2(oSheet)

    # salva il file col nome del file di origine
    if len(oDoc.getURL()) == 0:
        dest = filename[0:-5]+ '.ods'
        PL.salva_come(dest)

    # LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    Dialogs.Ok(Text=f'Importazione di {elaborato} eseguita con successo!')
    if 'giuserpe' not in os.getlogin():
        PL.dlg_donazioni()
