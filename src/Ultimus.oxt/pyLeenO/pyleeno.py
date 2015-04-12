#!/usr/bin/env python -c
# -*- coding: utf-8 -*-
########################################################################
import os, sys, uno, unohelper
# cos'e' il namespace:
# http://www.html.it/articoli/il-misterioso-mondo-dei-namespaces-1/
import xml.etree.ElementTree as etree
#xmL = 'W:/_dwg/ULTIMUSFREE/elenchi/Bolzano/altri_formati/HBED13.six---.xml'# file da convertire
xmL = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/Bolzano/2014/HBED14_OpereEdili_20141103_XmlSix.xml'
# fixtag serve per risalire al dict che contiene il namespace del file XML-six
def fixtag(ns, tag, nsmap):
    return '{' + nsmap[ns] + '}' + tag
########################################################################
def xmlsix2ods ():
    """
    Modulo per la conversione dei prezzari dal formato XML-six a LeenO
    (c) Giuseppe Vizziello 2015
    """
    lang = 1 # 0 tedesco - 1 italiano
    nsmap = {}
    listaUM = []
    listaprzDes = []
    listaPRO = []
    listaSO = []
    for event, elem in etree.iterparse(xmL, events=('end', 'start-ns')):
        if event == 'start-ns':
            ns, url = elem
            nsmap[ns] = url
        if event == 'end':
# elenco categorie SOA #################################################
            if elem.tag == fixtag('', 'categoriaSOA', nsmap):
                listaSO.append(elem)
# raccolgo gli elementi descrittivi del prezzario ######################
            if elem.tag == fixtag('', 'prezzario', nsmap):
                prezzarioId = elem.get('prezzarioId')
                przId = elem.get('przId')
                prdStruttura = elem.get('prdStruttura')
                categoriaPrezzario = elem.get('categoriaPrezzario')
                arrotondamento = elem.get('arrotondamento')
                arrotondamentoImporto = elem.get('arrotondamentoImporto')
                arrotondamentoPercentuale = elem.get('arrotondamentoPercentuale')
# lista delle unità di misura ##########################################
            elif elem.tag == fixtag('', 'unitaDiMisura', nsmap):
                listaUM.append(elem)
# titolo del prezzario #################################################
            elif elem.tag == fixtag('', 'przDescrizione', nsmap):
                listaprzDes.append(elem)
# lista completa delle voci ############################################
            elif elem.tag == fixtag('', 'prodotto', nsmap):
                listaPRO.append(elem)
# elenco categorie SOA #################################################
    listaSOA = []
    for elem in listaSO:
        soaCategoria = elem.get('soaCategoria')
        soaId = elem.get('soaId')
        breve = elem.getchildren()[0].get('breve')                # solo - 1 italiano
        soa = (soaCategoria, soaId, breve)
        listaSOA.append(soa)
# lista delle unità di misura ##########################################
    listaUMis = dict()
    for elem in listaUM:
        listaUMis[elem.get('unitaDiMisuraId')] = elem.getchildren()[lang].get('breve')# 0 tedesco - 1 italiano
########################################################################
# formazione della lista completa delle voci ###########################
# prodotto:
    listaPrz =[]
    prezzo =[]
    for elem in listaPRO:
        descEST = elem.getchildren()[lang].get('estesa')                # 0 tedesco - 1 italiano
        if descEST == None:
            descEST = elem.getchildren()[lang].get('breve')             # 0 tedesco - 1 italiano
        desc = elem.getchildren()[lang].get('breve')                    # 0 tedesco - 1 italiano
        um = dict(elem.items()).get('unitaDiMisuraId')
        prezzo = (elem.get('prdId'), desc)
        if um != None:
            prezzo = (elem.get('prdId'), descEST, desc, listaUMis[um])
        if len(elem.getchildren()) == 3:
            valore = dict(elem.getchildren()[2].items())['valore']
            prezzo = (elem.get('prdId'), descEST, desc, listaUMis[um], valore)
        listaPrz.append(prezzo)
    przDes_breve = dict(listaprzDes[lang].items())['breve']
# compilo la tabella ###################################################
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('Listino')
    oSheet.getCellRangeByName('C1').String = przDes_breve
    oSheet.getCellRangeByName('E3').String = przDes_breve
    i=7
    for elem in listaPrz:
        try:
            oSheet.getCellRangeByName('C'+ str (i)).String = elem[0]    #Codice
            oSheet.getCellRangeByName('E'+ str (i)).String = elem[1]    #Descrizione estesa
    #            oSheet.getCellRangeByName('E'+ str (i)).String = elem[2]    #Descrizione breve
            oSheet.getCellRangeByName('G'+ str (i)).String = elem[3]    #Unità di misura
            oSheet.getCellRangeByName('H'+ str (i)).String = elem[4]    #Prezzo
        except IndexError:
            pass
        i=i+1
    listaSOA.sort()
    for elem in listaSOA:
        oSheet.getCellRangeByName('B'+ str (i)).String = elem[0]
        oSheet.getCellRangeByName('C'+ str (i)).String = elem[1]    #Codice
        oSheet.getCellRangeByName('E'+ str (i)).String = elem[2]    #Codice
        i=i+1

import logging
from xml.etree.ElementTree import ElementTree

#~ filename = 'W:/_dwg/ULTIMUSFREE/elenchi/abruzzo/2014/AB2014_03-11-2014.xml'
filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/abruzzo/2014/AB2014_03-11-2014.xml'

#~ filename = 'W:/_dwg/ULTIMUSFREE/elenchi/Bolzano/2014/HBED14_OpereEdili_20141103_XmlSix.xml'
#~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/Bolzano/2014/HBED14_OpereEdili_20141103_XmlSix.xml'

def import_XML ():
    """Routine di importazione di un prezziario XML formato SIX"""
    # inizializzazioe delle variabili
    lista_articoli = list() # lista in cui memorizzare gli articoli da importare
    diz_um = dict() # array per le unità di misura
    # stringhe per nome capitoli
    #~ titolo_sup = str()
    #~ titolo_cap = str()
    #~ titolo_sub = str()
    # stringhe per descrizioni articoli
    desc_codice = str()
    desc_estesa = str()
    # effettua il parsing del file XML
    #~ logging.debug(_("Parsing del file XML: {0}").format(filename))
    tree = ElementTree()
    tree.parse(filename)
    # ottieni l'item root
    root = tree.getroot()
    logging.debug(list(root))
    # effettua il parsing di tutti gli elemnti dell'albero XML
    iter = tree.getiterator()
    listaSOA = []
    articolo = []
    for elem in iter:
        # esegui le verifiche sulla root dell'XML
        if elem.tag == "{six.xsd}intestazione":
            intestazioneId= elem.get("intestazioneId")
            lingua= elem.get("lingua")
            separatore= elem.get("separatore")
            separatoreParametri= elem.get("separatoreParametri")
            valuta= elem.get("valuta")
            autore= elem.get("autore")
            versione= elem.get("versione")
            # inserisci i dati generali
            #~ self.update_dati_generali (nome=None, cliente=None,
                                       #~ redattore=autore,
                                       #~ ricarico=1,
                                       #~ manodopera=None,
                                       #~ sicurezza=None,
                                       #~ indirizzo=None,
                                       #~ comune=None, provincia=None,
                                       #~ valuta=valuta)
        elif elem.tag == "{six.xsd}categoriaSOA":
            soaId = elem.get("soaId")
            soaCategoria = elem.get ("soaCategoria")
            soaDescrizione = elem.find("{six.xsd}soaDescrizione")
            if soaDescrizione != None:
                breveSOA = soaDescrizione.get("breve")
            voceSOA = (soaCategoria, soaId, breveSOA)
            listaSOA.append(voceSOA)
        elif elem.tag == "{six.xsd}prezzario":
            prezzarioId = elem.get("prezzarioId")
            przId= elem.get("przId")
            livelli_struttura= len(elem.get("prdStruttura").split("."))
            categoriaPrezzario= elem.get("categoriaPrezzario")
        elif elem.tag == "{six.xsd}przDescrizione":
            logging.debug(elem.get("breve"))
            # inserisci il titolo del prezziario
            #~ self.update_dati_generali (nome=elem.get("breve"))
            nome=elem.get("breve")
        elif elem.tag == "{six.xsd}unitaDiMisura":
            um_id= elem.get("unitaDiMisuraId")
            um_sim= elem.get("simbolo")
            um_dec= elem.get("decimali")
            # crea il dizionario dell'unita di misura
            diz_um[um_id] = um_sim
            udmDescrizione = elem.find("{six.xsd}udmDescrizione")
            if udmDescrizione != None:
                udmDescrizione.get("breve")
        # se il tag è un prodotto fa parte degli articoli da analizzare
        elif elem.tag == "{six.xsd}prodotto":
            prod_id = elem.get("prodottoId")
            if prod_id is not None:
                prod_id = int(prod_id)
            tariffa= elem.get("prdId")
            if diz_um.get(elem.get("unitaDiMisuraId")) != None:
                unita_misura = diz_um.get(elem.get("unitaDiMisuraId"))
            else:
                unita_misura = ""
            # verifica e ricava le sottosezioni
            sub_desc = elem.find("{six.xsd}prdDescrizione")
            if sub_desc != None:
                desc_codice = sub_desc.get("breve")
                if desc_codice == None:
                    desc_codice = ""
                desc_estesa = sub_desc.get("estesa")
                if desc_estesa == None:
                    desc_estesa = ""
            sub_quot = elem.find("{six.xsd}prdQuotazione")
            if sub_quot != None:
                list_nr = sub_quot.get("listaQuotazioneId")
                sic = elem.get("onereSicurezza")
                if sic != None:
                    sicurezza = float(sic)
                if sub_quot.get("valore") != None:
                    valore = float(sub_quot.get("valore"))
                if sub_quot.get("quantita") is not None:
                    quantita = float(sub_quot.get("quantita"))
            else:
                valore = ""
                sicurezza = ""
                quantita = ""
            articolo = (prod_id,            #0
                        tariffa,            #1
                        desc_codice,        #2
                        desc_estesa,        #3
                        unita_misura,       #4
                        valore,             #5
                        quantita,           #6
                        sicurezza)         #7
            lista_articoli.append(articolo)
# compilo la tabella ###################################################
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.getSheets().getByName('Listino')
    #~ oSheet.getCellRangeByName('C1').String = przDes_breve
    #~ oSheet.getCellRangeByName('E3').String = przDes_breve
# nome del prezzario ###################################################
    oSheet.getCellRangeByName('C1').String = nome
    i=7
    for elem in lista_articoli:
        try:
            oSheet.getCellRangeByName('A'+ str (i)).String = elem[0]    #prod_id
            oSheet.getCellRangeByName('C'+ str (i)).String = elem[1]    #tariffa
            oSheet.getCellRangeByName('E'+ str (i)).String = elem[2]    #desc_codice
# sovrappongo con la descrizione estesa ################################
            oSheet.getCellRangeByName('F'+ str (i)).String = elem[3]    #desc_estesa
            oSheet.getCellRangeByName('G'+ str (i)).String = elem[4]    #unita_misura
            if elem[5] == "":
                oSheet.getCellRangeByName('H'+ str (i)).String = ""
            else:
                oSheet.getCellRangeByName('H'+ str (i)).Value = elem[5] #valore
            if elem[7] == "":
                oSheet.getCellRangeByName('L'+ str (i)).String = ""
            else:
                #~ oSheet.getCellRangeByName('L'+ str (i)).Value = elem[7] #sicurezza
                oSheet.getCellRangeByName('L'+ str (i)).Value = float (elem[5] * elem[7] / 100) #sicurezza
        except IndexError:
            pass
        i=i+1
    #~ listaSOA.sort()
    #~ for elem in listaSOA:
        #~ oSheet.getCellRangeByName('B'+ str (i)).String = elem[0]
        #~ oSheet.getCellRangeByName('C'+ str (i)).String = elem[1]    #Codice
        #~ oSheet.getCellRangeByName('E'+ str (i)).String = elem[2]    #Codice
        #~ i=i+1

#g_exportedScripts = xmlsix2ods, import_XML,
########################################################################
#import pdb; pdb.set_trace() #debugger
