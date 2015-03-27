#!/usr/bin/env python -c
# -*- coding: utf-8 -*-
########################################################################
import os, sys, uno, unohelper
# cos'e' il namespace:
# http://www.html.it/articoli/il-misterioso-mondo-dei-namespaces-1/
import xml.etree.ElementTree as etree
#xmL = 'W:/_dwg/ULTIMUSFREE/elenchi/Bolzano/altri_formati/HBED13.six---.xml'# file da convertire
xmL = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/Bolzano/altri_formati/HBED13.six.xml'
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

g_exportedScripts = xmlsix2ods,
########################################################################
#import pdb; pdb.set_trace() #debugger
