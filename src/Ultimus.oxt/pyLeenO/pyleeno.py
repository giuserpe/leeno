#!/usr/bin/env python -c
# -*- coding: utf-8 -*-
########################################################################
import os, sys
# cos'e' il namespace:
# http://www.html.it/articoli/il-misterioso-mondo-dei-namespaces-1/
import xml.etree.ElementTree as etree
xmL = 'W:/_dwg/ULTIMUSFREE/elenchi/Bolzano/altri_formati/HBED13.six.xml'# file da convertire
#xmL = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/Bolzano/altri_formati/HBED13.six.xml'
# fixtag serve per risalire al dict che contiene il namespace del file XML-six
def fixtag(ns, tag, nsmap):
    return '{' + nsmap[ns] + '}' + tag
########################################################################
def xmlsix2ods ():
    """
    Modulo per la conversione dei prezzari dal formato XML-six a LeenO
    (c) Giuseppe Vizziello 2015
    """
    nsmap = {}
    listaUM = []
    listaprzDes = []
    listaPRO = []
    for event, elem in etree.iterparse(xmL, events=('end', 'start-ns')):
        if event == 'start-ns':
            ns, url = elem
            nsmap[ns] = url
        if event == 'end':
            if elem.tag == fixtag('', 'prezzario', nsmap):
                prezzarioId = elem.get('prezzarioId')
                przId = elem.get('przId')
                prdStruttura = elem.get('prdStruttura')
                categoriaPrezzario = elem.get('categoriaPrezzario')
                arrotondamento = elem.get('arrotondamento')
                arrotondamentoImporto = elem.get('arrotondamentoImporto')
                arrotondamentoPercentuale = elem.get('arrotondamentoPercentuale')
            elif elem.tag == fixtag('', 'unitaDiMisura', nsmap):
                listaUM.append(elem)
            elif elem.tag == fixtag('', 'przDescrizione', nsmap):# titolo del prezzario
                listaprzDes.append(elem)
            elif elem.tag == fixtag('', 'prodotto', nsmap):
                listaPRO.append(elem)
    #lista delle unita' di misura
    listaUMis = dict()
    #unitaDiMisura:
    for elem in listaUM:
        listaUMis[elem.get('unitaDiMisuraId')] = elem.getchildren()[1].get('breve')# in italiano
    ########################################################################
    #Formazione della lista dei prezzi
    #prodotto:
    listaPrz =[]
    prezzo =[]
    for elem in listaPRO:
        desc = elem.getchildren()[1].get('estesa')
        if desc == None:
            desc = elem.getchildren()[1].get('breve')
        um = dict(elem.items()).get('unitaDiMisuraId')
        prezzo = (elem.get('prdId'), desc)
        if um != None:
            prezzo = (elem.get('prdId'), desc, listaUMis[um])
        if len(elem.getchildren()) == 3:
            valore = dict(elem.getchildren()[2].items())['valore']
            prezzo = (elem.get('prdId'), desc, listaUMis[um], valore)
        listaPrz.append(prezzo)
    
    przDes_breve = dict(listaprzDes[1].items())['breve']
    for elem in  listaPrz:
        print(elem[1] + '\n')
    print ('mio')

g_exportedScripts = xmlsix2ods,
########################################################################
#import pdb; pdb.set_trace() #debugger
