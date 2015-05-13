#!/usr/bin/env python -c
# -*- coding: utf-8 -*-
########################################################################
import os, sys, uno, unohelper
# cos'e' il namespace:
# http://www.html.it/articoli/il-misterioso-mondo-dei-namespaces-1/
import logging
from xml.etree.ElementTree import ElementTree

#~ filename = 'W:/_dwg/ULTIMUSFREE/elenchi/abruzzo/2014/AB2014_03-11-2014.xml'
#~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/abruzzo/2014/AB2014_03-11-2014.xml'
#~ filename = "W:/_dwg/ULTIMUSFREE/elenchi/Calabria/2013/Urbanizzazioni_2013_UCU2.xml"

#~ filename = 'W:/_dwg/ULTIMUSFREE/elenchi/Bolzano/2014/HBED14_OpereEdili_20141103_XmlSix.xml'
#~ filename = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/elenchi/Bolzano/2014/HBED14_OpereEdili_20141103_XmlSix.xml'
#~ filename = 'W:/_dwg/ULTIMUSFREE/elenchi/Piemonte/Cuneo/2015/Prezzario_15BD.xml'

def import_XML (filename):
    """Routine di importazione di un prezziario XML formato SIX
    liberamente tratta da Preventares https://launchpad.net/preventares
    di <Davide Vescovini> <davide.vescovini@gmail.com>"""
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
                if valore == 0:
                    valore = ""
                #~ if sub_quot.get("quantita") is not None:
                    #~ quantita = float(sub_quot.get("quantita"))
            else:
                valore = ""
                sicurezza = ""
                #~ quantita = ""
            articolo = (prod_id,            #0
                        tariffa,            #1
                        desc_codice,        #2
                        desc_estesa,        #3
                        unita_misura,       #4
                        valore,             #5
                        #~ quantita,           #6
                        sicurezza)         #7 %
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
            if elem[3] !="":
                oSheet.getCellRangeByName('E'+ str (i)).String = elem[3]    #desc_estesa
#            oSheet.getCellRangeByName('F'+ str (i)).String = elem[3]    #desc_estesa
            oSheet.getCellRangeByName('G'+ str (i)).String = elem[4]    #unita_misura
            if elem[5] == "":
                oSheet.getCellRangeByName('H'+ str (i)).String = ""
            else:
                oSheet.getCellRangeByName('H'+ str (i)).Value = elem[5] #valore
            if elem[7] == "":
                oSheet.getCellRangeByName('L'+ str (i)).String = ""
            else:
                oSheet.getCellRangeByName('k'+ str (i)).Value = elem[7] #sicurezza %
                #~ oSheet.getCellRangeByName('L'+ str (i)).Value = float (elem[5] * elem[7] / 100) #sicurezza
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
