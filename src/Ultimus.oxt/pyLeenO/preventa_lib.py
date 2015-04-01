#!/usr/bin/env python
# -*- Mode: Python; coding: utf-8; indent-tabs-mode: nil; tab-width: 4 -*-
### BEGIN LICENSE
# Copyright (C) 2011 <Davide Vescovini> <davide.vescovini@gmail.com>
# This program is free software: you can redistribute it and/or modify it 
# under the terms of the GNU General Public License version 3, as published 
# by the Free Software Foundation.
# 
# This program is distributed in the hope that it will be useful, but 
# WITHOUT ANY WARRANTY; without even the implied warranties of 
# MERCHANTABILITY, SATISFACTORY QUALITY, or FITNESS FOR A PARTICULAR 
# PURPOSE.  See the GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License along 
# with this program.  If not, see <http://www.gnu.org/licenses/>.
### END LICENSE


################################################################################
############# IMPORTAZIONE MODULI ##############################################
################################################################################
import os
import sys
import datetime
import csv
import logging
from pysqlite2 import dbapi2 as sqlite3
from xml.etree.ElementTree import ElementTree
from string import ascii_lowercase
################################################################################
############ SUPPORTO INTERNAZIONALIZZAZIONE ###################################
################################################################################
import gettext
from gettext import gettext as _
gettext.textdomain('preventares')
################################################################################
############ SUPPORTO PER LA LOCALIZZAZIONE ####################################
################################################################################
import locale
locale.setlocale(locale.LC_ALL, '')
__currency_ = locale.localeconv().get('currency_symbol', 'euro')
################################################################################
########### COSTANTI DA INIZIALIZZARE ##########################################
################################################################################
VERSION = "1.3"
LOGGING = "PreventaLib"
DATI_GENERALI_DEFAULT = {"unita_misura": "cad.", 
                         "ricarico": 1.0,
                         "manodopera": 0.0, 
                         "sicurezza":0.0, 
                         "quantita": 1,
                         "nome":"Nuovo preventivo", 
                         "indirizzo":"", 
                         "comune":"",
                         "provincia":"", 
                         "cliente":"", 
                         "redattore":"", 
                         "valuta": __currency_}
DEFAULT_SETTINGS      = {"ricalcolo_automatico_prezzi":True,
                         "sostituisci_id_cancellati":False,
                         "elimina_da_epu_articoli_non_in_computo":False,
                         "inserisci_articoli_listino_con_articoli_epu": True,
                         "save_database_on_close_connection": True,
                         "decimals": 2,
                         "default_chapter_name": "none",
                         "default_category_name": "none",
                         "default_contract_type": [_("Lavori a Corpo"),
                                                   _("Lavori a Misura"),
                                                   _("Lavori in Economia")],
                         "default_contract_category": []}
################################################################################
########### STRINGHE ESPORTAZIONI ##############################################
################################################################################
STR_TOT_P = _("Totale Parziale")
STR_TOT_G = _("Totale Generale")
STR_MO = _("Manodopera")
STR_SIC = _("Sicurezza")
STR_TYP_NULL = _("-Nessuna categoria di Appalto-")
################################################################################
############ GESTIONE ERRORI ###################################################
################################################################################

class PreventiviError(Exception):
    """Classe per gestire le eccezioni della libreria prevanta_lib"""
    def __init__(self, value):
        date = datetime.datetime.now()
        error_info = sys.exc_info()
        self.value = "{0}: {1} {2}".format(date, value, error_info)
        logging.error(self.value)

    def __str__(self):
        return repr(self.value)

################################################################################
############ GESTIONE LOGGING ##################################################
################################################################################
#livello di logging lasciato da configurare all'applicazione che usa la libreria
logging.getLogger(LOGGING)
################################################################################
############ CLASSE - PREVENTIVO ###############################################
################################################################################

# una classe per gestire i computi
class Preventivo:
    """Classe per gestire il Preventivo"""
    def __init__(self, file_preventivo, lib_settings=None, dati_generali = None):
        self.__FileDB = file_preventivo
        # imposta il file di lavoro, se 'None' lavora in memoria ram
        if self.__FileDB is None:
            self.__FileDB = ":memory:"
        # informa l'utente sul tipo di linguaggio in uso sul suo sistema
        logging.info(_("Linguaggio e codifica locale: {0}").format(locale.getlocale()))
        # apre la connessione col database
        self.db = self.__connect_database()
        # database: use a text_factory that can interpret 8-bit bytestrings
        self.db.text_factory = str
        # inizializza il cursore
        self.c = self.db.cursor()
        # cancella i trigger esistenti nel database (permette l'aggiornamento
        # di database esistenti alle nuove vers. della libreria)
        self.__delete_triggers()
        # crea il database del computo
        self.__create_database()
        # imposta i settaggi generali dell'applicazione
        global settings
        settings = DEFAULT_SETTINGS
        # aggiorna le impostazioni di default
        if lib_settings is not None:
            self.set_settings(lib_settings)
        # imposta i dati generali del preventivo
        if dati_generali is not None:
            self.set_dati_generali(dati_generali)
        else:
            self.dati_generali = DATI_GENERALI_DEFAULT
        # inizializza il database
        self.__inizialize_database()
        #print 'ricarico', self.__FileDB, self.dati_generali_list()[0] #FIXME

    def __connect_database(self):
        """Stabilisci la connessione con il database"""
        try:
            return sqlite3.connect(self.__FileDB)
        except:
            raise PreventiviError(_("Impossibile connettersi col database, verificare che il file '%s' esista") % self.__FileDB)

    def __create_database(self):
        """crea il database SQL, le tavole e i trigger del database"""
        # Creazione delle tavole e dei trigger del database
        self.c.executescript("""
/*Creazione della tavola per Computo */
CREATE TABLE IF NOT EXISTS Computo (id INTEGER PRIMARY KEY, Supercategoria INTEGER, Categoria INTEGER, Subcategoria INTEGER, 
                      Tariffa TEXT, Quantita REAL, Prezzo_unitario REAL, Prezzo_totale REAL, Data DATE, Note TEXT,
                      Tipo_lavori TEXT, Cat_appalto TEXT, Image_art BLOB);

/*Creazione della tavola per Epu */
CREATE TABLE IF NOT EXISTS Epu (Tariffa PRIMARY KEY, Supercapitolo INTEGER, Capitolo INTEGER, Subcapitolo INTEGER, 
                  Descrizione_codice TEXT, Descrizione_voce TEXT, Descrizione_estesa TEXT, 
                  Unita_misura TEXT, Ricarico REAL, Tempo_inst INTEGER, Costo_materiali REAL, Prezzo_unitario REAL, Sicurezza REAL,
                  Cod_analisi TEXT, Note TEXT, CostoMat_1 REAL, CostoMat_2 REAL, CostoMat_3 REAL, CostoMat_4 REAL);

/*Creazione della tavola Analisi */
CREATE TABLE IF NOT EXISTS Analisi (id INTEGER PRIMARY KEY, Tariffa TEXT, Codice TEXT, Descrizione_codice TEXT, Unita_misura TEXT, 
                      Quantita REAL, Prezzo_unitario REAL, Sconto REAL, Accessori INTEGER, Prezzo_totale REAL, Note TEXT);

/*Creazione delle tavole Capitoli per EPU */
CREATE TABLE IF NOT EXISTS Supercapitolo (id INTEGER PRIMARY KEY, Nome TEXT, Descrizione TEXT, Aumento_prezzi REAL);
CREATE TABLE IF NOT EXISTS Capitolo (id INTEGER PRIMARY KEY, Nome TEXT, Descrizione TEXT, Aumento_prezzi REAL);
CREATE TABLE IF NOT EXISTS Subcapitolo (id INTEGER PRIMARY KEY, Nome TEXT, Descrizione TEXT, Aumento_prezzi REAL);

/*Creazione delle tavole Categorie Computo */
CREATE TABLE IF NOT EXISTS Supercategoria (id INTEGER PRIMARY KEY, Nome TEXT, Descrizione TEXT, Aumento_prezzi REAL);
CREATE TABLE IF NOT EXISTS Categoria (id INTEGER PRIMARY KEY, Nome TEXT, Descrizione TEXT, Aumento_prezzi REAL);
CREATE TABLE IF NOT EXISTS Subcategoria (id INTEGER PRIMARY KEY, Nome TEXT, Descrizione TEXT, Aumento_prezzi REAL);

/*Creazione delle tavole per i Dati Generali */
CREATE TABLE IF NOT EXISTS DatiGenerali (Ricarico REAL, Manodopera REAL, Sicurezza REAL, Valuta TEXT, Nome_lavoro TEXT, 
                           Indirizzo TEXT, Comune TEXT, Provincia TEXT, Cliente TEXT, Redattore TEXT);

/*Creazione delle tavole per Manodopera */
CREATE TABLE IF NOT EXISTS Manodopera (id INTEGER PRIMARY KEY, DescOper TEXT, CostOper REAL, PercOper REAL, Note TEXT);

/*Trigger per inserire la data ad ogni articolo modificato, inserito aggiornato in computo*/
CREATE TRIGGER IF NOT EXISTS Insert_Date AFTER INSERT ON Computo FOR EACH ROW
                               BEGIN 
                                 UPDATE Computo SET Data = CURRENT_TIMESTAMP WHERE Tariffa = new.Tariffa; 
                               END;
/*Trigger per: cancellare dall'Epu tutte le 'Tariffe' non più presenti in 'Computo'*/
CREATE TRIGGER IF NOT EXISTS Cancella_articoli_epu AFTER DELETE ON Computo FOR EACH ROW  
                               BEGIN
                                  DELETE FROM Epu WHERE Tariffa NOT IN (SELECT Tariffa FROM Computo); 
                               END;
/*Trigger per: cancellare dal listino tutte le 'Tariffe' non più presenti in 'Epu'*/
CREATE TRIGGER IF NOT EXISTS Cancella_articoli_listino AFTER DELETE ON Epu FOR EACH ROW  
                               BEGIN
                                  DELETE FROM Analisi WHERE Tariffa NOT IN (SELECT Tariffa FROM Epu); 
                               END;
/*Trigger per aggiornare i prezzi unitari dopo la modifica della manodopera in dati generali*/
CREATE TRIGGER IF NOT EXISTS Update_Manodopera AFTER UPDATE OF Manodopera ON DatiGenerali FOR EACH ROW 
                               BEGIN 
                                 UPDATE Epu SET Prezzo_unitario = (Costo_materiali * Ricarico) + 
                                                                  (Tempo_inst/60.0 * new.Manodopera) * 
                                                                  (1+(Sicurezza/100)); 
                               END;
/*Trigger per: Aggiorna in 'Epu' il 'Prezzo_unitario' della voce quando vengono cambiati
il 'Ricarico, Tempo_inst, Prezzo_unitario, Costo_materiali, Sicurezza'*/
CREATE TRIGGER IF NOT EXISTS Update_PU_Epu AFTER UPDATE OF Ricarico, Tempo_inst, Costo_materiali, Sicurezza ON Epu FOR EACH ROW 
                               BEGIN 
                                 UPDATE Epu SET Prezzo_unitario = (((new.Costo_materiali * new.Ricarico) + 
                                                                  (new.Tempo_inst/60.0 * (SELECT Manodopera FROM DatiGenerali))) *
                                                                  (1+(new.Sicurezza/100)) )
                                        WHERE Tariffa = new.Tariffa; 
                               END;
/*Trigger per: Aggiornare in 'Computo' il 'Prezzo_unitario' quando questo viene aggiornato in Epu*/
CREATE TRIGGER IF NOT EXISTS Update_PU_computo AFTER UPDATE OF Prezzo_unitario ON Epu FOR EACH ROW
                               BEGIN 
                                 UPDATE Computo SET Prezzo_unitario = new.Prezzo_unitario WHERE Tariffa = new.Tariffa;
                               END;

/*Trigger per: Aggiorna in 'Computo' il 'Prezzo_totale' quando viene cambiata la 'Quantita' o il 'Prezzo_unitario'*/
CREATE TRIGGER IF NOT EXISTS Update_PT_computo AFTER UPDATE OF Quantita, Prezzo_unitario ON Computo FOR EACH ROW 
                               BEGIN 
                                  UPDATE Computo SET Prezzo_totale = new.Quantita * new.Prezzo_unitario WHERE id= old.id;
                               END;

/*Trigger per: aggiorna la 'Tariffa' se modificata in 'Epu' */
CREATE TRIGGER IF NOT EXISTS Update_tariffa_computo AFTER UPDATE OF Tariffa ON Epu FOR EACH ROW
                               BEGIN
                                  UPDATE Computo SET Tariffa = new.Tariffa WHERE Tariffa = old.Tariffa; 
                                  UPDATE Analisi SET Tariffa = new.Tariffa WHERE Tariffa = old.Tariffa;
                               END;

/*Trigger per: Aggiorna in 'Analisi' il prezzo toale dopo aver modificato Quantita, Prezzo_unitario, Sconto, Accessori*/
CREATE TRIGGER IF NOT EXISTS Update_Pr_tot_after_mod_Analisi AFTER UPDATE OF Quantita, Prezzo_unitario, Sconto, Accessori ON Analisi FOR EACH ROW 
                               BEGIN 
                                 UPDATE Analisi SET Prezzo_totale = new.Quantita * new.Prezzo_unitario * (1- (new.Sconto/100)) + new.Accessori WHERE id = new.id;
                               END;

/*Trigger per l'aggiornamento automatico delle tabelle delle Categorie e Capitoli quando viene aggiunto un elemento in Computo o Epu*/ 
CREATE TRIGGER IF NOT EXISTS Insert_supercategorie AFTER INSERT ON Computo FOR EACH ROW WHEN new.Supercategoria NOT IN (SELECT id FROM Supercategoria)
                               BEGIN
                                  INSERT INTO Supercategoria VALUES (new.Supercategoria, 'Nuovo Supercategoria', '', 0);
                               END;
CREATE TRIGGER IF NOT EXISTS Update_supercategorie AFTER UPDATE OF Supercategoria ON Computo FOR EACH ROW WHEN new.Supercategoria NOT IN (SELECT id FROM Supercategoria)
                               BEGIN
                                  INSERT INTO Supercategoria VALUES (new.Supercategoria, 'Nuovo Supercategoria', '', 0);
                               END;
CREATE TRIGGER IF NOT EXISTS Insert_categorie AFTER INSERT ON Computo FOR EACH ROW WHEN new.Categoria NOT IN (SELECT id FROM Categoria)
                               BEGIN
                                  INSERT INTO Categoria VALUES (new.Categoria, 'Nuova Categoria', '', 0);
                               END;
CREATE TRIGGER IF NOT EXISTS Update_categorie AFTER UPDATE OF Categoria ON Computo FOR EACH ROW WHEN new.Categoria NOT IN (SELECT id FROM Categoria)
                               BEGIN
                                  INSERT INTO Categoria VALUES (new.Categoria, 'Nuova Categoria', '', 0);
                               END;
CREATE TRIGGER IF NOT EXISTS Insert_subcategorie AFTER INSERT ON Computo FOR EACH ROW WHEN new.Subcategoria NOT IN (SELECT id FROM Subcategoria)
                               BEGIN
                                  INSERT INTO Subcategoria VALUES (new.Subcategoria, 'Nuova Subcategoria', '', 0);
                               END;
CREATE TRIGGER IF NOT EXISTS Update_subcategorie UPDATE OF Subcategoria ON Computo FOR EACH ROW WHEN new.Subcategoria NOT IN (SELECT id FROM Subcategoria)
                               BEGIN
                                  INSERT INTO Subcategoria VALUES (new.Subcategoria, 'Nuova Subcategoria', '', 0);
                               END;

CREATE TRIGGER IF NOT EXISTS Insert_supercapitoli AFTER INSERT ON Epu FOR EACH ROW WHEN new.Supercapitolo NOT IN (SELECT id FROM Supercapitolo)
                               BEGIN
                                  INSERT INTO Supercapitolo VALUES (new.Supercapitolo, 'Nuovo SuperCapitolo', '', 0);
                               END;
CREATE TRIGGER IF NOT EXISTS Update_supercapitoli AFTER UPDATE OF Supercapitolo ON Epu FOR EACH ROW WHEN new.Supercapitolo NOT IN (SELECT id FROM Supercapitolo)
                               BEGIN
                                  INSERT INTO Supercapitolo VALUES (new.Supercapitolo, 'Nuovo SuperCapitolo', '', 0);
                               END;
CREATE TRIGGER IF NOT EXISTS Insert_capitoli AFTER INSERT ON Epu FOR EACH ROW WHEN new.Capitolo NOT IN (SELECT id FROM Capitolo)
                               BEGIN
                                  INSERT INTO Capitolo VALUES (new.Capitolo, 'Nuovo Capitolo', '', 0);
                               END;
CREATE TRIGGER IF NOT EXISTS Update_capitoli AFTER UPDATE OF Capitolo ON Epu FOR EACH ROW WHEN new.Capitolo NOT IN (SELECT id FROM Capitolo)
                               BEGIN
                                  INSERT INTO Capitolo VALUES (new.Capitolo, 'Nuovo Capitolo', '', 0);
                               END;
CREATE TRIGGER IF NOT EXISTS Insert_subcapitoli AFTER INSERT ON Epu FOR EACH ROW WHEN new.Subcapitolo NOT IN (SELECT id FROM Subcapitolo)
                               BEGIN
                                  INSERT INTO Subcapitolo VALUES (new.Subcapitolo, 'Nuovo Subcapitolo', '', 0);
                               END;
CREATE TRIGGER IF NOT EXISTS Update_subcapitoli UPDATE OF Subcapitolo ON Epu FOR EACH ROW WHEN new.Subcapitolo NOT IN (SELECT id FROM Subcapitolo)
                               BEGIN
                                  INSERT INTO Subcapitolo VALUES (new.Subcapitolo, 'Nuovo Subcapitolo', '', 0);
                               END;

/*Trigger per: Spostamento automatico alla cat/cap 0 delle voci di 'Computo' ed 'Epu' quando una Categoria o Capitolo viene cancellata*/
CREATE TRIGGER IF NOT EXISTS Cancella_Supercategoria AFTER DELETE ON Supercategoria FOR EACH ROW  
                               BEGIN 
                                  UPDATE Computo SET Supercategoria = 0 WHERE Supercategoria NOT IN (SELECT id FROM Supercategoria); 
                               END;
CREATE TRIGGER IF NOT EXISTS Cancella_Categoria AFTER DELETE ON Categoria FOR EACH ROW  
                               BEGIN 
                                  UPDATE Computo SET Categoria = 0 WHERE Categoria NOT IN (SELECT id FROM Categoria); 
                               END;
CREATE TRIGGER IF NOT EXISTS Cancella_Subcategoria AFTER DELETE ON Subcategoria FOR EACH ROW  
                               BEGIN 
                                  UPDATE Computo SET Subcategoria = 0 WHERE Subcategoria NOT IN (SELECT id FROM Subcategoria); 
                               END;
CREATE TRIGGER IF NOT EXISTS Cancella_Supercapitolo AFTER DELETE ON Supercapitolo FOR EACH ROW  
                               BEGIN 
                                  UPDATE Epu SET Supercapitolo = 0 WHERE Supercapitolo NOT IN (SELECT id FROM Supercapitolo); 
                               END;
CREATE TRIGGER IF NOT EXISTS Cancella_Capitolo AFTER DELETE ON Capitolo FOR EACH ROW  
                               BEGIN 
                                  UPDATE Epu SET Capitolo = 0 WHERE Capitolo NOT IN (SELECT id FROM Capitolo); 
                               END;
CREATE TRIGGER IF NOT EXISTS Cancella_Subcapitolo AFTER DELETE ON Subcapitolo FOR EACH ROW  
                               BEGIN 
                                  UPDATE Epu SET Subcapitolo = 0 WHERE Subcapitolo NOT IN (SELECT id FROM Subcapitolo); 
                               END;""")
        logging.debug(_("TABLE e TRIGGER creati nel database: %s") % self.__FileDB)

    def __delete_triggers(self):
        """cancella tutti i trigger del database - da usare solo PRIMA dell'apertura del DB"""
        self.c.executescript("""
             DROP TRIGGER IF EXISTS Insert_Date;
             DROP TRIGGER IF EXISTS Cancella_articoli_epu;
             DROP TRIGGER IF EXISTS Cancella_articoli_listino;
             DROP TRIGGER IF EXISTS Update_Manodopera;
             DROP TRIGGER IF EXISTS Update_PU_Epu;
             DROP TRIGGER IF EXISTS Update_PU_computo;
             DROP TRIGGER IF EXISTS Update_PT_computo;
             DROP TRIGGER IF EXISTS Update_tariffa_computo;
             DROP TRIGGER IF EXISTS Update_Pr_tot_after_mod_Analisi""")
        # informazione di debug
        logging.debug(_("DROP dei TRIGGER esistenti nel database %s effettuato") % self.__FileDB)

    def __inizialize_database(self):
        """inserisce alcuni valori di dafault ma solo se non già presenti in un nuovo database"""
        # nome di default per i capitoli/categorie 0
        n_cap = (0, settings["default_chapter_name"], '', 0)
        n_cat = (0, settings["default_category_name"], '', 0)
        # Inserimento dei Capitoli di default se non già esistenti
        self.c.execute("""SELECT id FROM Supercapitolo WHERE id=0;""")
        if len(self.c.fetchall()) == 0:
            self.c.execute("""INSERT INTO Supercapitolo VALUES (?,?,?,?);""", n_cap)
        self.c.execute("""SELECT id FROM Capitolo WHERE id=0;""")
        if len(self.c.fetchall()) == 0:
            self.c.execute("""INSERT INTO Capitolo VALUES (?,?,?,?);""", n_cap)
        self.c.execute("""SELECT id FROM Subcapitolo WHERE id=0;""")
        if len(self.c.fetchall()) == 0:
            self.c.execute("""INSERT INTO Subcapitolo VALUES (?,?,?,?);""", n_cap)
        # Inserimento delle Categorie di default se non già esistenti
        self.c.execute("""SELECT id FROM Supercategoria WHERE id=0;""")
        if len(self.c.fetchall()) == 0:
            self.c.execute("""INSERT INTO Supercategoria VALUES (?,?,?,?);""", n_cat)
        self.c.execute("""SELECT id FROM Categoria WHERE id=0;""")
        if len(self.c.fetchall()) == 0:
            self.c.execute("""INSERT INTO Categoria VALUES (?,?,?,?);""", n_cat)
        self.c.execute("""SELECT id FROM Subcategoria WHERE id=0;""")
        if len(self.c.fetchall()) == 0:
            self.c.execute("""INSERT INTO Subcategoria VALUES (?,?,?,?);""", n_cat)
        # Inserimento dei dati di default della tabella dei Dati Generali
        self.c.execute("""SELECT * FROM DatiGenerali;""")
        if len(self.c.fetchall()) == 0:
            self.c.execute("""INSERT INTO DatiGenerali VALUES (:ricarico, :manodopera, :sicurezza, :valuta, 
                              :nome, :indirizzo, :comune, :provincia, :cliente, :redattore);""", self.dati_generali)
        else:
            dati = self.__retrive_dati_generali()
            self.set_dati_generali(dati)
        # modifica della tabella computo per retrocompatibilità versioni precedenti
        try:
            self.c.executescript("""
                              ALTER TABLE Computo ADD COLUMN Tipo_lavori TEXT; 
                              ALTER TABLE Computo ADD COLUMN Cat_appalto TEXT; 
                              ALTER TABLE Computo ADD COLUMN Image_art BLOB;""")
            logging.debug("TABLE 'Computo' aggiornata, aggiunte colonne: 'Tipo_lavori', 'Cat_appalto', 'Image_art'")
        except sqlite3.OperationalError:
            # verificare l'esistenza delle colonne effettua il logging e non fare altro
            try:
                self.c.execute("""SELECT Tipo_lavori, Cat_appalto, Image_art FROM Computo;""")
                logging.debug(_("Colonne 'Tipo_lavori', 'Cat_appalto', 'Image_art' nella TABLE 'Computo' esistenti, aggiornamento non necessario."))
            except sqlite3.OperationalError:
                raise PreventiviError(_("Impossibile aggiungere le Colonne 'Tipo_lavori', 'Cat_appalto', 'Image_art' nella TABLE 'Computo' errore nel database %s.") % self.__FileDB)
        # consolida il database appena aperto per effettuare i salvataggi delle
        # tabelle e dei dati base appena inseriti
        self.db.commit()
        return

    def __retrive_dati_generali(self):
        """"""
        self.c.execute("""SELECT * FROM DatiGenerali;""")
        lista = self.c.fetchall()
        dati = {"ricarico":lista[0][0], "manodopera":lista[0][1], "sicurezza":lista[0][2], 
                "valuta":lista[0][3], "nome":lista[0][4], "indirizzo":lista[0][5], 
                "comune":lista[0][6], "provincia":lista[0][7], "cliente":lista[0][8]}
        return dati

    def delete_database(self):
        """cancella tutte le tavole del database"""
        self.__connect_database()
        try:
            self.c.execute("""DROP TABLE IF EXISTS Computo;""")
            self.c.execute("""DROP TABLE IF EXISTS Epu;""")
            self.c.execute("""DROP TABLE IF EXISTS Analisi;""")
            self.c.execute("""DROP TABLE IF EXISTS Supercapitolo;""")
            self.c.execute("""DROP TABLE IF EXISTS Capitolo;""")
            self.c.execute("""DROP TABLE IF EXISTS Subcapitolo;""")
            self.c.execute("""DROP TABLE IF EXISTS Supercategoria;""")
            self.c.execute("""DROP TABLE IF EXISTS Categoria;""")
            self.c.execute("""DROP TABLE IF EXISTS Subcategoria;""")
            self.c.execute("""DROP TABLE IF EXISTS DatiGenerali;""")
            self.c.execute("""DROP TABLE IF EXISTS Manodopera;""")
        except:
            raise PreventiviError(_("Errore, impossibile eliminare le tavole dal database. Tavole inesistenti o File database %s mancante") % self.__FileDB)
        finally:
            logging.debug(_("Tutte le tavole esistenti nel database %s sono state eliminate") % self.__FileDB)

    def get_database(self):
        """Resitituisce l'oggetto database del preventivo"""
        return self.db

    def get_filename(self):
        """Resitituisce il nome del file database del preventivo"""
        return self.__FileDB

    def get_settings(self):
        """Resitituisce le impostazioni del preventivo"""
        return settings

    def set_settings(self, new_settings):
        """modifica le impostazioni generali della libreria"""
        # se i parametri iniziali sono validi sostituisci i valori di defalut
        if new_settings is not None and type(new_settings) is dict:
            for key in new_settings.keys():
                if key in settings:
                    settings[key] = new_settings[key]
                else:
                    raise PreventiviError(_("Parametro iniziale: '%s' non riconosciuto") % key)
        else:
            raise PreventiviError(_("Nuove Impostazioni non riconosciute. 'new_settings' deve essere <type 'dict'>"))
        # infine aggiorna le impostazioni del database
        if settings["elimina_da_epu_articoli_non_in_computo"]:
            self.c.executescript("""CREATE TRIGGER IF NOT EXISTS Cancella_articoli_epu AFTER DELETE ON Computo FOR EACH ROW  
                               BEGIN
                                  DELETE FROM Epu WHERE Tariffa NOT IN (SELECT Tariffa FROM Computo); 
                               END;""")
        elif not settings["elimina_da_epu_articoli_non_in_computo"]:
            self.c.execute("""DROP TRIGGER IF EXISTS Cancella_articoli_epu""")
        # aggiorna l'aggiornamento automatico dei valori di Epu 
        if settings["ricalcolo_automatico_prezzi"]:
            self.c.executescript("""
            /*Trigger per: Aggiorna in 'Epu' il Costo_materiali dopo aver 
              inserito un articolo in 'Analisi'*/
            CREATE TRIGGER IF NOT EXISTS Update_Costo_mat_Epu_after_insert_Analisi AFTER INSERT ON Analisi FOR EACH ROW 
                               BEGIN 
                                 UPDATE Epu SET Costo_materiali = (SELECT total(Prezzo_totale) FROM Analisi WHERE Tariffa = Epu.Tariffa)
                                            WHERE Tariffa IN (SELECT Tariffa FROM Analisi);
                               END;
            /*Trigger per: Aggiorna in 'Epu' il Costo_materiali dopo aver 
              modificato il Prezzo_totale di un articoli in 'Analisi'*/
            CREATE TRIGGER IF NOT EXISTS Update_Costo_mat_Epu_after_mod_Analisi AFTER UPDATE OF Prezzo_totale ON Analisi FOR EACH ROW
                               BEGIN 
                                 UPDATE Epu SET Costo_materiali = (SELECT total(Prezzo_totale) FROM Analisi WHERE Tariffa = Epu.Tariffa) 
                                            WHERE Tariffa IN (SELECT Tariffa FROM Analisi);
                               END;""")
        elif not settings["ricalcolo_automatico_prezzi"]:
            self.c.executescript("""
                DROP TRIGGER IF EXISTS Update_Costo_mat_Epu_after_insert_Analisi;
                DROP TRIGGER IF EXISTS Update_Costo_mat_Epu_after_mod_Analisi;""")
        return settings

    def get_dati_generali(self):
        """Resitituisce i dati generali del preventivo"""
        return self.dati_generali

    def set_dati_generali(self, new_dati_generali):
        """modifica i dati generali della libreria"""
        # se i parametri iniziali sono validi sostituisci i valori di defalut
        if new_dati_generali is not None and type(new_dati_generali) is dict:
            for key in new_dati_generali.keys():
                if key in self.dati_generali:
                    self.dati_generali[key] = new_dati_generali[key]
                else:
                    raise PreventiviError(_("Dato generale iniziale: '%s' non riconosciuto") % key)
        else:
            raise PreventiviError(_("Nuove Impostazioni non riconosciute. Le impostazioni devono essere un oggetto type(dict)"))
        return self.dati_generali

################################################################################
############ PREVENTIVO - operazioni di connessione salvataggio DB #############
################################################################################
# Nota: alle seguenti operazioni servono a salvare, effettuare 
# il rollback e chiudere la connessione con il database

    def save_database (self):
        """salva le modifiche al database"""
        try:
            # Save (commit) the changes to database
            self.db.commit()
        except:
            raise PreventiviError(_("Impossibile salvare il database, verificare che il file %s esista") % self.__FileDB)
        finally:
            logging.debug(_("Modifiche al database %s salvate") % self.__FileDB)

    def rollback_database (self):
        """annulla le modifiche effettuate sul database dall'ultimo salvataggio"""
        try:
            # Rollback (undo) the changes to database
            self.db.rollback()
        except:
            raise PreventiviError(_("Impossibile annullare le modificha al database, verificare che il file %s esista") % self.__FileDB)
        finally:
            logging.debug(_("Modifiche al database '%s' annullate") % self.__FileDB)

    def connection_shutdown (self):
        """chiude la connessione col database"""
        # If settings allow save (commit) the changes to database
        if settings["save_database_on_close_connection"]:
            self.db.commit()
        else:
            self.db.rollback()
        # Chiudiamo il cursore quando abbiamo finito
        self.c.close()
        logging.debug(_("Connessione al database %s chiusa") % self.__FileDB)
        return True

################################################################################
############ DATI GENERALI - operazioni base sul database DatiGenerali #########
################################################################################
# Nota: *****

    def update_dati_generali (self, nome=None, cliente=None, redattore=None, 
                              ricarico=None, manodopera=None, sicurezza=None, 
                              indirizzo=None, comune=None, provincia=None, valuta=None):
        """Aggiorna i dati generali del preventivo"""
        result = True
        if nome is None: nome = self.dati_generali["nome"]
        if cliente is None: cliente = self.dati_generali["cliente"]
        if ricarico is None: ricarico = self.dati_generali["ricarico"]
        if manodopera is None: manodopera = self.dati_generali["manodopera"]
        if sicurezza is None: sicurezza = self.dati_generali["sicurezza"]
        if indirizzo is None: indirizzo = self.dati_generali["indirizzo"]
        if comune is None: comune = self.dati_generali["comune"]
        if provincia is None: provincia = self.dati_generali["provincia"]
        if valuta is None: valuta = self.dati_generali["valuta"]
        if redattore is None: redattore = self.dati_generali["redattore"]
        dati = {"ricarico":ricarico, "manodopera":manodopera, "sicurezza":sicurezza, 
                "valuta":valuta, "nome":nome, "cliente":cliente, "redattore":redattore,
                "indirizzo":indirizzo, "comune":comune, "provincia":provincia}
        try:
            self.c.execute("""UPDATE DatiGenerali SET 
                              Ricarico=:ricarico, Manodopera=:manodopera,
                              Sicurezza=:sicurezza, Valuta=:valuta,
                              Nome_lavoro=:nome, Cliente=:cliente,
                              Redattore=:redattore, Indirizzo=:indirizzo,
                              Comune=:comune, Provincia=:provincia;""", dati)
        except:
            raise PreventiviError(_("Errore dell'aggiornamento dei dati generali del preventivo"))
            result = False
        finally:
            dati = self.__retrive_dati_generali()
            self.set_dati_generali(dati)
            logging.debug(_("Dati generali aggiornati in 'DatiGenerali'"))
            return result

    def update_ricarico_epu (self, new_ricarico):
        """Imposta il nuovo ricarico in tutti gli articoli di Epu"""
        result = True
        self.dati_generali["ricarico"] = new_ricarico
        try:
            self.c.execute("""UPDATE Epu SET Ricarico=:ricarico;""", self.dati_generali)
            self.c.execute("""UPDATE DatiGenerali SET Ricarico=:ricarico;""", self.dati_generali)
        except:
            raise PreventiviError(_("Errore nell'aggiornamento del ricarico in tutti gli articoli di Epu"))
            result = False
        finally:
            logging.debug(_("Ricarico aggiornato in 'Epu' e 'DatiGenerali'"))
            return result

    def update_sicurezza_epu (self, new_sicurezza):
        """Imposta la nuova incidenza della sicurezza in tutti gli articoli di Epu"""
        result = True
        self.dati_generali["sicurezza"] = new_sicurezza
        try:
            self.c.execute("""UPDATE Epu SET Sicurezza=:sicurezza;""", self.dati_generali)
            self.c.execute("""UPDATE DatiGenerali SET Sicurezza=:sicurezza;""", self.dati_generali)
        except:
            raise PreventiviError(_("Errore nell'aggiornamento dell'incidenza della sicurezza in tutti gli articoli di Epu"))
            result = False
        finally:
            logging.debug(_("Sicurezza aggiornata in 'Epu' e 'DatiGenerali'"))
            return result

################################################################################
############ MANODOPERA - TABELLE COMPOSIZIONE DELLA MANODOPERA ################
################################################################################

    def insert_manodopera (self, nome=None, costo=None, perc=None, note=None):
        """Inserisce un nuovo valore di manodopera del preventivo"""
        result = True
        if nome is None: nome = str()
        if costo is None: 
            costo = self.dati_generali_list()[1]
        if perc is None: 
            perc = 100
        elif perc > 100:
            raise PreventiviError(_("La percentuale di manodopera '%s' è superiore a 100") % nome)
        if note is None: note = str()
        dati = {"nome":nome, "costo":costo, "perc":perc, "note":note}
        try:
            self.c.execute("""INSERT INTO Manodopera VALUES (NULL, :nome, 
                                          :costo, :perc, :note);""", dati)
        except:
            raise PreventiviError(_("Errore nell'inserimento del rigo di moanodopera '%s'") % nome)
            result = False
        finally:
            logging.debug(_("Rigo manodopera '%s' inserito.") % nome)
            return result

    def update_manodopera (self, key, nome= None, costo=None, perc=None, note=None):
        """Aggiorna un valore di manodopera del preventivo"""
        result = True
        if nome is None: 
            nome = self.get_table_manodopera (key=nome)[1]
        if costo is None: 
            costo = self.get_table_manodopera (key=nome)[2]
        if perc is None: 
            perc = self.get_table_manodopera (key=nome)[3]
        if note is None: 
            note = self.get_table_manodopera (key=nome)[4]
        dati = {"nome":nome, "costo":costo, "perc":perc, "note":note, "id":key}
        try:
            self.c.execute("""UPDATE Manodopera SET DescOper=:nome, CostOper=:costo, 
                                   PercOper=:perc, Note=:note WHERE id=:id;""", dati)
        except:
            raise PreventiviError(_("Errore nell'aggiornamento del rigo di moanodopera '%s'") % nome)
            result = False
        finally:
            logging.debug(_("Rigo manodopera '%s' aggiornato.") % nome)
            return result

    def delete_manodopera (self, key):
        """Cancella un valore di manodopera dal preventivo"""
        result = True
        dati = {"id":key}
        try:
            self.c.execute("""DELETE FROM Manodopera WHERE id=:id;""", dati)
        except:
            raise PreventiviError(_("Errore nella cancellazione del rigo di moanodopera '%s'") % key)
            result = False
        finally:
            logging.debug(_("Rigo manodopera '%s' cancellato.") % key)
            return result

    def calcola_media_manodopera (self, update_dg = True):
        """
        Calcola il valore medio della manodopera in base alla tabella 
        'Manodopera' e aggiorna i dati generali
        """
        # estrai i dati dalla tabella
        lista = self.get_table_manodopera()
        if len(lista) == 0:
            logging.debug(_("Tabella manodopera vuota, manodopera non aggiornata"))
            return None
        # inizializza le variabili
        costo_tot = 0
        perc_tot = 0
        # per ogni elemento della tabella aggiorna il costo totale
        for key, nome, costo, perc, note in lista:
            perc_tot += perc
            costo_tot += costo*perc
        # calcola il nuovo valore medio della manodopera
        if perc_tot != 0:
            mo = costo_tot/perc_tot
        else:
            logging.debug(_("Percentuale totale delle righe manodopera nulla, manodopera non aggiornata"))
            return None
        # aggiorna i dati generali
        if update_dg:
            self.update_dati_generali (manodopera= mo)
            logging.debug(_("Valore della manodopera aggiornata: '%s'") % mo)
        return mo

    def get_table_manodopera (self, key=None):
        """elenca i valori della tabella manodopera"""
        if key is not None:
            self.c.execute("""SELECT * FROM Manodopera WHERE id=:key;""", {"key":key})
        else:
            self.c.execute("""SELECT * FROM Manodopera ORDER BY id;""")
        return self.c.fetchall()

    def __get_perc_tot (self, key=None):
        """"""
        perc_tot = 0
        self.c.execute("""SELECT PercOper FROM Manodopera;""")
        l_perc = self.c.fetchall()
        if len(l_perc) > 0:
            for perc in l_perc:
                perc_tot += perc[0]
            return perc_tot
        else:
            return 1

################################################################################
############ CAPITOLI e CATEGORIE ##############################################
################################################################################

# Nota: le operazioni avvengono sempre su singole categorie (sub o super) e capitoli (sub o super)
# le categorie vengono create automaticamente con l'immissione degli articoli
# ma possono anche essere create o eliminate separatamente con i seguenti metodi

    def insert_capitoli_categorie (self, Tipo, Nome, Descrizione = None, Aumento_prezzi = None, codice = None):
        """Aggiunge un capitolo o una categoria al database"""
        if Tipo != "Supercapitolo" and Tipo != "Capitolo" and Tipo != "Subcapitolo" and Tipo != "Supercategoria" and Tipo != "Categoria" and Tipo != "Subcategoria":
            raise PreventiviError(_("Errore 'Tipo' deve essere Supercapitolo/Supercapitolo/Capitolo/Subcapitolo o /Supercategoria/Categoria/Subcategoria"))
            Tipo = "Subcategoria"
        if Descrizione is None:
            Descrizione = str()
        elif type(Descrizione) != str:
            raise PreventiviError(_("Errore nella tabella %s 'Descrizione' = %s deve essere una stringa") % (str(Tipo), str(Descrizione)))
            Descrizione = str()
        if Aumento_prezzi is None:
            Aumento_prezzi = 0.0
        elif type(Aumento_prezzi) != int:
            raise PreventiviError(_("Errore nella tabella %s 'Aumento_prezzi' = %s deve essere un numero intero") % (str(Tipo), str(Aumento_prezzi)))
            Aumento_prezzi = 0.0
        if codice is not None and type(codice) != int:
            codice = None
        # preparazione dell'istruzione di cancellazione
        row = (codice, Nome, Descrizione, Aumento_prezzi)
        istruzione = """INSERT INTO %s VALUES """ % (Tipo)
        try:
            self.c.execute(istruzione + """(?, ?, ?, ?);""", row)
            logging.debug(_("Inserita nella tabella '{0}': {1}").format(Tipo, Nome))
        except:
             raise PreventiviError(_("Errore di esecuzione inserimento nella tabella %s Capitoli e Categorie: %s") % (str(Tipo), str(Nome)))
             return False
        return True

    def update_capitoli_categorie (self, key, Tipo, Nome, Descrizione=None, Aumento_prezzi=None):
        """Aggiorna i valori di un capitolo o una categoria"""
        result = True
        if Descrizione is None:
            Descrizione = ""
        if Aumento_prezzi is None:
            Aumento_prezzi = 0
        diz = {"key":key, "Nome":Nome, "Descrizione": Descrizione, "Aumento_prezzi" : Aumento_prezzi}
        try:
            self.c.execute("""UPDATE %s 
                          SET Nome =:Nome, Descrizione=:Descrizione, Aumento_prezzi=:Aumento_prezzi
                          WHERE id =:key;""" % Tipo, diz)
            logging.debug("Tabella aggiornata '{0}'colonna 'id': {1}".format(Tipo, key))
        except:
             raise PreventiviError(_("Errore di aggiornamento da tabella %s nome %s id %s") % (str(Tipo), str(Nome), str(key)))
             result = False
        return result

    def delete_capitoli_categorie (self, Tipo, key, column = None):
        """Cancella dal database un capitolo o una categoria"""
        result = True
        if Tipo != "Supercapitolo" and Tipo != "Capitolo" and Tipo != "Subcapitolo" and Tipo != "Supercategoria" and Tipo != "Categoria" and Tipo != "Subcategoria":
            raise PreventiviError("Errore 'Tipo' deve essere Supercapitolo/Supercapitolo/Capitolo/Subcapitolo o /Supercategoria/Categoria/Subcategoria")
            Tipo = "Subcategoria"
        elif Tipo is None:
            Tipo = "Subcategoria"
        if column is None:
            column = "id"
        elif column != "id" or column != "Nome" or column != "Descrizione" or column != "Aumento_prezzi" or column != None:
            raise PreventiviError(_("Errore 'column' deve essere una colonna valida della tabella %s") % Tipo)
            column = "id"
        try:
            # eseguiamo l'istruzione SQL di cancellazione
            self.c.execute("""DELETE FROM %s WHERE %s = '%s' ;""" % (Tipo, column, key))
            logging.debug(_("Tabella cancellata '{0}'colonna 'id': {1}").format(Tipo, key))
        except:
            raise PreventiviError(_("Errore eliminazione da tabella %s colonna %s voce %s") % (str(Tipo), str(column), str(key)))
            result = False
        return result

################################################################################
############ ELENCO PREZZI - operazioni base sul database EPU ##################
################################################################################
# Nota: le operazioni avvengono sempre su oggetti appartenenti alla classe Articoli
    def __verifica_tariffa_epu (self, articolo):
        """verifica se la tariffa è presente in EPU, se non lo fosse la inserisce"""
        self.c.execute("""SELECT Tariffa FROM Epu""")
        lista_tariffe = self.c.fetchall()
        t = (articolo.tariffa,)
        if t not in lista_tariffe:
            self.insert_articoli_epu(None, articolo)
        return articolo

    def __verifica_tariffa_duplicata (self, articolo):
        """verifica se la tariffa è duplicata prima dell'inserimento e se la  
           tariffa esiste già ne cambia il valore con uno progressivo 
           (es. A001 --> A001 (1)"""
        self.c.execute("""SELECT Tariffa FROM Epu""")
        lista_tariffe = self.c.fetchall()
        t = (articolo.tariffa,)
        n = 1
        while t in lista_tariffe:
             new_tariffa = "%s (%d)" % (articolo.tariffa, n)
             t = (new_tariffa,)
             n+=1
        articolo.tariffa = t[0]
        return articolo

    def insert_articoli_epu (self, verifica_tariffa = False, *lista_articoli):
        """Inserisce un nuovo articolo nel database EPU"""
        lista_articoli_inseriti = list()
        for articolo in lista_articoli:
            tariffa_old = articolo.tariffa
            if verifica_tariffa:
                articolo = self.__verifica_tariffa_duplicata (articolo)
            rigo = articolo.row()
            try:
                self.c.execute("""INSERT INTO Epu VALUES (
                                  :Tar, :SupCap, :Cap, :SubCap,
                                  :DesCod, :DesVoc, :DesEst, 
                                  :UM, :Ric, :Temp, :CostoMat, :PrezzoUnit, :Sicurezza,
                                  :cod_listino, :Note, 
                                  :CostoMat1, :CostoMat2, :CostoMat3, :CostoMat4);""", rigo)
                lista_articoli_inseriti.append(articolo)
            except:
                raise PreventiviError(_("Errore di inserimento dell'articolo: %s in EPU") % str(articolo))
            finally:
                # se le impostazioni lo consentono inserisci anche gli articoli di listino del vecchio prezzo
                if settings["inserisci_articoli_listino_con_articoli_epu"] == True:
                    if len(articolo.art_listino) == 0:
                        art_listino = self.get_articoli_listino(tariffa_old)
                    else: 
                        art_listino = articolo.art_listino
                    if len(art_listino) > 0: 
                        self.insert_articoli_listino(articolo.tariffa, *art_listino)
                logging.debug(_("Inserito articolo tabella '{0}' col 'Tariffa': {1}").format('Epu', articolo.tariffa))
        return lista_articoli_inseriti

    def delete_articoli_epu (self, *lista_articoli):
        """Cancella un articolo esistente dal database epu"""
        result = True
        for articolo in lista_articoli:
            try:
                # eseguiamo l'istruzione SQL di cancellazione
                self.c.execute("""DELETE FROM Computo WHERE Tariffa=:tariffa;""", {"tariffa": articolo.tariffa})
                self.c.execute("""DELETE FROM Epu WHERE Tariffa=:tariffa;""", {"tariffa": articolo.tariffa})
            except:
                raise PreventiviError(_("Errore di cancellazione dell'articolo: %s") % str(articolo))
                result = False
            finally:
                logging.debug(_("Cancellato articolo tabella '{0}' col 'id': {1}").format('Epu', articolo.tariffa))
        return result

    def update_articolo_epu (self, tariffa_esistente, articolo):
        """Aggiorna un articolo di epu data la sua prymary key (tariffa)"""
        result = True
        rigo = articolo.row()
        rigo["old_Tariffa"] = tariffa_esistente
        try:
            self.c.execute("""UPDATE Epu SET Tariffa=:Tar, Supercapitolo=:SupCap, Capitolo=:Cap, Subcapitolo=:SubCap,
                                  Descrizione_codice=:DesCod, Descrizione_voce=:DesVoc, Descrizione_estesa=:DesEst, 
                                  Unita_misura=:UM, Ricarico=:Ric, Tempo_inst=:Temp, Costo_materiali=:CostoMat,
                                  Sicurezza=:Sicurezza, Cod_analisi=:cod_listino, Note=:Note, CostoMat_1=:CostoMat1, 
                                  CostoMat_2=:CostoMat2, CostoMat_3=:CostoMat3, CostoMat_4=:CostoMat4
                                  WHERE Tariffa =:old_Tariffa;""", rigo)
        except:
            raise PreventiviError(_("Errore dell'aggiornamento dell'articolo: %s in Epu") % str(articolo))
            result = False
        finally:
            logging.debug(_("Aggiornato articolo tabella '{0}' col 'id': {1}").format('Epu', articolo.tariffa))
        return result

    def nuovo_articolo_epu (self, tariffa, descrizione_codice, descrizione_voce=None, descrizione_estesa=None, 
                                supercapitolo = None, capitolo = None, subcapitolo = None):
        """Crea un nuovo articolo nella tabella Epu"""
        if supercapitolo is None: supercapitolo = 0
        if capitolo is None: capitolo = 0
        if subcapitolo is None: subcapitolo = 0
        articolo = ArticoloComputo(self.db, supercapitolo = supercapitolo, capitolo = capitolo, subcapitolo = subcapitolo, 
                                   supercategoria = 0, categoria = 0, subcategoria = 0, 
                                   tariffa= tariffa, codice=None, descrizione_codice=descrizione_codice,
                                   descrizione_voce=None, descrizione_estesa=None, 
                                   unita_misura=self.dati_generali["unita_misura"], quantita=0, 
                                   ricarico= self.dati_generali["ricarico"], 
                                   tempo_inst=0, costo_materiali=0.0, prezzo_unitario=0.0, sicurezza=self.dati_generali["sicurezza"], 
                                   cod_listino=str(), note=None)
        #inserisci il nuovo articolo nel database epu
        self.insert_articoli_epu (True, articolo)
        return articolo

    def copia_articolo_epu (self, tariffa, supercategoria = None, categoria = None, subcategoria = None, copia_nomi_capitoli= False):
        """Restituisce un articolo epu esistente nel database"""
        if supercategoria is None: supercategoria = 0
        if categoria is None: categoria = 0
        if subcategoria is None: subcategoria = 0
        self.c.execute("""SELECT * FROM Epu WHERE Tariffa = '%s'""" % tariffa)
        lista_epu = self.c.fetchall()
        if len(lista_epu) == 0:
             logging.debug(_("Impossibile copiare la tariffa '%s', la tariffa non esiste nel database") % tariffa)
             return None
        e = lista_epu[0]
        articolo = ArticoloComputo(database=self.db, supercapitolo = e[1], capitolo = e[2], subcapitolo = e[3], \
                       supercategoria = supercategoria, categoria = categoria, subcategoria = subcategoria, \
                       tariffa= tariffa, codice=None, descrizione_codice=e[4], descrizione_voce=e[5], descrizione_estesa=e[6], \
                       unita_misura=e[7], quantita=0, ricarico=e[8], tempo_inst=e[9], costo_materiali=e[10], prezzo_unitario=e[11], sicurezza=e[12], \
                       cod_listino=e[13], note=e[14], data= None, costo_mat1=e[15], costo_mat2=e[16], costo_mat3=e[17],costo_mat4=e[18])
        if copia_nomi_capitoli:
            # ricavo i nomi dei capitoli (id, Nome, descr, aumento_prezzi)
            articolo.nome_supercapitolo, articolo.nome_capitolo, articolo.nome_subcapitolo = self.get_capitoli_name_from_id (
                                         articolo.supercapitolo, articolo.capitolo, articolo.subcapitolo)
        #copiare (se le impostazioni lo consentono) gli articoli di listino nell'attributo articolo.art_listino
        if settings["inserisci_articoli_listino_con_articoli_epu"] == True:
            art_listino = self.get_articoli_listino(articolo.tariffa)
            if len(art_listino) > 0: 
                articolo.art_listino = art_listino
        return articolo

    def incolla_articolo_epu (self, supercapitolo = None, capitolo = None, subcapitolo = None, *lista_articoli):
        """Inserisce nei nuovi capitoli selezionati una lista di articoli passati come argomento"""
        for articolo in lista_articoli:
            if supercapitolo is not None and type(supercapitolo) is int: 
                articolo.supercapitolo = supercapitolo
            if capitolo is not None and type(capitolo) is int:
                articolo.capitolo = capitolo
            if subcapitolo is not None and type(subcapitolo) is int:
                articolo.subcapitolo = subcapitolo
        #inserisci gli articoli nel database
        self.insert_articoli_epu (True, *lista_articoli)
        return lista_articoli

    def update_articoli_epu (self, supercapitolo, capitolo, subcapitolo, un_mis, ric, temp,
                             mat, sicurezza, costi1, costi2, costi3, costi4, *lista_articoli):
        """Aggiorna simultaneamente un gruppo di articoli di elenco prezzi, 
           un parametro con valore None non viene aggiornato!"""
        result = True
        for articolo in lista_articoli:
            rigo = articolo.row()
            if supercapitolo is not None:
                rigo["SupCap"] = supercapitolo
            if capitolo is not None:
                rigo["Cap"] = capitolo
            if subcapitolo is not None:
                rigo["SubCap"] = subcapitolo
            if un_mis is not None:
                rigo["UM"] = un_mis
            if ric is not None:
                rigo["Ric"] = ric
            if temp is not None:
                rigo["Temp"] = temp
            if mat is not None:
                rigo["CostoMat"] = mat
            if sicurezza is not None:
                rigo["Sicurezza"] = sicurezza
            if costi1 is not None:
                rigo["CostoMat1"] = costi1
            if costi2 is not None:
                rigo["CostoMat2"] = costi2
            if costi3 is not None:
                rigo["CostoMat3"] = costi3
            if costi4 is not None:
                rigo["CostoMat4"] = costi4
            try:
                self.c.execute("""UPDATE Epu SET Supercapitolo=:SupCap, Capitolo=:Cap, Subcapitolo=:SubCap, 
                                  Unita_misura=:UM, Ricarico=:Ric, Tempo_inst=:Temp, Costo_materiali=:CostoMat,
                                  Sicurezza=:Sicurezza, CostoMat_1=:CostoMat1, CostoMat_2=:CostoMat2, 
                                  CostoMat_3=:CostoMat3, CostoMat_4=:CostoMat4 WHERE Tariffa =:Tar;""", rigo)
            except:
                raise PreventiviError(_("Errore nell'aggiornamento dell'articolo di Epu: %s") % str(articolo))
                result = False
        return result

    def scambia_prezzi_articoli_epu (self, costo_partenza, costo_arrivo, *lista_articoli):
        """La funzione copia il valore di una fascia di prezzo (es. fascia '0')
           in un'altra (es. fascia '3') per un gruppo di articoli"""
        if type(costo_partenza) is not int or type(costo_arrivo) is not int:
           raise PreventiviError(_("I valori 'costo_partenza' e 'costo_arrivo' devono essere numeri interi"))
           return True
        for articolo in lista_articoli:
            rigo = articolo.row()
            # carico in memoria il costo di 'partenza'
            if costo_partenza == 0:
                costo = rigo["CostoMat"]
            elif costo_partenza == 1:
                costo = rigo["CostoMat1"]
            elif costo_partenza == 2:
                costo = rigo["CostoMat2"]
            elif costo_partenza == 3:
                costo = rigo["CostoMat3"]
            elif costo_partenza == 4:
                costo = rigo["CostoMat4"]
            # sostituisco il costo di 'arrivo'
            if costo_arrivo == 0:
                rigo["CostoMat"] = costo
            elif costo_arrivo == 1:
                rigo["CostoMat1"] = costo
            elif costo_arrivo == 2:
                rigo["CostoMat2"] = costo
            elif costo_arrivo == 3:
                rigo["CostoMat3"] = costo
            elif costo_arrivo == 4:
                rigo["CostoMat4"] = costo
            try:
                self.c.execute("""UPDATE Epu SET Costo_materiali=:CostoMat, CostoMat_1=:CostoMat1, CostoMat_2=:CostoMat2, 
                                  CostoMat_3=:CostoMat3, CostoMat_4=:CostoMat4 WHERE Tariffa =:Tar;""", rigo)
            except:
                raise PreventiviError(_("Errore nell'aggiornamento delle fasce di costo dell'articolo: %s") % str(articolo))
            finally:
                logging.debug(_("Aggiornato prezzo articolo tabella '{0}' col 'id': {1}").format('Epu', articolo.tariffa))
        return False

################################################################################
############ COMPUTO - VOCI DI COMPUTO #########################################
################################################################################
# Nota: le operazioni avvengono sempre su oggetti appartenenti alla classe Articoli

    def __select_row_id (self):
        """Extract and return the primary integer key sorted"""
        if not settings["sostituisci_id_cancellati"]:
            return None
        self.c.execute("""SELECT id FROM Computo""")
        lista = self.c.fetchall()
        if len(lista) == 0:
            return 1
        elif len(lista) == 1:
            return 2
        else:
            sorted_list = sorted(lista)
            last_tuple = sorted_list[-1]
            for i in range(len(sorted_list)):
                t = (i,)
                if t not in sorted_list:
                    return i

    def insert_articoli_computo (self, *lista_articoli):
        """Inserisce un nuovo articolo nel database"""
        result = True
        for articolo in lista_articoli:
            articolo = self.__verifica_tariffa_epu(articolo)
            # Insert a row of data or report an error
            rigo = articolo.row(codice=self.__select_row_id ())
            try:
                self.c.execute("""INSERT INTO Computo VALUES (
                               :id, :SupCat, :Cat, :SubCat, 
                               :Tar, :Quant, :PrezzoUnit, :PrezzoTot,
                               CURRENT_TIMESTAMP, :Note,
                               :TipoLav, :CatApp, :Image);""", rigo)
            except:
                raise PreventiviError(_("Errore nell'inserimento dell'articolo: %s") % str(articolo))
                result = False
            finally:
                # se le impostazioni lo consentono inserire anche gli articoli di listino memorizzati nell'articolo
                if settings["inserisci_articoli_listino_con_articoli_epu"]:
                    if len(articolo.art_listino) > 0 and result:
                         self.insert_articoli_listino (articolo.tariffa, *articolo.art_listino)
                logging.debug("Inserito articolo tabella '{0}', 'id' {1}, 'Tariffa' {2}".format(
                              'Computo', articolo.codice, articolo.tariffa))
        return result

    def update_articolo_computo (self, prymary_key, articolo, new_id=None):
        """Aggiorna i valori di un articolo"""
        result = True
        if type(prymary_key) is not int:
            raise PreventiviError(_("Errore la chiave primaria per aggiornare un articolo deve essere un numero intero"))
        if new_id is not None and type(new_id) != int:
            raise PreventiviError(_("Errore la nuova chiave primaria deve essere un numero intero"))
        rigo = articolo.row(codice=new_id)
        rigo["old_id"] = prymary_key
        rigo["PrezzoTot"] = rigo["Quant"] * rigo["PrezzoUnit"]
        try:
            self.c.execute("""UPDATE Computo SET Supercategoria=:SupCat, 
                              Categoria=:Cat, Subcategoria=:SubCat, 
                              Quantita=:Quant, Prezzo_totale=:PrezzoTot, 
                              Data=CURRENT_TIMESTAMP, Note=:Note, 
                              Tipo_lavori=:TipoLav, Cat_appalto=:CatApp, 
                              Image_art=:Image WHERE id ==:old_id;""", rigo)
            if new_id is not None:
                # rendi negative le chiavi nell'intervallo tra la nuova PRIMARY KEY e quella esistente
                if prymary_key > new_id:
                    self.c.execute("""UPDATE Computo SET id = -id-1 WHERE id >= :id AND id < :old_id;""", rigo)
                elif prymary_key < new_id:
                    self.c.execute("""UPDATE Computo SET id = -id+1 WHERE id > :old_id AND id <= :id;""", rigo)
                # effettua la modifica della primary key da modificare
                self.c.execute("""UPDATE Computo SET id=:id WHERE id ==:old_id;""", rigo)
                # dopo la modifica dell'id riportare in positivo le chiavi modificate
                self.c.execute("""UPDATE Computo SET id = -id WHERE id < 0;""")
        except:
            raise PreventiviError(_("Errore di inserimento dell'articolo: %s") % str(articolo))
            result = False
        finally:
            logging.debug(_("Aggiornato articolo tabella '{0}', 'id' {1}, 'Tariffa' {2}").format(
                          'Computo', articolo.codice, articolo.tariffa))
        return result

    def delete_articoli_computo (self, *lista_articoli):
        """Cancella un articolo esistente dal database"""
        result = True
        for articolo in lista_articoli:
            try:
                # eseguiamo l'istruzione SQL di cancellazione
                self.c.execute("""DELETE FROM Computo WHERE id =:codice;""", {"codice": articolo.codice})
            except:
                raise PreventiviError("Errore di cancellazione dell'articolo: %s" % str(articolo))
                result = False
            finally:
                logging.debug(_("Cancellato articolo tabella '{0}', 'id' {1}, 'Tariffa' {2}").format(
                              'Computo', articolo.codice, articolo.tariffa))
        return result

    def nuovo_articolo_computo (self, tariffa, descrizione_codice, descrizione_voce=None, descrizione_estesa=None, 
                                supercategoria=None, categoria=None, subcategoria=None):
        """Crea in memoria una nuova istanza articolo"""
        if supercategoria is None: supercategoria = 0
        if categoria is None: categoria = 0
        if subcategoria is None: subcategoria = 0
        articolo = ArticoloComputo(self.db, supercapitolo = 0, capitolo = 0, subcapitolo = 0, 
                                   supercategoria = supercategoria, categoria = categoria, subcategoria = subcategoria, 
                                   tariffa= tariffa, codice=None, 
                                   descrizione_codice=descrizione_codice,
                                   descrizione_voce=descrizione_voce, 
                                   descrizione_estesa=descrizione_estesa, 
                                   unita_misura=self.dati_generali["unita_misura"], 
                                   quantita=self.dati_generali["quantita"], 
                                   ricarico= self.dati_generali["ricarico"],
                                   tempo_inst=0, costo_materiali=0.0, prezzo_unitario=0.0, 
                                   sicurezza=self.dati_generali["sicurezza"], 
                                   cod_listino=str(), note=None, 
                                   tipo_lavori=0, cat_appalto=None,
                                   image_art=None) 
        #inserisci il nuovo articolo nel database
        articolo = self.__verifica_tariffa_duplicata (articolo)
        self.insert_articoli_computo (articolo)
        return articolo

    def copia_articolo_computo (self, prymary_key, copia_nomi_categorie= False,
                                copia_art_listino=True):
        """Carica in memoria un articolo esistente nel database"""
        self.c.execute("""SELECT * FROM Computo WHERE id =:id""", {"id": prymary_key})
        lista_cmp = self.c.fetchall()
        if lista_cmp == []:
            return None
        c = lista_cmp[0]
        tariffa = c[4]
        self.c.execute("""SELECT * FROM Epu WHERE Tariffa = :tariffa""", {"tariffa": tariffa})
        lista_epu = self.c.fetchall()
        e = lista_epu[0]
        articolo = ArticoloComputo(database=self.db, supercapitolo = e[1], capitolo = e[2], 
                                   subcapitolo = e[3], supercategoria = c[1], 
                                   categoria = c[2], subcategoria = c[3], tariffa= c[4], 
                                   codice=c[0], descrizione_codice=e[4], 
                                   descrizione_voce=e[5], descrizione_estesa=e[6], 
                                   unita_misura=e[7], quantita=c[5], ricarico=e[8], 
                                   tempo_inst=e[9], costo_materiali=e[10], 
                                   prezzo_unitario=e[11], sicurezza=e[12], 
                                   costo_mat1=e[15], costo_mat2=e[16], 
                                   costo_mat3=e[17], costo_mat4=e[18],
                                   cod_listino=e[13], note=c[9],
                                   tipo_lavori=c[10], cat_appalto=c[11],
                                   image_art=c[12])
        # ricavo i nomi dei capitoli e delle categorie
        if copia_nomi_categorie:
            articolo.nome_supercapitolo, articolo.nome_capitolo, articolo.nome_subcapitolo = self.get_capitoli_name_from_id (
                                         articolo.supercapitolo, articolo.capitolo, articolo.subcapitolo)
            articolo.nome_supercategoria, articolo.nome_categoria, articolo.nome_subcategoria = self.get_categorie_name_from_id (
                                         articolo.supercategoria, articolo.categoria, articolo.subcategoria)
        # copiare (se le impostazioni lo consentono) gli articoli di listino in articolo.art_listino
        if settings["inserisci_articoli_listino_con_articoli_epu"] and copia_art_listino:
            art_listino = self.get_articoli_listino(articolo.tariffa)
            if len(art_listino) > 0: 
                articolo.art_listino = art_listino
        return articolo

    def incolla_articolo_computo (self, supercategoria = None, categoria = None, 
                                  subcategoria = None, *lista_articoli):
        """Inserisce nelle nuove categorie selezionate una lista di articoli passati come argomento"""
        for articolo in lista_articoli:
            if supercategoria is not None and type(supercategoria) == int: 
                articolo.supercategoria = supercategoria
            if categoria is not None and type(categoria) == int:
                articolo.categoria = categoria
            if subcategoria is not None and type(subcategoria) == int:
                articolo.subcategoria = subcategoria
        #inserisci gli articoli nel database
        self.insert_articoli_computo (*lista_articoli)
        return lista_articoli

    def update_articoli_computo (self, supercategoria, categoria, subcategoria, 
                                 quantita, note, tipo_lavori, cat_appalto, 
                                 image_art, *lista_articoli):
        """Aggiorna simultaneamente un gruppo di articoli di computo, 
           un parametro con valore None non viene aggiornato!"""
        result = True
        for articolo in lista_articoli:
            rigo = articolo.row(codice=articolo.codice)
            if supercategoria is not None:
                rigo["SupCat"] = supercategoria
            if categoria is not None:
                rigo["Cat"] = categoria
            if subcategoria is not None:
                rigo["SubCat"] = subcategoria
            if quantita is not None:
                rigo["Quant"] = quantita
            if note is not None:
                rigo["Note"] = note
            if tipo_lavori is not None:
                rigo["TipoLav"] = tipo_lavori
            if cat_appalto is not None:
                rigo["CatApp"] = cat_appalto
            try:
                self.c.execute("""UPDATE Computo SET Supercategoria=:SupCat, 
                                  Categoria=:Cat, Subcategoria=:SubCat, 
                                  Quantita=:Quant, Note=:Note, 
                                  Tipo_lavori=:TipoLav, Cat_appalto=:CatApp, 
                                  Image_art=:Image WHERE id=:id;""", rigo)
            except:
                raise PreventiviError(_("Errore nell'aggiornamento dell'articolo di computo: %s") % str(articolo))
                result = False
            finally:
                logging.debug(_("Aggiornato articolo tabella '{0}', 'id' {1}, 'Tariffa' {2}").format(
                          'Computo', articolo.codice, articolo.tariffa))
        return result

    def update_lista_articoli (self, colonna, nuovo_campo, *lista_articoli):
        """Aggiorna i valori di una colonna di un gruppo di articoli"""
        colonne_computo = ['Supercategoria', 'Categoria', 'Subcategoria']
        colonne_epu = ['Supercapitolo', 'Capitolo', 'Subcapitolo', 'Descrizione_codice', 'Descrizione_voce', 'Descrizione_estesa', \
                       'Unita_misura', 'Ricarico', 'Tempo_inst', 'Prezzo_unitario', 'Sicurezza']
        for articolo_esistente in lista_articoli:
            dict_modifiche = {"id": articolo_esistente.codice, "tariffa": articolo_esistente.tariffa, "column": colonna, "new_value": nuovo_campo}
            if colonna in colonne_computo:
                self.c.execute("""UPDATE Computo SET :column =:new_value WHERE id =:id;""", dict_modifiche)
            elif colonna in colonne_epu:
                self.c.execute("""UPDATE Epu SET :column =:new_value WHERE id =:tariffa;""", dict_modifiche)
            else: 
                return False
            return True

    def delete_row_categoria (self, tipo, id_categoria):
        """Cancella dal database tutte le righe che appartengono alla stessa super/sub/categoria"""
        if tipo != "Supercategoria" or tipo != "Categoria" or tipo != "Subcategoria":
            return False
        try:
            # eseguiamo l'istruzione SQL di cancellazione
            self.c.execute("""DELETE FROM Computo WHERE :colonna_cat =:id""", {'colonna_cat':tipo, "id":id_categoria})
            logging.debug("Tutti gli articoli della '%s' numero '%d'cancellati dal database" % (tipo, id_categoria))
            return True
        except: 
            raise PreventiviError(_("Errore: impossibile cancellare gli articoli della '%s' numero '%d'") % (tipo, id_categoria))
            return False

################################################################################
############ LISTINO ANALISI PREZZI - operazioni base sul database Analisi #####
################################################################################

    def __verifica_codice_listino_esistente (self, articolo_listino):
        """verifica se la voce di listino esaminata ha 2 campi che corrispondono
           ad un articolo già inserite cioè tariffa e codice, se è già nel
           database, ritorna True"""
        self.c.execute("""SELECT Tariffa, Codice FROM Analisi""")
        lista_codici = self.c.fetchall()
        t = (articolo_listino.tariffa, articolo_listino.codice)
        if t in lista_codici:
            return True
        else: return False

    def __verifica_codice_listino_duplicato (self, articolo_listino):
        """verifica se, per la tariffa scelta esiste un articolo di listino che
           non abbia il codice uguale ad un altro esistente, se esiste già un
           codice con lo stesso nome cambia il nome del codice che si stà 
           inserendo aggiungendo un numero progressivo 
           (es. L00XX001 --> L00XX001 (1)"""
        self.c.execute("""SELECT Tariffa, Codice FROM Analisi""")
        lista_codici = self.c.fetchall()
        t = (articolo_listino.tariffa, articolo_listino.codice)
        n = 1
        while t in lista_codici:
             new_codice = "%s (%d)" % (articolo_listino.codice, n)
             t = (articolo_listino.tariffa, new_codice)
             n+=1
        articolo_listino.codice = t[1]
        return articolo_listino

    def insert_articoli_listino (self, tariffa, *lista_listino):
        """Inserisce uno o più articoli nel database Analisi"""
        result = True
        for articolo_listino in lista_listino:
            # se il campo tariffa non è nullo inserisce l'articolo di list x la tariffa 
            # corrispondente, altrimenti utilizza la tariffa memorizzata nel listino
            if tariffa is not None:
                articolo_listino.tariffa = tariffa
            # verifica se l'elemento di listino è già esistente, se non esiste lo inserisce
            if not self.__verifica_codice_listino_esistente(articolo_listino):
                rigo = articolo_listino.row()
                try:
                    self.c.execute("""INSERT INTO Analisi VALUES (NULL,
                                  :Tariffa, :Codice, :DesCod, 
                                  :UM, :Quantita, :PrezzoUnit, 
                                  :Sconto, :Accessori, :PrezzoTot,
                                  :Note);""", rigo)
                except:
                    raise PreventiviError(_("Errore di inserimento dell'articolo: %s in Analisi") % str(articolo_listino))
                    result = False
                finally:
                    logging.debug(_("Inserito articolo tabella '{0}' col 'Tariffa': '{1}' col 'Codice': '{2}'").format(
                                    'Analisi', articolo_listino.tariffa, articolo_listino.codice))
            # se l'elemento di listino è già esistente non lo inserisce
            else:
                logging.warning(_("Articolo tab 'Analisi' esistente per la Tariffa %s: %s \nl'articolo duplicato non verrà inserito") % (articolo_listino.tariffa, str(articolo_listino)))
                result = False
        return result

    def delete_articoli_listino (self, *lista_listino):
        """Cancella un articolo nel database Analisi"""
        result = True
        for art_listino in lista_listino:
            try:
                # eseguiamo l'istruzione SQL di cancellazione
                self.c.execute("""DELETE FROM Analisi WHERE id=:id;""", {"id": art_listino.nr})
            except:
                raise PreventiviError(_("Errore di cancellazione dell'articolo di listino: %s") % str(art_listino))
                result = False
            finally:
                logging.debug(_("Cancellato articolo tabella '{0}' col 'id': {1}").format('Analisi', art_listino.nr))
        return result

    def update_articolo_listino (self, prymary_key, articolo):
        """Aggiorna un articolo nel database Analisi"""
        result = True
        rigo = articolo.row()
        rigo["id"] = prymary_key
        try:
            self.c.execute("""UPDATE Analisi SET Codice=:Codice, Descrizione_codice=:DesCod, 
                              Unita_misura=:UM, Quantita=:Quantita, Prezzo_unitario=:PrezzoUnit, 
                              Sconto=:Sconto, Accessori=:Accessori, Prezzo_totale=:PrezzoTot, 
                              Note=:Note WHERE id=:id;""", rigo)
        except:
            raise PreventiviError(_("Errore dell'aggiornamento dell'articolo di listino: %s in Analisi") % str(articolo))
            result = False
        finally:
            logging.debug(_("Aggiornato articolo tabella '{0}' col 'id': {1}").format('Analisi', prymary_key))
        return result

    def update_lista_art_listino (self, tariffa, quantita, prezzo_unitario, 
                                  sconto, accessori, *lista_articoli): 
        """
        Aggiorna una lista di articoli articolo di listino nella tabella Analisi
        i campi 'tariffa', 'quantita', 'prezzo_unitario', 'sconto' e 'accessori' 
        possono essere una stringa (in questo caso il valore verrà sostituito
        in ogni articolo) o None (in questo caso il valore dell'articolo non 
        sarà modificato.
        """
        result = True
        for articolo in lista_articoli:
            rigo = articolo.row()
            if tariffa is not None:
                rigo["Tariffa"] = tariffa
            if quantita is not None:
                rigo["Quantita"] = quantita
            if prezzo_unitario is not None:
                rigo["PrezzoUnit"] = prezzo_unitario
            if sconto is not None:
                rigo["Sconto"] = sconto
            if accessori is not None:
                rigo["Accessori"] = accessori
            try:
                self.c.execute("""UPDATE Analisi SET Codice=:Codice, Descrizione_codice=:DesCod, 
                              Unita_misura=:UM, Quantita=:Quantita, Prezzo_unitario=:PrezzoUnit, 
                              Sconto=:Sconto, Accessori=:Accessori, Prezzo_totale=:PrezzoTot, 
                              Note=:Note WHERE id=:id;""", rigo)
            except:
                raise PreventiviError(_("Errore nell'aggiornamento dell'articolo di listino: %s in Analisi") % str(articolo))
                result = False
            finally:
                logging.debug(_("Aggiornato articolo tabella '{0}' col 'id': {1}").format('Analisi', articolo.nr))
        return result

    def copia_articolo_listino (self, prymary_key):
        """Carica in memoria un articolo di listino"""
        self.c.execute("""SELECT * FROM Analisi WHERE id=:id""", {"id":prymary_key})
        lista_lis = self.c.fetchall()
        if lista_lis == []:
            articolo = ArticoloListino(self.db, None, None, None, None, 0, 0)
            return articolo
        l = lista_lis[0]
        articolo = ArticoloListino(tariffa = l[1], codice= l[2], descrizione_codice = l[3], unita_misura = l[4], 
                                   quantita = l[5], prezzo_unitario = l[6], sconto = l[7], accessori= l[8],  nr= l[0], 
                                   note=l[10], database = self.db)
        return articolo

    def nuovo_articolo_listino (self, tariffa, codice, descrizione_codice, 
                                unita_misura = None, quantita = None, 
                                prezzo_unitario = None, sconto = None, 
                                accessori = None, verifica_duplicato=True):
        """Inserisce nel database un nuovo articolo di listino"""
        if unita_misura is None: unita_misura = "N"
        if quantita is None: quantita = 1
        if sconto is None: sconto = 0
        if accessori is None: accessori = 0
        if prezzo_unitario is None: prezzo_unitario = 0
        articolo = ArticoloListino(tariffa = tariffa, codice = codice, 
                                   descrizione_codice = descrizione_codice, unita_misura = unita_misura, 
                                   quantita = quantita, prezzo_unitario = prezzo_unitario, sconto = sconto, 
                                   accessori= accessori,  nr= None, note="", database = self.db)
        # verifica se esiste già un codice per quell'articolo e se si lo 
        # sostituisce con un codice inseribile
        if verifica_duplicato:
            articolo = self.__verifica_codice_listino_duplicato (articolo)
        self.insert_articoli_listino (tariffa, articolo)
        return articolo

    def get_articoli_listino (self, tariffa):
        """Restituisce la lista degli articoli di listino della tariffa richiesta"""
        lista_art_listino = list()
        lista = list()
        try:
            self.c.execute("""SELECT * FROM Analisi WHERE Tariffa=:tariffa""", {"tariffa":tariffa})
            lista = self.c.fetchall()
        except:
            raise PreventiviError(_("Errore impossibile recuperare gli articoli di listino relativi alla tariffa: %s") % tariffa)
        for art_listino in lista:
            articolo = self.copia_articolo_listino (art_listino[0])
            lista_art_listino.append(articolo)
        return lista_art_listino

################################################################################
############ COPIA DI ARTICOLI da - COMPUTI ESTERNI ############################
################################################################################

    def __ricava_id_e_nomi_cat(self, lista_cap):
        diz = dict()
        if len(lista_cap) == 0:
            return diz
        for cap in lista_cap:
            diz[cap[0]] = cap[1]
        return diz

    def __get_key_from_value(self, dizionario, nome):
        for key, value in dizionario.items():
            if nome == value:
                return key
        return 0

    def __get_last_key(self, dizionario):
        last_key = sorted(dizionario.keys())[-1]
        return last_key+1

    def insert_articoli_from_archivio_epu (self, insert_cmp=True, *lista_articoli_copiati):
        """Incolla Articoli da un computo o archivio esterno. Questa funzione permette di 
           preservare anche il nome dei capitoli originari."""
        # lista articoli da incollare al termine dell'operazione
        lista_articoli = list() 
        # analizza ogni articolo in lista
        for articolo in lista_articoli_copiati:
            # ricavo la lista delle categorie e dei capitoli (id, Nome, descr, aumento_prezzi)
            sup_cap, cap, sub_cap = self.capitoli_rows_list()
            # per ogni lista crea un dizionario contenente (id come chiave e Nome come valore)
            diz_sup_cap = self.__ricava_id_e_nomi_cat(sup_cap)
            diz_cap = self.__ricava_id_e_nomi_cat(cap)
            diz_sub_cap = self.__ricava_id_e_nomi_cat(sub_cap)
            # analisi nome Capitoli
            if articolo.nome_supercapitolo in diz_sup_cap.values():
                articolo.supercapitolo = self.__get_key_from_value(diz_sup_cap, articolo.nome_supercapitolo)
            else: 
                self.insert_capitoli_categorie ("Supercapitolo", articolo.nome_supercapitolo)
                articolo.supercapitolo = self.__get_last_key(diz_sup_cap)
            if articolo.nome_capitolo in diz_cap.values():
                articolo.capitolo = self.__get_key_from_value(diz_cap, articolo.nome_capitolo)
            else: 
                self.insert_capitoli_categorie ("Capitolo", articolo.nome_capitolo)
                articolo.capitolo = self.__get_last_key(diz_cap)
            if articolo.nome_subcapitolo in diz_sub_cap.values():
                articolo.subcapitolo = self.__get_key_from_value(diz_sub_cap, articolo.nome_subcapitolo)
            else: 
                self.insert_capitoli_categorie ("Subcapitolo", articolo.nome_subcapitolo)
                articolo.subcapitolo = self.__get_last_key(diz_sub_cap)
            # copia l'articolo nella lista
            lista_articoli.append(articolo)
        # inserisce la lista in computo o elenco prezzi
        if type(insert_cmp) is not bool:
            raise PreventiviError(_("'insert_cmp' può essere 'True' o 'False'"))
        if not insert_cmp:
            # effettua l'inserimento degli articoli solo in elenco prezzi
            self.insert_articoli_epu (True, *lista_articoli)
        else:
            # effettua l'inserimento degli articoli in computo e in epu
            self.insert_articoli_computo (*lista_articoli)
        return True

    def insert_articoli_from_archivio_computo (self, *lista_articoli_copiati):
        """Incolla Articoli da un computo o archivio esterno. Questa funzione permette di 
           preservare anche il nome delle categorie originarie:
           - si confronta la lista degli id e dei nomi di capitolo x verificare 
             se il nome o il numero del capitolo è già esistente
           - si creano i capitoli e le categorie necessarie
           - infine si inserisce l'articolo modificato"""
        # lista articoli da incollare al termine dell'operazione
        lista_articoli = list() 
        # analizza ogni articolo in lista
        for articolo in lista_articoli_copiati:
            # ricavo la lista delle categorie e dei capitoli (id, Nome, descr, aumento_prezzi)
            sup_cat, cat, sub_cat = self.categorie_rows_list()
            sup_cap, cap, sub_cap = self.capitoli_rows_list()
            # per ogni lista crea un dizionario contenente (id come chiave e Nome come valore)
            diz_sup_cat = self.__ricava_id_e_nomi_cat(sup_cat)
            diz_cat = self.__ricava_id_e_nomi_cat(cat)
            diz_sub_cat = self.__ricava_id_e_nomi_cat(sub_cat)
            diz_sup_cap = self.__ricava_id_e_nomi_cat(sup_cap)
            diz_cap = self.__ricava_id_e_nomi_cat(cap)
            diz_sub_cap = self.__ricava_id_e_nomi_cat(sub_cap)
            # analisi nome Categorie
            if articolo.nome_supercategoria in diz_sup_cat.values():
                articolo.supercategoria = self.__get_key_from_value(diz_sup_cat, articolo.nome_supercategoria)
            else: 
                self.insert_capitoli_categorie ("Supercategoria", articolo.nome_supercategoria)
                articolo.supercategoria = self.__get_last_key(diz_sup_cat)
            if articolo.nome_categoria in diz_cat.values():
                articolo.categoria = self.__get_key_from_value(diz_cat, articolo.nome_categoria)
            else: 
                self.insert_capitoli_categorie ("Categoria", articolo.nome_categoria)
                articolo.categoria = self.__get_last_key(diz_cat)
            if articolo.nome_subcategoria in diz_sub_cat.values():
                articolo.subcategoria = self.__get_key_from_value(diz_sub_cat, articolo.nome_subcategoria)
            else: 
                self.insert_capitoli_categorie ("Subcategoria", articolo.nome_subcategoria)
                articolo.subcategoria = self.__get_last_key(diz_sub_cat)
            # analisi nome Capitoli
            if articolo.nome_supercapitolo in diz_sup_cap.values():
                articolo.supercapitolo = self.__get_key_from_value(diz_sup_cap, articolo.nome_supercapitolo)
            else: 
                self.insert_capitoli_categorie ("Supercapitolo", articolo.nome_supercapitolo)
                articolo.supercapitolo = self.__get_last_key(diz_sup_cap)
            if articolo.nome_capitolo in diz_cap.values():
                articolo.capitolo = self.__get_key_from_value(diz_cap, articolo.nome_capitolo)
            else: 
                self.insert_capitoli_categorie ("Capitolo", articolo.nome_capitolo)
                articolo.capitolo = self.__get_last_key(diz_cap)
            if articolo.nome_subcapitolo in diz_sub_cap.values():
                articolo.subcapitolo = self.__get_key_from_value(diz_sub_cap, articolo.nome_subcapitolo)
            else: 
                self.insert_capitoli_categorie ("Subcapitolo", articolo.nome_subcapitolo)
                articolo.subcapitolo = self.__get_last_key(diz_sub_cap)
            # copia l'articolo nella lista
            lista_articoli.append(articolo)
        # effettua l'inserimento degli articoli in computo
        self.insert_articoli_computo (*lista_articoli)
        return True

################################################################################
############ PREVENTIVO - FUNZIONI RICALCOLO PREZZI ############################
################################################################################
    def ricalcolo_generale(self):
        """effettua il ricalcolo generale di tutti i prezzi per tutte le tabelle"""
        try:
            if settings["ricalcolo_automatico_prezzi"]:
                self.c.executescript("""
                /*aggiorna tutti i costi totali del listino */
                UPDATE Analisi SET Prezzo_totale = (Prezzo_unitario * Quantita *
                                                   (1- (Sconto/100))) + Accessori;
                /*aggiorna tutti i costi materiali in Epu dal listino*/
                UPDATE Epu SET Costo_materiali = (SELECT total(Prezzo_totale) 
                       FROM Analisi WHERE Tariffa = Epu.Tariffa) 
                       WHERE Tariffa IN (SELECT Tariffa FROM Analisi);
                """)
            self.c.executescript("""
                /*aggiorna tutti i prezzi unitari in epu*/
                UPDATE Epu SET Prezzo_unitario = (Costo_materiali * Ricarico) + 
                                                 (Tempo_inst/60.0 * (SELECT 
                                                  Manodopera FROM DatiGenerali)); 

                /*aggiorna tutti i prezzi totali in computo*/
                UPDATE Computo SET Prezzo_totale = Quantita * Prezzo_unitario;
                """)
        except:
            raise PreventiviError(_("Errore nel ricalcolo generale dei prezzi"))
        finally:
            logging.debug(_("Ricalcolo generale di tutti i prezzi effettuato"))
            return self.calcolo_prezzi(convert_to_string=False)

    def ricalcolo_prezzo_tariffa(self, tariffa):
        """effettua il ricalcolo del prezzo di una tariffa"""
        tar = {"tar": tariffa}
        try:
            if len(self.get_articoli_listino(tariffa)) > 0:
                self.c.execute("""UPDATE Analisi SET Prezzo_totale = (
                                  Prezzo_unitario * Quantita * (1-(Sconto/100))) + Accessori 
                                  WHERE Tariffa=:tar;""", tar)
                if not settings["ricalcolo_automatico_prezzi"]:
                    self.c.execute("""UPDATE Epu SET Costo_materiali = (SELECT 
                                      total(Prezzo_totale) FROM Analisi 
                                      WHERE Tariffa = :tar) WHERE Tariffa = :tar;""", tar)
            else:
                self.c.executescript("""
                /*aggiorna tutti i prezzi unitari in epu*/
                UPDATE Epu SET Prezzo_unitario = (Costo_materiali * Ricarico) + 
                                                 (Tempo_inst/60.0 * (SELECT 
                                                  Manodopera FROM DatiGenerali))
                           WHERE Tariffa = :tar;""", tar)
        except:
            raise PreventiviError(_("Errore nel ricalcolo del prezzo dell'articolo: %s") % tariffa)
        finally:
            logging.debug(_("Ricalcolo prezzo tariffa {0} effettuato").format(tariffa))
            return True

################################################################################
############ PREVENTIVO - AUMENTI/DIMINUZIONE DEI PREZZI DELLE SINGOLE VOCI ####
################################################################################
    def modifica_prezzi(self, perc_var, varia_costi=True, varia_tempi=False, *lista_articoli):
        """
        Aumenta o diminuisce il costo di ogni materiale della lista 
        'lista_articoli' di 'perc_var' percento; se 'varia_tempi' è True,
        anche i tempi della manodopera vengono variati.
        """
        result = True
        if type(perc_var) is not int and type(perc_var) is not float:
            raise PreventiviError("Errore la variabile 'perc_var' deve essere un 'int' oppure un 'float'")
        if type(varia_tempi) is not bool: varia_tempi = False
        if type(varia_costi) is not bool: varia_costi = True
        mul = 1+(perc_var/100)
        # crea un set contenente solo le tariffe da modificare
        lista_tariffe = set()
        for articolo in lista_articoli:
            lista_tariffe.add(articolo.tariffa)
        # per ogni tariffa da modificare esegui le operazioni SQL di modifica
        for tariffa in lista_tariffe:
            diz = {"Tar": tariffa, "multiplier":mul}
            # se sono presenti articoli di listino aumenta il loro costo
            if len(self.get_articoli_listino(tariffa)) > 0 and varia_costi and \
                                        settings["ricalcolo_automatico_prezzi"]:
                try:
                    # aggiorna il prezzo unitario e aggiorna gli accassori
                    self.c.execute("""UPDATE Analisi SET Prezzo_unitario = Prezzo_unitario * :multiplier
                                      WHERE (Tariffa=:Tar AND Prezzo_unitario != 0);""", diz)
                    self.c.execute("""UPDATE Analisi SET Accessori = Accessori * :multiplier
                                      WHERE (Tariffa=:Tar AND Accessori != 0);""", diz)
                except:
                    raise PreventiviError(_("Errore nell'aumento dei prezzi unitari delle voci di analisi dell'articolo: %s") % tariffa)
                    result = False
                finally:
                    logging.debug(_("Modifica '{0}' tabella '{1}' con tariffa '{2}'").format('Prezzo_unitario', 'Analisi', tariffa))
                    logging.debug(_("Modifica '{0}' tabella '{1}' con tariffa '{2}'").format('Accessori', 'Analisi', tariffa))
            # aumenta/diminuisci il costo unitario di ogni voce in lista
            elif varia_costi:
                try:
                    self.c.execute("""UPDATE Epu SET Costo_materiali = Costo_materiali * :multiplier 
                                      WHERE (Tariffa=:Tar AND Costo_materiali != 0);""", diz)
                except:
                    raise PreventiviError(_("Errore nell'aumento del costo materiali dell'articolo: %s") % tariffa)
                    result = False
                finally:
                    logging.debug(_("Modifica '{0}' tabella '{1}' con tariffa '{2}'").format('Costo_materiali', 'Epu', tariffa))
            # aumenta/diminuisci il tempo di installazione di ogni voce in lista
            if varia_tempi:
                try:
                    self.c.execute("""UPDATE Epu SET Tempo_inst = Tempo_inst * :multiplier
                                      WHERE (Tariffa=:Tar AND Tempo_inst != 0);""", diz)
                except:
                    raise PreventiviError(_("Errore nell'aumento del tempo di installazione dell'articolo: %s") % tariffa)
                    result = False
                finally:
                    logging.debug(_("Modifica '{0}' tabella '{1}' con tariffa '{2}'").format('Tempo_inst', 'Epu', tariffa))
            # informazioni di debug
            logging.debug(_("Fine modifica tabella {0} articolo {1}").format('Epu', tariffa))
        return result

################################################################################
############ COMPUTO - Estrazione DATI #########################################
################################################################################

    def __select_category (self, supercategoria, categoria, subcategoria, 
                           contract_type, is_chapter=False):
        """codice ripetitivo utilizzato per definire le stringhe da passare alle
        istruzioni SQL relative alle categorie nelle fuc di estrazione dati"""
        if is_chapter:
            n_su, n_c, n_sb = ("Supercapitolo", "Capitolo", "Subcapitolo")
        else:
            n_su, n_c, n_sb = ("Supercategoria", "Categoria", "Subcategoria")
        # selezione delle righe in base alla tipologia di appalto
        if contract_type == "NULL":
            cntrct_tp = """Tipo_lavori IS NULL"""
        elif contract_type is not None:
            if type(contract_type) is int:
                # assegna come valore il numero corrispondente alla lista
                contract_type = settings["default_contract_type"][contract_type]
            cntrct_tp = """Tipo_lavori = '{0}'""".format(contract_type)
        else:
            cntrct_tp = """1 = 1"""
        # selezione delle righe in base alla categoria
        if supercategoria is not None and type(supercategoria) is int:
            sup_cat = """{0} = {1}""".format(n_su, supercategoria)
        else: 
            sup_cat = """1 = 1"""
        if categoria is not None and type(categoria) is int:
            cat = """{0} = {1}""".format(n_c, categoria)
        else: 
            cat = """1 = 1"""
        if subcategoria is not None and type(subcategoria) is int:
            sub_cat = """{0} = {1}""".format(n_sb, subcategoria)
        else: 
            sub_cat = """1 = 1"""
        return sup_cat, cat, sub_cat, cntrct_tp


    def get_computo_ids (self):
        """Estrae dal database una lista di chiavi univoche di codici valida"""
        self.c.execute("""SELECT id FROM Computo ORDER BY id;""")
        return self.c.fetchall()

    def dati_generali_list (self):
        """Estrae dal database i dati generali"""
        self.c.execute("""SELECT * FROM DatiGenerali;""")
        lista = self.c.fetchall()
        if len(lista) == 0:
            lista = [(self.dati_generali["ricarico"], self.dati_generali["manodopera"],
                     self.dati_generali["sicurezza"], self.dati_generali["valuta"],
                     self.dati_generali["nome"], self.dati_generali["indirizzo"],
                     self.dati_generali["comune"], self.dati_generali["provincia"],
                     self.dati_generali["cliente"], self.dati_generali["redattore"]) ]
        return lista[0]

    def computo_count_rows (self, supercategoria=None, categoria=None, subcategoria=None, 
                            count_if_quantity_not_0=False, contract_type=None):
        """Estrae il numero di elementi presenti nel computo per ogni categoria"""
        # selezione delle righe in base alla categoria e alla tipologia di appalto
        sup_cat, cat, sub_cat, cntrct_tp = self.__select_category(supercategoria,
                                           categoria, subcategoria, contract_type)
        # selezione solo delle righe con quantità (in computo) diversa da zero  
        if count_if_quantity_not_0:
            quant = """Quantita != 0"""
        else: 
            quant = """1 = 1"""
        self.c.execute("""SELECT id FROM Computo WHERE %s AND %s AND %s AND %s AND %s;""" % (sup_cat, cat, sub_cat, quant, cntrct_tp))
        lista = self.c.fetchall()
        return len(lista)

    def epu_count_rows (self, supercapitolo=None, capitolo=None, subcapitolo=None, 
                        count_if_computo_not_0 = False, contract_type=None):
        """Estrae il numero di elementi presenti nell'epu per ogni capitolo"""
        lista = self.epu_rows_list (supercapitolo, capitolo, subcapitolo, 
                                    count_if_computo_not_0, contract_type)
        return len(lista)

    def lavorazioni_count_rows (self, supercapitolo=None, capitolo=None, 
                                subcapitolo=None, count_if_computo_not_0=False, 
                                contract_type=None):
        """Estrae il numero di elementi presenti nella lista lavorazioni per ogni capitolo"""
        lista = self.epu_rows_list (supercapitolo, capitolo, subcapitolo, 
                                    count_if_computo_not_0, contract_type=None)
        return len(lista)

    def computo_rows_list (self, supercategoria=None , categoria=None, 
                                 subcategoria=None, list_if_quantity_not_0=False,
                                 column=None, search_key=None,
                                 convert_to_string=False,
                                 contract_type=None):
        """
        Estrae dal database le righe di computo richieste. Ritorna una lista in 
        questo ordine: 
        [Cod, Tariffa, Desc_cod, Desc_voce, UM, Quant, Ric, Temp, Pr_Un, Pr_Tot]          
        """
        # selezione delle righe in base alla categoria e alla tipologia di appalto
        sup_cat, cat, sub_cat, cntrct_tp = self.__select_category(supercategoria,
                                           categoria, subcategoria, contract_type)
        # selezione solo delle righe con quantità (in computo) diversa da zero
        if list_if_quantity_not_0:
            quant = """Quantita != 0"""
        else: 
            quant = """1 = 1"""
        # verifica problemi di SQL injection per il campo di ricerca
        if search_key is not None and search_key is str:
            search_key = search_key.replace("'", " ").replace("=", ' ')
        if column is None and search_key is None:
            self.c.execute("""SELECT * FROM Computo WHERE {0} AND {1} AND {2} AND {3} AND {4} ORDER BY id;""".format(
                           cntrct_tp, sup_cat, cat, sub_cat, quant))
        else:
            self.c.execute("""SELECT * FROM Computo WHERE {0} AND {1} AND {2} AND {3} AND {4} AND {5} LIKE '%%{6}%%' ORDER BY id;""".format(
                           cntrct_tp, sup_cat, cat, sub_cat, quant, column, search_key))
        lista_voci = self.c.fetchall()
        lista = list()
        for voce in lista_voci:
            Tariffa = voce[4]
            self.c.execute("""SELECT * FROM Epu WHERE Tariffa = '%s';""" % Tariffa)
            lista_epu = self.c.fetchall()
            epu = lista_epu[0]
            Desc_cod = epu[4]
            Desc_voce = epu[5]
            UM = epu[7]
            Ric = epu[8]
            Temp = epu[9]
            Costo_mat = epu[10]
            Pr_Un = epu[11]
            Sic = epu[12]
            Cod = voce[0]
            Quant = voce[5]
            Pr_Tot = voce[7]
            Note = voce[9]
            Data = voce[8]
            Tipo_lav = voce[10]
            Cat_app = voce[11]
            if convert_to_string:
                Quant = self.__convert_number_to_string(Quant, round_to_integer=True)
                Costo_mat = self.__convert_number_to_string(Costo_mat)
                Ric = self.__convert_number_to_string(Ric, decimals= 3)
                Temp = self.__convert_number_to_string(Temp, round_to_integer=True)
                Sic = self.__convert_number_to_string(Sic, decimals= 3)
                Pr_Un = self.__convert_number_to_string(Pr_Un)
                Pr_Tot = self.__convert_number_to_string(Pr_Tot)
            voce =[Cod, Tariffa, Desc_cod, Desc_voce, UM, Quant, Costo_mat, Ric, 
                   Temp, Sic, Pr_Un, Pr_Tot, Note, Tipo_lav, Cat_app, Data]
            lista.append(voce)
        return lista

    def epu_rows_list (self, supercapitolo= None, capitolo=None, subcapitolo=None, 
                       list_if_computo_not_0 = False, convert_to_string=False, 
                       contract_type=None):
        """Estrae tutte le righe dal database dalla tavola EPU"""
        # selezione delle righe in base al capitolo e alla tipologia di appalto
        sup_cap, cap, sub_cap, cntrct_tp = self.__select_category(supercapitolo, 
                                                capitolo, subcapitolo, 
                                                contract_type, is_chapter=True)
        if list_if_computo_not_0:
            self.c.execute("""SELECT * FROM Epu WHERE %s AND %s AND %s AND Tariffa IN 
                                (SELECT Tariffa FROM Computo WHERE Quantita != 0) 
                                ORDER BY Tariffa;""" % (sup_cap, cap, sub_cap)) 
                                #FEATURE: selezione della lista del computo col  'contract_type' ???
        elif supercapitolo is None and capitolo is None and subcapitolo is None:
            self.c.execute("""SELECT * FROM Epu ORDER BY Tariffa;""")
        else:
            self.c.execute("""SELECT * FROM Epu WHERE %s AND %s AND %s ORDER BY Tariffa;""" % (sup_cap, cap, sub_cap))
        lista = self.c.fetchall()
        if convert_to_string:
            list_new=list()
            for articolo in lista:
                art= list(articolo)
                art[8] = self.__convert_number_to_string(art[8], decimals=3)
                art[9] = self.__convert_number_to_string(art[9], round_to_integer=True)
                art[10] = self.__convert_number_to_string(art[10])
                art[11] = self.__convert_number_to_string(art[11])
                art[12] = self.__convert_number_to_string(art[12])
                art[15] = self.__convert_number_to_string(art[15])
                art[16] = self.__convert_number_to_string(art[16])
                art[17] = self.__convert_number_to_string(art[17])
                art[18] = self.__convert_number_to_string(art[18])
                list_new.append(art)
            lista=list_new
        return lista

    def listino_rows_list (self, tariffa=None, convert_to_string=False):
        """Estrae dal database le righe di listino richieste"""
        lista = list()
        if tariffa is None:
            self.c.execute("""SELECT * FROM Analisi ORDER BY Tariffa;""")
            lista = self.c.fetchall()
        else: 
             self.c.execute("""SELECT * FROM Analisi WHERE Tariffa=:Tariffa ORDER BY id;""", {"Tariffa":tariffa})
             lista = self.c.fetchall()
        if convert_to_string:
            list_new=list()
            for articolo in lista:
                art= list(articolo)
                #id, Tariffa, Codice, Descrizione_codice, Unita_misura, Quantita, Prezzo_unitario, Sconto, Accessori, Prezzo_totale, Note
                art[5] = self.__convert_number_to_string(art[5], round_to_integer=True)
                art[6] = self.__convert_number_to_string(art[6])
                art[7] = self.__convert_number_to_string(art[7])
                art[8] = self.__convert_number_to_string(art[8])
                art[9] = self.__convert_number_to_string(art[9])
                list_new.append(art)
            lista=list_new
        return lista

    def lavorazioni_rows_list (self, supercapitolo= None, capitolo=None, 
                               subcapitolo=None, list_if_quantity_not_0 = False,
                               column=None, search_key=None, 
                               convert_to_string=False, contract_type=None):
        """Estrae dal database una lista delle LAVORAZIONI"""
        # selezione delle righe in base al capitolo e alla tipologia di appalto
        sup_cap, cap, sub_cap, cntrct_tp = self.__select_category(supercapitolo, 
                                                capitolo, subcapitolo, 
                                                contract_type, is_chapter=True)
        # verifica problemi di SQL injection per il campo di ricerca
        if search_key is not None and search_key is str:
            search_key = search_key.replace("'", " ").replace("=", " ")
        if column is None and search_key is None:
            self.c.execute("""SELECT Tariffa, Descrizione_codice, Descrizione_voce, Descrizione_estesa, 
                           Unita_misura, Ricarico, Tempo_inst, Costo_materiali, Prezzo_unitario, Sicurezza, Note 
                           FROM Epu WHERE %s AND %s AND %s ORDER BY Tariffa;""" % (sup_cap, cap, sub_cap))
        else:
            self.c.execute("""SELECT Tariffa, Descrizione_codice, Descrizione_voce, Descrizione_estesa, 
                           Unita_misura, Ricarico, Tempo_inst, Costo_materiali, Prezzo_unitario, Sicurezza, Note 
                           FROM Epu WHERE %s AND %s AND %s AND %s LIKE '%%%s%%' ORDER BY Tariffa;""" % (sup_cap, cap, sub_cap, column, search_key))
        lista_epu = self.c.fetchall()
        self.c.execute("""SELECT Tariffa, Quantita, Prezzo_totale FROM Computo ORDER BY Tariffa;""") 
                       #FEATURE: selezione della lista del computo col  'contract_type' ???
        lista_computo = self.c.fetchall()
        tariffe_quantita = dict()
        for art in lista_computo:
            if art[0] not in tariffe_quantita.keys():
                tariffe_quantita[art[0]] = [art[1], art[2]]
            else: 
                tariffe_quantita[art[0]] = [art[1]+tariffe_quantita[art[0]][0], art[2]+tariffe_quantita[art[0]][1]]
        lista = list()
        for e in lista_epu:
            Tariffa = e[0]
            Desc_cod = e[1]
            Desc_voce = e[2]
            Desc_estesa = e[3]
            UM = e[4]
            Ric = e[5]
            Temp = e[6]
            Costo_mat = e[7]
            Pr_Un = e[8]
            Sic = e[9]
            if e[0] in tariffe_quantita:
                Quant = tariffe_quantita[e[0]][0]
                Pr_Tot = tariffe_quantita[e[0]][1]
            else: 
                Quant, Pr_Tot = (0,0)
            Note = e[10]
            if convert_to_string:
                Costo_mat = self.__convert_number_to_string(Costo_mat)
                Ric = self.__convert_number_to_string(Ric, decimals= 3)
                Temp = self.__convert_number_to_string(Temp, round_to_integer=True)
                Sic = self.__convert_number_to_string(Sic, decimals= 3)
                Quant = self.__convert_number_to_string(Quant, round_to_integer=True)
                Pr_Un = self.__convert_number_to_string(Pr_Un)
                Pr_Tot = self.__convert_number_to_string(Pr_Tot)
            # crea l'articolo
            voce =[Tariffa, Desc_cod, Desc_voce, Desc_estesa, UM, Costo_mat, Ric, Temp, Sic, Quant, Pr_Un, Pr_Tot, Note]
            if list_if_quantity_not_0:
                if int(Quant) != 0:
                    lista.append(voce)
            else:
                lista.append(voce)
        return lista

    def categorie_rows_list (self):
        """Estrae dal database una lista delle categorie di computo"""
        self.c.execute("""SELECT * FROM Supercategoria ORDER BY id;""")
        sup = self.c.fetchall()
        self.c.execute("""SELECT * FROM Categoria ORDER BY id;""")
        cat = self.c.fetchall()
        self.c.execute("""SELECT * FROM Subcategoria ORDER BY id;""")
        sub = self.c.fetchall()
        return sup, cat, sub

    def capitoli_rows_list (self):
        """Estrae dal database una lista dei capitoli di epu"""
        self.c.execute("""SELECT * FROM Supercapitolo ORDER BY id;""")
        sup = self.c.fetchall()
        self.c.execute("""SELECT * FROM Capitolo ORDER BY id;""")
        cap = self.c.fetchall()
        self.c.execute("""SELECT * FROM Subcapitolo ORDER BY id;""")
        sub = self.c.fetchall()
        return sup, cap, sub

    def get_categorie_name_from_id (self, id_sup, id_cat, id_sub):
        """Estrae dal database i nomi delle categorie di computo richieste"""
        self.c.execute("""SELECT Nome FROM Supercategoria WHERE id=:id;""", {"id":id_sup})
        ls = self.c.fetchall()    
        if len(ls) == 0: name_sup = None
        else: name_sup = ls[0][0]
        self.c.execute("""SELECT Nome FROM Categoria WHERE id=:id;""", {"id":id_cat})
        ls = self.c.fetchall()    
        if len(ls) == 0: name_cat = None
        else: name_cat = ls[0][0]
        self.c.execute("""SELECT Nome FROM Subcategoria WHERE id=:id;""", {"id":id_sub})
        ls = self.c.fetchall()    
        if len(ls) == 0: name_sub = None
        else: name_sub = ls[0][0]
        return name_sup, name_cat, name_sub

    def get_capitoli_name_from_id (self, id_sup, id_cap, id_sub):
        """Estrae dal database i nomi dei capitoli di epu richiesti"""
        self.c.execute("""SELECT Nome FROM Supercapitolo WHERE id=:id;""", {"id":id_sup})
        ls = self.c.fetchall()    
        if len(ls) == 0: name_sup = None
        else: name_sup = ls[0][0]
        self.c.execute("""SELECT Nome FROM Capitolo WHERE id=:id;""", {"id":id_cap})
        ls = self.c.fetchall()    
        if len(ls) == 0: name_cap = None
        else: name_cap = ls[0][0]
        self.c.execute("""SELECT Nome FROM Subcapitolo WHERE id=:id;""", {"id":id_sub})
        ls = self.c.fetchall()    
        if len(ls) == 0: name_sub = None
        else: name_sub = ls[0][0]
        return name_sup, name_cap, name_sub

    def list_contract_type (self):
        """Crea una lista di di tutte le tipologie contrattuali (col. Tipo_lavori) 
        presenti nella tabella Computo."""
        # recuperiamo la lista delle righe selezionate
        self.c.execute("""SELECT Tipo_lavori FROM Computo;""")
        lista = set(self.c.fetchall())
        # se esistono righe con il campo 'Tipo_lavori' vuoto (valore di ritorno
        # 'None') sostituisce quel valore con il parametro NULL, adeguato a effettuare
        # le query SQL per estrarre i valori con campo 'Tipo_lavori' vuoto.
        if (None,) in lista:
            lista = lista ^ set([('NULL',), (None,)])
        return lista

################################################################################
############ COMPUTO - Calcolo PREZZI ##########################################
################################################################################
    def calcolo_prezzi (self, supercategoria= None , categoria=None, 
                        subcategoria=None, contract_type=None, 
                        convert_to_string=True):
        """Estrae dal database e calcola i totali parziali per categoria o anche il totale generale"""
        # selezione delle righe in base alla categoria e alla tipologia di appalto
        sup_cat, cat, sub_cat, cntrct_tp = self.__select_category(supercategoria,
                                           categoria, subcategoria, contract_type)
        # estrae e calcola tutti i prezzi
        if supercategoria is None and categoria is None and subcategoria is None and contract_type is None:
            self.c.execute("""SELECT Prezzo_totale FROM Computo;""")
        else:
            self.c.execute("""SELECT Prezzo_totale FROM Computo WHERE {0} AND {1} AND {2} AND {3};""".format(sup_cat, cat, sub_cat, cntrct_tp))
        lista = self.c.fetchall()
        totale_prezzi = 0
        for prezzo in lista:
            totale_prezzi += round(prezzo[0], settings["decimals"])
        if convert_to_string:
            totale_prezzi = self.__convert_number_to_string(totale_prezzi)
        return totale_prezzi

    def calcolo_prezzi_lavorazioni (self, supercapitolo=None , capitolo=None, 
                                    subcapitolo=None, contract_type=None, 
                                    convert_to_string=True):
        """Estrae dal database e calcola i total generale"""
        # selezione delle righe in base alla categoria e alla tipologia di appalto
        sup_cap, cap, sub_cap, cntrct_tp = self.__select_category(supercapitolo,
                          capitolo, subcapitolo, contract_type, is_chapter=True)
        self.c.execute("""SELECT Tariffa, Prezzo_totale FROM Computo;""")
        lista_computo = self.c.fetchall()
        if supercapitolo is None and capitolo is None and subcapitolo is None:
            self.c.execute("""SELECT Tariffa, Prezzo_unitario FROM Epu;""")
        else:
            self.c.execute("""SELECT Tariffa, Prezzo_unitario FROM Epu WHERE {0} AND {1} AND {2};""".format(sup_cap, cap, sub_cap))
        lista_epu = self.c.fetchall()
        diz_tariffe = dict()
        totale_prezzi = 0
        for tariffa, prezzo_unitario in lista_epu:
            diz_tariffe[tariffa] = prezzo_unitario
        for art in lista_computo:
            if art[0] in diz_tariffe.keys():
                totale_prezzi += round(art[1], settings["decimals"])
        if convert_to_string:
            totale_prezzi = self.__convert_number_to_string(totale_prezzi)
        return totale_prezzi

    def calcolo_incidenza_manodopera (self, supercategoria=None , categoria=None, 
                                      subcategoria=None, contract_type=None, 
                                      convert_to_string=True, min_to_hour=True):
        """Calcola il valore totale dei min e il costo della manodopera per un gruppo di articoli"""
        # selezione delle righe in base alla categoria e alla tipologia di appalto
        sup_cat, cat, sub_cat, cntrct_tp = self.__select_category(supercategoria,
                                           categoria, subcategoria, contract_type)
        # estrae la quantità e la tariffa
        self.c.execute("""SELECT Tariffa, Quantita FROM Computo WHERE {0} AND {1} AND {2} AND {3};""".format(
                                              sup_cat, cat, sub_cat, cntrct_tp))
        lista_tar_quant = self.c.fetchall()
        # ricava il costo della manodopera
        mo = self.dati_generali_list()[1]
        # il tempo di installazione di tutte le tariffe
        self.c.execute("""SELECT Tariffa, Tempo_inst FROM Epu;""")
        lista_tempi = self.c.fetchall()
        diz_tempi = dict(lista_tempi)
        # calcola il prezzo totale della manodopera e i minuti totali di installazione
        min_tot = 0
        for tariffa, quant in lista_tar_quant:
            tempo = diz_tempi[tariffa]
            min_tot += tempo * quant
        costo_tot_mo = min_tot/60 * mo
        # calcola l'incidenza percentuale
        costo_totale = self.calcolo_prezzi (supercategoria=supercategoria, 
                                            categoria=categoria, 
                                            subcategoria=subcategoria,
                                            contract_type=contract_type,
                                            convert_to_string=False)
        try:
            inc = 100 * costo_tot_mo/costo_totale
        except ZeroDivisionError:
            inc = 0
        # converte i minuti in ore di installazione
        if min_to_hour: tmp_tot = min_tot/60
        else: tmp_tot = min_tot
        # converte i valori numerici in stringhe
        if convert_to_string:
            costo_tot_mo = self.__convert_number_to_string(costo_tot_mo)
            tmp_tot = self.__convert_number_to_string(tmp_tot)
            inc = self.__convert_number_to_string(inc)
        return (tmp_tot, costo_tot_mo, inc)

    def calcolo_inc_manodopera_lavorazioni (self, supercapitolo=None , capitolo=None, 
                                            subcapitolo=None, min_to_hour=True,
                                            convert_to_string=True):
        """Calcola il valore totale dei min e il costo della manodopera per un gruppo di articoli"""
        # selezione delle righe in base alla categoria e alla tipologia di appalto
        sup_cat, cat, sub_cat, cntrct_tp = self.__select_category(supercapitolo, 
                                   capitolo, subcapitolo, None, is_chapter=True)
        # estrae la quantità e la tariffa di tutte le voci
        self.c.execute("""SELECT Tariffa, Quantita FROM Computo""")
        lista_tar_quant = self.c.fetchall()
        # ricava il costo della manodopera
        mo = self.dati_generali_list()[1]
        # il tempo di installazione di tutte le tariffe
        self.c.execute("""SELECT Tariffa, Tempo_inst FROM Epu WHERE {0} AND {1} AND {2};""".format(
                       sup_cat, cat, sub_cat))
        lista_tempi = self.c.fetchall()
        # calcola il prezzo totale della manodopera e i minuti totali di installazione
        min_tot = 0
        for tariffa, tempo in lista_tempi:
            for tar, quant in lista_tar_quant:
                if tariffa == tar:
                    min_tot += tempo * quant
        costo_tot_mo = min_tot/60 * mo
        # calcola l'incidenza percentuale
        costo_totale = self.calcolo_prezzi_lavorazioni (supercapitolo= supercapitolo, 
                                            capitolo=capitolo, 
                                            subcapitolo=subcapitolo, 
                                            convert_to_string=False)
        try:
            inc = 100 * costo_tot_mo/costo_totale
        except ZeroDivisionError:
            inc = 0
        # converte i minuti in ore di installazione
        if min_to_hour: tmp_tot = min_tot/60
        else: tmp_tot = min_tot
        # converte i valori numerici in stringhe
        if convert_to_string:
            costo_tot_mo = self.__convert_number_to_string(costo_tot_mo)
            tmp_tot = self.__convert_number_to_string(tmp_tot)
            inc = self.__convert_number_to_string(inc)
        return (tmp_tot, costo_tot_mo, inc)

    def calcolo_incidenza_sicurezza (self, supercategoria=None, categoria=None, 
                                     subcategoria=None, contract_type=None, 
                                     convert_to_string=True):
        """Calcola il valore totale dei min e il costo della manodopera per un gruppo di articoli"""
        # selezione delle righe in base alla categoria e alla tipologia di appalto
        sup_cat, cat, sub_cat, cntrct_tp = self.__select_category(supercategoria, 
                                          categoria, subcategoria, contract_type)
        # estrae la quantità e la tariffa
        self.c.execute("""SELECT Tariffa, Prezzo_totale FROM Computo WHERE {0} AND {1} AND {2} AND {3};""".format(
                                              sup_cat, cat, sub_cat, cntrct_tp))
        lista_tar_pr = self.c.fetchall()
        # crea un dizionario con la sicurezza per tutte le tariffe
        self.c.execute("""SELECT Tariffa, Sicurezza FROM Epu;""")
        diz_sic = dict(self.c.fetchall())
        # calcola il prezzo totale della sicurezza
        sic_tot = 0
        for tariffa, pr_tot in lista_tar_pr:
            sic = diz_sic.get(tariffa)
            sic_tot += pr_tot * sic/100
        # calcola l'incidenza percentuale
        costo_totale = self.calcolo_prezzi (supercategoria= supercategoria, 
                                            categoria=categoria, 
                                            subcategoria=subcategoria, 
                                            convert_to_string=False)
        try:
            inc = 100 * sic_tot/costo_totale
        except ZeroDivisionError:
            inc = 0
        # converte i valori numerici in stringhe
        if convert_to_string:
            costo_totale = self.__convert_number_to_string(costo_totale)
            sic_tot = self.__convert_number_to_string(sic_tot)
            inc = self.__convert_number_to_string(inc)
        return (sic_tot, costo_totale, inc)

    def calcolo_incidenza_sic_lavorazioni (self, supercapitolo= None , capitolo=None, 
                                           subcapitolo=None, contract_type=None, 
                                           convert_to_string=True):
        """Calcola il valore totale dei min e il costo della manodopera per un gruppo di articoli"""
        sup_cat, cat, sub_cat, cntrct_tp = self.__select_category(supercapitolo, 
                                   capitolo, subcapitolo, None, is_chapter=True)
        # estrae la quantità e la tariffa
        self.c.execute("""SELECT Tariffa, Prezzo_totale FROM Computo;""")
        lista_tar_pr = self.c.fetchall()
        # crea un dizionario con la sicurezza e le tariffe
        self.c.execute("""SELECT Tariffa, Sicurezza FROM Epu WHERE %s AND %s AND %s;""" % (sup_cat, cat, sub_cat))
        diz_sic = dict(self.c.fetchall())
        # calcola il prezzo totale della sicurezza
        sic_tot = 0
        for tariffa, sic in diz_sic.items():
            for tar, pr_tot in lista_tar_pr:
                if tar == tariffa:
                    sic_tot += pr_tot * sic/100
        # calcola l'incidenza percentuale
        costo_totale = self.calcolo_prezzi_lavorazioni (supercapitolo=supercapitolo, 
                                                        capitolo=capitolo, 
                                                        subcapitolo=subcapitolo, 
                                                        convert_to_string=False)
        try:
            inc = 100 * sic_tot/costo_totale
        except ZeroDivisionError:
            inc = 0
        # converte i valori numerici in stringhe
        if convert_to_string:
            costo_totale = self.__convert_number_to_string(costo_totale)
            sic_tot = self.__convert_number_to_string(sic_tot)
            inc = self.__convert_number_to_string(inc)
        return (sic_tot, costo_totale, inc)

    def search_tables (self, table, column, search_key, order):
        """Esegue un'istruzione SQL di ricerca estrazione sul database EPU"""
        #dict_par = {"table":table, "column":column, "search_key":search_key, "order":order}
        lista = list()
        try:
            # eseguiamo l'istruzione SQL
            self.c.execute("""SELECT * FROM %s WHERE %s LIKE '%%%s%%' ORDER BY %s;""" % (table, column, search_key, order))
            # recuperiamo la lista delle righe selezionate
            lista = self.c.fetchall()
        except: 
            raise PreventiviError("Errore di esecuzione istruzione SQL: %s" % istruzione)
        finally:
            return lista

    def get_costi_mat_extra (self, tariffa):
        """Estrae dal database dalla tavola EPU i costi dei materiali opzionali 1...4"""
        self.c.execute("""SELECT CostoMat_1, CostoMat_2, CostoMat_3, CostoMat_4 FROM Epu 
                             WHERE Tariffa=:tariffa;""", {"tariffa":tariffa})
        lista = self.c.fetchall()
        return lista

    def __convert_number_to_string(self, number, decimals= None, 
                                   round_to_integer = False, is_courrency=False,
                                   international=False):
        """
        Funzione per convertire i valori numerici in stringhe con il numero di 
        decimali stabilito da DEFAULT_SETTINGS['decimals'] secondo gli standard 
        internazionali definiti dal modulo locale
        """
        if number is None:
            return number
        if decimals is None:
            decimals = settings["decimals"]
        # se il parametro is_courrency è vero trasforma il numero in una valuta
        if is_courrency:
            number = round(number, decimals)
            locale.currency(number, symbol=True, grouping=True, 
                            international=international)
        # se i decimali del numero sono 0 trasformalo in un intero
        if type(number) is float and round_to_integer:
            if number.is_integer():
                number = int(number)
        # se il numero è con virgola mobile arrotonda a n. 'decimals' di decimali
        if type(number) is float:
            number = round(number, decimals)
            text = locale.format("%.{0}f".format(decimals), number, grouping=True)
        elif type(number) is int:
            text = locale.format("%d", number, grouping=True)
        else:
            # se il valore numerico non è riconosciuto solleva un'eccezione
            raise PreventiviError("Il valore '{0}' number non è un valore numerico".format(number))
        return text
################################################################################
############ COMPUTO - STAMPA DATI #############################################
################################################################################

    def print_computo (self, lista_supercategorie=None, lista_categorie=None, lista_subcategorie=None, 
                       stampa_descr_breve = True, stampa_descr_estesa=True, stampa_note = True,
                       print_only_if_not_0 = False, print_without_chapter = False, 
                       print_analisi= False):
        """Crea una lista di tuple in cui ciascun elemento rappresenta una colonna da stampare per 
        il computo metrico: 'Tariffa - Descrizione - UM - Quantita - Pr.Unit. - Pr. Tot."""
        # se non ho una lista di articoli da stampare usa tutte le categorie
        lista_sup, lista_cat, lista_sub = self.categorie_rows_list()
        if lista_supercategorie is not None and len(lista_supercategorie) > 0:
            lista_sup = lista_supercategorie
        if lista_categorie is not None and len(lista_categorie) > 0:
            lista_cat = lista_categorie
        if lista_subcategorie is not None and len(lista_subcategorie) > 0:
            lista_sub = lista_subcategorie
        # ricava la lista delle tipologie contrattuali da stampare presenti in computo
        lista_type = self.list_contract_type()
        if len(lista_type) == 0: lista_type = [(None,)]
        # inizializzo la lista degli articoli da stampare, i dizionari ecc.
        diz_desc_estese = dict()
        lista_articoli_stampare = list()
        # ricava la lista degli elementi da stampare
        lista = self.__get_list_to_print (self.computo_rows_list, 
                       self.computo_count_rows, self.calcolo_prezzi,
                       lista_type, lista_sup, lista_cat, lista_sub,
                       print_only_if_not_0 = print_only_if_not_0, 
                       print_without_chapter = print_without_chapter)
        # ricava il dizionario delle descrizioni estese
        if stampa_descr_estesa:
            lista_epu = self.epu_rows_list()
            for articolo in lista_epu:
                tariffa = articolo[0]
                descr_estesa = articolo[6]
                diz_desc_estese[tariffa] = descr_estesa
        # dalla lista ricavata estrai una lista di tuple ciascuna delle quali contiene le seguenti colonne:
        # articolo da inserire: [0 - Cod + Tariffa, 1-Descrizione (desc.codice + estesa + voce + note), 2-UM, 3-Quant, 4-Pr_Un, 5-Pr_Tot, 6-Costo_mat, 7-Ric, 8-Temp]
        for articolo in lista:
            # articolo =[0-Cod, 1-Tariffa, 2-Desc_cod, 3-Desc_voce, 4-UM, 5-Quant, 6-Costo_mat, 7-Ric, 8-Temp, 9-Sic, 10-Pr_Un, 11-Pr_Tot, 12-Note, 13-Data]
            code = str(articolo[0]) 
            if code.startswith("$$$"):
                tariffa = str(articolo[1])
            else:
                tariffa = "n.{0}\n{1}".format(articolo[0], articolo[1])
            descr_codice = articolo[2]
            um, quantita, pr_un, pr_tot = articolo[4], articolo[5], articolo[10], articolo[11]
            costo, ric, tempo, sic = articolo[6], articolo[7], articolo[8], articolo[9]
            # verifica se stampare la descrizione estesa
            descr_estesa = str()
            if stampa_descr_estesa:
                if articolo[1] in diz_desc_estese: 
                    descr_estesa = diz_desc_estese[articolo[1]]
            # verifica se stampare la descrizione breve
            descr_breve = str()
            if stampa_descr_breve: descr_breve = articolo[3]
            # verifica se stampare le note
            note = str()
            if stampa_note: note = articolo[12]
            # crea la tupla della colonna Descrizione
            Descrizione = (descr_codice, descr_estesa, descr_breve, note)
            # crea la tupla con le colonne del computo
            tupla = (code, tariffa, Descrizione, um, quantita, pr_un, pr_tot, costo, ric, tempo, sic)
            lista_articoli_stampare.append(tupla)
            # verifica se stampare gli elementi di listino (per l'analisi prezzi)
            if print_analisi:
                lista_articoli_stampare = self.__print_analisi (articolo[1], tupla, lista_articoli_stampare)
        return lista_articoli_stampare


    def print_lavorazioni(self, lista_supercapitoli=None, lista_capitoli=None, lista_subcapitoli=None, 
                       stampa_descr_breve = True, stampa_descr_estesa = True, stampa_note = True, 
                       print_only_if_not_0 = False, print_without_chapter = False,
                       print_analisi = False):
        """Crea una lista di tuple in cui ciascun elemento rappresenta una colonna da stampare per 
        la lista lavorazioni: 'Tariffa - Descrizione - UM - Quantita - Pr.Unit. - Pr. Tot."""
        # inizializzo la lista degli articoli da stampare, i dizionari ecc.
        lista_articoli_stampare = list()
        # se non ho una lista di capitoli da stampare usa tutti i capitoli
        lista_sup, lista_cap, lista_sub = self.capitoli_rows_list()
        if lista_supercapitoli is not None and len(lista_supercapitoli) > 0:
            lista_sup = lista_supercapitoli
        if lista_capitoli is not None and len(lista_capitoli) > 0:
            lista_cap = lista_capitoli
        if lista_subcapitoli is not None and len(lista_subcapitoli) > 0:
            lista_sub = lista_subcapitoli
        # ricava la lista delle tipologie contrattuali da stampare
        lista_type = [(None,)]
        # ricava la lista degli elementi da stampare
        lista = self.__get_list_to_print(self.lavorazioni_rows_list, 
                       self.lavorazioni_count_rows, 
                       self.calcolo_prezzi_lavorazioni,
                       lista_type, lista_sup, lista_cap, lista_sub,
                       print_only_if_not_0=print_only_if_not_0,
                       print_without_chapter=print_without_chapter)
        # dalla lista ricavata estrai una lista di tuple ciascuna delle quali contiene le seguenti colonne:
        # articolo da inserire: [0-Cod + Tariffa, 1-Descrizione (desc.codice + estesa + voce + note), 2-UM, 3-Quant, 4-Pr_Un, 5-Pr_Tot, 6-Costo_mat, 7-Ric, 8-Temp]
        n=0
        for articolo in lista:
            # 0-Tariffa, 1-Desc_cod, 2-Desc_voce, 3-Desc_estesa, 4-UM, 5-Costo_mat, 6-Ric, 7-Temp, 8-Sic, 9-Quant, 10-Pr_Un, 11-Pr_Tot, 12-Note
            # unità di misura, ricarichi tempi e altro ...
            um, costo, ric, tempo, sic, quantita, pr_un, pr_tot = (
                                    articolo[4], articolo[5], articolo[6], 
                                    articolo[7], articolo[8], articolo[9], 
                                    articolo[10], articolo[11])
            code = str(articolo[0])
            # descrizione codice, tariffa
            if code.startswith("$$$"):
                tariffa = str(articolo[1])
                descr_codice = articolo[2]
            else:
                n+=1
                tariffa = "n.{0}\n{1}".format(n, articolo[0])
                descr_codice = articolo[1]
            # verifica se stampare la descrizione breve
            descr_breve = str()
            if stampa_descr_breve and not code.startswith("$$$"): 
                descr_breve = articolo[2]
            # verifica se stampare la descrizione estesa
            descr_estesa = str()
            if stampa_descr_estesa and not code.startswith("$$$"): 
                descr_estesa = articolo[3]
            # verifica se stampare le note
            note = str()
            if stampa_note and not code.startswith("$$$"):
                note = articolo[12]
            # crea la tupla della colonna Descrizione
            Descrizione = (descr_codice, descr_estesa, descr_breve, note)
            # crea la tupla con le colonne del computo
            tupla = (code, tariffa, Descrizione, um, quantita, pr_un, pr_tot, costo, ric, tempo, sic)
            lista_articoli_stampare.append(tupla)
            # verifica se stampare gli elementi di listino (per l'analisi prezzi)
            if print_analisi:
                lista_articoli_stampare = self.__print_analisi (articolo[0], tupla, lista_articoli_stampare)
        return lista_articoli_stampare

    def print_epu(self, lista_supercapitoli=None, lista_capitoli=None, lista_subcapitoli=None, 
                       stampa_descr_breve = True, stampa_descr_estesa = True, stampa_note = True, 
                       print_only_if_not_0 = False, print_without_chapter = False, 
                       print_analisi= False):
        """Crea una lista di tuple in cui ciascun elemento rappresenta una colonna da stampare per 
        l'elenco prezzi: 'Tariffa - Descrizione - UM - Quantita - Pr.Unit. - Pr. Tot."""
        # inizializzo la lista degli articoli da stampare, idizionari ecc.
        lista_articoli_stampare = list()
        # se non ho una lista di capitoli da stampare usa tutti i capitoli
        lista_sup, lista_cap, lista_sub = self.capitoli_rows_list()
        if lista_supercapitoli is not None and len(lista_supercapitoli) > 0:
            lista_sup = lista_supercapitoli
        if lista_capitoli is not None and len(lista_capitoli) > 0:
            lista_cap = lista_capitoli
        if lista_subcapitoli is not None and len(lista_subcapitoli) > 0:
            lista_sub = lista_subcapitoli
        # ricava la lista delle tipologie contrattuali da stampare
        lista_type = [(None,)] 
        # ricava la lista degli elementi da stampare
        lista = self.__get_list_to_print(self.epu_rows_list, self.epu_count_rows,
                       None, lista_type,lista_sup, lista_cap, lista_sub,
                       print_only_if_not_0=print_only_if_not_0,
                       print_without_chapter=print_without_chapter, is_epu=True)
        # 0-Tariffa, 1-Supercapitolo, 2-Capitolo, 3-Subcapitolo, 4-Descrizione_codice, 5-Descrizione_voce, 
        # 6-Descrizione_estesa, 7- Unita_misura, 8-Ricarico, 9-Tempo_inst, 10-Costo_materiali, 11-Prezzo_unitario, 
        # 12-Sicurezza, 13-Cod_analisi, 14-Note, 15-CostoMat_1, 16-CostoMat_2, 17-CostoMat_3, 18-CostoMat_4
        n=0
        for articolo in lista:
            # se l'articolo è un titolo di capitolo non stampare quantità e prezzo tot
            # unità di misura, ricarichi tempi e altro ...
            um, ric, tempo, costo, pr_un, sic = (articolo[7], articolo[8], articolo[9], 
                                                 articolo[10], articolo[11], articolo[12])
            code = str(articolo[0])
            # descrizione codice, tariffa
            if code.startswith("$$$"):
                tariffa = str(articolo[1])
                quantita, pr_tot = ("", "")
                descr_codice = articolo[2]
            else:
                n+=1
                tariffa = "n.{0}\n{1}".format(n, articolo[0])
                quantita, pr_tot = (1, pr_un)
                descr_codice = articolo[4]
            # verifica se stampare la descrizione breve
            descr_breve = str()
            if stampa_descr_breve and not code.startswith("$$$"): 
                descr_breve = articolo[5]
            # verifica se stampare la descrizione estesa
            descr_estesa = str()
            if stampa_descr_estesa and not code.startswith("$$$"):
                descr_estesa = articolo[6]
            # verifica se stampare le note
            note = str()
            if stampa_note and not code.startswith("$$$"): 
                note = articolo[14]
            # crea la tupla della colonna Descrizione
            Descrizione = (descr_codice, descr_estesa, descr_breve, note)
            # crea la tupla con le colonne dell'epu
            tupla = (code, tariffa, Descrizione, um, quantita, pr_un, pr_tot, costo, ric, tempo, sic)
            lista_articoli_stampare.append(tupla)
            # verifica se stampare gli elementi di listino (per l'analisi prezzi)
            if print_analisi:
                lista_articoli_stampare = self.__print_analisi (articolo[0], tupla, lista_articoli_stampare)
        return lista_articoli_stampare

    def __get_list_to_print (self, func_list_rows, func_count_rows, func_calc,
                             list_type, list_sup, list_cat, list_sub,
                             print_only_if_not_0 = False, 
                             print_without_chapter = False, is_epu=False):
        # tutte le righe di intestazione capitoli e totali devono avere lo stesso 
        # codice iniziale identificativo
        COD_CAP = '$$$*$$$' #codice per nome capitoli/categorie/paragrafi
        COD_TOT = '$$$+$$$' #codice per riga dei totali parziali o generali
        # inizializzo la lista degli articoli da stampare, idizionari ecc.
        lista = list()
        n=0
        # ricavo il valore del campo valuta (da mostrare nei totali)
        valuta = self.dati_generali_list()[3]
        # prepara la lista di oggetti da stampare
        if print_without_chapter:
            if is_epu: 
                lista.extend(func_list_rows(list_if_computo_not_0 = print_only_if_not_0))
            else: 
                lista.extend(func_list_rows())
        else:
            for tp in list_type:
                if func_count_rows(None, None, None, print_only_if_not_0, tp[0]) > 0:
                    if tp[0] is not None:
                        # modifica il nome di default per la tipologia con campo nullo
                        if tp[0] == 'NULL': n_tp = STR_TYP_NULL
                        else: n_tp = tp[0]
                        # aggiungi alla lista il titolo della tipologia di appalto
                        c_tp = ascii_lowercase[n:n+1]+'-'
                        titoli_tp = (COD_CAP, c_tp, n_tp, '', '', '', '', '', '', '', '', '', '')
                        lista.append(titoli_tp)
                        n+=1
                    else: c_tp = ''
                    for sup in list_sup:
                        if func_count_rows(sup[0], None, None, print_only_if_not_0, tp[0]) > 0:
                            for cat in list_cat:
                                if func_count_rows(sup[0], cat[0], None,
                                               print_only_if_not_0, tp[0]) > 0:
                                    for sub in list_sub:
                                        if func_count_rows(sup[0], cat[0], sub[0],
                                                   print_only_if_not_0, tp[0]) > 0:
                                           # aggiungi alla lista il titolo delle categorie
                                           code = "{0}{1}-{2}-{3}".format(c_tp, sup[0], cat[0], sub[0])
                                           text = "%s\n%s\n%s" % (sup[1], cat[1], sub[1])
                                           titoli_cat = (COD_CAP, code, text, '', '', '', '', '', '', '', '', '', '')
                                           # aggiungi alla lista il prezzo totale della categoria
                                           if func_calc is not None:
                                               totale_sub = func_calc(sup[0], cat[0], sub[0], tp[0])
                                               text1 = "(%s -%s - %s)" % (sup[1], cat[1], sub[1])
                                               prezzi_cat = (COD_TOT, code, 
                                                             '{0} {1} ({2}):'.format(STR_TOT_P, text1, valuta),
                                                             '', '', '', '', '', '', '', totale_sub, '', '')
                                           #per mantenere la stessa struttura del documento inserisce
                                           #una riga di 'termine paragrafo anche quando non è necessario
                                           #inserire nel documento una riga con i prezzi tot. (es. EPU)
                                           else:
                                               prezzi_cat = (COD_TOT,'','','', '', '', '', '', '', '', '', '', '')
                                           # aggiungi alla lista gli articoli di computo
                                           lista.append(titoli_cat)
                                           lista.extend(func_list_rows(sup[0], cat[0], 
                                                        sub[0], print_only_if_not_0, 
                                                        contract_type=tp[0]))
                                           lista.append(prezzi_cat)
                    # aggiungere totale parziale tipologia di appalto
                    if func_calc is not None:
                        if len(list_type) > 1:
                            tot_str = STR_TOT_P
                        else: 
                            tot_str = STR_TOT_G
                            n_tp = ''
                        totale_tp = func_calc(None, None, None, tp[0])
                        prezzi_tp = (COD_TOT, '', "{0} {1} ({2}):".format(tot_str, n_tp, valuta),'', '', '', '', '', '', '', totale_tp, '', '')
                        lista.append(prezzi_tp)
        # restituisce la lista degli articoli
        return lista

    def print_report_computo (self, lista_supercategorie=None, lista_categorie=None, lista_subcategorie=None, 
                              print_only_if_not_0 = False, print_only_grand_total = False):
        """"""
        # se non ho una lista di categorie da stampare usa tutte le categorie
        lista_sup, lista_cat, lista_sub = self.categorie_rows_list()
        if lista_supercategorie is not None and len(lista_supercategorie) > 0:
            lista_sup = lista_supercategorie
        if lista_categorie is not None and len(lista_categorie) > 0:
            lista_cat = lista_categorie
        if lista_subcategorie is not None and len(lista_subcategorie) > 0:
            lista_sub = lista_subcategorie
        # ricava la lista delle tipologie contrattuali da stampare
        lista_type = self.list_contract_type()
        if len(lista_type) == 0: lista_type = [(None,)]
        return self.__report_st (self.computo_count_rows, self.calcolo_prezzi,
                                 lista_sup, lista_cat, lista_sub, lista_type,
                                 print_only_if_not_0=print_only_if_not_0,
                                 print_only_grand_total=print_only_grand_total)

    def print_report_lavorazioni (self, lista_supercapitoli=None, lista_capitoli=None, lista_subcapitoli=None, 
                              print_only_if_not_0 = False, print_only_grand_total = False):
        """"""
        # se non ho una lista di categorie da stampare usa tutte le categorie
        lista_sup, lista_cat, lista_sub = self.capitoli_rows_list()
        if lista_supercapitoli is not None and len(lista_supercapitoli) > 0:
            lista_sup = lista_supercapitoli
        if lista_capitoli is not None and len(lista_capitoli) > 0:
            lista_cat = lista_capitoli
        if lista_subcapitoli is not None and len(lista_subcapitoli) > 0:
            lista_sub = lista_capitoli
        # ricava la lista delle tipologie contrattuali da stampare
        lista_type = [(None,)]
        return self.__report_st (self.epu_count_rows, self.calcolo_prezzi_lavorazioni,
                                 lista_sup, lista_cat, lista_sub, lista_type,
                                 print_only_if_not_0=print_only_if_not_0,
                                 print_only_grand_total=print_only_grand_total)

    def __report_st (self, func_count_rows, func_calc,
                     list_sup, list_cat, list_sub, list_type,
                     print_only_if_not_0 = False,
                     print_only_grand_total = False):
        # inizializzo la lista degli articoli da stampare
        lista = list()
        n=0
        # ricavo il valore del campo valuta (da mostrare nei totali)
        valuta = self.dati_generali_list()[3]
        # prepara la lista di oggetti da stampare
        totale_globale = func_calc (convert_to_string=False)
        if not print_only_grand_total:
            for tp in list_type:
                if func_count_rows(None, None, None, print_only_if_not_0, tp[0]) > 0:
                    if tp[0] is not None:
                        # modifica il nome di default per la tipologia con campo nullo
                        if tp[0] == 'NULL': n_tp = STR_TYP_NULL
                        else: n_tp = tp[0]
                        # aggiungi alla lista il titolo della tipologia di appalto
                        c_tp = ascii_lowercase[n:n+1] + '-'
                        titoli_tp = ('TYP', c_tp, n_tp, '', '')
                        lista.append(titoli_tp)
                        n+=1
                    else: c_tp = ''
                    for sup in list_sup:
                        if func_count_rows(sup[0], None, None, print_only_if_not_0, tp[0]) > 0:
                            totale = func_calc(sup[0], None, None, tp[0], convert_to_string=False)
                            incid_perc =  self.__calcola_incidenza_percentuale (totale, totale_globale)
                            cod = "{0}{1}".format(c_tp, sup[0])
                            titoli = ('SUP', cod, "%s\n" % sup[1], 
                                      self.__convert_number_to_string(totale), incid_perc)
                            lista.append(titoli)
                            for cat in list_cat:
                                if func_count_rows(sup[0], cat[0], None, print_only_if_not_0, tp[0]) > 0:
                                    totale = func_calc(sup[0], cat[0], None, tp[0], convert_to_string=False)
                                    incid_perc =  self.__calcola_incidenza_percentuale (totale, totale_globale)
                                    cod = "{0}{1}-{2}".format(c_tp, sup[0], cat[0])
                                    titoli = ('CAT', cod, "%s\n" % cat[1],
                                              self.__convert_number_to_string(totale), incid_perc)
                                    lista.append(titoli)
                                    for sub in list_sub:
                                        if func_count_rows(sup[0], cat[0], sub[0], print_only_if_not_0, tp[0]) > 0:
                                            totale = func_calc(sup[0], cat[0], sub[0], tp[0], convert_to_string=False)
                                            incid_perc =  self.__calcola_incidenza_percentuale (totale, totale_globale)
                                            cod = "{0}{1}-{2}-{3}".format(c_tp, sup[0], cat[0], sub[0])
                                            titoli = ('SUB', cod, "%s\n" % sub[1], 
                                                      self.__convert_number_to_string(totale), incid_perc)
                                            lista.append(titoli)
        # aggiungi il totale complessivo
        incid_perc =  self.__calcola_incidenza_percentuale (totale_globale, totale_globale)
        titoli = ('TOT', "", "{0} ({1}):".format(STR_TOT_G, valuta), 
                  self.__convert_number_to_string(totale_globale), incid_perc)
        lista.append(titoli)
        return lista

    def print_report_categorie (self, lista_supercategorie=None, 
                                lista_categorie=None, 
                                lista_subcategorie=None, 
                                print_only_if_not_0 = False):
        """"""
        # se non ho una lista di categorie da stampare usa tutte le categorie
        lista_sup, lista_cat, lista_sub = self.categorie_rows_list()
        if lista_supercategorie is not None and len(lista_supercategorie) > 0:
            lista_sup = lista_supercategorie
        if lista_categorie is not None and len(lista_categorie) > 0:
            lista_cat = lista_categorie
        if lista_subcategorie is not None and len(lista_subcategorie) > 0:
            lista_sub = lista_subcategorie
        # ricava la lista delle tipologie contrattuali da stampare
        lista_type = self.list_contract_type()
        if len(lista_type) == 0: lista_type = [(None,)]
        return self.__report_cp (self.computo_count_rows, self.calcolo_prezzi,
                                 list_sup=lista_sup, 
                                 list_cat=lista_cat, 
                                 list_sub=lista_sub, 
                                 list_type=lista_type,
                                 print_only_if_not_0 = print_only_if_not_0)

    def print_report_capitoli (self, lista_supercapitoli=None, 
                               lista_capitoli=None, 
                               lista_subcapitoli=None, 
                               print_only_if_not_0 = False):
        """"""
        # se non ho una lista di categorie da stampare usa tutte le categorie
        lista_sup, lista_cat, lista_sub = self.capitoli_rows_list()
        if lista_supercapitoli is not None and len(lista_supercapitoli) > 0:
            lista_sup = lista_supercapitoli
        if lista_capitoli is not None and len(lista_capitoli) > 0:
            lista_cat = lista_capitoli
        if lista_subcapitoli is not None and len(lista_subcapitoli) > 0:
            lista_sub = lista_capitoli
        # ricava la lista delle tipologie contrattuali da stampare
        lista_type = [(None,)]
        return self.__report_cp (self.epu_count_rows, self.calcolo_prezzi_lavorazioni,
                              list_sup=lista_sup, 
                              list_cat=lista_cat, 
                              list_sub=lista_sub,
                              list_type=lista_type,
                              print_only_if_not_0 = print_only_if_not_0)

    def __report_cp (self, func_count_rows, func_calc,
                     list_sup, list_cat, list_sub, list_type,
                     print_only_if_not_0 = False):
        # inizializzo la lista degli articoli da stampare
        lista = list()
        n=0
        # ricavo il valore del campo valuta (da mostrare nei totali)
        valuta = self.dati_generali_list()[3]
        # prepara la lista di oggetti da stampare
        totale_globale = func_calc (convert_to_string=False)
        incid_perc =  self.__calcola_incidenza_percentuale (totale_globale, totale_globale)
        totale_gen = ('TOT', "", "{0} ({1}):".format(STR_TOT_G, valuta), 
                    self.__convert_number_to_string(totale_globale), incid_perc)
        # pagina per le tipologie in appalto (da aggiungere solo se ne esiste più di una)
        if len(list_type) > 1:
            for tp in list_type:
                if func_count_rows(None, None, None, print_only_if_not_0, tp[0]) > 0:
                    # aggiungi alla lista il titolo della tipologia di appalto
                    totale = func_calc(None, None, None, tp[0], convert_to_string=False)
                    incid_perc =  self.__calcola_incidenza_percentuale (totale, totale_globale)
                    if tp[0] is not None:
                        # modifica il nome di default per la tipologia con campo nullo
                        if tp[0] == 'NULL': n_tp = STR_TYP_NULL
                        else: n_tp = tp[0]
                        # aggiungi alla lista il titolo della tipologia di appalto
                        c_tp = ascii_lowercase[n:n+1] + '-'
                        titoli = ('SUP', c_tp, "%s\n" % n_tp, self.__convert_number_to_string(totale), incid_perc)
                        lista.append(titoli)
                        n+=1
            # aggiungi il totale complessivo
            lista.append(totale_gen)
        # pagina per le supercategorie
        for sup in list_sup:
            if func_count_rows(sup[0], None, None, print_only_if_not_0) > 0:
                totale = func_calc(sup[0], None, None, convert_to_string=False)
                incid_perc =  self.__calcola_incidenza_percentuale (totale, totale_globale)
                titoli = ('SUP', "%s" % sup[0], "%s\n" % sup[1], 
                          self.__convert_number_to_string(totale), incid_perc)
                lista.append(titoli)
        # aggiungi il totale complessivo
        lista.append(totale_gen)
        # pagina per le categorie
        for cat in list_cat:
            if func_count_rows(None, cat[0], None, print_only_if_not_0) > 0:
                totale = func_calc(None, cat[0], None, convert_to_string=False)
                incid_perc =  self.__calcola_incidenza_percentuale (totale, totale_globale)
                titoli = ('SUP', "%s" % cat[0], "%s\n" % cat[1], 
                          self.__convert_number_to_string(totale), incid_perc)
                lista.append(titoli)
        # aggiungi il totale complessivo
        lista.append(totale_gen)
        # pagina per le subcategorie
        for sub in list_sub:
            if func_count_rows(None, None, sub[0], print_only_if_not_0) > 0:
                totale = func_calc(None, None, sub[0], convert_to_string=False)
                incid_perc =  self.__calcola_incidenza_percentuale (totale, totale_globale)
                titoli = ('SUP', "%s" % sub[0], "%s\n" % sub[1], 
                          self.__convert_number_to_string(totale), incid_perc)
                lista.append(titoli)
        # aggiungi il totale complessivo
        lista.append(totale_gen)
        return lista

    def print_report_inc_manodopera (self, lista_supercategorie=None, 
                                    lista_categorie=None, 
                                    lista_subcategorie=None, 
                                    print_only_if_not_0 = False,
                                    min_to_hour=True):
        """crea un report con le incidenze perc suddivise per categorie di cm"""
        # se non ho una lista di categorie da stampare usa tutte le categorie
        lista_sup, lista_cat, lista_sub = self.categorie_rows_list()
        if lista_supercategorie is not None and len(lista_supercategorie) > 0:
            lista_sup = lista_supercategorie
        if lista_categorie is not None and len(lista_categorie) > 0:
            lista_cat = lista_categorie
        if lista_subcategorie is not None and len(lista_subcategorie) > 0:
            lista_sub = lista_subcategorie
        return self.__report_inc (self.computo_count_rows, 
                                  self.calcolo_incidenza_manodopera,
                                  list_sup=lista_sup, 
                                  list_cat=lista_cat, 
                                  list_sub=lista_sub, 
                                  print_only_if_not_0 = print_only_if_not_0,
                                  title = STR_MO)
            
    def print_report_inc_mo_lavorazioni (self, lista_supercapitoli=None, 
                                    lista_capitoli=None, 
                                    lista_subcapitoli=None, 
                                    print_only_if_not_0 = False):
        """crea un report con le incidenze perc suddivise per capitoli di epu"""
        # se non ho una lista di capitoli da stampare usa tutti i capitoli
        lista_sup, lista_cat, lista_sub = self.capitoli_rows_list()
        if lista_supercapitoli is not None and len(lista_supercapitoli) > 0:
            lista_sup = lista_supercapitoli
        if lista_capitoli is not None and len(lista_capitoli) > 0:
            lista_cat = lista_capitoli
        if lista_subcapitoli is not None and len(lista_subcapitoli) > 0:
            lista_sub = lista_subcapitoli
        return self.__report_inc (self.epu_count_rows, 
                                  self.calcolo_inc_manodopera_lavorazioni,
                                  list_sup=lista_sup, 
                                  list_cat=lista_cat, 
                                  list_sub=lista_sub, 
                                  print_only_if_not_0 = print_only_if_not_0,
                                  title = STR_MO)

    def print_report_inc_sicurezza (self, lista_supercategorie=None, 
                                    lista_categorie=None, 
                                    lista_subcategorie=None, 
                                    print_only_if_not_0 = False):
        """crea un report con le incidenze perc della sicurezza suddivise per categorie di computo"""
        # se non ho una lista di categorie da stampare usa tutte le categorie
        lista_sup, lista_cat, lista_sub = self.categorie_rows_list()
        if lista_supercategorie is not None and len(lista_supercategorie) > 0:
            lista_sup = lista_supercategorie
        if lista_categorie is not None and len(lista_categorie) > 0:
            lista_cat = lista_categorie
        if lista_subcategorie is not None and len(lista_subcategorie) > 0:
            lista_sub = lista_subcategorie
        return self.__report_inc (self.computo_count_rows, 
                                  self.calcolo_incidenza_sicurezza,
                                  list_sup=lista_sup, 
                                  list_cat=lista_cat, 
                                  list_sub=lista_sub, 
                                  print_only_if_not_0 = print_only_if_not_0,
                                  title = STR_SIC)

    def print_report_inc_sic_lavorazioni (self, lista_supercapitoli=None, 
                                          lista_capitoli=None, 
                                          lista_subcapitoli=None, 
                                          print_only_if_not_0 = False):
        """crea un report con le incidenze perc della sicurezza suddivise per capitoli di epu"""
        # se non ho una lista di capitoli da stampare usa tutti i capitoli
        lista_sup, lista_cat, lista_sub = self.capitoli_rows_list()
        if lista_supercapitoli is not None and len(lista_supercapitoli) > 0:
            lista_sup = lista_supercapitoli
        if lista_capitoli is not None and len(lista_capitoli) > 0:
            lista_cat = lista_capitoli
        if lista_subcapitoli is not None and len(lista_subcapitoli) > 0:
            lista_sub = lista_subcapitoli
        return self.__report_inc (self.lavorazioni_count_rows, 
                                  self.calcolo_incidenza_sic_lavorazioni,
                                  list_sup=lista_sup, 
                                  list_cat=lista_cat, 
                                  list_sub=lista_sub, 
                                  print_only_if_not_0 = print_only_if_not_0,
                                  title=STR_SIC)

    def __report_inc (self, func_count_rows, func_calc,
                      list_sup=None, 
                      list_cat=None, 
                      list_sub=None, 
                      print_only_if_not_0 = False,
                      title=None):
        if title is None: title = ""
        # inizializzo la lista degli articoli da stampare
        lista = list()
        # ricavo il valore del campo valuta (da mostrare nei totali)
        valuta = self.dati_generali_list()[3]
        # prepara la lista di oggetti da stampare
        tmp_tot, totale_globale, incid_perc = func_calc()
        totale_gen = ('TOT', "", _("Totale Complessivo {0} ({1}):").format(title, valuta), 
                      tmp_tot, totale_globale, incid_perc)
        for sup in list_sup:
            if func_count_rows(sup[0], None, None, print_only_if_not_0) > 0:
                tmp_tot, totale, incid_perc = func_calc (sup[0], None, None)
                titoli = ('SUP', "%s" % sup[0], "%s\n" % sup[1], tmp_tot, totale, incid_perc)
                lista.append(titoli)
        # aggiungi il totale complessivo
        lista.append(totale_gen)
        for cat in list_cat:
            if func_count_rows(None, cat[0], None, print_only_if_not_0) > 0:
                tmp_tot, totale, incid_perc = func_calc (None, cat[0], None)
                titoli = ('SUP', "%s" % cat[0], "%s\n" % cat[1], tmp_tot, totale, incid_perc)
                lista.append(titoli)
        # aggiungi il totale complessivo
        lista.append(totale_gen)
        for sub in list_sub:
            if func_count_rows(None, None, sub[0], print_only_if_not_0) > 0:
                tmp_tot, totale, incid_perc = func_calc (None, None, sub[0])
                titoli = ('SUP', "%s" % sub[0], "%s\n" % sub[1], 
                          tmp_tot, totale, incid_perc)
                lista.append(titoli)
        # aggiungi il totale complessivo
        lista.append(totale_gen)
        return lista

    def __print_analisi (self, tariffa, tupla_articolo, lista_articoli_stampare):
        """la funzione aggiunge in coda ad ogni articolo i suoi elementi di analisi prezzi"""
        # tutte le righe di listino devono avere lo stesso codice iniziale che le identifichi
        COD_LIS = '$$$LIS$$$' #codice per singolo elemento di listino
        COD_LIS_T = '$$$LIT$$$' #codice per intestazione articolo, totali parziali, ecc.
        COD_A = "A" #codice per elementi per materiali
        COD_B = "B" #codice per elementi di manodopera
        COD_C = "C" #codice per elementi per la sicurezza
        COD_D = "D" #codice per arrotondamento
        i=0
        # ricavo i campi testo dell'articolo da analizzare
        (code, tar_art, Descrizione, um, quantita,  
         pr_un_art, pr_tot_art, costo, ric, tempo, sic) = tupla_articolo
        # se le righe sono speciali (es. per capitoli subtotali, ecc. saltale)
        if code.startswith('$$$'):
            return lista_articoli_stampare
        # ricava la lista degli elementi di listino dell'articolo
        lista_art_list = self.listino_rows_list (tariffa)
        # stampa una riga di introduzione dell'analisi prezzi
        tar = "{0}".format(tariffa)
        Descrizione = ("{0} {1}:".format(_("Elementi di Analisi art."), tariffa), '', '', '')
        lista_articoli_stampare.append((COD_LIS_T, tar, Descrizione,'', '', '', 
                                                            '', '', '', '', ''))
        # A) se non ci sono elementi di listino stampa una riga generica
        if len(lista_art_list) == 0:
            tar = "{0} ({1}.{2})".format(tariffa, COD_A, 1)
            pr_un = costo
            pr_tot = costo*quantita
            Descrizione = ("{0}.{1} - {2} {3}".format(COD_A, 1, 
                           _("Costo Materiali Art."), tariffa), '', '', '')
            lista_articoli_stampare.append((COD_LIS, tar, Descrizione, "corpo", 
                                    quantita, pr_un, pr_tot, '', '', '', ''))
        else:
            # A) inizializza l'indice per i codici di listino
            for art_list in lista_art_list:
                #id INTEGER PRIMARY KEY, Tariffa TEXT, Codice TEXT, Descrizione_codice TEXT, 
                #Unita_misura TEXT, Quantita REAL, Prezzo_unitario REAL, Sconto REAL, Accessori INTEGER, Prezzo_totale REAL, Note TEXT
                i+=1
                descr_codice = art_list[3]
                note = art_list[10]
                tar = "{0} ({1}.{2})".format(tariffa, COD_A, i)
                Descrizione = ("{0}.{1} - {2}".format(COD_A,i, descr_codice), '', art_list[2], note)
                um = art_list[4]
                quant = art_list[5]*quantita
                sconto = 1-(art_list[7]/100)
                pr_un = art_list[6] * (sconto) # il prezzo totale viene modificato con lo sconto
                accessori = art_list[8]
                pr_tot =  pr_un*quant
                tupla_list = (COD_LIS, tar, Descrizione, um, quant, pr_un, pr_tot, art_list[6], sconto, '', '')
                lista_articoli_stampare.append(tupla_list)
                # stampa la riga degli accessori
                if accessori > 0:
                    i+=1
                    tar = "{0} ({1}.{2})".format(tariffa, COD_A, i)
                    Descrizione = ("{0}.{1} - {2} {3}".format(COD_A, i, _("Accessori per art."), descr_codice), '', '', '')
                    lista_articoli_stampare.append((COD_LIS, tar, Descrizione, "corpo", 
                                  1, accessori, accessori, accessori, '', '', ''))
        # A) stampa la riga di riepilogo dei costi
        tar = "{0} ({1}.{2})".format(tariffa, COD_A, i+1)
        ricarico = str((ric-1)*100)
        pr_un = costo*quantita
        pr_tot_ric = pr_un*ric
        Descrizione = ("{0}.{1} - {2} ({3}%)".format(COD_A,i+1, _('Parziale Costo materiali con ricarico'), ricarico), '', '', '')
        lista_articoli_stampare.append((COD_LIS_T, tar, Descrizione, "-", ric, 
                                        pr_un, pr_tot_ric, '', ric, '', ''))
        # B) stampa la riga della manodopera
        if tempo > 0:
            # inizializza l'indice per i codici di manodopera
            i=0
            tmp = float(tempo)/60
            lista_mo = self.get_table_manodopera()
            # B) iniz. somma dei totali della manodopera per calcolo finale
            sum_mo = 0 
            if len(lista_mo) != 0:
                perc_tot = self.__get_perc_tot()
                tmp_a = tmp/perc_tot
                for key, nome, costo, perc, note in lista_mo:
                    i+=1
                    tar = "{0} ({1}.{2})".format(tariffa, COD_B, i)
                    tmp_mo = tmp_a*perc*quantita
                    tempo_mo = tmp_mo*60
                    pr_un = costo/ric
                    pr_tot_mo = pr_un*tmp_mo
                    Descrizione = ("{0}.{1} - {2}".format(COD_B, i, nome), '', '', note)
                    lista_articoli_stampare.append((COD_LIS, tar, Descrizione, "h", 
                                tmp_mo, pr_un, pr_tot_mo, '', '', tempo_mo, ''))
                    sum_mo += pr_tot_mo
            else:
                # B) inserisci il totale della manodopera
                i+=1
                tar = "{0} ({1}.{2})".format(tariffa, COD_B, i)
                tmp_mo = tmp*quantita
                pr_un = self.dati_generali_list()[1]/ric
                pr_tot_mo = pr_un*tmp_mo
                Descrizione = ("{0}.{1} - {2}".format(COD_B, i,_("Totale manodopera")), '', '', '')
                lista_articoli_stampare.append((COD_LIS_T, tar, Descrizione, "h", 
                                      tmp_mo, pr_un, pr_tot_mo, '', '', tempo, ''))
                sum_mo = pr_tot_mo
            # B) inserisci la riga finale della manodopera
            pr_un = sum_mo
            pr_tot_mo = pr_un*ric
            Descrizione = ("{0}.{1} - {2} ({3}%)".format(COD_B, i+1, _('Parziale Manodopera con ricarico'), ricarico), '', '', '')
            lista_articoli_stampare.append((COD_LIS_T, tar, Descrizione, "-", ric, 
                                             pr_un, pr_tot_mo, '', ric, '', ''))
        else: pr_tot_mo = 0
        # C) stampa la riga per la sicurezza
        if sic > 0:
            tar = "{0} ({1}.{2})".format(tariffa, COD_C, 1)
            pr_un_sic = pr_tot_ric + pr_tot_mo
            pr_tot_sic = pr_un_sic*sic/100
            Descrizione = ("{0}.{1} - {2}".format(COD_C, 1, _("Oneri di Sicurezza")), '', '', '')
            lista_articoli_stampare.append((COD_LIS_T, tar, Descrizione, "%", 
                                   sic, pr_un_sic, pr_tot_sic, '', '', '', sic))
        else: pr_tot_sic = 0
        # D) stampa l'arrotondamento per l'articolo di listino
        arrotondamento = round((pr_un_art*quantita),2)-round(pr_tot_mo,2)-round(pr_tot_ric,2)-round(pr_tot_sic,2)
        if abs(round(arrotondamento, 2)) > 0.00:
            tar = "{0} ({1}.{2})".format(tariffa, COD_D, 1)
            Descrizione = ("{0}.{1} - {2}".format(COD_D,1,_("Arrotondamento")), '', '', '')
            lista_articoli_stampare.append((COD_LIS_T, tar, Descrizione, "-", 
                           '-', arrotondamento, arrotondamento, '', '', '', ''))
        return lista_articoli_stampare

    def __calcola_incidenza_percentuale (self, parziale, totale):
        if totale == 0:
            return 0.0
        else:
            return (parziale/totale)*100
################################################################################
############ COMPUTO - ESPORTAZIONI ############################################
################################################################################

    def export_to_dumpfile (self, filename):
        """esporta il database con le istruzioni SQL in un file di testo"""
        con = self.db
        full_dump = os.linesep.join([line for line in con.iterdump()])
        f = open(filename, 'w')
        f.writelines(full_dump)
        f.close()
        return True

    def export_epu_to_csv (self, filename, stampa_descr_estesa=False, 
                           stampa_descr_breve = False, stampa_note=False,
                           without_chapter=False, print_analisi= True):
        """export epu to csv file"""
        epu = self.print_epu (lista_supercapitoli=None, 
                                  lista_capitoli=None, lista_subcapitoli=None, 
                                      stampa_descr_breve = stampa_descr_breve, 
                                      stampa_descr_estesa = stampa_descr_estesa, 
                                      stampa_note = stampa_note,
                                      print_only_if_not_0 = False, 
                                      print_without_chapter = without_chapter, 
                                      print_analisi= print_analisi)
        # processa e modifica la lista computo per adattare la descrizione
        lista_epu = list()
        for line in epu:
            #tupla = (code, tariffa, Descrizione, um, quantita, pr_un, pr_tot, costo, ric, tempo, sic)
            #Descrizione = (descr_codice, descr_estesa, descr_breve, note)
            descr_codice, descr_estesa, descr_breve, note = line[2]
            lista_epu.append([line[1], "{0}\n{1}\n{2}\n{3}".format(descr_codice, 
                              descr_estesa, descr_breve, note), line[3], 
                              line[4], line[5], line[6]])
        # are il file csv in scrittura
        computo_csv = csv.writer(open(filename, 'wb'), dialect=csv.excel, 
                                 quoting=csv.QUOTE_NONNUMERIC)
        # scrive sul file i dati estratti e processati
        computo_csv.writerow([_("No.\nTariffa"), _("Descrizione"), _("UM"), 
                              _("Quantita"), _("P.Unitario"), _("P.Totale")])
        computo_csv.writerows(lista_epu)
        return True

    def export_computo_to_csv (self, filename, stampa_descr_estesa=False, 
                               stampa_descr_breve = False, stampa_note=False, 
                               export_report=True,
                               without_chapter=False, print_analisi= False):
        """export computo to csv file"""
        computo = self.print_computo (lista_supercategorie=None, 
                                      lista_categorie=None, 
                                      lista_subcategorie=None, 
                                      stampa_descr_breve = stampa_descr_breve, 
                                      stampa_descr_estesa = stampa_descr_estesa, 
                                      stampa_note = stampa_note,
                                      print_only_if_not_0 = False, 
                                      print_without_chapter = without_chapter, 
                                      print_analisi= print_analisi)
        # processa e modifica la lista computo per adattare la descrizione
        lista_computo = list()
        for line in computo:
            #tupla = (code, tariffa, Descrizione, um, quantita, pr_un, pr_tot, costo, ric, tempo, sic)
            #Descrizione = (descr_codice, descr_estesa, descr_breve, note)
            descr_codice, descr_estesa, descr_breve, note = line[2]
            lista_computo.append([line[1],"{0}\n{1}\n{2}\n{3}".format(descr_codice, 
                                 descr_estesa, descr_breve, note), line[3], 
                                 line[4], line[5], line[6]])
        # estrae il report finale se richiesto
        lista_report = list()
        if export_report:
            report = self.print_report_computo (lista_supercategorie=None, 
                                   lista_categorie=None, lista_subcategorie=None, 
                                   print_only_if_not_0 = False, 
                                   print_only_grand_total = False)
            for line in report:
                #('SUP', "%s" % cat[0], "%s\n" % cat[1], totale, incid_perc)
                if line[0] == 'CAT': indent = " "*2
                elif line[0] == 'SUB': indent = " "*4
                else: indent = str()
                lista_report.append([line[1], indent + line[2], "", line[3], line[4]])
        # are il file csv in scrittura
        computo_csv = csv.writer(open(filename, 'wb'), dialect=csv.excel, 
                                 quoting=csv.QUOTE_NONNUMERIC)
        # scrive sul file i dati estratti e processati
        computo_csv.writerow([_("No.\nTariffa"), _("Descrizione"), _("UM"), 
                              _("Quantita"), _("P.Unitario"), _("P.Totale")])
        computo_csv.writerows(lista_computo)
        if export_report:
            computo_csv.writerow([_("Cat."), _("Categoria\n"), "", _("Totale"),
                                  _("Incidenza"), ""])
            computo_csv.writerows(lista_report)
        return True

    def export_lavorazioni_to_csv (self, filename, stampa_descr_estesa=False, 
                               stampa_descr_breve = False, stampa_note=False, 
                               export_report=True,
                               without_chapter=False, print_analisi = False):
        """export lista lavorazioni to csv file"""
        computo = self.print_lavorazioni(lista_supercapitoli=None, 
                                      lista_capitoli=None, 
                                      lista_subcapitoli=None, 
                                      stampa_descr_breve = stampa_descr_breve, 
                                      stampa_descr_estesa = stampa_descr_estesa, 
                                      stampa_note = stampa_note,
                                      print_only_if_not_0 = False, 
                                      print_without_chapter = without_chapter,
                                      print_analisi= print_analisi)
        # processa e modifica la lista computo per adattare la descrizione
        lista_computo = list()
        for line in computo:
            #tupla = (code, tariffa, Descrizione, um, quantita, pr_un, pr_tot, costo, ric, tempo, sic)
            #Descrizione = (descr_codice, descr_estesa, descr_breve, note)
            descr_codice, descr_estesa, descr_breve, note = line[2]
            lista_computo.append([line[1],"{0}\n{1}\n{2}\n{3}".format(descr_codice, 
                                 descr_estesa, descr_breve, note), line[3], 
                                 line[4], line[5], line[6]])
        # estrae il report finale se richiesto
        lista_report = list()
        if export_report:
            report = self.print_report_lavorazioni (lista_supercapitoli=None, 
                                   lista_capitoli=None, lista_subcapitoli=None, 
                                   print_only_if_not_0 = False, 
                                   print_only_grand_total = False)
            for line in report:
                #('SUP', "%s" % cat[0], "%s\n" % cat[1], totale, incid_perc)
                if line[0] == 'CAT': indent = " "*2
                elif line[0] == 'SUB': indent = " "*4
                else: indent = str()
                lista_report.append([line[1], indent + line[2], "", line[3], line[4]])
        # are il file csv in scrittura
        computo_csv = csv.writer(open(filename, 'wb'), dialect=csv.excel, 
                                 quoting=csv.QUOTE_NONNUMERIC)
        # scrive sul file i dati estratti e processati
        computo_csv.writerow([_("No.\nTariffa"), _("Descrizione"), _("UM"), 
                              _("Quantita"), _("P.Unitario"), _("P.Totale")])
        computo_csv.writerows(lista_computo)
        if export_report:
            computo_csv.writerow([_("Cat."), _("Categoria\n"), "", _("Totale"),
                                  _("Incidenza"), ""])
            computo_csv.writerows(lista_report)
        return True

    def export_to_primus_pwe (self, filename, esporta_elementi_listino=False):
        """esportazione del preventivo nel file di interscambio di Primus PWE"""
        # costanti e inizializzazione variabili
        cr = "\r\n" # carriage return da inserire alla fine di ogni rigo
        diz_tariffe = dict() #dizionario tariffe-codici epu
        # ricava una lista di oggetti 'articolo' ricavata dalla tabella epu
        lista_articoli_epu = list()
        self.c.execute("""SELECT Tariffa FROM Epu ORDER BY Tariffa;""") 
        lista = self.c.fetchall()
        for tariffa in lista:
            articolo = self.copia_articolo_epu (tariffa[0], copia_nomi_capitoli= True)
            lista_articoli_epu.append(articolo)
        # ricava una lista di oggetti 'articolo' ricavata dalla tabella computo
        lista_articoli_cmp = list()
        self.c.execute("""SELECT id FROM Computo ORDER BY id;""")
        lista = self.c.fetchall()
        for prymary_key in lista:
            articolo = self.copia_articolo_computo (prymary_key[0], copia_nomi_categorie= True)
            lista_articoli_cmp.append(articolo)
        # ricava una lista di oggetti 'listino' ricavata dalla tabella listino
        lista_listino = list()
        self.c.execute("""SELECT id FROM Analisi ORDER BY id;""")
        lista = self.c.fetchall()
        for prymary_key in lista:
            articolo = self.copia_articolo_listino (prymary_key[0])
            lista_listino.append(articolo)
        # crea le seguenti liste: capitoli e categorie
        sup_cat, cat, sub_cat = self.categorie_rows_list()
        sup_cap, cap, sub_cap = self.capitoli_rows_list()
        # crea una lista di dati per la sez 'dati generali' del file PWE
        lista_dg = self.dati_generali_list ()
        lista_dati_generali_PWE = list()
        (ricarico, manodopera, sicurezza, valuta, nome, indirizzo, comune, 
            provincia, cliente, redattore)= lista_dg
        # creazione dei campi capitoli e categorie
        Supercategoria = self.__somma_nomi(sup_cat)
        Categoria = self.__somma_nomi(cat)
        Supercapitolo = self.__somma_nomi(sup_cap)
        Subcategoria = self.__somma_nomi(sub_cat)
        Capitolo = self.__somma_nomi(cap)
        Subcapitolo = self.__somma_nomi(sub_cap)
        # verificare se mantenere formula per utile d'impresa
        spese_generali = (ricarico-1)*100/2
        utile_impresa = (ricarico-1)*100/2
        sicurezza = sicurezza*100
        # cifre decimali
        dec = settings["decimals"]
        dec_b = 3
        # lista conversioni
        lista_dati_generali_PWE.append("@a%s%s" % (comune, cr))
        lista_dati_generali_PWE.append("@b%s%s" % (provincia, cr))
        lista_dati_generali_PWE.append("@c%s%s" % (cliente, cr))
        lista_dati_generali_PWE.append("@d%s%s" % (indirizzo, cr))
        lista_dati_generali_PWE.append("@e%s%s" % (nome, cr))
        lista_dati_generali_PWE.append("@f%s%s" % (redattore, cr))
        lista_dati_generali_PWE.append("@g%s%s" % (Supercategoria, cr)) #Supercategoria
        lista_dati_generali_PWE.append("@h%s%s" % (Categoria, cr)) #Categoria
        lista_dati_generali_PWE.append("@i%s%s" % (Supercapitolo, cr)) #Supercapitolo
        lista_dati_generali_PWE.append("@j%s%s" % (Subcategoria, cr)) #Subcategoria
        lista_dati_generali_PWE.append("@k%s%s" % (Capitolo, cr)) #Capitolo
        lista_dati_generali_PWE.append("@l%s%s" % (Subcapitolo, cr)) #Subcapitolo
        campo_m = "%f|%f|%f" % (spese_generali, utile_impresa, sicurezza)
        lista_dati_generali_PWE.append("@m{0}|{1}|{2}{3}".format(spese_generali, utile_impresa, sicurezza, cr))
        lista_dati_generali_PWE.append("@n%s%s" % (0, cr))
        # campi impostazioni cifre decimali primus (dalla 'a-j' e campo 'l')
        lista_dati_generali_PWE.append("@9a{0}.{1}|{2}{3}".format(8, dec, 0, cr))
        lista_dati_generali_PWE.append("@9b{0}.{1}|{2}{3}".format(8, dec, 0, cr))
        lista_dati_generali_PWE.append("@9c{0}.{1}|{2}{3}".format(9, dec_b, 0, cr))
        lista_dati_generali_PWE.append("@9d{0}.{1}|{2}{3}".format(9, dec_b, 0, cr))
        lista_dati_generali_PWE.append("@9e{0}.{1}|{2}{3}".format(10, dec, 1, cr))
        lista_dati_generali_PWE.append("@9f{0}.{1}|{2}{3}".format(14, dec, 1, cr))
        lista_dati_generali_PWE.append("@9g{0}.{1}|{2}{3}".format(18, dec, 1, cr))
        lista_dati_generali_PWE.append("@9h{0}.{1}|{2}{3}".format(10, dec, 1, cr))
        lista_dati_generali_PWE.append("@9i{0}.{1}|{2}{3}".format(14, dec, 1, cr))
        lista_dati_generali_PWE.append("@9j{0}.{1}|{2}{3}".format(7, 3, 0, cr))
        lista_dati_generali_PWE.append("@9k{0}{1}".format(valuta, cr))
        lista_dati_generali_PWE.append("@9l{0}.{1}|{2}{3}".format(11, 3, 1, cr))
        # crea una lista di voci di elenco prezzi per sezione 'listino' del file PWE
        lista_listino_PWE = list()
        campo_N = 0
        campo_F = 0
        campo_O = 0
        campo_U = ""
        n_epu = 1
        for articolo in lista_articoli_epu:
            lista_listino_PWE.append("@V%s%s" % (articolo.tariffa, cr))
            lista_listino_PWE.append("@A%s%s" % (articolo.tariffa, cr))
            lista_listino_PWE.append("@I%s%s" % (n_epu, cr))
            lista_listino_PWE.append("@R%s%s" % (articolo.descrizione_codice, cr))
            lista_listino_PWE.append("@D%s%s" % (articolo.descrizione_estesa, cr))
            lista_listino_PWE.append("@U%s%s" % (articolo.unita_misura, cr))
            lista_listino_PWE.append("@H%s%s" % (articolo.supercapitolo, cr))
            lista_listino_PWE.append("@J%s%s" % (articolo.capitolo, cr))
            lista_listino_PWE.append("@K%s%s" % (articolo.subcapitolo, cr))
            lista_listino_PWE.append("@N%s%s" % (campo_N, cr))
            # sistemazione della data x primus
            if articolo.data is None:
                articolo.data = str(datetime.datetime.now())
            data = "%s/%s/%s" % (articolo.data[8:10], articolo.data[5:7], articolo.data[0:4])
            lista_listino_PWE.append("@Y%s%s" % (data, cr)) 
            lista_listino_PWE.append("@F%s%s" % (campo_F, cr))
            lista_listino_PWE.append("@O%s%s" % (campo_F, cr))
            lista_listino_PWE.append("@0B%s%s" % (articolo.descrizione_voce, cr))
            lista_listino_PWE.append("@0U%s%s" % (campo_U, cr))
            # calcolo dei prezzi unitari dei singoli costi materiali
            p1 = articolo._ArticoloComputo__calcola_prezzo_unitario (costo_materiali= articolo.costo_mat1)
            p2 = articolo._ArticoloComputo__calcola_prezzo_unitario (costo_materiali= articolo.costo_mat2)
            p3 = articolo._ArticoloComputo__calcola_prezzo_unitario (costo_materiali= articolo.costo_mat3)
            p4 = articolo._ArticoloComputo__calcola_prezzo_unitario (costo_materiali= articolo.costo_mat4)
            campo_P = "{0:.{5}f}|{1:.{5}f}|{2:.{5}f}|{3:.{5}f}|{4:.{5}f}".format( 
                      articolo.prezzo_unitario, p1, p2, p3, p4, dec)
            lista_listino_PWE.append("@P%s%s" % (campo_P, cr))
            # dizionario tariffe-codici (utile per analisi e misurazioni)
            diz_tariffe[articolo.tariffa] = n_epu
            n_epu += 1
        # crea una lista di voci di computo per sezione 'misurazioni' del file PWE
        lista_misurazioni_PWE = list()
        for articolo in lista_articoli_cmp:
            lista_misurazioni_PWE.append("@V%s%s" % (articolo.tariffa, cr))
            lista_misurazioni_PWE.append("@I%s%s" % (diz_tariffe[articolo.tariffa], cr))
            lista_misurazioni_PWE.append("@0I%s%s" % (articolo.codice, cr))
            lista_misurazioni_PWE.append("@B%s%s" % (articolo.categoria, cr))
            lista_misurazioni_PWE.append("@E%s%s" % (articolo.subcategoria, cr))
            lista_misurazioni_PWE.append("@W%s%s" % (articolo.supercategoria, cr))
            lista_misurazioni_PWE.append("@Q%s%s" % (articolo.quantita, cr))
            # sistemazione della data x primus
            if articolo.data is None:
                articolo.data = str(datetime.datetime.now())
            data = "%s/%s/%s" % (articolo.data[8:10], articolo.data[5:7], articolo.data[0:4])
            lista_misurazioni_PWE.append("@T%s%s" % (data, cr))
            # rigo misurazione
            # verificare se il campo quantita deve essere un intero
            campo_M = "%s|%.2f||||00000|%.2f" % (articolo.note, articolo.quantita, articolo.quantita)
            lista_misurazioni_PWE.append("@M%s%s" % (campo_M, cr))
        # crea una lista di voci di elementi di listino per sezione 'analisi' del file PWE
        lista_analisi_PWE = list()
        if esporta_elementi_listino:
            for articolo in lista_listino:
                #@L7|0|0|Cavo N07V-K sezione mmq 1.5|LM|12.0000|0.1102|00000
                #if articolo.tariffa in diz_tariffe:
                n_epu = diz_tariffe[articolo.tariffa]
                #TODO lista_analisi_PWE.append("@L%s|?|?|%s|%s|%f|%f|00000%s" % (
                      #n_epu, ?,?, articolo.descrizione_codice, articolo.unita_misura, 
                      #articolo.quantita, articolo.prezzo_unitario, cr))

        # apri e scrivi nel file PWE il contenuto delle precedenti liste
        intestazione_file = ["PWE (PriMus EXCHANGE) - by ACCA%s" % cr, "2.00%s" % cr, 
                             "ANSI%s" % cr, "@;Inizio Dati Generali%s" % cr]
        with open(filename, 'w') as fpwe:
            # scrivi l'intestazione del file
            fpwe.writelines(intestazione_file)
            # scrivi i dati generali
            fpwe.writelines(lista_dati_generali_PWE)
            fpwe.write("@;Fine Dati Generali%s" % cr)
            # scrivi la sezione 'listino'
            fpwe.write("@;Inizio Voci di Listino%s" % cr)
            if len(lista_listino_PWE) > 0:
                fpwe.writelines(lista_listino_PWE)
            else:
                fpwe.write("@;Non ci sono Voci di Listino%s" % cr)
            fpwe.write("@;Fine Voci di Listino%s" % cr)
            # scrivi la sezione 'misurazioni'
            fpwe.write("@;Inizio Voci di Misurazione%s" % cr)
            if len(lista_misurazioni_PWE) > 0:
                fpwe.writelines(lista_misurazioni_PWE)
            else:
                fpwe.write("@;Non ci sono Voci di Misurazione%s" % cr)
            fpwe.write("@;Fine Voci di Misurazione%s" % cr)
            # scrivi la sezione 'analisi'
            fpwe.write("@;Inizio Elementi di Analisi%s" % cr)
            if len(lista_analisi_PWE) > 0:
                fpwe.writelines(lista_analisi_PWE)
            else:
                fpwe.write("@;Non ci sono Elementi di Analisi%s" % cr)
            fpwe.write("@;Fine Elementi di Analisi%s" % cr)
            # scrivi chiusura del file PWE
            fpwe.write("@;Fine del file PWE")

    def __somma_nomi(self, list_cat, remove_cat_0=True):
        testo = str()
        if remove_cat_0:
            list_cat = list_cat[1:]
        for cat in list_cat:
            testo += "%s|" % cat[1]
        return testo[:-1]
################################################################################
############ COMPUTO - IMPORTAZIONI ############################################
################################################################################

    def import_from_dumpfile (self, filename):
        """Importa un file di testo costituito da istruzioni SQL valide, solitamente costituito da un salvataggio dell'archivio"""
        self. delete_database()
        self.save_database()
        con = self.db
        f = open(filename, 'r')
        full_dump = f.read()
        f.close()
        try:
            self.c.executescript(unicode(full_dump))
            return True
        except:
            raise PreventiviError("Errore di importazione, non è possibile importare il file: %s" % filename)
            return False

################################################################################
############ CLASSE - IMPORTAZIONE FORMATO PWE (PRIMUS) ########################
################################################################################

    def import_from_primus_pwe (self, filename):
        """routine di importazione da file di testo primus PWE"""
        # disattiva l'eliminazione degli articoli epu
        save_epu_settings = settings["elimina_da_epu_articoli_non_in_computo"]
        if settings["elimina_da_epu_articoli_non_in_computo"] == True:
            settings["elimina_da_epu_articoli_non_in_computo"] = False
        # funzione per convertire una stringa in un numero a virgola mobile
        def converte_prezzo_pwe (prezzo):
            try:
                p = float(prezzo)
                return p
            except ValueError:
                return 0.0
        # funzione per decodificare i file di testo, spesso scritti da primus con codifica 'iso-8859-1'
        def decode_line (line):
            line = line.replace("\r\n", "").replace("\n", "")
            try:
                line = unicode(line)
            except UnicodeDecodeError:
                line = unicode(line, 'iso-8859-1')
            key = line[0:2]
            value = line[2:]
            return key, value
        # funzione per importare nel computo i dati generali
        def pwe_import_dati_generali (dati_generali):
            """"""
            Supercapitoli, Capitoli, SubCapitoli = [], [], []
            Supercategorie, Categorie, Subcategorie = [], [], []
            (nome, cliente, redattore, ricarico, manodopera, sicurezza, indirizzo, 
            comune, provincia, valuta) = (None, None, None, None, None, None, None, None, None, None)
            for line in dati_generali:
                key, value = decode_line (line)
                if key == "@a":
                    comune = value # comune
                elif key == "@b":
                    provincia = value # provincia
                elif key == "@c":
                    cliente = value # Cliente A
                elif key == "@d":
                    indirizzo = value # Cliente B
                elif key == "@e":
                    nome = value # Nome preventivo
                elif key == "@f":
                    redattore = value # Intestazione del redattore computo
                elif key == "@g":
                    Supercategorie = value.split("|") # Supercategoria
                elif key == "@h":
                    Categorie = value.split("|") # categoria
                elif key == "@i":
                    Supercapitoli = value.split("|") # Supercapitolo
                elif key == "@j":
                    Subcategorie = value.split("|") # Subcategoria
                elif key == "@k":
                    Capitoli = value.split("|") # Capitolo
                elif key == "@l":
                    SubCapitoli = value.split("|") # SubCapitolo
                elif key == "@m":
                    spese_generali, utile_impresa, sicurezza = value.split("|") #spese generali | utile d'impresa | sicurezza
                    if spese_generali == str(): spese_generali = 0
                    if utile_impresa == str(): utile_impresa = 0
                    if sicurezza == str(): sicurezza = 0
                    ricarico = 1 + ((float(spese_generali) + float(utile_impresa))/100)
                    sicurezza = float(sicurezza)/100
                elif key == "@n":
                    n = value #TODO ???
                elif key == "@9":
                    # impostazione nr. cifre decimali e intere (da non importare)
                    if line[0:3] == "@9a": a = value[1:].split("|")
                    if line[0:3] == "@9b": b = value[1:].split("|")
                    if line[0:3] == "@9c": c = value[1:].split("|")
                    if line[0:3] == "@9d": d = value[1:].split("|")
                    if line[0:3] == "@9e": e = value[1:].split("|")
                    if line[0:3] == "@9f": f = value[1:].split("|")
                    if line[0:3] == "@9g": g = value[1:].split("|")
                    if line[0:3] == "@9h": h = value[1:].split("|")
                    if line[0:3] == "@9i": i = value[1:].split("|")
                    if line[0:3] == "@9j": j = value[1:].split("|")
                    # valuta (da importare)
                    if line[0:3] == "@9k": valuta = value[1:]
                    # impostazione nr. cifre decimali e intere (da non importare)
                    if line[0:3] == "@9l": l = value[1:].split("|")
            # inserisce i dati generali e i capitoli
            self.update_dati_generali (nome=nome, cliente=cliente, redattore=redattore, ricarico=ricarico, manodopera=None, 
                 sicurezza=sicurezza, indirizzo=indirizzo, comune=comune, provincia=provincia, valuta=valuta)
            if len(Supercapitoli) > 0:
                for Nome in Supercapitoli:
                    self.insert_capitoli_categorie ("Supercapitolo", Nome)
            if len(Capitoli) > 0:
                for Nome in Capitoli:
                    self.insert_capitoli_categorie ("Capitolo", Nome)
            if len(SubCapitoli) > 0:
                for Nome in SubCapitoli:
                    self.insert_capitoli_categorie ("Subcapitolo", Nome)
            if len(Supercategorie) > 0:
                for Nome in Supercategorie:
                    self.insert_capitoli_categorie ("Supercategoria", Nome)
            if len(Categorie) > 0:
                for Nome in Categorie:
                    self.insert_capitoli_categorie ("Categoria", Nome)
            if len(Subcategorie) > 0:
                for Nome in Subcategorie:
                    self.insert_capitoli_categorie ("Subcategoria", Nome)
            return
        # funzione per importare nel computo le voci di elenco prezzi
        def pwe_import_listino (listino):
            supercapitolo, capitolo, subcapitolo = (0,0,0)
            # Creo una lista per immettere i dati di epu.
            articolo = None
            lista_voci_epu = list()
            diz_tariffe_num_listino = dict()
            for line in listino:
                key, value = decode_line (line)
                if key == "@V":
                    # Crea un articolo vuoto da impostare
                    articolo = ArticoloComputo(self.db, 0, 0, 0, 0, 0, 0, value)
                    articolo.tariffa =  value #tariffa
                elif key == "@A":
                    pass # articolo - non usato
                elif key == "@I":
                    diz_tariffe_num_listino[value] = articolo.tariffa # numero di elenco prezzi (non utilizzato)
                    articolo.cod_listino = int(value) # numero di elenco prezzi (non utilizzato)
                elif key == "@R":
                    articolo.descrizione_codice = value # descrizione sintetica
                elif key == "@D":
                    articolo.descrizione_estesa = value # descrizione estesa
                elif key == "@U":
                    articolo.unita_misura = value #unita di misura
                elif key == "@H":
                    articolo.supercapitolo = int(value) #Supercapitolo
                elif key == "@J":
                    articolo.capitolo = int(value) #Capitolo
                elif key == "@K":
                    articolo.subcapitolo = int(value) #Subcapitolo
                elif key == "@N":
                    pass #TODO ??? (questo valore sembra essere sempre 0)
                elif key == "@Y":
                    pass # data
                elif key == "@F":
                    pass #TODO ??? (questo valore sembra essere sempre 0)
                elif key == "@O":
                    pass #TODO ??? (questo valore sembra essere sempre 0 oppure 1)
                elif key == "@0":
                    if line[0:3] == "@0B":
                        articolo.descrizione_voce = value[1:] # descrizione breve
                    elif line[0:3] == "@0U":
                        pass #TODO ???
                elif key == "@P":
                    prezzo1, prezzo2, prezzo3, prezzo4, prezzo5 = value.split("|",4)
                    articolo.costo_materiali = converte_prezzo_pwe(prezzo1) # prezzo unitario
                    articolo.costo_mat1 = converte_prezzo_pwe(prezzo2)
                    articolo.costo_mat2 = converte_prezzo_pwe(prezzo3)
                    articolo.costo_mat3 = converte_prezzo_pwe(prezzo4)
                    articolo.costo_mat4 = converte_prezzo_pwe(prezzo5)
                    articolo.ricarico = 1
                    articolo.tempo_inst = 0
                    # qui copia l'articolo è completo e vuol trasferito nel database
                    # trasferisco l'articolo in una lista dove poi verrà copiato nel db
                    lista_voci_epu.append(articolo)
                else: 
                    if articolo is not None: 
                        articolo.descrizione_estesa += '\n'+ key+value #unicode(line, 'iso-8859-1')
            self.insert_articoli_epu (True, *lista_voci_epu)
            return diz_tariffe_num_listino
        # funzione per importare nel computo le voci di computo (collegate all'epu)
        def pwe_import_misurazioni (misurazioni):
            supercategoria, categoria, subcategoria = (0,0,0)
            # Creo una lista per immettere i dati di computo
            lista_voci_computo = list()
            for line in misurazioni:
                key, value = decode_line (line)
                if key == "@V":
                    tariffa = value
                elif key == "@I":
                    cod_listino =  int(value) # numero di elenco prezzi 
                    # non utilizzato, eventualmente il valore può essere utilizzato per 
                    # effettuare una verifica sul corrispondente in epu
                elif line[0:3] == "@0I":
                    codice = value[1:] # numero progressivo computo
                elif key == "@B":
                    categoria = int(value) # Categoria
                elif key == "@E":
                    subcategoria = int(value) # Subcategoria
                elif key == "@W":
                    supercategoria = int(value) # Supercategoria
                elif key == "@Q":
                    quantita_totale = float(value) # Quantità totale viene computata dalle righe di misurazione
                elif key == "@T":
                    pass # data immissione
                elif key == "@M":
                    descrizione, part_ug, lung, larg, peso, mul, tot = value.split("|")
                    #composizione dell'articolo a partire dal rigo di misurazione
                    articolo = self.copia_articolo_epu(tariffa)
                    if articolo is None:
                        articolo = ArticoloComputo(self.db, 0, 0, 0, 0, 0, 0, '@!0000')
                    articolo.codice = codice
                    articolo.categoria = categoria
                    articolo.subcategoria = subcategoria
                    articolo.supercategoria = supercategoria
                    articolo.quantita = converte_prezzo_pwe(tot)
                    articolo.note = descrizione # rigo di misurazione
                    # trasferisco l'articolo in una lista dove poi verrà copiato nel db
                    lista_voci_computo.append(articolo)
            self.insert_articoli_computo (*lista_voci_computo)
            return
        # funzione per importare le voci di listino
        def pwe_import_analisi (analisi, diz_tariffe_num_listino):
            #Crea una lista per l'immissione degli elementi di listino
            for line in analisi:
                key, value = decode_line (line)
                if key == "@0":
                    if line[0:3] == "@0L":
                        # es. rigo: @0L16|12|10|0|11.3|1
                        a, b, c, d, e, f  = value.split("|", 5)
                        pass #TODO cosa fare con questi dati?
                elif key == "@L":
                    # es. rigo: @L7|0|0|Cavo N07V-K sezione mmq 1.5|LM|12.0000|0.1102|00000
                    # oppure anche: @L16|0|0|Quota di ammortamento oraria|lire|1.00|1140/10000000*35000000|00000
                    cod_epu, in_elenco, n_elenco, desc, um, quant, pu, cod = value.split("|", 7)
                    # trasformo le stringhe in valori numerici leggibili da python
                    quant = eval(quant)
                    pu = eval(pu)
                    # il valore 'cod' sembra essere sempre uguale a: 00000
                    if cod_epu in diz_tariffe_num_listino:
                        tariffa = diz_tariffe_num_listino[cod_epu]
                        art_listino = ArticoloListino(self.db, tariffa= tariffa, codice= cod,
                                                  descrizione_codice=desc, unita_misura=um, 
                                                  quantita=quant, prezzo_unitario=pu, nr=None)
                        # inserimento degli articoli in computo
                        self.insert_articoli_listino (tariffa, art_listino)
                    else:
                        raise PreventiviError("DEBUG: articolo di listino non importato perchè non collegato a tariffa: \n%s" % line)
            return
        # apertura del file in sola lettura e uso delle funzioni precedenti per
        # importare voci, articoli e dati generali
        with open(filename) as fpwe:
            lines = fpwe.readlines()
            # esegue l'importazione per singole sezioni
            pwe_import_dati_generali(lines)
            diz_tariffe_num_listino = pwe_import_listino(lines)
            pwe_import_misurazioni(lines)
            pwe_import_analisi(lines, diz_tariffe_num_listino)
        #ripristina le impostazioni iniziali
        settings["elimina_da_epu_articoli_non_in_computo"] = save_epu_settings
        return True

################################################################################
############ CLASSE - IMPORTAZIONE FORMATO SIX (XML) ###########################
################################################################################

    def import_XML (self, filename):
        """Routine di importazione di un prezziario XML formato SIX"""
        # inizializzazioe delle variabili
        lista_articoli = list() # lista in cui memorizzare gli articoli da importare
        diz_um = dict() # array per le unità di misura
        # stringhe per nome capitoli
        titolo_sup = str()
        titolo_cap = str()
        titolo_sub = str()
        # stringhe per descrizioni articoli
        desc_codice = str()
        desc_voce = str()
        desc_estesa = str()
        # effettua il parsing del file XML
        logging.debug(_("Parsing del file XML: {0}").format(filename))
        tree = ElementTree()
        tree.parse(filename)
        # ottieni l'item root
        root = tree.getroot()
        logging.debug(list(root))
        # effettua il parsing di tutti gli elemnti dell'albero XML
        iter = tree.getiterator()
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
                self.update_dati_generali (nome=None, cliente=None, 
                                           redattore=autore, 
                                           ricarico=1, 
                                           manodopera=None, 
                                           sicurezza=None, 
                                           indirizzo=None, 
                                           comune=None, provincia=None, 
                                           valuta=valuta)
            elif elem.tag == "{six.xsd}prezzario":  
                prezzarioId= elem.get("prezzarioId")
                przId= elem.get("przId")
                livelli_struttura= len(elem.get("prdStruttura").split("."))
                categoriaPrezzario= elem.get("categoriaPrezzario")
            elif elem.tag == "{six.xsd}przDescrizione":
                logging.debug(elem.get("breve"))
                # inserisci il titolo del prezziario
                self.update_dati_generali (nome=elem.get("breve"))                
            elif elem.tag == "{six.xsd}unitaDiMisura":
                um_id= elem.get("unitaDiMisuraId")
                um_sim= elem.get("simbolo")
                um_dec= elem.get("decimali")
                # crea il dizionario dell'unita di misura
                diz_um[um_id] = um_sim
            # se il tag è un prodotto fa parte degli articoli da analizzare
            elif elem.tag == "{six.xsd}prodotto":
                # verifica e ricava le sottosezioni
                sub_desc = elem.find("{six.xsd}prdDescrizione")
                sub_quot = elem.find("{six.xsd}prdQuotazione")
                # ricava gli attributi base per la costruzione dell'articolo
                prod_id = elem.get("prodottoId")
                if prod_id is not None:
                    prod_id = int(prod_id)
                tariffa= elem.get("prdId")
                # crea l'oggetto Articolo di Computo (verifica dopo se deve essere inserito o no)
                articolo = ArticoloComputo(self.get_database(), 0, 0, 0, 0, 0, 0, 
                           tariffa, codice=prod_id)
                # verifica se l'elemento è un titolo (da inserire come Capitolo)
                if elem.get("titolo") == "true" and sub_desc is not None and sub_quot is None:
                    livello = len(tariffa.split("."))
                    # descrizione o titolo capitolo
                    titolo = sub_desc.get("breve")
                    # verifica quanto il titolo è un capitolo o un sub capitolo
                    if livello == 1:
                        titolo_sup = titolo
                        if len(titolo_sup) > 70: titolo_sup = "{0}...".format(titolo_sup[:67])
                        logging.debug(_("SUPERCAPITOLO: {0}").format(titolo_sup))
                    elif livello == 2:
                        titolo_cap = titolo
                        if len(titolo_cap) > 70: titolo_cap = "{0}...".format(titolo_cap[:67])
                        logging.debug(_("CAPITOLO: {0}").format(titolo_cap))
                    elif livello == 3:
                        titolo_sub = titolo
                        if len(titolo_sub) > 70: titolo_sub = "{0}...".format(titolo_sub[:67])
                        logging.debug(_("SUBCAPITOLO: {0}").format(titolo_sub))
                # se è presente una sezione 'prdDescrizione' ricavarne le descrizioni
                if sub_desc is not None:
                    # se l'elemento è una voce memorizza le descrizioni estese, altrimenti anche la voce
                    is_voce = elem.get("voce")
                    if (is_voce == "false" or is_voce is None) and (sub_quot is None):
                        desc_codice = sub_desc.get("breve")
                        desc_estesa = sub_desc.get("estesa")
                        desc_voce = str()
                    # imposta la descr. voce nel caso in cui vi sia una quotazione o il campo: voce="true"
                    elif is_voce == "true" and (sub_quot is not None):
                        desc_voce = sub_desc.get("breve")
                    elif is_voce is None and (sub_quot is not None):
                        desc_codice = sub_desc.get("breve")
                        desc_estesa = sub_desc.get("estesa")
                        desc_voce = sub_desc.get("breve")
                    else:
                        desc_codice = sub_desc.get("breve")
                        desc_estesa = sub_desc.get("estesa")
                        desc_voce = sub_desc.get("breve")
                # se è presente una sezione 'prdQuotazione' inserire in EPU l'articolo
                if sub_quot is not None:
                    # imposta la descrizione breve e estesa
                    articolo.descrizione_codice = desc_codice
                    articolo.descrizione_estesa = desc_estesa
                    articolo.descrizione_voce = desc_voce
                    # imposta il costo e quantità dell'articolo
                    if sub_quot.get("valore") is not None:
                        articolo.costo_materiali = float(sub_quot.get("valore"))
                    if sub_quot.get("quantita") is not None:
                        articolo.quantita = float(sub_quot.get("quantita"))
                    articolo.unita_misura = diz_um.get(elem.get("unitaDiMisuraId"), self.dati_generali["unita_misura"])
                    list_nr = sub_quot.get("listaQuotazioneId")
                    # sicurezza
                    sic = elem.get("onereSicurezza")
                    if sic is not None:
                        articolo.sicurezza = float(sic)
                    # corpo d'opera
                    if elem.get("corpoDOpera") == "true":
                        corpoDOpera = True
                    else:
                        corpoDOpera = False
                    # nomi categorie e capitoli
                    articolo.nome_supercapitolo = titolo_sup
                    articolo.nome_capitolo = titolo_cap
                    articolo.nome_subcapitolo = titolo_sub
                    articolo.nome_supercategoria = str()
                    articolo.nome_categoria = str()
                    articolo.nome_subcategoria = str()
                    # aggiunge l'articolo creato alla lista
                    lista_articoli.append(articolo)
        # importazione degli articoli nel database
        logging.debug("Importazione di {0} articoli nel database".format(len(lista_articoli)))
        self.insert_articoli_from_archivio_epu(True, *lista_articoli) 

################################################################################
############ CLASSE - ARTICOLI DI COMPUTO ######################################
################################################################################

# una classe per gestire gli articoli
class ArticoloComputo(Preventivo):
    """Classe per gestire gli Articoli del computo"""
    def __init__(self, database, supercapitolo, capitolo, subcapitolo,
                 supercategoria, categoria, subcategoria, tariffa, 
                 codice=None, descrizione_codice=None, 
                 descrizione_voce=None, descrizione_estesa=None, 
                 unita_misura=None, quantita=None, ricarico=None, 
                 tempo_inst=None, costo_materiali=None, prezzo_unitario=None,
                 sicurezza=None, cod_listino=None, note=None, data= None,
                 costo_mat1=None, costo_mat2=None, costo_mat3=None, costo_mat4=None, 
                 tipo_lavori=None, cat_appalto=None, image_art=None,
                 dati_generali=None):
        # DATABASE: imposta la connessione al database
        self.db = database
        # Imposta un cursore
        self.c = self.db.cursor()
        # imposta i dati generali per l'articolo
        if dati_generali is not None:
            self.dati_generali = dati_generali
        else:
            self.dati_generali = DATI_GENERALI_DEFAULT
            (self.dati_generali["ricarico"], self.dati_generali["manodopera"],
            self.dati_generali["sicurezza"], self.dati_generali["valuta"],
            self.dati_generali["nome"], self.dati_generali["indirizzo"],
            self.dati_generali["comune"], self.dati_generali["provincia"],
            self.dati_generali["cliente"], self.dati_generali["redattore"]
            ) = self.dati_generali_list()
        # importa i settaggi della libreria Preventa
        self.__lib_settings = self.get_settings()
        # acquisisci gli altri dati
        self.codice = codice
        self.supercapitolo = int(supercapitolo)
        self.capitolo = int(capitolo)
        self.subcapitolo = int(subcapitolo)
        self.supercategoria = int(supercategoria)
        self.categoria = int(categoria)
        self.subcategoria = int(subcategoria)
        # nomi categorie e capitoli
        self.nome_supercategoria = str()
        self.nome_categoria = str()
        self.nome_subcategoria = str()
        self.nome_supercapitolo = str()
        self.nome_capitolo = str()
        self.nome_subcapitolo = str()
        # tariffa
        self.tariffa = str(tariffa)
        if descrizione_codice is None:
            self.descrizione_codice = str()
        else: self.descrizione_codice = str(descrizione_codice)
        if descrizione_voce is None:
            self.descrizione_voce = str()
        else: self.descrizione_voce = str(descrizione_voce)
        if descrizione_estesa is None:
            self.descrizione_estesa = str()
        else: self.descrizione_estesa = str(descrizione_estesa)
        if unita_misura is None:
            self.unita_misura = self.dati_generali["unita_misura"]
        else: self.unita_misura = str(unita_misura)
        if quantita is None:
            self.quantita = self.dati_generali["quantita"]
        else: self.quantita = float(quantita)
        if ricarico is None or ricarico == 0 or ricarico == 0.0:
            self.ricarico = self.dati_generali["ricarico"]
        else: self.ricarico = float(ricarico)
        if tempo_inst is None:
            self.tempo_inst = 0
        else: self.tempo_inst = float(tempo_inst)
        if costo_materiali is None:
            self.costo_materiali = 0
        else: self.costo_materiali = float(costo_materiali)
        # sicurezza
        if sicurezza is None:
            self.sicurezza = self.dati_generali["sicurezza"]
        else: self.sicurezza = float(sicurezza)
        # prezzo unitario
        if prezzo_unitario is None:
            self.prezzo_unitario = self.__calcola_prezzo_unitario()
        else:
            self.prezzo_unitario = float(prezzo_unitario)
        # lista dei codici di listino
        self.cod_listino = cod_listino
        # note
        if note is None:
            self.note = str()
        else: self.note = note
        self.data = data
        # calcola il prezzo totale
        self.prezzo_totale = self.__calcola_prezzo_totale()
        if costo_mat1 is None:
            self.costo_mat1 = self.costo_materiali
        else: self.costo_mat1 = float(costo_mat1)
        if costo_mat2 is None:
            self.costo_mat2 = 0.0
        else: self.costo_mat2 = float(costo_mat2)
        if costo_mat3 is None:
            self.costo_mat3 = 0.0
        else: self.costo_mat3 = float(costo_mat3)
        if costo_mat4 is None:
            self.costo_mat4 = 0.0
        else: self.costo_mat4 = float(costo_mat4)
        # tipologia di contratto (corpo, economia, ecc.)
        if tipo_lavori is None:
            # assegna come valore di default il primo valore della lista ( [0] )
            # o se la lista è vuota una stringa vuota
            if len(self.__lib_settings["default_contract_type"]) == 0:
                self.tipo_lavori = str()
            else:
                self.tipo_lavori = self.__lib_settings["default_contract_type"][0]
        else:
            # assegna come valore il numero corrispondente alla lista
            if type(tipo_lavori) is int:
                self.tipo_lavori = self.__lib_settings["default_contract_type"][tipo_lavori]
            elif type(tipo_lavori) is str:
                self.tipo_lavori = tipo_lavori
            else: self.tipo_lavori = str()
        # categoria dei lavori in appalto
        if cat_appalto is None:
            # assegna come valore di default il primo valore della lista ( [0] )
            # o se la lista è vuota una stringa vuota
            if len(self.__lib_settings["default_contract_category"]) == 0:
                self.cat_appalto = str()
            else:
                self.cat_appalto = self.__lib_settings["default_contract_category"][0]
        else:
            # se il valore è diverso da 'None' attribuisce l'indice della lista
            # se il valore è 'int' oppure il valore assegnato se è una 'str'
            if type(cat_appalto) is int:
                self.cat_appalto = self.__lib_settings["default_contract_category"][cat_appalto]
            elif type(cat_appalto) is str:
                self.cat_appalto = cat_appalto
            else: self.cat_appalto = str()
        # per futura implementazione di immagine collegata all'articolo
        self.image_art=image_art
        # attributo (non assegnato in fase di creazione dell'oggetto) a disposizione
        # per memorizzare gli articoli di listino relativi all'articolo di computo
        self.art_listino = list()

    def __calcola_prezzo_unitario (self, costo_materiali = None):
        """calcola il prezzo unitario dell'articolo di computo. 
           IMPORTANTE!!! il costo della manodopera si intende all'ora,
           mentre il tempo di installazione è in minuti!!"""
        # se il parametro costo materiali è Nullo il valore è calcolato utilizzando self.costo_materiali
        if costo_materiali is None: costo_materiali = self.costo_materiali
        # ricavo il costo della manodopera
        costo_manodopera_ora = self.dati_generali_list()[1]
        return ((((costo_manodopera_ora/60) * self.tempo_inst) + (costo_materiali * self.ricarico)) * (1+(self.sicurezza/100)))

    def __calcola_prezzo_totale (self):
        """calcola il prezzo totale"""
        self.prezzo_unitario = self.__calcola_prezzo_unitario()
        return self.quantita * self.prezzo_unitario

    def row(self, codice=None):
        """
        This method return a list of attributes suitable to be used with 
        sqlite .execute() instance
        """
        self.prezzo_totale = self.__calcola_prezzo_totale()
        if codice is not None and type(codice) == int:
            self.codice = codice
        dizionario = {"id":codice,
                      "SupCap":self.supercapitolo, "Cap":self.capitolo, 
                      "SubCap":self.subcapitolo,
                      "SupCat":self.supercategoria, "Cat":self.categoria, 
                      "SubCat":self.subcategoria,
                      "Tar":self.tariffa, 
                      "DesCod":self.descrizione_codice,
                      "DesVoc":self.descrizione_voce, 
                      "DesEst":self.descrizione_estesa,
                      "UM":self.unita_misura, 
                      "Quant":self.quantita,
                      "Ric":self.ricarico, 
                      "Temp":self.tempo_inst,
                      "PrezzoUnit":self.prezzo_unitario, 
                      "PrezzoTot": self.prezzo_totale, 
                      "CostoMat":self.costo_materiali,
                      "Sicurezza":self.sicurezza, 
                      "cod_listino":self.cod_listino, 
                      "Note":self.note,
                      "CostoMat1":self.costo_mat1, "CostoMat2":self.costo_mat2, 
                      "CostoMat3":self.costo_mat3, "CostoMat4":self.costo_mat4,
                      "TipoLav":self.tipo_lavori, 
                      "CatApp":self.cat_appalto, 
                      "Image":self.image_art}
        return dizionario

    def set_chapters_name (self, supercapitolo=None, capitolo=None, subcapitolo=None):
        """set the chapters name attribute"""
        if supercapitolo is None:
            self.nome_supercapitolo = str()
        else:
            self.nome_supercapitolo = supercapitolo
        if capitolo is None:
            self.nome_capitolo = str()
        else:
            self.nome_capitolo = capitolo
        if subcapitolo is None:
            self.nome_subcapitolo = str()
        else:
            self.nome_subcapitolo = subcapitolo
        return True

    def set_categorys_name (self, supercategoria=None, categoria=None, subcategoria=None):
        """set the categorys name attribute"""
        if supercategoria is None:
            self.nome_supercategoria = str()
        else:
            self.nome_supercategoria = supercategoria
        if categoria is None:
            self.nome_categoria = str()
        else:
            self.nome_categoria = categoria
        if subcategoria is None:
            self.nome_subcategoria = str()
        else:
            self.nome_subcategoria = subcategoria
        return True
   
    def set_unit_price(self, value):
        """
        When value is given set the unit_price by lowering the installation time or
        if it would be negative lower the materials costs.
        """
        # verifica che il valore 'value' sia un numero a doppia precisione
        if type(value) != float:
            raise PreventiviError("'value' deve essere 'float' invece di {0} ".format(type(value)))
            return False
        # se il costo materiali è pari a 0 fissa il costo pari al nuovo valore
        if self.costo_materiali == 0: cm = value
        else: cm = None
        # ricavo il costo della manodopera
        costo_mo_ora = self.dati_generali_list()[1]
        # calcolo la differenza tra i due prezzi per poi ridurre il tempo di installazione
        diff = self.prezzo_unitario - value
        # calcolo il nuovo tempo di installazione
        try:
            self.tempo_inst = (value - ((self.costo_materiali * self.ricarico) * (1+(self.sicurezza/100))))/(costo_mo_ora/60)
        except ZeroDivisionError:
            self.tempo_inst = 0
        # se il tempo di installazione è negativo riduco il costo dei materiali
        if self.tempo_inst < 0:
            self.tempo_inst = 0
            try:
                self.costo_materiali = (value/(1+(self.sicurezza/100)))/ self.ricarico
            except ZeroDivisionError:
                self.costo_materiali = value
        # ricalcolo il nuovo prezzo unitario
        self.prezzo_unitario = self.__calcola_prezzo_unitario ()
        # infine calcola il prezzo totale
        self.prezzo_totale = self.__calcola_prezzo_totale()
        return True

    def __str__(self):
        """Metodo per la conversione dell'articolo in una stringa"""
        # se il codice non è 'valido' viene ignorato
        if self.codice is not None or type(self.codice) == int:
            self.codice = None
        # verifica se sono memorizzati codici di listino e li 'atomizza'
        if self.art_listino != []:
            lista_art_listino = list()
            for articolo_listino in self.art_listino:
                lista_art_listino.append(str(articolo_listino))
            self.art_listino = lista_art_listino
        # compone la stringa per "l'atomizzazione" dell'articolo
        stringa = """ArtCmp|%s|%s|%s|%s|%s|%d|%s|%d|%s|%d|%s|%d|%s|%d|%s|%d|%s|%s|%f|%f|%f|%f|%f|%f|%s|%s|%s|%f|%f|%f|%f|%f|%s|%s|%s""" % (
                  self.tariffa, self.codice, self.descrizione_codice, 
                  self.descrizione_voce, self.descrizione_estesa,
                  self.supercapitolo, self.nome_supercapitolo, 
                  self.capitolo, self.nome_capitolo,
                  self.subcapitolo, self.nome_subcapitolo, 
                  self.supercategoria, self.nome_supercategoria, 
                  self.categoria, self.nome_categoria, 
                  self.subcategoria, self.nome_subcategoria,
                  self.unita_misura, self.quantita, self.ricarico, 
                  self.tempo_inst, self.costo_materiali, self.prezzo_unitario,
                  self.sicurezza, self.cod_listino, self.note, self.data, 
                  self.prezzo_totale, self.costo_mat1, self.costo_mat2, 
                  self.costo_mat3, self.costo_mat4, self.tipo_lavori,
                  self.cat_appalto, repr(self.art_listino))
        return stringa

def convert_string_to_articolo (stringa, database):
    """Converte una stringa (es. output dela funz. str(ArticoloComputo)) in un articolo della classe ArticoloComputo"""
    # separa i campi della stringa e assegna i campi all'articolo
    (intestazione, tariffa, codice, descrizione_codice, 
    descrizione_voce, descrizione_estesa, supercapitolo, nome_supercapitolo, 
    capitolo, nome_capitolo, subcapitolo, nome_subcapitolo, 
    supercategoria, nome_supercategoria, categoria, nome_categoria, 
    subcategoria, nome_subcategoria, unita_misura, quantita, ricarico, 
    tempo_inst, costo_materiali, prezzo_unitario, sicurezza, cod_listino, 
    note, data, prezzo_totale, costo_mat1, costo_mat2, costo_mat3, costo_mat4, 
    tipo_lavori, cat_appalto, art_listino) = stringa.split("|") 
    # inizializza l'articolo con i campi ricavati
    articolo = ArticoloComputo(database=database, supercapitolo=supercapitolo,  
                       capitolo=capitolo, subcapitolo=subcapitolo, 
                       supercategoria=supercategoria, categoria=categoria, 
                       subcategoria=subcategoria, tariffa=tariffa,
                       codice=codice, descrizione_codice=descrizione_codice, 
                       descrizione_voce=descrizione_voce, 
                       descrizione_estesa=descrizione_estesa, 
                       unita_misura=unita_misura, quantita=quantita, 
                       ricarico=ricarico, tempo_inst=tempo_inst, 
                       costo_materiali=costo_materiali, 
                       prezzo_unitario=prezzo_unitario, sicurezza=sicurezza, 
                       cod_listino=cod_listino, note=note, data= data,
                       costo_mat1=costo_mat1, costo_mat2=costo_mat2, 
                       costo_mat3=costo_mat3, costo_mat4=costo_mat4,
                       tipo_lavori=tipo_lavori, cat_appalto=cat_appalto)
    articolo.nome_supercapitolo = nome_supercapitolo
    articolo.nome_capitolo = nome_capitolo
    articolo.nome_subcapitolo = nome_subcapitolo
    articolo.nome_supercategoria = nome_supercategoria
    articolo.nome_categoria = nome_categoria
    articolo.nome_subcategoria = nome_subcategoria
    # se ci sono stringhe articoli di listino li converto
    if eval(art_listino) is not None:
        lista_str_listino = eval(art_listino)
        for stringa in lista_str_listino:
            listino = convert_string_to_listino (stringa, database)
            articolo.art_listino.append(listino)
    return articolo

################################################################################
############ CLASSE - ARTICOLI DI LISTINO ######################################
################################################################################

# una classe per gestire gli articoli
class ArticoloListino(Preventivo):
    """Classe per gestire gli Articoli di listino"""
    def __init__(self, database, tariffa, codice, descrizione_codice, unita_misura, 
                       quantita, prezzo_unitario, accessori=None, sconto=None, nr=None, 
                       note=None):
        # DATABASE: imposta la connessione al database
        self.db = database
        # Imposta un cursore
        self.c = self.db.cursor()
        # Imposta gli altri dati
        self.tariffa = tariffa
        self.codice = codice
        self.descrizione_codice = str(descrizione_codice)
        self.unita_misura = str(unita_misura)
        self.quantita = float(quantita)
        self.prezzo_unitario = float(prezzo_unitario)
        if accessori is None:
            self.accessori = 0.0
        else: self.accessori = float(accessori)
        if sconto is None:
            self.sconto = 0.0
        else: self.sconto = float(sconto)
        if note is None:
            self.note = str()
        else: self.note = note
        self.nr = nr #numero progressivo di inserimento nel database 'Analisi'
        # calcola il prezzo totale
        self.prezzo_tot = self.__calcola_prezzo()

    def __calcola_prezzo (self):
        """calcola il prezzo totale"""
        return self.prezzo_unitario * self.quantita * (1 - (self.sconto/100)) + self.accessori

    def row(self, nr=None):
        """This method return a list of attributes suitable to be used with sqlite .execute instance"""
        self.prezzo_tot = self.__calcola_prezzo()
        if nr is not None:
            self.nr = nr
        dizionario = {"id":self.nr, "Tariffa":self.tariffa, "Codice":self.codice, "DesCod":self.descrizione_codice, \
                      "UM":self.unita_misura, "Quantita":self.quantita, "PrezzoUnit":self.prezzo_unitario, \
                      "Accessori":self.accessori, "Sconto":self.sconto, "PrezzoTot":self.prezzo_tot, "Note":self.note}
        return dizionario

    def __str__(self):
        """Metodo di stampa dell'articolo"""
        stringa = "ArtLis§%s§%s§%s§%s§%f§%f§%f§%f§%s§%s§%f" % (self.tariffa, self.codice, self.descrizione_codice, 
                                                     self.unita_misura, self.quantita, self.prezzo_unitario, self.accessori, 
                                                     self.sconto, self.note, self.nr, self.prezzo_tot)
        return stringa

def convert_string_to_listino (stringa, database):
    """Converte una stringa (es. output dela funz. str(ArticoloListino)) in un articolo della classe ArticoloListino"""
    # separa i campi della stringa e assegna i campi all'articolo
    (intestazione, tariffa, codice, descrizione_codice, unita_misura, quantita,  
     prezzo_unitario, accessori, sconto, note, nr, prezzo_tot) = stringa.split("§")
    # inizializza l'articolo con i campi ricavati
    listino = ArticoloListino(database, tariffa, codice, descrizione_codice,  
                              unita_misura, quantita, prezzo_unitario, 
                              accessori=accessori, sconto=sconto, nr=nr, note=note)
    return listino


################################################################################
###################### SHELL INTERATTIVA #######################################
################################################################################
# A minimal shell for experiments
def shell(preventivo=None):
    """Shell minimale per il test della libreria preventARES"""
    if preventivo is None:
        preventivo = ":memory:"
    Preventa = Preventivo(preventivo)
    print """\nSHELL - Preventa_lib\n
    Database used: '{0}'

    Enter your commands to execute in Preventa_lib.
    Enter 'dir' or 'help' to list commands.
    Enter a blank line to exit.""".format(preventivo)
    while True:
        line = raw_input("$preventa_lib.")
        if line == "":
            break
        elif line == "dir" or line == "help":
            print dir(Preventa)
        elif line.startswith("doc("):
            print eval("Preventa."+line+".__doc__")
        else:
            try:
                print eval("Preventa."+line)
            except:
                print "An error occurred:", sys.exc_info()
    Preventa.connection_shutdown()
    print "Exit Shell"

################################################################################
###################### FUNZIONI DI TEST DELLA LIBRERIA #########################
################################################################################
def show(file_preventivo):
     """funzione di test per stampa a video di tutte le tabelle del database"""
     preventivo = Preventivo(file_preventivo)
     print "\nDati Generali del preventivo: \n"
     print preventivo.dati_generali_list()
     print "\nCategorie di Computo: \n"
     for elemento in preventivo.categorie_rows_list():
        print elemento
     print "\nCapitoli di Epu: \n"
     for elemento in preventivo.capitoli_rows_list():
        print elemento
     print "\nComputo: \n"
     for articolo in preventivo.computo_rows_list(): 
        print articolo
     print "\nEpu: \n"
     for articolo in preventivo.epu_rows_list(): 
        print articolo
     print "\nTavola listino: \n"
     for articolo in preventivo.listino_rows_list(): 
        print articolo
        print "\n"
     print "\nTabella manodopera: \n"
     for articolo in preventivo.get_table_manodopera(): 
        print articolo
     preventivo.connection_shutdown()

def test(preventivo=None):
    """Funzione di test della libreria preventARES"""
    print "\nInizio Test della libreria preventa_lib\n"
    if preventivo is None:
        preventivo = ":memory:"
    Nuovo_preventivo = Preventivo(preventivo)
    db = Nuovo_preventivo.get_database()
    #articoli di prova:
    articolo1 = ArticoloComputo(codice = 1, supercapitolo = 1, capitolo = 1, subcapitolo = 1, supercategoria = 1, categoria = 1, subcategoria = 1, 
                       tariffa= "A110", descrizione_codice="Articolo 1", unita_misura="N", quantita=10, costo_materiali=22.5, tempo_inst=60, database=db)
    articolo2 = ArticoloComputo(codice = 2, supercapitolo = 1, capitolo = 1, subcapitolo = 1, supercategoria = 1, categoria = 1, subcategoria = 1, 
                       tariffa= "A101", descrizione_codice="Articolo 2", unita_misura="N", quantita=20, costo_materiali=100.0, tempo_inst=60, database=db)
    articolo3 = ArticoloComputo(codice = None, supercapitolo = 1, capitolo = 2, subcapitolo = 1, supercategoria = 1, categoria = 1, subcategoria = 1, 
                       tariffa= "B100", descrizione_codice="Articolo 3", unita_misura="N", quantita=15, costo_materiali=120.0, tempo_inst=60, database=db)

    articolo4 = ArticoloComputo(codice = None, supercapitolo = 1, capitolo = 2, subcapitolo = 1, supercategoria = 1, categoria = 1, subcategoria = 1, 
                       tariffa= "E100", descrizione_codice="Articolo 4", unita_misura="N", quantita=10, costo_materiali=135.5, tempo_inst=60, database=db)
    # articoli di listino di prova
    listino1 = ArticoloListino(tariffa = None, codice = "SMU2050", descrizione_codice = "cassetta da incasso 196x118", unita_misura="N", 
                       quantita= 2, prezzo_unitario= 3.25, accessori=None, sconto=None, note=None, database=db)
    listino2 = ArticoloListino(tariffa = None, codice = "BTI    457850", descrizione_codice = "interruttore unipolare LIVING", unita_misura="N", 
                       quantita= 1, prezzo_unitario= 6.5, accessori=None, sconto=10, note=None,database=db)
    listino3 = ArticoloListino(tariffa = None, codice = "GEW 30012", descrizione_codice = "quadretto con portella trasp.", unita_misura="N", 
                       quantita= 3, prezzo_unitario= 6.20, accessori=None, sconto=None, note=None,database=db)
    listino4 = ArticoloListino(tariffa = None, codice = "BTI    123456", descrizione_codice = "interruttore MTD 2x6 6kA 0,03", unita_misura="N", 
                       quantita= 15, prezzo_unitario= 6.5, accessori=None, sconto=10, note=None,database=db)
    Nuovo_preventivo.insert_articoli_computo(articolo1, articolo2, articolo3, articolo4)
    Nuovo_preventivo.insert_articoli_listino(listino1, listino2, listino3)
    Nuovo_preventivo.connection_shutdown()
    show(preventivo)
    print "End Test"

if __name__ == "__main__":
    test()
