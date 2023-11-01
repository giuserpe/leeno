'''
Funzioni di utilità extra
'''
# ~import subprocess
# ~import LeenoDialogs as DLG

import xml.etree.ElementTree as ET
import Dialogs
import LeenoUtils
import pyleeno as PL

def ricevuta_pec():
    '''
    Copia i dati dal file xml di accettazione / avvenuta-consegna PEC in un nuovo foglio
    '''
    try:
        nfile = Dialogs.FileSelect(est='*.xml')
        tree = ET.parse(nfile)
        root = tree.getroot()

        ricevuta_elem = root.find("dati/ricevuta")
        tipo_ricevuta = ricevuta_elem.get("tipo") if ricevuta_elem is not None else ''

        dati = {
            "Tipo": root.get("tipo", ""),
            "Errore": root.get("errore", ""),
            "Mittente": root.findtext("intestazione/mittente"),
            "Destinatario Certificato": root.findtext("intestazione/destinatari[@tipo='certificato']"),
            "Destinatario Esterno": root.findtext("intestazione/destinatari[@tipo='esterno']"),
            "Risposte": root.findtext("intestazione/risposte"),
            "Oggetto": root.findtext("intestazione/oggetto"),
            "Gestore Emittente": root.findtext("dati/gestore-emittente"),
            "Data": root.findtext("dati/data/giorno"),
            "Ora": root.findtext("dati/data/ora"),
            "Identificativo": root.findtext("dati/identificativo"),
            "Message ID": root.findtext("dati/msgid"),
            "Tipo Ricevuta": tipo_ricevuta,
            "Consegna": root.findtext("dati/consegna") or ''
        }

        # Converte il dizionario in una lista di coppie chiave-valore
        dati_lista = list(dati.items())
        oDoc = LeenoUtils.getDocument()
        nome = 'PEC_'+ dati['Tipo']
        try:
            oDoc.Sheets.insertNewByName(nome, 100)
        except:
            pass
        PL.GotoSheet(nome)
        oSheet = oDoc.CurrentController.ActiveSheet

        # Imposta i dati nella riga 0 del foglio
        for idx, (key, value) in enumerate(dati_lista, start=0):
            oSheet.getCellByPosition(0, idx).String = key
            oSheet.getCellByPosition(1, idx).String = value

        # Converti il foglio di calcolo in PDF
        # ~convert_to_pdf(pdf_output_file)
        oSheet.getCellRangeByName('A1:B1048576').Columns.OptimalWidth = True

    except ET.ParseError as e:
        DLG.chi(f"Errore durante il parsing del file XML: {e}")
    
    return

########################################################################

def convert_to_pdf(output_file):
    '''
    da testare
    '''
    try:
        # Comando per convertire il foglio di calcolo in PDF utilizzando unoconv
        command = f"soffice --headless --convert-to pdf:calc_pdf_Export --outdir {output_file} {output_file}.ods"


        DLG.chi(command)
        # Esegui il comando di shell
        subprocess.run(command, shell=True)
        print(f"File PDF creato con successo: {output_file}")
    except Exception as e:
        print(f"Errore durante la conversione in PDF: {e}")


    
    return
