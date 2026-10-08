#!/usr/bin/env python3
"""
Aggiorna i numeri di pagina dell'Indice generale di documentazione/MANUALE_LeenO.fodt.

Come funziona: LibreOffice (headless, via UNO) carica una COPIA temporanea del manuale,
aggiorna gli indici e la salva; dall'indice aggiornato si leggono i numeri di pagina e si
riportano, voce per voce (chiave: bookmark __RefHeading__...), nel FODT originale. Il resto
del file non viene toccato, quindi nessuna riscrittura dello XML da parte di LibreOffice.
Si ripete finche' i numeri non cambiano piu' (di solito una sola passata).

Uso: python3 .agent/skills/leeno-aggiorna-manuale/scripts/aggiorna_indice.py
Richiede: LibreOffice con il modulo Python `uno` (python3-uno) e `soffice` nel PATH.
Eseguirlo DOPO ogni modifica al manuale e PRIMA di genera_pdf.py.
"""
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', '..', '..', '..'))
FODT = os.path.join(REPO_ROOT, 'documentazione', 'MANUALE_LeenO.fodt')
VOCE = re.compile(r'(xlink:href="#([^"]+)"[^>]*>(?:(?!</text:a>).)*?<text:tab/>)(\d+)(</text:a>)', re.S)


def leggi(percorso):
    with open(percorso, encoding='utf-8', newline='') as f:
        return f.read()


def limiti_indice(testo):
    i = testo.index('<text:table-of-content ')
    return i, testo.index('</text:table-of-content>', i)


def pagine(testo):
    i, e = limiti_indice(testo)
    return {m.group(2): int(m.group(3)) for m in VOCE.finditer(testo[i:e])}


def aggiorna_con_libreoffice(sorgente, uscita):
    import uno
    from com.sun.star.beans import PropertyValue
    soffice = shutil.which('soffice')
    if not soffice:
        sys.exit('soffice non trovato nel PATH')
    with socket.socket() as s:
        s.bind(('localhost', 0))
        porta = s.getsockname()[1]
    profilo = tempfile.mkdtemp(prefix='lo_indice_')
    proc = subprocess.Popen([soffice, '--headless', '--norestore', '-env:UserInstallation=file://' + profilo.replace('\\', '/'),
                             '--accept=socket,host=localhost,port=%d;urp;' % porta],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        ctx = None
        for _ in range(90):
            try:
                locale = uno.getComponentContext()
                res = locale.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', locale)
                ctx = res.resolve('uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext' % porta)
                break
            except Exception:
                time.sleep(1)
        if ctx is None:
            sys.exit('Impossibile collegarsi a LibreOffice')
        desk = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)

        def pv(nome, valore):
            p = PropertyValue()
            p.Name, p.Value = nome, valore
            return p
        doc = desk.loadComponentFromURL(uno.systemPathToFileUrl(sorgente), '_blank', 0, (pv('Hidden', True),))
        indici = doc.getDocumentIndexes()
        for k in range(indici.getCount()):
            indici.getByIndex(k).update()
        doc.storeToURL(uno.systemPathToFileUrl(uscita), (pv('FilterName', 'OpenDocument Text Flat XML'),))
        doc.close(True)
        try:
            desk.terminate()
        except Exception:
            pass
    finally:
        try:
            proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            proc.kill()
        shutil.rmtree(profilo, ignore_errors=True)


def main():
    cambiate_tot = 0
    with tempfile.TemporaryDirectory() as tmp:
        for passata in range(1, 5):
            copia = os.path.join(tmp, 'copia.fodt')
            uscita = os.path.join(tmp, 'aggiornato.fodt')
            shutil.copyfile(FODT, copia)
            aggiorna_con_libreoffice(copia, uscita)
            nuove = pagine(leggi(uscita))
            testo = leggi(FODT)
            attuali = pagine(testo)
            if set(nuove) != set(attuali):
                sys.exit('Le voci dell\'indice non coincidono (titoli aggiunti o rimossi): '
                         'aggiornare l\'indice da LibreOffice (Strumenti > Aggiorna > Tutti gli indici)')
            da_cambiare = [h for h in attuali if attuali[h] != nuove[h]]
            print('Passata %d: %d voci da aggiornare su %d' % (passata, len(da_cambiare), len(attuali)))
            if not da_cambiare:
                break
            i, e = limiti_indice(testo)
            corpo = VOCE.sub(lambda m: m.group(1) + str(nuove[m.group(2)]) + m.group(4), testo[i:e])
            with open(FODT, 'w', encoding='utf-8', newline='') as f:
                f.write(testo[:i] + corpo + testo[e:])
            cambiate_tot += len(da_cambiare)
        else:
            sys.exit('I numeri di pagina non si stabilizzano: controllare il manuale')
    print('Indice aggiornato (%d modifiche ai numeri di pagina).' % cambiate_tot)


if __name__ == '__main__':
    main()
