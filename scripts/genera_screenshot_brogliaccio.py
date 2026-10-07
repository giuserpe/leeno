#!/usr/bin/env python3
"""Genera gli screenshot di Brogliaccio usati nel manuale (dati di esempio fittizi).

Uso: python3 scripts/genera_screenshot_brogliaccio.py CARTELLA_OUTPUT
Produce: brog_principale.png, brog_giornata.png, brog_stampa.png
Richiede: playwright (con Chromium), Pillow, pdftoppm.
"""
import json
import os
import subprocess
import sys
import time

from PIL import Image
from playwright.sync_api import sync_playwright

RADICE = os.path.normpath(os.path.join(os.path.dirname(__file__), '..'))
APP = os.path.join(RADICE, 'tools', 'appunti-cantiere')
CHIAVE = 'appunti-cantiere-v1'
PORTA = '8766'

STATO = {
    'cantieri': {'a': {
        'nome': 'Via Roma 10 - Ristrutturazione',
        'ultimo_export': None,
        'giornate': {
            '2026-10-02': {'campi': {'meteo': 'Sereno', 'annotazioni': 'Getto del solaio al primo piano.'},
                           'modificato_il': '2026-10-02T10:00:00Z'},
            '2026-10-05': {'campi': {'meteo': 'Nuvoloso', 'annotazioni': 'Posa dei tubi di scarico.'},
                           'modificato_il': '2026-10-05T10:00:00Z'},
            '2026-10-06': {'campi': {}, 'modificato_il': '2026-10-06T10:00:00Z'},
        }}},
    'attivo': 'a',
}


def ritaglia(percorso, altezza_max):
    im = Image.open(percorso).convert('RGB')
    if im.height > altezza_max:
        im = im.crop((0, 0, im.width, altezza_max))
    im.save(percorso, optimize=True)


def main(out):
    os.makedirs(out, exist_ok=True)
    srv = subprocess.Popen(['python3', '-m', 'http.server', PORTA, '-d', APP],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(1)
    try:
        with sync_playwright() as p:
            b = p.chromium.launch()
            ctx = b.new_context(viewport={'width': 390, 'height': 780}, device_scale_factor=2, locale='it-IT')
            pg = ctx.new_page()
            pg.goto('http://localhost:%s/index.html' % PORTA)
            pg.evaluate('([k,v])=>localStorage.setItem(k,v)', [CHIAVE, json.dumps(STATO)])
            pg.reload(); time.sleep(.5)
            f1 = os.path.join(out, 'brog_principale.png')
            pg.screenshot(path=f1); ritaglia(f1, 1515)
            pg.click('ul.lista li button'); time.sleep(.3)  # giornata piu' recente
            pg.fill('#c_meteo', 'Sereno')
            pg.fill('#c_annotazioni', 'Posa dei tubi di scarico al piano terra.')
            pg.fill('#c_presenti', 'Impresa, D.L.')
            f2 = os.path.join(out, 'brog_giornata.png')
            pg.evaluate('document.activeElement.blur()')
            pg.screenshot(path=f2); ritaglia(f2, 1560)
            pg.click('text=Chiudi'); time.sleep(.3)
            pg.evaluate('window.print = function(){}')
            pg.click('text=Stampa o PDF'); time.sleep(.8)
            pdf = os.path.join(out, 'brog_stampa.pdf')
            pg.pdf(path=pdf, format='A4', print_background=True)
            b.close()
        subprocess.run(['pdftoppm', '-png', '-r', '110', '-f', '1', '-l', '1', '-singlefile', pdf,
                        os.path.join(out, 'brog_stampa')], check=True)
        os.remove(pdf)
        im = Image.open(os.path.join(out, 'brog_stampa.png')).convert('RGB')
        im = im.crop((0, 0, im.width, 690))
        im.save(os.path.join(out, 'brog_stampa.png'), optimize=True)
    finally:
        srv.terminate()


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else '.')
