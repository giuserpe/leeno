# Copyright (C) Giuseppe Vizziello - supporto@leeno.org. Licenza LGPL 2.1 (vedi LICENSE nella radice del repository).
"""Test di integrazione Playwright per l'esportazione e condivisione file .zip in Brogliaccio."""

import json
import os
import subprocess
import sys
import time
import pytest
from playwright.sync_api import sync_playwright

RADICE = os.path.normpath(os.path.join(os.path.dirname(__file__), '..'))
APP = os.path.join(RADICE, 'tools', 'appunti-cantiere')
CHIAVE = 'appunti-cantiere-v1'
PORTA = '8767'

STATO_SINGOLO = {
    'cantieri': {
        'c1': {
            'nome': 'Cantiere Test Zip',
            'ultimo_export': None,
            'giornate': {
                '2026-10-10': {
                    'campi': {'meteo': 'Sereno', 'annotazioni': 'Lavori di prova per export zip.'},
                    'modificato_il': '2026-10-10T09:00:00Z'
                }
            }
        }
    },
    'attivo': 'c1',
    'rev': 1
}

STATO_SENSIVILE = {
    'cantieri': {
        'c1': {
            'nome': 'Cantiere Privacy Zip',
            'ultimo_export': None,
            'giornate': {
                '2026-10-10': {
                    'campi': {
                        'meteo': 'Sereno',
                        'annotazioni': 'Lavori con infortunio.',
                        'infortuni': 'Infortunio ad un operaio'
                    },
                    'modificato_il': '2026-10-10T09:00:00Z'
                }
            }
        }
    },
    'attivo': 'c1',
    'rev': 1
}

STATO_MOLTEPLICI = {
    'cantieri': {
        'c1': {
            'nome': 'Cantiere A',
            'ultimo_export': None,
            'giornate': {
                '2026-10-10': {
                    'campi': {'meteo': 'Sereno', 'annotazioni': 'Lavori A.'},
                    'modificato_il': '2026-10-10T09:00:00Z'
                }
            }
        },
        'c2': {
            'nome': 'Cantiere B',
            'ultimo_export': None,
            'giornate': {
                '2026-10-11': {
                    'campi': {'meteo': 'Pioggia', 'annotazioni': 'Lavori B.'},
                    'modificato_il': '2026-10-11T09:00:00Z'
                }
            }
        }
    },
    'attivo': 'c1',
    'rev': 1
}

INIT_SCRIPT = """
    window.__sharedData = [];
    navigator.canShare = function(data) { return true; };
    navigator.share = function(data) {
        window.__sharedData.push({
            title: data.title,
            files: data.files ? Array.from(data.files).map(f => ({ name: f.name, type: f.type, size: f.size })) : []
        });
        return Promise.resolve();
    };
"""


@pytest.fixture(scope="module")
def server_app():
    srv = subprocess.Popen([sys.executable, '-m', 'http.server', PORTA, '-d', APP],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(0.5)
    yield "http://localhost:%s/index.html" % PORTA
    srv.terminate()


def test_export_zip_singolo_cantiere_con_foto(server_app):
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(viewport={'width': 390, 'height': 780}, locale='it-IT')
        ctx.add_init_script(INIT_SCRIPT)
        
        page = ctx.new_page()
        page.goto(server_app)
        page.evaluate('([k,v])=>localStorage.setItem(k,v)', [CHIAVE, json.dumps(STATO_SINGOLO)])
        
        page.evaluate("""async () => {
            if (window.FotoStore) {
                const jpegDummy = new Uint8Array([0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01, 0xFF, 0xD9]);
                const blob = new Blob([jpegDummy], { type: 'image/jpeg' });
                await window.FotoStore.aggiungi('c1', '2026-10-10', blob);
            }
        }""")
        
        page.reload()
        page.wait_for_selector('text=Esporta per LeenO')
        page.click('text=Esporta per LeenO')
        
        page.wait_for_selector('#pannello-export')
        pannello_testo = page.text_content('#pannello-export')
        assert 'File pronto:' in pannello_testo
        assert '.zip' in pannello_testo
        
        page.click('text=Invia o condividi')
        page.wait_for_timeout(300)
        
        shared = page.evaluate('window.__sharedData')
        assert len(shared) == 1
        assert shared[0]['files'][0]['name'].endswith('.zip')
        
        browser.close()


def test_export_zip_con_avviso_privacy(server_app):
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(viewport={'width': 390, 'height': 780}, locale='it-IT')
        ctx.add_init_script(INIT_SCRIPT)
        
        page = ctx.new_page()
        page.goto(server_app)
        page.evaluate('([k,v])=>localStorage.setItem(k,v)', [CHIAVE, json.dumps(STATO_SENSIVILE)])
        
        page.evaluate("""async () => {
            if (window.FotoStore) {
                const jpegDummy = new Uint8Array([0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01, 0xFF, 0xD9]);
                const blob = new Blob([jpegDummy], { type: 'image/jpeg' });
                await window.FotoStore.aggiungi('c1', '2026-10-10', blob);
            }
        }""")
        
        page.reload()
        page.wait_for_selector('text=Esporta per LeenO')
        
        dialogs = []
        page.on('dialog', lambda d: (dialogs.append(d.message), d.accept()))
        
        page.click('text=Esporta per LeenO')
        page.wait_for_selector('#pannello-export')
        page.click('text=Invia o condividi')
        page.wait_for_timeout(300)
        
        assert len(dialogs) == 2
        assert 'Evento infortunistico' in dialogs[0]
        
        shared = page.evaluate('window.__sharedData')
        assert len(shared) == 1
        assert shared[0]['files'][0]['name'].endswith('.zip')
        
        browser.close()


def test_export_zip_tutti_i_cantieri(server_app):
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(viewport={'width': 390, 'height': 780}, locale='it-IT')
        ctx.add_init_script(INIT_SCRIPT)
        
        page = ctx.new_page()
        page.goto(server_app)
        page.evaluate('([k,v])=>localStorage.setItem(k,v)', [CHIAVE, json.dumps(STATO_MOLTEPLICI)])
        
        page.reload()
        page.wait_for_selector('text=Esporta per LeenO')
        page.click('text=Esporta per LeenO')
        
        page.wait_for_selector('text=Tutti i cantieri (2)')
        page.click('text=Tutti i cantieri (2)')
        
        page.wait_for_selector('#pannello-export')
        pannello_testo = page.text_content('#pannello-export')
        assert 'File pronto: agenda-tutti-i-cantieri-' in pannello_testo
        assert '.zip' in pannello_testo
        
        page.click('text=Invia o condividi')
        page.wait_for_timeout(300)
        
        shared = page.evaluate('window.__sharedData')
        assert len(shared) == 1
        assert 'agenda-tutti-i-cantieri-' in shared[0]['files'][0]['name']
        assert shared[0]['files'][0]['name'].endswith('.zip')
        
        browser.close()


def test_export_fallback_quando_can_share_falso(server_app):
    """Verifica che in Firefox (o quando canShare è falso) il pannello si apra mostrando 'Salva sul telefono'."""
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(viewport={'width': 390, 'height': 780}, locale='it-IT')
        ctx.add_init_script("""
            navigator.canShare = function(data) { return false; };
            navigator.share = undefined;
        """)
        
        page = ctx.new_page()
        page.goto(server_app)
        page.evaluate('([k,v])=>localStorage.setItem(k,v)', [CHIAVE, json.dumps(STATO_SINGOLO)])
        page.reload()
        page.wait_for_selector('text=Esporta per LeenO')
        page.click('text=Esporta per LeenO')
        
        page.wait_for_selector('#pannello-export')
        pannello_testo = page.text_content('#pannello-export')
        assert 'File pronto:' in pannello_testo
        assert 'non supporta la condivisione diretta' in pannello_testo
        assert page.locator('text=Invia o condividi').count() == 0
        assert page.locator('text=Salva sul telefono').count() == 1
        
        browser.close()


def test_export_fallback_quando_share_fallisce_brave(server_app):
    """Verifica che in Brave (quando share() viene rifiutato) l'errore venga intercettato ed eseguito il salvataggio."""
    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(viewport={'width': 390, 'height': 780}, locale='it-IT')
        ctx.add_init_script("""
            navigator.canShare = function(data) { return true; };
            navigator.share = function(data) {
                return Promise.reject(new DOMException("Permission denied", "NotAllowedError"));
            };
        """)
        
        page = ctx.new_page()
        page.goto(server_app)
        page.evaluate('([k,v])=>localStorage.setItem(k,v)', [CHIAVE, json.dumps(STATO_SINGOLO)])
        page.reload()
        page.wait_for_selector('text=Esporta per LeenO')
        page.click('text=Esporta per LeenO')
        
        page.wait_for_selector('#pannello-export')
        page.click('text=Invia o condividi')
        page.wait_for_timeout(300)
        
        msg_testo = page.text_content('#msg')
        assert 'Condivisione non riuscita' in msg_testo or 'salvato' in msg_testo
        
        browser.close()
