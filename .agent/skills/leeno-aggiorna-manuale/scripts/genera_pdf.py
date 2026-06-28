#!/usr/bin/env python3
"""
Genera il PDF del manuale di LeenO a partire dal file FODT sorgente.

Converte:
  documentazione/MANUALE_LeenO.fodt  →  src/Ultimus.oxt/MANUALE_LeenO.pdf

Richiede LibreOffice installato e `soffice` raggiungibile nel PATH.
"""

import os
import shutil
import subprocess
import sys
import tempfile

# ── Percorsi ─────────────────────────────────────────────────────────────
REPO_ROOT = os.path.normpath(
    os.path.join(os.path.dirname(__file__), '..', '..', '..', '..')
)
FODT_SRC = os.path.join(REPO_ROOT, 'documentazione', 'MANUALE_LeenO.fodt')
PDF_DEST = os.path.join(REPO_ROOT, 'src', 'Ultimus.oxt', 'MANUALE_LeenO.pdf')


def find_soffice():
    """Restituisce il percorso di soffice, cercando nel PATH e nei luoghi comuni."""
    # 1. Prova il PATH di sistema
    soffice = shutil.which('soffice')
    if soffice:
        return soffice
    # 2. Percorsi comuni Windows
    for candidate in (
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    ):
        if os.path.isfile(candidate):
            return candidate
    return None


def main():
    if not os.path.isfile(FODT_SRC):
        print(f'ERRORE: File sorgente non trovato: {FODT_SRC}', file=sys.stderr)
        sys.exit(1)

    soffice = find_soffice()
    if not soffice:
        print('ERRORE: LibreOffice (soffice) non trovato nel PATH.', file=sys.stderr)
        sys.exit(1)

    # Usa una directory temporanea per l'output, poi sposta il file.
    # Questo evita problemi con --outdir che non permette di rinominare.
    with tempfile.TemporaryDirectory() as tmpdir:
        cmd = [
            soffice,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', tmpdir,
            FODT_SRC,
        ]
        print(f'Esecuzione: {" ".join(cmd)}')
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

        if result.returncode != 0:
            print(f'ERRORE soffice (exit {result.returncode}):', file=sys.stderr)
            print(result.stderr, file=sys.stderr)
            sys.exit(1)

        # Il file generato ha lo stesso basename del sorgente, estensione .pdf
        generated = os.path.join(tmpdir, 'MANUALE_LeenO.pdf')
        if not os.path.isfile(generated):
            print(f'ERRORE: PDF non generato. Output soffice:\n{result.stdout}',
                  file=sys.stderr)
            sys.exit(1)

        # Assicura che la directory di destinazione esista
        os.makedirs(os.path.dirname(PDF_DEST), exist_ok=True)
        shutil.move(generated, PDF_DEST)

    size_kb = os.path.getsize(PDF_DEST) / 1024
    print(f'PDF generato con successo: {PDF_DEST}  ({size_kb:.0f} KB)')


if __name__ == '__main__':
    main()
