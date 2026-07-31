#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rigenera_bmp_icone.py
======================

Rigenera i file "_16.bmp", "_16h.bmp", "_26.bmp", "_26h.bmp" richiesti da
LibreOffice a partire dalle icone SVG sorgente in "icons/svg/".

Le icone LeenO non sono bitmap vere: LibreOffice accetta contenuto SVG
anche con estensione .bmp, quindi rigenerare un set significa semplicemente
copiare l'SVG sorgente 4 volte con i suffissi richiesti dentro "icons/"
(la cartella padre di "icons/svg/"). E' l'esatta logica gia' presente in
icons/svg/leeno_icons_2.sh, riscritta in Python per poter girare anche su
Windows senza bash/WSL.

USO (da lanciare dentro icons/svg/, oppure passando --svg-dir):
    python rigenera_bmp_icone.py                     # rigenera TUTTE le icone
    python rigenera_bmp_icone.py somme_sicurezza      # solo una o piu' icone
    python rigenera_bmp_icone.py somme_sicurezza parz
    python rigenera_bmp_icone.py --svg-dir "W:\\_dwg\\ULTIMUSFREE\\_SRC\\leeno\\src\\Ultimus.oxt\\icons\\svg"
    python rigenera_bmp_icone.py --dry-run

Nessuna dipendenza esterna: solo libreria standard.
"""

import argparse
import shutil
import sys
from pathlib import Path

SUFFISSI = ("_16.bmp", "_16h.bmp", "_26.bmp", "_26h.bmp")


def rigenera_icona(svg_path: Path, dest_dir: Path, dry_run: bool = False) -> list[Path]:
    """Copia svg_path in dest_dir con i 4 suffissi bmp richiesti da LO.

    Ritorna la lista dei file (ri)scritti.
    """
    scritti = []
    for suffisso in SUFFISSI:
        target = dest_dir / f"{svg_path.stem}{suffisso}"
        if dry_run:
            print(f"[dry-run] {svg_path} -> {target}")
        else:
            shutil.copyfile(svg_path, target)
        scritti.append(target)
    return scritti


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "icone",
        nargs="*",
        help="Nomi icona senza estensione (es. somme_sicurezza). "
             "Se omesso, rigenera tutte le SVG trovate in --svg-dir.",
    )
    parser.add_argument(
        "--svg-dir",
        default=".",
        help="Cartella con le SVG sorgente (default: cartella corrente, "
             "pensata per essere lanciata dentro icons/svg/).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Mostra cosa verrebbe copiato senza scrivere nulla.",
    )
    args = parser.parse_args()

    svg_dir = Path(args.svg_dir).resolve()
    dest_dir = svg_dir.parent  # icons/ (la cartella padre di icons/svg/)

    if not svg_dir.is_dir():
        print(f"Errore: cartella SVG non trovata: {svg_dir}", file=sys.stderr)
        sys.exit(1)

    if args.icone:
        svg_paths = [svg_dir / f"{nome}.svg" for nome in args.icone]
        mancanti = [p for p in svg_paths if not p.is_file()]
        if mancanti:
            for p in mancanti:
                print(f"Errore: SVG non trovata: {p}", file=sys.stderr)
            sys.exit(1)
    else:
        svg_paths = sorted(svg_dir.glob("*.svg"))
        if not svg_paths:
            print(f"Nessuna SVG trovata in {svg_dir}", file=sys.stderr)
            sys.exit(1)

    totale = 0
    for svg_path in svg_paths:
        scritti = rigenera_icona(svg_path, dest_dir, dry_run=args.dry_run)
        totale += len(scritti)
        print(f"{svg_path.name}: {len(scritti)} file rigenerati in {dest_dir}")

    print(f"\nCompletato: {len(svg_paths)} icone, {totale} file bmp "
          f"{'(dry-run, nessuna scrittura)' if args.dry_run else 'scritti'}.")


if __name__ == "__main__":
    main()
