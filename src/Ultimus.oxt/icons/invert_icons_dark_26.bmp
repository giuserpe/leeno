#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
invert_icons_dark.py
=====================

Genera le varianti "tema scuro" delle icone SVG di LeenO.

Le icone LeenO sono costruite con shape piene (no stroke): i bordi sono
ottenuti sovrapponendo forme concentriche, in genere un anello quasi nero
(es. "#1A2010") seguito da un riempimento quasi bianco (es. "#F0F4E0"), con
sopra i colori "accento"/semantici (blu, rosso, verde, giallo, ecc.).

Un invert "globale" dei colori romperebbe gli accenti (es. un blu diventa
un arancione). Questo script quindi inverte SOLO i colori neutri o agli
estremi di luminosita' (near-black, near-white, grigi puri), lasciando
intatti i colori saturi usati come accento/semantica.

Criterio (in spazio HSL):
    - se saturazione <= NEUTRAL_SAT_THRESHOLD  -> considerato neutro (grigio)
    - oppure se lightness <= DARK_L_THRESHOLD  -> considerato "nero/bordo"
    - oppure se lightness >= LIGHT_L_THRESHOLD -> considerato "bianco/sfondo"
  In questi 3 casi si inverte la lightness (L' = 1 - L) mantenendo hue e
  saturazione originali. Altrimenti il colore resta invariato.

USO:
    python invert_icons_dark.py "W:\\tmp\\svg"
    python invert_icons_dark.py "W:\\tmp\\svg" --outdir "W:\\tmp\\svg\\dark"
    python invert_icons_dark.py "W:\\tmp\\svg" --suffix _dark --recursive
    python invert_icons_dark.py "W:\\tmp\\svg" --dry-run

Nessuna dipendenza esterna: solo libreria standard.
"""

import argparse
import colorsys
import os
import re
import sys

# ---------------------------------------------------------------------------
# Parametri di classificazione colore (regolabili da riga di comando)
# ---------------------------------------------------------------------------
NEUTRAL_SAT_THRESHOLD = 0.25   # saturazione <= a questo valore => "grigio"
DARK_L_THRESHOLD = 0.15        # lightness <= a questo valore   => "nero/bordo"
LIGHT_L_THRESHOLD = 0.85       # lightness >= a questo valore   => "bianco/sfondo"

HEX3_RE = re.compile(r'^#([0-9a-fA-F])([0-9a-fA-F])([0-9a-fA-F])$')
HEX6_RE = re.compile(r'^#([0-9a-fA-F]{6})$')

# Cattura fill="#xxxxxx", stroke="#xxxxxx", stop-color="#xxxxxx"
# e le stesse proprieta' dentro un attributo style="...fill:#xxxxxx;..."
ATTR_COLOR_RE = re.compile(
    r'((?:fill|stroke|stop-color)\s*=\s*")(#[0-9a-fA-F]{3,6})(")'
)
STYLE_COLOR_RE = re.compile(
    r'((?:fill|stroke|stop-color)\s*:\s*)(#[0-9a-fA-F]{3,6})(\s*;?)'
)


def hex_to_rgb(hex_color: str):
    """'#RGB' o '#RRGGBB' -> (r, g, b) in 0..255"""
    m3 = HEX3_RE.match(hex_color)
    if m3:
        r, g, b = (int(c * 2, 16) for c in m3.groups())
        return r, g, b
    m6 = HEX6_RE.match(hex_color)
    if m6:
        h = m6.group(1)
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return None


def rgb_to_hex(r: int, g: int, b: int) -> str:
    return "#{:02X}{:02X}{:02X}".format(
        max(0, min(255, round(r))),
        max(0, min(255, round(g))),
        max(0, min(255, round(b))),
    )


def invert_if_neutral_or_extreme(
    hex_color: str,
    neutral_sat: float = NEUTRAL_SAT_THRESHOLD,
    dark_l: float = DARK_L_THRESHOLD,
    light_l: float = LIGHT_L_THRESHOLD,
) -> str:
    """Applica il criterio di inversione selettiva a un colore esadecimale."""
    rgb = hex_to_rgb(hex_color)
    if rgb is None:
        return hex_color  # formato non riconosciuto, lascia invariato

    r, g, b = (c / 255.0 for c in rgb)
    h, l, s = colorsys.rgb_to_hls(r, g, b)  # nota: colorsys usa HLS, non HSL

    is_neutral = s <= neutral_sat
    is_dark_extreme = l <= dark_l
    is_light_extreme = l >= light_l

    if not (is_neutral or is_dark_extreme or is_light_extreme):
        return hex_color  # colore accento/semantico: non toccare

    new_l = 1.0 - l
    nr, ng, nb = colorsys.hls_to_rgb(h, new_l, s)
    return rgb_to_hex(nr * 255, ng * 255, nb * 255)


def process_svg_text(svg_text: str, neutral_sat: float, dark_l: float, light_l: float):
    """Ritorna (nuovo_testo, lista_sostituzioni) per un SVG in memoria."""
    changes = []

    def make_repl():
        def repl(m):
            prefix, color, suffix = m.groups()
            new_color = invert_if_neutral_or_extreme(color, neutral_sat, dark_l, light_l)
            if new_color != color:
                changes.append((color, new_color))
            return f"{prefix}{new_color}{suffix}"
        return repl

    new_text = ATTR_COLOR_RE.sub(make_repl(), svg_text)
    new_text = STYLE_COLOR_RE.sub(make_repl(), new_text)
    return new_text, changes


def find_svg_files(input_dir: str, recursive: bool):
    if recursive:
        for root, _dirs, files in os.walk(input_dir):
            for name in files:
                if name.lower().endswith(".svg"):
                    yield os.path.join(root, name)
    else:
        for name in os.listdir(input_dir):
            if name.lower().endswith(".svg"):
                yield os.path.join(input_dir, name)


def main():
    parser = argparse.ArgumentParser(
        description="Inverte selettivamente i colori neutri/estremi delle icone SVG LeenO per il tema scuro."
    )
    parser.add_argument("input_dir", help="Cartella contenente le icone SVG (es. W:\\tmp\\svg)")
    parser.add_argument(
        "--outdir",
        default=None,
        help="Cartella di destinazione (default: <input_dir>\\dark)",
    )
    parser.add_argument(
        "--suffix",
        default="",
        help="Suffisso da aggiungere al nome file, prima di '.svg' (es. _dark). "
             "Utile se si vuole scrivere nella stessa cartella di input.",
    )
    parser.add_argument(
        "--recursive", action="store_true", help="Cerca gli SVG anche nelle sottocartelle"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Mostra solo cosa verrebbe cambiato, senza scrivere alcun file",
    )
    parser.add_argument(
        "--neutral-sat",
        type=float,
        default=NEUTRAL_SAT_THRESHOLD,
        help=f"Soglia di saturazione sotto la quale un colore e' considerato grigio (default {NEUTRAL_SAT_THRESHOLD})",
    )
    parser.add_argument(
        "--dark-l",
        type=float,
        default=DARK_L_THRESHOLD,
        help=f"Soglia di lightness sotto la quale un colore e' considerato 'nero/bordo' (default {DARK_L_THRESHOLD})",
    )
    parser.add_argument(
        "--light-l",
        type=float,
        default=LIGHT_L_THRESHOLD,
        help=f"Soglia di lightness sopra la quale un colore e' considerato 'bianco/sfondo' (default {LIGHT_L_THRESHOLD})",
    )
    args = parser.parse_args()

    neutral_sat = args.neutral_sat
    dark_l = args.dark_l
    light_l = args.light_l

    input_dir = os.path.abspath(args.input_dir)
    if not os.path.isdir(input_dir):
        print(f"ERRORE: la cartella non esiste: {input_dir}")
        sys.exit(1)

    outdir = os.path.abspath(args.outdir) if args.outdir else os.path.join(input_dir, "dark")
    if not args.dry_run:
        os.makedirs(outdir, exist_ok=True)

    svg_files = sorted(find_svg_files(input_dir, args.recursive))
    if not svg_files:
        print(f"Nessun file .svg trovato in: {input_dir}")
        sys.exit(0)

    total_files_changed = 0
    total_colors_changed = 0

    for path in svg_files:
        with open(path, "r", encoding="utf-8") as f:
            svg_text = f.read()

        new_text, changes = process_svg_text(svg_text, neutral_sat, dark_l, light_l)
        rel_name = os.path.relpath(path, input_dir)

        if changes:
            total_files_changed += 1
            total_colors_changed += len(changes)
            unique_changes = sorted(set(changes))
            change_desc = ", ".join(f"{old}->{new}" for old, new in unique_changes)
            print(f"[MODIFICATO] {rel_name}: {change_desc}")
        else:
            print(f"[invariato]  {rel_name}")

        if not args.dry_run:
            name, ext = os.path.splitext(os.path.basename(path))
            out_name = f"{name}{args.suffix}{ext}"
            out_path = os.path.join(outdir, out_name)
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(new_text)

    print("-" * 60)
    if args.dry_run:
        print(f"DRY RUN: {total_files_changed}/{len(svg_files)} file avrebbero subito modifiche "
              f"({total_colors_changed} sostituzioni colore totali). Nessun file scritto.")
    else:
        print(f"Fatto: {total_files_changed}/{len(svg_files)} file modificati "
              f"({total_colors_changed} sostituzioni colore totali).")
        print(f"Output scritto in: {outdir}")


if __name__ == "__main__":
    main()
