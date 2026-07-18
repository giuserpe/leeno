#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Generatore automatico di icone SVG per LeenO secondo il Design System v2.0 e Addendum v2.1.
Produce 51 icone × 2 temi (chiaro/scuro) = 102 file SVG in:
- src/Ultimus.oxt/icons/svg/ (Tema Chiaro)
- src/Ultimus.oxt/icons/scuro/ (Tema Scuro)
Tutte le icone sono "fill-only", senza 'stroke', con allineamento snap-to-grid,
massimo 2 colori semantici oltre al tratto principale, e spessori massimi delle linee di 1.5px (su griglia 24x24).
"""

import os
import re

# Tavolozza dei colori (Design System v2.0 / v2.1)
PALETTE = {
    "verde_primario": "#5D7400",
    "lime_accento": "#AAD400",
    "arancione_azione": "#FF4D2E",
    "blu_info": "#3B82F6",
    "giallo_avviso": "#F4B400",
    "sfondo": "#F0F4E0",
    "grigio": "#808080",
    "chiaro_bg": "#FAFAFA",
}

# Dimensioni griglia
GRID_SIZE = 24

def get_colors(is_dark):
    """
    Ritorna i colori per il tema specificato.
    """
    if is_dark:
        return {
            "main": "#FAFAFA",          # Bianco sporco/chiaro
            "accent": PALETTE["lime_accento"],
            "success": PALETTE["lime_accento"],
            "warning": PALETTE["giallo_avviso"],
            "danger": PALETTE["arancione_azione"],
            "info": PALETTE["blu_info"],
            "bg": "#1A2010",            # Sfondo scuro per contrasto o alone invertito
            "gray": "#A0A0A0",
        }
    else:
        return {
            "main": "#1A2010",          # Scuro
            "accent": PALETTE["verde_primario"],
            "success": PALETTE["verde_primario"],
            "warning": PALETTE["giallo_avviso"],
            "danger": PALETTE["arancione_azione"],
            "info": PALETTE["blu_info"],
            "bg": PALETTE["sfondo"],
            "gray": PALETTE["grigio"],
        }

# --- Primitive Vettoriali Riutilizzabili (Esclusivamente FILL, NO STROKE) ---

def make_svg_wrapper(paths_html):
    """Ritorna l'elemento radice SVG con il viewBox standard 24x24."""
    return f"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="100%" height="100%">{paths_html}</svg>"""

def rect_fill(x, y, w, h, color, rx=0):
    """Rettangolo standard riempito (usato per linee o blocchi piatti)."""
    rx_attr = f' rx="{rx}" ry="{rx}"' if rx > 0 else ""
    return f'<rect x="{x}" y="{y}" width="{w}" height="{h}" fill="{color}"{rx_attr} />'

def line_fill(x, y, length, is_vertical, thickness, color, rx=0.4):
    """Una linea disegnata come un rettangolo sottile secondo la regola 12.3 (spessore <= 1.5px)."""
    if is_vertical:
        return rect_fill(x - thickness/2, y, thickness, length, color, rx)
    else:
        return rect_fill(x, y - thickness/2, length, thickness, color, rx)

def doc_primitive(color, scale=1.0, offset_x=0.0, offset_y=0.0):
    """
    Forma Documento: rettangolo verticale con l'angolo in alto a destra piegato.
    Su griglia 24x24: larghezza 14, altezza 18. Centrato: x=5, y=3.
    """
    t = 1.5
    outer = f"M {5+offset_x} {3+offset_y} L {15+offset_x} {3+offset_y} L {19+offset_x} {7+offset_y} L {19+offset_x} {21+offset_y} L {5+offset_x} {21+offset_y} Z"
    inner = f"M {5+t+offset_x} {3+t+offset_y} L {5+t+offset_x} {21-t+offset_y} L {19-t+offset_x} {21-t+offset_y} L {19-t+offset_x} {7+t/2+offset_y} L {15-t/2+offset_x} {3+t+offset_y} Z"
    fold_line = f"M {15+offset_x} {3+offset_y} L {15+offset_x} {7+offset_y} L {19+offset_x} {7+offset_y} L {19-t+offset_x} {7-t+offset_y} L {15-t+offset_x} {7-t+offset_y} L {15-t+offset_x} {3+t+offset_y} Z"
    return f'<path d="{outer} {inner}" fill="{color}" fill-rule="evenodd" />'

def folder_primitive(color, offset_x=0.0, offset_y=0.0):
    """
    Forma Cartella: classica scheda di cartella.
    Su griglia 24x24: larghezza 20, altezza 16. Centrato: x=2, y=4.
    """
    t = 1.5
    outer = f"M {2+offset_x} {4+offset_y} L {10+offset_x} {4+offset_y} L {12+offset_x} {7+offset_y} L {22+offset_x} {7+offset_y} L {22+offset_x} {20+offset_y} L {2+offset_x} {20+offset_y} Z"
    inner = f"M {2+t+offset_x} {7+t+offset_y} L {22-t+offset_x} {7+t+offset_y} L {22-t+offset_x} {20-t+offset_y} L {2+t+offset_x} {20-t+offset_y} Z"
    tab_inner = f"M {2+t+offset_x} {4+t+offset_y} L {10-t+offset_x} {4+t+offset_y} L {11.2+offset_x} {7+offset_y} L {2+t+offset_x} {7+offset_y} Z"
    return f'<path d="{outer} {inner} {tab_inner}" fill="{color}" fill-rule="evenodd" />'

def wbs_list_primitive(color, offset_x=0.0, offset_y=0.0):
    """Forma WBS / Elenco: linee parallele orizzontali e verticali di guida."""
    t = 1.5
    l1 = rect_fill(3+offset_x, 6+offset_y, 18, t, color, rx=0.4)
    l2 = rect_fill(6+offset_x, 11+offset_y, 15, t, color, rx=0.4)
    l3 = rect_fill(9+offset_x, 16+offset_y, 12, t, color, rx=0.4)
    l_vert = rect_fill(3+offset_x, 6+offset_y, t, 11.5, color, rx=0.4)
    return l1 + l2 + l3 + l_vert

# --- Decoratori / Badge con Knockout (Alone) ---

def apply_knockout(badge_html, bg_color, offset_x=0.0, offset_y=0.0):
    """
    Raddoppia il badge disegnando prima una versione leggermente più grande in colore di sfondo (knockout/alone),
    quindi sopra il badge reale per evitare sovrapposizioni sgradevoli (Addendum v2.1 Section 12).
    """
    # Usiamo una regex sicura per sostituire qualsiasi fill="..." con fill="{bg_color}"
    bg_paths = []
    for dx in [-1.2, 0, 1.2]:
        for dy in [-1.2, 0, 1.2]:
            if dx != 0 or dy != 0:
                bg_paths.append(re.sub(r'fill="[^"]+"', f'fill="{bg_color}"', badge_html))

    bg_layer = "".join(bg_paths)
    return bg_layer + badge_html

def badge_plus(color, bg_color, x=16, y=16):
    """Badge Più (+) di 6x6px con alone."""
    t = 1.5
    html = rect_fill(x-3, y-t/2, 6, t, color, rx=0.4) + rect_fill(x-t/2, y-3, t, 6, color, rx=0.4)
    return apply_knockout(html, bg_color)

def badge_minus(color, bg_color, x=16, y=16):
    """Badge Meno (-) di 6x6px con alone."""
    t = 1.5
    html = rect_fill(x-3, y-t/2, 6, t, color, rx=0.4)
    return apply_knockout(html, bg_color)

def badge_check(color, bg_color, x=16, y=16):
    """Badge Spunta (V) con alone."""
    t = 1.5
    html = f'<path d="M {x-3} {y} L {x-1} {y+2} L {x+4} {y-3} L {x+4-t} {y-3-t/2} L {x-1} {y+2-t} L {x-3+t/2} {y-t/2} Z" fill="{color}" />'
    return apply_knockout(html, bg_color)

def ring_primitive(outer_color, inner_color, center_color, r_outer=5.0, x=16, y=16):
    """
    Tecnica dell'anello (ring) per badge e cerchi (sezione 12.4):
    outer_color -> inner_color -> center_color
    """
    c1 = f'<circle cx="{x}" cy="{y}" r="{r_outer}" fill="{outer_color}" />'
    c2 = f'<circle cx="{x}" cy="{y}" r="{r_outer - 1.2}" fill="{inner_color}" />'
    c3 = f'<circle cx="{x}" cy="{y}" r="{r_outer - 2.5}" fill="{center_color}" />'
    return c1 + c2 + c3

# --- Generatore di icone specifico ---

def build_icon(name, colors):
    """
    Genera il codice SVG dell'icona in base al nome e ai colori del tema.
    """
    m = colors["main"]
    acc = colors["accent"]
    warn = colors["warning"]
    dang = colors["danger"]
    inf = colors["info"]
    bg = colors["bg"]
    gray = colors["gray"]

    html = ""

    # --- CATEGORIA 1: Principale e Navigazione ---
    if name == "leeno":
        html += ring_primitive(m, bg, acc, r_outer=9, x=12, y=12)
        html += rect_fill(7, 6, 2.5, 12, m, rx=0.5)
        html += rect_fill(7, 15.5, 10, 2.5, m, rx=0.5)
    elif name == "manuale":
        html += doc_primitive(m)
        html += apply_knockout(rect_fill(11.25, 10, 1.5, 6, inf, rx=0.4) + rect_fill(11.25, 7.5, 1.5, 1.5, inf, rx=0.75), bg)
    elif name == "teleg":
        # Disegno fill-only di Telegram
        html = f'<path d="M 3 12 L 21 3 L 13 21 L 11 13 Z" fill="{inf}" opacity="0.8"/>'
        html += f'<path d="M 21 3 L 13 21 L 11 13 Z" fill="{inf}"/>'
        # Disegniamo la linea di cucitura con un poligono fill-only sottile
        html += f'<path d="M 11 13 L 21 3 L 11 11.5 Z" fill="{bg}"/>'

    # --- CATEGORIA 2: Struttura di Scomposizione (WBS) ---
    elif name == "supcat":
        html += folder_primitive(m)
        html += apply_knockout(rect_fill(11.25, 10, 1.5, 6, acc, rx=0.4), bg)
    elif name == "cat":
        html += folder_primitive(m)
        html += apply_knockout(rect_fill(9.5, 10, 1.5, 6, acc, rx=0.4) + rect_fill(13, 10, 1.5, 6, acc, rx=0.4), bg)
    elif name == "subcat":
        html += folder_primitive(m)
        html += apply_knockout(rect_fill(8, 10, 1.5, 6, acc, rx=0.4) + rect_fill(11.25, 10, 1.5, 6, acc, rx=0.4) + rect_fill(14.5, 10, 1.5, 6, acc, rx=0.4), bg)
    elif name == "struttura_on":
        html += wbs_list_primitive(m)
        html += badge_plus(acc, bg, x=18, y=16)
    elif name == "struttura_off":
        html += wbs_list_primitive(m)
        html += badge_minus(dang, bg, x=18, y=16)
    elif name == "rinumCap":
        html += wbs_list_primitive(m)
        t = 1.0
        hash_tag = (
            rect_fill(15, 6, t, 12, warn) + rect_fill(18, 6, t, 12, warn) +
            rect_fill(13, 9, 7, t, warn) + rect_fill(13, 13, 7, t, warn)
        )
        html += apply_knockout(hash_tag, bg)

    # --- CATEGORIA 3: Voci di Lavoro (Voci) ---
    elif name == "nuova_voce":
        html += doc_primitive(m)
        html += badge_plus(acc, bg, x=16, y=16)
    elif name == "voce_breve":
        html += doc_primitive(m)
        cut_line = rect_fill(3, 12, 18, 1.5, dang, rx=0.4)
        html += apply_knockout(cut_line, bg)
    elif name == "vedivoce":
        html += doc_primitive(m)
        eye = (
            f'<path d="M 8 13 C 10 10, 14 10, 16 13 C 14 16, 10 16, 8 13 Z" fill="{inf}" />' +
            f'<circle cx="12" cy="13" r="2" fill="{bg}" />' +
            f'<circle cx="12" cy="13" r="1" fill="{inf}" />'
        )
        html += apply_knockout(eye, bg)
    elif name == "pesca":
        html += f'<path d="M 12 4 L 13.5 4 L 13.5 14 C 13.5 17, 7.5 17, 7.5 14 L 9 14 C 9 15.5 12 15.5 12 14 L 12 4 Z" fill="{acc}" />'
        html += f'<path d="M 7.5 14 L 6 12.5 L 7 11.5 Z" fill="{acc}" />'
    elif name == "invia_voce_ep":
        html += doc_primitive(m)
        arrow = rect_fill(12, 12, 7, 1.5, inf, rx=0.4) + f'<path d="M 16 9.5 L 19.5 12.75 L 16 16 Z" fill="{inf}" />'
        html += apply_knockout(arrow, bg)
    elif name == "aggiungi_misura":
        html += doc_primitive(m)
        measure_line = rect_fill(8, 12, 8, 1.5, acc, rx=0.4)
        html += apply_knockout(measure_line, bg)
        html += badge_plus(acc, bg, x=17, y=17)
    elif name == "sposta_voce":
        arrow_up = rect_fill(9, 6, 1.5, 12, acc, rx=0.4) + f'<path d="M 7 9 L 9.75 5.5 L 12.5 9 Z" fill="{acc}" />'
        arrow_down = rect_fill(15, 6, 1.5, 12, acc, rx=0.4) + f'<path d="M 13 15 L 15.75 18.5 L 18.5 15 Z" fill="{acc}" />'
        html += arrow_up + arrow_down

    # --- CATEGORIA 4: Elenchi Prezzi e Analisi dei Costi ---
    elif name == "analisi_a_prezzo":
        html += doc_primitive(gray, offset_x=-2, offset_y=-2)
        html += apply_knockout(doc_primitive(m, offset_x=2, offset_y=2), bg)
        arrow = rect_fill(9, 13, 6, 1.5, acc, rx=0.4) + f'<path d="M 13 10.5 L 16.5 13.75 L 13 17 Z" fill="{acc}" />'
        html += apply_knockout(arrow, bg)
    elif name == "utili_maggiorazioni":
        html += ring_primitive(acc, bg, acc, r_outer=3, x=8, y=8)
        html += ring_primitive(acc, bg, acc, r_outer=3, x=16, y=16)
        html += f'<path d="M 16 6 L 17.5 7 L 8 18 L 6.5 17 Z" fill="{acc}" />'
    elif name == "elimina_doppioni":
        html += doc_primitive(m, offset_x=-1, offset_y=-1)
        html += apply_knockout(doc_primitive(m, offset_x=2, offset_y=2), bg)
        cross_fill = f'<path d="M 13 14 L 14 13 L 18 17 L 17 18 Z M 17 13 L 18 14 L 14 18 L 13 17 Z" fill="{dang}" />'
        html += apply_knockout(cross_fill, bg)
    elif name == "riordina":
        html += rect_fill(5, 4, 1.5, 16, acc, rx=0.4) + f'<path d="M 3 16 L 5.75 19.5 L 8.5 16 Z" fill="{acc}" />'
        html += f'<path d="M 13 4 L 16 4 L 18 10 L 16.5 10 L 16 8 L 13 8 L 13 10 L 11.5 10 Z M 13.5 5.5 L 15.5 5.5 L 15.5 7 L 13.5 7 Z" fill="{m}" fill-rule="evenodd" />'
        html += f'<path d="M 12 14 L 17 14 L 17 15.5 L 13.5 18.5 L 17 18.5 L 17 20 L 12 20 L 12 18.5 L 15.5 15.5 L 12 15.5 Z" fill="{m}" fill-rule="evenodd" />'

    # --- CATEGORIA 5: Quantità e Contabilità ---
    elif name == "parz":
        html = f'<path d="M 8 5 L 16 5 L 12 12 L 16 19 L 8 19 L 8 17.5 L 13.5 17.5 L 10.5 12 L 13.5 6.5 L 8 6.5 Z" fill="{acc}" />'
        html += f'<path d="M 5 4 C 3 8, 3 16, 5 20 L 6.2 19.5 C 4.5 16, 4.5 8, 6.2 4.5 Z" fill="{m}" />'
        html += f'<path d="M 18 4 C 20 8, 20 16, 18 20 L 16.8 19.5 C 18.5 16, 18.5 8, 16.8 4.5 Z" fill="{m}" />'
    elif name == "inverti_segno":
        html += rect_fill(5, 11, 4, 1.5, acc, rx=0.4) + rect_fill(6.25, 9, 1.5, 5.5, acc, rx=0.4)
        html += rect_fill(14, 11, 4, 1.5, acc, rx=0.4)
        html += rect_fill(8, 16, 7, 1.5, warn, rx=0.4) + f'<path d="M 13 14 L 16.5 16.75 L 13 19.5 Z" fill="{warn}" />'
    elif name == "azzera":
        t = 1.5
        outer_ellipse = f"M 12 4 C 7.5 4, 6 8, 6 12 C 6 16, 7.5 20, 12 20 C 16.5 20, 18 16, 18 12 C 18 8, 16.5 4, 12 4 Z"
        inner_ellipse = f"M 12 {4+t} C {7.5+t} {4+t}, {6+t} 8, {6+t} 12 C {6+t} 16, {7.5+t} {20-t}, 12 {20-t} C {16.5-t} {20-t}, {18-t} 16, {18-t} 12 C {18-t} 8, {16.5-t} {4+t}, 12 {4+t} Z"
        html += f'<path d="{outer_ellipse} {inner_ellipse}" fill="{dang}" fill-rule="evenodd" />'
    elif name in ("partita_provvisoria_piu", "part_agg"):
        html += folder_primitive(m, offset_y=-2)
        html += apply_knockout(folder_primitive(m, offset_y=2), bg)
        html += badge_plus(acc, bg, x=17, y=17)
    elif name in ("partita_provvisoria_meno", "part_det"):
        html += folder_primitive(m, offset_y=-2)
        html += apply_knockout(folder_primitive(m, offset_y=2), bg)
        html += badge_minus(dang, bg, x=17, y=17)
    elif name == "strutt_voci_zero":
        html += wbs_list_primitive(m)
        t = 1.0
        zero = (
            f'<path d="M 16 11 C 14 11, 13 12.5, 13 15 C 13 17.5, 14 19, 16 19 C 18 19, 19 17.5, 19 15 C 19 12.5, 18 11, 16 11 Z M 16 {11+t} C 14.8 {11+t}, 14 {12.5}, 14 15 C 14 17.5, 14.8 {19-t}, 16 {19-t} C 17.2 {19-t}, 18 17.5, 18 15 C 18 12.5, 17.2 {11+t}, 16 {11+t} Z" fill="{dang}" fill-rule="evenodd" />' +
            f'<path d="M 13.5 18.5 L 18.5 11.5 L 19.5 12.2 L 14.5 19.2 Z" fill="{dang}" />'
        )
        html += apply_knockout(zero, bg)
    elif name == "elimina_azzerate":
        html += doc_primitive(m)
        zero = f'<path d="M 10 10 C 8.5 10, 8 11, 8 13 C 8 15, 8.5 16, 10 16 C 11.5 16, 12 15, 12 13 C 12 11, 11.5 10, 10 10 Z" fill="{dang}" />'
        html += apply_knockout(zero, bg)
        html += badge_minus(dang, bg, x=16, y=16)
    elif name == "elimina_vuote":
        html += wbs_list_primitive(m)
        cross_fill = f'<path d="M 14 14 L 15 13 L 19 17 L 18 18 Z M 18 13 L 19 14 L 15 18 L 14 17 Z" fill="{dang}" />'
        html += apply_knockout(cross_fill, bg)

    # --- CATEGORIA 6: Layout, Fogli e Viste ---
    elif name == "scelta_viste":
        html += rect_fill(2, 4, 20, 16, m, rx=1.5)
        html += rect_fill(3.5, 5.5, 17, 13, bg, rx=0.5)
        html += rect_fill(4.5, 7, 4, 10, acc, rx=0.4)
        html += rect_fill(10, 7, 4, 10, warn, rx=0.4)
        html += rect_fill(15.5, 7, 4, 10, inf, rx=0.4)
    elif name == "adattaH":
        html += rect_fill(3, 11.25, 18, 1.5, m, rx=0.4)
        arrow_up = rect_fill(12, 4, 1.5, 5, acc, rx=0.4) + f'<path d="M 10 7 L 12.75 3.5 L 15.5 7 Z" fill="{acc}" />'
        arrow_down = rect_fill(12, 15, 1.5, 5, acc, rx=0.4) + f'<path d="M 10 17 L 12.75 20.5 L 15.5 17 Z" fill="{acc}" />'
        html += arrow_up + arrow_down
    elif name in ("griglia3", "mostra_griglia"):
        t = 1.5
        html += rect_fill(3, 3, 18, 18, m, rx=1.5)
        html += rect_fill(3+t, 3+t, 18-2*t, 18-2*t, bg, rx=0.5)
        html += rect_fill(8.5, 3, 1.5, 18, m)
        html += rect_fill(14, 3, 1.5, 18, m)
        html += rect_fill(3, 8.5, 18, 1.5, m)
        html += rect_fill(3, 14, 18, 1.5, m)
    elif name in ("vintage", "copertine"):
        html += rect_fill(4, 3, 16, 18, m, rx=1.5)
        html += rect_fill(5.5, 4.5, 13, 15, bg, rx=0.5)
        html += rect_fill(2, 6, 4, 1.5, acc, rx=0.4)
        html += rect_fill(2, 11, 4, 1.5, acc, rx=0.4)
        html += rect_fill(2, 16, 4, 1.5, acc, rx=0.4)
    elif name == "colore_tematico":
        # Disegno 100% fill-only del secchio e goccia
        html += f'<path d="M 6 10 L 14 4 L 20 10 L 12 16 Z" fill="{m}" />'
        # Manico secchio disegnato come nastro ad arco fill-only (ribbon)
        html += f'<path d="M 6 10 C 6 6, 20 6, 20 10 L 20 9 C 20 4.5, 6 4.5, 6 9 Z" fill="{m}" />'
        html += f'<path d="M 14 17 C 14 19, 16 19, 16 17 C 16 15, 14 15, 14 17 Z" fill="{acc}" />'

    # --- CATEGORIA 7: Reporting, Stampa ed Esportazione ---
    elif name == "riepilogo":
        html += doc_primitive(m)
        lines = rect_fill(8, 7, 6, 1.0, gray) + rect_fill(8, 10, 8, 1.0, gray) + rect_fill(8, 13, 8, 1.0, gray)
        html += apply_knockout(lines, bg)
        pen = f'<path d="M 14 18 L 18 14 L 21 17 L 17 21 Z" fill="{acc}" />' + f'<path d="M 13 19 L 14 18 L 13 21 Z" fill="{m}" />'
        html += apply_knockout(pen, bg)
    elif name == "riepilogo_quantita":
        html += doc_primitive(m)
        chart = rect_fill(8, 8, 4, 1.5, acc) + rect_fill(8, 11, 8, 1.5, warn) + rect_fill(8, 14, 6, 1.5, inf)
        html += apply_knockout(chart, bg)
    elif name == "riepilogo_a2":
        html += doc_primitive(m)
        grid = rect_fill(8, 7, 8, 8, gray, rx=0.5) + rect_fill(9, 8, 6, 6, bg, rx=0.4)
        html += apply_knockout(grid, bg)
    elif name == "print_ok":
        html += rect_fill(4, 8, 16, 8, m, rx=1.5)
        html += rect_fill(7, 4, 10, 4, gray, rx=0.5)
        html += rect_fill(6, 14, 12, 6, bg, rx=0.5) + rect_fill(8, 16, 8, 1.0, acc) + rect_fill(8, 18, 6, 1.0, acc)
    elif name in ("image100", "riga_rossa"):
        html += rect_fill(3, 14, 18, 3, dang, rx=0.5)
        pen = rect_fill(10, 4, 4, 6, warn, rx=0.5) + f'<path d="M 10 10 L 12 13 L 14 10 Z" fill="{dang}" />'
        html += apply_knockout(pen, bg)

    # --- CATEGORIA 8: Utility e Configurazioni ---
    elif name == "config":
        html += ring_primitive(m, bg, gray, r_outer=8, x=11, y=11)
        html += ring_primitive(m, bg, acc, r_outer=5, x=17, y=17)
    elif name in ("image16", "stringhe_numeri"):
        # Lettera A disegnata come path fill-only
        html += f'<path d="M 5 15 L 8 5 L 11 15 L 9.5 15 L 9 12 L 7 12 L 6.5 15 Z M 8 7 L 8.7 10.5 L 7.3 10.5 Z" fill="{m}" fill-rule="evenodd" />'
        # Freccia
        html += rect_fill(11, 10, 5, 1.5, acc, rx=0.4) + f'<path d="M 14 7.5 L 17.5 10.75 L 14 14 Z" fill="{acc}" />'
        # Numero 1
        html += rect_fill(19, 5, 1.5, 10, m, rx=0.4) + rect_fill(17.5, 5, 2, 1.5, m, rx=0.4)
    elif name in ("image17", "sproteggi_tutto"):
        body = rect_fill(5, 10, 14, 10, warn, rx=1.5) + rect_fill(10, 13, 4, 4, bg, rx=0.5)
        shackle_fill = f'<path d="M 8 10 L 8 7 C 8 4, 16 4, 16 7 L 16 12 L 14.5 12 L 14.5 7 C 14.5 5.5, 9.5 5.5, 9.5 7 L 9.5 10 Z" fill="{warn}" />'
        html += shackle_fill + body
    elif name in ("sfera_gialla", "importa_stili"):
        html += folder_primitive(gray)
        brush = rect_fill(12, 6, 2.5, 8, m, rx=0.5) + rect_fill(10, 12, 6.5, 4, warn, rx=1.0)
        html += apply_knockout(brush, bg)
    elif name in ("sf_Ver", "numeri_lettere"):
        # Numero 1 (fill-only)
        html += rect_fill(5, 5, 1.5, 10, m, rx=0.4) + rect_fill(3.5, 5, 2, 1.5, m, rx=0.4)
        # Freccia
        html += rect_fill(10, 12, 4, 1.5, acc, rx=0.4)
        # Lettera a (fill-only)
        html += f'<path d="M 16 14 C 15 14, 14 15, 14 16 L 14 18 C 14 19, 15 20, 16 20 M 16 14 L 16 20" fill="{m}" />'
        html += f'<path d="M 14.5 14.5 C 13 14.5, 13 19.5, 14.5 19.5 L 14.5 18 C 13.8 18, 13.8 16, 14.5 16 Z" fill="{m}" fill-rule="evenodd" />'

    # --- CATEGORIA 9: Sviluppatore e Importazioni Legacy ---
    elif name in ("py", "python_debug"):
        html += f'<path d="M 12 2 C 9 2, 8 4, 8 6 L 8 8 L 12 8 L 12 9 L 15 9 C 17 9, 18 8, 18 6 C 18 4, 17 2, 15 2 Z" fill="{inf}" />'
        html += f'<path d="M 12 22 C 15 22, 16 20, 16 18 L 16 16 L 12 16 L 12 15 L 9 15 C 7 15, 6 16, 6 18 C 6 20, 7 22, 9 22 Z" fill="{warn}" />'
        html += f'<circle cx="10" cy="4" r="0.6" fill="{bg}" />'
        html += f'<circle cx="14" cy="20" r="0.6" fill="{bg}" />'
    elif name == "refresh":
        html += f'<path d="M 12 4 C 16.42 4, 20 7.58, 20 12 L 18.5 12 C 18.5 8.41, 15.59 5.5, 12 5.5 C 8.41 5.5, 5.5 8.41, 5.5 12 L 4 12 C 4 7.58, 7.58 4, 12 4 Z" fill="{acc}" />'
        html += f'<path d="M 12 20 C 7.58 20, 4 16.42, 4 12 L 5.5 12 C 5.5 15.59, 8.41 18.5, 12 18.5 C 15.59 18.5, 18.5 15.59, 18.5 12 L 20 12 C 20 16.42, 16.42 20, 12 20 Z" fill="{acc}" />'
        html += f'<path d="M 2 12 L 4.75 15.5 L 7.5 12 Z" fill="{acc}" />'
        html += f'<path d="M 16.5 12 L 19.25 8.5 L 22 12 Z" fill="{acc}" />'
    elif name in ("falegname", "importa_dat"):
        html += f'<path d="M 7 6 L 3 12 L 7 18 L 8 17 L 4.7 12 L 8 7 Z" fill="{m}" />'
        html += f'<path d="M 13 6 L 17 12 L 13 18 L 12 17 L 15.3 12 L 12 7 Z" fill="{m}" />'
        html += rect_fill(10, 4, 1.5, 12, acc, rx=0.4) + f'<path d="M 7.5 13 L 10.75 16.5 L 14 13 Z" fill="{acc}" />'

    # --- NUOVE ICONE (Sezione 8) ---
    elif name == "importa_xml":
        html += doc_primitive(m)
        text = rect_fill(8, 8, 8, 1.5, inf) + rect_fill(8, 11, 6, 1.5, inf)
        html += apply_knockout(text, bg)
        arrow = rect_fill(6, 14, 1.5, 6, acc, rx=0.4) + f'<path d="M 3.5 17.5 L 6.75 21 L 10 17.5 Z" fill="{acc}" />'
        html += apply_knockout(arrow, bg)
    elif name == "esporta_gantt":
        html += folder_primitive(gray)
        gantt_b1 = rect_fill(5, 8, 6, 1.5, acc, rx=0.4)
        gantt_b2 = rect_fill(9, 11, 8, 1.5, warn, rx=0.4)
        gantt_b3 = rect_fill(13, 14, 5, 1.5, inf, rx=0.4)
        html += apply_knockout(gantt_b1 + gantt_b2 + gantt_b3, bg)
        arrow = rect_fill(15, 6, 6, 1.5, acc, rx=0.4) + f'<path d="M 18.5 3.5 L 22 6.75 L 18.5 10 Z" fill="{acc}" />'
        html += apply_knockout(arrow, bg)
    elif name == "documento_bollo":
        html += doc_primitive(m)
        stamp = ring_primitive(dang, bg, dang, r_outer=4, x=15, y=15)
        html += apply_knockout(stamp, bg)
    elif name == "unisci_fogli":
        sheet1 = folder_primitive(gray, offset_x=-3, offset_y=-2)
        sheet2 = folder_primitive(m, offset_x=2, offset_y=2)
        html += sheet1 + apply_knockout(sheet2, bg)
        arrow = rect_fill(6, 12, 5, 1.5, acc, rx=0.4) + f'<path d="M 8.5 9.5 L 12 12.75 L 8.5 16 Z" fill="{acc}" />'
        html += apply_knockout(arrow, bg)
    elif name == "somma_colore":
        html += f'<path d="M 5 5 L 13 5 L 9 12 L 13 19 L 5 19 L 5 17.5 L 10.5 17.5 L 7.5 12 L 10.5 6.5 L 5 6.5 Z" fill="{acc}" />'
        pen = rect_fill(14, 4, 3, 10, warn, rx=0.5) + f'<path d="M 14 14 L 15.5 17 L 17 14 Z" fill="{acc}" />'
        html += apply_knockout(pen, bg)

    else:
        html += doc_primitive(m)
        html += badge_check(acc, bg, x=16, y=16)

    return make_svg_wrapper(html)

def main():
    icon_names = [
        "leeno", "manuale", "teleg", "supcat", "cat", "subcat",
        "struttura_on", "struttura_off", "rinumCap", "nuova_voce",
        "voce_breve", "vedivoce", "pesca", "invia_voce_ep",
        "aggiungi_misura", "sposta_voce", "analisi_a_prezzo",
        "utili_maggiorazioni", "elimina_doppioni", "riordina",
        "parz", "inverti_segno", "azzera", "partita_provvisoria_piu", "part_agg",
        "partita_provvisoria_meno", "part_det", "strutt_voci_zero",
        "elimina_azzerate", "elimina_vuote", "scelta_viste", "adattaH",
        "griglia3", "copertine", "colore_tematico", "riepilogo",
        "riepilogo_quantita", "riepilogo_a2", "print_ok", "image100", "riga_rossa",
        "config", "stringhe_numeri", "sproteggi_tutto", "importa_stili",
        "numeri_lettere", "python_debug", "refresh", "importa_dat",
        "importa_xml", "esporta_gantt", "documento_bollo", "unisci_fogli",
        "somma_colore",
        # Altre icone storiche scoperte nella directory
        "Caschetto", "Corta", "azzera", "cestino", "compo", "cross", "decimali",
        "elimina_azzerate", "elimina_vuote", "falegname", "findcode", "image14",
        "image15", "image16", "image17", "image18", "image21B", "image37", "image444",
        "image8", "image9", "image93", "imageNA", "invert", "invia_voce_ep", "parz",
        "perc", "pesca", "print_ok", "py", "refresh", "ricicla", "riepilogo",
        "riepilogo_a2", "riepilogo_quantita", "rinumCap", "sf_Ver", "sfera_gialla",
        "strutt_voci_zero", "unisci_fogli", "vedivoce", "vintage"
    ]

    unique_names = []
    for name in icon_names:
        if name not in unique_names:
            unique_names.append(name)

    base_path = "src/Ultimus.oxt/icons"
    svg_dir = os.path.join(base_path, "svg")
    scuro_dir = os.path.join(base_path, "scuro")

    os.makedirs(svg_dir, exist_ok=True)
    os.makedirs(scuro_dir, exist_ok=True)

    print(f"Inizio generazione di {len(unique_names)} icone per entrambi i temi...")

    # Genera icone in chiaro
    colors_light = get_colors(is_dark=False)
    for name in unique_names:
        svg_content = build_icon(name, colors_light)
        filepath = os.path.join(svg_dir, f"{name}.svg")
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(svg_content)

    # Genera icone in scuro
    colors_dark = get_colors(is_dark=True)
    for name in unique_names:
        svg_content = build_icon(name, colors_dark)
        filepath = os.path.join(scuro_dir, f"{name}.svg")
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(svg_content)

    print("Generazione completata con successo!")

if __name__ == "__main__":
    main()
