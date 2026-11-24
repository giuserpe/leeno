# -*- coding: utf-8 -*-
"""
Helper dei test: compila funzioni e costanti di un modulo di pythonpath senza importarlo
(i moduli di LeenO dipendono da uno e da LibreOffice) e un foglio finto con celle in memoria.
"""
import ast
from pathlib import Path
from types import SimpleNamespace


def carica(percorso, funzioni=(), costanti=(), ns=None, classi=()):
    """Ritorna il namespace con le funzioni e le costanti indicate, compilate dal sorgente."""
    percorso = Path(percorso)
    albero = ast.parse(percorso.read_text(encoding="utf-8"))
    nodi = []
    for n in albero.body:
        if isinstance(n, ast.FunctionDef) and n.name in funzioni:
            nodi.append(n)
        elif isinstance(n, ast.ClassDef) and n.name in classi:
            nodi.append(n)
        elif isinstance(n, ast.Assign) and getattr(n.targets[0], "id", None) in costanti:
            nodi.append(n)
    trovati = {getattr(n, "name", None) or n.targets[0].id for n in nodi}
    mancanti = (set(funzioni) | set(costanti) | set(classi)) - trovati
    assert not mancanti, "non trovati in %s: %s" % (percorso.name, sorted(mancanti))
    ns = {} if ns is None else ns
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(percorso), "exec"), ns)
    return ns


class FoglioFinto:
    """Foglio con celle in un dizionario {(colonna, riga): testo}."""

    def __init__(self, celle, righe=None):
        self.celle = celle
        self.ultima = righe if righe is not None else max((r for _, r in celle), default=0)

    def getCellByPosition(self, col, riga):
        return SimpleNamespace(String=self.celle.get((col, riga), ""))

    def createCursorByRange(self, _cella):
        return SimpleNamespace(
            gotoEndOfUsedArea=lambda _x: None,
            RangeAddress=SimpleNamespace(EndRow=self.ultima))
