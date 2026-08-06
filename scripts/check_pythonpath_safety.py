#!/usr/bin/env python3
"""
check_pythonpath_safety.py

Verifica che src/Ultimus.oxt/python/pythonpath/ rispetti le regole di
sicurezza descritte in AGENTS.md, sezione "Sicurezza dei moduli in
pythonpath/":

1. Nessun file di test (test_*.py, *_test.py) nella cartella: qualunque
   file qui dentro puo' essere importato dal processo di LibreOffice
   per motivi indipendenti dal task che lo ha creato (esplorazione
   macro, importlib.reload di recupero, tool di indicizzazione).

2. Nessuna assegnazione a sys.modules[...] a livello di modulo (fuori
   da funzioni/classi). Questo e' il pattern che, se importato anche
   una sola volta dentro LibreOffice, sostituisce silenziosamente i
   moduli reali (uno, unohelper, pyleeno, Dialogs, ecc.) con dei mock
   per l'intera sessione.

Uso:
    python3 scripts/check_pythonpath_safety.py
    python3 scripts/check_pythonpath_safety.py --path <cartella_pythonpath>

Exit code 0 = ok, 1 = violazioni trovate (pensato per bloccare la CI).
"""

import argparse
import ast
import sys
from pathlib import Path

DEFAULT_PYTHONPATH = "src/Ultimus.oxt/python/pythonpath"


def is_test_filename(filename: str) -> bool:
    return filename.startswith("test_") or filename.endswith("_test.py")


def find_module_level_sys_modules_assignments(tree: ast.Module, source: str):
    """
    Ritorna una lista di (lineno, snippet) per ogni assegnazione a
    sys.modules[...] (o <alias>.modules[...] se sys e' importato con
    un alias) che si trova nel BODY DI MODULO, cioe' non annidata in
    FunctionDef/AsyncFunctionDef/ClassDef.

    Nota: un'assegnazione dentro una ClassDef ma fuori da un metodo e'
    comunque eseguita all'import della classe (corpo di classe eseguito
    a import-time), quindi viene trattata come "a livello di modulo".
    Solo l'annidamento dentro Function/AsyncFunctionDef mette in salvo.
    """
    violations = []
    lines = source.splitlines()

    # Alias con cui 'sys' e' stato importato in questo file (di solito 'sys')
    sys_aliases = {"sys"}
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "sys":
                    sys_aliases.add(alias.asname or alias.name)

    def is_sys_modules_subscript(target: ast.expr) -> bool:
        # target atteso: Subscript(value=Attribute(value=Name(id='sys'), attr='modules'), ...)
        if not isinstance(target, ast.Subscript):
            return False
        value = target.value
        return (
            isinstance(value, ast.Attribute)
            and value.attr == "modules"
            and isinstance(value.value, ast.Name)
            and value.value.id in sys_aliases
        )

    class FunctionBoundaryVisitor(ast.NodeVisitor):
        """Visita l'intero albero ma NON scende dentro i corpi di
        funzione/funzione-async: tutto cio' che vede e' quindi
        eseguito a import-time (modulo o corpo di classe)."""

        def visit_FunctionDef(self, node):
            return  # non entrare: codice eseguito solo alla chiamata

        def visit_AsyncFunctionDef(self, node):
            return  # idem

        def visit_Lambda(self, node):
            return  # idem

        def _check_assign_targets(self, targets, node):
            for target in targets:
                if is_sys_modules_subscript(target):
                    lineno = node.lineno
                    snippet = lines[lineno - 1].strip() if 0 < lineno <= len(lines) else ""
                    violations.append((lineno, snippet))

        def visit_Assign(self, node):
            self._check_assign_targets(node.targets, node)
            self.generic_visit(node)

        def visit_AugAssign(self, node):
            self._check_assign_targets([node.target], node)
            self.generic_visit(node)

    FunctionBoundaryVisitor().visit(tree)
    return violations


def check_file(path: Path):
    """Ritorna una lista di stringhe di errore per il file dato (vuota se ok)."""
    errors = []

    if is_test_filename(path.name):
        errors.append(
            f"{path}: file di test in pythonpath/ (vietato — spostare in tests/, "
            f"escluso dal sys.path dell'estensione, oppure rimuovere se non "
            f"serve al funzionamento di LeenO)"
        )
        # Non serve nemmeno fare il parse AST: e' gia' una violazione di per se'.
        return errors

    try:
        source = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        source = path.read_text(encoding="utf-16")

    try:
        tree = ast.parse(source, filename=str(path))
    except SyntaxError as e:
        errors.append(f"{path}: impossibile analizzare il file ({e}) — controllare manualmente")
        return errors

    for lineno, snippet in find_module_level_sys_modules_assignments(tree, source):
        errors.append(
            f"{path}:{lineno}: assegnazione a sys.modules[...] a livello di modulo "
            f"('{snippet}') — se questo file viene importato dentro LibreOffice, "
            f"sostituisce silenziosamente i moduli reali per l'intera sessione"
        )

    return errors


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--path",
        default=DEFAULT_PYTHONPATH,
        help=f"Cartella pythonpath da controllare (default: {DEFAULT_PYTHONPATH})",
    )
    args = parser.parse_args()

    root = Path(args.path)
    if not root.is_dir():
        print(f"❌ Cartella non trovata: {root}", file=sys.stderr)
        return 1

    all_errors = []
    for py_file in sorted(root.rglob("*.py")):
        all_errors.extend(check_file(py_file))

    if all_errors:
        print(f"❌ Violazioni delle regole di sicurezza pythonpath/ ({len(all_errors)}):\n")
        for err in all_errors:
            print(f"  - {err}")
        print(
            "\nVedi AGENTS.md, sezione 'Sicurezza dei moduli in pythonpath/', "
            "per la motivazione completa."
        )
        return 1

    print(f"✅ Nessuna violazione trovata in {root}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
