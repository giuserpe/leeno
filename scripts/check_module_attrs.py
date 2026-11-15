#!/usr/bin/env python3
"""
check_module_attrs.py

Verifica che ogni riferimento a un attributo di un modulo locale esista
davvero nel modulo. pyflakes non controlla gli attributi di modulo, quindi
una chiamata come `LeenoUtils.funzioneInesistente()` passa inosservata e
fallisce solo a runtime, dentro LibreOffice.

Controlla, nei moduli di python/ e python/pythonpath/:

1. `import Modulo` / `import Modulo as M`: ogni `M.nome` deve essere definito
   nel modulo (funzione, classe, variabile, import, o nome dichiarato
   `global` in una funzione).
2. `from Modulo import nome`: `nome` deve essere definito in Modulo.

Sono ignorati i moduli con `__getattr__`, `import *` o scritture dinamiche
su `globals()` (`.update()`, chiave non costante), perche' i loro nomi non
sono determinabili staticamente, e gli attributi protetti da
`hasattr(M, 'nome')` o `getattr(M, 'nome', default)`.

Uso:
    python3 scripts/check_module_attrs.py
    python3 scripts/check_module_attrs.py --path <cartella_python>

Exit code 0 = ok, 1 = riferimenti non risolti.
"""

import argparse
import ast
import sys
from pathlib import Path

DEFAULT_PYTHON = "src/Ultimus.oxt/python"


def _targets(node):
    """Nomi assegnati da un target di assegnazione."""
    if isinstance(node, ast.Name):
        yield node.id
    elif isinstance(node, (ast.Tuple, ast.List)):
        for e in node.elts:
            yield from _targets(e)
    elif isinstance(node, ast.Starred):
        yield from _targets(node.value)


def _is_globals_call(node):
    return (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
            and node.func.id == "globals" and not node.args)


def _scope_statements(body):
    """Istruzioni di livello modulo, entrando in if/try/with/for/while."""
    for st in body:
        yield st
        for attr in ("body", "orelse", "finalbody"):
            sub = getattr(st, attr, None)
            if sub and not isinstance(st, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                yield from _scope_statements(sub)
        for h in getattr(st, "handlers", []) or []:
            yield from _scope_statements(h.body)


def module_names(tree):
    """(nomi definiti, True se i nomi non sono determinabili)."""
    names = set()
    dynamic = False
    for st in _scope_statements(tree.body):
        if isinstance(st, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            names.add(st.name)
            if st.name == "__getattr__":
                dynamic = True
        elif isinstance(st, ast.Assign):
            for t in st.targets:
                names.update(_targets(t))
        elif isinstance(st, (ast.AnnAssign, ast.AugAssign)):
            names.update(_targets(st.target))
        elif isinstance(st, (ast.For, ast.AsyncFor)):
            names.update(_targets(st.target))
        elif isinstance(st, ast.With):
            for it in st.items:
                if it.optional_vars is not None:
                    names.update(_targets(it.optional_vars))
        elif isinstance(st, ast.Import):
            for a in st.names:
                names.add(a.asname or a.name.split(".")[0])
        elif isinstance(st, ast.ImportFrom):
            for a in st.names:
                if a.name == "*":
                    dynamic = True
                else:
                    names.add(a.asname or a.name)
    for node in ast.walk(tree):
        if isinstance(node, ast.Global):
            names.update(node.names)
        elif (isinstance(node, ast.Subscript) and isinstance(node.ctx, (ast.Store, ast.Del))
              and _is_globals_call(node.value)):
            # globals()['nome'] = ...: definisce 'nome'; chiave non costante: dinamico
            if isinstance(node.slice, ast.Constant) and isinstance(node.slice.value, str):
                names.add(node.slice.value)
            else:
                dynamic = True
        elif (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
              and _is_globals_call(node.func.value)
              and node.func.attr in ("update", "setdefault", "pop", "clear")):
            dynamic = True
    return names, dynamic


def collect_modules(root: Path):
    mods = {}
    for folder in (root, root / "pythonpath"):
        if not folder.is_dir():
            continue
        for f in sorted(folder.glob("*.py")):
            try:
                tree = ast.parse(f.read_text(encoding="utf-8", errors="replace"))
            except SyntaxError:
                continue
            names, dynamic = module_names(tree)
            mods[f.stem] = (f, tree, names, dynamic)
    return mods


def find_problems(mods):
    problems = []
    for stem, (path, tree, _names, _dyn) in mods.items():
        aliases = {}
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for a in node.names:
                    if a.name in mods and a.name != stem:
                        aliases[a.asname or a.name] = a.name
            elif isinstance(node, ast.ImportFrom) and node.module in mods and node.module != stem:
                _, _, names, dynamic = mods[node.module]
                if dynamic:
                    continue
                for a in node.names:
                    if a.name != "*" and a.name not in names:
                        problems.append((path, node.lineno, f"from {node.module} import {a.name}", node.module))
        # attributi protetti da hasattr(M, 'nome') / getattr(M, 'nome', default)
        guarded = set()
        for node in ast.walk(tree):
            if (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
                    and node.func.id in ("hasattr", "getattr") and len(node.args) >= 2
                    and isinstance(node.args[0], ast.Name)
                    and isinstance(node.args[1], ast.Constant)
                    and (node.func.id == "hasattr" or len(node.args) >= 3)):
                guarded.add((node.args[0].id, node.args[1].value))
        # un alias riassegnato a livello di modulo non e' piu' un modulo
        rebound = set()
        for st in _scope_statements(tree.body):
            if isinstance(st, ast.Assign):
                for t in st.targets:
                    rebound.update(_targets(t))
        for node in ast.walk(tree):
            if isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name):
                alias = node.value.id
                if alias in aliases and alias not in rebound and (alias, node.attr) not in guarded:
                    modname = aliases[alias]
                    _, _, names, dynamic = mods[modname]
                    if not dynamic and node.attr not in names:
                        problems.append((path, node.lineno, f"{alias}.{node.attr}", modname))
    return sorted(set(problems), key=lambda p: (str(p[0]), p[1], p[2]))


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--path", default=DEFAULT_PYTHON, help="cartella python/ dell'estensione")
    args = ap.parse_args(argv)
    root = Path(args.path)
    if not root.is_dir():
        print(f"Cartella non trovata: {root}", file=sys.stderr)
        return 2
    mods = collect_modules(root)
    problems = find_problems(mods)
    if not problems:
        print(f"Nessun riferimento non risolto in {root} ({len(mods)} moduli)")
        return 0
    print(f"{len(problems)} riferimenti non risolti in {root}:")
    for path, line, ref, modname in problems:
        print(f"  {path.relative_to(root)}:{line}: {ref}  (non definito in {modname}.py)")
    return 1


if __name__ == "__main__":
    sys.exit(main())
