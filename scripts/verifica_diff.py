#!/usr/bin/env python3
"""
verifica_diff.py

Controlli meccanici per rivedere una PR (propria o di un agente) confrontando
due revisioni git sui file .py sotto src/, tests/ e scripts/.

Uso:
    python3 scripts/verifica_diff.py <base> <head>
    python3 scripts/verifica_diff.py origin/dev pr118

Per ogni file riporta: righe aggiunte e tolte, conteggio CRLF/LF, se l'AST e'
identico a parte le docstring, funzioni senza docstring, except silenziosi
(solo pass), avvisi pyflakes (totale e nomi non definiti, se pyflakes e'
installato) e i nomi di modulo rimossi che altri file usano ancora.
Non modifica nulla. Le righe ATTENZIONE vanno guardate a mano.
Exit code sempre 0: e' un ausilio alla revisione, non un controllo di CI.

Copyright Giuseppe Vizziello - supporto@leeno.org - LGPLv2.1
"""

import argparse
import ast
import os
import re
import subprocess
import sys
import tempfile

PREFISSI = ('src/', 'tests/', 'scripts/')


def git(*args):
    r = subprocess.run(['git', *args], capture_output=True)
    return r.returncode, r.stdout


def contenuto(rev, path):
    rc, out = git('show', f'{rev}:{path}')
    return out if rc == 0 else None


def parse(dati):
    return ast.parse(dati.decode('utf-8', errors='replace'))


def ast_senza_docstring(dati):
    tree = parse(dati)
    for n in ast.walk(tree):
        if isinstance(n, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)) and n.body:
            primo = n.body[0]
            if (isinstance(primo, ast.Expr) and isinstance(getattr(primo, 'value', None), ast.Constant)
                    and isinstance(primo.value.value, str)):
                n.body = n.body[1:] or [ast.Pass()]
    return ast.dump(tree)


def funzioni_senza_docstring(dati):
    return sum(1 for n in ast.walk(parse(dati))
               if isinstance(n, ast.FunctionDef) and not ast.get_docstring(n))


def except_silenziosi(dati):
    tot = 0
    for n in ast.walk(parse(dati)):
        if (isinstance(n, ast.ExceptHandler) and len(n.body) == 1 and isinstance(n.body[0], ast.Pass)
                and (n.type is None or getattr(n.type, 'id', '') == 'Exception')):
            tot += 1
    return tot


def nomi_di_modulo(dati):
    nomi = set()
    tree = parse(dati)
    for n in ast.walk(tree):
        if isinstance(n, (ast.Import, ast.ImportFrom)) and n.col_offset == 0:
            for a in n.names:
                nomi.add((a.asname or a.name).split('.')[0])
    for n in tree.body:
        if isinstance(n, (ast.FunctionDef, ast.ClassDef)):
            nomi.add(n.name)
        elif isinstance(n, ast.Assign):
            for t in n.targets:
                if isinstance(t, ast.Name):
                    nomi.add(t.id)
    return nomi


def pyflakes(dati):
    """Restituisce (avvisi_totali, nomi_non_definiti) oppure None se pyflakes manca."""
    with tempfile.NamedTemporaryFile('wb', suffix='.py', delete=False) as f:
        f.write(dati)
        nome = f.name
    try:
        r = subprocess.run([sys.executable, '-m', 'pyflakes', nome], capture_output=True, text=True)
    finally:
        os.unlink(nome)
    if 'No module named pyflakes' in r.stderr:
        return None
    righe = [x for x in r.stdout.splitlines() if x.strip()]
    return len(righe), sum(1 for x in righe if 'undefined name' in x)


def usi_esterni(head, path, nome):
    """Cerca nome usato da altri file come Modulo.nome, 'from Modulo import nome' o dispatcher."""
    modulo = os.path.splitext(os.path.basename(path))[0]
    alias = [modulo] + (['PL'] if modulo == 'pyleeno' else [])
    n = re.escape(nome)
    pattern = '|'.join([rf'\b({"|".join(alias)})\.{n}\b',
                        rf'from {modulo} import[^\n]*\b{n}\b',
                        rf'\?{modulo}\.{n}\b'])
    rc, out = git('grep', '-n', '-E', '-e', pattern, head, '--', 'src')
    righe = out.decode('utf-8', errors='replace').splitlines() if rc == 0 else []
    return [x for x in righe if not x.split(':', 2)[1].endswith(os.path.basename(path))]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('base')
    ap.add_argument('head')
    args = ap.parse_args()

    rc, out = git('diff', '--numstat', args.base, args.head)
    if rc != 0:
        print('ERRORE: revisioni non valide')
        return 0
    file_diff = []
    for riga in out.decode('utf-8', errors='replace').splitlines():
        a, d, path = riga.split('\t', 2)
        file_diff.append((path, a, d))
    if not file_diff:
        print('ATTENZIONE: nessun file modificato tra le due revisioni (PR vuota?)')
        return 0

    print(f'{len(file_diff)} file modificati ({args.base} -> {args.head})')
    for path, a, d in file_diff:
        print(f'\n== {path}  (+{a} -{d})')
        if not path.endswith('.py') or not path.startswith(PREFISSI):
            print('   non Python o fuori da src/tests/scripts: controllo manuale')
            continue
        prima, dopo = contenuto(args.base, path), contenuto(args.head, path)
        if prima is None:
            print('   file nuovo: controllo manuale del contenuto')
            continue
        if dopo is None:
            print('   ATTENZIONE: file rimosso')
            continue
        crlf = (prima.count(b'\r\n'), dopo.count(b'\r\n'))
        lf = (prima.count(b'\n') - crlf[0], dopo.count(b'\n') - crlf[1])
        flag = '' if (crlf[0] == 0) == (crlf[1] == 0) else '  ATTENZIONE: cambia lo stile di fine riga'
        print(f'   CRLF {crlf[0]} -> {crlf[1]} | LF {lf[0]} -> {lf[1]}{flag}')
        try:
            ident = ast_senza_docstring(prima) == ast_senza_docstring(dopo)
            print(f'   AST identico a parte le docstring: {"si" if ident else "no"}')
            print(f'   funzioni senza docstring: {funzioni_senza_docstring(prima)} -> {funzioni_senza_docstring(dopo)}')
            e0, e1 = except_silenziosi(prima), except_silenziosi(dopo)
            print(f'   except silenziosi (solo pass): {e0} -> {e1}'
                  + ('  ATTENZIONE: aumentano' if e1 > e0 else ''))
            p0, p1 = pyflakes(prima), pyflakes(dopo)
            if p0 is None:
                print('   pyflakes non installato: saltato')
            else:
                print(f'   pyflakes: {p0[0]} -> {p1[0]} avvisi | nomi non definiti {p0[1]} -> {p1[1]}'
                      + ('  ATTENZIONE: aumentano i nomi non definiti' if p1[1] > p0[1] else ''))
            if '/pythonpath/' in path:
                for nome in sorted(nomi_di_modulo(prima) - nomi_di_modulo(dopo)):
                    usi = usi_esterni(args.head, path, nome)
                    if usi:
                        print(f'   ATTENZIONE: nome rimosso "{nome}" ancora usato in: {"; ".join(usi[:3])}')
                    else:
                        print(f'   nome di modulo rimosso: {nome} (nessun uso esterno trovato)')
        except SyntaxError as e:
            print(f'   ATTENZIONE: errore di sintassi ({e})')
    print('\nDa controllare comunque a mano: dispatcher e menu (.xcu, .xba, .xdl), docstring a campione, prova in Calc.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
