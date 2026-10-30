# Copyright (C) Giuseppe Vizziello - supporto@leeno.org. Licenza LGPL 2.1 (vedi LICENSE nella radice del repository).
"""Test AST per verificare l'assenza di 'except Exception: pass' o 'except: pass' nei try modificati."""

import ast
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
LEENO_CONTAB_PATH = REPO_ROOT / "src" / "Ultimus.oxt" / "python" / "pythonpath" / "LeenoContab.py"
PYLEENO_PATH = REPO_ROOT / "src" / "Ultimus.oxt" / "python" / "pythonpath" / "pyleeno.py"

TARGET_MODIFIED_TRIES = [
    ("LeenoContab.py", "firme_contabili_sal", "oDoc.calculate()"),
    ("pyleeno.py", "recupera_voce", "RangeAddress"),
    ("pyleeno.py", "_applica_altezza_pending", "oSheet.protect"),
    ("pyleeno.py", "rompi_collegamenti_esterni", "AreaLinks"),
    ("pyleeno.py", "rompi_collegamenti_esterni", "SheetLinkMode"),
    ("pyleeno.py", "rompi_collegamenti_esterni", "getSheets"),
    ("pyleeno.py", "rompi_collegamenti_esterni", "createReplaceDescriptor"),
    ("pyleeno.py", "MENU_prefisso_VDS_", "CurrentController.ActiveSheet"),
    ("pyleeno.py", "MENU_prefisso_VDS_", "circoscriveVoceComputo"),
    ("pyleeno.py", "MENU_costo_elementare_", "LeggiPosizioneCorrente"),
    ("pyleeno.py", "MENU_prefisso_codice", "getSheets().getByName(el)"),
    ("pyleeno.py", "scelta_viste_run", "int(oDialog1.getControl('ComboBox1').getText())"),
    ("pyleeno.py", "MENU_azzera_voce", "CurrentController.ActiveSheet"),
    ("pyleeno.py", "MENU_azzera_voce", "circoscriveVoceComputo"),
    ("pyleeno.py", "MENU_converti_stringhe", "Type.value == 'TEXT'"),
    ("pyleeno.py", "adegua_tmpl", "rigenera_tutte()"),
    ("pyleeno.py", "MENU_hl", "getCellByPosition(lcol, el)"),
]


def test_no_silent_except_in_modified_tries():
    file_map = {
        "LeenoContab.py": LEENO_CONTAB_PATH,
        "pyleeno.py": PYLEENO_PATH,
    }

    for filename, fn_name, first_stmt_sub in TARGET_MODIFIED_TRIES:
        filepath = file_map[filename]
        content = filepath.read_text(encoding="utf-8")
        tree = ast.parse(content, str(filepath))

        matching_tries = []
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name in (fn_name, "invia_voce"):
                for child in ast.walk(node):
                    if isinstance(child, ast.Try) and child.body:
                        body_code = ast.unparse(child.body)
                        if first_stmt_sub in body_code:
                            matching_tries.append(child)

        assert len(matching_tries) > 0, (
            f"Try contenente '{first_stmt_sub}' non trovato in {filename}:{fn_name}"
        )

        for target_try in matching_tries:
            for handler in target_try.handlers:
                if len(handler.body) == 1 and isinstance(handler.body[0], ast.Pass):
                    h_type = ast.unparse(handler.type) if handler.type else "bare"
                    assert False, (
                        f"Trovato 'except {h_type}: pass' silenzioso nel try alla riga {target_try.lineno} "
                        f"di '{fn_name}' in {filename}"
                    )
