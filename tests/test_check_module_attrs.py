import importlib.util
import tempfile
import unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parent.parent / "scripts" / "check_module_attrs.py"
_spec = importlib.util.spec_from_file_location("check_module_attrs", SCRIPT)
cma = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(cma)


class TestCheckModuleAttrs(unittest.TestCase):
    def _run(self, files):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "pythonpath").mkdir()
            for name, src in files.items():
                (root / "pythonpath" / f"{name}.py").write_text(src, encoding="utf-8")
            problems = cma.find_problems(cma.collect_modules(root))
            return [p[2] for p in problems]

    def test_attributo_inesistente_segnalato(self):
        r = self._run({
            "A": "def esiste():\n    pass\n",
            "B": "import A\n\ndef f():\n    A.esiste()\n    A.manca()\n",
        })
        self.assertEqual(r, ["A.manca"])

    def test_alias_e_from_import(self):
        r = self._run({
            "A": "x = 1\n",
            "B": "import A as M\nfrom A import y\n\ndef f():\n    return M.x, M.z\n",
        })
        self.assertCountEqual(r, ["M.z", "from A import y"])

    def test_nomi_definiti_in_vari_modi(self):
        r = self._run({
            "A": ("import os\nfrom math import pi\n\nclass C:\n    pass\n\n"
                  "a, b = 1, 2\ntry:\n    c = 1\nexcept Exception:\n    c = 2\n\n"
                  "def f():\n    global dinamico\n    dinamico = 1\n"),
            "B": "import A\n\nA.os, A.pi, A.C, A.a, A.b, A.c, A.f, A.dinamico\n",
        })
        self.assertEqual(r, [])

    def test_hasattr_e_getattr_con_default_non_segnalati(self):
        r = self._run({
            "A": "x = 1\n",
            "B": ("import A\n\nif hasattr(A, 'opzionale'):\n    A.opzionale()\n"
                  "g = getattr(A, 'altro', None)\n"),
        })
        self.assertEqual(r, [])

    def test_globals_in_lettura_non_rende_dinamico(self):
        r = self._run({
            "A": "def f():\n    return 'x' in globals()\n",
            "B": "import A\n\nA.manca\n",
        })
        self.assertEqual(r, ["A.manca"])

    def test_globals_scrittura_con_chiave_costante_definisce_il_nome(self):
        r = self._run({
            "A": "def f():\n    globals()['creato'] = 1\n",
            "B": "import A\n\nA.creato\n",
        })
        self.assertEqual(r, [])

    def test_globals_update_rende_dinamico(self):
        r = self._run({
            "A": "globals().update({'x': 1})\n",
            "B": "import A\n\nA.qualsiasi\n",
        })
        self.assertEqual(r, [])

    def test_modulo_con_getattr_o_star_ignorato(self):
        r = self._run({
            "A": "def __getattr__(name):\n    return 1\n",
            "S": "from os import *\n",
            "B": "import A\nimport S\n\nA.qualsiasi\nS.altro\n",
        })
        self.assertEqual(r, [])


if __name__ == "__main__":
    unittest.main()
