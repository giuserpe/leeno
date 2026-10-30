import ast
import os
import shutil
import tempfile
import unittest
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src/Ultimus.oxt/python/pythonpath/LeenoUtils.py"


def _carica_log_errore():
    """Estrae log_errore e il suo contatore da LeenoUtils senza importare il modulo (richiede uno)."""
    tree = ast.parse(SRC.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name == "log_errore":
            nodi.append(n)
        elif isinstance(n, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id == "_LOG_ERRORI_CONTEGGI" for t in n.targets
        ):
            nodi.append(n)
    ns = {}
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(SRC), "exec"), ns)
    return ns


class TestLogErrore(unittest.TestCase):
    def setUp(self):
        self.ns = _carica_log_errore()
        self.dir = tempfile.mkdtemp()
        self.file = os.path.join(self.dir, "leeno_errori.log")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _leggi(self):
        with open(self.file, encoding="utf-8") as f:
            return f.read()

    def test_scrive_contesto_e_tipo(self):
        try:
            raise ValueError("cella non valida")
        except ValueError as e:
            self.assertTrue(self.ns["log_errore"]("prova", e, percorso=self.file))
        testo = self._leggi()
        self.assertIn("[prova] ValueError: cella non valida", testo)
        self.assertIn("Traceback", testo)

    def test_senza_eccezione(self):
        self.assertTrue(self.ns["log_errore"]("solo contesto", percorso=self.file))
        self.assertIn("[solo contesto]", self._leggi())

    def test_limite_tre_occorrenze(self):
        esiti = []
        for _ in range(6):
            try:
                raise KeyError("x")
            except KeyError as e:
                esiti.append(self.ns["log_errore"]("ciclo", e, percorso=self.file))
        self.assertEqual(esiti, [True, True, True, False, False, False])
        testo = self._leggi()
        self.assertEqual(testo.count("[ciclo]"), 3)
        self.assertIn("ulteriori occorrenze non registrate", testo)

    def test_contesti_diversi_non_si_limitano(self):
        for i in range(5):
            self.assertTrue(self.ns["log_errore"]("ctx%d" % i, percorso=self.file))

    def test_rotazione_oltre_soglia(self):
        with open(self.file, "w", encoding="utf-8") as f:
            f.write("x" * 524289)
        self.assertTrue(self.ns["log_errore"]("dopo rotazione", percorso=self.file))
        self.assertTrue(os.path.exists(self.file + ".old"))
        self.assertLess(os.path.getsize(self.file), 1000)

    def test_non_solleva_con_percorso_non_valido(self):
        percorso = os.path.join(self.dir, "inesistente", "sub", "log.txt")
        self.assertFalse(self.ns["log_errore"]("percorso rotto", percorso=percorso))


if __name__ == "__main__":
    unittest.main()
