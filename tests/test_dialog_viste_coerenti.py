"""Coerenza tra i dialoghi delle viste (xdl), le loro stringhe e il codice che li usa.

Un xdl che richiama una stringa assente dal .properties fa fallire la creazione
del dialogo: scelta_viste ritorna senza aprirlo e senza segnalare errori.
"""
import os
import re
import unittest

BASE = os.path.join(os.path.dirname(__file__), "..", "src", "Ultimus.oxt")
DIALOGHI = os.path.join(BASE, "UltimusFree2")
PROPERTIES = os.path.join(DIALOGHI, "DialogStrings_it_IT.properties")
PYLEENO = os.path.join(BASE, "python", "pythonpath", "pyleeno.py")
DIALOGHI_VISTE = ("Dialogviste_N", "DialogViste_A", "DialogViste_EP")

# controlli letti con getControl nel ramo CONTABILITA ma non presenti nell'xdl:
# il codice li gestisce con un controllo "is not None"
CONTROLLI_FACOLTATIVI = {"EliminaMisure"}


def _leggi(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def _chiavi_properties():
    chiavi = set()
    for riga in _leggi(PROPERTIES).splitlines():
        if "=" in riga and not riga.startswith("#"):
            chiavi.add(riga.split("=", 1)[0].strip())
    return chiavi


class TestDialogViste(unittest.TestCase):
    def test_stringhe_richiamate_esistono(self):
        chiavi = _chiavi_properties()
        for nome in DIALOGHI_VISTE:
            xdl = _leggi(os.path.join(DIALOGHI, nome + ".xdl"))
            riferimenti = set(re.findall(r"&amp;(\d+\.[\w.]+)", xdl))
            mancanti = sorted(r for r in riferimenti if r not in chiavi)
            self.assertEqual(mancanti, [], "%s richiama stringhe assenti" % nome)

    def test_combovista_rimossa_da_dialogviste_n(self):
        xdl = _leggi(os.path.join(DIALOGHI, "Dialogviste_N.xdl"))
        self.assertNotIn("ComboVISTA", xdl)
        self.assertFalse(
            [c for c in _chiavi_properties() if "Dialogviste_N.ComboVISTA" in c],
            "stringhe di ComboVISTA ancora presenti per Dialogviste_N",
        )

    def test_controlli_usati_da_contabilita_esistono(self):
        sorgente = _leggi(PYLEENO)
        inizio = sorgente.index("UltimusFree2.Dialogviste_N?language=Basic")
        fine = sorgente.index("#Elenco Prezzi", inizio)
        usati = set(re.findall(r"getControl\('(\w+)'\)", sorgente[inizio:fine]))
        xdl = _leggi(os.path.join(DIALOGHI, "Dialogviste_N.xdl"))
        presenti = set(re.findall(r'dlg:id="(\w+)"', xdl))
        self.assertEqual(sorted(usati - presenti - CONTROLLI_FACOLTATIVI), [])


if __name__ == "__main__":
    unittest.main()
