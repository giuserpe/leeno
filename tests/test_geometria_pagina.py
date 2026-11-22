# -*- coding: utf-8 -*-
"""
Test della geometria di pagina centralizzata (LeenoSettings).

Tre livelli:
- unitari, senza UNO: valori scritti, ordine dei nomi, ripiego per singola proprieta';
- guardia anti-duplicazione: i margini di stampa sono definiti in un solo punto;
- documento reale (opzionale): confronta il comportamento precedente (assegnazione
  una per una, come faceva MENU_sistema_pagine) con quello nuovo su tutti gli stili
  di pagina di tests/data/LeenO_Esempio.ods. Gira solo con LEENO_TEST_UNO=1 e
  LibreOffice installato; usa una copia temporanea del file. Serve il Python di
  LibreOffice (ha il modulo uno), ad esempio su Windows:
    $env:LEENO_TEST_UNO = "1"
    & "C:/Program Files/LibreOffice/program/python.exe" -m unittest tests.test_geometria_pagina -v
"""
import ast
import os
import re
import shutil
import socket
import subprocess
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

ROOT = Path(__file__).resolve().parent.parent
PYTHONPATH_DIR = ROOT / "src/Ultimus.oxt/python/pythonpath"
SETTINGS = PYTHONPATH_DIR / "LeenoSettings.py"
ESEMPIO = ROOT / "tests/data/LeenO_Esempio.ods"

FUNZIONI = ("_imposta_proprieta", "set_page_margins", "set_page_scale", "set_page_borders")
COSTANTI = ("_PAGINA_MARGINI", "_PAGINA_SCALA")

MARGINI_ATTESI = {
    "TopMargin": 1000, "BottomMargin": 1350, "LeftMargin": 1000, "RightMargin": 1000,
    "FooterLeftMargin": 0, "FooterRightMargin": 0, "HeaderLeftMargin": 0, "HeaderRightMargin": 0,
    "HeaderBodyDistance": 0, "FooterBodyDistance": 0,
}
SCALA_ATTESA = {"PageScale": 0, "CenterHorizontally": True, "ScaleToPagesX": 1, "ScaleToPagesY": 0}


def _carica(uno):
    """Compila costanti e funzioni di geometria da LeenoSettings senza importare il modulo."""
    tree = ast.parse(SETTINGS.read_text(encoding="utf-8"))
    nodi = []
    for n in tree.body:
        if isinstance(n, ast.FunctionDef) and n.name in FUNZIONI:
            nodi.append(n)
        elif isinstance(n, ast.Assign) and getattr(n.targets[0], "id", None) in COSTANTI:
            nodi.append(n)
    ns = {"uno": uno}
    exec(compile(ast.Module(body=nodi, type_ignores=[]), str(SETTINGS), "exec"), ns)
    return ns


class PaginaFinta:
    """Stile di pagina che registra le scritture."""

    def __init__(self, multi=True):
        self.chiamate_multiple = []
        self._multi = multi

    def setPropertyValues(self, nomi, valori):
        if not self._multi:
            raise RuntimeError("XMultiPropertySet non supportato")
        self.chiamate_multiple.append(dict(zip(nomi, valori)))
        self.chiamate_nomi = nomi


def _uno_finto():
    uno = MagicMock()
    uno.createUnoStruct.side_effect = lambda _nome: SimpleNamespace(LineWidth=7, OuterLineWidth=7)
    return uno


class TestGeometriaUnitari(unittest.TestCase):
    def setUp(self):
        self.ns = _carica(_uno_finto())

    def test_margini_e_scala_in_una_sola_chiamata(self):
        p = PaginaFinta()
        self.ns["set_page_margins"](p)
        self.assertEqual(len(p.chiamate_multiple), 1)
        atteso = {**MARGINI_ATTESI, **SCALA_ATTESA}
        self.assertEqual(p.chiamate_multiple[0], atteso)

    def test_nomi_ordinati_e_scala_automatica_prevale(self):
        p = PaginaFinta()
        self.ns["set_page_margins"](p)
        nomi = list(p.chiamate_nomi)
        self.assertEqual(nomi, sorted(nomi))
        # PageScale deve precedere ScaleToPagesX/Y: altrimenti la scala fissa annulla l'adattamento
        self.assertLess(nomi.index("PageScale"), nomi.index("ScaleToPagesX"))
        self.assertLess(nomi.index("PageScale"), nomi.index("ScaleToPagesY"))

    def test_solo_scala_non_tocca_i_margini(self):
        p = PaginaFinta()
        self.ns["set_page_scale"](p)
        self.assertEqual(p.chiamate_multiple, [SCALA_ATTESA])

    def test_bordi_azzerati_su_tutti_e_quattro_i_lati(self):
        p = PaginaFinta()
        self.ns["set_page_borders"](p)
        (chiamata,) = p.chiamate_multiple
        self.assertEqual(set(chiamata), {"TopBorder", "BottomBorder", "RightBorder", "LeftBorder"})
        for bordo in chiamata.values():
            self.assertEqual((bordo.LineWidth, bordo.OuterLineWidth), (0, 0))

    def test_ripiego_se_la_scrittura_multipla_non_e_supportata(self):
        p = PaginaFinta(multi=False)
        self.ns["set_page_margins"](p)
        for nome, valore in {**MARGINI_ATTESI, **SCALA_ATTESA}.items():
            self.assertEqual(getattr(p, nome), valore, nome)

    def test_dizionario_costanti_non_modificato_dalle_chiamate(self):
        copia = (dict(self.ns["_PAGINA_MARGINI"]), dict(self.ns["_PAGINA_SCALA"]))
        self.ns["set_page_margins"](PaginaFinta())
        self.assertEqual((self.ns["_PAGINA_MARGINI"], self.ns["_PAGINA_SCALA"]), copia)


class TestNessunaDuplicazione(unittest.TestCase):
    def test_margine_inferiore_definito_in_un_solo_punto(self):
        trovati = []
        for f in PYTHONPATH_DIR.glob("*.py"):
            for riga in f.read_text(encoding="utf-8").splitlines():
                if re.search(r"\b1350\b", riga) and "Margin" in riga:
                    trovati.append(f.name)
        self.assertEqual(trovati, ["LeenoSettings.py"], trovati)

    def test_sistema_pagine_usa_gli_helper(self):
        src = (PYTHONPATH_DIR / "pyleeno.py").read_text(encoding="utf-8")
        inizio = src.index("def MENU_sistema_pagine")
        corpo = src[inizio:inizio + 12000]
        self.assertIn("LS.set_page_margins(", corpo)
        self.assertIn("LS.set_page_borders(", corpo)
        self.assertNotIn("PROPS_COMUNI", corpo)


# --------------------------------------------------------------------------------------
# Documento reale
# --------------------------------------------------------------------------------------
HOST = "127.0.0.1"
PORTA = int(os.environ.get("LEENO_UNO_PORTA", "2002"))


def _trova_soffice():
    """Percorso di soffice: dal PATH o, su Windows, dalle cartelle di installazione standard."""
    if os.name == "nt":
        for var in ("ProgramFiles", "ProgramFiles(x86)"):
            base = os.environ.get(var)
            if base:
                candidato = Path(base) / "LibreOffice" / "program" / "soffice.exe"
                if candidato.exists():
                    return str(candidato)
    return shutil.which("soffice")


def _uno_disponibile():
    if os.environ.get("LEENO_TEST_UNO") != "1" or not ESEMPIO.exists():
        return False
    try:
        import uno  # noqa: F401
    except ImportError:
        return False
    return _trova_soffice() is not None


def _porta_aperta():
    with socket.socket() as s:
        s.settimeout(0.5)
        return s.connect_ex((HOST, PORTA)) == 0


def _connetti(attesa=120):
    """Attende che la porta UNO sia aperta e si collega; l'avvio con profilo nuovo e' lento."""
    import uno
    ctx = uno.getComponentContext()
    resolver = ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx)
    ultimo = None
    fine = time.time() + attesa
    while time.time() < fine:
        if _porta_aperta():
            try:
                return resolver.resolve(
                    "uno:socket,host=%s,port=%d;urp;StarOffice.ComponentContext" % (HOST, PORTA))
            except Exception as exc:
                ultimo = exc
        time.sleep(1)
    raise RuntimeError("LibreOffice non raggiungibile su %s:%d entro %d s (porta aperta: %s, ultimo errore: %s)"
                       % (HOST, PORTA, attesa, _porta_aperta(), ultimo))


PROPRIETA_LETTE = tuple(MARGINI_ATTESI) + tuple(SCALA_ATTESA) + (
    "Width", "Height", "IsLandscape", "HeaderIsOn", "FooterIsOn", "HeaderHeight", "FooterHeight")


def _bordi(p):
    return tuple((getattr(p, n).LineWidth, getattr(p, n).OuterLineWidth)
                 for n in ("TopBorder", "BottomBorder", "RightBorder", "LeftBorder"))


def _applica_riferimento(uno, p):
    """Comportamento precedente di MENU_sistema_pagine: una proprieta' alla volta."""
    for nome, valore in {**MARGINI_ATTESI, **SCALA_ATTESA}.items():
        setattr(p, nome, valore)
    bordo = uno.createUnoStruct("com.sun.star.table.BorderLine2")
    bordo.LineWidth = 0
    bordo.OuterLineWidth = 0
    p.TopBorder = bordo
    p.BottomBorder = bordo
    p.RightBorder = bordo
    p.LeftBorder = bordo


@unittest.skipUnless(_uno_disponibile(), "richiede LEENO_TEST_UNO=1, uno, soffice e LeenO_Esempio.ods")
class TestDocumentoReale(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import uno
        cls.uno = uno
        cls.processo = None
        cls.tmp = tempfile.mkdtemp(prefix="leeno_geom_")
        cls.addClassCleanup(shutil.rmtree, cls.tmp, ignore_errors=True)
        cls.log = os.path.join(cls.tmp, "soffice.log")
        if not _porta_aperta():
            extra = {"start_new_session": True} if os.name == "posix" else {}
            with open(cls.log, "w") as log:
                cls.processo = subprocess.Popen(
                    [_trova_soffice(), "--headless", "--invisible", "--norestore", "--nologo",
                     "-env:UserInstallation=" + (Path(cls.tmp) / "profilo").as_uri(),
                     "--accept=socket,host=%s,port=%d;urp;" % (HOST, PORTA)],
                    stdout=log, stderr=log, **extra)
            cls.addClassCleanup(cls._chiudi_processo)
        try:
            ctx = _connetti()
        except RuntimeError as exc:
            dettagli = ""
            if cls.processo is not None:
                dettagli = " [codice uscita avvio: %s]" % cls.processo.poll()
                try:
                    dettagli += " log: " + Path(cls.log).read_text(errors="replace")[-500:]
                except OSError:
                    pass
            raise RuntimeError(str(exc) + dettagli)
        cls.desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        cls.docs = []
        for tag in ("nuovo", "riferimento"):
            copia = os.path.join(cls.tmp, tag + ".ods")
            shutil.copy(ESEMPIO, copia)
            doc = cls.desktop.loadComponentFromURL(
                uno.systemPathToFileUrl(copia), "_blank", 0,
                (uno.createUnoStruct("com.sun.star.beans.PropertyValue", "Hidden", 0, True, 0),))
            cls.docs.append(doc)
        cls.ns = _carica(uno)

    @classmethod
    def _chiudi_processo(cls):
        try:
            cls.desktop.terminate()
        except Exception:
            pass
        try:
            cls.processo.terminate()
        except Exception:
            pass

    @classmethod
    def tearDownClass(cls):
        for d in cls.docs:
            try:
                d.close(True)
            except Exception:
                pass

    def _stili(self, doc):
        return doc.StyleFamilies.getByName("PageStyles")

    def test_stessi_valori_di_prima_su_tutti_gli_stili(self):
        nuovo, rif = self.docs
        sn, sr = self._stili(nuovo), self._stili(rif)
        self.assertGreater(sn.Count, 5)
        self.assertEqual(sn.Count, sr.Count)
        for i in range(sn.Count):
            pn, pr = sn.getByIndex(i), sr.getByIndex(i)
            self.ns["set_page_margins"](pn)
            self.ns["set_page_borders"](pn)
            _applica_riferimento(self.uno, pr)
            nome = pn.DisplayName
            for prop in PROPRIETA_LETTE:
                self.assertEqual(getattr(pn, prop), getattr(pr, prop), "%s.%s" % (nome, prop))
            self.assertEqual(_bordi(pn), _bordi(pr), nome)

    def test_valori_attesi_e_stili_contabili_presenti(self):
        nuovo = self.docs[0]
        stili = self._stili(nuovo)
        for nome in ("PageStyle_COMPUTO_A4", "PageStyle_Elenco Prezzi"):
            self.assertTrue(stili.hasByName(nome), nome)
        p = stili.getByName("PageStyle_COMPUTO_A4")
        self.ns["set_page_margins"](p)
        self.ns["set_page_borders"](p)
        # LibreOffice converte le misure in unita' interne: 1350 viene riletto come 1349.
        # PageScale non e' confrontato: con "adatta alla larghezza" attiva viene riletto come 100.
        for prop, valore in {**MARGINI_ATTESI, **SCALA_ATTESA}.items():
            if prop != "PageScale":
                self.assertAlmostEqual(getattr(p, prop), valore, delta=1, msg=prop)
        self.assertEqual(_bordi(p), ((0, 0),) * 4)

    def test_set_page_scale_non_modifica_i_margini(self):
        nuovo = self.docs[0]
        p = self._stili(nuovo).getByName("PageStyle_COMPUTO_A4")
        p.TopMargin = 2222
        self.ns["set_page_scale"](p)
        self.assertAlmostEqual(p.TopMargin, 2222, delta=1)
        self.assertEqual((p.ScaleToPagesX, p.ScaleToPagesY, p.CenterHorizontally), (1, 0, True))


if __name__ == "__main__":
    unittest.main()
