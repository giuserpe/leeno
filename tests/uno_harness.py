# -*- coding: utf-8 -*-
"""
Sessione LibreOffice headless condivisa dai test su documento reale.

Attiva solo con LEENO_TEST_UNO=1, con il modulo uno (su Windows il Python di
LibreOffice) e con soffice raggiungibile. Se la porta UNO e' gia' aperta si collega
a quella istanza, altrimenti ne avvia una con profilo temporaneo e la chiude a fine
prova. Porta configurabile con LEENO_UNO_PORTA (predefinita 2002).
"""
import os
import shutil
import socket
import subprocess
import tempfile
import time
from pathlib import Path

HOST = "127.0.0.1"
PORTA = int(os.environ.get("LEENO_UNO_PORTA", "2002"))


def trova_soffice():
    """Percorso di soffice: su Windows dalle cartelle di installazione, poi dal PATH."""
    if os.name == "nt":
        for var in ("ProgramFiles", "ProgramFiles(x86)"):
            base = os.environ.get(var)
            if base:
                candidato = Path(base) / "LibreOffice" / "program" / "soffice.exe"
                if candidato.exists():
                    return str(candidato)
    return shutil.which("soffice")


def uno_disponibile(*file_richiesti):
    """True se i test su documento reale possono girare."""
    if os.environ.get("LEENO_TEST_UNO") != "1":
        return False
    if any(not Path(f).exists() for f in file_richiesti):
        return False
    try:
        # importati ora, in fase di raccolta: altri test sostituiscono moduli UNO con oggetti finti
        import uno  # noqa: F401
        import unohelper  # noqa: F401
    except ImportError:
        return False
    return trova_soffice() is not None


def porta_aperta():
    with socket.socket() as s:
        s.settimeout(0.5)
        return s.connect_ex((HOST, PORTA)) == 0


def connetti(attesa=120):
    """Attende che la porta UNO sia aperta e si collega; l'avvio con profilo nuovo e' lento."""
    import uno
    ctx = uno.getComponentContext()
    resolver = ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx)
    ultimo = None
    fine = time.time() + attesa
    while time.time() < fine:
        if porta_aperta():
            try:
                return resolver.resolve(
                    "uno:socket,host=%s,port=%d;urp;StarOffice.ComponentContext" % (HOST, PORTA))
            except Exception as exc:
                ultimo = exc
        time.sleep(1)
    raise RuntimeError("LibreOffice non raggiungibile su %s:%d entro %d s (porta aperta: %s, ultimo errore: %s)"
                       % (HOST, PORTA, attesa, porta_aperta(), ultimo))


class SessioneUno:
    """
    Uso in un TestCase:
        cls.sessione = SessioneUno()
        cls.addClassCleanup(cls.sessione.chiudi)   # prima di avvia(), cosi' pulisce anche se fallisce
        cls.sessione.avvia()
        doc = cls.sessione.apri_copia(percorso_ods, "nome")
    """

    def __init__(self):
        import uno
        self.uno = uno
        self.processo = None
        self.desktop = None
        self.docs = []
        self.tmp = tempfile.mkdtemp(prefix="leeno_uno_")
        self.log = os.path.join(self.tmp, "soffice.log")

    def avvia(self):
        if not porta_aperta():
            extra = {"start_new_session": True} if os.name == "posix" else {}
            with open(self.log, "w") as log:
                self.processo = subprocess.Popen(
                    [trova_soffice(), "--headless", "--invisible", "--norestore", "--nologo",
                     "-env:UserInstallation=" + (Path(self.tmp) / "profilo").as_uri(),
                     "--accept=socket,host=%s,port=%d;urp;" % (HOST, PORTA)],
                    stdout=log, stderr=log, **extra)
        try:
            ctx = connetti()
        except RuntimeError as exc:
            dettagli = ""
            if self.processo is not None:
                dettagli = " [codice uscita avvio: %s]" % self.processo.poll()
                try:
                    dettagli += " log: " + Path(self.log).read_text(errors="replace")[-500:]
                except OSError:
                    pass
            raise RuntimeError(str(exc) + dettagli)
        self.desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    def apri_copia(self, origine, tag):
        """Apre (nascosta) una copia temporanea di origine: i file del repo non vengono mai modificati."""
        copia = os.path.join(self.tmp, tag + ".ods")
        shutil.copy(origine, copia)
        uno = self.uno
        doc = self.desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(copia), "_blank", 0,
            (uno.createUnoStruct("com.sun.star.beans.PropertyValue", "Hidden", 0, True, 0),))
        self.docs.append(doc)
        return doc

    def chiudi(self):
        for d in self.docs:
            try:
                d.close(True)
            except Exception:
                pass
        if self.desktop is not None and self.processo is not None:
            try:
                self.desktop.terminate()
            except Exception:
                pass
        if self.processo is not None:
            try:
                self.processo.terminate()
            except Exception:
                pass
        shutil.rmtree(self.tmp, ignore_errors=True)
