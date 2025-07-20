#!/usr/bin/env python3
"""
Script per generare automaticamente la versione del progetto Leeno.
Legge da un file VERSION e genera un file version.h per C/C++.
"""

import os
import re
import requests
from datetime import datetime

def read_current_version():
    """Legge la versione corrente dal file VERSION."""
    try:
        with open("VERSION", "r") as f:
            version = f.read().strip()
            # Verifica il formato semantico (es. 1.2.3)
            if not re.match(r"^\d+\.\d+\.\d+$", version):
                raise ValueError(f"Formato versione non valido: {version}")
            return version
    except FileNotFoundError:
        print("ERRORE: File VERSION non trovato nella directory corrente")
        print("Creane uno con il formato MAJOR.MINOR.PATCH (es. 1.0.0)")
        raise

def generate_build_number():
    """Genera un numero di build univoco basato sulla data."""
    now = datetime.utcnow()
    return now.strftime("%Y%m%d%H%M")

def generate_version_header(version, build_number):
    """Genera il file version.h per progetti C/C++."""
    header = f"""
#ifndef LEENO_VERSION_H
#define LEENO_VERSION_H

#define LEENO_VERSION_MAJOR {version.split('.')[0]}
#define LEENO_VERSION_MINOR {version.split('.')[1]}
#define LEENO_VERSION_PATCH {version.split('.')[2]}
#define LEENO_VERSION_STRING "{version}"
#define LEENO_BUILD_NUMBER "{build_number}"
#define LEENO_BUILD_DATE __DATE__
#define LEENO_BUILD_TIME __TIME__

#endif // LEENO_VERSION_H
"""
    return header.strip()

def write_version_header(content, output_dir="include"):
    """Scrive il file version.h nella directory specificata."""
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "version.h")
    
    with open(output_path, "w") as f:
        f.write(content)
    
    print(f"File generato: {output_path}")

def main():
    try:
        # Leggi la versione corrente
        version = read_current_version()
        print(f"Versione corrente: {version}")
        
        # Genera numero di build
        build_number = generate_build_number()
        print(f"Numero di build: {build_number}")
        
        # Genera il file header
        header_content = generate_version_header(version, build_number)
        write_version_header(header_content)
        
        # Se necessario, puoi aggiungere qui la logica per aggiornare il changelog
        # o altre operazioni correlate alla gestione delle versioni
        
    except Exception as e:
        print(f"ERRORE durante la generazione della versione: {str(e)}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())