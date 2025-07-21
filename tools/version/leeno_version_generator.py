#!/usr/bin/env python3
"""
Script completo per la gestione delle versioni LeenO con archivio .oxt
"""

import os
import re
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, List

# Configurazione logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

class VersionManager:
    VERSION_PATTERN = re.compile(
        r'^LeenO-(?P<major>\d+)\.(?P<minor>\d+)\.(?P<patch>\d+)\.(?P<build>\d+)-(?P<type>STABLE|TESTING)-(?P<date>\d{8})$'
    )
    
    def __init__(self, repo_root: Path):
        self.repo_root = repo_root
        self.version_file = repo_root / 'src' / 'Ultimus.oxt' / 'leeno_version_code'
        self.include_dir = repo_root / 'include'
        self.web_dir = repo_root / 'tools' / 'version'
        
        # Crea directory se non esistono
        self.include_dir.mkdir(exist_ok=True)
        self.web_dir.mkdir(exist_ok=True)

    def _parse_oxt_list(self) -> List[Dict[str, str]]:
        """Parsa l'elenco dei file .oxt con gestione errori migliorata"""
        oxt_list = []
        oxt_file_path = os.getenv('OXT_LIST_PATH', '')
        
        if not oxt_file_path or not Path(oxt_file_path).exists():
            logger.warning("Nessun file lista .oxt trovato")
            return oxt_list
            
        try:
            with open(oxt_file_path, 'r') as f:
                for line in f:
                    line = line.strip()
                    if not line or '.oxt' not in line:
                        continue
                        
                    try:
                        parts = line.split()
                        if len(parts) >= 6:  # Formato minimo atteso
                            filename = parts[-1].split('/')[-1]
                            file_date = f"{parts[5]} {parts[6]}" if len(parts) > 6 else "Data sconosciuta"
                            file_size = f"{int(parts[4])/1024:.1f} KB" if parts[4].isdigit() else "Dimensione sconosciuta"
                            
                            oxt_list.append({
                                "name": filename,
                                "size": file_size,
                                "date": file_date,
                                "url": f"{os.getenv('SFTP_BASE_URL', '')}/{filename}"
                            })
                    except Exception as e:
                        logger.warning(f"Errore processamento linea: {line} - {str(e)}")
                        continue
                        
            logger.info(f"Trovati {len(oxt_list)} file .oxt validi")
        except Exception as e:
            logger.error(f"Errore lettura file lista: {str(e)}")
            
        return oxt_list[:10]

    def update_version_files(self, version_info: Dict[str, str]):
        """Genera tutti i file necessari"""
        try:
            # File versione principale
            with open(self.version_file, 'w') as f:
                f.write(version_info['full'])
            
            # File C++ header
            self._generate_version_header(version_info)
            
            # Pagina HTML
            self._generate_versions_html(version_info)
            
            logger.info("File generati con successo")
        except Exception as e:
            logger.error(f"Errore generazione file: {str(e)}")
            raise

    def _generate_version_header(self, version_info: Dict[str, str]):
        """Genera version.h per C++"""
        content = f"""// Auto-generated
#ifndef LEENO_VERSION_H
#define LEENO_VERSION_H
#define LEENO_VERSION_FULL "{version_info['full']}"
#define LEENO_VERSION_MAJOR {version_info['major']}
#define LEENO_VERSION_MINOR {version_info['minor']}
#define LEENO_VERSION_PATCH {version_info['patch']}
#define LEENO_BUILD_NUMBER "{version_info['build_number']}"
#define LEENO_BUILD_DATE "{version_info['build_date']}"
#define LEENO_GIT_SHA "{version_info['git_sha']}"
#endif
"""
        with open(self.include_dir / 'version.h', 'w') as f:
            f.write(content)

    def _generate_versions_html(self, version_info: Dict[str, str]):
        """Genera la pagina HTML completa"""
        oxt_files = self._parse_oxt_list()
        
        html = f"""<!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="UTF-8">
    <title>LeenO {version_info['full']} - Archivio Versioni</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        h1, h2 {{ color: #2c3e50; }}
        .version-info {{ 
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin: 20px 0;
        }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ padding: 8px 12px; text-align: left; border-bottom: 1px solid #ddd; }}
        th {{ background-color: #e9ecef; }}
        .git-sha {{ font-family: monospace; }}
        a {{ color: #0066cc; text-decoration: none; }}
        a:hover {{ text-decoration: underline; }}
        .file-list {{ margin-top: 30px; }}
    </style>
</head>
<body>
    <h1>Versione LeenO {version_info['full']}</h1>
    
    <div class="version-info">
        <h2>Informazioni Build</h2>
        <table>
            <tr><th>Componente</th><th>Valore</th></tr>
            <tr><td>Versione completa</td><td>{version_info['full']}</td></tr>
            <tr><td>Tipo build</td><td>{version_info['type']}</td></tr>
            <tr><td>Data build</td><td>{version_info['build_date']}</td></tr>
            <tr><td>Commit Git</td><td class="git-sha">{version_info['git_sha']}</td></tr>
        </table>
    </div>

    <div class="version-info file-list">
        <h2>Archivio Versioni (.oxt)</h2>
        <table>
            <tr><th>Nome File</th><th>Dimensione</th><th>Ultima Modifica</th></tr>
            {"".join(
                f'<tr><td><a href="{file["name"]}">{file["name"]}</a></td>'
                f'<td>{file["size"]}</td>'
                f'<td>{file["date"]}</td></tr>'
                for file in oxt_files
            )}
        </table>
    </div>
</body>
</html>
"""
        with open(self.web_dir / 'versions.html', 'w', encoding='utf-8') as f:
            f.write(html)

def main():
    try:
        logger.info("Avvio generazione versione...")
        repo_root = Path(__file__).parent.parent.parent
        vm = VersionManager(repo_root)
        
        # Leggi versione corrente
        with open(vm.version_file, 'r') as f:
            current_version = f.read().strip()
        
        match = vm.VERSION_PATTERN.match(current_version)
        if not match:
            raise ValueError(f"Formato versione non valido: {current_version}")
        
        # Prepara nuova versione
        new_version = {
            'full': f"LeenO-{match.group('major')}.{match.group('minor')}.{match.group('patch')}.{os.getenv('BUILD_NUMBER', match.group('build'))}-{match.group('type')}-{datetime.now().strftime('%Y%m%d')}",
            'major': match.group('major'),
            'minor': match.group('minor'),
            'patch': match.group('patch'),
            'build_number': os.getenv('BUILD_NUMBER', match.group('build')),
            'build_date': datetime.now().strftime("%Y-%m-%d"),
            'git_sha': os.getenv('GITHUB_SHA', 'local')[:7],
            'type': match.group('type')
        }
        
        vm.update_version_files(new_version)
        logger.info(f"Versione generata: {new_version['full']}")
        
    except Exception as e:
        logger.critical(f"Errore: {str(e)}")
        raise SystemExit(1)

if __name__ == "__main__":
    main()