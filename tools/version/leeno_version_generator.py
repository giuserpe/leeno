#!/usr/bin/env python3
"""
Script completo per la gestione delle versioni LeenO
"""

import os
import re
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict

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
        """Genera la pagina HTML"""
        html = f"""<!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="UTF-8">
    <title>LeenO {version_info['full']}</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        h1 {{ color: #2c3e50; }}
        .version-info {{ 
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin-top: 20px;
        }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }}
        .git-sha {{ font-family: monospace; }}
    </style>
</head>
<body>
    <h1>Versione LeenO {version_info['full']}</h1>
    <div class="version-info">
        <table>
            <tr><th>Componente</th><th>Valore</th></tr>
            <tr><td>Versione completa</td><td>{version_info['full']}</td></tr>
            <tr><td>Tipo build</td><td>{version_info['type']}</td></tr>
            <tr><td>Data build</td><td>{version_info['build_date']}</td></tr>
            <tr><td>Commit Git</td><td class="git-sha">{version_info['git_sha']}</td></tr>
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
            'build_time': datetime.now().strftime("%H:%M:%S"),
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