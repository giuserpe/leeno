#!/usr/bin/env python3
"""
Script completo per la gestione delle versioni LeenO
Genera version.h, versions.html e aggiorna leeno_version_code
"""

import os
import re
import logging
import argparse
from datetime import datetime
from pathlib import Path
from typing import Dict

# Configurazione logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

class VersionManager:
    """Gestisce tutte le operazioni di versioning"""
    
    VERSION_PATTERN = re.compile(
        r'^LeenO-'
        r'(?P<major>\d+)\.(?P<minor>\d+)\.(?P<patch>\d+)\.(?P<build>\d+)-'
        r'(?P<type>STABLE|TESTING)-'
        r'(?P<date>\d{8})$'
    )
    
    def __init__(self, repo_root: Path):
        self.repo_root = repo_root
        self.version_file = repo_root / 'src' / 'Ultimus.oxt' / 'leeno_version_code'
        self.include_dir = repo_root / 'include'
        self.tools_version_dir = repo_root / 'tools' / 'version'
        
        # Crea directory se non esistono
        self.include_dir.mkdir(exist_ok=True)
        self.tools_version_dir.mkdir(exist_ok=True)

    def parse_current_version(self) -> Dict[str, str]:
        """Legge la versione corrente"""
        try:
            with open(self.version_file, 'r') as f:
                version_str = f.read().strip()
                match = self.VERSION_PATTERN.match(version_str)
                if not match:
                    raise ValueError(f"Formato versione non valido: {version_str}")
                
                return {
                    'full': version_str,
                    'major': match.group('major'),
                    'minor': match.group('minor'),
                    'patch': match.group('patch'),
                    'build': match.group('build'),
                    'type': match.group('type'),
                    'date': match.group('date'),
                    'semver': f"{match.group('major')}.{match.group('minor')}.{match.group('patch')}"
                }
        except Exception as e:
            logger.error(f"Errore parsing versione: {str(e)}")
            raise

    def generate_new_version(self, current: Dict[str, str]) -> Dict[str, str]:
        """Genera una nuova versione"""
        now = datetime.utcnow()
        return {
            'full': f"LeenO-{current['major']}.{current['minor']}.{current['patch']}.{os.getenv('BUILD_NUMBER', current['build'])}-{current['type']}-{now.strftime('%Y%m%d')}",
            'major': current['major'],
            'minor': current['minor'],
            'patch': current['patch'],
            'build_number': os.getenv('BUILD_NUMBER', current['build']),
            'build_date': now.strftime("%Y-%m-%d"),
            'build_time': now.strftime("%H:%M:%S"),
            'git_sha': os.getenv('GITHUB_SHA', 'local')[:7],
            'type': current['type'],
            'semver': current['semver']
        }

    def _generate_version_header(self, version_info: Dict[str, str], build_info: Dict[str, str]) -> None:
        """Genera il file C++ version.h"""
        header_content = f"""// Auto-generated
#ifndef LEENO_VERSION_H
#define LEENO_VERSION_H
#define LEENO_VERSION_FULL "{build_info['full']}"
#define LEENO_VERSION_MAJOR {version_info['major']}
#define LEENO_VERSION_MINOR {version_info['minor']}
#define LEENO_VERSION_PATCH {version_info['patch']}
#define LEENO_BUILD_NUMBER "{build_info['build_number']}"
#define LEENO_BUILD_DATE "{build_info['build_date']}"
#define LEENO_GIT_SHA "{build_info['git_sha']}"
#endif
"""
        with open(self.include_dir / 'version.h', 'w') as f:
            f.write(header_content)

    def _generate_versions_html(self, version_info: Dict[str, str], build_info: Dict[str, str]) -> None:
        """Genera la pagina HTML completa"""
        html_content = f"""<!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="UTF-8">
    <title>LeenO {version_info['full']}</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        h1 {{ color: #2c3e50; border-bottom: 1px solid #eee; }}
        .version-info {{ 
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin-top: 20px;
        }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ padding: 8px 12px; text-align: left; border-bottom: 1px solid #ddd; }}
        th {{ background-color: #f0f0f0; }}
        .git-sha {{ font-family: monospace; }}
    </style>
</head>
<body>
    <h1>LeenO {version_info['semver']}</h1>
    <div class="version-info">
        <table>
            <tr><th>Versione completa</th><td>{version_info['full']}</td></tr>
            <tr><th>Tipo build</th><td>{build_info['type']}</td></tr>
            <tr><th>Data build</th><td>{build_info['build_date']}</td></tr>
            <tr><th>Commit Git</th><td class="git-sha">{build_info['git_sha']}</td></tr>
        </table>
    </div>
</body>
</html>
"""
        with open(self.tools_version_dir / 'versions.html', 'w', encoding='utf-8') as f:
            f.write(html_content)

    def update_version_files(self, current: Dict[str, str], new: Dict[str, str]) -> None:
        """Aggiorna tutti i file di versione"""
        try:
            # File versione principale
            with open(self.version_file, 'w') as f:
                f.write(new['full'])
            
            # File C++
            self._generate_version_header(current, new)
            
            # Pagina HTML
            self._generate_versions_html(current, new)
            
            logger.info("File di versione aggiornati con successo")
        except Exception as e:
            logger.error(f"Errore aggiornamento file: {str(e)}")
            raise

def main():
    parser = argparse.ArgumentParser(description='Generatore versioni LeenO')
    parser.add_argument('--debug', action='store_true', help='Abilita debug')
    args = parser.parse_args()
    
    if args.debug:
        logger.setLevel(logging.DEBUG)
    
    try:
        logger.info("Avvio generazione versione...")
        repo_root = Path(__file__).parent.parent.parent
        vm = VersionManager(repo_root)
        
        current = vm.parse_current_version()
        new = vm.generate_new_version(current)
        vm.update_version_files(current, new)
        
        logger.info(f"Versione generata: {new['full']}")
    except Exception as e:
        logger.critical(f"Errore: {str(e)}")
        raise SystemExit(1)

if __name__ == "__main__":
    main()