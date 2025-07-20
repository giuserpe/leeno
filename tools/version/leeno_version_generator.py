#!/usr/bin/env python3
"""
Script per la gestione delle versioni LeenO con generazione di versions.html
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
        
        # Debug: stampa i percorsi
        logger.debug(f"Repo root: {self.repo_root}")
        logger.debug(f"Version file: {self.version_file}")
        logger.debug(f"Include dir: {self.include_dir}")
        logger.debug(f"Tools version dir: {self.tools_version_dir}")

        # Verifica e crea directory
        self._ensure_directories_exist()

    def _ensure_directories_exist(self):
        """Crea le directory necessarie se non esistono"""
        try:
            self.include_dir.mkdir(exist_ok=True)
            self.tools_version_dir.mkdir(parents=True, exist_ok=True)
            logger.debug("Directory verificate/creata con successo")
        except Exception as e:
            logger.error(f"Errore nella creazione delle directory: {str(e)}")
            raise

    def parse_current_version(self) -> Dict[str, str]:
        try:
            with open(self.version_file, 'r') as f:
                version_str = f.read().strip()
                logger.info(f"Versione corrente: {version_str}")
                
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
            logger.error(f"Errore nel parsing della versione: {str(e)}")
            raise

    def generate_new_version(self, current: Dict[str, str], version_type: str = None) -> Dict[str, str]:
        now = datetime.utcnow()
        build_number = os.getenv('BUILD_NUMBER', current['build'])
        git_sha = os.getenv('GITHUB_SHA', 'local')[:7]
        
        version_type = version_type if version_type else current['type']
        
        return {
            'full': f"LeenO-{current['major']}.{current['minor']}.{current['patch']}.{build_number}-{version_type}-{now.strftime('%Y%m%d')}",
            'major': current['major'],
            'minor': current['minor'],
            'patch': current['patch'],
            'build_number': build_number,
            'build_date': now.strftime("%Y-%m-%d"),
            'build_time': now.strftime("%H:%M:%S"),
            'git_sha': git_sha,
            'type': version_type,
            'semver': current['semver']
        }

    def _generate_version_header(self, version_info: Dict[str, str], build_info: Dict[str, str]):
        header_file = self.include_dir / 'version.h'
        with open(header_file, 'w') as f:
            f.write(f"""// Auto-generated
#ifndef LEENO_VERSION_H
#define LEENO_VERSION_H
#define LEENO_VERSION_FULL "{build_info['full']}"
#endif
""")
        logger.info(f"Generato {header_file}")

    def _generate_versions_html(self, version_info: Dict[str, str], build_info: Dict[str, str]):
        html_file = self.tools_version_dir / 'versions.html'
        
        content = f"""<!DOCTYPE html>
<html>
<head>
    <title>Versioni LeenO</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        h1 {{ color: #2c3e50; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }}
    </style>
</head>
<body>
    <h1>Versione {version_info['full']}</h1>
    <table>
        <tr><th>Componente</th><th>Valore</th></tr>
        <tr><td>Versione completa</td><td>{version_info['full']}</td></tr>
        <tr><td>Tipo build</td><td>{build_info['type']}</td></tr>
        <tr><td>Data build</td><td>{build_info['build_date']}</td></tr>
    </table>
</body>
</html>
"""
        try:
            with open(html_file, 'w') as f:
                f.write(content)
            logger.info(f"Generato {html_file}")
            logger.debug(f"Contenuto di {html_file}:\n{content[:200]}...")  # Log parziale del contenuto
        except Exception as e:
            logger.error(f"Errore nella generazione di {html_file}: {str(e)}")
            raise

    def update_version_files(self, current: Dict[str, str], new: Dict[str, str]):
        try:
            # Aggiorna file versione principale
            with open(self.version_file, 'w') as f:
                f.write(new['full'])
            
            # Genera gli altri file
            self._generate_version_header(current, new)
            self._generate_versions_html(current, new)
            
            logger.info("Tutti i file generati con successo")
            
            # Verifica che i file esistano
            assert (self.tools_version_dir / 'versions.html').exists(), "versions.html non creato!"
            assert (self.include_dir / 'version.h').exists(), "version.h non creato!"
            
        except Exception as e:
            logger.error(f"Errore nell'aggiornamento dei file: {str(e)}")
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
        
        logger.info("Operazione completata con successo")
        return 0
    except Exception as e:
        logger.critical(f"Errore: {str(e)}", exc_info=args.debug)
        return 1

if __name__ == "__main__":
    exit(main())