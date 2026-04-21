#!/usr/bin/env python3
"""
Script completo per la gestione delle versioni LeenO con archivio .oxt
"""
import json
import os
import re
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, List

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

        self.include_dir.mkdir(exist_ok=True)
        self.web_dir.mkdir(exist_ok=True)

    def _parse_oxt_list(self) -> List[Dict[str, str]]:
        """
        Legge oxt_list.txt generato da parse_webdav.py.
        Formato riga: "2026-03-20 18:30 4.4MB LeenO-xxx.oxt"
          parts[0] = data, parts[1] = ora, parts[2] = size, parts[3] = nome

        PUBLIC_DOWNLOAD_URL è già nella forma:
          https://dev.leeno.org/index.php/s/TOKEN/download?path=&files=
        quindi l'URL finale è semplicemente base_url + nome_file.
        """
        oxt_list = []
        raw_url = (os.getenv('PUBLIC_DOWNLOAD_URL') or os.getenv('OXT_BASE_URL', '')).rstrip('#')
        # Normalizza: rimuovi trailing slash solo se NON termina con '='
        # (se termina con '=' è già pronto per appendere il nome file)
        if raw_url.endswith('='):
            base_url = raw_url
        else:
            base_url = raw_url.rstrip('/')

        try:
            oxt_list_path = os.getenv('OXT_LIST_PATH', '')
            if not oxt_list_path:
                raise ValueError("OXT_LIST_PATH non impostato")
            with open(oxt_list_path, 'r') as f:
                for line in f:
                    line = line.strip()
                    if not line or '.oxt' not in line.lower():
                        continue
                    parts = line.split()
                    if len(parts) >= 4:
                        date = f"{parts[0]} {parts[1]}"
                        size = parts[2]
                        name = parts[3]
                    elif len(parts) == 3:
                        date = parts[0]
                        size = parts[1]
                        name = parts[2]
                    else:
                        logger.warning(f"Riga non parsabile: {line!r}")
                        continue
                    # URL: se base_url termina con '=' appendi direttamente il nome
                    # altrimenti aggiungi /download?path=&files=
                    if base_url.endswith('='):
                        url = f"{base_url}{name}"
                    elif base_url:
                        url = f"{base_url}/download?path=&files={name}"
                    else:
                        url = '#'
                    oxt_list.append({
                        'name': name,
                        'size': size,
                        'date': date,
                        'url': url,
                    })
        except Exception as e:
            logger.error(f"Errore lettura lista file: {str(e)}")

        if oxt_list:
            logger.info(f"Trovati {len(oxt_list)} file .oxt, uso i primi 5")
            return oxt_list[:5]

        logger.warning("Nessun file .oxt trovato nella lista")
        return [{
            'name': 'Nessun file disponibile',
            'size': '0KB',
            'date': datetime.now().strftime('%Y-%m-%d'),
            'url': '#'
        }]

    def _parse_commits(self) -> List[Dict[str, str]]:
        """Legge commits.json generato da parse_commits.py."""
        commits_path = os.getenv('COMMITS_PATH', 'commits.json')
        try:
            with open(commits_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.warning(f"Impossibile leggere commits.json: {e}")
            return []

    def update_version_files(self, version_info: Dict[str, str]):
        """Genera tutti i file necessari"""
        try:
            with open(self.version_file, 'w') as f:
                f.write(version_info['full'])

            self._generate_version_header(version_info)
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
        """Genera la pagina HTML con le ultime 5 versioni e gli ultimi commit"""
        oxt_files = self._parse_oxt_list()
        commits = self._parse_commits()
        now_utc = datetime.utcnow().strftime('%Y-%m-%d %H:%M')
        base_url = (os.getenv('PUBLIC_DOWNLOAD_URL') or os.getenv('OXT_BASE_URL', '')).rstrip('#').rstrip('/')

        # Righe tabella download
        rows = []
        for i, file in enumerate(oxt_files):
            name = file['name']
            badge = '<span class="badge badge-latest">ULTIMA</span>' if i == 0 else ''
            url = file.get('url', '#')
            rows.append(f"""
            <tr>
                <td>{name} {badge}</td>
                <td><a href="{url}" download>Scarica</a></td>
                <td>{file['date']}</td>
                <td>{file['size']}</td>
            </tr>""")

        # Sezione commit
        commit_rows = []
        for c in commits:
            commit_rows.append(f"""
            <tr>
                <td>{c['date']}</td>
                <td>{c.get('author', 'N/A')}</td>
                <td><a href="{c['url']}" target="_blank" rel="noopener"><code>{c['sha']}</code></a></td>
                <td>{c['msg']}</td>
            </tr>""")

        commits_section = ''
        if commit_rows:
            commits_section = f"""
    <h2>Attività di sviluppo recente</h2>
    <table>
        <thead>
            <tr>
                <th style="width:140px">Data</th>
                <th style="width:120px">Autore</th>
                <th style="width:80px">Commit</th>
                <th>Descrizione</th>
            </tr>
        </thead>
        <tbody>
            {"".join(commit_rows)}
        </tbody>
    </table>"""
        else:
            commits_section = """
    <h2>Attività di sviluppo recente</h2>
    <p><i>Nessun commit recente trovato o errore nel recupero dei dati.</i></p>"""

        html = f"""<!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="utf-8">
    <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Expires" content="0">
    <title>Versioni Nightly Builds LeenO</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 20px;
            color: #333;
            max-width: 1200px;
            margin: auto;
        }}
        h1, h2 {{
            color: #2c3e50;
            border-bottom: 2px solid #3498db;
            padding-bottom: 10px;
        }}
        .info-box {{
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
            border-left: 4px solid #3498db;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
        }}
        th, td {{
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }}
        th {{
            background-color: #3498db;
            color: white;
        }}
        tr:nth-child(even) {{ background-color: #f2f2f2; }}
        tr:hover {{ background-color: #e9f7fe; }}
        code {{
            font-family: monospace;
            font-size: 0.9em;
            background: #eef;
            padding: 1px 5px;
            border-radius: 3px;
        }}
        .badge {{
            display: inline-block;
            padding: 3px 7px;
            border-radius: 3px;
            font-size: 0.8em;
            font-weight: bold;
            color: white;
        }}
        .badge-latest {{ background-color: #2ecc71; }}
        .footer {{
            margin-top: 30px;
            font-size: 0.9em;
            color: #7f8c8d;
            text-align: center;
        }}
        a {{ color: #0066cc; text-decoration: none; }}
        a:hover {{ text-decoration: underline; }}
        @media (max-width: 768px) {{
            th, td {{ padding: 8px; }}
        }}
    </style>
</head>
<body>
    <h1>Nightly Builds LeenO</h1>

    <div class="info-box">
        <h2>Informazioni</h2>
        <p>Questa pagina elenca le ultime 5 versioni di sviluppo disponibili sul server.</p>
        <p><strong>Build Commit:</strong> <code>{version_info['git_sha']}</code></p>
    </div>

    <h2>Download disponibili</h2>
    <table>
        <thead>
            <tr>
                <th>Versione</th>
                <th>Download</th>
                <th>Data</th>
                <th>Dimensione</th>
            </tr>
        </thead>
        <tbody>
            {"".join(rows)}
        </tbody>
    </table>
    {commits_section}
    <div class="footer">
        <p>Generato automaticamente il {now_utc} UTC</p>
        <p>Server: {base_url}</p>
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

        with open(vm.version_file, 'r') as f:
            current_version = f.read().strip()

        match = vm.VERSION_PATTERN.match(current_version)
        if not match:
            raise ValueError(f"Formato versione non valido: {current_version}")

        new_version = {
            'full': f"LeenO-{match.group('major')}.{match.group('minor')}.{match.group('patch')}.{os.getenv('BUILD_NUMBER', match.group('build'))}-{match.group('type')}-{datetime.now().strftime('%Y%m%d')}",
            'major': match.group('major'),
            'minor': match.group('minor'),
            'patch': match.group('patch'),
            'build_number': os.getenv('BUILD_NUMBER', match.group('build')),
            'build_date': datetime.now().strftime('%Y-%m-%d'),
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
