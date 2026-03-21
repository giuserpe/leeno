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
        oxt_list = []
        # base_url es: https://dev.leeno.org/index.php/s/jLnxqWRzSD7MqFB#
        base_url = (os.getenv('PUBLIC_DOWNLOAD_URL') or os.getenv('OXT_BASE_URL', '')).rstrip('#').rstrip('/')

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
                        name = parts[-1]
                        size = parts[-2]
                        date = ' '.join(parts[:3])
                        url = f"{base_url}/download?path=&files={name}" if base_url else '#'
                        oxt_list.append({
                            'name': name,
                            'size': size,
                            'date': date,
                            'url': url,
                        })
                    else:
                        logger.warning(f"Riga non parsabile in oxt_list: {line!r}")
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
        """Genera la pagina HTML con le ultime 5 versioni"""
        oxt_files = self._parse_oxt_list()
        now_utc = datetime.utcnow().strftime('%Y-%m-%d %H:%M')
        base_url = os.getenv('PUBLIC_DOWNLOAD_URL') or os.getenv('OXT_BASE_URL', '')

        rows = []
        for i, file in enumerate(oxt_files):
            name = file['name']
            badge = '<span class="badge badge-latest">ULTIMA</span>' if i == 0 else ''
            url = file.get('url', '#')
            sha256 = file.get('sha256', '')
            rows.append(f"""
            <tr>
                <td>{name} {badge}</td>
                <td><a href="{url}" download>Scarica</a></td>
                <td class="hash">{sha256}</td>
                <td>{file['date']}</td>
                <td>{file['size']}</td>
            </tr>""")

        html = f"""<!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="utf-8">
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
        .hash {{
            font-family: monospace;
            font-size: 0.85em;
            word-break: break-all;
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
        <p><strong>Ultima versione:</strong> {version_info['full']}</p>
    </div>

    <h2>Download disponibili</h2>
    <table>
        <thead>
            <tr>
                <th>Versione</th>
                <th>Download</th>
                <th>SHA256</th>
                <th>Data</th>
                <th>Dimensione</th>
            </tr>
        </thead>
        <tbody>
            {"".join(rows)}
        </tbody>
    </table>

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
