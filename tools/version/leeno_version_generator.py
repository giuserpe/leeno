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
        """
        oxt_list = []
        raw_url = (os.getenv('PUBLIC_DOWNLOAD_URL') or os.getenv('OXT_BASE_URL', '')).rstrip('#')
        base_url = raw_url if raw_url.endswith('=') else raw_url.rstrip('/')

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
                        name_raw = " ".join(parts[3:])
                    elif len(parts) == 3:
                        date = parts[0]
                        size = parts[1]
                        name_raw = parts[2]
                    else:
                        logger.warning(f"Riga non parsabile: {line!r}")
                        continue
                        
                    import urllib.parse
                    name_decoded = urllib.parse.unquote(name_raw)
                    
                    if base_url.endswith('='):
                        url = f"{base_url}{name_raw}"
                    elif base_url:
                        # Assicuriamoci che name_raw sia URL-encoded (se proviene da WebDAV solitamente lo è)
                        url = f"{base_url}/download?path=&files={name_raw}"
                    else:
                        url = '#'
                    oxt_list.append({
                        'name': name_decoded,
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
            badge = '<span class="badge">ULTIMA</span>' if i == 0 else ''
            url = file.get('url', '#')
            rows.append(f"""
            <tr>
                <td class="col-name"><a href="{url}">{name}</a> {badge}</td>
                <td class="col-dim">{file['size']}</td>
                <td class="col-extra">{file['date']}</td>
                <td style="text-align: right;">
                    <a href="{url}" class="btn-download-link" download>
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                        Scarica
                    </a>
                </td>
            </tr>""")

        # Sezione commit
        if commits:
            commit_rows = []
            for c in commits:
                commit_rows.append(f"""
            <tr>
                <td>{c['date']}</td>
                <td><a href="{c['url']}" target="_blank" rel="noopener"><code>{c['sha']}</code></a></td>
                <td class="commit-msg">{c['msg']}</td>
            </tr>""")
            commits_section = f"""
        <div class="activity-section">
            <h2>Attività di sviluppo recente</h2>
            <div class="table-wrap">
                <table>
                    <thead>
                        <tr>
                            <th style="width:140px">Data</th>
                            <th style="width:90px">Commit</th>
                            <th>Descrizione</th>
                        </tr>
                    </thead>
                    <tbody>
                        {"".join(commit_rows)}
                    </tbody>
                </table>
            </div>
        </div>"""
        else:
            commits_section = """
        <div class="activity-section">
            <h2>Attività di sviluppo recente</h2>
            <p class="no-activity">Nessun commit recente trovato o errore nel recupero dei dati.</p>
        </div>"""
        html = f"""<!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="utf-8">
    <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
    <meta http-equiv="Pragma" content="no-cache">
    <meta http-equiv="Expires" content="0">
    <title>Versioni Nightly Builds LeenO</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;700&family=Space+Grotesk:wght@600;700&display=swap" rel="stylesheet">
    <style>
        /* ——— CSS Variables ——— */
        :root {{
            --bg-primary: #f0f4e0;
            --bg-dark: #121a23;
            --bg-green: #0e1a08;
            --accent-cyan: #aad400;
            --accent-cyan-hover: #8fb200;
            --accent-rust: #ff4d2e;
            --accent-rust-hover: #ff6b50;
            --text-primary: #ffffff;
            --text-secondary: #8896ab;
            --text-dark: #1a2010;
            --text-green: #5d7400;
            --font-display: 'Space Grotesk', sans-serif;
            --font-body: 'Inter', sans-serif;
            --font-mono: 'JetBrains Mono', monospace;
            --container: 1200px;
        }}

        /* ——— Reset & Base ——— */
        *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

        html {{
            font-size: 16px;
            scroll-behavior: smooth;
        }}

        body {{
            font-family: var(--font-body);
            background: var(--bg-primary);
            color: var(--text-dark);
            line-height: 1.6;
            overflow-x: hidden;
        }}

        h1, h2, h3 {{
            font-family: var(--font-display);
            font-weight: 700;
            line-height: 1.1;
            letter-spacing: -0.02em;
        }}

        a {{
            color: inherit;
            text-decoration: none;
            transition: color 0.2s, transform 0.2s;
        }}

        /* ——— Layout ——— */
        .container {{
            max-width: var(--container);
            margin: 0 auto;
            padding: 0 24px;
        }}

        /* ——— HEADER ——— */
        .site-header {{
            background: var(--bg-dark);
            padding: 15px 0;
            border-bottom: 1px solid rgba(170, 212, 0, 0.1);
        }}

        .header-inner {{
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}

        .site-logo {{
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .site-logo img {{
            height: 36px;
            width: auto;
        }}

        .site-name {{
            font-family: var(--font-display);
            font-weight: 700;
            color: #fff;
            letter-spacing: 0.1em;
            font-size: 1.1rem;
        }}

        .main-nav {{
            display: flex;
            gap: 20px;
        }}

        .main-nav a {{
            color: rgba(255, 255, 255, 0.7);
            font-size: 0.7rem;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }}

        .main-nav a:hover {{
            color: var(--accent-cyan);
        }}

        /* ——— PAGE HEADER (HERO) ——— */
        /* ——— CONTENT ——— */
        .content-main {{
            padding: 40px 0;
        }}

        /* ——— TABLE ——— */
        .table-wrap {{
            background: #fff;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.05);
            margin-bottom: 40px;
            overflow-x: auto;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            background: #fff;
        }}

        thead tr {{
            border-top: 3px solid #000;
            border-bottom: 3px solid #000;
        }}

        th {{
            background: #f9f9f9;
            color: #000;
            padding: 12px 20px;
            font-family: var(--font-display);
            font-weight: 800;
            text-transform: uppercase;
            font-size: 0.75rem;
            letter-spacing: 0.05em;
        }}

        td {{
            padding: 12px 20px;
            border-bottom: 1px solid #f0f0f0;
            vertical-align: middle;
        }}

        .col-name {{
            color: var(--accent-cyan);
            font-weight: 600;
            font-size: 1rem;
        }}

        .col-name a:hover {{
            text-decoration: underline;
        }}

        .col-dim, .col-extra {{
            color: #aaa;
            font-family: var(--font-mono);
            font-size: 0.85rem;
            text-align: right;
        }}

        tr:hover td {{
            background: rgba(0, 0, 0, 0.01);
        }}

        .badge {{
            display: inline-block;
            padding: 3px 8px;
            background: var(--accent-cyan);
            color: var(--bg-dark);
            font-family: var(--font-display);
            font-weight: 700;
            font-size: 0.65rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-left: 8px;
        }}

        .btn-download-link {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 12px;
            background: #ff4d2e;
            color: #ffffff !important;
            font-family: var(--font-display);
            font-weight: 800;
            padding: 12px 32px;
            font-size: 0.95rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            border: none;
            border-radius: 2px;
            cursor: pointer;
            box-shadow: 0 4px 15px rgba(255, 77, 46, 0.3);
            text-decoration: none !important;
        }}

        .btn-download-link:hover {{
            background: var(--accent-rust-hover);
            transform: translateY(-2px);
            box-shadow: 0 8px 25px rgba(255, 77, 46, 0.4);
            color: #ffffff !important;
        }}

        .btn-download-link svg {{
            transition: transform 0.3s;
            stroke: #ffffff !important;
            stroke-width: 3;
        }}

        .btn-download-link:hover svg {{
            transform: translateY(2px);
        }}

        /* ——— ACTIVITY SECTIONS ——— */
        .activity-section {{
            margin-bottom: 60px;
        }}

        .activity-section h2 {{
            font-size: 1.6rem;
            margin-bottom: 24px;
            text-transform: uppercase;
            color: var(--bg-dark);
            border-left: 4px solid var(--accent-cyan);
            padding-left: 16px;
        }}

        .no-activity {{
            color: var(--text-secondary);
            font-style: italic;
            padding: 24px;
            background: #fff;
            border: 1px dashed rgba(0,0,0,0.1);
            text-align: center;
        }}

        .commit-msg {{
            font-size: 0.9em;
            color: var(--text-secondary);
            max-width: 600px;
        }}

        code {{
            font-family: var(--font-mono);
            font-size: 0.85em;
            background: rgba(0,0,0,0.05);
            padding: 3px 6px;
            color: var(--text-green);
        }}

        /* ——— FOOTER ——— */
        .site-footer {{
            background: var(--bg-dark);
            color: #fff;
            padding: 40px 0;
            margin-top: 60px;
        }}

        .footer-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 24px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
            padding-bottom: 24px;
            margin-bottom: 24px;
        }}

        .footer-brand {{
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .footer-logo {{
            height: 28px;
            width: auto;
        }}

        .footer-license {{
            font-size: 0.7rem;
            color: var(--text-secondary);
        }}

        .footer-links {{
            display: flex;
            gap: 16px;
            flex-wrap: wrap;
        }}

        .footer-links a {{
            font-size: 0.75rem;
            color: var(--text-secondary);
        }}

        .footer-links a:hover {{
            color: var(--accent-cyan);
        }}

        .footer-bottom {{
            text-align: center;
            font-family: var(--font-mono);
            font-size: 0.7rem;
            color: var(--text-secondary);
        }}

        /* ——— Responsive ——— */
        @media (max-width: 768px) {{
            .header-inner {{
                flex-direction: column;
                gap: 12px;
            }}
            .main-nav {{
                gap: 12px;
                flex-wrap: wrap;
                justify-content: center;
            }}
            .footer-row {{
                flex-direction: column;
                text-align: center;
            }}
            .footer-brand {{
                flex-direction: column;
            }}
            .footer-links {{
                justify-content: center;
            }}
            td, th {{
                padding: 12px 16px;
            }}
        }}
    </style>
</head>
<body>
    <main class="container content-main">
        <div class="activity-section">
            <h2>Download disponibili</h2>
            <div class="table-wrap">
                <table>
                    <thead>
                        <tr>
                            <th>Nome</th>
                            <th style="text-align: right;">Dim.</th>
                            <th style="text-align: right;">&darr;</th>
                            <th style="width: 150px;"></th>
                        </tr>
                    </thead>
                    <tbody>
                        {"".join(rows)}
                    </tbody>
                </table>
            </div>
        </div>

        {commits_section}
    </main>

    <footer class="site-footer">
        <div class="container">
            <div class="footer-row">
                <div class="footer-brand">
                    <img src="https://leeno.org/wp-content/themes/leeno-theme/assets/images/logo-leeno.png" alt="LeenO" class="footer-logo">
                    <span class="footer-license">Open Source — LGPL v3</span>
                </div>
                <div class="footer-links">
                    <a href="https://www.libreoffice.org/" target="_blank">LibreOffice</a>
                    <a href="https://gitlab.com/giuserpe/leeno" target="_blank">GitLab</a>
                    <a href="https://github.com/giuserpe/leeno" target="_blank">GitHub</a>
                    <a href="https://t.me/leeno_computometrico" target="_blank">Telegram</a>
                    <a href="https://leeno.org/donazioni/">Dona!</a>
                </div>
            </div>
            <div class="footer-bottom">
                <p>Generato automaticamente il {now_utc} UTC</p>
                <p>Build Commit: <code>{version_info['git_sha']}</code> — Server: GitHub Actions</p>
            </div>
        </div>
    </footer>
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
