#!/usr/bin/env python3
"""
Legge commits_raw.json (risposta API GitHub), filtra i commit automatici
e salva i primi 8 significativi in commits.json.
"""
import json
from datetime import datetime, timezone

SKIP_PREFIXES = (
    'chore: aggiorna versions.html',
    'chore: aggiorna version',
    'Merge branch',
    'merge branch',
)

with open('commits_raw.json', 'r', encoding='utf-8') as f:
    raw = json.load(f)

commits = []
for item in raw:
    msg_full = item.get('commit', {}).get('message', '').strip()
    msg = msg_full.splitlines()[0]  # solo prima riga

    # Salta commit automatici
    if any(msg.startswith(p) for p in SKIP_PREFIXES):
        continue

    sha = item.get('sha', '')[:7]
    url = item.get('html_url', '')
    date_raw = item.get('commit', {}).get('author', {}).get('date', '')
    try:
        dt = datetime.fromisoformat(date_raw.replace('Z', '+00:00'))
        date = dt.strftime('%Y-%m-%d %H:%M')
    except Exception:
        date = date_raw[:10] if date_raw else 'N/A'

    commits.append({
        'sha': sha,
        'msg': msg,
        'date': date,
        'url': url,
    })

    if len(commits) >= 8:
        break

with open('commits.json', 'w', encoding='utf-8') as f:
    json.dump(commits, f, ensure_ascii=False, indent=2)

print(f"Salvati {len(commits)} commit in commits.json")
