#!/usr/bin/env python3
"""
Legge commits_raw.json (risposta API GitHub), filtra i commit automatici
e salva i primi 10 significativi in commits.json.
"""
import json
from datetime import datetime
import os, sys

RAW_PATH = 'commits_raw.json'
JSON_PATH = 'commits.json'

if not os.path.exists(RAW_PATH):
    print(f"DEBUG: {RAW_PATH} non trovato, creo lista vuota.")
    with open(JSON_PATH, 'w') as f:
        f.write('[]')
    sys.exit(0)

SKIP_PREFIXES = (
    'Merge branch',
    'merge branch',
)

try:
    with open(RAW_PATH, 'r', encoding='utf-8') as f:
        raw = json.load(f)
except Exception as e:
    print(f"ERROR: Fallito caricamento JSON da {RAW_PATH}: {e}")
    with open(JSON_PATH, 'w') as f:
        f.write('[]')
    sys.exit(1)

if not isinstance(raw, list):
    print(f"ERROR: La risposta API non è una lista. Ricevuto: {type(raw)}")
    if isinstance(raw, dict):
        print(f"Messaggio API: {raw.get('message', 'Nessun messaggio')}")
    with open(JSON_PATH, 'w') as f:
        f.write('[]')
    sys.exit(0)

commits = []
print(f"DEBUG: Processo {len(raw)} elementi da raw JSON")
for i, item in enumerate(raw):
    if not isinstance(item, dict):
        print(f"DEBUG: Elemento {i} non è un dict: {type(item)}")
        continue
        
    commit_data = item.get('commit', {})
    msg_full = commit_data.get('message', '').strip()
    if not msg_full:
        continue
        
    msg = msg_full.splitlines()[0]  # solo prima riga

    # Salta commit automatici
    if any(msg.startswith(p) for p in SKIP_PREFIXES):
        continue

    sha = item.get('sha', '')[:7]
    url = item.get('html_url', '')
    date_raw = commit_data.get('author', {}).get('date', '')
    
    try:
        # Gestione formati data GitHub
        dt = datetime.fromisoformat(date_raw.replace('Z', '+00:00'))
        date = dt.strftime('%Y-%m-%d %H:%M')
    except Exception:
        date = date_raw[:10] if date_raw else 'N/A'

    author = commit_data.get('author', {}).get('name', 'N/A')
    commits.append({
        'sha': sha,
        'msg': msg,
        'date': date,
        'url': url,
        'author': author,
    })

    if len(commits) >= 10:
        break

with open(JSON_PATH, 'w', encoding='utf-8') as f:
    json.dump(commits, f, ensure_ascii=False, indent=2)

print(f"SUCCESS: Salvati {len(commits)} commit in {JSON_PATH}")
