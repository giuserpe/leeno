import os
import hashlib
import requests
from xml.etree import ElementTree
from requests.auth import HTTPBasicAuth
from datetime import datetime

# --- CONFIGURAZIONE ---
WEBDAV_URL = "https://dev.leeno.org/remote.php/dav/files/loadmin/LeenoNigthlyBuilds/"
WEBDAV_USER = "loadmin"
WEBDAV_PASS = "yGT.:jya09--"
GITHUB_API_TAGS = "https://api.github.com/repos/giuserpe/leeno/tags"
PUBLIC_DOWNLOAD_BASE = "https://dev.leeno.org/index.php/s/jLnxqWRzSD7MqFB/download?path=&files="

# --- FUNZIONI ---
def list_webdav_oxt_files():
    print("📥 Connessione a Nextcloud WebDAV...")
    headers = {"Depth": "infinity"}
    propfind = """
    <d:propfind xmlns:d='DAV:'>
        <d:prop>
            <d:displayname/>
            <d:getlastmodified/>
            <d:getcontentlength/>
        </d:prop>
    </d:propfind>
    """
    res = requests.request("PROPFIND", WEBDAV_URL, data=propfind, headers=headers,
                           auth=HTTPBasicAuth(WEBDAV_USER, WEBDAV_PASS))
    res.raise_for_status()

    tree = ElementTree.fromstring(res.content)
    files = []
    for resp in tree.findall("{DAV:}response"):
        href = resp.find("{DAV:}href")
        if href is not None and href.text.endswith(".oxt"):
            filename = os.path.basename(href.text)
            if filename == "LeenO.oxt":
                continue  # Escludi LeenO.oxt
            remote_path = WEBDAV_URL + filename

            lastmod_elem = resp.find(".//{DAV:}getlastmodified")
            size_elem = resp.find(".//{DAV:}getcontentlength")

            lastmod = lastmod_elem.text if lastmod_elem is not None else ""
            size = int(size_elem.text) if size_elem is not None else 0

            files.append({
                "filename": filename,
                "remote_url": remote_path,
                "public_url": PUBLIC_DOWNLOAD_BASE + filename,
                "lastmod": lastmod,
                "size": size
            })
            print(f"  ➕ Trovato: {filename}")

    print(f"📦 Trovati {len(files)} file .oxt (escluso LeenO.oxt)")
    return files

def calculate_sha256(url):
    print(f"🔢 Calcolo SHA256 per: {url}")
    sha256 = hashlib.sha256()
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        for chunk in r.iter_content(chunk_size=8192):
            sha256.update(chunk)
    return sha256.hexdigest()

def get_github_tags():
    print("🔗 Recupero tag GitHub...")
    r = requests.get(GITHUB_API_TAGS)
    r.raise_for_status()
    tags = [tag["name"] for tag in r.json()]
    return tags

def format_size(size_bytes):
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"

def build_html(files, github_tags):
    latest_file = sorted(files, key=lambda x: x['filename'], reverse=True)[0] if files else None
    files = sorted(files, key=lambda x: x['filename'], reverse=True)[:30]  # Solo le ultime 30

    style = """
    <style>
    body { font-family: sans-serif; max-width: 1000px; margin: auto; padding: 20px; }
    h1 { text-align: center; }
    table { width: 100%; border-collapse: collapse; }
    th, td { padding: 8px 12px; border: 1px solid #ccc; text-align: left; }
    th { background-color: #f4f4f4; }
    tr.latest { background-color: #e0ffe0; }
    .hash { font-family: monospace; word-break: break-all; font-size: 0.85em; }
    .badge-latest { background: #28a745; color: white; padding: 2px 6px; border-radius: 4px; font-size: 0.75em; }
    </style>
    <link rel=\"stylesheet\" href=\"https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css\">
    <script src=\"https://code.jquery.com/jquery-3.6.0.min.js\"></script>
    <script src=\"https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js\"></script>
    <script>
    jQuery(document).ready(function($) {
      $('#leenotable').DataTable({
        pageLength: 10,
        order: [[3, 'desc']]
      });
    });
    </script>
    """

    now = datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')

    html = ["<html><head><meta charset='utf-8'><title>Versioni LeenO</title>", style, "</head><body>"]
    html.append("<h1>Versioni di sviluppo LeenO</h1>")
    html.append(f"<p><em>Aggiornato l'ultima volta il {now}</em></p>")
    html.append("<table id='leenotable'>")
    html.append("<thead><tr><th>Versione</th><th>Download</th><th>SHA256</th><th>Data</th><th>Dimensione</th></tr></thead><tbody>")

    for file in files:
        is_latest = file == latest_file
        version = file['filename'].replace("Leeno-", "").replace(".oxt", "")
        sha256 = calculate_sha256(file['public_url'])
        data = file['lastmod'] or ""
        size = format_size(file['size'])

        row_class = " class='latest'" if is_latest else ""
        badge = " <span class='badge-latest'>Ultima</span>" if is_latest else ""

        html.append(f"<tr{row_class}>")
        html.append(f"<td>{file['filename']}{badge}</td>")
        html.append(f"<td><a href='{file['public_url']}'>Download</a></td>")
        html.append(f"<td class='hash'>{sha256}</td>")
        html.append(f"<td>{data}</td>")
        html.append(f"<td>{size}</td>")
        html.append("</tr>")

    html.append("</tbody></table>")
    html.append("</body></html>")
    return "\n".join(html)

def main():
    files = list_webdav_oxt_files()
    github_tags = get_github_tags()
    html = build_html(files, github_tags)
    with open("versions.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ File 'versions.html' generato correttamente.")

if __name__ == "__main__":
    main()