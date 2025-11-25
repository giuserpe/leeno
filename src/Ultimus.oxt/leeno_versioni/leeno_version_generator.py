# WEBDAV_PASS = "yGT.:jya09--"
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
            <d:getcontentlength/>
            <d:getlastmodified/>
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
            props = resp.find("{DAV:}propstat/{DAV:}prop")
            size = props.find("{DAV:}getcontentlength")
            modified = props.find("{DAV:}getlastmodified")

            size = int(size.text) if size is not None and size.text else 0
            modified_str = modified.text if modified is not None else ""
            try:
                modified_dt = datetime.strptime(modified_str, "%a, %d %b %Y %H:%M:%S %Z")
                modified_fmt = modified_dt.strftime("%Y-%m-%d %H:%M")
            except:
                modified_fmt = ""

            files.append({
                "filename": filename,
                "remote_url": remote_path,
                "public_url": PUBLIC_DOWNLOAD_BASE + filename,
                "size": size,
                "modified": modified_fmt
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

def build_html(files, github_tags):
    latest_file = sorted(files, key=lambda x: x['filename'], reverse=True)[0] if files else None

    style = """
    <style>
    body { font-family: sans-serif; max-width: 1000px; margin: auto; padding: 20px; }
    h1 { text-align: center; }
    table { width: 100%; border-collapse: collapse; }
    th, td { padding: 8px 12px; border: 1px solid #ccc; text-align: left; }
    th { background-color: #f4f4f4; }
    tr.latest { background-color: #e0ffe0; }
    .hash { font-family: monospace; word-break: break-all; font-size: 0.85em; }
    </style>
    """

    html = ["<html><head><meta charset='utf-8'><title>Versioni LeenO</title>", style, "</head><body>"]
    html.append("<h1>Versioni di sviluppo LeenO</h1>")
    html.append("<table>")
    html.append("<tr><th>Versione</th><th>Download</th><th>SHA256</th><th>Data</th><th>Dimensione</th></tr>")

    for file in sorted(files, key=lambda x: x['filename'], reverse=True):
        is_latest = file == latest_file
        sha256 = calculate_sha256(file['public_url'])
        size_kb = f"{file['size'] // 1024} KB" if file['size'] else ""

        row_class = " class='latest'" if is_latest else ""
        html.append(f"<tr{row_class}>")
        html.append(f"<td>{file['filename']}</td>")
        html.append(f"<td><a href='{file['public_url']}'>Download</a></td>")
        html.append(f"<td class='hash'>{sha256}</td>")
        html.append(f"<td>{file['modified']}</td>")
        html.append(f"<td>{size_kb}</td>")
        html.append("</tr>")

    html.append("</table>")
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
