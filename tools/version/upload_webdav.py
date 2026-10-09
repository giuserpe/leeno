#!/usr/bin/env python3
"""
Carica i file generati su WebDAV con gestione dei tentativi e risoluzione
dell'errore FileLocked di Nextcloud/SabreDAV.
Supporta `requests` (se disponibile) e fallback robusto su `urllib`.
"""
import base64
import http.client
import os
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

try:
    import requests

    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False


def _make_request_requests(
    url: str, method: str, data: bytes = None, auth: tuple = None, timeout: int = 30
):
    headers = {
        "X-Requested-With": "XMLHttpRequest",
        "OCS-APIREQUEST": "true",
    }
    kwargs = {
        "headers": headers,
        "timeout": timeout,
    }
    if auth and (auth[0] or auth[1]):
        kwargs["auth"] = auth

    if method == "PUT":
        kwargs["data"] = data
        resp = requests.put(url, **kwargs)
    elif method == "DELETE":
        resp = requests.delete(url, **kwargs)
    else:
        resp = requests.request(method, url, **kwargs)

    return resp.status_code, resp.text


def _make_request_urllib(
    url: str, method: str, data: bytes = None, auth: tuple = None, timeout: int = 30
):
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("X-Requested-With", "XMLHttpRequest")
    req.add_header("OCS-APIREQUEST", "true")

    if auth and (auth[0] or auth[1]):
        user, password = auth
        credentials = f"{user}:{password}"
        b64_creds = base64.b64encode(credentials.encode("utf-8")).decode("utf-8")
        req.add_header("Authorization", f"Basic {b64_creds}")

    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            status = resp.status
            try:
                body = resp.read().decode("utf-8", errors="replace")
            except http.client.IncompleteRead as e:
                body = e.partial.decode("utf-8", errors="replace")
            return status, body
    except urllib.error.HTTPError as err:
        try:
            body_bytes = err.read()
        except http.client.IncompleteRead as e:
            body_bytes = e.partial
        except Exception:
            body_bytes = b""
        body = body_bytes.decode("utf-8", errors="replace")
        return err.code, body
    except http.client.IncompleteRead as e:
        body = e.partial.decode("utf-8", errors="replace")
        status = 423 if ("FileLocked" in body or "is locked" in body) else 500
        return status, body
    except Exception as err:
        raise err


def _make_request(
    url: str, method: str, data: bytes = None, auth: tuple = None, timeout: int = 30
):
    if HAS_REQUESTS:
        try:
            return _make_request_requests(url, method, data, auth, timeout)
        except Exception as req_err:
            print(f"   Note: requests call failed ({req_err}), falling back to urllib...")
    return _make_request_urllib(url, method, data, auth, timeout)


def upload_file(
    file_path: str,
    base_webdav_url: str,
    user: str,
    password: str,
    max_retries: int = 8,
    delays: list = None,
) -> bool:
    path = Path(file_path)
    if not path.is_file():
        print(f"❌ Errore: File non trovato: {file_path}")
        return False

    filename = path.name
    target_url = f"{base_webdav_url.rstrip('/')}/{filename}"
    auth = (user, password) if user or password else None

    print(f"📤 Caricamento in corso: {filename} -> {target_url}")

    if delays is None:
        delays = [3, 5, 10, 15, 20, 30, 40, 60]

    for attempt in range(1, max_retries + 1):
        try:
            with open(path, "rb") as f:
                file_data = f.read()

            status_code, response_text = _make_request(
                target_url, method="PUT", data=file_data, auth=auth, timeout=30
            )

            if status_code in (200, 201, 204):
                print(f"✅ {filename} caricato con successo (HTTP {status_code}).")
                return True

            is_locked = (
                status_code == 423
                or "FileLocked" in response_text
                or "is locked" in response_text
            )
            is_retryable = is_locked or status_code in (500, 502, 503, 504)

            print(
                f"⚠️ Tentativo {attempt}/{max_retries} per {filename} fallito (HTTP {status_code})."
            )
            if is_locked:
                print("   Dettaglio: Nextcloud segnala file bloccato (FileLocked).")

            if is_retryable and attempt < max_retries:
                # Tentativo di DELETE per rimuovere il file bloccato o lo lock stale in Nextcloud
                try:
                    print("   Tentativo di rimozione del file/lock remoto via DELETE...")
                    del_status, _ = _make_request(
                        target_url, method="DELETE", auth=auth, timeout=15
                    )
                    print(f"   Esito DELETE: HTTP {del_status}")
                except Exception as del_err:
                    print(f"   Richiesta DELETE fallita: {del_err}")

                delay = delays[min(attempt - 1, len(delays) - 1)]
                print(f"   Attesa di {delay}s prima di riprovare...")
                time.sleep(delay)
            elif not is_retryable:
                print(
                    f"❌ Errore non ripristinabile per {filename} (HTTP {status_code}):"
                )
                print(f"   {response_text[:300]}")
                return False

        except Exception as err:
            print(f"⚠️ Tentativo {attempt}/{max_retries} eccezione per {filename}: {err}")
            if attempt < max_retries:
                delay = delays[min(attempt - 1, len(delays) - 1)]
                time.sleep(delay)

    print(f"❌ Impossibile caricare {filename} dopo {max_retries} tentativi.")
    return False


def main():
    if len(sys.argv) < 2:
        print("Uso: python upload_webdav.py <file1> [file2 ...]")
        sys.exit(1)

    webdav_url = os.environ.get("WEBDAV_URL", "").rstrip("/")
    user = os.environ.get("WEBDAV_USER", "")
    password = os.environ.get("WEBDAV_PASS", "")

    if not webdav_url:
        print("❌ Errore: WEBDAV_URL non impostato in ambiente.")
        sys.exit(1)

    if not webdav_url.endswith("/LeenoNigthlyBuilds"):
        target_folder_url = f"{webdav_url}/LeenoNigthlyBuilds"
    else:
        target_folder_url = webdav_url

    files = sys.argv[1:]
    success = True

    for f in files:
        if not upload_file(f, target_folder_url, user, password):
            success = False

    if not success:
        sys.exit(1)


if __name__ == "__main__":
    main()
