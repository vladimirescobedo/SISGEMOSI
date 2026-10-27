import hashlib
import http.client
import json
import os
import ssl
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlencode, urlsplit

from google.auth.transport.requests import Request
from google.auth.exceptions import GoogleAuthError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload


RECORDINGS_DIR = Path(os.getenv("APOWERREC_RECORDINGS_DIR", r"D:\GRABACIONES"))
SISGEMOSI_URL = os.getenv("SISGEMOSI_URL", "https://10.5.3.116:5000").rstrip("/")
SYNC_TOKEN = os.getenv("APOWERREC_SYNC_TOKEN", "")
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.file"]
POLL_SECONDS = 30
STABLE_SECONDS = 60
MAX_UPLOAD_BYTES = 4 * 1024 * 1024 * 1024
CHUNK_SIZE = 1024 * 1024


def log(message):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}", flush=True)


def state_file_path():
    local_app_data = os.getenv("LOCALAPPDATA")
    root = Path(local_app_data) if local_app_data else Path.home()
    state_dir = root / "SISGEMOSI"
    state_dir.mkdir(parents=True, exist_ok=True)
    return state_dir / "apowerrec_sync_state.json"


def load_synced_hashes(path):
    if not path.exists():
        return set()
    try:
        values = json.loads(path.read_text(encoding="utf-8"))
        return set(values if isinstance(values, list) else [])
    except (OSError, json.JSONDecodeError) as error:
        log(f"No se pudo leer el historial local de sincronización: {error}")
        return set()


def save_synced_hashes(path, hashes):
    temporary_path = path.with_suffix(".tmp")
    temporary_path.write_text(
        json.dumps(sorted(hashes), ensure_ascii=False),
        encoding="utf-8"
    )
    temporary_path.replace(path)


def create_ssl_context():
    ca_path = os.getenv("SISGEMOSI_CA_CERT", "").strip()
    if not ca_path:
        candidates = (
            Path(__file__).with_name("rootCA.pem"),
            Path(__file__).parent / "certificados" / "rootCA.pem"
        )
        ca_path = next((str(path) for path in candidates if path.is_file()), "")
    if ca_path:
        return ssl.create_default_context(cafile=ca_path)
    return ssl.create_default_context()


def get_drive_service(state_path):
    token_path = state_path.parent / "google_drive_token.json"
    client_file = Path(os.getenv(
        "GOOGLE_OAUTH_CLIENT_FILE",
        str(Path(__file__).with_name("client_secret.json"))
    ))
    if not client_file.is_file():
        raise RuntimeError(
            f"No se encontró el cliente OAuth de escritorio de Google: {client_file}"
        )

    credentials = None
    if token_path.is_file():
        credentials = Credentials.from_authorized_user_file(str(token_path), DRIVE_SCOPES)
    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
    if not credentials or not credentials.valid:
        flow = InstalledAppFlow.from_client_secrets_file(str(client_file), DRIVE_SCOPES)
        credentials = flow.run_local_server(port=0, access_type="offline", prompt="consent")
    temporary_path = token_path.with_suffix(".tmp")
    temporary_path.write_text(credentials.to_json(), encoding="utf-8")
    temporary_path.replace(token_path)
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def server_request(method, endpoint, context, query=None, json_payload=None):
    parsed_url = urlsplit(SISGEMOSI_URL)
    if parsed_url.scheme != "https" or not parsed_url.hostname:
        raise ValueError("SISGEMOSI_URL debe ser una dirección HTTPS válida.")
    if parsed_url.path not in ("", "/") or parsed_url.query or parsed_url.fragment:
        raise ValueError("SISGEMOSI_URL debe contener solo el origen, sin rutas ni parámetros.")

    route = endpoint
    if query:
        route = f"{route}?{urlencode(query)}"
    connection = http.client.HTTPSConnection(
        parsed_url.hostname,
        parsed_url.port or 443,
        context=context,
        timeout=600
    )
    request_headers = {
        "Authorization": f"Bearer {SYNC_TOKEN}",
        "Accept": "application/json"
    }
    try:
        body = None
        if json_payload is not None:
            body = json.dumps(json_payload).encode("utf-8")
            request_headers["Content-Type"] = "application/json"
            request_headers["Content-Length"] = str(len(body))
        if body is not None:
            connection.request(method, route, body=body, headers=request_headers)
        else:
            connection.request(method, route, headers=request_headers)

        response = connection.getresponse()
        raw_body = response.read()
        try:
            payload = json.loads(raw_body.decode("utf-8")) if raw_body else {}
        except (UnicodeDecodeError, json.JSONDecodeError):
            payload = {}
        if response.status < 200 or response.status >= 300:
            message = payload.get("error") or f"El servidor respondió HTTP {response.status}."
            raise RuntimeError(message)
        return payload
    finally:
        connection.close()


def file_sha256(file_path):
    digest = hashlib.sha256()
    with file_path.open("rb") as recording:
        while True:
            chunk = recording.read(CHUNK_SIZE)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def find_or_upload_drive_file(service, file_path, digest, event_title):
    existing_files = service.files().list(
        q=f"appProperties has {{ key='sisgemosi_sha256' and value='{digest}' }} and trashed=false",
        pageSize=1,
        fields="files(id,name,webViewLink)"
    ).execute().get("files", [])
    if existing_files:
        drive_file = existing_files[0]
    else:
        metadata = {
            "name": file_path.name,
            "description": f"Grabación ApowerREC de la actividad TI: {event_title}",
            "appProperties": {"sisgemosi_sha256": digest}
        }
        media = MediaFileUpload(
            str(file_path),
            mimetype="video/mp4",
            chunksize=8 * 1024 * 1024,
            resumable=True
        )
        drive_file = service.files().create(
            body=metadata,
            media_body=media,
            fields="id,name,webViewLink"
        ).execute()

    drive_file_id = drive_file.get("id")
    if not drive_file_id:
        raise RuntimeError("Google Drive no devolvió el identificador del video.")
    drive_url = drive_file.get("webViewLink") or f"https://drive.google.com/file/d/{drive_file_id}/view"
    return drive_file_id, drive_url


def upload_recording(file_path, context, drive_service, activity_id, event_title):
    file_stat = file_path.stat()
    if file_stat.st_size <= 0 or file_stat.st_size > MAX_UPLOAD_BYTES:
        raise ValueError(f"El archivo está vacío o supera el límite de 4 GB: {file_path.name}")
    digest = file_sha256(file_path)
    drive_file_id, drive_url = find_or_upload_drive_file(
        drive_service, file_path, digest, event_title
    )
    payload = server_request(
        "POST",
        "/api/apowerrec/recordings",
        context,
        json_payload={
            "filename": file_path.name,
            "sha256": digest,
            "actividad_id": activity_id,
            "drive_file_id": drive_file_id,
            "drive_url": drive_url,
            "fecha_grabacion": datetime.fromtimestamp(file_stat.st_mtime).date().isoformat(),
            "mtime": file_stat.st_mtime
        }
    )
    if payload.get("estado") != "SINCRONIZADA":
        raise RuntimeError("El servidor no confirmó el estado SINCRONIZADA.")
    if payload.get("advertencia"):
        log(f"Advertencia del servidor: {payload['advertencia']}")
    return digest, payload


def run():
    if not SYNC_TOKEN:
        raise RuntimeError("Configure la variable APOWERREC_SYNC_TOKEN en esta PC.")
    if not RECORDINGS_DIR.is_dir():
        raise RuntimeError(f"No existe la carpeta de grabaciones: {RECORDINGS_DIR}")

    context = create_ssl_context()
    state_path = state_file_path()
    synced_hashes = load_synced_hashes(state_path)
    drive_service = get_drive_service(state_path)
    retry_after = {}
    observed_sizes = {}
    stable_since = {}
    reported_issues = set()

    log(f"Vigilando grabaciones MP4 en {RECORDINGS_DIR}")
    log(f"SISGEMOSI: {SISGEMOSI_URL}")
    while True:
        now = time.time()
        try:
            files = sorted(
                (path for path in RECORDINGS_DIR.rglob("*.mp4") if path.is_file() and not path.name.startswith(".")),
                key=lambda path: path.stat().st_mtime
            )
        except OSError as error:
            log(f"No se pudo revisar la carpeta de grabaciones: {error}")
            time.sleep(POLL_SECONDS)
            continue

        candidates = []
        for file_path in files:
            try:
                stats = file_path.stat()
            except OSError:
                continue
            key = str(file_path.resolve())
            if stats.st_size > MAX_UPLOAD_BYTES:
                issue = (key, stats.st_mtime_ns, "oversize")
                if issue not in reported_issues:
                    log(f"Se omite {file_path.name}: supera el límite de 4 GB.")
                    reported_issues.add(issue)
                continue
            if stats.st_size <= 0:
                continue
            if observed_sizes.get(key) != stats.st_size:
                observed_sizes[key] = stats.st_size
                stable_since[key] = now
                continue
            if now - stable_since.get(key, now) < STABLE_SECONDS or now - stats.st_mtime < STABLE_SECONDS:
                continue
            candidates.append(file_path)

        if candidates:
            try:
                status = server_request("GET", "/api/apowerrec/status", context)
                if not status.get("ready"):
                    motivo = status.get("motivo") or (
                        "Debe existir exactamente una actividad GRABAR SESION en progreso "
                        f"(actuales: {status.get('tareas_en_progreso', '?')})."
                    )
                    log(f"Esperando actividad de grabación: {motivo}")
                else:
                    started_at = float(status["iniciada_el"])
                    for file_path in candidates:
                        key = str(file_path.resolve())
                        if retry_after.get(key, 0) > now:
                            continue
                        file_stat = file_path.stat()
                        if file_stat.st_mtime < started_at - 2:
                            issue = (key, file_stat.st_mtime_ns, "before-activity")
                            if issue not in reported_issues:
                                log(f"Se omite {file_path.name}: se guardó antes de iniciar la actividad actual.")
                                reported_issues.add(issue)
                            continue
                        digest = file_sha256(file_path)
                        if digest in synced_hashes:
                            continue
                        try:
                            _, result = upload_recording(
                                file_path,
                                context,
                                drive_service,
                                status["actividad_id"],
                                status["titulo"]
                            )
                            synced_hashes.add(digest)
                            save_synced_hashes(state_path, synced_hashes)
                            retry_after.pop(key, None)
                            log(
                                f"Sincronizado: {file_path.name} → {result.get('nombre_evento')} "
                                f"| {result.get('drive_url')}"
                            )
                        except (OSError, RuntimeError, ValueError, http.client.HTTPException, ssl.SSLError, HttpError, GoogleAuthError) as error:
                            retry_after[key] = time.time() + 120
                            log(f"No se pudo sincronizar {file_path.name}; se reintentará: {error}")
                        break
            except (OSError, RuntimeError, ValueError, http.client.HTTPException, ssl.SSLError, GoogleAuthError) as error:
                log(f"No se pudo consultar SISGEMOSI; se reintentará: {error}")

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    try:
        run()
    except KeyboardInterrupt:
        log("Sincronización detenida.")
    except Exception as error:
        log(f"Error de configuración: {error}")
        sys.exit(1)
