import os
import re
import io
import asyncio
import pandas as pd
import requests
from urllib.parse import urlsplit, parse_qs, urlencode, urlunsplit

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# === Google Drive API ===
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.service_account import Credentials

from orchestrator import evaluate_file_for_main
from utils import load_document_content, raw_total, weighted_total  # compatibility

INPUT_EXCEL = './SIH 2025 project submission for prescreening.xlsx'
OUTPUT_EXCEL = 'evaluation_results.xlsx'
FILE_DIR = 'ppt'

ALLOWED_EXTS = {'.pdf', '.ppt', '.pptx'}
GDRIVE_EXPORTS = {
    'application/vnd.google-apps.presentation': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',  # -> .pptx
}
MIMETYPE_TO_EXT = {
    'application/pdf': '.pdf',
    'application/vnd.ms-powerpoint': '.ppt',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
}

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)

# -------- Utils --------
def make_session():
    retries = Retry(
        total=5, connect=5, read=5,
        backoff_factor=0.6,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["HEAD", "GET", "OPTIONS"]
    )
    s = requests.Session()
    s.headers.update({"User-Agent": UA})
    s.mount("https://", HTTPAdapter(max_retries=retries))
    s.mount("http://", HTTPAdapter(max_retries=retries))
    return s

def sanitize(value: str) -> str:
    value = value or ""
    value = re.sub(r"[\\/:*?\"<>|\n\r\t]+", "_", str(value))
    value = re.sub(r"\s+", "_", value.strip())
    return value[:180] or "file"

def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

# -------- Drive helpers --------
def get_drive_service():
    """
    Uses service_account.json in CWD or GOOGLE_APPLICATION_CREDENTIALS env.
    Scope: drive.readonly
    """
    sa_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")
    creds = Credentials.from_service_account_file(
        sa_path,
        scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)

def is_drive_url(url: str) -> bool:
    host = urlsplit(url).netloc.lower()
    return "drive.google.com" in host or "docs.google.com" in host

def parse_drive_ids(url: str):
    """
    Returns (folder_id, file_id, is_presentation_link)
    Supports:
      - https://drive.google.com/drive/folders/{FOLDER_ID}
      - https://drive.google.com/open?id={FILE_ID}
      - https://drive.google.com/file/d/{FILE_ID}/view
      - https://docs.google.com/presentation/d/{FILE_ID}/edit
    """
    u = urlsplit(url)
    parts = u.path.strip("/").split("/")
    qs = parse_qs(u.query)

    # folder
    if "folders" in parts:
        try:
            return parts[parts.index("folders")+1], None, False
        except Exception:
            pass

    # file
    if "file" in parts and "d" in parts:
        try:
            return None, parts[parts.index("d")+1], False
        except Exception:
            pass

    # presentation
    if "presentation" in parts and "d" in parts:
        try:
            return None, parts[parts.index("d")+1], True
        except Exception:
            pass

    # open?id=...
    if "id" in qs:
        return None, qs["id"][0], False

    return None, None, False

def list_drive_folder_files(svc, folder_id: str):
    """
    Yields file dicts for PDFs and PPTs and Google Slides inside folder (non-recursive).
    """
    q = f"'{folder_id}' in parents and trashed=false and (" \
        f"mimeType='application/pdf' or " \
        f"mimeType='application/vnd.ms-powerpoint' or " \
        f"mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation' or " \
        f"mimeType='application/vnd.google-apps.presentation'" \
        f")"
    page_token = None
    while True:
        resp = svc.files().list(
            q=q,
            fields="nextPageToken, files(id, name, mimeType)",
            pageSize=1000,
            pageToken=page_token
        ).execute()
        for f in resp.get("files", []):
            yield f
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

def download_drive_file(svc, file_id: str, name: str, mime_type: str, out_dir: str) -> str | None:
    """
    Downloads a single Drive file. Exports Slides to PPTX.
    Returns saved path or None.
    """
    ensure_dir(out_dir)

    if mime_type in GDRIVE_EXPORTS:
        export_mime = GDRIVE_EXPORTS[mime_type]
        ext = MIMETYPE_TO_EXT[export_mime]
        req = svc.files().export_media(fileId=file_id, mimeType=export_mime)
    else:
        # Only allow target types
        if mime_type not in MIMETYPE_TO_EXT:
            return None
        ext = MIMETYPE_TO_EXT[mime_type]
        req = svc.files().get_media(fileId=file_id)

    safe_name = sanitize(os.path.splitext(name)[0]) + ext
    dst = os.path.join(out_dir, safe_name)

    # avoid overwrite
    base_no_ext, ext_only = os.path.splitext(dst)
    k = 1
    while os.path.exists(dst):
        dst = f"{base_no_ext}({k}){ext_only}"
        k += 1

    fh = io.FileIO(dst, mode='wb')
    downloader = MediaIoBaseDownload(fh, req, chunksize=1024*1024)
    done = False
    try:
        while not done:
            status, done = downloader.next_chunk()
        return dst
    finally:
        fh.close()

def download_drive_folder_recursive(svc, folder_id: str, out_dir: str) -> list[str]:
    """
    Non-recursive by default as requested. To include subfolders, extend here.
    """
    saved = []
    for f in list_drive_folder_files(svc, folder_id):
        path = download_drive_file(svc, f["id"], f["name"], f["mimeType"], out_dir)
        if path:
            saved.append(path)
    return saved

def download_drive_single(svc, file_id: str, is_presentation_hint: bool, out_dir: str) -> str | None:
    # Fetch file metadata to know mimeType and name
    meta = svc.files().get(fileId=file_id, fields="id, name, mimeType").execute()
    mime = meta["mimeType"]
    name = meta["name"]

    # If user pasted a presentation URL, we will export to PPTX
    return download_drive_file(svc, file_id, name, mime, out_dir)

# -------- Non-Drive downloader (kept for completeness) --------
def resolve_download_url(url: str) -> str:
    u = urlsplit(url)
    host = u.netloc.lower()
    # OneDrive/SharePoint force download
    if any(x in host for x in ["onedrive.live.com", "1drv.ms", "sharepoint.com"]):
        q = parse_qs(u.query)
        q["download"] = ["1"]
        new_query = urlencode({k: (v[0] if isinstance(v, list) else v) for k, v in q.items()})
        return urlunsplit((u.scheme, u.netloc, u.path, new_query, u.fragment))
    return url

def guess_ext_from_url(url: str) -> str:
    path = urlsplit(url).path
    ext = os.path.splitext(path)[1].lower()
    return ext if ext in ALLOWED_EXTS else ""

def http_download(session: requests.Session, url: str, out_dir: str, base_name: str) -> str | None:
    ensure_dir(out_dir)
    url = resolve_download_url(url)
    r = session.get(url, stream=True, timeout=90)
    r.raise_for_status()
    # extension
    ct = r.headers.get("Content-Type", "").split(";")[0].strip().lower()
    ext = MIMETYPE_TO_EXT.get(ct) or guess_ext_from_url(url)
    if ext not in ALLOWED_EXTS:
        # ignore non target files
        return None

    safe = sanitize(base_name) + ext
    dst = os.path.join(out_dir, safe)
    base_no_ext, ext_only = os.path.splitext(dst)
    k = 1
    while os.path.exists(dst):
        dst = f"{base_no_ext}({k}){ext_only}"
        k += 1

    with open(dst, "wb") as f:
        for chunk in r.iter_content(8192):
            if chunk:
                f.write(chunk)
    return dst

# -------- Evaluation (unchanged logic) --------
def evaluate_file(file_path: str):
    try:
        report = asyncio.run(evaluate_file_for_main(file_path, agent_mode="combined"))
        return {
            'status': 'Success',
            'num_images': len(report.get('workflow_analysis', {}).get('images', [])),
            'total_raw': report.get('total_raw', 0),
            'total_weighted': report.get('total_weighted', 0),
            'scores': report.get('scores', {}),
            'num_slides': report.get('workflow_analysis', {}).get('num_slides', ''),
            'summary': report.get('summary', ''),
            'feedback': report.get('feedback', {}),
        }
    except Exception as e:
        return {'status': f'Error: {e}'}

# -------- Main --------
def main():
    ensure_dir(FILE_DIR)
    df = pd.read_excel(INPUT_EXCEL)

    session = make_session()
    drive = None  # lazy init only if needed

    results = []

    for _, row in df.iterrows():
        team = str(row.get('Team name') or "").strip()
        leader = str(row.get('Team leader name') or "").strip()
        link = row.get('Idea PPT')

        # skip empty
        if not isinstance(link, str) or not link.strip():
            results.append({
                'Team name': team,
                'Team leader name': leader,
                'File Link': link,
                'File': '',
                'Evaluation Status': 'No URL',
                'Number of Images': 0,
                'Total Raw Score': 0,
                'Total Weighted Score': 0,
                'Scores': '{}',
                'Number of Slides': '',
                'Summary': '',
                'Feedback': '{}',
            })
            continue

        saved_paths = []

        if is_drive_url(link):
            # parse ids
            folder_id, file_id, is_pres = parse_drive_ids(link)
            if drive is None:
                # init once
                try:
                    drive = get_drive_service()
                except Exception as e:
                    drive = None
            if drive is None:
                # cannot use Drive API; skip to HTTP fallback
                base_name = f"{sanitize(team)}_{sanitize(leader)}"
                p = http_download(session, link.strip(), FILE_DIR, base_name)
                if p:
                    saved_paths.append(p)
            else:
                if folder_id:
                    # download all PDFs/PPTs in folder (non-recursive)
                    saved_paths.extend(download_drive_folder_recursive(drive, folder_id, FILE_DIR))
                elif file_id:
                    p = download_drive_single(drive, file_id, is_pres, FILE_DIR)
                    if p: saved_paths.append(p)
        else:
            base_name = f"{sanitize(team)}_{sanitize(leader)}"
            p = http_download(session, link.strip(), FILE_DIR, base_name)
            if p:
                saved_paths.append(p)

        # If nothing downloaded, record failure
        if not saved_paths:
            results.append({
                'Team name': team,
                'Team leader name': leader,
                'File Link': link,
                'File': '',
                'Evaluation Status': 'Download Failed or No PDF/PPT found',
                'Number of Images': 0,
                'Total Raw Score': 0,
                'Total Weighted Score': 0,
                'Scores': '{}',
                'Number of Slides': '',
                'Summary': '',
                'Feedback': '{}',
            })
            continue

        # Evaluate each saved file; append one row per file
        for p in saved_paths:
            eval_result = evaluate_file(p)
            results.append({
                'Team name': team,
                'Team leader name': leader,
                'File Link': link,
                'File': os.path.basename(p),
                'Evaluation Status': eval_result['status'],
                'Number of Images': eval_result.get('num_images', 0),
                'Total Raw Score': eval_result.get('total_raw', 0),
                'Total Weighted Score': eval_result.get('total_weighted', 0),
                'Scores': str(eval_result.get('scores', {})),
                'Number of Slides': eval_result.get('num_slides', ''),
                'Summary': eval_result.get('summary', ''),
                'Feedback': str(eval_result.get('feedback', {})),
            })

    pd.DataFrame(results).to_excel(OUTPUT_EXCEL, index=False)
    print(f"Saved results to {OUTPUT_EXCEL}")

if __name__ == "__main__":
    main()
