"""
sharepoint_indexer.py — SharePoint-to-Memory Indexer for MCS CoWorker.

Walks every client folder under Server/Clients/ in SharePoint, extracts text
from supported file types, and stores a summary + metadata into the MemoryStore
documents collection for semantic search.

Supported file types:
  - .txt, .md, .csv  → read directly as UTF-8
  - .pdf             → extract text via pdfminer.six
  - .docx            → extract text via python-docx

Only files modified since the last index run are re-indexed (incremental).
"""
from __future__ import annotations

import io
import logging
import re
import time
from datetime import datetime, timezone
from typing import TYPE_CHECKING, Optional

from config import get_setting, set_setting

if TYPE_CHECKING:
    from graph_client import GraphClient
    from memory_store import MemoryStore

logger = logging.getLogger(__name__)

# ── Constants ─────────────────────────────────────────────────────────────────

SETTING_LAST_RUN     = "sharepoint_index_last_run"       # ISO timestamp
SETTING_LAST_COUNT   = "sharepoint_index_last_count"     # int — docs indexed last run
SETTING_TOTAL_COUNT  = "sharepoint_index_total_count"    # int — total docs in store
SETTING_STATUS       = "sharepoint_index_status"         # "idle" | "running" | "error"
SETTING_LAST_ERROR   = "sharepoint_index_last_error"

SUPPORTED_EXTENSIONS = {".txt", ".md", ".csv", ".pdf", ".docx"}
MAX_TEXT_CHARS       = 2000   # chars stored per document (enough for semantic search)
MAX_FILES_PER_CLIENT = 100    # safety cap per client folder

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


# ── Text extraction helpers ───────────────────────────────────────────────────

def _extract_text_from_bytes(filename: str, content: bytes) -> str:
    """Extract plain text from file bytes based on extension."""
    ext = "." + filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

    if ext in (".txt", ".md", ".csv"):
        try:
            return content.decode("utf-8", errors="replace")
        except Exception:
            return ""

    if ext == ".pdf":
        try:
            from pdfminer.high_level import extract_text as pdf_extract
            return pdf_extract(io.BytesIO(content)) or ""
        except Exception as e:
            logger.debug("PDF extract failed for %s: %s", filename, e)
            return ""

    if ext == ".docx":
        try:
            import docx
            doc = docx.Document(io.BytesIO(content))
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        except Exception as e:
            logger.debug("DOCX extract failed for %s: %s", filename, e)
            return ""

    return ""


def _truncate(text: str, max_chars: int = MAX_TEXT_CHARS) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > max_chars:
        return text[:max_chars] + "…"
    return text


# ── Core indexer ─────────────────────────────────────────────────────────────

class SharePointIndexer:
    """Indexes SharePoint client files into MemoryStore."""

    def __init__(self, graph: "GraphClient", memory: "MemoryStore"):
        self._graph  = graph
        self._memory = memory

    # ── Public API ────────────────────────────────────────────────────────────

    def run(self, incremental: bool = True) -> dict:
        """Run the full index.

        Args:
            incremental: If True, only re-index files modified since last run.
                         If False, re-index everything (full rebuild).

        Returns:
            dict with keys: indexed, skipped, errors, duration_seconds
        """
        set_setting(SETTING_STATUS, "running")
        set_setting(SETTING_LAST_ERROR, "")
        start = time.time()

        last_run_iso = get_setting(SETTING_LAST_RUN, "") if incremental else ""
        last_run_dt  = _parse_iso(last_run_iso)

        stats = {"indexed": 0, "skipped": 0, "errors": 0}

        try:
            site_id  = self._graph.get_sharepoint_site_id()
            drive_id = self._graph.get_sharepoint_drive_id(site_id) if site_id else None
            if not site_id or not drive_id:
                raise RuntimeError("SharePoint not connected — cannot index")

            # List all client folders under Server/Clients/
            from graph_client import SHAREPOINT_CLIENT_BASE
            client_folders = self._list_children(drive_id, SHAREPOINT_CLIENT_BASE)
            logger.info("[Indexer] Found %d client folders", len(client_folders))

            for client_folder in client_folders:
                if not client_folder.get("is_folder"):
                    continue
                client_name = client_folder["name"]
                self._index_client_folder(
                    drive_id, SHAREPOINT_CLIENT_BASE,
                    client_name, entity_name=None,
                    last_run_dt=last_run_dt, stats=stats,
                )

            # Persist results
            now_iso = datetime.now(timezone.utc).isoformat()
            set_setting(SETTING_LAST_RUN,   now_iso)
            set_setting(SETTING_LAST_COUNT, str(stats["indexed"]))
            set_setting(SETTING_STATUS,     "idle")
            logger.info(
                "[Indexer] Done — indexed=%d skipped=%d errors=%d in %.1fs",
                stats["indexed"], stats["skipped"], stats["errors"],
                time.time() - start,
            )

        except Exception as e:
            logger.error("[Indexer] Failed: %s", e, exc_info=True)
            set_setting(SETTING_STATUS,     "error")
            set_setting(SETTING_LAST_ERROR, str(e))
            stats["errors"] += 1

        stats["duration_seconds"] = round(time.time() - start, 1)
        return stats

    # ── Private helpers ───────────────────────────────────────────────────────

    def _index_client_folder(
        self,
        drive_id: str,
        base_path: str,
        client_name: str,
        entity_name: Optional[str],
        last_run_dt: Optional[datetime],
        stats: dict,
        depth: int = 0,
    ):
        """Recursively index a client (or entity) folder."""
        if depth > 3:
            return  # safety — don't recurse too deep

        folder_path = f"{base_path}/{client_name}"
        if entity_name:
            folder_path += f"/{entity_name}"

        items = self._list_children_by_path(drive_id, folder_path)
        file_count = 0

        for item in items:
            if file_count >= MAX_FILES_PER_CLIENT:
                logger.debug("[Indexer] Hit file cap for %s/%s", client_name, entity_name or "")
                break

            name = item.get("name", "")
            is_folder = item.get("is_folder", False)

            if is_folder:
                # Recurse into entity sub-folders (one level down from client)
                if depth == 0:
                    self._index_client_folder(
                        drive_id, f"{base_path}/{client_name}",
                        name, entity_name=None,
                        last_run_dt=last_run_dt, stats=stats,
                        depth=depth + 1,
                    )
                continue

            ext = "." + name.rsplit(".", 1)[-1].lower() if "." in name else ""
            if ext not in SUPPORTED_EXTENSIONS:
                stats["skipped"] += 1
                continue

            # Skip if not modified since last run
            modified_iso = item.get("modified", "")
            if last_run_dt and modified_iso:
                modified_dt = _parse_iso(modified_iso)
                if modified_dt and modified_dt <= last_run_dt:
                    stats["skipped"] += 1
                    continue

            # Download and index
            try:
                content = self._download_file(drive_id, folder_path, name)
                if not content:
                    stats["skipped"] += 1
                    continue

                text = _extract_text_from_bytes(name, content)
                if not text.strip():
                    stats["skipped"] += 1
                    continue

                text = _truncate(text)
                doc_id = f"sp_{_safe_id(client_name)}_{_safe_id(name)}"

                self._memory._documents.upsert(
                    ids=[doc_id],
                    documents=[text],
                    metadatas=[{
                        "source":           "sharepoint",
                        "client_name":      client_name,
                        "entity_name":      entity_name or "",
                        "filename":         name,
                        "folder_path":      folder_path,
                        "sharepoint_url":   item.get("url", ""),
                        "modified":         modified_iso,
                        "indexed_at":       datetime.now(timezone.utc).isoformat(),
                        "epoch":            int(time.time()),
                    }],
                )
                stats["indexed"] += 1
                file_count += 1
                logger.debug("[Indexer] Indexed: %s / %s", client_name, name)

            except Exception as e:
                logger.warning("[Indexer] Error indexing %s/%s: %s", client_name, name, e)
                stats["errors"] += 1

    def _list_children(self, drive_id: str, folder_path: str) -> list:
        """List immediate children of a folder path."""
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{folder_path}:/children?$top=200"
        result = self._graph._make_request("GET", url)
        if not result or "value" not in result:
            return []
        return [
            {
                "name":      item["name"],
                "size":      item.get("size", 0),
                "modified":  item.get("lastModifiedDateTime", ""),
                "url":       item.get("webUrl", ""),
                "is_folder": "folder" in item,
                "item_id":   item.get("id", ""),
            }
            for item in result["value"]
        ]

    def _list_children_by_path(self, drive_id: str, folder_path: str) -> list:
        return self._list_children(drive_id, folder_path)

    def _download_file(self, drive_id: str, folder_path: str, filename: str) -> Optional[bytes]:
        """Download a file's raw bytes from SharePoint."""
        file_path = f"{folder_path}/{filename}"
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{file_path}:/content"
        try:
            import requests as _requests
            headers = self._graph._headers()
            r = _requests.get(url, headers=headers, timeout=30, allow_redirects=True)
            if r.status_code == 200:
                return r.content
            logger.debug("[Indexer] Download %s returned %s", file_path, r.status_code)
            return None
        except Exception as e:
            logger.warning("[Indexer] Download failed for %s: %s", file_path, e)
            return None


# ── Status helper ─────────────────────────────────────────────────────────────

def get_index_status() -> dict:
    """Return current index status for the Settings UI."""
    return {
        "status":      get_setting(SETTING_STATUS, "idle"),
        "last_run":    get_setting(SETTING_LAST_RUN, ""),
        "last_count":  int(get_setting(SETTING_LAST_COUNT, "0") or "0"),
        "last_error":  get_setting(SETTING_LAST_ERROR, ""),
    }


# ── Scheduler helper ──────────────────────────────────────────────────────────

def is_index_due() -> bool:
    """Return True if the index should run now.

    Triggers when:
    1. Never run before, OR
    2. Last run was before the most recent Monday 10am (catch-up logic).
    """
    last_run_iso = get_setting(SETTING_LAST_RUN, "")
    if not last_run_iso:
        return True  # never run

    last_run_dt = _parse_iso(last_run_iso)
    if not last_run_dt:
        return True

    now = datetime.now(timezone.utc)
    # Find the most recent Monday 10:00 UTC
    days_since_monday = now.weekday()  # 0 = Monday
    last_monday = now.replace(hour=10, minute=0, second=0, microsecond=0)
    last_monday = last_monday.replace(
        day=last_monday.day - days_since_monday
    ) if days_since_monday > 0 else last_monday

    # If it's Monday but before 10am, the "last Monday 10am" was 7 days ago
    if now.weekday() == 0 and now.hour < 10:
        from datetime import timedelta
        last_monday -= timedelta(days=7)

    return last_run_dt < last_monday


# ── Utilities ─────────────────────────────────────────────────────────────────

def _parse_iso(iso: str) -> Optional[datetime]:
    if not iso:
        return None
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        return None


def _safe_id(s: str) -> str:
    """Convert a string to a safe ChromaDB document ID component."""
    return re.sub(r"[^a-zA-Z0-9_-]", "_", s)[:40]
