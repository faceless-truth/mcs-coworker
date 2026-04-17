"""
MC & S Desktop Agent — Vector Memory Store (Stream 2 — APEX Upgrade)
=====================================================================
Provides semantic (vector) memory alongside the existing SQLite relational
store in config.py.  Uses ChromaDB running embedded/locally — no server,
no Docker, no internet required after first model download.

COLLECTIONS
-----------
  "client_interactions"  — emails sent/received, NOA outcomes, outreach history
  "lessons"              — extracted lessons from user feedback (mirrors memory_lessons)
  "documents"            — processed PDF content (NOAs, ASIC returns, etc.)
  "general"              — catch-all for anything that doesn't fit above

USAGE IN PLUGINS
----------------
    from memory_store import MemoryStore

    # Store something
    MemoryStore.store(
        content="Client John Smith received a $2,400 refund for FY2024",
        metadata={"client_email": "john@example.com", "type": "noa_outcome", "fy": "2024"},
        collection="client_interactions"
    )

    # Retrieve semantically relevant memories
    results = MemoryStore.search(
        query="John Smith tax refund history",
        collection="client_interactions",
        n_results=5
    )
    for doc, meta, distance in results:
        print(f"  [{distance:.2f}] {doc}  |  {meta}")

    # Or use the context helper (context.memory)
    results = context.memory.search("John Smith previous contact", n_results=3)

DESIGN NOTES
------------
- IDs are deterministic SHA-256 hashes of the content so duplicate
  documents are silently ignored (upsert behaviour).
- Metadata values must be str, int, float, or bool — no nested dicts.
- The ChromaDB path is ~/.mcs_email_automation/chroma_db/
- Thread-safe: ChromaDB PersistentClient is safe for concurrent reads;
  writes are serialised via a module-level lock.
"""

import hashlib
import threading
import time
from pathlib import Path
from typing import Any

# ── ChromaDB path ─────────────────────────────────────────────────────────────

CHROMA_PATH = Path.home() / ".mcs_email_automation" / "chroma_db"

# Valid collection names
COLLECTIONS = {
    "client_interactions",
    "lessons",
    "documents",
    "general",
}

DEFAULT_COLLECTION = "general"

# ── Module-level state ────────────────────────────────────────────────────────

_client = None
_client_lock = threading.Lock()


def _get_client():
    """Return the singleton ChromaDB PersistentClient, creating it if needed."""
    global _client
    if _client is not None:
        return _client
    with _client_lock:
        if _client is not None:
            return _client
        try:
            import chromadb
            CHROMA_PATH.mkdir(parents=True, exist_ok=True)
            _client = chromadb.PersistentClient(path=str(CHROMA_PATH))
        except ImportError:
            raise ImportError(
                "ChromaDB is not installed. Run: pip install chromadb>=0.4.0"
            )
    return _client


def _make_id(content: str) -> str:
    """Generate a deterministic ID from content so duplicates are ignored."""
    return hashlib.sha256(content.encode("utf-8")).hexdigest()[:32]


def _sanitise_metadata(metadata: dict) -> dict:
    """
    ChromaDB only accepts str/int/float/bool metadata values.
    Convert anything else to a string representation.
    """
    safe = {}
    for k, v in (metadata or {}).items():
        if isinstance(v, (str, int, float, bool)):
            safe[k] = v
        else:
            safe[k] = str(v)
    # Always stamp with a timestamp if not provided
    if "stored_at" not in safe:
        safe["stored_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    return safe


# ── Public API ────────────────────────────────────────────────────────────────

class MemoryStore:
    """
    Static helper class for semantic memory operations.
    All methods are class-level so plugins can call MemoryStore.store(...)
    without instantiation.
    """

    @staticmethod
    def store(
        content: str,
        metadata: dict = None,
        collection: str = DEFAULT_COLLECTION,
        doc_id: str = None,
    ) -> str:
        """
        Store a text document in the vector memory.

        Parameters
        ----------
        content    : The text to embed and store.
        metadata   : Optional dict of str/int/float/bool key-value pairs.
        collection : Which collection to store in (default: "general").
        doc_id     : Optional explicit ID. If omitted, a hash of content is used.

        Returns
        -------
        The document ID that was stored.
        """
        if not content or not content.strip():
            return ""

        client = _get_client()
        col = client.get_or_create_collection(collection)

        doc_id = doc_id or _make_id(content)
        safe_meta = _sanitise_metadata(metadata or {})

        # Upsert — safe to call multiple times with the same content
        col.upsert(
            documents=[content],
            metadatas=[safe_meta],
            ids=[doc_id],
        )
        return doc_id

    @staticmethod
    def search(
        query: str,
        collection: str = DEFAULT_COLLECTION,
        n_results: int = 5,
        where: dict = None,
    ) -> list[tuple[str, dict, float]]:
        """
        Semantic search across a collection.

        Parameters
        ----------
        query      : Natural language query string.
        collection : Which collection to search (default: "general").
        n_results  : Maximum number of results to return.
        where      : Optional ChromaDB metadata filter dict,
                     e.g. {"client_email": "john@example.com"}

        Returns
        -------
        List of (document_text, metadata_dict, distance) tuples,
        sorted by relevance (lowest distance = most relevant).
        Returns [] if the collection is empty or query fails.
        """
        if not query or not query.strip():
            return []

        try:
            client = _get_client()
            col = client.get_or_create_collection(collection)

            # Don't query an empty collection
            if col.count() == 0:
                return []

            # Clamp n_results to collection size
            actual_n = min(n_results, col.count())

            kwargs: dict[str, Any] = {
                "query_texts": [query],
                "n_results": actual_n,
                "include": ["documents", "metadatas", "distances"],
            }
            if where:
                kwargs["where"] = where

            results = col.query(**kwargs)

            docs = results.get("documents", [[]])[0]
            metas = results.get("metadatas", [[]])[0]
            distances = results.get("distances", [[]])[0]

            return list(zip(docs, metas, distances))

        except Exception:
            return []

    @staticmethod
    def delete(doc_id: str, collection: str = DEFAULT_COLLECTION) -> bool:
        """
        Delete a specific document by ID.

        Returns True if deleted, False if not found or error.
        """
        try:
            client = _get_client()
            col = client.get_or_create_collection(collection)
            col.delete(ids=[doc_id])
            return True
        except Exception:
            return False

    @staticmethod
    def count(collection: str = DEFAULT_COLLECTION) -> int:
        """Return the number of documents in a collection."""
        try:
            client = _get_client()
            col = client.get_or_create_collection(collection)
            return col.count()
        except Exception:
            return 0

    @staticmethod
    def format_for_prompt(
        results: list[tuple[str, dict, float]],
        max_chars: int = 2000,
        header: str = "RELEVANT MEMORY:",
    ) -> str:
        """
        Format search results into a clean string for injection into a
        Claude prompt.

        Parameters
        ----------
        results   : Output from MemoryStore.search()
        max_chars : Truncate total output to this many characters.
        header    : Section header to prepend.

        Returns
        -------
        A formatted string ready to inject into a prompt, or "" if no results.
        """
        if not results:
            return ""

        lines = [header]
        total = len(header)
        for doc, meta, distance in results:
            # Skip very low relevance results (high distance)
            if distance > 1.5:
                continue
            line = f"- {doc}"
            if total + len(line) > max_chars:
                break
            lines.append(line)
            total += len(line)

        return "\n".join(lines) if len(lines) > 1 else ""

    @staticmethod
    def store_client_interaction(
        content: str,
        client_email: str = "",
        interaction_type: str = "general",
        extra_meta: dict = None,
    ) -> str:
        """
        Convenience method: store a client interaction with standard metadata.

        interaction_type examples: "email_sent", "email_received", "noa_outcome",
                                   "asic_reminder", "outreach_draft", "tax_note"
        """
        meta = {
            "client_email": client_email,
            "type": interaction_type,
        }
        if extra_meta:
            meta.update(extra_meta)
        return MemoryStore.store(content, meta, collection="client_interactions")

    @staticmethod
    def store_lesson(lesson: str, source: str = "") -> str:
        """
        Convenience method: store a lesson in the 'lessons' collection.
        Also mirrors to the SQLite memory_lessons table for UI display.
        """
        try:
            from config import add_lesson
            add_lesson(lesson, source)
        except Exception:
            pass
        return MemoryStore.store(
            lesson,
            {"source": source, "type": "lesson"},
            collection="lessons",
        )

    @staticmethod
    def store_document(
        content: str,
        doc_type: str = "general",
        client_email: str = "",
        filename: str = "",
    ) -> str:
        """
        Convenience method: store processed document content (NOA, ASIC PDF, etc.)
        """
        return MemoryStore.store(
            content,
            {
                "doc_type": doc_type,
                "client_email": client_email,
                "filename": filename,
            },
            collection="documents",
        )

    @staticmethod
    def get_client_context(
        client_email: str,
        query: str = "",
        n_results: int = 5,
    ) -> str:
        """
        Retrieve all relevant memory for a specific client and format it
        for prompt injection.

        If query is provided, uses semantic search filtered to that client.
        If no query, returns the most recent interactions.
        """
        if not client_email:
            return ""

        search_query = query or f"history and context for {client_email}"
        results = MemoryStore.search(
            query=search_query,
            collection="client_interactions",
            n_results=n_results,
            where={"client_email": client_email} if client_email else None,
        )
        return MemoryStore.format_for_prompt(
            results,
            header=f"CLIENT HISTORY ({client_email}):",
        )
