"""
MemoryStore — ChromaDB-backed semantic memory for MCS CoWorker.

Provides persistent vector storage for client interactions, lessons learned,
and documents. Supports semantic search (meaning-based, not just keyword
matching) so plugins can retrieve relevant context when drafting emails or
processing documents.

Collections:
  - client_interactions: Emails sent/received, NOAs processed, debtor
    follow-ups, FuseSign events, correspondence — anything involving a client.
  - lessons: Patterns learned from the approval queue (e.g., "Ross always
    edits the greeting on debtor emails for clients over 60 days").
  - documents: Stored document content (knowledge base entries, templates).
"""
from __future__ import annotations

import logging
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


class MemoryStore:
    """ChromaDB-backed semantic memory store."""

    def __init__(self, persist_dir: Optional[str] = None):
        """Initialise the memory store.

        Args:
            persist_dir: Directory for ChromaDB persistence. If None, uses
                         DATA_DIR / "memory" from config.
        """
        import chromadb

        if persist_dir is None:
            from config import DATA_DIR
            persist_dir = str(DATA_DIR / "memory")

        Path(persist_dir).mkdir(parents=True, exist_ok=True)

        self._client = chromadb.PersistentClient(path=persist_dir)

        # ChromaDB's default embedding function (all-MiniLM-L6-v2 via ONNX)
        # is used automatically. The model is ~80MB and downloaded on first
        # use to ~/.cache/chroma/onnx_models/ — warn the user so a blank
        # startup log isn't mistaken for a hang.
        model_cache = Path.home() / ".cache" / "chroma" / "onnx_models" / "all-MiniLM-L6-v2"
        if not model_cache.exists():
            logger.info("Downloading embedding model (first run only, ~80MB)...")

        self._interactions = self._client.get_or_create_collection(
            name="client_interactions",
            metadata={"description": "Client emails, NOAs, debtor follow-ups, correspondence"},
        )
        self._lessons = self._client.get_or_create_collection(
            name="lessons",
            metadata={"description": "Patterns learned from approval queue edits"},
        )
        self._documents = self._client.get_or_create_collection(
            name="documents",
            metadata={"description": "Knowledge base entries, templates, reference docs"},
        )

        logger.info(
            "MemoryStore initialised (ChromaDB at %s) — "
            "%d interactions, %d lessons, %d documents",
            persist_dir,
            self._interactions.count(),
            self._lessons.count(),
            self._documents.count(),
        )

    # ── Client Interactions ─────────────────────────────────────────────

    def store_client_interaction(
        self,
        # Accept BOTH calling conventions in use by plugins:
        #   NOA-style:    (content=, client_email=, interaction_type=, extra_meta=)
        #   Plugin-style: (client_name=, interaction_type=, summary=, metadata=)
        content: Optional[str] = None,
        client_name: Optional[str] = None,
        client_email: Optional[str] = None,
        interaction_type: str = "email",
        summary: Optional[str] = None,
        extra_meta: Optional[Dict] = None,
        metadata: Optional[Dict] = None,
        **_kwargs,
    ) -> str:
        """Store a client interaction for semantic retrieval.

        Returns the generated document ID, or "" on failure.
        """
        doc_text = content or summary or ""
        if not doc_text:
            logger.warning("store_client_interaction called with empty content — skipping")
            return ""

        meta: Dict[str, Any] = {
            "interaction_type": interaction_type,
            "timestamp": datetime.now().isoformat(),
            "epoch": int(time.time()),
        }

        if client_name:
            meta["client_name"] = client_name
        if client_email:
            meta["client_email"] = client_email
            if not client_name:
                # Derive a readable client_name from email for later filtering.
                meta["client_name"] = client_email.split("@")[0].replace(".", " ").title()

        extra = extra_meta or metadata or {}
        for k, v in extra.items():
            meta[k] = v if isinstance(v, (str, int, float, bool)) else str(v)

        doc_id = f"int_{uuid.uuid4().hex[:12]}"

        try:
            self._interactions.add(
                ids=[doc_id],
                documents=[doc_text],
                metadatas=[meta],
            )
            logger.debug(
                "Stored interaction %s: %s (%s) — %d chars",
                doc_id, meta.get("client_name", "unknown"),
                interaction_type, len(doc_text),
            )
        except Exception as e:
            logger.error("Failed to store interaction: %s", e, exc_info=True)
            return ""

        return doc_id

    def search(
        self,
        query: str,
        n_results: int = 5,
        filters: Optional[Dict] = None,
        collection: str = "interactions",
        where: Optional[Dict] = None,
        **_kwargs,
    ) -> List[Dict[str, Any]]:
        """Semantic search across stored memories.

        Args:
            query: Natural language search query.
            n_results: Max results to return.
            filters: ChromaDB metadata filters, e.g. {"client_name": "Chen Family Trust"}.
                     Accepts `where=` as an alias for backwards-compat.
            collection: "interactions", "lessons", or "documents".

        Returns:
            List of dicts with keys: id, content, metadata, distance, relevance.
        """
        col = self._get_collection(collection)
        total = col.count()
        if total == 0:
            return []

        n_results = min(n_results, total)
        effective_filters = filters if filters is not None else where

        try:
            results = col.query(
                query_texts=[query],
                n_results=n_results,
                where=effective_filters,
            )
        except Exception as e:
            logger.error("Memory search failed: %s", e, exc_info=True)
            return []

        output: List[Dict[str, Any]] = []
        if results and results.get("ids") and results["ids"][0]:
            ids = results["ids"][0]
            docs = results.get("documents", [[]])[0] or []
            metas = results.get("metadatas", [[]])[0] or []
            dists = results.get("distances", [[]])[0] or []
            for i, doc_id in enumerate(ids):
                distance = dists[i] if i < len(dists) else 0.0
                output.append({
                    "id": doc_id,
                    "content": docs[i] if i < len(docs) else "",
                    "metadata": metas[i] if i < len(metas) else {},
                    "distance": distance,
                    "relevance": round(max(0.0, 1.0 - distance), 2),
                })
        return output

    def get_client_history(
        self,
        client_name: str,
        limit: int = 10,
    ) -> List[Dict[str, Any]]:
        """All stored interactions for a client, sorted by recency (newest first)."""
        if self._interactions.count() == 0:
            return []

        try:
            results = self._interactions.get(
                where={"client_name": client_name},
                limit=limit,
            )
        except Exception as e:
            logger.error("get_client_history failed: %s", e, exc_info=True)
            return []

        output: List[Dict[str, Any]] = []
        if results and results.get("ids"):
            ids = results["ids"]
            docs = results.get("documents") or []
            metas = results.get("metadatas") or []
            for i, doc_id in enumerate(ids):
                output.append({
                    "id": doc_id,
                    "content": docs[i] if i < len(docs) else "",
                    "metadata": metas[i] if i < len(metas) else {},
                })
            output.sort(
                key=lambda x: x.get("metadata", {}).get("epoch", 0),
                reverse=True,
            )
        return output[:limit]

    def get_client_context(
        self,
        client_name: Optional[str] = None,
        client_email: Optional[str] = None,
        query: Optional[str] = None,
        n_results: int = 5,
        limit: Optional[int] = None,
        **_kwargs,
    ) -> List[Dict[str, Any]]:
        """Context for a client — semantic search if `query` given, else recent history.

        Accepts either `client_name` or `client_email` (the latter is derived
        to a name using the same rule as `store_client_interaction`).
        """
        if limit is not None:
            n_results = limit

        resolved_name = client_name
        if not resolved_name and client_email:
            resolved_name = client_email.split("@")[0].replace(".", " ").title()
        if not resolved_name:
            return []

        if query:
            return self.search(
                query=query,
                n_results=n_results,
                filters={"client_name": resolved_name},
                collection="interactions",
            )
        return self.get_client_history(resolved_name, n_results)

    # ── Lessons (from approval queue) ───────────────────────────────────

    def store_lesson(
        self,
        # Accept BOTH shapes: (lesson=, category=, metadata=) and (content=, ...)
        content: Optional[str] = None,
        lesson: Optional[str] = None,
        category: Optional[str] = None,
        plugin_name: Optional[str] = None,
        source: Optional[str] = None,
        metadata: Optional[Dict] = None,
        **_kwargs,
    ) -> str:
        """Store a lesson learned from an approval queue edit."""
        text = content or lesson or ""
        if not text:
            return ""

        meta: Dict[str, Any] = {
            "timestamp": datetime.now().isoformat(),
            "epoch": int(time.time()),
        }
        if category:
            meta["category"] = category
        if plugin_name:
            meta["plugin_name"] = plugin_name
        if source:
            meta["source"] = source
        if metadata:
            for k, v in metadata.items():
                meta[k] = v if isinstance(v, (str, int, float, bool)) else str(v)

        doc_id = f"les_{uuid.uuid4().hex[:12]}"

        try:
            self._lessons.add(
                ids=[doc_id],
                documents=[text],
                metadatas=[meta],
            )
            logger.debug("Stored lesson %s: %s", doc_id, text[:80])
        except Exception as e:
            logger.error("Failed to store lesson: %s", e, exc_info=True)
            return ""

        # Mirror to SQLite so the lessons UI (which reads from SQLite) still sees it.
        try:
            from config import add_lesson
            add_lesson(text, source or category or "")
        except Exception:
            pass

        return doc_id

    # ── Documents (knowledge base) ──────────────────────────────────────

    def store_document(
        self,
        # Accept (content=, ...) or plugin shorthand (document=, ...)
        content: Optional[str] = None,
        document: Optional[str] = None,
        title: Optional[str] = None,
        doc_type: Optional[str] = None,
        client_email: Optional[str] = None,
        filename: Optional[str] = None,
        metadata: Optional[Dict] = None,
        **_kwargs,
    ) -> str:
        """Store a document or knowledge base entry for retrieval."""
        text = content or document or ""
        if not text:
            return ""

        meta: Dict[str, Any] = {
            "timestamp": datetime.now().isoformat(),
            "epoch": int(time.time()),
        }
        if title:
            meta["title"] = title
        if doc_type:
            meta["doc_type"] = doc_type
        if client_email:
            meta["client_email"] = client_email
        if filename:
            meta["filename"] = filename
        if metadata:
            for k, v in metadata.items():
                meta[k] = v if isinstance(v, (str, int, float, bool)) else str(v)

        doc_id = f"doc_{uuid.uuid4().hex[:12]}"

        try:
            self._documents.add(
                ids=[doc_id],
                documents=[text],
                metadatas=[meta],
            )
            logger.debug("Stored document %s: %s", doc_id, (title or text[:40]))
        except Exception as e:
            logger.error("Failed to store document: %s", e, exc_info=True)
            return ""

        return doc_id

    # ── Generic helpers ─────────────────────────────────────────────────

    def store(
        self,
        key: Optional[str] = None,
        value: Any = None,
        # Plugin shorthand: context.memory.store(document=text, metadata={...})
        document: Optional[str] = None,
        content: Optional[str] = None,
        metadata: Optional[Dict] = None,
        **_kwargs,
    ) -> str:
        """Generic key-value store (persisted in the documents collection)."""
        text = document or content or (str(value) if value is not None else "")
        if not text:
            return ""
        return self.store_document(
            content=text,
            title=key,
            metadata=metadata,
        )

    def get(self, key: str) -> Optional[str]:
        """Generic key-value get (searches documents by title)."""
        try:
            results = self._documents.get(where={"title": key}, limit=1)
            docs = results.get("documents") if results else None
            if docs:
                return docs[0]
        except Exception:
            pass
        return None

    def delete(self, doc_id: str = "", collection: Optional[str] = None, **_kwargs) -> bool:
        """Delete a document by ID. If collection is given, only try that one."""
        if not doc_id:
            return False
        cols = [self._get_collection(collection)] if collection else [
            self._interactions, self._lessons, self._documents
        ]
        for col in cols:
            try:
                col.delete(ids=[doc_id])
                return True
            except Exception:
                continue
        return False

    def count(self, collection: str = "interactions") -> int:
        """Return the number of documents in a collection."""
        return self._get_collection(collection).count()

    def format_for_prompt(
        self,
        query: str,
        n_results: int = 5,
        collection: str = "interactions",
    ) -> str:
        """Search and format results as a string block for prompt injection.

        Returns "" if no results.
        """
        results = self.search(query, n_results=n_results, collection=collection)
        if not results:
            return ""

        lines = ["Relevant context from memory:"]
        for r in results:
            meta = r.get("metadata", {})
            client = meta.get("client_name", "Unknown")
            itype = meta.get("interaction_type", "")
            ts = meta.get("timestamp", "")
            lines.append(f"- [{client}] ({itype}, {ts}): {r['content'][:200]}")
        return "\n".join(lines)

    # ── Internal ────────────────────────────────────────────────────────

    def _get_collection(self, name: Optional[str]):
        """Map collection short names to ChromaDB collection objects."""
        mapping = {
            "interactions": self._interactions,
            "client_interactions": self._interactions,
            "lessons": self._lessons,
            "documents": self._documents,
        }
        return mapping.get(name or "interactions", self._interactions)
