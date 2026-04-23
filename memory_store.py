"""
MemoryStore — no-op stub (chromadb is explicitly forbidden in requirements.txt).

Plugins call ``context.memory.search(...)``, ``context.memory.store_*(...)``,
``context.memory.get_client_context(...)``, etc. Without this stub those
calls would raise ``ImportError`` at load time (chromadb missing) or
``AttributeError`` / ``TypeError`` at call time. This module provides a
concrete ``MemoryStore`` class where every method is a no-op with the right
signature so plugins can run; reads return empty lists / empty strings so the
plugin-side for-loops and ``if history:`` checks naturally skip.

``plugin_loader.py`` wires the bare class (not an instance) onto
``context.memory``, so every public method is a ``@staticmethod``.

TODO: Replace with a real implementation (SQLite FTS5, or chromadb once/if
the dependency is approved) when memory-backed plugins actually need history.
"""
from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


DEFAULT_COLLECTION = "general"


class MemoryStore:
    """No-op memory store. Writes are discarded, reads return empty."""

    def __init__(self, persist_dir: Optional[str] = None):
        # Kept for backwards compat — some callers may instantiate. Most use
        # the class directly (via plugin_loader).
        logger.info("MemoryStore initialised (stub — no persistence)")

    # ── Writes ────────────────────────────────────────────────────────────────

    @staticmethod
    def store(
        content: str = "",
        metadata: Optional[Dict[str, Any]] = None,
        collection: str = DEFAULT_COLLECTION,
        doc_id: Optional[str] = None,
        **_kwargs,
    ) -> str:
        logger.debug(
            "MemoryStore.store stub: collection=%s len(content)=%d",
            collection, len(content or ""),
        )
        return ""

    @staticmethod
    def store_client_interaction(
        content: str = "",
        client_email: str = "",
        interaction_type: str = "general",
        extra_meta: Optional[Dict[str, Any]] = None,
        # Plugins also pass these kwarg shapes; tolerate both.
        client_name: str = "",
        summary: str = "",
        metadata: Optional[Dict[str, Any]] = None,
        **_kwargs,
    ) -> str:
        logger.debug(
            "MemoryStore.store_client_interaction stub: type=%s client=%s",
            interaction_type, client_email or client_name,
        )
        return ""

    @staticmethod
    def store_lesson(
        lesson: str = "",
        source: str = "",
        category: str = "",
        metadata: Optional[Dict[str, Any]] = None,
        **_kwargs,
    ) -> str:
        # Mirror to the SQLite lessons table so UI-visible lessons survive
        # the stub — the real chromadb layer did this too.
        try:
            if lesson:
                from config import add_lesson
                add_lesson(lesson, source or category)
        except Exception:
            pass
        return ""

    @staticmethod
    def store_document(
        content: str = "",
        doc_type: str = "general",
        client_email: str = "",
        filename: str = "",
        **_kwargs,
    ) -> str:
        return ""

    # ── Reads ─────────────────────────────────────────────────────────────────

    @staticmethod
    def search(
        query: str = "",
        collection: str = DEFAULT_COLLECTION,
        n_results: int = 5,
        where: Optional[Dict[str, Any]] = None,
        **_kwargs,
    ) -> List[Any]:
        logger.debug(
            "MemoryStore.search stub: collection=%s query=%r n=%d",
            collection, (query or "")[:40], n_results,
        )
        # Return [] so plugin-side ``if history:`` and ``for r in history``
        # loops naturally no-op.
        return []

    @staticmethod
    def get_client_context(
        client_email: str = "",
        query: str = "",
        n_results: int = 5,
        **_kwargs,
    ) -> str:
        return ""

    @staticmethod
    def count(collection: str = DEFAULT_COLLECTION) -> int:
        return 0

    # ── Formatting / housekeeping ────────────────────────────────────────────

    @staticmethod
    def format_for_prompt(
        results: List[Any],
        max_chars: int = 2000,
        header: str = "RELEVANT MEMORY:",
    ) -> str:
        return ""

    @staticmethod
    def delete(doc_id: str = "", collection: str = DEFAULT_COLLECTION) -> bool:
        return False

    # Generic key/value convenience — some older plugin scaffolds reach for these.
    @staticmethod
    def get(key: str) -> Optional[Any]:
        return None
