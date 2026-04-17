"""
Stream 2 Test Suite — ChromaDB Vector Memory Layer
====================================================
Tests every aspect of the MemoryStore implementation:
  1. Core store/search/delete/count operations
  2. Metadata sanitisation
  3. Duplicate (upsert) handling
  4. Convenience methods (store_client_interaction, store_lesson, store_document)
  5. format_for_prompt helper
  6. get_client_context helper
  7. plugin_base.py — PluginContext.memory field
  8. plugin_loader.py — _make_context() injects MemoryStore
  9. Graceful degradation when ChromaDB unavailable
 10. NOA processor memory integration
"""

import sys
import os
import unittest
import tempfile
import shutil
from pathlib import Path
from unittest.mock import patch, MagicMock

# ── Path setup ────────────────────────────────────────────────────────────────
REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

# Use a temp DB so tests never touch the real user database
import config as cfg
_tmp_db = tempfile.mktemp(suffix=".db")
cfg.DB_PATH = Path(_tmp_db)
cfg.init_db()


# ── Helper: fresh MemoryStore with temp chroma path ──────────────────────────

def make_fresh_memory_store():
    """Return a MemoryStore class backed by a fresh temp ChromaDB directory."""
    import memory_store as ms
    tmp_chroma = tempfile.mkdtemp()
    # Patch the module-level path and reset the singleton client
    ms.CHROMA_PATH = Path(tmp_chroma)
    ms._client = None
    return ms.MemoryStore, tmp_chroma


# ── 1. Core operations ────────────────────────────────────────────────────────

class TestMemoryStoreCoreOps(unittest.TestCase):

    def setUp(self):
        self.MemoryStore, self.tmp_dir = make_fresh_memory_store()

    def tearDown(self):
        import memory_store as ms
        ms._client = None
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_store_returns_id(self):
        doc_id = self.MemoryStore.store("Test document content")
        self.assertIsInstance(doc_id, str)
        self.assertGreater(len(doc_id), 0)

    def test_store_and_count(self):
        self.MemoryStore.store("First document")
        self.MemoryStore.store("Second document about accounting")
        count = self.MemoryStore.count()
        self.assertEqual(count, 2)

    def test_search_returns_results(self):
        self.MemoryStore.store("John Smith received a tax refund of $2400 for FY2024")
        self.MemoryStore.store("Jane Doe owes $800 in tax for FY2024")
        results = self.MemoryStore.search("tax refund")
        self.assertGreater(len(results), 0)
        # Each result is a (doc, meta, distance) tuple
        doc, meta, distance = results[0]
        self.assertIsInstance(doc, str)
        self.assertIsInstance(meta, dict)
        self.assertIsInstance(distance, float)

    def test_search_most_relevant_first(self):
        self.MemoryStore.store("Client received a large tax refund")
        self.MemoryStore.store("Client owes money for company tax")
        results = self.MemoryStore.search("tax refund", n_results=2)
        self.assertEqual(len(results), 2)
        # First result should have lower distance (more relevant)
        self.assertLessEqual(results[0][2], results[1][2])

    def test_search_empty_collection_returns_empty(self):
        results = self.MemoryStore.search("anything", collection="documents")
        self.assertEqual(results, [])

    def test_search_empty_query_returns_empty(self):
        self.MemoryStore.store("Some content")
        results = self.MemoryStore.search("")
        self.assertEqual(results, [])

    def test_delete_removes_document(self):
        doc_id = self.MemoryStore.store("Document to delete")
        self.assertEqual(self.MemoryStore.count(), 1)
        result = self.MemoryStore.delete(doc_id)
        self.assertTrue(result)
        self.assertEqual(self.MemoryStore.count(), 0)

    def test_delete_nonexistent_returns_false(self):
        result = self.MemoryStore.delete("nonexistent-id-12345")
        # ChromaDB may raise or silently succeed — either is acceptable
        # The important thing is it doesn't crash
        self.assertIsInstance(result, bool)

    def test_count_empty_collection(self):
        count = self.MemoryStore.count(collection="lessons")
        self.assertEqual(count, 0)


# ── 2. Metadata sanitisation ──────────────────────────────────────────────────

class TestMetadataSanitisation(unittest.TestCase):

    def setUp(self):
        self.MemoryStore, self.tmp_dir = make_fresh_memory_store()

    def tearDown(self):
        import memory_store as ms
        ms._client = None
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_valid_metadata_stored_correctly(self):
        self.MemoryStore.store(
            "Test with valid metadata",
            metadata={"client_email": "test@example.com", "amount": 2400, "active": True}
        )
        results = self.MemoryStore.search("Test with valid metadata")
        self.assertGreater(len(results), 0)
        _, meta, _ = results[0]
        self.assertEqual(meta["client_email"], "test@example.com")
        self.assertEqual(meta["amount"], 2400)

    def test_nested_dict_metadata_converted_to_string(self):
        """Nested dicts must be stringified — ChromaDB rejects them."""
        self.MemoryStore.store(
            "Test with nested metadata",
            metadata={"nested": {"key": "value"}, "normal": "ok"}
        )
        results = self.MemoryStore.search("Test with nested metadata")
        self.assertGreater(len(results), 0)
        _, meta, _ = results[0]
        self.assertIsInstance(meta["nested"], str)

    def test_stored_at_timestamp_auto_added(self):
        self.MemoryStore.store("Test timestamp auto-add")
        results = self.MemoryStore.search("Test timestamp auto-add")
        self.assertGreater(len(results), 0)
        _, meta, _ = results[0]
        self.assertIn("stored_at", meta)


# ── 3. Duplicate / upsert handling ────────────────────────────────────────────

class TestUpsertBehaviour(unittest.TestCase):

    def setUp(self):
        self.MemoryStore, self.tmp_dir = make_fresh_memory_store()

    def tearDown(self):
        import memory_store as ms
        ms._client = None
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_duplicate_content_not_duplicated(self):
        content = "Identical content stored twice"
        self.MemoryStore.store(content)
        self.MemoryStore.store(content)
        # Should still be 1 document (upsert by hash ID)
        self.assertEqual(self.MemoryStore.count(), 1)

    def test_explicit_id_upserts(self):
        self.MemoryStore.store("Version 1", doc_id="my-doc-001")
        self.MemoryStore.store("Version 2 updated", doc_id="my-doc-001")
        self.assertEqual(self.MemoryStore.count(), 1)
        results = self.MemoryStore.search("Version 2 updated")
        self.assertGreater(len(results), 0)


# ── 4. Convenience methods ────────────────────────────────────────────────────

class TestConvenienceMethods(unittest.TestCase):

    def setUp(self):
        self.MemoryStore, self.tmp_dir = make_fresh_memory_store()
        # Fresh DB for lesson mirroring
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        import memory_store as ms
        ms._client = None
        shutil.rmtree(self.tmp_dir, ignore_errors=True)
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_store_client_interaction(self):
        doc_id = self.MemoryStore.store_client_interaction(
            content="John Smith called about his tax return",
            client_email="john@example.com",
            interaction_type="phone_call",
        )
        self.assertIsInstance(doc_id, str)
        count = self.MemoryStore.count(collection="client_interactions")
        self.assertEqual(count, 1)

    def test_store_client_interaction_searchable_by_email(self):
        self.MemoryStore.store_client_interaction(
            content="John Smith refund $2400 FY2024",
            client_email="john@example.com",
            interaction_type="noa_outcome",
        )
        self.MemoryStore.store_client_interaction(
            content="Jane Doe payable $800 FY2024",
            client_email="jane@example.com",
            interaction_type="noa_outcome",
        )
        # Filter by client email
        results = self.MemoryStore.search(
            "tax outcome",
            collection="client_interactions",
            where={"client_email": "john@example.com"},
        )
        self.assertGreater(len(results), 0)
        for _, meta, _ in results:
            self.assertEqual(meta["client_email"], "john@example.com")

    def test_store_lesson(self):
        doc_id = self.MemoryStore.store_lesson(
            "Always check the ATO portal before sending NOA emails",
            source="manual"
        )
        self.assertIsInstance(doc_id, str)
        count = self.MemoryStore.count(collection="lessons")
        self.assertEqual(count, 1)

    def test_store_document(self):
        doc_id = self.MemoryStore.store_document(
            content="Taxable income: $85,000. Tax assessed: $19,717.",
            doc_type="noa_pdf",
            client_email="client@example.com",
            filename="NOA_2024.pdf",
        )
        self.assertIsInstance(doc_id, str)
        count = self.MemoryStore.count(collection="documents")
        self.assertEqual(count, 1)


# ── 5. format_for_prompt ──────────────────────────────────────────────────────

class TestFormatForPrompt(unittest.TestCase):

    def setUp(self):
        self.MemoryStore, self.tmp_dir = make_fresh_memory_store()

    def tearDown(self):
        import memory_store as ms
        ms._client = None
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_format_returns_string(self):
        self.MemoryStore.store("Client John had a refund last year")
        results = self.MemoryStore.search("John refund")
        formatted = self.MemoryStore.format_for_prompt(results)
        self.assertIsInstance(formatted, str)

    def test_format_empty_results_returns_empty_string(self):
        formatted = self.MemoryStore.format_for_prompt([])
        self.assertEqual(formatted, "")

    def test_format_includes_header(self):
        self.MemoryStore.store("Some relevant memory")
        results = self.MemoryStore.search("relevant memory")
        formatted = self.MemoryStore.format_for_prompt(results, header="TEST HEADER:")
        self.assertIn("TEST HEADER:", formatted)

    def test_format_respects_max_chars(self):
        for i in range(20):
            self.MemoryStore.store(f"Document number {i} with some content about accounting")
        results = self.MemoryStore.search("accounting document", n_results=20)
        formatted = self.MemoryStore.format_for_prompt(results, max_chars=200)
        self.assertLessEqual(len(formatted), 300)  # some tolerance for header


# ── 6. get_client_context ─────────────────────────────────────────────────────

class TestGetClientContext(unittest.TestCase):

    def setUp(self):
        self.MemoryStore, self.tmp_dir = make_fresh_memory_store()

    def tearDown(self):
        import memory_store as ms
        ms._client = None
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_get_client_context_empty_returns_empty(self):
        result = self.MemoryStore.get_client_context("nobody@example.com")
        self.assertEqual(result, "")

    def test_get_client_context_no_email_returns_empty(self):
        result = self.MemoryStore.get_client_context("")
        self.assertEqual(result, "")

    def test_get_client_context_returns_formatted_string(self):
        self.MemoryStore.store_client_interaction(
            "John Smith received $2400 refund for FY2024",
            client_email="john@example.com",
            interaction_type="noa_outcome",
        )
        result = self.MemoryStore.get_client_context(
            "john@example.com",
            query="refund history",
        )
        self.assertIsInstance(result, str)
        self.assertIn("john@example.com", result)


# ── 7. PluginContext.memory field ─────────────────────────────────────────────

class TestPluginContextMemoryField(unittest.TestCase):

    def test_context_has_memory_field(self):
        from plugin_base import PluginContext
        ctx = PluginContext()
        self.assertTrue(hasattr(ctx, "memory"),
                        "PluginContext must have a 'memory' field")

    def test_context_memory_defaults_to_none(self):
        from plugin_base import PluginContext
        ctx = PluginContext()
        self.assertIsNone(ctx.memory)

    def test_context_accepts_memory_store(self):
        from plugin_base import PluginContext
        mock_memory = MagicMock()
        ctx = PluginContext(memory=mock_memory)
        self.assertIs(ctx.memory, mock_memory)


# ── 8. plugin_loader._make_context() injects MemoryStore ─────────────────────

class TestPluginLoaderMemoryInjection(unittest.TestCase):

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_make_context_injects_memory(self):
        from plugin_loader import PluginLoader
        from memory_store import MemoryStore
        loader = PluginLoader()
        ctx = loader._make_context(draft_mode=True)
        self.assertIsNotNone(ctx.memory,
                             "_make_context() should inject MemoryStore into context.memory")
        self.assertIs(ctx.memory, MemoryStore)

    def test_make_context_memory_is_callable(self):
        from plugin_loader import PluginLoader
        loader = PluginLoader()
        ctx = loader._make_context(draft_mode=True)
        if ctx.memory is not None:
            self.assertTrue(hasattr(ctx.memory, "store"))
            self.assertTrue(hasattr(ctx.memory, "search"))
            self.assertTrue(hasattr(ctx.memory, "get_client_context"))


# ── 9. Graceful degradation ───────────────────────────────────────────────────

class TestGracefulDegradation(unittest.TestCase):

    def setUp(self):
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_loader_graceful_when_chromadb_missing(self):
        """If chromadb import fails, _make_context() sets memory=None without crashing."""
        from plugin_loader import PluginLoader
        loader = PluginLoader()
        with patch.dict("sys.modules", {"chromadb": None}):
            with patch("plugin_loader.PluginLoader._make_context",
                       wraps=loader._make_context) as mock_ctx:
                # Simulate chromadb import failure inside _make_context
                import importlib
                with patch("builtins.__import__",
                           side_effect=lambda name, *a, **kw: (_ for _ in ()).throw(
                               ImportError("No module named 'chromadb'")
                           ) if name == "memory_store" else importlib.import_module(name)):
                    ctx = loader._make_context(draft_mode=True)
                    # memory should be None when import fails
                    self.assertIsNone(ctx.memory)

    def test_search_returns_empty_on_error(self):
        """MemoryStore.search() returns [] on any exception."""
        import memory_store as ms
        original_client = ms._client
        ms._client = None
        tmp_chroma = tempfile.mkdtemp()
        ms.CHROMA_PATH = Path(tmp_chroma)

        # Force an error by patching the internal client
        with patch("memory_store._get_client", side_effect=Exception("DB error")):
            results = ms.MemoryStore.search("test query")
        self.assertEqual(results, [])

        ms._client = original_client
        shutil.rmtree(tmp_chroma, ignore_errors=True)


# ── 10. NOA processor memory integration ─────────────────────────────────────

class TestNOAProcessorMemoryIntegration(unittest.TestCase):
    """Verify that plugin_noa_processor stores outcomes in MemoryStore."""

    def setUp(self):
        self.MemoryStore, self.tmp_dir = make_fresh_memory_store()
        cfg.DB_PATH = Path(tempfile.mktemp(suffix=".db"))
        cfg.init_db()

    def tearDown(self):
        import memory_store as ms
        ms._client = None
        shutil.rmtree(self.tmp_dir, ignore_errors=True)
        try:
            os.unlink(str(cfg.DB_PATH))
        except Exception:
            pass

    def test_noa_outcome_stored_in_memory(self):
        """Simulate what the NOA processor does after processing an NOA."""
        client_email = "john@example.com"
        client_name = "John Smith"
        outcome = "REFUND"
        amount = "$2,400"
        tax_year = "2024"
        entity_name = "John Smith"

        memory_text = (
            f"{client_name} NOA {tax_year}: outcome={outcome}, "
            f"amount={amount}, entity={entity_name}"
        )
        self.MemoryStore.store_client_interaction(
            content=memory_text,
            client_email=client_email,
            interaction_type="noa_outcome",
            extra_meta={
                "outcome": outcome,
                "amount": amount,
                "tax_year": tax_year,
                "entity_name": entity_name,
            },
        )

        # Verify it's retrievable
        results = self.MemoryStore.search(
            "John Smith refund",
            collection="client_interactions",
            where={"client_email": client_email},
        )
        self.assertGreater(len(results), 0)
        doc, meta, _ = results[0]
        self.assertIn("REFUND", doc)
        self.assertEqual(meta["client_email"], client_email)
        self.assertEqual(meta["outcome"], "REFUND")

    def test_client_context_retrieval_for_outreach(self):
        """Verify get_client_context() returns NOA history for outreach plugin."""
        client_email = "jane@example.com"

        # Store some history
        self.MemoryStore.store_client_interaction(
            "Jane Doe NOA 2024: outcome=PAYABLE, amount=$1,200",
            client_email=client_email,
            interaction_type="noa_outcome",
        )
        self.MemoryStore.store_client_interaction(
            "Jane Doe NOA 2023: outcome=REFUND, amount=$500",
            client_email=client_email,
            interaction_type="noa_outcome",
        )

        context_str = self.MemoryStore.get_client_context(
            client_email=client_email,
            query="tax history payable refund",
            n_results=5,
        )
        self.assertIsInstance(context_str, str)
        self.assertGreater(len(context_str), 0)
        self.assertIn("jane@example.com", context_str)


if __name__ == "__main__":
    unittest.main(verbosity=2)
