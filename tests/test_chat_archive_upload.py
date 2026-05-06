"""
Tests for the chat archive SharePoint upload plumbing (Fix B2).

Pins the contract from fix_b2_sharepoint_upload.md:
  - SPECIALIST_FOLDER_MAP uses the corrected v2.3 agent_id keys (no
    `trust_tax`, `individual_deductions`, `net_wealth_smsf_audit`; the
    folder NAMES stay as-is because they are real SharePoint folders).
  - plugin_builder is intentionally absent.
  - _build_chat_archive_paths returns (folder_path, file_path) under
    CHAT_ARCHIVE_ROOT for a known agent_id.
  - Unknown agent_id -> ValueError("No SharePoint folder mapped...").
  - SharePointFolderMissing exists, subclasses Exception, no extra behaviour.

The actual Graph upload is NOT exercised here (per the fix file: mocking
_make_request would test the mock, not the code). End-to-end smoke test
happens manually after B5.
"""
import sys
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

from graph_client import (  # noqa: E402
    CHAT_ARCHIVE_ROOT,
    SPECIALIST_FOLDER_MAP,
    SharePointFolderMissing,
    _build_chat_archive_paths,
)


class SpecialistFolderMapTests(unittest.TestCase):
    def test_chat_archive_root_is_mcs_coworkers_path(self):
        self.assertEqual(CHAT_ARCHIVE_ROOT, "Server/Clients/MCS Coworkers")

    def test_corrected_agent_ids_are_present(self):
        # Phase B kickoff renamed three keys. Pin the corrected ones so a
        # future drift back to the task-spec drafts is caught here.
        for key in ("trusts", "individual_tax", "net_wealth"):
            self.assertIn(key, SPECIALIST_FOLDER_MAP)

    def test_old_drafted_agent_ids_are_absent(self):
        for stale in ("trust_tax", "individual_deductions", "net_wealth_smsf_audit"):
            self.assertNotIn(stale, SPECIALIST_FOLDER_MAP)

    def test_plugin_builder_is_not_in_map(self):
        # Excluded by design — /finish blocks it at the endpoint layer.
        self.assertNotIn("plugin_builder", SPECIALIST_FOLDER_MAP)

    def test_full_expected_mapping(self):
        # Folder NAMES (right-hand side) are real SharePoint folder names
        # and must stay verbatim. Any change here means a SharePoint-side
        # rename also happened; investigate before adjusting the test.
        self.assertEqual(SPECIALIST_FOLDER_MAP, {
            "general": "General",
            "gst": "GST Specialist",
            "smsf": "SMSF Specialist",
            "div7a": "Division 7A Specialist",
            "trusts": "Trust Tax Specialist",
            "tax_structure": "Tax Structure Specialist",
            "individual_tax": "Individual Tax Deductions",
            "payroll": "Payroll Specialist",
            "net_wealth": "Net Wealth SMSF Audit",
            "ato_portal": "General",
            "ato_correspondence": "General",
        })


class BuildChatArchivePathsTests(unittest.TestCase):
    def test_known_agent_returns_folder_and_file_paths(self):
        folder, file_path = _build_chat_archive_paths(
            "gst", "2026-05-06 GST - Going Concern Sale.docx",
        )
        self.assertEqual(folder, "Server/Clients/MCS Coworkers/GST Specialist")
        self.assertEqual(
            file_path,
            "Server/Clients/MCS Coworkers/GST Specialist/"
            "2026-05-06 GST - Going Concern Sale.docx",
        )

    def test_renamed_agent_uses_corrected_key(self):
        # `trusts` (not `trust_tax`) is the live key — confirm the path
        # build reaches the "Trust Tax Specialist" folder.
        folder, file_path = _build_chat_archive_paths(
            "trusts", "2026-05-06 Trust Tax - Sub-Trust UPE.docx",
        )
        self.assertEqual(
            folder, "Server/Clients/MCS Coworkers/Trust Tax Specialist"
        )
        self.assertTrue(file_path.startswith(folder + "/"))

    def test_ato_portal_routes_to_general(self):
        # ATO Portal Assistant has no dedicated folder yet; spec routes
        # it to General. Pin that wiring.
        folder, _file_path = _build_chat_archive_paths(
            "ato_portal", "anything.docx",
        )
        self.assertEqual(folder, "Server/Clients/MCS Coworkers/General")

    def test_unknown_agent_id_raises_valueerror_with_repr(self):
        with self.assertRaises(ValueError) as ctx:
            _build_chat_archive_paths("not_a_real_agent", "file.docx")
        self.assertIn("not_a_real_agent", str(ctx.exception))
        self.assertIn("No SharePoint folder mapped", str(ctx.exception))

    def test_plugin_builder_raises_valueerror(self):
        # plugin_builder hits the same ValueError — endpoint layer is the
        # primary gate, but this is the defence-in-depth check.
        with self.assertRaises(ValueError):
            _build_chat_archive_paths("plugin_builder", "file.docx")


class SharePointFolderMissingTests(unittest.TestCase):
    def test_subclasses_exception(self):
        self.assertTrue(issubclass(SharePointFolderMissing, Exception))

    def test_message_is_plain_str(self):
        # No special behaviour beyond Exception — the spec is explicit.
        e = SharePointFolderMissing("hello")
        self.assertEqual(str(e), "hello")


if __name__ == "__main__":
    unittest.main()
