"""Tests for _ensure_client_folder_exists and the related governance
exceptions in graph_client.

The helper is read-only by design. These tests pin its three branches —
single match (case/whitespace-insensitive), missing, ambiguous — plus the
pagination edge case where the listing is truncated.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

from graph_client import (  # noqa: E402
    GraphClient,
    SharePointFolderAmbiguous,
    SharePointFolderMissing,
)


def _folder(name: str) -> dict:
    """Shape of a Graph driveItem entry that represents a folder."""
    return {"id": f"id-{name}", "name": name, "folder": {"childCount": 0}}


def _file(name: str) -> dict:
    """Shape of a Graph driveItem that's a file (no folder facet)."""
    return {"id": f"file-{name}", "name": name}


class TestEnsureClientFolderExists(unittest.TestCase):
    """Pin the contract: exactly-one match returns verbatim, zero raises
    Missing, two or more raises Ambiguous, pagination edge raises Missing
    with a clear hint."""

    def setUp(self):
        # GraphClient.__init__ touches MSAL; we never call any unmocked path.
        self.graph = GraphClient.__new__(GraphClient)

    def _patch_request(self, response: dict | None):
        return patch.object(self.graph, "_make_request", return_value=response)

    # 1. Single exact match → verbatim name.
    def test_single_exact_match_returns_verbatim(self):
        response = {"value": [_folder("ABC Pty Ltd")]}
        with self._patch_request(response):
            actual = self.graph._ensure_client_folder_exists(
                site_id="SITE", drive_id="DRIVE",
                parent_path="Server/Clients", folder_name="ABC Pty Ltd",
            )
        self.assertEqual(actual, "ABC Pty Ltd")

    # 2. Single case-insensitive match → returns verbatim (existing wins).
    def test_single_case_insensitive_match_returns_verbatim(self):
        response = {"value": [_folder("abc pty ltd")]}
        with self._patch_request(response):
            actual = self.graph._ensure_client_folder_exists(
                site_id="SITE", drive_id="DRIVE",
                parent_path="Server/Clients", folder_name="ABC Pty Ltd",
            )
        self.assertEqual(actual, "abc pty ltd")

    # 3. Single whitespace-trimmed match → returns verbatim including the
    #    trailing space. The verbatim name is what goes into upload paths.
    def test_single_whitespace_match_returns_verbatim_with_trailing_space(self):
        response = {"value": [_folder("ABC Pty Ltd ")]}
        with self._patch_request(response):
            actual = self.graph._ensure_client_folder_exists(
                site_id="SITE", drive_id="DRIVE",
                parent_path="Server/Clients", folder_name="ABC Pty Ltd",
            )
        self.assertEqual(actual, "ABC Pty Ltd ")

    # 4. Zero matches → SharePointFolderMissing, message names the client.
    def test_zero_matches_raises_missing(self):
        response = {"value": [
            _folder("Other Pty Ltd"),
            _folder("Different Holdings"),
        ]}
        with self._patch_request(response):
            with self.assertRaises(SharePointFolderMissing) as cm:
                self.graph._ensure_client_folder_exists(
                    site_id="SITE", drive_id="DRIVE",
                    parent_path="Server/Clients", folder_name="ABC Pty Ltd",
                )
        self.assertIn("ABC Pty Ltd", str(cm.exception))

    # 5. Two matches → SharePointFolderAmbiguous, message lists candidates.
    def test_two_matches_raises_ambiguous_with_both_names(self):
        response = {"value": [
            _folder("Beta Holdings"),
            _folder("Beta  Holdings"),  # double internal space
        ]}
        with self._patch_request(response):
            with self.assertRaises(SharePointFolderAmbiguous) as cm:
                self.graph._ensure_client_folder_exists(
                    site_id="SITE", drive_id="DRIVE",
                    parent_path="Server/Clients", folder_name="Beta Holdings",
                )
        msg = str(cm.exception)
        self.assertIn("Beta Holdings", msg)
        self.assertIn("Beta  Holdings", msg)  # verbatim, double-space preserved

    # 6. Empty parent → Missing.
    def test_empty_parent_raises_missing(self):
        response = {"value": []}
        with self._patch_request(response):
            with self.assertRaises(SharePointFolderMissing):
                self.graph._ensure_client_folder_exists(
                    site_id="SITE", drive_id="DRIVE",
                    parent_path="Server/Clients", folder_name="ABC Pty Ltd",
                )

    # 7. Pagination edge: nextLink + no match in first page → Missing with
    #    a pagination hint, not a silent miss.
    def test_pagination_with_no_match_raises_missing_with_hint(self):
        response = {
            "value": [_folder("Other Folder")],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/...next...",
        }
        with self._patch_request(response):
            with self.assertRaises(SharePointFolderMissing) as cm:
                self.graph._ensure_client_folder_exists(
                    site_id="SITE", drive_id="DRIVE",
                    parent_path="Server/Clients", folder_name="ABC Pty Ltd",
                )
        msg = str(cm.exception)
        self.assertIn("paginated", msg.lower())
        self.assertIn("ABC Pty Ltd", msg)

    # 8. Sanity — a file with the same name shouldn't be considered a match.
    #    Folders are identified by the presence of the "folder" facet.
    def test_file_with_same_name_is_not_a_match(self):
        response = {"value": [_file("ABC Pty Ltd")]}
        with self._patch_request(response):
            with self.assertRaises(SharePointFolderMissing):
                self.graph._ensure_client_folder_exists(
                    site_id="SITE", drive_id="DRIVE",
                    parent_path="Server/Clients", folder_name="ABC Pty Ltd",
                )


if __name__ == "__main__":
    unittest.main()
