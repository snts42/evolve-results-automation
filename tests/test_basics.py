"""Basic offline unit tests for the pure helper functions.

These tests cover the functions whose correctness is critical for data
integrity (hashing, dedup, date parsing, PDF filename generation) and
security (credential encryption round-trip). No Selenium, no GUI, no
network, no live E-volve access - everything runs instantly and
deterministically on any machine.

Run with::

    python -m pytest tests/ -v
"""
from datetime import datetime
from unittest.mock import patch

import pytest

import evolve_results_automation.main as main_mod
from evolve_results_automation.main import compute_pdf_cutoff_date
from evolve_results_automation.parsing_utils import (
    unique_row_hash, report_filename
)
from evolve_results_automation.excel_utils import format_ddmmyyyy
from evolve_results_automation.secure_credentials import SecureCredentialManager


# ---------------------------------------------------------------------------
# Fixed datetime helpers (no third-party dependency)
# ---------------------------------------------------------------------------

def _fixed_datetime(fixed):
    """Build a datetime subclass whose .now() returns the given value.

    Preserves the regular datetime constructor so ``datetime(y, m, d)`` keeps
    working inside the function under test.
    """
    class _Fixed(datetime):
        @classmethod
        def now(cls, tz=None):
            return fixed
    return _Fixed


# ---------------------------------------------------------------------------
# unique_row_hash
# ---------------------------------------------------------------------------

_ROW = {
    "Enrolment no.": "12345678",
    "First name": "Emma",
    "Last name": "Smith",
    "Completed": "15/06/2026",
    "Test Name": "Functional Skills English Level 2",
    "Result": "Pass",
}


def test_unique_row_hash_deterministic():
    """Same row must always produce the same hash."""
    assert unique_row_hash(_ROW) == unique_row_hash(dict(_ROW))


def test_unique_row_hash_different_enrolment():
    """Two candidates with the same name but different enrolments must not collide."""
    other = dict(_ROW, **{"Enrolment no.": "99999999"})
    assert unique_row_hash(_ROW) != unique_row_hash(other)


def test_unique_row_hash_missing_fields():
    """Missing fields must not crash; hash stays a pipe-delimited string."""
    partial = {"First name": "Emma"}
    h = unique_row_hash(partial)
    assert isinstance(h, str)
    assert "|" in h


def test_unique_row_hash_case_insensitive():
    """Result 'PASS' and 'Pass' must produce the same hash so existing rows match."""
    lower = dict(_ROW, Result="Pass")
    upper = dict(_ROW, Result="PASS")
    assert unique_row_hash(lower) == unique_row_hash(upper)


# ---------------------------------------------------------------------------
# compute_pdf_cutoff_date (uses months_back + 1 rule)
# ---------------------------------------------------------------------------

def test_compute_pdf_cutoff_date_basic():
    """months_back=1 in June 2026 -> cutoff is 1 April 2026.

    The function uses ``months_back + 1`` as the effective lookback because
    the E-volve calendar opens one month behind current. A months_back=1
    setting therefore covers the two most recent months.
    """
    with patch.object(main_mod, "datetime", _fixed_datetime(datetime(2026, 6, 15))):
        cutoff = compute_pdf_cutoff_date(1)
    assert cutoff == datetime(2026, 4, 1)


def test_compute_pdf_cutoff_date_year_boundary():
    """months_back=1 in January 2027 must wrap back into the previous year.

    Effective lookback = 2 months, so the cutoff is 1 November 2026. This
    protects against off-by-one bugs in the month arithmetic.
    """
    with patch.object(main_mod, "datetime", _fixed_datetime(datetime(2027, 1, 15))):
        cutoff = compute_pdf_cutoff_date(1)
    assert cutoff == datetime(2026, 11, 1)


def test_compute_pdf_cutoff_date_large():
    """months_back=24 in June 2026 -> cutoff is 1 May 2024 (25 months back)."""
    with patch.object(main_mod, "datetime", _fixed_datetime(datetime(2026, 6, 15))):
        cutoff = compute_pdf_cutoff_date(24)
    assert cutoff == datetime(2024, 5, 1)


# ---------------------------------------------------------------------------
# report_filename (NEW-1 regression test is first)
# ---------------------------------------------------------------------------

def test_report_filename_includes_enrolment():
    """Two candidates who share first+last+test+result+date must get different
    filenames because Enrolment no. is now part of the filename. This is the
    regression test for the PDF filename collision bug fixed in v1.3.3."""
    base = {
        "First name": "Emma",
        "Last name": "Smith",
        "Test Name": "Functional Skills English Level 2",
        "Result": "Pass",
        "Completed": "15/06/2026",
    }
    a = dict(base, **{"Enrolment no.": "11111111"})
    b = dict(base, **{"Enrolment no.": "22222222"})
    fa = report_filename(a)
    fb = report_filename(b)
    assert fa != fb
    assert "11111111" in fa
    assert "22222222" in fb


def test_report_filename_sanitizes_forbidden_chars():
    """Windows-forbidden characters must be stripped from the filename."""
    row = {
        "First name": "Emma/Odd",
        "Last name": "Smith:Test",
        "Enrolment no.": "12345678",
        "Test Name": 'Bad<>|"?*Name',
        "Result": "Pass",
    }
    fname = report_filename(row)
    for forbidden in r'\/:*?"<>|':
        assert forbidden not in fname


# ---------------------------------------------------------------------------
# format_ddmmyyyy
# ---------------------------------------------------------------------------

def test_format_ddmmyyyy_iso_input():
    """E-volve-style ISO timestamp must be converted to UK dd/mm/yyyy."""
    assert format_ddmmyyyy("2026-04-15 10:30:00") == "15/04/2026"


def test_format_ddmmyyyy_already_formatted():
    """An already-formatted dd/mm/yyyy string must pass through unchanged."""
    assert format_ddmmyyyy("15/04/2026") == "15/04/2026"


def test_format_ddmmyyyy_empty_and_none():
    """Empty, whitespace, and None inputs must return an empty string."""
    assert format_ddmmyyyy("") == ""
    assert format_ddmmyyyy(None) == ""
    assert format_ddmmyyyy("   ") == ""


# ---------------------------------------------------------------------------
# Credential encryption round-trip
# ---------------------------------------------------------------------------

def test_credential_roundtrip(tmp_path):
    """Add two credentials, decrypt, remove one, decrypt again. Wrong password
    must raise. Covers AES-256-CBC + HMAC-SHA256 + PBKDF2 end-to-end."""
    enc_file = tmp_path / "credentials.enc"
    mgr = SecureCredentialManager(str(enc_file))
    master = "correct horse battery staple"

    assert mgr.add_credential("alice@example.com", "alice-pass", master) is True
    assert mgr.add_credential("bob@example.com", "bob-pass", master) is True

    creds = mgr.decrypt_credentials(master)
    assert len(creds) == 2
    usernames = {c["username"] for c in creds}
    assert usernames == {"alice@example.com", "bob@example.com"}

    # Wrong password must not decrypt
    with pytest.raises(ValueError):
        mgr.decrypt_credentials("wrong-password")

    # Remove one and verify
    assert mgr.remove_credential("alice@example.com", master) is True
    remaining = mgr.decrypt_credentials(master)
    assert len(remaining) == 1
    assert remaining[0]["username"] == "bob@example.com"
