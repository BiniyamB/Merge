"""Tests for the Digital Transaction Value Snapshot report helpers."""

from __future__ import annotations

from pathlib import Path
import sys

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import snapshot_module as sm  # noqa: E402


# ── key messages become one bullet per point ────────────────────────────────
@pytest.mark.parametrize("text, expected", [
    ("One message", ["One message"]),
    ("First; second", ["First", "second"]),
    ("First\nSecond\nThird", ["First", "Second", "Third"]),
    ("First\r\nSecond", ["First", "Second"]),
    # a bullet the user typed is dropped, the renderer adds its own
    ("\u2022 First", ["First"]),
    ("- First", ["First"]),
    ("1. First; 2) Second", ["First", "Second"]),
    # blank points never become empty bullets
    ("First;;Second;", ["First", "Second"]),
    ("\n\n", []),
    ("", []),
    (None, []),
])
def test_key_message_points(text, expected):
    assert sm.key_message_points(text) == expected


def test_key_message_renders_one_li_per_point():
    html = sm.key_message_html("Alpha; Beta")
    assert html.count("<li>") == 2
    assert html.startswith('<ul class="msg-list">')
    assert "Alpha" in html and "Beta" in html


def test_single_message_is_still_a_bullet():
    assert sm.key_message_html("Only one").count("<li>") == 1


def test_blank_message_renders_nothing():
    assert sm.key_message_html("") == ""
    assert sm.key_message_html(None) == ""


def test_key_message_is_escaped():
    assert "&lt;script&gt;" in sm.key_message_html("<script>alert(1)</script>")


def test_points_do_not_run_together_on_one_line():
    """The bug: newlines were collapsed, so every point landed on one line."""
    html = sm.key_message_html("Alpha\nBeta\nGamma")
    assert html.count("<li>") == 3
    # each point lives in its own element rather than one run of text
    assert "<li>Alpha</li><li>Beta</li><li>Gamma</li>" in html


# ── the ATM / POS success rate message ──────────────────────────────────────
def _defaults_by_name():
    return {s["name"]: s for s in sm.SERVICE_DEFAULTS}


def test_atm_and_pos_share_one_merged_key_message():
    rows = _defaults_by_name()
    atm = rows["ATM SUCCESS RATE"]["keyMessage"]
    pos = rows["POS SUCCESS RATE"]["keyMessage"]
    assert atm.strip(), "the merged message must actually say something"
    assert pos == "", "the message is written once, so POS must not repeat it"
    # the merged message speaks for both services
    assert "ATM" in atm.upper()
    assert "POS" in atm.upper()


def test_atm_and_pos_rows_stay_separate():
    """Only the key message is merged - the two rates remain distinct rows."""
    names = [s["name"] for s in sm.SERVICE_DEFAULTS]
    assert "ATM SUCCESS RATE" in names
    assert "POS SUCCESS RATE" in names
    rows = _defaults_by_name()
    assert rows["ATM SUCCESS RATE"]["transactionVolume"] != \
        rows["POS SUCCESS RATE"]["transactionVolume"]
    assert rows["ATM SUCCESS RATE"]["target"] != rows["POS SUCCESS RATE"]["target"]


def test_success_rates_still_excluded_from_the_total():
    calc = sm.calc_all(sm.SERVICE_DEFAULTS)
    counted = [x["name"] for x in calc["services"]
               if not x["isSuccessRate"] and "RTP" not in x["name"]
               and "NPG" not in x["name"]]
    total = calc["total"]
    assert total["performance"] == pytest.approx(
        sum(x["transactionVolume"] for x in calc["services"] if x["name"] in counted))


# ── the rendered report ─────────────────────────────────────────────────────
def test_report_renders_key_messages_as_bullets():
    services = [dict(s) for s in sm.SERVICE_DEFAULTS]
    services[3]["keyMessage"] = "Point one; Point two"
    calc = sm.calc_all(services)
    html = sm.build_report_html(dict(sm.REPORT_DEFAULTS), calc)
    assert "<li>Point one</li><li>Point two</li>" in html


def test_report_escapes_a_hostile_key_message():
    services = [dict(s) for s in sm.SERVICE_DEFAULTS]
    services[3]["keyMessage"] = "<img src=x onerror=alert(1)>"
    calc = sm.calc_all(services)
    html = sm.build_report_html(dict(sm.REPORT_DEFAULTS), calc)
    assert "<img src=x onerror=" not in html
    assert "&lt;img src=x" in html