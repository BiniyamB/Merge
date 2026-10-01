"""Tests for the Digital Transaction Value Snapshot report helpers."""

from __future__ import annotations

from pathlib import Path
import re
import sys

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import snapshot_module as sm  # noqa: E402


# ── key messages become one bullet per point ────────────────────────────────
def _defaults_by_name():
    return {s["name"]: s for s in sm.SERVICE_DEFAULTS}


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


# ── ATM and POS success rate share one justification, drawn as one part ─────
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
    """Only the justification is merged - the two rates remain distinct rows."""
    names = [s["name"] for s in sm.SERVICE_DEFAULTS]
    assert "ATM SUCCESS RATE" in names
    assert "POS SUCCESS RATE" in names
    rows = _defaults_by_name()
    assert rows["ATM SUCCESS RATE"]["transactionVolume"] != \
        rows["POS SUCCESS RATE"]["transactionVolume"]
    assert rows["ATM SUCCESS RATE"]["target"] != rows["POS SUCCESS RATE"]["target"]


def test_pos_row_is_folded_into_the_atm_row():
    groups = sm.merge_success_rate_groups(sm.SERVICE_DEFAULTS)
    owner, rest = groups[0]
    assert owner == 0, "the ATM row owns the shared justification"
    assert rest == [1], "the POS row continues it"
    # every other row keeps its own message
    assert all(not rest for _o, rest in groups[1:])


def test_a_row_with_its_own_message_starts_its_own_group():
    services = [
        {"name": "ATM SUCCESS RATE", "type": "success-rate", "keyMessage": "Shared"},
        {"name": "POS SUCCESS RATE", "type": "success-rate", "keyMessage": ""},
        {"name": "P2P SUCCESS RATE", "type": "success-rate", "keyMessage": "Its own"},
        {"name": "QR", "type": "financial", "keyMessage": ""},
    ]
    assert sm.merge_success_rate_groups(services) == [(0, [1]), (2, []), (3, [])]


def test_only_the_atm_pos_pair_shares_a_message():
    """The pair is matched by name, so two unrelated rate rows never merge."""
    services = [
        {"name": "ATM SUCCESS RATE", "type": "success-rate", "keyMessage": "Shared"},
        {"name": "P2P SUCCESS RATE", "type": "success-rate", "keyMessage": ""},
        {"name": "CASH SUCCESS RATE", "type": "success-rate", "keyMessage": ""},
    ]
    assert sm.merge_success_rate_groups(services) == [(0, []), (1, []), (2, [])]


def test_p2p_stays_separate_even_with_a_blank_message():
    """P2P sits under the pair; clearing its message must not widen the merge."""
    services = [dict(s) for s in sm.SERVICE_DEFAULTS]
    services[2]["keyMessage"] = ""
    groups = sm.merge_success_rate_groups(services)
    assert groups[0] == (0, [1]), "ATM and POS still merge"
    assert groups[1] == (2, []), "P2P keeps its own justification"


def test_a_rate_row_without_a_message_is_not_folded_into_a_non_rate_row():
    services = [
        {"name": "CASH WITHDRAWAL", "type": "financial", "keyMessage": "Cash"},
        {"name": "ATM SUCCESS RATE", "type": "success-rate", "keyMessage": ""},
    ]
    # nothing above it is a success rate, so it stands on its own
    assert sm.merge_success_rate_groups(services) == [(0, []), (1, [])]


def test_the_pair_merges_in_either_order():
    services = [
        {"name": "POS SUCCESS RATE", "type": "success-rate", "keyMessage": ""},
        {"name": "ATM SUCCESS RATE", "type": "success-rate", "keyMessage": "Shared"},
    ]
    # the row above carries the message, so nothing is folded away here
    assert sm.merge_success_rate_groups(services) == [(0, []), (1, [])]


def test_report_draws_the_merged_pair_without_a_separating_line():
    calc = sm.calc_all(sm.SERVICE_DEFAULTS)
    html = sm.build_report_html(dict(sm.REPORT_DEFAULTS), calc)
    table = re.search(r"<table class=\"report-table\">.*?</table>", html, re.S).group(0)

    atm = re.search(r'<tr class="msg-merge-top">(.*?)</tr>', table, re.S).group(1)
    pos = re.search(r'<tr class="msg-merge-bottom">(.*?)</tr>', table, re.S).group(1)

    # one justification cell, written once, spanning both rows
    assert 'rowspan="2"' in atm
    assert "ATM and POS acceptance stay above plan" in atm
    assert "msg-cell" not in pos, "the shared cell already covers the POS row"
    # and neither row is left with a rule between them
    assert "border-bottom: none" in sm.REPORT_CSS
    assert "msg-merge-bottom" in sm.REPORT_CSS


def test_the_merged_pair_keeps_the_table_grid_consistent():
    calc = sm.calc_all(sm.SERVICE_DEFAULTS)
    html = sm.build_report_html(dict(sm.REPORT_DEFAULTS), calc)
    table = re.search(r"<table class=\"report-table\">.*?</table>", html, re.S).group(0)
    columns = len(re.findall(r"<th[ >]", re.search(r"<thead>.*?</thead>", table, re.S).group(0)))

    carry = 0
    for _attrs, body in re.findall(r"<tr([^>]*)>(.*?)</tr>", table, re.S):
        if "<th" in body:
            continue
        cells = re.findall(r"<td([^>]*)>", body)
        assert len(cells) + carry == columns, "row does not fill the grid"
        carry = sum(int(m.group(1)) - 1 for m in
                    (re.search(r'rowspan="(\d+)"', td) for td in cells) if m)


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