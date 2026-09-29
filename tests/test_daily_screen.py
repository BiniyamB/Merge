"""Wiring tests for the "Daily Compiled" screen in the Streamlit app.

These need streamlit, which only the venv that runs the app has, so they skip
in the lighter test environment.
"""

from __future__ import annotations

import calendar
from datetime import date
from pathlib import Path

import pytest

pytest.importorskip("streamlit")
from streamlit.testing.v1 import AppTest  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "28.09.2026"
BOOK = SAMPLE / "September_Successful_Financial_and_Decline_Transaction_Report_28.xlsx"

pytestmark = pytest.mark.skipif(
    not BOOK.exists(), reason="reference workbook is not available")


@pytest.fixture
def screen():
    app = AppTest.from_file(str(ROOT / "streamlit_app.py"), default_timeout=120)
    app.run()
    app.session_state["mode_key"] = "daily"
    app.session_state["daily_workbook"] = BOOK.read_bytes()
    app.run()
    return app


def test_daily_screen_renders_without_errors(screen):
    assert [e.value for e in screen.error] == []


def test_workbook_and_the_eight_sources_are_offered(screen):
    # one uploader for the workbook plus one per source report
    assert len(screen.file_uploader) == 9


def test_each_upload_is_labelled_with_the_report_it_needs(screen):
    labels = [f.label for f in screen.file_uploader[1:]]
    assert labels == [
        "Issuer card - successful transactions",
        "Acquirer card - successful transactions",
        "Issuer card - declined transactions",
        "Acquirer card - declined transactions",
        "IPS - successful interbank transfers",
        "IPS - declined interbank transfers",
        "QR - successful interbank payments",
        "QR - declined interbank payments",
    ]
    # the file names are never asked for
    assert not any(".xlsx" in label for label in labels)


def test_month_list_covers_the_workbook_this_month_and_the_next(screen):
    month = screen.selectbox[0]
    # the months the workbook already holds, with how many days are in them
    assert month.options[0] == "September 2026 (28 day(s) already)"
    assert month.value == (2026, 9)

    today = date.today()
    following = ((today.year + 1, 1) if today.month == 12
                 else (today.year, today.month + 1))
    name = calendar.month_name[following[1]]
    # a month with no sheets yet can still be started, so the next one is offered
    assert f"{name} {following[0]} (new month)" in month.options


def test_day_list_follows_the_month(screen):
    assert len(screen.selectbox[1].options) == 30, "September has 30 days"


def test_first_missing_day_is_offered_by_default(screen):
    # the workbook stops at 28.09, so 29.09 is the next new day
    assert screen.selectbox[1].value == 29
    assert "does not exist yet" in screen.caption[0].value
    assert "28.09.2026" in screen.caption[0].value


def test_a_new_month_offers_every_day_and_an_empty_footer(screen):
    label = [o for o in screen.selectbox[0].options if "new month" in o][0]
    name, year = label.split()[0], int(label.split()[1])
    month = {calendar.month_name[i]: i for i in range(1, 13)}[name]
    screen.selectbox[0].select((year, month))
    screen.run()
    assert [e.value for e in screen.error] == []
    # every valid day of that month is on offer
    assert len(screen.selectbox[1].options) == calendar.monthrange(year, month)[1]
    plans = {n.label: n.value for n in screen.number_input}
    # nothing carried over from the previous month: blank plans, zero RTP
    assert set(plans.values()) == {0.0}
    assert any("no sheet in this workbook yet" in c.value for c in screen.caption)


def test_choosing_an_existing_day_names_the_previous_day_as_donor(screen):
    screen.selectbox[1].select(28)
    screen.run()
    caption = screen.caption[0].value
    assert "will replace the existing sheet" in caption
    # a replaced day must not donate its own layout
    assert "27.09.2026" in caption


def test_monthly_footer_is_read_from_the_sheet_and_prefilled(screen):
    plans = {n.label: n.value for n in screen.number_input}
    assert plans["Cash Withdrawal plan"] == pytest.approx(347974.3)
    assert plans["POS Purchase plan"] == pytest.approx(11727.78, rel=1e-3)
    assert plans["IPS RTP transactions"] == 7.0
    assert plans["IPS RTP value"] == 502.0
    assert [t.label for t in screen.text_area] == [
        "Card services note", "P2P, IPS and ETH QR note"]


def test_build_is_blocked_until_every_source_is_uploaded(screen):
    build = [b for b in screen.button if b.label == "Build report sheet"][0]
    assert build.disabled
    assert any("Still missing" in w.value for w in screen.warning)
