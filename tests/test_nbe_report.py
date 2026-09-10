"""Unit tests for nbe_report module."""

import io
import openpyxl
import pandas as pd
import pytest

from nbe_report import (
    STANDARD_NBE_BANKS,
    build_nbe_report_excel,
    build_sett_sum_report_excel,
    generate_nbe_report,
    generate_sett_sum_report,
    normalize_nbe_bank,
)


def test_normalize_nbe_bank():
    """Verify alias mapping for issuer & acquirer bank names."""
    assert normalize_nbe_bank("Abay") == "Abay Bank"
    assert normalize_nbe_bank("Abay Bank") == "Abay Bank"
    assert normalize_nbe_bank("Abyssinia") == "BOA"
    assert normalize_nbe_bank("Abyssinia Bank") == "BOA"
    assert normalize_nbe_bank("BOA") == "BOA"
    assert normalize_nbe_bank("Commercial Bank") == "CBE"
    assert normalize_nbe_bank("Commercial Bank of Ethiopia") == "CBE"
    assert normalize_nbe_bank("CBE") == "CBE"
    assert normalize_nbe_bank("Hibret Bank") == "United Bank"
    assert normalize_nbe_bank("UB") == "United Bank"
    assert normalize_nbe_bank("Global") == "Global Bank"
    assert normalize_nbe_bank("Debub Bank") == "Global Bank"
    assert normalize_nbe_bank("Berhan Bank") == "Birhan Bank"
    assert normalize_nbe_bank("Tsedey Bank") == "Tseday Bank"
    assert normalize_nbe_bank(None) == ""


def test_generate_nbe_report_pos():
    """Verify POS NBE Report filtering and aggregation."""
    records = [
        # Match (POS purchase, RESP -1)
        {"TRANS_TYPE": "POS purchase", "RESP": "-1", "ISSUER": "Commercial Bank", "ACQUIRER": "Abyssinia Bank", "AMOUNT": 1000.0},
        {"TRANS_TYPE": "purchase", "RESP": -1.0, "ISSUER": "Awash Bank", "ACQUIRER": "CBE", "AMOUNT": 500.0},
        # Non-matching response code
        {"TRANS_TYPE": "POS purchase", "RESP": "901", "ISSUER": "Awash Bank", "ACQUIRER": "CBE", "AMOUNT": 200.0},
        # Non-matching trans type
        {"TRANS_TYPE": "POS balance inquiry", "RESP": "-1", "ISSUER": "Awash Bank", "ACQUIRER": "CBE", "AMOUNT": 0.0},
    ]

    df = generate_nbe_report(records, mode_key="pos")
    assert not df.empty
    assert "BANKS" in df.columns
    assert "PURCHASE As Issuer (Count)" in df.columns
    assert "PURCHASE As Acquirer (Amount ETB)" in df.columns

    # CBE as Issuer: 1 txn (1000.0)
    cbe_row = df[df["BANKS"] == "CBE"].iloc[0]
    assert cbe_row["PURCHASE As Issuer (Count)"] == 1
    assert cbe_row["PURCHASE As Issuer (Amount ETB)"] == 1000.0
    assert cbe_row["PURCHASE As Acquirer (Count)"] == 1
    assert cbe_row["PURCHASE As Acquirer (Amount ETB)"] == 500.0

    # Total row
    tot_row = df[df["BANKS"] == "Total"].iloc[0]
    assert tot_row["PURCHASE As Issuer (Count)"] == 2
    assert tot_row["PURCHASE As Issuer (Amount ETB)"] == 1500.0
    assert tot_row["PURCHASE As Acquirer (Count)"] == 2
    assert tot_row["PURCHASE As Acquirer (Amount ETB)"] == 1500.0


def test_generate_nbe_report_atm():
    """Verify ATM NBE Report filtering and aggregation."""
    records = [
        {"TRANS_TYPE": "ATM Cash withdrawal", "RESP": "-1.0", "ISSUER": "Abay Bank", "ACQUIRER": "Wegagen Bank", "AMOUNT": 400.0},
        {"TRANS_TYPE": "cash withdrawal", "RESP": -1, "ISSUER": "Wegagen Bank", "ACQUIRER": "BOA", "AMOUNT": 600.0},
    ]

    df = generate_nbe_report(records, mode_key="atm")
    tot_row = df[df["BANKS"] == "Total"].iloc[0]
    assert tot_row["CASH WITHDRAWAL As Issuer (Count)"] == 2
    assert tot_row["CASH WITHDRAWAL As Issuer (Amount ETB)"] == 1000.0
    assert tot_row["CASH WITHDRAWAL As Acquirer (Count)"] == 2
    assert tot_row["CASH WITHDRAWAL As Acquirer (Amount ETB)"] == 1000.0


def test_build_nbe_report_excel():
    """Verify Excel workbook generation."""
    records = [
        {"TRANS_TYPE": "POS purchase", "RESP": "-1", "ISSUER": "CBE", "ACQUIRER": "BOA", "AMOUNT": 1250.50},
    ]
    df = generate_nbe_report(records, mode_key="pos")
    excel_bytes = build_nbe_report_excel(df, mode_key="pos")
    assert len(excel_bytes) > 0

    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active
    assert "NBE REPORT" in ws["A1"].value
    assert ws["B2"].value == "BANKS"
    assert ws["C2"].value == "PURCHASE As Issuer"


def test_generate_nbe_report_pos_decline():
    """Verify POS decline NBE Report counts all decline txn and includes amounts."""
    records = [
        # Both declined purchases counted, regardless of RESP code
        {"TRANS_TYPE": "POS purchase", "RESP_CODE": "901", "ISSUER": "Commercial Bank", "ACQUIRER": "Abyssinia Bank", "AMOUNT": 1000.0},
        {"TRANS_TYPE": "purchase", "RESP_CODE": "914", "ISSUER": "Awash Bank", "ACQUIRER": "CBE", "AMOUNT": 500.0},
        # Non-matching trans type (fee rows etc.) excluded
        {"TRANS_TYPE": "QR purchase", "RESP_CODE": "901", "ISSUER": "Awash Bank", "ACQUIRER": "CBE", "AMOUNT": 300.0},
    ]

    df = generate_nbe_report(records, mode_key="pos_decline")
    assert not df.empty
    assert "PURCHASE As Issuer (Count)" in df.columns
    assert "PURCHASE As Acquirer (Amount ETB)" in df.columns

    cbe_row = df[df["BANKS"] == "CBE"].iloc[0]
    assert cbe_row["PURCHASE As Issuer (Count)"] == 1
    assert cbe_row["PURCHASE As Acquirer (Count)"] == 1
    assert cbe_row["PURCHASE As Acquirer (Amount ETB)"] == 500.0

    tot_row = df[df["BANKS"] == "Total"].iloc[0]
    assert tot_row["PURCHASE As Issuer (Count)"] == 2
    assert tot_row["PURCHASE As Issuer (Amount ETB)"] == 1500.0


def test_generate_nbe_report_balance_inquiry():
    """Verify balance inquiry report counts all records without amounts."""
    records = [
        {"TRANS_TYPE": "POS balance inquiry", "RESP": "801", "ISSUER": "Abay Bank", "ACQUIRER": "Wegagen Bank", "AMOUNT": 0.0},
        {"TRANS_TYPE": "balance inquiry", "RESP": "-1", "ISSUER": "CBE", "ACQUIRER": "Dashen Bank", "AMOUNT": 0.0},
        {"TRANS_TYPE": "POS balance inquiry", "RESP": "901", "ISSUER": "CBE", "ACQUIRER": "BOA", "AMOUNT": 0.0},
        # Excluded: purchase transaction
        {"TRANS_TYPE": "purchase", "RESP": "-1", "ISSUER": "CBE", "ACQUIRER": "BOA", "AMOUNT": 100.0},
    ]

    df = generate_nbe_report(records, mode_key="balance_inquiry")
    assert not df.empty
    assert "BALANCE INQUIRY As Issuer (Count)" in df.columns
    assert "BALANCE INQUIRY As Acquirer (Count)" in df.columns
    assert "Amount" not in " ".join(df.columns)

    cbe_row = df[df["BANKS"] == "CBE"].iloc[0]
    assert cbe_row["BALANCE INQUIRY As Issuer (Count)"] == 2
    assert cbe_row["BALANCE INQUIRY As Acquirer (Count)"] == 0

    dashen_row = df[df["BANKS"] == "Dashen Bank"].iloc[0]
    assert dashen_row["BALANCE INQUIRY As Acquirer (Count)"] == 1

    tot_row = df[df["BANKS"] == "Total"].iloc[0]
    assert tot_row["BALANCE INQUIRY As Issuer (Count)"] == 3
    assert tot_row["BALANCE INQUIRY As Acquirer (Count)"] == 3


def test_build_nbe_report_excel_balance_inquiry_4col():
    """Verify 4-column layout for balance inquiry workbook (no amount columns)."""
    records = [
        {"TRANS_TYPE": "POS balance inquiry", "RESP": "801", "ISSUER": "Abay Bank", "ACQUIRER": "Wegagen Bank", "AMOUNT": 0.0},
        {"TRANS_TYPE": "balance inquiry", "RESP": "-1", "ISSUER": "CBE", "ACQUIRER": "Dashen Bank", "AMOUNT": 0.0},
    ]
    df = generate_nbe_report(records, mode_key="balance_inquiry")
    excel_bytes = build_nbe_report_excel(df, mode_key="balance_inquiry")
    assert len(excel_bytes) > 0

    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active
    assert "NBE REPORT" in ws["A1"].value
    assert ws["A3"].value == "S/N"
    assert ws["B3"].value == "BANKS"
    assert ws["C3"].value == "Count"
    assert ws["D3"].value == "Count"
    assert ws["E3"].value is None
    assert ws["F3"].value is None


def test_generate_sett_sum_report_and_excel():
    """Verify Sett(Sum) merges same bank name across both sheets and sums columns."""
    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "iss"
    ws1.append(["ISS_BANKS", "CASH_WITHDRAWAL", "AMOUNT_CW", "BALANCE_INQUIRY", "PURCHASE", "AMOUNT_POS", "STATEMENT"])
    ws1.append(["Abay Bank", 11, 6000, 2, 0, 0, 0])
    ws1.append(["Commercial Bank", 5, 10000, 1, 2, 300, 0])
    ws1.append(["Abyssinia", 3, 2000, 0, 1, 100, 0])
    ws1.append([None, 19, 18000, 3, 3, 400, 0])  # total-ish row (no bank name) -> skipped

    ws2 = wb.create_sheet("acq")
    ws2.append(["ACQ_BANKS", "CASH_WITHDRAWAL", "AMOUNT_CW", "BALANCE_INQUIRY", "PURCHASE", "AMOUNT_POS", "STATEMENT"])
    ws2.append(["Abay", 37, 402790, 50, 4, 100, 0])
    ws2.append(["CBE", 1, 500, 0, 0, 0, 1])

    buf = io.BytesIO()
    wb.save(buf)
    wb.close()

    df = generate_sett_sum_report(buf.getvalue())
    assert not df.empty

    # names normalized to STANDARD_NBE_BANKS
    banks = set(df["BANKS"].tolist())
    assert "Abay Bank" in banks
    assert "CBE" in banks
    assert "BOA" in banks

    # Abay: 11+37=48 CW, 6000+402790=408790 AMOUNT_CW, 2+50=52 BI, 0+4=4 PURCHASE
    abay = df[df["BANKS"] == "Abay Bank"].iloc[0]
    assert abay["CASH_WITHDRAWAL"] == 48
    assert abay["AMOUNT_CW"] == 408790
    assert abay["BALANCE_INQUIRY"] == 52
    assert abay["PURCHASE"] == 4

    # CBE alias ("Commercial Bank" -> CBE) merged: 5+1=6 CW
    cbe = df[df["BANKS"] == "CBE"].iloc[0]
    assert cbe["CASH_WITHDRAWAL"] == 6
    assert cbe["STATEMENT"] == 1

    # Total equals grand sum of every numeric column
    tot = df[df["BANKS"] == "Total"].iloc[0]
    numeric_cols = ["CASH_WITHDRAWAL", "AMOUNT_CW", "BALANCE_INQUIRY", "PURCHASE", "AMOUNT_POS", "STATEMENT"]
    data = df[df["BANKS"] != "Total"]
    for col in numeric_cols:
        assert tot[col] == sum(data[col].fillna(0))

    excel_bytes = build_sett_sum_report_excel(df)
    assert len(excel_bytes) > 0
    wb2 = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws_out = wb2.active
    assert ws_out["A1"].value is not None
    assert "BANKS" in str(ws_out["B2"].value)
