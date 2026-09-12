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


def test_generate_nbe_report_atm_decline():
    """ATM Decline report counts all response codes except the excluded ones.

    Excluded codes: -1, 503, 821, 862, 901, 904, 911, 912, 915. Records with
    those codes are skipped; every other code (any trans type) is counted.
    """
    records = [
        # success code -1 excluded
        {"TRANS_TYPE": "ATM Cash withdrawal", "RESP": "-1.0", "ISSUER": "Abay Bank", "ACQUIRER": "Wegagen Bank", "AMOUNT": 400.0},
        # known decline codes excluded
        {"TRANS_TYPE": "cash withdrawal", "RESP": "503", "ISSUER": "Abay Bank", "ACQUIRER": "BOA", "AMOUNT": 0},
        {"TRANS_TYPE": "Cash withdrawal", "RESP": 915, "ISSUER": "CBE", "ACQUIRER": "BOA", "AMOUNT": 0},
        {"TRANS_TYPE": "cash withdrawal", "RESP": "901.0", "ISSUER": "CBE", "ACQUIRER": "Abay Bank", "AMOUNT": 0},
        {"TRANS_TYPE": "ATM Cash withdrawal", "RESP": "912", "ISSUER": "CBE", "ACQUIRER": "BOA", "AMOUNT": 0},
        # other response codes are counted (both trans types)
        {"TRANS_TYPE": "Cash withdrawal", "RESP": "802", "ISSUER": "CBE", "ACQUIRER": "Wegagen Bank", "AMOUNT": 0},
        {"TRANS_TYPE": "ATM Balance inquiry", "RESP": "820", "ISSUER": "CBE", "ACQUIRER": "BOA", "AMOUNT": 0},
        {"TRANS_TYPE": "balance inquiry", "RESP": "953.0", "ISSUER": "Abay Bank", "ACQUIRER": "CBE", "AMOUNT": 0},
        {"TRANS_TYPE": "Cash withdrawal", "RESP": 906, "ISSUER": "Wegagen Bank", "ACQUIRER": "Abay Bank", "AMOUNT": 0},
    ]

    df = generate_nbe_report(records, mode_key="atm_decline")
    assert "ATM DECLINE RESPONSE CODES As Issuer (Count)" in df.columns
    assert "ATM DECLINE RESPONSE CODES As Acquirer (Count)" in df.columns
    assert "ATM DECLINE RESPONSE CODES As Issuer (Amount ETB)" in df.columns
    assert "ATM DECLINE RESPONSE CODES As Acquirer (Amount ETB)" in df.columns

    # Abay: issuer for 953.0 (amount 0), acquirer for 906 (amount 0); the -1
    # row (400.0) and the 503 / 901 rows are excluded so no amount leaks in
    abay_row = df[df["BANKS"] == "Abay Bank"].iloc[0]
    assert abay_row["ATM DECLINE RESPONSE CODES As Issuer (Count)"] == 1
    assert abay_row["ATM DECLINE RESPONSE CODES As Acquirer (Count)"] == 1
    assert abay_row["ATM DECLINE RESPONSE CODES As Issuer (Amount ETB)"] == 0.0
    assert abay_row["ATM DECLINE RESPONSE CODES As Acquirer (Amount ETB)"] == 0.0

    cbe_row = df[df["BANKS"] == "CBE"].iloc[0]
    assert cbe_row["ATM DECLINE RESPONSE CODES As Issuer (Count)"] == 2
    assert cbe_row["ATM DECLINE RESPONSE CODES As Acquirer (Count)"] == 1

    wegagen_row = df[df["BANKS"] == "Wegagen Bank"].iloc[0]
    assert wegagen_row["ATM DECLINE RESPONSE CODES As Issuer (Count)"] == 1
    assert wegagen_row["ATM DECLINE RESPONSE CODES As Acquirer (Count)"] == 1

    tot_row = df[df["BANKS"] == "Total"].iloc[0]
    assert tot_row["ATM DECLINE RESPONSE CODES As Issuer (Count)"] == 4
    assert tot_row["ATM DECLINE RESPONSE CODES As Acquirer (Count)"] == 4
    assert tot_row["ATM DECLINE RESPONSE CODES As Issuer (Amount ETB)"] == 0.0
    assert tot_row["ATM DECLINE RESPONSE CODES As Acquirer (Amount ETB)"] == 0.0


def test_generate_nbe_report_atm_decline_amounts():
    """Amounts from counted (non-excluded) transactions are aggregated."""
    records = [
        {"TRANS_TYPE": "ATM Cash withdrawal", "RESP": "802", "ISSUER": "Abay Bank", "ACQUIRER": "Wegagen Bank", "AMOUNT": 500.0},
        {"TRANS_TYPE": "Cash withdrawal", "RESP": "820", "ISSUER": "CBE", "ACQUIRER": "Abay Bank", "AMOUNT": 250.5},
        {"TRANS_TYPE": "ATM Balance inquiry", "RESP": "820", "ISSUER": "Wegagen Bank", "ACQUIRER": "CBE", "AMOUNT": 300.0},
        # excluded codes: even though they carry amounts, nothing is counted
        {"TRANS_TYPE": "Cash withdrawal", "RESP": "901", "ISSUER": "Abay Bank", "ACQUIRER": "CBE", "AMOUNT": 9999.0},
        {"TRANS_TYPE": "Cash withdrawal", "RESP": "-1", "ISSUER": "CBE", "ACQUIRER": "Abay Bank", "AMOUNT": 8888.0},
    ]

    df = generate_nbe_report(records, mode_key="atm_decline")

    abay = df[df["BANKS"] == "Abay Bank"].iloc[0]
    assert abay["ATM DECLINE RESPONSE CODES As Issuer (Count)"] == 1
    assert abay["ATM DECLINE RESPONSE CODES As Issuer (Amount ETB)"] == 500.0
    assert abay["ATM DECLINE RESPONSE CODES As Acquirer (Count)"] == 1
    assert abay["ATM DECLINE RESPONSE CODES As Acquirer (Amount ETB)"] == 250.5

    cbe = df[df["BANKS"] == "CBE"].iloc[0]
    assert cbe["ATM DECLINE RESPONSE CODES As Issuer (Count)"] == 1
    assert cbe["ATM DECLINE RESPONSE CODES As Issuer (Amount ETB)"] == 250.5
    assert cbe["ATM DECLINE RESPONSE CODES As Acquirer (Count)"] == 1
    assert cbe["ATM DECLINE RESPONSE CODES As Acquirer (Amount ETB)"] == 300.0

    wegagen = df[df["BANKS"] == "Wegagen Bank"].iloc[0]
    assert wegagen["ATM DECLINE RESPONSE CODES As Issuer (Count)"] == 1
    assert wegagen["ATM DECLINE RESPONSE CODES As Issuer (Amount ETB)"] == 300.0
    assert wegagen["ATM DECLINE RESPONSE CODES As Acquirer (Count)"] == 1
    assert wegagen["ATM DECLINE RESPONSE CODES As Acquirer (Amount ETB)"] == 500.0

    tot = df[df["BANKS"] == "Total"].iloc[0]
    assert tot["ATM DECLINE RESPONSE CODES As Issuer (Count)"] == 3
    assert tot["ATM DECLINE RESPONSE CODES As Issuer (Amount ETB)"] == 1050.5
    assert tot["ATM DECLINE RESPONSE CODES As Acquirer (Count)"] == 3
    assert tot["ATM DECLINE RESPONSE CODES As Acquirer (Amount ETB)"] == 1050.5


def test_build_nbe_report_excel_atm_decline_6col():
    """Verify the 6-column Count + Amount layout for the ATM decline workbook."""
    records = [
        {"TRANS_TYPE": "ATM Cash withdrawal", "RESP": "802", "ISSUER": "Abay Bank", "ACQUIRER": "Wegagen Bank", "AMOUNT": 500.0},
        {"TRANS_TYPE": "Cash withdrawal", "RESP": "915", "ISSUER": "CBE", "ACQUIRER": "BOA", "AMOUNT": 0},
    ]
    df = generate_nbe_report(records, mode_key="atm_decline")
    excel_bytes = build_nbe_report_excel(df, mode_key="atm_decline")
    assert len(excel_bytes) > 0

    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active
    assert "ATM DECLINE" in ws["A1"].value
    assert ws["A3"].value == "S/N"
    assert ws["B3"].value == "BANKS"
    assert ws["C3"].value == "Count"
    assert ws["D3"].value == "Amount (ETB)"
    assert ws["E3"].value == "Count"
    assert ws["F3"].value == "Amount (ETB)"
    # the CBE 915 row is excluded, so only Abay/Wegagen are counted
    tot_row = df[df["BANKS"] == "Total"].iloc[0]
    assert tot_row["ATM DECLINE RESPONSE CODES As Issuer (Count)"] == 1
    assert tot_row["ATM DECLINE RESPONSE CODES As Acquirer (Count)"] == 1
    assert tot_row["ATM DECLINE RESPONSE CODES As Issuer (Amount ETB)"] == 500.0


def test_generate_sett_sum_report_label_collision():
    """Both sheets may spell the bank header 'ISS_BANKS'; sides must not collide."""
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "s1"
    ws1.append(["ISS_BANKS", "CASH_WITHDRAWAL", "AMOUNT_CW"])
    ws1.append(["Abay Bank", 10, 100])
    ws2 = wb.create_sheet("s2")
    ws2.append(["ISS_BANKS", "CASH_WITHDRAWAL", "AMOUNT_CW"])
    ws2.append(["Abay Bank", 20, 200])
    buf = io.BytesIO()
    wb.save(buf)
    wb.close()

    reports = generate_sett_sum_report(buf.getvalue())
    assert set(reports.keys()) == {"ISS_Banks", "Acquirer_Banks"}
    iss = reports["ISS_Banks"]
    acq = reports["Acquirer_Banks"]
    assert iss[iss["BANKS"] == "Abay Bank"].iloc[0]["CASH_WITHDRAWAL"] == 10
    assert acq[acq["BANKS"] == "Abay Bank"].iloc[0]["CASH_WITHDRAWAL"] == 20


def test_generate_sett_sum_report_and_excel():
    """Verify Sett(Sum) keeps ISS and ACQ sides in separate sheets and sums columns."""
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

    reports = generate_sett_sum_report(buf.getvalue())
    assert set(reports.keys()) == {"ISS_Banks", "Acquirer_Banks"}

    iss = reports["ISS_Banks"]
    assert not iss.empty

    # names normalized to STANDARD_NBE_BANKS
    banks = set(iss["BANKS"].tolist())
    assert "Abay Bank" in banks
    assert "CBE" in banks
    assert "BOA" in banks

    # Abay (issuer side): 11 CW, 6000 AMOUNT_CW, 2 BI, 0 PURCHASE
    abay = iss[iss["BANKS"] == "Abay Bank"].iloc[0]
    assert abay["CASH_WITHDRAWAL"] == 11
    assert abay["AMOUNT_CW"] == 6000
    assert abay["BALANCE_INQUIRY"] == 2
    assert abay["PURCHASE"] == 0

    # CBE alias ("Commercial Bank" -> CBE): 5 CW on issuer side only
    cbe = iss[iss["BANKS"] == "CBE"].iloc[0]
    assert cbe["CASH_WITHDRAWAL"] == 5
    assert cbe["STATEMENT"] == 0

    # ISS Total equals grand sum of every numeric column
    tot = iss[iss["BANKS"] == "Total"].iloc[0]
    numeric_cols = ["CASH_WITHDRAWAL", "AMOUNT_CW", "BALANCE_INQUIRY", "PURCHASE", "AMOUNT_POS", "STATEMENT"]
    data = iss[iss["BANKS"] != "Total"]
    for col in numeric_cols:
        assert tot[col] == sum(data[col].fillna(0))

    # Acquirer side is separate and carries its own figures
    acq = reports["Acquirer_Banks"]
    abay_acq = acq[acq["BANKS"] == "Abay Bank"].iloc[0]
    assert abay_acq["CASH_WITHDRAWAL"] == 37
    assert abay_acq["AMOUNT_CW"] == 402790
    assert abay_acq["BALANCE_INQUIRY"] == 50
    cbe_acq = acq[acq["BANKS"] == "CBE"].iloc[0]
    assert cbe_acq["STATEMENT"] == 1
    tot_acq = acq[acq["BANKS"] == "Total"].iloc[0]
    for col in numeric_cols:
        assert tot_acq[col] == sum(acq[acq["BANKS"] != "Total"][col].fillna(0))

    excel_bytes = build_sett_sum_report_excel(reports)
    assert len(excel_bytes) > 0
    wb2 = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    assert wb2.sheetnames == ["ISS_Banks", "Acquirer_Banks"]
    for sheet_name in wb2.sheetnames:
        ws_out = wb2[sheet_name]
        assert "Report name:" in str(ws_out["A1"].value)
        assert "SETT(SUM)" in str(ws_out["B1"].value)
        assert "BANKS" in str(ws_out["B2"].value)
        assert ws_out.freeze_panes == "A3"
        assert ws_out.sheet_view.showGridLines is False
