"""Top acquirer analysis for ATM, POS and IPS reports.

Parses the raw transaction reports, normalises bank names so that
variants like 'Abay' / 'Abay Bank', 'CBE' / 'Commercial Bank', etc.
are treated as the same bank, filters by transaction type, and returns
the top-3 acquirers (ATM, POS) or top-3 senders/receivers (IPS).
"""

from __future__ import annotations

import io
from collections import Counter
from dataclasses import dataclass, field
from typing import Any

import pandas as pd


# ---------------------------------------------------------------------------
# Bank name normalisation
# ---------------------------------------------------------------------------
# Every alias is mapped to a canonical short name.  The lookup is
# case-insensitive and strips trailing 'bank', 'ebirr', etc.
_BANK_CANONICAL: dict[str, str] = {
    "abay": "ABAY",
    "abay bank": "ABAY",
    "addis": "ADDIS",
    "addis bank": "ADDIS",
    "addis int": "ADDIS",
    "addis international": "ADDIS",
    "ahadu": "AHADU",
    "ahadu ebirr": "AHADU",
    "amhara": "AMHARA",
    "amhara bank": "AMHARA",
    "amharaethbirr": "AMHARA",
    "awash": "AIB",
    "awash bank": "AIB",
    "boa": "BOA",
    "abyssinia": "BOA",
    "abyssinia bank": "BOA",
    "berhan": "BERHAN",
    "berhan bank": "BERHAN",
    "birhan": "BERHAN",
    "birhan bank": "BERHAN",
    "bunna": "BUNNA",
    "bunna bank": "BUNNA",
    "buna": "BUNNA",
    "buna bank": "BUNNA",
    "cbe": "CBE",
    "commercial bank": "CBE",
    "commercial bank of ethiopia": "CBE",
    "cbé": "CBE",
    "cbébirr": "CBE",
    "cbébírr": "CBE",
    "coop": "COOP",
    "coopay-e-birr": "COOP",
    "cbo": "CBO",
    "cbo switch": "CBO",
    "dashen": "DB",
    "dashen bank": "DB",
    "db": "DB",
    "debub": "DEBUB",
    "debub bank": "DEBUB",
    "global": "DEBUB",
    "dedebit": "DEBUB",
    "enat": "ENAT",
    "enat bank": "ENAT",
    "gadaa": "GADA",
    "gadaa bank": "GADA",
    "gada": "GADA",
    "goh betoch": "GOH",
    "h-cash": "HIBRET",
    "hijra": "HIJRA",
    "hijra bank": "HIJRA",
    "hibret": "UB",
    "hibret bank": "UB",
    "kaafi": "KAAFI",
    "kacha": "KACHA",
    "lib": "LIB",
    "lion": "LIB",
    "lion bank": "LIB",
    "mpesa": "MPESA",
    "nib": "NIB",
    "nib bank": "NIB",
    "nib int bank": "NIB",
    "nib international bank": "NIB",
    "nibbirr": "NIB",
    "nisir": "NISIR",
    "oib": "OIB",
    "oromia": "OIB",
    "oromia bank": "OIB",
    "omo": "OMO",
    "rays": "RAYS",
    "raamis": "RAMMIS",
    "raammis bank": "RAMMIS",
    "rammis": "RAMMIS",
    "rammis bank": "RAMMIS",
    "saha": "SAHA",
    "santimpay": "SANTIMPAY",
    "shabelle": "SHABELLE",
    "siinqee": "SINQEE",
    "siinqee bank": "SINQEE",
    "siinqee wallet": "SINQEE",
    "siket": "SIKET",
    "sidama": "SIDAMA",
    "sidama bank": "SIDAMA",
    "tseday": "TSEDAY",
    "tseday bank": "TSEDAY",
    "tsehay": "TSEDAY",
    "tsehay bank": "TSEDAY",
    "ub": "UB",
    "vision fund": "VISION",
    "vitabirr": "VITABIRR",
    "wb": "WB",
    "wegagen": "WB",
    "wegagen bank": "WB",
    "wegagen e-birr": "WB",
    "yagoutpay": "YAGOUT",
    "yaya wallet": "YAYA",
    "zb": "ZB",
    "zemen": "ZB",
    "zemen bank": "ZB",
    "zamzam": "ZAMZAM",
}


def normalize_bank(name: Any) -> str:
    """Return the canonical short name for a bank, or the upper-stripped
    raw value if no mapping exists."""
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    s = str(name).strip()
    if not s:
        return ""
    key = s.lower()
    # try exact match first
    if key in _BANK_CANONICAL:
        return _BANK_CANONICAL[key]
    # strip common suffixes and retry
    for suffix in (" bank", " ebirr", " e-birr", " int", " international"):
        if key.endswith(suffix):
            key = key[: -len(suffix)].strip()
            break
    if key in _BANK_CANONICAL:
        return _BANK_CANONICAL[key]
    return s.upper()


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------
@dataclass
class TopBank:
    name: str           # canonical name shown to the user
    count: int
    percentage: float   # 0-100


@dataclass
class AcquirerResult:
    top3: list[TopBank] = field(default_factory=list)
    total: int = 0
    filtered_rows: int = 0
    warnings: list[str] = field(default_factory=list)


@dataclass
class IpsResult:
    top3_senders: list[TopBank] = field(default_factory=list)
    top3_receivers: list[TopBank] = field(default_factory=list)
    total_banks: int = 0
    warnings: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# ATM analysis
# ---------------------------------------------------------------------------
_ATM_CASH_TYPES = {"atm cash withdrawal", "cash withdrawal"}


def analyze_atm(file_bytes: bytes, filename: str = "ATM") -> AcquirerResult:
    """Parse an ATM daily report and return the top-3 acquirers for
    cash withdrawal transactions."""
    try:
        engine = "xlrd" if file_bytes[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" else "openpyxl"
        df = pd.read_excel(io.BytesIO(file_bytes), header=0, engine=engine)
    except Exception as exc:
        return AcquirerResult(warnings=[f"Could not read '{filename}': {exc}"])

    if "ACQUIRER" not in df.columns or "TRANS_TYPE" not in df.columns:
        return AcquirerResult(warnings=[f"'{filename}' is missing required columns (ACQUIRER, TRANS_TYPE)."])

    # Filter to cash withdrawal only
    mask = df["TRANS_TYPE"].astype(str).str.strip().str.lower().isin(_ATM_CASH_TYPES)
    filtered = df[mask].copy()

    if filtered.empty:
        return AcquirerResult(
            warnings=[f"No ATM cash withdrawal transactions found in '{filename}'."]
        )

    # Normalise and count
    acquirers = filtered["ACQUIRER"].apply(normalize_bank)
    # Drop empty / unknown
    acquirers = acquirers[acquirers != ""]
    total = len(acquirers)
    counts = Counter(acquirers)
    top3 = [
        TopBank(name=name, count=cnt, percentage=round(cnt / total * 100, 1) if total else 0)
        for name, cnt in counts.most_common(3)
    ]
    return AcquirerResult(top3=top3, total=total, filtered_rows=len(filtered))


# ---------------------------------------------------------------------------
# POS analysis
# ---------------------------------------------------------------------------
_POS_PURCHASE_TYPES = {"pos purchase", "purchase"}


def analyze_pos(file_bytes: bytes, filename: str = "POS") -> AcquirerResult:
    """Parse a POS daily report and return the top-3 acquirers for
    purchase transactions."""
    try:
        engine = "xlrd" if file_bytes[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" else "openpyxl"
        df = pd.read_excel(io.BytesIO(file_bytes), header=0, engine=engine)
    except Exception as exc:
        return AcquirerResult(warnings=[f"Could not read '{filename}': {exc}"])

    if "ACQUIRER" not in df.columns or "TRANS_TYPE" not in df.columns:
        return AcquirerResult(warnings=[f"'{filename}' is missing required columns (ACQUIRER, TRANS_TYPE)."])

    # Filter to POS purchase only
    mask = df["TRANS_TYPE"].astype(str).str.strip().str.lower().isin(_POS_PURCHASE_TYPES)
    filtered = df[mask].copy()

    if filtered.empty:
        return AcquirerResult(
            warnings=[f"No POS purchase transactions found in '{filename}'."]
        )

    acquirers = filtered["ACQUIRER"].apply(normalize_bank)
    acquirers = acquirers[acquirers != ""]
    total = len(acquirers)
    counts = Counter(acquirers)
    top3 = [
        TopBank(name=name, count=cnt, percentage=round(cnt / total * 100, 1) if total else 0)
        for name, cnt in counts.most_common(3)
    ]
    return AcquirerResult(top3=top3, total=total, filtered_rows=len(filtered))


# ---------------------------------------------------------------------------
# IPS analysis
# ---------------------------------------------------------------------------
def analyze_ips(file_bytes: bytes, filename: str = "IPS") -> IpsResult:
    """Parse an IPS success report and return the top-3 senders
    (ACQUIRER_TXN_COUNT) and top-3 receivers (ISSUER_TXN_COUNT)."""
    try:
        engine = "xlrd" if file_bytes[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" else "openpyxl"
        df = pd.read_excel(io.BytesIO(file_bytes), header=0, engine=engine)
    except Exception as exc:
        return IpsResult(warnings=[f"Could not read '{filename}': {exc}"])

    required = {"BANK_NAME", "ACQUIRER_TXN_COUNT", "ISSUER_TXN_COUNT"}
    missing = required - set(df.columns)
    if missing:
        return IpsResult(warnings=[f"'{filename}' is missing columns: {', '.join(sorted(missing))}"])

    # Normalise bank names
    df = df.copy()
    df["_norm"] = df["BANK_NAME"].apply(normalize_bank)

    # Aggregate by normalised bank (some variants may map to the same name)
    grouped = df.groupby("_norm", as_index=False).agg({
        "ACQUIRER_TXN_COUNT": "sum",
        "ISSUER_TXN_COUNT": "sum",
    })

    total_banks = len(grouped)

    # Top 3 senders
    top_send = grouped.nlargest(3, "ACQUIRER_TXN_COUNT")
    top3_senders = [
        TopBank(
            name=row["_norm"],
            count=int(row["ACQUIRER_TXN_COUNT"]),
            percentage=0,
        )
        for _, row in top_send.iterrows()
    ]

    # Top 3 receivers
    top_recv = grouped.nlargest(3, "ISSUER_TXN_COUNT")
    top3_receivers = [
        TopBank(
            name=row["_norm"],
            count=int(row["ISSUER_TXN_COUNT"]),
            percentage=0,
        )
        for _, row in top_recv.iterrows()
    ]

    return IpsResult(
        top3_senders=top3_senders,
        top3_receivers=top3_receivers,
        total_banks=total_banks,
    )
