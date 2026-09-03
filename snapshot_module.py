"""Digital Transaction Value Snapshot - pure-Python port.

Ported from digital-transaction-snapshot/public/js/app.js so the same
module is available on the deployed Streamlit app without a Node server.
"""

import base64
import html as _html
import os

import pandas as pd

REPORT_DEFAULTS = {
    "title": "DIGITAL TRANSACTION VALUE SNAPSHOT",
    "organization": "ETHSWITCH",
    "brand": "EthioPay",
    "tagline1": "Making Payment Simple and Affordable",
    "tagline2": "One Payment. Every Possibility.",
    "subtitle": "Performance Overview by Service",
    "date": "31.08.26",
}

SERVICE_DEFAULTS = [
    {"name": "ATM SUCCESS RATE", "type": "success-rate", "transactionVolume": 98.7,
     "totalValue": 0, "target": 98, "keyMessage": "ATM success rate", "highlighted": False},
    {"name": "POS SUCCESS RATE", "type": "success-rate", "transactionVolume": 97.5,
     "totalValue": 0, "target": 97, "keyMessage": "POS success rate", "highlighted": False},
    {"name": "P2P SUCCESS RATE", "type": "success-rate", "transactionVolume": 99.1,
     "totalValue": 0, "target": 99, "keyMessage": "P2P success rate", "highlighted": False},
    {"name": "CASH WITHDRAWAL", "type": "financial", "transactionVolume": 306455,
     "totalValue": 455114740.00, "target": 340000, "keyMessage": "Lower-value transactions", "highlighted": False},
    {"name": "BALANCE INQUIRY & MINI STATEMENT", "type": "non-financial",
     "transactionVolume": 18406, "totalValue": 0, "target": 20000, "keyMessage": "Non-financial service",
     "highlighted": False},
    {"name": "POS PURCHASE", "type": "financial", "transactionVolume": 10189,
     "totalValue": 33248484.72, "target": 12000, "keyMessage": "Moderate transaction value", "highlighted": False},
    {"name": "IPS P2P", "type": "financial", "transactionVolume": 926648,
     "totalValue": 4049226900.62, "target": 1000000, "keyMessage": "Volume leader and value driver", "highlighted": False},
    {"name": "QR", "type": "financial", "transactionVolume": 30248,
     "totalValue": 268061663.69, "target": 35000, "keyMessage": "", "highlighted": True},
    {"name": "RTP", "type": "financial", "transactionVolume": 5400,
     "totalValue": 1850000000.00, "target": 6000, "keyMessage": "", "highlighted": False},
    {"name": "NPG (CARD AND ONLINE)", "type": "financial", "transactionVolume": 21203,
     "totalValue": 6925000000.00, "target": 25000, "keyMessage": "", "highlighted": False},
]

_IMG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "snapshot")


def _img_data_uri(filename):
    """Return a base64 jpeg data URI for an image in static/snapshot, or ''."""
    path = os.path.join(_IMG_DIR, filename)
    try:
        with open(path, "rb") as f:
            return "data:image/jpeg;base64," + base64.b64encode(f.read()).decode("ascii")
    except OSError:
        return ""


def _phone_mockup():
    """Decorative CSS phone mockup: orange screen + QR-style pattern."""
    return (
        '<div class="phone" title="QR">'
        '<div class="phone-screen">'
        '<svg class="qr-deco" viewBox="0 0 21 21" aria-hidden="true">'
        '<rect x="1" y="1" width="5" height="5" fill="#fff"/>'
        '<rect x="2" y="2" width="3" height="3" fill="#F4511E"/>'
        '<rect x="15" y="1" width="5" height="5" fill="#fff"/>'
        '<rect x="16" y="2" width="3" height="3" fill="#F4511E"/>'
        '<rect x="1" y="15" width="5" height="5" fill="#fff"/>'
        '<rect x="2" y="16" width="3" height="3" fill="#F4511E"/>'
        '<rect x="8" y="3" width="2" height="2" fill="#fff"/>'
        '<rect x="11" y="2" width="2" height="1" fill="#fff"/>'
        '<rect x="7" y="6" width="1" height="2" fill="#fff"/>'
        '<rect x="13" y="6" width="2" height="1" fill="#fff"/>'
        '<rect x="9" y="8" width="3" height="2" fill="#fff"/>'
        '<rect x="5" y="9" width="2" height="1" fill="#fff"/>'
        '<rect x="14" y="10" width="2" height="1" fill="#fff"/>'
        '<rect x="7" y="12" width="2" height="1" fill="#fff"/>'
        '<rect x="11" y="13" width="2" height="2" fill="#fff"/>'
        '<rect x="3" y="13" width="2" height="1" fill="#fff"/>'
        '<rect x="16" y="13" width="2" height="2" fill="#fff"/>'
        '<rect x="8" y="16" width="2" height="1" fill="#fff"/>'
        '<rect x="12" y="17" width="2" height="1" fill="#fff"/>'
        '<rect x="5" y="17" width="1" height="2" fill="#fff"/>'
        '<rect x="17" y="9" width="1" height="2" fill="#fff"/>'
        '<rect x="4" y="8" width="2" height="1" fill="#fff"/>'
        '<rect x="16" y="6" width="1" height="2" fill="#fff"/>'
        '<rect x="9" y="15" width="2" height="1" fill="#fff"/>'
        "</svg></div><div class='phone-notch'></div></div>"
    )


REPORT_CSS = """
:root {
  --navy-900: #0B2A5B; --navy-100: #d9e2f0; --navy-500: #416eb4;
  --brand-500: #F4511E; --orange-light: #FFF3E0;
}
* { box-sizing: border-box; }
body { margin: 0; background: #eef1f6; font-family: 'Plus Jakarta Sans', 'Segoe UI', sans-serif; }
.toolbar { display: flex; justify-content: flex-end; gap: 10px; padding: 12px 20px; }
.toolbar button {
  font-family: inherit; font-size: 12px; font-weight: 700; cursor: pointer;
  border: 1px solid #d7dbe3; background: #ffffff; color: #0B2A5B;
  border-radius: 8px; padding: 8px 16px; transition: all .2s;
}
.toolbar button:hover { border-color: #F4511E; color: #F4511E; }
.report-landscape {
  width: 1038px; height: 735px; background: #ffffff; margin: 0 auto;
  display: flex; flex-direction: column; overflow: hidden;
  box-shadow: 0 4px 24px rgba(11,42,91,0.12), 0 1px 4px rgba(0,0,0,0.06);
  border: 1px solid #e5e7eb;
}
.report-header { display: flex; align-items: center; justify-content: space-between;
  padding: 8px 20px; min-height: 56px;
  background: linear-gradient(135deg, #0B2A5B 0%, #1e3a6b 100%); color: #fff; }
.header-left { display: flex; align-items: center; }
.header-center { flex: 1; text-align: center; padding: 0 16px; }
.header-right { display: flex; align-items: center; flex-shrink: 0; }
.header-right-inner { display: flex; align-items: center; gap: 10px; }
.header-sep { width: 1px; align-self: stretch; background: rgba(255,255,255,.2); }
.org-badge { display: flex; align-items: center; gap: 8px; }
.org-icon { width: 36px; height: 36px; background: rgba(255,255,255,.15);
  border: 1.5px solid rgba(255,255,255,.3); border-radius: 6px;
  display: flex; align-items: center; justify-content: center;
  font-size: 11px; font-weight: 800; letter-spacing: .5px; color: #fff; }
.org-name { font-size: 12px; font-weight: 700; letter-spacing: 1px; }
.org-sub { font-size: 8px; color: rgba(255,255,255,.6); letter-spacing: .5px;
  text-transform: none; white-space: nowrap; max-width: 150px; }
.org-logo { height: 34px; width: auto; border-radius: 4px; }
.bird-logo { height: 30px; width: auto; border-radius: 4px; flex-shrink: 0; }
.bird-group { display: flex; flex-direction: column; align-items: flex-end; gap: 2px; }
.bird-text { text-align: right; line-height: 1.1; }
.bird-name { font-size: 10px; font-weight: 800; letter-spacing: .5px; color: #fff; }
.bird-tagline { font-size: 6.5px; color: rgba(255,255,255,.7); letter-spacing: .5px; white-space: nowrap; }
.phone { width: 24px; height: 42px; background: #0B2A5B; border: 1.5px solid rgba(255,255,255,.35);
  border-radius: 6px; padding: 3px; flex-shrink: 0; position: relative; }
.phone-screen { width: 100%; height: 100%; border-radius: 3px; overflow: hidden;
  background: linear-gradient(160deg, #F4511E 0%, #ff8a55 100%);
  display: flex; align-items: center; justify-content: center; }
.qr-deco { width: 15px; height: 15px; display: block; }
.phone-notch { position: absolute; top: 0; left: 50%; transform: translateX(-50%);
  width: 9px; height: 2px; background: rgba(255,255,255,.6); border-radius: 2px; }
.brand-badge { display: inline-block; background: #F4511E; color: #fff; font-size: 10px;
  font-weight: 800; letter-spacing: 2px; padding: 2px 12px; border-radius: 3px; margin-bottom: 4px; }
.report-title { font-size: 15px; font-weight: 900; letter-spacing: 2px;
  text-transform: uppercase; margin: 0; line-height: 1.2; }
.report-subtitle { font-size: 9px; color: rgba(255,255,255,.65); margin: 2px 0 0;
  letter-spacing: 1px; text-transform: uppercase; }
.date-badge { display: flex; flex-direction: column; align-items: flex-end; gap: 1px; }
.date-label { font-size: 7px; color: rgba(255,255,255,.5); letter-spacing: 1.5px; text-transform: uppercase; }
.date-value { font-size: 14px; font-weight: 700; letter-spacing: .5px;
  background: rgba(255,255,255,.1); padding: 2px 10px; border-radius: 4px;
  border: 1px solid rgba(255,255,255,.15); }
.report-table-wrapper { flex: 1; overflow: hidden; }
.report-table { width: 100%; border-collapse: collapse; font-size: 11px; }
.report-table thead tr { background: #0B2A5B; color: #fff; }
.report-table th { padding: 5px 8px; font-size: 8px; font-weight: 700; letter-spacing: 1px;
  text-transform: uppercase; text-align: left; white-space: nowrap; }
.th-inner { display: flex; align-items: center; gap: 4px; }
.th-num { justify-content: flex-end; }
.th-inner svg { width: 10px; height: 10px; flex-shrink: 0; }
.th-icon { display: inline-flex; width: 16px; height: 16px; border-radius: 3px;
  align-items: center; justify-content: center; flex-shrink: 0; opacity: .9; }
.th-icon-blue { background: #2d5190; }
.th-icon-purple { background: #7a4fa0; }
.th-icon-green { background: #2f9e62; }
.th-icon-teal { background: #156e8a; }
.th-icon svg { width: 9px; height: 9px; }
.th-service { width: 17%; }
.th-volume { width: 12%; text-align: right; }
.th-target { width: 13%; text-align: right; }
.th-ach { width: 12%; text-align: right; }
.th-value { width: 15%; text-align: right; }
.th-avg { width: 18%; text-align: right; }
.th-message { width: 13%; text-align: left; }
.report-table td { padding: 4px 8px; vertical-align: middle; border-bottom: 1px solid #eef1f5; height: 30px; }
.report-table tbody tr:last-child td { border-bottom: none; }
.report-table tbody tr:nth-child(even) { background: #fafbfc; }
.report-table tbody tr.highlight-row { background: #FFF3E0 !important; }
.report-table tbody tr.highlight-row td { border-bottom-color: #ffe0cc; }
.svc-cell { display: flex; align-items: center; gap: 7px; }
.svc-icon { width: 22px; height: 22px; border-radius: 5px; display: flex;
  align-items: center; justify-content: center; flex-shrink: 0; }
.svc-icon.financial { background: #d9e2f0; color: #0B2A5B; }
.svc-icon.non-financial { background: #e8eaed; color: #5f6368; }
.svc-icon.cash { background: #dff0e2; color: #1a7a3a; }
.svc-icon.pos { background: #e3edfb; color: #1e56c0; }
.svc-icon.p2p { background: #f3e3fb; color: #7a1fa2; }
.svc-icon.qr { background: #fff0e0; color: #e07b0a; }
.svc-icon.landmark { background: #e2f0f7; color: #0f7a9e; }
.svc-icon.smartphone { background: #fde8e8; color: #c0392b; }
.svc-icon.default { background: #d9e2f0; color: #0B2A5B; }
.svc-icon svg { width: 12px; height: 12px; }
.svc-name { font-weight: 700; color: #0B2A5B; font-size: 10px; letter-spacing: .2px; line-height: 1.15; }
.highlight-row .svc-name { color: #F4511E; }
.num-cell { text-align: right; font-variant-numeric: tabular-nums; }
.num-primary { font-weight: 700; font-size: 11px; color: #0B2A5B; line-height: 1.15; }
.highlight-row .num-primary { color: #F4511E; }
.num-dash { color: #9ca3af; font-style: italic; }
.num-total { color: #F4511E; font-weight: 800; }
.report-table tbody tr.total-row td { background: #eef3fb; border-top: 2px solid #0B2A5B; font-weight: 800; color: #0B2A5B; }
.metric-bar-wrap { margin-top: 2px; }
.metric-bar { height: 2px; border-radius: 2px; background: #d9e2f0; overflow: hidden; }
.metric-bar-fill { height: 100%; border-radius: 2px; background: #416eb4; transition: width .4s ease; }
.highlight-row .metric-bar-fill { background: #F4511E; }
.msg-cell { font-size: 9.5px; color: #6b7280; line-height: 1.3; }
.msg-cell.msg-highlight { color: #F4511E; font-weight: 600; }
.msg-badge { display: inline-block; background: #F4511E; color: #fff; font-size: 7px;
  font-weight: 700; letter-spacing: .8px; padding: 1px 6px; border-radius: 3px;
  text-transform: uppercase; margin-top: 2px; }
.report-insights { display: grid; grid-template-columns: 1fr 1.35fr 1.1fr; gap: 8px;
  padding: 8px 12px 10px; background: #f6f8fc; border-top: 2px solid #0B2A5B; }
.insight-card { border-radius: 8px; padding: 8px 12px; }
.volume-leader-card { background: linear-gradient(135deg, #0B2A5B 0%, #1e3a6b 100%); }
.qr-advantage-card { background: linear-gradient(135deg, #F4511E 0%, #ff7a3d 100%); }
.takeaway-card { background: linear-gradient(135deg, #0b7a45 0%, #14a46b 100%); }
.insight-label { font-size: 7.5px; font-weight: 800; letter-spacing: 1.5px;
  color: rgba(255,255,255,.75); text-transform: uppercase; margin-bottom: 4px; }
.insight-service { font-size: 14px; font-weight: 800; color: #fff; margin-bottom: 2px; }
.insight-detail { font-size: 9px; color: rgba(255,255,255,.85); line-height: 1.35; }
.qr-avg-main { text-align: center; padding: 4px 0 6px; border-bottom: 1px solid rgba(255,255,255,.3); margin-bottom: 6px; }
.qr-avg-label { font-size: 6.5px; font-weight: 700; letter-spacing: 1px; color: rgba(255,255,255,.8); text-transform: uppercase; }
.qr-avg-value { font-size: 18px; font-weight: 900; color: #fff; line-height: 1.3; }
.qr-avg-sub { font-size: 8px; color: rgba(255,255,255,.75); }
.qr-comparisons { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 4px; }
.qr-comp { text-align: center; }
.qr-comp-multiplier { font-size: 14px; font-weight: 800; color: #fff; line-height: 1.2; }
.qr-comp-label { font-size: 7px; color: rgba(255,255,255,.8); line-height: 1.2; }
.takeaway-text { font-size: 9.5px; color: #fff; line-height: 1.5; }
.report-footer { display: flex; align-items: center; justify-content: space-between;
  padding: 6px 20px; background: #0B2A5B; color: #fff; min-height: 30px; }
.footer-org-name { font-size: 10px; font-weight: 700; letter-spacing: .5px; }
.footer-tagline { font-size: 8px; color: rgba(255,255,255,.55); margin-left: 8px; font-style: italic; }
.footer-web { display: flex; align-items: center; gap: 5px; font-size: 9px; color: rgba(255,255,255,.7); }
.footer-web svg { width: 12px; height: 12px; }
@media print {
  .toolbar, .toolbar * { display: none !important; }
  body { background: #fff !important; }
  .report-landscape { box-shadow: none; border: none; }
}
"""


def _num(v, default=0):
    if v is None:
        return default
    if isinstance(v, float) and pd.isna(v):
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _bool(v):
    return bool(v)


def fmt_int(n):
    n = _num(n)
    if n <= 0:
        return "-"
    if n.is_integer():
        return f"{int(n):,}"
    return f"{n:,.2f}".rstrip('0').rstrip('.')


def fmt_dec(n):
    n = _num(n)
    if n <= 0:
        return "-"
    return f"{n:,.2f}"


def _icon(name):
    n = (name or "").upper()
    if "CASH" in n or "WITHDRAW" in n:
        return "banknote"
    if "POS" in n or "PURCHASE" in n:
        return "credit-card"
    if "P2P" in n or "IPS" in n:
        return "users"
    if "QR" in n:
        return "qr-code"
    if "BALANCE" in n or "INQUIRY" in n or "MINI" in n:
        return "landmark"
    if "RTP" in n:
        return "arrow-left-right"
    if "NPG" in n or "ONLINE" in n:
        return "globe"
    if "SUCCESS RATE" in n:
        return "percent"
    return "circle-dot"


def _icon_class(name):
    n = (name or "").upper()
    if "CASH" in n or "WITHDRAW" in n:
        return "cash"
    if "POS" in n or "PURCHASE" in n:
        return "pos"
    if "P2P" in n or "IPS" in n:
        return "p2p"
    if "QR" in n:
        return "qr"
    if "BALANCE" in n or "INQUIRY" in n or "MINI" in n:
        return "landmark"
    if "RTP" in n or "NPG" in n or "ONLINE" in n:
        return "landmark"
    if "SUCCESS RATE" in n:
        return "pos"
    return "default"


def _short_name(name):
    n = (name or "").upper()
    if "CASH" in n:
        return "Cash Withdrawal"
    if "POS" in n:
        return "POS"
    if "IPS" in n or "P2P" in n:
        return "P2P"
    if "QR" in n:
        return "QR"
    return name or ""


def _initials(org):
    org = (org or "ES")
    words = [w for w in org.split() if w]
    chars = "".join(w[0] for w in words[:4]) if words else "ES"
    if len(chars) < 2:
        chars = org[:2]
    return chars[:2].upper()


def _esc(text):
    return _html.escape(str(text if text is not None else ""))


def _is_success_rate(s):
    return s.get("type") == "success-rate" or "SUCCESS RATE" in (s.get("name") or "").upper()


def calc_all(services):
    """Mirror of app.js calcAll()."""
    enriched = []
    for idx, s in enumerate(services):
        vol = max(0, _num(s.get("transactionVolume")))
        val = max(0, _num(s.get("totalValue")))
        target = max(0, _num(s.get("target")))
        is_rate = _is_success_rate(s)
        is_fin = s.get("type") == "financial" and val > 0
        avg = (val / vol) if (is_fin and vol > 0) else 0
        ach = None
        if target > 0:
            ach = (vol / target) * 100
        enriched.append({
            "uid": s.get("uid") or f"row_{idx}",
            "name": s.get("name") or f"Service {idx + 1}",
            "type": s.get("type") or "financial",
            "transactionVolume": vol,
            "totalValue": val,
            "target": target,
            "keyMessage": s.get("keyMessage") or "",
            "highlighted": _bool(s.get("highlighted")),
            "isFinancial": is_fin,
            "isSuccessRate": is_rate,
            "achievementPercent": ach,
            "averageTransactionValue": avg,
        })

    countables = [x for x in enriched if not x["isSuccessRate"]]
    max_vol = max((x["transactionVolume"] for x in countables), default=0)
    feas = [x for x in enriched if x["isFinancial"]]
    max_avg = max((x["averageTransactionValue"] for x in feas), default=0)
    for x in enriched:
        x["volumePercent"] = (x["transactionVolume"] / max_vol * 100) if max_vol > 0 else 0
        x["avgPercent"] = (x["averageTransactionValue"] / max_avg * 100) if max_avg > 0 else 0

    volume_leader = max(feas, key=lambda x: x["transactionVolume"], default=None) if feas else None
    by_avg = sorted(feas, key=lambda x: x["averageTransactionValue"], reverse=True) if feas else []
    highest_avg = by_avg[0] if by_avg else None
    highest_total = max(feas, key=lambda x: x["totalValue"], default=None) if feas else None

    qr = next((x for x in enriched if "QR" in (x.get("name") or "").upper()), None)
    qr_advantages = []
    if qr and qr["isFinancial"]:
        others = [x for x in enriched if x["isFinancial"] and x["uid"] != qr["uid"]]

        def _order_key(x):
            n = (x.get("name") or "").upper()
            for i, token in enumerate(["IPS", "POS", "CASH"]):
                if token in n:
                    return i
            return 99

        others.sort(key=lambda x: (_order_key(x), -x["averageTransactionValue"]))
        for o in others:
            if o["averageTransactionValue"] > 0:
                ratio = qr["averageTransactionValue"] / o["averageTransactionValue"]
                qr_advantages.append({
                    "service": o["name"],
                    "ratio": ratio,
                    "label": f"{ratio:.1f}x",
                    "description": "Higher than " + _short_name(o["name"]),
                })

    if qr and qr["isFinancial"] and not (qr.get("keyMessage") or "").strip():
        if highest_avg and highest_avg.get("name") == qr.get("name"):
            qr["keyMessage"] = "HIGHEST average transaction value"

    total = {
        "performance": sum(x["transactionVolume"] for x in countables),
        "target": sum(x["target"] for x in countables),
        "totalValue": sum(x["totalValue"] for x in countables),
    }
    total["achievementPercent"] = (total["performance"] / total["target"] * 100) if total["target"] > 0 else None

    return {
        "services": enriched,
        "volumeLeader": volume_leader,
        "highestAvg": highest_avg,
        "highestTotal": highest_total,
        "qr": qr,
        "qrAdvantages": qr_advantages,
        "takeaway": _takeaway(qr, volume_leader, feas),
        "total": total,
    }


def _takeaway(qr, volume_leader, feas):
    """Mirror of app.js generateTakeaway()."""
    if not qr or not qr["isFinancial"] or not feas:
        return ""
    by_vol = sorted(feas, key=lambda x: x["transactionVolume"], reverse=True)
    qr_vol_rank = next((i + 1 for i, x in enumerate(by_vol) if x["uid"] == qr["uid"]), len(by_vol))
    by_avg = sorted(feas, key=lambda x: x["averageTransactionValue"], reverse=True)
    qr_avg_rank = next((i + 1 for i, x in enumerate(by_avg) if x["uid"] == qr["uid"]), len(by_avg))

    leader_name = _short_name(volume_leader["name"]) if volume_leader else "other services"
    qr_name = qr["name"]

    if qr_avg_rank == 1 and qr_vol_rank > 1:
        return (qr_name + " remains smaller in volume compared to " + leader_name
                + ", but each transaction carries significantly more value \u2014 making it a "
                + "high-potential growth channel for digital merchant payments.")
    if qr_avg_rank == 1 and qr_vol_rank == 1:
        return (qr_name + " leads in both volume and average transaction value, "
                + "demonstrating strong adoption and high-value usage across the payment ecosystem.")
    if qr_avg_rank <= 2:
        return (qr_name + " shows competitive average transaction value. With growing merchant adoption, "
                + "it represents a key channel for high-value digital payments.")
    return (qr_name + " has room for growth in average transaction value. Continued merchant onboarding "
            + "and user education could drive higher-value transactions over time.")


def build_report_html(report, calc, show_bars=True, auto_highlight=True, takeaway_override=""):
    """Render the 1038x735 landscape report as a self-contained HTML string."""
    services = calc["services"]
    highest_avg = calc["highestAvg"]
    highest_avg_name = highest_avg.get("name") if highest_avg else None
    qr = calc["qr"]
    takeaway = takeaway_override.strip() or calc["takeaway"]

    rows = []
    for s in services:
        is_rate = s["isSuccessRate"]
        is_highest = (not is_rate) and highest_avg_name and highest_avg_name == s["name"]
        effective_h1 = s["highlighted"] or (auto_highlight and is_highest)
        tr_class = ' class="highlight-row"' if effective_h1 else ""

        icon_class = "financial" if (s["isFinancial"] and not is_rate) else "non-financial"
        svc = ('<td><div class="svc-cell"><div class="svc-icon ' + icon_class + ' ' + _icon_class(s["name"]) + '">'
               '<i data-lucide="' + _icon(s["name"]) + '"></i></div>'
               '<span class="svc-name">' + _esc(s["name"]) + "</span></div></td>")

        bar = lambda pct: (
            ('<div class="metric-bar-wrap"><div class="metric-bar">'
             '<div class="metric-bar-fill" style="width:' + f"{pct:.1f}" + '%"></div></div></div>')
            if show_bars else "")

        if is_rate:
            perf = ('<div class="num-primary">' + (fmt_dec(s["transactionVolume"]) + "%" if s["transactionVolume"] > 0
                                                   else '<span class="num-dash">-</span>') + "</div>")
        else:
            perf = ('<div class="num-primary">' + fmt_int(s["transactionVolume"]) + "</div>" + bar(s["volumePercent"]))
        if is_rate:
            target = ('<div class="num-primary">' + (fmt_dec(s["target"]) + "%" if s["target"] > 0
                                                    else '<span class="num-dash">-</span>') + "</div>")
        else:
            target = ('<div class="num-primary">' + (fmt_int(s["target"]) if s["target"] > 0
                                                    else '<span class="num-dash">-</span>') + "</div>")
        if s["achievementPercent"] is None:
            ach = '<span class="num-dash">-</span>'
        else:
            ach = f"{s['achievementPercent']:.2f}%"
        ach = '<div class="num-primary">' + ach + "</div>"
        tot = ('<div class="num-primary">' + ("ETB " + fmt_dec(s["totalValue"]) if s["isFinancial"]
                                              else '<span class="num-dash">-</span>') + "</div>")
        avg = ('<div class="num-primary">' + ("ETB " + fmt_dec(s["averageTransactionValue"]) if s["isFinancial"]
                                              else '<span class="num-dash">-</span>') + "</div>"
               + (bar(s["avgPercent"]) if s["isFinancial"] else ""))

        msg = "msg-highlight" if is_highest else ""
        badge = ""
        if is_highest:
            badge = ('<div class="msg-badge"><i data-lucide="star"></i> HIGHEST AVG VALUE</div>')
        key_msg = ('<td class="msg-cell ' + msg + '">' + _esc(s["keyMessage"]) + badge + "</td>")

        rows.append("<tr" + tr_class + ">" + svc
                    + '<td class="num-cell">' + target + "</td>"
                    + '<td class="num-cell">' + perf + "</td>"
                    + '<td class="num-cell">' + tot + "</td>"
                    + '<td class="num-cell">' + ach + "</td>"
                    + '<td class="num-cell">' + avg + "</td>"
                    + key_msg + "</tr>")

    tot = calc["total"]
    if tot["achievementPercent"] is None:
        tot_ach = '<span class="num-dash">-</span>'
    else:
        tot_ach = f"{tot['achievementPercent']:.2f}%" if not _num(tot['achievementPercent']).is_integer() else f"{int(tot['achievementPercent'])}%"
    total_row = ('<tr class="total-row">'
                 '<td><div class="svc-cell"><span class="svc-name">TOTAL</span></div></td>'
                 '<td class="num-cell"><div class="num-primary num-total">' + fmt_int(tot["target"]) + "</div></td>"
                 '<td class="num-cell"><div class="num-primary num-total">' + fmt_int(tot["performance"]) + "</div></td>"
                 '<td class="num-cell"><div class="num-primary num-total">ETB ' + fmt_dec(tot["totalValue"]) + "</div></td>"
                 '<td class="num-cell"><div class="num-primary num-total">' + tot_ach + "</div></td>"
                 '<td class="num-cell"><div class="num-dash">-</div></td>'
                 '<td class="msg-cell"></td></tr>')

    table = ('<div class="report-table-wrapper"><table class="report-table"><thead><tr>'
             '<th class="th-service"><span class="th-inner"><span class="th-icon th-icon-blue">'
             '<i data-lucide="list"></i></span>SERVICE</span></th>'
             '<th class="th-target"><span class="th-inner th-num"><span class="th-icon th-icon-purple">'
             '<i data-lucide="target"></i></span>MONTHLY PLAN (TARGET)</span></th>'
             '<th class="th-volume"><span class="th-inner th-num"><span class="th-icon th-icon-purple">'
             '<i data-lucide="hash"></i></span>PERFORMANCE</span></th>'
             '<th class="th-value"><span class="th-inner th-num"><span class="th-icon th-icon-teal">'
             '<i data-lucide="wallet"></i></span>TOTAL VALUE (ETB)</span></th>'
             '<th class="th-ach"><span class="th-inner th-num"><span class="th-icon th-icon-green">'
             '<i data-lucide="trending-up"></i></span>ACHIEVEMENT %</span></th>'
             '<th class="th-avg"><span class="th-inner th-num"><span class="th-icon th-icon-green">'
             '<i data-lucide="trending-up"></i></span>AVG TRANSACTION VALUE (ETB)</span></th>'
             '<th class="th-message"><span class="th-inner"><span class="th-icon th-icon-purple">'
             '<i data-lucide="message-square"></i></span>KEY MESSAGE</span></th>'
             "</tr></thead><tbody>" + "".join(rows) + total_row + "</tbody></table></div>")

    leader = calc["volumeLeader"]
    vl_name = _esc(leader["name"]) if leader else "-"
    vl_detail = _esc(fmt_int(leader["transactionVolume"]) + " transactions drive the ecosystem's scale.") if leader else "&mdash;"

    qr_avg = "-"
    comparisons = ""
    if qr and qr["isFinancial"]:
        qr_avg = "ETB " + fmt_dec(qr["averageTransactionValue"])
        comps = []
        for adv in calc["qrAdvantages"]:
            comps.append('<div class="qr-comp"><div class="qr-comp-multiplier">' + adv["label"]
                         + '</div><div class="qr-comp-label">' + _esc(adv["description"]) + "</div></div>")
        comparisons = '<div class="qr-comparisons">' + "".join(comps) + "</div>"

    insights = ('<div class="report-insights">'
                '<div class="insight-card volume-leader-card">'
                '<div class="insight-label">VOLUME LEADER</div>'
                '<div class="insight-service">' + vl_name + "</div>"
                '<div class="insight-detail">' + vl_detail + "</div></div>"
                '<div class="insight-card qr-advantage-card">'
                '<div class="insight-label">QR VALUE ADVANTAGE</div>'
                '<div class="qr-avg-main"><div class="qr-avg-label">AVERAGE TRANSACTION VALUE</div>'
                '<div class="qr-avg-value">' + qr_avg + '</div><div class="qr-avg-sub">per transaction</div></div>'
                + comparisons + "</div>"
                '<div class="insight-card takeaway-card">'
                '<div class="insight-label">KEY TAKEAWAY</div>'
                '<div class="takeaway-text">' + _esc(takeaway) + "</div></div>"
                "</div>")

    org_logo = _img_data_uri("ethswitch.jpg")
    bird_logo = _img_data_uri("ethiopay-bird.jpg")
    header = ('<div class="report-header">'
              '<div class="header-left"><div class="org-badge">'
              '<img class="org-logo" src="' + org_logo + '" alt="' + _esc(report.get("organization")) + '">'
              '<div><div class="org-name">' + _esc(report.get("organization")) + "</div>"
              '<div class="org-sub">Making Payments Simple and Affordable</div></div></div></div>'
              '<div class="header-center"><div class="brand-badge">' + _esc(report.get("brand")) + "</div>"
              '<h1 class="report-title">' + _esc(report.get("title")) + "</h1>"
              '<p class="report-subtitle">' + _esc(report.get("subtitle")) + "</p></div>"
              '<div class="header-right"><div class="header-right-inner">'
              '<div class="bird-group"><div class="bird-text">'
              '<div class="bird-name">Ethiopay</div>'
              '<div class="bird-tagline">One Payment, Every Possibility</div></div>'
              '<img class="bird-logo" src="' + bird_logo + '" alt="EthioPay"></div>'
              '<div class="header-sep"></div>'
              '<div class="date-badge"><span class="date-label">DATE</span>'
              '<span class="date-value">' + _esc(report.get("date")) + "</span></div>"
              + _phone_mockup() +
              "</div></div></div>")

    footer = ('<div class="report-footer"><div>'
              '<span class="footer-org-name">' + _esc(report.get("organization")) + " S.C.</span>"
              '<span class="footer-tagline">' + _esc(report.get("tagline1")) + "</span></div>"
              '<div class="footer-web"><i data-lucide="globe"></i><span>www.ethswitch.et</span></div></div>')

    return ("<!DOCTYPE html><html><head><meta charset='utf-8'/><style>" + REPORT_CSS + "</style>"
            "<script src='https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js'></script>"
            "<script src='https://unpkg.com/lucide@latest'></script></head><body>"
            "<div class='toolbar'>"
            "<button onclick='window.print()'>Print / Save as PDF</button>"
            "<button onclick='capturePNG()'>Download PNG</button></div>"
            "<div id='report-content'>" + header + table + insights + footer + "</div>"
            "<script>lucide.createIcons();"
            "function capturePNG(){var t=document.querySelector('.toolbar');var e=document.getElementById('report-content');"
            "t.style.visibility='hidden';"
            "html2canvas(e,{scale:2,useCORS:true,backgroundColor:'#ffffff',windowWidth:1038,windowHeight:735})"
            ".then(function(c){t.style.visibility='visible';var a=document.createElement('a');"
            "a.download='Digital_Transaction_Value_Snapshot.png';a.href=c.toDataURL('image/png');a.click();})"
            ".catch(function(err){t.style.visibility='visible';"
            "alert('PNG export needs internet (html2canvas loads from CDN): '+err.message);});}"
            "</script></body></html>")