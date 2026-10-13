"""Visual design system for the Streamlit UI.

`streamlit_app` owns behaviour; this module owns everything you can see: one
token table per theme, a small line-icon set and a single stylesheet that every
component reads from, so the light and dark themes can never drift apart.

Two things are deliberate here:

* Colours are declared once as custom properties. Rules use ``var(--...)`` only,
  which is what makes a theme switch a single-line change.
* Motion is progressive. Time-based keyframes are the baseline; scroll-driven
  animations (``animation-timeline``) are added behind ``@supports`` so a
  browser without them simply shows the page without the effect.
"""

from __future__ import annotations

# ── Tokens ──────────────────────────────────────────────────────────────────
# `scheme` keeps native widgets (scrollbars, date pickers, autofill) in step
# with the theme; the rest are consumed by `build_css`.

TOKENS: dict[str, dict[str, str]] = {
    "dark": {
        "scheme": "dark",
        "bg": "#070B14",
        "bg2": "#0B1220",
        "panel": "#0E1626",
        "surface": "rgba(255,255,255,0.030)",
        "surface2": "rgba(255,255,255,0.062)",
        "border": "rgba(255,255,255,0.075)",
        "border2": "rgba(255,255,255,0.170)",
        "text": "#ECF1F9",
        "dim": "#97A2BB",
        "faint": "#7A86A1",
        "accent": "#F2A93B",
        "accent-soft": "rgba(242,169,59,0.13)",
        "accent-line": "rgba(242,169,59,0.34)",
        "accent-ink": "#140D02",
        "accent2": "#8C9BFF",
        "accent2-soft": "rgba(140,155,255,0.14)",
        "ok": "#34D399",
        "ok-soft": "rgba(52,211,153,0.12)",
        "ok-ink": "#03251B",
        "warn": "#FBBF24",
        "warn-soft": "rgba(251,191,36,0.12)",
        "bad": "#FB7185",
        "bad-soft": "rgba(251,113,133,0.12)",
        "info": "#60A5FA",
        "info-soft": "rgba(96,165,250,0.12)",
        "violet": "#C4B5FD",
        "violet-soft": "rgba(196,181,253,0.12)",
        "glow1": "rgba(242,169,59,0.17)",
        "glow2": "rgba(140,155,255,0.15)",
        "glow3": "rgba(52,211,153,0.10)",
        "grid": "rgba(255,255,255,0.028)",
        "shadow1": "0 1px 2px rgba(0,0,0,0.45), 0 10px 28px -14px rgba(0,0,0,0.70)",
        "shadow2": "0 2px 6px rgba(0,0,0,0.50), 0 28px 56px -24px rgba(0,0,0,0.80)",
    },
    "light": {
        "scheme": "light",
        "bg": "#F4F6F9",
        "bg2": "#EAEFF6",
        "panel": "#FFFFFF",
        "surface": "rgba(255,255,255,0.80)",
        "surface2": "rgba(13,23,38,0.040)",
        "border": "rgba(12,22,40,0.105)",
        "border2": "rgba(12,22,40,0.200)",
        "text": "#0D1726",
        "dim": "#4A5670",
        "faint": "#5D6A82",
        "accent": "#B45309",
        "accent-soft": "rgba(180,83,9,0.085)",
        "accent-line": "rgba(180,83,9,0.30)",
        "accent-ink": "#FFFFFF",
        "accent2": "#4F46E5",
        "accent2-soft": "rgba(79,70,229,0.085)",
        "ok": "#047857",
        "ok-soft": "rgba(4,120,87,0.085)",
        "ok-ink": "#FFFFFF",
        "warn": "#B45309",
        "warn-soft": "rgba(180,83,9,0.090)",
        "bad": "#BE123C",
        "bad-soft": "rgba(190,18,60,0.085)",
        "info": "#1D4ED8",
        "info-soft": "rgba(29,78,216,0.085)",
        "violet": "#6D28D9",
        "violet-soft": "rgba(109,40,217,0.085)",
        "glow1": "rgba(244,169,37,0.24)",
        "glow2": "rgba(99,102,241,0.17)",
        "glow3": "rgba(16,185,129,0.13)",
        "grid": "rgba(12,22,40,0.040)",
        "shadow1": "0 1px 2px rgba(13,23,38,0.06), 0 10px 26px -16px rgba(13,23,38,0.28)",
        "shadow2": "0 2px 6px rgba(13,23,38,0.07), 0 26px 52px -26px rgba(13,23,38,0.35)",
    },
}

# ── Icons ───────────────────────────────────────────────────────────────────
# Lucide-flavoured 24x24 strokes. `icon()` wraps a name in an <svg> that
# inherits `currentColor`, so an icon always matches the text beside it.

ICONS: dict[str, str] = {
    "trending-down": '<polyline points="22 17 13.5 8.5 8.5 13.5 2 7"/><polyline points="16 17 22 17 22 11"/>',
    "trending-up": '<polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/>',
    "check": '<path d="M21.8 10A10 10 0 1 1 17 3.34"/><path d="m9 11 3 3L22 4"/>',
    "credit-card": '<rect width="20" height="14" x="2" y="5" rx="2"/><line x1="2" x2="22" y1="10" y2="10"/>',
    "landmark": '<line x1="3" x2="21" y1="22" y2="22"/><line x1="6" x2="6" y1="18" y2="11"/><line x1="10" x2="10" y1="18" y2="11"/><line x1="14" x2="14" y1="18" y2="11"/><line x1="18" x2="18" y1="18" y2="11"/><polygon points="12 2 20 7 4 7"/>',
    "contactless": '<path d="M6 8.32a7.43 7.43 0 0 1 0 7.36"/><path d="M9.46 6.21a11.76 11.76 0 0 1 0 11.58"/><path d="M12.91 4.1a15.91 15.91 0 0 1 .01 15.79"/><path d="M16.37 2a20.16 20.16 0 0 1 0 20"/>',
    "link": '<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>',
    "swap": '<path d="M8 3 4 7l4 4"/><path d="M4 7h16"/><path d="m16 21 4-4-4-4"/><path d="M20 17H4"/>',
    "repeat": '<path d="m17 2 4 4-4 4"/><path d="M3 11v-1a4 4 0 0 1 4-4h14"/><path d="m7 22-4-4 4-4"/><path d="M21 13v1a4 4 0 0 1-4 4H3"/>',
    "calculator": '<rect x="4" y="2" width="16" height="20" rx="2"/><line x1="8" x2="16" y1="6" y2="6"/><line x1="8" x2="8" y1="14" y2="14"/><line x1="12" x2="12" y1="14" y2="14"/><line x1="16" x2="16" y1="14" y2="14"/><line x1="8" x2="8" y1="18" y2="18"/><line x1="12" x2="12" y1="18" y2="18"/><line x1="16" x2="16" y1="18" y2="18"/>',
    "calendar": '<path d="M8 2v4"/><path d="M16 2v4"/><rect width="18" height="18" x="3" y="4" rx="2"/><path d="M3 10h18"/>',
    "layers": '<path d="M12.83 2.18a2 2 0 0 0-1.66 0L2.6 6.08a1 1 0 0 0 0 1.83l8.58 3.91a2 2 0 0 0 1.66 0l8.58-3.9a1 1 0 0 0 0-1.83Z"/><path d="m6.08 9.5-3.5 1.6a1 1 0 0 0 0 1.81l8.6 3.91a2 2 0 0 0 1.65 0l8.58-3.9a1 1 0 0 0 0-1.83l-3.5-1.6"/><path d="m6.08 14.5-3.5 1.6a1 1 0 0 0 0 1.81l8.6 3.91a2 2 0 0 0 1.65 0l8.58-3.9a1 1 0 0 0 0-1.83l-3.5-1.6"/>',
    "file": '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/>',
    "eye": '<path d="M2.06 12.35a1 1 0 0 1 0-.7 10.75 10.75 0 0 1 19.88 0 1 1 0 0 1 0 .7 10.75 10.75 0 0 1-19.88 0"/><circle cx="12" cy="12" r="3"/>',
    "search": '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>',
    "alert": '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
    "download": '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/>',
    "upload": '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" x2="12" y1="3" y2="15"/>',
    "chart": '<line x1="12" x2="12" y1="20" y2="10"/><line x1="18" x2="18" y1="20" y2="4"/><line x1="6" x2="6" y1="20" y2="16"/>',
    "chart-line": '<path d="M3 3v16a2 2 0 0 0 2 2h16"/><path d="m7 15 4-5 3 3 5-7"/>',
    "building": '<path d="M6 22V4a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v18Z"/><path d="M6 12H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h2"/><path d="M18 9h2a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2h-2"/><path d="M10 6h4"/><path d="M10 10h4"/><path d="M10 14h4"/><path d="M10 18h4"/>',
    "wallet": '<path d="M19 7V4a1 1 0 0 0-1-1H5a2 2 0 0 0 0 4h15a1 1 0 0 1 1 1v4h-3a2 2 0 0 0 0 4h3a1 1 0 0 0 1-1v-2a1 1 0 0 0-1-1"/><path d="M3 5v14a2 2 0 0 0 2 2h15a1 1 0 0 0 1-1v-4"/>',
    "sliders": '<line x1="21" x2="14" y1="4" y2="4"/><line x1="10" x2="3" y1="4" y2="4"/><line x1="21" x2="12" y1="12" y2="12"/><line x1="8" x2="3" y1="12" y2="12"/><line x1="21" x2="16" y1="20" y2="20"/><line x1="12" x2="3" y1="20" y2="20"/><line x1="14" x2="14" y1="2" y2="6"/><line x1="8" x2="8" y1="10" y2="14"/><line x1="16" x2="16" y1="18" y2="22"/>',
    "folder": '<path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/>',
    "merge": '<path d="M8 3H5a2 2 0 0 0-2 2v3"/><path d="M21 8V5a2 2 0 0 0-2-2h-3"/><path d="M3 16v3a2 2 0 0 0 2 2h3"/><path d="M16 21h3a2 2 0 0 0 2-2v-3"/><line x1="12" x2="12" y1="3" y2="21"/>',
    "shield": '<path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1Z"/>',
    "moon": '<path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/>',
    "sparkles": '<path d="M9.94 14.06 2 21"/><path d="M11.34 2.5 12.5 6.5l4 1.1-4 1.1-1.16 4-1.16-4-4-1.1 4-1.1Z"/><path d="M18.5 13.5 19.4 17l3.5 1-3.5 1-.9 3.5-.9-3.5-3.5-1 3.5-1Z"/>',
    "arrow-right": '<path d="M5 12h14"/><path d="m12 5 7 7-7 7"/>',
    "plus": '<path d="M5 12h14"/><path d="M12 5v14"/>',
    "info": '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><path d="M12 8h.01"/>',
}


def icon(name: str, size: int = 18, cls: str = "") -> str:
    """An inline SVG for `name`, inheriting the surrounding text colour."""
    body = ICONS.get(name, ICONS["file"])
    cls_attr = f' class="{cls}"' if cls else ""
    return (
        f'<svg{cls_attr} width="{size}" height="{size}" viewBox="0 0 24 24" '
        'fill="none" stroke="currentColor" stroke-width="1.75" '
        'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" '
        'focusable="false">' + body + "</svg>"
    )


# ── Stylesheet ──────────────────────────────────────────────────────────────

_FONT_HREF = (
    "https://fonts.googleapis.com/css2?"
    "family=IBM+Plex+Sans:wght@300;400;500;600;700&"
    "family=IBM+Plex+Mono:wght@400;500;600&display=swap"
)

_STYLESHEET = """
@import url('@@FONT@@');

/* ── Tokens ─────────────────────────────────────────────────────────── */
:root {
  --font: 'IBM Plex Sans', ui-sans-serif, system-ui, -apple-system, 'Segoe UI', sans-serif;
  --mono: 'IBM Plex Mono', ui-monospace, 'SFMono-Regular', Menlo, monospace;

  --bg: @@bg@@; --bg2: @@bg2@@; --panel: @@panel@@;
  --surface: @@surface@@; --surface2: @@surface2@@;
  --border: @@border@@; --border2: @@border2@@;
  --text: @@text@@; --dim: @@dim@@; --faint: @@faint@@;
  --accent: @@accent@@; --accent-soft: @@accent-soft@@;
  --accent-line: @@accent-line@@; --accent-ink: @@accent-ink@@;
  --accent2: @@accent2@@; --accent2-soft: @@accent2-soft@@;
  --ok: @@ok@@; --ok-soft: @@ok-soft@@; --ok-ink: @@ok-ink@@;
  --warn: @@warn@@; --warn-soft: @@warn-soft@@;
  --bad: @@bad@@; --bad-soft: @@bad-soft@@;
  --info: @@info@@; --info-soft: @@info-soft@@;
  --violet: @@violet@@; --violet-soft: @@violet-soft@@;
  --glow1: @@glow1@@; --glow2: @@glow2@@; --glow3: @@glow3@@;
  --grid: @@grid@@; --shadow1: @@shadow1@@; --shadow2: @@shadow2@@;

  --r-xs: 8px; --r-sm: 11px; --r: 15px; --r-lg: 20px; --r-xl: 28px;
  --ease: cubic-bezier(.22, 1, .36, 1);
  --ease-out: cubic-bezier(.16, 1, .3, 1);
  --t-fast: .16s; --t: .24s; --t-slow: .52s;

  color-scheme: @@scheme@@;
}

/* ── Base ────────────────────────────────────────────────────────────── */
* { -webkit-font-smoothing: antialiased; text-rendering: optimizeLegibility; }

.stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"],
[data-testid="stHeader"], [data-testid="stBottom"] { background: transparent !important; }

.stApp {
  font-family: var(--font);
  color: var(--text);
  font-size: 15px;
  line-height: 1.55;
}

/* Strip Streamlit's own chrome; the app supplies its own header. */
#MainMenu, footer, header[data-testid="stHeader"],
div[data-testid="stToolbar"], div[data-testid="stDecoration"],
div[data-testid="stStatusWidget"], div[data-testid="stDeployButton"],
button[title="View fullscreen"], div.stProfiler,
[data-testid="stMainBlockContainer"] > div[data-testid="stVerticalBlock"]:first-child > div:empty {
  visibility: hidden !important; height: 0 !important;
  margin: 0 !important; padding: 0 !important; display: none !important;
}

section[data-testid="stSidebar"] { background: var(--panel) !important; }

.block-container, [data-testid="stMainBlockContainer"] {
  max-width: 1180px;
  padding: 1.5rem 2.2rem 5rem !important;
  position: relative;
  z-index: 2;
}
.block-container [data-testid="stVerticalBlock"] { gap: 0.5rem; }
.block-container [data-testid="stHorizontalBlock"] { gap: 0.7rem; }

::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb {
  background: var(--border2); border-radius: 99px;
  border: 3px solid transparent; background-clip: content-box; }
::-webkit-scrollbar-thumb:hover { background: var(--faint); background-clip: content-box; }

::selection { background: var(--accent-soft); color: var(--text); }

:focus-visible {
  outline: 2px solid var(--accent);
  outline-offset: 2px;
  border-radius: 6px;
}

/* ── Canvas: aurora wash + Swiss grid ───────────────────────────────── */
.stApp::before {
  content: ''; position: fixed; inset: -22%; z-index: 0; pointer-events: none;
  background:
    radial-gradient(46rem 30rem at 10% -6%, var(--glow1), transparent 62%),
    radial-gradient(40rem 28rem at 94% 2%,  var(--glow2), transparent 60%),
    radial-gradient(48rem 34rem at 58% 106%, var(--glow3), transparent 62%);
  filter: blur(26px);
  animation: drift 34s var(--ease) infinite alternate;
}
.stApp::after {
  content: ''; position: fixed; inset: 0; z-index: 0; pointer-events: none;
  background-image:
    linear-gradient(var(--grid) 1px, transparent 1px),
    linear-gradient(90deg, var(--grid) 1px, transparent 1px);
  background-size: 76px 76px;
  -webkit-mask-image: radial-gradient(78% 52% at 50% 0%, #000 0%, transparent 76%);
          mask-image: radial-gradient(78% 52% at 50% 0%, #000 0%, transparent 76%);
}

@keyframes drift {
  from { transform: translate3d(-2%, -1%, 0) scale(1); }
  to   { transform: translate3d(2%, 2%, 0) scale(1.08); }
}

/* ── Scroll rail ─────────────────────────────────────────────────────── */
.scroll-rail {
  position: sticky; top: 0; z-index: 40;
  height: 2px; margin: 0 0 1.15rem;
  background: var(--surface2); border-radius: 99px; overflow: hidden;
}
.scroll-fill {
  display: block; height: 100%; width: 100%;
  transform: scaleX(0); transform-origin: 0 50%;
  background: linear-gradient(90deg, var(--accent), var(--accent2));
}
@supports (animation-timeline: scroll(nearest block)) {
  .scroll-fill { animation: railGrow linear both; animation-timeline: scroll(nearest block); }
}
@keyframes railGrow { from { transform: scaleX(0); } to { transform: scaleX(1); } }

/* ── Typography ──────────────────────────────────────────────────────── */
.eyebrow {
  display: inline-flex; align-items: center; gap: .5rem;
  margin: 0 0 .55rem; font-size: .68rem; font-weight: 600;
  letter-spacing: .18em; text-transform: uppercase; color: var(--accent);
}
.eyebrow::before {
  content: ''; width: 18px; height: 1px; background: var(--accent); opacity: .6; }

.hero-title {
  margin: 0; font-size: clamp(2rem, 1.2rem + 2.4vw, 2.9rem);
  font-weight: 700; letter-spacing: -.032em; line-height: 1.04; color: var(--text);
}
.hero-title em {
  font-style: normal;
  background: linear-gradient(100deg, var(--accent), var(--accent2) 78%);
  -webkit-background-clip: text; background-clip: text; color: transparent;
}
.subtitle {
  margin: .7rem 0 0; max-width: 62ch;
  font-size: .95rem; line-height: 1.6; color: var(--dim);
}

.hr-rule { height: 1px; background: var(--border); border: 0; margin: 1.6rem 0; }

h1, h2, h3, h4 { color: var(--text); letter-spacing: -.018em; font-weight: 650; }
h1 { font-size: 1.6rem; } h2 { font-size: 1.28rem; } h3 { font-size: 1.08rem; }
a { color: var(--accent); text-decoration-color: var(--accent-line); }
a:hover { text-decoration-color: var(--accent); }
code {
  font-family: var(--mono); font-size: .82em;
  background: var(--surface2); color: var(--text);
  padding: .1em .38em; border-radius: 5px; border: 1px solid var(--border);
}
.stCaption, [data-testid="stCaptionContainer"] p,
caption, .stMarkdown p em + p { color: var(--dim) !important; }
.caption { color: var(--faint); font-size: .8rem; }

/* ── Section header ("card") ─────────────────────────────────────────── */
.card {
  margin: 1.7rem 0 0; padding: 0;
  border: 0; background: none; animation: none;
}
.card-head {
  display: flex; align-items: center; gap: .9rem;
  padding: 1rem 1.15rem;
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--r-lg); box-shadow: var(--shadow1);
  backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px);
}
.card-title {
  margin: 0; font-size: 1rem; font-weight: 650; color: var(--text);
  letter-spacing: -.012em; line-height: 1.3;
}
.card-sub {
  margin: .18rem 0 0; font-size: .82rem; line-height: 1.5; color: var(--dim);
}
.card-icon {
  flex: 0 0 auto; display: inline-flex; align-items: center; justify-content: center;
  width: 38px; height: 38px; border-radius: var(--r-sm);
  font-size: .8rem; font-weight: 700; letter-spacing: .02em;
  background: var(--surface2); border: 1px solid var(--border);
  color: var(--dim);
}
.icon-accent  { background: var(--accent-soft);  border-color: var(--accent-line); color: var(--accent); }
.icon-ok      { background: var(--ok-soft);      border-color: var(--ok); color: var(--ok); }
.icon-warn    { background: var(--warn-soft);    border-color: var(--warn); color: var(--warn); }
.icon-bad     { background: var(--bad-soft);     border-color: var(--bad); color: var(--bad); }
.icon-info    { background: var(--info-soft);    border-color: var(--info); color: var(--info); }
.icon-violet  { background: var(--violet-soft);  border-color: var(--violet); color: var(--violet); }
.icon-accent2 { background: var(--accent2-soft); border-color: var(--accent2); color: var(--accent2); }

/* ── Section separators ──────────────────────────────────────────────── */
.section-sep {
  display: flex; align-items: center; gap: 1rem;
  margin: 2.4rem 0 .2rem;
}
.section-sep::after {
  content: ''; flex: 1; height: 1px;
  background: linear-gradient(90deg, var(--border2), transparent);
}
.section-sep span {
  font-size: .68rem; font-weight: 650; letter-spacing: .17em;
  text-transform: uppercase; color: var(--faint); white-space: nowrap;
}

/* ── Report type picker ──────────────────────────────────────────────── */
.mode-grid {
  display: grid; gap: .8rem; margin: 1rem 0 .1rem;
  grid-template-columns: repeat(4, minmax(0, 1fr));
}
.mode-card {
  position: relative; display: flex; flex-direction: column; gap: .1rem;
  padding: 1.05rem 1rem 1rem;
  border: 1px solid var(--border); border-radius: var(--r-lg);
  background: var(--surface); box-shadow: var(--shadow1);
  transition: transform var(--t) var(--ease), border-color var(--t) var(--ease),
              box-shadow var(--t) var(--ease), background var(--t) var(--ease);
}
.mode-card::before {
  content: ''; position: absolute; inset: 0; border-radius: inherit;
  background: linear-gradient(180deg, var(--mc-glow, transparent), transparent 62%);
  opacity: .55; pointer-events: none;
}
.mode-card:hover {
  transform: translateY(-3px);
  border-color: var(--border2);
  box-shadow: var(--shadow2);
}
.mode-card.mode-active {
  border-color: var(--accent-line);
  box-shadow: var(--shadow1), 0 0 0 1px var(--accent-line);
  background: var(--accent-soft);
}
.mc-icon {
  display: inline-flex; align-items: center; justify-content: center;
  width: 34px; height: 34px; margin-bottom: .7rem;
  border-radius: var(--r-sm);
  background: var(--mc-tint, var(--surface2));
  border: 1px solid var(--mc-edge, var(--border));
  color: var(--mc, var(--dim));
}
.mc-name {
  font-size: .93rem; font-weight: 650; color: var(--text);
  letter-spacing: -.01em; line-height: 1.25;
}
.mc-desc { margin-top: .2rem; font-size: .755rem; line-height: 1.4; color: var(--faint); }
.mode-card.mode-active .mc-name { color: var(--accent); }
.mc-check {
  position: absolute; top: .85rem; right: .85rem; display: none;
  color: var(--accent);
}
.mode-card.mode-active .mc-check { display: flex; }

.t-red    { --mc: var(--bad);     --mc-tint: var(--bad-soft);     --mc-edge: var(--bad);     --mc-glow: var(--bad-soft); }
.t-green  { --mc: var(--ok);      --mc-tint: var(--ok-soft);      --mc-edge: var(--ok);      --mc-glow: var(--ok-soft); }
.t-blue   { --mc: var(--info);    --mc-tint: var(--info-soft);    --mc-edge: var(--info);    --mc-glow: var(--info-soft); }
.t-cyan   { --mc: var(--accent2); --mc-tint: var(--accent2-soft); --mc-edge: var(--accent2); --mc-glow: var(--accent2-soft); }
.t-amber  { --mc: var(--warn);    --mc-tint: var(--warn-soft);    --mc-edge: var(--warn);    --mc-glow: var(--warn-soft); }
.t-purple,
.t-violet { --mc: var(--violet);  --mc-tint: var(--violet-soft);  --mc-edge: var(--violet);  --mc-glow: var(--violet-soft); }

/* ── Metrics ─────────────────────────────────────────────────────────── */
div[data-testid="stMetric"] {
  position: relative; overflow: hidden;
  padding: .85rem .95rem .9rem !important;
  background: var(--surface) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--r) !important;
  box-shadow: var(--shadow1) !important;
  transition: transform var(--t) var(--ease), border-color var(--t) var(--ease);
}
div[data-testid="stMetric"]::after {
  content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 2px;
  background: var(--accent); opacity: 0; transition: opacity var(--t) var(--ease);
}
div[data-testid="stMetric"]:hover { transform: translateY(-2px); border-color: var(--border2) !important; }
div[data-testid="stMetric"]:hover::after { opacity: .8; }
div[data-testid="stMetric"] label,
div[data-testid="stMetric"] [data-testid="stMetricLabel"] {
  color: var(--faint) !important; font-size: .655rem !important; font-weight: 650 !important;
  letter-spacing: .13em !important; text-transform: uppercase;
}
div[data-testid="stMetric"] [data-testid="stMetricValue"],
div[data-testid="stMetricValue"] {
  color: var(--text) !important;
  font-family: var(--mono) !important; font-size: 1.32rem !important;
  font-weight: 600 !important; letter-spacing: -.02em !important;
}
div[data-testid="stMetric"] [data-testid="stMetricDelta"] { color: var(--dim) !important; }

/* ── File rows ───────────────────────────────────────────────────────── */
.sheet-item {
  display: flex; align-items: center; gap: .7rem;
  padding: .58rem .85rem; margin-bottom: .38rem;
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--r-sm);
  transition: border-color var(--t) var(--ease), transform var(--t) var(--ease);
}
.sheet-item:hover { border-color: var(--border2); transform: translateX(2px); }
.sheet-num {
  flex: 0 0 auto; display: inline-flex; align-items: center; justify-content: center;
  width: 22px; height: 22px; border-radius: 6px;
  background: var(--accent-soft); color: var(--accent);
  font-family: var(--mono); font-size: .66rem; font-weight: 600;
}
.sheet-name {
  min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  font-family: var(--mono); font-size: .78rem; color: var(--text);
}
.sheet-desc { font-size: .8rem; color: var(--dim); }
.sheet-rows {
  margin-left: auto; padding-left: .6rem; white-space: nowrap;
  font-family: var(--mono); font-size: .72rem; color: var(--faint);
}

/* ── Badges & pills ──────────────────────────────────────────────────── */
.badge {
  display: inline-flex; align-items: center; gap: .35rem;
  padding: .22rem .6rem; border-radius: 99px;
  font-size: .655rem; font-weight: 650; letter-spacing: .12em;
  text-transform: uppercase; border: 1px solid transparent;
}
.badge-green  { background: var(--ok-soft);     color: var(--ok);     border-color: var(--ok); }
.badge-purple { background: var(--violet-soft); color: var(--violet); border-color: var(--violet); }
.badge-red    { background: var(--bad-soft);    color: var(--bad);    border-color: var(--bad); }
.badge-blue   { background: var(--info-soft);   color: var(--info);   border-color: var(--info); }
.badge-amber  { background: var(--warn-soft);   color: var(--warn);   border-color: var(--warn); }
.badge-accent { background: var(--accent-soft); color: var(--accent); border-color: var(--accent-line); }

.date-pill {
  display: inline-flex; align-items: center;
  padding: .24rem .6rem; margin: 0 .4rem .4rem 0;
  border-radius: 99px; font-family: var(--mono); font-size: .72rem;
  background: var(--accent-soft); color: var(--accent);
  border: 1px solid var(--accent-line);
}
.date-pill.is-empty { background: var(--bad-soft); color: var(--bad); border-color: var(--bad); }

/* ── Buttons ─────────────────────────────────────────────────────────── */
.stButton > button, .stDownloadButton > button,
button[data-testid^="stBaseButton"] {
  min-height: 44px;
  border-radius: var(--r-sm) !important;
  font-family: var(--font) !important;
  font-size: .875rem !important; font-weight: 600 !important;
  letter-spacing: -.005em !important;
  transition: transform var(--t-fast) var(--ease), box-shadow var(--t) var(--ease),
              background var(--t) var(--ease), border-color var(--t) var(--ease),
              color var(--t) var(--ease) !important;
}
.stButton > button:hover:not(:disabled),
.stDownloadButton > button:hover:not(:disabled),
button[data-testid^="stBaseButton"]:hover:not(:disabled) {
  transform: translateY(-1px) !important;
}
.stButton > button:active, .stDownloadButton > button:active,
button[data-testid^="stBaseButton"]:active:not(:disabled) { transform: translateY(0) !important; }
.stButton > button:disabled, .stDownloadButton > button:disabled {
  opacity: .45 !important; box-shadow: none !important; transform: none !important;
}

/* Primary: solid accent, one soft ring — no multi-hue gradients. */
.stButton > button[kind="primary"],
button[data-testid="stBaseButton-primary"],
.stButton > button[kind="primaryFormSubmit"] {
  background: var(--accent) !important;
  color: var(--accent-ink) !important;
  border: 1px solid transparent !important;
  box-shadow: 0 1px 2px rgba(0,0,0,.18), 0 10px 26px -14px var(--accent) !important;
}
.stButton > button[kind="primary"]:hover,
button[data-testid="stBaseButton-primary"]:hover {
  box-shadow: 0 2px 4px rgba(0,0,0,.2), 0 16px 34px -14px var(--accent) !important;
}
.stDownloadButton > button {
  background: var(--ok) !important;
  color: var(--ok-ink) !important;
  border: 1px solid transparent !important;
  box-shadow: 0 1px 2px rgba(0,0,0,.18), 0 10px 26px -14px var(--ok) !important;
}
.stDownloadButton > button:hover { box-shadow: 0 16px 34px -14px var(--ok) !important; }

/* Secondary / ghost: a quiet outline that warms on hover. */
.stButton > button:not([kind="primary"]),
button[data-testid="stBaseButton-secondary"],
button[data-testid="stBaseButton-tertiary"] {
  background: var(--surface) !important;
  color: var(--text) !important;
  border: 1px solid var(--border) !important;
  box-shadow: none !important;
}
.stButton > button:not([kind="primary"]):hover,
button[data-testid="stBaseButton-secondary"]:hover,
button[data-testid="stBaseButton-tertiary"]:hover {
  border-color: var(--accent-line) !important;
  color: var(--accent) !important;
  background: var(--accent-soft) !important;
}

/* ── Segmented navigation ────────────────────────────────────────────── */
[data-testid="stButtonGroup"][data-variant="segmented_control"] {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--r); padding: .28rem; box-shadow: var(--shadow1);
  gap: .25rem;
}
[data-testid="stButtonGroup"][data-variant="segmented_control"]
  [role="radiogroup"] { background: transparent !important; gap: .25rem; }
button[data-testid="stBaseButton-segmented_controlActive"],
button[data-testid="stBaseButton-segmented_controlActive"]:hover {
  background: var(--accent) !important; color: var(--accent-ink) !important;
  border-radius: var(--r-sm) !important; box-shadow: none !important;
  font-weight: 650 !important; transform: none !important;
}
button[data-testid="stBaseButton-segmented_control"],
button[data-testid="stBaseButton-segmented_control"]:hover {
  background: transparent !important; color: var(--dim) !important;
  border-radius: var(--r-sm) !important; box-shadow: none !important;
  transform: none !important;
}
button[data-testid="stBaseButton-segmented_control"]:hover {
  color: var(--text) !important; background: var(--surface2) !important;
}

/* ── Inputs ──────────────────────────────────────────────────────────── */
[data-testid="stWidgetLabel"] p, .stTextInput label p,
.stNumberInput label p, .stDateInput label p, .stTextArea label p,
.stSelectbox label p, .stMultiselect label p {
  color: var(--dim) !important;
  font-size: .7rem !important; font-weight: 650 !important;
  letter-spacing: .1em !important; text-transform: uppercase;
}
/* Checkbox, radio and toggle labels read as sentences, not field captions. */
.stCheckbox label p, .stRadio label p, [data-testid="stToggle"] label p,
[data-testid="stCheckbox"] label p, [data-testid="stRadio"] label p {
  color: var(--text) !important;
  font-size: .85rem !important; font-weight: 500 !important;
  letter-spacing: 0 !important; text-transform: none !important;
}
div[data-baseweb="input"], div[data-baseweb="textarea"] > textarea {
  background: var(--surface) !important;
  border-color: var(--border) !important;
  border-radius: var(--r-sm) !important;
  color: var(--text) !important;
  font-family: var(--font) !important;
}
div[data-baseweb="input"]:hover { border-color: var(--border2) !important; }
div[data-baseweb="input"]:focus-within {
  border-color: var(--accent) !important;
  box-shadow: 0 0 0 3px var(--accent-soft) !important;
}
div[data-baseweb="input"] input::placeholder,
div[data-baseweb="textarea"] textarea::placeholder { color: var(--faint) !important; }

div[data-baseweb="select"] > div, div[data-baseweb="popover"] > div {
  background: var(--panel) !important; color: var(--text) !important;
  border-color: var(--border) !important; border-radius: var(--r-sm) !important;
}
div[data-baseweb="select"] > div { background: var(--surface) !important; }
div[data-baseweb="select"] > div:hover { border-color: var(--border2) !important; }
div[data-baseweb="select"] > div:focus-within {
  border-color: var(--accent) !important;
  box-shadow: 0 0 0 3px var(--accent-soft) !important;
}
div[data-baseweb="select"] svg { fill: var(--dim) !important; }
div[role="listbox"] { background: var(--panel) !important; }
div[role="option"]:hover { background: var(--surface2) !important; color: var(--text) !important; }
div[role="option"][aria-selected="true"] { color: var(--accent) !important; font-weight: 600 !important; }

div[data-baseweb="tag"] {
  background: var(--accent-soft) !important; color: var(--accent) !important;
  border: 1px solid var(--accent-line) !important; border-radius: 7px !important;
  font-family: var(--mono) !important; font-size: .74rem !important;
}
div[data-baseweb="tag"] svg { fill: var(--accent) !important; }
div[data-baseweb="tag"] [data-testid="stTagCloseIcon"] svg { fill: var(--accent) !important; }
ul[data-testid="stMultiselectDropdown"], ul[role="listbox"] { background: var(--panel) !important; }

div[data-baseweb="radio"] > label,
div[data-testid="stRadio"] > div[role="radiogroup"] > label {
  cursor: pointer; color: var(--dim) !important; border-radius: var(--r-sm);
  transition: background var(--t) var(--ease), color var(--t) var(--ease);
  padding: .45rem .6rem;
}
div[data-baseweb="radio"] > label:hover { background: var(--surface2) !important; color: var(--text) !important; }
div[data-baseweb="radio"] > label[data-checked="true"] > div { color: var(--accent) !important; font-weight: 600 !important; }
div[data-baseweb="radio"] > label[data-checked="true"] { background: var(--accent-soft) !important; }

div[data-testid="stToggle"] > label { cursor: pointer; }
div[data-testid="stToggle"] > div[role="switch"] { background: var(--border2) !important; }
div[data-testid="stToggle"] > div[role="switch"][aria-checked="true"] { background: var(--accent) !important; }
div[data-testid="stCheckbox"] svg, div[data-testid="stCheckbox"] path { fill: var(--dim) !important; }

section[data-testid="stFileUploadDropzone"] {
  background: var(--surface) !important;
  border: 1.5px dashed var(--border2) !important;
  border-radius: var(--r-lg) !important;
  padding: 1.6rem 1.2rem !important;
  transition: border-color var(--t) var(--ease), background var(--t) var(--ease);
}
section[data-testid="stFileUploadDropzone"]:hover {
  border-color: var(--accent) !important; background: var(--accent-soft) !important;
}
section[data-testid="stFileUploadDropzone"] button {
  background: var(--accent) !important; color: var(--accent-ink) !important;
  border-radius: var(--r-sm) !important; font-weight: 600 !important;
  box-shadow: none !important;
}
div[data-testid="stFileUploaderDropzoneInstructions"] div,
div[data-testid="stFileUploaderDropzoneInstructions"] span { color: var(--dim) !important; }
div[data-testid="stFileUploaderDropzoneInstructions"] small { color: var(--faint) !important; }

/* ── Containers: expanders, tabs, editors, tables ────────────────────── */
details[data-testid="stExpander"] {
  background: var(--surface) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--r) !important;
  overflow: hidden;
}
details[data-testid="stExpander"] summary {
  min-height: 42px; color: var(--text) !important;
  font-size: .875rem !important; font-weight: 600 !important;
  transition: color var(--t) var(--ease);
}
details[data-testid="stExpander"] summary:hover { color: var(--accent) !important; }
details[data-testid="stExpander"] summary svg { fill: var(--dim) !important; }
details[data-testid="stExpander"][open] summary { border-bottom: 1px solid var(--border) !important; }

div[data-baseweb="tab-list"] {
  gap: .3rem; background: var(--surface); border: 1px solid var(--border);
  border-radius: var(--r); padding: .28rem;
}
div[data-baseweb="tab-list"] > button {
  min-height: 36px; border-radius: var(--r-sm); color: var(--dim);
  font-size: .82rem; font-weight: 600; padding: 0 .85rem;
  transition: color var(--t) var(--ease), background var(--t) var(--ease);
}
div[data-baseweb="tab-list"] > button:hover { color: var(--text); background: var(--surface2); }
div[data-baseweb="tab-list"] > button[aria-selected="true"] {
  background: var(--accent-soft); color: var(--accent);
  box-shadow: inset 0 0 0 1px var(--accent-line);
}
div[data-baseweb="tab-border"], div[data-baseweb="tab-highlight"] { background: transparent !important; }

div[data-testid="stDataEditor"], div[data-testid="stDataFrame"] {
  border: 1px solid var(--border) !important; border-radius: var(--r) !important;
  overflow: hidden; box-shadow: var(--shadow1);
}
[data-testid="stDataFrameResizable"] > div { border-radius: var(--r) !important; }

/* ── Alerts ──────────────────────────────────────────────────────────── */
div[data-testid="stAlert"] {
  border-radius: var(--r-sm) !important; border: 1px solid var(--border) !important;
  background: var(--surface) !important;
}
div[data-testid="stAlert"] [data-testid="stAlertMessage"] { color: var(--text) !important; }
div[data-testid="stAlert"][data-baseweb="notification"] { border-left-width: 3px !important; }

/* ── Progress, spinner, bar charts ───────────────────────────────────── */
div[data-testid="stProgress"] > div > div > div {
  background: linear-gradient(90deg, var(--accent), var(--accent2)) !important;
  border-radius: 99px !important;
}
div[data-testid="stProgress"] > div { background: var(--surface2) !important; }
div[data-testid="stSpinner"] > div { border-top-color: var(--accent) !important; border-right-color: var(--accent) !important; }
div[data-testid="stToolbar"] { background: var(--panel) !important; color: var(--text) !important; }
[data-testid="stChart"] svg text { fill: var(--dim) !important; }
[data-testid="stChart"] svg path { stroke: var(--accent2) !important; }

/* ── Motion: entrance ────────────────────────────────────────────────── */
@keyframes rise {
  from { opacity: 0; transform: translateY(16px); }
  to   { opacity: 1; transform: translateY(0); }
}
@keyframes slide {
  from { opacity: 0; transform: translateX(-14px); }
  to   { opacity: 1; transform: translateX(0); }
}
.hero-anim { animation: rise var(--t-slow) var(--ease-out) both; }
.d1 { animation-delay: .05s; } .d2 { animation-delay: .11s; }
.d3 { animation-delay: .17s; } .d4 { animation-delay: .23s; } .d5 { animation-delay: .29s; }

/* ── Motion: scroll-driven reveals (progressive enhancement) ─────────── */
@keyframes reveal {
  from { opacity: 0; transform: translateY(22px) scale(.988); }
  to   { opacity: 1; transform: translateY(0) scale(1); }
}
@supports (animation-timeline: view()) {
  @media (prefers-reduced-motion: no-preference) {
    .card-head, .mode-card, .section-sep, .sheet-item,
    div[data-testid="stMetric"], details[data-testid="stExpander"] {
      animation: reveal linear both;
      animation-timeline: view();
      animation-range: entry 0% entry 42%;
    }
    /* The mode grid reads as one gesture: the left column leads. */
    .mode-card:nth-child(4n+1) { animation-range: entry 0%  entry 48%; }
    .mode-card:nth-child(4n+2) { animation-range: entry 2%  entry 54%; }
    .mode-card:nth-child(4n+3) { animation-range: entry 4%  entry 60%; }
    .mode-card:nth-child(4n)   { animation-range: entry 6%  entry 66%; }
  }
}

/* ── Motion opt-out ──────────────────────────────────────────────────── */
@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after {
    animation-duration: .001ms !important;
    animation-iteration-count: 1 !important;
    transition-duration: .001ms !important;
    scroll-behavior: auto !important;
  }
  .scroll-fill { transform: scaleX(0) !important; }
}

/* ── Responsive ──────────────────────────────────────────────────────── */
@media (max-width: 900px) {
  .block-container, [data-testid="stMainBlockContainer"] { padding: 1.1rem 1.1rem 4rem !important; }
  .mode-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .card-head { padding: .9rem 1rem; }
}
@media (max-width: 560px) {
  .mode-grid { grid-template-columns: minmax(0, 1fr); }
  .card-title { font-size: .95rem; }
}
@media print {
  .stApp::before, .stApp::after, .scroll-rail { display: none !important; }
}
"""


def build_css(theme: str) -> str:
    """Render the stylesheet for `theme` (unknown names fall back to dark)."""
    values = TOKENS.get(theme, TOKENS["dark"])
    css = _STYLESHEET.replace("@@FONT@@", _FONT_HREF)
    for name, value in values.items():
        css = css.replace(f"@@{name}@@", value)
    return css


def inject(theme: str = "dark") -> None:
    """Install the design system into the running Streamlit app."""
    import streamlit as st

    st.markdown(f"<style>{build_css(theme)}</style>", unsafe_allow_html=True)


def scroll_rail() -> None:
    """A 2px sticky progress bar driven by the app's scroll position."""
    import streamlit as st

    st.markdown(
        '<div class="scroll-rail" aria-hidden="true">'
        '<span class="scroll-fill"></span></div>',
        unsafe_allow_html=True,
    )