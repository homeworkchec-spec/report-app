"""
Design tokens — 3-theme system (Mono / Editorial / Vivid)

3-layer: primitive → semantic → component
사용자가 UI에서 테마를 선택하면 build_css(theme) 가 그에 맞는 단일 CSS 를 반환합니다.

Theme 의도:
  · mono       — 색 없음. 흑백+회색조. 인쇄물처럼 가장 단정.
  · editorial  — 종이(크림) + 잉크 + 절제된 액센트. (기본)
  · vivid      — 흰 바탕 + 활발한 차트 색. 블로그 친화.

각 테마는 동일한 의미 토큰(token) 이름을 공유 — 컴포넌트 CSS는 한 번만 작성.
"""

from typing import Literal

ThemeName = Literal["mono", "editorial", "vivid"]


# ─────────────────────────────────────────────────────────────
# 1. 공통 타이포/스페이싱/라디우스 — 테마 무관
# ─────────────────────────────────────────────────────────────
TYPE = {
    "font-sans": "'Pretendard Variable', 'Pretendard', -apple-system, BlinkMacSystemFont, 'Malgun Gothic', sans-serif",
    "font-mono": "'IBM Plex Mono', 'JetBrains Mono', 'D2Coding', 'Consolas', monospace",
    "size-display": "32px",
    "size-h1":      "26px",
    "size-h2":      "19px",
    "size-h3":      "15px",
    "size-body":    "14.5px",
    "size-meta":    "12.5px",
    "size-caption": "11.5px",
    "leading-tight":  "1.2",
    "leading-normal": "1.5",
    "leading-loose":  "1.7",
    "track-tight":  "-0.022em",
    "track-normal": "0",
    "track-wide":   "0.06em",
}
SPACE = {f"{k}": v for k, v in [
    ("0", "0"), ("1", "4px"), ("2", "8px"), ("3", "12px"), ("4", "16px"),
    ("5", "20px"), ("6", "24px"), ("8", "32px"), ("10", "40px"), ("12", "56px"),
]}
RADIUS = {"sm": "4px", "md": "8px", "lg": "12px", "xl": "16px", "full": "9999px"}


# ─────────────────────────────────────────────────────────────
# 2. 테마별 SEMANTIC 매핑 — 컴포넌트가 참조하는 추상 이름
# ─────────────────────────────────────────────────────────────
def _theme_tokens(theme: ThemeName) -> dict[str, str]:
    if theme == "mono":
        return {
            "text-heading":  "#0A0A0A",
            "text-body":     "#1F1F1F",
            "text-muted":    "#6B7280",
            "text-faint":    "#A1A6B0",
            "text-accent":   "#0A0A0A",
            "text-on-dark":  "#FFFFFF",

            "bg-app":        "#FFFFFF",
            "bg-surface":    "#FFFFFF",
            "bg-elevated":   "#FAFAFA",
            "bg-subtle":     "#F4F4F5",
            "bg-inset":      "#EEEEEF",

            "border-default":  "#E5E7EB",
            "border-strong":   "#D4D4D8",
            "rule-page":       "#0A0A0A",

            "action-primary":        "#0A0A0A",
            "action-primary-hover":  "#262626",

            "status-ok-fg":      "#1F2937",
            "status-ok-bg":      "#F4F4F5",
            "status-warn-fg":    "#1F2937",
            "status-warn-bg":    "#F4F4F5",
            "status-danger-fg":  "#1F2937",
            "status-danger-bg":  "#F4F4F5",
            "killer-fg":  "#0A0A0A",
            "killer-bg":  "#F4F4F5",
        }
    if theme == "vivid":
        return {
            "text-heading":  "#1A1F36",
            "text-body":     "#1F2937",
            "text-muted":    "#6B7280",
            "text-faint":    "#9CA3AF",
            "text-accent":   "#4A5A8C",
            "text-on-dark":  "#FFFFFF",

            "bg-app":        "#F8F9FC",
            "bg-surface":    "#FFFFFF",
            "bg-elevated":   "#F0F2F8",
            "bg-subtle":     "#E5EAF3",
            "bg-inset":      "#EEF1F8",

            "border-default":  "#E5E7EB",
            "border-strong":   "#CBD2DD",
            "rule-page":       "#4A5A8C",

            "action-primary":        "#4A5A8C",
            "action-primary-hover":  "#2D3A5C",

            "status-ok-fg":      "#0F766E",
            "status-ok-bg":      "#CCFBF1",
            "status-warn-fg":    "#92400E",
            "status-warn-bg":    "#FEF3C7",
            "status-danger-fg":  "#991B1B",
            "status-danger-bg":  "#FEE2E2",
            "killer-fg":  "#9A3412",
            "killer-bg":  "#FED7AA",
        }
    # editorial — default
    return {
        "text-heading":  "#1A1F36",
        "text-body":     "#1F2937",
        "text-muted":    "#6B7280",
        "text-faint":    "#9CA3AF",
        "text-accent":   "#2D3A5C",
        "text-on-dark":  "#FFFFFF",

        "bg-app":        "#F7F4ED",
        "bg-surface":    "#FFFFFF",
        "bg-elevated":   "#FBF9F4",
        "bg-subtle":     "#EFEAE0",
        "bg-inset":      "#F3F4F6",

        "border-default":  "#D4CFC0",
        "border-strong":   "#E2DCCE",
        "rule-page":       "#2D3A5C",

        "action-primary":        "#2D3A5C",
        "action-primary-hover":  "#1A1F36",

        "status-ok-fg":      "#3F5F4F",
        "status-ok-bg":      "#E5EEE9",
        "status-warn-fg":    "#8A5A0E",
        "status-warn-bg":    "#F7EBC8",
        "status-danger-fg":  "#8E3A2D",
        "status-danger-bg":  "#F2DDD7",
        "killer-fg":  "#8E3A2D",
        "killer-bg":  "#F2DDD7",
    }


# ─────────────────────────────────────────────────────────────
# 3. 테마별 차트 팔레트 — exam_analysis 가 import 해서 사용
# ─────────────────────────────────────────────────────────────
CHART_PALETTES: dict[str, dict] = {
    "mono": {
        "ink":      "#0A0A0A",
        "paper":    "#FFFFFF",
        "rule":     "#D4D4D8",
        "accents":  ["#0A0A0A", "#404040", "#737373", "#A3A3A3", "#262626", "#525252", "#171717"],
    },
    "editorial": {
        "ink":      "#1A1F36",
        "paper":    "#FFFFFF",
        "rule":     "#D4CFC0",
        "accents":  ["#1A1F36", "#6B8E7F", "#B07814", "#B5503F", "#2D3A5C",
                     "#8E97B5", "#6B73B5", "#5E8DA8", "#7B9B8E", "#C18A65"],
    },
    "vivid": {
        "ink":      "#1A1F36",
        "paper":    "#FFFFFF",
        "rule":     "#E5E7EB",
        "accents":  ["#4A5A8C", "#0EA5E9", "#10B981", "#F59E0B", "#EF4444",
                     "#8B5CF6", "#EC4899", "#14B8A6", "#F97316", "#6366F1"],
    },
}


def chart_palette(theme: ThemeName) -> dict:
    return CHART_PALETTES.get(theme, CHART_PALETTES["editorial"])


# ─────────────────────────────────────────────────────────────
# 4. CSS template — 컴포넌트 규칙은 모두 의미 토큰 참조
# ─────────────────────────────────────────────────────────────
def build_css(theme: ThemeName = "editorial") -> str:
    s = _theme_tokens(theme)
    t = TYPE
    sp = SPACE
    r = RADIUS

    return f"""
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css');
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&display=swap');

:root {{
    --text-heading:  {s["text-heading"]};
    --text-body:     {s["text-body"]};
    --text-muted:    {s["text-muted"]};
    --text-faint:    {s["text-faint"]};
    --text-accent:   {s["text-accent"]};
    --text-on-dark:  {s["text-on-dark"]};

    --bg-app:        {s["bg-app"]};
    --bg-surface:    {s["bg-surface"]};
    --bg-elevated:   {s["bg-elevated"]};
    --bg-subtle:     {s["bg-subtle"]};
    --bg-inset:      {s["bg-inset"]};

    --border-default: {s["border-default"]};
    --border-strong:  {s["border-strong"]};
    --rule-page:      {s["rule-page"]};

    --action-primary:        {s["action-primary"]};
    --action-primary-hover:  {s["action-primary-hover"]};

    --status-ok-fg:     {s["status-ok-fg"]};
    --status-ok-bg:     {s["status-ok-bg"]};
    --status-warn-fg:   {s["status-warn-fg"]};
    --status-warn-bg:   {s["status-warn-bg"]};
    --status-danger-fg: {s["status-danger-fg"]};
    --status-danger-bg: {s["status-danger-bg"]};
    --killer-fg:        {s["killer-fg"]};
    --killer-bg:        {s["killer-bg"]};

    --font-sans: {t["font-sans"]};
    --font-mono: {t["font-mono"]};

    --r-sm: {r["sm"]};
    --r-md: {r["md"]};
    --r-lg: {r["lg"]};
    --r-xl: {r["xl"]};
}}

/* ── Global ── */
html, body, .stApp,
.stApp [class*="css"] {{
    font-family: var(--font-sans) !important;
    color: var(--text-body);
    -webkit-font-smoothing: antialiased;
    text-rendering: optimizeLegibility;
}}
.stApp {{ background: var(--bg-app); }}

/* ── Sidebar ── */
section[data-testid="stSidebar"] {{
    background: var(--bg-surface);
    border-right: 1px solid var(--border-default);
}}
section[data-testid="stSidebar"] .stMarkdown p {{
    font-size: {t["size-meta"]};
    color: var(--text-muted);
}}

/* ── Typography ── */
h1, .editorial-title {{
    color: var(--text-heading) !important;
    font-weight: 700 !important;
    font-size: {t["size-display"]} !important;
    letter-spacing: {t["track-tight"]};
    line-height: {t["leading-tight"]};
}}
h2 {{
    color: var(--text-heading) !important;
    font-weight: 600 !important;
    font-size: {t["size-h1"]} !important;
    letter-spacing: {t["track-tight"]};
}}
h3 {{
    color: var(--text-heading) !important;
    font-weight: 600 !important;
    font-size: {t["size-h2"]} !important;
}}
h4, h5 {{
    color: var(--text-heading) !important;
    font-weight: 600 !important;
    font-size: {t["size-h3"]} !important;
}}

.eyebrow {{
    font-size: {t["size-caption"]};
    font-weight: 600;
    color: var(--text-muted);
    text-transform: uppercase;
    letter-spacing: {t["track-wide"]};
    margin-bottom: {sp["2"]};
}}

/* ── Card ── */
.card {{
    background: var(--bg-surface);
    border: 1px solid var(--border-default);
    border-radius: var(--r-lg);
    padding: {sp["5"]} {sp["6"]};
    margin-bottom: {sp["3"]};
}}
.card-eyebrow {{
    font-size: {t["size-caption"]};
    font-weight: 600;
    color: var(--text-faint);
    text-transform: uppercase;
    letter-spacing: {t["track-wide"]};
    margin-bottom: {sp["2"]};
}}
.card-value {{
    font-family: var(--font-mono);
    font-size: 28px;
    font-weight: 600;
    color: var(--text-heading);
    line-height: 1.1;
    letter-spacing: -0.01em;
}}
.card-meta {{
    font-size: {t["size-meta"]};
    color: var(--text-muted);
    margin-top: {sp["2"]};
}}

.kpi-strip {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: {sp["3"]};
    margin: {sp["4"]} 0;
}}

/* ── Tag ── */
.tag {{
    display: inline-block;
    padding: 2px 8px;
    border-radius: var(--r-sm);
    font-size: {t["size-caption"]};
    font-weight: 600;
    letter-spacing: 0.02em;
    border: 1px solid transparent;
}}
.tag-ok      {{ background: var(--status-ok-bg);     color: var(--status-ok-fg); }}
.tag-warn    {{ background: var(--status-warn-bg);   color: var(--status-warn-fg); }}
.tag-danger  {{ background: var(--status-danger-bg); color: var(--status-danger-fg); }}
.tag-killer  {{ background: var(--killer-bg);        color: var(--killer-fg);
                border-color: var(--killer-fg); }}
.tag-neutral {{ background: var(--bg-inset);         color: var(--text-muted); }}

/* ── Buttons ── */
.stButton > button {{
    border-radius: var(--r-md) !important;
    font-weight: 500 !important;
    font-size: {t["size-body"]} !important;
    padding: 8px 18px !important;
    border: 1px solid var(--border-strong) !important;
    background: var(--bg-surface) !important;
    color: var(--text-heading) !important;
    transition: background .12s ease, border-color .12s ease !important;
}}
.stButton > button:hover {{
    background: var(--bg-elevated) !important;
    border-color: var(--text-heading) !important;
}}
.stButton > button[kind="primary"] {{
    background: var(--action-primary) !important;
    color: var(--text-on-dark, #fff) !important;
    border-color: var(--action-primary) !important;
}}
.stButton > button[kind="primary"]:hover {{
    background: var(--action-primary-hover) !important;
    border-color: var(--action-primary-hover) !important;
}}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {{
    gap: 0;
    border-bottom: 1px solid var(--border-default);
    background: transparent;
}}
.stTabs [data-baseweb="tab"] {{
    border-radius: 0 !important;
    padding: 12px 22px !important;
    font-weight: 500 !important;
    font-size: {t["size-body"]} !important;
    color: var(--text-muted) !important;
    border-bottom: 2px solid transparent;
    margin-bottom: -1px;
    background: transparent !important;
}}
.stTabs [aria-selected="true"] {{
    color: var(--text-heading) !important;
    border-bottom-color: var(--text-heading) !important;
}}

.stProgress > div > div > div {{
    background: var(--text-heading) !important;
    border-radius: var(--r-full) !important;
}}

/* ── Dividers ── */
.divider {{
    height: 1px;
    background: var(--border-default);
    margin: {sp["5"]} 0;
}}
.divider-strong {{
    height: 1px;
    background: var(--text-heading);
    margin: {sp["6"]} 0 {sp["4"]} 0;
    opacity: 0.85;
}}
.section-mark {{
    font-family: var(--font-mono);
    font-size: {t["size-caption"]};
    color: var(--text-faint);
    letter-spacing: {t["track-wide"]};
    margin: {sp["6"]} 0 {sp["2"]} 0;
}}

/* ── Empty ── */
.empty {{
    text-align: center;
    padding: 56px 24px;
    color: var(--text-faint);
    font-size: {t["size-body"]};
    background: var(--bg-elevated);
    border: 1px dashed var(--border-default);
    border-radius: var(--r-lg);
}}

.mono, .mono * {{
    font-family: var(--font-mono) !important;
    font-variant-numeric: tabular-nums !important;
}}

.killer-flag {{
    color: var(--killer-fg);
    font-weight: 700;
    font-family: var(--font-mono);
    margin-right: 4px;
}}

.summary-box {{
    background: var(--bg-elevated);
    border: 1px solid var(--border-default);
    border-left: 3px solid var(--text-heading);
    padding: {sp["5"]} {sp["6"]};
    border-radius: 0 var(--r-md) var(--r-md) 0;
    font-size: {t["size-body"]};
    line-height: {t["leading-loose"]};
    color: var(--text-body);
}}
.summary-box .lead {{
    font-weight: 600;
    color: var(--text-heading);
    margin-bottom: {sp["2"]};
}}

.section-label {{
    font-size: {t["size-caption"]};
    font-weight: 700;
    color: var(--text-faint);
    text-transform: uppercase;
    letter-spacing: {t["track-wide"]};
    margin: {sp["6"]} 0 {sp["2"]} 0;
}}

/* ── Theme picker (사이드바) ── */
.theme-picker .stRadio > div {{
    flex-direction: row;
    gap: 6px;
}}
.theme-picker label {{
    font-size: 12px !important;
}}

#MainMenu, footer {{ visibility: hidden; }}
.stDeployButton {{ display: none; }}
</style>
"""
