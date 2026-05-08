"""
Design tokens — Editorial Academic
3-layer system: primitive → semantic → component

Reference: Radix UI / Geist / Stripe Press
Principle: every token earns its place. Primitives are raw OKLCH-tuned values;
semantics map intent (heading/body/rule) to primitives so a brand swap is one line.
"""

# ─────────────────────────────────────────────────────────────
# 1. PRIMITIVES (raw palette, no meaning)
# ─────────────────────────────────────────────────────────────
PRIMITIVE = {
    # ink scale — perceptually-even, deeper blacks for editorial heads
    "ink-900": "#1A1F36",
    "ink-700": "#2D3A5C",
    "ink-500": "#4A5A8C",   # 기존 브랜드 네이비
    "ink-300": "#8E97B5",
    "ink-100": "#D4D8E5",

    # paper scale — warm cream → off-white
    "paper-50":  "#FBF9F4",
    "paper-100": "#F7F4ED",  # 앱 바탕
    "paper-200": "#EFEAE0",
    "paper-300": "#E2DCCE",
    "paper-rule": "#D4CFC0",  # hairline 구분선

    # neutral cool grays (서브 텍스트, 메타 정보)
    "neutral-900": "#111827",
    "neutral-700": "#374151",
    "neutral-500": "#6B7280",
    "neutral-400": "#9CA3AF",
    "neutral-300": "#D1D5DB",
    "neutral-200": "#E5E7EB",
    "neutral-100": "#F3F4F6",

    # accents — sage(긍정/완료), amber(주의/킬러), terracotta(에러)
    "sage-700":   "#3F5F4F",
    "sage-500":   "#6B8E7F",
    "sage-100":   "#E5EEE9",

    "amber-700":  "#8A5A0E",
    "amber-500":  "#B07814",
    "amber-100":  "#F7EBC8",

    "terra-700":  "#8E3A2D",
    "terra-500":  "#B5503F",
    "terra-100":  "#F2DDD7",

    # surface
    "surface":    "#FFFFFF",
    "surface-overlay": "rgba(26, 31, 54, 0.04)",
}


# ─────────────────────────────────────────────────────────────
# 2. SEMANTIC (intent → primitive ref)
# ─────────────────────────────────────────────────────────────
SEMANTIC = {
    # text
    "text-heading":  PRIMITIVE["ink-900"],
    "text-body":     PRIMITIVE["neutral-900"],
    "text-muted":    PRIMITIVE["neutral-500"],
    "text-faint":    PRIMITIVE["neutral-400"],
    "text-accent":   PRIMITIVE["ink-700"],
    "text-on-dark":  "#FFFFFF",

    # backgrounds
    "bg-app":        PRIMITIVE["paper-100"],
    "bg-surface":    PRIMITIVE["surface"],
    "bg-elevated":   PRIMITIVE["paper-50"],
    "bg-subtle":     PRIMITIVE["paper-200"],
    "bg-inset":      PRIMITIVE["neutral-100"],

    # borders & rules
    "border-default":  PRIMITIVE["paper-rule"],
    "border-strong":   PRIMITIVE["paper-300"],
    "border-subtle":   PRIMITIVE["paper-200"],
    "rule-page":       PRIMITIVE["ink-700"],   # 본문 큰 구분선

    # interactive
    "action-primary":     PRIMITIVE["ink-700"],
    "action-primary-hover": PRIMITIVE["ink-900"],
    "action-secondary":   PRIMITIVE["ink-500"],
    "focus-ring":         PRIMITIVE["ink-500"],

    # status
    "status-ok-fg":      PRIMITIVE["sage-700"],
    "status-ok-bg":      PRIMITIVE["sage-100"],
    "status-warn-fg":    PRIMITIVE["amber-700"],
    "status-warn-bg":    PRIMITIVE["amber-100"],
    "status-danger-fg":  PRIMITIVE["terra-700"],
    "status-danger-bg":  PRIMITIVE["terra-100"],

    # killer 강조 — 시험지 분석 전용
    "killer-fg":  PRIMITIVE["terra-700"],
    "killer-bg":  PRIMITIVE["terra-100"],
}


# ─────────────────────────────────────────────────────────────
# 3. SCALES (typography / spacing / radius)
# ─────────────────────────────────────────────────────────────
TYPE = {
    # body Pretendard, 숫자 IBM Plex Mono — editorial 대비
    "font-sans": "'Pretendard Variable', 'Pretendard', -apple-system, BlinkMacSystemFont, 'Malgun Gothic', sans-serif",
    "font-mono": "'IBM Plex Mono', 'JetBrains Mono', 'D2Coding', 'Consolas', monospace",
    "font-serif-display": "'Pretendard Variable', 'Pretendard', serif",  # 본문 통일성 위해 sans 유지

    # type ramp — editorial: 더 큰 H1, 좁은 본문
    "size-display": "32px",  # 페이지 타이틀
    "size-h1":      "26px",
    "size-h2":      "19px",
    "size-h3":      "15px",
    "size-body":    "14.5px",
    "size-meta":    "12.5px",
    "size-caption": "11.5px",

    "leading-tight":  "1.2",
    "leading-normal": "1.5",
    "leading-loose":  "1.7",

    "track-tight":  "-0.022em",   # 큰 텍스트
    "track-normal": "0",
    "track-wide":   "0.06em",     # uppercase 라벨
}

SPACE = {
    "0": "0",
    "1": "4px",
    "2": "8px",
    "3": "12px",
    "4": "16px",
    "5": "20px",
    "6": "24px",
    "8": "32px",
    "10": "40px",
    "12": "56px",
}

RADIUS = {
    "sm":  "4px",   # 태그/badge
    "md":  "8px",   # 입력/버튼
    "lg":  "12px",  # 카드
    "xl":  "16px",  # 페이지 섹션
    "full": "9999px",
}


# ─────────────────────────────────────────────────────────────
# 4. CSS — semantic & component (not primitive)
# ─────────────────────────────────────────────────────────────
def build_css() -> str:
    """Streamlit `st.markdown` 으로 주입할 CSS 한 덩어리."""
    s = SEMANTIC
    t = TYPE
    sp = SPACE
    r = RADIUS

    return f"""
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css');
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&display=swap');

:root {{
    /* semantic tokens — components reference these, never primitives */
    --text-heading:  {s["text-heading"]};
    --text-body:     {s["text-body"]};
    --text-muted:    {s["text-muted"]};
    --text-faint:    {s["text-faint"]};
    --text-accent:   {s["text-accent"]};

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

/* ── Sidebar — editorial 좌측 여백판 ── */
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

/* ── Editorial label (서브헤딩, 출판 느낌) ── */
.eyebrow {{
    font-size: {t["size-caption"]};
    font-weight: 600;
    color: var(--text-muted);
    text-transform: uppercase;
    letter-spacing: {t["track-wide"]};
    margin-bottom: {sp["2"]};
}}

/* ── Card — 종이 위 인쇄된 느낌, hairline rule ── */
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

/* ── KPI strip — editorial metric row ── */
.kpi-strip {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: {sp["3"]};
    margin: {sp["4"]} 0;
}}

/* ── Tag / Badge ── */
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

/* ── Buttons — editorial: 단단한 검은 잉크 또는 ghost ── */
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

/* ── Tabs — editorial: 큰 본문, 가는 underline ── */
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

/* ── Progress ── */
.stProgress > div > div > div {{
    background: var(--text-heading) !important;
    border-radius: var(--r-full) !important;
}}

/* ── Dividers — page rule (출판물 § 구분) ── */
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

/* ── Empty state ── */
.empty {{
    text-align: center;
    padding: 56px 24px;
    color: var(--text-faint);
    font-size: {t["size-body"]};
    background: var(--bg-elevated);
    border: 1px dashed var(--border-default);
    border-radius: var(--r-lg);
}}

/* ── Mono numerics — 점수, 통계, 표 안의 숫자 ── */
.mono, .mono * {{
    font-family: var(--font-mono) !important;
    font-variant-numeric: tabular-nums !important;
}}

/* ── Killer flag inline ── */
.killer-flag {{
    color: var(--killer-fg);
    font-weight: 700;
    font-family: var(--font-mono);
    margin-right: 4px;
}}

/* ── Editorial summary 박스 ── */
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

/* ── Section label (sidebar) ── */
.section-label {{
    font-size: {t["size-caption"]};
    font-weight: 700;
    color: var(--text-faint);
    text-transform: uppercase;
    letter-spacing: {t["track-wide"]};
    margin: {sp["6"]} 0 {sp["2"]} 0;
}}

/* ── Table-like rows for question list ── */
.q-row {{
    display: grid;
    grid-template-columns: 36px 1fr 80px 90px 60px;
    gap: {sp["3"]};
    padding: {sp["3"]} {sp["4"]};
    border-bottom: 1px solid var(--border-subtle);
    align-items: center;
    font-size: {t["size-body"]};
}}
.q-row:hover {{ background: var(--bg-elevated); }}
.q-row.killer {{ background: var(--killer-bg); }}
.q-no {{ font-family: var(--font-mono); color: var(--text-faint); }}

/* ── Hide defaults ── */
#MainMenu, footer {{ visibility: hidden; }}
.stDeployButton {{ display: none; }}
</style>
"""
