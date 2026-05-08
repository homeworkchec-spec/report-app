"""
시험지 분석 자동화 모듈
─────────────────────────────────────────────────────────────
워크플로우:
  1) 시험지 이미지 업로드  → GPT-4o Vision OCR + 구조화
  2) 메타정보 검토/수정    → st.data_editor 로 문항별 편집
  3) 킬러문항 제안 + 총평 → GPT 분석 + matplotlib 차트
  4) Word 보고서 다운로드  → 한들/레이더차트 패턴 재사용

본 모듈은 app.py 의 4번째 탭에서 호출됩니다.
"""

from __future__ import annotations

import base64
import io
import json
import re
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import Any

import platform
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image


# ─────────────────────────────────────────────────────────────
# Korean font setup — chart 의 한글 깨짐 방지
# 우선순위: 레포 .fonts 폴더 → 시스템 한글 폰트 → fallback
# ─────────────────────────────────────────────────────────────
def _setup_korean_font():
    candidates = []
    repo_font = Path(__file__).parent / ".fonts"
    if repo_font.exists():
        candidates += list(repo_font.glob("*.ttf")) + list(repo_font.glob("*.otf"))
    # 시스템 후보
    if platform.system() == "Windows":
        for n in ("malgun.ttf", "malgunbd.ttf", "NanumGothic.ttf"):
            p = Path("C:/Windows/Fonts") / n
            if p.exists():
                candidates.append(p)
    elif platform.system() == "Darwin":
        for n in ("AppleSDGothicNeo.ttc", "AppleGothic.ttf"):
            p = Path("/System/Library/Fonts") / n
            if p.exists():
                candidates.append(p)
    else:  # Linux
        for n in ("NanumGothic.ttf", "NotoSansCJK-Regular.ttc", "malgun.ttf"):
            for base in ("/usr/share/fonts", str(Path.home() / ".fonts")):
                for p in Path(base).rglob(n) if Path(base).exists() else []:
                    candidates.append(p)
                    break

    for path in candidates:
        try:
            fm.fontManager.addfont(str(path))
            family = fm.FontProperties(fname=str(path)).get_name()
            matplotlib.rcParams["font.family"] = family
            matplotlib.rcParams["axes.unicode_minus"] = False
            return family
        except Exception:
            continue
    matplotlib.rcParams["axes.unicode_minus"] = False
    return None


_KO_FONT = _setup_korean_font()

import openai

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


# ─────────────────────────────────────────────────────────────
# 0. CONSTANTS
# ─────────────────────────────────────────────────────────────
DIFFICULTY_LEVELS = ["하", "중하", "중", "중상", "상"]
DIFFICULTY_NUM = {"하": 1, "중하": 2, "중": 3, "중상": 4, "상": 5}

QUESTION_TYPE_PRESETS = {
    "영어": ["일치/불일치", "빈칸", "어법", "어휘", "순서", "삽입", "요약",
             "함의", "대의", "지칭", "영영풀이", "조건영작", "서술형", "단답형"],
    "수학": ["계산", "개념", "응용", "증명", "도형", "함수", "확률통계", "킬러"],
    "국어": ["독서", "문학", "문법", "화법작문", "어휘"],
    "기타": ["객관식", "주관식", "서술형", "단답형", "수행평가"],
}

VISION_MODEL = "gpt-4o"           # 이미지→구조화 추출
TEXT_MODEL_DEFAULT = "gpt-4o"     # 분석/총평 (app.py 의 모델 변수와 별도)


# ─────────────────────────────────────────────────────────────
# 1. SCHEMA
# ─────────────────────────────────────────────────────────────
@dataclass
class ExamMeta:
    title: str = ""
    school: str = ""
    grade: str = ""
    subject: str = "영어"
    exam_type: str = "중간고사"     # 중간/기말/모의/기타
    exam_date: str = ""
    duration_min: int = 50
    total_score: int = 100
    total_questions: int = 0
    notes: str = ""


@dataclass
class Question:
    no: int = 0
    type: str = ""            # 유형
    difficulty: str = "중"     # 하/중하/중/중상/상
    score: float = 0.0
    is_subjective: bool = False  # 객관식 False, 서답형 True
    is_killer: bool = False
    scope: str = ""            # 출제 범위/단원
    memo: str = ""

    def difficulty_num(self) -> int:
        return DIFFICULTY_NUM.get(self.difficulty, 3)


# ─────────────────────────────────────────────────────────────
# 2. OCR — GPT-4o Vision 으로 한 번에 메타+문항 추출
# ─────────────────────────────────────────────────────────────
OCR_SYSTEM = (
    "당신은 한국 중·고등학교 시험지를 분석하는 전문가입니다. "
    "주어진 시험지 이미지를 읽고, 시험 메타 정보와 모든 문항의 유형·난이도·배점을 추출합니다. "
    "추정이 어려운 값은 빈 문자열로 두되, 문항 번호는 반드시 채우세요. "
    "응답은 반드시 단일 JSON 객체로만 출력하세요."
)

OCR_USER_TEMPLATE = """다음 시험지 이미지를 분석하여 JSON 으로 반환하세요.

[과목 힌트] {subject}
[학교급 힌트] {grade}

[필수 JSON 스키마]
{{
  "exam_meta": {{
    "title": "...",          // 시험지 상단 제목
    "school": "...",         // 학교명 (없으면 "")
    "grade": "...",          // 예: "고2", "중3"
    "subject": "...",        // 과목
    "exam_type": "...",      // 중간고사/기말고사/모의고사/수행평가
    "exam_date": "YYYY-MM-DD",  // 추정 어려우면 ""
    "duration_min": 50,
    "total_score": 100,
    "total_questions": 25,
    "notes": "..."           // 출제범위·시험범위 등 시험지에 명시된 안내
  }},
  "questions": [
    {{
      "no": 1,
      "type": "어법",            // 한국 시험지 표준 유형명 사용
      "difficulty": "중",       // 하/중하/중/중상/상 중 하나
      "score": 4.0,
      "is_subjective": false,   // 서답형(서술/단답)이면 true
      "scope": "Lesson 3",      // 시험범위 명시 시
      "memo": ""
    }}
    // ... 모든 문항
  ]
}}

[난이도 판단 기준]
- 하: 단순 일치/암기/기본 어휘
- 중하: 표준 개념 적용
- 중: 다단계 추론, 복합 어법
- 중상: 변형/응용, 복합 조건
- 상(킬러 후보): 다중 변형·시간소모·고난도 추론

JSON 외 다른 텍스트는 절대 출력하지 마세요."""


def _img_to_data_url(img_bytes: bytes, mime: str = "image/png") -> str:
    b64 = base64.b64encode(img_bytes).decode("ascii")
    return f"data:{mime};base64,{b64}"


def _normalize_image(file_bytes: bytes, max_side: int = 2000) -> tuple[bytes, str]:
    """과대 이미지 축소 + JPEG 통일 — 비용/지연 절감."""
    try:
        im = Image.open(io.BytesIO(file_bytes))
        im = im.convert("RGB")
        w, h = im.size
        scale = min(1.0, max_side / max(w, h))
        if scale < 1.0:
            im = im.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=88, optimize=True)
        return buf.getvalue(), "image/jpeg"
    except Exception:
        return file_bytes, "image/png"


def _safe_json_loads(text: str) -> dict:
    """모델이 코드블록을 감쌀 때 대비."""
    text = text.strip()
    m = re.search(r"```(?:json)?\s*(.+?)```", text, re.DOTALL)
    if m:
        text = m.group(1).strip()
    return json.loads(text)


def ocr_exam_images(
    api_key: str,
    images: list[bytes],
    subject_hint: str = "영어",
    grade_hint: str = "고2",
) -> tuple[ExamMeta, list[Question]]:
    """여러 페이지 이미지를 한 번에 분석. 멀티 페이지면 모두 한 메시지로 묶어 일관성 유지."""
    client = openai.OpenAI(api_key=api_key)

    content: list[dict[str, Any]] = [
        {
            "type": "text",
            "text": OCR_USER_TEMPLATE.format(subject=subject_hint, grade=grade_hint),
        }
    ]
    for raw in images:
        norm, mime = _normalize_image(raw)
        content.append({
            "type": "image_url",
            "image_url": {"url": _img_to_data_url(norm, mime), "detail": "high"},
        })

    resp = client.chat.completions.create(
        model=VISION_MODEL,
        messages=[
            {"role": "system", "content": OCR_SYSTEM},
            {"role": "user", "content": content},
        ],
        temperature=0.1,
        max_tokens=4096,
        response_format={"type": "json_object"},
    )
    raw_text = resp.choices[0].message.content or "{}"
    data = _safe_json_loads(raw_text)

    meta_d = data.get("exam_meta", {}) or {}
    meta = ExamMeta(
        title=str(meta_d.get("title") or "").strip(),
        school=str(meta_d.get("school") or "").strip(),
        grade=str(meta_d.get("grade") or grade_hint).strip(),
        subject=str(meta_d.get("subject") or subject_hint).strip(),
        exam_type=str(meta_d.get("exam_type") or "중간고사").strip(),
        exam_date=str(meta_d.get("exam_date") or "").strip(),
        duration_min=int(meta_d.get("duration_min") or 50),
        total_score=int(meta_d.get("total_score") or 100),
        total_questions=int(meta_d.get("total_questions") or 0),
        notes=str(meta_d.get("notes") or "").strip(),
    )

    qs: list[Question] = []
    for q in data.get("questions", []) or []:
        try:
            diff = q.get("difficulty", "중")
            if diff not in DIFFICULTY_LEVELS:
                diff = "중"
            qs.append(Question(
                no=int(q.get("no", 0) or 0),
                type=str(q.get("type") or "").strip(),
                difficulty=diff,
                score=float(q.get("score") or 0),
                is_subjective=bool(q.get("is_subjective", False)),
                scope=str(q.get("scope") or "").strip(),
                memo=str(q.get("memo") or "").strip(),
            ))
        except Exception:
            continue
    qs.sort(key=lambda x: x.no)

    if meta.total_questions == 0:
        meta.total_questions = len(qs)

    return meta, qs


# ─────────────────────────────────────────────────────────────
# 3. KILLER 제안 — 메타데이터 휴리스틱 + GPT 보강
# ─────────────────────────────────────────────────────────────
def heuristic_killer_flags(meta: ExamMeta, qs: list[Question]) -> list[bool]:
    """난이도 상 + 배점 상위 25% + 서답형 복합 → 킬러 후보."""
    if not qs:
        return []
    scores = [q.score for q in qs if q.score > 0]
    score_threshold = sorted(scores)[int(len(scores) * 0.75)] if scores else 0

    flags = []
    for q in qs:
        is_killer = (
            q.difficulty == "상"
            or (q.difficulty == "중상" and q.score >= score_threshold)
            or (q.is_subjective and q.score >= score_threshold and q.difficulty in ("중상", "상"))
        )
        flags.append(is_killer)
    return flags


KILLER_SYSTEM = (
    "당신은 입시 전문 학원의 시험 분석 전문가입니다. "
    "주어진 시험지 메타데이터와 문항 리스트를 보고, "
    "변별력의 핵심이 되는 '킬러문항' 후보를 선정하고 그 이유를 한국어로 설명합니다."
)

KILLER_USER = """[시험 정보]
{meta_brief}

[문항 메타]
{table}

지시:
1) 위 문항 중 변별력 핵심인 킬러문항 후보 2~5개를 고르세요.
2) 각각의 선정 사유를 한 문장으로 적되, 어떤 능력을 변별하는지 명확히 하세요.
3) 응답은 JSON: {{"killers":[{{"no":3,"reason":"..."}}, ...], "rationale":"전체 변별 전략 1문단"}}

JSON 외 다른 텍스트는 출력하지 마세요."""


def suggest_killers(api_key: str, meta: ExamMeta, qs: list[Question]) -> dict:
    if not qs:
        return {"killers": [], "rationale": ""}
    client = openai.OpenAI(api_key=api_key)

    table_lines = ["번호 | 유형 | 난이도 | 배점 | 서답형 | 메모"]
    for q in qs:
        table_lines.append(
            f"{q.no} | {q.type} | {q.difficulty} | {q.score} | {'Y' if q.is_subjective else 'N'} | {q.memo}"
        )
    table = "\n".join(table_lines)

    meta_brief = (
        f"{meta.school} {meta.grade} {meta.subject} {meta.exam_type} · "
        f"총 {meta.total_questions}문항 · {meta.total_score}점 · {meta.duration_min}분"
    )

    resp = client.chat.completions.create(
        model=TEXT_MODEL_DEFAULT,
        messages=[
            {"role": "system", "content": KILLER_SYSTEM},
            {"role": "user", "content": KILLER_USER.format(meta_brief=meta_brief, table=table)},
        ],
        temperature=0.3,
        max_tokens=900,
        response_format={"type": "json_object"},
    )
    return _safe_json_loads(resp.choices[0].message.content or "{}")


# ─────────────────────────────────────────────────────────────
# 4. 총평 — editorial 스타일 분석문
# ─────────────────────────────────────────────────────────────
SUMMARY_SYSTEM = (
    "당신은 학원 원장이 학부모에게 보내는 시험 분석문을 작성하는 전문가입니다. "
    "출판물 같은 단정한 한국어, 과장 없는 분석, 학습 전략까지 담은 3~4단락 글을 씁니다. "
    "이모지·해시태그·과도한 강조는 사용하지 마세요. 단단한 문어체를 유지하세요."
)

SUMMARY_USER = """[시험 정보]
{meta_brief}

[유형 분포]
{type_dist}

[난이도 분포]
{diff_dist}

[킬러 후보]
{killer_brief}

요청: 다음 4개 단락으로 구성된 총평을 작성해 주세요.
1) 시험의 한 줄 요약 + 출제 의도
2) 유형/난이도 분포에서 드러나는 출제 경향
3) 변별력의 핵심 (킬러 문항 중심)
4) 다음 시험을 위한 학습 전략 (3가지 구체 액션)

각 단락 사이에는 빈 줄을 한 줄 두세요."""


def gen_summary(
    api_key: str,
    meta: ExamMeta,
    qs: list[Question],
    killers: dict,
) -> str:
    if not qs:
        return ""
    client = openai.OpenAI(api_key=api_key)

    type_counts = pd.Series([q.type for q in qs if q.type]).value_counts()
    type_dist = "\n".join(f"- {k}: {v}문항" for k, v in type_counts.items())

    diff_counts = pd.Series([q.difficulty for q in qs]).value_counts().reindex(DIFFICULTY_LEVELS, fill_value=0)
    diff_dist = "\n".join(f"- {k}: {v}문항" for k, v in diff_counts.items())

    killer_list = killers.get("killers", []) if killers else []
    killer_brief = "\n".join(f"- {k['no']}번: {k.get('reason', '')}" for k in killer_list) or "- 없음"

    meta_brief = (
        f"{meta.school} {meta.grade} {meta.subject} {meta.exam_type} · "
        f"총 {meta.total_questions}문항 · {meta.total_score}점 · {meta.duration_min}분"
    )

    resp = client.chat.completions.create(
        model=TEXT_MODEL_DEFAULT,
        messages=[
            {"role": "system", "content": SUMMARY_SYSTEM},
            {"role": "user", "content": SUMMARY_USER.format(
                meta_brief=meta_brief,
                type_dist=type_dist,
                diff_dist=diff_dist,
                killer_brief=killer_brief,
            )},
        ],
        temperature=0.55,
        max_tokens=1200,
    )
    return (resp.choices[0].message.content or "").strip()


# ─────────────────────────────────────────────────────────────
# 5. CHARTS — editorial 톤 (검은 잉크 + 종이 배경)
# ─────────────────────────────────────────────────────────────
EDITORIAL_INK = "#1A1F36"
EDITORIAL_PAPER = "#F7F4ED"
EDITORIAL_RULE = "#D4CFC0"
EDITORIAL_ACCENTS = ["#1A1F36", "#6B8E7F", "#B07814", "#B5503F", "#2D3A5C", "#8E97B5"]


def _editorial_style(ax):
    ax.set_facecolor(EDITORIAL_PAPER)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    for spine in ("left", "bottom"):
        ax.spines[spine].set_color(EDITORIAL_RULE)
        ax.spines[spine].set_linewidth(0.8)
    ax.tick_params(colors=EDITORIAL_INK, labelsize=9)
    ax.grid(axis="y", color=EDITORIAL_RULE, linestyle="-", linewidth=0.5, alpha=0.6)
    ax.set_axisbelow(True)


def chart_type_distribution(qs: list[Question]) -> bytes:
    counts = pd.Series([q.type for q in qs if q.type]).value_counts()
    if counts.empty:
        return b""
    fig, ax = plt.subplots(figsize=(8, 4.5), facecolor=EDITORIAL_PAPER)
    bars = ax.barh(counts.index[::-1], counts.values[::-1],
                   color=EDITORIAL_INK, edgecolor=EDITORIAL_PAPER, height=0.65)
    for bar, v in zip(bars, counts.values[::-1]):
        ax.text(v + 0.15, bar.get_y() + bar.get_height() / 2, str(v),
                va="center", fontsize=9.5, color=EDITORIAL_INK, family="monospace")
    ax.set_xlim(0, max(counts.values) * 1.18)
    ax.set_title("유형별 출제 비중", loc="left",
                 fontsize=13, color=EDITORIAL_INK, fontweight="bold", pad=14)
    _editorial_style(ax)
    ax.spines["bottom"].set_visible(False)
    ax.tick_params(axis="x", labelsize=0)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, facecolor=EDITORIAL_PAPER, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def chart_difficulty_distribution(qs: list[Question]) -> bytes:
    counts = pd.Series([q.difficulty for q in qs]).value_counts().reindex(DIFFICULTY_LEVELS, fill_value=0)
    fig, ax = plt.subplots(figsize=(7, 3.5), facecolor=EDITORIAL_PAPER)
    bars = ax.bar(counts.index, counts.values,
                  color=[EDITORIAL_ACCENTS[i % len(EDITORIAL_ACCENTS)] for i in range(len(counts))],
                  edgecolor=EDITORIAL_PAPER, linewidth=2, width=0.55)
    for bar, v in zip(bars, counts.values):
        if v > 0:
            ax.text(bar.get_x() + bar.get_width() / 2, v + 0.1, str(int(v)),
                    ha="center", fontsize=10, color=EDITORIAL_INK,
                    family="monospace", fontweight="bold")
    ax.set_title("난이도 분포", loc="left",
                 fontsize=13, color=EDITORIAL_INK, fontweight="bold", pad=14)
    _editorial_style(ax)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, facecolor=EDITORIAL_PAPER, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def chart_killer_map(qs: list[Question]) -> bytes:
    """문항 번호 순으로 난이도 + 킬러 표시 — 시험지 위치 감각."""
    if not qs:
        return b""
    fig, ax = plt.subplots(figsize=(9, 3.0), facecolor=EDITORIAL_PAPER)
    nos = [q.no for q in qs]
    diffs = [q.difficulty_num() for q in qs]
    colors = [EDITORIAL_ACCENTS[3] if q.is_killer else EDITORIAL_INK for q in qs]
    sizes = [110 if q.is_killer else 36 for q in qs]
    ax.scatter(nos, diffs, c=colors, s=sizes, edgecolor=EDITORIAL_PAPER, linewidth=1.2, zorder=3)
    ax.plot(nos, diffs, color=EDITORIAL_INK, alpha=0.18, linewidth=1, zorder=1)
    for q in qs:
        if q.is_killer:
            ax.annotate(f"#{q.no}", (q.no, q.difficulty_num()),
                        xytext=(0, 12), textcoords="offset points",
                        ha="center", fontsize=9, color=EDITORIAL_ACCENTS[3],
                        fontweight="bold")
    ax.set_yticks(list(DIFFICULTY_NUM.values()))
    ax.set_yticklabels(list(DIFFICULTY_NUM.keys()))
    ax.set_xlabel("문항 번호", fontsize=10, color=EDITORIAL_INK)
    ax.set_title("문항 위치별 난이도 & 킬러 분포", loc="left",
                 fontsize=13, color=EDITORIAL_INK, fontweight="bold", pad=14)
    _editorial_style(ax)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, facecolor=EDITORIAL_PAPER, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def chart_radar(qs: list[Question]) -> bytes:
    counts = pd.Series([q.type for q in qs if q.type]).value_counts()
    if len(counts) < 3:
        return b""
    cats = list(counts.index)
    values = list(counts.values)
    angles = np.linspace(0, 2 * np.pi, len(cats), endpoint=False).tolist()
    angles += angles[:1]
    values += values[:1]

    fig, ax = plt.subplots(figsize=(6, 6), facecolor=EDITORIAL_PAPER,
                           subplot_kw=dict(projection="polar"))
    ax.set_facecolor(EDITORIAL_PAPER)
    ax.plot(angles, values, color=EDITORIAL_INK, linewidth=1.6)
    ax.fill(angles, values, color=EDITORIAL_INK, alpha=0.12)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(cats, fontsize=10, color=EDITORIAL_INK)
    ax.set_yticks([])
    ax.spines["polar"].set_color(EDITORIAL_RULE)
    ax.grid(color=EDITORIAL_RULE, linewidth=0.6, alpha=0.7)
    ax.set_title("유형별 레이더", color=EDITORIAL_INK, fontsize=13,
                 fontweight="bold", pad=18)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, facecolor=EDITORIAL_PAPER, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# 6. WORD 보고서
# ─────────────────────────────────────────────────────────────
def _set_east_asia(run, name="맑은 고딕"):
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), name)


def build_word_report(
    meta: ExamMeta,
    qs: list[Question],
    killers: dict,
    summary_text: str,
    charts: dict[str, bytes],
) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21.0), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(1.5)
    section.left_margin = section.right_margin = Cm(1.8)

    # 표지
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("EXAM ANALYSIS REPORT")
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor.from_string("8E97B5")
    r.font.name = "IBM Plex Mono"

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(meta.title or f"{meta.school} {meta.grade} {meta.subject} {meta.exam_type}")
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string("1A1F36")
    _set_east_asia(r)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = " · ".join(filter(None, [meta.school, meta.grade, meta.subject, meta.exam_type, meta.exam_date]))
    r = p.add_run(sub)
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor.from_string("4A5A8C")
    _set_east_asia(r)

    doc.add_paragraph()

    # 메타 KPI
    kpi = doc.add_table(rows=1, cols=4)
    kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    labels = ["총 문항", "총 배점", "시험 시간", "킬러 후보"]
    killer_count = sum(1 for q in qs if q.is_killer)
    values = [f"{meta.total_questions}",
              f"{meta.total_score}점",
              f"{meta.duration_min}분",
              f"{killer_count}문항"]
    for i, (lab, val) in enumerate(zip(labels, values)):
        cell = kpi.cell(0, i)
        cell.text = ""
        para_lab = cell.paragraphs[0]
        para_lab.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rl = para_lab.add_run(lab)
        rl.font.size = Pt(9)
        rl.font.color.rgb = RGBColor.from_string("8E97B5")
        _set_east_asia(rl)
        para_val = cell.add_paragraph()
        para_val.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rv = para_val.add_run(val)
        rv.font.size = Pt(16)
        rv.font.bold = True
        rv.font.color.rgb = RGBColor.from_string("1A1F36")
        rv.font.name = "IBM Plex Mono"

    doc.add_paragraph()

    # 차트들
    def _add_chart(title: str, key: str):
        if not charts.get(key):
            return
        h = doc.add_paragraph()
        rh = h.add_run(title)
        rh.font.size = Pt(13)
        rh.font.bold = True
        rh.font.color.rgb = RGBColor.from_string("1A1F36")
        _set_east_asia(rh)
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture(io.BytesIO(charts[key]), width=Cm(16))

    _add_chart("§1.  유형별 출제 비중", "type")
    _add_chart("§2.  난이도 분포", "difficulty")
    _add_chart("§3.  문항 위치별 난이도 & 킬러", "killer_map")

    # 킬러 문항 리스트
    if killers and killers.get("killers"):
        h = doc.add_paragraph()
        rh = h.add_run("§4.  킬러 문항")
        rh.font.size = Pt(13)
        rh.font.bold = True
        rh.font.color.rgb = RGBColor.from_string("1A1F36")
        _set_east_asia(rh)
        for k in killers["killers"]:
            p = doc.add_paragraph()
            r1 = p.add_run(f"#{k.get('no')}  ")
            r1.font.bold = True
            r1.font.color.rgb = RGBColor.from_string("B5503F")
            r1.font.name = "IBM Plex Mono"
            r2 = p.add_run(k.get("reason", ""))
            r2.font.color.rgb = RGBColor.from_string("1A1F36")
            r2.font.size = Pt(10.5)
            _set_east_asia(r2)

    # 총평
    if summary_text:
        h = doc.add_paragraph()
        rh = h.add_run("§5.  총평")
        rh.font.size = Pt(13)
        rh.font.bold = True
        rh.font.color.rgb = RGBColor.from_string("1A1F36")
        _set_east_asia(rh)
        for para in summary_text.split("\n\n"):
            p = doc.add_paragraph()
            r = p.add_run(para.strip())
            r.font.size = Pt(10.5)
            r.font.color.rgb = RGBColor.from_string("1A1F36")
            _set_east_asia(r)
            p.paragraph_format.line_spacing = 1.6

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# 7. UI — 4번째 탭 본체
# ─────────────────────────────────────────────────────────────
SS_PREFIX = "exam_"  # session_state 네임스페이스


def _ss(key: str, default=None):
    full = SS_PREFIX + key
    if full not in st.session_state:
        st.session_state[full] = default
    return st.session_state[full]


def _set_ss(key: str, value):
    st.session_state[SS_PREFIX + key] = value


def _meta_card(meta: ExamMeta, qs: list[Question]) -> str:
    killer_count = sum(1 for q in qs if q.is_killer)
    return f"""
<div class='kpi-strip'>
  <div class='card'>
    <div class='card-eyebrow'>총 문항</div>
    <div class='card-value'>{meta.total_questions}</div>
    <div class='card-meta'>{meta.subject} · {meta.exam_type}</div>
  </div>
  <div class='card'>
    <div class='card-eyebrow'>총 배점</div>
    <div class='card-value'>{meta.total_score}<span style='font-size:14px;color:var(--text-muted)'> 점</span></div>
    <div class='card-meta'>{meta.duration_min}분</div>
  </div>
  <div class='card'>
    <div class='card-eyebrow'>학교 · 학년</div>
    <div class='card-value' style='font-size:18px'>{meta.grade or '—'}</div>
    <div class='card-meta'>{meta.school or '—'}</div>
  </div>
  <div class='card'>
    <div class='card-eyebrow'>킬러 후보</div>
    <div class='card-value' style='color:var(--killer-fg)'>{killer_count}</div>
    <div class='card-meta'>{meta.exam_date or '날짜 미정'}</div>
  </div>
</div>
"""


def _questions_to_df(qs: list[Question]) -> pd.DataFrame:
    if not qs:
        return pd.DataFrame(columns=["no", "type", "difficulty", "score",
                                     "is_subjective", "is_killer", "scope", "memo"])
    return pd.DataFrame([asdict(q) for q in qs])


def _df_to_questions(df: pd.DataFrame) -> list[Question]:
    qs = []
    for _, row in df.iterrows():
        try:
            qs.append(Question(
                no=int(row.get("no") or 0),
                type=str(row.get("type") or "").strip(),
                difficulty=str(row.get("difficulty") or "중"),
                score=float(row.get("score") or 0),
                is_subjective=bool(row.get("is_subjective") or False),
                is_killer=bool(row.get("is_killer") or False),
                scope=str(row.get("scope") or "").strip(),
                memo=str(row.get("memo") or "").strip(),
            ))
        except Exception:
            continue
    qs.sort(key=lambda x: x.no)
    return qs


def _init_state():
    _ss("meta", None)
    _ss("questions", [])
    _ss("killers", None)
    _ss("summary", "")
    _ss("uploaded_keys", [])
    _ss("subject_hint", "영어")
    _ss("grade_hint", "고2")


def render_sidebar():
    """시험 분석 모드의 사이드바. app.py 의 with st.sidebar: 안에서 호출."""
    _init_state()
    st.markdown("### 분석 설정")
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    st.markdown('<p class="section-label">OCR 힌트</p>', unsafe_allow_html=True)
    st.selectbox(
        "과목", ["영어", "수학", "국어", "기타"],
        key=SS_PREFIX + "subject_hint",
        help="과목을 알려주면 OCR 의 유형 분류 정확도가 높아집니다.",
    )
    st.text_input(
        "학교/학년", key=SS_PREFIX + "grade_hint",
        help="예: 고2, 중3, 고1",
    )

    st.markdown('<p class="section-label">초기화</p>', unsafe_allow_html=True)
    if st.button("분석 결과 초기화", use_container_width=True):
        for k in ("meta", "questions", "killers", "summary", "uploaded_keys"):
            st.session_state[SS_PREFIX + k] = None if k in ("meta", "killers") else ([] if k in ("questions", "uploaded_keys") else "")
        st.rerun()

    meta: ExamMeta | None = _ss("meta")
    if meta is not None:
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('<p class="section-label">현재 시험</p>', unsafe_allow_html=True)
        st.markdown(
            f"<div style='font-size:13px;color:var(--text-body);line-height:1.6'>"
            f"<b>{meta.title or '제목 없음'}</b><br/>"
            f"{meta.school or ''} {meta.grade or ''} {meta.subject or ''}<br/>"
            f"<span style='color:var(--text-muted)'>"
            f"{meta.exam_type} · {meta.total_questions}문항 · {meta.total_score}점"
            f"</span></div>",
            unsafe_allow_html=True,
        )


def render_main(api_key: str):
    """시험 분석 모드의 본문. app.py 의 메인 영역에서 호출."""
    if not api_key:
        st.error("OpenAI API Key 가 필요합니다. .env 또는 Streamlit Secrets 에 OPENAI_API_KEY 를 설정하세요.")
        return

    _init_state()

    # ── Stage 1. 업로드 & OCR ──
    st.markdown('<div class="section-mark">§ 1. 업로드 & OCR</div>', unsafe_allow_html=True)

    files = st.file_uploader(
        "시험지 이미지 업로드 (여러 페이지 동시 가능)",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        key="exam_uploader",
        help="시험지 한 부 전체를 페이지별 이미지로 올리세요. PDF는 미리 이미지로 변환해 주세요.",
    )

    subject_hint = _ss("subject_hint", "영어")
    grade_hint = _ss("grade_hint", "고2")

    if files:
        keys = [f"{f.name}_{f.size}" for f in files]
        if keys != _ss("uploaded_keys"):
            _set_ss("uploaded_keys", keys)
            # 미리보기
        thumb_cols = st.columns(min(len(files), 4))
        for i, f in enumerate(files):
            with thumb_cols[i % 4]:
                st.image(f, use_container_width=True, caption=f.name)

        if st.button("OCR 실행 — 메타정보 & 문항 추출", type="primary", use_container_width=True):
            with st.status("이미지 분석 중...", expanded=True) as status:
                try:
                    img_bytes_list = []
                    for f in files:
                        f.seek(0)
                        img_bytes_list.append(f.read())
                    status.update(label=f"GPT-4o Vision 으로 {len(files)}장 분석 중...")
                    meta, qs = ocr_exam_images(api_key, img_bytes_list, subject_hint, grade_hint)
                    flags = heuristic_killer_flags(meta, qs)
                    for q, flag in zip(qs, flags):
                        q.is_killer = flag
                    _set_ss("meta", meta)
                    _set_ss("questions", qs)
                    _set_ss("killers", None)
                    _set_ss("summary", "")
                    status.update(label=f"완료 — 문항 {len(qs)}개 추출", state="complete")
                except openai.AuthenticationError:
                    status.update(label="API Key 오류", state="error")
                    st.error("OpenAI API Key 가 유효하지 않습니다.")
                except Exception as e:
                    status.update(label="OCR 실패", state="error")
                    st.error(f"분석 중 오류: {e}")

    # ── Stage 2. 메타정보 검토 & 수정 ──
    meta: ExamMeta | None = _ss("meta")
    qs: list[Question] = _ss("questions")

    if meta is None:
        st.markdown('<div class="empty">이미지를 업로드하고 OCR 을 실행하면 여기에 분석 결과가 표시됩니다.</div>',
                    unsafe_allow_html=True)
        return

    st.markdown('<div class="section-mark" style="margin-top:32px">§ 2. 메타정보 검토 & 수정</div>',
                unsafe_allow_html=True)
    st.markdown(_meta_card(meta, qs), unsafe_allow_html=True)

    with st.expander("시험 메타정보 수정", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            meta.title = st.text_input("시험 제목", meta.title, key="m_title")
            meta.school = st.text_input("학교명", meta.school, key="m_school")
            meta.grade = st.text_input("학년", meta.grade, key="m_grade")
        with c2:
            meta.subject = st.text_input("과목", meta.subject, key="m_subject")
            meta.exam_type = st.selectbox("시험 종류",
                                          ["중간고사", "기말고사", "모의고사", "수행평가", "기타"],
                                          index=["중간고사", "기말고사", "모의고사", "수행평가", "기타"].index(meta.exam_type)
                                          if meta.exam_type in ["중간고사", "기말고사", "모의고사", "수행평가", "기타"] else 0,
                                          key="m_type")
            meta.exam_date = st.text_input("시험일자", meta.exam_date, key="m_date")
        with c3:
            meta.duration_min = st.number_input("시험 시간 (분)", 10, 200, meta.duration_min, key="m_dur")
            meta.total_score = st.number_input("총 배점", 10, 200, meta.total_score, key="m_tot")
            meta.notes = st.text_input("출제 범위/메모", meta.notes, key="m_notes")
        _set_ss("meta", meta)

    st.markdown("##### 문항 메타 — 표에서 직접 수정")
    st.caption("난이도·유형·배점·킬러 여부를 자유롭게 고치세요. 변경 즉시 차트에 반영됩니다.")

    df = _questions_to_df(qs)
    edited = st.data_editor(
        df,
        column_config={
            "no": st.column_config.NumberColumn("번호", width=60, format="%d"),
            "type": st.column_config.TextColumn("유형", width=100),
            "difficulty": st.column_config.SelectboxColumn("난이도", options=DIFFICULTY_LEVELS, width=80),
            "score": st.column_config.NumberColumn("배점", min_value=0, max_value=20, step=0.5, width=70),
            "is_subjective": st.column_config.CheckboxColumn("서답형", width=70),
            "is_killer": st.column_config.CheckboxColumn("킬러", width=60),
            "scope": st.column_config.TextColumn("범위", width=100),
            "memo": st.column_config.TextColumn("메모", width=200),
        },
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="exam_q_editor",
    )
    if not edited.equals(df):
        _set_ss("questions", _df_to_questions(edited))
        qs = _ss("questions")

    # 빠른 액션
    qa, qb, qc = st.columns(3)
    with qa:
        if st.button("휴리스틱으로 킬러 자동 표시", use_container_width=True):
            flags = heuristic_killer_flags(meta, qs)
            for q, f in zip(qs, flags):
                q.is_killer = f
            _set_ss("questions", qs)
            st.rerun()
    with qb:
        if st.button("킬러 표시 모두 해제", use_container_width=True):
            for q in qs:
                q.is_killer = False
            _set_ss("questions", qs)
            st.rerun()
    with qc:
        st.download_button(
            "메타정보 JSON 다운로드",
            data=json.dumps(
                {"meta": asdict(meta), "questions": [asdict(q) for q in qs]},
                ensure_ascii=False, indent=2,
            ).encode("utf-8"),
            file_name=f"exam_meta_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json",
            use_container_width=True,
        )

    # ── Stage 3. 분석 (킬러 + 총평 + 차트) ──
    st.markdown('<div class="section-mark" style="margin-top:32px">§ 3. 분석 & 총평</div>',
                unsafe_allow_html=True)

    a1, a2 = st.columns(2)
    with a1:
        if st.button("킬러문항 제안 (GPT)", type="primary", use_container_width=True):
            with st.spinner("변별력 분석 중..."):
                try:
                    res = suggest_killers(api_key, meta, qs)
                    _set_ss("killers", res)
                    # 모델이 짚은 번호를 메타에도 반영
                    suggested = {k.get("no") for k in res.get("killers", [])}
                    for q in qs:
                        if q.no in suggested:
                            q.is_killer = True
                    _set_ss("questions", qs)
                except Exception as e:
                    st.error(f"킬러 분석 실패: {e}")
    with a2:
        if st.button("총평 생성 (GPT)", type="primary", use_container_width=True):
            with st.spinner("총평 작성 중..."):
                try:
                    text = gen_summary(api_key, meta, qs, _ss("killers") or {})
                    _set_ss("summary", text)
                except Exception as e:
                    st.error(f"총평 생성 실패: {e}")

    killers = _ss("killers")
    if killers and killers.get("killers"):
        st.markdown("##### 킬러문항 제안")
        for k in killers["killers"]:
            st.markdown(
                f"<div class='card'>"
                f"<span class='killer-flag'>#{k.get('no')}</span>"
                f"<span style='color:var(--text-heading);font-weight:600'>킬러 후보</span>"
                f"<div style='margin-top:6px;color:var(--text-body);font-size:14px;line-height:1.6'>"
                f"{k.get('reason', '')}</div></div>",
                unsafe_allow_html=True,
            )
        if killers.get("rationale"):
            st.markdown(f"<div class='summary-box'>"
                        f"<div class='lead'>변별 전략 요약</div>{killers['rationale']}"
                        f"</div>", unsafe_allow_html=True)

    summary_text = _ss("summary")
    if summary_text:
        st.markdown("##### 총평")
        st.markdown(f"<div class='summary-box'>"
                    f"<div class='lead'>{meta.title or meta.subject + ' ' + meta.exam_type}</div>"
                    f"{summary_text.replace(chr(10), '<br/>')}"
                    f"</div>", unsafe_allow_html=True)

    # ── 차트 ──
    if qs:
        st.markdown('<div class="section-mark" style="margin-top:32px">§ 4. 시각화</div>',
                    unsafe_allow_html=True)
        charts: dict[str, bytes] = {}
        c1, c2 = st.columns(2)
        with c1:
            charts["type"] = chart_type_distribution(qs)
            if charts["type"]:
                st.image(charts["type"], use_container_width=True)
            charts["difficulty"] = chart_difficulty_distribution(qs)
            if charts["difficulty"]:
                st.image(charts["difficulty"], use_container_width=True)
        with c2:
            charts["killer_map"] = chart_killer_map(qs)
            if charts["killer_map"]:
                st.image(charts["killer_map"], use_container_width=True)
            charts["radar"] = chart_radar(qs)
            if charts["radar"]:
                st.image(charts["radar"], use_container_width=True)

        # ── Word 다운로드 ──
        st.markdown('<div class="section-mark" style="margin-top:32px">§ 5. 보고서 다운로드</div>',
                    unsafe_allow_html=True)
        try:
            word_bytes = build_word_report(meta, qs, killers or {}, summary_text, charts)
            fname_base = re.sub(r'[\\/*?:"<>|]', "_",
                                meta.title or f"{meta.school}_{meta.grade}_{meta.subject}_{meta.exam_type}")
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                "Word 보고서 다운로드 (.docx)",
                data=word_bytes,
                file_name=f"{fname_base}_분석보고서_{ts}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"보고서 생성 중 오류: {e}")
