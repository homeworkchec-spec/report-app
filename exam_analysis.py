"""
시험지 분석 자동화 모듈 — 블로그 발행용

워크플로우:
    1) 이미지 업로드 → GPT-4o Vision OCR
    2) 메타정보 확인/수정 (인라인 편집)
    3) [보고서 생성] 단일 버튼 → 한 호흡으로:
            난도 라벨 → 킬러 자동 표시 → GPT 킬러 추천(기승전결 포함)
            → 블로그 톤 총평 → 차트 → Word + PNG 동시 생성

사람이 쓴 듯한 결과물을 위해:
    · 멀티스테이지 LLM (분석 → 본문 → 폴리시)
    · LEEPIN(최상위학원) 블로그 톤 few-shot
    · AI tells 차단 및 후처리 치환
    · 도입부 패턴 무작위 (시작 문장이 매번 동일하지 않게)
"""

from __future__ import annotations

import base64
import io
import json
import platform
import random
import re
import textwrap
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
from typing import Any

import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

import openai

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


# ─────────────────────────────────────────────────────────────
# Korean font setup — matplotlib + PIL 모두 동일 파일을 씀
# ─────────────────────────────────────────────────────────────
def _setup_korean_font() -> tuple[str | None, str | None]:
    """등록된 폰트 패밀리명과 ttf 경로를 모두 반환."""
    candidates: list[Path] = []
    repo_font = Path(__file__).parent / ".fonts"
    if repo_font.exists():
        candidates += list(repo_font.glob("*.ttf")) + list(repo_font.glob("*.otf"))
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
    else:
        for n in ("NanumGothic.ttf", "NotoSansCJK-Regular.ttc", "malgun.ttf"):
            for base in ("/usr/share/fonts", str(Path.home() / ".fonts")):
                bp = Path(base)
                if not bp.exists():
                    continue
                for p in bp.rglob(n):
                    candidates.append(p)
                    break

    for path in candidates:
        try:
            fm.fontManager.addfont(str(path))
            family = fm.FontProperties(fname=str(path)).get_name()
            matplotlib.rcParams["font.family"] = family
            matplotlib.rcParams["axes.unicode_minus"] = False
            return family, str(path)
        except Exception:
            continue
    matplotlib.rcParams["axes.unicode_minus"] = False
    return None, None


_KO_FONT, _KO_FONT_PATH = _setup_korean_font()


# ─────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────
DIFFICULTY_LEVELS = ["하", "중하", "중", "중상", "상"]
DIFFICULTY_NUM = {"하": 1, "중하": 2, "중": 3, "중상": 4, "상": 5}

VISION_MODEL = "gpt-4o"
TEXT_MODEL = "gpt-4o"

# 출판물 톤의 잉크/페이퍼 — 차트
EDITORIAL_INK = "#1A1F36"
EDITORIAL_PAPER = "#FFFFFF"  # 블로그용은 순백색
EDITORIAL_RULE = "#D4CFC0"
EDITORIAL_ACCENTS = ["#1A1F36", "#6B8E7F", "#B07814", "#B5503F", "#2D3A5C", "#8E97B5",
                     "#6B73B5", "#5E8DA8", "#7B9B8E", "#C18A65"]

# 학원 정보 — 사이드바에서 변경 가능
DEFAULT_ACADEMY = "최상위학원"
DEFAULT_PHONE = "0507-1385-4320"


# ─────────────────────────────────────────────────────────────
# Schema
# ─────────────────────────────────────────────────────────────
@dataclass
class ExamMeta:
    title: str = ""
    school: str = ""
    grade: str = ""
    subject: str = "영어"
    exam_type: str = "중간고사"
    exam_date: str = ""
    duration_min: int = 50
    total_score: int = 100
    total_questions: int = 0
    notes: str = ""


@dataclass
class Question:
    no: int = 0
    type: str = ""
    difficulty: str = "중"
    score: float = 0.0
    is_subjective: bool = False
    is_killer: bool = False
    scope: str = ""
    memo: str = ""

    def difficulty_num(self) -> int:
        return DIFFICULTY_NUM.get(self.difficulty, 3)


# ─────────────────────────────────────────────────────────────
# OCR
# ─────────────────────────────────────────────────────────────
OCR_SYSTEM = (
    "당신은 한국 중·고등학교 시험지를 분석하는 전문가입니다. "
    "주어진 시험지 이미지를 읽고, 시험 메타 정보와 모든 문항의 유형·난이도·배점·시험범위를 추출합니다. "
    "추정이 어려운 값은 빈 문자열로 두되, 문항 번호는 반드시 채우세요. "
    "응답은 반드시 단일 JSON 객체로만 출력하세요."
)

OCR_USER_TEMPLATE = """다음 시험지 이미지를 분석하여 JSON 으로 반환하세요.

[과목 힌트] {subject}
[학교급 힌트] {grade}

[필수 JSON 스키마]
{{
  "exam_meta": {{
    "title": "...",
    "school": "...",
    "grade": "...",
    "subject": "...",
    "exam_type": "...",
    "exam_date": "YYYY-MM-DD",
    "duration_min": 50,
    "total_score": 100,
    "total_questions": 25,
    "notes": "..."
  }},
  "questions": [
    {{
      "no": 1,
      "type": "어법",
      "difficulty": "중",
      "score": 4.0,
      "is_subjective": false,
      "scope": "Lesson 3",
      "memo": ""
    }}
  ]
}}

[난이도 판단 기준]
- 하: 단순 일치/암기/기본 어휘
- 중하: 표준 개념 적용
- 중: 다단계 추론, 복합 어법
- 중상: 변형/응용, 복합 조건
- 상(킬러 후보): 다중 변형·시간소모·고난도 추론

[유형명 규칙(영어)]
일치/불일치, 빈칸, 어법, 어휘, 순서, 삽입, 요약문, 함의, 대의, 지칭,
영영풀이, 조건영작, 서술형, 단답형 — 위 표준명을 우선 사용.

JSON 외 다른 텍스트는 절대 출력하지 마세요."""


def _img_to_data_url(img_bytes: bytes, mime: str = "image/png") -> str:
    return f"data:{mime};base64,{base64.b64encode(img_bytes).decode('ascii')}"


def _normalize_image(file_bytes: bytes, max_side: int = 2000) -> tuple[bytes, str]:
    try:
        im = Image.open(io.BytesIO(file_bytes)).convert("RGB")
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
    text = text.strip()
    m = re.search(r"```(?:json)?\s*(.+?)```", text, re.DOTALL)
    if m:
        text = m.group(1).strip()
    return json.loads(text)


def ocr_exam_images(api_key: str, images: list[bytes],
                    subject_hint: str = "영어", grade_hint: str = "") -> tuple[ExamMeta, list[Question]]:
    client = openai.OpenAI(api_key=api_key)
    content: list[dict[str, Any]] = [
        {"type": "text", "text": OCR_USER_TEMPLATE.format(subject=subject_hint, grade=grade_hint)}
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
    data = _safe_json_loads(resp.choices[0].message.content or "{}")

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
# 분포 / 라벨
# ─────────────────────────────────────────────────────────────
def auto_killer_flags(meta: ExamMeta, qs: list[Question]) -> list[bool]:
    """난이도 상 + 배점 상위 25% + 서답형 복합 → 어려운 문항(킬러) 후보."""
    if not qs:
        return []
    scores = sorted([q.score for q in qs if q.score > 0])
    threshold = scores[int(len(scores) * 0.75)] if scores else 0
    flags = []
    for q in qs:
        is_killer = (
            q.difficulty == "상"
            or (q.difficulty == "중상" and q.score >= threshold)
            or (q.is_subjective and q.score >= threshold and q.difficulty in ("중상", "상"))
        )
        flags.append(is_killer)
    return flags


def derive_difficulty_label(qs: list[Question]) -> str:
    """객관 난이도 라벨 — 평균 + 킬러 가중."""
    if not qs:
        return "중"
    nums = [q.difficulty_num() for q in qs]
    avg = sum(nums) / len(nums)
    killer_ratio = sum(1 for q in qs if q.is_killer) / len(qs)
    score = avg + killer_ratio * 1.2
    if score < 1.7:   return "下"
    if score < 2.4:   return "中下"
    if score < 3.1:   return "中"
    if score < 3.8:   return "中上"
    return "上"


def type_distribution(qs: list[Question]) -> list[tuple[str, int, float]]:
    if not qs:
        return []
    s = pd.Series([q.type for q in qs if q.type])
    counts = s.value_counts()
    total = counts.sum()
    return [(k, int(v), v / total * 100) for k, v in counts.items()]


def scope_distribution(qs: list[Question]) -> list[tuple[str, int, float]]:
    """범위별 분포. scope 비어 있으면 빈 리스트."""
    s = pd.Series([q.scope for q in qs if q.scope.strip()])
    if s.empty:
        return []
    counts = s.value_counts()
    total = counts.sum()
    return [(k, int(v), v / total * 100) for k, v in counts.items()]


def difficulty_distribution(qs: list[Question]) -> list[tuple[str, int, float]]:
    counts = pd.Series([q.difficulty for q in qs]).value_counts().reindex(DIFFICULTY_LEVELS, fill_value=0)
    total = counts.sum() or 1
    return [(k, int(v), v / total * 100) for k, v in counts.items()]


# ─────────────────────────────────────────────────────────────
# Anti-AI humanize layer
# ─────────────────────────────────────────────────────────────
# AI 가 좋아하는 어휘를 자연스러운 한국어로 치환. 한 글에서 한 번씩만.
AI_TELLS: list[tuple[str, list[str]]] = [
    ("다음과 같이",        ["이렇게", "아래처럼"]),
    ("다음과 같습니다",    ["아래와 같습니다", "정리하면 이렇습니다"]),
    ("종합적으로",        ["전반적으로", "전체를 놓고 보면"]),
    ("효과적으로",        ["제대로", "확실하게"]),
    ("효과적인",          ["제대로 된", "실제로 통하는"]),
    ("매우 ",             ["", "꽤 ", "상당히 "]),
    ("살펴보겠습니다",    ["짚어보면 이렇습니다", "정리해 보면"]),
    ("살펴보면",          ["짚어보면", "들여다보면"]),
    ("이를 통해",         ["이로써", "이렇게 해서"]),
    ("이러한 ",           ["이런 ", ""]),
    ("저희는",            ["", "이번 시험은"]),
    ("여러분",            [""]),
    ("결론적으로",        ["정리하면", "요약하면"]),
    ("나타났습니다",      ["드러났습니다", "확인됐습니다"]),
    ("확인할 수 있습니다", ["확인됩니다", "분명합니다"]),
    ("필요합니다.",        ["필수적입니다.", "요구됩니다."]),
    ("중요합니다.",        ["관건입니다.", "핵심입니다."]),
    ("출제되었습니다",    ["출제됐습니다", "출제된 셈입니다"]),
]

# 도입부 패턴 — 매번 다른 시작점
OPENING_VARIANTS = [
    "이번 시험은",
    "전반적으로 보면 이번 시험은",
    "결과만 놓고 보면",
    "체감 난이도부터 짚어보면",
    "큰 틀에서 이번 시험은",
]


def humanize_text(text: str, seed: int | None = None) -> str:
    """LLM 결과 후처리 — AI 흔적을 자연스럽게."""
    rng = random.Random(seed)
    out = text

    # 1) 단어/구 치환 (한 패턴 첫 등장만)
    for ai_phrase, alts in AI_TELLS:
        if ai_phrase in out:
            out = out.replace(ai_phrase, rng.choice(alts), 1)

    # 2) 너무 매끄러운 "~입니다. ~입니다." 연속 줄이기 — 가끔 "~죠." 로
    # 보수적으로 적용: 연속 3회 이상이면 한 번 변형
    sentences = re.split(r"(?<=[.!?])\s+", out)
    if len(sentences) >= 3:
        ips = [i for i, s in enumerate(sentences) if s.endswith("입니다.")]
        # 너무 자주 바꾸면 어색 → 25% 확률로만 끝 한 곳 변환
        if ips and rng.random() < 0.25:
            i = rng.choice(ips)
            sentences[i] = sentences[i][:-4] + "라 할 수 있습니다."
        out = " ".join(sentences)

    # 3) AI 가 자주 쓰는 뻔한 표현 일부 제거
    out = re.sub(r"분석한 결과,?\s*", "", out)
    out = re.sub(r"^\s*결과적으로,?\s*", "", out, flags=re.MULTILINE)

    # 4) 공백 정리
    out = re.sub(r"[ \t]+", " ", out)
    out = re.sub(r"\n{3,}", "\n\n", out)

    return out.strip()


# ─────────────────────────────────────────────────────────────
# 블로그 톤 분석문 — LEEPIN(최상위학원) few-shot
# ─────────────────────────────────────────────────────────────
LEEPIN_SAMPLE = """이순신고 1학년 2025년 2학기 1회고사 영어 시험 분석

총평 (난도: 中下)

이번 이순신고 1학년 영어 시험은 객관적 난이도 '중하' 수준으로, 기본적인 학습이 충실히 이루어진 학생이라면 무난히 해결할 수 있는 문항들로 구성되었습니다. 다만, 일부 문항의 선택지가 교묘하게 구성되고 초반부터 변형 문제가 출제되어 학생들의 체감 난이도는 '중' 수준까지 상승했을 것으로 판단됩니다. 서술형 난도가 하락한 점은 주목할 만한 변화이며, 이는 객관식에서의 단 한 번의 실수가 등급에 결정적인 영향을 미치는 구조로 작용했습니다.

📌 주요 특징 및 분석

1. 특정 범위 편중 없는 균형 잡힌 출제
이번 시험의 가장 큰 특징은 교과서 5과와 6과에서 각각 36%씩 동일한 비중으로 출제되었다는 점입니다. 모의고사 역시 28%를 차지하여, 전 범위를 고르게 학습했는지를 평가하려는 출제 의도가 명확합니다. 어느 한 부분도 소홀히 할 수 없는, 학습의 균형이 요구되는 시험이었습니다.

2. 변별력의 핵심, '일치/불일치'와 '어휘'
객관식 문항의 성패는 '일치/불일치(6문항)'와 '어휘(5문항)' 유형에서 갈렸습니다. 이 두 유형이 전체의 44%를 차지하며, 지문 내용을 얼마나 세밀하고 정확하게 이해하고 있는지를 집중적으로 평가했습니다. 이는 고난도 어법보다 꼼꼼한 지문 분석 능력의 중요성이 부각되었음을 의미합니다.

3. 쉬워진 서답형, 기회이자 위기
서답형 3문항 모두 평이한 단답형으로 출제되어 상위권 학생들에게는 변별력을 갖지 못했을 것으로 분석됩니다. 이는 서답형이 점수를 얻는 구간이 아닌, 실점하지 않고 반드시 지켜내야 하는 구간임을 의미합니다. 조건에 맞춘 정확한 단어 사용, 사소한 철자 오류 등 기본적인 부분에서의 실수가 등급에 치명적인 영향을 미쳤을 것입니다.

💡 학습 전략 제언

이러한 출제 경향을 고려할 때, 안정적인 상위 등급 확보를 위한 학습 전략은 다음과 같이 수립되어야 합니다.

첫째, 전 범위에 대한 균형 있는 학습이 필수적입니다. 특정 단원에 집중하기보다, 교과서와 모의고사 등 시험 범위 전체를 반복 학습하여 모든 지문에 대한 숙지도를 동일한 수준으로 유지해야 합니다.

둘째, 꼼꼼함을 바탕으로 한 지문 완전 정복에 집중해야 합니다. 단순히 지문을 해석하는 수준을 넘어, 내용의 순서나 세부 정보까지 정확히 파악하고 근거를 찾아가며 푸는 훈련이 요구됩니다.

셋째, 만점을 목표로 한 서답형 실수 방지 훈련이 필요합니다. 교과서의 주요 구문과 핵심 어휘는 의미뿐만 아니라 정확한 철자까지 완벽하게 쓰는 연습을 반복하여 실전에서 사소한 실수로 감점당하는 일을 원천적으로 차단해야 합니다.

현재의 평이한 기조에 안주하기보다는, 기본기를 더욱 공고히 다지면서 고난도 유형에 대한 대비를 병행하여 어떠한 변화에도 흔들리지 않는 실력을 갖추는 것이 현명한 장기적 전략입니다."""


BLOG_SYSTEM = """당신은 학원 원장 LEEPIN 입니다. 영어 입시 전문이며, 중·고등 내신 시험을 직접 분석해 블로그 칼럼을 씁니다.

[당신의 글쓰기 스타일]
- 학부모와 학생이 함께 읽는 블로그 톤. 단단한 문어체이지만 "갈렸습니다", "치명적인 영향", "변별의 핵심", "기회이자 위기"처럼 한국 입시 지도자 특유의 표현을 자유롭게 사용
- 객관 데이터(분포·비중·문항수)와 전문가의 판단을 한 단락 안에서 자연스럽게 섞음
- 단순 나열이 아닌 의미 부여: 이 분포가 어떤 학습 신호인지, 다음 시험에서 어떤 행동이 필요한지
- 짧은 문장과 긴 문장을 섞어 리듬을 만듦 — 같은 길이 문장이 4개 이상 연속되면 어색함

[금지 표현 — 절대 사용하지 말 것]
"다음과 같이", "분석한 결과", "살펴보겠습니다", "여러분", "효과적으로",
"종합적으로", "다양한", "이를 통해", "이러한 측면에서", "결론적으로 말씀드리면",
"~을 통해서", "~에 대한 분석", "체계적으로", "활용하여",
"~인 점은 인상적입니다", "~을 알 수 있습니다", "~로 사료됩니다"
이모지 남발 금지. 본문에는 이모지를 쓰지 말고, 섹션 제목(📌, 💡)에만 1개씩.

[반드시 지킬 출력 구조]
{title}

총평 (난도: {difficulty})
[4~6 문장. 객관 난이도와 체감 난이도를 구분해 진단. 출제 의도 한마디. 다음 시험 학습 방향을 한 줄 시사.]

📌 주요 특징 및 분석

1. {특징1 한 줄 헤드라인}
[2~3 문장. 데이터+의미.]

2. {특징2 한 줄 헤드라인}
[2~3 문장.]

3. {특징3 한 줄 헤드라인}
[2~3 문장.]

💡 학습 전략 제언

이러한 출제 경향을 고려할 때, 안정적인 상위 등급 확보를 위한 학습 전략은 다음과 같이 수립되어야 합니다.

첫째, [전략1 명령형 한 문장]. [부연 1~2 문장.]
둘째, [전략2 명령형 한 문장]. [부연 1~2 문장.]
셋째, [전략3 명령형 한 문장]. [부연 1~2 문장.]

[마무리 한 단락 — 장기 관점의 학습 태도. "안주하기보다는" 같은 능동적 끝맺음.]"""


def gen_blog_analysis(api_key: str, meta: ExamMeta, qs: list[Question],
                      academy: str = DEFAULT_ACADEMY, phone: str = DEFAULT_PHONE,
                      seed: int | None = None) -> str:
    """블로그용 분석문 — 두 단계 호출: 1) 초안 2) 폴리시."""
    if not qs:
        return ""
    rng = random.Random(seed)
    client = openai.OpenAI(api_key=api_key)

    diff_label = derive_difficulty_label(qs)
    type_dist = type_distribution(qs)
    scope_dist = scope_distribution(qs)
    diff_dist = difficulty_distribution(qs)
    killer_qs = [q for q in qs if q.is_killer]

    title = (meta.title or
             " ".join(filter(None, [meta.school, meta.grade,
                                    f"{datetime.now().year}년" if not meta.exam_date else meta.exam_date[:4] + "년",
                                    meta.exam_type, "영어 시험 분석"])))

    type_lines = "\n".join(f"- {t}: {n}문항 ({p:.1f}%)" for t, n, p in type_dist)
    scope_lines = "\n".join(f"- {s}: {n}문항 ({p:.1f}%)" for s, n, p in scope_dist) or "(범위 정보 미제공)"
    diff_lines = "\n".join(f"- {d}: {n}문항 ({p:.1f}%)" for d, n, p in diff_dist if n > 0)
    killer_lines = "\n".join(f"- {q.no}번 ({q.type}, {q.difficulty}, {q.score}점)" for q in killer_qs) or "- 명시적 킬러 없음"

    user = f"""[참고: 당신이 과거에 쓴 글의 톤]
{LEEPIN_SAMPLE}

────────────────
[이번 시험 데이터]
제목: {title}
난도 라벨: {diff_label}
총 {meta.total_questions}문항 · {meta.total_score}점 · {meta.duration_min}분

[유형별 분포]
{type_lines}

[범위별 분포]
{scope_lines}

[난이도 분포]
{diff_lines}

[어려운(킬러) 문항]
{killer_lines}

────────────────
지시:
- 위 데이터로 블로그 칼럼을 작성하세요.
- 시작 단락은 "{rng.choice(OPENING_VARIANTS)}" 형태로 자연스럽게 시작하되, 너무 정형화되지 않게 변형해도 좋습니다.
- 데이터(분포 비율, 문항수)는 본문 안에 자연스럽게 녹여 쓰세요. 별도의 표나 불릿 나열은 만들지 마세요(상단 문항 구성은 본문 외 자동 생성).
- 마지막 줄에 "수강 문의: {academy} ☎️ {phone}" 한 줄을 그대로 붙이세요."""

    # Stage 1 — 초안
    draft = client.chat.completions.create(
        model=TEXT_MODEL,
        messages=[
            {"role": "system", "content": BLOG_SYSTEM},
            {"role": "user", "content": user},
        ],
        temperature=0.85,
        max_tokens=1800,
    ).choices[0].message.content or ""

    # Stage 2 — 폴리시: AI 흔적 제거 지시
    polish_sys = (
        "당신은 한국 입시 학원 원장의 글을 자연스럽게 다듬는 편집자입니다. "
        "AI 가 쓴 듯한 매끄러움을 줄이고, 사람이 손으로 쓴 듯한 호흡과 억양을 살리세요. "
        "아래 금지 표현은 모두 자연스러운 표현으로 대체하세요: "
        "'다음과 같이', '분석한 결과', '살펴보겠습니다', '효과적으로', '종합적으로', "
        "'다양한', '이를 통해', '체계적으로', '활용하여'. "
        "원문의 의미와 구조(섹션 헤딩, 첫째/둘째/셋째)는 보존하되, 어휘와 리듬만 자연스럽게 다듬으세요. "
        "수정된 전체 글만 그대로 출력하세요."
    )
    polished = client.chat.completions.create(
        model=TEXT_MODEL,
        messages=[
            {"role": "system", "content": polish_sys},
            {"role": "user", "content": draft},
        ],
        temperature=0.6,
        max_tokens=1800,
    ).choices[0].message.content or draft

    # Stage 3 — Python 후처리
    return humanize_text(polished, seed=seed)


# ─────────────────────────────────────────────────────────────
# 킬러문항 기승전결
# ─────────────────────────────────────────────────────────────
KSSJG_SYSTEM = """당신은 입시 영어 분석 전문가입니다.
주어진 어려운 문항(킬러)에 대해 '기승전결' 4단계로 풀어줍니다.

각 단계 정의:
- 기 (출제 의도): 어떤 능력을 묻는지 한 줄로
- 승 (난이도 핵심): 왜 어려운가. 함정·복합 요소 한 줄로
- 전 (학생 오답 패턴): 흔히 빠지는 사고 흐름 한 줄로
- 결 (대비 학습법): 다음 시험 대비 구체적 액션 한 줄로

문체: 학원 원장이 학부모에게 설명하듯, 단정하지만 따뜻한 한국어. "~합니다" 종결.
절대 금지: "다음과 같이", "다양한", "효과적으로", "이를 통해", "여러분"
JSON 으로만 응답하세요."""


def gen_killer_kssjg(api_key: str, meta: ExamMeta, qs: list[Question]) -> list[dict]:
    """어려운 문항별 기승전결. 결과: [{no, type, ki, seung, jeon, gyeol, headline}]."""
    killers = [q for q in qs if q.is_killer]
    if not killers:
        return []

    client = openai.OpenAI(api_key=api_key)
    table = "\n".join(
        f"- {q.no}번 | {q.type} | {q.difficulty} | {q.score}점 | "
        f"서답형={q.is_subjective} | 범위={q.scope or '미지정'} | 메모={q.memo or '-'}"
        for q in killers
    )
    user = f"""[시험 정보] {meta.school} {meta.grade} {meta.subject} {meta.exam_type}, 총 {meta.total_questions}문항

[어려운 문항 리스트]
{table}

요청: 각 문항별로 기·승·전·결 + 한 줄 헤드라인을 JSON 으로 반환하세요.
{{
  "items": [
    {{
      "no": 4,
      "type": "조건영작",
      "headline": "본문 변형 + 어법 조건이 결합된 최고난도 서술형",
      "ki":    "본문의 핵심 구문을 변형해 정확히 재구성할 수 있는지를 묻는 문항입니다.",
      "seung": "어법 조건 3가지가 동시에 걸려 있어 한 곳만 어긋나도 감점이 발생합니다.",
      "jeon": "구문은 맞췄지만 어순·시제·일치 중 한 곳을 놓쳐 부분 감점되는 학생이 다수입니다.",
      "gyeol": "교과서 핵심 구문을 영작 카드로 정리해 어법 조건과 함께 매일 5문항씩 훈련하시면 됩니다."
    }}
  ]
}}

JSON 외 텍스트는 출력하지 마세요."""

    resp = client.chat.completions.create(
        model=TEXT_MODEL,
        messages=[
            {"role": "system", "content": KSSJG_SYSTEM},
            {"role": "user", "content": user},
        ],
        temperature=0.55,
        max_tokens=1400,
        response_format={"type": "json_object"},
    )
    data = _safe_json_loads(resp.choices[0].message.content or "{}")
    items = data.get("items", []) or []
    # 후처리
    cleaned = []
    for it in items:
        cleaned.append({
            "no": it.get("no"),
            "type": it.get("type", ""),
            "headline": humanize_text(it.get("headline", "")),
            "ki": humanize_text(it.get("ki", "")),
            "seung": humanize_text(it.get("seung", "")),
            "jeon": humanize_text(it.get("jeon", "")),
            "gyeol": humanize_text(it.get("gyeol", "")),
        })
    return cleaned


# ─────────────────────────────────────────────────────────────
# 차트 — 블로그용 (흰 배경, 잉크 톤)
# ─────────────────────────────────────────────────────────────
def _editorial_style(ax):
    ax.set_facecolor(EDITORIAL_PAPER)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color(EDITORIAL_RULE)
        ax.spines[s].set_linewidth(0.8)
    ax.tick_params(colors=EDITORIAL_INK, labelsize=10)
    ax.grid(axis="y", color=EDITORIAL_RULE, linestyle="-", linewidth=0.5, alpha=0.6)
    ax.set_axisbelow(True)


def chart_type_distribution(qs: list[Question]) -> bytes:
    dist = type_distribution(qs)
    if not dist:
        return b""
    labels, counts, pcts = zip(*dist)
    fig, ax = plt.subplots(figsize=(9, max(4.5, 0.55 * len(labels))), facecolor=EDITORIAL_PAPER)
    bars = ax.barh(labels[::-1], counts[::-1],
                   color=[EDITORIAL_ACCENTS[i % len(EDITORIAL_ACCENTS)] for i in range(len(labels))][::-1],
                   edgecolor=EDITORIAL_PAPER, height=0.7)
    for bar, n, p in zip(bars, counts[::-1], pcts[::-1]):
        ax.text(n + max(counts) * 0.02, bar.get_y() + bar.get_height() / 2,
                f"{n}문항 ({p:.1f}%)", va="center", fontsize=10.5,
                color=EDITORIAL_INK)
    ax.set_xlim(0, max(counts) * 1.32)
    ax.set_title("유형별 출제 비중", loc="left", fontsize=14,
                 color=EDITORIAL_INK, fontweight="bold", pad=14)
    _editorial_style(ax)
    ax.tick_params(axis="x", labelsize=0)
    ax.spines["bottom"].set_visible(False)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, facecolor=EDITORIAL_PAPER, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def chart_scope_distribution(qs: list[Question]) -> bytes:
    dist = scope_distribution(qs)
    if not dist:
        return b""
    labels, counts, pcts = zip(*dist)
    fig, ax = plt.subplots(figsize=(8, max(3.5, 0.55 * len(labels))), facecolor=EDITORIAL_PAPER)
    bars = ax.barh(labels[::-1], counts[::-1],
                   color=[EDITORIAL_ACCENTS[(i + 2) % len(EDITORIAL_ACCENTS)] for i in range(len(labels))][::-1],
                   edgecolor=EDITORIAL_PAPER, height=0.7)
    for bar, n, p in zip(bars, counts[::-1], pcts[::-1]):
        ax.text(n + max(counts) * 0.02, bar.get_y() + bar.get_height() / 2,
                f"{n}문항 ({p:.1f}%)", va="center", fontsize=10.5,
                color=EDITORIAL_INK)
    ax.set_xlim(0, max(counts) * 1.32)
    ax.set_title("범위별 문항 분포", loc="left", fontsize=14,
                 color=EDITORIAL_INK, fontweight="bold", pad=14)
    _editorial_style(ax)
    ax.tick_params(axis="x", labelsize=0)
    ax.spines["bottom"].set_visible(False)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, facecolor=EDITORIAL_PAPER, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def chart_difficulty_distribution(qs: list[Question]) -> bytes:
    dist = difficulty_distribution(qs)
    fig, ax = plt.subplots(figsize=(8, 3.6), facecolor=EDITORIAL_PAPER)
    labels, counts, pcts = zip(*dist)
    bars = ax.bar(labels, counts,
                  color=[EDITORIAL_ACCENTS[i % len(EDITORIAL_ACCENTS)] for i in range(len(labels))],
                  edgecolor=EDITORIAL_PAPER, linewidth=2, width=0.55)
    for bar, n, p in zip(bars, counts, pcts):
        if n > 0:
            ax.text(bar.get_x() + bar.get_width() / 2, n + max(counts) * 0.02,
                    f"{n}\n({p:.1f}%)", ha="center", fontsize=9.5,
                    color=EDITORIAL_INK, fontweight="bold")
    ax.set_title("난이도 분포", loc="left", fontsize=14,
                 color=EDITORIAL_INK, fontweight="bold", pad=14)
    _editorial_style(ax)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, facecolor=EDITORIAL_PAPER, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def chart_killer_map(qs: list[Question]) -> bytes:
    if not qs:
        return b""
    fig, ax = plt.subplots(figsize=(9, 3.0), facecolor=EDITORIAL_PAPER)
    nos = [q.no for q in qs]
    ds = [q.difficulty_num() for q in qs]
    colors = [EDITORIAL_ACCENTS[3] if q.is_killer else EDITORIAL_INK for q in qs]
    sizes = [120 if q.is_killer else 36 for q in qs]
    ax.scatter(nos, ds, c=colors, s=sizes, edgecolor=EDITORIAL_PAPER, linewidth=1.2, zorder=3)
    ax.plot(nos, ds, color=EDITORIAL_INK, alpha=0.16, linewidth=1, zorder=1)
    for q in qs:
        if q.is_killer:
            ax.annotate(f"#{q.no}", (q.no, q.difficulty_num()),
                        xytext=(0, 12), textcoords="offset points",
                        ha="center", fontsize=9.5, color=EDITORIAL_ACCENTS[3], fontweight="bold")
    ax.set_yticks(list(DIFFICULTY_NUM.values()))
    ax.set_yticklabels(list(DIFFICULTY_NUM.keys()))
    ax.set_xlabel("문항 번호", fontsize=10, color=EDITORIAL_INK)
    ax.set_title("문항 위치별 난이도 & 어려운 문항 분포", loc="left",
                 fontsize=14, color=EDITORIAL_INK, fontweight="bold", pad=14)
    _editorial_style(ax)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, facecolor=EDITORIAL_PAPER, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# 블로그용 단일 PNG 이미지 (PIL)
# ─────────────────────────────────────────────────────────────
def _font(size: int, bold: bool = False) -> ImageFont.ImageFont:
    """한글 폰트. bold 는 같은 파일을 그대로 두께 처리(matplotlib 와 호환).
    실제 굵은 ttf 가 있으면 우선 사용."""
    path = _KO_FONT_PATH
    bold_path = None
    if path:
        p = Path(path)
        for cand in (p.parent / "malgunbd.ttf", p.parent / "NanumGothicBold.ttf"):
            if cand.exists():
                bold_path = str(cand)
                break
    use = (bold_path if bold and bold_path else path)
    if use:
        try:
            return ImageFont.truetype(use, size)
        except Exception:
            pass
    return ImageFont.load_default()


def _wrap(text: str, font: ImageFont.ImageFont, max_w: int) -> list[str]:
    """한국어 줄바꿈 — 글자 단위로 측정하여 max_w 안에 들어가게."""
    out = []
    for paragraph in text.split("\n"):
        if not paragraph.strip():
            out.append("")
            continue
        line = ""
        for ch in paragraph:
            test = line + ch
            w = font.getbbox(test)[2]
            if w > max_w and line:
                out.append(line)
                line = ch
            else:
                line = test
        if line:
            out.append(line)
    return out


def _draw_paragraph(draw: ImageDraw.ImageDraw, text: str, x: int, y: int, max_w: int,
                    font: ImageFont.ImageFont, color=(26, 31, 54), line_height: float = 1.65) -> int:
    """줄바꿈된 문단을 그리고 끝 y 반환."""
    if not text.strip():
        return y
    lines = _wrap(text, font, max_w)
    line_h = int(font.size * line_height)
    for ln in lines:
        draw.text((x, y), ln, font=font, fill=color)
        y += line_h
    return y


def render_blog_image(meta: ExamMeta, qs: list[Question], analysis_text: str,
                      kssjg: list[dict], charts: dict[str, bytes],
                      academy: str = DEFAULT_ACADEMY, phone: str = DEFAULT_PHONE) -> bytes:
    """단일 PNG (1080 폭, 흰 배경) — 네이버 블로그 본문에 그대로 붙여넣기 가능."""
    W = 1080
    PAD = 60
    INNER = W - PAD * 2
    BG = (255, 255, 255)
    INK = (26, 31, 54)
    MUTED = (110, 115, 130)
    ACCENT = (45, 58, 92)
    RULE = (212, 207, 192)

    F_TITLE = _font(38, bold=True)
    F_H1    = _font(26, bold=True)
    F_H2    = _font(20, bold=True)
    F_BODY  = _font(17)
    F_META  = _font(14)
    F_KICK  = _font(15)

    # ── 1차 패스: 전체 높이 계산을 위해 더미 캔버스에 측정 ──
    dummy = Image.new("RGB", (W, 100), BG)
    draw0 = ImageDraw.Draw(dummy)

    def measure_paragraph(text: str, font, line_h_mult=1.65) -> int:
        if not text.strip():
            return 0
        lines = _wrap(text, font, INNER)
        return int(font.size * line_h_mult) * len(lines)

    # 본문 길이 추정으로 캔버스 높이 산출
    H = PAD * 2
    H += int(F_TITLE.size * 1.4) + 30                # 타이틀
    H += int(F_META.size * 1.6) + 28                  # 메타 줄
    H += 1 + 24                                       # 룰
    H += int(F_H1.size * 1.4) + 8                     # 총평 헤드
    # 분석 본문 높이 (대략)
    H += measure_paragraph(analysis_text, F_BODY) + 40
    # 차트
    for key in ("type", "scope", "difficulty", "killer_map"):
        if charts.get(key):
            H += 600 + 30
    # 기승전결
    if kssjg:
        H += int(F_H1.size * 1.4) + 16
        for it in kssjg:
            H += int(F_H2.size * 1.4) + 6
            for label in ("ki", "seung", "jeon", "gyeol"):
                H += measure_paragraph(it.get(label, ""), F_BODY) + 6
            H += 20
    # 푸터
    H += int(F_KICK.size * 1.6) + 20

    # ── 본 캔버스 ──
    canvas = Image.new("RGB", (W, max(H, 1200)), BG)
    draw = ImageDraw.Draw(canvas)
    y = PAD

    # 타이틀
    title = (meta.title or
             " ".join(filter(None, [meta.school, meta.grade,
                                    meta.exam_date[:7] if meta.exam_date else "",
                                    meta.exam_type, "영어 시험 분석"])))
    title_lines = _wrap(title, F_TITLE, INNER)
    for ln in title_lines:
        draw.text((PAD, y), ln, font=F_TITLE, fill=INK)
        y += int(F_TITLE.size * 1.2)
    y += 10

    # 메타 한 줄
    meta_str = " · ".join(filter(None, [
        meta.school or "", meta.grade or "", f"총 {meta.total_questions}문항",
        f"{meta.total_score}점", f"{meta.duration_min}분",
        f"난도 {derive_difficulty_label(qs)}",
    ]))
    draw.text((PAD, y), meta_str, font=F_META, fill=MUTED)
    y += int(F_META.size * 1.4) + 18

    # 룰
    draw.line([(PAD, y), (W - PAD, y)], fill=ACCENT, width=2)
    y += 26

    # 본문 — 분석 텍스트 (총평/특징/전략 포함)
    # 섹션 헤딩 분리 렌더 — "총평", "📌 주요 특징", "💡 학습 전략" 만 굵은 H1, 나머지 본문
    section_re = re.compile(
        r"^(총평\s*\(난도:.+?\)|📌\s*주요 특징.*|💡\s*학습 전략.*)$",
        re.MULTILINE,
    )
    blocks = []
    last = 0
    for m in section_re.finditer(analysis_text):
        if m.start() > last:
            blocks.append(("body", analysis_text[last:m.start()].strip()))
        blocks.append(("head", m.group(1).strip()))
        last = m.end()
    if last < len(analysis_text):
        blocks.append(("body", analysis_text[last:].strip()))

    for kind, text in blocks:
        if not text:
            continue
        if kind == "head":
            y += 16
            for ln in _wrap(text, F_H1, INNER):
                draw.text((PAD, y), ln, font=F_H1, fill=INK)
                y += int(F_H1.size * 1.25)
            y += 8
        else:
            y = _draw_paragraph(draw, text, PAD, y, INNER, F_BODY, color=INK)
            y += 14

    # 차트들
    chart_titles = [
        ("type",       "유형별 출제 비중"),
        ("scope",      "범위별 문항 분포"),
        ("difficulty", "난이도 분포"),
        ("killer_map", "문항 위치별 난이도 & 어려운 문항"),
    ]
    for key, _ in chart_titles:
        if not charts.get(key):
            continue
        chart_im = Image.open(io.BytesIO(charts[key])).convert("RGB")
        # 폭 INNER 에 맞게 리사이즈
        ratio = INNER / chart_im.width
        new_h = int(chart_im.height * ratio)
        chart_im = chart_im.resize((INNER, new_h), Image.LANCZOS)
        canvas.paste(chart_im, (PAD, y))
        y += new_h + 24

    # 기승전결
    if kssjg:
        y += 12
        draw.text((PAD, y), "🎯 어려운 문항 기승전결", font=F_H1, fill=INK)
        y += int(F_H1.size * 1.4) + 8
        for it in kssjg:
            head = f"{it.get('no')}번 · {it.get('type', '')} — {it.get('headline', '')}"
            for ln in _wrap(head, F_H2, INNER):
                draw.text((PAD, y), ln, font=F_H2, fill=ACCENT)
                y += int(F_H2.size * 1.25)
            y += 4
            for label, key in (("기 · 출제 의도", "ki"),
                               ("승 · 난이도 핵심", "seung"),
                               ("전 · 학생 오답 패턴", "jeon"),
                               ("결 · 대비 학습법", "gyeol")):
                txt = it.get(key, "").strip()
                if not txt:
                    continue
                draw.text((PAD, y), label, font=F_KICK, fill=ACCENT)
                y += int(F_KICK.size * 1.4)
                y = _draw_paragraph(draw, txt, PAD + 18, y, INNER - 18, F_BODY, color=INK)
                y += 6
            y += 14

    # 푸터
    y += 16
    draw.line([(PAD, y), (W - PAD, y)], fill=RULE, width=1)
    y += 18
    cta = f"수강 문의: {academy} ☎️ {phone}"
    draw.text((PAD, y), cta, font=F_H2, fill=INK)

    # 캔버스 자르기 — 실제 사용 높이까지만
    final_h = y + int(F_H2.size * 1.4) + PAD
    if final_h < canvas.height:
        canvas = canvas.crop((0, 0, W, final_h))

    out = io.BytesIO()
    canvas.save(out, format="PNG", optimize=True)
    return out.getvalue()


# ─────────────────────────────────────────────────────────────
# Word 보고서 — 블로그와 동일한 텍스트 흐름
# ─────────────────────────────────────────────────────────────
def _set_east_asia(run, name="맑은 고딕"):
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), name)


def _add_run(p, text, *, size=10.5, bold=False, color="1A1F36", mono=False):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = RGBColor.from_string(color)
    if mono:
        r.font.name = "IBM Plex Mono"
    else:
        r.font.name = "맑은 고딕"
        _set_east_asia(r)
    return r


def build_word_report(meta: ExamMeta, qs: list[Question], analysis_text: str,
                      kssjg: list[dict], charts: dict[str, bytes],
                      academy: str = DEFAULT_ACADEMY, phone: str = DEFAULT_PHONE) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21.0), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(1.6)
    section.left_margin = section.right_margin = Cm(1.8)

    # 표지
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    title = (meta.title or
             " ".join(filter(None, [meta.school, meta.grade,
                                    meta.exam_date[:7] if meta.exam_date else "",
                                    meta.exam_type, "영어 시험 분석"])))
    _add_run(p, title, size=22, bold=True, color="1A1F36")

    p = doc.add_paragraph()
    meta_str = " · ".join(filter(None, [
        meta.school or "", meta.grade or "",
        f"총 {meta.total_questions}문항", f"{meta.total_score}점", f"{meta.duration_min}분",
        f"난도 {derive_difficulty_label(qs)}",
    ]))
    _add_run(p, meta_str, size=10, color="6E7382")

    doc.add_paragraph()  # spacer

    # 본문 — analysis_text 를 단락별로 풀어 넣고 섹션 헤드는 굵게
    section_re = re.compile(
        r"^(총평\s*\(난도:.+?\)|📌\s*주요 특징.*|💡\s*학습 전략.*)$",
        re.MULTILINE,
    )
    parts = []
    last = 0
    for m in section_re.finditer(analysis_text):
        if m.start() > last:
            parts.append(("body", analysis_text[last:m.start()].strip()))
        parts.append(("head", m.group(1).strip()))
        last = m.end()
    if last < len(analysis_text):
        parts.append(("body", analysis_text[last:].strip()))

    for kind, text in parts:
        if not text:
            continue
        if kind == "head":
            doc.add_paragraph()
            p = doc.add_paragraph()
            _add_run(p, text, size=14, bold=True, color="1A1F36")
        else:
            for para in text.split("\n\n"):
                para = para.strip()
                if not para:
                    continue
                p = doc.add_paragraph()
                _add_run(p, para, size=10.5, color="1A1F36")
                p.paragraph_format.line_spacing = 1.65
                p.paragraph_format.space_after = Pt(8)

    # 차트
    chart_order = [
        ("type",       "유형별 출제 비중"),
        ("scope",      "범위별 문항 분포"),
        ("difficulty", "난이도 분포"),
        ("killer_map", "문항 위치별 난이도 & 어려운 문항"),
    ]
    for key, label in chart_order:
        if not charts.get(key):
            continue
        doc.add_paragraph()
        p = doc.add_paragraph()
        _add_run(p, label, size=12.5, bold=True, color="1A1F36")
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture(io.BytesIO(charts[key]), width=Cm(16))

    # 기승전결
    if kssjg:
        doc.add_paragraph()
        p = doc.add_paragraph()
        _add_run(p, "🎯 어려운 문항 기승전결", size=14, bold=True, color="1A1F36")
        for it in kssjg:
            p = doc.add_paragraph()
            _add_run(p, f"{it.get('no')}번 · {it.get('type', '')} ", size=11.5, bold=True, color="2D3A5C", mono=True)
            _add_run(p, f"— {it.get('headline', '')}", size=11.5, bold=True, color="2D3A5C")
            for label, key in (("기 · 출제 의도", "ki"),
                               ("승 · 난이도 핵심", "seung"),
                               ("전 · 학생 오답 패턴", "jeon"),
                               ("결 · 대비 학습법", "gyeol")):
                txt = it.get(key, "").strip()
                if not txt:
                    continue
                p = doc.add_paragraph()
                _add_run(p, label, size=10.5, bold=True, color="2D3A5C")
                p = doc.add_paragraph()
                _add_run(p, txt, size=10.5, color="1A1F36")
                p.paragraph_format.line_spacing = 1.55
                p.paragraph_format.left_indent = Cm(0.5)

    # 푸터
    doc.add_paragraph()
    p = doc.add_paragraph()
    _add_run(p, f"수강 문의: {academy} ☎️ {phone}", size=11.5, bold=True, color="1A1F36")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────
SS_PREFIX = "exam_"


def _ss(key: str, default=None):
    full = SS_PREFIX + key
    if full not in st.session_state:
        st.session_state[full] = default
    return st.session_state[full]


def _set_ss(key: str, value):
    st.session_state[SS_PREFIX + key] = value


def _init_state():
    _ss("meta", None)
    _ss("questions", [])
    _ss("kssjg", [])
    _ss("blog_text", "")
    _ss("blog_image", b"")
    _ss("blog_word", b"")
    _ss("uploaded_keys", [])
    _ss("academy", DEFAULT_ACADEMY)
    _ss("phone", DEFAULT_PHONE)


def render_sidebar():
    """시험 분석 모드의 사이드바."""
    _init_state()

    st.markdown("### 분석 설정")
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    st.markdown(
        "<p class='section-label'>OCR 설정</p>"
        "<div style='font-size:13px;color:var(--text-body);line-height:1.7'>"
        "과목 <span style='color:var(--text-muted)'>·</span> "
        "<b>영어</b> <span class='tag tag-neutral' style='margin-left:4px'>고정</span><br/>"
        "학교/학년 <span style='color:var(--text-muted)'>·</span> "
        "<span style='color:var(--text-muted)'>OCR 자동 추출</span>"
        "</div>",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<p class="section-label">학원 정보 (보고서 푸터)</p>', unsafe_allow_html=True)
    st.text_input("학원명", key=SS_PREFIX + "academy", value=_ss("academy", DEFAULT_ACADEMY))
    st.text_input("연락처", key=SS_PREFIX + "phone", value=_ss("phone", DEFAULT_PHONE))

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<p class="section-label">초기화</p>', unsafe_allow_html=True)
    if st.button("분석 결과 초기화", use_container_width=True):
        for k in ("meta", "questions", "kssjg", "blog_text", "blog_image", "blog_word", "uploaded_keys"):
            st.session_state[SS_PREFIX + k] = None if k == "meta" else ([] if k in ("questions", "kssjg", "uploaded_keys") else (b"" if k in ("blog_image", "blog_word") else ""))
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
    <div class='card-eyebrow'>어려운 문항</div>
    <div class='card-value' style='color:var(--killer-fg)'>{killer_count}</div>
    <div class='card-meta'>난도 {derive_difficulty_label(qs)}</div>
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


def render_main(api_key: str):
    """시험 분석 모드의 본문."""
    if not api_key:
        st.error("OpenAI API Key 가 필요합니다. .env 또는 Streamlit Secrets 에 OPENAI_API_KEY 를 설정하세요.")
        return
    _init_state()

    # ── §1. 업로드 & OCR ──
    st.markdown('<div class="section-mark">§ 1. 업로드 & OCR</div>', unsafe_allow_html=True)

    files = st.file_uploader(
        "시험지 이미지 업로드 (여러 페이지 동시 가능)",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        key="exam_uploader",
        help="시험지 한 부 전체를 페이지별 이미지로 올리세요. PDF는 미리 이미지로 변환해 주세요.",
    )

    if files:
        keys = [f"{f.name}_{f.size}" for f in files]
        if keys != _ss("uploaded_keys"):
            _set_ss("uploaded_keys", keys)
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
                    meta, qs = ocr_exam_images(api_key, img_bytes_list, "영어", "")
                    flags = auto_killer_flags(meta, qs)
                    for q, flag in zip(qs, flags):
                        q.is_killer = flag
                    _set_ss("meta", meta)
                    _set_ss("questions", qs)
                    _set_ss("kssjg", [])
                    _set_ss("blog_text", "")
                    _set_ss("blog_image", b"")
                    _set_ss("blog_word", b"")
                    status.update(label=f"완료 — 문항 {len(qs)}개 추출", state="complete")
                except openai.AuthenticationError:
                    status.update(label="API Key 오류", state="error")
                    st.error("OpenAI API Key 가 유효하지 않습니다.")
                except Exception as e:
                    status.update(label="OCR 실패", state="error")
                    st.error(f"분석 중 오류: {e}")

    meta: ExamMeta | None = _ss("meta")
    qs: list[Question] = _ss("questions")
    if meta is None:
        st.markdown('<div class="empty">이미지를 업로드하고 OCR 을 실행하면 여기에 분석 결과가 표시됩니다.</div>',
                    unsafe_allow_html=True)
        return

    # ── §2. 메타정보 검토 ──
    st.markdown('<div class="section-mark" style="margin-top:32px">§ 2. 메타정보 확인 & 수정</div>',
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
            options = ["중간고사", "기말고사", "모의고사", "수행평가", "기타"]
            meta.exam_type = st.selectbox(
                "시험 종류", options,
                index=options.index(meta.exam_type) if meta.exam_type in options else 0,
                key="m_type",
            )
            meta.exam_date = st.text_input("시험일자", meta.exam_date, key="m_date")
        with c3:
            meta.duration_min = st.number_input("시험 시간 (분)", 10, 200, meta.duration_min, key="m_dur")
            meta.total_score = st.number_input("총 배점", 10, 200, meta.total_score, key="m_tot")
            meta.notes = st.text_input("출제 범위/메모", meta.notes, key="m_notes")
        _set_ss("meta", meta)

    st.markdown("##### 문항 정보 — 표에서 직접 수정")
    st.caption("난이도·유형·배점·범위·어려운 문항 표시를 자유롭게 고치세요. 표에서 고치는 즉시 반영됩니다.")

    df = _questions_to_df(qs)
    edited = st.data_editor(
        df,
        column_config={
            "no": st.column_config.NumberColumn("번호", width=60, format="%d"),
            "type": st.column_config.TextColumn("유형", width=100),
            "difficulty": st.column_config.SelectboxColumn("난이도", options=DIFFICULTY_LEVELS, width=80),
            "score": st.column_config.NumberColumn("배점", min_value=0, max_value=20, step=0.5, width=70),
            "is_subjective": st.column_config.CheckboxColumn("서답형", width=70),
            "is_killer": st.column_config.CheckboxColumn("어려운 문항", width=90),
            "scope": st.column_config.TextColumn("범위", width=110),
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

    qa, qb, qc = st.columns(3)
    with qa:
        if st.button("어려운 문항 자동 표시", use_container_width=True,
                     help="난이도 상 + 배점 상위 25% + 서답형 복합 기준으로 자동 표시합니다."):
            flags = auto_killer_flags(meta, qs)
            for q, f in zip(qs, flags):
                q.is_killer = f
            _set_ss("questions", qs)
            st.rerun()
    with qb:
        if st.button("어려운 문항 모두 해제", use_container_width=True):
            for q in qs:
                q.is_killer = False
            _set_ss("questions", qs)
            st.rerun()
    with qc:
        st.download_button(
            "메타정보 JSON 저장",
            data=json.dumps({"meta": asdict(meta), "questions": [asdict(q) for q in qs]},
                            ensure_ascii=False, indent=2).encode("utf-8"),
            file_name=f"exam_meta_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json",
            use_container_width=True,
        )

    # ── §3. 보고서 생성 (단일 버튼, 한 번에 다 만듦) ──
    st.markdown('<div class="section-mark" style="margin-top:32px">§ 3. 블로그용 보고서 생성</div>',
                unsafe_allow_html=True)
    st.caption("메타정보를 확인한 뒤 아래 버튼 한 번이면 — 어려운 문항 기승전결, 블로그 톤 총평, 차트, Word, 이미지까지 한 번에 만듭니다.")

    academy = _ss("academy", DEFAULT_ACADEMY)
    phone = _ss("phone", DEFAULT_PHONE)

    if st.button("블로그용 분석 보고서 생성", type="primary", use_container_width=True):
        if not qs:
            st.warning("문항 정보가 비어 있습니다. 먼저 OCR 을 실행하거나 표에 추가하세요.")
        else:
            try:
                # Stage A — 어려운 문항 기승전결
                with st.status("어려운 문항 분석 중...", expanded=True) as status:
                    kssjg = gen_killer_kssjg(api_key, meta, qs)
                    _set_ss("kssjg", kssjg)
                    status.update(label=f"기승전결 작성 완료 — {len(kssjg)}문항")

                    # Stage B — 블로그 톤 총평 (LLM × 2 + Python humanize)
                    status.update(label="블로그 칼럼 초안 작성 중...")
                    blog_text = gen_blog_analysis(api_key, meta, qs, academy, phone,
                                                  seed=hash((meta.title, meta.exam_date)) & 0xFFFFFFFF)
                    _set_ss("blog_text", blog_text)
                    status.update(label="블로그 칼럼 폴리시 완료")

                    # Stage C — 차트
                    status.update(label="차트 생성 중...")
                    charts = {
                        "type":       chart_type_distribution(qs),
                        "scope":      chart_scope_distribution(qs),
                        "difficulty": chart_difficulty_distribution(qs),
                        "killer_map": chart_killer_map(qs),
                    }

                    # Stage D — 산출물
                    status.update(label="Word & 이미지 만드는 중...")
                    word_b = build_word_report(meta, qs, blog_text, kssjg, charts, academy, phone)
                    img_b = render_blog_image(meta, qs, blog_text, kssjg, charts, academy, phone)
                    _set_ss("blog_word", word_b)
                    _set_ss("blog_image", img_b)
                    status.update(label="완료 — Word + 이미지 + 텍스트 준비됨", state="complete")
            except openai.AuthenticationError:
                st.error("OpenAI API Key 가 유효하지 않습니다.")
            except Exception as e:
                st.error(f"분석 실패: {e}")

    # ── §4. 결과 미리보기 + 다운로드 ──
    blog_text = _ss("blog_text", "")
    blog_word = _ss("blog_word", b"")
    blog_image = _ss("blog_image", b"")
    kssjg = _ss("kssjg", [])

    if blog_text or blog_image:
        st.markdown('<div class="section-mark" style="margin-top:32px">§ 4. 결과 미리보기 & 다운로드</div>',
                    unsafe_allow_html=True)

        prev_col, dl_col = st.columns([3, 2])
        with prev_col:
            if blog_image:
                st.markdown("##### 블로그 이미지 미리보기")
                st.image(blog_image, use_container_width=True)
            if blog_text:
                with st.expander("블로그용 텍스트 보기", expanded=False):
                    st.text_area("blog_text_preview", value=blog_text, height=400,
                                 label_visibility="collapsed")
        with dl_col:
            st.markdown("##### 다운로드")
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            base = re.sub(r'[\\/*?:"<>|]', "_",
                          meta.title or f"{meta.school}_{meta.grade}_{meta.subject}_{meta.exam_type}")
            if blog_word:
                st.download_button(
                    "Word 보고서 (.docx)",
                    data=blog_word,
                    file_name=f"{base}_분석_{ts}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True,
                )
            if blog_image:
                st.download_button(
                    "블로그 이미지 (.png)",
                    data=blog_image,
                    file_name=f"{base}_분석_{ts}.png",
                    mime="image/png",
                    type="primary",
                    use_container_width=True,
                )
            if blog_text:
                st.download_button(
                    "블로그 텍스트 (.txt)",
                    data=blog_text.encode("utf-8"),
                    file_name=f"{base}_분석_{ts}.txt",
                    mime="text/plain",
                    use_container_width=True,
                )

        # 기승전결 카드 (요약)
        if kssjg:
            st.markdown('<div class="section-mark" style="margin-top:32px">어려운 문항 기승전결</div>',
                        unsafe_allow_html=True)
            for it in kssjg:
                st.markdown(
                    f"<div class='card'>"
                    f"<span class='killer-flag'>#{it.get('no')}</span>"
                    f"<span style='color:var(--text-heading);font-weight:700'>{it.get('type', '')}</span>"
                    f"<span style='color:var(--text-muted);font-size:13px;margin-left:8px'>"
                    f"{it.get('headline', '')}</span>"
                    f"<div style='margin-top:10px;font-size:14px;line-height:1.7'>"
                    f"<b style='color:var(--text-accent)'>기.</b> {it.get('ki','')}<br/>"
                    f"<b style='color:var(--text-accent)'>승.</b> {it.get('seung','')}<br/>"
                    f"<b style='color:var(--text-accent)'>전.</b> {it.get('jeon','')}<br/>"
                    f"<b style='color:var(--text-accent)'>결.</b> {it.get('gyeol','')}"
                    f"</div></div>",
                    unsafe_allow_html=True,
                )
