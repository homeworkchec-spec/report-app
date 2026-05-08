"""
시험지 분석 자동화 모듈 — 블로그 발행용 (기승전결 구조)

워크플로우:
    §1) 이미지 업로드 → GPT-4o Vision OCR (메타 + 문항 + 지문발췌 + 선지)
    §2) 메타정보 확인/수정 (인라인 편집)
    §3) [블로그용 분석 보고서 생성] 단일 버튼 → 자동 진행:
         · 어려운 문항 깊이 분석 (지문 페러프레이징 + 선지 페러프레이징
           + 함정 분석 + 풀이 방법)
         · 기·승·결 본문 + 차트별 캡션 (단일 GPT 호출 → JSON)
         · 차트 4종
         · Word + PNG (이미지 다음에 설명) 동시 생성
    §4) 결과 미리보기 + 다운로드

사람이 쓴 듯한 글을 위해:
    · LEEPIN(최상위학원) 톤 few-shot
    · 멀티스테이지 LLM (초안 → 폴리시) + Python humanize 후처리
    · 도입부 패턴 5종 무작위
"""

from __future__ import annotations

import base64
import io
import json
import platform
import random
import re
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

from design_tokens import chart_palette, ThemeName


# ─────────────────────────────────────────────────────────────
# Korean font setup
# ─────────────────────────────────────────────────────────────
def _setup_korean_font() -> tuple[str | None, str | None]:
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
    # 어려운 문항 깊이 분석을 위해 OCR 단계에서 함께 추출
    passage_excerpt: str = ""           # 지문 핵심 문장 1~2개
    choices: list[str] = field(default_factory=list)  # 선지 텍스트 목록

    def difficulty_num(self) -> int:
        return DIFFICULTY_NUM.get(self.difficulty, 3)


@dataclass
class KillerDeep:
    """어려운 문항 깊이 분석 결과 (페러프레이징 + 함정 + 풀이)."""
    no: int = 0
    type: str = ""
    headline: str = ""
    paraphrase_passage: str = ""        # 1. 지문의 어려운 부분 풀어쓰기
    paraphrase_choices: list[dict] = field(default_factory=list)  # 2. [{"label":"①","text":"..."}]
    trap_analysis: str = ""             # 3. 함정 분석
    solution_method: str = ""           # 4. 풀이 방법


# ─────────────────────────────────────────────────────────────
# OCR — 메타 + 문항 + 지문발췌 + 선지를 한 번에
# ─────────────────────────────────────────────────────────────
OCR_SYSTEM = (
    "당신은 한국 중·고등 영어 시험지를 분석하는 전문가입니다. "
    "이미지에서 시험 메타 정보, 모든 문항의 메타데이터, "
    "그리고 어려운 문항의 깊이 분석에 쓸 지문 핵심 문장과 선지 텍스트를 추출합니다. "
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
      "difficulty": "중하",
      "score": 4.0,
      "is_subjective": false,
      "scope": "Lesson 3",
      "memo": "관계대명사 변형",
      "passage_excerpt": "본문에서 정답 단서가 되는 핵심 문장 1~2개. 가능한 정확히 옮겨 적되, 너무 길면 핵심부만.",
      "choices": ["① ...", "② ...", "③ ...", "④ ...", "⑤ ..."]
    }}
  ]
}}

[난이도 판단 — 매우 정밀하게 분포시킬 것]
- 상 (Killer): 다중 변형 + 시간 소모 + 고난도 추론. **시험 전체에서 보통 1~2문항.** 매우 엄격하게 판단.
- 중상: 응용·복합. 한 가지 핵심 함정. 보통 2~4문항.
- 중: 표준 응용, 다단계 추론. 보통 5~8문항.
- 중하: 표준 개념 적용. 보통 5~8문항.
- 하: 단순 일치/암기. 보통 3~6문항.

⚠ 모든 문항을 같은 라벨(특히 "중")로 몰지 말 것. 다섯 단계가 정밀하게 분포하도록
   미세 차이를 구분해 분류하세요. "상"은 빡빡하게(엄격하게), 나머지는 다양하게.

[유형명(영어 표준)]
일치/불일치, 빈칸, 어법, 어휘, 순서, 삽입, 요약문, 함의, 대의, 지칭,
영영풀이, 조건영작, 서술형, 단답형 — 위 표준명을 우선 사용.

[passage_excerpt 와 choices]
- 객관식: passage_excerpt 는 정답 단서 핵심 문장. choices 는 ①~⑤ 텍스트 그대로.
- 서답형: passage_excerpt 는 답안의 단서가 되는 본문 부분. choices 는 빈 배열 [].
- 본문 분량이 너무 길면 정답 단서 핵심부만 남기되, 어려운 문항(상/중상)일수록 더 풍부하게.

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
        max_tokens=6500,
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
            ch = q.get("choices") or []
            if not isinstance(ch, list):
                ch = []
            qs.append(Question(
                no=int(q.get("no", 0) or 0),
                type=str(q.get("type") or "").strip(),
                difficulty=diff,
                score=float(q.get("score") or 0),
                is_subjective=bool(q.get("is_subjective", False)),
                scope=str(q.get("scope") or "").strip(),
                memo=str(q.get("memo") or "").strip(),
                passage_excerpt=str(q.get("passage_excerpt") or "").strip(),
                choices=[str(c).strip() for c in ch if str(c).strip()],
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
    """상 난이도는 엄격하게: 기본은 difficulty=='상' 만. 단, 상이 0개이면
    중상 + 배점 상위 25% 만 보조 표시."""
    if not qs:
        return []
    flags = [q.difficulty == "상" for q in qs]
    if not any(flags):
        scores = sorted([q.score for q in qs if q.score > 0])
        threshold = scores[int(len(scores) * 0.75)] if scores else 0
        for i, q in enumerate(qs):
            if q.difficulty == "중상" and q.score >= threshold:
                flags[i] = True
    return flags


def derive_difficulty_label(qs: list[Question]) -> str:
    if not qs:
        return "中"
    nums = [q.difficulty_num() for q in qs]
    avg = sum(nums) / len(nums)
    killer_ratio = sum(1 for q in qs if q.is_killer) / len(qs)
    score = avg + killer_ratio * 1.0
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

OPENING_VARIANTS = [
    "이번 시험은",
    "전반적으로 보면 이번 시험은",
    "결과만 놓고 보면",
    "체감 난이도부터 짚어보면",
    "큰 틀에서 이번 시험은",
]


def humanize_text(text: str, seed: int | None = None) -> str:
    rng = random.Random(seed)
    out = text
    for ai_phrase, alts in AI_TELLS:
        if ai_phrase in out:
            out = out.replace(ai_phrase, rng.choice(alts), 1)
    sentences = re.split(r"(?<=[.!?])\s+", out)
    if len(sentences) >= 3:
        ips = [i for i, s in enumerate(sentences) if s.endswith("입니다.")]
        if ips and rng.random() < 0.25:
            i = rng.choice(ips)
            sentences[i] = sentences[i][:-4] + "라 할 수 있습니다."
        out = " ".join(sentences)
    out = re.sub(r"분석한 결과,?\s*", "", out)
    out = re.sub(r"^\s*결과적으로,?\s*", "", out, flags=re.MULTILINE)
    out = re.sub(r"[ \t]+", " ", out)
    out = re.sub(r"\n{3,}", "\n\n", out)
    return out.strip()


# ─────────────────────────────────────────────────────────────
# 어려운 문항 깊이 분석 (지문 페러프레이징 + 선지 페러프레이징 + 함정 + 풀이)
# ─────────────────────────────────────────────────────────────
KILLER_DEEP_SYSTEM = """당신은 입시 영어 분석 전문가입니다.
어려운 문항을 학부모와 학생이 이해할 수 있는 깊이 분석으로 풀어주세요.

각 문항당 4가지 항목을 작성합니다:

1. paraphrase_passage — 지문의 어려운 부분 풀어쓰기
   원문에서 학생이 막히는 핵심 문장을 한국어로 풀어 씁니다.
   "원문은 ~ 라고 표현했는데, 풀어 쓰면 ~ 라는 의미입니다." 형태도 자연스럽습니다.

2. paraphrase_choices — 선지 풀어쓰기 (객관식만, 서답형은 빈 배열)
   각 선지를 학생 입장에서 한 줄로 풀어 씁니다.
   [{"label": "①", "text": "이 선지는 ~ 라는 뜻입니다."}, ...]

3. trap_analysis — 함정 분석
   학생이 흔히 빠지는 오답 흐름. 어떤 선지가 왜 매력적으로 보이는지.

4. solution_method — 풀이 방법
   접근 순서. 어떤 단서를 먼저 보고, 어떤 순서로 좁혀가야 하는지 구체적 행동 지침.

추가:
- headline: 한 줄로 이 문항의 본질을 요약 (예: "본문 변형 + 어법 조건 결합형")

[금지 표현]
"다음과 같이", "다양한", "효과적으로", "이를 통해", "여러분", "체계적으로",
"~을 통해서", "결론적으로 말씀드리면"

문체: 학원 원장이 학부모에게 차분히 설명하는 단정한 한국어. "~합니다" 종결.
JSON 으로만 응답하세요."""


def gen_killer_deep(api_key: str, meta: ExamMeta, qs: list[Question]) -> list[KillerDeep]:
    killers = [q for q in qs if q.is_killer]
    if not killers:
        return []
    client = openai.OpenAI(api_key=api_key)

    items_input = []
    for q in killers:
        items_input.append({
            "no": q.no,
            "type": q.type,
            "difficulty": q.difficulty,
            "score": q.score,
            "is_subjective": q.is_subjective,
            "scope": q.scope,
            "memo": q.memo,
            "passage_excerpt": q.passage_excerpt,
            "choices": q.choices,
        })

    user = f"""[시험 정보]
{meta.school} {meta.grade} {meta.subject} {meta.exam_type}
총 {meta.total_questions}문항 · {meta.total_score}점 · {meta.duration_min}분

[어려운 문항 데이터]
{json.dumps(items_input, ensure_ascii=False, indent=2)}

요청: 위 각 문항에 대해 깊이 분석을 작성하세요.
객관식 문항은 paraphrase_choices 를 모든 선지에 대해 작성하고,
서답형은 paraphrase_choices 를 빈 배열 [] 로 두되 paraphrase_passage 와
solution_method 를 더 풍부하게 써주세요.

응답:
{{
  "items": [
    {{
      "no": 4,
      "type": "조건영작",
      "headline": "본문 변형 + 어법 조건 결합형",
      "paraphrase_passage": "원문은 ... 라고 했는데, 풀어 쓰면 ...",
      "paraphrase_choices": [
        {{"label": "①", "text": "이 선지는 ..."}},
        {{"label": "②", "text": "이 선지는 ..."}}
      ],
      "trap_analysis": "...",
      "solution_method": "..."
    }}
  ]
}}

JSON 외 텍스트 없이 반환하세요."""

    resp = client.chat.completions.create(
        model=TEXT_MODEL,
        messages=[
            {"role": "system", "content": KILLER_DEEP_SYSTEM},
            {"role": "user", "content": user},
        ],
        temperature=0.55,
        max_tokens=2500,
        response_format={"type": "json_object"},
    )
    data = _safe_json_loads(resp.choices[0].message.content or "{}")
    out: list[KillerDeep] = []
    for it in data.get("items", []) or []:
        ch = it.get("paraphrase_choices") or []
        if not isinstance(ch, list):
            ch = []
        out.append(KillerDeep(
            no=int(it.get("no", 0) or 0),
            type=str(it.get("type", "")),
            headline=humanize_text(it.get("headline", "")),
            paraphrase_passage=humanize_text(it.get("paraphrase_passage", "")),
            paraphrase_choices=[
                {"label": str(c.get("label", "")), "text": humanize_text(str(c.get("text", "")))}
                for c in ch if isinstance(c, dict)
            ],
            trap_analysis=humanize_text(it.get("trap_analysis", "")),
            solution_method=humanize_text(it.get("solution_method", "")),
        ))
    return out


# ─────────────────────────────────────────────────────────────
# 보고서 본문 — 기 / 승 / 결 + 차트 캡션 (단일 호출, JSON)
#   "전(轉)" 자리는 코드가 killer_deep 카드를 끼워 넣음.
# ─────────────────────────────────────────────────────────────
LEEPIN_SAMPLE = """이번 시험은 객관적 난이도 '중하' 수준으로, 기본기가 충실한 학생이라면 무난히 풀 수 있는 구성이었습니다. 다만 일부 문항의 선택지가 교묘하게 짜여 체감 난이도는 '중'까지 올라갔을 것으로 보입니다. 서술형 난도가 하락한 점은 주목할 변화이며, 객관식에서의 한 번 실수가 등급을 가르는 구조였습니다.

균형 잡힌 범위 분포가 이번 시험의 큰 그림입니다. 5과·6과가 36%씩 동등하게 출제되었고 모의고사가 28%를 차지해, 단원 편식을 허용하지 않는 구성이었습니다. 변별의 분기점은 '일치/불일치'와 '어휘'에 있었습니다. 두 유형이 전체의 44%를 차지하며 꼼꼼한 지문 분석 능력을 직접 평가했습니다. 서답형은 평이해 점수를 얻는 구간이 아닌, 실점하지 않고 지켜내야 하는 구간으로 작용했습니다.

다음 시험을 위한 학습 전략은 분명합니다. 첫째, 전 범위 균형 학습이 필수적입니다. 한 단원에만 시간을 쏟는 방식으로는 상위 등급이 어렵습니다. 둘째, 지문 완전 정복에 집중해야 합니다. 단순 해석을 넘어 세부 정보까지 근거를 잡아 푸는 훈련이 요구됩니다. 셋째, 서답형 실수 방지 루틴이 필요합니다. 핵심 구문과 어휘는 의미뿐 아니라 정확한 철자까지 완벽하게 쓰는 연습을 매일 반복하셔야 합니다. 평이한 기조에 안주하기보다, 기본기를 단단히 다지면서 고난도 변화에도 흔들리지 않는 실력을 만드는 것이 장기적 전략입니다."""


BLOG_BODY_SYSTEM = """당신은 학원 원장 LEEPIN 입니다. 영어 입시 전문이며, 중·고등 내신 시험을 직접 분석해 블로그 칼럼을 씁니다.

[당신의 글쓰기 스타일]
- 학부모와 학생이 함께 읽는 블로그 톤. 단단한 문어체이지만 "갈렸습니다", "치명적인", "변별의 분기점", "기회이자 위기"처럼 한국 입시 지도자의 표현을 자유롭게 사용
- 객관 데이터와 전문가 판단을 한 단락 안에서 자연스럽게 섞음
- 짧은 문장과 긴 문장을 섞어 리듬을 만듦

[금지 표현 — 절대 사용 금지]
"다음과 같이", "분석한 결과", "살펴보겠습니다", "여러분", "효과적으로",
"종합적으로", "다양한", "이를 통해", "이러한 측면에서", "결론적으로 말씀드리면",
"~을 통해서", "체계적으로", "활용하여",
"~인 점은 인상적입니다", "~을 알 수 있습니다", "~로 사료됩니다"
이모지 본문 사용 금지(섹션 헤딩에는 1개씩 허용 — 출력은 본문만이므로 불필요).

[기·승·결 작성 분담]
- 기 (gi): 시험 한눈 진단. 4~5 문장. 객관 vs 체감 난이도, 출제 의도 한마디, 다음 시험 시사.
- 승 (seung): 출제 분포의 의미. 5~7 문장 한 단락. 유형/범위/난이도 분포가 무엇을 신호하는지.
- 결 (gyeol): 학습 전략과 대비 방법. 첫째/둘째/셋째 형태로 3가지 명령형 + 부연. 마지막 한 단락은 장기 학습 태도. 대비 방법은 구체적 행동(루틴, 분량, 순서)이 드러나야 함.

[차트 캡션 4종]
각 차트 아래 본문에 들어갈 짧은 설명을 작성합니다. 1~3 문장씩.
- type_caption: 유형별 분포 차트 아래
- scope_caption: 범위별 분포 차트 아래 (범위 데이터가 있을 때만)
- difficulty_caption: 난이도 분포 차트 아래
- location_caption: 위치별 난이도 + 어려운 문항 차트 아래

전부 한 번에 JSON 으로 응답하세요."""


def gen_blog_body(api_key: str, meta: ExamMeta, qs: list[Question],
                  killer_deeps: list[KillerDeep], academy: str, phone: str,
                  seed: int | None = None) -> dict:
    """반환: {gi, seung, gyeol, captions:{type,scope,difficulty,location}}"""
    if not qs:
        return {"gi": "", "seung": "", "gyeol": "", "captions": {}}
    rng = random.Random(seed)
    client = openai.OpenAI(api_key=api_key)

    diff_label = derive_difficulty_label(qs)
    type_dist = type_distribution(qs)
    scope_dist = scope_distribution(qs)
    diff_dist = difficulty_distribution(qs)

    title = (meta.title or " ".join(filter(None, [
        meta.school, meta.grade, meta.exam_date[:7] if meta.exam_date else "",
        meta.exam_type, "영어 시험 분석",
    ])))

    type_lines = "\n".join(f"- {t}: {n}문항 ({p:.1f}%)" for t, n, p in type_dist)
    scope_lines = "\n".join(f"- {s}: {n}문항 ({p:.1f}%)" for s, n, p in scope_dist) or "(범위 정보 미제공)"
    diff_lines = "\n".join(f"- {d}: {n}문항 ({p:.1f}%)" for d, n, p in diff_dist if n > 0)
    killer_brief = "\n".join(
        f"- {kd.no}번 ({kd.type}): {kd.headline}" for kd in killer_deeps
    ) or "- 명시적 어려운 문항 없음"

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

[어려운 문항 헤드라인 (전(轉) 자리에 코드가 별도 카드로 삽입)]
{killer_brief}

────────────────
지시:
- 기·승·결 본문은 LEEPIN 톤으로 작성 (시작 어구는 "{rng.choice(OPENING_VARIANTS)}" 등 자연스럽게 변형 가능).
- 데이터(분포/문항수)는 본문 안에 자연스럽게 녹여 쓰되, 표·불릿 나열은 만들지 마세요.
- 결(gyeol) 마지막 단락은 학습 태도/대비 방법으로 강하게 마무리.
- 차트 캡션 4종은 각 1~3 문장. {('범위 데이터가 없으므로 scope_caption 은 빈 문자열' if not scope_dist else '')}

응답 JSON 스키마:
{{
  "gi": "...",
  "seung": "...",
  "gyeol": "...",
  "captions": {{
    "type": "...",
    "scope": "...",
    "difficulty": "...",
    "location": "..."
  }}
}}

JSON 외 텍스트 없이 반환하세요."""

    # Stage 1 — 초안
    draft = client.chat.completions.create(
        model=TEXT_MODEL,
        messages=[
            {"role": "system", "content": BLOG_BODY_SYSTEM},
            {"role": "user", "content": user},
        ],
        temperature=0.85,
        max_tokens=2400,
        response_format={"type": "json_object"},
    ).choices[0].message.content or "{}"

    # Stage 2 — 폴리시 (텍스트 필드만)
    polish_sys = (
        "당신은 한국 입시 학원 원장의 글을 자연스럽게 다듬는 편집자입니다. "
        "AI 가 쓴 듯한 매끄러움을 줄이고 사람이 손으로 쓴 호흡을 살리세요. "
        "금지 표현을 자연스러운 표현으로 모두 대체: "
        "'다음과 같이', '분석한 결과', '살펴보겠습니다', '효과적으로', '종합적으로', "
        "'다양한', '이를 통해', '체계적으로', '활용하여'. "
        "원문 JSON 키 구조와 의미는 보존하되, 각 텍스트 값의 어휘와 리듬만 다듬어 같은 JSON 으로 반환하세요."
    )
    polished = client.chat.completions.create(
        model=TEXT_MODEL,
        messages=[
            {"role": "system", "content": polish_sys},
            {"role": "user", "content": draft},
        ],
        temperature=0.6,
        max_tokens=2400,
        response_format={"type": "json_object"},
    ).choices[0].message.content or draft

    data = _safe_json_loads(polished)
    captions = data.get("captions", {}) or {}
    return {
        "gi": humanize_text(data.get("gi", ""), seed=seed),
        "seung": humanize_text(data.get("seung", ""), seed=(seed or 0) + 1),
        "gyeol": humanize_text(data.get("gyeol", ""), seed=(seed or 0) + 2),
        "captions": {
            "type": humanize_text(captions.get("type", "")),
            "scope": humanize_text(captions.get("scope", "")),
            "difficulty": humanize_text(captions.get("difficulty", "")),
            "location": humanize_text(captions.get("location", "")),
        },
    }


# ─────────────────────────────────────────────────────────────
# 차트 — 테마별 팔레트
# ─────────────────────────────────────────────────────────────
def _editorial_style(ax, palette):
    ax.set_facecolor(palette["paper"])
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color(palette["rule"])
        ax.spines[s].set_linewidth(0.8)
    ax.tick_params(colors=palette["ink"], labelsize=10)
    ax.grid(axis="y", color=palette["rule"], linestyle="-", linewidth=0.5, alpha=0.6)
    ax.set_axisbelow(True)


def chart_type_distribution(qs: list[Question], theme: ThemeName = "editorial") -> bytes:
    pal = chart_palette(theme)
    dist = type_distribution(qs)
    if not dist:
        return b""
    labels, counts, pcts = zip(*dist)
    fig, ax = plt.subplots(figsize=(9, max(4.5, 0.55 * len(labels))), facecolor=pal["paper"])
    bars = ax.barh(labels[::-1], counts[::-1],
                   color=[pal["accents"][i % len(pal["accents"])] for i in range(len(labels))][::-1],
                   edgecolor=pal["paper"], height=0.7)
    for bar, n, p in zip(bars, counts[::-1], pcts[::-1]):
        ax.text(n + max(counts) * 0.02, bar.get_y() + bar.get_height() / 2,
                f"{n}문항 ({p:.1f}%)", va="center", fontsize=10.5, color=pal["ink"])
    ax.set_xlim(0, max(counts) * 1.32)
    ax.set_title("유형별 출제 비중", loc="left", fontsize=14,
                 color=pal["ink"], fontweight="bold", pad=14)
    _editorial_style(ax, pal)
    ax.tick_params(axis="x", labelsize=0)
    ax.spines["bottom"].set_visible(False)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, facecolor=pal["paper"], bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def chart_scope_distribution(qs: list[Question], theme: ThemeName = "editorial") -> bytes:
    pal = chart_palette(theme)
    dist = scope_distribution(qs)
    if not dist:
        return b""
    labels, counts, pcts = zip(*dist)
    fig, ax = plt.subplots(figsize=(8, max(3.5, 0.55 * len(labels))), facecolor=pal["paper"])
    bars = ax.barh(labels[::-1], counts[::-1],
                   color=[pal["accents"][(i + 2) % len(pal["accents"])] for i in range(len(labels))][::-1],
                   edgecolor=pal["paper"], height=0.7)
    for bar, n, p in zip(bars, counts[::-1], pcts[::-1]):
        ax.text(n + max(counts) * 0.02, bar.get_y() + bar.get_height() / 2,
                f"{n}문항 ({p:.1f}%)", va="center", fontsize=10.5, color=pal["ink"])
    ax.set_xlim(0, max(counts) * 1.32)
    ax.set_title("범위별 문항 분포", loc="left", fontsize=14,
                 color=pal["ink"], fontweight="bold", pad=14)
    _editorial_style(ax, pal)
    ax.tick_params(axis="x", labelsize=0)
    ax.spines["bottom"].set_visible(False)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, facecolor=pal["paper"], bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def chart_difficulty_distribution(qs: list[Question], theme: ThemeName = "editorial") -> bytes:
    pal = chart_palette(theme)
    dist = difficulty_distribution(qs)
    fig, ax = plt.subplots(figsize=(8, 3.6), facecolor=pal["paper"])
    labels, counts, pcts = zip(*dist)
    bars = ax.bar(labels, counts,
                  color=[pal["accents"][i % len(pal["accents"])] for i in range(len(labels))],
                  edgecolor=pal["paper"], linewidth=2, width=0.55)
    for bar, n, p in zip(bars, counts, pcts):
        if n > 0:
            ax.text(bar.get_x() + bar.get_width() / 2, n + max(counts) * 0.02,
                    f"{n}\n({p:.1f}%)", ha="center", fontsize=9.5,
                    color=pal["ink"], fontweight="bold")
    ax.set_title("난이도 분포", loc="left", fontsize=14,
                 color=pal["ink"], fontweight="bold", pad=14)
    _editorial_style(ax, pal)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, facecolor=pal["paper"], bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def chart_killer_map(qs: list[Question], theme: ThemeName = "editorial") -> bytes:
    pal = chart_palette(theme)
    if not qs:
        return b""
    fig, ax = plt.subplots(figsize=(9, 3.0), facecolor=pal["paper"])
    nos = [q.no for q in qs]
    ds = [q.difficulty_num() for q in qs]
    killer_color = pal["accents"][3] if len(pal["accents"]) > 3 else pal["accents"][-1]
    colors = [killer_color if q.is_killer else pal["ink"] for q in qs]
    sizes = [120 if q.is_killer else 36 for q in qs]
    ax.scatter(nos, ds, c=colors, s=sizes, edgecolor=pal["paper"], linewidth=1.2, zorder=3)
    ax.plot(nos, ds, color=pal["ink"], alpha=0.16, linewidth=1, zorder=1)
    for q in qs:
        if q.is_killer:
            ax.annotate(f"#{q.no}", (q.no, q.difficulty_num()),
                        xytext=(0, 12), textcoords="offset points",
                        ha="center", fontsize=9.5, color=killer_color, fontweight="bold")
    ax.set_yticks(list(DIFFICULTY_NUM.values()))
    ax.set_yticklabels(list(DIFFICULTY_NUM.keys()))
    ax.set_xlabel("문항 번호", fontsize=10, color=pal["ink"])
    ax.set_title("문항 위치별 난이도 & 어려운 문항 분포", loc="left",
                 fontsize=14, color=pal["ink"], fontweight="bold", pad=14)
    _editorial_style(ax, pal)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, facecolor=pal["paper"], bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# PIL utilities
# ─────────────────────────────────────────────────────────────
def _font(size: int, bold: bool = False) -> ImageFont.ImageFont:
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
                    font: ImageFont.ImageFont, color, line_height: float = 1.65) -> int:
    if not text.strip():
        return y
    lines = _wrap(text, font, max_w)
    line_h = int(font.size * line_height)
    for ln in lines:
        draw.text((x, y), ln, font=font, fill=color)
        y += line_h
    return y


# ─────────────────────────────────────────────────────────────
# 블로그 단일 PNG 렌더 — 기·승·전·결 + 이미지 → 설명
# ─────────────────────────────────────────────────────────────
THEME_INKS = {
    "mono":      {"ink": (10, 10, 10),  "muted": (107, 114, 128), "accent": (10, 10, 10),  "rule": (212, 212, 216)},
    "editorial": {"ink": (26, 31, 54),  "muted": (110, 115, 130), "accent": (45, 58, 92),  "rule": (212, 207, 192)},
    "vivid":     {"ink": (26, 31, 54),  "muted": (107, 114, 128), "accent": (74, 90, 140), "rule": (229, 231, 235)},
}


def render_blog_image(meta: ExamMeta, qs: list[Question],
                      gi: str, seung: str, gyeol: str,
                      captions: dict, killer_deeps: list[KillerDeep],
                      charts: dict[str, bytes],
                      academy: str = DEFAULT_ACADEMY, phone: str = DEFAULT_PHONE,
                      theme: ThemeName = "editorial") -> bytes:
    W = 1080
    PAD = 60
    INNER = W - PAD * 2
    BG = (255, 255, 255)
    inks = THEME_INKS.get(theme, THEME_INKS["editorial"])
    INK = inks["ink"]
    MUTED = inks["muted"]
    ACCENT = inks["accent"]
    RULE = inks["rule"]

    F_TITLE = _font(38, bold=True)
    F_H1    = _font(26, bold=True)
    F_H2    = _font(20, bold=True)
    F_H3    = _font(17, bold=True)
    F_BODY  = _font(17)
    F_BODY_S = _font(15)
    F_META  = _font(14)
    F_KICK  = _font(15, bold=True)

    # ── 측정 ──
    def m_para(text, font, lh=1.65) -> int:
        if not text or not text.strip():
            return 0
        return int(font.size * lh) * len(_wrap(text, font, INNER))

    H = PAD * 2
    H += int(F_TITLE.size * 1.4) + 30
    H += int(F_META.size * 1.6) + 28
    H += 1 + 24
    # 기 (총평)
    H += int(F_H1.size * 1.4) + 8 + m_para(gi, F_BODY) + 26
    # 승 (출제 분석) + 차트 4종 (이미지 → 캡션)
    if seung.strip():
        H += int(F_H1.size * 1.4) + 8 + m_para(seung, F_BODY) + 22
    for key, cap_key in [("type", "type"), ("scope", "scope"),
                         ("difficulty", "difficulty"), ("killer_map", "location")]:
        if charts.get(key):
            H += 600 + 12
            cap = (captions or {}).get(cap_key, "")
            if cap:
                H += m_para(cap, F_BODY_S) + 18
    # 전 (어려운 문항 깊이 분석)
    if killer_deeps:
        H += int(F_H1.size * 1.4) + 16
        for kd in killer_deeps:
            H += int(F_H2.size * 1.4) + 8                  # 헤드라인
            # 1. paraphrase passage
            H += int(F_H3.size * 1.4) + m_para(kd.paraphrase_passage, F_BODY) + 16
            # 2. paraphrase choices
            if kd.paraphrase_choices:
                H += int(F_H3.size * 1.4)
                for ch in kd.paraphrase_choices:
                    H += m_para(f"{ch.get('label','')} {ch.get('text','')}", F_BODY) + 4
                H += 12
            # 3. trap
            H += int(F_H3.size * 1.4) + m_para(kd.trap_analysis, F_BODY) + 16
            # 4. solution
            H += int(F_H3.size * 1.4) + m_para(kd.solution_method, F_BODY) + 22
    # 결 (학습 전략 + 대비)
    if gyeol.strip():
        H += int(F_H1.size * 1.4) + 8 + m_para(gyeol, F_BODY) + 26
    H += int(F_H2.size * 1.6) + 24  # 푸터

    canvas = Image.new("RGB", (W, max(H, 1200)), BG)
    draw = ImageDraw.Draw(canvas)
    y = PAD

    # ── 타이틀 ──
    title = (meta.title or " ".join(filter(None, [
        meta.school, meta.grade, meta.exam_date[:7] if meta.exam_date else "",
        meta.exam_type, "영어 시험 분석",
    ])))
    for ln in _wrap(title, F_TITLE, INNER):
        draw.text((PAD, y), ln, font=F_TITLE, fill=INK)
        y += int(F_TITLE.size * 1.2)
    y += 10

    # 메타
    diff_label = derive_difficulty_label(qs)
    meta_str = " · ".join(filter(None, [
        meta.school or "", meta.grade or "", f"총 {meta.total_questions}문항",
        f"{meta.total_score}점", f"{meta.duration_min}분", f"난도 {diff_label}",
    ]))
    draw.text((PAD, y), meta_str, font=F_META, fill=MUTED)
    y += int(F_META.size * 1.4) + 18
    draw.line([(PAD, y), (W - PAD, y)], fill=ACCENT, width=2)
    y += 26

    # ── 기. 총평 ──
    if gi.strip():
        draw.text((PAD, y), f"총평 (난도: {diff_label})", font=F_H1, fill=INK)
        y += int(F_H1.size * 1.3) + 10
        y = _draw_paragraph(draw, gi, PAD, y, INNER, F_BODY, color=INK)
        y += 24

    # ── 승. 출제 분석 + 차트 (이미지 → 캡션) ──
    if seung.strip():
        draw.text((PAD, y), "📊 출제 분석", font=F_H1, fill=INK)
        y += int(F_H1.size * 1.3) + 10
        y = _draw_paragraph(draw, seung, PAD, y, INNER, F_BODY, color=INK)
        y += 18

    chart_specs = [
        ("type", "type", "유형별 출제 비중"),
        ("scope", "scope", "범위별 문항 분포"),
        ("difficulty", "difficulty", "난이도 분포"),
        ("killer_map", "location", "문항 위치별 난이도 & 어려운 문항"),
    ]
    for key, cap_key, _ in chart_specs:
        if not charts.get(key):
            continue
        chart_im = Image.open(io.BytesIO(charts[key])).convert("RGB")
        ratio = INNER / chart_im.width
        new_h = int(chart_im.height * ratio)
        chart_im = chart_im.resize((INNER, new_h), Image.LANCZOS)
        canvas.paste(chart_im, (PAD, y))
        y += new_h + 10
        cap = (captions or {}).get(cap_key, "")
        if cap.strip():
            y = _draw_paragraph(draw, cap, PAD, y, INNER, F_BODY_S, color=MUTED, line_height=1.6)
        y += 14

    # ── 전. 어려운 문항 깊이 분석 ──
    if killer_deeps:
        y += 10
        draw.text((PAD, y), "🎯 변별의 분기점 — 어려운 문항 깊이 분석", font=F_H1, fill=INK)
        y += int(F_H1.size * 1.4) + 10
        for kd in killer_deeps:
            head = f"{kd.no}번 · {kd.type} — {kd.headline}"
            for ln in _wrap(head, F_H2, INNER):
                draw.text((PAD, y), ln, font=F_H2, fill=ACCENT)
                y += int(F_H2.size * 1.25)
            y += 6
            # 1. paraphrase passage
            draw.text((PAD, y), "1. 어려운 지문 부분 풀어쓰기", font=F_H3, fill=ACCENT)
            y += int(F_H3.size * 1.4)
            y = _draw_paragraph(draw, kd.paraphrase_passage, PAD + 18, y, INNER - 18, F_BODY, color=INK)
            y += 12
            # 2. choices
            if kd.paraphrase_choices:
                draw.text((PAD, y), "2. 선지 풀어쓰기", font=F_H3, fill=ACCENT)
                y += int(F_H3.size * 1.4)
                for ch in kd.paraphrase_choices:
                    line = f"{ch.get('label','')} {ch.get('text','')}"
                    y = _draw_paragraph(draw, line, PAD + 18, y, INNER - 18, F_BODY, color=INK, line_height=1.55)
                    y += 2
                y += 8
            # 3. trap
            draw.text((PAD, y), "3. 함정 분석", font=F_H3, fill=ACCENT)
            y += int(F_H3.size * 1.4)
            y = _draw_paragraph(draw, kd.trap_analysis, PAD + 18, y, INNER - 18, F_BODY, color=INK)
            y += 12
            # 4. solution
            draw.text((PAD, y), "4. 풀이 방법", font=F_H3, fill=ACCENT)
            y += int(F_H3.size * 1.4)
            y = _draw_paragraph(draw, kd.solution_method, PAD + 18, y, INNER - 18, F_BODY, color=INK)
            y += 22

    # ── 결. 학습 전략 + 대비 방법 ──
    if gyeol.strip():
        draw.text((PAD, y), "💡 학습 전략과 대비 방법", font=F_H1, fill=INK)
        y += int(F_H1.size * 1.3) + 10
        y = _draw_paragraph(draw, gyeol, PAD, y, INNER, F_BODY, color=INK)
        y += 22

    # 푸터
    draw.line([(PAD, y), (W - PAD, y)], fill=RULE, width=1)
    y += 18
    draw.text((PAD, y), f"수강 문의: {academy} ☎️ {phone}", font=F_H2, fill=INK)

    final_h = y + int(F_H2.size * 1.4) + PAD
    if final_h < canvas.height:
        canvas = canvas.crop((0, 0, W, final_h))
    out = io.BytesIO()
    canvas.save(out, format="PNG", optimize=True)
    return out.getvalue()


# ─────────────────────────────────────────────────────────────
# Word 보고서 — PNG 와 동일 흐름 (이미지 → 캡션)
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


def build_word_report(meta: ExamMeta, qs: list[Question],
                      gi: str, seung: str, gyeol: str,
                      captions: dict, killer_deeps: list[KillerDeep],
                      charts: dict[str, bytes],
                      academy: str = DEFAULT_ACADEMY, phone: str = DEFAULT_PHONE) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21.0), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(1.6)
    section.left_margin = section.right_margin = Cm(1.8)

    diff_label = derive_difficulty_label(qs)

    # 표지
    title = (meta.title or " ".join(filter(None, [
        meta.school, meta.grade, meta.exam_date[:7] if meta.exam_date else "",
        meta.exam_type, "영어 시험 분석",
    ])))
    p = doc.add_paragraph()
    _add_run(p, title, size=22, bold=True, color="1A1F36")
    p = doc.add_paragraph()
    meta_str = " · ".join(filter(None, [
        meta.school or "", meta.grade or "",
        f"총 {meta.total_questions}문항", f"{meta.total_score}점", f"{meta.duration_min}분",
        f"난도 {diff_label}",
    ]))
    _add_run(p, meta_str, size=10, color="6E7382")
    doc.add_paragraph()

    # 기 — 총평
    if gi.strip():
        p = doc.add_paragraph()
        _add_run(p, f"총평 (난도: {diff_label})", size=14, bold=True, color="1A1F36")
        for para in gi.split("\n\n"):
            para = para.strip()
            if not para:
                continue
            p = doc.add_paragraph()
            _add_run(p, para, size=10.5, color="1A1F36")
            p.paragraph_format.line_spacing = 1.65
            p.paragraph_format.space_after = Pt(8)

    # 승 — 출제 분석 + 차트 (이미지 → 캡션)
    if seung.strip():
        doc.add_paragraph()
        p = doc.add_paragraph()
        _add_run(p, "📊 출제 분석", size=14, bold=True, color="1A1F36")
        for para in seung.split("\n\n"):
            para = para.strip()
            if not para:
                continue
            p = doc.add_paragraph()
            _add_run(p, para, size=10.5, color="1A1F36")
            p.paragraph_format.line_spacing = 1.65
            p.paragraph_format.space_after = Pt(8)

    chart_order = [
        ("type", "type", "유형별 출제 비중"),
        ("scope", "scope", "범위별 문항 분포"),
        ("difficulty", "difficulty", "난이도 분포"),
        ("killer_map", "location", "문항 위치별 난이도 & 어려운 문항"),
    ]
    for key, cap_key, label in chart_order:
        if not charts.get(key):
            continue
        doc.add_paragraph()
        p = doc.add_paragraph()
        _add_run(p, label, size=12, bold=True, color="1A1F36")
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture(io.BytesIO(charts[key]), width=Cm(16))
        cap = (captions or {}).get(cap_key, "")
        if cap.strip():
            p = doc.add_paragraph()
            _add_run(p, cap, size=10, color="6E7382")
            p.paragraph_format.line_spacing = 1.6

    # 전 — 어려운 문항 깊이 분석
    if killer_deeps:
        doc.add_paragraph()
        p = doc.add_paragraph()
        _add_run(p, "🎯 변별의 분기점 — 어려운 문항 깊이 분석", size=14, bold=True, color="1A1F36")
        for kd in killer_deeps:
            p = doc.add_paragraph()
            _add_run(p, f"{kd.no}번 · {kd.type} ", size=12, bold=True, color="2D3A5C", mono=True)
            _add_run(p, f"— {kd.headline}", size=12, bold=True, color="2D3A5C")

            def _section(title_text, body_text):
                p = doc.add_paragraph()
                _add_run(p, title_text, size=10.5, bold=True, color="2D3A5C")
                p = doc.add_paragraph()
                _add_run(p, body_text, size=10.5, color="1A1F36")
                p.paragraph_format.line_spacing = 1.6
                p.paragraph_format.left_indent = Cm(0.5)

            _section("1. 어려운 지문 부분 풀어쓰기", kd.paraphrase_passage)
            if kd.paraphrase_choices:
                p = doc.add_paragraph()
                _add_run(p, "2. 선지 풀어쓰기", size=10.5, bold=True, color="2D3A5C")
                for ch in kd.paraphrase_choices:
                    p = doc.add_paragraph()
                    p.paragraph_format.left_indent = Cm(0.5)
                    _add_run(p, f"{ch.get('label','')} ", size=10.5, bold=True, color="2D3A5C", mono=True)
                    _add_run(p, ch.get("text", ""), size=10.5, color="1A1F36")
                    p.paragraph_format.line_spacing = 1.55
            _section("3. 함정 분석", kd.trap_analysis)
            _section("4. 풀이 방법", kd.solution_method)
            doc.add_paragraph()

    # 결 — 학습 전략 + 대비
    if gyeol.strip():
        p = doc.add_paragraph()
        _add_run(p, "💡 학습 전략과 대비 방법", size=14, bold=True, color="1A1F36")
        for para in gyeol.split("\n\n"):
            para = para.strip()
            if not para:
                continue
            p = doc.add_paragraph()
            _add_run(p, para, size=10.5, color="1A1F36")
            p.paragraph_format.line_spacing = 1.65
            p.paragraph_format.space_after = Pt(8)

    # 푸터
    doc.add_paragraph()
    p = doc.add_paragraph()
    _add_run(p, f"수강 문의: {academy} ☎️ {phone}", size=11.5, bold=True, color="1A1F36")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# 텍스트 합본 — 블로그 본문 그대로 복붙용
# ─────────────────────────────────────────────────────────────
def build_blog_text(meta: ExamMeta, qs: list[Question],
                    gi: str, seung: str, gyeol: str,
                    captions: dict, killer_deeps: list[KillerDeep],
                    academy: str, phone: str) -> str:
    diff_label = derive_difficulty_label(qs)
    title = (meta.title or " ".join(filter(None, [
        meta.school, meta.grade, meta.exam_date[:7] if meta.exam_date else "",
        meta.exam_type, "영어 시험 분석",
    ])))
    out = [title, ""]
    out.append(f"총평 (난도: {diff_label})")
    out.append("")
    out.append(gi.strip())
    out.append("")
    if seung.strip():
        out.append("📊 출제 분석")
        out.append("")
        out.append(seung.strip())
        out.append("")
    cap_keys = [("type", "유형별 분포"), ("scope", "범위별 분포"),
                ("difficulty", "난이도 분포"), ("location", "문항 위치별 난이도")]
    for k, label in cap_keys:
        cap = (captions or {}).get(k, "").strip()
        if cap:
            out.append(f"[{label}]")
            out.append(cap)
            out.append("")
    if killer_deeps:
        out.append("🎯 변별의 분기점 — 어려운 문항 깊이 분석")
        out.append("")
        for kd in killer_deeps:
            out.append(f"{kd.no}번 · {kd.type} — {kd.headline}")
            out.append("")
            out.append("1. 어려운 지문 부분 풀어쓰기")
            out.append(kd.paraphrase_passage)
            out.append("")
            if kd.paraphrase_choices:
                out.append("2. 선지 풀어쓰기")
                for ch in kd.paraphrase_choices:
                    out.append(f"{ch.get('label', '')} {ch.get('text', '')}")
                out.append("")
            out.append("3. 함정 분석")
            out.append(kd.trap_analysis)
            out.append("")
            out.append("4. 풀이 방법")
            out.append(kd.solution_method)
            out.append("")
    if gyeol.strip():
        out.append("💡 학습 전략과 대비 방법")
        out.append("")
        out.append(gyeol.strip())
        out.append("")
    out.append(f"수강 문의: {academy} ☎️ {phone}")
    return "\n".join(out)


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


CHART_THEME_OPTIONS = ["editorial", "mono", "vivid"]
CHART_THEME_LABELS = {
    "editorial": "Editorial · 종이 톤",
    "mono":      "Mono · 흑백",
    "vivid":     "Vivid · 활발한 색",
}


def _init_state():
    _ss("meta", None)
    _ss("questions", [])
    _ss("killer_deeps", [])
    _ss("blog_body", {})       # gi/seung/gyeol/captions
    _ss("blog_text", "")
    _ss("blog_image", b"")
    _ss("blog_word", b"")
    _ss("uploaded_keys", [])
    _ss("academy", DEFAULT_ACADEMY)
    _ss("phone", DEFAULT_PHONE)
    _ss("chart_theme", "editorial")


def render_sidebar():
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
    st.markdown('<p class="section-label">차트 테마</p>', unsafe_allow_html=True)
    st.caption("보고서에 들어갈 차트 4종의 색감을 정합니다.")
    cur = _ss("chart_theme", "editorial")
    new_ct = st.radio(
        "chart_theme_radio",
        CHART_THEME_OPTIONS,
        index=CHART_THEME_OPTIONS.index(cur) if cur in CHART_THEME_OPTIONS else 0,
        format_func=lambda k: CHART_THEME_LABELS[k],
        label_visibility="collapsed",
        key="chart_theme_radio",
    )
    if new_ct != cur:
        _set_ss("chart_theme", new_ct)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<p class="section-label">학원 정보 (보고서 푸터)</p>', unsafe_allow_html=True)
    st.text_input("학원명", key=SS_PREFIX + "academy", value=_ss("academy", DEFAULT_ACADEMY))
    st.text_input("연락처", key=SS_PREFIX + "phone", value=_ss("phone", DEFAULT_PHONE))

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<p class="section-label">초기화</p>', unsafe_allow_html=True)
    if st.button("분석 결과 초기화", use_container_width=True):
        for k in ("meta", "questions", "killer_deeps", "blog_body",
                  "blog_text", "blog_image", "blog_word", "uploaded_keys"):
            if k == "meta":
                st.session_state[SS_PREFIX + k] = None
            elif k in ("questions", "killer_deeps", "uploaded_keys"):
                st.session_state[SS_PREFIX + k] = []
            elif k == "blog_body":
                st.session_state[SS_PREFIX + k] = {}
            elif k in ("blog_image", "blog_word"):
                st.session_state[SS_PREFIX + k] = b""
            else:
                st.session_state[SS_PREFIX + k] = ""
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
    rows = [{
        "no": q.no, "type": q.type, "difficulty": q.difficulty, "score": q.score,
        "is_subjective": q.is_subjective, "is_killer": q.is_killer,
        "scope": q.scope, "memo": q.memo,
    } for q in qs]
    return pd.DataFrame(rows)


def _df_to_questions(df: pd.DataFrame, prev_qs: list[Question]) -> list[Question]:
    """DataFrame 수정 사항을 prev_qs(passage_excerpt/choices 포함) 위에 덮어씀."""
    by_no = {q.no: q for q in prev_qs}
    qs = []
    for _, row in df.iterrows():
        try:
            no = int(row.get("no") or 0)
            base = by_no.get(no, Question(no=no))
            base.no = no
            base.type = str(row.get("type") or "").strip()
            base.difficulty = str(row.get("difficulty") or "중")
            base.score = float(row.get("score") or 0)
            base.is_subjective = bool(row.get("is_subjective") or False)
            base.is_killer = bool(row.get("is_killer") or False)
            base.scope = str(row.get("scope") or "").strip()
            base.memo = str(row.get("memo") or "").strip()
            qs.append(base)
        except Exception:
            continue
    qs.sort(key=lambda x: x.no)
    return qs


def render_main(api_key: str):
    if not api_key:
        st.error("OpenAI API Key 가 필요합니다.")
        return
    _init_state()
    chart_theme: ThemeName = _ss("chart_theme", "editorial")

    # ── §1. 업로드 & OCR ──
    st.markdown('<div class="section-mark">§ 1. 업로드 & OCR</div>', unsafe_allow_html=True)

    files = st.file_uploader(
        "시험지 이미지 업로드 (여러 페이지 동시 가능)",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        key="exam_uploader",
        help="시험지 한 부 전체를 페이지별 이미지로 올리세요.",
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
                    _set_ss("killer_deeps", [])
                    _set_ss("blog_body", {})
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

    # ── §2. 메타정보 확인 & 수정 ──
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
    st.caption("난이도·유형·배점·범위·어려운 문항 표시를 자유롭게 고치세요. 변경 즉시 반영됩니다.")

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
        _set_ss("questions", _df_to_questions(edited, qs))
        qs = _ss("questions")

    qa, qb, qc = st.columns(3)
    with qa:
        if st.button("어려운 문항 자동 표시", use_container_width=True,
                     help="난이도 '상' 우선. '상'이 없을 경우만 '중상+상위 배점' 보조 표시."):
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

    # ── §3. 보고서 생성 (단일 버튼) ──
    st.markdown('<div class="section-mark" style="margin-top:32px">§ 3. 블로그용 보고서 생성</div>',
                unsafe_allow_html=True)
    st.caption("메타정보 확인 후 이 버튼 한 번이면 끝까지 만듭니다 — 어려운 문항 깊이 분석 → 기·승·결 본문 → 차트 → Word + 이미지.")

    academy = _ss("academy", DEFAULT_ACADEMY)
    phone = _ss("phone", DEFAULT_PHONE)

    if st.button("블로그용 분석 보고서 생성", type="primary", use_container_width=True):
        if not qs:
            st.warning("문항 정보가 비어 있습니다. 먼저 OCR 을 실행하거나 표에 추가하세요.")
        else:
            try:
                with st.status("어려운 문항 깊이 분석 중...", expanded=True) as status:
                    kds = gen_killer_deep(api_key, meta, qs)
                    _set_ss("killer_deeps", kds)
                    status.update(label=f"깊이 분석 완료 — {len(kds)}문항")

                    status.update(label="블로그 본문 작성 중 (기·승·결 + 캡션)...")
                    body = gen_blog_body(api_key, meta, qs, kds, academy, phone,
                                         seed=hash((meta.title, meta.exam_date)) & 0xFFFFFFFF)
                    _set_ss("blog_body", body)

                    status.update(label="차트 생성 중...")
                    charts = {
                        "type":       chart_type_distribution(qs, chart_theme),
                        "scope":      chart_scope_distribution(qs, chart_theme),
                        "difficulty": chart_difficulty_distribution(qs, chart_theme),
                        "killer_map": chart_killer_map(qs, chart_theme),
                    }

                    status.update(label="Word & 이미지 생성 중...")
                    word_b = build_word_report(meta, qs, body["gi"], body["seung"], body["gyeol"],
                                               body["captions"], kds, charts, academy, phone)
                    img_b = render_blog_image(meta, qs, body["gi"], body["seung"], body["gyeol"],
                                              body["captions"], kds, charts, academy, phone, theme=chart_theme)
                    text_b = build_blog_text(meta, qs, body["gi"], body["seung"], body["gyeol"],
                                             body["captions"], kds, academy, phone)
                    _set_ss("blog_word", word_b)
                    _set_ss("blog_image", img_b)
                    _set_ss("blog_text", text_b)
                    status.update(label="완료 — Word + 이미지 + 텍스트 준비됨", state="complete")
            except openai.AuthenticationError:
                st.error("OpenAI API Key 가 유효하지 않습니다.")
            except Exception as e:
                st.error(f"분석 실패: {e}")

    # ── §4. 결과 ──
    blog_body = _ss("blog_body", {})
    blog_word = _ss("blog_word", b"")
    blog_image = _ss("blog_image", b"")
    blog_text = _ss("blog_text", "")
    kds = _ss("killer_deeps", [])

    if blog_image or blog_body:
        st.markdown('<div class="section-mark" style="margin-top:32px">§ 4. 결과 미리보기 & 다운로드</div>',
                    unsafe_allow_html=True)
        prev_col, dl_col = st.columns([3, 2])
        with prev_col:
            if blog_image:
                st.markdown("##### 블로그 이미지 미리보기")
                st.image(blog_image, use_container_width=True)
            if blog_text:
                with st.expander("블로그용 텍스트 보기 (복붙용)", expanded=False):
                    st.text_area("blog_text_preview", value=blog_text, height=400,
                                 label_visibility="collapsed")
        with dl_col:
            st.markdown("##### 다운로드")
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            base = re.sub(r'[\\/*?:"<>|]', "_",
                          meta.title or f"{meta.school}_{meta.grade}_{meta.subject}_{meta.exam_type}")
            if blog_word:
                st.download_button(
                    "Word 보고서 (.docx)", data=blog_word,
                    file_name=f"{base}_분석_{ts}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary", use_container_width=True,
                )
            if blog_image:
                st.download_button(
                    "블로그 이미지 (.png)", data=blog_image,
                    file_name=f"{base}_분석_{ts}.png", mime="image/png",
                    type="primary", use_container_width=True,
                )
            if blog_text:
                st.download_button(
                    "블로그 텍스트 (.txt)", data=blog_text.encode("utf-8"),
                    file_name=f"{base}_분석_{ts}.txt", mime="text/plain",
                    use_container_width=True,
                )

        if kds:
            st.markdown('<div class="section-mark" style="margin-top:32px">어려운 문항 깊이 분석</div>',
                        unsafe_allow_html=True)
            for kd in kds:
                choices_html = ""
                if kd.paraphrase_choices:
                    choices_html = "<div style='margin-top:8px'><b style='color:var(--text-accent)'>2. 선지 풀어쓰기</b><div style='margin-left:14px;margin-top:4px'>"
                    for ch in kd.paraphrase_choices:
                        choices_html += f"<div style='margin:3px 0'><b>{ch.get('label','')}</b> {ch.get('text','')}</div>"
                    choices_html += "</div></div>"
                st.markdown(
                    f"<div class='card'>"
                    f"<span class='killer-flag'>#{kd.no}</span>"
                    f"<span style='color:var(--text-heading);font-weight:700'>{kd.type}</span>"
                    f"<span style='color:var(--text-muted);font-size:13px;margin-left:8px'>"
                    f"— {kd.headline}</span>"
                    f"<div style='margin-top:12px;font-size:14px;line-height:1.7'>"
                    f"<div><b style='color:var(--text-accent)'>1. 어려운 지문 부분 풀어쓰기</b><div style='margin-left:14px;margin-top:4px'>{kd.paraphrase_passage}</div></div>"
                    f"{choices_html}"
                    f"<div style='margin-top:8px'><b style='color:var(--text-accent)'>3. 함정 분석</b><div style='margin-left:14px;margin-top:4px'>{kd.trap_analysis}</div></div>"
                    f"<div style='margin-top:8px'><b style='color:var(--text-accent)'>4. 풀이 방법</b><div style='margin-left:14px;margin-top:4px'>{kd.solution_method}</div></div>"
                    f"</div></div>",
                    unsafe_allow_html=True,
                )
