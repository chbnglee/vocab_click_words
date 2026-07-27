from __future__ import annotations

import io
import importlib.util
import json
import re
import time
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from google import genai
from google.genai import types as genai_types
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Side, Border
from openpyxl.utils import get_column_letter


APP_DIR = Path(__file__).resolve().parent
CEFR_PATH = APP_DIR / "lcms_cefr.csv"
ANALYZER_PATH = APP_DIR / "vocab_analyzer.py"

INPUT_COLUMNS = ["ID", "Title", "Level", "Base Text"]
CEFR_ORDER = ["Pre A1", "A1", "A2", "B1", "B2", "C1", "C2"]
LEVEL_MAP = {
    "1": "Pre A1",
    "2": "A1",
    "3": "A2",
    "4": "B1",
    "pre-a1": "Pre A1",
    "pre a1": "Pre A1",
    "pre_a1": "Pre A1",
    "prea1": "Pre A1",
    "a1": "A1",
    "a2": "A2",
    "b1": "B1",
    "b2": "B2",
    "c1": "C1",
    "c2": "C2",
}
BOOK_MOODS = [
    "Warm",
    "Playful",
    "Adventurous",
    "Mysterious",
    "Emotional",
    "Calm",
    "Humorous",
    "Dramatic",
    "Inspirational",
]
STORY_CATEGORIES = [
    "Classic",
    "Emotion",
    "Growth & Self",
    "Family",
    "Friends & School",
    "Hobbies & Sports",
    "Career & Dreams",
    "History & Heroes",
    "World",
    "Nature & Animals",
    "Body",
    "Science & Space",
    "Technology",
    "Arts & Music",
]
PREFERRED_MODELS = [
    "gemini-3.5-flash",
    "gemini-3.1-flash-lite",
    "gemini-flash-latest",
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
]
DEFAULT_MODEL = PREFERRED_MODELS[0]
MODEL_SESSION_KEY = "story_vocab_available_models"


@st.cache_resource
def load_vocab_analyzer():
    spec = importlib.util.spec_from_file_location("story_vocab_analyzer", ANALYZER_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load analyzer: {ANALYZER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


st.set_page_config(
    page_title="Story Info + Vocab & Click Words",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed",
)

vocab_analyzer = load_vocab_analyzer()

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.4rem; }
    div[data-testid="stMetricValue"] { font-size: 1.35rem; }
    .small-note { color: #5f6368; font-size: 0.9rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


def normalize_model_name(model_name: str) -> str:
    normalized = str(model_name or "").strip()
    if normalized.startswith("models/"):
        normalized = normalized.removeprefix("models/")
    return normalized


def sort_model_names(models: list[str]) -> list[str]:
    preferred_order = {model: idx for idx, model in enumerate(PREFERRED_MODELS)}
    normalized_models: list[str] = []
    for model in models:
        normalized = normalize_model_name(model)
        if normalized and normalized not in normalized_models:
            normalized_models.append(normalized)
    return sorted(
        normalized_models,
        key=lambda model: (preferred_order.get(model, len(PREFERRED_MODELS)), model),
    )


def list_available_text_models(api_key: str) -> list[str]:
    client = genai.Client(api_key=api_key.strip())
    models: list[str] = []
    excluded_fragments = ["aqa", "audio", "embedding", "imagen", "live", "tts", "veo"]
    for model in client.models.list():
        name = normalize_model_name(getattr(model, "name", ""))
        if not name or "gemini" not in name.lower():
            continue
        lower_name = name.lower()
        if any(fragment in lower_name for fragment in excluded_fragments):
            continue
        supported_actions = getattr(model, "supported_actions", None)
        if supported_actions is None:
            supported_actions = getattr(model, "supported_generation_methods", [])
        if supported_actions and not any(
            str(action).lower() in {"generatecontent", "generate_content"}
            for action in supported_actions
        ):
            continue
        models.append(name)
    return sort_model_names(models)


def normalize_level(raw: Any) -> str:
    value = str(raw or "").strip()
    if not value:
        return ""
    key = value.lower().replace("-", " ").replace("_", " ")
    key = re.sub(r"\s+", " ", key).strip()
    mapped = LEVEL_MAP.get(value.lower()) or LEVEL_MAP.get(key)
    if mapped:
        return mapped
    upper = value.upper().replace("-", " ")
    if upper == "PRE A1":
        return "Pre A1"
    if upper in CEFR_ORDER:
        return upper
    return value


def adjacent_level(level: str, delta: int) -> str:
    normalized = normalize_level(level)
    if normalized not in CEFR_ORDER:
        return ""
    idx = CEFR_ORDER.index(normalized)
    idx = max(0, min(len(CEFR_ORDER) - 1, idx + delta))
    return CEFR_ORDER[idx]


def count_story_words(text: str) -> int:
    body = re.sub(r"#SC\d+\b", " ", str(text or ""))
    return len(re.findall(r"\b[a-zA-Z0-9']+\b", body))


def count_story_scenes(text: str) -> int:
    return len(re.findall(r"#SC\d+\b", str(text or "")))


def clean_json_text(text: str) -> str:
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return text.strip()


def describe_gemini_error(message: str) -> str:
    lower = message.lower()
    if any(token in lower for token in ["not_found", "model not", "404"]):
        return "선택한 Gemini 모델을 현재 API 키에서 사용할 수 없을 가능성이 큽니다. 모델 목록을 확인해 다른 모델을 선택해 주세요."
    if any(token in lower for token in ["quota", "rate limit", "429", "billing"]):
        return "Gemini API 사용량, 쿼터, 결제 또는 속도 제한 문제일 가능성이 큽니다."
    if any(token in lower for token in ["invalid api key", "permission_denied", "401", "403"]):
        return "Gemini API 키가 잘못되었거나 해당 키에 권한이 부족할 가능성이 큽니다."
    return "Gemini API 호출 중 오류가 발생했습니다. 원문 메시지를 확인해 주세요."


def create_template_workbook() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Story_Confirmed"
    ws.append(INPUT_COLUMNS)
    ws.append(
        [
            "OG0001",
            "Sample Title",
            "A2",
            "#SC01\nA small turtle finds a shiny seed.\n#SC02\nThe seed begins to glow.",
        ]
    )
    widths = [14, 30, 14, 80]
    header_fill = PatternFill("solid", start_color="1F3864")
    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    for col_idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width
        cell = ws.cell(1, col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2, max_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def read_story_input(uploaded_file) -> pd.DataFrame:
    if uploaded_file.name.lower().endswith(".csv"):
        return pd.read_csv(uploaded_file, dtype=str).fillna("")
    return pd.read_excel(uploaded_file, dtype=str).fillna("")


def validate_story_df(df: pd.DataFrame) -> list[str]:
    missing = [col for col in INPUT_COLUMNS if col not in df.columns]
    if missing:
        return [f"필수 컬럼 없음: {', '.join(missing)}"]
    errors: list[str] = []
    for idx, row in df.iterrows():
        sid = str(row.get("ID", "")).strip() or f"{idx + 2}행"
        title = str(row.get("Title", "")).strip()
        if not str(row.get("Base Text", "")).strip():
            errors.append(f"{sid} / {title}: Base Text가 비어 있음")
        if normalize_level(row.get("Level", "")) not in CEFR_ORDER:
            errors.append(f"{sid} / {title}: Level은 1-4 또는 Pre A1/A1/A2/B1/B2/C1/C2 중 하나여야 함")
    return errors


def build_story_prompt(
    story_id: str,
    title: str,
    input_level: str,
    easy_target: str,
    difficult_target: str,
    base_story: str,
    word_count: int,
    scene_count: int,
) -> str:
    fallback_level = input_level or "the independently detected CEFR level"
    easy_rule = easy_target or f"one level below {fallback_level}"
    difficult_rule = difficult_target or f"one level above {fallback_level}"
    categories = " | ".join(STORY_CATEGORIES)
    moods = " | ".join(BOOK_MOODS)
    return f"""You are an expert children's story editor and English language teacher.

Analyze the story and return ONLY a valid JSON object. Do not use markdown fences.

Required JSON keys:
{{
  "detected_level": "Pre A1 | A1 | A2 | B1 | B2 | C1 | C2",
  "detected_level_rationale": "1-2 concise sentences with textual evidence",
  "lexile": "estimated Lexile as a number with L, e.g. 520L",
  "lexile_rationale": "1 concise sentence based on sentence length and vocabulary difficulty",
  "category": "one category exactly from the list",
  "book_mood": "one mood exactly from the list",
  "book_info": "1-2 short declarative summary sentences for young learners, no spoilers, end with a period",
  "keywords": ["6-12 lowercase content words, no proper nouns"],
  "intro": "4-6 sentence spoken intro script in the main character's voice",
  "easy_version": "rewritten story preserving every #SC marker",
  "difficult_version": "rewritten story preserving every #SC marker",
  "learning_focus_1": "Type/Focus/Items/Prompt 1-3 block",
  "learning_focus_2": "Type/Focus/Items/Prompt 1-3 block",
  "learning_focus_3": "Type/Focus/Items/Prompt 1-3 block"
}}

Story ID: {story_id}
Title: {title}
Input production level for rewriting: {input_level or "(blank)"}
Easy rewrite target: {easy_rule}
Difficult rewrite target: {difficult_rule}
Known word count: {word_count}
Known scene count: {scene_count}

Important separation of level logic:
- Detect "detected_level" and "lexile" independently from the Base Story only.
- Do not force detected_level to match the input production level.
- Use the input production level only to decide the Easy/Difficult rewrite targets.
- Preserve all #SC markers exactly. Do not remove, merge, or invent scenes.
- Easy version should be clearly simpler than the input production level.
- Difficult version should be richer than the input production level but keep the same plot.

Category choices:
{categories}

Book mood choices:
{moods}

Base Story:
=== BASE STORY START ===
{base_story}
=== BASE STORY END ===
"""


def call_story_info(
    client,
    model_name: str,
    row: pd.Series,
    max_retries: int = 3,
) -> dict[str, Any] | None:
    story_id = str(row["ID"]).strip()
    title = str(row["Title"]).strip()
    base_story = str(row["Base Text"]).strip()
    input_level = normalize_level(row.get("Level", ""))
    word_count = count_story_words(base_story)
    scene_count = count_story_scenes(base_story)
    prompt = build_story_prompt(
        story_id=story_id,
        title=title,
        input_level=input_level,
        easy_target=adjacent_level(input_level, -1),
        difficult_target=adjacent_level(input_level, 1),
        base_story=base_story,
        word_count=word_count,
        scene_count=scene_count,
    )
    full_model_name = model_name if model_name.startswith("models/") else f"models/{model_name}"

    for attempt in range(1, max_retries + 1):
        try:
            response = client.models.generate_content(
                model=full_model_name,
                contents=prompt,
                config=genai_types.GenerateContentConfig(
                    temperature=0.45,
                    max_output_tokens=16000,
                    response_mime_type="application/json",
                ),
            )
            parsed = json.loads(clean_json_text(response.text))
            return normalize_story_info(parsed, row)
        except json.JSONDecodeError as exc:
            last_error = f"JSONDecodeError: {exc}"
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
            if any(token in last_error.lower() for token in ["401", "403", "429", "quota", "invalid api key"]):
                st.session_state["last_story_info_error"] = last_error
                return None
        if attempt < max_retries:
            time.sleep(4 * attempt)

    st.session_state["last_story_info_error"] = last_error
    return None


def normalize_story_info(parsed: dict[str, Any], row: pd.Series) -> dict[str, Any]:
    base_story = str(row["Base Text"]).strip()
    input_level = normalize_level(row.get("Level", ""))
    keywords = parsed.get("keywords", [])
    if isinstance(keywords, str):
        keywords = [item.strip() for item in keywords.split(",") if item.strip()]
    elif not isinstance(keywords, list):
        keywords = []

    detected_level = normalize_level(parsed.get("detected_level", ""))
    if detected_level not in CEFR_ORDER:
        detected_level = ""

    mood = str(parsed.get("book_mood", "")).strip()
    if mood not in BOOK_MOODS:
        mood = mood.title() if mood else ""

    category = str(parsed.get("category", "")).strip()
    if category not in STORY_CATEGORIES:
        category = ""

    return {
        "id": str(row["ID"]).strip(),
        "title": str(row["Title"]).strip(),
        "input_level": input_level,
        "base_text": base_story,
        "detected_level": detected_level,
        "detected_level_rationale": str(parsed.get("detected_level_rationale", "")).strip(),
        "lexile": str(parsed.get("lexile", "")).strip(),
        "lexile_rationale": str(parsed.get("lexile_rationale", "")).strip(),
        "word_count": count_story_words(base_story),
        "scene_count": count_story_scenes(base_story),
        "category": category,
        "book_mood": mood,
        "book_info": str(parsed.get("book_info", "")).strip(),
        "keywords": [str(item).strip().lower() for item in keywords if str(item).strip()],
        "intro": str(parsed.get("intro", "")).strip(),
        "easy_version": str(parsed.get("easy_version", "")).strip(),
        "difficult_version": str(parsed.get("difficult_version", "")).strip(),
        "learning_focus_1": str(parsed.get("learning_focus_1", "")).strip(),
        "learning_focus_2": str(parsed.get("learning_focus_2", "")).strip(),
        "learning_focus_3": str(parsed.get("learning_focus_3", "")).strip(),
    }


def call_vocab_analysis(client, model_name: str, story_info: dict[str, Any]) -> dict[str, Any]:
    normal = story_info["base_text"]
    easy = story_info.get("easy_version", "")
    difficult = story_info.get("difficult_version", "")
    raw = vocab_analyzer.call_gemini(
        client,
        model_name,
        vocab_analyzer.build_prompt(story_info["title"], normal, easy, difficult),
    )
    if not raw:
        raw = {"vocab": [], "normal_vocab": [], "easy_vocab": [], "difficult_vocab": []}

    detected_level = story_info.get("detected_level") or normalize_level(raw.get("cefr", ""))
    raw["cefr"] = detected_level
    raw["cefr_rationale"] = story_info.get("detected_level_rationale", raw.get("cefr_rationale", ""))
    raw["lexile"] = story_info.get("lexile", raw.get("lexile", ""))
    raw["lexile_rationale"] = story_info.get("lexile_rationale", raw.get("lexile_rationale", ""))
    raw = vocab_analyzer.apply_db_vocab_filter(raw, normal, easy, difficult)
    raw["id"] = story_info["id"]
    raw["title"] = story_info["title"]
    return raw


def vocab_to_cell(vocab: Any) -> str:
    return vocab_analyzer.vocab_to_cell(vocab)


def result_is_complete(item: dict[str, Any] | None) -> bool:
    if not item:
        return False
    story_info = item.get("story_info", {})
    vocab = item.get("vocab", {})
    return bool(story_info.get("detected_level") and story_info.get("easy_version") and vocab.get("normal_vocab"))


def load_checkpoint(uploaded_file) -> dict[str, Any]:
    if uploaded_file is None:
        return {}
    return json.load(uploaded_file)


def checkpoint_to_bytes(results_by_id: dict[str, Any]) -> bytes:
    return json.dumps(results_by_id, ensure_ascii=False, indent=2).encode("utf-8")


def style_header(cell, fill_color: str):
    cell.fill = PatternFill("solid", start_color=fill_color)
    cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_table_sheet(
    wb: Workbook,
    title: str,
    headers: list[str],
    rows: list[list[Any]],
    widths: list[int],
    fill_color: str,
):
    ws = wb.create_sheet(title)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.row_dimensions[1].height = 32
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(1, col_idx, header)
        style_header(cell, fill_color)
        cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = widths[col_idx - 1]
    fills = [PatternFill("solid", start_color="FFFFFF"), PatternFill("solid", start_color="F7F9FC")]
    for row_idx, values in enumerate(rows, 2):
        fill = fills[row_idx % 2]
        for col_idx, value in enumerate(values, 1):
            cell = ws.cell(row_idx, col_idx, "" if pd.isna(value) else value)
            cell.fill = fill
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center" if col_idx in {1, 3, 5, 7, 8, 9} else "left",
                vertical="top",
                wrap_text=True,
            )
        ws.row_dimensions[row_idx].height = 52
    ws.freeze_panes = "A2"
    return ws


def build_output_workbook(source_df: pd.DataFrame, results_by_id: dict[str, Any]) -> bytes:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    combined_headers = [
        "ID",
        "Title",
        "Input Level",
        "Base Text",
        "Detected CEFR",
        "CEFR Rationale",
        "Lexile",
        "Lexile Rationale",
        "Word Count",
        "Scene Count",
        "Category",
        "Book Mood",
        "Summary",
        "Keywords",
        "Intro Script",
        "Easy Version",
        "Difficult Version",
        "Vocab",
        "Normal Words",
        "Easy Words",
        "Difficult Words",
    ]
    combined_rows: list[list[Any]] = []
    story_rows: list[list[Any]] = []
    vocab_rows: list[list[Any]] = []

    for _, row in source_df.iterrows():
        sid = str(row["ID"])
        item = results_by_id.get(sid, {})
        story = item.get("story_info", {})
        vocab = item.get("vocab", {})
        keywords = ", ".join(story.get("keywords", []))
        combined_rows.append(
            [
                sid,
                row["Title"],
                story.get("input_level", row.get("Level", "")),
                story.get("base_text", row.get("Base Text", "")),
                story.get("detected_level", ""),
                story.get("detected_level_rationale", ""),
                story.get("lexile", ""),
                story.get("lexile_rationale", ""),
                story.get("word_count", ""),
                story.get("scene_count", ""),
                story.get("category", ""),
                story.get("book_mood", ""),
                story.get("book_info", ""),
                keywords,
                story.get("intro", ""),
                story.get("easy_version", ""),
                story.get("difficult_version", ""),
                vocab_to_cell(vocab.get("vocab", [])),
                vocab_to_cell(vocab.get("normal_vocab", [])),
                vocab_to_cell(vocab.get("easy_vocab", [])),
                vocab_to_cell(vocab.get("difficult_vocab", [])),
            ]
        )
        story_rows.append(combined_rows[-1][:17])
        vocab_rows.append(
            [
                sid,
                row["Title"],
                story.get("detected_level", ""),
                story.get("lexile", ""),
                vocab_to_cell(vocab.get("vocab", [])),
                vocab_to_cell(vocab.get("normal_vocab", [])),
                vocab_to_cell(vocab.get("easy_vocab", [])),
                vocab_to_cell(vocab.get("difficult_vocab", [])),
            ]
        )

    write_table_sheet(
        wb,
        "Combined",
        combined_headers,
        combined_rows,
        [12, 26, 13, 58, 13, 42, 12, 42, 12, 12, 18, 16, 46, 36, 48, 58, 58, 32, 46, 46, 46],
        "1F3864",
    )
    write_table_sheet(
        wb,
        "Story_Info",
        combined_headers[:17],
        story_rows,
        [12, 26, 13, 58, 13, 42, 12, 42, 12, 12, 18, 16, 46, 36, 48, 58, 58],
        "385723",
    )
    write_table_sheet(
        wb,
        "Vocab_Click_Words",
        ["ID", "Title", "Detected CEFR", "Lexile", "Vocab", "Normal Words", "Easy Words", "Difficult Words"],
        vocab_rows,
        [12, 28, 14, 12, 36, 54, 54, 54],
        "7030A0",
    )

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


@st.cache_data
def load_cefr_wordlist(path: str, mtime: float) -> pd.DataFrame:
    _ = mtime
    df = pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")
    if "cefr_level" in df.columns:
        order = {level: idx for idx, level in enumerate(CEFR_ORDER)}
        df["cefr_sort"] = df["cefr_level"].map(order).fillna(999)
        df = df.sort_values(["cefr_sort", "word"]).drop(columns=["cefr_sort"])
    return df


def guide_tab():
    st.subheader("작업 흐름")
    st.markdown(
        """
        1. `Story_Confirmed_Template.xlsx`를 내려받아 `ID`, `Title`, `Level`, `Base Text`를 입력합니다.
        2. `Level`은 Easy/Difficult 버전 생성 기준으로만 사용합니다.
        3. `Detected CEFR`와 `Lexile`은 Base Text를 API가 별도로 추정합니다.
        4. Vocab & Click Words는 추정 CEFR을 기준으로 LCMS CEFR DB를 먼저 적용하고, API는 핵심/클릭 어휘 보완 판단에만 사용합니다.
        """
    )
    st.divider()
    st.subheader("LCMS CEFR 단어 리스트")
    if not CEFR_PATH.exists():
        st.warning("lcms_cefr.csv 파일을 찾을 수 없습니다.")
        return
    cefr_df = load_cefr_wordlist(str(CEFR_PATH), CEFR_PATH.stat().st_mtime)
    cols = st.columns(3)
    cols[0].metric("전체 단어", f"{len(cefr_df):,}")
    cols[1].metric("CEFR 단계", f"{cefr_df['cefr_level'].nunique() if 'cefr_level' in cefr_df else 0}")
    cols[2].metric("출처", "LCMS_0727 단어메타")
    display_columns = [col for col in ["word", "cefr_level", "source", "pos", "category"] if col in cefr_df.columns]
    st.dataframe(
        cefr_df[display_columns].rename(
            columns={
                "word": "단어",
                "cefr_level": "CEFR",
                "source": "출처",
                "pos": "품사",
                "category": "카테고리",
            }
        ),
        use_container_width=True,
        height=420,
        hide_index=True,
    )


def extraction_tab():
    st.subheader("Story Info + Vocab 분석")
    st.download_button(
        "Story_Confirmed_Template.xlsx 다운로드",
        create_template_workbook(),
        file_name="Story_Confirmed_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    uploaded_story = st.file_uploader(
        "분석할 story_confirmed 파일 업로드",
        type=["xlsx", "csv"],
        accept_multiple_files=False,
    )
    if uploaded_story is None:
        st.info("템플릿을 내려받아 작성한 뒤 업로드해 주세요.")
        return

    checkpoint_file = st.file_uploader(
        "이전 checkpoint JSON 업로드 (선택)",
        type=["json"],
        accept_multiple_files=False,
    )

    try:
        story_df = read_story_input(uploaded_story)
    except Exception as exc:
        st.error(f"입력 파일을 읽지 못했습니다: {exc}")
        return

    errors = validate_story_df(story_df)
    preview_cols = [col for col in INPUT_COLUMNS if col in story_df.columns]
    metric_cols = st.columns(4)
    metric_cols[0].metric("스토리", f"{len(story_df):,}")
    metric_cols[1].metric("필수 컬럼", f"{len(preview_cols)}/{len(INPUT_COLUMNS)}")
    metric_cols[2].metric("입력 오류", f"{len(errors):,}")
    metric_cols[3].metric("LCMS DB", "있음" if CEFR_PATH.exists() else "없음")

    with st.expander("입력 미리보기", expanded=True):
        st.dataframe(story_df[preview_cols].head(20), use_container_width=True, hide_index=True)

    if errors:
        st.error("입력 파일을 먼저 수정해야 합니다.")
        st.code("\n".join(errors[:30]), language="text")
        return

    checkpoint = load_checkpoint(checkpoint_file)
    completed_ids = {sid for sid, item in checkpoint.items() if result_is_complete(item)}
    labels = [f"{row['ID']} | {row['Title']}" for _, row in story_df.iterrows()]
    selected_labels = st.multiselect(
        "분석 대상",
        labels,
        default=[label for label in labels if label.split(" | ", 1)[0] not in completed_ids] or labels,
    )
    selected_ids = {label.split(" | ", 1)[0] for label in selected_labels}
    selected_df = story_df[story_df["ID"].astype(str).isin(selected_ids)].copy()

    api_key = st.text_input(
        "Gemini API Key",
        type="password",
        placeholder="분석 실행 시에만 사용하며 앱에 저장하지 않습니다.",
    )
    with st.expander("고급 설정"):
        check_models_clicked = st.button(
            "API 키로 사용 가능한 모델 확인",
            disabled=not api_key.strip(),
            use_container_width=True,
        )
        if check_models_clicked:
            try:
                available_models = list_available_text_models(api_key)
            except Exception as exc:
                message = f"{type(exc).__name__}: {exc}"
                st.error(describe_gemini_error(message))
                st.code(message, language="text")
            else:
                st.session_state[MODEL_SESSION_KEY] = available_models
                st.success(f"사용 가능한 텍스트 모델 {len(available_models):,}개를 확인했습니다.")

        model_options = st.session_state.get(MODEL_SESSION_KEY) or PREFERRED_MODELS
        default_index = model_options.index(DEFAULT_MODEL) if DEFAULT_MODEL in model_options else 0
        selected_model = st.selectbox("모델", options=model_options, index=default_index)
        custom_model = st.text_input("직접 입력할 모델명", placeholder="예: gemini-2.5-flash")
        model_name = normalize_model_name(custom_model) or selected_model
        st.caption(f"현재 실행 모델: {model_name}")

    st.caption(
        f"checkpoint 재사용 가능: {len(completed_ids):,}개 / 현재 선택: {len(selected_df):,}개"
    )
    run_clicked = st.button(
        "Story Info + Vocab 분석 실행",
        type="primary",
        disabled=selected_df.empty,
        use_container_width=True,
    )

    if run_clicked:
        if not api_key.strip():
            st.error("Gemini API Key가 필요합니다.")
            return
        if not model_name.strip():
            st.error("Gemini 모델명이 필요합니다.")
            return

        try:
            client = genai.Client(api_key=api_key.strip())
        except Exception as exc:
            st.error(f"Gemini 클라이언트를 만들지 못했습니다: {type(exc).__name__}: {exc}")
            return

        results_by_id: dict[str, Any] = dict(checkpoint)
        progress = st.progress(0)
        log_box = st.empty()

        for idx, (_, row) in enumerate(selected_df.iterrows(), 1):
            sid = str(row["ID"])
            title = str(row["Title"])
            if result_is_complete(results_by_id.get(sid)):
                progress.progress(idx / len(selected_df))
                continue

            log_box.info(f"{idx}/{len(selected_df)} Story Info 생성 중: {sid} / {title}")
            story_info = call_story_info(client, model_name, row)
            if not story_info:
                message = st.session_state.get("last_story_info_error", "")
                st.warning(f"{sid} Story Info 생성 실패")
                if message:
                    st.error(describe_gemini_error(message))
                    st.code(message, language="text")
                progress.progress(idx / len(selected_df))
                continue

            log_box.info(f"{idx}/{len(selected_df)} Vocab & Click Words 분석 중: {sid} / {title}")
            try:
                vocab_result = call_vocab_analysis(client, model_name, story_info)
            except Exception as exc:
                message = f"{type(exc).__name__}: {exc}"
                st.warning(f"{sid} Vocab 분석 중 오류가 발생했습니다. Story Info 결과는 유지합니다.")
                st.error(describe_gemini_error(message))
                st.code(message, language="text")
                vocab_result = {"id": sid, "title": title}

            results_by_id[sid] = {"story_info": story_info, "vocab": vocab_result}
            progress.progress(idx / len(selected_df))

        output_bytes = build_output_workbook(story_df, results_by_id)
        checkpoint_bytes = checkpoint_to_bytes(results_by_id)
        st.session_state["story_vocab_output_bytes"] = output_bytes
        st.session_state["story_vocab_checkpoint_bytes"] = checkpoint_bytes
        st.session_state["story_vocab_result_count"] = sum(
            1 for item in results_by_id.values() if result_is_complete(item)
        )
        log_box.success(f"완료: {st.session_state['story_vocab_result_count']:,}개 결과")

    if "story_vocab_output_bytes" in st.session_state:
        download_cols = st.columns(2)
        with download_cols[0]:
            st.download_button(
                "Story_Info_Vocab_Analysis.xlsx 다운로드",
                st.session_state["story_vocab_output_bytes"],
                file_name="Story_Info_Vocab_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with download_cols[1]:
            st.download_button(
                "checkpoint JSON 다운로드",
                st.session_state["story_vocab_checkpoint_bytes"],
                file_name="story_info_vocab_checkpoint.json",
                mime="application/json",
                use_container_width=True,
            )


st.title("Story Info + Vocab & Click Words")
tab_guide, tab_extract = st.tabs(["가이드", "통합 분석"])

with tab_guide:
    guide_tab()

with tab_extract:
    extraction_tab()
