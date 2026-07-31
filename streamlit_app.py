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

PLATFORM_LEVEL_COLUMN = "Platform Level"
LEGACY_LEVEL_COLUMN = "Level"
INPUT_COLUMNS = ["ID", "Title", PLATFORM_LEVEL_COLUMN, "Base Text"]
REQUIRED_STORY_COLUMNS = ["ID", "Title", "Base Text"]
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
PLATFORM_LEVEL_TO_MAX_CEFR = {
    "1": "A2",
    "2": "B1",
    "3": "B2",
    "4": "C1",
}
BOOK_MOODS = [
    "Exciting",
    "Playful",
    "Warm",
    "Loving",
    "Proud",
    "Successful",
    "Inspired",
    "Fierce",
    "Competitive",
    "Brave",
    "Strict",
    "Touching",
    "Healing",
    "Lonely",
    "Longing",
    "Serious",
    "Sensitive",
    "Calm",
    "Peaceful",
    "Thoughtful",
    "Curious",
    "Adventurous",
    "Righteous",
]
STORY_CATEGORIES = [
    "Classics",
    "Emotion",
    "Growth",
    "Self",
    "Family",
    "Friends",
    "School",
    "Hobbies",
    "Sports",
    "Dreams",
    "History",
    "Heroes",
    "World",
    "Nature",
    "Animals",
    "Science",
    "Space",
    "Technology",
    "Arts",
    "Music",
    "Fantasy",
    "Body",
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
    page_title="Story Info.",
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
    .click-rule-table { width: 100%; border-collapse: collapse; font-size: 0.94rem; }
    .click-rule-table th { background: #E6EEF7; text-align: center; font-weight: 700; }
    .click-rule-table th, .click-rule-table td { border: 1px solid #D0D7DE; padding: 0.62rem 0.7rem; vertical-align: middle; }
    .click-rule-table td:first-child { text-align: center; font-weight: 700; white-space: nowrap; }
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


def get_platform_level(row: pd.Series) -> str:
    for column in (PLATFORM_LEVEL_COLUMN, LEGACY_LEVEL_COLUMN):
        if column in row:
            value = str(row.get(column, "")).strip()
            if value:
                return value
    return ""


def get_rewrite_level(row: pd.Series) -> str:
    return normalize_level(get_platform_level(row))


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


def platform_level_key(raw: Any) -> str:
    value = str(raw or "").strip()
    if re.fullmatch(r"\d+\.0+", value):
        value = str(int(float(value)))
    return value


def get_platform_max_cefr(platform_level: Any) -> str:
    key = platform_level_key(platform_level)
    if key in PLATFORM_LEVEL_TO_MAX_CEFR:
        return PLATFORM_LEVEL_TO_MAX_CEFR[key]
    normalized = normalize_level(key)
    return normalized if normalized in CEFR_ORDER else ""


@st.cache_data
def load_platform_cefr_index(path: str, mtime: float) -> tuple[dict[str, dict[str, Any]], dict[str, str]]:
    _ = mtime
    entries: dict[str, dict[str, Any]] = {}
    form_to_word: dict[str, str] = {}
    csv_path = Path(path)
    if not csv_path.exists():
        return entries, form_to_word

    df = pd.read_csv(csv_path, dtype=str, encoding="utf-8-sig").fillna("")
    for _, row in df.iterrows():
        word = vocab_analyzer.normalize_vocab_item(row.get("word", ""))
        level = normalize_level(row.get("cefr_level", ""))
        if not word or level not in CEFR_ORDER:
            continue
        entry = entries.setdefault(word, {"level": level, "forms": set(), "pos": set(), "synonyms": []})
        if CEFR_ORDER.index(level) < CEFR_ORDER.index(entry["level"]):
            entry["level"] = level
        pos = str(row.get("pos", "")).strip()
        if pos:
            entry["pos"].add(pos.lower())
        entry["forms"].add(word)
        for form in re.split(r"[,;]", str(row.get("forms", ""))):
            normalized_form = vocab_analyzer.normalize_vocab_item(form)
            if normalized_form:
                entry["forms"].add(normalized_form)
        for synonym in re.split(r"[,;]", str(row.get("synonyms", ""))):
            normalized_synonym = vocab_analyzer.normalize_vocab_item(synonym)
            if normalized_synonym and normalized_synonym != word and normalized_synonym not in entry["synonyms"]:
                entry["synonyms"].append(normalized_synonym)

    for word, entry in entries.items():
        for form in entry["forms"]:
            previous = form_to_word.get(form)
            if previous is None or CEFR_ORDER.index(entry["level"]) < CEFR_ORDER.index(entries[previous]["level"]):
                form_to_word[form] = word
        form_to_word.setdefault(word, word)
    return entries, form_to_word


def platform_replacement_suggestions(
    word: str,
    max_rank: int,
    entries: dict[str, dict[str, Any]],
    form_to_word: dict[str, str],
    max_count: int = 3,
) -> list[tuple[str, str]]:
    entry = entries.get(word, {})
    synonyms = entry.get("synonyms", [])
    suggestions: list[tuple[str, str]] = []
    seen: set[str] = set()

    def add(candidate: str, level: str):
        candidate = vocab_analyzer.normalize_vocab_item(candidate)
        if candidate and candidate != word and candidate not in seen:
            seen.add(candidate)
            suggestions.append((candidate, level))

    for synonym in synonyms:
        candidate_word = form_to_word.get(synonym)
        candidate_entry = entries.get(candidate_word or "")
        if not candidate_entry:
            continue
        candidate_rank = CEFR_ORDER.index(candidate_entry["level"])
        if candidate_rank <= max_rank:
            add(candidate_word or synonym, candidate_entry["level"])
        if len(suggestions) >= max_count:
            return suggestions

    for synonym in synonyms:
        if form_to_word.get(synonym):
            continue
        add(synonym, "-")
        if len(suggestions) >= max_count:
            return suggestions

    return suggestions


def flag_platform_cefr_words(text: str, platform_level: Any) -> str:
    max_cefr = get_platform_max_cefr(platform_level)
    if not max_cefr or not CEFR_PATH.exists():
        return ""

    entries, form_to_word = load_platform_cefr_index(str(CEFR_PATH), CEFR_PATH.stat().st_mtime)
    max_rank = CEFR_ORDER.index(max_cefr)
    normalized = vocab_analyzer.normalize_vocab_item(re.sub(r"#SC\d+\b", " ", str(text or "")))
    tokens = vocab_analyzer.TOKEN_PATTERN.findall(normalized)
    flagged: list[str] = []
    seen: set[str] = set()

    for token in tokens:
        for candidate in vocab_analyzer.lemma_candidates(token):
            word = form_to_word.get(candidate)
            if not word:
                continue
            entry = entries.get(word)
            if not entry:
                continue
            level = entry["level"]
            if CEFR_ORDER.index(level) > max_rank and word not in seen:
                suggestions = platform_replacement_suggestions(word, max_rank, entries, form_to_word)
                suggestion_text = (
                    ", ".join(f"{suggestion} ({suggestion_level})" for suggestion, suggestion_level in suggestions)
                    if suggestions
                    else "no suggestion"
                )
                seen.add(word)
                flagged.append(f"{word} ({level}) => {suggestion_text}")
            break

    return "\n".join(flagged)


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
            "2",
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


def create_vocab_template_workbook() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Vocab_Only"
    headers = ["ID", "Title", "Normal Ver.", "Easy Ver.", "Difficult Ver."]
    ws.append(headers)
    ws.append(
        [
            "OG0001",
            "Sample Title",
            "#SC01\nA small turtle finds a shiny seed.",
            "#SC01\nA turtle sees a seed.",
            "#SC01\nA curious turtle discovers a shining seed.",
        ]
    )
    widths = [14, 30, 80, 80, 80]
    header_fill = PatternFill("solid", start_color="1F3864")
    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col_idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width
        cell = ws.cell(1, col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2, max_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center" if cell.column == 1 else "left",
                vertical="top",
                wrap_text=True,
            )
    ws.freeze_panes = "A2"
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def read_story_input(uploaded_file) -> pd.DataFrame:
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file, dtype=str).fillna("")
    else:
        df = pd.read_excel(uploaded_file, dtype=str).fillna("")
    if PLATFORM_LEVEL_COLUMN not in df.columns and LEGACY_LEVEL_COLUMN in df.columns:
        df[PLATFORM_LEVEL_COLUMN] = df[LEGACY_LEVEL_COLUMN]
    return df


def validate_story_df(df: pd.DataFrame) -> list[str]:
    missing = [col for col in REQUIRED_STORY_COLUMNS if col not in df.columns]
    if PLATFORM_LEVEL_COLUMN not in df.columns and LEGACY_LEVEL_COLUMN not in df.columns:
        missing.append(PLATFORM_LEVEL_COLUMN)
    if missing:
        return [f"필수 컬럼 없음: {', '.join(missing)}"]
    errors: list[str] = []
    for idx, row in df.iterrows():
        sid = str(row.get("ID", "")).strip() or f"{idx + 2}행"
        title = str(row.get("Title", "")).strip()
        if not str(row.get("Base Text", "")).strip():
            errors.append(f"{sid} / {title}: Base Text가 비어 있음")
        if get_rewrite_level(row) not in CEFR_ORDER:
            errors.append(
                f"{sid} / {title}: Platform Level은 1-4 또는 기존 호환용 Pre A1/A1/A2/B1/B2/C1/C2 중 하나여야 함"
            )
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
  "category": "exactly 3 categories from the list, comma-separated, best match first",
  "book_mood": "exactly 3 moods from the list, comma-separated, best match first",
  "book_info": "Easy English pre-reading story introduction, 1-3 very short declarative sentences, 35 words maximum, do not end with a question",
  "keywords": ["6-12 lowercase content words, no proper nouns"],
  "intro": "Easy English spoken intro script for elementary/middle school video, preferably in the main character's voice, 35 words maximum",
  "movie_book_script": "Short movie book script summary, 2-4 short paragraphs, easy picture-book narration",
  "easy_version": "rewritten story preserving every #SC marker",
  "difficult_version": "rewritten story preserving every #SC marker"
}}

Story ID: {story_id}
Title: {title}
Input production level for rewriting (Platform Level): {input_level or "(blank)"}
Easy rewrite target: {easy_rule}
Difficult rewrite target: {difficult_rule}
Known word count: {word_count}
Known scene count: {scene_count}

Important separation of level logic:
- Detect "detected_level" and "lexile" independently from the Base Story only.
- Do not force detected_level to match the input production level.
- Use the input production level only to decide the Easy/Difficult rewrite targets. In the rewrite rules below, "level" means Platform Level.
- Preserve all #SC markers exactly. Do not remove, merge, or invent scenes.
- Easy version: rewrite the Base Story one Platform Level below the input production level.
  - Replace every word above the Easy target level with an easier synonym.
  - Keep every scene's event order. Do not omit, merge, or skip scenes.
  - Keep a similar or shorter length than the Base Story. Do not pad.
  - Fairy-tale expressions such as "Off she ran!" are allowed when natural.
  - Each scene must be clearly easier than the matching Base scene.
- Difficult version: rewrite the Base Story one Platform Level above the input production level.
  - Upgrade words only to slightly more precise or slightly more difficult expressions that fit the Difficult target level.
  - Avoid overly difficult words, archaic words, stiff literary diction, or formal written expressions that do not fit a children's storybook.
  - Scale the upgrade by the level gap: A1 to A2 should be a small lift, and higher levels may be more expressive but must still sound natural.
  - Keep every event in order. Do not add a new plot.
  - The text may become slightly longer than the Base Story because of added description, but do not pad.
  - Keep a storybook tone, not an academic essay tone.
- The category field must contain exactly 3 items from Category choices, ordered by relevance.
- The book_mood field must contain exactly 3 items from Book mood choices, ordered by relevance.
- Write book_info and intro in easy English for elementary/middle school learners. Do not use Korean.
- book_info is not a full spoiler summary. It should introduce the setup before reading, using simple words and only story events or characters that appear in the Base Story.
- book_info should be declarative, not a teaser question. Avoid endings like "Can ...?", "Will ...?", or "What will happen?". Prefer sentences like "Milo goes to find his colors." or "The girl tries to solve the problem."
- intro is for video narration, so keep it shorter and easier than ordinary reading text.
- Keep both book_info and intro within 35 words each.
- movie_book_script should feel like a children's movie book narration for Korean elementary English learners aged 6-8, CEFR A0-A2. Use language similar to the Base Story. Do not use harder words or longer sentences than the Base Story.
- movie_book_script should keep the story flow, not just summarize the plot. Include the setting/background, main character, important events, problem or challenge, solution, and ending or lesson only when they exist in the Base Story. Do not add new events, morals, emotions, or background information.
- movie_book_script should use 2-4 short paragraphs. Each paragraph should show a clear scene. Prefer one idea per sentence and short, simple sentences. Reuse simple original expressions when they are important. Keep the tone simple, warm, visual, easy to read aloud, and suitable for animation.
- For very short or low-plot stories, do not force drama. Keep the original feeling and focus on introduction, meeting, activity, and happy ending.

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
    input_level = get_rewrite_level(row)
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


def normalize_ordered_choices(raw: Any, valid_choices: list[str], limit: int = 3) -> str:
    if isinstance(raw, list):
        candidates = [str(item).strip() for item in raw]
    else:
        candidates = [
            item.strip()
            for item in re.split(r"[,/|;\n]", str(raw or ""))
            if item.strip()
        ]

    canonical = {choice.lower(): choice for choice in valid_choices}
    selected: list[str] = []
    for candidate in candidates:
        normalized = re.sub(r"\s+", " ", candidate).strip()
        choice = canonical.get(normalized.lower())
        if choice and choice not in selected:
            selected.append(choice)
        if len(selected) >= limit:
            break
    return ", ".join(selected)


def normalize_book_info(value: Any) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    if text.endswith("?"):
        question = text[:-1].strip()
        if re.match(r"(?i)^what\s+(?:will\s+)?happens?\s+(?:next|now)?$", question):
            return "The story follows the next part of the adventure."
        match = re.match(r"(?i)^what\s+(?:will\s+)?happens?\s+when\s+(.+)$", question)
        if match:
            return f"The story begins when {match.group(1).strip()}."
        verbs = (
            "look for|search for|find|get|save|help|learn|make|reach|escape|solve|bring|"
            "turn|become|discover|return|stop|win|fix|follow|meet|catch|use|remember|"
            "choose|share|finish|cross|open|close|keep|wake|fly|grow|glow|sing|dance|"
            "play|build|protect|rescue|see|hear|feel|understand|take|give|show|tell|be"
        )
        match = re.match(rf"(?i)^(?:can|will)\s+(.+?)\s+({verbs})(.*)$", question)
        if match:
            subject = match.group(1).strip()
            verb = match.group(2).strip()
            rest = match.group(3).strip()
            text = f"{subject[:1].upper()}{subject[1:]} tries to {verb}{(' ' + rest) if rest else ''}."
        else:
            text = question + "."
    return text


def normalize_story_info(parsed: dict[str, Any], row: pd.Series) -> dict[str, Any]:
    base_story = str(row["Base Text"]).strip()
    platform_level = get_platform_level(row)
    input_level = normalize_level(platform_level)
    keywords = parsed.get("keywords", [])
    if isinstance(keywords, str):
        keywords = [item.strip() for item in keywords.split(",") if item.strip()]
    elif not isinstance(keywords, list):
        keywords = []

    detected_level = normalize_level(parsed.get("detected_level", ""))
    if detected_level not in CEFR_ORDER:
        detected_level = ""

    mood = normalize_ordered_choices(parsed.get("book_mood", ""), BOOK_MOODS)
    category = normalize_ordered_choices(parsed.get("category", ""), STORY_CATEGORIES)

    return {
        "id": str(row["ID"]).strip(),
        "title": str(row["Title"]).strip(),
        "platform_level": platform_level,
        "input_level": input_level,
        "base_text": base_story,
        "flagged_words": flag_platform_cefr_words(base_story, platform_level),
        "detected_level": detected_level,
        "detected_level_rationale": str(parsed.get("detected_level_rationale", "")).strip(),
        "lexile": str(parsed.get("lexile", "")).strip(),
        "lexile_rationale": str(parsed.get("lexile_rationale", "")).strip(),
        "word_count": count_story_words(base_story),
        "scene_count": count_story_scenes(base_story),
        "category": category,
        "book_mood": mood,
        "book_info": normalize_book_info(parsed.get("book_info", "")),
        "keywords": [str(item).strip().lower() for item in keywords if str(item).strip()],
        "intro": str(parsed.get("intro", "")).strip(),
        "movie_book_script": str(parsed.get("movie_book_script", "")).strip(),
        "easy_version": str(parsed.get("easy_version", "")).strip(),
        "difficult_version": str(parsed.get("difficult_version", "")).strip(),
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
    centered_headers = {
        "ID",
        "Platform Level",
        "Input Level",
        "Detected CEFR",
        "CEFR",
        "Lexile",
        "Word Count",
        "Scene Count",
    }
    for row_idx, values in enumerate(rows, 2):
        fill = fills[row_idx % 2]
        for col_idx, value in enumerate(values, 1):
            header = headers[col_idx - 1]
            cell = ws.cell(row_idx, col_idx, "" if pd.isna(value) else value)
            cell.fill = fill
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center" if header in centered_headers else "left",
                vertical="top",
                wrap_text=True,
            )
        ws.row_dimensions[row_idx].height = 52
    ws.freeze_panes = "A2"
    return ws


def story_info_to_vocab_df(source_df: pd.DataFrame, story_info_by_id: dict[str, Any]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for _, row in source_df.iterrows():
        sid = str(row["ID"])
        story = story_info_by_id.get(sid, {})
        rows.append(
            {
                "ID": sid,
                "Title": row["Title"],
                "Normal Ver.": story.get("base_text", row.get("Base Text", "")),
                "Easy Ver.": story.get("easy_version", ""),
                "Difficult Ver.": story.get("difficult_version", ""),
                "Detected CEFR": story.get("detected_level", ""),
                "CEFR Rationale": story.get("detected_level_rationale", ""),
                "Lexile": story.get("lexile", ""),
                "Lexile Rationale": story.get("lexile_rationale", ""),
            }
        )
    return pd.DataFrame(rows)


def build_story_info_workbook(source_df: pd.DataFrame, story_info_by_id: dict[str, Any]) -> bytes:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    story_headers = [
        "ID",
        "Title",
        "Platform Level",
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
        "Movie Book Script",
        "Easy Version",
        "Difficult Version",
        "Flagged Words",
    ]
    story_rows: list[list[Any]] = []
    for _, row in source_df.iterrows():
        sid = str(row["ID"])
        story = story_info_by_id.get(sid, {})
        flagged_words = flag_platform_cefr_words(
            story.get("base_text", row.get("Base Text", "")),
            story.get("platform_level", get_platform_level(row)),
        )
        if story:
            story["flagged_words"] = flagged_words
        story_rows.append(
            [
                sid,
                row["Title"],
                story.get("platform_level", get_platform_level(row)),
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
                ", ".join(story.get("keywords", [])),
                story.get("intro", ""),
                story.get("movie_book_script", ""),
                story.get("easy_version", ""),
                story.get("difficult_version", ""),
                flagged_words,
            ]
        )

    write_table_sheet(
        wb,
        "Story_Info",
        story_headers,
        story_rows,
        [12, 30, 15, 80, 14, 42, 12, 42, 12, 12, 22, 28, 46, 36, 48, 70, 80, 80, 42],
        "385723",
    )

    vocab_df = story_info_to_vocab_df(source_df, story_info_by_id)
    write_table_sheet(
        wb,
        "Vocab_Input",
        list(vocab_df.columns),
        vocab_df.fillna("").values.tolist(),
        [12, 30, 80, 80, 80, 14, 42, 12, 42],
        "1F3864",
    )

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def build_vocab_output_workbook(vocab_df: pd.DataFrame, vocab_results_by_id: dict[str, Any]) -> bytes:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    headers = [
        "ID",
        "Title",
        "Normal Ver.",
        "Easy Ver.",
        "Difficult Ver.",
        "CEFR",
        "CEFR Rationale",
        "Lexile",
        "Lexile Rationale",
        "Vocab",
        "Normal Words",
        "Easy Words",
        "Difficult Words",
    ]
    rows: list[list[Any]] = []
    for _, row in vocab_df.iterrows():
        sid = str(row["ID"])
        vocab = vocab_results_by_id.get(sid, {})
        rows.append(
            [
                sid,
                row["Title"],
                row.get("Normal Ver.", ""),
                row.get("Easy Ver.", ""),
                row.get("Difficult Ver.", ""),
                vocab.get("cefr", row.get("Detected CEFR", "")),
                vocab.get("cefr_rationale", row.get("CEFR Rationale", "")),
                vocab.get("lexile", row.get("Lexile", "")),
                vocab.get("lexile_rationale", row.get("Lexile Rationale", "")),
                vocab_to_cell(vocab.get("vocab", [])),
                vocab_to_cell(vocab.get("normal_vocab", [])),
                vocab_to_cell(vocab.get("easy_vocab", [])),
                vocab_to_cell(vocab.get("difficult_vocab", [])),
            ]
        )

    write_table_sheet(
        wb,
        "Vocab_Click_Words",
        headers,
        rows,
        [12, 28, 55, 55, 55, 12, 42, 12, 42, 35, 50, 50, 50],
        "7030A0",
    )

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def normalize_vocab_input_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.fillna("").copy()
    rename_map = {}
    if "Base Text" in df.columns and "Normal Ver." not in df.columns:
        rename_map["Base Text"] = "Normal Ver."
    if "Easy Version" in df.columns and "Easy Ver." not in df.columns:
        rename_map["Easy Version"] = "Easy Ver."
    if "Difficult Version" in df.columns and "Difficult Ver." not in df.columns:
        rename_map["Difficult Version"] = "Difficult Ver."
    if rename_map:
        df = df.rename(columns=rename_map)
    return df


def read_vocab_input(uploaded_file) -> pd.DataFrame:
    if uploaded_file.name.lower().endswith(".csv"):
        return normalize_vocab_input_df(pd.read_csv(uploaded_file, dtype=str).fillna(""))

    xls = pd.ExcelFile(uploaded_file)
    if "Vocab_Input" in xls.sheet_names:
        sheet_name = "Vocab_Input"
    elif "Story_Info" in xls.sheet_names:
        sheet_name = "Story_Info"
    else:
        sheet_name = xls.sheet_names[0]
    return normalize_vocab_input_df(pd.read_excel(xls, sheet_name=sheet_name, dtype=str).fillna(""))


def validate_vocab_df(df: pd.DataFrame) -> list[str]:
    required = ["ID", "Title", "Normal Ver.", "Easy Ver.", "Difficult Ver."]
    missing = [col for col in required if col not in df.columns]
    if missing:
        return [f"필수 컬럼 없음: {', '.join(missing)}"]

    errors: list[str] = []
    for _, row in df.iterrows():
        blanks = [
            col
            for col in ["Normal Ver.", "Easy Ver.", "Difficult Ver."]
            if not str(row.get(col, "")).strip()
        ]
        if blanks:
            errors.append(f"{row.get('ID', '')} / {row.get('Title', '')}: {', '.join(blanks)} 비어 있음")
    return errors


def call_vocab_for_row(client, model_name: str, row: pd.Series) -> dict[str, Any] | None:
    detected_level = normalize_level(row.get("Detected CEFR", ""))
    if detected_level in CEFR_ORDER:
        story_info = {
            "id": str(row["ID"]),
            "title": str(row["Title"]),
            "base_text": str(row["Normal Ver."]),
            "easy_version": str(row["Easy Ver."]),
            "difficult_version": str(row["Difficult Ver."]),
            "detected_level": detected_level,
            "detected_level_rationale": str(row.get("CEFR Rationale", "")),
            "lexile": str(row.get("Lexile", "")),
            "lexile_rationale": str(row.get("Lexile Rationale", "")),
        }
        return call_vocab_analysis(client, model_name, story_info)
    return vocab_analyzer.analyze(client, model_name, row)


@st.cache_data
def load_cefr_wordlist(path: str, mtime: float) -> pd.DataFrame:
    _ = mtime
    df = pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")
    if "cefr_level" in df.columns:
        order = {level: idx for idx, level in enumerate(CEFR_ORDER)}
        df["cefr_sort"] = df["cefr_level"].map(order).fillna(999)
        df = df.sort_values(["cefr_sort", "word"]).drop(columns=["cefr_sort"])
    return df


def model_settings(prefix: str, api_key: str) -> str:
    with st.expander("고급 설정"):
        check_models_clicked = st.button(
            "API 키로 사용 가능한 모델 확인",
            disabled=not api_key.strip(),
            use_container_width=True,
            key=f"{prefix}_check_models",
        )
        if check_models_clicked:
            try:
                available_models = list_available_text_models(api_key)
            except Exception as exc:
                message = f"{type(exc).__name__}: {exc}"
                st.error(describe_gemini_error(message))
                st.code(message, language="text")
            else:
                st.session_state[f"{prefix}_{MODEL_SESSION_KEY}"] = available_models
                st.success(f"사용 가능한 텍스트 모델 {len(available_models):,}개를 확인했습니다.")

        model_options = st.session_state.get(f"{prefix}_{MODEL_SESSION_KEY}") or PREFERRED_MODELS
        default_index = model_options.index(DEFAULT_MODEL) if DEFAULT_MODEL in model_options else 0
        selected_model = st.selectbox(
            "모델",
            options=model_options,
            index=default_index,
            key=f"{prefix}_model",
        )
        custom_model = st.text_input(
            "직접 입력할 모델명",
            placeholder="예: gemini-2.5-flash",
            key=f"{prefix}_custom_model",
        )
        model_name = normalize_model_name(custom_model) or selected_model
        st.caption(f"현재 실행 모델: {model_name}")
        return model_name


def guide_tab():
    st.subheader("작업 흐름")
    st.markdown(
        """
        **1단계 | Story Info**

        `ID`, `Title`, `Platform Level`, `Base Text`만 입력합니다.

        - Base Text만으로 Story Info를 먼저 생성합니다.
        - 추정 CEFR, Lexile, 단어 수, 장면 수
        - Category, Book Mood, Summary, Intro Script, Movie Book Script
        - Platform Level 기준 Flagged Words와 대체어 제안
        - Easy Version, Difficult Version

        **2단계 | Vocab & Click Words**

        `Normal Ver.`, `Easy Ver.`, `Difficult Ver.`가 모두 준비된 파일을 입력합니다.

        - 세 수준의 텍스트에서 각각 클릭 단어를 추출합니다.
        - 1단계 결과를 사용할 경우, `Detected CEFR`와 LCMS CEFR DB를 우선 적용합니다.
        - 2단계만 단독 진행할 경우, Normal 텍스트 기준으로 CEFR/Lexile을 먼저 추정한 뒤 LCMS CEFR DB를 적용합니다.
        - Gemini가 제안한 단어도 LCMS DB 기준을 다시 적용합니다. 기준보다 낮은 DB 단어는 원칙적으로 제외하되, 한 단계 낮은 스토리 핵심 클릭 후보만 좁게 허용합니다.
        - 1단계 결과 엑셀의 `Vocab_Input` 시트 또는 2단계 단독 템플릿 파일을 사용할 수 있습니다.

        `Platform Level`은 Easy/Difficult 생성 기준이고, Vocab 필터링 기준은 별도로 추정한 `Detected CEFR`입니다.
        """
    )
    st.markdown("**클릭 단어 추출 기준**")
    st.markdown(
        """
        <table class="click-rule-table">
          <thead>
            <tr>
              <th>분류</th>
              <th>설명</th>
              <th>예시</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>기준 레벨 이상 어휘</td>
              <td>EVP/LCMS 기준에서 텍스트 CEFR 밴드와 같거나 그보다 높게 분류된 어휘</td>
              <td>A2 텍스트의 age (A2) - 허용<br>A2 텍스트의 blossomed (B2) - 허용<br>A2 텍스트의 all (A1) - 불가</td>
            </tr>
            <tr>
              <td>콘텐츠 특화 어휘</td>
              <td>LCMS DB에 없거나, 기준보다 한 단계 낮아도 스토리 이해에 꼭 필요한 핵심 캐릭터, 사물, 배경, 행동, 감정 후보</td>
              <td>pea-shooter, pod, moss</td>
            </tr>
            <tr>
              <td>의미 확장 어휘</td>
              <td>문자적 의미만으로 파악하기 어려운 비유, 관용 표현, 뜻이 다양하게 쓰이는 동사</td>
              <td>have, take</td>
            </tr>
            <tr>
              <td>제외 대상</td>
              <td>고유명사, 맥락만으로 100% 추론 가능한 어휘</td>
              <td>Hans, Christmas Eve</td>
            </tr>
          </tbody>
        </table>
        """,
        unsafe_allow_html=True,
    )
    st.divider()
    st.subheader("LCMS CEFR 단어 리스트")
    if not CEFR_PATH.exists():
        st.warning("lcms_cefr.csv 파일을 찾을 수 없습니다.")
        return
    cefr_df = load_cefr_wordlist(str(CEFR_PATH), CEFR_PATH.stat().st_mtime)
    source_label = ""
    if "source" in cefr_df.columns:
        source_values = [str(item).strip() for item in cefr_df["source"].dropna().unique() if str(item).strip()]
        source_label = ", ".join(source_values[:2])
    search_col, meta_col = st.columns([2, 3])
    with search_col:
        search_word = st.text_input("단어 검색", placeholder="예: apple", key="cefr_word_search")
    with meta_col:
        st.markdown(
            (
                "<div style='text-align:right; color:#5f6368; font-size:0.85rem; padding-top:2rem;'>"
                f"전체 단어 {len(cefr_df):,} · "
                f"CEFR 단계 {cefr_df['cefr_level'].nunique() if 'cefr_level' in cefr_df else 0} · "
                f"출처 {source_label or 'lcms_cefr.csv'}"
                "</div>"
            ),
            unsafe_allow_html=True,
        )
    display_columns = [col for col in ["word", "cefr_level", "source", "pos", "category"] if col in cefr_df.columns]
    display_df = cefr_df
    query = search_word.strip().lower()
    if query and "word" in cefr_df.columns:
        word_series = cefr_df["word"].astype(str).str.lower()
        exact_mask = word_series == query
        contains_mask = word_series.str.contains(query, regex=False, na=False)
        display_df = pd.concat([cefr_df[exact_mask], cefr_df[contains_mask & ~exact_mask]])
        st.caption(f"검색 결과 {len(display_df):,}개")
    st.dataframe(
        display_df[display_columns].rename(
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


def story_info_tab():
    st.subheader("1단계. Story Info 생성")
    st.caption("입력: ID / Title / Platform Level / Base Text")
    st.download_button(
        "Story_Confirmed_Template.xlsx 다운로드",
        create_template_workbook(),
        file_name="Story_Confirmed_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    uploaded_story = st.file_uploader(
        "story_confirmed 파일 업로드",
        type=["xlsx", "csv"],
        accept_multiple_files=False,
        key="story_info_upload",
    )
    if uploaded_story is None:
        st.info("템플릿을 내려받아 Base Text까지 작성한 뒤 업로드해 주세요.")
        return

    checkpoint_file = st.file_uploader(
        "이전 Story Info checkpoint JSON 업로드 (선택)",
        type=["json"],
        accept_multiple_files=False,
        key="story_info_checkpoint_upload",
    )
    with st.expander("checkpoint JSON은 무엇인가요?"):
        st.markdown(
            """
            긴 파일을 처리하다가 중간에 멈추거나 일부 행만 실패했을 때 이어서 작업하기 위한 임시 저장 파일입니다.

            - 입력: 이전에 다운로드한 checkpoint JSON을 올리면 이미 완료된 ID는 다시 처리하지 않습니다.
            - 출력: 이번 실행에서 성공한 Story Info 결과를 JSON으로 저장합니다.
            - 용도: API 오류, 쿼터 제한, 브라우저 새로고침이 있어도 완료된 결과를 재사용합니다.
            - 엑셀 최종 결과만 필요하다면 사용하지 않아도 됩니다.
            """
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
    metric_cols[3].metric("다음 단계 출력", "Vocab_Input")

    with st.expander("입력 미리보기", expanded=True):
        st.dataframe(story_df[preview_cols].head(20), use_container_width=True, hide_index=True)

    if errors:
        st.error("입력 파일을 먼저 수정해야 합니다.")
        st.code("\n".join(errors[:30]), language="text")
        return

    checkpoint = load_checkpoint(checkpoint_file)
    story_info_checkpoint = {
        sid: item.get("story_info", item)
        for sid, item in checkpoint.items()
        if isinstance(item, dict)
    }
    completed_ids = {sid for sid, item in story_info_checkpoint.items() if item.get("easy_version")}

    labels = [f"{row['ID']} | {row['Title']}" for _, row in story_df.iterrows()]
    selected_labels = st.multiselect(
        "Story Info 생성 대상",
        labels,
        default=[label for label in labels if label.split(" | ", 1)[0] not in completed_ids] or labels,
        key="story_info_selection",
    )
    selected_ids = {label.split(" | ", 1)[0] for label in selected_labels}
    selected_df = story_df[story_df["ID"].astype(str).isin(selected_ids)].copy()

    api_key = st.text_input(
        "Gemini API Key",
        type="password",
        placeholder="Story Info 생성 시에만 사용하며 앱에 저장하지 않습니다.",
        key="story_info_api_key",
    )
    model_name = model_settings("story_info", api_key)

    st.caption(
        f"checkpoint 재사용 가능: {len(completed_ids):,}개 / 현재 선택: {len(selected_df):,}개"
    )
    run_clicked = st.button(
        "Story Info 생성 실행",
        type="primary",
        disabled=selected_df.empty,
        use_container_width=True,
        key="story_info_run",
    )

    if run_clicked:
        if not api_key.strip():
            st.error("Gemini API Key가 필요합니다.")
            return
        try:
            client = genai.Client(api_key=api_key.strip())
        except Exception as exc:
            st.error(f"Gemini 클라이언트를 만들지 못했습니다: {type(exc).__name__}: {exc}")
            return

        story_info_by_id: dict[str, Any] = dict(story_info_checkpoint)
        progress = st.progress(0)
        log_box = st.empty()

        for idx, (_, row) in enumerate(selected_df.iterrows(), 1):
            sid = str(row["ID"])
            title = str(row["Title"])
            if story_info_by_id.get(sid, {}).get("easy_version"):
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

            story_info_by_id[sid] = story_info
            progress.progress(idx / len(selected_df))

        output_bytes = build_story_info_workbook(story_df, story_info_by_id)
        checkpoint_bytes = checkpoint_to_bytes(story_info_by_id)
        st.session_state["story_info_source_df"] = story_df
        st.session_state["story_info_by_id"] = story_info_by_id
        st.session_state["story_info_output_bytes"] = output_bytes
        st.session_state["story_info_checkpoint_bytes"] = checkpoint_bytes
        st.session_state["story_info_result_count"] = sum(
            1 for item in story_info_by_id.values() if item.get("easy_version")
        )
        log_box.success(f"완료: {st.session_state['story_info_result_count']:,}개 Story Info")

    if "story_info_output_bytes" in st.session_state:
        st.success("1단계 결과가 준비되었습니다. 2단계 탭에서 현재 세션 결과를 바로 사용할 수 있습니다.")
        download_cols = st.columns(2)
        with download_cols[0]:
            st.download_button(
                "Story_Info_Result.xlsx 다운로드",
                st.session_state["story_info_output_bytes"],
                file_name="Story_Info_Result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with download_cols[1]:
            st.download_button(
                "Story Info checkpoint JSON 다운로드",
                st.session_state["story_info_checkpoint_bytes"],
                file_name="story_info_checkpoint.json",
                mime="application/json",
                use_container_width=True,
            )


def vocab_tab():
    st.subheader("2단계. Vocab & Click Words 분석")
    st.caption("입력: ID / Title / Normal Ver. / Easy Ver. / Difficult Ver.")
    st.download_button(
        "2단계 단독 템플릿 다운로드",
        create_vocab_template_workbook(),
        file_name="Vocab_Only_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("2단계 단독 템플릿에는 CEFR/Lexile 입력칸이 없습니다. 분석 실행 시 Normal Ver. 기준으로 CEFR/Lexile을 먼저 추정합니다.")

    source_mode = "2단계 단독 파일 업로드"
    if "story_info_source_df" in st.session_state and "story_info_by_id" in st.session_state:
        source_mode = st.radio(
            "Vocab 입력 방식",
            ["현재 세션의 1단계 결과 사용", "2단계 단독 파일 업로드"],
            horizontal=True,
            key="vocab_source_mode",
        )
    else:
        st.info("현재 세션에 1단계 결과가 없으므로 2단계 단독 파일 업로드로 진행합니다.")

    vocab_df: pd.DataFrame | None = None
    if source_mode == "현재 세션의 1단계 결과 사용":
        vocab_df = story_info_to_vocab_df(
            st.session_state["story_info_source_df"],
            st.session_state["story_info_by_id"],
        )
    else:
        uploaded_vocab = st.file_uploader(
            "Vocab 입력 파일 업로드",
            type=["xlsx", "csv"],
            accept_multiple_files=False,
            help="1단계 결과 엑셀의 Vocab_Input 시트 또는 2단계 단독 템플릿 형식을 사용할 수 있습니다.",
            key="vocab_upload",
        )
        if uploaded_vocab is None:
            st.info("1단계에서 생성한 Story_Info_Result.xlsx 또는 2단계 단독 템플릿 파일을 업로드해 주세요.")
            return
        try:
            vocab_df = read_vocab_input(uploaded_vocab)
        except Exception as exc:
            st.error(f"Vocab 입력 파일을 읽지 못했습니다: {exc}")
            return

    errors = validate_vocab_df(vocab_df)
    preview_cols = [col for col in ["ID", "Title", "Normal Ver.", "Easy Ver.", "Difficult Ver.", "Detected CEFR", "Lexile"] if col in vocab_df.columns]
    metric_cols = st.columns(4)
    metric_cols[0].metric("스토리", f"{len(vocab_df):,}")
    metric_cols[1].metric("필수 컬럼", f"{min(5, len([c for c in ['ID','Title','Normal Ver.','Easy Ver.','Difficult Ver.'] if c in vocab_df.columns]))}/5")
    metric_cols[2].metric("입력 오류", f"{len(errors):,}")
    metric_cols[3].metric("LCMS DB", "있음" if CEFR_PATH.exists() else "없음")

    with st.expander("Vocab 입력 미리보기", expanded=True):
        st.dataframe(vocab_df[preview_cols].head(20), use_container_width=True, hide_index=True)

    if errors:
        st.error("Vocab 입력 파일을 먼저 수정해야 합니다.")
        st.code("\n".join(errors[:30]), language="text")
        return

    if "Detected CEFR" in vocab_df.columns and vocab_df["Detected CEFR"].astype(str).str.strip().any():
        st.caption("Detected CEFR가 있는 행은 해당 값을 우선 사용합니다.")
    else:
        st.caption("Detected CEFR/Lexile이 없는 입력이므로 분석 실행 시 Normal Ver. 기준으로 먼저 추정합니다.")

    labels = [f"{row['ID']} | {row['Title']}" for _, row in vocab_df.iterrows()]
    selected_labels = st.multiselect(
        "Vocab 분석 대상",
        labels,
        default=labels,
        key="vocab_selection",
    )
    selected_ids = {label.split(" | ", 1)[0] for label in selected_labels}
    selected_df = vocab_df[vocab_df["ID"].astype(str).isin(selected_ids)].copy()

    api_key = st.text_input(
        "Gemini API Key",
        type="password",
        placeholder="Vocab 분석 시에만 사용하며 앱에 저장하지 않습니다.",
        key="vocab_api_key",
    )
    model_name = model_settings("vocab", api_key)

    run_clicked = st.button(
        "Vocab & Click Words 분석 실행",
        type="primary",
        disabled=selected_df.empty,
        use_container_width=True,
        key="vocab_run",
    )

    if run_clicked:
        if not api_key.strip():
            st.error("Gemini API Key가 필요합니다.")
            return
        try:
            client = genai.Client(api_key=api_key.strip())
        except Exception as exc:
            st.error(f"Gemini 클라이언트를 만들지 못했습니다: {type(exc).__name__}: {exc}")
            return

        vocab_results_by_id: dict[str, Any] = {}
        progress = st.progress(0)
        log_box = st.empty()

        for idx, (_, row) in enumerate(selected_df.iterrows(), 1):
            sid = str(row["ID"])
            log_box.info(f"{idx}/{len(selected_df)} Vocab 분석 중: {sid} / {row['Title']}")
            try:
                result = call_vocab_for_row(client, model_name, row)
            except Exception as exc:
                message = f"{type(exc).__name__}: {exc}"
                st.warning(f"{sid} Vocab 분석 중 오류가 발생했습니다.")
                st.error(describe_gemini_error(message))
                st.code(message, language="text")
                result = None
            if result:
                vocab_results_by_id[sid] = result
            progress.progress(idx / len(selected_df))

        output_bytes = build_vocab_output_workbook(vocab_df, vocab_results_by_id)
        checkpoint_bytes = checkpoint_to_bytes(vocab_results_by_id)
        st.session_state["vocab_output_bytes"] = output_bytes
        st.session_state["vocab_checkpoint_bytes"] = checkpoint_bytes
        st.session_state["vocab_result_count"] = len(vocab_results_by_id)
        log_box.success(f"완료: {len(vocab_results_by_id):,}개 Vocab 결과")

    if "vocab_output_bytes" in st.session_state:
        download_cols = st.columns(2)
        with download_cols[0]:
            st.download_button(
                "Vocab_Click_Words_Analysis.xlsx 다운로드",
                st.session_state["vocab_output_bytes"],
                file_name="Vocab_Click_Words_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with download_cols[1]:
            st.download_button(
                "Vocab checkpoint JSON 다운로드",
                st.session_state["vocab_checkpoint_bytes"],
                file_name="vocab_checkpoint.json",
                mime="application/json",
                use_container_width=True,
            )


st.title("Story Info.")
tab_guide, tab_story_info, tab_vocab = st.tabs(["가이드", "1단계 Story Info", "2단계 Vocab & Click Words"])

with tab_guide:
    guide_tab()

with tab_story_info:
    story_info_tab()

with tab_vocab:
    vocab_tab()
