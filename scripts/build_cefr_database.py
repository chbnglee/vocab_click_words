from __future__ import annotations

import argparse
import csv
import json
import re
from collections import Counter, OrderedDict
from pathlib import Path

import pdfplumber


SOURCE_FILES = [
    (
        "Cambridge_Starters",
        "Cambridge-starters-movers-flyers-word-list-2025.pdf",
        "parse_cambridge_starters",
    ),
    (
        "Cambridge_B1 Preliminary",
        "Cambridge-b1-preliminary-vocabulary-list.pdf",
        "parse_cambridge_b1",
    ),
    (
        "Oxford 3000",
        "The_Oxford_3000_by_CEFR_level.pdf",
        "parse_oxford",
    ),
    (
        "Oxford 5000",
        "The_Oxford_5000_by_CEFR_level.pdf",
        "parse_oxford",
    ),
]

CEFR_ORDER = ["Pre A1", "A1", "A2", "B1", "B2", "C1", "C2"]
POS_SET = {"adj", "adv", "conj", "det", "dis", "excl", "int", "n", "poss", "prep", "pron", "v", "title"}
NAME_WORDS = set(
    """
    alex alice ann anna ben bill dan eva grace hugo jill kim lucy mark matt nick
    pat sam sue tom charlie clare daisy fred jack jane jim julia lily mary paul
    peter sally vicky zoe betty david emma frank george harry helen holly katy
    michael oliver richard robert sarah sophia william
    """.split()
)

LEVEL_RE = re.compile(r"^(Pre A1|A1|A2|B1|B2|C1|C2)$")
OXFORD_POS_RE = re.compile(
    r"\s(?:modal\s+v\.|auxiliary\s+v\.|indefinite\s+article|"
    r"adj\.|adv\.|n\.|v\.|prep\.|pron\.|det\.|conj\.|number|exclam\.)"
)
B1_ENTRY_RE = re.compile(
    r"^(?P<head>.+?)\s+\((?P<pos>[^)]*"
    r"(?:adj|adv|n|v|prep|pron|det|conj|exclam|phr|mv|av|abbrev|pl|sing)[^)]*)\)"
)


def normalize_text(value: str) -> str:
    return (
        value.strip()
        .replace("\u2019", "'")
        .replace("\u2018", "'")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
    )


def clean_word(raw: str) -> list[str]:
    word = normalize_text(raw)
    word = re.sub(r"\((?:s \+ pl|pl|sing)\)", "", word, flags=re.I)
    word = re.sub(
        r"\((?:as in|e\.g\.|i\.e\.|for|long flat surface|money|river|animal|"
        r"final|taking time|deal with|tell a lie|find sb/sth pleasant|caring|"
        r"computer|theatre|football|time|ride|mail|direction|car; bike|"
        r"a competition|sth|sb|past of can)[^)]*\)",
        "",
        word,
        flags=re.I,
    )
    word = re.sub(r"\bsth/sb\b|\bsb/sth\b|\bsth\b|\bsb\b", "", word)
    word = re.sub(r"\b(?:UK|US|Br Eng|Am Eng):?\b", "", word, flags=re.I)
    word = word.replace("...", " ")
    word = re.sub(r"\s+", " ", word).strip(" .;,")
    word = re.sub(r"(?<=[a-zA-Z])\d+$", "", word)
    word = word.lower().replace("\u2019", "'")

    if not word or word == "no words at this level":
        return []

    variants: list[str] = []
    if "/" in word:
        parts = word.split("/")
        if len(parts) == 2:
            left, right = parts[0].strip(), parts[1].strip()
            if " " in right and " " not in left:
                suffix = right.split(" ", 1)[1]
                variants.extend([f"{left} {suffix}", right])
            elif " " in left and " " not in right:
                prefix = left.rsplit(" ", 1)[0]
                variants.extend([left, f"{prefix} {right}"])
            elif left.endswith("man") and right == "woman":
                variants.extend([left, f"{left[:-3]}woman"])
            else:
                variants.extend([left, right])
        else:
            variants.append(word)
    elif re.search(r",\s+", word) and len(word) < 12:
        variants.extend(part.strip() for part in word.split(","))
    else:
        variants.append(word)

    cleaned: list[str] = []
    for variant in variants:
        variant = re.sub(r"\s+", " ", variant).strip(" .;,")
        variant = re.sub(r"(?<=[a-zA-Z])\d+$", "", variant)
        if variant:
            cleaned.append(variant)
    return cleaned


def expand_cambridge_variants(head: str) -> list[str]:
    candidates = [head]
    for match in re.finditer(r"\((?:UK|US)\s+([^)]*)\)", head, flags=re.I):
        candidates.append(re.sub(r"\(s\)", "s", match.group(1)))

    words: list[str] = []
    for candidate in candidates:
        candidate = re.sub(r"\((?:UK|US)\s+[^)]*\)", "", candidate, flags=re.I)
        words.extend(clean_word(candidate))
    return words


def parse_cambridge_line(line: str) -> list[str]:
    line = normalize_text(line)
    if not line or re.fullmatch(r"[A-Z]", line):
        return []

    skip_starts = [
        "Grammatical key",
        "Candidates",
        "Names",
        "Numbers",
        "Pre A1 Starters",
        "A1 Movers",
        "A2 Flyers",
        "Letters & numbers",
    ]
    if any(line.startswith(prefix) for prefix in skip_starts):
        return []

    grammar_terms = [
        "adjective",
        "adverb",
        "conjunction",
        "interrogative",
        "possessive",
        "determiner",
        "discourse marker",
        "exclamation",
        "preposition",
        "pronoun",
        "verb",
        "noun",
    ]
    if any(term in line.lower() for term in grammar_terms) or "wordlist" in line.lower():
        return []

    line = re.sub(r"^\d+\s+", "", line)
    line = re.sub(r"\s+\d+\s+[A-Za-z0-9 ].*$", "", line)

    tokens = line.split()
    words: list[str] = []
    start = 0
    idx = 0
    while idx < len(tokens):
        token = tokens[idx].strip(";,")
        if token in POS_SET:
            head = " ".join(tokens[start:idx]).strip()
            if head:
                for word in expand_cambridge_variants(head):
                    if word not in NAME_WORDS:
                        words.append(word)

            idx += 1
            while idx < len(tokens):
                current = tokens[idx].strip(";,")
                if current == "+" and idx + 1 < len(tokens) and tokens[idx + 1].strip(";,") in POS_SET:
                    idx += 2
                    continue
                if current == "of" and idx + 1 < len(tokens) and tokens[idx + 1].strip(";,") in {"place", "time"}:
                    idx += 2
                    continue
                if current == "or" and idx + 1 < len(tokens) and tokens[idx + 1].strip(";,") == "conj":
                    idx += 2
                    continue
                break
            start = idx
            continue
        idx += 1

    return words


def parse_cambridge_starters(path: Path) -> list[tuple[str, str]]:
    # Only parse the three individual A-Z lists. The combined/thematic lists are reprints.
    ranges = [(4, 7, "Pre A1"), (8, 11, "A1"), (12, 15, "A2")]
    rows: list[tuple[str, str]] = []
    with pdfplumber.open(path) as pdf:
        for start, end, level in ranges:
            for page_num in range(start, end + 1):
                text = pdf.pages[page_num - 1].extract_text(x_tolerance=1, y_tolerance=3) or ""
                for line in text.splitlines():
                    rows.extend((word, level) for word in parse_cambridge_line(line))
    return rows


def column_lines(page, bins: list[tuple[int, int]]) -> list[str]:
    words = page.extract_words(x_tolerance=1, y_tolerance=3, keep_blank_chars=False, use_text_flow=False)
    lines: list[str] = []
    for left, right in bins:
        column_words = [
            word for word in words if left <= word["x0"] < right and 35 <= word["top"] <= 780
        ]
        column_words = sorted(column_words, key=lambda word: (word["top"], word["x0"]))

        current_line: list[dict] = []
        current_top = None
        for word in column_words:
            top = word["top"]
            if current_top is None or abs(top - current_top) <= 3:
                current_line.append(word)
                current_top = top if current_top is None else min(current_top, top)
            else:
                lines.append(" ".join(item["text"] for item in sorted(current_line, key=lambda item: item["x0"])))
                current_line = [word]
                current_top = top
        if current_line:
            lines.append(" ".join(item["text"] for item in sorted(current_line, key=lambda item: item["x0"])))
    return lines


def parse_oxford(path: Path) -> list[tuple[str, str]]:
    rows: list[tuple[str, str]] = []
    level: str | None = None
    bins = [(35, 160), (160, 292), (292, 423), (423, 560)]

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for raw_line in column_lines(page, bins):
                line = re.sub(r"\s+", " ", normalize_text(raw_line)).strip()
                if not line or line.startswith("\u00a9") or "Oxford University Press" in line:
                    continue

                level_match = LEVEL_RE.match(line)
                if level_match:
                    level = level_match.group(1)
                    continue
                if level is None:
                    continue

                pos_match = OXFORD_POS_RE.search(line)
                if not pos_match:
                    continue

                head = line[: pos_match.start()].strip()
                if any(
                    fragment in head
                    for fragment in ["Oxford 3000", "Oxford 5000", "CEFR level", "learners of English", "additional"]
                ):
                    continue

                rows.extend((word, level) for word in clean_word(head))
    return rows


def parse_b1_head(head: str) -> list[str]:
    return clean_word(re.sub(r"\(.*?\)", "", head))


def parse_cambridge_b1(path: Path) -> list[tuple[str, str]]:
    rows: list[tuple[str, str]] = []
    bins = [(45, 300), (315, 560)]
    with pdfplumber.open(path) as pdf:
        for page_num in range(4, 41):
            for raw_line in column_lines(pdf.pages[page_num - 1], bins):
                line = re.sub(r"\s+", " ", normalize_text(raw_line)).strip()
                if (
                    not line
                    or line.startswith("\u2022")
                    or line.startswith("\u00a9")
                    or "Page " in line
                    or re.fullmatch(r"[A-Z]", line)
                    or line.startswith(("Preliminary", "Schools", "Vocabulary", "List", "Appendix"))
                ):
                    continue

                match = B1_ENTRY_RE.match(line)
                if not match:
                    continue

                for word in parse_b1_head(match.group("head")):
                    if word not in NAME_WORDS and not word.startswith("\u2022"):
                        rows.append((word, "B1"))
    return rows


def merge_sources(reference_dir: Path) -> tuple[list[dict], Counter, Counter]:
    database: OrderedDict[str, dict] = OrderedDict()
    raw_counts: Counter = Counter()
    unique_counts: Counter = Counter()

    parsers = {
        "parse_cambridge_starters": parse_cambridge_starters,
        "parse_cambridge_b1": parse_cambridge_b1,
        "parse_oxford": parse_oxford,
    }

    for source, filename, parser_name in SOURCE_FILES:
        path = reference_dir / filename
        if not path.exists():
            raise FileNotFoundError(f"Missing reference PDF: {path}")

        rows = parsers[parser_name](path)
        raw_counts[source] = len(rows)
        unique_counts[source] = len({word for word, _ in rows})

        for word, cefr_level in rows:
            if not word or word in NAME_WORDS:
                continue

            record = database.setdefault(
                word,
                {
                    "word": word,
                    "cefr_level": cefr_level,
                    "source": source,
                    "references": [],
                },
            )
            reference = {"source": source, "cefr_level": cefr_level}
            if reference not in record["references"]:
                record["references"].append(reference)

    records: list[dict] = []
    level_order = {level: idx for idx, level in enumerate(CEFR_ORDER)}
    for record in database.values():
        references = record["references"]
        additional = references[1:]
        records.append(
            {
                "word": record["word"],
                "cefr_level": record["cefr_level"],
                "source": record["source"],
                "all_references": " | ".join(f"{item['source']}: {item['cefr_level']}" for item in references),
                "additional_references": " | ".join(f"{item['source']}: {item['cefr_level']}" for item in additional),
                "reference_details": json.dumps(references, ensure_ascii=False),
            }
        )

    records.sort(key=lambda item: (level_order.get(item["cefr_level"], 999), item["word"]))
    return records, raw_counts, unique_counts


def write_csv(records: list[dict], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "word",
        "cefr_level",
        "source",
        "all_references",
        "additional_references",
        "reference_details",
    ]
    with output_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(records)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build merged CEFR vocabulary CSV from local reference PDFs.")
    parser.add_argument(
        "--reference-dir",
        default=str(Path(__file__).resolve().parents[2] / "Reference_Vocab"),
        help="Directory containing the four reference PDF files.",
    )
    parser.add_argument(
        "--output",
        default=str(Path(__file__).resolve().parents[1] / "merged_cefr.csv"),
        help="Output CSV path.",
    )
    args = parser.parse_args()

    records, raw_counts, unique_counts = merge_sources(Path(args.reference_dir))
    write_csv(records, Path(args.output))

    print(f"Wrote {len(records):,} unique words to {args.output}")
    for source, _, _ in SOURCE_FILES:
        print(f"- {source}: raw={raw_counts[source]:,}, unique={unique_counts[source]:,}")


if __name__ == "__main__":
    main()
