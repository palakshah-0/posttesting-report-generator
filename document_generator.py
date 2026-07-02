from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
from docxtpl import DocxTemplate

from field_mapping import FIELD_MAP

# Classification rules are where the score for the child falls 

TEST_CLASSIFICATION_RULES = {
    "kbit_standard": [
        (131, "Superior"),
        (116, "Above Average"),
        (85, "Average"),
        (70, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],
    "ctopp_elision_standard": [
        (15, "Superior"),
        (13, "Above Average"),
        (8, "Average"),
        (6, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],
    "ctopp_nwr_standard": [
        (15, "Superior"),
        (13, "Above Average"),
        (8, "Average"),
        (6, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],
    "wrmt_word_id_standard": [
        (131, "Superior"),
        (116, "Above Average"),
        (85, "Average"),
        (70, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],
    "wrmt_word_attack_standard": [
        (131, "Superior"),
        (116, "Above Average"),
        (85, "Average"),
        (70, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],
    "wrmt_pc_standard": [
        (131, "Superior"),
        (116, "Above Average"),
        (85, "Average"),
        (70, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],
    "towre_swe_standard": [
        (121, "Superior"),
        (111, "Above Average"),
        (90, "Average"),
        (80, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],
    "towre_pde_standard": [
        (121, "Superior"),
        (111, "Above Average"),
        (90, "Average"),
        (80, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],
    "ppvt_standard": [
        (131, "Superior"),
        (116, "Above Average"),
        (85, "Average"),
        (70, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],
}


CATEGORY_FIELDS = {
    "kbit_standard": "kbit_category",
    "ctopp_elision_standard": "ctopp_elision_category",
    "ctopp_nwr_standard": "ctopp_nwr_category",
    "wrmt_word_id_standard": "wrmt_word_id_category",
    "wrmt_word_attack_standard": "wrmt_word_attack_category",
    "wrmt_pc_standard": "wrmt_pc_category",
    "towre_swe_standard": "towre_swe_category",
    "towre_pde_standard": "towre_pde_category",
    "ppvt_standard" : "ppvt_category"
}

def clean_value(value: Any) -> str:
    """
    Convert missing values to an empty string.
    Convert all other values to cleaned text.
    """
    if pd.isna(value):
        return ""

    return str(value).strip()

def classify_score(value: Any, rules: list[tuple[float, str]]) -> str:
    """
    Return the interpretation category for a score.
    """
    if pd.isna(value) or str(value).strip() == "":
        return ""

    try:
        score = float(str(value).strip())
    except ValueError:
        return ""

    for cutoff, category in rules:
        if score >= cutoff:
            return category

    return ""

def format_date(value: Any) -> str:
    """
    Convert a date into a readable format.
    Example: 2026-06-23 becomes June 23, 2026.
    """
    if pd.isna(value) or str(value).strip() == "":
        return ""

    parsed_date = pd.to_datetime(value, errors="coerce")

    if pd.isna(parsed_date):
        return clean_value(value)

    return parsed_date.strftime("%B %d, %Y")


def build_context(row: pd.Series) -> dict:
    """
    Turn one CSV row into the dictionary used by the Word template.
    """
    context = {}

    for csv_field, template_field in FIELD_MAP.items():
        context[template_field] = clean_value(row.get(csv_field))

    context["visit_date"] = format_date(row.get("visit_date"))

    # Add interpretation categories
    for score_field, category_field in CATEGORY_FIELDS.items():
        context[category_field] = classify_score(
            row.get(score_field),
            TEST_CLASSIFICATION_RULES[score_field],
        )

    return context


def create_document(
    template_path: str | Path,
    context: dict,
) -> BytesIO:
    """
    Fill the Word template and return the completed document.
    """
    template = DocxTemplate(template_path)
    template.render(context)

    output = BytesIO()
    template.save(output)
    output.seek(0)

    return output
