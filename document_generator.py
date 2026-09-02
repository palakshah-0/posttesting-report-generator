from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
from docxtpl import DocxTemplate

from field_mapping import FIELD_MAP
from post_test_field_mapping import POST_TEST_FIELD_MAP


# ---------------------------------------------------------
# CLASSIFICATION RULES
# ---------------------------------------------------------

TEST_CLASSIFICATION_RULES = {

    "kbit_standard": [
        (131, "Upper Extreme"),
        (116, "Above Average"),
        (85, "Average"),
        (70, "Below Average"),
        (float("-inf"), "Lower Extreme"),
    ],

    "ctopp_elision_standard": [
        (17, "Very Superior"),
        (15, "Superior"),
        (13, "Above Average"),
        (8, "Average"),
        (6, "Below Average"),
        (4, "Poor"),
        (float("-inf"), "Very Poor"),
    ],

    "ctopp_nwr_standard": [
        (17, "Very Superior"),
        (15, "Superior"),
        (13, "Above Average"),
        (8, "Average"),
        (6, "Below Average"),
        (4, "Poor"),
        (float("-inf"), "Very Poor"),
    ],

    "wrmt_word_id_standard": [
        (131, "Well Above Average"),
        (116, "Above Average"),
        (85, "Average"),
        (70, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],

    "wrmt_word_attack_standard": [
        (131, "Well Above Average"),
        (116, "Above Average"),
        (85, "Average"),
        (70, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],

    "wrmt_pc_standard": [
        (131, "Well Above Average"),
        (116, "Above Average"),
        (85, "Average"),
        (70, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],

    "towre_swe_standard": [
        (130, "Very Superior"),
        (121, "Superior"),
        (111, "Above Average"),
        (90, "Average"),
        (80, "Below Average"),
        (70, "Poor"),
        (float("-inf"), "Very Poor"),
    ],

    "towre_pde_standard": [
        (130, "Very Superior"),
        (121, "Superior"),
        (111, "Above Average"),
        (90, "Average"),
        (80, "Below Average"),
        (70, "Poor"),
        (float("-inf"), "Very Poor"),
    ],

    "ppvt_standard": [
        (130, "Extremely high score"),
        (116, "Moderately high score"),
        (85, "Average score"),
        (70, "Moderately low score"),
        (float("-inf"), "Extremely low score"),
    ],
}


# ---------------------------------------------------------
# CATEGORY FIELD NAMES
# ---------------------------------------------------------

CATEGORY_FIELDS = {
    "kbit_standard": "kbit_category",
    "ctopp_elision_standard": "ctopp_elision_category",
    "ctopp_nwr_standard": "ctopp_nwr_category",
    "wrmt_word_id_standard": "wrmt_word_id_category",
    "wrmt_word_attack_standard": "wrmt_word_attack_category",
    "wrmt_pc_standard": "wrmt_pc_category",
    "towre_swe_standard": "towre_swe_category",
    "towre_pde_standard": "towre_pde_category",
    "ppvt_standard": "ppvt_category",
}

POST_TEST_CLASSIFICATION_FIELDS = {
    "ctopp_elision_std_post": (
        "ctopp_elision_standard",
        "ctopp_elision_category",
    ),

    "ctopp_nwr_std_post": (
        "ctopp_nwr_standard",
        "ctopp_nwr_category",
    ),

    "wrmt_wid_std_post": (
        "wrmt_word_id_standard",
        "wrmt_word_id_category",
    ),

    "wrmt_wa_std_post": (
        "wrmt_word_attack_standard",
        "wrmt_word_attack_category",
    ),

    "wrmt_pc_std_post": (
        "wrmt_pc_standard",
        "wrmt_pc_category",
    ),

    "towre_swe_std_post": (
        "towre_swe_standard",
        "towre_swe_category",
    ),

    "towre_pde_std_post": (
        "towre_pde_standard",
        "towre_pde_category",
    ),

    "ppvt_std_post": (
        "ppvt_standard",
        "ppvt_category",
    ),
}


# ---------------------------------------------------------
# GENERAL HELPER FUNCTIONS
# ---------------------------------------------------------

def clean_value(value: Any) -> str:
    """
    Convert missing values to an empty string.
    Convert all other values to cleaned text.
    """

    if pd.isna(value):
        return ""

    return str(value).strip()


def format_date(value: Any) -> str:
    """
    Convert a date into a readable format.

    Example:
    2026-06-23 -> June 23, 2026
    """

    if pd.isna(value) or str(value).strip() == "":
        return ""

    parsed_date = pd.to_datetime(
        value,
        errors="coerce",
    )

    if pd.isna(parsed_date):
        return clean_value(value)

    return parsed_date.strftime("%B %d, %Y")


# ---------------------------------------------------------
# SCORE CLASSIFICATION
# ---------------------------------------------------------

def classify_score(
    score_value: Any,
    rules: list[tuple[float, str]],
) -> str:
    """
    Convert a numeric score into its classification category.
    """

    cleaned_score = clean_value(score_value)

    if cleaned_score == "":
        return ""

    try:
        numeric_score = float(cleaned_score)

    except ValueError:
        return ""

    for minimum_score, category in rules:

        if numeric_score >= minimum_score:
            return category

    return ""


def add_classification_categories(
    context: dict,
) -> dict:
    """
    Add classification labels to the template context.
    """

    for score_field, category_field in CATEGORY_FIELDS.items():

        rules = TEST_CLASSIFICATION_RULES.get(
            score_field
        )

        if rules is None:
            continue

        score_value = context.get(
            score_field,
            "",
        )

        context[category_field] = classify_score(
            score_value,
            rules,
        )

    return context

def add_post_test_classification_categories(
    context: dict,
) -> dict:
    """
    Add category labels for post-test score fields.
    """

    for post_score_field, (
        classification_rule_name,
        category_field,
    ) in POST_TEST_CLASSIFICATION_FIELDS.items():

        score_value = context.get(
            post_score_field,
            "",
        )

        rules = TEST_CLASSIFICATION_RULES.get(
            classification_rule_name
        )

        if rules is None:
            continue

        context[category_field] = classify_score(
            score_value,
            rules,
        )

    return context
# ---------------------------------------------------------
# PRE-TEST CONTEXT
# ---------------------------------------------------------

def calculate_dorf_accuracy(
    words_correct,
    total_words,
) -> str:
    """
    Calculate DORF accuracy as a whole-number percentage.
    Example: 95 correct out of 100 returns '95%'.
    """

    try:
        correct = float(clean_value(words_correct))
        total = float(clean_value(total_words))
    except ValueError:
        return ""

    if total == 0:
        return ""

    accuracy = (correct / total) * 100

    return f"{accuracy:.0f}%"

def build_context(
    row: pd.Series,
) -> dict:
    """
    Build the data used by the pre-test Word template.
    """

    context = {}

    for csv_field, template_field in FIELD_MAP.items():

        context[template_field] = clean_value(
            row.get(csv_field)
        )

    context["visit_date"] = format_date(
        row.get("visit_date")
    )

    context = add_classification_categories(
        context
    )

    return context


# ---------------------------------------------------------
# POST-TEST CONTEXT
# ---------------------------------------------------------

def build_post_test_context(
    row: pd.Series,
) -> dict:
    """
    Build the data used by the post-test Word template.
    """

    context = {}

    for csv_field, template_field in POST_TEST_FIELD_MAP.items():

        context[template_field] = clean_value(
            row.get(csv_field)
        )

    context = add_post_test_classification_categories(
        context
    )

    context["dorf_accuracy"] = calculate_dorf_accuracy(
    row.get("dibels_words_correct_post"),
    row.get("dibels_total_words_post"),
)

    return context
# ---------------------------------------------------------
# WORD DOCUMENT GENERATION
# ---------------------------------------------------------

def create_document(
    template_path: str | Path,
    context: dict,
) -> BytesIO:
    """
    Fill a Word template and return the completed document.
    """

    template = DocxTemplate(
        template_path
    )

    template.render(
        context
    )

    output = BytesIO()

    template.save(
        output
    )

    output.seek(0)

    return output
