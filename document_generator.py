from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
from docxtpl import DocxTemplate

from post_test_field_mapping import POST_TEST_FIELD_MAP


# ---------------------------------------------------------
# CLASSIFICATION RULES
# ---------------------------------------------------------

TEST_CLASSIFICATION_RULES = {

    # CTOPP-2 scaled scores
    "ctopp_elision": [
        (17, "Very Superior"),
        (15, "Superior"),
        (13, "Above Average"),
        (8, "Average"),
        (6, "Below Average"),
        (4, "Poor"),
        (float("-inf"), "Very Poor"),
    ],

    "ctopp_nwr": [
        (17, "Very Superior"),
        (15, "Superior"),
        (13, "Above Average"),
        (8, "Average"),
        (6, "Below Average"),
        (4, "Poor"),
        (float("-inf"), "Very Poor"),
    ],

    # WRMT-III standard scores
    "wrmt_word_id": [
        (131, "Well Above Average"),
        (116, "Above Average"),
        (85, "Average"),
        (70, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],

    "wrmt_word_attack": [
        (131, "Well Above Average"),
        (116, "Above Average"),
        (85, "Average"),
        (70, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],

    "wrmt_pc": [
        (131, "Well Above Average"),
        (116, "Above Average"),
        (85, "Average"),
        (70, "Below Average"),
        (float("-inf"), "Well Below Average"),
    ],

    # TOWRE-2 standard scores
    "towre_swe": [
        (130, "Very Superior"),
        (121, "Superior"),
        (111, "Above Average"),
        (90, "Average"),
        (80, "Below Average"),
        (70, "Poor"),
        (float("-inf"), "Very Poor"),
    ],

    "towre_pde": [
        (130, "Very Superior"),
        (121, "Superior"),
        (111, "Above Average"),
        (90, "Average"),
        (80, "Below Average"),
        (70, "Poor"),
        (float("-inf"), "Very Poor"),
    ],

    # PPVT-5 standard scores
    "ppvt": [
        (130, "Extremely high score"),
        (116, "Moderately high score"),
        (85, "Average score"),
        (70, "Moderately low score"),
        (float("-inf"), "Extremely low score"),
    ],
}


# ---------------------------------------------------------
# POST-TEST SCORE -> CATEGORY MAPPING
# ---------------------------------------------------------

POST_TEST_CLASSIFICATION_FIELDS = {

    "ctopp_elision_std_post": (
        "ctopp_elision",
        "ctopp_elision_category",
    ),

    "ctopp_nwr_std_post": (
        "ctopp_nwr",
        "ctopp_nwr_category",
    ),

    "wrmt_wid_std_post": (
        "wrmt_word_id",
        "wrmt_word_id_category",
    ),

    "wrmt_wa_std_post": (
        "wrmt_word_attack",
        "wrmt_word_attack_category",
    ),

    "wrmt_pc_std_post": (
        "wrmt_pc",
        "wrmt_pc_category",
    ),

    "towre_swe_std_post": (
        "towre_swe",
        "towre_swe_category",
    ),

    "towre_pde_std_post": (
        "towre_pde",
        "towre_pde_category",
    ),

    "ppvt_std_post": (
        "ppvt",
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

    if pd.isna(value):
        return ""

    cleaned_value = str(value).strip()

    if cleaned_value == "":
        return ""

    parsed_date = pd.to_datetime(
        cleaned_value,
        errors="coerce",
    )

    if pd.isna(parsed_date):
        return cleaned_value

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

    except (ValueError, TypeError):
        return ""

    for minimum_score, category in rules:

        if numeric_score >= minimum_score:
            return category

    return ""


def add_post_test_classification_categories(
    context: dict,
) -> dict:
    """
    Add classification/category labels for all post-test scores.
    """

    for (
        post_score_field,
        classification_info,
    ) in POST_TEST_CLASSIFICATION_FIELDS.items():

        rule_name, category_field = classification_info

        score_value = context.get(
            post_score_field,
            "",
        )

        rules = TEST_CLASSIFICATION_RULES.get(
            rule_name
        )

        if rules is None:
            context[category_field] = ""
            continue

        context[category_field] = classify_score(
            score_value,
            rules,
        )

    return context


# ---------------------------------------------------------
# DIBELS ACCURACY
# ---------------------------------------------------------

def calculate_dorf_accuracy(
    words_correct: Any,
    total_words: Any,
) -> str:
    """
    Calculate DIBELS Oral Reading Fluency accuracy.

    Example:
    95 correct / 100 total words -> 95%
    """

    cleaned_correct = clean_value(
        words_correct
    )

    cleaned_total = clean_value(
        total_words
    )

    if cleaned_correct == "" or cleaned_total == "":
        return ""

    try:
        correct = float(cleaned_correct)
        total = float(cleaned_total)

    except (ValueError, TypeError):
        return ""

    if total <= 0:
        return ""

    accuracy = (correct / total) * 100

    return f"{accuracy:.0f}%"


# ---------------------------------------------------------
# POST-TEST CONTEXT
# ---------------------------------------------------------

def build_context(
    row: pd.Series,
) -> dict:
    """
    Build the context dictionary used by the post-test
    Word document template.
    """

    context = {}

    # Copy post-test scores from the CSV into the template context
    for csv_field, template_field in POST_TEST_FIELD_MAP.items():

        context[template_field] = clean_value(
            row.get(csv_field)
        )

    # Add classification categories
    context = add_post_test_classification_categories(
        context
    )

    # Calculate DIBELS Oral Reading Fluency accuracy
    context["dorf_accuracy"] = calculate_dorf_accuracy(
        row.get("dibels_words_correct_post"),
        row.get("dibels_total_words_post"),
    )

    return context


# Optional alias so either function name works
def build_post_test_context(
    row: pd.Series,
) -> dict:
    """
    Alias for build_context().
    """

    return build_context(row)


# ---------------------------------------------------------
# WORD DOCUMENT GENERATION
# ---------------------------------------------------------

def create_document(
    template_path: str | Path,
    context: dict,
) -> BytesIO:
    """
    Fill the post-test Word template and return
    the completed document as an in-memory file.
    """

    template_path = Path(
        template_path
    )

    if not template_path.exists():
        raise FileNotFoundError(
            f"Word template not found: {template_path}"
        )

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
