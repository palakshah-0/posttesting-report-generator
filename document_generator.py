from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
from docxtpl import DocxTemplate

from field_mapping import FIELD_MAP


# Category fields = sentence fields for paragraph text
TEST_CATEGORY_RULES = {
    "kbit_standard": [
        (131, "suggests advanced nonverbal reasoning abilities"),
        (116, "suggests strong nonverbal reasoning abilities"),
        (85, "suggests typical nonverbal reasoning abilities"),
        (70, "suggests some difficulties with nonverbal reasoning abilities"),
        (float("-inf"), "suggests significant difficulties with nonverbal reasoning abilities"),
    ],
    "ctopp_elision_standard": [
        (15, "suggests advanced phonological awareness and ability to manipulate sound structures in words"),
        (13, "suggests strong phonological awareness and ability to manipulate sound structures in words"),
        (8, "suggests typical phonological awareness and ability to manipulate sound structures in words"),
        (6, "suggests some difficulties with phonological awareness and the ability to manipulate sound structures in words"),
        (float("-inf"), "suggests significant difficulties with phonological awareness and the ability to manipulate sound structures in words"),
    ],
    "ctopp_nwr_standard": [
        (15, "suggests advanced phonological memory and ability to reproduce unfamiliar sound sequences"),
        (13, "suggests strong phonological memory and ability to reproduce unfamiliar sound sequences"),
        (8, "suggests typical phonological memory and ability to reproduce unfamiliar sound sequences"),
        (6, "suggests some difficulties with phonological memory and the ability to reproduce unfamiliar sound sequences"),
        (float("-inf"), "suggests significant difficulties with phonological memory and the ability to reproduce unfamiliar sound sequences"),
    ],
    "wrmt_word_id_standard": [
        (131, "suggests an advanced ability to recognize and identify written words"),
        (116, "suggests a strong ability to recognize and identify written words"),
        (85, "suggests a typical ability to recognize and identify written words"),
        (70, "suggests some difficulties with recognizing and identifying written words"),
        (float("-inf"), "suggests significant difficulties with recognizing and identifying written words"),
    ],
    "wrmt_word_attack_standard": [
        (131, "suggests an advanced ability to apply phonological and structural knowledge to unfamiliar words"),
        (116, "suggests a strong ability to apply phonological and structural knowledge to unfamiliar words"),
        (85, "suggests a typical ability to apply phonological and structural knowledge to unfamiliar words"),
        (70, "suggests some difficulties with applying phonological and structural knowledge to unfamiliar words"),
        (float("-inf"), "suggests significant difficulties with applying phonological and structural knowledge to unfamiliar words"),
    ],
    "wrmt_pc_standard": [
        (131, "suggests an advanced ability to understand written text and identify vocabulary in context"),
        (116, "suggests a strong ability to understand written text and identify vocabulary in context"),
        (85, "suggests a typical ability to understand written text and identify vocabulary in context"),
        (70, "suggests some difficulties with understanding written text and identifying vocabulary in context"),
        (float("-inf"), "suggests significant difficulties with understanding written text and identifying vocabulary in context"),
    ],
    "towre_swe_standard": [
        (121, "suggests an advanced ability to rapidly recognize familiar words"),
        (111, "suggests a strong ability to rapidly recognize familiar words"),
        (90, "suggests a typical ability to rapidly recognize familiar words"),
        (80, "suggests some difficulties rapidly recognizing familiar words"),
        (float("-inf"), "suggests significant difficulties rapidly recognizing familiar words"),
    ],
    "towre_pde_standard": [
        (121, "suggests an advanced ability to rapidly decode unfamiliar words"),
        (111, "suggests a strong ability to rapidly decode unfamiliar words"),
        (90, "suggests a typical ability to rapidly decode unfamiliar words"),
        (80, "suggests some difficulties rapidly decoding unfamiliar words"),
        (float("-inf"), "suggests significant difficulties rapidly decoding unfamiliar words"),
    ],
    "ppvt_standard": [
        (131, "suggests advanced receptive vocabulary for their age"),
        (116, "suggests strong receptive vocabulary for their age"),
        (85, "suggests typical receptive vocabulary for their age"),
        (70, "suggests some difficulty understanding receptive vocabulary for their age"),
        (float("-inf"), "suggests significant difficulty understanding receptive vocabulary for their age"),
    ],
    "dibels_orf_words_correct": [
        (76, "suggests a strong ability to read connected text fluently"),
        (39, "suggests a typical ability to read connected text fluently"),
        (26, "suggests some difficulties with the ability to read connected text fluently"),
        (float("-inf"), "suggests significant difficulties with the ability to read connected text fluently"),
    ],
}


# Interpretation fields = Average, Below Average, etc. for table column
TEST_INTERPRETATION_RULES = {
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
    "dibels_orf_words_correct": [
        (76, "Above Average"),
        (39, "Average"),
        (26, "Below Average"),
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
    "ppvt_standard": "ppvt_category",
    "dibels_orf_words_correct": "dibels_orf_words_correct_category",
}


def clean_value(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def classify_score(value: Any, rules: list[tuple[float, str]]) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return ""

    try:
        score = float(str(value).strip())
    except ValueError:
        return ""

    for cutoff, label in rules:
        if score >= cutoff:
            return label

    return ""


def format_date(value: Any) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return ""

    parsed_date = pd.to_datetime(value, errors="coerce")

    if pd.isna(parsed_date):
        return clean_value(value)

    return parsed_date.strftime("%B %d, %Y")


def build_context(row: pd.Series) -> dict:
    context = {}

    for csv_field, template_field in FIELD_MAP.items():
        context[template_field] = clean_value(row.get(csv_field))

    context["visit_date"] = format_date(row.get("visit_date"))

    for score_field, category_field in CATEGORY_FIELDS.items():
        # Paragraph sentence placeholders, e.g. {{ kbit_category }}
        context[category_field] = classify_score(
            row.get(score_field),
            TEST_CATEGORY_RULES[score_field],
        )

        # Table interpretation placeholders, e.g. {{ kbit_interpretation }}
        interpretation_field = category_field.replace("_category", "_interpretation")
        context[interpretation_field] = classify_score(
            row.get(score_field),
            TEST_INTERPRETATION_RULES[score_field],
        )

    return context


def create_document(template_path: str | Path, context: dict) -> BytesIO:
    template = DocxTemplate(template_path)
    template.render(context)

    output = BytesIO()
    template.save(output)
    output.seek(0)

    return output
