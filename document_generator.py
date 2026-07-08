from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
from docxtpl import DocxTemplate

from field_mapping import FIELD_MAP


from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

def remove_content_controls(docx_bytes: BytesIO) -> BytesIO:
    docx_bytes.seek(0)
    cleaned_output = BytesIO()

    with ZipFile(docx_bytes, "r") as zin:
        with ZipFile(cleaned_output, "w", ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    try:
                        root = etree.fromstring(data)
                        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

                        for sdt in root.xpath(".//w:sdt", namespaces=ns):
                            parent = sdt.getparent()
                            sdt_content = sdt.find("w:sdtContent", namespaces=ns)

                            if parent is not None and sdt_content is not None:
                                index = parent.index(sdt)

                                for child in list(sdt_content):
                                    parent.insert(index, child)
                                    index += 1

                                parent.remove(sdt)

                        data = etree.tostring(
                            root,
                            xml_declaration=True,
                            encoding="UTF-8",
                            standalone="yes"
                        )
                    except Exception:
                        pass

                zout.writestr(item, data)

    cleaned_output.seek(0)
    return cleaned_output
# -------------------------
# General assessment rules
# -------------------------

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
}


TEST_INTERPRETATION_RULES = {
    "kbit_standard": [(131, "Superior"), (116, "Above Average"), (85, "Average"), (70, "Below Average"), (float("-inf"), "Well Below Average")],
    "ctopp_elision_standard": [(15, "Superior"), (13, "Above Average"), (8, "Average"), (6, "Below Average"), (float("-inf"), "Well Below Average")],
    "ctopp_nwr_standard": [(15, "Superior"), (13, "Above Average"), (8, "Average"), (6, "Below Average"), (float("-inf"), "Well Below Average")],
    "wrmt_word_id_standard": [(131, "Superior"), (116, "Above Average"), (85, "Average"), (70, "Below Average"), (float("-inf"), "Well Below Average")],
    "wrmt_word_attack_standard": [(131, "Superior"), (116, "Above Average"), (85, "Average"), (70, "Below Average"), (float("-inf"), "Well Below Average")],
    "wrmt_pc_standard": [(131, "Superior"), (116, "Above Average"), (85, "Average"), (70, "Below Average"), (float("-inf"), "Well Below Average")],
    "towre_swe_standard": [(121, "Superior"), (111, "Above Average"), (90, "Average"), (80, "Below Average"), (float("-inf"), "Well Below Average")],
    "towre_pde_standard": [(121, "Superior"), (111, "Above Average"), (90, "Average"), (80, "Below Average"), (float("-inf"), "Well Below Average")],
    "ppvt_standard": [(131, "Superior"), (116, "Above Average"), (85, "Average"), (70, "Below Average"), (float("-inf"), "Well Below Average")],
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
}


# -------------------------
# DIBELS grade-based rules
# -------------------------

DIBELS_WORDS_CORRECT_SENTENCES = {
    "1": [
        (76, "suggests a strong ability to read connected text fluently"),
        (39, "suggests a typical ability to read connected text fluently"),
        (26, "suggests some difficulties with the ability to read connected text fluently"),
        (float("-inf"), "suggests significant difficulties with the ability to read connected text fluently"),
    ],
    "2": [
        (128, "suggests a strong ability to read connected text fluently"),
        (94, "suggests a typical ability to read connected text fluently"),
        (77, "suggests some difficulties with the ability to read connected text fluently"),
        (float("-inf"), "suggests significant difficulties with the ability to read connected text fluently"),
    ],
    "3": [
        (136, "suggests a strong ability to read connected text fluently"),
        (114, "suggests a typical ability to read connected text fluently"),
        (96, "suggests some difficulties with the ability to read connected text fluently"),
        (float("-inf"), "suggests significant difficulties with the ability to read connected text fluently"),
    ],
    "4": [
        (159, "suggests a strong ability to read connected text fluently"),
        (125, "suggests a typical ability to read connected text fluently"),
        (99, "suggests some difficulties with the ability to read connected text fluently"),
        (float("-inf"), "suggests significant difficulties with the ability to read connected text fluently"),
    ],
    "5": [
        (157, "suggests a strong ability to read connected text fluently"),
        (137, "suggests a typical ability to read connected text fluently"),
        (124, "suggests some difficulties with the ability to read connected text fluently"),
        (float("-inf"), "suggests significant difficulties with the ability to read connected text fluently"),
    ],
}


DIBELS_WORDS_CORRECT_INTERPRETATIONS = {
    "1": [(76, "Above Average"), (39, "Average"), (26, "Below Average"), (float("-inf"), "Well Below Average")],
    "2": [(128, "Above Average"), (94, "Average"), (77, "Below Average"), (float("-inf"), "Well Below Average")],
    "3": [(136, "Above Average"), (114, "Average"), (96, "Below Average"), (float("-inf"), "Well Below Average")],
    "4": [(159, "Above Average"), (125, "Average"), (99, "Below Average"), (float("-inf"), "Well Below Average")],
    "5": [(157, "Above Average"), (137, "Average"), (124, "Below Average"), (float("-inf"), "Well Below Average")],
}


DIBELS_WORDS_CORRECT_TYPICAL_RANGE = {
    "1": "39–75",
    "2": "94–127",
    "3": "114–135",
    "4": "125–158",
    "5": "137–156",
}


DIBELS_ACCURACY_SENTENCES = {
    "1": [
        (91, "suggests a typical level of accuracy when reading connected text"),
        (85, "suggests some difficulties with accuracy when reading connected text"),
        (float("-inf"), "suggests significant difficulties with accuracy when reading connected text"),
    ],
    "2": [
        (96, "suggests a typical level of accuracy when reading connected text"),
        (85, "suggests some difficulties with accuracy when reading connected text"),
        (float("-inf"), "suggests significant difficulties with accuracy when reading connected text"),
    ],
    "3": [
        (96, "suggests a typical level of accuracy when reading connected text"),
        (91, "suggests some difficulties with accuracy when reading connected text"),
        (float("-inf"), "suggests significant difficulties with accuracy when reading connected text"),
    ],
    "4": [
        (96, "suggests a typical level of accuracy when reading connected text"),
        (91, "suggests some difficulties with accuracy when reading connected text"),
        (float("-inf"), "suggests significant difficulties with accuracy when reading connected text"),
    ],
    "5": [
        (96, "suggests a typical level of accuracy when reading connected text"),
        (91, "suggests some difficulties with accuracy when reading connected text"),
        (float("-inf"), "suggests significant difficulties with accuracy when reading connected text"),
    ],
}


DIBELS_ACCURACY_INTERPRETATIONS = {
    "1": [(91, "Average"), (85, "Below Average"), (float("-inf"), "Well Below Average")],
    "2": [(96, "Average"), (85, "Below Average"), (float("-inf"), "Well Below Average")],
    "3": [(96, "Average"), (91, "Below Average"), (float("-inf"), "Well Below Average")],
    "4": [(96, "Average"), (91, "Below Average"), (float("-inf"), "Well Below Average")],
    "5": [(96, "Average"), (91, "Below Average"), (float("-inf"), "Well Below Average")],
}


DIBELS_ACCURACY_TYPICAL_RANGE = {
    "1": "91–100%",
    "2": "96–100%",
    "3": "96–100%",
    "4": "96–100%",
    "5": "96–100%",
}


# -------------------------
# Helper functions
# -------------------------

def clean_value(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def classify_score(value: Any, rules: list[tuple[float, str]]) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return ""

    try:
        score = float(str(value).replace("%", "").strip())
    except ValueError:
        return ""

    for cutoff, label in rules:
        if score >= cutoff:
            return label

    return ""


def calculate_dorf_accuracy(words_correct: Any, total_words: Any) -> str:
    if pd.isna(words_correct) or pd.isna(total_words):
        return ""

    try:
        correct = float(str(words_correct).strip())
        total = float(str(total_words).strip())

        if total == 0:
            return ""

        accuracy = (correct / total) * 100
        return f"{accuracy:.0f}%"

    except ValueError:
        return ""


def get_grade(value: Any) -> str:
    value = clean_value(value).lower()

    if value in {"1", "1.0", "1st", "first", "grade 1", "1st grade"}:
        return "1"
    if value in {"2", "2.0", "2nd", "second", "grade 2", "2nd grade"}:
        return "2"
    if value in {"3", "3.0", "3rd", "third", "grade 3", "3rd grade"}:
        return "3"
    if value in {"4", "4.0", "4th", "fourth", "grade 4", "4th grade"}:
        return "4"
    if value in {"5", "5.0", "5th", "fifth", "grade 5", "5th grade"}:
        return "5"

    return ""


def get_student_grade(row: pd.Series) -> str:
    possible_grade_fields = [
        "grade_2025_2026",
        "Grade in 2025-2026 school year",
        "grade",
        "student_grade",
    ]

    for field in possible_grade_fields:
        grade = get_grade(row.get(field))
        if grade:
            return grade

    return ""


def format_date(value: Any) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return ""

    parsed_date = pd.to_datetime(value, errors="coerce")

    if pd.isna(parsed_date):
        return clean_value(value)

    return parsed_date.strftime("%B %d, %Y")


# -------------------------
# Main context builder
# -------------------------

def build_context(row: pd.Series) -> dict:
    context = {}

    for csv_field, template_field in FIELD_MAP.items():
        context[template_field] = clean_value(row.get(csv_field))

    context["visit_date"] = format_date(row.get("doe1"))
    context["age"] = clean_value(row.get("calc_age1"))

    for score_field, category_field in CATEGORY_FIELDS.items():
        context[category_field] = classify_score(
            row.get(score_field),
            TEST_CATEGORY_RULES[score_field],
        )

        interpretation_field = category_field.replace("_category", "_interpretation")

        context[interpretation_field] = classify_score(
            row.get(score_field),
            TEST_INTERPRETATION_RULES[score_field],
        )

    grade = get_student_grade(row)

    context["dibels_orf_words_correct_typical_range"] = DIBELS_WORDS_CORRECT_TYPICAL_RANGE.get(grade, "")

    context["dibels_orf_words_correct_category"] = classify_score(
        row.get("dibels_orf_words_correct"),
        DIBELS_WORDS_CORRECT_SENTENCES.get(grade, []),
    )

    context["dibels_orf_words_correct_interpretation"] = classify_score(
        row.get("dibels_orf_words_correct"),
        DIBELS_WORDS_CORRECT_INTERPRETATIONS.get(grade, []),
    )

    context["dibels_orf_accuracy"] = calculate_dorf_accuracy(
        row.get("dibels_orf_words_correct"),
        row.get("dibels_orf_total_words"),
    )

    context["dibels_orf_accuracy_typical_range"] = DIBELS_ACCURACY_TYPICAL_RANGE.get(grade, "")

    context["dibels_orf_accuracy_category"] = classify_score(
        context.get("dibels_orf_accuracy"),
        DIBELS_ACCURACY_SENTENCES.get(grade, []),
    )

    context["dibels_orf_accuracy_interpretation"] = classify_score(
        context.get("dibels_orf_accuracy"),
        DIBELS_ACCURACY_INTERPRETATIONS.get(grade, []),
    )

    return context


# -------------------------
# Document creation
# -------------------------

def create_document(template_path: str | Path, context: dict) -> BytesIO:
    template = DocxTemplate(template_path)
    template.render(context)

    output = BytesIO()
    template.save(output)
    output.seek(0)

    output = remove_content_controls(output)

    return output
