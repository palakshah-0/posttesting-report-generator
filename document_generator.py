from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
from docxtpl import DocxTemplate

from field_mapping import FIELD_MAP

# I never even touched the code in this file

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