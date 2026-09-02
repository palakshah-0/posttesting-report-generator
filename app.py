from pathlib import Path

import pandas as pd
import streamlit as st

from document_generator import build_context, create_document


TEMPLATE_PATH = Path(
    "templates/TEST_of_EVAL_Pre_Testing_Template_2.docx"
)


st.set_page_config(
    page_title="REDCap Document Generator",
    page_icon="📄",
)


st.title("REDCap Word Document Generator")

st.write(
    "Upload a CSV export from REDCap, select a record, "
    "and download a completed Word document."
)


# ---------------------------------------------------------
# CHECK TEMPLATE
# ---------------------------------------------------------

if not TEMPLATE_PATH.exists():
    st.error(
        f"The pre-test Word template was not found: {TEMPLATE_PATH}"
    )
    st.stop()


# ---------------------------------------------------------
# UPLOAD CSV
# ---------------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload REDCap CSV file",
    type=["csv"],
)


# Do not continue until a CSV has been uploaded.
if uploaded_file is None:
    st.stop()


# ---------------------------------------------------------
# READ CSV
# ---------------------------------------------------------

try:
    data = pd.read_csv(
        uploaded_file,
        dtype=str,
    )

except Exception as error:
    st.error(
        f"The CSV could not be read: {error}"
    )
    st.stop()


st.success("CSV uploaded successfully.")


# ---------------------------------------------------------
# CSV PREVIEW
# ---------------------------------------------------------

st.subheader("CSV preview")

st.dataframe(
    data.head(),
    use_container_width=True,
)


# ---------------------------------------------------------
# REQUIRED COLUMNS
# ---------------------------------------------------------

required_columns = {
    "sub_id_number",
    "kbit_standard",
    "ctopp_elision_standard",
    "ctopp_nwr_standard",
    "dibels_orf_total_words",
    "dibels_orf_total_errors",
    "dibels_orf_words_correct",
    "wrmt_word_id_standard",
    "wrmt_word_attack_standard",
    "wrmt_pc_standard",
    "towre_swe_standard",
    "towre_pde_standard",
    "ppvt_standard",
}


missing_columns = (
    required_columns
    - set(data.columns)
)


if missing_columns:
    st.error(
        "The CSV is missing these columns: "
        + ", ".join(
            sorted(missing_columns)
        )
    )
    st.stop()


if data.empty:
    st.error(
        "The uploaded CSV does not contain any records."
    )
    st.stop()


# ---------------------------------------------------------
# SELECT RECORD
# ---------------------------------------------------------

record_ids = (
    data["sub_id_number"]
    .dropna()
    .drop_duplicates()
    .tolist()
)


if not record_ids:
    st.error(
        "No record IDs were found in the CSV."
    )
    st.stop()


selected_record = st.selectbox(
    "Select a record",
    record_ids,
)


selected_rows = data[
    data["sub_id_number"] == selected_record
]


if len(selected_rows) > 1:
    st.warning(
        "This record appears more than once in the CSV. "
        "The app is currently using the first matching row."
    )


selected_row = selected_rows.iloc[0]


# ---------------------------------------------------------
# BUILD CONTEXT
# ---------------------------------------------------------

context = build_context(
    selected_row
)


# ---------------------------------------------------------
# DISPLAY SELECTED RECORD
# ---------------------------------------------------------

st.subheader("Selected record")


st.write(
    f"kbit_standard: {context.get('kbit_standard', '')}"
)

st.write(
    f"ctopp_elision_standard: "
    f"{context.get('ctopp_elision_standard', '')}"
)

st.write(
    f"ctopp_nwr_standard: "
    f"{context.get('ctopp_nwr_standard', '')}"
)

st.write(
    f"dibels_orf_total_words: "
    f"{context.get('dibels_orf_total_words', '')}"
)

st.write(
    f"dibels_orf_total_errors: "
    f"{context.get('dibels_orf_total_errors', '')}"
)

st.write(
    f"dibels_orf_words_correct: "
    f"{context.get('dibels_orf_words_correct', '')}"
)

st.write(
    f"wrmt_word_id_standard: "
    f"{context.get('wrmt_word_id_standard', '')}"
)

st.write(
    f"wrmt_word_attack_standard: "
    f"{context.get('wrmt_word_attack_standard', '')}"
)

st.write(
    f"wrmt_pc_standard: "
    f"{context.get('wrmt_pc_standard', '')}"
)

st.write(
    f"towre_swe_standard: "
    f"{context.get('towre_swe_standard', '')}"
)

st.write(
    f"towre_pde_standard: "
    f"{context.get('towre_pde_standard', '')}"
)

st.write(
    f"ppvt_standard: "
    f"{context.get('ppvt_standard', '')}"
)


# ---------------------------------------------------------
# CREATE WORD DOCUMENT
# ---------------------------------------------------------

try:
    completed_document = create_document(
        TEMPLATE_PATH,
        context,
    )

except Exception as error:
    st.error(
        f"The Word document could not be created: {error}"
    )
    st.stop()


# ---------------------------------------------------------
# DOWNLOAD
# ---------------------------------------------------------

safe_record_id = str(
    selected_record
).replace("/", "-")


st.download_button(
    label="Download completed Word document",
    data=completed_document,
    file_name=f"report_{safe_record_id}.docx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "wordprocessingml.document"
    ),
)
