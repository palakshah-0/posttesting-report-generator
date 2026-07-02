from pathlib import Path

import pandas as pd
import streamlit as st

from document_generator import build_context, create_document


TEMPLATE_FOLDER = Path("templates")


st.set_page_config(
    page_title="REDCap Document Generator",
    page_icon="📄",
)


st.title("REDCap Word Document Generator")

st.write(
    "Upload a CSV export from REDCap, select a record, "
    "and download a completed Word document."
)


uploaded_file = st.file_uploader(
    "Upload REDCap CSV file",
    type=["csv"],
)


if uploaded_file is not None:
    try:
        data = pd.read_csv(
            uploaded_file,
            dtype=str,
    )
        # Remove leading/trailing spaces from column names
        data.columns = data.columns.str.strip()

    # Rename REDCap export columns to the names expected by the app
        data = data.rename(columns={
            "Amira_ID": "sub_id_number",
            "Matrices Standard": "kbit_standard",
            "Elision Words Standard": "ctopp_elision_standard",
            "Nonword Repetition Standard": "ctopp_nwr_standard",
            "Oral Reading Fluency Total Words": "dibels_orf_total_words",
            "Oral Reading Fluency Total Errors": "dibels_orf_total_errors",
            "Oral Reading Fluency Total Words Correct": "dibels_orf_words_correct",
            "Word Identification Standard": "wrmt_word_id_standard",
            "Word Attack Standard": "wrmt_word_attack_standard",
            "Passage Comprehension Standard": "wrmt_pc_standard",
            "Sight word efficiency standard": "towre_swe_standard",
            "Phonemic decoding efficiency standard": "towre_pde_standard",
            "PPVT Standard": "ppvt_standard",
        })

    except Exception as error:
        st.error(f"The CSV could not be read: {error}")
        st.stop()

    st.success("CSV uploaded successfully.")

    st.subheader("CSV preview")
    st.dataframe(data.head())

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
        "ppvt_standard"
    }

    
    missing_columns = required_columns - set(data.columns)

    if missing_columns:
        st.error(
            "The CSV is missing these columns: "
            + ", ".join(sorted(missing_columns))
        )
        st.stop()

    if data.empty:
        st.error("The uploaded CSV does not contain any records.")
        st.stop()

    record_ids = (
        data["sub_id_number"]
        .dropna()
        .drop_duplicates()
        .tolist()
    )

    if not record_ids:
        st.error("No record IDs were found in the CSV.")
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

    context = build_context(selected_row)

    st.subheader("Selected record")

    st.write(f"kbit_standard: {context['kbit_standard']}")
    st.write(f"ctopp_elision_standard: {context['ctopp_elision_standard']}")
    st.write(f"ctopp_nwr_standard: {context['ctopp_nwr_standard']}")
    st.write(f"dibels_orf_total_words: {context['dibels_orf_total_words']}")
    st.write(f"dibels_orf_total_errors: {context['dibels_orf_total_errors']}")
    st.write(f"dibels_orf_words_correct: {context['dibels_orf_words_correct']}")
    st.write(f"wrmt_word_id_standard: {context['wrmt_word_id_standard']}")
    st.write(f"wrmt_word_attack_standard: {context['wrmt_word_attack_standard']}")
    st.write(f"wrmt_pc_standard: {context['wrmt_pc_standard']}")
    st.write(f"towre_swe_standard: {context['towre_swe_standard']}")
    st.write(f"towre_pde_standard: {context['towre_pde_standard']}")
    st.write(f"ppvt_standard: {context['ppvt_standard']}")

    template_files = list(TEMPLATE_FOLDER.glob("*.docx"))

    if not template_files:
        st.error(
        "The Word template was not found. "
        "Make sure report_template.docx is inside the templates folder."
        )
        st.stop()

    TEMPLATE_PATH = template_files[0]

    try:
        completed_document = create_document(
            TEMPLATE_PATH,
            context,
        )
    except Exception as error:
        st.error(f"The Word document could not be created: {error}")
        st.stop()

    safe_record_id = str(selected_record).replace("/", "-")

    st.download_button(
        label="Download completed Word document",
        data=completed_document,
        file_name=f"report_{safe_record_id}.docx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
    )
