from pathlib import Path

import pandas as pd
import streamlit as st

from document_generator import build_context, create_document


# ---------------------------------------------------------
# SETTINGS
# ---------------------------------------------------------

TEMPLATE_FOLDER = Path("templates")


st.set_page_config(
    page_title="EVAL Post-Test Report Generator",
    page_icon="📄",
)


# ---------------------------------------------------------
# PAGE TITLE
# ---------------------------------------------------------

st.title(
    "EVAL Post-Test Report Generator"
)

st.write(
    "Upload a post-test REDCap CSV, select a record, "
    "and download a completed post-test Word report."
)


# ---------------------------------------------------------
# UPLOAD CSV
# ---------------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload post-test REDCap CSV file",
    type=["csv"],
)


if uploaded_file is not None:

    try:

        data = pd.read_csv(
            uploaded_file,
            dtype=str,
        )

        data.columns = (
            data.columns
            .str.strip()
        )

    except Exception as error:

        st.error(
            f"The CSV could not be read: {error}"
        )

        st.stop()


    # ---------------------------------------------------------
    # REDCAP COLUMN RENAMING
    # ---------------------------------------------------------
    #
    # Later, when you receive the scored REDCap export,
    # place the real REDCap score column names on the LEFT.
    #
    # The standardized generator field names stay on the RIGHT.
    # ---------------------------------------------------------

    data = data.rename(
        columns={

            # Example mappings for future score export:
            #
            # "Post Elision Standard Score":
            #     "ctopp_elision_std_post",
            #
            # "Post Nonword Repetition Standard Score":
            #     "ctopp_nwr_std_post",
            #
            # "Post DIBELS Words Correct":
            #     "dibels_orf_words_correct_post",
            #
            # "Post DIBELS Total Words":
            #     "dibels_total_words_post",
            #
            # "Post Word Identification Standard":
            #     "wrmt_wid_std_post",
            #
            # "Post Word Attack Standard":
            #     "wrmt_wa_std_post",
            #
            # "Post Passage Comprehension Standard":
            #     "wrmt_pc_std_post",
            #
            # "Post Sight Word Efficiency Standard":
            #     "towre_swe_std_post",
            #
            # "Post Phonemic Decoding Efficiency Standard":
            #     "towre_pde_std_post",
            #
            # "Post PPVT Standard":
            #     "ppvt_std_post",
        }
    )


    # ---------------------------------------------------------
    # EXPECTED SCORE COLUMNS
    # ---------------------------------------------------------
    #
    # If they do not exist yet, create them as blank columns.
    # This allows development/testing before final score data
    # becomes available.
    # ---------------------------------------------------------

    expected_score_columns = [
        "ctopp_elision_std_post",
        "ctopp_nwr_std_post",
        "dibels_orf_words_correct_post",
        "dibels_total_words_post",
        "wrmt_wid_std_post",
        "wrmt_wa_std_post",
        "wrmt_pc_std_post",
        "towre_swe_std_post",
        "towre_pde_std_post",
        "ppvt_std_post",
    ]


    for column in expected_score_columns:

        if column not in data.columns:
            data[column] = ""


    # ---------------------------------------------------------
    # CHECK RECORD ID
    # ---------------------------------------------------------

    if "sub_id_number" not in data.columns:

        st.error(
            "The CSV does not contain "
            "'sub_id_number'."
        )

        st.stop()


    if data.empty:

        st.error(
            "The CSV does not contain any records."
        )

        st.stop()


    st.success(
        "Post-test CSV uploaded successfully."
    )


    # ---------------------------------------------------------
    # CSV PREVIEW
    # ---------------------------------------------------------

    st.subheader(
        "CSV preview"
    )

    st.dataframe(
        data.head()
    )


    # ---------------------------------------------------------
    # SELECT RECORD
    # ---------------------------------------------------------

    record_ids = (
        data["sub_id_number"]
        .dropna()
        .astype(str)
        .str.strip()
    )


    record_ids = (
        record_ids[
            record_ids != ""
        ]
        .drop_duplicates()
        .tolist()
    )


    if not record_ids:

        st.error(
            "No record IDs were found."
        )

        st.stop()


    selected_record = (
        st.selectbox(
            "Select a record",
            record_ids,
        )
    )


    selected_rows = data[
        data["sub_id_number"]
        .astype(str)
        .str.strip()
        ==
        str(selected_record).strip()
    ]


    if selected_rows.empty:

        st.error(
            "The selected record "
            "could not be found."
        )

        st.stop()


    if len(selected_rows) > 1:

        st.warning(
            "This record appears more than once. "
            "The first matching row will be used."
        )


    selected_row = (
        selected_rows.iloc[0]
    )


    # ---------------------------------------------------------
    # BUILD REPORT DATA
    # ---------------------------------------------------------

    try:

        context = build_context(
            selected_row
        )

    except Exception as error:

        st.error(
            "The selected record could not "
            f"be processed: {error}"
        )

        st.stop()


    # ---------------------------------------------------------
    # SELECTED RECORD PREVIEW
    # ---------------------------------------------------------

    st.subheader(
        "Selected record"
    )


    st.write(
        f"**Record ID:** "
        f"{context.get('sub_id_number', '')}"
    )


    st.write(
        f"**CTOPP Elision:** "
        f"{context.get('ctopp_elision_std_post', '')}"
    )


    st.write(
        f"**CTOPP Nonword Repetition:** "
        f"{context.get('ctopp_nwr_std_post', '')}"
    )


    st.write(
        f"**DIBELS Words Correct:** "
        f"{context.get('dibels_orf_words_correct_post', '')}"
    )


    st.write(
        f"**DIBELS Accuracy:** "
        f"{context.get('dorf_accuracy', '')}"
    )


    st.write(
        f"**WRMT Word Identification:** "
        f"{context.get('wrmt_wid_std_post', '')}"
    )


    st.write(
        f"**WRMT Word Attack:** "
        f"{context.get('wrmt_wa_std_post', '')}"
    )


    st.write(
        f"**WRMT Passage Comprehension:** "
        f"{context.get('wrmt_pc_std_post', '')}"
    )


    st.write(
        f"**TOWRE Sight Word Efficiency:** "
        f"{context.get('towre_swe_std_post', '')}"
    )


    st.write(
        f"**TOWRE Phonemic Decoding:** "
        f"{context.get('towre_pde_std_post', '')}"
    )


    st.write(
        f"**PPVT:** "
        f"{context.get('ppvt_std_post', '')}"
    )


    # ---------------------------------------------------------
    # FIND TEMPLATE
    # ---------------------------------------------------------

    template_files = list(
        TEMPLATE_FOLDER.glob(
            "*.docx"
        )
    )


    if not template_files:

        st.error(
            "No Word template was found. "
            "Place the post-test .docx file "
            "inside the templates folder."
        )

        st.stop()


    if len(template_files) > 1:

        st.error(
            "More than one Word template was found. "
            "Keep only the post-test template "
            "inside the templates folder."
        )

        st.stop()


    template_path = (
        template_files[0]
    )


    # ---------------------------------------------------------
    # CREATE DOCUMENT
    # ---------------------------------------------------------

    try:

        completed_document = (
            create_document(
                template_path,
                context,
            )
        )

    except Exception as error:

        st.error(
            "The Word document could not "
            f"be created: {error}"
        )

        st.stop()


    # ---------------------------------------------------------
    # DOWNLOAD
    # ---------------------------------------------------------

    safe_record_id = (
        str(selected_record)
        .replace("/", "-")
        .replace("\\", "-")
        .replace(" ", "_")
    )


    st.download_button(
        label=(
            "Download completed "
            "post-test Word report"
        ),
        data=completed_document,
        file_name=(
            f"posttest_report_"
            f"{safe_record_id}.docx"
        ),
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
    )
