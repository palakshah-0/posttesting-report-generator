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

st.title("EVAL Post-Test Report Generator")

st.write(
    "Upload a post-test CSV export from REDCap, select a record, "
    "and download the completed post-test Word report."
)


# ---------------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload post-test REDCap CSV file",
    type=["csv"],
)


# ---------------------------------------------------------
# PROCESS CSV
# ---------------------------------------------------------

if uploaded_file is not None:

    try:
        data = pd.read_csv(
            uploaded_file,
            dtype=str,
        )

        # Remove leading/trailing spaces from CSV column names
        data.columns = data.columns.str.strip()

        # -------------------------------------------------
        # RENAME REDCAP COLUMNS
        # -------------------------------------------------
        #
        # The LEFT side should match the column names
        # appearing in the REDCap CSV.
        #
        # The RIGHT side is the standardized field name
        # used by the post-test report generator.
        #
        # If your REDCap export already uses the names on
        # the right, this rename operation will simply leave
        # them unchanged.
        # -------------------------------------------------

        data = data.rename(
            columns={
                # Record ID
                "Amira_ID": "sub_id_number",

                # CTOPP-2
                "Elision Words Standard": "ctopp_elision_std_post",
                "Nonword Repetition Standard": "ctopp_nwr_std_post",

                # DIBELS
                "Oral Reading Fluency Total Words Correct":
                    "dibels_words_correct_post",

                "Oral Reading Fluency Total Words":
                    "dibels_total_words_post",

                # WRMT-III
                "Word Identification Standard":
                    "wrmt_wid_std_post",

                "Word Attack Standard":
                    "wrmt_wa_std_post",

                "Passage Comprehension Standard":
                    "wrmt_pc_std_post",

                # TOWRE-2
                "Sight word efficiency standard":
                    "towre_swe_std_post",

                "Phonemic decoding efficiency standard":
                    "towre_pde_std_post",

                # PPVT-5
                "PPVT Standard":
                    "ppvt_std_post",
            }
        )

    except Exception as error:
        st.error(
            f"The CSV could not be read: {error}"
        )
        st.stop()


    st.success(
        "Post-test CSV uploaded successfully."
    )


    # ---------------------------------------------------------
    # CSV PREVIEW
    # ---------------------------------------------------------

    st.subheader("CSV preview")

    st.dataframe(
        data.head()
    )


    # ---------------------------------------------------------
    # REQUIRED COLUMNS
    # ---------------------------------------------------------

    required_columns = {
        "sub_id_number",
        "ctopp_elision_std_post",
        "ctopp_nwr_std_post",
        "dibels_words_correct_post",
        "dibels_total_words_post",
        "wrmt_wid_std_post",
        "wrmt_wa_std_post",
        "wrmt_pc_std_post",
        "towre_swe_std_post",
        "towre_pde_std_post",
        "ppvt_std_post",
    }


    missing_columns = (
        required_columns
        - set(data.columns)
    )


    if missing_columns:

        st.error(
            "The CSV is missing these required columns: "
            + ", ".join(
                sorted(missing_columns)
            )
        )

        st.write(
            "Check that the REDCap column names match "
            "the names expected by the app."
        )

        st.stop()


    # ---------------------------------------------------------
    # CHECK FOR EMPTY CSV
    # ---------------------------------------------------------

    if data.empty:

        st.error(
            "The uploaded CSV does not contain any records."
        )

        st.stop()


    # ---------------------------------------------------------
    # FIND RECORD IDS
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
            "No record IDs were found in the CSV."
        )

        st.stop()


    # ---------------------------------------------------------
    # SELECT RECORD
    # ---------------------------------------------------------

    selected_record = st.selectbox(
        "Select a record",
        record_ids,
    )


    selected_rows = data[
        data["sub_id_number"].astype(str).str.strip()
        == str(selected_record).strip()
    ]


    if selected_rows.empty:

        st.error(
            "The selected record could not be found."
        )

        st.stop()


    if len(selected_rows) > 1:

        st.warning(
            "This record appears more than once in the CSV. "
            "The app is using the first matching row."
        )


    selected_row = (
        selected_rows.iloc[0]
    )


    # ---------------------------------------------------------
    # BUILD TEMPLATE CONTEXT
    # ---------------------------------------------------------

    try:

        context = build_context(
            selected_row
        )

    except Exception as error:

        st.error(
            "The selected student's data could not be processed. "
            f"Error: {error}"
        )

        st.stop()


    # ---------------------------------------------------------
    # SHOW SELECTED RECORD
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
        f"{context.get('ctopp_elision_std_post', '')} "
        f"— {context.get('ctopp_elision_category', '')}"
    )


    st.write(
        f"**CTOPP Nonword Repetition:** "
        f"{context.get('ctopp_nwr_std_post', '')} "
        f"— {context.get('ctopp_nwr_category', '')}"
    )


    st.write(
        f"**DIBELS Words Correct:** "
        f"{context.get('dibels_words_correct_post', '')}"
    )


    st.write(
        f"**DIBELS Total Words:** "
        f"{context.get('dibels_total_words_post', '')}"
    )


    st.write(
        f"**DIBELS Accuracy:** "
        f"{context.get('dorf_accuracy', '')}"
    )


    st.write(
        f"**WRMT Word Identification:** "
        f"{context.get('wrmt_wid_std_post', '')} "
        f"— {context.get('wrmt_word_id_category', '')}"
    )


    st.write(
        f"**WRMT Word Attack:** "
        f"{context.get('wrmt_wa_std_post', '')} "
        f"— {context.get('wrmt_word_attack_category', '')}"
    )


    st.write(
        f"**WRMT Passage Comprehension:** "
        f"{context.get('wrmt_pc_std_post', '')} "
        f"— {context.get('wrmt_pc_category', '')}"
    )


    st.write(
        f"**TOWRE Sight Word Efficiency:** "
        f"{context.get('towre_swe_std_post', '')} "
        f"— {context.get('towre_swe_category', '')}"
    )


    st.write(
        f"**TOWRE Phonemic Decoding Efficiency:** "
        f"{context.get('towre_pde_std_post', '')} "
        f"— {context.get('towre_pde_category', '')}"
    )


    st.write(
        f"**PPVT:** "
        f"{context.get('ppvt_std_post', '')} "
        f"— {context.get('ppvt_category', '')}"
    )


    # ---------------------------------------------------------
    # FIND WORD TEMPLATE
    # ---------------------------------------------------------

    template_files = list(
        TEMPLATE_FOLDER.glob("*.docx")
    )


    if not template_files:

        st.error(
            "The post-test Word template was not found. "
            "Make sure a .docx template is inside "
            "the templates folder."
        )

        st.stop()


    if len(template_files) > 1:

        st.error(
            "More than one Word template was found in the "
            "templates folder. Keep only the post-test "
            "template in this folder."
        )

        st.stop()


    template_path = (
        template_files[0]
    )


    # ---------------------------------------------------------
    # GENERATE WORD DOCUMENT
    # ---------------------------------------------------------

    try:

        completed_document = create_document(
            template_path,
            context,
        )

    except Exception as error:

        st.error(
            "The Word document could not be created. "
            f"Error: {error}"
        )

        st.stop()


    # ---------------------------------------------------------
    # DOWNLOAD FILE
    # ---------------------------------------------------------

    safe_record_id = (
        str(selected_record)
        .replace("/", "-")
        .replace("\\", "-")
        .replace(" ", "_")
    )


    st.download_button(
        label="Download completed post-test Word report",
        data=completed_document,
        file_name=(
            f"posttest_report_{safe_record_id}.docx"
        ),
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
    )
