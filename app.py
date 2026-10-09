"""Web interface for the spreadsheet cleaner.

Run with:
    streamlit run app.py

Uploaded files are only ever held in memory; nothing is written to disk.
"""

import io

import pandas as pd
import streamlit as st

from clean import clean, detect_headings, write_workbook
from make_demo_data import build_demo_files

NO_KEY = "None (only remove exact duplicates)"
PREVIEW_ROWS = 100
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def as_sources(files):
    """Turn (name, bytes) pairs into the (name, open file) pairs clean.py expects."""
    return [(name, io.BytesIO(content)) for name, content in files]


@st.cache_data(show_spinner=False)
def demo_files():
    return sorted(build_demo_files().items())


@st.cache_data(show_spinner=False)
def headings_for(files):
    return detect_headings(as_sources(files))


def run_clean(files, key, dayfirst):
    table, summary, per_file, skipped = clean(as_sources(files), key=key, dayfirst=dayfirst)
    workbook = io.BytesIO()
    write_workbook(table, summary, per_file, workbook)
    return {"table": table, "summary": summary, "per_file": per_file, "skipped": skipped, "workbook": workbook.getvalue()}


def show_result(result):
    summary = dict(result["summary"])
    merged_label = next((label for label in summary if label.startswith("Rows merged by key")), None)

    st.subheader("Summary")
    shown = ["Files read", "Rows in", "Duplicates removed"] + ([merged_label] if merged_label else []) + ["Rows out"]
    for column, label in zip(st.columns(len(shown)), shown):
        column.metric(label, summary[label])

    notes = [
        f"{summary[label]} {text}"
        for label, text in [
            ("Blank rows removed", "blank rows removed"),
            ("Dates not recognized (left as-is)", "dates not recognized (left as-is)"),
            ("Phone numbers not recognized (left as-is)", "phone numbers not recognized (left as-is)"),
        ]
        if summary[label]
    ]
    if notes:
        st.caption(" · ".join(notes))
    for name, error in result["skipped"]:
        st.warning(f"Skipped {name}: {error}")
    with st.expander("Rows read per file"):
        st.dataframe(pd.DataFrame(result["per_file"], columns=["File", "Rows read"]), hide_index=True)

    table = result["table"]
    st.subheader("Preview")
    if len(table) > PREVIEW_ROWS:
        st.caption(f"Showing the first {PREVIEW_ROWS} of {len(table)} rows. The download has all of them.")
    st.dataframe(table.head(PREVIEW_ROWS), hide_index=True)

    st.download_button("Download cleaned.xlsx", result["workbook"], file_name="cleaned.xlsx", mime=XLSX_MIME, type="primary")


st.set_page_config(page_title="Spreadsheet Cleaner", layout="wide")
st.title("Spreadsheet Cleaner")
st.write(
    "Combine messy customer lists into one clean Excel file: matching column headings, "
    "tidy dates and phone numbers, duplicates removed."
)

uploads = st.file_uploader("Upload .xlsx or .csv files", type=["xlsx", "csv"], accept_multiple_files=True)
if st.button("Load demo files", help="Five made-up customer lists with messy headings, dates and duplicates."):
    st.session_state["demo"] = True

if uploads:
    # Alphabetical, like the command-line tool, so "first value wins" behaves the same.
    files = tuple(sorted((upload.name, upload.getvalue()) for upload in uploads))
elif st.session_state.get("demo"):
    files = tuple(demo_files())
    st.info("Using the demo files: " + ", ".join(name for name, _ in files))
    if st.button("Remove demo files"):
        st.session_state["demo"] = False
        st.rerun()
else:
    st.stop()

headings = headings_for(files)
if not headings:
    st.error("None of those files could be read. Check they are real .xlsx or .csv files.")
    st.stop()

key_choice = st.selectbox(
    "Merge rows that share the same value in this column",
    [NO_KEY] + headings,
    help="For example, choose Email to get one row per customer. Blanks in one file are filled from another.",
)
dayfirst = st.checkbox("Dates are day-first (03/04/2024 means 3 April, not March 4)")
key = None if key_choice == NO_KEY else key_choice

# Identifies exactly what was cleaned, so an old result is never shown against new inputs.
signature = (files, key, dayfirst)

if st.button("Clean", type="primary"):
    try:
        with st.spinner("Cleaning..."):
            st.session_state["result"] = (signature, run_clean(files, key, dayfirst))
    except ValueError as error:
        st.session_state.pop("result", None)
        st.error(str(error))

stored = st.session_state.get("result")
if stored and stored[0] == signature:
    show_result(stored[1])
elif stored:
    st.caption("Files or options have changed. Click Clean to update the result.")
