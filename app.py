"""Web interface for the spreadsheet cleaner.

Run with:
    streamlit run app.py

Uploaded files are only ever held in memory; nothing is written to disk.

Set DEMO_MODE to "true" (environment variable or Streamlit secret) to cap uploads
for a public demo. Without it there are no limits.
"""

import io
import os

import pandas as pd
import streamlit as st

from clean import clean, detect_headings, read_file, write_workbook
from make_demo_data import build_demo_files

NO_KEY = "None (only remove exact duplicates)"
PREVIEW_ROWS = 100
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

DEMO_MAX_FILES = 3
DEMO_MAX_ROWS = 200
DEMO_MESSAGE = "This is a demo. For a version tailored to your data, contact info.sendovia@gmail.com"


def demo_mode():
    """True when DEMO_MODE is "true" in the environment or in Streamlit secrets."""
    value = os.environ.get("DEMO_MODE")
    if value is None:
        try:
            value = st.secrets.get("DEMO_MODE")
        except Exception:
            # No secrets file, which is the normal case when running locally.
            value = None
    return str(value).strip().lower() == "true"


def as_sources(files):
    """Turn (name, bytes) pairs into the (name, open file) pairs clean.py expects."""
    return [(name, io.BytesIO(content)) for name, content in files]


@st.cache_data(show_spinner=False)
def demo_files():
    return sorted(build_demo_files().items())


@st.cache_data(show_spinner=False)
def headings_for(files):
    return detect_headings(as_sources(files))


@st.cache_data(show_spinner=False)
def over_demo_limits(files):
    """Describe how the uploads exceed the demo limits, or return None if they don't."""
    if len(files) > DEMO_MAX_FILES:
        return f"You uploaded {len(files)} files; the demo allows up to {DEMO_MAX_FILES}."
    for name, source in as_sources(files):
        try:
            rows = len(read_file(name, source))
        except Exception:
            continue  # unreadable files are reported when cleaning
        if rows > DEMO_MAX_ROWS:
            return f"{name} has {rows} rows; the demo allows up to {DEMO_MAX_ROWS} rows per file."
    return None


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

DEMO = demo_mode()


def page():
    uploads = st.file_uploader("Upload .xlsx or .csv files", type=["xlsx", "csv"], accept_multiple_files=True)
    if DEMO:
        st.caption(f"Demo limits: up to {DEMO_MAX_FILES} files, {DEMO_MAX_ROWS} rows per file.")
    if st.button("Load demo files", help="Five made-up customer lists with messy headings, dates and duplicates."):
        st.session_state["demo"] = True

    if uploads:
        # Alphabetical, like the command-line tool, so "first value wins" behaves the same.
        files = tuple(sorted((upload.name, upload.getvalue()) for upload in uploads))
        problem = over_demo_limits(files) if DEMO else None
        if problem:
            st.error(f"{problem} {DEMO_MESSAGE}")
            return
    elif st.session_state.get("demo"):
        files = tuple(demo_files())
        st.info("Using the demo files: " + ", ".join(name for name, _ in files))
        if st.button("Remove demo files"):
            st.session_state["demo"] = False
            st.rerun()
    else:
        return

    headings = headings_for(files)
    if not headings:
        st.error("None of those files could be read. Check they are real .xlsx or .csv files.")
        return

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


page()
if DEMO:
    st.divider()
    st.caption(DEMO_MESSAGE)
