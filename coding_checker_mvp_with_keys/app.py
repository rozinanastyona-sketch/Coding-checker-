from __future__ import annotations

from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st

from checker_engine import (
    add_reference_to_library,
    compare_files,
    create_reference_passport,
    load_grammar,
    load_reference_passports,
    read_excel_first_sheet_or_named,
    suggest_references,
    write_annotated_excel,
)

ROOT = Path(__file__).parent
GRAMMAR_PATH = ROOT / "grammar.yaml"
REFERENCE_DIR = ROOT / "reference_keys"

st.set_page_config(page_title="Coding Checker MVP", layout="wide")
st.title("Coding Checker MVP")

grammar = load_grammar(GRAMMAR_PATH)
REFERENCE_DIR.mkdir(exist_ok=True)

page = st.sidebar.radio("Page", ["Check student file", "Reference Library"])

if page == "Reference Library":
    st.header("Reference Library")
    st.caption("Add key files once. The app creates a passport with the first 5 utterances.")

    uploaded_ref = st.file_uploader("Add reference Excel file", type=["xlsx"], key="ref_upload")
    display_name = st.text_input("Display name (optional)")

    if uploaded_ref is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_ref.getbuffer())
            tmp_path = Path(tmp.name)
        try:
            passport_preview = create_reference_passport(tmp_path, grammar, display_name=display_name or None)
            st.subheader("Passport preview")
            st.write(passport_preview)
            if st.button("Save this reference"):
                saved_path = REFERENCE_DIR / uploaded_ref.name
                saved_path.write_bytes(uploaded_ref.getbuffer())
                passport = create_reference_passport(saved_path, grammar, display_name=display_name or None)
                saved_path.with_suffix(".json").write_text(
                    __import__("json").dumps(passport, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
                st.success("Reference saved.")
        except Exception as e:
            st.error(f"Could not create passport: {e}")

    st.subheader("Existing references")
    passports = load_reference_passports(REFERENCE_DIR)
    if not passports:
        st.info("No references yet.")
    else:
        for p in passports:
            with st.expander(p.get("display_name", p.get("id", "Reference"))):
                st.write(f"File: {p.get('file_name')}")
                st.write(f"Utterances: {p.get('utterance_count')}")
                st.write("First utterances:")
                for utt in p.get("first_utterances", []):
                    st.write(f"- {utt}")

else:
    st.header("Check student file")
    student_upload = st.file_uploader("Upload student Observer Excel file", type=["xlsx"])

    if student_upload is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(student_upload.getbuffer())
            student_path = Path(tmp.name)

        try:
            student_df = read_excel_first_sheet_or_named(student_path)
            student_first, suggestions = suggest_references(student_df, grammar, REFERENCE_DIR, top_k=5)

            st.subheader("First utterances detected in student file")
            if student_first:
                for utt in student_first:
                    st.write(f"- {utt}")
            else:
                st.warning("No utterance text detected. You can still choose a reference manually.")

            passports = load_reference_passports(REFERENCE_DIR)
            if not passports:
                st.error("No reference keys in library yet. Add keys on the Reference Library page.")
                st.stop()

            st.subheader("Choose reference")
            ordered = suggestions + [p for p in passports if p not in suggestions]
            labels = []
            for p in ordered:
                first = " · ".join(p.get("first_utterances", [])[:3])
                labels.append(f"{p.get('display_name', p.get('id'))} — {first}")

            choice = st.selectbox("Reference key", options=list(range(len(ordered))), format_func=lambda i: labels[i])
            selected = ordered[choice]
            st.write("Selected reference passport:", selected)

            if st.button("Run comparison"):
                ref_path = REFERENCE_DIR / selected["file_name"]
                key_df = read_excel_first_sheet_or_named(ref_path, sheet_name=selected.get("sheet_name"))
                issues, student_utts, key_utts = compare_files(student_df, key_df, grammar)

                st.subheader("Summary")
                st.write(f"Student utterances: {len(student_utts)}")
                st.write(f"Key utterances: {len(key_utts)}")
                st.write(f"Issues found: {len(issues)}")

                if issues:
                    st.dataframe(pd.DataFrame([i.__dict__ for i in issues]))
                else:
                    st.success("No issues found by MVP checker.")

                output_path = student_path.with_name("student_checked.xlsx")
                write_annotated_excel(student_path, output_path, issues, grammar)
                st.download_button(
                    "Download checked Excel",
                    data=output_path.read_bytes(),
                    file_name=f"{Path(student_upload.name).stem}_checked.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error(f"Error: {e}")
