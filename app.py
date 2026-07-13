"""Coding Checker - Streamlit app.

Key persistence
---------------
Keys in reference_keys/ IN THE REPOSITORY survive every redeploy: the host
re-clones the repo on each update and index_reference_folder() rebuilds the
passports at startup. Keys uploaded through the UI go to the same folder, but on
a hosted server that folder is temporary - they vanish when the app restarts.
The Key Library page says so explicitly.
"""
from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st

import checker_engine as ce
from checker_engine import (
    add_reference_to_library,
    build_scores_rows,
    compare_files_with_alignment,
    create_reference_passport,
    index_reference_folder,
    load_grammar,
    load_reference_passports,
    read_excel_first_sheet_or_named,
    suggest_references,
    write_annotated_excel,
)

ROOT = Path(__file__).parent
GRAMMAR_PATH = ROOT / "grammar.yaml"
REFERENCE_DIR = ROOT / "reference_keys"

st.set_page_config(page_title="Coding Checker", layout="wide")

# ---------------------------------------------------------------------------
# Visual theme — applies the reviewed mockup design over native Streamlit.
# ---------------------------------------------------------------------------
THEME_CSS = """
<style>
:root{
  --brand:#e2604a; --brand-soft:#3a2622; --brand-text:#f0a189;
  --ink:#e8ebf1; --muted:#9aa4b5;
  --line:#333a48; --card:#262a36; --page:#1b1e27;
  --green:#7a9c54; --green-soft:#2b3324;
  --red:#c17356; --red-soft:#33241d;
}
/* Card-like buttons */
.stButton>button, .stDownloadButton>button{
  border-radius:10px; font-weight:600; border:1px solid var(--line);
  transition:.15s;
}
.stButton>button[kind="primary"], .stDownloadButton>button[kind="primary"]{
  border:none; background:var(--brand); color:#1b1013;
}
/* Metrics as cards */
[data-testid="stMetric"]{
  background:var(--card); border:1px solid var(--line); border-radius:12px;
  padding:14px 18px;
}
[data-testid="stMetricValue"]{ font-weight:700; }
/* Expanders as cards */
[data-testid="stExpander"]{
  border:1px solid var(--line); border-radius:10px;
  background:var(--card); margin-bottom:8px; overflow:hidden;
}
[data-testid="stExpander"] summary{ font-weight:500; }
/* Stepper */
.stepper{display:flex;gap:6px;align-items:center;margin:2px 0 20px;flex-wrap:wrap;}
.stp{display:flex;align-items:center;gap:8px;color:var(--muted);background:#2b3038;
  padding:6px 14px;border-radius:20px;font-size:13px;font-weight:600;}
.stp .num{width:22px;height:22px;border-radius:50%;background:#3a414d;display:grid;
  place-items:center;font-size:12px;}
.stp.active{color:var(--brand-text);background:var(--brand-soft);}
.stp.active .num{background:var(--brand);color:#1b1013;}
.stp.done{color:var(--green);background:var(--green-soft);}
.stp.done .num{background:var(--green);color:#1b1e13;}
.stp .arw{color:#556074;font-size:15px;}
/* Category bars */
.catbar{display:grid;grid-template-columns:190px 1fr 34px;gap:12px;align-items:center;
  margin:8px 0;font-size:13px;}
.catbar .track{height:9px;background:#2b3038;border-radius:6px;overflow:hidden;}
.catbar .fill{height:100%;background:var(--red);border-radius:6px;}
.catbar .fill.ok{background:var(--green);}
.catbar .cat-n{text-align:right;color:var(--muted);font-variant-numeric:tabular-nums;}
/* Flag cards inside review */
.flagcard{background:var(--red-soft);border:1px solid #4a3226;border-radius:8px;
  padding:9px 12px;margin:6px 0;font-size:13px;line-height:1.4;color:var(--ink);}
.flagcard .cat{color:var(--red);font-weight:700;}
/* Hide Streamlit chrome for a standalone-product look */
#MainMenu{visibility:hidden;}
[data-testid="stToolbar"]{display:none !important;}
[data-testid="stStatusWidget"]{display:none !important;}
[data-testid="stDecoration"]{display:none !important;}
[data-testid="manage-app-button"]{display:none !important;}
.stAppDeployButton{display:none !important;}
footer{visibility:hidden; height:0;}
[class*="viewerBadge"]{display:none !important;}
</style>
"""
st.markdown(THEME_CSS, unsafe_allow_html=True)


def render_stepper(step: int) -> None:
    labels = ["Upload", "Summary", "Review"]
    html = '<div class="stepper">'
    for i, lab in enumerate(labels):
        cls = "stp" + (" active" if i == step else "") + (" done" if i < step else "")
        num = "&#10003;" if i < step else str(i + 1)
        html += f'<div class="{cls}"><span class="num">{num}</span>{lab}</div>'
        if i < len(labels) - 1:
            html += '<span class="arw">&rsaquo;</span>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def password_ok() -> bool:
    """Ask for a password only if one is configured.

    Set it in .streamlit/secrets.toml locally, or in the host's Secrets panel:

        password = "your-password-here"

    With no password configured the app behaves exactly as before, so local use
    stays frictionless.
    """
    try:
        expected = st.secrets.get("password")
    except Exception:
        expected = None
    if not expected:
        return True
    if st.session_state.get("authenticated"):
        return True

    st.title("Coding Checker")
    st.caption("This tool contains answer keys. Please enter the lab password.")
    entered = st.text_input("Password", type="password")
    if entered:
        if entered == expected:
            st.session_state["authenticated"] = True
            st.rerun()
        st.error("Incorrect password.")
    return False


if not password_ok():
    st.stop()

grammar = load_grammar(GRAMMAR_PATH)
REFERENCE_DIR.mkdir(exist_ok=True)
index_reference_folder(REFERENCE_DIR, grammar)

CATEGORY_OF_KIND = {
    "missing_utterance": "Missing utterance",
    "extra_utterance": "Extra utterance",
    "transcript_mismatch": "Transcript",
    "usv_missing_note": "USV note",
    "boundary_unclear": "Boundary",
    "extra": "Extra code",
}

st.sidebar.title("Coding Checker")
page = st.sidebar.radio("Menu", ["New Check", "Key Library"], label_visibility="collapsed")
st.sidebar.markdown("---")
st.sidebar.caption("Developed for research and coder training in AAC lab.")


if page == "New Check":

    def goto(step: int) -> None:
        st.session_state["step"] = step
        st.rerun()

    step = st.session_state.get("step", 0)
    passports = load_reference_passports(REFERENCE_DIR)
    if not passports:
        st.title("New Check")
        st.warning("No reference keys loaded. Add them on the Key Library page.")
        st.stop()

    render_stepper(step)

    # ---------------------------------------------------------------- STEP 0
    if step == 0:
        st.title("Load a student's file")
        st.caption("Drop the Observer XT export. The answer key is matched automatically.")

        uploaded = st.file_uploader(
            "Upload the student's Observer export (.xlsx)", type=["xlsx"]
        )
        if uploaded is None:
            st.info("Upload a student file to begin.")
            st.stop()

        if st.session_state.get("uploaded_name") != uploaded.name:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded.getbuffer())
            st.session_state["uploaded_name"] = uploaded.name
            st.session_state["student_path"] = tmp.name
            st.session_state.pop("results", None)
            st.session_state["notes"] = {}

        student_path = Path(st.session_state["student_path"])
        student_df = pd.read_excel(student_path)

        first_utts, suggestions = suggest_references(student_df, grammar, REFERENCE_DIR, top_k=5)
        ordered = suggestions or passports
        labels = [
            f"{p.get('display_name', p.get('id'))} ({p.get('utterance_count')} utterances)"
            for p in ordered
        ]
        best = labels[0] if labels else "—"
        st.markdown(
            f"**{uploaded.name}** &nbsp;→&nbsp; best match "
            f'<span style="background:#e6f6ec;color:#1f9d55;border:1px solid #bfe3cd;'
            f'padding:3px 9px;border-radius:8px;font-weight:600;font-size:12px">{best}</span>',
            unsafe_allow_html=True,
        )
        choice = st.selectbox(
            "Reference key (change if the match is wrong)",
            options=list(range(len(ordered))),
            format_func=lambda i: labels[i],
        )
        selected = ordered[choice]

        with st.expander("First utterances in this file"):
            for u in first_utts:
                st.write(f"- {u}")

        if st.button("Run check  ›", type="primary"):
            key_df = read_excel_first_sheet_or_named(
                REFERENCE_DIR / selected["file_name"], sheet_name=selected.get("sheet_name")
            )
            issues, student_utts, key_utts, alignment = compare_files_with_alignment(
                student_df, key_df, grammar
            )
            st.session_state["results"] = (issues, student_utts, key_utts, alignment)
            st.session_state["notes"] = {}
            goto(1)
        st.stop()

    # ------------------------------------------------- shared for steps 1-3
    results = st.session_state.get("results")
    if not results:
        goto(0)
    issues, student_utts, key_utts, alignment = results

    scores_rows = build_scores_rows(student_utts, key_utts, alignment, issues, grammar)
    category_cols = [label for _, label in ce.SCORES_COLUMN_ORDER]
    total_cells = len(scores_rows) * len(category_cols)
    matched = sum(r[c] for r in scores_rows for c in category_cols)
    match_pct = matched / total_cells if total_cells else 0

    issues_by_utt = {}
    for i in issues:
        if i.utterance_id is not None:
            issues_by_utt.setdefault(i.utterance_id, []).append(i)
    missing_utts = [i for i in issues if i.kind == "missing_utterance"]
    student_name = Path(st.session_state.get("uploaded_name", "student")).stem

    # ---------------------------------------------------------------- STEP 1
    if step == 1:
        st.title(f"Results — {student_name}")
        st.caption("What went wrong, at a glance. Details are one click away in Review.")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Utterances", len(student_utts))
        c2.metric("With issues", len(issues_by_utt))
        c3.metric("Missing utterances", len(missing_utts))
        c4.metric("Agreement", f"{match_pct:.0%}", help="Coding cells that agree with the key")

        st.markdown("##### Errors by category")
        scores_df = pd.DataFrame(scores_rows)
        cat_errors = {c: int((scores_df[c] == 0).sum()) for c in category_cols} if scores_rows else {}
        max_err = max(cat_errors.values()) if cat_errors and max(cat_errors.values()) else 1
        bars = ""
        for cat, n in sorted(cat_errors.items(), key=lambda kv: kv[1], reverse=True):
            width = int(n / max_err * 100) if n else 100
            klass = "fill" if n else "fill ok"
            bars += (
                f'<div class="catbar"><span>{cat}</span>'
                f'<div class="track"><div class="{klass}" style="width:{width}%"></div></div>'
                f'<span class="cat-n">{n}</span></div>'
            )
        st.markdown(bars, unsafe_allow_html=True)

        st.markdown("")
        b1, b2 = st.columns([1, 1])
        if b1.button("Review flagged utterances  ›", type="primary"):
            goto(2)
        if b2.button("‹  Back to upload"):
            goto(0)
        st.stop()

    # ---------------------------------------------------------------- STEP 2
    if step == 2:
        st.title(f"Review — {student_name}")

        notes = st.session_state.setdefault("notes", {})
        flagged = [u for u in student_utts if issues_by_utt.get(u.uid)]
        noted = sum(1 for u in flagged if notes.get(u.uid, "").strip())

        top = st.columns([2, 1])
        with top[0]:
            st.caption(f"Noted {noted} of {len(flagged)} flagged utterances")
            st.progress(noted / len(flagged) if flagged else 0.0)
        with top[1]:
            only_issues = st.toggle("Only show disagreements", value=True)

        if missing_utts:
            with st.container(border=True):
                st.markdown("**Missing utterances** — in the key, not coded by the student")
                for i in missing_utts:
                    st.write(f'- "{i.expected or "no transcript"}"')

        for utt in student_utts:
            utt_issues = issues_by_utt.get(utt.uid, [])
            if only_issues and not utt_issues:
                continue

            light = "🟢" if not utt_issues else "🟤"
            badge = "matches key" if not utt_issues else f"{len(utt_issues)} issue(s)"
            text = utt.utterance_text or "no transcript"
            title = f'{light}  Utterance {utt.uid:02d} — "{text}"  ·  {badge}'
            with st.expander(title, expanded=bool(utt_issues) and len(flagged) <= 3):
                if utt_issues:
                    cards = ""
                    for i in utt_issues:
                        cat = CATEGORY_OF_KIND.get(i.kind)
                        if cat:
                            cards += (
                                f'<div class="flagcard"><span class="cat">{cat}</span> — {i.message}</div>'
                            )
                        else:
                            cards += f'<div class="flagcard">{i.message}</div>'
                    st.markdown(cards, unsafe_allow_html=True)
                else:
                    st.success("Matches the key.")

                notes[utt.uid] = st.text_area(
                    "Your note (added to the Excel note for this utterance)",
                    value=notes.get(utt.uid, ""),
                    key=f"note_{utt.uid}",
                    height=80,
                )

        st.markdown("---")

        # Build the annotated report on the fly so it always reflects the notes
        # typed above, then offer it as a direct download — no separate step.
        student_path = Path(st.session_state["student_path"])
        out_path = Path(tempfile.gettempdir()) / f"{student_name}_checked.xlsx"
        write_annotated_excel(
            student_path, out_path, issues, grammar,
            student_utts=student_utts, key_utts=key_utts, alignment=alignment,
            reviewer_notes={k: v for k, v in st.session_state.get("notes", {}).items() if v and v.strip()},
        )
        with open(out_path, "rb") as f:
            st.download_button(
                "⬇  Download report (with your notes)",
                data=f.read(),
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        st.caption(
            "Scores sheet plus your notes as Excel hover comments. "
            "Totals and percentages are live Excel formulas — edit a cell and they recalculate."
        )

        c1, c2 = st.columns([1, 1])
        if c1.button("Check next student  ›"):
            for k in ("results", "notes", "uploaded_name", "student_path"):
                st.session_state.pop(k, None)
            goto(0)
        if c2.button("‹  Back to summary"):
            goto(1)
        st.stop()


elif page == "Key Library":
    st.title("Key Library")

    passports = load_reference_passports(REFERENCE_DIR)
    st.write(f"**{len(passports)} keys loaded.**")
    for p in passports:
        with st.expander(f"{p.get('display_name')} — {p.get('utterance_count')} utterances"):
            for u in p.get("first_utterances", []):
                st.write(f"- {u}")

    st.markdown("---")
    st.subheader("Add a key")
    st.warning(
        "Permanent keys live in the reference_keys/ folder of the repository and are reloaded "
        "on every deploy. A key uploaded here is available for this session only — on a hosted "
        "server it is lost when the app restarts. To add one permanently, commit the .xlsx file "
        "to reference_keys/ in the repository."
    )

    display_name = st.text_input("Display name")
    uploaded_ref = st.file_uploader("Reference .xlsx", type=["xlsx"], key="ref_upload")
    if uploaded_ref is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_ref.getbuffer())
            tmp_path = Path(tmp.name)
        preview = create_reference_passport(tmp_path, grammar, display_name=display_name or None)
        st.caption(f"{preview['utterance_count']} utterances detected. First utterances:")
        for u in preview.get("first_utterances", []):
            st.write(f"- {u}")
        if st.button("Add for this session"):
            add_reference_to_library(
                tmp_path, REFERENCE_DIR, grammar,
                display_name=display_name or uploaded_ref.name.replace(".xlsx", ""),
            )
            st.success("Added.")
            st.rerun()
