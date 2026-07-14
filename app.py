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
TRAINING_PATH = ROOT / "training_items.yaml"

st.set_page_config(page_title="Coding Checker", layout="wide", initial_sidebar_state="expanded")

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
/* Hide Streamlit chrome for a standalone-product look.
   NOTE: we hide only specific items (menu, deploy, badge, status) and NOT the
   whole toolbar/header, so the sidebar open/close control keeps working. */
#MainMenu{visibility:hidden;}
[data-testid="stStatusWidget"]{display:none !important;}
[data-testid="stDecoration"]{display:none !important;}
[data-testid="manage-app-button"]{display:none !important;}
.stAppDeployButton{display:none !important;}
footer{visibility:hidden; height:0;}
[class*="viewerBadge"]{display:none !important;}
/* Keep the sidebar open/close control always visible and clickable. */
[data-testid="stSidebarCollapsedControl"],
[data-testid="stSidebarCollapseButton"],
[data-testid="collapsedControl"]{
  display:flex !important; visibility:visible !important; opacity:1 !important;
}
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


# ---------------------------------------------------------------------------
# Training mode
# ---------------------------------------------------------------------------
import re as _re
import yaml as _yaml


@st.cache_data(show_spinner=False)
def load_training_items():
    if not TRAINING_PATH.exists():
        return {"video_themes": {}, "sv_columns": [], "items": []}
    with open(TRAINING_PATH, encoding="utf-8") as f:
        return _yaml.safe_load(f) or {}


# Item-bank category -> the Scores-sheet label used to count that category's errors.
TRAINABLE = {
    "Listing": "Listing",
    "Word Order": "Word order",
    "SV": "SV",
    "Parts of Speech": "Parts of speech",
    "Inflectional Morphemes": "Inflectional morph",
    "Grammatical Intent": "Grammatical Intent",
}
POS_OPTIONS = ["Noun", "Pronoun", "Verb", "Preposition", "Adjective", "Determiner", "Conjunction", "Absent"]
ITEM_OPTIONS = {
    "Listing": ["Listing Present", "Listing Not Present"],
    "Word Order": ["1.0", "0.5", "0"],
    "Grammatical Intent": ["Grammatical Intent Clear", "Grammatical Intent NOT Clear"],
    "Inflectional Morphemes": ["At least one inflectional morpheme", "No appropriate inflectional morphemes"],
}


def _grade_item(item, given):
    """Return (all_correct, detail_markdown)."""
    cat = item["category"]
    ans = item["answer"]
    if cat == "SV":
        lines = []
        all_ok = True
        for col, correct_v in ans.items():
            g = str(given.get(col, "")).strip()
            cv = str(correct_v).strip()
            if col == "USV":
                ok = g.upper() == cv.upper() or (g.upper() in ("", "NO") and cv.upper() == "NO")
            else:
                ok = g.upper() == cv.upper()
            all_ok = all_ok and ok
            mark = "✓" if ok else "✗"
            lines.append(f'- {col}: you coded "{g or "-"}" · key "{cv}"  {mark}')
        return all_ok, "\n".join(lines)
    if cat == "Parts of Speech":
        want = {"Absent"} if ans == "Absent" else {x.strip() for x in ans.split(",")}
        got = set(given)
        return got == want, f'Key: {", ".join(sorted(want))}'
    return str(given).strip() == str(ans).strip(), f"Key: {ans}"


def render_training_item(item, seq):
    """One interactive practice item: input widgets, a Check button, then feedback."""
    cat = item["category"]
    key = f'ti_{item["video"]}_{cat}_{item["n"]}'
    with st.container(border=True):
        st.markdown(f'**{seq}.**  "{item["utterance"]}"')
        if cat == "SV":
            sv_cols = load_training_items().get("sv_columns", [])
            given = {}
            cols = st.columns(2)
            for i, colname in enumerate(sv_cols):
                with cols[i % 2]:
                    if colname == "USV":
                        given[colname] = st.text_input(
                            "USV (type the subject+verb, or NO)", key=key + "_USV"
                        )
                    else:
                        given[colname] = st.radio(
                            colname, ["YES", "NO"], key=key + "_" + colname, horizontal=True
                        )
        elif cat == "Parts of Speech":
            given = st.multiselect("Parts of speech present", POS_OPTIONS, key=key + "_ms")
        else:
            given = st.radio("Your answer", ITEM_OPTIONS[cat], key=key + "_r", index=None)

        if st.button("Check answer", key=key + "_btn"):
            st.session_state[key + "_done"] = True

        if st.session_state.get(key + "_done"):
            correct, detail = _grade_item(item, given)
            if correct:
                st.success("Correct.")
            else:
                st.error("Not quite — compare with the key below.")
            if detail:
                st.markdown(detail)
            st.info(f'**Why:** {item["rationale"]}\n\n_Operational Definitions: {item["ref"]}_')


def _flag_html(issues_list):
    cards = ""
    for i in issues_list:
        cat = CATEGORY_OF_KIND.get(i.kind)
        prefix = f'<span class="cat">{cat}</span> — ' if cat else ""
        cards += f'<div class="flagcard">{prefix}{i.message}</div>'
    return cards


def password_ok() -> bool:
    """Two-level access. Passwords come only from st.secrets, never hard-coded:

        teacher_password = "..."   (full access; falls back to legacy 'password')
        student_password = "..."   (Training page only)

    On success the role is stored in st.session_state['role'] = 'teacher' | 'student'.
    With no passwords configured the app opens as teacher, so local use stays frictionless.
    """
    try:
        teacher_pw = st.secrets.get("teacher_password") or st.secrets.get("password")
        student_pw = st.secrets.get("student_password")
    except Exception:
        teacher_pw = student_pw = None

    if not teacher_pw and not student_pw:
        st.session_state["role"] = "teacher"
        return True
    if st.session_state.get("role"):
        return True

    st.title("Coding Checker")
    st.caption("Please enter your password.")
    entered = st.text_input("Password", type="password")
    if entered:
        if teacher_pw and entered == teacher_pw:
            st.session_state["role"] = "teacher"
            st.rerun()
        elif student_pw and entered == student_pw:
            st.session_state["role"] = "student"
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


if not password_ok():
    st.stop()

ROLE = st.session_state.get("role", "teacher")

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
if ROLE == "teacher":
    page = st.sidebar.radio(
        "Menu", ["New Check", "Key Library", "Training", "Item Bank"], label_visibility="collapsed"
    )
else:
    # Students only ever see the Training page. No keys, no key library, no item bank.
    page = "Training"
    st.sidebar.markdown("### Training")
    st.sidebar.caption("Practice mode")
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

            badge = "matches key" if not utt_issues else f"{len(utt_issues)} issue(s)"
            text = utt.utterance_text or "no transcript"
            title = f'Utterance {utt.uid:02d} — "{text}"  ·  {badge}'
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


elif page == "Training":
    ti = load_training_items()
    themes = ti.get("video_themes", {})
    st.title("Training")

    # ---------------------------------------------------------- TEACHER MODE
    if ROLE == "teacher":
        st.caption("Practice any video and category, independent of any student's errors.")
        vnum = st.selectbox(
            "Video", [1, 2, 3, 4],
            format_func=lambda v: f"Video {v} — {themes.get(str(v), '')}",
        )
        cat = st.selectbox("Category", list(TRAINABLE.keys()))
        items = [x for x in ti["items"] if x["video"] == vnum and x["category"] == cat]
        st.markdown(f"##### {cat} — {len(items)} items")
        for i, item in enumerate(items, 1):
            render_training_item(item, i)
        st.stop()

    # ---------------------------------------------------------- STUDENT MODE
    st.caption("Upload your Observer XT export. You'll see your scores and get practice on your weakest areas.")
    uploaded = st.file_uploader("Upload your Observer export (.xlsx)", type=["xlsx"])
    if uploaded is None:
        st.info("Upload your file to begin.")
        st.stop()

    if st.session_state.get("tr_name") != uploaded.name:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded.getbuffer())
        st.session_state["tr_name"] = uploaded.name
        st.session_state["tr_path"] = tmp.name

    student_df = pd.read_excel(st.session_state["tr_path"])
    passports = load_reference_passports(REFERENCE_DIR)
    _first, suggestions = suggest_references(student_df, grammar, REFERENCE_DIR, top_k=5)
    ordered = suggestions or passports
    if not ordered:
        st.error("Could not match your file to a training video. Check with your instructor.")
        st.stop()

    selected = ordered[0]
    m = _re.search(r"(\d+)", selected.get("file_name", ""))
    vnum = int(m.group(1)) if m else None
    st.markdown(f"**Video:** {themes.get(str(vnum), 'unknown')}")

    # Hooks into the existing engine: same comparison, scoring and error counting
    # used by the teacher's New Check — students just never see the key itself.
    key_df = read_excel_first_sheet_or_named(
        REFERENCE_DIR / selected["file_name"], sheet_name=selected.get("sheet_name")
    )
    issues, su, ku, al = compare_files_with_alignment(student_df, key_df, grammar)
    scores_rows = build_scores_rows(su, ku, al, issues, grammar)
    category_cols = [label for _, label in ce.SCORES_COLUMN_ORDER]
    scores_df = pd.DataFrame(scores_rows)
    total_cells = len(scores_rows) * len(category_cols)
    matched = int(scores_df[category_cols].values.sum()) if scores_rows else 0
    match_pct = matched / total_cells if total_cells else 0

    issues_by_utt = {}
    for i in issues:
        if i.utterance_id is not None:
            issues_by_utt.setdefault(i.utterance_id, []).append(i)

    c1, c2, c3 = st.columns(3)
    c1.metric("Utterances", len(su))
    c2.metric("With issues", len(issues_by_utt))
    c3.metric("Agreement", f"{match_pct:.0%}")

    cat_errors = (
        {bank: int((scores_df[label] == 0).sum()) for bank, label in TRAINABLE.items()}
        if scores_rows else {b: 0 for b in TRAINABLE}
    )
    max_err = max(cat_errors.values()) if cat_errors and max(cat_errors.values()) else 1
    st.markdown("##### Your errors by category")
    bars = ""
    for bank in TRAINABLE:
        n = cat_errors.get(bank, 0)
        width = int(n / max_err * 100) if n else 100
        klass = "fill" if n else "fill ok"
        bars += (
            f'<div class="catbar"><span>{bank}</span>'
            f'<div class="track"><div class="{klass}" style="width:{width}%"></div></div>'
            f'<span class="cat-n">{n}</span></div>'
        )
    st.markdown(bars, unsafe_allow_html=True)

    with st.expander("Your flagged utterances and feedback"):
        any_flag = False
        for utt in su:
            ui = issues_by_utt.get(utt.uid, [])
            if not ui:
                continue
            any_flag = True
            st.markdown(f'**"{utt.utterance_text or "no transcript"}"**')
            st.markdown(_flag_html(ui), unsafe_allow_html=True)
        if not any_flag:
            st.write("No issues flagged.")

    st.markdown("##### Practice your top areas")
    ranked = [b for b, _ in sorted(cat_errors.items(), key=lambda kv: kv[1], reverse=True)
              if cat_errors[b] > 0][:3]
    if not ranked:
        st.success("No errors in the trainable categories — nothing to practice. Great work!")
        st.stop()

    tabs = st.tabs([f"{b} ({cat_errors[b]})" for b in ranked])
    for tab, bank in zip(tabs, ranked):
        with tab:
            items = [x for x in ti["items"] if x["video"] == vnum and x["category"] == bank]
            if not items:
                st.info("No practice items for this category and video yet.")
            for i, item in enumerate(items, 1):
                render_training_item(item, i)


elif page == "Item Bank":
    ti = load_training_items()
    themes = ti.get("video_themes", {})
    st.title("Item Bank")
    st.caption("All training items (teacher view). Students never see this page.")

    rows = []
    for x in ti["items"]:
        ans = x["answer"]
        if isinstance(ans, dict):
            ans = " | ".join(f"{k}={v}" for k, v in ans.items())
        rows.append({
            "Video": f'{x["video"]} — {themes.get(str(x["video"]), "")}',
            "Category": x["category"],
            "#": x["n"],
            "Utterance": x["utterance"],
            "Answer": ans,
            "Rationale": x["rationale"],
            "Ref": x["ref"],
        })
    df = pd.DataFrame(rows)

    fcol1, fcol2 = st.columns(2)
    fv = fcol1.multiselect("Filter by video", sorted(df["Video"].unique()))
    fc = fcol2.multiselect("Filter by category", list(TRAINABLE.keys()))
    view = df
    if fv:
        view = view[view["Video"].isin(fv)]
    if fc:
        view = view[view["Category"].isin(fc)]
    st.caption(f"{len(view)} of {len(df)} items")
    st.dataframe(view, use_container_width=True, hide_index=True)
