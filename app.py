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
    passport_match_quality,
    read_excel_first_sheet_or_named,
    suggest_references,
    write_annotated_excel,
    write_student_feedback_excel,
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
/* The stepper chips themselves are styled in render_stepper(), which rebuilds
   their colours and number badges for the current step on every run. */
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


def render_stepper(step: int, unlocked: int = 0) -> None:
    """The step indicator, and a second way to navigate the wizard.

    Buttons rather than links: a link would navigate the browser, which starts a
    new Streamlit session and would ask for the password again. Buttons rerun in
    place, so the file, the results and the login all survive.

    To keep the original chip look the button carries only its label, and the CSS
    below - rebuilt on every run for the current step - adds the round number
    badge and the active/done/locked colours.
    """
    labels = ["Upload", "Summary", "Review", "Training"]
    css = [
        # st.columns guarantees the row; the chip and its arrow then share one
        # vertically-centred flex line inside the column, so they cannot drift apart.
        'div[class*="st-key-stepper_"]{display:flex;align-items:center;gap:6px;}',
        # Measured in the browser: a weighted column is wider than its chip, and the
        # leftover pushed the next chip ~20px further away. Letting these columns
        # shrink to their content puts the chips back at the original ~18px apart,
        # whatever the window width. Scoped by :has so other columns are untouched.
        '[data-testid="stColumn"]:has(div[class*="st-key-stepper_"]){'
        "flex:0 0 auto!important;width:auto!important;min-width:0!important;}",
        '[data-testid="stHorizontalBlock"]:has(div[class*="st-key-stepper_"]){gap:6px;}',
        # The chip itself, sized to its own label like the original inline chips.
        'div[class*="st-key-stepper_"] button{'
        "display:flex;align-items:center;justify-content:center;gap:8px;width:auto;"
        "color:var(--muted);background:#2b3038;border:none!important;"
        "padding:6px 14px;border-radius:20px;font-size:13px;font-weight:600;"
        "line-height:1.2;min-height:0;transition:.15s;white-space:nowrap;}",
        # The round number badge, drawn in front of the label.
        'div[class*="st-key-stepper_"] button::before{'
        "width:22px;height:22px;border-radius:50%;background:#3a414d;color:inherit;"
        "display:grid;place-items:center;font-size:12px;flex:none;}",
        'div[class*="st-key-stepper_"] button:disabled{opacity:.45;cursor:not-allowed;}',
        'div[class*="st-key-stepper_"] button:hover:not(:disabled){filter:brightness(1.18);}',
        # No focus ring: it left a stray outline around whichever chip was clicked.
        'div[class*="st-key-stepper_"] button:focus,'
        'div[class*="st-key-stepper_"] button:focus-visible,'
        'div[class*="st-key-stepper_"] button:active{'
        "outline:none!important;box-shadow:none!important;border:none!important;}",
        'div[class*="st-key-stepper_"] button p{margin:0;}',
        # The arrow rides in the chip's own flex line, so it cannot fall out of line.
        'div[class*="st-key-stepper_"]::after{'
        'content:"\\203A";color:#556074;font-size:20px;font-weight:600;line-height:1;}',
        'div[class*="st-key-stepper_3"]::after{content:none;}',
    ]
    for i in range(len(labels)):
        sel = f'div[class*="st-key-stepper_{i}"] button'
        if i < step:  # completed
            css.append(f'{sel}::before{{content:"\\2713";background:var(--green);color:#1b1e13;}}')
            css.append(f"{sel}{{color:var(--green);background:var(--green-soft);}}")
        elif i == step:  # current
            css.append(f'{sel}::before{{content:"{i + 1}";background:var(--brand);color:#1b1013;}}')
            css.append(f"{sel}{{color:var(--brand-text);background:var(--brand-soft);}}")
        else:
            css.append(f'{sel}::before{{content:"{i + 1}";}}')
    st.markdown("<style>" + "".join(css) + "</style>", unsafe_allow_html=True)

    # One column per chip, sized to its label, plus a spacer so they stay left.
    widths = [len(lab) + 3.0 for lab in labels] + [26.0]
    cols = st.columns(widths, vertical_alignment="center")
    for i, lab in enumerate(labels):
        with cols[i]:
            if st.button(lab, key=f"stepper_{i}", disabled=i > unlocked):
                st.session_state["step"] = i
                st.rerun()
    st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)


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
# For the "which categories are coded when Listing is present?" item. The student
# picks the correct subset; the rest are distractors from the normal coding set.
LISTING_MOD_OPTIONS = [
    "Number of Symbols", "Number of Relevant Symbols", "Parts of Speech",
    "Word Order", "SV", "Inflectional Morphemes", "Grammatical Intent",
]
# Student upload gate: the best-matching key must reach this share of the maximum
# possible first-utterance match score, or the file is refused and nothing opens.
# Keys are a serious research instrument, so a non-matching file reveals nothing.
MATCH_THRESHOLD = 0.5
# How a category is named to a student. A message can start with either the
# grammar's group name ("Imitative", "Independent Aided Utterances" — used by the
# "not coded" template) or the friendlier wording feedback.yaml uses in its own
# sentences ("Imitativeness", "Independence"), so both forms are listed and map to
# one display label. Longest first, so a longer name never loses to a shorter one.
CATEGORY_ALIASES = {
    "Listing": "Listing",
    "Communicative Intent": "Communicative Intent",
    "Imitative": "Imitativeness",
    "Imitativeness": "Imitativeness",
    "Independent Aided Utterances": "Independence",
    "Independence": "Independence",
    "Number of Symbols": "Number of Symbols",
    "Number of Relevant Symbols": "Number of Relevant Symbols",
    "Word Order": "Word Order",
    "SV": "SV",
    "Parts of Speech": "Parts of Speech",
    "Inflectional Morphemes": "Inflectional Morphemes",
    "Grammatical Intent": "Grammatical Intent",
}
CATEGORY_NEEDLES = sorted(CATEGORY_ALIASES, key=len, reverse=True)
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
    if item.get("mode") == "modifiers":
        want = {x.strip() for x in ans}
        got = set(given)
        return got == want, f'Key: {", ".join(ans)}'
    if cat == "SV":
        lines = []
        all_ok = True
        for col, correct_v in ans.items():
            g = str(given.get(col, "")).strip()
            cv = str(correct_v).strip()
            if col == "USV":
                # Accept the coding-sheet form the student is practising, e.g.
                # "USV: DOG EAT" or just "DOG EAT"; "NO"/"NONE" means no USV.
                g_norm = _re.sub(r"^\s*usv\s*:\s*", "", g, flags=_re.I).strip()
                if cv.upper() == "NO":
                    ok = g_norm.upper() in ("NO", "NONE")
                else:
                    ok = g_norm.upper() == cv.upper()
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
        incomplete = False
        if item.get("mode") == "modifiers":
            given = st.multiselect(
                "Listing is present here. Which categories do you code? Select all that apply.",
                LISTING_MOD_OPTIONS, key=key + "_mods",
            )
            incomplete = not given
        elif cat == "SV":
            sv_cols = load_training_items().get("sv_columns", [])
            given = {}
            cols = st.columns(2)
            for i, colname in enumerate(sv_cols):
                with cols[i % 2]:
                    if colname == "USV":
                        # Mandatory: the student writes the USV note the way they
                        # would in Observer's Comment column — this is the skill.
                        given[colname] = st.text_input(
                            'USV — write it out as "USV: SUBJECT VERB" (or NO if none)',
                            key=key + "_USV", placeholder="USV: DOG EAT",
                        )
                    else:
                        # index=None so nothing is pre-selected — the student must
                        # decide each YES/NO rather than accept a default.
                        given[colname] = st.radio(
                            colname, ["YES", "NO"], key=key + "_" + colname,
                            horizontal=True, index=None,
                        )
            missing_radio = any(given.get(c) is None for c in sv_cols if c != "USV")
            missing_usv = not str(given.get("USV", "")).strip()
            incomplete = missing_radio or missing_usv
        elif cat == "Parts of Speech":
            given = st.multiselect("Parts of speech present", POS_OPTIONS, key=key + "_ms")
        else:
            given = st.radio("Your answer", ITEM_OPTIONS[cat], key=key + "_r", index=None)
            incomplete = given is None

        if st.button("Check answer", key=key + "_btn"):
            if incomplete:
                st.warning(
                    "Answer every field before checking — choose YES or NO for each row "
                    "and write out the USV."
                )
            else:
                st.session_state[key + "_done"] = True

        if st.session_state.get(key + "_done"):
            correct, detail = _grade_item(item, given)
            if correct:
                st.success("Correct.")
            else:
                st.error("Not quite — compare with the key below.")
            if detail:
                st.markdown(detail)
            st.info(f'**Why:** {item["rationale"]}')


def _flag_html(issues_list):
    cards = ""
    for i in issues_list:
        cat = CATEGORY_OF_KIND.get(i.kind)
        prefix = f'<span class="cat">{cat}</span> — ' if cat else ""
        cards += f'<div class="flagcard">{prefix}{i.message}</div>'
    return cards


def issue_category(issue):
    """A safe category label for one issue — never exposes the key's value.

    Coding-category issues carry their category as the message prefix before
    ' - ' (e.g. "Word Order - key: ..."); non-coding issues are mapped by kind.
    Students see only this label, never the rest of the message.
    """
    cat = CATEGORY_OF_KIND.get(issue.kind)
    if cat:
        return cat
    msg = issue.message or ""
    # Both "<Category> - key: X, you coded: Y" and "<Category> is not coded for
    # this utterance. Key: X" start with the category, so one prefix test covers
    # every situation.
    for needle in CATEGORY_NEEDLES:
        if msg.startswith(needle):
            return CATEGORY_ALIASES[needle]
    # Never fall back to the raw message: it can contain the key's value.
    return "Coding"


def student_hint(grammar, category_label):
    """The answer-free operational-definition hint for a category, or ''."""
    hints = (grammar.get("feedback", {}) or {}).get("student_hints", {}) or {}
    return str(hints.get(category_label, "") or "")


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
    # Students run only the wizard (Upload -> Summary -> Review -> Training) and have
    # a single page, so the sidebar shows no menu.
    page = "New Check"

st.sidebar.markdown("---")
# Switch role without reloading the page: drop the role and every file/result from
# this session, then rerun — password_ok() asks for the password again.
if st.sidebar.button(f"Switch role  ({ROLE})", use_container_width=True):
    for k in list(st.session_state.keys()):
        st.session_state.pop(k, None)
    st.rerun()
st.sidebar.caption("Developed for research and coder training in AAC lab.")
st.sidebar.caption("by Anastasiia Volkova")


if page == "New Check":

    is_student = ROLE == "student"

    def goto(step: int) -> None:
        st.session_state["step"] = step
        st.rerun()

    def reset_flow() -> None:
        for k in ("results", "notes", "uploaded_name", "student_path",
                  "uploaded_bytes", "selected_ref"):
            st.session_state.pop(k, None)
        goto(0)

    step = st.session_state.get("step", 0)
    passports = load_reference_passports(REFERENCE_DIR)
    if not passports:
        st.title("Practice" if is_student else "New Check")
        st.warning(
            "No practice videos are loaded yet. Check with your instructor."
            if is_student else
            "No reference keys loaded. Add them on the Key Library page."
        )
        st.stop()

    # Summary/Review/Training only become reachable once a file has been checked.
    unlocked = 3 if st.session_state.get("results") else 0
    render_stepper(step, unlocked)
    themes = load_training_items().get("video_themes", {})

    # ---------------------------------------------------------------- STEP 0
    if step == 0:
        if is_student:
            st.title("Upload your file")
            st.caption(
                "Drop your Observer XT export. It is checked against the answer key so you "
                "can see where to improve — you never see the key itself."
            )
        else:
            st.title("Load a student's file")
            st.caption("Drop the Observer XT export. The answer key is matched automatically.")

        uploaded = st.file_uploader("Upload the Observer export (.xlsx)", type=["xlsx"])
        if uploaded is None:
            st.info("Upload a file to begin.")
            st.stop()

        if st.session_state.get("uploaded_name") != uploaded.name:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded.getbuffer())
            st.session_state["uploaded_name"] = uploaded.name
            st.session_state["student_path"] = tmp.name
            # Keep the exact original bytes so a student can download their file back
            # untouched — every Observer XT column intact, reopens in Observer.
            st.session_state["uploaded_bytes"] = uploaded.getvalue()
            st.session_state.pop("results", None)
            st.session_state["notes"] = {}

        student_path = Path(st.session_state["student_path"])
        student_df = pd.read_excel(student_path)

        first_utts, suggestions = suggest_references(student_df, grammar, REFERENCE_DIR, top_k=5)
        ordered = suggestions or passports

        # Match-quality gate, order-independent: a coder who missed an utterance
        # must still be recognised, since that is exactly what we are checking for.
        best_ref = ordered[0]
        quality = passport_match_quality(first_utts, best_ref)

        def run_check(selected):
            key_df = read_excel_first_sheet_or_named(
                REFERENCE_DIR / selected["file_name"], sheet_name=selected.get("sheet_name")
            )
            issues, student_utts, key_utts, alignment = compare_files_with_alignment(
                student_df, key_df, grammar
            )
            st.session_state["results"] = (issues, student_utts, key_utts, alignment)
            st.session_state["selected_ref"] = selected
            st.session_state["notes"] = {}
            goto(1)

        if is_student:
            # A file that matches no key opens nothing: the keys are a serious
            # research instrument and must not be handed to arbitrary uploads.
            if quality < MATCH_THRESHOLD:
                st.error(
                    "This file doesn't match any of the training videos closely enough to check. "
                    "Make sure you uploaded the right Observer export for your assigned video, "
                    "then try again."
                )
                st.stop()
            st.success("File recognized. Ready to check.")
            if st.button("Run check  ›", type="primary"):
                run_check(best_ref)
            st.stop()

        # Teacher: full visibility, may override the matched key.
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
            run_check(selected)
        st.stop()

    # ------------------------------------------------- shared for steps 1-3
    results = st.session_state.get("results")
    if not results:
        goto(0)
    issues, student_utts, key_utts, alignment = results
    selected_ref = st.session_state.get("selected_ref", {})
    _m = _re.search(r"(\d+)", selected_ref.get("file_name", "") or "")
    vnum = int(_m.group(1)) if _m else None
    video_theme = themes.get(str(vnum), "")

    scores_rows = build_scores_rows(student_utts, key_utts, alignment, issues, grammar)
    category_cols = [label for _, label in ce.SCORES_COLUMN_ORDER]
    total_cells = len(scores_rows) * len(category_cols)
    matched = sum(r[c] for r in scores_rows for c in category_cols)
    match_pct = matched / total_cells if total_cells else 0
    scores_df = pd.DataFrame(scores_rows)

    issues_by_utt = {}
    for i in issues:
        if i.utterance_id is not None:
            issues_by_utt.setdefault(i.utterance_id, []).append(i)
    missing_utts = [i for i in issues if i.kind == "missing_utterance"]
    student_name = Path(st.session_state.get("uploaded_name", "student")).stem

    # Errors per trainable category (bank name -> count), for the Training step.
    if scores_rows:
        cat_errors_bank = {b: int((scores_df[label] == 0).sum()) for b, label in TRAINABLE.items()}
    else:
        cat_errors_bank = {b: 0 for b in TRAINABLE}
    ranked_banks = [b for b, _ in sorted(cat_errors_bank.items(), key=lambda kv: kv[1], reverse=True)
                    if cat_errors_bank[b] > 0][:3]

    # ---------------------------------------------------------------- STEP 1
    if step == 1:
        title = "Your results" if is_student else f"Results — {student_name}"
        st.title(title)
        if is_student and video_theme:
            st.caption(f"Video: {video_theme}. A quick look at how your coding compares with the key.")
        else:
            st.caption("What went wrong, at a glance. Details are one click away in Review.")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Utterances", len(student_utts))
        c2.metric("With issues", len(issues_by_utt))
        c3.metric("Missing utterances", len(missing_utts))
        c4.metric("Agreement", f"{match_pct:.0%}", help="Coding cells that agree with the key")

        st.markdown("##### Errors by category")
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
        review_label = "See what to fix  ›" if is_student else "Review flagged utterances  ›"
        if b1.button(review_label, type="primary"):
            goto(2)
        if b2.button("‹  Start over" if is_student else "‹  Back to upload"):
            reset_flow() if is_student else goto(0)
        st.stop()

    # ---------------------------------------------------------------- STEP 2
    if step == 2:
        # -------------------------------------------------- STUDENT REVIEW
        if is_student:
            st.title("What to fix")
            st.caption(
                "These are the utterances where your coding differs from the key. You're shown "
                "which category to re-check — not the answer. Rewatch, rethink, then fix it in "
                "Observer and run the check again."
            )
            flagged = [u for u in student_utts if issues_by_utt.get(u.uid)]

            def _chips(pairs):
                # pairs: list of (label, is_missing)
                out = ""
                for label, miss in pairs:
                    tail = " — missing" if miss else ""
                    out += (
                        '<span class="flagcard" style="display:inline-block;margin:3px 6px 3px 0;">'
                        f'<span class="cat">{label}</span>{tail}</span>'
                    )
                return out

            if not flagged and not missing_utts:
                st.success("No differences from the key. Excellent work!")
            else:
                total = len(flagged) + len(missing_utts)
                st.markdown(f"**{total} item(s)** need another look.")

                # Flagged and missing utterances are shown in one list, ordered by
                # where they sit in the file — a missing utterance appears between
                # the utterances it belongs between, exactly as in the marked-up
                # file, rather than being appended at the end away from its context.
                ordered = [(u.anchor_row_index, "utt", u) for u in flagged]
                ordered += [
                    ((i.insert_after_row_index or -1) + 0.5, "missing", i) for i in missing_utts
                ]
                ordered.sort(key=lambda x: x[0])

                for _, kind, obj in ordered:
                    if kind == "missing":
                        text = obj.expected or "no transcript"
                        with st.container(border=True):
                            st.markdown(f'**Utterance not coded** — "{text}"')
                            st.markdown(
                                "Re-check: " + _chips([("Missing utterance", True)])
                                + "  <span style='color:var(--muted)'>please code this "
                                "utterance here</span>",
                                unsafe_allow_html=True,
                            )
                        continue
                    utt = obj
                    ui = issues_by_utt[utt.uid]
                    wrong_cats = sorted({issue_category(i) for i in ui if i.kind != "missing"})
                    miss_cats = sorted({issue_category(i) for i in ui if i.kind == "missing"})
                    text = utt.utterance_text or "no transcript"
                    with st.container(border=True):
                        st.markdown(f'**Utterance {utt.uid:02d}** — "{text}"')
                        pairs = [(c, False) for c in wrong_cats] + [(c, True) for c in miss_cats]
                        st.markdown("Re-check: " + _chips(pairs), unsafe_allow_html=True)

                # How this is coded — rule for each category with a genuine coding
                # difference (not the merely-missing ones), shown once, no answers.
                rule_cats = sorted({
                    issue_category(i) for i in issues
                    if i.kind not in ("missing", "missing_utterance")
                })
                rule_cats = [c for c in rule_cats if student_hint(grammar, c)]
                if rule_cats:
                    st.markdown("##### How these categories are coded")
                    st.caption("The rule for each category you need to re-check — not the answer.")
                    for c in rule_cats:
                        st.markdown(f"**{c}** — {student_hint(grammar, c)}")

            st.markdown("---")
            st.markdown("##### Your marked-up file")
            st.caption(
                "One file: your coding with the utterances to re-check shaded and a hover note on "
                "each. Categories you didn't code are marked as missing, and any utterance you "
                "missed is inserted with its transcript and a note to code it — no answers. "
                "The Scores sheet shows your totals and percentage per category."
            )

            # File inputs, all answer-free for coded values:
            #  - wrong_rows: rows with a wrong coded value -> rule comment
            #  - missing_cat_rows: a category not coded at all -> "missing" note only
            #  - missing_utts_arg: whole utterances to insert with their transcript
            utt_by_uid = {u.uid: u for u in student_utts}
            wrong_rows, missing_cat_rows = {}, {}
            for i in issues:
                if i.kind == "missing_utterance":
                    continue
                if i.kind == "missing":
                    u = utt_by_uid.get(i.utterance_id)
                    if u is not None:
                        missing_cat_rows.setdefault(u.anchor_row_index, set()).add(issue_category(i))
                elif i.row_index is not None:
                    wrong_rows.setdefault(i.row_index, set()).add(issue_category(i))
            wrong_rows = {r: sorted(c) for r, c in wrong_rows.items()}
            missing_cat_rows = {r: sorted(c) for r, c in missing_cat_rows.items()}
            missing_utts_arg = [
                (i.insert_after_row_index, i.expected or "") for i in missing_utts
            ]

            base = Path(st.session_state.get("uploaded_name", "my_file.xlsx")).stem
            fb_path = Path(tempfile.gettempdir()) / f"{base}_feedback.xlsx"
            try:
                write_student_feedback_excel(
                    Path(st.session_state["student_path"]), fb_path, grammar,
                    wrong_rows, missing_cat_rows, missing_utts_arg,
                    scores_rows=scores_rows,
                )
                fb_bytes = fb_path.read_bytes()
            except Exception:
                fb_bytes = None

            if fb_bytes is not None:
                st.download_button(
                    "⬇  Download my marked-up file",
                    data=fb_bytes,
                    file_name=fb_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            else:
                st.caption("_(Marked-up file unavailable for this file.)_")

            st.markdown("")
            c1, c2, c3 = st.columns(3)
            if c1.button("Practice these areas  ›", type="primary"):
                goto(3)
            if c2.button("‹  Back to results"):
                goto(1)
            if c3.button("Start over"):
                reset_flow()
            st.stop()

        # -------------------------------------------------- TEACHER REVIEW
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

        c1, c2, c3 = st.columns(3)
        if c1.button("Practice weak areas  ›", type="primary"):
            goto(3)
        if c2.button("Check next student"):
            reset_flow()
        if c3.button("‹  Back to summary"):
            goto(1)
        st.stop()

    # ---------------------------------------------------------------- STEP 3
    if step == 3:
        ti = load_training_items()
        st.title("Training")
        if video_theme:
            st.caption(f"Targeted practice for {video_theme}, focused on your weakest categories.")
        else:
            st.caption("Targeted practice focused on the weakest categories in this file.")

        if not ranked_banks:
            st.success("No errors in the trainable categories — nothing to practice. Great work!")
        else:
            tabs = st.tabs([f"{b} ({cat_errors_bank[b]})" for b in ranked_banks])
            for tab, bank in zip(tabs, ranked_banks):
                with tab:
                    items = [x for x in ti["items"] if x["video"] == vnum and x["category"] == bank]
                    if not items:
                        st.info("No practice items for this category and video yet.")
                    for i, item in enumerate(items, 1):
                        render_training_item(item, i)

        st.markdown("---")
        c1, c2 = st.columns(2)
        if c1.button("‹  Back to review"):
            goto(2)
        if c2.button("Start over" if is_student else "Check next student"):
            reset_flow()
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
    # Teacher-only free practice. Students never reach this page — their practice
    # is the Training step of the wizard, scoped to their own video and errors.
    ti = load_training_items()
    themes = ti.get("video_themes", {})
    st.title("Training")
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
        elif isinstance(ans, list):
            ans = ", ".join(ans)
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
