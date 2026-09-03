"""
Validate the checker against Naomi's manual review of Anastasia's practice work.

The *_AV.xlsx files are real coded sessions that a human reviewer went through
row by row, leaving a threaded comment on each row she considered wrong. That
gives us ground truth to measure the checker against:

  - a row Naomi commented on but the program did not flag  -> MISS
  - a row the program flagged but Naomi did not comment on -> possible FALSE POSITIVE
    (possible, not certain: Naomi may simply have missed it, or chosen not to
    comment on a minor point)
  - both flagged the same row -> AGREEMENT

Comparison is at the level of the UTTERANCE, not the row: Naomi often attaches
one comment to an utterance's first row while the actual wrong code sits a few
rows below, so a row-exact match would understate agreement badly.
"""
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pandas as pd

import checker_engine as ce

NS = {"tc": "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"}
ROOT = Path(__file__).parent
UPLOADS = Path("/mnt/user-data/uploads")
grammar = ce.load_grammar(ROOT / "grammar.yaml")

CASES = [
    ("Video 2", UPLOADS / "Coding_Training-PracticeVideo2_AV.xlsx", ROOT / "reference_keys/KEY_Video_2.xlsx"),
    ("Video 3", UPLOADS / "Coding_Training-PracticeVideo3_AV.xlsx", ROOT / "reference_keys/KEY_Video_3.xlsx"),
    ("Video 4", UPLOADS / "Coding_Training-PracticeVideo4_AV.xlsx", ROOT / "reference_keys/KEY_Video_4.xlsx"),
]


def manual_comment_rows(path):
    """Excel row number -> Naomi's comment text."""
    z = zipfile.ZipFile(path)
    out = {}
    for f in z.namelist():
        if "threadedComment" not in f:
            continue
        root = ET.fromstring(z.read(f))
        for c in root.findall("tc:threadedComment", NS):
            ref = c.get("ref")  # e.g. "A138"
            row = int("".join(ch for ch in ref if ch.isdigit()))
            txt = c.find("tc:text", NS)
            out[row] = (txt.text or "").strip() if txt is not None else ""
    return out


totals = {"agree": 0, "miss": 0, "extra": 0}

for label, student_path, key_path in CASES:
    student_df = pd.read_excel(student_path)
    student_df.columns = [str(c).strip() for c in student_df.columns]
    key_df = pd.read_excel(key_path)

    issues, s_utts, k_utts, pairs = ce.compare_files_with_alignment(student_df, key_df, grammar)

    # Map: dataframe row index -> utterance uid
    row_to_utt = {}
    for u in s_utts:
        for r in u.rows:
            row_to_utt[r.original_index] = u.uid
    utt_text = {u.uid: u.utterance_text for u in s_utts}

    # Program-flagged utterances
    prog_utts = set()
    for i in issues:
        if i.utterance_id is not None:
            prog_utts.add(i.utterance_id)
        elif i.insert_after_row_index is not None:
            uid = row_to_utt.get(i.insert_after_row_index)
            if uid:
                prog_utts.add(uid)

    # Naomi-flagged utterances (Excel row N -> df index N-2)
    manual = manual_comment_rows(student_path)
    manual_utts = {}
    for excel_row, text in manual.items():
        uid = row_to_utt.get(excel_row - 2)
        if uid is not None:
            manual_utts.setdefault(uid, []).append(text)

    agree_m = {uid for uid in manual_utts if any(abs(uid - p) <= 1 for p in prog_utts)}
    miss = sorted(set(manual_utts) - agree_m)
    extra = sorted(p for p in prog_utts if not any(abs(p - m) <= 1 for m in manual_utts))
    agree = sorted(agree_m)

    totals["agree"] += len(agree)
    totals["miss"] += len(miss)
    totals["extra"] += len(extra)

    print("=" * 78)
    print(f"{label}: {len(s_utts)} student utterances vs {len(k_utts)} in key")
    print(f"  Naomi flagged {len(manual_utts)} utterances | program flagged {len(prog_utts)}")
    print(f"  AGREE {len(agree)}  |  MISSED BY PROGRAM {len(miss)}  |  PROGRAM-ONLY {len(extra)}")
    print("=" * 78)

    if miss:
        print("\n  -- Naomi flagged, program did NOT (real misses):")
        for uid in miss:
            print(f'     utt #{uid} "{utt_text.get(uid, "")}"')
            for t in manual_utts[uid]:
                print(f"        Naomi: {t}")

    if extra:
        print("\n  -- Program flagged, Naomi did not (check if false positive):")
        for uid in extra:
            msgs = [i.message for i in issues if i.utterance_id == uid]
            if not msgs:
                msgs = [i.message for i in issues
                        if i.insert_after_row_index is not None
                        and row_to_utt.get(i.insert_after_row_index) == uid]
            print(f'     utt #{uid} "{utt_text.get(uid, "")}"')
            for m in msgs[:3]:
                print(f"        program: {m[:88]}")
    print()

print("=" * 78)
tot = totals["agree"] + totals["miss"]
print(f"OVERALL: program caught {totals['agree']}/{tot} of the utterances Naomi flagged "
      f"({totals['agree'] / tot:.0%})")
print(f"         program flagged {totals['extra']} utterances Naomi did not")
print("=" * 78)
