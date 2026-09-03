"""
Validation suite for the coding checker.

Two kinds of test:

1. SELF-CHECK: compare each key against itself. A perfect submission must
   produce zero issues. Any issue here is a false positive in the engine.

2. MUTATION TEST: take a key, inject one known error, and check the engine
   reports exactly that error and nothing else. This is the only way to prove
   the checker actually catches what it claims to catch - a checker that
   reports nothing would pass the self-check perfectly.

Run:  python3 test_engine.py
"""
import sys
from pathlib import Path

import pandas as pd

import checker_engine as ce

ROOT = Path(__file__).parent
KEYS = ROOT / "reference_keys"
grammar = ce.load_grammar(ROOT / "grammar.yaml")

EXPECTED_UTTERANCES = {1: 47, 2: 34, 3: 32, 4: 34}

passed, failed = 0, 0


def check(name, condition, detail=""):
    global passed, failed
    if condition:
        passed += 1
        print(f"  PASS  {name}")
    else:
        failed += 1
        print(f"  FAIL  {name}   {detail}")


def load_key(n):
    return pd.read_excel(KEYS / f"KEY_Video_{n}.xlsx")


print("=" * 70)
print("1. SEGMENTATION: does each key parse into the expected utterances?")
print("=" * 70)
for n, expected in EXPECTED_UTTERANCES.items():
    df = load_key(n)
    events = ce.dataframe_to_events(ce.prepare_dataframe(df, grammar), grammar)
    utts = ce.segment_utterances(events, grammar)
    rows_kept = sum(len(u.rows) for u in utts)
    no_text = [u.uid for u in utts if not u.utterance_text]
    unclear = [u.uid for u in utts if u.boundary_unclear]
    check(f"KEY_Video_{n}: {len(utts)} utterances", len(utts) == expected, f"expected {expected}")
    check(f"KEY_Video_{n}: no rows lost", rows_kept == len(events), f"{rows_kept}/{len(events)}")
    check(f"KEY_Video_{n}: every utterance has a transcript", not no_text, f"empty: {no_text}")
    check(f"KEY_Video_{n}: every utterance has a CI code", not unclear, f"missing CI: {unclear}")

print()
print("=" * 70)
print("2. SELF-CHECK: a key compared to itself must produce ZERO issues")
print("=" * 70)
for n in EXPECTED_UTTERANCES:
    df = load_key(n)
    issues, s, k, pairs = ce.compare_files_with_alignment(df, df, grammar)
    unmatched = [p for p in pairs if p[0] is None or p[1] is None]
    check(f"KEY_Video_{n}: 0 issues against itself", len(issues) == 0,
          f"got {len(issues)}: {[i.message[:60] for i in issues[:3]]}")
    check(f"KEY_Video_{n}: all utterances aligned 1:1", not unmatched, f"unmatched: {len(unmatched)}")

print()
print("=" * 70)
print("3. MUTATION TESTS: inject one known error, expect it to be caught")
print("=" * 70)

base = load_key(2)
mods = [c for c in base.columns if str(c).startswith("Modifier")]


def run(df):
    return ce.compare_files_with_alignment(df, load_key(2), grammar)[0]


# --- 3a. Wrong behavior: change an Independent Utterance to Co-constructed
df = base.copy()
idx = df.index[df["Behavior"].astype(str).str.strip() == "Independent Utterance"][0]
df.at[idx, "Behavior"] = "Co-constructed Utterance"
issues = run(df)
hits = [i for i in issues if "Independence" in i.message and "independent" in i.message]
check("wrong behavior (Independent -> Co-constructed) detected", len(hits) == 1,
      f"{len(hits)} hits, {len(issues)} issues total")
check("  ...and nothing else flagged", len(issues) == 1, f"{[i.message[:50] for i in issues]}")

# --- 3b. Wrong modifier value: change a Number of Symbols count
df = base.copy()
idx = df.index[df["Behavior"].astype(str).str.strip() == "Number of Symbols"][0]
old = df.at[idx, mods[0]]
df.at[idx, mods[0]] = "99"
issues = run(df)
hits = [i for i in issues if "Number of Symbols" in i.message]
check(f"wrong modifier (Number of Symbols {old} -> 99) detected", len(hits) >= 1,
      f"{len(hits)} hits")
check("  ...and nothing else flagged", len(issues) == len(hits), f"{len(issues)} total")

# --- 3c. Missing code row: delete a Word Order Score row
df = base.copy()
idx = df.index[df["Behavior"].astype(str).str.strip() == "Word Order Score"][0]
df = df.drop(index=idx).reset_index(drop=True)
issues = run(df)
hits = [i for i in issues if i.kind == "missing" and "Word Order" in (i.missing_behavior or "")]
check("missing code row (Word Order Score) detected", len(hits) == 1, f"{len(hits)} hits, {len(issues)} total")

# --- 3d. Whole utterance skipped: delete every row of utterance #5
df = base.copy()
events = ce.dataframe_to_events(ce.prepare_dataframe(df, grammar), grammar)
utts = ce.segment_utterances(events, grammar)
victim = utts[4]
victim_text = victim.utterance_text
df = df.drop(index=[r.original_index for r in victim.rows]).reset_index(drop=True)
issues = run(df)
hits = [i for i in issues if i.kind == "missing_utterance"]
check(f"skipped utterance ({victim_text!r}) detected", len(hits) == 1, f"{len(hits)} hits")
check("  ...no cascade of false errors after it", len(issues) <= 2,
      f"{len(issues)} issues: {[i.message[:45] for i in issues]}")

# --- 3e. Extra utterance: duplicate an utterance block
df = base.copy()
events = ce.dataframe_to_events(ce.prepare_dataframe(df, grammar), grammar)
utts = ce.segment_utterances(events, grammar)
dup = utts[9]
dup_rows = df.loc[[r.original_index for r in dup.rows]]
df = pd.concat([df.loc[:dup.last_row_index], dup_rows, df.loc[dup.last_row_index + 1:]]).reset_index(drop=True)
issues = run(df)
hits = [i for i in issues if i.kind == "extra_utterance"]
check("extra/duplicated utterance detected", len(hits) == 1, f"{len(hits)} hits, {len(issues)} total")

# --- 3f. Missing CI code (the KEY_Video_1 bug): delete a CI row
df = base.copy()
idx = df.index[df["Behavior"].astype(str).str.strip() == "Communicative Intent Present"][3]
df = df.drop(index=idx).reset_index(drop=True)
issues = run(df)
hits = [i for i in issues if "Communicative Intent" in (i.missing_behavior or "") or "Communicative Intent" in i.message]
check("missing Communicative Intent row detected", len(hits) >= 1, f"{len(hits)} hits, {len(issues)} total")
check("  ...utterance still found (not swallowed)",
      len([i for i in issues if i.kind == "missing_utterance"]) == 0,
      "utterance was lost by the segmenter")

# --- 3g. Duplicate Observer start/stop markers must NOT be treated as errors
df = base.copy()
start_idx = df.index[df["Behavior"].astype(str).str.strip().str.lower() == "utterance start"][2]
extra = df.loc[[start_idx]]
df = pd.concat([df.loc[:start_idx], extra, df.loc[start_idx + 1:]]).reset_index(drop=True)
issues = run(df)
check("duplicate 'utterance start' marker ignored (Observer bug, not a student error)",
      len(issues) == 0, f"{len(issues)} false issues: {[i.message[:50] for i in issues]}")

print()
print("=" * 70)
print(f"RESULT: {passed} passed, {failed} failed")
print("=" * 70)
sys.exit(1 if failed else 0)
