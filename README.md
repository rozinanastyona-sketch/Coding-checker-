# Coding Checker MVP

This is a first MVP for checking student Observer coding files against reference keys.

## What it currently does

- Reads a student Excel file.
- Reads a selected reference key from the local reference library.
- Detects utterances using `Utterance start`, `Utterance stop`, and `Communicative Intent` rows.
- Treats duplicated start/stop rows as Observer app noise, not student error.
- Treats all `Modifier*` columns as dynamic: the column number does not matter.
- Compares coding structure inside each utterance against the key.
- Inserts red `MISSING` rows into the output Excel where a key code was expected but missing.
- Highlights wrong/extra rows red.
- Highlights missing `USV:` comments orange when `SV + Lexical Verb + Unique Subject-Verb Combination (USV)` is present.
- Highlights unclear utterance boundaries purple.
- Keeps student row order. It does not reorder existing rows.

## Project structure

```text
coding_checker_mvp/
  app.py
  checker_engine.py
  grammar.yaml
  requirements.txt
  reference_keys/
```

## Reference key library

`reference_keys/` holds one `.xlsx` per key plus a `.json` passport with the same
stem. The passport is what the app lists; `index_reference_folder()` writes one
automatically at startup for any `.xlsx` that lacks it, so dropping a file into
the folder is enough. Both files are committed to the repository, which is what
makes a key survive a redeploy.

Current library (8 keys):

| File stem | Shown in the app as | Utterances |
|---|---|---|
| `KEY_Video_1` … `KEY_Video_4` | same as the stem | 47 / 34 / 32 / 34 |
| `KEY_PlayInterventionTraining_1` … `_4` | `Training Video 1…4 (Play Intervention)` | 36 / 29 / 14 / 26 |

The `KEY_PlayInterventionTraining_*` keys are the four Play Intervention training
videos, converted from the Revised coding scheme to the classic KEY format and
checked against the operational definitions manual (Sept 2026). They are separate
videos from `KEY_Video_1–4`; the names were kept distinct so neither set
overwrites the other.

## How to run

1. Open this folder in VS Code.
2. In Terminal, install requirements:

```bash
pip install -r requirements.txt
```

3. Run Streamlit:

```bash
streamlit run app.py
```

4. First go to `Reference Library` and add your key files.
5. Then go to `Check student file`, upload a student file, choose the suggested key, and download the checked Excel.

## Important MVP limitation

This first version compares utterances mostly by order after the reference is selected. The passport feature helps select the right reference by the first 5 utterances, but deeper utterance-to-utterance fuzzy matching can be added later.

## Files to edit without programming

- `grammar.yaml` — coding logic, Behavior names, aliases, modifier rules, colors.
- Files inside `reference_keys/` — reference key library.
