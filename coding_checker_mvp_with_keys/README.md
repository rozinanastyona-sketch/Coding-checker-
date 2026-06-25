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
