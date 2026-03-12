# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

A command-line tool that grades student essays using Claude AI and outputs a color-coded Excel gradebook. Teachers configure it with a rubric file and folder of student `.docx` essays.

## Running the Tool

```bash
# Install dependencies (one time)
pip install -r requirements.txt

# Run with default config
python grade_essays.py

# Run with a specific period's config (multi-class use)
python grade_essays.py P1_config.txt
```

No build step, test suite, or CI configuration exists.

## Architecture

Everything lives in `grade_essays.py` (~772 lines, single file).

**Data flow:**
1. `load_config()` — reads a `key = value` config file (path optionally passed as CLI arg)
2. `parse_rubric_metadata()` — extracts assignment title, category names, weights, and total points from rubric `.docx` or `.txt`
   - Primary path: direct `.docx` table parsing via `_parse_rubric_from_docx()`
   - Fallback: `_parse_rubric_with_ai()` sends plain-text rubric to Claude
3. `load_examples()` — optional few-shot examples (paired student essays + teacher-graded rubrics) loaded from `examples_folder`
4. Per-essay loop: `extract_student_name()` → `read_docx()` → `grade_with_ai()`
5. `grade_with_ai()` — calls Claude (`claude-opus-4-6`) with rubric + essay, expects strict JSON back: `{status, categories, notes}`. Retries up to 3× on parse/rate-limit errors.
6. Output: `build_excel()` creates color-coded `.xlsx`; `write_score_sheets()` writes a parallel `.txt` summary

## Config Format

Config files are simple `key = value` text files:

| Key | Required | Description |
|---|---|---|
| `essays_folder` | Yes | Path to folder of student `.docx` files |
| `rubric_file` | Yes | Path to rubric (`.txt` or `.docx`) |
| `api_key` | Yes* | Anthropic API key (or set `ANTHROPIC_API_KEY` env var) |
| `output_file` | No | Output `.xlsx` path (default: `grades.xlsx`) |
| `examples_folder` | No | Folder of paired example essays + rubrics for few-shot prompting |
| `grade_level` | No | Shown in spreadsheet header (e.g. `10th Grade`) |
| `period` | No | Class period shown in header (e.g. `1`, `5`, `B`) |

## Security Note

The `P*_config.txt` files may contain live API keys — do not commit them. Consider adding `P*_config.txt` to `.gitignore`.
