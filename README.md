# Essay Grader

Automatically grades student `.docx` essays using Claude AI and outputs a
color-coded Excel gradebook.

---

## Setup (one time)

**1. Install Python 3.9+**
If you don't already have it: https://www.python.org/downloads/

**2. Install dependencies**
Open Terminal (Mac) or Command Prompt (Windows), navigate to this folder, and run:
```
pip install -r requirements.txt
```

**3. Get an Anthropic API key**
- Go to https://console.anthropic.com
- Create an account and generate an API key
- Paste it into `config.txt` (see below)

---

## How to use

### Step 1 — Fill in config.txt

Open `config.txt` and edit the values:

```
# Path to the folder containing student .docx essay files
essays_folder = ./essays

# Path to your rubric file (.txt or .docx)
rubric_file = rubric.txt

# Your Anthropic API key
api_key = sk-ant-...

# Output filename for the Excel gradebook
output_file = grades.xlsx

# Grade level shown in the spreadsheet header (e.g. 10th Grade)
grade_level = 10th Grade

# Class period shown in the spreadsheet header (e.g. 1, 5, B)
period = 5

# How strictly Claude should grade (easy, medium, hard)
grading_difficulty = medium
```

`grade_level`, `period`, and `grading_difficulty` are optional — leave them blank to use the defaults.

### Step 2 — Run the program

```
python grade_essays.py
```

That's it. The program reads everything from `config.txt`.

### Using a different config file

If you teach multiple classes, you can keep a separate config file for each:

```
python grade_essays.py period1_config.txt
python grade_essays.py period5_config.txt
```

---

## Folder structure

```
essay_grader/
├── grade_essays.py       ← the program
├── config.txt            ← your settings (edit this)
├── requirements.txt
├── rubric.txt            ← your rubric
├── essays/               ← folder of student .docx submissions
│   ├── Sara_S_-_Essay.docx
│   ├── Chana_M_-_Essay.docx
│   └── ...
└── grades.xlsx           ← output (created automatically)
```

---

## Rubric format

The rubric can be a plain `.txt` file or a `.docx`. The program automatically
detects the assignment title and grading categories. Just write your rubric
naturally — anything with scoring criteria and point values will work.

**Example rubric.txt:**
```
Essay Rubric: Three Ineffective Solutions for Peace

Introduction (4 points)
  4 - Compelling, specific, fully sets up the topic. 5+ sentences.
  3 - Complete and mostly clear.
  2 - Vague or only partially addresses the topic.
  1 - Missing information or off-topic.

Body Paragraphs (4 points each, x3)
  4 - Complete, flows well, highly convincing.
  ...

Conclusion (4 points)
  ...

Spelling and Grammar (4 points)
  4 - 0-1 errors
  3 - 2-4 errors
  ...

Formatting (4 points)
  ...

Transitions (4 points)
  ...
```

---

## Output

The Excel file includes:
- Title row with the essay name (pulled from the rubric automatically)
- Subtitle row with grade level and/or period (if provided in config.txt)
- One row per student
- Color-coded scores (green=4, yellow=3, orange=2, red=1, gray=missing)
- Status column (Complete / Incomplete / Minimal / Blank)
- Total score and letter grade
- Brief AI-generated teacher notes for each student
- A Legend tab explaining the color coding

---

## Grading difficulty

The `grading_difficulty` setting controls how strictly Claude applies the rubric:

| Value | Behavior |
|-------|----------|
| `easy` | Generous; gives benefit of the doubt and liberal partial credit |
| `medium` | Straightforward rubric application — neither inflated nor deflated (default) |
| `hard` | High standard; top score only for fully convincing work, deducts for vagueness or weak evidence |

---

## Tips

- **File naming:** Student names are extracted from MLA header automatically.
- **API costs:** Each essay costs roughly $0.01–0.03 to grade (using Claude Opus).
  A class of 40 essays costs under $1.
- **Multiple classes:** Use separate config files per class and run with
  `python grade_essays.py period1_config.txt`.
- **API key security:** Never share your `config.txt` if it contains your API key.
  Alternatively, leave `api_key` blank and set the `ANTHROPIC_API_KEY`
  environment variable on your computer instead.
