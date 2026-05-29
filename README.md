# docforge

> Metadata analysis & copy-paste detection for Microsoft Office files

`docforge` is a Python CLI toolkit that dissects the internal metadata of Office files to detect copy-paste patterns, trace edit history, and produce a human- and machine-readable risk score. Includes a cleaner to sanitize metadata and rotate RSIDs.

---

## Table of Contents

- [Background](#background)
- [Supported Formats](#supported-formats)
- [Installation](#installation)
- [Usage](#usage)
  - [docforge.py — Analyzer](#docforgepy--analyzer)
  - [docforge_cleaner.py — Cleaner](#docforge_cleanerpy--cleaner)
- [How It Works](#how-it-works)
- [Output](#output)
- [Interpreting Results](#interpreting-results)
- [Limitations](#limitations)
- [Project Structure](#project-structure)

---

## Background

Modern Office files (`.docx`, `.pptx`, `.xlsx`) are ZIP archives containing a collection of XML files. Inside these XMLs lies metadata that most users are unaware of, including:

- **Who created** and **who last edited** the document
- **How many times** the document was saved
- **RSID** (*Revision Save ID*) — a unique ID that Word assigns to each editing session

By analyzing this metadata, `docforge` can identify patterns indicating that document content was not written organically, but instead moved from another source in a single large operation.

---

## Supported Formats

| Extension | Format | RSID Analysis |
|---|---|---|
| `.docx` `.dotx` | Word Document | ✅ Full |
| `.pptx` `.potx` | PowerPoint | ✅ Metadata + slide count |
| `.xlsx` `.xlsm` | Excel | ✅ Metadata + sheet info |

---

## Installation

**Prerequisites:** Python 3.8 or newer.

```bash
# 1. Clone the repo
git clone https://github.com/ElNu-mp4/docforge.git
cd docforge

# 2. Install dependencies
pip install rich

# 3. (Optional) Make globally executable on Linux/macOS
chmod +x docforge.py
sudo cp docforge.py /usr/local/bin/docforge
```

No other dependencies — modules like `zipfile`, `xml.etree`, and `json` are all part of the Python standard library.

---

## Usage

### docforge.py — Analyzer

```bash
python3 docforge.py <file> [options]
```

| Argument/Option | Description |
|---|---|
| `file` | Path to the Office file to analyze |
| `--json` | Print JSON output below the dashboard |
| `--out <path>` | Save JSON report to a file |
| `--no-dashboard` | JSON only, no terminal dashboard |
| `-h`, `--help` | Show help |

**Examples:**

```bash
# Basic analysis — show terminal dashboard
python3 docforge.py report.docx

# Analyze and save JSON report
python3 docforge.py report.docx --out analysis.json

# JSON only (useful for pipelines/scripting)
python3 docforge.py report.docx --no-dashboard

# Show dashboard AND print JSON
python3 docforge.py report.docx --json

# Analyze a PowerPoint or Excel file
python3 docforge.py slides.pptx
python3 docforge.py data.xlsx --out report.json
```

**Example terminal output:**

```
──────── docforge  •  report.docx ────────

╭──────────────── Copy-Paste Risk Score ───────────────╮
│                                                       │
│  MEDIUM   ██████████░░░░░░░░░░   50/100              │
│                                                       │
│  ⚠  Document created by 'Sandy', edited by 'M S I'   │
│  ⚠  Only saved 3 times (very few)                    │
│  ⚠  Dominant RSID 95.7%: almost all text one session │
│                                                       │
╰───────────────────────────────────────────────────────╯

📄  Document Metadata
┌────────────────────────┬──────────────────────────┐
│ Created by             │ Sandy Kurniawan           │
│ Last edited by         │ M S I                     │
│ Revision               │ 3                         │
│ Date created           │ 27 Feb 2026, 05:41 UTC    │
│ Last saved             │ 02 May 2026, 23:17 UTC    │
│ Application            │ Microsoft Office Word     │
└────────────────────────┴──────────────────────────┘

📈  RSID Distribution (run level)
  00C81389    1080   95.7%   ███████████████████░
  00B96628      11    1.0%   ░░░░░░░░░░░░░░░░░░░░
  003030D0       7    0.6%   ░░░░░░░░░░░░░░░░░░░░
```

---

### docforge_cleaner.py — Cleaner

Cleans metadata and rotates RSIDs in `.docx` files so the docforge risk score = 0.

```bash
python3 docforge_cleaner.py <file> [options]
```

| Option | Description |
|---|---|
| `--author` | Author name to embed in metadata |
| `--revision` | Revision count (default: auto-calculated from word count) |
| `--total-time` | Total edit time in minutes (default: calculated from word count) |
| `--rotation-freq N` | RSID rotation frequency per N paragraphs (default: 5) |
| `--out`, `-o` | Output path (default: `input_clean.docx`) |
| `--dry-run` | Simulate without writing any file |
| `--verbose`, `-v` | Show RSID details |
| `--skip-rsid` | Skip RSID rotation |
| `--skip-core` | Skip core.xml patch |
| `--skip-app` | Skip app.xml patch |

**Examples:**

```bash
# Clean with a new author
python3 docforge_cleaner.py assignment.docx --author "Your Name"

# Manually set revision count and total edit time
python3 docforge_cleaner.py assignment.docx --author "Name" --revision 18 --total-time 150

# Set RSID rotation frequency and output path
python3 docforge_cleaner.py assignment.docx --rotation-freq 4 --out clean.docx

# Dry run
python3 docforge_cleaner.py assignment.docx --dry-run --verbose
```

**`--rotation-freq` guide:**

| Value | Effect |
|---|---|
| 2–3 | Many sessions, very even RSID distribution |
| 5–7 | Default, suitable for most documents |
| 10+ | Few sessions (short documents, edited in one sitting) |

---

## How It Works

### Analyzer (`docforge.py`)

Office files are ZIP archives. `docforge` opens them without extracting to disk, then reads three layers of metadata:

**1. Core Metadata (`docProps/core.xml`)**

| XML Field | Meaning |
|---|---|
| `dc:creator` | Username who first created the document |
| `cp:lastModifiedBy` | Username who last saved |
| `cp:revision` | Cumulative save count |
| `dcterms:created` | Creation timestamp |
| `dcterms:modified` | Last saved timestamp |

**2. App Metadata (`docProps/app.xml`)**

| XML Field | Meaning |
|---|---|
| `TotalTime` | Total minutes the document was open in edit mode |
| `Words` | Word count |
| `Pages` | Page count |
| `Application` | Application used (e.g. "Microsoft Office Word") |
| `Template` | Template used at creation |

**3. RSID Analysis — `.docx` only (`word/document.xml` + `word/settings.xml`)**

RSID (*Revision Save ID*) is an 8-character hex value that Word randomly assigns to each new editing session, recorded on `w:rsidR` in `<w:p>` (paragraph) and `<w:r>` (run/text) elements.

| Indicator | Detection Method | Interpretation |
|---|---|---|
| **Foreign RSIDs** | RSIDs in `document.xml` not registered in `settings.xml` | Content pasted from a different Word document |
| **RSID dominance** | % of the most common RSID out of all runs | >80% = most text came from a single session |
| **Paragraph↔run mismatch** | Paragraph RSID ≠ run RSIDs within it | Paragraph structure from one session, text content from another |
| **rsidRoot** | Value of `<w:rsidRoot>` | Origin identity of the document / first creation session |

**4. Risk Scoring**

| Condition | Score Added |
|---|---|
| Creator ≠ lastModifiedBy | +15 |
| Revision = 1 | +20 |
| Revision ≤ 3 | +10 |
| Dominant RSID ≥ 95% | +25 |
| Dominant RSID ≥ 80% | +15 |
| Foreign RSIDs found | +30 |
| Paragraph↔run mismatch ≥ 20% | +15 |
| Edit speed > 100 words/min | +20 |
| Edit speed > 50 words/min | +10 |

| Score | Level | Color |
|---|---|---|
| 0 – 39 | LOW | Green |
| 40 – 69 | MEDIUM | Yellow |
| 70 – 100 | HIGH | Red |

### Cleaner (`docforge_cleaner.py`)

- Rotates RSIDs in `document.xml` and `settings.xml` via raw bytes (no full XML re-parse)
- Synchronizes run RSIDs with their parent paragraph
- Adds random jitter between sessions for natural distribution
- Patches `core.xml`: author, timestamps, revision
- Patches `app.xml`: application, total edit time (calculated from word count)

---

## Output

### Terminal Dashboard

Displayed by default. Consists of:

1. **Risk banner** — score, level, and list of flags
2. **Document metadata table** — identity and application info
3. **Statistics table** — words, pages, time, speed
4. **RSID table** — session analysis summary (`.docx` only)
5. **RSID distribution** — ASCII bar chart per session
6. **Mismatch examples** — text snippets with inconsistent RSIDs

### JSON Report

Generated with `--json` or `--out`. Structure:

```json
{
  "file": "report.docx",
  "format": "docx",
  "analyzed": "2026-05-03T06:26:33+00:00",
  "risk": {
    "score": 50,
    "level": "MEDIUM",
    "flags": ["Document created by 'Sandy', edited by 'M S I'", "..."],
    "notes": []
  },
  "core_metadata": {
    "creator": "Sandy Kurniawan",
    "last_modified_by": "M S I",
    "revision": "3",
    "created": "2026-02-27T05:41:00Z",
    "modified": "2026-05-02T23:17:00Z"
  },
  "app_metadata": {
    "application": "Microsoft Office Word",
    "total_time": "426",
    "words": "8083",
    "pages": "50"
  },
  "rsid_analysis": {
    "rsid_root": "0092077F",
    "registered_count": 107,
    "unique_run_rsids": 12,
    "dominant_rsid": "00C81389",
    "dominant_count": 1080,
    "dominant_pct": 95.7,
    "foreign_rsids": [],
    "total_para": 516,
    "mismatch_para": 5,
    "mismatch_pct": 1.0
  }
}
```

---

## Interpreting Results

### High RSID Dominance (>80%)

This does not automatically indicate plagiarism. Two common scenarios:

**Normal scenario:** A student fills in a lecturer's template in one long work session. All text typed in that session gets the same RSID. This is normal and common.

**Suspicious scenario:** Someone pastes content from outside Word (Google Docs, Notepad, web) and saves. All pasted text gets the RSID of that session, causing extreme dominance.

Distinguish between them by looking at **edit speed** (words/min) and **revision count**. High dominance + very few revisions + unreasonable speed → more suspicious.

### Foreign RSIDs Found

This is the **strongest indicator**. Foreign RSIDs appear when someone opens two different Word documents on the same computer and copy-pastes between them. The pasted text carries the RSID from the source document, which is not registered in the destination document's `settings.xml`.

### Creator ≠ Last Modified By

Very common and not always a problem — for example, a lecturer shares a template and a student fills it in. But in the context of assignment submissions between students, this flag warrants closer inspection.

### Paragraph↔Run Mismatch

Occurs when a paragraph structure (the `¶` mark) originates from one session, but the text content inside was written in a different session. This can happen due to deletion and retyping, or pasting that overwrites existing text.

---

## Limitations

- **Not legal evidence.** Analysis results are indicative, not conclusive. Metadata can be manually manipulated.
- **RSID only exists in `.docx`.** Formats `.pptx` and `.xlsx` have no equivalent RSID system, so their analysis is limited to general metadata.
- **Does not detect paraphrasing.** `docforge` analyzes technical metadata, not content similarity. For text-based plagiarism detection, use tools like Turnitin.
- **Paste from outside Word leaves no foreign RSIDs.** Content copied from a browser, PDF, or Google Docs produces no foreign RSIDs — only single-session RSID dominance.
- **Files re-saved in Word Online** may normalize RSIDs, erasing traces of previous sessions.

---

## Project Structure

```
docforge/
├── README.md
├── requirements.txt
├── docforge.py           # analyzer
└── docforge_cleaner.py   # cleaner
```

## Author

Elang N
