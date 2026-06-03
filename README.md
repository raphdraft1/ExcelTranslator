# ExcelTranslator

Translate Excel and PowerPoint files from any language to any language via Google Translate — preserving formatting, merged cells, and slide structure.

Built to solve a real problem: cross-border teams sharing bilingual documents shouldn't need to translate them manually.

---

## Features

- **Excel & PowerPoint support** — `.xlsx`, `.xls`, `.xlsm`, `.pptx`
- **Merged cell preservation** — unmerges, translates, re-merges; no data loss
- **Auto language detection** — no need to specify source language
- **Translation cache** — skips re-translating identical strings, speeds up large files significantly
- **Exponential backoff** — handles rate limiting gracefully with 2s → 4s → 8s retry intervals
- **Interactive column selection** — choose which columns to translate, or pass `--all-columns`
- **CLI interface** — clean `argparse`-based usage with sensible defaults

---

## Installation

```bash
git clone https://github.com/raphdraft1/ExcelTranslator
cd ExcelTranslator
pip install -r requirements.txt
```

**Requirements:**
```
pandas
openpyxl
deep-translator
tqdm
python-pptx
```

---

## Usage

**Translate an Excel file (interactive column selection):**
```bash
python ExcelTranslator.py --input report.xlsx
```

**Translate all text columns without prompting:**
```bash
python ExcelTranslator.py --input report.xlsx --all-columns
```

**Translate a PowerPoint file:**
```bash
python ExcelTranslator.py --input deck.pptx
```

**Specify source and target language:**
```bash
python ExcelTranslator.py --input data.xlsx --src ja --dest fr
```

**Custom output directory:**
```bash
python ExcelTranslator.py --input report.xlsx --output ./translated/
```

### All options

| Flag | Default | Description |
|------|---------|-------------|
| `--input` / `-i` | *(required)* | Path to input file |
| `--output` / `-o` | `./translated` | Output directory |
| `--src` / `-s` | `auto` | Source language code |
| `--dest` / `-d` | `en` | Target language code |
| `--all-columns` | `false` | Translate all text columns without prompting |

---

## Language Codes

Common codes for the `--src` and `--dest` flags:

| Language | Code |
|----------|------|
| Chinese (Simplified) | `zh-CN` |
| Chinese (Traditional) | `zh-TW` |
| English | `en` |
| Japanese | `ja` |
| Korean | `ko` |
| French | `fr` |
| Spanish | `es` |
| German | `de` |

Full list: [Google Translate language codes](https://cloud.google.com/translate/docs/languages)

---

## How It Works

1. **Excel**: Opens the workbook with `openpyxl`, records all merged cell ranges, temporarily unmerges, translates cell-by-cell, then re-merges and saves — preserving formatting throughout.

2. **PowerPoint**: Iterates over every shape's text frame and translates at the `run` level, keeping bold, font size, colour, and other formatting intact.

3. **Caching**: Every translated string is stored in an in-memory dict keyed by `src:dest:text`. Duplicate strings (common in structured documents) are returned instantly from cache.

4. **Rate limiting**: Uses `deep-translator` with exponential backoff on `TooManyRequests` errors — safer and more stable than `googletrans`.

---

## Example Output

```
17:42:03  INFO      Loading workbook: Q3_report.xlsx

Available sheets:
   1. Summary
   2. Raw Data
   3. Charts

Enter sheet number (or press Enter for all sheets): 1

Available columns:
    1. 项目名称  (text)
    2. 负责人    (text)
    3. 状态      (text)
    4. 预算      (non-text, skipped)

Enter column numbers to translate (comma-separated): 1,2,3

Translating columns: 100%|████████████████| 3/3 [00:14<00:00]

17:42:18  INFO      Saved: ./translated/Q3_report_translated.xlsx

✓  Done in 14.3s  →  ./translated/Q3_report_translated.xlsx
   Cache hits saved ~47 repeat translations.
```

---

## Limitations

- Uses the free Google Translate endpoint via `deep-translator` — no API key required, but rate limits apply on very large files. Add `time.sleep()` delays or process in batches if hitting limits consistently.
- Translation quality is Google Translate quality — suitable for operational/business documents, not literary or legal text.
- In-memory cache resets between runs. For repeated large-file use, extending the cache to disk (e.g. `shelve` or SQLite) would be a worthwhile addition.

---

## Background

This tool was originally built and deployed at [Perennial Holdings](https://perennialholdings.com.au/) to automate translation of internal operational documents between Chinese and English, eliminating a recurring manual bottleneck for cross-border teams.

---

## License

MIT
