# sysjoinworkify-plugins

Convert SYS Daily Cost Sheet Excel files into ERPNext Journal Entry import format.

## Requirements

- Python 3.12+
- [uv](https://github.com/astral-sh/uv) (recommended) or pip

Install dependencies:

```bash
uv sync
```

## Usage

```bash
python main.py --input <file> --output <file> [--sheets <sheet1> <sheet2> ...] [--output-flagged <file>]
```

### Arguments

| Argument | Short | Required | Description |
|---|---|---|---|
| `--input` | `-i` | Yes | Path to the input `.xlsx` file |
| `--output` | `-o` | Yes | Path to save the **clean** output `.xlsx` file |
| `--sheets` | `-s` | No | Sheet name(s) to process. If omitted, an interactive selector appears. |
| `--output-flagged` | `-of` | No | Path for the flagged output file. Auto-derived from `--output` if omitted (e.g. `out.xlsx` → `out_flagged.xlsx`). |

### Interactive sheet selection

If `--sheets` is not provided, the program reads the available sheets from the input file and shows an interactive selector:

```
Select sheets to process (Space to select, Enter to confirm):
 ○ RealReferenceSheet - Jan
 ● RealReferenceSheet - Feb
 ● RealReferenceSheet - Mar
 ○ RealReferenceSheet - Apr
```

- `Space` — toggle selection
- `Enter` — confirm and run

## Output files

The program produces **two output files**:

| File | Content | Header color |
|---|---|---|
| `--output` | Clean entries only — ready to import into ERPNext | Blue |
| `--output-flagged` | Flagged entries that need manual review | Orange |

### Flagged conditions

A journal entry is flagged and excluded from the clean file when:

| Reason | Meaning |
|---|---|
| `SRC=N/A` | The source account cell is `#N/A` or empty — no valid debit/credit account |
| `SRC==SUB_CAT` | Source account and Sub Category account resolve to the same value — debit and credit would be the same account |

Flagged entries have a `[REASON - needs manual account assignment]` prefix in the **User Remark** column.

### Terminal output example

```
  RealReferenceSheet - Feb: 142 clean, 8 flagged
  RealReferenceSheet - Mar: 98 clean, 4 flagged

Saved (clean): erpnext_jv_import.xlsx
  240 journal entries

WARNING: 12 flagged journal entries need manual review
  - SRC=N/A      : 8 entries
  - SRC==SUB_CAT : 4 entries
Saved (flagged): erpnext_jv_import_flagged.xlsx
```

## Examples

**With sheet names specified:**

```bash
python main.py \
  --input "Detail_Daily_Cost_Yangon_2026.xlsx" \
  --output "erpnext_jv_import.xlsx" \
  --sheets "RealReferenceSheet - Feb" "RealReferenceSheet - Mar" "RealReferenceSheet - Apr"
```

**Short flags:**

```bash
python main.py -i input.xlsx -o output.xlsx -s "Sheet Feb" "Sheet Mar"
```

**Custom flagged output path:**

```bash
python main.py -i input.xlsx -o output.xlsx -s "Sheet Feb" -of review/flagged_entries.xlsx
```

**Interactive mode (no --sheets):**

```bash
python main.py -i input.xlsx -o output.xlsx
```
