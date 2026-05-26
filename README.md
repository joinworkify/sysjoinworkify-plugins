# sysjoinworkify-plugins

Convert SYS Daily Cost Sheet Excel files into ERPNext Journal Entry import format.

## Requirements

- Python 3.12+
- [uv](https://github.com/astral-sh/uv) (recommended) or pip

Install dependencies:

```bash
uv sync
# or
pip install -r requirements.txt
```

## Usage

```bash
python main.py --input <file> --output <file> [--sheets <sheet1> <sheet2> ...]
```

### Arguments

| Argument | Short | Required | Description |
|---|---|---|---|
| `--input` | `-i` | Yes | Path to the input `.xlsx` file |
| `--output` | `-o` | Yes | Path to save the output `.xlsx` file |
| `--sheets` | `-s` | No | Sheet name(s) to process. If omitted, an interactive selector appears. |

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

**Interactive mode (no --sheets):**

```bash
python main.py -i input.xlsx -o output.xlsx
```

The program will list all sheets in `input.xlsx` and prompt you to select.

## Output

Produces a styled `.xlsx` file formatted for ERPNext Journal Entry bulk import:

- Each transaction generates two rows (debit leg + credit leg)
- Rows with `#N/A` source account are flagged for manual review
- Header row is frozen at row 1
