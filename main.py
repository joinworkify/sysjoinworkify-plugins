import argparse
import sys
import zipfile
from pathlib import Path

import pandas as pd
import questionary

from plugins.convert_to_erpnext import run

# ── Output configuration ─────────────────────────────────────────────────────
DEFAULT_OUTPUT_DIR = Path.home() / "Downloads"
# Change to e.g. Path.home() / "Documents" / "ERPNext" if preferred
# ─────────────────────────────────────────────────────────────────────────────


def resolve_output_path(raw: str) -> str:
    """If raw is a plain filename (no directory), place it in DEFAULT_OUTPUT_DIR."""
    p = Path(raw)
    if p.parent == Path("."):
        return str(DEFAULT_OUTPUT_DIR / p)
    return str(p)


def get_sheets_interactive(input_file):
    try:
        xl = pd.ExcelFile(input_file)
        available = xl.sheet_names
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)

    if not available:
        print("No sheets found in file.")
        sys.exit(1)

    selected = questionary.checkbox(
        "Select sheets to process (Space to select, Enter to confirm):",
        choices=available,
    ).ask()

    if not selected:
        print("No sheets selected. Exiting.")
        sys.exit(0)

    return selected


def main():
    parser = argparse.ArgumentParser(
        description="Convert SYS Daily Cost Sheet to ERPNext Journal Entry import format."
    )
    parser.add_argument(
        "--input", "-i",
        required=True,
        help="Path to input .xlsx file",
    )
    parser.add_argument(
        "--output", "-o",
        required=True,
        help=(
            "Output filename or full path for clean entries. "
            f"Plain filenames are saved to {DEFAULT_OUTPUT_DIR}."
        ),
    )
    parser.add_argument(
        "--sheets", "-s",
        nargs="+",
        required=False,
        help="Sheet name(s) to process. If omitted, interactive selector appears.",
    )
    parser.add_argument(
        "--output-flagged", "-of",
        required=False,
        default=None,
        dest="output_flagged",
        help="Path for flagged output file. Auto-derived from --output if omitted.",
    )
    parser.add_argument(
        "--zip",
        action="store_true",
        help="Zip both output files into a single .zip and remove the .xlsx files.",
    )

    args = parser.parse_args()

    output = resolve_output_path(args.output)
    output_flagged = resolve_output_path(args.output_flagged) if args.output_flagged else None

    # Ensure output directory exists
    Path(output).parent.mkdir(parents=True, exist_ok=True)

    sheets = args.sheets if args.sheets else get_sheets_interactive(args.input)

    written_clean, written_flagged = run(
        input_file=args.input,
        output_file=output,
        sheets=sheets,
        output_flagged_file=output_flagged,
    )

    if args.zip:
        files_to_zip = [p for p in [written_clean, written_flagged] if p and Path(p).exists()]
        if files_to_zip:
            zip_path = Path(output).with_suffix(".zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for f in files_to_zip:
                    zf.write(f, Path(f).name)
            for f in files_to_zip:
                Path(f).unlink()
            print(f"\nZipped: {zip_path}")


if __name__ == "__main__":
    main()
