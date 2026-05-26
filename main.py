import argparse
import sys
import pandas as pd
import questionary
from plugins.convert_to_erpnext import run


def get_sheets_interactive(input_file):
    """Read available sheets from the file and let user pick via checkbox prompt."""
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
        help="Path to output .xlsx file",
    )
    parser.add_argument(
        "--sheets", "-s",
        nargs="+",
        required=False,
        help="Sheet name(s) to process. If omitted, interactive selector will appear.",
    )

    args = parser.parse_args()

    sheets = args.sheets if args.sheets else get_sheets_interactive(args.input)

    run(input_file=args.input, output_file=args.output, sheets=sheets)


if __name__ == "__main__":
    main()
