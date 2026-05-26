import argparse
from plugins.convert_to_erpnext import run


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
        required=True,
        help="Sheet name(s) to process (space-separated, quote names with spaces)",
    )

    args = parser.parse_args()
    run(input_file=args.input, output_file=args.output, sheets=args.sheets)


if __name__ == "__main__":
    main()
