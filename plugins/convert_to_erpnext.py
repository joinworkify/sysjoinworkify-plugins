"""
Convert SYS Daily Cost Sheet -> ERPNext Journal Entry Import Format

Logic per row:
  - Cost column (positive)  → Debit Src account, Credit Sub Cat account
  - Cost column (negative)  → Credit Src account, Debit Sub Cat account
  - Other column (positive) → Credit Src account, Debit Sub Cat account
  - Other column (negative) → Debit Src account, Credit Sub Cat account
  - Row with #N/A as Src    → use Sub Cat as both debit/credit account pair
"""

import pandas as pd
from datetime import datetime

COMPANY = "Seinn Yaung So MFG"
ENTRY_TYPE = "Journal Entry"
SERIES = "ACC-JV-.YYYY.-"
IS_OPENING = "No"


def extract_account_code(account_str):
    """Return the account string as-is (ERPNext uses full name)."""
    if pd.isna(account_str) or str(account_str).strip() in ("#N/A", ""):
        return None
    return str(account_str).strip()


def process_sheet(input_file, sheet_name):
    df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)

    # Row 1 (index 1) is the header
    df.columns = df.iloc[1]
    df = df.iloc[2:].reset_index(drop=True)
    df.columns = [
        "Date",
        "Place",
        "Src",
        "Category",
        "Sub Cat",
        "Particular",
        "Cost",
        "Other",
    ]

    rows = []
    jv_counter = 0

    for _, row in df.iterrows():
        date = row["Date"]
        if pd.isna(date):
            continue

        # Convert serial date if needed
        if isinstance(date, (int, float)):
            try:
                date = pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(date))
            except Exception:
                continue
        elif not isinstance(date, (pd.Timestamp, datetime)):
            try:
                date = pd.Timestamp(date)
            except Exception:
                continue

        posting_date = date.strftime("%Y-%m-%d")

        src = extract_account_code(row["Src"])
        sub_cat = extract_account_code(row["Sub Cat"])
        particular = (
            str(row["Particular"]).strip() if not pd.isna(row["Particular"]) else ""
        )

        cost_raw = row["Cost"]
        other_raw = row["Other"]

        cost = float(cost_raw) if not pd.isna(cost_raw) else None
        other = float(other_raw) if not pd.isna(other_raw) else None

        # Determine which amount column is used
        if cost is not None and cost != 0:
            amount = cost
            use_cost = True
        elif other is not None and other != 0:
            amount = other
            use_cost = False
        else:
            continue  # no amount, skip

        src_is_na = (src is None) or (str(row["Src"]).strip() == "#N/A")

        abs_amount = abs(amount)

        if use_cost:
            if amount > 0:
                debit_acct = sub_cat
                credit_acct = src
            else:
                debit_acct = src
                credit_acct = sub_cat
        else:
            if amount > 0:
                debit_acct = sub_cat
                credit_acct = src
            else:
                debit_acct = src
                credit_acct = sub_cat

        if src_is_na:
            credit_acct = sub_cat
            debit_acct = sub_cat
            note = "[SRC=N/A - needs manual account assignment] " + particular
        else:
            note = particular

        jv_id = f"ACC-JV-2026-{jv_counter:05d}"
        jv_counter += 1

        rows.append(
            {
                "ID": jv_id,
                "Company": COMPANY,
                "Entry Type": ENTRY_TYPE,
                "Series": SERIES,
                "Title": particular[:140] if particular else "",
                "Posting Date": posting_date,
                "Debit (Accounting Entries)": None,
                "Credit (Accounting Entries)": abs_amount,
                "Is Opening": IS_OPENING,
                "ID (Accounting Entries)": credit_acct,
                "Account (Accounting Entries)": credit_acct,
                "User Remark (Accounting Entries)": note,
            }
        )

        rows.append(
            {
                "ID": None,
                "Company": None,
                "Entry Type": None,
                "Series": None,
                "Title": None,
                "Posting Date": None,
                "Debit (Accounting Entries)": abs_amount,
                "Credit (Accounting Entries)": None,
                "Is Opening": None,
                "ID (Accounting Entries)": debit_acct,
                "Account (Accounting Entries)": debit_acct,
                "User Remark (Accounting Entries)": note,
            }
        )

    return rows


def run(input_file, output_file, sheets):
    all_rows = []

    for sheet in sheets:
        try:
            sheet_rows = process_sheet(input_file, sheet)
            all_rows.extend(sheet_rows)
            print(f"  {sheet}: {len(sheet_rows)//2} journal entries")
        except Exception as e:
            print(f"  Skipped {sheet}: {e}")

    if not all_rows:
        print("No data processed.")
        return

    out_df = pd.DataFrame(
        all_rows,
        columns=[
            "ID",
            "Company",
            "Entry Type",
            "Series",
            "Title",
            "Posting Date",
            "Debit (Accounting Entries)",
            "Credit (Accounting Entries)",
            "Is Opening",
            "ID (Accounting Entries)",
            "Account (Accounting Entries)",
            "User Remark (Accounting Entries)",
        ],
    )

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "JV Import"

    headers = list(out_df.columns)
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    jv_fill = PatternFill("solid", fgColor="DEEAF1")
    normal_font = Font(name="Arial", size=9)
    jv_font = Font(name="Arial", size=9, bold=True)

    for row_idx, (_, data_row) in enumerate(out_df.iterrows(), 2):
        is_jv_header = (
            not pd.isna(data_row["ID"]) if data_row["ID"] is not None else False
        )
        for col_idx, col in enumerate(headers, 1):
            val = data_row[col]
            if pd.isna(val) if not isinstance(val, str) else False:
                val = None
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = jv_font if is_jv_header else normal_font
            if is_jv_header:
                cell.fill = jv_fill

    col_widths = [22, 20, 15, 18, 60, 14, 28, 28, 12, 45, 45, 60]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"

    wb.save(output_file)
    print(f"\nSaved: {output_file}")
    print(f"Total rows: {len(all_rows)} ({len(all_rows)//2} journal entries)")


if __name__ == "__main__":
    # Legacy direct run with defaults
    run(
        input_file="/mnt/user-data/uploads/shh_Detail_Daily_Cost_-_Yangon__2_3_4_2026_.xlsx",
        output_file="/mnt/user-data/outputs/erpnext_jv_import.xlsx",
        sheets=[
            "RealReferenceSheet - Feb",
            "RealReferenceSheet - Mar",
            "RealReferenceSheet - Apr",
        ],
    )
