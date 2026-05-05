"""Utility script to clean AccountsToBeAdded.xlsx.

The `update_account_csv.py` script merges account data from this Excel file
into exported CSV files before importing them back into Hyperion. A duplicate
member name (e.g., ``LOB_1001``) can cause the import job to fail with the
error:

    Member name and its alias must be unique among siblings.

This script removes any rows in the Excel file where the **Name** column
matches a given list of problematic members. By default it removes
``LOB_1001`` but the list can be extended via the ``PROBLEMATIC_MEMBERS``
constant.

The script is safe to run multiple times; if the problematic rows are
already absent, the workbook is left unchanged.
"""

from pathlib import Path
import openpyxl

# Path to the Excel file relative to the project root
EXCEL_PATH = Path(__file__).resolve().parent / "AccountsToBeAdded.xlsx"

# List of member names that should be removed to avoid duplicate errors.
# Extend this list if other duplicate members are discovered.
PROBLEMATIC_MEMBERS = {"LOB_1001"}

def clean_excel(excel_path: Path) -> None:
    """Remove rows whose first column (Name) is in ``PROBLEMATIC_MEMBERS``.

    The function preserves the header row and rewrites the worksheet in‑place.
    """
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    # Read all rows (including header)
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        print("Excel file is empty.")
        return

    header = rows[0]
    # Filter out rows where the Name column (index 0) is problematic
    filtered = [header] + [r for r in rows[1:] if (r[0] not in PROBLEMATIC_MEMBERS)]

    # Clear the worksheet
    ws.delete_rows(1, ws.max_row)

    # Write back filtered rows
    for r_idx, row in enumerate(filtered, start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # Save to a temporary file first to avoid permission issues if the original
    # file is locked by another process. Once saved we atomically replace the
    # original file.
    import os
    # Save the cleaned workbook to a new file to avoid permission issues.
    cleaned_path = excel_path.with_name(excel_path.stem + "_clean" + excel_path.suffix)
    wb.save(cleaned_path)
    wb.close()
    print(f"Cleaned data saved to {cleaned_path.name}. Original file left unchanged.")


if __name__ == "__main__":
    clean_excel(EXCEL_PATH)