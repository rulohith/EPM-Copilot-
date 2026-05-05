"""Utility script to update the description of a LOB member in the LOBsToBeAdded.xlsx file.

It looks for a row where the member Name matches the target (default "L1") and updates
its Description column to the desired new description.
"""

import sys
from pathlib import Path

import openpyxl

def update_description(excel_path: Path, member_name: str, new_description: str) -> bool:
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    updated = False
    for row in ws.iter_rows(min_row=2, values_only=False):
        # Assuming columns: Name, Parent, Description
        name_cell, _, desc_cell = row[0], row[1], row[2]
        if name_cell.value == member_name:
            desc_cell.value = new_description
            updated = True
            print(f"Updated {member_name} description to '{new_description}'")
            break
    if not updated:
        # Append a new row if not found
        ws.append([member_name, "Total Division", new_description])
        print(f"Added new member {member_name} with description '{new_description}'")
    wb.save(excel_path)
    return updated

if __name__ == "__main__":
    """Entry point for the script.

    The original script used hard‑coded values (member="L1", description="NewBarcelona").
    For the automation rules we need to be able to add *any* member with a custom
    description (the parent is always "Total Division" for LOB dimension updates).

    The script now accepts command‑line arguments:
        python edit_lob_excel.py <member_name> <description>
    If the arguments are not supplied, it falls back to the original defaults
    to retain backward compatibility.
    """
    excel_file = Path("LOBsToBeAdded.xlsx")

    # Determine member name and description from CLI if provided
    if len(sys.argv) >= 3:
        member = sys.argv[1]
        description = sys.argv[2]
    else:
        # Preserve original behaviour for legacy calls
        member = "L1"
        description = "NewBarcelona"

    update_description(excel_file, member, description)