#!/usr/bin/python

from openpyxl import load_workbook
import shutil
import sys


def process_excel(file_name):
    """
    Process an Excel file to copy values (not formulas) from closing stock columns
    to opening stock columns, leaving closing stock data unchanged.

    Args:
        file_name (str): The name of the Excel file to process.
    """

    # Create a copy to edit
    updated_file_name = f"updated-{file_name}"
    shutil.copy(file_name, updated_file_name)

    # Define sheet details
    sheets = [
        {"name": "LUBES STORE", "closing_stock_column": "H",
            "opening_stock_column": "D", "cells_to_clear": ["E", "F"]},
        {"name": "LUBES SHOP SUMMARY", "closing_stock_column": "H",
            "opening_stock_column": "D", "cells_to_clear": ["E", "F", "G"]},
        {"name": "LUBES CAGE SUMMARY", "closing_stock_column": "G",
            "opening_stock_column": "D", "cells_to_clear": ["E", "F"]},
    ]
    password = "AA"

    try:
        # Load original (read-only) and updated (editable) workbooks
        original_workbook = load_workbook(file_name, data_only=True)
        updated_workbook = load_workbook(updated_file_name)
    except FileNotFoundError:
        print(f"Error: File '{file_name}' not found.")
        return

    # Iterate over each sheet in the configuration
    for sheet_config in sheets:
        sheet_name = sheet_config["name"]
        closing_stock_column = sheet_config["closing_stock_column"]
        opening_stock_column = sheet_config["opening_stock_column"]
        columns_to_clear = sheet_config["cells_to_clear"]

        # Check if the sheet exists
        if sheet_name not in original_workbook.sheetnames:
            print(f"Sheet '{sheet_name}' not found. Skipping...")
            continue

        original_sheet = original_workbook[sheet_name]
        updated_sheet = updated_workbook[sheet_name]

        # Copy data from closing_stock_column to opening_stock_column
        for row in range(6, 30):  # Start from row 6
            closing_cell = original_sheet[f"{closing_stock_column}{row}"]
            opening_cell = updated_sheet[f"{opening_stock_column}{row}"]

            # Copy value if present
            if closing_cell.value is not None:
                opening_cell.value = closing_cell.value
                print(f"{opening_cell} changed to {closing_cell}")

            # Clear needful cells
            for column_to_clear in columns_to_clear:
                cell_to_clear = updated_sheet[f"{column_to_clear}{row}"]

                if type(cell_to_clear.value) is not str:
                    cell_to_clear.value = None

    # Save the updated workbook
    updated_workbook.save(updated_file_name)

    # Close workbooks
    original_workbook.close()
    updated_workbook.close()

    print(f"Workbook updated and saved as '{updated_file_name}'.")


# Entry point for the program
if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script_name.py <excel_file_name>")
    else:
        excel_file_name = sys.argv[1]
        process_excel(excel_file_name)
