# Check if collection summary file exists within COLLECTION SUMMARY folder in parent folder

import calendar
from datetime import datetime
import os
from openpyxl import Workbook


def get_days_in_month(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]


def month_str_to_int(month_name: str) -> int:
    return datetime.strptime(month_name, "%B").month


def get_year_from_file_name(file_name: str) -> int:
    file_name = os.path.basename(file_name)
    year = file_name.split(" ")[2].split(".")[0]

    return int(year)


def get_month_from_file_name(file_name: str) -> str:
    file_name = os.path.basename(file_name)
    month = file_name.split(" ")[1]

    return month


def get_day_from_file_name(file_name: str) -> int:
    file_name = os.path.basename(file_name)
    day_with_suffix = file_name.split(" ")[0]
    day_without_suffix = day_with_suffix[:-2]

    return int(day_without_suffix)


def create_summary_file_path(shift_change_file_path):
    """
    Gets the path for the summary file, given a shift change file path

    Args:
        shift_change_file_path (str): The path of the shift change excel file
    """
    month = get_month_from_file_name(shift_change_file_path)
    year = get_year_from_file_name(shift_change_file_path)
    parent_directory = os.path.dirname(shift_change_file_path)
    summary_file_path = os.path.join(
        parent_directory, "COLLECTION SUMMARY", f"{month} {year}.xlsx")

    return summary_file_path


def create_summary_file(shift_change_file_path, summary_workbook_map):
    """
    Creates the summary Excel workbook, adds the sheets required
    and creates the tables in each sheet
    """
    summary_workbook = Workbook()
    summary_file_path = create_summary_file_path(shift_change_file_path)

    # Retrieve summary workbook metadata
    sheet_names = summary_workbook_map["sheets"]["names"]
    table_title_cell = summary_workbook_map["table"]["title_cell"]
    table_column_titles = summary_workbook_map["table"]["column_titles"]
    date_column = summary_workbook_map["table"]["date_column"]
    first_date_row = summary_workbook_map["table"]["first_date_row"]

    # Get date metadata
    month = month_str_to_int(get_month_from_file_name(shift_change_file_path))
    year = get_year_from_file_name(shift_change_file_path)
    number_of_days = get_days_in_month(year, month)

    # Create sheets
    for sheet_name in sheet_names:
        summary_workbook.create_sheet(sheet_name)

        # Change column width
        for column in range(ord("C"), ord("J") + 1):
            summary_workbook[sheet_name].column_dimensions[chr(
                column)].width = 22

    summary_workbook.remove(summary_workbook["Sheet"])

    # Create tables in sheets
    for sheet_name in sheet_names:
        sheet = summary_workbook[sheet_name]
        sheet[table_title_cell] = sheet_name

        # Add table titles
        for cell, title in table_column_titles.items():
            sheet[cell] = title

        # Add dates
        for day in range(1, number_of_days + 1):
            cell = f"{date_column}{(first_date_row - 1) + day}"
            sheet[cell] = day

        # Add total formulas
        for column in range(ord("C"), ord("J") + 1):
            formula_cell = f"{chr(column)}{first_date_row + number_of_days}"
            sheet[formula_cell].value = f"=SUM({chr(column)}{first_date_row}:{chr(column)}{(number_of_days + first_date_row) - 1})"

    # Create COLLECTION SUMMARY directory
    try:
        os.makedirs(os.path.dirname(summary_file_path,), exist_ok=False)
    except FileExistsError:
        print("Error: A COLLECTION SUMMARY folder already exists in this directory!")
        exit(1)

    # Save workbook
    summary_workbook.save(summary_file_path)
    summary_workbook.close()

# TODO: Add formulas to calculate totals at bottom


def transfer_data(shift_change_file: object, summary_file: object, summary_map: dict,  day_number: int,):
    source_sheet = shift_change_file[summary_map["data_transfer"]
                                     ["source_sheet"]]
    first_date_row = summary_map["table"]["first_date_row"]
    day = str(day_number + (first_date_row - 1))

    for summary_sheet_name, source_column in summary_map["data_transfer"]["source_columns"].items():
        destination_sheet = summary_file[summary_sheet_name]

        for source_row, destination_column in summary_map["data_transfer"]["map"].items():
            source_cell = source_sheet[f"{source_column}{source_row}"]
            destination_cell = destination_sheet[
                f"{destination_column}{day}"]

            if source_cell.value:
                destination_cell.value = source_cell.value
            else:
                destination_cell.value = 0
