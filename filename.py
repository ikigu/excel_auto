import calendar
import os
from datetime import datetime
from pathlib import Path
from datetime import datetime, timedelta


def get_day_with_suffix(day):
    """
    Formats a date with extensions: Xst, Xnd, Xrd or Xth


    Args:
        day (day object): The current date
    """
    if 10 <= day % 100 <= 20:  # Handle special cases for 11th, 12th, 13th, etc.
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{day}{suffix}"


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


def create_new_file_path(original_filename):
    month_directory = os.path.dirname(original_filename)
    year_directory = Path(month_directory).parent.absolute()
    shift_change_directory = year_directory.parent.absolute()

    file_year = get_year_from_file_name(original_filename)
    file_month = get_month_from_file_name(original_filename).capitalize()
    file_day = get_day_from_file_name(original_filename)

    current_file_date = datetime(file_year, list(
        calendar.month_name).index(file_month), file_day)

    new_file_date = current_file_date + timedelta(days=1)

    new_file_year = new_file_date.year
    new_file_month = new_file_date.strftime("%B").upper()
    new_file_day = new_file_date.day
    day_with_suffix = get_day_with_suffix(new_file_day)

    new_file_name = f"{day_with_suffix} {new_file_month} {new_file_year}.xlsx"

    new_path = os.path.join(shift_change_directory,
                            f"YEAR {new_file_year}", f"{new_file_month} {new_file_year}")

    os.makedirs(new_path, exist_ok=True)

    new_file_path = os.path.join(new_path, new_file_name)

    return new_file_path
