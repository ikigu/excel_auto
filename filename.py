import os
from datetime import datetime


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


def create_new_file_path(old_file_path):
    """
    Creates the a path with the new file name according to today's date

    Args:
        old_file_path (str): The path to the existing .xlsx file.

    Returns:
        str: The new file path if successful.

    Raises:
        FileNotFoundError: If the old file path does not exist.
        ValueError: If the new file name does not have the .xlsx extension.
    """
    if not os.path.isfile(old_file_path):
        raise FileNotFoundError(f"The file '{old_file_path}' does not exist.")

    # Extract the directory of the old file
    directory = os.path.dirname(old_file_path)

    # Get today's date
    date = datetime.now()

    # Create the new file name with today's date
    new_file_name = f"{get_day_with_suffix(date.day)} {date.strftime('%B %Y')}.xlsx"

    # Construct the new file path
    new_file_path = os.path.join(directory, new_file_name)

    return new_file_path
