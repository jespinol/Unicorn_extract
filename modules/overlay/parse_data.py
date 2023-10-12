import os.path

import openpyxl

from modules.constants import *


def parse_files(paths):
    out = {}
    for path in paths:
        if os.path.isfile(path) and path.lower().endswith(".xlsx"):
            out[os.path.splitext(os.path.basename(path))[0]] = copy_values_from_excel(path)
        elif os.path.isdir(path):
            out = process_directory(path)
        else:
            raise f"Error: '{path}' is not a valid input file or directory"

    return out


def copy_values_from_excel(filename, header_name: str = A260_HEADER):
    try:
        workbook = openpyxl.load_workbook(filename)
        sheet = workbook.active
        headers = sheet[1]

        header_col_index = None
        for col_index, cell in enumerate(headers, 1):
            if cell.value == header_name:
                header_col_index = col_index
                break

        if header_col_index is None:
            raise ValueError(f"The header '{header_name}' was not found in the file.")

        data_dict = {}

        for row_index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
            x_value = row[header_col_index - 2]
            y_value = row[header_col_index - 1]

            data_dict[x_value] = y_value

        return data_dict
    except Exception as e:
        print(f"An error occurred: {e}")
        return {}


def process_directory(directory):
    out = {}
    for root, _, files in os.walk(directory):
        files = [file for file in files if file.lower().endswith(".xlsx") and not file.lower().endswith("compare.xlsx")]
        for file in files:
            filename = os.path.join(root, file)
            out[os.path.splitext(file)[0]] = copy_values_from_excel(filename)

    return out
