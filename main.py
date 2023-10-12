import sys
import os

from modules.extract.parse_data import *
from modules.extract.save_xlxs import *


def main():
    args = sys.argv[1:]
    if len(args) == 0:
        print("Error: No files provided")
    else:
        for path in args:
            try:
                if os.path.isfile(path) and path.lower().endswith(".zip"):
                    process_file(path)
                elif os.path.isdir(path):
                    process_directory(path)
                else:
                    print(f"Error: '{path}' is neither a zip file nor a directory")
            except Exception as e:
                print(f"Could not process file(s). Error: '{e}'")

    print(f"Done!")


def process_file(filename):
    print(f"Processing {filename}")
    try:
        raw_data = decode(filename)
        data_a260, data_a280 = trim_data(raw_data)
        ratio = calculate_a260a280_ratio(data_a260, data_a280)
        generate_xlsx(filename, data_a260, data_a280, ratio)
    except Exception as e:
        print(f"\tCould not process file. Error: {e}")


def process_directory(directory):
    for root, _, files in os.walk(directory):
        files = [file for file in files if file.lower().endswith(".zip")]
        for file in files:
            filename = os.path.join(root, file)
            process_file(filename)


main()
