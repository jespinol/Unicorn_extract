import sys

from modules.parse_data import parse_files
from modules.save_xlsx import generate_xlsx


def main():
    args = sys.argv[1:]
    if len(args) == 0:
        print("Error: No files provided")
    else:
        data = parse_files(args)
        generate_xlsx(data)


main()
