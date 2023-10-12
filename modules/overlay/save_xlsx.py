import time

import openpyxl.chart.label
import openpyxl.utils.cell
import pandas as pd

from modules.constants import *
from modules.plotting import *


def generate_xlsx(dataset):
    workbook = create_workbook()
    writer = pd.ExcelWriter(workbook, engine="openpyxl", mode="a", if_sheet_exists="overlay")

    row = 0
    col = 0
    for data in dataset:
        write_data(writer, data, dataset[data], row, col)
        col += 3

    chart_row = 2
    chart_col = chr(ord('A') + col + 3)
    worksheet = writer.sheets[WS_NAME]
    chart = create_chart(worksheet, len(dataset))
    worksheet.add_chart(chart, f"{chart_col}{chart_row}")

    writer.close()


def write_data(writer,name, data, row, col):
    df = create_dataframe(name, data)
    df.to_excel(writer, sheet_name=WS_NAME, index=False, startrow=row, startcol=col)


def create_dataframe(name, data):
    y_header = name
    out = {VOLUME_HEADER: [], y_header: []}
    for x in data:
        out[VOLUME_HEADER].append(x)
        out[y_header].append(data[x])

    df = pd.DataFrame(out)
    return df


def create_workbook():
    filename_suffix = time.strftime("%Y%m%d-%H%M%S")
    wb_filename = f"compare_{filename_suffix}.xlsx"

    workbook = openpyxl.Workbook()
    workbook["Sheet"].title = WS_NAME

    workbook.save(wb_filename)

    return wb_filename
