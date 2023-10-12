from decimal import *

from modules.constants import *
from modules.extract.pycorn import pc_uni6
from modules.helpers import *


def decode(input_file):
    data = pc_uni6(input_file)
    data.load()
    data.xml_parse()
    data.clean_up()
    return data


def trim_data(data):
    a260 = make_a260_dict(data[A260_UNICORN_HEADER][UNICORN_DATA])
    a280 = make_a280_dict(a260, data[A280_UNICORN_HEADER][UNICORN_DATA])

    return a260, a280


def make_a260_dict(a260_raw):
    out = {}
    total_dps = len(a260_raw)
    step = total_dps // MAX_DATAPOINTS
    for i in range(0, total_dps, step):
        x, y = a260_raw[i]
        out[x] = y
    return out


def make_a280_dict(a260_trimmed, a280_raw):
    out = {}
    for x, y in a280_raw:
        if x in a260_trimmed:
            out[x] = y

    return out


def calculate_a260a280_ratio(a260, a280):
    peak_pos = find_peak(a260)
    ratio_start = peak_pos - 1.0
    ratio_end = peak_pos + 1.5

    out = {}

    for x in a260:
        if ratio_start <= x <= ratio_end:
            ratio = float(Decimal(a260[x]) / Decimal(a280[x]))
            if 0.5 <= ratio <= 2:
                out[x] = ratio

    return out
