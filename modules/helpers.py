def find_peak(data):
    peak_x = 0.0
    peak_y = 0.0

    for x in data:
        if data[x] > peak_y:
            peak_y = data[x]
            peak_x = x

    return peak_x
