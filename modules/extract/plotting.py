from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.series_factory import SeriesFactory
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import Font, CharacterProperties, Paragraph, ParagraphProperties


def create_chart(ws, a260a280_dps, ratio_dps, peak_pos, reduced_window):
    # initialize a chart object, set chart dimensions, set font and font size
    chart = ScatterChart()
    chart.height = 20.96
    chart.width = 31.44
    chart.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    font = Font(typeface="Calibri")
    axis_font = CharacterProperties(latin=font, sz=2000)
    axis_label_font = CharacterProperties(latin=font, sz=1800)
    legend_font = CharacterProperties(latin=font, sz=1800)

    # legend properties
    chart.legend.overlay = True
    chart.legend.txPr = RichText(
        p=[Paragraph(pPr=ParagraphProperties(defRPr=legend_font), endParaRPr=legend_font)])

    # y-axis properties
    chart.y_axis.majorGridlines = None
    chart.y_axis.majorTickMark = "in"
    chart.y_axis.title = "mAU"
    chart.y_axis.title.tx.rich.p[0].pPr.defRPr = axis_font
    chart.y_axis.scaling.min = 0.0
    chart.y_axis.txPr = RichText(
        p=[Paragraph(pPr=ParagraphProperties(defRPr=axis_label_font), endParaRPr=axis_label_font)])
    chart.y_axis.number_format = "#0"

    # x-axis properties
    chart.x_axis.majorGridlines = None
    chart.x_axis.majorTickMark = "in"
    chart.x_axis.title = "Volume (mL)"
    chart.x_axis.title.tx.rich.p[0].pPr.defRPr = axis_font
    chart.x_axis.txPr = RichText(
        p=[Paragraph(pPr=ParagraphProperties(defRPr=axis_label_font), endParaRPr=axis_label_font)])
    chart.x_axis.number_format = "#0"
    if reduced_window:
        chart.x_axis.scaling.min = peak_pos - 1.5
        chart.x_axis.scaling.max = peak_pos + 2.5
        chart.x_axis.majorUnit = 1.0
    else:
        offset = 5.0 if peak_pos - 5.0 > 0 else peak_pos
        chart.x_axis.scaling.min = peak_pos - offset
        chart.x_axis.scaling.max = peak_pos + offset
        chart.x_axis.majorUnit = 2.0

    # plot A260
    x_values = Reference(ws, min_col=1, min_row=2, max_row=a260a280_dps + 1)
    y_values = Reference(ws, min_col=2, min_row=2, max_row=a260a280_dps + 1)
    series = SeriesFactory(y_values, x_values, title=" A260")
    series.graphicalProperties.line.solidFill = "D22B2B"
    series.graphicalProperties.line.width = 60000
    chart.series.append(series)

    # plot A280
    x_values = Reference(ws, min_col=4, min_row=2, max_row=a260a280_dps + 1)
    y_values = Reference(ws, min_col=5, min_row=2, max_row=a260a280_dps + 1)
    series = SeriesFactory(y_values, x_values, title=" A280")
    series.graphicalProperties.line.solidFill = "0047AB"
    series.graphicalProperties.line.width = 60000
    chart.series.append(series)

    # chart with A260/A280 plot
    ratio_chart = create_ratio_chart(ws, ratio_dps)

    chart += ratio_chart
    return chart


def create_ratio_chart(ws, ratio_dps):
    # initialize a chart object for a260/a280, set axis properties
    ratio_chart = ScatterChart()
    ratio_chart.y_axis.majorGridlines = None
    ratio_chart.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    ratio_chart_font = CharacterProperties(sz=100, b=False, solidFill="FF0000")
    ratio_chart.y_axis.txPr = RichText(
        p=[Paragraph(pPr=ParagraphProperties(defRPr=ratio_chart_font), endParaRPr=ratio_chart_font)])
    ratio_chart.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    ratio_chart.y_axis.axId = 200
    ratio_chart.y_axis.crosses = "max"
    ratio_chart.y_axis.scaling.min = -2.0
    ratio_chart.y_axis.scaling.max = 5.0

    # plot A260/A280
    x_values = Reference(ws, min_col=7, min_row=2, max_row=ratio_dps + 1)
    y_values = Reference(ws, min_col=8, min_row=2, max_row=ratio_dps + 1)
    series = Series(y_values, x_values, title=" A260/A280")
    series.graphicalProperties.line.solidFill = "BF40BF"
    series.graphicalProperties.line.width = 60000
    ratio_chart.series.append(series)

    return ratio_chart
