import io
from typing import Any

import matplotlib
import pandas
import zipfile

from . import xml_relationships, xml_drawing, xml_themes, xml_colours, xml_styles, xml_chart, embedded_workbook, \
    clipboard
from .themes import Theme
from .drawingml import Style, Emu, ML_LineType


def create_chart(data: pandas.DataFrame, theme: Theme = Theme(),
                 styles: list[Style] | None = None,
                 colours: list[Any] | Any | None = "tab10",
                 line_width: list[float] | float = 0.75,
                 line_type: list[ML_LineType] | ML_LineType = ML_LineType.SINGLE,
                 width: Emu = Emu(cm=23.5), height: Emu = Emu(cm=14.5)) -> None:
    if colours is None:
        colours = [matplotlib.colormaps.get_cmap("tab10")(i) for i in range(0, data.shape[1])]
    elif isinstance(colours, str):
        colours = [matplotlib.colormaps.get_cmap(colours)(i) for i in range(0, data.shape[1])]

    zip_buffer = io.BytesIO()

    contents = [
        ("[Content_Types].xml", xml_relationships.content_types()),
        ("_rels/.rels", xml_relationships.container_relationships()),
        ("clipboard/charts/style1.xml", xml_styles.chart_style()),
        ("clipboard/charts/colors1.xml", xml_colours.chart_colours()),
        ("clipboard/charts/chart1.xml", xml_chart.container_chart(data, theme, colours)),
        ("clipboard/charts/_rels/chart1.xml.rels", xml_relationships.chart_relationships()),
        ("clipboard/drawings/drawing1.xml", xml_drawing.container_drawing(width=width, height=height)),
        ("clipboard/drawings/_rels/drawing1.xml.rels", xml_relationships.drawing_relationships()),
        ("clipboard/theme/theme1.xml", xml_themes.container_theme()),
        ("clipboard/embeddings/Microsoft_Excel_Worksheet.xlsx", embedded_workbook.create(data))
    ]

    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for archive, stream in contents:
            zip_file.writestr(archive, stream, zipfile.ZIP_STORED if 'xlsx' in archive else zipfile.ZIP_DEFLATED)

    clipboard.send_officedrawing(zip_buffer)
