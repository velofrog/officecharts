import io
from typing import Any

import matplotlib
import pandas
import zipfile

from . import xml_relationships, xml_drawing, xml_themes, xml_colours, xml_styles, xml_linechart, embedded_workbook, \
    clipboard
from .themes import Theme, ChartProperties
from .drawingml import Style, Emu, ML_LineType, ML_LineCap, ML_LineJoin


def create_linechart(data: pandas.DataFrame, theme: Theme = Theme(),
                     title: str = None,
                     axis_x_title: str = None,
                     axis_y_title: str = None,
                     label_endpoints: bool = False,
                     styles: list[Style] | None = None,
                     colours: list[Any] | Any | None = "tab10",
                     line_width: list[float] | float = 2.25,
                     line_type: list[ML_LineType] | ML_LineType = ML_LineType.SINGLE,
                     line_cap: list[ML_LineCap] | ML_LineType = ML_LineCap.ROUND,
                     line_join: list[ML_LineJoin] | ML_LineJoin = ML_LineJoin.ROUND,
                     width: Emu = Emu(cm=23.5), height: Emu = Emu(cm=14.5)) -> None:

    if data is None or data.shape[1] == 0:
        raise ValueError("No data supplied")

    if not isinstance(theme, Theme):
        theme = Theme()

    if colours is None:
        colours = [matplotlib.colormaps.get_cmap("tab10")(i) for i in range(0, data.shape[1])]
    elif isinstance(colours, str):
        colours = [matplotlib.colormaps.get_cmap(colours)(i) for i in range(0, data.shape[1])]

    def _aslist(obj: Any, length: int) -> list[Any]:
        """Returns a list of length length, recycling if necessary"""
        if not isinstance(obj, (list, tuple)):
            obj = [obj]
        return [obj[i % len(obj)] for i in range(0, length)]

    if styles is None:
        styles = [Style() for _ in range(data.shape[1])]
        for i, v in enumerate(_aslist(colours, len(styles))):
            styles[i].colour = v
        for i, v in enumerate(_aslist(line_width, len(styles))):
            styles[i].width = v
        for i, v in enumerate(_aslist(line_type, len(styles))):
            styles[i].line_type = v
        for i, v in enumerate(_aslist(line_cap, len(styles))):
            styles[i].line_cap = v
        for i, v in enumerate(_aslist(line_join, len(styles))):
            styles[i].line_join = v
    else:
        if isinstance(styles, list):
            if len(styles) < data.shape[1]:
                styles = _aslist(styles, data.shape[1])
        else:
            styles = [styles] * data.shape[1]

    properties = ChartProperties(title=title, axis_x_title=axis_x_title, axis_y_title=axis_y_title,
                                 label_endpoints=label_endpoints)

    zip_buffer = io.BytesIO()

    contents = [
        ("[Content_Types].xml", xml_relationships.content_types()),
        ("_rels/.rels", xml_relationships.container_relationships()),
        ("clipboard/charts/style1.xml", xml_styles.chart_style()),
        ("clipboard/charts/colors1.xml", xml_colours.chart_colours()),
        ("clipboard/charts/chart1.xml", xml_linechart.container_linechart(data, theme, styles, properties)),
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
    # with open("/Users/michael/clip4/last_output.zip", "wb") as f:
    #     f.write(zip_buffer.getvalue())

