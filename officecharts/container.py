import io
import pandas
import zipfile

from . import xml_relationships, xml_drawing, xml_themes, xml_colours, xml_styles, xml_chart, embedded_workbook, \
    clipboard


def create_chart(data: pandas.DataFrame) -> None:
    zip_buffer = io.BytesIO()

    contents = [
        ("[Content_Types].xml", xml_relationships.content_types()),
        ("_rels/.rels", xml_relationships.container_relationships()),
        ("clipboard/charts/style1.xml", xml_styles.chart_style()),
        ("clipboard/charts/colors1.xml", xml_colours.chart_colours()),
        ("clipboard/charts/chart1.xml", xml_chart.container_chart(data)),
        ("clipboard/charts/_rels/chart1.xml.rels", xml_relationships.chart_relationships()),
        ("clipboard/drawings/drawing1.xml", xml_drawing.container_drawing()),
        ("clipboard/drawings/_rels/drawing1.xml.rels", xml_relationships.drawing_relationships()),
        ("clipboard/theme/theme1.xml", xml_themes.container_theme()),
        ("clipboard/embeddings/Microsoft_Excel_Worksheet.xlsx", embedded_workbook.create(data))
    ]

    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for archive, stream in contents:
            zip_file.writestr(archive, stream, zipfile.ZIP_STORED if 'xlsx' in archive else zipfile.ZIP_DEFLATED)

    clipboard.send_officedrawing(zip_buffer)
