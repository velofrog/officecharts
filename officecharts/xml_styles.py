import xml.etree.ElementTree as ET
from officecharts.drawingml import xml_header, xml_append, xml_tag, xml_tagWithAttributes, ml_outline, Style, \
    ML_LineCap, ML_LineJoin, ML_PenAlignment, ML_LineType


def style_standard(include_spPr: bool = True) -> list[ET.Element]:
    result = [
        xml_tagWithAttributes("cs:lnRef", {"idx": "0"}),
        xml_tagWithAttributes("cs:fillRef", {"idx": "0"}),
        xml_tagWithAttributes("cs:effectRef", {"idx": "0"}),
        xml_tagWithAttributes("cs:fontRef", {"idx": "minor"})
    ]

    if include_spPr:
        result.extend([
            xml_tag(
                "cs:spPr",
                ml_outline(Style(0.75, line_cap=ML_LineCap.ROUND, line_join=ML_LineJoin.ROUND,
                                 line_type=ML_LineType.SINGLE, alignment=ML_PenAlignment.CENTER))
            )])

    result.extend([
        xml_tagWithAttributes("cs:defRPrRef", {"sz": "1200", "kern": "1200"})
    ])

    return result


def chart_style() -> bytes:
    root = ET.Element("cs:chartStyle", {"xmlns:cs": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
                                        "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                                        "id": "255"})

    xml_append(
        root,
        xml_tag("cs:axisTitle", *style_standard(include_spPr=False)),
        xml_tag("cs:categoryAxis", *style_standard()),
        xml_tagWithAttributes("cs:chartArea", {"mods": "allowNoFillOverride allowNoLineOverride"}, *style_standard()),
        xml_tag("cs:dataLabel", *style_standard(include_spPr=False)),
        xml_tag("cs:dataLabelCallout", *style_standard()),
        xml_tag("cs:dataPoint", *style_standard()),
        xml_tag("cs:dataPoint3D", *style_standard()),
        xml_tag("cs:dataPointLine", *style_standard()),
        xml_tag("cs:dataPointMarker", *style_standard()),
        xml_tagWithAttributes("cs:dataPointMarkerLayout", {"symbol": "circle", "size": "5"}),
        xml_tag("cs:dataPointWireframe", *style_standard()),
        xml_tag("cs:dataTable", *style_standard()),
        xml_tag("cs:downBar", *style_standard()),
        xml_tag("cs:dropLine", *style_standard()),
        xml_tag("cs:errorBar", *style_standard()),
        xml_tag("cs:floor", *style_standard()),
        xml_tag("cs:gridlineMajor", *style_standard()),
        xml_tag("cs:gridlineMinor", *style_standard()),
        xml_tag("cs:hiLoLine", *style_standard()),
        xml_tag("cs:leaderLine", *style_standard()),
        xml_tag("cs:legend", *style_standard()),
        xml_tagWithAttributes("cs:plotArea", {"mods": "allowNoFillOverride allowNoLineOverride"},
                              *style_standard(include_spPr=False)),
        xml_tagWithAttributes("cs:plotArea3D", {"mods": "allowNoFillOverride allowNoLineOverride"},
                              *style_standard(include_spPr=False)),
        xml_tag("cs:seriesAxis", *style_standard(include_spPr=False)),
        xml_tag("cs:seriesLine", *style_standard()),
        xml_tag("cs:title", *style_standard(include_spPr=False)),
        xml_tag("cs:trendline", *style_standard()),
        xml_tag("cs:trendlineLabel", *style_standard(include_spPr=False)),
        xml_tag("cs:upBar", *style_standard()),
        xml_tag("cs:valueAxis", *style_standard(include_spPr=False)),
        xml_tag("cs:wall", *style_standard())
    )

    return xml_header() + ET.tostring(root)