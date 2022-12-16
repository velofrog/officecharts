import xml.etree.ElementTree as ET
from .drawingml import xml_header, xml_tagWithAttributes


def chart_colours() -> bytes:
    root = ET.Element("cs:colorStyle",
                      {"xmlns:cs": "http://schemas.microsoft.com/office/drawing/2012/chartStyle",
                       "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                       "meth": "cycle", "id": "10"})
    root.append(
        xml_tagWithAttributes("a:schemeClr", {"val": "accent1"})
    )

    return xml_header() + ET.tostring(root)
