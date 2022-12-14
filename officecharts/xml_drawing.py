import xml.etree.ElementTree as ET
from officecharts.drawingml import Emu, xml_header, xml_append, xml_tag, xml_tagWithAttributes, ml_group


def container_drawing(x: Emu = Emu(0), y: Emu = Emu(0), width: Emu = Emu(cm=23.5), height: Emu = Emu(cm=14.5)) -> bytes:
    root = ET.Element("a:graphic", {"xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
    xml_append(
        root,
        xml_tagWithAttributes(
            "a:graphicData", {"uri": "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"},
            xml_tagWithAttributes(
                "lc:lockedCanvas", {"xmlns:lc": "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"},
                *ml_group(0, "canvas", x, y, width, height),
                xml_tag(
                    "a:graphicFrame",
                    xml_tag(
                        "a:nvGraphicFramePr",
                        xml_tagWithAttributes("a:cNvPr", {"id": 1, "name": "Chart"}),
                        xml_tag("a:cNvGraphicFramePr")
                    ),
                    xml_tag(
                        "a:graphic",
                        xml_tagWithAttributes(
                            "a:graphicData", {"uri": "http://schemas.openxmlformats.org/drawingml/2006/chart"},
                            xml_tagWithAttributes(
                                "c:chart", {
                                    "xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                                    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                                    "r:id": "rId1"
                                }
                            )
                        )
                    ),
                    xml_tag(
                        "a:xfrm",
                        xml_tagWithAttributes("a:off", {"x": str(x), "y": str(y)}),
                        xml_tagWithAttributes("a:ext", {"cx": str(width), "cy": str(height)})
                    )
                )
            )
        )
    )

    return xml_header() + ET.tostring(root)
