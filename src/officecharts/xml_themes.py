import xml.etree.ElementTree as ET
import matplotlib
from .drawingml import xml_header, xml_tag, xml_tagWithAttributes


def container_theme() -> bytes:
    """Minimum set of theme elements required for embedding into Word"""
    colours = [matplotlib.colors.to_hex(colour, False).upper()[1:] for colour in
               [matplotlib.colormaps.get_cmap("tab10")(i) for i in range(0, 6)]]
    colour_tags = [xml_tag(f"a:accent{idx+1}", xml_tagWithAttributes("a:srgbClr", {"val": hex_colour})) for
                   idx, hex_colour in enumerate(colours)]
    clr_scheme = [
        xml_tag("a:dk1", xml_tagWithAttributes("a:sysClr", {"val": "windowText", "lastClr": "000000"})),
        xml_tag("a:lt1", xml_tagWithAttributes("a:sysClr", {"val": "window", "lastClr": "FFFFFF"})),
        xml_tag("a:dk2", xml_tagWithAttributes("a:srgbClr", {"val": "44546A"})),
        xml_tag("a:lt2", xml_tagWithAttributes("a:srgbClr", {"val": "E7E6E6"})),
        *colour_tags,
        xml_tag("a:hlink", xml_tagWithAttributes("a:srgbClr", {"val": "0563C1"})),
        xml_tag("a:folHlink", xml_tagWithAttributes("a:srgbClr", {"val": "954F72"}))
    ]

    root = xml_tagWithAttributes(
        "a:clipboardTheme", {"xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main"},
        xml_tag(
            "a:themeElements",
            xml_tagWithAttributes(
                "a:clrScheme", {"name": "Custom"},
                *clr_scheme
            ),
            xml_tagWithAttributes(
                "a:fontScheme", {"name": "Custom"},
                xml_tag(
                    "a:majorFont",
                    xml_tagWithAttributes("a:latin", {"typeface": "Calibri"}),
                    xml_tagWithAttributes("a:ea", {"typeface": ""}),
                    xml_tagWithAttributes("a:cs", {"typeface": ""})
                ),
                xml_tag(
                    "a:minorFont",
                    xml_tagWithAttributes("a:latin", {"typeface": "Calibri"}),
                    xml_tagWithAttributes("a:ea", {"typeface": ""}),
                    xml_tagWithAttributes("a:cs", {"typeface": ""})
                )
            ),
            xml_tagWithAttributes(
                "a:fmtScheme", {"name": "Custom"},
                xml_tag(
                    "a:fillStyleLst",
                    xml_tag("a:solidFill", xml_tagWithAttributes("a:schemeClr", {"val": "accent1"})),
                    xml_tag("a:solidFill", xml_tagWithAttributes("a:schemeClr", {"val": "accent2"})),
                    xml_tag("a:solidFill", xml_tagWithAttributes("a:schemeClr", {"val": "accent3"}))
                ),
                xml_tag(
                    "a:lnStyleLst",
                    xml_tagWithAttributes("a:ln", {"w": "6350", "cap": "flat", "cmpd": "sng", "algn": "ctr"},
                                          xml_tag("a:solidFill",
                                                  xml_tagWithAttributes("a:schemeClr", {"val": "accent1"}))),
                    xml_tagWithAttributes("a:ln", {"w": "12700", "cap": "flat", "cmpd": "sng", "algn": "ctr"},
                                          xml_tag("a:solidFill",
                                                  xml_tagWithAttributes("a:schemeClr", {"val": "accent2"}))),
                    xml_tagWithAttributes("a:ln", {"w": "19050", "cap": "flat", "cmpd": "sng", "algn": "ctr"},
                                          xml_tag("a:solidFill",
                                                  xml_tagWithAttributes("a:schemeClr", {"val": "accent3"})))
                ),
                xml_tag(
                    "a:effectStyleLst",
                    xml_tag("a:effectStyle", xml_tag("a:effectLst")),
                    xml_tag("a:effectStyle", xml_tag("a:effectLst")),
                    xml_tag("a:effectStyle", xml_tag("a:effectLst")),
                ),
                xml_tag(
                    "a:bgFillStyleLst",
                    xml_tag("a:solidFill", xml_tagWithAttributes("a:schemeClr", {"val": "accent1"})),
                    xml_tag("a:solidFill", xml_tagWithAttributes("a:schemeClr", {"val": "accent2"})),
                    xml_tag("a:solidFill", xml_tagWithAttributes("a:schemeClr", {"val": "accent3"}))
                )
            )
        ),
        xml_tagWithAttributes(
            "a:clrMap", {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2",
                         "accent1": "accent1", "accent2": "accent2", "accent3": "accent3", "accent4": "accent4",
                         "accent5": "accent5", "accent6": "accent6", "hlink": "hlink", "folHlink": "folHlink"}
        )
    )

    return xml_header() + ET.tostring(root)
