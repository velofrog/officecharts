from typing import Any
from enum import Enum
from dataclasses import dataclass
from .utilities import *
import xml.etree.ElementTree as ET
import matplotlib.colors
import pandas
import re


class ML_LineCap(Enum):
    FLAT = "flat"
    ROUND = "rnd"
    SQUARE = "sq"


class ML_PenAlignment(Enum):
    CENTER = "ctr"
    INSET = "in"


class ML_LineType(Enum):
    DOUBLE = "dbl"
    SINGLE = "sng"
    THICK_THIN = "thickThin"
    THIN_THICK = "thinThick"
    TRIPLE = "tri"


class ML_LineJoin(Enum):
    BEVEL = "bevel"
    MITER = "mitre"
    ROUND = "round"


class ML_TextStrikeType(Enum):
    NONE = "noStrike"
    SINGLE = "sngStrike"
    DOUBLE = "dblStrike"


class ML_TextUnderline(Enum):
    NONE = "none"
    SINGLE = "sng"
    DOUBLE = "dbl"
    HEAVY = "heavy"
    WAVY = "wavy"
    WORDS = "words"


class ML_LegendPosition(Enum):
    TOP = "t"
    BOTTOM = "b"
    LEFT = "l"
    RIGHT = "r"
    TOP_RIGHT = "tr"


class LayoutTarget(Enum):
    INNER = "inner"
    OUTER = "outer"


class LayoutMode(Enum):
    EDGE = "edge"
    FACTOR = "factor"


@dataclass
class Layout:
    target: LayoutTarget | None = None
    x_mode: LayoutMode | None = LayoutMode.EDGE
    y_mode: LayoutMode | None = LayoutMode.FACTOR
    w_mode: LayoutMode | None = LayoutMode.FACTOR
    h_mode: LayoutMode | None = LayoutMode.FACTOR
    x: float | None = 0.05
    y: float | None = 0
    w: float | None = 0.9
    h: float | None = 0.9


@dataclass
class Emu:
    pt: float = None
    cm: float = None
    inches: float = None
    emu: float = None

    def __str__(self) -> str:
        if self.pt is not None:
            return f"{self.pt * 12_700.0:.0f}"
        elif self.cm is not None:
            return f"{self.cm * 360_000.0:.0f}"
        elif self.inches is not None:
            return f"{self.inches * 914_400:.0f}"
        elif self.emu is not None:
            return f"{self.emu:.0f}"
        else:
            return ""


def cm(val: float) -> Emu:
    return Emu(cm=val)


def inches(val: float) -> Emu:
    return Emu(inches=val)


def points(val: float) -> Emu:
    return Emu(pt=val)


@dataclass
class Font:
    family: str | None = None
    size: float = None
    bold: bool = None
    italic: bool = None
    underline: ML_TextUnderline = None
    colour: object = None
    kerning: float = None
    spacing: float = None
    baseline: float = None
    strike: ML_TextStrikeType = None

    @staticmethod
    def default() -> 'Font':
        font = Font()
        font.family = "Helvetica"
        font.size = 12
        font.bold = False
        font.italic = False
        font.underline = ML_TextUnderline.NONE
        font.colour = "black"
        font.kerning = 1200
        font.spacing = 0
        font.baseline = 0
        font.strike = ML_TextStrikeType.NONE
        return font

    @staticmethod
    def use(font: 'Font | None', parent: 'Font | None') -> 'Font | None':
        if parent is None:
            return font

        if font is None:
            return parent

        result = Font()
        result.family = parent.family if font.family is None else font.family
        result.size = parent.size if font.size is None else font.size
        result.bold = parent.bold if font.bold is None else font.bold
        result.italic = parent.italic if font.italic is None else font.italic
        result.underline = parent.underline if font.underline is None else font.underline
        result.colour = parent.colour if font.colour is None else font.colour
        result.kerning = parent.kerning if font.kerning is None else font.kerning
        result.spacing = parent.spacing if font.spacing is None else font.spacing
        result.baseline = parent.baseline if font.baseline is None else font.baseline
        result.strike = parent.strike if font.strike is None else font.strike
        return result

    def attributes(self) -> dict:
        result = {}
        if self.size is not None:
            result['sz'] = self.str_size()

        if self.bold is not None:
            result['b'] = self.str_bold()

        if self.italic is not None:
            result['i'] = self.str_italic()

        if self.underline is not None:
            result['u'] = self.str_underline()

        if self.strike is not None:
            result['strike'] = self.str_strike()

        if self.kerning is not None:
            result['kern'] = self.str_kerning()

        if self.spacing is not None:
            result['spc'] = self.str_spacing()

        if self.baseline is not None:
            result['baseline'] = self.str_baseline()

        return result

    def str_size(self):
        return f"{self.size*100:.0f}"

    def str_baseline(self):
        return f"{self.baseline:.0f}"

    def str_kerning(self):
        return "" if self.kerning is None else f"{self.kerning:.0f}"

    def str_spacing(self):
        return "" if self.spacing is None else f"{self.spacing:.0f}"

    def str_bold(self):
        return "1" if self.bold else "0"

    def str_italic(self):
        return "1" if self.italic else "0"

    def str_underline(self):
        return self.underline.value

    def str_strike(self):
        return self.strike.value


@dataclass
class Style:
    width: float = None
    colour: object = None
    fill_colour: object = None
    line_cap: ML_LineCap = None
    line_type: ML_LineType = None
    line_join: ML_LineJoin = None
    alignment: ML_PenAlignment = None

    @staticmethod
    def use(style: 'Style | None', parent: 'Style | None') -> 'Style | None':
        if parent is None:
            return style

        if style is None:
            return parent

        result = Style()
        result.width = parent.width if style.width is None else style.width
        result.colour = parent.colour if style.colour is None else style.colour
        result.fill_colour = parent.fill_colour if style.fill_colour is None else style.fill_colour
        result.line_cap = parent.line_cap if style.line_cap is None else style.line_cap
        result.line_type = parent.line_type if style.line_type is None else style.line_type
        result.line_join = parent.line_join if style.line_join is None else style.line_join
        result.alignment = parent.alignment if style.alignment is None else style.alignment
        return result


def properties(**kwargs) -> dict:
    return dict(**kwargs)


def _format_code(fmt_code: str = "General"):
    """Escape special characters in format code"""
    fmt_code = re.sub(r"([-\/])", r"\\\1", fmt_code)
    fmt_code = re.sub(r"\"", r"&quot;", fmt_code)
    return fmt_code


def _format_str(s: Any):
    """Escape special characters for in general strings"""
    if s is None:
        s = ""

    if not isinstance(s, str):
        s = str(s)

    # s = re.sub(r"\"", r"&quot;", s)
    s = re.sub("&", "&amp;", s)
    s = re.sub("<", "&lt;", s)
    s = re.sub(">", "&gt;", s)

    return s


def xml_append(parent: ET.Element, *args) -> None:
    for fn in args:
        if fn is not None:
            parent.append(fn)


def xml_header() -> bytes:
    return b'<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>'


def xml_tag(tag: str, *children) -> ET.Element:
    ele = ET.Element(tag)
    for child in children:
        if child is not None:
            ele.append(child)

    return ele


def xml_tagWithAttributes(tag: str, attrib: dict = {}, *children) -> ET.Element:
    ele = xml_tag(tag, *children)

    if isinstance(attrib, dict):
        for key, value in attrib.items():
            ele.set(key, str(value))

    return ele

def ml_solidFill(colour, default_colour="black") -> ET.Element:
    if colour is None:
        return ET.Element("a:noFill")

    if not matplotlib.colors.is_color_like(colour):
        colour = default_colour

    rgb = matplotlib.colors.to_hex(colour, False).upper()[1:]
    alpha = f"{matplotlib.colors.to_rgba(colour)[3]*100_000:.0f}"

    ele = ET.Element("a:solidFill")
    ET.SubElement(ET.SubElement(ele, "a:srgbClr", val=rgb), "a:alpha", val=alpha)
    return ele


def ml_numCache(ele, values: list[object], format="General") -> ET.Element:
    ele_cache = ET.SubElement(ele, "c:numCache")
    ET.SubElement(ele_cache, "c:formatCode").text = format
    ET.SubElement(ele_cache, "c:ptCount", val=str(len(values)))
    for idx, v in enumerate(values):
        ET.SubElement(ET.SubElement(ele_cache, "c:pt", idx=str(idx)), "c:v").text = str(v)
    return ele


def ml_defaultTextRunProperties(font: Font = None) -> ET.Element:
    """a:defRPr tag. Default Text Paragraph Properties. Embeds font properties if supplied"""
    if font is None:
        return ET.Element("a:defRPr")

    ele = ET.Element("a:defRPr", font.attributes())
    if font.colour is not None:
        ele.append(ml_solidFill(font.colour))
    if font.family is not None:
        ET.SubElement(ele, "a:latin", typeface=font.family)
    return ele


def ml_paragraphProperties(*args) -> ET.Element:
    """a:pPr tag. Text Paragraph Properties"""
    ele = ET.Element("a:pPr")
    for fn in args:
        if fn is not None:
            ele.append(fn)
    return ele


def ml_textParagraph(*args) -> ET.Element:
    """a:p tag. Text Paragraph"""
    ele = ET.Element("a:p")
    for fn in args:
        if fn is not None:
            ele.append(fn)
    return ele


def ml_textRun(text: str, *args) -> ET.Element:
    """a:r tag with embedded a:t tag. Text Run"""
    ele = ET.Element("a:r")
    ET.SubElement(ele, "a:t").text = text
    for fn in args:
        if fn is not None:
            ele.append(fn)
    return ele


def ml_tag(tag: str, *args) -> ET.Element | None:
    """generic tag"""
    ele = ET.Element(tag)
    if ele is None:
        return None

    for fn in args:
        if fn is not None:
            ele.append(fn)
    return ele


def ml_tagWithProperties(tag: str, tag_properties: dict = None, *args) -> ET.Element | None:
    ele = ml_tag(tag, *args)
    if ele is None:
        return None

    if tag_properties is None:
        return None

    for key, value in tag_properties.items():
        ele.set(key, str(value))

    return ele


def ml_richText(*args) -> ET.Element:
    """generic tag"""
    return ml_tag("c:rich", *args)


def ml_title(*args) -> ET.Element:
    """c:title tag"""
    return ml_tag("c:title", *args)


def ml_listStyle(*args) -> ET.Element:
    """a:lstStyle tag"""
    return ml_tag("a:lstStyle", *args)


def ml_bodyProperties(body_properties: dict = None, *args) -> ET.Element:
    """a:bodyPr tag. Body properties"""
    ele = ml_tag("a:bodyPr", *args)
    if body_properties is not None:
        for key, value in body_properties.items():
            ele.set(key, str(value))

    return ele


def ml_shapeProperties(*args) -> ET.Element:
    """c:spPr tag. Shape properties"""
    return ml_tag("c:spPr", *args)


def ml_outline(style: Style, *args) -> ET.Element:
    """a:ln tag. Outline"""
    ele = ml_tag("a:ln", *args)
    if style is None:
        for fn in args:
            if fn is not None:
                ele.append(fn)
        return ele

    if style.width is not None:
        ele.set("w", str(Emu(pt=style.width)))

    if style.line_cap is not None:
        ele.set("cap", style.line_cap.value)

    if style.line_type is not None:
        ele.set("cmpd", style.line_type.value)

    if style.alignment is not None:
        ele.set("algn", style.alignment.value)

    if style.colour is None:
        ele.append(ml_tag("a:noFill"))
    else:
        ele.append(ml_solidFill(style.colour))

    if style.line_join is not None:
        ele.append(ml_tag("a:" + style.line_join.value))

    for fn in args:
        if fn is not None:
            ele.append(fn)

    return ele


def ml_chartText(*args) -> ET.Element:
    """c:tx tag. Chart Text"""
    return ml_tag("c:tx", *args)


def ml_stringReference(*args) -> ET.Element:
    """c:strRef tag. String Reference"""
    return ml_tag("c:strRef", *args)


def ml_numberReference(*args) -> ET.Element:
    """c:numRef tag. Number Reference"""
    return ml_tag("c:numRef", *args)


def ml_excelColumnAsString(colIndex: int) -> str:
    label = ''
    while colIndex >= 0:
        label = chr(ord('A') + (colIndex % 26)) + label
        colIndex = (colIndex // 26) - 1

    return label


def ml_formulaExcelAddress(sheet: str = "Sheet1", start_row: int = 0, start_col: int = 0,
                           end_row: int = None, end_col: int = None, absolute: bool = True) -> ET.Element:
    """c:f tag. Creates Excel A1 notation style reference"""
    ele = ml_tag("c:f")

    if not sheet.isalnum():
        ref = f"'{sheet}'!"
    else:
        ref = sheet + "!"

    if end_row is not None and end_col is None:
        end_col = start_col

    if end_col is not None and end_row is None:
        end_row = start_row

    ref = ref + ("$" if absolute else "") + ml_excelColumnAsString(start_col) + ("$" if absolute else "") + str(
        start_row + 1)
    if end_row is not None and end_col is not None:
        ref = ref + ":" + ("$" if absolute else "") + ml_excelColumnAsString(end_col) + ("$" if absolute else "") + str(
            end_row + 1)

    ele.text = ref
    return ele


def ml_value(value: any) -> ET.Element:
    ele = ET.Element("c:v")
    ele.text = str(value)
    return ele


def ml_stringCache(array: list[str]) -> ET.Element:
    ele = ET.Element("c:strCache")
    if array is None:
        ele.append(ml_tagWithProperties("c:ptCount", {"val": "0"}))
        return ele

    if not is_iterable(array):
        array = [array]

    ele.append(ml_tagWithProperties("c:ptCount", {"val": str(len(array))}))
    for idx, value in enumerate(array):
        ele.append(
            ml_tagWithProperties("c:pt", {"idx": str(idx)}, ml_value(value))
        )

    return ele


def ml_formatCode(format_code: str) -> ET.Element:
    ele = ml_tag("c:formatCode")
    if format_code is not None:
        ele.text = format_code
    return ele


def ml_numberCache(array: list[object], format_code: str = None) -> ET.Element:
    ele = ET.Element("c:numCache")
    if array is None:
        ele.append(ml_tagWithProperties("c:ptCount", {"val": "0"}))
        return ele

    if not is_iterable(array):
        array = [array]

    if is_datetime(array):
        array = excel_datetime(as_datetime(array))
        if format_code is None:
            format_code = "yyyy-mm-dd"

    if format_code is not None:
        ele.append(ml_formatCode(format_code))

    ele.append(ml_tagWithProperties("c:ptCount", {"val": str(len(array))}))
    for idx, value in enumerate(array):
        ele.append(
            ml_tagWithProperties("c:pt", {"idx": str(idx)}, ml_value(value))
        )

    return ele


def ml_dataLabel_last(data: pandas.DataFrame, theme: 'Theme') -> ET.Element:
    ele = xml_tag("c:dLbls")

    xml_append(
        ele,
        xml_tag("c:dLbl",
                xml_tagWithAttributes("c:idx", {"val": str(data.shape[0]-1)}),
                xml_tagWithAttributes("c:showLegendKey", {"val": "1"}),
                xml_tagWithAttributes("c:showVal", {"val": "1"}),
                xml_tagWithAttributes("c:showCatName", {"val": "0"}),
                xml_tagWithAttributes("c:showSerName", {"val": "1"}),
                xml_tagWithAttributes("c:showPercent", {"val": "0"}),
                xml_tagWithAttributes("c:showBubbleSize", {"val": "0"}),
                xml_tag("c:extLst",
                        xml_tagWithAttributes(
                            "c:ext", {"uri": "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}",
                                      "xmlns:c15": "http://schemas.microsoft.com/office/drawing/2012/chart"})
                        )
                )
    )

    xml_append(
        ele,
        xml_tagWithAttributes("c:numFmt", {"formatCode": _format_code(theme.axis_y_format), "sourceLinked": "0"}),
        xml_tag("c:spPr",
                xml_tag("c:noFill"),
                xml_tag("a:ln",
                        xml_tag("a:noFill")),
                xml_tag("a:effectLst")
                )
    )

    xml_append(
        ele,
        ml_tag("c:txPr",
               ml_bodyProperties(properties(rot="0", spcFirstLastPara="1", vertOverflow="ellipsis", vert="horz",
                                            wrap="square", anchor="ctr", anchorCtr="1")),
               ml_listStyle(),
               ml_textParagraph(
                   ml_paragraphProperties(
                       ml_defaultTextRunProperties(Font.use(theme.axis_labels_x, Font.use(theme.axis_labels,
                                                                                          theme.font)))
                   )
               )
               )
    )

    return ele


def ml_chartSeries(data: pandas.DataFrame, theme: 'Theme', styles: list[Style],
                   chart_properties: 'ChartProperties') -> list[ET.Element]:
    if data is None:
        return []

    series_array = []
    if styles is None:
        styles = [Style(width=2.25, colour="black", line_type=ML_LineType.SINGLE,
                        line_cap=ML_LineCap.ROUND, line_join=ML_LineJoin.ROUND)] * data.shape[1]

    for idx in range(0, data.shape[1]):
        series = ET.Element('c:ser')
        xml_append(
            series,
            ml_tagWithProperties("c:idx", {"val": str(idx)}),
            ml_tagWithProperties("c:order", {"val": str(idx)}),
            # Series name
            ml_chartText(
                ml_stringReference(
                    ml_formulaExcelAddress('Sheet1', 0, idx + 1),
                    ml_stringCache(data.columns[idx])
                )
            ),
            # Series style
            ml_shapeProperties(
                ml_outline(styles[idx]),
                ml_tag("a:effectLst")
            ),
            ml_tag(
                "c:marker",
                ml_tagWithProperties("c:symbol", {"val": "none"})
            ),
            ml_dataLabel_last(data, theme) if chart_properties.label_endpoints else None,
            # Series category values (x-axis)
            ml_tag(
                "c:cat",
                ml_numberReference(
                    ml_formulaExcelAddress('Sheet1', 1, 0, data.shape[0]),
                    ml_numberCache(data.index)
                )
            ),
            # Series values (y-axis)
            ml_tag(
                "c:val",
                ml_numberReference(
                    ml_formulaExcelAddress('Sheet1', 1, idx + 1, data.shape[0]),
                    ml_numberCache(data.iloc[:, idx])
                )
            ),
            ml_tagWithProperties("c:smooth", {"val": "0"})
        )
        series_array.extend([series])

    return series_array


def ml_group(identifier: int = 0, name: str = "", x: Emu = Emu(0), y: Emu = Emu(0),
             cx: Emu = Emu(cm=10), cy: Emu = Emu(cm=10)) -> list[ET.Element]:
    return [
        xml_tag(
            "a:nvGrpSpPr",
            xml_tagWithAttributes("a:cNvPr", {"id": identifier, "name": name}),
            xml_tag("a:cNvGrpSpPr")
        ),
        xml_tag(
            "a:grpSpPr",
            xml_tag(
                "a:xfrm",
                xml_tagWithAttributes("a:off", {"x": "0", "y": "0"}),
                xml_tagWithAttributes("a:ext", {"cx": str(cx), "cy": str(cy)}),
                xml_tagWithAttributes("a:chOff", {"x": str(x), "y": str(y)}),
                xml_tagWithAttributes("a:chExt", {"cx": str(cx), "cy": str(cy)})
            )
        )
    ]


def ml_layout(layout: Layout | None) -> ET.Element:
    if layout is None or layout.target is None:
        return ml_tag("c:layout")

    return ml_tag(
        "c:layout",
        ml_tag(
            "c:manualLayout",
            ml_tagWithProperties("c:layoutTarget", {"val": layout.target.value}),
            ml_tagWithProperties("c:xMode", {"val": layout.x_mode.value}),
            ml_tagWithProperties("c:yMode", {"val": layout.y_mode.value}),
            ml_tagWithProperties("c:wMode", {"val": layout.w_mode.value}),
            ml_tagWithProperties("c:hMode", {"val": layout.h_mode.value}),
            ml_tagWithProperties("c:x", {"val": str(layout.x)}),
            ml_tagWithProperties("c:y", {"val": str(layout.y)}),
            ml_tagWithProperties("c:w", {"val": str(layout.w)}),
            ml_tagWithProperties("c:h", {"val": str(layout.h)})
        )
    )
