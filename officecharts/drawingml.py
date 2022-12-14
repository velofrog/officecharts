from enum import Enum
from dataclasses import dataclass
from officecharts.utilities import *
import xml.etree.ElementTree as ET
import matplotlib.colors
import pandas


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


@dataclass
class Font:
    family: str = "Helvetica"
    size: float = 12
    bold: bool = False
    italic: bool = False
    underline: ML_TextUnderline = ML_TextUnderline.NONE
    colour: object = "black"
    kerning: float = 1200
    spacing: float = 0
    baseline: float = 0
    strike: ML_TextStrikeType = ML_TextStrikeType.NONE

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


def properties(**kwargs) -> dict:
    return dict(**kwargs)


def xml_append(parent: ET.Element, *args) -> None:
    for fn in args:
        parent.append(fn)


def xml_header() -> bytes:
    return b'<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>'


def xml_tag(tag: str, *children) -> ET.Element:
    ele = ET.Element(tag)
    for child in children:
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
        ele.append(fn)
    return ele


def ml_textParagraph(*args) -> ET.Element:
    """a:p tag. Text Paragraph"""
    ele = ET.Element("a:p")
    for fn in args:
        ele.append(fn)
    return ele


def ml_textRun(text: str, *args) -> ET.Element:
    """a:r tag with embedded a:t tag. Text Run"""
    ele = ET.Element("a:r")
    ET.SubElement(ele, "a:t").text = text
    for fn in args:
        ele.append(fn)
    return ele


def ml_tag(tag: str, *args) -> ET.Element | None:
    """generic tag"""
    ele = ET.Element(tag)
    if ele is None:
        return None

    for fn in args:
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


def ml_chartSeries(data: pandas.DataFrame) -> list[ET.Element]:
    if data is None:
        return []

    series_array = []

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
                ml_outline(Style(width=2.25, line_cap=ML_LineCap.ROUND, line_join=ML_LineJoin.ROUND,
                                 colour="black")),
                ml_tag("a:effectLst")
            ),
            ml_tag(
                "c:marker",
                ml_tagWithProperties("c:symbol", {"val": "none"})
            ),
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
             cx: Emu = Emu(cm=23.5), cy: Emu = Emu(cm=14.5)) -> list[ET.Element]:
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
