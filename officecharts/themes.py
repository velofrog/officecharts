from dataclasses import dataclass
from enum import Enum
from .drawingml import Font, Style, Emu


class ChartPosition(Enum):
    TOP = "t"
    BOTTOM = "b"
    LEFT = "l"
    RIGHT = "r"


@dataclass
class Theme:
    font: Font = Font(family="Helvetica", size=12)
    line: Style = Style(width=0.75, colour="#cbcbcb")
    title: Font = Font(family=None, size=14, bold=True)
    axis_title: Font = Font(family=None, size=12, bold=True)
    axis_title_x: Font = None
    axis_title_y: Font = None
    axis_x: Style = Style(width=0.75, colour="black")
    axis_y: Style = Style(width=0.75, colour="black")
    axis_format_x: str = "yyyy\\-mm\\-dd"
    axis_format_y: str = "General"
    legend_position: ChartPosition = ChartPosition.BOTTOM
    grid_major_x: Style = None
    grid_minor_x: Style = None
    grid_major_y: Style = Style(width=0.75, colour="#cbcbcb")
    grid_minor_y: Style = None



