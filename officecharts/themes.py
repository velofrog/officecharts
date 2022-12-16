from dataclasses import dataclass, field
from enum import Enum
from .drawingml import Font, Style, Emu


class ChartPosition(Enum):
    TOP = "t"
    BOTTOM = "b"
    LEFT = "l"
    RIGHT = "r"


@dataclass
class Theme:
    font: Font = field(default_factory=lambda: Font(family="Helvetica", size=12))
    line: Style = field(default_factory=lambda: Style(width=0.75, colour="#cbcbcb"))
    title: Font = field(default_factory=lambda: Font(family=None, size=14, bold=True))
    axis_title: Font = field(default_factory=lambda: Font(family=None, size=12, bold=True))
    axis_title_x: Font = None
    axis_title_y: Font = None
    axis_x: Style = field(default_factory=lambda: Style(width=0.75, colour="black"))
    axis_y: Style = field(default_factory=lambda: Style(width=0.75, colour="black"))
    axis_format_x: str = "yyyy\\-mm\\-dd"
    axis_format_y: str = "General"
    legend_position: ChartPosition = ChartPosition.BOTTOM
    grid_major_x: Style = None
    grid_minor_x: Style = None
    grid_major_y: Style = field(default_factory=lambda: Style(width=0.75, colour="#cbcbcb"))
    grid_minor_y: Style = None



