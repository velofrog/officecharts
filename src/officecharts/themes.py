from dataclasses import dataclass, field
from enum import Enum
from .drawingml import Font, Style, Emu, ML_LineType, ML_LineJoin, ML_LineCap, ML_PenAlignment


class LegendPosition(Enum):
    TOP = "t"
    BOTTOM = "b"
    LEFT = "l"
    RIGHT = "r"


@dataclass
class Theme:
    font: Font = field(default_factory=lambda: Font(family="Helvetica", size=12))
    line: Style = field(default_factory=lambda: Style(width=0.75, colour="#cbcbcb", line_join=ML_LineJoin.ROUND,
                                                      line_cap=ML_LineCap.FLAT, line_type=ML_LineType.SINGLE,
                                                      alignment=ML_PenAlignment.CENTER))
    title: Font = field(default_factory=lambda: Font(size=18, bold=True))
    plot_area: Style = None
    axis_labels: Font = field(default_factory=lambda: Font(size=11))
    axis_labels_x: Font = None
    axis_labels_y: Font = None
    axis_title: Font = field(default_factory=lambda: Font(size=12, bold=True))
    axis_title_x: Font = None
    axis_title_y: Font = None
    axis: Style | bool = field(default_factory=lambda: Style(colour="black"))
    axis_x: Style | bool = None
    axis_y: Style | bool = None
    axis_x_format: str = "yyyy-mm-dd"
    axis_y_format: str = "General"
    legend_position: LegendPosition = LegendPosition.BOTTOM
    legend: Font = None
    grid_major: Style = None
    grid_minor: Style = None
    grid_major_x: Style = None
    grid_minor_x: Style = None
    grid_major_y: Style = field(default_factory=lambda: Style(width=0.75, colour="#cbcbcb"))
    grid_minor_y: Style = None

    def has_axis_x(self):
        if isinstance(self.axis_x, bool):
            return self.axis_x

        if isinstance(self.axis, bool):
            return self.axis

        return not (self.axis_x is None and self.axis is None)

    def has_axis_y(self):
        if isinstance(self.axis_y, bool):
            return self.axis_y

        if isinstance(self.axis, bool):
            return self.axis

        return not (self.axis_y is None and self.axis is None)


@dataclass
class ChartProperties:
    title: str = None
    axis_x_title: str = None
    axis_y_title: str = None
    label_endpoints: str = None
