import xml.etree.ElementTree as ET
from .drawingml import xml_header, xml_append, ml_tagWithProperties, ml_title, ml_chartText, ml_richText, \
    ml_bodyProperties, properties, ml_listStyle, ml_textParagraph, ml_paragraphProperties, \
    ml_defaultTextRunProperties, ml_textRun, Font, ml_tag, ml_chartSeries, ml_outline, Style, \
    ml_solidFill, _format_str, _format_code
from .themes import Theme, ChartProperties

import pandas as pd
from random import randint


def container_linechart(data: pd.DataFrame, theme: Theme, styles: list[Style],
                        chart_properties: ChartProperties) -> bytes:

    root = ET.Element("c:chartSpace", {"xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
                                       "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                                       "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                                       "xmlns:c16r2": "http://schemas.microsoft.com/office/drawing/2015/06/chart"})

    xml_append(
        root,
        ml_tagWithProperties("c:date1904", {"val": "0"}),
        ml_tagWithProperties("c:lang", {"val": "en-US"}),
        ml_tagWithProperties("c:roundedCorners", {"val": "0"})
    )
    
    chart = ET.SubElement(root, "c:chart")
    
    # Chart Title
    if chart_properties.title is not None:
        chart.append(
            ml_title(
                ml_chartText(
                    ml_richText(
                        ml_bodyProperties(properties(rot="0", spcFirstLastPara="1", vertOverflow="ellipsis",
                                                     vert="horz", wrap="square", anchor="ctr", anchorCtr="1")),
                        ml_listStyle(),
                        ml_textParagraph(
                            ml_paragraphProperties(
                                ml_defaultTextRunProperties(Font.use(theme.title, theme.font))
                            ),
                            ml_textRun(_format_str(chart_properties.title))
                        )
                    )
                )
            )
        )
        chart.append(ml_tagWithProperties("c:autoTitleDeleted", properties(val="0")))
    else:
        chart.append(ml_tagWithProperties("c:autoTitleDeleted", properties(val="1")))
    
    # Chart Plot Area
    # Generate random ids for x and y axes
    x_axis_id = randint(11_111_111, 99_999_999) * 10 + 5
    y_axis_id = x_axis_id + 1
    
    plotarea = ET.SubElement(chart, "c:plotArea")
    plotarea.append(
        ml_tag(
            "c:layout",
            ml_tag(
                "c:manualLayout",
                ml_tagWithProperties("c:targetLayout", {"val": "inner"}),
                ml_tagWithProperties("c:xMode", {"val": "edge"}),
                ml_tagWithProperties("c:yMode", {"val": "factor"}),
                ml_tagWithProperties("c:x", {"val": "0.05"}),
                ml_tagWithProperties("c:y", {"val": "0"}),
                ml_tagWithProperties("c:w", {"val": "0.8"}),
                ml_tagWithProperties("c:h", {"val": "0.8"})
            )
        )
    )

    # Chart Plot Area: Line Chart Data
    plotarea.append(
        ml_tag("c:lineChart",
               ml_tagWithProperties("c:grouping", {"val": "standard"}),
               ml_tagWithProperties("c:varyColors", {"val": "0"}),
               *ml_chartSeries(data, theme, styles, chart_properties),
               ml_tag("c:dLbls",
                      ml_tagWithProperties("c:showLegendKey", {"val": "0"}),
                      ml_tagWithProperties("c:showVal", {"val": "0"}),
                      ml_tagWithProperties("c:showCatName", {"val": "0"}),
                      ml_tagWithProperties("c:showSerName", {"val": "0"}),
                      ml_tagWithProperties("c:showPercent", {"val": "0"}),
                      ml_tagWithProperties("c:showBubbleSize", {"val": "0"})
                      ),
               ml_tagWithProperties("c:smooth", {"val": "0"}),
               ml_tagWithProperties("c:axId", {"val": str(x_axis_id)}),
               ml_tagWithProperties("c:axId", {"val": str(y_axis_id)})
               )
    )
    
    # Chart Plot Area: x-axis (category axis). Place at bottom of chart
    plotarea.append(
        ml_tag(
            "c:dateAx",
            ml_tagWithProperties("c:axId", {"val": str(x_axis_id)}),
            ml_tag("c:scaling",
                   ml_tagWithProperties("c:orientation", {"val": "minMax"})
                   ),
            ml_tagWithProperties("c:delete", {"val": "0"}),
            ml_tagWithProperties("c:axPos", {"val": "b"}),
            None if theme.grid_major is None and theme.grid_major_x is None else
            ml_tag("c:majorGridlines",
                   ml_tag("c:spPr",
                          ml_outline(Style.use(theme.grid_major_x, Style.use(theme.grid_major, theme.line))),
                          ml_tag("a:effectLst")
                          )
                   ),
            None if theme.grid_minor is None and theme.grid_minor_x is None else
            ml_tag("c:minorGridlines",
                   ml_tag("c:spPr",
                          ml_outline(Style.use(theme.grid_minor_x, Style.use(theme.grid_minor, theme.line))),
                          ml_tag("a:effectLst")
                          )
                   ),
            None if chart_properties.axis_x_title is None else ml_title(
                ml_chartText(
                    ml_richText(
                        ml_bodyProperties(properties(rot="0", spcFirstLastPara="1", vertOverflow="ellipsis",
                                                     vert="horz", wrap="square", anchor="ctr", anchorCtr="1")),
                        ml_listStyle(),
                        ml_textParagraph(
                            ml_paragraphProperties(
                                ml_defaultTextRunProperties(Font.use(theme.axis_title_x,
                                                                     Font.use(theme.axis_title, theme.font)))
                            ),
                            ml_textRun(_format_str(chart_properties.axis_x_title))
                        )
                    )
                )
            ),
            ml_tagWithProperties("c:numFmt", {"formatCode": _format_code(theme.axis_x_format), "sourceLinked": "0"}),
            ml_tagWithProperties("c:majorTickMark", {"val": "out"}),
            ml_tagWithProperties("c:minorTickMark", {"val": "none"}),
            ml_tagWithProperties("c:tickLblPos", {"val": "low"}),
            ml_tag("c:spPr",
                   ml_tag("a:noFill"),
                   None if not theme.has_axis_x() else ml_outline(
                       Style.use(theme.axis_x, Style.use(theme.axis, theme.line))),
                   ml_tag("a:effectLst")
                   ),
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
                   ),
            ml_tagWithProperties("c:crossAx", {"val": str(y_axis_id)}),
            ml_tagWithProperties("c:crosses", {"val": "autoZero"}),
            ml_tagWithProperties("c:auto", {"val": "1"}),
            ml_tagWithProperties("c:lblOffset", {"val": "100"}),
            ml_tagWithProperties("c:baseTimeUnit", {"val": "days"})
        )
    )
    
    # Chart Plot Area: y-axis (value axis). Place on left side of chart
    plotarea.append(
        ml_tag(
            "c:valAx",
            ml_tagWithProperties("c:axId", {"val": str(y_axis_id)}),
            ml_tag("c:scaling",
                   ml_tagWithProperties("c:orientation", {"val": "minMax"})
                   ),
            ml_tagWithProperties("c:delete", {"val": "0"}),
            ml_tagWithProperties("c:axPos", {"val": "l"}),
            None if theme.grid_major is None and theme.grid_major_y is None else
            ml_tag("c:majorGridlines",
                   ml_tag("c:spPr",
                          ml_outline(Style.use(theme.grid_major_y, Style.use(theme.grid_major, theme.line))),
                          ml_tag("a:effectLst")
                          )
                   ),
            None if theme.grid_minor is None and theme.grid_minor_y is None else
            ml_tag("c:minorGridlines",
                   ml_tag("c:spPr",
                          ml_outline(Style.use(theme.grid_minor_y, Style.use(theme.grid_minor, theme.line))),
                          ml_tag("a:effectLst")
                          )
                   ),
            None if chart_properties.axis_y_title is None else ml_title(
                ml_chartText(
                    ml_richText(
                        ml_bodyProperties(properties(rot="-5400000", spcFirstLastPara="1", vertOverflow="ellipsis",
                                                     vert="horz", wrap="square", anchor="ctr", anchorCtr="1")),
                        ml_listStyle(),
                        ml_textParagraph(
                            ml_paragraphProperties(
                                ml_defaultTextRunProperties(Font.use(theme.axis_title_y,
                                                                     Font.use(theme.axis_title, theme.font)))
                            ),
                            ml_textRun(_format_str(chart_properties.axis_y_title))
                        )
                    )
                )
            ),
            ml_tagWithProperties("c:numFmt", {"formatCode": _format_code(theme.axis_y_format), "sourceLinked": "0"}),
            ml_tagWithProperties("c:majorTickMark", {"val": "out"}),
            ml_tagWithProperties("c:minorTickMark", {"val": "none"}),
            ml_tagWithProperties("c:tickLblPos", {"val": "nextTo"}),
            ml_tag("c:spPr",
                   ml_tag("a:noFill"),
                   None if not theme.has_axis_x() else ml_outline(
                       Style.use(theme.axis_x, Style.use(theme.axis, theme.line))),
                   ml_tag("a:effectLst")
                   ),
            ml_tag("c:txPr",
                   ml_bodyProperties(properties(rot="0", spcFirstLastPara="1", vertOverflow="ellipsis", vert="horz",
                                                wrap="square", anchor="ctr", anchorCtr="1")),
                   ml_listStyle(),
                   ml_textParagraph(
                       ml_paragraphProperties(
                           ml_defaultTextRunProperties(Font.use(theme.axis_labels_y, Font.use(theme.axis_labels,
                                                                                              theme.font)))
                       )
                   )
                   ),
            ml_tagWithProperties("c:crossAx", {"val": str(x_axis_id)}),
            ml_tagWithProperties("c:crosses", {"val": "autoZero"}),
            ml_tagWithProperties("c:crossesBetween", {"val": "between"})
        )
    )

    # Set plot area style, if defined
    if theme.plot_area is not None:
        plotarea.append(
            ml_tag(
                "c:spPr",
                None if theme.plot_area.fill_colour is None else ml_solidFill(theme.plot_area.fill_colour),
                None if theme.plot_area.colour is None else ml_outline(Style.use(theme.plot_area, theme.line))
            )
        )
    
    # Chart Legend
    if theme.legend_position is not None:
        legend = ET.SubElement(chart, "c:legend")
        xml_append(
            legend,
            ml_tagWithProperties("c:legendPos", {"val": theme.legend_position.value}),
            ml_tagWithProperties("c:overlay", {"val": "0"}),
            ml_tag("c:spPr",
                   ml_tag("a:noFill"),
                   ml_outline(Style(colour=None)),
                   ml_tag("a:effectLst")
                   ),
            ml_tag("c:txPr",
                   ml_bodyProperties(properties(rot="0", spcFirstLastPara="1", vertOverflow="ellipsis", vert="horz",
                                                wrap="square", anchor="ctr", anchorCtr="1")),
                   ml_listStyle(),
                   ml_textParagraph(
                       ml_paragraphProperties(
                           ml_defaultTextRunProperties(Font.use(theme.legend, theme.font))
                       )
                   )
                   )
        )
    
    # Finalise Chart
    xml_append(
        chart,
        ml_tagWithProperties("c:plotVisOnly", properties(val="1")),
        ml_tagWithProperties("c:dispBlanksAs", properties(val="gap")),
        ml_tag("c:extLst",
               ml_tagWithProperties("c:ext", {"uri": "{56B9EC1D-385E-4148-901F-78D8002777C0}",
                                              "xmlns:c16r3":
                                                  "http://schemas.microsoft.com/office/drawing/2017/03/chart"},
                                    ml_tag("c16r3:dataDisplayOptions16",
                                           ml_tagWithProperties("c16r3:dispNaAsBlank", {"val": "1"})
                                           )
                                    )
               )
    )
    
    # Final set of tags
    xml_append(
        root,
        ml_tagWithProperties("c:externalData", {"r:id": "rId3"},
                             ml_tagWithProperties("c:autoUpdate", {"val": "0"})
                             )
    )
    
    return xml_header() + ET.tostring(root)
