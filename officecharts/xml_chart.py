import xml.etree.ElementTree as ET
from officecharts.drawingml import xml_header, xml_append, ml_tagWithProperties, ml_title, ml_chartText, ml_richText, \
    ml_bodyProperties, properties, ml_listStyle, ml_textParagraph, ml_paragraphProperties, \
    ml_defaultTextRunProperties, ml_textRun, Font, ml_tag, ml_chartSeries, ml_outline, Style, ML_LineCap, ML_LineType, \
    ML_PenAlignment, ML_LineJoin
import pandas as pd
from random import randint


def container_chart(data: pd.DataFrame) -> bytes:

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
    chart.append(
        ml_title(
            ml_chartText(
                ml_richText(
                    ml_bodyProperties(properties(rot="0", spcFirstLastPara="1", vertOverflow="ellipsis", vert="horz",
                                                 wrap="square", anchor="ctr", anchorCtr="1")),
                    ml_listStyle(),
                    ml_textParagraph(
                        ml_paragraphProperties(
                            ml_defaultTextRunProperties(Font())
                        ),
                        ml_textRun("My Chart Title")
                    )
                )
            )
        )
    )
    chart.append(ml_tagWithProperties("c:autoTitleDeleted", properties(val="0")))
    
    # Chart Plot Area
    x_axis_id = randint(11_111_111, 99_999_999) * 10 + 5
    y_axis_id = x_axis_id + 1
    
    plotarea = ET.SubElement(chart, "c:plotArea")
    plotarea.append(ml_tag("c:layout"))
    
    # Chart Plot Area: Line Chart Data
    plotarea.append(
        ml_tag("c:lineChart",
               ml_tagWithProperties("c:grouping", {"val": "standard"}),
               ml_tagWithProperties("c:varyColors", {"val": "0"}),
               *ml_chartSeries(data),
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
            # majorGridLines
            # minorGridLines
            ml_tagWithProperties("c:numFmt", {"formatCode": "yyyy\\-mm\\-dd", "sourceLinked": "0"}),
            ml_tagWithProperties("c:majorTickMark", {"val": "out"}),
            ml_tagWithProperties("c:minorTickMark", {"val": "none"}),
            ml_tagWithProperties("c:tickLblPos", {"val": "low"}),
            ml_tag("c:spPr",
                   ml_tag("a:noFill"),
                   ml_outline(Style(width=0.75, line_cap=ML_LineCap.FLAT, line_type=ML_LineType.SINGLE,
                                    line_join=ML_LineJoin.ROUND, alignment=ML_PenAlignment.CENTER, colour="black")),
                   ml_tag("a:effectLst")
                   ),
            ml_tag("c:txPr",
                   ml_bodyProperties(properties(rot="0", spcFirstLastPara="1", vertOverflow="ellipsis", vert="horz",
                                                wrap="square", anchor="ctr", anchorCtr="1")),
                   ml_listStyle(),
                   ml_textParagraph(
                       ml_paragraphProperties(
                           ml_defaultTextRunProperties(Font())
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
            # majorGridLines
            # minorGridLines
            ml_tagWithProperties("c:numFmt", {"formatCode": "General", "sourceLinked": "1"}),
            ml_tagWithProperties("c:majorTickMark", {"val": "out"}),
            ml_tagWithProperties("c:minorTickMark", {"val": "none"}),
            ml_tagWithProperties("c:tickLblPos", {"val": "nextTo"}),
            ml_tag("c:spPr",
                   ml_tag("a:noFill"),
                   ml_outline(Style(width=0.75, line_cap=ML_LineCap.FLAT, line_type=ML_LineType.SINGLE,
                                    line_join=ML_LineJoin.ROUND, alignment=ML_PenAlignment.CENTER, colour="black")),
                   ml_tag("a:effectLst")
                   ),
            ml_tag("c:txPr",
                   ml_bodyProperties(properties(rot="0", spcFirstLastPara="1", vertOverflow="ellipsis", vert="horz",
                                                wrap="square", anchor="ctr", anchorCtr="1")),
                   ml_listStyle(),
                   ml_textParagraph(
                       ml_paragraphProperties(
                           ml_defaultTextRunProperties(Font())
                       )
                   )
                   ),
            ml_tagWithProperties("c:crossAx", {"val": str(x_axis_id)}),
            ml_tagWithProperties("c:crosses", {"val": "autoZero"}),
            ml_tagWithProperties("c:crossesBetween", {"val": "between"})
        )
    )
    
    # Chart Legend
    legend = ET.SubElement(chart, "c:legend")
    xml_append(
        legend,
        ml_tagWithProperties("c:legendPos", {"val": "b"}),
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
                       ml_defaultTextRunProperties(Font())
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
