import xml.etree.ElementTree as ET
from officecharts.drawingml import xml_append, xml_header


def content_types() -> bytes:
    root = ET.Element("Types", xmlns="http://schemas.openxmlformats.org/package/2006/content-types")
    xml_append(
        root,
        ET.Element("Default", Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml"),
        ET.Element("Default", Extension="xlsx",
                   ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        ET.Element("Default", Extension="xml", ContentType="application/xml"),
        ET.Element("Override", PartName="/clipboard/drawings/drawing1.xml",
                   ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"),
        ET.Element("Override", PartName="/clipboard/charts/chart1.xml",
                   ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"),
        ET.Element("Override", PartName="/clipboard/charts/style1.xml",
                   ContentType="application/vnd.ms-office.chartstyle+xml"),
        ET.Element("Override", PartName="/clipboard/charts/colors1.xml",
                   ContentType="application/vnd.ms-office.chartcolorstyle+xml"),
        ET.Element("Override", PartName="/clipboard/theme/theme1.xml",
                   ContentType="application/vnd.openxmlformats-officedocument.theme+xml")
    )
    return xml_header() + ET.tostring(root)


def container_relationships() -> bytes:
    root = ET.Element("Relationships", xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
    xml_append(
        root,
        ET.Element("Relationship", Id="rId1",
                   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                   Target="clipboard/drawings/drawing1.xml")
    )
    return xml_header() + ET.tostring(root)


def drawing_relationships() -> bytes:
    root = ET.Element("Relationships", xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
    xml_append(
        root,
        ET.Element("Relationship", Id="rId1",
                   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                   Target="../charts/chart1.xml"),
        ET.Element("Relationship", Id="rId2",
                   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                   Target="../theme/theme1.xml")
    )
    return xml_header() + ET.tostring(root)


def chart_relationships() -> bytes:
    root = ET.Element("Relationships", xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
    xml_append(
        root,
        ET.Element("Relationship", Id="rId1", Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle",
                   Target="style1.xml"),
        ET.Element("Relationship", Id="rId2",
                   Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle", Target="colors1.xml"),
        ET.Element("Relationship", Id="rId3",
                   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
                   Target="../embeddings/Microsoft_Excel_Worksheet.xlsx")
    )
    return xml_header() + ET.tostring(root)


