import xml.etree.ElementTree as ET
from .drawingml import xml_header


def container_theme() -> bytes:
    root = ET.Element("clipboardTheme", {"xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
    return xml_header() + ET.tostring(root)
