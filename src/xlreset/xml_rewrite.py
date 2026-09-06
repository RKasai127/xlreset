"""Find a workbook's worksheet parts and reset their <sheetView> position/zoom."""

import re
import xml.etree.ElementTree as ET

_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_WORKSHEET_RELATIONSHIP_TYPE = f"{_RELATIONSHIPS_NAMESPACE}/worksheet"
_SHEET_SEARCH_PATH = ".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"

# A <Relationship> is OOXML's generic "this part links to that part" record --
# not worksheet-specific. A .rels file lists links to worksheets, styles,
# themes, etc. all the same way; only its Type attribute says which.
_RELATIONSHIP_SEARCH_PATH = (
    "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
)
_REL_ID_ATTR = f"{{{_RELATIONSHIPS_NAMESPACE}}}id"


# [^>]* would miss a tag whose attribute value contains a literal ">" --
# not schema-forbidden, but no real producer writes one here.
_SHEET_VIEW_RE = re.compile(rb"<sheetView\b[^>]*/>|<sheetView\b[^>]*>.*?</sheetView>", re.DOTALL)
_SELECTION_RE = re.compile(rb"<selection\b[^>]*/>")


def resolve_worksheet_parts(workbook_xml: bytes, workbook_rels_xml: bytes) -> set[str]:
    """Resolve the zip paths of the real worksheet parts.

    Reads xl/workbook.xml's ``<sheets><sheet r:id="..."/></sheets>`` list
    and follows each r:id through xl/_rels/workbook.xml.rels to its
    relationship Target, keeping only relationships whose Type identifies
    a worksheet (this excludes chartsheets and dialogsheets, which also
    appear in ``<sheets>`` but point at a different relationship Type).
    Raises xml.etree.ElementTree.ParseError if either input is malformed.
    """
    rels_root = ET.fromstring(workbook_rels_xml)
    worksheet_path_by_rel_id = {
        rel.get("Id"): rel.get("Target", "")
        for rel in rels_root.findall(_RELATIONSHIP_SEARCH_PATH)
        if rel.get("Type") == _WORKSHEET_RELATIONSHIP_TYPE
    }

    workbook_root = ET.fromstring(workbook_xml)
    sheet_rel_ids = {sheet.get(_REL_ID_ATTR) for sheet in workbook_root.findall(_SHEET_SEARCH_PATH)}

    return {
        _resolve_zip_path(worksheet_path_by_rel_id[rel_id])
        for rel_id in sheet_rel_ids
        if rel_id in worksheet_path_by_rel_id
    }


def _resolve_zip_path(raw_worksheet_path: str) -> str:
    """Resolve a relationship Target to a full zip path.

    A Target is usually relative to xl/ (where workbook.xml lives), e.g.
    "worksheets/sheet1.xml" -> "xl/worksheets/sheet1.xml". A leading "/"
    instead means it's already package-root-relative.
    """
    if raw_worksheet_path.startswith("/"):
        return raw_worksheet_path.lstrip("/")
    return f"xl/{raw_worksheet_path}"


def rewrite_worksheet_xml(data: bytes) -> bytes:
    """Reset every sheetView's position, zoom, and selection to A1/100%."""
    data = _SHEET_VIEW_RE.sub(_rewrite_sheet_view, data)
    return _SELECTION_RE.sub(_rewrite_selection, data)


def _rewrite_sheet_view(sheet_view_match: re.Match[bytes]) -> bytes:
    """Set position/zoom on one <sheetView>'s own opening tag only."""
    view = sheet_view_match.group(0)
    if view.rstrip().endswith(b"/>"):
        opening, rest = view, b""
    else:
        end = view.index(b">") + 1
        opening, rest = view[:end], view[end:]

    opening = _upsert_attr(opening, b"topLeftCell", b"A1")
    opening = _upsert_attr(opening, b"zoomScale", b"100")
    opening = _upsert_attr(opening, b"zoomScaleNormal", b"100")
    return opening + rest


def _rewrite_selection(selection_match: re.Match[bytes]) -> bytes:
    """Set the active cell on one <selection> tag."""
    tag = selection_match.group(0)
    tag = _upsert_attr(tag, b"activeCell", b"A1")
    return _upsert_attr(tag, b"sqref", b"A1")


def _upsert_attr(tag: bytes, name: bytes, value: bytes) -> bytes:
    """Upsert an attribute on a single opening/self-closing tag."""
    pattern = re.compile(name + rb'="[^"]*"')
    replacement = name + b'="' + value + b'"'

    if pattern.search(tag):
        return pattern.sub(replacement, tag, count=1)

    stripped = tag.rstrip()
    if stripped.endswith(b"/>"):
        return stripped[:-2].rstrip() + b" " + replacement + b"/>"

    return stripped[:-1].rstrip() + b" " + replacement + b">"
