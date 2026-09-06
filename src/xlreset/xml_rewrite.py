"""Find a workbook's worksheet parts and reset their view (position/zoom/active tab)."""

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
_WORKBOOK_VIEW_RE = re.compile(
    rb"<workbookView\b[^>]*/>|<workbookView\b[^>]*>.*?</workbookView>", re.DOTALL
)


def resolve_worksheet_parts(workbook_xml: bytes, workbook_rels_xml: bytes) -> set[str]:
    """Resolve the zip paths of the real worksheet parts.

    Reads xl/workbook.xml's ``<sheets><sheet r:id="..."/></sheets>`` list
    and follows each r:id through xl/_rels/workbook.xml.rels to its
    relationship Target, keeping only relationships whose Type identifies
    a worksheet (this excludes chartsheets and dialogsheets, which also
    appear in ``<sheets>`` but point at a different relationship Type).
    Raises xml.etree.ElementTree.ParseError if either input is malformed.
    """
    worksheet_path_by_rel_id = _worksheet_path_by_rel_id(workbook_rels_xml)
    workbook_root = ET.fromstring(workbook_xml)
    sheet_rel_ids = {sheet.get(_REL_ID_ATTR) for sheet in workbook_root.findall(_SHEET_SEARCH_PATH)}

    return {
        worksheet_path_by_rel_id[rel_id]
        for rel_id in sheet_rel_ids
        if rel_id in worksheet_path_by_rel_id
    }


def resolve_first_worksheet_part(workbook_xml: bytes, workbook_rels_xml: bytes) -> str | None:
    """Resolve the zip path of the first sheet in tab order (the first <sheet> in
    xl/workbook.xml's <sheets> list), i.e. the sheet that should become the active tab.

    Returns None if the workbook declares no sheets, or if the first one isn't
    a worksheet (e.g. a chartsheet) or its r:id doesn't resolve -- in these
    cases there's no worksheet part to mark as the active tab.
    Raises xml.etree.ElementTree.ParseError if either input is malformed.
    """
    worksheet_path_by_rel_id = _worksheet_path_by_rel_id(workbook_rels_xml)
    workbook_root = ET.fromstring(workbook_xml)
    sheets = workbook_root.findall(_SHEET_SEARCH_PATH)
    if not sheets:
        return None
    first_sheet_rel_id = sheets[0].get(_REL_ID_ATTR)
    return worksheet_path_by_rel_id.get(first_sheet_rel_id)


def _worksheet_path_by_rel_id(workbook_rels_xml: bytes) -> dict[str | None, str]:
    """Map each relationship Id to its resolved zip path, for worksheet relationships only."""
    rels_root = ET.fromstring(workbook_rels_xml)
    return {
        rel.get("Id"): _resolve_zip_path(rel.get("Target", ""))
        for rel in rels_root.findall(_RELATIONSHIP_SEARCH_PATH)
        if rel.get("Type") == _WORKSHEET_RELATIONSHIP_TYPE
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


def rewrite_worksheet_xml(worksheet_xml: bytes, is_first_sheet: bool = False) -> bytes:
    """Reset every sheetView's position, zoom, selection, and tab-selected state.

    `is_first_sheet` marks this as the first sheet in tab order, which
    becomes the active tab when the workbook is opened.
    """
    rewritten_xml = _SHEET_VIEW_RE.sub(
        lambda match: _rewrite_sheet_view(match, is_first_sheet), worksheet_xml
    )
    rewritten_xml = _SELECTION_RE.sub(_rewrite_selection, rewritten_xml)
    return rewritten_xml


def rewrite_workbook_xml(workbook_xml: bytes) -> bytes:
    """Set xl/workbook.xml's active tab (<workbookView activeTab="...">) to the first sheet."""
    rewritten_xml = _WORKBOOK_VIEW_RE.sub(_rewrite_workbook_view, workbook_xml)
    return rewritten_xml


def _rewrite_sheet_view(sheet_view_match: re.Match[bytes], is_first_sheet: bool) -> bytes:
    """Set position/zoom/tab-selected state on one <sheetView>'s own opening tag only."""
    opening, rest = _split_opening_tag(sheet_view_match.group(0))
    opening = _upsert_attr(opening, b"topLeftCell", b"A1")
    opening = _upsert_attr(opening, b"zoomScale", b"100")
    opening = _upsert_attr(opening, b"zoomScaleNormal", b"100")
    opening = _upsert_attr(opening, b"tabSelected", b"1" if is_first_sheet else b"0")
    return opening + rest


def _rewrite_selection(selection_match: re.Match[bytes]) -> bytes:
    """Set the active cell on one <selection> tag."""
    tag = selection_match.group(0)
    tag = _upsert_attr(tag, b"activeCell", b"A1")
    return _upsert_attr(tag, b"sqref", b"A1")


def _rewrite_workbook_view(workbook_view_match: re.Match[bytes]) -> bytes:
    """Set activeTab=0 on the <workbookView>'s own opening tag only."""
    opening, rest = _split_opening_tag(workbook_view_match.group(0))
    return _upsert_attr(opening, b"activeTab", b"0") + rest


def _split_opening_tag(tag: bytes) -> tuple[bytes, bytes]:
    """Split a matched element into its opening/self-closing tag and everything after."""
    if tag.rstrip().endswith(b"/>"):
        return tag, b""
    end = tag.index(b">") + 1
    return tag[:end], tag[end:]


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
