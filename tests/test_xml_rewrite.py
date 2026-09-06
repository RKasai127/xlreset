import xml.etree.ElementTree as ET

import pytest

from xlreset.xml_rewrite import (
    resolve_first_worksheet_part,
    resolve_worksheet_parts,
    rewrite_workbook_xml,
    rewrite_worksheet_xml,
)

_MAIN_NAMESPACE_ATTR = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
_RELATIONSHIPS_NAMESPACE_ATTR = (
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)
_PACKAGE_RELATIONSHIPS_NAMESPACE_ATTR = (
    'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
)
_WORKSHEET_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
_CHARTSHEET_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"

_WRAP = b'<?xml version="1.0"?><worksheet>%b<sheetData/></worksheet>'


def _worksheet(sheet_views: bytes) -> bytes:
    return _WRAP % sheet_views


def test_self_closing_sheet_view_gains_attributes():
    xml = _worksheet(b'<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    result = rewrite_worksheet_xml(xml)
    assert b'topLeftCell="A1"' in result
    assert b'zoomScale="100"' in result
    assert b'zoomScaleNormal="100"' in result
    assert b'workbookViewId="0"' in result


def test_space_before_self_closing_slash_is_handled():
    xml = _worksheet(b'<sheetViews><sheetView topLeftCell="F5" /></sheetViews>')
    result = rewrite_worksheet_xml(xml)
    assert b'topLeftCell="A1"' in result
    assert b'zoomScale="100"' in result
    assert b'zoomScaleNormal="100"' in result


def test_multiple_sheet_views_each_rewritten_independently():
    xml = _worksheet(
        b"<sheetViews>"
        b'<sheetView topLeftCell="F20" zoomScale="70" zoomScaleNormal="70" workbookViewId="0"/>'
        b'<sheetView topLeftCell="Z99" zoomScale="40" zoomScaleNormal="40" workbookViewId="1"/>'
        b"</sheetViews>"
    )
    result = rewrite_worksheet_xml(xml)
    assert result.count(b'topLeftCell="A1"') == 2
    assert result.count(b'zoomScale="100"') == 2
    assert result.count(b'zoomScaleNormal="100"') == 2
    assert b'workbookViewId="0"' in result
    assert b'workbookViewId="1"' in result


def test_existing_attributes_are_overwritten():
    xml = _worksheet(
        b'<sheetViews><sheetView topLeftCell="C5" zoomScale="85" zoomScaleNormal="85">'
        b'<selection activeCell="C5" sqref="C5:D10"/>'
        b"</sheetView></sheetViews>"
    )
    result = rewrite_worksheet_xml(xml)
    assert b'topLeftCell="C5"' not in result
    assert b'zoomScale="85"' not in result
    assert b'topLeftCell="A1"' in result
    assert b'zoomScale="100"' in result
    assert b'zoomScaleNormal="100"' in result
    assert b'activeCell="A1"' in result
    assert b'sqref="A1"' in result


def test_freeze_pane_untouched_but_all_selections_reset():
    pane = b'<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>'
    xml = _worksheet(
        b'<sheetViews><sheetView workbookViewId="0">'
        + pane
        + b'<selection pane="topRight"/>'
        + b'<selection pane="bottomLeft"/>'
        + b'<selection pane="bottomRight" activeCell="B2" sqref="B2"/>'
        + b"</sheetView></sheetViews>"
    )
    result = rewrite_worksheet_xml(xml)

    assert pane in result

    assert result.count(b'activeCell="A1"') == 3
    assert result.count(b'sqref="A1"') == 3
    assert b'<selection pane="topRight" activeCell="A1" sqref="A1"/>' in result
    assert b'<selection pane="bottomLeft" activeCell="A1" sqref="A1"/>' in result
    assert b'<selection pane="bottomRight" activeCell="A1" sqref="A1"/>' in result


def test_cell_text_resembling_attributes_is_untouched():
    xml = (
        b'<?xml version="1.0"?><worksheet>'
        b'<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        b'<sheetData><row><c t="inlineStr"><is><t>topLeftCell selection zoomScale</t></is></c></row></sheetData>'
        b"</worksheet>"
    )
    result = rewrite_worksheet_xml(xml)
    assert b"<t>topLeftCell selection zoomScale</t>" in result


def test_no_sheet_views_block_is_untouched():
    xml = b'<?xml version="1.0"?><worksheet><sheetData/></worksheet>'
    assert rewrite_worksheet_xml(xml) == xml


def test_is_first_sheet_true_sets_tab_selected():
    xml = _worksheet(b'<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    result = rewrite_worksheet_xml(xml, is_first_sheet=True)
    assert b'tabSelected="1"' in result


def test_is_first_sheet_false_clears_existing_tab_selected():
    xml = _worksheet(b'<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>')
    result = rewrite_worksheet_xml(xml, is_first_sheet=False)
    assert b'tabSelected="0"' in result


def test_is_first_sheet_defaults_to_false():
    xml = _worksheet(b'<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    result = rewrite_worksheet_xml(xml)
    assert b'tabSelected="0"' in result


def test_workbook_view_active_tab_reset_to_zero():
    xml = b'<?xml version="1.0"?><workbook><bookViews><workbookView activeTab="3"/></bookViews></workbook>'
    result = rewrite_workbook_xml(xml)
    assert b'activeTab="0"' in result


def test_workbook_view_gains_active_tab_when_absent():
    xml = b'<?xml version="1.0"?><workbook><bookViews><workbookView windowWidth="100"/></bookViews></workbook>'
    result = rewrite_workbook_xml(xml)
    assert b'activeTab="0"' in result
    assert b'windowWidth="100"' in result


def test_workbook_xml_without_workbook_view_is_untouched():
    xml = b'<?xml version="1.0"?><workbook><sheets/></workbook>'
    assert rewrite_workbook_xml(xml) == xml


def test_resolve_worksheet_parts_standard_name():
    workbook = (
        f"<workbook {_MAIN_NAMESPACE_ATTR} {_RELATIONSHIPS_NAMESPACE_ATTR}>"
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    ).encode()
    rels = (
        f"<Relationships {_PACKAGE_RELATIONSHIPS_NAMESPACE_ATTR}>"
        f'<Relationship Id="rId1" Type="{_WORKSHEET_TYPE}" Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    ).encode()
    assert resolve_worksheet_parts(workbook, rels) == {"xl/worksheets/sheet1.xml"}


def test_resolve_worksheet_parts_multiple_sheets():
    workbook = (
        f"<workbook {_MAIN_NAMESPACE_ATTR} {_RELATIONSHIPS_NAMESPACE_ATTR}>"
        "<sheets>"
        '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>'
        '<sheet name="Sheet2" sheetId="2" r:id="rId2"/>'
        '<sheet name="Sheet3" sheetId="3" r:id="rId3"/>'
        "</sheets>"
        "</workbook>"
    ).encode()
    rels = (
        f"<Relationships {_PACKAGE_RELATIONSHIPS_NAMESPACE_ATTR}>"
        f'<Relationship Id="rId1" Type="{_WORKSHEET_TYPE}" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{_WORKSHEET_TYPE}" Target="worksheets/sheet2.xml"/>'
        f'<Relationship Id="rId3" Type="{_WORKSHEET_TYPE}" Target="worksheets/sheet3.xml"/>'
        "</Relationships>"
    ).encode()
    assert resolve_worksheet_parts(workbook, rels) == {
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
        "xl/worksheets/sheet3.xml",
    }


def test_resolve_worksheet_parts_nonstandard_name():
    workbook = (
        f"<workbook {_MAIN_NAMESPACE_ATTR} {_RELATIONSHIPS_NAMESPACE_ATTR}>"
        '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    ).encode()
    rels = (
        f"<Relationships {_PACKAGE_RELATIONSHIPS_NAMESPACE_ATTR}>"
        f'<Relationship Id="rId1" Type="{_WORKSHEET_TYPE}" Target="worksheets/MyCustomSheet.xml"/>'
        "</Relationships>"
    ).encode()
    assert resolve_worksheet_parts(workbook, rels) == {"xl/worksheets/MyCustomSheet.xml"}


def test_resolve_worksheet_parts_excludes_chartsheet():
    workbook = (
        f"<workbook {_MAIN_NAMESPACE_ATTR} {_RELATIONSHIPS_NAMESPACE_ATTR}>"
        "<sheets>"
        '<sheet name="Data" sheetId="1" r:id="rId1"/>'
        '<sheet name="Chart1" sheetId="2" r:id="rId2"/>'
        "</sheets>"
        "</workbook>"
    ).encode()
    rels = (
        f"<Relationships {_PACKAGE_RELATIONSHIPS_NAMESPACE_ATTR}>"
        f'<Relationship Id="rId1" Type="{_WORKSHEET_TYPE}" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{_CHARTSHEET_TYPE}" Target="chartsheets/sheet1.xml"/>'
        "</Relationships>"
    ).encode()
    assert resolve_worksheet_parts(workbook, rels) == {"xl/worksheets/sheet1.xml"}


def test_resolve_worksheet_parts_absolute_target():
    workbook = (
        f"<workbook {_MAIN_NAMESPACE_ATTR} {_RELATIONSHIPS_NAMESPACE_ATTR}>"
        '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    ).encode()
    rels = (
        f"<Relationships {_PACKAGE_RELATIONSHIPS_NAMESPACE_ATTR}>"
        f'<Relationship Id="rId1" Type="{_WORKSHEET_TYPE}" '
        'Target="/xl/worksheets/sheet1.xml"/>'
        "</Relationships>"
    ).encode()
    assert resolve_worksheet_parts(workbook, rels) == {"xl/worksheets/sheet1.xml"}


def test_resolve_worksheet_parts_malformed_xml_raises():
    with pytest.raises(ET.ParseError):
        resolve_worksheet_parts(b"not xml <<<", b"<Relationships></Relationships>")


def test_resolve_first_worksheet_part_returns_first_in_tab_order():
    workbook = (
        f"<workbook {_MAIN_NAMESPACE_ATTR} {_RELATIONSHIPS_NAMESPACE_ATTR}>"
        "<sheets>"
        '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>'
        '<sheet name="Sheet2" sheetId="2" r:id="rId2"/>'
        "</sheets>"
        "</workbook>"
    ).encode()
    rels = (
        f"<Relationships {_PACKAGE_RELATIONSHIPS_NAMESPACE_ATTR}>"
        f'<Relationship Id="rId1" Type="{_WORKSHEET_TYPE}" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{_WORKSHEET_TYPE}" Target="worksheets/sheet2.xml"/>'
        "</Relationships>"
    ).encode()
    assert resolve_first_worksheet_part(workbook, rels) == "xl/worksheets/sheet1.xml"


def test_resolve_first_worksheet_part_none_when_first_sheet_is_chartsheet():
    workbook = (
        f"<workbook {_MAIN_NAMESPACE_ATTR} {_RELATIONSHIPS_NAMESPACE_ATTR}>"
        "<sheets>"
        '<sheet name="Chart1" sheetId="1" r:id="rId1"/>'
        '<sheet name="Data" sheetId="2" r:id="rId2"/>'
        "</sheets>"
        "</workbook>"
    ).encode()
    rels = (
        f"<Relationships {_PACKAGE_RELATIONSHIPS_NAMESPACE_ATTR}>"
        f'<Relationship Id="rId1" Type="{_CHARTSHEET_TYPE}" Target="chartsheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{_WORKSHEET_TYPE}" Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    ).encode()
    assert resolve_first_worksheet_part(workbook, rels) is None


def test_resolve_first_worksheet_part_none_when_no_sheets():
    workbook = f"<workbook {_MAIN_NAMESPACE_ATTR} {_RELATIONSHIPS_NAMESPACE_ATTR}><sheets/></workbook>".encode()
    rels = f"<Relationships {_PACKAGE_RELATIONSHIPS_NAMESPACE_ATTR}></Relationships>".encode()
    assert resolve_first_worksheet_part(workbook, rels) is None
