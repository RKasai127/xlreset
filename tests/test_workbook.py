import zipfile

import pytest

from xlreset.workbook import XlresetError, reset_view

_SHEET_SCROLLED = b"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView topLeftCell="F20" zoomScale="70" zoomScaleNormal="70" workbookViewId="0">
<selection activeCell="G25" sqref="G25"/></sheetView></sheetViews>
<sheetData/>
</worksheet>
"""

_SHEET_PROTECTED = b"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView topLeftCell="B3" zoomScale="120" zoomScaleNormal="120" workbookViewId="0">
<selection activeCell="B3" sqref="B3"/></sheetView></sheetViews>
<sheetProtection password="C6F3" sheet="1" objects="1" scenarios="1"/>
<sheetData/>
</worksheet>
"""

_SHEET_FROZEN = b"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView workbookViewId="0">
<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
<selection pane="bottomRight" activeCell="B2" sqref="B2"/>
</sheetView></sheetViews>
<sheetData/>
</worksheet>
"""


def test_multi_sheet_workbook_all_sheets_rewritten(tmp_xlsx):
    path = tmp_xlsx(
        {
            "xl/worksheets/sheet1.xml": _SHEET_SCROLLED,
            "xl/worksheets/sheet2.xml": _SHEET_SCROLLED,
        },
        worksheet_paths=("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"),
    )
    reset_view(path)
    with zipfile.ZipFile(path) as zf:
        for name in ("sheet1.xml", "sheet2.xml"):
            data = zf.read(f"xl/worksheets/{name}")
            assert b'topLeftCell="A1"' in data
            assert b'zoomScale="100"' in data
            assert b'zoomScaleNormal="100"' in data
            assert b'activeCell="A1"' in data
            assert b'sqref="A1"' in data


def test_nonstandard_worksheet_filename_is_still_rewritten(tmp_xlsx):
    path = tmp_xlsx(
        {"xl/worksheets/MyCustomSheet.xml": _SHEET_SCROLLED},
        worksheet_paths=("xl/worksheets/MyCustomSheet.xml",),
    )
    reset_view(path)
    with zipfile.ZipFile(path) as zf:
        data = zf.read("xl/worksheets/MyCustomSheet.xml")
    assert b'topLeftCell="A1"' in data
    assert b'zoomScale="100"' in data


def test_freeze_pane_untouched_but_sheet_view_still_rewritten(tmp_xlsx):
    path = tmp_xlsx({"xl/worksheets/sheet1.xml": _SHEET_FROZEN})
    reset_view(path)
    with zipfile.ZipFile(path) as zf:
        data = zf.read("xl/worksheets/sheet1.xml")

    assert (
        b'<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>'
        in data
    )
    assert b'<sheetView workbookViewId="0" topLeftCell="A1"' in data


def test_non_worksheet_parts_untouched(tmp_xlsx):
    path = tmp_xlsx({"xl/worksheets/sheet1.xml": _SHEET_SCROLLED})
    with zipfile.ZipFile(path) as zf:
        before = {
            name: zf.read(name) for name in zf.namelist() if not name.startswith("xl/worksheets/")
        }

    reset_view(path)

    with zipfile.ZipFile(path) as zf:
        for name, data in before.items():
            assert zf.read(name) == data


def test_xlsm_macro_bytes_untouched(tmp_xlsx):
    macro_bytes = bytes(range(256)) * 4
    path = tmp_xlsx(
        {
            "xl/worksheets/sheet1.xml": _SHEET_SCROLLED,
            "xl/vbaProject.bin": macro_bytes,
        },
        filename="sample.xlsm",
    )
    reset_view(path)
    with zipfile.ZipFile(path) as zf:
        assert zf.read("xl/vbaProject.bin") == macro_bytes


def test_sheet_protection_untouched(tmp_xlsx):
    path = tmp_xlsx({"xl/worksheets/sheet1.xml": _SHEET_PROTECTED})
    reset_view(path)
    with zipfile.ZipFile(path) as zf:
        data = zf.read("xl/worksheets/sheet1.xml")
    assert b'<sheetProtection password="C6F3" sheet="1" objects="1" scenarios="1"/>' in data


def test_no_worksheet_parts_raises(tmp_xlsx):
    path = tmp_xlsx({"xl/worksheets/sheet1.bin": b"\x00\x01"}, worksheet_paths=())
    with pytest.raises(XlresetError, match="no worksheet parts"):
        reset_view(path)


def test_missing_workbook_xml_raises(tmp_path):
    path = tmp_path / "broken.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/worksheets/sheet1.xml", _SHEET_SCROLLED)
    with pytest.raises(XlresetError, match="missing required part"):
        reset_view(path)


def test_malformed_workbook_xml_raises(tmp_xlsx):
    path = tmp_xlsx({"xl/workbook.xml": b"not valid xml <<<"})
    with pytest.raises(XlresetError, match="could not parse"):
        reset_view(path)


def test_not_a_zip_raises(tmp_path):
    path = tmp_path / "garbage.xlsx"
    path.write_bytes(b"this is not a zip file")
    with pytest.raises(XlresetError):
        reset_view(path)


def test_missing_file_raises_oserror(tmp_path):
    path = tmp_path / "missing.xlsx"
    with pytest.raises(OSError):
        reset_view(path)


def test_directory_path_raises_oserror(tmp_path):
    path = tmp_path / "a_directory.xlsx"
    path.mkdir()
    with pytest.raises(OSError):
        reset_view(path)
