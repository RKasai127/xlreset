import pytest

from xlreset.cli import _EXIT_FAILURE, _EXIT_SUCCESS, main

_SHEET = b"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetViews><sheetView topLeftCell="C5" zoomScale="80" zoomScaleNormal="80" workbookViewId="0">
<selection activeCell="C5" sqref="C5"/></sheetView></sheetViews>
<sheetData/>
</worksheet>
"""


def test_main_success(tmp_xlsx):
    path = tmp_xlsx({"xl/worksheets/sheet1.xml": _SHEET})
    assert main([str(path)]) == _EXIT_SUCCESS
    data = path.read_bytes()
    assert b'topLeftCell="A1"' in data
    assert b'zoomScale="100"' in data
    assert b'zoomScaleNormal="100"' in data
    assert b'activeCell="A1"' in data
    assert b'sqref="A1"' in data


def test_main_missing_file_reports_error(tmp_path, capsys):
    path = tmp_path / "missing.xlsx"
    assert main([str(path)]) == _EXIT_FAILURE
    captured = capsys.readouterr()
    assert "xlreset:" in captured.err
    assert str(path) in captured.err


def test_main_not_a_valid_workbook_reports_error(tmp_path, capsys):
    path = tmp_path / "garbage.xlsx"
    path.write_bytes(b"this is not a zip file")
    assert main([str(path)]) == _EXIT_FAILURE
    captured = capsys.readouterr()
    assert "xlreset:" in captured.err
    assert str(path) in captured.err


def test_main_non_excel_file_reports_error(tmp_path, capsys):
    path = tmp_path / "notes.txt"
    path.write_text("just a regular text file, not an xlsx")
    assert main([str(path)]) == _EXIT_FAILURE
    captured = capsys.readouterr()
    assert "xlreset:" in captured.err
    assert str(path) in captured.err


def test_main_no_arguments_exits_with_usage_error():
    with pytest.raises(SystemExit) as exc_info:
        main([])
    assert exc_info.value.code == 2


def test_main_too_many_arguments_exits_with_usage_error():
    with pytest.raises(SystemExit) as exc_info:
        main(["a.xlsx", "b.xlsx"])
    assert exc_info.value.code == 2
