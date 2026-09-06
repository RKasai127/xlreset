"""Fixture builders for constructing minimal valid xlsx/xlsm zips in memory."""

import io
import zipfile
from collections.abc import Sequence
from pathlib import Path

import pytest

_CONTENT_TYPES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
</Types>
"""

_ROOT_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

_WORKSHEET_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
)


def _make_workbook_xml(sheet_count: int) -> bytes:
    sheets = "".join(
        f'<sheet name="Sheet{i}" sheetId="{i}" r:id="rId{i}"/>' for i in range(1, sheet_count + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheets}</sheets>"
        "</workbook>"
    ).encode()


def _make_workbook_rels_xml(worksheet_paths: Sequence[str]) -> bytes:
    relationships = "".join(
        f'<Relationship Id="rId{i}" Type="{_WORKSHEET_REL_TYPE}" '
        f'Target="{path.removeprefix("xl/")}"/>'
        for i, path in enumerate(worksheet_paths, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{relationships}"
        "</Relationships>"
    ).encode()


def make_xlsx_bytes(
    entries: dict[str, bytes],
    worksheet_paths: Sequence[str] = ("xl/worksheets/sheet1.xml",),
) -> bytes:
    """Build a minimal valid .xlsx/.xlsm zip in memory.

    `entries` maps zip paths, given verbatim, to their raw content bytes.
    `worksheet_paths` lists the paths that xl/workbook.xml and
    xl/_rels/workbook.xml.rels should declare as real worksheets, in
    order -- this is what a test overrides to prove worksheet discovery
    does not depend on the "sheetN.xml" filename convention. An entry in
    `entries` for "xl/workbook.xml" or "xl/_rels/workbook.xml.rels"
    overrides the generated default entirely, for tests that need to
    craft those parts by hand.
    """
    defaults = {
        "[Content_Types].xml": _CONTENT_TYPES,
        "_rels/.rels": _ROOT_RELS,
        "xl/workbook.xml": _make_workbook_xml(len(worksheet_paths)),
        "xl/_rels/workbook.xml.rels": _make_workbook_rels_xml(worksheet_paths),
    }
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for name, data in {**defaults, **entries}.items():
            zf.writestr(name, data)
    return buf.getvalue()


@pytest.fixture
def tmp_xlsx(tmp_path):
    """Return a function that writes an xlsx built from given parts to disk."""

    def _write(
        entries: dict[str, bytes],
        worksheet_paths: Sequence[str] = ("xl/worksheets/sheet1.xml",),
        filename: str = "sample.xlsx",
    ) -> Path:
        path = tmp_path / filename
        path.write_bytes(make_xlsx_bytes(entries, worksheet_paths))
        return path

    return _write
