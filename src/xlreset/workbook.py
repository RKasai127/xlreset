"""Zip-level I/O: rewrite one .xlsx/.xlsm file in place."""

import os
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from xlreset.xml_rewrite import (
    resolve_first_worksheet_part,
    resolve_worksheet_parts,
    rewrite_workbook_xml,
    rewrite_worksheet_xml,
)

_WORKBOOK_XML_PATH = "xl/workbook.xml"


class XlresetError(Exception):
    """Raised when `path` can't be processed as an xlsx/xlsm workbook."""


def reset_view(path: Path) -> None:
    """Rewrite every worksheet's view in the .xlsx/.xlsm at `path`, in place.

    For every sheet, resets scroll position/zoom/selection to A1/100%, and
    makes the first sheet (in tab order) the active tab.

    Raises XlresetError for a bad zip, a zip missing xl/workbook.xml or
    its relationships, or a workbook with no worksheet parts (e.g. an old
    .xls, an .xlsb, or an unrelated file); propagates OSError for
    filesystem problems (missing file, permissions, ...).
    """
    try:
        with zipfile.ZipFile(path, "r") as src:
            zip_infos: list[zipfile.ZipInfo] = src.infolist()
            original_content_by_path: dict[str, bytes] = {
                zip_info.filename: src.read(zip_info.filename) for zip_info in zip_infos
            }
    except zipfile.BadZipFile as exc:
        raise XlresetError(f"not a valid zip/xlsx file: {exc}") from exc

    try:
        workbook_xml = original_content_by_path[_WORKBOOK_XML_PATH]
        workbook_rels_xml = original_content_by_path["xl/_rels/workbook.xml.rels"]
    except KeyError as exc:
        raise XlresetError(
            f"missing required part {exc} "
            "(not a supported .xlsx/.xlsm workbook — old .xls, .xlsb, "
            "or an unrelated zip file?)"
        ) from exc

    try:
        worksheet_paths: set[str] = resolve_worksheet_parts(workbook_xml, workbook_rels_xml)
        active_worksheet_path = resolve_first_worksheet_part(workbook_xml, workbook_rels_xml)
    except ET.ParseError as exc:
        raise XlresetError(f"could not parse workbook structure: {exc}") from exc

    if not worksheet_paths:
        raise XlresetError("xl/workbook.xml declares no worksheet parts")

    tmp_file_descriptor, tmp_path_str = tempfile.mkstemp(
        dir=path.parent or ".", prefix=".xlreset-", suffix=".tmp"
    )
    tmp_path = Path(tmp_path_str)

    try:
        with (
            os.fdopen(tmp_file_descriptor, "wb") as tmp_file,
            zipfile.ZipFile(tmp_file, "w", strict_timestamps=False) as dst,
        ):
            for zip_info in zip_infos:
                content = original_content_by_path[zip_info.filename]
                if zip_info.filename == _WORKBOOK_XML_PATH:
                    content = rewrite_workbook_xml(content)
                elif zip_info.filename in worksheet_paths:
                    content = rewrite_worksheet_xml(
                        content, is_first_sheet=zip_info.filename == active_worksheet_path
                    )
                dst.writestr(zip_info, content)
        os.replace(tmp_path, path)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise
