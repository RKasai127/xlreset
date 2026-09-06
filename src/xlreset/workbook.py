"""Zip-level I/O: rewrite one .xlsx/.xlsm file in place."""

import os
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from xlreset.xml_rewrite import resolve_worksheet_parts, rewrite_worksheet_xml


class XlresetError(Exception):
    """Raised when `path` can't be processed as an xlsx/xlsm workbook."""


def reset_view(path: Path) -> None:
    """Rewrite every worksheet's view in the .xlsx/.xlsm at `path`, in place.

    Raises XlresetError for a bad zip, a zip missing xl/workbook.xml or
    its relationships, or a workbook with no worksheet parts (e.g. an old
    .xls, an .xlsb, or an unrelated file); propagates OSError for
    filesystem problems (missing file, permissions, ...).
    """
    try:
        with zipfile.ZipFile(path, "r") as src:
            infos: list[zipfile.ZipInfo] = src.infolist()
            originals = {info.filename: src.read(info.filename) for info in infos}
    except zipfile.BadZipFile as exc:
        raise XlresetError(f"not a valid zip/xlsx file: {exc}") from exc

    try:
        workbook_xml = originals["xl/workbook.xml"]
        workbook_rels_xml = originals["xl/_rels/workbook.xml.rels"]
    except KeyError as exc:
        raise XlresetError(
            f"missing required part {exc} "
            "(not a supported .xlsx/.xlsm workbook — old .xls, .xlsb, "
            "or an unrelated zip file?)"
        ) from exc

    try:
        worksheet_paths: set[str] = resolve_worksheet_parts(workbook_xml, workbook_rels_xml)
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
            for info in infos:
                data = originals[info.filename]
                if info.filename in worksheet_paths:
                    data = rewrite_worksheet_xml(data)
                dst.writestr(info, data)
        os.replace(tmp_path, path)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise
