from __future__ import annotations

import argparse
from pathlib import Path
from tempfile import NamedTemporaryFile
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

CORE_PROPERTIES = "docProps/core.xml"
WORKSHEET_PREFIX = "xl/worksheets/sheet"
FIXED_TIMESTAMP = "2026-01-01T00:00:00Z"


def sanitize_xlsx_fixture(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(source) as input_archive:
        with NamedTemporaryFile(dir=destination.parent, suffix=".xlsx", delete=False) as temp_file:
            temp_path = Path(temp_file.name)
        try:
            with ZipFile(temp_path, "w", compression=ZIP_DEFLATED) as output_archive:
                for entry in input_archive.infolist():
                    if _should_remove(entry.filename):
                        continue
                    data = input_archive.read(entry.filename)
                    if entry.filename == CORE_PROPERTIES:
                        data = _sanitize_core_properties(data)
                    elif _is_worksheet(entry.filename):
                        data = _remove_printer_relationship(data)
                    clean_entry = ZipInfo(entry.filename, date_time=(2026, 1, 1, 0, 0, 0))
                    clean_entry.compress_type = ZIP_DEFLATED
                    clean_entry.external_attr = entry.external_attr
                    output_archive.writestr(clean_entry, data)
            temp_path.replace(destination)
        finally:
            temp_path.unlink(missing_ok=True)


def _should_remove(name: str) -> bool:
    return (
        name.startswith("[trash]/")
        or name.startswith("xl/printerSettings/")
        or (name.startswith("xl/worksheets/_rels/sheet") and name.endswith(".rels"))
    )


def _is_worksheet(name: str) -> bool:
    return name.startswith(WORKSHEET_PREFIX) and name.endswith(".xml")


def _sanitize_core_properties(data: bytes) -> bytes:
    root = ElementTree.fromstring(data)
    for element in root.iter():
        local_name = element.tag.rsplit("}", 1)[-1]
        if local_name in {"creator", "lastModifiedBy"}:
            element.text = "Excel Git Viewer Test Fixture"
        elif local_name in {"created", "modified"}:
            element.text = FIXED_TIMESTAMP
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)


def _remove_printer_relationship(data: bytes) -> bytes:
    root = ElementTree.fromstring(data)
    for element in root.iter():
        if element.tag.rsplit("}", 1)[-1] != "pageSetup":
            continue
        for attribute in list(element.attrib):
            if attribute.rsplit("}", 1)[-1] == "id":
                del element.attrib[attribute]
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    arguments = parser.parse_args()
    sanitize_xlsx_fixture(arguments.source, arguments.destination)


if __name__ == "__main__":
    main()
