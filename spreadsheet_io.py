from __future__ import annotations

import csv
import posixpath
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Tuple
from xml.etree import ElementTree as ET

from validation_utils import clean_text


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

SPREADSHEET_SUFFIXES = {".csv", ".txt", ".tsv", ".xlsx", ".xls", ".xlsm", ".ods"}


def column_index(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref.upper())
    if not letters:
        return 0
    total = 0
    for char in letters.group(0):
        total = total * 26 + (ord(char) - ord("A") + 1)
    return total - 1


def read_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings = []
    for si in root.findall("main:si", NS):
        parts = []
        for text in si.findall(".//main:t", NS):
            parts.append(text.text or "")
        strings.append("".join(parts))
    return strings


def workbook_sheets(zf: zipfile.ZipFile) -> List[Tuple[str, str]]:
    names = set(zf.namelist())
    missing = [
        name
        for name in ("xl/workbook.xml", "xl/_rels/workbook.xml.rels")
        if name not in names
    ]
    if missing:
        raise ValueError(
            "Uploaded file is not a valid .xlsx workbook. "
            "Please open it in Excel/Google Sheets and save/download it as "
            "Excel Workbook (.xlsx), then upload again."
        )

    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_lookup = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall("pkgrel:Relationship", NS)
    }

    sheets = []
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        name = sheet.attrib["name"]
        rel_id = sheet.attrib[f"{{{NS['rel']}}}id"]
        target = rel_lookup[rel_id]
        path = target if target.startswith("xl/") else posixpath.normpath("xl/" + target)
        sheets.append((name, path))
    return sheets


def cell_value(cell: ET.Element, shared_strings: List[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.findall(".//main:t", NS))

    value_node = cell.find("main:v", NS)
    if value_node is None or value_node.text is None:
        return ""

    value = value_node.text
    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (IndexError, ValueError):
            return value
    return value


def read_xlsx_rows(path: Path) -> Tuple[str, List[str], List[Dict[str, str]]]:
    first_bytes = path.read_bytes()[:8]
    if first_bytes.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        raise ValueError(
            "Uploaded file appears to be old Excel .xls format. "
            "This format must be converted before reading."
        )

    if not zipfile.is_zipfile(path):
        raise ValueError(
            "Uploaded file is not a valid .xlsx workbook. "
            "Please upload an Excel Workbook (.xlsx) file."
        )

    try:
        zf_context = zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        raise ValueError(
            "Uploaded file is not a valid .xlsx workbook. "
            "Please upload an Excel Workbook (.xlsx) file."
        ) from exc

    with zf_context as zf:
        shared_strings = read_shared_strings(zf)
        for sheet_name, sheet_path in workbook_sheets(zf):
            root = ET.fromstring(zf.read(sheet_path))
            raw_rows: List[List[str]] = []
            for row in root.findall(".//main:sheetData/main:row", NS):
                values: Dict[int, str] = {}
                max_index = -1
                for cell in row.findall("main:c", NS):
                    ref = cell.attrib.get("r", "A1")
                    idx = column_index(ref)
                    values[idx] = clean_text(cell_value(cell, shared_strings))
                    max_index = max(max_index, idx)
                if max_index >= 0:
                    raw_rows.append([values.get(i, "") for i in range(max_index + 1)])

            non_empty = [row for row in raw_rows if any(clean_text(value) for value in row)]
            if not non_empty:
                continue

            headers = [clean_text(value) for value in non_empty[0]]
            records = []
            for row in non_empty[1:]:
                padded = row + [""] * (len(headers) - len(row))
                record = {headers[i]: padded[i] for i in range(len(headers)) if headers[i]}
                if any(clean_text(value) for value in record.values()):
                    records.append(record)
            return sheet_name, headers, records
    return "", [], []


def is_xlsx_workbook(path: Path) -> bool:
    if not zipfile.is_zipfile(path):
        return False
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
    except zipfile.BadZipFile:
        return False
    return "xl/workbook.xml" in names and "xl/_rels/workbook.xml.rels" in names


def read_csv_rows(path: Path) -> Tuple[str, List[str], List[Dict[str, str]]]:
    with path.open(newline="", encoding="utf-8-sig") as f:
        sample = f.read(4096)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
        except csv.Error:
            dialect = csv.excel
        reader = csv.DictReader(f, dialect=dialect)
        headers = [clean_text(header) for header in (reader.fieldnames or [])]
        rows: List[Dict[str, str]] = []
        for row in reader:
            cleaned = {
                clean_text(key): clean_text(value)
                for key, value in row.items()
                if key is not None and clean_text(key)
            }
            if any(clean_text(value) for value in cleaned.values()):
                rows.append(cleaned)
    return "CSV", headers, rows


def spreadsheet_kind(path: Path) -> str:
    sample = path.read_bytes()[:4096]
    first_bytes = sample[:8]
    suffix = path.suffix.lower()
    if first_bytes.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return "xls"
    if suffix == ".zip":
        return "zip"
    if suffix == ".csv":
        return "csv"
    if is_xlsx_workbook(path):
        return "xlsx"
    if zipfile.is_zipfile(path):
        return "zip"
    if suffix in {".xls", ".xlsm", ".ods"}:
        return "convertible"
    try:
        text_sample = sample.decode("utf-8-sig")
    except UnicodeDecodeError:
        text_sample = ""
    if text_sample and "\x00" not in text_sample:
        stripped = text_sample.lstrip().lower()
        if not stripped.startswith(("<html", "<table")) and any(
            delimiter in text_sample for delimiter in [",", "\t", ";", "\n"]
        ):
            return "csv"
    return "csv" if suffix in {".txt", ".tsv"} else "unknown"


def convert_to_xlsx(path: Path, kind: str) -> Path:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise ValueError(
            "Uploaded file is not directly readable. For .xls/.ods/other spreadsheet "
            "formats, Jenkins must have LibreOffice installed for automatic conversion."
        )

    suffix = {
        "xls": ".xls",
        "convertible": path.suffix if path.suffix else ".xls",
        "xlsx": ".xlsx",
    }.get(kind, path.suffix or ".xls")

    tmp_dir = Path(tempfile.mkdtemp(prefix="fm_spreadsheet_"))
    input_path = tmp_dir / f"upload{suffix}"
    input_path.write_bytes(path.read_bytes())

    result = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(tmp_dir),
            str(input_path),
        ],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    converted_files = sorted(
        [candidate for candidate in tmp_dir.glob("*.xlsx") if candidate != input_path],
        key=lambda candidate: candidate.stat().st_mtime,
    )
    if result.returncode != 0 or not converted_files:
        message = clean_text(result.stderr) or clean_text(result.stdout) or "conversion failed"
        raise ValueError(f"Unable to convert uploaded spreadsheet to .xlsx: {message}")
    return converted_files[-1]


def spreadsheet_files(root: Path) -> List[Path]:
    if root.is_file():
        return [root]

    files = []
    for candidate in root.rglob("*"):
        if not candidate.is_file():
            continue
        name = candidate.name
        if name.startswith(".") or name.startswith("~$") or "__MACOSX" in candidate.parts:
            continue
        if candidate.suffix.lower() in SPREADSHEET_SUFFIXES or not candidate.suffix:
            files.append(candidate)
    return sorted(files, key=lambda item: str(item).lower())


def extract_spreadsheet_archive(path: Path) -> Path:
    tmp_dir = Path(tempfile.mkdtemp(prefix="fm_ops_uploads_"))
    root = tmp_dir / "files"
    root.mkdir(parents=True, exist_ok=True)
    root_resolved = root.resolve()

    with zipfile.ZipFile(path) as zf:
        for member in zf.infolist():
            member_path = Path(member.filename)
            if member.is_dir() or member_path.name.startswith(".") or "__MACOSX" in member_path.parts:
                continue
            target = (root / member.filename).resolve()
            if not target.is_relative_to(root_resolved):
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(member) as source, target.open("wb") as dest:
                shutil.copyfileobj(source, dest)
    return root


def merge_spreadsheet_rows(paths: List[Path]) -> Tuple[str, List[str], List[Dict[str, str]]]:
    headers: List[str] = []
    rows: List[Dict[str, str]] = []
    used_files = 0

    for path in paths:
        try:
            sheet_name, file_headers, file_rows = read_single_spreadsheet_rows(path)
        except ValueError as exc:
            raise ValueError(f"{path.name}: {exc}") from exc

        if not file_headers and not file_rows:
            continue
        used_files += 1
        for header in file_headers:
            if header and header not in headers:
                headers.append(header)
        for row in file_rows:
            row["__source_file"] = path.name
            row["__source_sheet"] = sheet_name
            rows.append(row)

    if not used_files:
        raise ValueError("No readable spreadsheet files found to merge.")
    return f"Merged {used_files} files", headers, rows


def read_single_spreadsheet_rows(path: Path) -> Tuple[str, List[str], List[Dict[str, str]]]:
    kind = spreadsheet_kind(path)
    if kind == "csv":
        return read_csv_rows(path)
    if kind == "xlsx":
        try:
            return read_xlsx_rows(path)
        except ValueError:
            converted = convert_to_xlsx(path, kind)
            return read_xlsx_rows(converted)
    if kind in {"xls", "convertible", "unknown"}:
        converted = convert_to_xlsx(path, kind)
        return read_xlsx_rows(converted)
    raise ValueError("Unsupported spreadsheet file format")


def read_spreadsheet_rows(path: Path) -> Tuple[str, List[str], List[Dict[str, str]]]:
    if path.is_dir():
        return merge_spreadsheet_rows(spreadsheet_files(path))

    kind = spreadsheet_kind(path)
    if kind == "zip":
        return merge_spreadsheet_rows(spreadsheet_files(extract_spreadsheet_archive(path)))
    return read_single_spreadsheet_rows(path)


def write_simple_csv(path: Path, headers: List[str], rows: List[List[Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)
