#!/usr/bin/env python3
"""Single-file AI entrypoint for FM onboarding.

This file contains the FM onboarding validation logic, pincode insert support,
report generation, downstream Jenkins trigger logic, and AI/ticket response output.
It intentionally does not contain any password, token, or user secret.

Run:
    python ai_fm_onboarding_single.py /path/to/ops_file.xlsx

Required runtime secrets are read from environment variables:
    PINCODE_LOGIN_USERNAME, PINCODE_LOGIN_PASSWORD, JENKINS_USER, JENKINS_TOKEN
"""
from __future__ import annotations




# ==============================================================================
# validation_utils.py
# ==============================================================================

import os
import re
from datetime import datetime
from typing import Any

TRUE_VALUES = {"1", "true", "yes", "y"}
INACTIVE_VALUES = {"0", "false", "no", "n", "inactive"}
LOCATION_DATE_FORMATS = ("%d-%m-%Y, %H:%M", "%d-%m-%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S")


def clean_text(value: Any) -> str:
    return "" if value is None else str(value).replace("\u00a0", " ").strip()


def clean_digits(value: Any) -> str:
    return re.sub(r"\D", "", clean_text(value))


def first_comma_value(value: Any) -> str:
    return clean_text(value).split(",", 1)[0].strip()


def prod_branch_admin_name(value: Any) -> str:
    return " ".join(clean_text(value).split()[:2])


def normalize_key(value: str) -> str:
    return clean_text(value).lower()


def normalize_words(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()


def parse_int_text(value: Any) -> int | None:
    digits = clean_digits(value)
    return int(digits) if digits else None


def env_flag(name: str, default: str = "0") -> bool:
    return clean_text(os.getenv(name, default)).lower() in TRUE_VALUES


def id_text(value: Any) -> str:
    return clean_text(value).replace(",", "")


def is_active_value(value: Any) -> bool:
    return clean_text(value).lower() not in INACTIVE_VALUES


def parse_location_datetime(value: Any) -> datetime | None:
    text = clean_text(value)
    for date_format in LOCATION_DATE_FORMATS if text else ():
        try:
            return datetime.strptime(text, date_format)
        except ValueError:
            pass
    return None



# ==============================================================================
# spreadsheet_io.py
# ==============================================================================

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



# ==============================================================================
# titan_lookup.py
# ==============================================================================

import os
from typing import Any, Dict, List, Optional, Sequence



DEFAULT_TITAN_CITIES_QUERY = """
SELECT id, name, state_id, priority, is_active, uber_city_name
FROM ls_silver.titan__cities
"""
DEFAULT_TITAN_CITIES_TABLE = "ls_silver.titan__cities"
DEFAULT_TITAN_STATES_TABLE = "ls_silver.titan__states"
DEFAULT_TITAN_STATE_QUERIES = [
    "SELECT id, name, is_active FROM ls_silver.titan__states",
    "SELECT id, name, is_active FROM ls_silver.titan__state",
]
DEFAULT_TRINO_MAX_ATTEMPTS = 1


def env_float(*names: str) -> Optional[float]:
    for name in names:
        value = clean_text(os.getenv(name))
        if not value:
            continue
        try:
            return float(value)
        except ValueError:
            continue
    return None


def env_int(*names: str, default: int) -> int:
    for name in names:
        value = clean_text(os.getenv(name))
        if not value:
            continue
        try:
            return int(value)
        except ValueError:
            continue
    return default


def first_present(row: Dict[str, Any], *keys: str, default: Any = "") -> Any:
    for key in keys:
        value = row.get(key)
        if value is not None:
            return value
    return default


class TrinoTitanLookup:
    def __init__(self) -> None:
        self.states: List[Dict[str, str]] = []
        self.states_by_id: Dict[str, Dict[str, str]] = {}
        self.cities: List[Dict[str, str]] = []
        self.cities_by_id: Dict[str, Dict[str, str]] = {}
        self._state_by_name_cache: Dict[str, Optional[Dict[str, str]]] = {}
        self._first_city_by_state_cache: Dict[str, Optional[Dict[str, str]]] = {}

    @property
    def cities_table(self) -> str:
        return clean_text(os.getenv("TITAN_TRINO_CITIES_TABLE")) or DEFAULT_TITAN_CITIES_TABLE

    @property
    def states_table(self) -> str:
        return clean_text(os.getenv("TITAN_TRINO_STATES_TABLE")) or DEFAULT_TITAN_STATES_TABLE

    @staticmethod
    def connect():
        try:
            from trino.dbapi import connect
        except ImportError as exc:
            raise SystemExit("Python package 'trino' is required. Install with: pip install trino") from exc

        kwargs = {
            "host": os.getenv("TITAN_TRINO_HOST", "os-trino-ls-net1.prd.meesho.int"),
            "port": int(os.getenv("TITAN_TRINO_PORT", "80")),
            "user": os.getenv("TITAN_TRINO_USER", "log10_scripts"),
            "catalog": os.getenv("TITAN_TRINO_CATALOG", "delta"),
            "schema": os.getenv("TITAN_TRINO_SCHEMA", "ls_silver"),
            "max_attempts": env_int(
                "TRINO_MAX_ATTEMPTS",
                "TITAN_TRINO_MAX_ATTEMPTS",
                default=DEFAULT_TRINO_MAX_ATTEMPTS,
            ),
        }
        request_timeout = env_float("TRINO_REQUEST_TIMEOUT_SECONDS", "TITAN_TRINO_REQUEST_TIMEOUT_SECONDS")
        if request_timeout is not None:
            kwargs["request_timeout"] = request_timeout
        query_max_run_time = clean_text(
            os.getenv("TRINO_QUERY_MAX_RUN_TIME") or os.getenv("TITAN_TRINO_QUERY_MAX_RUN_TIME")
        )
        if query_max_run_time:
            kwargs["session_properties"] = {"query_max_run_time": query_max_run_time}
        return connect(**kwargs)

    @staticmethod
    def fetch_rows(query: str) -> List[Dict[str, Any]]:
        conn = TrinoTitanLookup.connect()
        try:
            cursor = conn.cursor()
            try:
                cursor.execute(query)
                columns = [clean_text(column[0]).lower() for column in cursor.description]
                return [dict(zip(columns, row)) for row in cursor.fetchall()]
            except Exception as exc:
                compact_query = " ".join(query.split())
                raise RuntimeError(f"Trino query failed: {compact_query}: {exc}") from exc
        finally:
            conn.close()

    @staticmethod
    def fetch_first_working(queries: Sequence[str]) -> List[Dict[str, Any]]:
        errors = []
        for query in queries:
            try:
                return TrinoTitanLookup.fetch_rows(query)
            except Exception as exc:
                errors.append(f"{query}: {exc}")
        raise RuntimeError("Unable to fetch TITAN state data from Trino. Tried: " + " | ".join(errors))

    def fetch_state_rows(self) -> List[Dict[str, Any]]:
        query = os.getenv("TITAN_TRINO_STATE_QUERY", "")
        return self.fetch_rows(query) if query else self.fetch_first_working(DEFAULT_TITAN_STATE_QUERIES)

    def fetch_city_rows(self) -> List[Dict[str, Any]]:
        return self.fetch_rows(os.getenv("TITAN_TRINO_CITIES_QUERY", DEFAULT_TITAN_CITIES_QUERY))

    def parse_states(self, rows: List[Dict[str, Any]]) -> List[Dict[str, str]]:
        states = []
        for row in rows:
            state = {
                "id": id_text(row.get("id")),
                "name": clean_text(row.get("name") or row.get("state_name") or row.get("state")),
                "is_active": clean_text(first_present(row, "is_active", "active", default="TRUE")),
            }
            if state["id"] and state["name"]:
                states.append(state)
        return states

    def parse_cities(self, rows: List[Dict[str, Any]]) -> List[Dict[str, str]]:
        cities = []
        for row in rows:
            state_id = id_text(row.get("state_id") or row.get("stateid"))
            state = self.states_by_id.get(state_id, {})
            city = {
                "id": id_text(row.get("id")),
                "name": clean_text(row.get("name") or row.get("city_name") or row.get("city")),
                "state_id": state_id,
                "priority": id_text(row.get("priority")) or "999999",
                "is_active": clean_text(first_present(row, "is_active", "active", default="TRUE")),
                "state_name": clean_text(row.get("state_name") or row.get("state") or state.get("name")),
                "uber_city_name": clean_text(row.get("uber_city_name")),
            }
            if city["id"]:
                cities.append(city)
        return cities

    def active_states(self) -> List[Dict[str, str]]:
        return [dict(state) for state in self.states if is_active_value(state.get("is_active"))]

    def active_cities(self) -> List[Dict[str, str]]:
        return [dict(city) for city in self.cities if is_active_value(city.get("is_active"))]

    def state_by_id(self, state_id: str) -> Optional[Dict[str, str]]:
        key = id_text(state_id)
        if not key:
            return None
        if key in self.states_by_id:
            return dict(self.states_by_id[key])

        rows = self.fetch_rows(
            f"""
            SELECT id, name, is_active
            FROM {self.states_table}
            WHERE CAST(id AS VARCHAR) = '{key}'
            LIMIT 1
            """
        )
        states = self.parse_states(rows)
        if not states:
            return None
        self.states_by_id[key] = states[0]
        return dict(states[0])

    def state_by_name(self, names: Sequence[str]) -> Optional[Dict[str, str]]:
        cleaned_names = [clean_text(name) for name in names if clean_text(name)]
        if not cleaned_names:
            return None
        cache_key = "|".join(sorted(name.lower() for name in cleaned_names))
        if cache_key in self._state_by_name_cache:
            cached = self._state_by_name_cache[cache_key]
            return dict(cached) if cached else None

        quoted_names = ", ".join("'" + name.lower().replace("'", "''") + "'" for name in cleaned_names)
        rows = self.fetch_rows(
            f"""
            SELECT id, name, is_active
            FROM {self.states_table}
            WHERE lower(trim(CAST(name AS VARCHAR))) IN ({quoted_names})
            ORDER BY id
            LIMIT 1
            """
        )
        states = self.parse_states(rows)
        state = states[0] if states else None
        if state and state.get("id"):
            self.states_by_id[state["id"]] = state
        self._state_by_name_cache[cache_key] = state
        return dict(state) if state else None

    def city_by_id(self, city_id: str) -> Optional[Dict[str, str]]:
        key = id_text(city_id)
        if not key:
            return None
        if key in self.cities_by_id:
            return dict(self.cities_by_id[key])

        rows = self.fetch_rows(
            f"""
            SELECT id, name, state_id, priority, is_active, uber_city_name
            FROM {self.cities_table}
            WHERE CAST(id AS VARCHAR) = '{key}'
            LIMIT 1
            """
        )
        cities = self.parse_cities(rows)
        if not cities:
            return None
        city = cities[0]
        if city.get("state_id") and not city.get("state_name"):
            state = self.state_by_id(city["state_id"])
            if state:
                city["state_name"] = clean_text(state.get("name"))
        self.cities_by_id[key] = city
        return dict(city)

    def pick_active_city(self, state_id: str) -> Optional[Dict[str, str]]:
        key = id_text(state_id)
        if not key:
            return None
        if key in self._first_city_by_state_cache:
            cached = self._first_city_by_state_cache[key]
            return dict(cached) if cached else None

        rows = self.fetch_rows(
            f"""
            SELECT id, name, state_id, priority, is_active, uber_city_name
            FROM {self.cities_table}
            WHERE CAST(state_id AS VARCHAR) = '{key}'
              AND UPPER(COALESCE(CAST(is_active AS VARCHAR), 'TRUE')) IN ('1', 'TRUE', 'YES', 'Y')
            ORDER BY
              COALESCE(try_cast(priority AS INTEGER), 999999),
              COALESCE(try_cast(id AS INTEGER), 999999)
            LIMIT 1
            """
        )
        cities = self.parse_cities(rows)
        city = cities[0] if cities else None
        if city:
            state = self.state_by_id(key)
            if state:
                city["state_name"] = clean_text(state.get("name"))
            if city.get("id"):
                self.cities_by_id[city["id"]] = city
        self._first_city_by_state_cache[key] = city
        return dict(city) if city else None



# ==============================================================================
# pincode_service.py
# ==============================================================================

import json
import os
import re
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple



PINCODE_RE = re.compile(r"^\d{6}$")
JENKINS_CITY_ID_RULES = (
    ({11}, "188"),
    ({22}, "45"),
    ({56, 57, 58}, "1"),
    ({60, 61, 62, 63, 64}, "2075"),
    ({67, 68, 69}, "2471"),
)


def split_pickup_pincodes(value: str) -> Tuple[List[str], List[str]]:
    errors = []
    pincodes = []
    seen = set()
    for part in re.split(r"[,;\n]", clean_text(value)):
        cleaned = clean_digits(part)
        if not cleaned:
            continue
        if not PINCODE_RE.match(cleaned):
            errors.append(f"Invalid pickup pincode: {part}")
            continue
        if cleaned not in seen:
            seen.add(cleaned)
            pincodes.append(cleaned)
    if not pincodes:
        errors.append("pickupPincodes has no valid 6 digit pincode")
    return pincodes, errors


def jenkins_city_id_for_pincode(pincode: str) -> str:
    code = int(str(pincode)[:2])
    for prefixes, city_id in JENKINS_CITY_ID_RULES:
        if code in prefixes:
            return city_id
    return "454"


def duplicate_insert_message(value: Any) -> bool:
    text = clean_text(value).lower()
    return "duplicate" in text and ("pincode" in text or "zipcode" in text or "entry" in text)


class PincodeAuthError(RuntimeError):
    pass


class PincodeInsertApiClient:
    """Calls the prod mutation API for pincode inserts using a fresh Hydra login."""

    def __init__(
        self,
        login_url: str,
        mutation_url: str,
        username: str,
        password: str,
        device_id: str,
        timeout_seconds: float = 30.0,
    ) -> None:
        self.login_url = login_url
        self.mutation_url = mutation_url
        self.username = username
        self.password = password
        self.device_id = device_id
        self.timeout_seconds = timeout_seconds
        self.access_token = ""
        self.token_id = ""

    @classmethod
    def from_env(cls) -> Optional["PincodeInsertApiClient"]:
        if not env_flag("PINCODE_INSERT_API_ENABLED"):
            return None

        login_url = clean_text(os.getenv("PINCODE_LOGIN_URL"))
        mutation_url = clean_text(os.getenv("PINCODE_INSERT_API_URL"))
        username = clean_text(
            os.getenv("PINCODE_LOGIN_USERNAME")
            or os.getenv("HYDRA_LOGIN_USERNAME")
            or os.getenv("Username")
            or os.getenv("USERNAME")
        )
        password = clean_text(
            os.getenv("PINCODE_LOGIN_PASSWORD")
            or os.getenv("HYDRA_LOGIN_PASSWORD")
            or os.getenv("Password")
            or os.getenv("PASSWORD")
        )
        device_id = clean_text(os.getenv("PINCODE_LOGIN_DEVICE_ID")) or "123123123123"
        timeout = float(clean_text(os.getenv("PINCODE_API_TIMEOUT_SECONDS")) or "30")

        missing = []
        for name, value in [
            ("PINCODE_LOGIN_URL", login_url),
            ("PINCODE_INSERT_API_URL", mutation_url),
            ("PINCODE_LOGIN_USERNAME", username),
            ("PINCODE_LOGIN_PASSWORD", password),
        ]:
            if not value:
                missing.append(name)
        if missing:
            raise ValueError(
                "PINCODE_INSERT_API_ENABLED=true but missing env vars: "
                + ", ".join(missing)
            )

        return cls(
            login_url=login_url,
            mutation_url=mutation_url,
            username=username,
            password=password,
            device_id=device_id,
            timeout_seconds=timeout,
        )

    def login(self) -> None:
        response = self.post_json(
            self.login_url,
            {
                "username": self.username,
                "password": self.password,
            },
            {
                "Content-Type": "application/json",
                "deviceId": self.device_id,
            },
        )
        status = response.get("status") or {}
        status_code = parse_int_text(status.get("code"))
        if status_code != 200:
            raise RuntimeError(
                "Hydra login failed: " + clean_text(status.get("message") or response)
            )

        token = response.get("response", {}).get("token", {})
        access_token = clean_text(token.get("accessToken"))
        token_id = clean_text(token.get("tokenId"))
        if not access_token or not token_id:
            raise RuntimeError("Hydra login response did not contain accessToken/tokenId")

        self.access_token = access_token
        self.token_id = token_id

    def insert_pincode(self, parsed: Dict[str, str]) -> Dict[str, Any]:
        if not self.access_token or not self.token_id:
            self.login()

        try:
            return self.insert_pincode_once(parsed)
        except PincodeAuthError:
            self.login()
            return self.insert_pincode_once(parsed)

    def insert_pincode_once(self, parsed: Dict[str, str]) -> Dict[str, Any]:
        payload = {
            "operation": "INSERT",
            "tableName": "pincodes",
            "values": {
                "zipcode": int(parsed["zipcode"]),
                "city_id": int(parsed["city_id"]),
                "status": int(parsed["status"]),
            },
        }
        response = self.post_json(
            self.mutation_url,
            payload,
            {
                "Accept": "*/*",
                "Cache-Control": "no-cache",
                "Content-Type": "application/json",
                "token": self.access_token,
                "tokenId": self.token_id,
            },
        )

        status = response.get("status") or {}
        status_code = parse_int_text(status.get("code"))
        status_message = clean_text(status.get("message"))
        if status_code in {401, 403}:
            raise PincodeAuthError(status_message or "Mutation API token expired")

        api_response = response.get("response") or {}
        rows_affected = parse_int_text(api_response.get("rowsAffected")) or 0
        response_message = clean_text(api_response.get("message") or status_message)
        if status_code != 200 or rows_affected < 1:
            if duplicate_insert_message(response_message):
                return {
                    "status_code": status_code,
                    "message": response_message,
                    "rows_affected": rows_affected,
                    "zipcode": parsed["zipcode"],
                    "city_id": parsed["city_id"],
                    "duplicate": True,
                }
            raise RuntimeError(
                "Pincode insert API failed: "
                + (response_message or str(response))
            )

        return {
            "status_code": status_code,
            "message": response_message,
            "rows_affected": rows_affected,
            "zipcode": parsed["zipcode"],
            "city_id": parsed["city_id"],
            "duplicate": False,
        }

    def post_json(
        self,
        url: str,
        payload: Dict[str, Any],
        headers: Dict[str, str],
    ) -> Dict[str, Any]:
        request = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers=headers,
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                body = response.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            if exc.code in {401, 403}:
                raise PincodeAuthError(f"HTTP {exc.code}: {body}") from exc
            if duplicate_insert_message(body):
                return {
                    "status": {"code": exc.code, "message": body},
                    "response": {"rowsAffected": 0, "message": body, "duplicate": True},
                }
            raise RuntimeError(f"HTTP {exc.code} from {url}: {body}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"Unable to call {url}: {exc.reason}") from exc

        try:
            parsed = json.loads(body)
        except json.JSONDecodeError as exc:
            raise RuntimeError(f"Invalid JSON response from {url}: {body}") from exc
        if not isinstance(parsed, dict):
            raise RuntimeError(f"Unexpected response from {url}: {parsed}")
        return parsed


class RecentPincodeCache:
    """Small Jenkins-side cache for the Trino ingestion delay after API inserts."""

    def __init__(self, path: Path, ttl_seconds: int) -> None:
        self.path = path
        self.ttl_seconds = ttl_seconds

    @classmethod
    def from_env(cls) -> Optional["RecentPincodeCache"]:
        path_text = clean_text(os.getenv("PINCODE_RECENT_CACHE_PATH"))
        if not path_text and clean_text(os.getenv("JENKINS_HOME")):
            path_text = str(
                Path(clean_text(os.getenv("JENKINS_HOME")))
                / "fm-onboard-pincode-cache"
                / "recent_pincodes.json"
            )
        if not path_text:
            return None

        ttl_minutes = parse_int_text(os.getenv("PINCODE_RECENT_CACHE_TTL_MINUTES")) or 180
        return cls(Path(path_text), ttl_minutes * 60)

    @staticmethod
    def allocation_key(context_key: str, input_pincode: str) -> str:
        return f"{clean_text(context_key).lower()}|{clean_digits(input_pincode)}"

    def load(self) -> Dict[str, Any]:
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except Exception:
            return {"allocations": []}
        if not isinstance(data, dict):
            return {"allocations": []}
        allocations = data.get("allocations")
        if not isinstance(allocations, list):
            data["allocations"] = []
        return data

    def active_allocations(self) -> List[Dict[str, Any]]:
        now = time.time()
        rows = []
        for row in self.load().get("allocations", []):
            if not isinstance(row, dict):
                continue
            inserted_at = float(row.get("inserted_at") or 0)
            if inserted_at and now - inserted_at <= self.ttl_seconds:
                rows.append(row)
        return rows

    def save(self, allocations: List[Dict[str, Any]]) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            temp_path = self.path.with_suffix(self.path.suffix + ".tmp")
            temp_path.write_text(
                json.dumps({"allocations": allocations}, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            temp_path.replace(self.path)
        except OSError:
            return

    def get_allocation(self, context_key: str, input_pincode: str) -> Optional[Dict[str, Any]]:
        key = self.allocation_key(context_key, input_pincode)
        for row in self.active_allocations():
            if clean_text(row.get("key")) == key:
                return dict(row)
        return None

    def reserved_pincodes_between(
        self,
        min_pin: int,
        max_pin: int,
        exclude_context_key: str = "",
    ) -> set[int]:
        excluded = clean_text(exclude_context_key).lower()
        reserved = set()
        for row in self.active_allocations():
            if excluded and clean_text(row.get("context_key")).lower() == excluded:
                continue
            zipcode = parse_int_text(row.get("output_pincode"))
            if zipcode is not None and int(min_pin) <= zipcode <= int(max_pin):
                reserved.add(zipcode)
        return reserved

    def remember(
        self,
        *,
        context_key: str,
        input_pincode: str,
        parsed: Dict[str, str],
        api_response: Optional[Dict[str, Any]],
    ) -> None:
        key = self.allocation_key(context_key, input_pincode)
        allocations = [row for row in self.active_allocations() if clean_text(row.get("key")) != key]
        allocations.append(
            {
                "key": key,
                "context_key": clean_text(context_key),
                "input_pincode": clean_digits(input_pincode),
                "output_pincode": clean_digits(parsed.get("zipcode")),
                "city_id": clean_text(parsed.get("city_id")),
                "status": clean_text(parsed.get("status") or "1"),
                "service_provider": clean_text(parsed.get("service_provider")),
                "inserted_at": time.time(),
                "insert_api_called": api_response is not None,
                "insert_api_duplicate": bool((api_response or {}).get("duplicate")),
                "insert_api_message": clean_text((api_response or {}).get("message")),
            }
        )
        self.save(allocations)


class PincodeResolver:
    """Pincode Jenkins selection logic backed by live VALMO lookup data."""

    def __init__(
        self,
        valmo_lookup: Any,
        insert_api_client: Optional["PincodeInsertApiClient"] = None,
        recent_cache: Optional[RecentPincodeCache] = None,
    ) -> None:
        self.valmo_lookup = valmo_lookup
        self.insert_api_client = insert_api_client
        self.recent_cache = recent_cache

    def close(self) -> None:
        return None

    def active_db_pincodes(self) -> set[int]:
        return self.valmo_lookup.active_pincode_numbers()

    def active_db_pincodes_between(self, min_pin: int, max_pin: int) -> set[int]:
        return self.valmo_lookup.active_pincode_numbers_between(min_pin, max_pin)

    def used_pincodes(self) -> set[int]:
        return self.valmo_lookup.used_pincode_numbers()

    def used_pincodes_between(self, min_pin: int, max_pin: int) -> set[int]:
        return self.valmo_lookup.used_pincode_numbers_between(min_pin, max_pin)

    def resolve(
        self,
        input_pincode: str,
        batch_used: set[int],
        context_key: str = "",
    ) -> Dict[str, Any]:
        pin = int(input_pincode)
        min_pin = pin - 100
        max_pin = pin + 100

        context = clean_text(context_key) or clean_digits(input_pincode)
        recent_allocation = (
            self.recent_cache.get_allocation(context, input_pincode)
            if self.recent_cache is not None
            else None
        )
        if recent_allocation:
            output_pin = int(clean_digits(recent_allocation.get("output_pincode")))
            batch_used.add(output_pin)
            self.valmo_lookup.cache_pincode(
                {
                    "zipcode": str(output_pin),
                    "city_id": clean_text(recent_allocation.get("city_id")),
                    "status": clean_text(recent_allocation.get("status") or "1"),
                    "service_provider": clean_text(recent_allocation.get("service_provider")),
                }
            )
            return {
                "input_pincode": str(input_pincode),
                "output_pincode": str(output_pin),
                "insert_sql": "",
                "inserted": False,
                "insert_api_called": False,
                "insert_api_response": None,
                "reused_recent_cache": True,
            }

        db_pincodes = self.active_db_pincodes_between(min_pin, max_pin)
        used_pincodes = self.used_pincodes_between(min_pin, max_pin)
        if self.recent_cache is not None:
            reserved = self.recent_cache.reserved_pincodes_between(
                min_pin,
                max_pin,
                exclude_context_key=context,
            )
            db_pincodes.update(reserved)
            used_pincodes.update(reserved)

        last_duplicate = ""
        for output_pin, insert_sql in self.new_pincode_candidates(
            pin,
            db_pincodes,
            used_pincodes,
            batch_used,
        ):
            parsed = self.parse_insert_sql(insert_sql)
            insert_api_response: Optional[Dict[str, Any]] = None
            inserted = False
            if self.insert_api_client is not None:
                insert_api_response = self.insert_api_client.insert_pincode(parsed)
                if insert_api_response.get("duplicate"):
                    db_pincodes.add(output_pin)
                    last_duplicate = clean_text(insert_api_response.get("message"))
                    continue
                inserted = True

            batch_used.add(output_pin)
            self.valmo_lookup.cache_pincode(parsed)
            if self.recent_cache is not None and self.insert_api_client is not None:
                self.recent_cache.remember(
                    context_key=context,
                    input_pincode=input_pincode,
                    parsed=parsed,
                    api_response=insert_api_response,
                )

            return {
                "input_pincode": str(input_pincode),
                "output_pincode": str(output_pin),
                "insert_sql": insert_sql,
                "inserted": inserted,
                "insert_api_called": insert_api_response is not None,
                "insert_api_response": insert_api_response,
                "reused_recent_cache": False,
            }

        reason = f": {last_duplicate}" if last_duplicate else ""
        raise ValueError(f"No new pincode found near {input_pincode}{reason}")

    def new_pincode_candidates(
        self,
        pin: int,
        db_pincodes: set[int],
        used_pincodes: set[int],
        batch_used: set[int],
    ) -> List[Tuple[int, str]]:
        min_pin = pin - 100
        max_pin = pin + 100
        candidates = []
        seen: set[int] = set()
        for offset in range(0, 101):
            for candidate in [pin + offset, pin - offset]:
                if candidate in seen:
                    continue
                seen.add(candidate)
                if (
                    min_pin <= candidate <= max_pin
                    and candidate not in db_pincodes
                    and candidate not in used_pincodes
                    and candidate not in batch_used
                ):
                    city_id = jenkins_city_id_for_pincode(str(candidate))
                    candidates.append(
                        (
                            candidate,
                            f"INSERT INTO pincodes VALUES "
                            f"({candidate}, {city_id}, 1, 'RELIANCE');",
                        )
                    )
        return candidates

    def parse_insert_sql(self, insert_sql: str) -> Dict[str, str]:
        match = re.match(
            r"^\s*INSERT\s+INTO\s+pincodes\s+VALUES\s*\(\s*"
            r"(?P<zipcode>\d{6})\s*,\s*"
            r"(?P<city_id>\d+)\s*,\s*"
            r"(?P<status>[01])\s*,\s*"
            r"'(?P<service_provider>[^']+)'\s*"
            r"\)\s*;?\s*$",
            insert_sql,
            re.IGNORECASE,
        )
        if not match:
            raise ValueError(f"Unsupported Jenkins insert SQL: {insert_sql}")
        return match.groupdict()

    def pincode_exists(self, zipcode: str) -> bool:
        return self.valmo_lookup.pincode_exists(zipcode)



# ==============================================================================
# valmo_lookup.py
# ==============================================================================

import difflib
import os
import re
from typing import Any, Dict, List, Optional, Tuple


SERVICE_SUFFIX_NORMALIZATIONS = {"services": "service"}
ACTIVE_LOCATION_STATUS_SQL = "UPPER(COALESCE(status, '1')) IN ('1', 'TRUE', 'YES', 'Y')"
DEFAULT_VALMO_PARTNERS_QUERY = """
SELECT id, name, status, updated_at
FROM ls_silver.logten__partners
"""
DEFAULT_VALMO_PARTNERS_TABLE = "ls_silver.logten__partners"
DEFAULT_VALMO_PINCODES_QUERY = """
SELECT id, zipcode, city_id, status, service_provider, updated_at
FROM ls_silver.logten__pincodes
"""
DEFAULT_VALMO_LOCATIONS_TABLE = "ls_silver.logten__locations"
DEFAULT_VALMO_PINCODES_TABLE = "ls_silver.logten__pincodes"

CITY_ALIASES_BY_NAME = {
    "Allahabad": ["prayagraj"],
    "Bengaluru": ["bangalore"],
    "Belgaum": ["belagavi"],
    "Chennai": ["madras"],
    "Cochin": ["kochi"],
    "Gurgaon": ["gurugram"],
    "Kolkata": ["calcutta"],
    "Mangalore": ["mangaluru"],
    "Mumbai": ["bombay"],
    "Mysore": ["mysuru"],
    "Trivandrum": ["thiruvananthapuram"],
    "Vadodara": ["baroda"],
}

STATE_ALIASES_BY_NAME = {
    "Andaman and Nicobar Islands": ["andaman", "nicobar"],
    "Andhra Pradesh": ["andhra", "ap"],
    "Arunachal Pradesh": ["arunachal"],
    "Chhattisgarh": ["chattisgarh", "cg"],
    "Dadra and Nagar Haveli": ["dadra", "nagar haveli"],
    "Delhi": ["new delhi", "ncr"],
    "Gujarat": ["gj"],
    "Haryana": ["hr"],
    "Himachal Pradesh": ["himachal", "hp"],
    "Jammu and Kashmir": ["jammu", "kashmir", "j k"],
    "Karnataka": ["ka"],
    "Madhya Pradesh": ["madhya", "mp"],
    "Maharashtra": ["mh"],
    "Orissa": ["odisha"],
    "Puducherry": ["pondicherry"],
    "Punjab": ["pb"],
    "Rajasthan": ["rj"],
    "Tamil Nadu": ["tn"],
    "Telengana": ["telangana", "ts"],
    "Uttar Pradesh": ["uttar pradesh", "up"],
    "Uttarakhand": ["uk"],
    "West Bengal": ["wb"],
}

PINCODE_PREFIX_STATE_RANGES = (
    (11, 11, "Delhi"),
    (12, 13, "Haryana"),
    (14, 16, "Punjab"),
    (17, 17, "Himachal Pradesh"),
    (18, 19, "Jammu and Kashmir"),
    (20, 28, "Uttar Pradesh"),
    (30, 34, "Rajasthan"),
    (36, 39, "Gujarat"),
    (40, 44, "Maharashtra"),
    (45, 48, "Madhya Pradesh"),
    (49, 49, "Chhattisgarh"),
    (50, 50, "Telengana"),
    (51, 53, "Andhra Pradesh"),
    (56, 59, "Karnataka"),
    (60, 64, "Tamil Nadu"),
    (67, 69, "Kerala"),
    (70, 74, "West Bengal"),
    (75, 77, "Orissa"),
    (78, 79, "Assam"),
    (80, 81, "Bihar"),
    (82, 83, "Jharkhand"),
    (84, 85, "Bihar"),
)
PINCODE_PREFIX_STATE_NAMES = {
    str(prefix): state
    for start, end, state in PINCODE_PREFIX_STATE_RANGES
    for prefix in range(start, end + 1)
}
DEFAULT_PARTNER_ID_OVERRIDES = {
    "sslogictic": "130544",
    "sslogistic": "130544",
    "sslogistics": "130544",
}


def canonical_partner_name(value: Any) -> str:
    tokens = partner_name_tokens(value)
    return "".join(tokens)


def compact_partner_name(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_text(value).lower())


def partner_name_tokens(value: Any) -> List[str]:
    tokens = re.findall(r"[a-z0-9]+", clean_text(value).lower())
    if not tokens:
        return []
    tokens[-1] = SERVICE_SUFFIX_NORMALIZATIONS.get(tokens[-1], tokens[-1])
    return tokens


def partner_id_overrides() -> Dict[str, str]:
    overrides = dict(DEFAULT_PARTNER_ID_OVERRIDES)
    raw = clean_text(os.getenv("PARTNER_ID_OVERRIDES"))
    for item in [part.strip() for part in raw.split(",") if part.strip()]:
        separator = "=" if "=" in item else ":" if ":" in item else ""
        if not separator:
            continue
        name, partner_id = item.split(separator, 1)
        key = canonical_partner_name(name)
        if key and clean_text(partner_id):
            overrides[key] = clean_text(partner_id)
    return overrides


def partner_lookup_keys(value: Any) -> List[str]:
    keys = {compact_partner_name(value), canonical_partner_name(value)}
    return [key for key in keys if key]


def aliases_for(name: Any, alias_map: Dict[str, List[str]]) -> List[str]:
    base = clean_text(name)
    return [alias for alias in [base, *alias_map.get(base, [])] if clean_text(alias)]


def best_alias_match(
    text: str,
    records: List[Dict[str, str]],
    alias_map: Dict[str, List[str]],
    min_alias_len: int,
) -> Optional[Tuple[str, Dict[str, str]]]:
    words = normalize_words(text)
    if not words:
        return None

    haystack = f" {words} "
    matches: List[Tuple[int, str, Dict[str, str]]] = []
    for record in records:
        record_aliases = record.get("name_aliases")
        if not isinstance(record_aliases, list):
            record_aliases = aliases_for(record.get("name"), alias_map)
        for alias in record_aliases:
            alias_words = normalize_words(alias)
            if len(alias_words) >= min_alias_len and f" {alias_words} " in haystack:
                matches.append((len(alias_words), alias, record))
                break
    if not matches:
        return None
    _, alias, record = max(matches, key=lambda item: item[0])
    return alias, record


def city_aliases(city: Dict[str, str]) -> List[str]:
    aliases = aliases_for(city.get("name"), CITY_ALIASES_BY_NAME)
    uber_city_name = clean_text(city.get("uber_city_name"))
    return aliases + ([uber_city_name] if uber_city_name else [])


def one_character_spelling_match(left: str, right: str) -> bool:
    if left == right:
        return True
    if min(len(left), len(right)) < 8 or abs(len(left) - len(right)) > 1:
        return False

    if len(left) == len(right):
        return sum(1 for a, b in zip(left, right) if a != b) == 1

    longer, shorter = (left, right) if len(left) > len(right) else (right, left)
    i = j = edits = 0
    while i < len(longer) and j < len(shorter):
        if longer[i] == shorter[j]:
            i += 1
            j += 1
        else:
            edits += 1
            if edits > 1:
                return False
            i += 1
    return True


def token_spelling_match(left: Any, right: Any) -> bool:
    left_tokens = partner_name_tokens(left)
    right_tokens = partner_name_tokens(right)
    if len(left_tokens) != len(right_tokens) or not left_tokens:
        return False

    has_difference = False
    for left_token, right_token in zip(left_tokens, right_tokens):
        if left_token == right_token:
            continue
        has_difference = True
        if min(len(left_token), len(right_token)) < 5:
            return False
        if difflib.SequenceMatcher(None, left_token, right_token).ratio() < 0.82:
            return False
    return has_difference


def compact_subset_match(left: Any, right: Any) -> bool:
    left_key = compact_partner_name(left)
    right_key = compact_partner_name(right)
    if left_key == right_key:
        return True
    shorter, longer = sorted([left_key, right_key], key=len)
    return len(shorter) >= 6 and shorter in longer


def partner_match_type(partner_name: str, target: str, candidate_name: Any) -> str:
    candidate_key = canonical_partner_name(candidate_name)
    if candidate_key == target:
        return "exact"
    if compact_subset_match(partner_name, candidate_name):
        return "compact_subset"
    if one_character_spelling_match(target, candidate_key):
        return "one_character_spelling"
    return "token_spelling" if token_spelling_match(partner_name, candidate_name) else ""


def summarize_partner_matches(matches: List[Dict[str, str]]) -> str:
    return "; ".join(
        f"id={clean_text(match.get('id'))}, name={clean_text(match.get('name'))}, "
        f"status={clean_text(match.get('status'))}"
        for match in matches
    )


def sql_quote(value: str) -> str:
    return "'" + clean_text(value).replace("'", "''") + "'"


def sql_in(values: List[str]) -> str:
    return ", ".join(sql_quote(value) for value in values)


def active_sql(alias: str = "") -> str:
    prefix = f"{alias}." if alias else ""
    return (
        f"UPPER(COALESCE(CAST({prefix}status AS VARCHAR), '1')) "
        "IN ('1', 'TRUE', 'YES', 'Y')"
    )


def text_sql(column: str) -> str:
    return f"LOWER(TRIM(COALESCE(CAST({column} AS VARCHAR), '')))"


class ValmoValidator:
    def __init__(self, titan: TrinoTitanLookup) -> None:
        self.titan = titan
        self._partners: Optional[List[Dict[str, Any]]] = None
        self._pincodes: Optional[List[Dict[str, Any]]] = None
        self._pincode_by_zipcode: Dict[str, Dict[str, Any]] = {}
        self._active_pincode_range_cache: Dict[Tuple[int, int], set[int]] = {}
        self._used_pincode_range_cache: Dict[Tuple[int, int], set[int]] = {}

    def close(self) -> None:
        return None

    @staticmethod
    def fetch_rows(query: str) -> List[Dict[str, Any]]:
        return TrinoTitanLookup.fetch_rows(query)

    @property
    def partners_table(self) -> str:
        return clean_text(os.getenv("VALMO_TRINO_PARTNERS_TABLE")) or DEFAULT_VALMO_PARTNERS_TABLE

    @property
    def locations_table(self) -> str:
        return clean_text(os.getenv("VALMO_TRINO_LOCATIONS_TABLE")) or DEFAULT_VALMO_LOCATIONS_TABLE

    @property
    def pincodes_table(self) -> str:
        return clean_text(os.getenv("VALMO_TRINO_PINCODES_TABLE")) or DEFAULT_VALMO_PINCODES_TABLE

    def partners(self) -> List[Dict[str, Any]]:
        if self._partners is None:
            query = clean_text(os.getenv("VALMO_TRINO_PARTNERS_QUERY")) or DEFAULT_VALMO_PARTNERS_QUERY
            self._partners = self.fetch_rows(query)
        return self._partners

    def partner_candidates(self, partner_name: str) -> List[Dict[str, Any]]:
        raw_name = clean_text(partner_name).lower()
        tokens = partner_name_tokens(partner_name)
        exact_names = {raw_name, " ".join(tokens)}
        if tokens and tokens[-1] == "service":
            exact_names.add(" ".join([*tokens[:-1], "services"]))

        trim_expr = "lower(trim(CAST(name AS VARCHAR)))"
        compact_expr = "regexp_replace(lower(CAST(name AS VARCHAR)), '[^a-z0-9]', '')"
        filters = []
        exact_names = {name for name in exact_names if name}
        if exact_names:
            filters.append(f"{trim_expr} IN ({sql_in(sorted(exact_names))})")
        lookup_keys = partner_lookup_keys(partner_name)
        if lookup_keys:
            filters.append(f"{compact_expr} IN ({sql_in(lookup_keys)})")
        if not filters:
            return self.partners()

        return self.fetch_rows(
            f"""
            SELECT id, name, status, updated_at
            FROM {self.partners_table}
            WHERE {" OR ".join(filters)}
            LIMIT 500
            """
        )

    def pincodes(self) -> List[Dict[str, Any]]:
        if self._pincodes is None:
            query = clean_text(os.getenv("VALMO_TRINO_PINCODES_QUERY")) or DEFAULT_VALMO_PINCODES_QUERY
            self._pincodes = self.fetch_rows(query)
        return self._pincodes

    def partner_exists(self, partner_name: str) -> Optional[Dict[str, str]]:
        target = canonical_partner_name(partner_name)
        if not target:
            return None

        exact_matches: List[Dict[str, str]] = []
        fuzzy_matches: List[Dict[str, str]] = []
        for row in self.partner_candidates(partner_name):
            candidate = {
                "id": clean_text(row.get("id")),
                "name": clean_text(row.get("name")),
                "status": clean_text(row.get("status")),
            }
            match_type = partner_match_type(partner_name, target, candidate.get("name"))
            if not match_type:
                continue

            candidate["normalized_match"] = target
            candidate["match_type"] = match_type
            if match_type == "exact":
                exact_matches.append(candidate)
            else:
                fuzzy_matches.append(candidate)

        if exact_matches:
            return self.resolve_partner_matches(exact_matches, "multiple_exact")
        return self.resolve_partner_matches(fuzzy_matches, "multiple_spelling")

    def resolve_partner_matches(
        self,
        matches: List[Dict[str, str]],
        match_type: str,
    ) -> Optional[Dict[str, str]]:
        if not matches:
            return None
        if len(matches) == 1:
            return matches[0]
        override = self.resolve_partner_by_override(matches)
        if override:
            override["match_type"] = f"{match_type}_resolved_by_partner_override"
            override["match_summary"] = summarize_partner_matches(matches)
            return override
        return self.resolve_partner_by_recent_location(matches, match_type) or {
            "ambiguous_match": "1",
            "match_type": match_type,
            "match_summary": summarize_partner_matches(matches),
        }

    def resolve_partner_by_override(
        self,
        matches: List[Dict[str, str]],
    ) -> Optional[Dict[str, str]]:
        normalized_match = clean_text(matches[0].get("normalized_match"))
        partner_id = partner_id_overrides().get(normalized_match)
        if not partner_id:
            return None
        for match in matches:
            if clean_text(match.get("id")) == partner_id:
                return dict(match)
        return None

    def resolve_partner_by_recent_location(
        self,
        matches: List[Dict[str, str]],
        match_type: str,
    ) -> Optional[Dict[str, str]]:
        if len(matches) != 2:
            return None

        partner_ids = [clean_text(match.get("id")) for match in matches if clean_text(match.get("id"))]
        if len(partner_ids) != 2:
            return None

        rows = self.fetch_rows(
            f"""
            SELECT entity_id, id AS location_id, updated_at, created_at
            FROM {self.locations_table}
            WHERE UPPER(COALESCE(CAST(entity_type AS VARCHAR), '')) = 'PARTNER'
              AND CAST(entity_id AS VARCHAR) IN ({sql_in(partner_ids)})
            """
        )

        latest_by_partner: Dict[str, Tuple[Any, str, str]] = {}
        for row in rows:
            partner_id = clean_text(row["entity_id"])
            updated_at = clean_text(row["updated_at"])
            created_at = clean_text(row["created_at"])
            parsed_at = parse_location_datetime(updated_at) or parse_location_datetime(created_at)
            if not parsed_at:
                continue
            current = latest_by_partner.get(partner_id)
            if current is None or parsed_at > current[0]:
                latest_by_partner[partner_id] = (
                    parsed_at,
                    clean_text(row["location_id"]),
                    updated_at or created_at,
                )

        if len(latest_by_partner) != 2:
            return None

        sorted_latest = sorted(
            latest_by_partner.items(),
            key=lambda item: item[1][0],
            reverse=True,
        )
        if sorted_latest[0][1][0] == sorted_latest[1][1][0]:
            return None

        chosen_partner_id = sorted_latest[0][0]
        chosen = next(
            (dict(match) for match in matches if clean_text(match.get("id")) == chosen_partner_id),
            None,
        )
        if not chosen:
            return None

        chosen["match_type"] = f"{match_type}_resolved_by_recent_location"
        chosen["match_summary"] = summarize_partner_matches(matches)
        chosen["recent_location_id"] = sorted_latest[0][1][1]
        chosen["recent_location_updated_at"] = sorted_latest[0][1][2]
        return chosen

    def location_exists(self, fmcode: str, client_location_name: str = "") -> Optional[Dict[str, str]]:
        candidates = list(
            dict.fromkeys(
                normalized
                for normalized in (normalize_key(fmcode), normalize_key(client_location_name))
                if normalized
            )
        )
        if not candidates:
            return None

        values = sql_in(candidates)
        rows = self.fetch_rows(
            f"""
            SELECT
                id,
                alias,
                client_location_name,
                entity_id,
                entity_type,
                status,
                location_hash
            FROM {self.locations_table}
            WHERE UPPER(COALESCE(CAST(entity_type AS VARCHAR), '')) = 'PARTNER'
              AND (
                {text_sql("alias")} IN ({values})
                OR {text_sql("client_location_name")} IN ({values})
                OR {text_sql("billing_name")} IN ({values})
              )
            ORDER BY
              CASE WHEN {active_sql()} THEN 0 ELSE 1 END,
              id
            LIMIT 1
            """
        )
        return dict(rows[0]) if rows else None

    def pick_active_city(self, state_id: str) -> Optional[Dict[str, str]]:
        return self.titan.pick_active_city(state_id)

    def state_by_name(self, state_name: str) -> Optional[Dict[str, str]]:
        aliases = aliases_for(state_name, STATE_ALIASES_BY_NAME)
        return self.titan.state_by_name(aliases)

    def state_from_address(self, address: str) -> Optional[Dict[str, str]]:
        words = normalize_words(address)
        if not words:
            return None

        haystack = f" {words} "
        matches: List[Tuple[int, str]] = []
        for state_name, aliases in STATE_ALIASES_BY_NAME.items():
            for alias in [state_name, *aliases]:
                alias_words = normalize_words(alias)
                if len(alias_words) >= 2 and f" {alias_words} " in haystack:
                    matches.append((len(alias_words), state_name))
                    break
        if not matches:
            return None
        _, state_name = max(matches, key=lambda item: item[0])
        return self.state_by_name(state_name)

    def state_from_pincode(self, pincode: str) -> Optional[Dict[str, str]]:
        if not PINCODE_RE.match(clean_digits(pincode)):
            return None
        state_name = PINCODE_PREFIX_STATE_NAMES.get(clean_digits(pincode)[:2])
        return self.state_by_name(state_name or "")

    def first_city_from_state_fallback(
        self,
        *,
        address: str,
        pincode: str,
    ) -> Optional[Dict[str, str]]:
        state = self.state_from_address(address) or self.state_from_pincode(pincode)
        if not state:
            return None
        city = self.pick_active_city(clean_text(state.get("id")))
        if not city:
            return None
        city["source"] = f"state fallback: first active city for {state.get('name')}"
        return city

    def city_by_id(self, city_id: str) -> Optional[Dict[str, str]]:
        return self.titan.city_by_id(city_id)

    def city_from_address(self, address: str) -> Optional[Dict[str, str]]:
        return None

    def derive_city_from_loczipcode(
        self,
        loczipcode: str,
    ) -> Tuple[Optional[Dict[str, str]], List[str]]:
        warnings: List[str] = []
        pincode_row = self.loczipcode_in_pincodes(loczipcode)
        valmo_city_id = ""
        if pincode_row:
            valmo_city_id = clean_text(pincode_row.get("city_id"))
            if valmo_city_id:
                city = self.city_by_id(valmo_city_id)
                if city and is_active_value(city.get("is_active")):
                    city["source"] = "VALMO.pincodes.city_id"
                    return city, warnings

        fallback_city_id = jenkins_city_id_for_pincode(loczipcode)
        fallback_city = self.city_by_id(fallback_city_id)
        if fallback_city and is_active_value(fallback_city.get("is_active")):
            if valmo_city_id and valmo_city_id != fallback_city_id:
                warnings.append(
                    f"VALMO.pincodes city_id {valmo_city_id} for loczipcode {loczipcode} "
                    f"is not active/present in TITAN.cities; used Jenkins prefix city_id "
                    f"{fallback_city_id}"
                )
            elif not pincode_row:
                warnings.append(
                    f"pincode {loczipcode} is not present in VALMO.pincodes; "
                    f"used Jenkins prefix city_id {fallback_city_id}"
                )
            fallback_city["source"] = "Jenkins pincode prefix"
            return fallback_city, warnings

        return None, warnings

    def loczipcode_in_pincodes(self, loczipcode: str) -> Optional[Dict[str, str]]:
        zipcode = clean_digits(loczipcode)
        if not zipcode:
            return None

        if zipcode in self._pincode_by_zipcode:
            result = dict(self._pincode_by_zipcode[zipcode])
            for column in ("id", "zipcode", "city_id", "status"):
                result[column] = clean_digits(result.get(column))
            return result

        rows = self.fetch_rows(
            f"""
            SELECT id, zipcode, city_id, status, service_provider
            FROM {self.pincodes_table}
            WHERE regexp_replace(CAST(zipcode AS VARCHAR), '[^0-9]', '') = {sql_quote(zipcode)}
            LIMIT 1
            """
        )
        if not rows:
            return None
        result = dict(rows[0])
        self._pincode_by_zipcode[zipcode] = dict(result)
        for column in ("id", "zipcode", "city_id", "status"):
            result[column] = clean_digits(result.get(column))
        return result

    def cache_pincode(self, row: Dict[str, Any]) -> None:
        zipcode = clean_digits(row.get("zipcode"))
        if zipcode:
            self._pincode_by_zipcode[zipcode] = {
                "id": clean_text(row.get("id")),
                "zipcode": zipcode,
                "city_id": clean_text(row.get("city_id")),
                "status": clean_text(row.get("status") or "1"),
                "service_provider": clean_text(row.get("service_provider")),
            }

    def active_pincode_numbers(self) -> set[int]:
        return self.active_pincode_numbers_between(0, 999999)

    def active_pincode_numbers_between(self, min_pin: int, max_pin: int) -> set[int]:
        cache_key = (int(min_pin), int(max_pin))
        if cache_key in self._active_pincode_range_cache:
            return set(self._active_pincode_range_cache[cache_key])

        rows = self.fetch_rows(
            f"""
            SELECT id, zipcode, city_id, status, service_provider
            FROM {self.pincodes_table}
            WHERE {active_sql()}
              AND try_cast(regexp_replace(CAST(zipcode AS VARCHAR), '[^0-9]', '') AS INTEGER)
                BETWEEN {int(min_pin)} AND {int(max_pin)}
            """
        )
        pincodes = set()
        for row in rows:
            zipcode = clean_digits(row.get("zipcode"))
            if not zipcode:
                continue
            self._pincode_by_zipcode[zipcode] = dict(row)
            pincodes.add(int(zipcode))
        self._active_pincode_range_cache[cache_key] = pincodes
        return set(pincodes)

    def pincode_exists(self, zipcode: str) -> bool:
        return self.loczipcode_in_pincodes(zipcode) is not None

    def used_pincode_numbers(self) -> set[int]:
        return self.used_pincode_numbers_between(0, 999999)

    def used_pincode_numbers_between(self, min_pin: int, max_pin: int) -> set[int]:
        cache_key = (int(min_pin), int(max_pin))
        if cache_key in self._used_pincode_range_cache:
            return set(self._used_pincode_range_cache[cache_key])

        rows = self.fetch_rows(
            f"""
            SELECT DISTINCT p.zipcode AS used_pincode
            FROM {self.locations_table} l
            JOIN {self.pincodes_table} p
              ON CAST(l.pincode_id AS VARCHAR) = CAST(p.id AS VARCHAR)
            WHERE {active_sql("l")}
              AND LOWER(COALESCE(CAST(l.client_location_name AS VARCHAR), ''))
                NOT LIKE '%\\_old%' ESCAPE '\\'
              AND LOWER(COALESCE(CAST(l.alias AS VARCHAR), ''))
                NOT LIKE '%\\_old%' ESCAPE '\\'
              AND try_cast(regexp_replace(CAST(p.zipcode AS VARCHAR), '[^0-9]', '') AS INTEGER)
                BETWEEN {int(min_pin)} AND {int(max_pin)}
            """
        )
        used = {
            int(zipcode)
            for row in rows
            if (zipcode := clean_digits(row.get("used_pincode")))
        }
        self._used_pincode_range_cache[cache_key] = used
        return set(used)



# ==============================================================================
# row_validator.py
# ==============================================================================

import re
from typing import Any, Dict, List, Optional


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
EXISTING_PARTNER_ERROR_PREFIX = "Partner_name already exists"
EXISTING_LOCATION_ERROR_PREFIX = "Location already exists"
REQUIRED_COLUMNS = [
    "fmcode",
    "fmsc",
    "Partner_name",
    "contactNumber",
    "branch_admin_name",
    "email",
    "fmcodeaddress",
    "pickupPincodes",
]


def validate_row(
    row: Dict[str, str],
    row_number: int,
    validator: ValmoValidator,
    pincode_resolver: Optional[PincodeResolver] = None,
    batch_used_pincodes: Optional[set[int]] = None,
) -> Dict[str, Any]:
    errors: List[str] = []
    warnings: List[str] = []

    for column in REQUIRED_COLUMNS:
        if not clean_text(row.get(column)):
            errors.append(f"Missing required field: {column}")

    fmcode = clean_text(row.get("fmcode"))
    fmsc = first_comma_value(row.get("fmsc"))
    partner_name = clean_text(row.get("Partner_name"))
    original_branch_admin_name = clean_text(row.get("branch_admin_name"))
    branch_admin_name = prod_branch_admin_name(original_branch_admin_name)
    contact_number = clean_digits(row.get("contactNumber"))
    email = clean_text(row.get("email")).lower()
    loczipcode = clean_digits(row.get("loczipcode"))
    pickup_pincodes, pincode_errors = split_pickup_pincodes(row.get("pickupPincodes", ""))
    errors.extend(pincode_errors)
    existing_partner: Optional[Dict[str, str]] = None
    existing_location: Optional[Dict[str, str]] = None

    if email and not EMAIL_RE.match(email):
        errors.append(f"Invalid email: {email}")
    if contact_number and not re.fullmatch(r"\d{10,15}", contact_number):
        errors.append("contactNumber must contain 10 to 15 digits")
    existing_partner = validator.partner_exists(partner_name) if partner_name else None
    if existing_partner:
        match_type = clean_text(existing_partner.get("match_type"))
        if existing_partner.get("ambiguous_match"):
            errors.append(
                "Partner_name matches multiple partners in VALMO.partners: "
                f"{clean_text(existing_partner.get('match_summary'))}. "
                "Confirm the correct partner_id before onboarding."
            )
        else:
            existing_name = clean_text(existing_partner.get("name"))
            existing_id = clean_text(existing_partner.get("id"))
            existing_status = clean_text(existing_partner.get("status"))
            errors.append(
                f"Partner_name already exists as '{existing_name}' "
                f"in VALMO.partners (id={existing_id}, status={existing_status})"
            )
        if match_type in {"compact_subset", "one_character_spelling", "token_spelling"}:
            existing_name = clean_text(existing_partner.get("name"))
            warnings.append(
                f"Partner_name matched existing partner '{existing_name}' by spacing/spelling check"
            )
        if match_type.endswith("_resolved_by_recent_location"):
            warnings.append(
                "Partner_name matched multiple partners; selected "
                f"id={clean_text(existing_partner.get('id'))} using latest VALMO.locations update "
                f"(location_id={clean_text(existing_partner.get('recent_location_id'))}, "
                f"updated_at={clean_text(existing_partner.get('recent_location_updated_at'))})"
            )
        if match_type.endswith("_resolved_by_partner_override"):
            warnings.append(
                "Partner_name matched multiple partners; selected "
                f"id={clean_text(existing_partner.get('id'))} using approved partner override"
            )

    existing_location = validator.location_exists(fmcode, clean_text(row.get("clientLocationName")))
    if existing_location:
        location_id = clean_text(existing_location.get("id"))
        alias = clean_text(existing_location.get("alias"))
        existing_client_location = clean_text(existing_location.get("client_location_name"))
        entity_id = clean_text(existing_location.get("entity_id"))
        status_value = clean_text(existing_location.get("status"))
        errors.append(
            f"{EXISTING_LOCATION_ERROR_PREFIX} in VALMO.locations "
            f"(id={location_id}, alias={alias}, client_location_name={existing_client_location}, "
            f"partner_id={entity_id}, status={status_value})"
        )

    state_id = ""
    state_name = ""
    city_id = ""
    city_name = ""
    pincode_resolution: Optional[Dict[str, Any]] = None

    jenkins_pincode_needed = False
    should_resolve_pincode = False
    pincode_seed = loczipcode
    if not loczipcode:
        jenkins_pincode_needed = True
        should_resolve_pincode = True
        pincode_seed = pickup_pincodes[0] if pickup_pincodes else ""
        warnings.append("loczipcode is blank; pincode Jenkins must generate unique loczipcode")
    elif not PINCODE_RE.match(loczipcode):
        errors.append("loczipcode must be a 6 digit pincode")
    else:
        existing_zip = validator.loczipcode_in_pincodes(loczipcode)
        if existing_zip:
            jenkins_pincode_needed = True
            should_resolve_pincode = True
            warnings.append(
                f"loczipcode {loczipcode} exists in VALMO.pincodes; "
                "pincode Jenkins should verify/generate unique value"
            )
        else:
            jenkins_pincode_needed = True
            should_resolve_pincode = True
            warnings.append(
                f"loczipcode {loczipcode} is not present in VALMO.pincodes; "
                "pincode Jenkins must create/insert it"
            )

    if should_resolve_pincode and pincode_resolver is not None:
        if not pincode_seed:
            errors.append("Unable to call pincode Jenkins because no seed pincode is available")
        else:
            try:
                pincode_resolution = pincode_resolver.resolve(
                    pincode_seed,
                    batch_used_pincodes if batch_used_pincodes is not None else set(),
                    context_key=fmcode,
                )
                loczipcode = pincode_resolution["output_pincode"]
                warnings.append(
                    "loczipcode resolved by pincode Jenkins logic: "
                    f"{pincode_resolution['input_pincode']} -> {loczipcode}"
                )
                if pincode_resolution.get("reused_recent_cache"):
                    warnings.append(
                        "loczipcode reused from recent pincode API cache while Trino catches up"
                    )
                if pincode_resolution.get("insert_api_called"):
                    api_response = pincode_resolution.get("insert_api_response") or {}
                    warnings.append(
                        "pincode insert API called: "
                        + (
                            clean_text(api_response.get("message"))
                            or f"rows affected {api_response.get('rows_affected', '')}"
                        )
                    )
            except Exception as exc:
                errors.append(f"Pincode Jenkins resolution failed: {exc}")

    city_lookup_pincode = loczipcode if loczipcode and PINCODE_RE.match(loczipcode) else ""
    if not city_lookup_pincode and pincode_seed and PINCODE_RE.match(pincode_seed):
        city_lookup_pincode = pincode_seed

    address_text = row.get("fmcodeaddress", "")
    address_state = validator.state_from_address(address_text)
    pincode_state = validator.state_from_pincode(city_lookup_pincode) if city_lookup_pincode else None
    address_city = validator.city_from_address(address_text)

    if city_lookup_pincode:
        loc_city, loc_city_warnings = validator.derive_city_from_loczipcode(city_lookup_pincode)
        warnings.extend(loc_city_warnings)
    else:
        loc_city = None

    if loc_city and address_state:
        address_state_id = clean_text(address_state.get("id"))
        loc_city_state_id = clean_text(loc_city.get("state_id"))
        if address_state_id and loc_city_state_id and address_state_id != loc_city_state_id:
            original_city = clean_text(loc_city.get("name")) or clean_text(loc_city.get("id"))
            original_state = clean_text(loc_city.get("state_name")) or loc_city_state_id
            preferred_city: Optional[Dict[str, str]] = None
            if address_city and clean_text(address_city.get("state_id")) == address_state_id:
                preferred_city = address_city
            else:
                preferred_city = validator.pick_active_city(address_state_id)
                if preferred_city:
                    preferred_city["source"] = (
                        f"address state fallback: first active city for {address_state.get('name')}"
                    )
            if preferred_city:
                warnings.append(
                    f"pincode {city_lookup_pincode} mapped to {original_city}, {original_state}; "
                    f"address has {address_state.get('name')}, so city_id derived from "
                    f"{preferred_city.get('source') or 'address state'}"
                )
                loc_city = preferred_city

    if loc_city and not address_state and pincode_state:
        pincode_state_id = clean_text(pincode_state.get("id"))
        loc_city_state_id = clean_text(loc_city.get("state_id"))
        if pincode_state_id and loc_city_state_id and pincode_state_id != loc_city_state_id:
            original_city = clean_text(loc_city.get("name")) or clean_text(loc_city.get("id"))
            original_state = clean_text(loc_city.get("state_name")) or loc_city_state_id
            preferred_city = validator.pick_active_city(pincode_state_id)
            if preferred_city:
                preferred_city["source"] = (
                    f"pincode prefix state fallback: first active city for {pincode_state.get('name')}"
                )
                warnings.append(
                    f"pincode {city_lookup_pincode} mapped to {original_city}, {original_state}; "
                    f"pincode prefix has {pincode_state.get('name')}, so city_id derived from "
                    f"{preferred_city.get('source')}"
                )
                loc_city = preferred_city

    if not loc_city:
        if address_city:
            source = clean_text(address_city.get("source")) or "address city match"
            warnings.append(
                f"city_id derived from {source}"
                + (f" after pincode {city_lookup_pincode} did not map to active TITAN city" if city_lookup_pincode else "")
            )
            loc_city = address_city

    if not loc_city:
        state_city = validator.first_city_from_state_fallback(
            address=address_text,
            pincode=city_lookup_pincode,
        )
        if state_city:
            source = clean_text(state_city.get("source")) or "state fallback"
            warnings.append(
                f"city_id derived from {source}"
                + (f" after pincode {city_lookup_pincode} and address city did not map to active TITAN city" if city_lookup_pincode else "")
            )
            loc_city = state_city

    if loc_city:
        city_id = clean_text(loc_city.get("id"))
        city_name = clean_text(loc_city.get("name"))
        state_id = clean_text(loc_city.get("state_id"))
        state_name = clean_text(loc_city.get("state_name"))
    elif city_lookup_pincode:
        errors.append(f"Unable to derive city_id from pincode {city_lookup_pincode}")

    status = "valid" if not errors else "failed"
    return {
        "row_number": row_number,
        "status": status,
        "errors": errors,
        "warnings": warnings,
        "fmcode": fmcode,
        "fmsc": fmsc,
        "partner_name": partner_name,
        "existing_partner_id": clean_text(existing_partner.get("id")) if existing_partner else "",
        "existing_partner_name": clean_text(existing_partner.get("name")) if existing_partner else "",
        "existing_partner_status": clean_text(existing_partner.get("status")) if existing_partner else "",
        "partner_match_blocked": bool(existing_partner and existing_partner.get("ambiguous_match")),
        "partner_match_summary": clean_text(existing_partner.get("match_summary")) if existing_partner else "",
        "existing_location_id": clean_text(existing_location.get("id")) if existing_location else "",
        "existing_location_alias": clean_text(existing_location.get("alias")) if existing_location else "",
        "existing_location_clientLocationName": clean_text(
            existing_location.get("client_location_name")
        )
        if existing_location
        else "",
        "existing_location_status": clean_text(existing_location.get("status")) if existing_location else "",
        "existing_location_hash": clean_text(existing_location.get("location_hash")) if existing_location else "",
        "contact_number": contact_number,
        "branch_admin_name": branch_admin_name,
        "email": email,
        "fmcodeaddress": clean_text(row.get("fmcodeaddress")),
        "original_clientLocationName": clean_text(row.get("clientLocationName")),
        "original_city_id": clean_text(row.get("city_id")),
        "original_isLoadsharePartner": clean_text(row.get("isLoadsharePartner")),
        "original_isMigratedLocation": clean_text(row.get("isMigratedLocation")),
        "clientLocationName": fmcode,
        "loczipcode": loczipcode,
        "jenkins_pincode_needed": jenkins_pincode_needed,
        "pincode_resolution": pincode_resolution,
        "pincode_insert_sql": pincode_resolution["insert_sql"] if pincode_resolution else "",
        "pincode_inserted": pincode_resolution["inserted"] if pincode_resolution else False,
        "pincode_insert_api_called": pincode_resolution["insert_api_called"] if pincode_resolution else False,
        "pincode_insert_api_message": clean_text(
            (pincode_resolution.get("insert_api_response") or {}).get("message")
            if pincode_resolution
            else ""
        ),
        "pickup_pincodes": pickup_pincodes,
        "state_id": state_id,
        "state_name": state_name,
        "city_id": city_id,
        "city_name": city_name,
        "isLoadsharePartner": 0,
        "isMigratedLocation": 0,
        "original_row": row,
    }



# ==============================================================================
# report_writer.py
# ==============================================================================

import csv
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List


EXISTING_PARTNER_UPLOAD_HEADERS = [
    "fmcode",
    "fmsc",
    "partner_id",
    "contactNumber",
    "branch_admin_name",
    "email",
    "clientLocationName",
    "fmcodeaddress",
    "loczipcode",
    "pickupPincodes",
    "city_id",
    "isMigratedLocation",
]

VALIDATION_REPORT_FIELDS = [
    "row_number",
    "status",
    "errors",
    "warnings",
    "fmcode",
    "fmsc",
    "partner_name",
    "existing_partner_id",
    "existing_partner_name",
    "existing_partner_status",
    "partner_match_blocked",
    "partner_match_summary",
    "existing_location_id",
    "existing_location_alias",
    "existing_location_clientLocationName",
    "existing_location_status",
    "existing_location_hash",
    "contact_number",
    "branch_admin_name",
    "email",
    "fmcodeaddress",
    "original_clientLocationName",
    "clientLocationName",
    "loczipcode",
    "jenkins_pincode_needed",
    "pincode_insert_sql",
    "pincode_inserted",
    "pincode_insert_api_called",
    "pincode_insert_api_message",
    "state_id",
    "state_name",
    "city_id",
    "city_name",
    "original_city_id",
    "original_isLoadsharePartner",
    "isLoadsharePartner",
    "original_isMigratedLocation",
    "isMigratedLocation",
    "pickup_pincodes",
]


def now_stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def short_reason(messages: List[str], fallback: str) -> str:
    if not messages:
        return fallback
    return "NOT UPDATED: " + " | ".join(messages)


def field_reason(result: Dict[str, Any], field_name: str) -> str:
    errors = result.get("errors", [])
    warnings = result.get("warnings", [])
    if field_name == "city_id":
        relevant = [
            message
            for message in errors
            if "Pickup pincode" in message
            or "state_id" in message
            or "TITAN.cities" in message
            or "derive" in message
        ]
        return short_reason(relevant or errors, "city_id not derived")
    if field_name == "loczipcode":
        relevant = [message for message in warnings if "loczipcode" in message]
        return short_reason(relevant, "pincode Jenkins must generate unique loczipcode")
    if field_name == "Partner_name":
        relevant = [message for message in errors if "Partner_name" in message]
        return short_reason(relevant, "partner validation failed")
    return short_reason(errors, f"{field_name} not updated")


def build_validated_upload_rows(headers: List[str], results: List[Dict[str, Any]]) -> List[List[Any]]:
    output_rows: List[List[Any]] = []
    for result in results:
        if (
            result.get("existing_partner_id")
            or result.get("existing_location_id")
            or result.get("partner_match_blocked")
        ):
            continue

        original = result.get("original_row", {})
        row = {header: clean_text(original.get(header)) for header in headers}

        for field in ("fmsc", "clientLocationName", "loczipcode", "city_id", "branch_admin_name"):
            if field in row:
                row[field] = result.get(field) or field_reason(result, field)
        for field in ("isLoadsharePartner", "isMigratedLocation"):
            if field in row:
                row[field] = "0"
        if "Partner_name" in row and any("Partner_name" in msg for msg in result.get("errors", [])):
            row["Partner_name"] = field_reason(result, "Partner_name")

        output_rows.append([row.get(header, "") for header in headers])
    return output_rows


def existing_partner_ready(result: Dict[str, Any]) -> bool:
    errors = result.get("errors") or []
    return (
        bool(result.get("existing_partner_id"))
        and not result.get("existing_location_id")
        and all(
            clean_text(error).startswith(EXISTING_PARTNER_ERROR_PREFIX) for error in errors
        )
    )


def build_existing_partner_upload_rows(results: List[Dict[str, Any]]) -> List[List[Any]]:
    output_rows: List[List[Any]] = []
    for result in results:
        if not existing_partner_ready(result):
            continue

        row = {
            "fmcode": result.get("fmcode"),
            "fmsc": result.get("fmsc"),
            "partner_id": result.get("existing_partner_id"),
            "contactNumber": "",
            "branch_admin_name": result.get("branch_admin_name"),
            "email": result.get("email"),
            "clientLocationName": result.get("clientLocationName"),
            "fmcodeaddress": result.get("fmcodeaddress"),
            "loczipcode": result.get("loczipcode"),
            "pickupPincodes": ", ".join(result.get("pickup_pincodes") or []),
            "city_id": result.get("city_id"),
            "isMigratedLocation": "0",
        }
        output_rows.append([clean_text(row.get(header)) for header in EXISTING_PARTNER_UPLOAD_HEADERS])
    return output_rows


def write_reports(
    results: List[Dict[str, Any]],
    output_dir: Path,
    source_file: Path,
    source_headers: List[str],
) -> Dict[str, str]:
    output_dir.mkdir(parents=True, exist_ok=True)
    base_name = source_file.stem.replace(" ", "_")
    stamp = now_stamp()
    json_path = output_dir / f"{base_name}_validation_{stamp}.json"
    csv_path = output_dir / f"{base_name}_validation_{stamp}.csv"
    upload_csv_path = output_dir / f"{base_name}_new_partner_fm_{stamp}.csv"
    existing_partner_upload_path = output_dir / f"{base_name}_existing_partner_upload_{stamp}.csv"
    existing_partner_rows = build_existing_partner_upload_rows(results)

    summary = {
        "source_file": str(source_file),
        "total_rows": len(results),
        "valid_rows": sum(1 for result in results if result["status"] == "valid"),
        "failed_rows": sum(1 for result in results if result["status"] == "failed"),
        "existing_partner_upload_rows": len(existing_partner_rows),
        "existing_location_rows": sum(1 for result in results if result.get("existing_location_id")),
        "partner_match_blocked_rows": sum(
            1 for result in results if result.get("partner_match_blocked")
        ),
        "results": results,
    }
    json_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")

    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=VALIDATION_REPORT_FIELDS)
        writer.writeheader()
        for result in results:
            row = dict(result)
            row["errors"] = " | ".join(result["errors"])
            row["warnings"] = " | ".join(result["warnings"])
            row["pickup_pincodes"] = ", ".join(result["pickup_pincodes"])
            writer.writerow({key: row.get(key, "") for key in VALIDATION_REPORT_FIELDS})

    new_partner_rows = build_validated_upload_rows(source_headers, results)
    write_simple_csv(upload_csv_path, source_headers, new_partner_rows)
    write_simple_csv(existing_partner_upload_path, EXISTING_PARTNER_UPLOAD_HEADERS, existing_partner_rows)

    return {
        "json_report": str(json_path),
        "csv_report": str(csv_path),
        "new_partner_fm_csv_report": str(upload_csv_path),
        "validated_csv_report": str(upload_csv_path),
        "existing_partner_csv_report": str(existing_partner_upload_path),
    }



# ==============================================================================
# scripts/trigger_jenkins_file_job.py
# ==============================================================================

#!/usr/bin/env python3
"""Trigger a Jenkins job with one uploaded file parameter using stdlib only."""


import argparse
import base64
import http.cookiejar
import json
import mimetypes
import os
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional


DEFAULT_TIMEOUT_SECONDS = 20.0
DEFAULT_QUEUE_TIMEOUT_SECONDS = 90.0
DEFAULT_BUILD_TIMEOUT_SECONDS = 300.0
DEFAULT_POLL_SECONDS = 1.0
USER_AGENT = "fm-onboard-automation/1.0"


def jenkins_job_url(base_url: str, job_name: str) -> str:
    base = base_url.rstrip("/")
    parts = [part for part in job_name.strip("/").split("/") if part]
    encoded = "/job/".join(urllib.parse.quote(part) for part in parts)
    return f"{base}/job/{encoded}"


def jenkins_api_url(resource_url: str) -> str:
    return f"{resource_url.rstrip('/')}/api/json"


def default_headers() -> Dict[str, str]:
    return {
        "Accept": "application/json",
        "User-Agent": USER_AGENT,
    }


def basic_auth_headers(user: str, token: str) -> Dict[str, str]:
    if not user and not token:
        return {}
    if not user or not token:
        raise ValueError("Both Jenkins user and token are required for basic auth")

    credentials = f"{user}:{token}".encode("utf-8")
    encoded = base64.b64encode(credentials).decode("ascii")
    return {"Authorization": f"Basic {encoded}"}


def format_http_error(exc: urllib.error.HTTPError, url: str) -> str:
    body = ""
    try:
        body = exc.read().decode("utf-8", errors="replace").strip()
    except Exception:
        body = ""
    if len(body) > 500:
        body = f"{body[:500]}..."

    message = f"HTTP {exc.code} while calling {url}"
    if body:
        message = f"{message}: {body}"
    return message


def read_json(
    url: str,
    *,
    headers: Mapping[str, str],
    opener: urllib.request.OpenerDirector,
    timeout: float,
) -> Dict[str, Any]:
    request = urllib.request.Request(url, headers=dict(headers), method="GET")
    try:
        with opener.open(request, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise RuntimeError(format_http_error(exc, url)) from exc


def crumb_headers(
    base_url: str,
    *,
    headers: Mapping[str, str],
    opener: urllib.request.OpenerDirector,
    timeout: float,
) -> Dict[str, str]:
    url = f"{base_url.rstrip('/')}/crumbIssuer/api/json"
    request = urllib.request.Request(url, headers=dict(headers), method="GET")
    try:
        with opener.open(request, timeout=timeout) as response:
            data = json.loads(response.read().decode())
        return {data["crumbRequestField"]: data["crumb"]}
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return {}
        raise RuntimeError(format_http_error(exc, url)) from exc
    except Exception as exc:
        raise RuntimeError(f"Unable to fetch Jenkins crumb from {url}: {exc}") from exc


def quote_header_value(value: str) -> str:
    return (
        str(value)
        .replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "")
        .replace("\n", "")
    )


def build_multipart(
    *,
    file_param: str,
    file_path: Path,
    fields: Dict[str, str],
) -> tuple[bytes, str]:
    boundary = f"----fmOnboardBoundary{uuid.uuid4().hex}"
    parts: List[bytes] = []

    def add_field(name: str, value: str) -> None:
        parts.append(f"--{boundary}\r\n".encode())
        safe_name = quote_header_value(name)
        parts.append(f'Content-Disposition: form-data; name="{safe_name}"\r\n\r\n'.encode())
        parts.append(str(value).encode("utf-8"))
        parts.append(b"\r\n")

    def add_file(name: str, path: Path) -> None:
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        safe_name = quote_header_value(name)
        safe_filename = quote_header_value(path.name)
        parts.append(f"--{boundary}\r\n".encode())
        parts.append(
            f'Content-Disposition: form-data; name="{safe_name}"; '
            f'filename="{safe_filename}"\r\n'.encode()
        )
        parts.append(f"Content-Type: {mime}\r\n\r\n".encode())
        parts.append(path.read_bytes())
        parts.append(b"\r\n")

    for key, value in fields.items():
        add_field(key, value)
    add_file(file_param, file_path)
    parts.append(f"--{boundary}--\r\n".encode())
    return b"".join(parts), boundary


def trigger_file_job(
    *,
    base_url: str,
    job_name: str,
    file_param: str,
    file_path: Path,
    fields: Dict[str, str],
    wait: bool,
    auth_headers: Optional[Mapping[str, str]] = None,
    request_timeout: float = DEFAULT_TIMEOUT_SECONDS,
    queue_timeout: float = DEFAULT_QUEUE_TIMEOUT_SECONDS,
    build_timeout: float = DEFAULT_BUILD_TIMEOUT_SECONDS,
    poll_seconds: float = DEFAULT_POLL_SECONDS,
) -> int:
    if not file_path.is_file():
        raise FileNotFoundError(file_path)
    if queue_timeout <= 0 or build_timeout <= 0 or request_timeout <= 0:
        raise ValueError("Timeout values must be greater than zero")
    if poll_seconds <= 0:
        raise ValueError("--poll-seconds must be greater than zero")

    job_url = jenkins_job_url(base_url, job_name)
    body, boundary = build_multipart(file_param=file_param, file_path=file_path, fields=fields)

    opener = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar())
    )
    shared_headers = default_headers()
    shared_headers.update(auth_headers or {})
    headers = dict(shared_headers)
    headers.update(
        crumb_headers(
            base_url,
            headers=shared_headers,
            opener=opener,
            timeout=request_timeout,
        )
    )
    headers["Content-Type"] = f"multipart/form-data; boundary={boundary}"
    request = urllib.request.Request(
        f"{job_url}/buildWithParameters",
        data=body,
        headers=headers,
        method="POST",
    )
    try:
        with opener.open(request, timeout=request_timeout) as response:
            queue_url = response.headers.get("Location")
    except urllib.error.HTTPError as exc:
        raise RuntimeError(format_http_error(exc, request.full_url)) from exc

    print(f"Queued Jenkins job: {queue_url}")

    if not wait:
        return 0
    if not queue_url:
        raise RuntimeError("Jenkins did not return a queue URL")

    build_url: Optional[str] = None
    queue_deadline = time.monotonic() + queue_timeout
    while time.monotonic() < queue_deadline:
        item = read_json(
            jenkins_api_url(queue_url),
            headers=shared_headers,
            opener=opener,
            timeout=request_timeout,
        )
        if item.get("cancelled"):
            reason = item.get("why") or "Queue item was cancelled"
            raise RuntimeError(str(reason))

        executable = item.get("executable") or {}
        if isinstance(executable, dict) and executable.get("url"):
            build_url = str(executable["url"])
            break
        time.sleep(poll_seconds)

    if not build_url:
        raise RuntimeError(f"Timed out waiting for Jenkins build to start: {queue_url}")
    print(f"Started Jenkins build: {build_url}")

    build_deadline = time.monotonic() + build_timeout
    while time.monotonic() < build_deadline:
        build = read_json(
            jenkins_api_url(build_url),
            headers=shared_headers,
            opener=opener,
            timeout=request_timeout,
        )
        if not build.get("building"):
            result = build.get("result")
            print(f"Jenkins build result: {result}")
            print(f"Jenkins build URL: {build_url}")
            return 0 if result == "SUCCESS" else 1
        time.sleep(poll_seconds)

    raise RuntimeError(f"Timed out waiting for Jenkins build to finish: {build_url}")


def parse_fields(raw_fields: List[str]) -> Dict[str, str]:
    fields: Dict[str, str] = {}
    for raw in raw_fields:
        if not raw:
            continue
        if "=" not in raw:
            raise ValueError(f"Field must be KEY=VALUE: {raw}")
        key, value = raw.split("=", 1)
        if not key:
            raise ValueError(f"Field key cannot be blank: {raw}")
        fields[key] = value
    return fields


def main() -> int:
    parser = argparse.ArgumentParser(description="Trigger Jenkins file upload job")
    parser.add_argument("--base-url", required=True)
    parser.add_argument("--job", required=True)
    parser.add_argument("--file-param", required=True)
    parser.add_argument("--file", required=True)
    parser.add_argument("--field", action="append", default=[])
    parser.add_argument("--wait", action="store_true")
    parser.add_argument("--user", default=os.getenv("JENKINS_USER", ""))
    parser.add_argument("--token", default=os.getenv("JENKINS_TOKEN", ""))
    parser.add_argument("--request-timeout", type=float, default=DEFAULT_TIMEOUT_SECONDS)
    parser.add_argument("--queue-timeout", type=float, default=DEFAULT_QUEUE_TIMEOUT_SECONDS)
    parser.add_argument("--build-timeout", type=float, default=DEFAULT_BUILD_TIMEOUT_SECONDS)
    parser.add_argument("--poll-seconds", type=float, default=DEFAULT_POLL_SECONDS)
    args = parser.parse_args()

    return trigger_file_job(
        base_url=args.base_url,
        job_name=args.job,
        file_param=args.file_param,
        file_path=Path(args.file),
        fields=parse_fields(args.field),
        wait=args.wait,
        auth_headers=basic_auth_headers(args.user, args.token),
        request_timeout=args.request_timeout,
        queue_timeout=args.queue_timeout,
        build_timeout=args.build_timeout,
        poll_seconds=args.poll_seconds,
    )



# ==============================================================================
# AI FM onboarding entrypoint
# ==============================================================================

import contextlib
import io

DEFAULT_OUTPUT_DIR = Path("outputs/ai-fm-onboarding")
DEFAULT_DOWNSTREAM_JENKINS_BASE_URL = "https://jenkins.prd.valmo.in"
DEFAULT_PINCODE_LOGIN_URL = "http://hydra.lgtc-prd.valmo.in/v1/login"
DEFAULT_PINCODE_INSERT_API_URL = "http://hub-log10.lgtc-prd.valmo.in/log10/update/mutation"
DEFAULT_NEW_PARTNER_JOB_NAME = "support/log10/Regular_tasks/fm_new_partner_location_onboarding"
DEFAULT_NEW_PARTNER_FILE_PARAM = "New_partner_location_onboarding_input"
DEFAULT_EXISTING_PARTNER_JOB_NAME = "support/log10/Regular_tasks/fm_existing_partner_location_onboarding"
DEFAULT_EXISTING_PARTNER_FILE_PARAM = "old_partner_location_onboarding_input.csv"


def bool_env(name: str, default: str = "true") -> bool:
    return clean_text(os.getenv(name, default)).lower() in {"1", "true", "yes", "y"}


def apply_runtime_defaults() -> None:
    defaults = {
        "PINCODE_LOGIN_URL": DEFAULT_PINCODE_LOGIN_URL,
        "PINCODE_INSERT_API_URL": DEFAULT_PINCODE_INSERT_API_URL,
        "DOWNSTREAM_JENKINS_BASE_URL": DEFAULT_DOWNSTREAM_JENKINS_BASE_URL,
        "NEW_PARTNER_JOB_NAME": DEFAULT_NEW_PARTNER_JOB_NAME,
        "NEW_PARTNER_FILE_PARAM": DEFAULT_NEW_PARTNER_FILE_PARAM,
        "EXISTING_PARTNER_JOB_NAME": DEFAULT_EXISTING_PARTNER_JOB_NAME,
        "EXISTING_PARTNER_FILE_PARAM": DEFAULT_EXISTING_PARTNER_FILE_PARAM,
        "PINCODE_INSERT_API_ENABLED": "true",
        "AUTO_RESOLVE_PINCODES": "1",
        "RESOLVE_PINCODES": "true",
        "TRIGGER_DOWNSTREAM_ONBOARDING": "true",
    }
    for key, value in defaults.items():
        os.environ.setdefault(key, value)


def prepare_ai_input(files: list[Path], input_dir: Path | None) -> Path:
    if input_dir:
        if not input_dir.is_dir():
            raise FileNotFoundError(f"Input folder not found: {input_dir}")
        return input_dir
    if not files:
        raise ValueError("Provide at least one Ops input file, a zip file, or --input-dir")
    for file in files:
        if not file.is_file():
            raise FileNotFoundError(f"Input file not found: {file}")
    if len(files) == 1:
        return files[0]

    merged_dir = Path(tempfile.mkdtemp(prefix="fm_ai_uploads_"))
    for index, file in enumerate(files, start=1):
        shutil.copy2(file, merged_dir / f"ops-upload-{index:02d}{file.suffix}")
    return merged_dir


def validate_ops_input(input_path: Path, output_dir: Path, resolve_pincodes: bool) -> tuple[int, dict[str, Any], list[dict[str, Any]]]:
    try:
        sheet_name, headers, rows = read_spreadsheet_rows(input_path)
    except ValueError as exc:
        summary = {
            "success": False,
            "source_file": str(input_path),
            "total_rows": 0,
            "valid_rows": 0,
            "failed_rows": 0,
            "error": str(exc),
        }
        return 2, summary, []

    try:
        titan_lookup = TrinoTitanLookup()
        validator = ValmoValidator(titan_lookup)
        pincode_insert_api_client = PincodeInsertApiClient.from_env() if resolve_pincodes else None
        pincode_resolver = (
            PincodeResolver(
                validator,
                insert_api_client=pincode_insert_api_client,
                recent_cache=RecentPincodeCache.from_env(),
            )
            if resolve_pincodes
            else None
        )
        batch_used_pincodes: set[int] = set()
        try:
            results = [
                validate_row(
                    row,
                    row_number=index + 2,
                    validator=validator,
                    pincode_resolver=pincode_resolver,
                    batch_used_pincodes=batch_used_pincodes,
                )
                for index, row in enumerate(rows)
            ]
        finally:
            validator.close()
            if pincode_resolver is not None:
                pincode_resolver.close()
    except Exception as exc:
        summary = {
            "success": False,
            "source_file": str(input_path),
            "sheet": sheet_name,
            "total_rows": len(rows),
            "valid_rows": 0,
            "failed_rows": len(rows),
            "error": str(exc),
        }
        return 2, summary, []

    reports = write_reports(results, output_dir, input_path, headers)
    summary = {
        "success": all(result["status"] == "valid" for result in results),
        "sheet": sheet_name,
        "total_rows": len(results),
        "valid_rows": sum(1 for result in results if result["status"] == "valid"),
        "failed_rows": sum(1 for result in results if result["status"] == "failed"),
        **reports,
    }
    return (0 if summary["success"] else 1), summary, results


def csv_dict_rows(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        return []
    with path.open(encoding="utf-8-sig", newline="") as file_obj:
        return [
            {clean_text(key): clean_text(value) for key, value in row.items() if key}
            for row in csv.DictReader(file_obj)
            if any(clean_text(value) for value in row.values())
        ]


def csv_location_names(path: Path) -> list[str]:
    names: list[str] = []
    for row in csv_dict_rows(path):
        name = clean_text(row.get("clientLocationName")) or clean_text(row.get("fmcode"))
        if name and name not in names:
            names.append(name)
    return names


def row_to_fix(row: dict[str, Any]) -> dict[str, Any]:
    return {
        "row_number": row.get("row_number"),
        "fmcode": row.get("fmcode"),
        "partner_name": row.get("partner_name"),
        "clientLocationName": row.get("clientLocationName"),
        "errors": row.get("errors") or [],
        "warnings": row.get("warnings") or [],
    }


def build_ai_decision(summary: dict[str, Any], results: list[dict[str, Any]]) -> dict[str, Any]:
    failed_rows = [row for row in results if row.get("status") == "failed"]
    existing_rows = [row for row in failed_rows if existing_partner_ready(row)]
    correction_rows = [row for row in failed_rows if not existing_partner_ready(row)]

    new_csv = Path(clean_text(summary.get("new_partner_fm_csv_report")))
    existing_csv = Path(clean_text(summary.get("existing_partner_csv_report")))
    new_locations = csv_location_names(new_csv)
    existing_locations = csv_location_names(existing_csv)

    return {
        "ready": not correction_rows and bool(new_locations or existing_locations),
        "new_partner_csv": str(new_csv) if new_locations else "",
        "existing_partner_csv": str(existing_csv) if existing_locations else "",
        "new_partner_rows": len(csv_dict_rows(new_csv)),
        "existing_partner_rows": len(csv_dict_rows(existing_csv)),
        "new_partner_locations": new_locations,
        "existing_partner_locations": existing_locations,
        "rows_to_fix": [row_to_fix(row) for row in correction_rows],
        "existing_partner_ready_rows": [row_to_fix(row) for row in existing_rows],
    }


def failure_ticket_reply(rows_to_fix: list[dict[str, Any]]) -> str:
    if not rows_to_fix:
        return "FM onboarding validation failed. Please check the validation output."

    lines = ["FM onboarding validation failed. Please correct the below issue(s):"]
    for row in rows_to_fix[:10]:
        fmcode = clean_text(row.get("fmcode")) or "-"
        errors = [clean_text(error) for error in row.get("errors", []) if clean_text(error)]
        lines.append(
            f"- Row {row.get('row_number')}: fmcode={fmcode}, "
            f"reason={' | '.join(errors[:3]) or 'Validation error'}"
        )
    if len(rows_to_fix) > 10:
        lines.append(f"- {len(rows_to_fix) - 10} more row(s) need correction.")
    return "\n".join(lines)


def success_ticket_reply(decision: dict[str, Any], triggered_jobs: list[dict[str, Any]]) -> str:
    location_lines: list[str] = []
    for name in decision.get("new_partner_locations") or []:
        location_lines.append(f"- {name}: onboarded as new partner location")
    for name in decision.get("existing_partner_locations") or []:
        location_lines.append(f"- {name}: onboarded under existing partner")

    message = "FM onboarding completed successfully." if triggered_jobs else "FM onboarding validation completed successfully."
    if location_lines:
        message += "\n\nLocation(s):\n" + "\n".join(location_lines[:25])
        if len(location_lines) > 25:
            message += f"\n- {len(location_lines) - 25} more location(s)"
    return message


def trigger_downstream_job(
    *,
    label: str,
    csv_file: Path,
    job_env: str,
    file_param_env: str,
    default_job_name: str,
    default_file_param: str,
) -> dict[str, Any]:
    base_url = clean_text(os.getenv("DOWNSTREAM_JENKINS_BASE_URL") or os.getenv("JENKINS_URL")) or DEFAULT_DOWNSTREAM_JENKINS_BASE_URL
    job_name = clean_text(os.getenv(job_env)) or default_job_name
    file_param = clean_text(os.getenv(file_param_env)) or default_file_param
    auth_headers = basic_auth_headers(
        clean_text(os.getenv("JENKINS_USER")),
        clean_text(os.getenv("JENKINS_TOKEN")),
    )

    output = io.StringIO()
    response = {
        "type": label,
        "success": False,
        "job_name": job_name,
        "file_param": file_param,
        "input_csv": str(csv_file),
    }
    try:
        with contextlib.redirect_stdout(output):
            exit_code = trigger_file_job(
                base_url=base_url,
                job_name=job_name,
                file_param=file_param,
                file_path=csv_file,
                fields={},
                wait=True,
                auth_headers=auth_headers,
                request_timeout=float(os.getenv("JENKINS_REQUEST_TIMEOUT_SECONDS", "20")),
                queue_timeout=float(os.getenv("JENKINS_QUEUE_TIMEOUT_SECONDS", "90")),
                build_timeout=float(os.getenv("JENKINS_BUILD_TIMEOUT_SECONDS", "900")),
                poll_seconds=float(os.getenv("JENKINS_POLL_SECONDS", "2")),
            )
        response["success"] = exit_code == 0
        response["logs"] = output.getvalue().strip()
        if exit_code != 0:
            response["error"] = "Downstream Jenkins job failed"
    except Exception as exc:
        response["error"] = str(exc)
        response["logs"] = output.getvalue().strip()
    return response


def ai_main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Single-file AI runner for FM onboarding")
    parser.add_argument("files", nargs="*", help="Ops spreadsheet file(s) or zip file")
    parser.add_argument("--input-dir", help="Folder containing Ops spreadsheet files")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR))
    parser.add_argument("--no-trigger", action="store_true", help="Validate only; do not trigger onboarding Jenkins")
    parser.add_argument("--no-resolve-pincodes", action="store_true", help="Do not resolve or insert loczipcode")
    args = parser.parse_args(argv)

    apply_runtime_defaults()
    resolve_pincodes = not args.no_resolve_pincodes and bool_env("RESOLVE_PINCODES", "true")

    try:
        input_path = prepare_ai_input(
            [Path(file) for file in args.files],
            Path(args.input_dir) if args.input_dir else None,
        )
        output_dir = Path(args.output_dir)
        shutil.rmtree(output_dir, ignore_errors=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        validation_returncode, summary, results = validate_ops_input(input_path, output_dir, resolve_pincodes)
        decision = build_ai_decision(summary, results)

        if not decision["ready"]:
            response = {
                "success": False,
                "stage": "validation",
                "message": "Validation failed. Fix the listed rows before onboarding.",
                "ticket_reply": failure_ticket_reply(decision["rows_to_fix"]),
                "validation_returncode": validation_returncode,
                "summary": summary,
                **decision,
            }
            print(json.dumps(response, indent=2, ensure_ascii=False))
            return 1

        triggered_jobs: list[dict[str, Any]] = []
        if not args.no_trigger and bool_env("TRIGGER_DOWNSTREAM_ONBOARDING", "true"):
            if decision["new_partner_csv"]:
                triggered_jobs.append(
                    trigger_downstream_job(
                        label="New Partner",
                        csv_file=Path(decision["new_partner_csv"]),
                        job_env="NEW_PARTNER_JOB_NAME",
                        file_param_env="NEW_PARTNER_FILE_PARAM",
                        default_job_name=DEFAULT_NEW_PARTNER_JOB_NAME,
                        default_file_param=DEFAULT_NEW_PARTNER_FILE_PARAM,
                    )
                )
            if decision["existing_partner_csv"]:
                triggered_jobs.append(
                    trigger_downstream_job(
                        label="Existing Partner",
                        csv_file=Path(decision["existing_partner_csv"]),
                        job_env="EXISTING_PARTNER_JOB_NAME",
                        file_param_env="EXISTING_PARTNER_FILE_PARAM",
                        default_job_name=DEFAULT_EXISTING_PARTNER_JOB_NAME,
                        default_file_param=DEFAULT_EXISTING_PARTNER_FILE_PARAM,
                    )
                )

        failed_jobs = [job for job in triggered_jobs if not job.get("success")]
        ticket_reply = (
            failure_ticket_reply(
                [
                    {
                        "row_number": "",
                        "fmcode": job.get("type"),
                        "errors": [job.get("error") or "Onboarding job failed"],
                    }
                    for job in failed_jobs
                ]
            )
            if failed_jobs
            else success_ticket_reply(decision, triggered_jobs)
        )
        response = {
            "success": not failed_jobs,
            "stage": "completed" if not failed_jobs else "trigger",
            "message": (
                "Validation passed and onboarding jobs completed."
                if triggered_jobs and not failed_jobs
                else "Validation passed. No downstream job was triggered."
                if not triggered_jobs
                else "Validation passed, but one or more onboarding jobs failed."
            ),
            "ticket_reply": ticket_reply,
            "summary": summary,
            **decision,
            "triggered_jobs": triggered_jobs,
        }
        print(json.dumps(response, indent=2, ensure_ascii=False))
        return 1 if failed_jobs else 0
    except Exception as exc:
        print(
            json.dumps(
                {
                    "success": False,
                    "stage": "setup",
                    "message": str(exc),
                    "ticket_reply": f"FM onboarding automation could not start: {exc}",
                },
                indent=2,
                ensure_ascii=False,
            )
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(ai_main())

