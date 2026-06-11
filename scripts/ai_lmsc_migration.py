#!/usr/bin/env python3
"""AI entrypoint for LMSC migration validation, formatting, and Jenkins trigger."""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from spreadsheet_io import read_spreadsheet_rows
from validation_utils import clean_text


DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "outputs" / "lmsc-migration"
DEFAULT_JENKINS_BASE_URL = "https://jenkins.prd.valmo.in"
DEFAULT_LMSC_JOB_NAME = "support/log10/Regular_tasks/lmsc_migration_cd"
DEFAULT_LMSC_FILE_PARAM = "lmscmigration_cd_input.csv"

OUTPUT_HEADERS = [
    "LMDC",
    "Current Sort Centre",
    "New Sort Centre",
    "Current Sort Code",
    "New Sort Code",
]

FIELD_ALIASES = {
    "LMDC": ["lmdc", "dc", "lmdc code", "dc code"],
    "Current Sort Centre": [
        "current sort centre",
        "current sort center",
        "current lmsc",
        "old sort centre",
        "old sort center",
        "old lmsc",
        "current sc",
        "old sc",
    ],
    "New Sort Centre": [
        "new sort centre",
        "new sort center",
        "new lmsc",
        "new sc",
    ],
    "Current Sort Code": [
        "current sort code",
        "old sort code",
        "current sortcode",
        "old sortcode",
    ],
    "New Sort Code": ["new sort code", "new sortcode"],
}


def flag(name: str, default: str = "true") -> bool:
    return os.getenv(name, default).strip().lower() in {"1", "true", "yes", "y"}


def canonical_key(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_text(value).lower())


def compact_code(value: object) -> str:
    return clean_text(value).upper().replace(" ", "")


def normalize_sort_code(value: object) -> str:
    text = clean_text(value).upper().replace(" ", "")
    return "/".join(part for part in text.split("/") if part)


def map_headers(headers: list[str]) -> dict[str, str]:
    by_key = {canonical_key(header): header for header in headers if clean_text(header)}
    mapped: dict[str, str] = {}
    for field, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            source = by_key.get(canonical_key(alias))
            if source:
                mapped[field] = source
                break
    return mapped


def looks_like_text_table(headers: list[str]) -> bool:
    if len(headers) != 1:
        return False
    header_key = canonical_key(headers[0])
    return all(token in header_key for token in ["lmdc", "current", "new", "sort", "code"])


def normalize_text_table(headers: list[str], rows: list[dict[str, str]]) -> list[dict[str, Any]]:
    source = headers[0]
    normalized: list[dict[str, Any]] = []
    for row_number, row in enumerate(rows, start=2):
        values = clean_text(row.get(source)).split()
        row_data = dict(zip(OUTPUT_HEADERS, values[:5]))
        errors = [] if len(values) >= 5 else ["Unable to split row into 5 values"]
        normalized.append({"row_number": row_number, "row": row_data, "errors": errors})
    return normalized


def normalize_table(headers: list[str], rows: list[dict[str, str]]) -> list[dict[str, Any]]:
    mapped = map_headers(headers)
    missing = [field for field in OUTPUT_HEADERS if field not in mapped]
    normalized: list[dict[str, Any]] = []
    for row_number, row in enumerate(rows, start=2):
        row_data = {
            field: clean_text(row.get(mapped[field], "")) if field in mapped else ""
            for field in OUTPUT_HEADERS
        }
        errors = [f"Missing required column(s): {', '.join(missing)}"] if missing else []
        normalized.append({"row_number": row_number, "row": row_data, "errors": errors})
    return normalized


def normalize_rows(headers: list[str], rows: list[dict[str, str]]) -> list[dict[str, Any]]:
    if looks_like_text_table(headers):
        return normalize_text_table(headers, rows)
    return normalize_table(headers, rows)


def validate_row(row_number: int, row: dict[str, str], base_errors: list[str]) -> dict[str, Any]:
    result = {
        "row_number": row_number,
        "LMDC": compact_code(row.get("LMDC")),
        "Current Sort Centre": compact_code(row.get("Current Sort Centre")),
        "New Sort Centre": compact_code(row.get("New Sort Centre")),
        "Current Sort Code": normalize_sort_code(row.get("Current Sort Code")),
        "New Sort Code": normalize_sort_code(row.get("New Sort Code")),
        "errors": list(base_errors),
    }
    for field in OUTPUT_HEADERS:
        if not result[field]:
            result["errors"].append(f"{field} is mandatory")
    result["status"] = "failed" if result["errors"] else "valid"
    return result


def write_output_csv(path: Path, rows: list[dict[str, Any]]) -> int:
    valid_rows = [row for row in rows if row["status"] == "valid"]
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_HEADERS)
        writer.writeheader()
        for row in valid_rows:
            writer.writerow({field: row[field] for field in OUTPUT_HEADERS})
    return len(valid_rows)


def trigger_lmsc_job(output_csv: Path) -> dict[str, Any]:
    base_url = clean_text(os.getenv("DOWNSTREAM_JENKINS_BASE_URL")) or DEFAULT_JENKINS_BASE_URL
    job_name = clean_text(os.getenv("LMSC_MIGRATION_JOB_NAME")) or DEFAULT_LMSC_JOB_NAME
    file_param = clean_text(os.getenv("LMSC_MIGRATION_FILE_PARAM")) or DEFAULT_LMSC_FILE_PARAM

    result = subprocess.run(
        [
            sys.executable,
            "scripts/trigger_jenkins_file_job.py",
            "--base-url",
            base_url,
            "--job",
            job_name,
            "--file-param",
            file_param,
            "--file",
            str(output_csv),
            "--wait",
        ],
        cwd=PROJECT_ROOT,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    return {
        "success": result.returncode == 0,
        "job_name": job_name,
        "file_param": file_param,
        "error": "" if result.returncode == 0 else clean_text(result.stderr or result.stdout),
    }


def failure_reply(rows: list[dict[str, Any]]) -> str:
    failed = [row for row in rows if row["status"] == "failed"]
    lines = ["LMSC migration validation failed. Please correct the below issue(s):"]
    for row in failed[:10]:
        lines.append(
            f"- Row {row['row_number']}: LMDC={row.get('LMDC') or '-'}, "
            f"reason={' | '.join(row['errors'][:3])}"
        )
    if len(failed) > 10:
        lines.append(f"- {len(failed) - 10} more row(s) need correction.")
    return "\n".join(lines)


def success_reply(rows: list[dict[str, Any]], triggered: bool) -> str:
    message = (
        "LMSC migration completed successfully."
        if triggered
        else "LMSC migration validation completed successfully."
    )
    lines = [message, "", "Migration row(s):"]
    for row in rows[:20]:
        lines.append(f"- {row['LMDC']}: {row['Current Sort Centre']} -> {row['New Sort Centre']}")
    if len(rows) > 20:
        lines.append(f"- {len(rows) - 20} more row(s)")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate and trigger LMSC migration")
    parser.add_argument("input_file", help="Ops spreadsheet, CSV, folder, or zip")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR))
    parser.add_argument("--no-trigger", action="store_true", help="Validate/format only")
    args = parser.parse_args()

    try:
        sheet, headers, raw_rows = read_spreadsheet_rows(Path(args.input_file))
        normalized = normalize_rows(headers, raw_rows)
        results = [
            validate_row(item["row_number"], item["row"], item["errors"])
            for item in normalized
        ]

        output_csv = Path(args.output_dir) / DEFAULT_LMSC_FILE_PARAM
        write_output_csv(output_csv, results)

        valid_rows = [row for row in results if row["status"] == "valid"]
        failed_rows = [row for row in results if row["status"] == "failed"]
        trigger_result = None
        trigger_enabled = not args.no_trigger and flag("TRIGGER_DOWNSTREAM_LMSC", "true")

        success = bool(results) and not failed_rows
        if success and trigger_enabled:
            trigger_result = trigger_lmsc_job(output_csv)
            success = bool(trigger_result["success"])

        if failed_rows:
            ticket_reply = failure_reply(results)
            stage = "validation"
        elif trigger_result and not trigger_result["success"]:
            ticket_reply = (
                "LMSC migration validation passed, but Jenkins trigger failed: "
                f"{trigger_result['error']}"
            )
            stage = "trigger"
        else:
            ticket_reply = success_reply(valid_rows, bool(trigger_result))
            stage = "completed"

        print(
            json.dumps(
                {
                    "success": success,
                    "stage": stage,
                    "sheet": sheet,
                    "total_rows": len(results),
                    "valid_rows": len(valid_rows),
                    "failed_rows": len(failed_rows),
                    "output_csv": str(output_csv),
                    "rows_to_fix": [
                        {
                            "row_number": row["row_number"],
                            "LMDC": row.get("LMDC"),
                            "errors": row["errors"],
                        }
                        for row in failed_rows
                    ],
                    "trigger_result": trigger_result,
                    "ticket_reply": ticket_reply,
                },
                indent=2,
                ensure_ascii=False,
            )
        )
        return 0 if success else 1
    except Exception as exc:
        print(
            json.dumps(
                {
                    "success": False,
                    "stage": "setup",
                    "message": str(exc),
                    "ticket_reply": f"LMSC migration automation could not start: {exc}",
                },
                indent=2,
                ensure_ascii=False,
            )
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
