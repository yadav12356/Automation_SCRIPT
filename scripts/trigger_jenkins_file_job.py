#!/usr/bin/env python3
"""Trigger a Jenkins job with one uploaded file parameter using stdlib only."""

from __future__ import annotations

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


if __name__ == "__main__":
    raise SystemExit(main())
