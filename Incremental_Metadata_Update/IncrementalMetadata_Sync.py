from __future__ import annotations

import csv
import json
import logging
import os
import re
import time
import zipfile
from datetime import datetime, timedelta
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import httpx

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover
    def load_dotenv(*_args: Any, **_kwargs: Any) -> bool:
        return False

logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)

DEFAULT_TEMPLATE_CSV_REL_PATH = Path("Files") / "Metadata_Template.csv"
DEFAULT_CUBE_NAME = "OEP_FS"

class IncrementalMetadataSyncError(RuntimeError):
    """Raised for functional failures in incremental metadata sync."""

    def __init__(
        self,
        message: str,
        *,
        error_code: str = "INC_METADATA_SYNC_ERROR",
        exit_code: int = 1,
    ) -> None:
        super().__init__(message)
        self.error_code = error_code
        self.exit_code = exit_code

@dataclass(frozen=True)
class MappingEntry:
    column_number: int  # 1-based
    dimension_name: str

@dataclass(frozen=True)
class ExistingMember:
    name: str
    parent_name: str | None
    path: str | None
    generation: int | None
    level: int | None
    object_type: str | None
    property_values: dict[str, str]

@dataclass(frozen=True)
class LoadRowSpec:
    member_name: str
    parent_name: str
    sibling_source_name: str | None

@dataclass(frozen=True)
class RunConfig:
    base_url: str
    app_name: str
    username: str | None
    password: str | None
    password_file: str | None
    oauth_bearer_token: str | None
    verify_ssl: bool
    timeout_seconds: float
    planning_api_version: str
    interop_api_version: str
    poll_interval_seconds: float
    poll_timeout_seconds: float
    working_dir: Path
    similarity_threshold: float

def _as_bool(value: str | None, *, default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}

def _normalize_base_url(raw: str | None) -> str:
    if not raw:
        return ""
    return raw.strip().rstrip("/")

def _normalize_key(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (text or "").lower())

def _member_compare_key(text: str) -> str:
    # Normalized key for safe membership checks across minor formatting differences.
    collapsed = re.sub(r"\s+", " ", (text or "").strip())
    return collapsed.upper()

def _safe_name(text: str) -> str:
    cleaned = re.sub(r"[^\w\-]+", "_", text.strip())
    cleaned = cleaned.strip("_")
    return cleaned or "Unknown"

def _read_password_file(password_file: str | None) -> str | None:
    if not password_file:
        return None
    path = Path(password_file).expanduser()
    if not path.exists():
        return None
    # utf-8-sig removes BOM if the file was saved with UTF-8 BOM.
    value = path.read_text(encoding="utf-8-sig").strip()
    return value or None

def _find_first_existing(paths: list[Path]) -> Path | None:
    for path in paths:
        if path.exists():
            return path
    return None

def _resolve_mapping_path(mapping_csv_path: str | None, prompt_if_missing: bool) -> Path:
    if mapping_csv_path:
        candidate = Path(mapping_csv_path).expanduser()
        if candidate.exists():
            return candidate.resolve()

    cwd = Path.cwd()
    script_dir = Path(__file__).resolve().parent
    defaults = [
        cwd / "FileColtoDim.csv",
        cwd / "Files" / "FileColtoDim.csv",
        script_dir / "FileColtoDim.csv",
        script_dir / "Files" / "FileColtoDim.csv",
    ]
    default_hit = _find_first_existing(defaults)
    if default_hit:
        return default_hit.resolve()

    if prompt_if_missing:
        raw = input(
            "FileColtoDim.csv was not found in default paths. "
            "Please enter full path to FileColtoDim.csv: "
        ).strip()
        if raw:
            prompted = Path(raw).expanduser()
            if prompted.exists():
                return prompted.resolve()

    raise IncrementalMetadataSyncError(
        "Could not find FileColtoDim.csv. Pass mapping_csv_path or place it under ./Files."
    )

def _resolve_default_template_path() -> Path:
    return (Path.cwd() / DEFAULT_TEMPLATE_CSV_REL_PATH).resolve()

def _resolve_cube_name() -> str:
    return (os.getenv("EPBCS_CUBE_NAME") or DEFAULT_CUBE_NAME).strip()

def _read_template_header(template_csv_path: Path) -> list[str]:
    if not template_csv_path.exists():
        raise IncrementalMetadataSyncError(f"Template file not found: {template_csv_path}")
    with template_csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.reader(handle)
        header = next(reader, None)
    if not header:
        raise IncrementalMetadataSyncError(
            f"Template file has no header row: {template_csv_path}"
        )
    return [cell.strip() for cell in header]

def _load_mapping_entries(mapping_csv_path: Path) -> list[MappingEntry]:
    entries: list[MappingEntry] = []
    with mapping_csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.reader(handle)
        for row in reader:
            if not row or not row[0].strip():
                continue
            try:
                column_number = int(row[0].strip())
            except ValueError:
                # Allows header rows in mapping file.
                continue
            if column_number <= 0:
                raise IncrementalMetadataSyncError(
                    f"Invalid column number {column_number} in mapping file {mapping_csv_path}"
                )
            if len(row) < 2 or not row[1].strip():
                raise IncrementalMetadataSyncError(
                    f"Missing dimension name for column {column_number} in {mapping_csv_path}"
                )
            entries.append(
                MappingEntry(
                    column_number=column_number,
                    dimension_name=row[1].strip(),
                )
            )
    if not entries:
        raise IncrementalMetadataSyncError(
            f"No usable mappings were found in {mapping_csv_path}"
        )
    entries.sort(key=lambda item: item.column_number)
    return entries

def _read_source_unique_values(
    source_csv_path: Path, mapping_entries: list[MappingEntry]
) -> tuple[dict[int, set[str]], list[str] | None]:
    if not source_csv_path.exists():
        raise IncrementalMetadataSyncError(f"Source data file not found: {source_csv_path}")

    target_cols = {entry.column_number for entry in mapping_entries}
    unique_by_col = {col: set() for col in target_cols}
    header: list[str] | None = None

    with source_csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.reader(handle)
        for row_index, row in enumerate(reader):
            if row_index == 0:
                header = row
                continue
            for col_num in target_cols:
                col_index = col_num - 1
                if col_index >= len(row):
                    continue
                value = row[col_index].strip()
                if value:
                    unique_by_col[col_num].add(value)
    return unique_by_col, header

def _stringify_property_value(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, str):
        text = value.strip()
        return text or None
    if isinstance(value, list):
        if not value:
            return None
        parts: list[str] = []
        for item in value:
            item_text = _stringify_property_value(item)
            if item_text:
                parts.append(item_text)
        return ", ".join(parts) if parts else None
    return str(value).strip() or None

def _extract_member_property_values(node: dict[str, Any]) -> dict[str, str]:
    properties: dict[str, str] = {}

    def _store_property(raw_key: Any, raw_value: Any) -> None:
        key = _normalize_key(str(raw_key))
        if not key:
            return
        text = _stringify_property_value(raw_value)
        if text is None:
            return
        properties[key] = text

    for key, value in node.items():
        if key in {"children", "links"}:
            continue

        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                _store_property(sub_key, sub_value)
            continue

        if isinstance(value, list):
            if value and all(isinstance(item, dict) for item in value):
                consumed = False
                for item in value:
                    name = item.get("name") or item.get("property") or item.get("key")
                    if not name:
                        continue
                    _store_property(name, item.get("value"))
                    consumed = True
                if consumed:
                    continue
            _store_property(key, value)
            continue

        _store_property(key, value)

    return properties

def _flatten_dimension_members(payload: dict[str, Any]) -> list[ExistingMember]:
    results: list[ExistingMember] = []
    seen: set[str] = set()

    def walk(node: dict[str, Any]) -> None:
        name = str(node.get("name") or "").strip()
        if not name:
            return
        path = node.get("path")
        key = str(path or name)
        if key in seen:
            return
        seen.add(key)
        property_values = _extract_member_property_values(node)
        results.append(
            ExistingMember(
                name=name,
                parent_name=node.get("parentName"),
                path=path,
                generation=node.get("generation"),
                level=node.get("level"),
                object_type=node.get("objectType"),
                property_values=property_values,
            )
        )
        children = node.get("children") or []
        if isinstance(children, list):
            for child in children:
                if isinstance(child, dict):
                    walk(child)

    walk(payload)
    children = payload.get("children") or []
    if isinstance(children, list):
        for child in children:
            if isinstance(child, dict):
                walk(child)
    return results

def _best_parent_by_similarity(
    new_member: str,
    existing_members: list[ExistingMember],
    dimension_name: str,
    threshold: float,
) -> tuple[str | None, dict[str, Any]]:
    candidates = [
        member
        for member in existing_members
        if _normalize_key(member.name) != _normalize_key(dimension_name)
    ]
    if not candidates:
        return None, {
            "reason": "no_candidates",
            "closest_member": None,
            "score": 0.0,
        }

    best_member: ExistingMember | None = None
    best_score = -1.0
    best_numeric_distance = float("inf")
    new_norm = _normalize_key(new_member)

    def numeric_distance(a: str, b: str) -> float:
        if a.isdigit() and b.isdigit() and len(a) == len(b):
            return abs(int(a) - int(b))
        return float("inf")

    for member in candidates:
        score = SequenceMatcher(None, new_norm, _normalize_key(member.name)).ratio()
        distance = numeric_distance(new_member.strip(), member.name.strip())
        if score > best_score:
            best_score = score
            best_member = member
            best_numeric_distance = distance
            continue
        # When similarity ties, prefer the numerically closest code-like member.
        if abs(score - best_score) < 1e-12 and distance < best_numeric_distance:
            best_score = score
            best_member = member
            best_numeric_distance = distance

    if not best_member:
        return None, {
            "reason": "no_closest",
            "closest_member": None,
            "score": 0.0,
        }

    if best_score < threshold:
        return None, {
            "reason": "below_threshold",
            "closest_member": best_member.name,
            "score": round(best_score, 4),
            "numeric_distance": None
            if best_numeric_distance == float("inf")
            else int(best_numeric_distance),
        }

    parent = best_member.parent_name or dimension_name
    return parent, {
        "reason": "matched",
        "closest_member": best_member.name,
        "score": round(best_score, 4),
        "numeric_distance": None
        if best_numeric_distance == float("inf")
        else int(best_numeric_distance),
    }

class EpbcsRestClient:
    """Minimal REST client for required EPBCS operations."""

    def __init__(self, config: RunConfig) -> None:
        self.config = config
        self.config.working_dir.mkdir(parents=True, exist_ok=True)

    def _auth(self) -> httpx.BasicAuth | None:
        if self.config.oauth_bearer_token or not self.config.username:
            return None
        password = self.config.password or _read_password_file(self.config.password_file)
        return httpx.BasicAuth(self.config.username, password) if password else None

    def _headers(self, extra: dict[str, str] | None = None) -> dict[str, str]:
        headers = {
            "Accept": "application/json",
            "User-Agent": "IncrementalMetadata-Sync/1.0",
        }
        if self.config.oauth_bearer_token:
            headers["Authorization"] = f"Bearer {self.config.oauth_bearer_token}"
        if extra:
            headers.update(extra)
        return headers

    def _build_url(self, service: str, path: str) -> str:
        normalized = path if path.startswith("/") else f"/{path}"
        if service == "planning":
            return (
                f"{self.config.base_url}/HyperionPlanning/rest/"
                f"{self.config.planning_api_version}{normalized}"
            )
        if service == "interop":
            return (
                f"{self.config.base_url}/interop/rest/"
                f"{self.config.interop_api_version}{normalized}"
            )
        return f"{self.config.base_url}{normalized}"

    def request(
        self,
        method: str,
        service: str,
        path: str,
        *,
        params: dict[str, str] | None = None,
        json_body: Any | None = None,
        headers: dict[str, str] | None = None,
        content: bytes | None = None,
    ) -> dict[str, Any]:
        url = self._build_url(service, path)
        try:
            with httpx.Client(
                verify=self.config.verify_ssl,
                timeout=self.config.timeout_seconds,
                follow_redirects=True,
            ) as client:
                response = client.request(
                    method=method.upper(),
                    url=url,
                    params=params,
                    json=json_body,
                    headers=self._headers(headers),
                    auth=self._auth(),
                    content=content,
                )
        except httpx.RequestError as exc:
            return self._request_error_result(method, url, exc)

        data = self._response_payload(response)
        ok_http = 200 <= response.status_code < 300
        epm_status, ok_epm = self._epm_status(data)

        return {
            "ok": ok_http and ok_epm,
            "ok_http": ok_http,
            "ok_epm": ok_epm,
            "http_status": response.status_code,
            "epm_status": epm_status,
            "method": method.upper(),
            "url": url,
            "data": data,
        }

    @staticmethod
    def _request_error_result(method: str, url: str, exc: httpx.RequestError) -> dict[str, Any]:
        return {
            "ok": False,
            "ok_http": False,
            "ok_epm": False,
            "http_status": None,
            "epm_status": None,
            "method": method.upper(),
            "url": url,
            "data": None,
            "error": str(exc),
            "error_type": "request_error",
        }

    @staticmethod
    def _response_payload(response: httpx.Response) -> Any:
        content_type = response.headers.get("content-type", "").lower()
        if "application/json" not in content_type:
            return response.text
        try:
            return response.json()
        except ValueError:
            return response.text

    @staticmethod
    def _epm_status(data: Any) -> tuple[int | None, bool]:
        if not isinstance(data, dict) or "status" not in data:
            return None, True
        try:
            status = int(data["status"])
        except (TypeError, ValueError):
            return None, True
        return status, status in {-1, 0}

    @staticmethod
    def _step_failure(step: str, response: dict[str, Any], **extra: Any) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "ok": False,
            "step": step,
            "response": response,
        }
        payload.update(extra)
        return payload

    @staticmethod
    def _payload_status(data: Any) -> int | None:
        if not isinstance(data, dict):
            return None
        try:
            return int(data.get("status"))
        except (TypeError, ValueError):
            return None

    def _wait_for_job(
        self,
        *,
        service: str,
        path: str,
        fetch_error: str,
    ) -> dict[str, Any]:
        started = time.time()
        attempts = 0
        while True:
            attempts += 1
            status_resp = self.request("GET", service, path)
            if not status_resp.get("ok_http"):
                return {
                    "ok": False,
                    "error": fetch_error,
                    "attempts": attempts,
                    "response": status_resp,
                }

            parsed_status = self._payload_status(status_resp.get("data"))
            elapsed_seconds = round(time.time() - started, 2)
            if parsed_status is not None and parsed_status != -1:
                return {
                    "ok": parsed_status == 0,
                    "final_status": parsed_status,
                    "attempts": attempts,
                    "elapsed_seconds": elapsed_seconds,
                    "response": status_resp,
                }
            if elapsed_seconds > self.config.poll_timeout_seconds:
                return {
                    "ok": False,
                    "timed_out": True,
                    "attempts": attempts,
                    "elapsed_seconds": elapsed_seconds,
                    "response": status_resp,
                }
            time.sleep(self.config.poll_interval_seconds)

    @staticmethod
    def _extract_job_id(payload: Any) -> str | None:
        if not isinstance(payload, dict):
            return None
        for key in ("jobId", "jobID", "jobIdentifier", "id"):
            value = payload.get(key)
            if value is not None:
                return str(value)
        patterns = (r"/jobs/([^/?#]+)", r"/files/upload/([^/?#]+)", r"/status/download/([^/?#]+)")
        for link in payload.get("links") or []:
            if not isinstance(link, dict):
                continue
            href = str(link.get("href") or "")
            if not href:
                continue
            for pattern in patterns:
                match = re.search(pattern, href)
                if match:
                    return match.group(1)
        return None

    def get_dimension(self, app_name: str, cube_name: str, dimension_name: str) -> dict[str, Any]:
        return self.request(
            "GET",
            "planning",
            f"/applications/{app_name}/plantypes/{cube_name}/dimensions/{dimension_name}",
        )

    @staticmethod
    def _interop_details_text(response: dict[str, Any]) -> str:
        data = response.get("data")
        return str(data.get("details") if isinstance(data, dict) else data or "")

    @staticmethod
    def _safe_int(value: Any) -> int | None:
        try:
            return int(value)
        except (TypeError, ValueError):
            return None

    def _is_missing_interop_file_response(self, response: dict[str, Any]) -> bool:
        details = self._interop_details_text(response).lower()
        parsed_status = self._safe_int(response.get("epm_status"))
        missing_tokens = (
            "not exist",
            "not found",
            "does not exist",
            "invalid file",
            "invalid filename",
            "no such file",
        )
        if any(token in details for token in missing_tokens):
            return True
        # Some environments return status=4 for delete of a non-existent interop file.
        return parsed_status in {4, 404} and ("file" in details or not details)

    def _create_upload(self, remote_file_name: str, file_size: int) -> dict[str, Any]:
        return self.request(
            "POST",
            "interop",
            "/files/upload",
            json_body={"fileName": remote_file_name, "fileSize": str(file_size)},
        )

    def delete_file_from_interop(
        self,
        remote_file_name: str,
        *,
        ignore_missing: bool = True,
    ) -> dict[str, Any]:
        delete_resp = self.request(
            "DELETE",
            "interop",
            "/files/delete",
            json_body={"fileName": remote_file_name},
        )
        if delete_resp.get("ok"):
            return {"ok": True, "response": delete_resp}

        if ignore_missing and self._is_missing_interop_file_response(delete_resp):
            return {"ok": True, "response": delete_resp, "ignored_missing": True}

        return {"ok": False, "response": delete_resp}

    def _is_upload_conflict(self, create_resp: dict[str, Any]) -> bool:
        details = self._interop_details_text(create_resp).lower()
        return (
            bool(create_resp.get("ok_http"))
            and int(create_resp.get("epm_status") or 0) != 0
            and ("already exists" in details or "upload is in progress" in details)
        )

    def upload_file_to_interop(
        self,
        local_zip_path: Path,
        remote_file_name: str,
        *,
        overwrite: bool = True,
    ) -> dict[str, Any]:
        file_size = local_zip_path.stat().st_size
        overwrite_delete_resp: dict[str, Any] | None = None

        if overwrite:
            overwrite_delete_resp = self.delete_file_from_interop(
                remote_file_name,
                ignore_missing=True,
            )
            if not overwrite_delete_resp.get("ok"):
                return self._step_failure(
                    "overwrite_delete",
                    overwrite_delete_resp,
                )

        create_resp = self._create_upload(remote_file_name, file_size)
        if self._is_upload_conflict(create_resp):
            # Retry once after delete for eventual consistency windows.
            retry_delete = self.delete_file_from_interop(remote_file_name, ignore_missing=True)
            if not retry_delete.get("ok"):
                return self._step_failure(
                    "retry_overwrite_delete",
                    retry_delete,
                    create_response=create_resp,
                )
            time.sleep(1)
            create_resp = self._create_upload(remote_file_name, file_size)
        if not create_resp.get("ok_http"):
            return self._step_failure("create_upload", create_resp)

        upload_job_id = self._extract_job_id(create_resp.get("data"))
        if not upload_job_id:
            return self._step_failure(
                "create_upload",
                create_resp,
                error="Could not derive upload job id",
            )

        file_bytes = local_zip_path.read_bytes()
        upload_resp = self.request(
            "PATCH",
            "interop",
            f"/files/upload/{upload_job_id}",
            headers={
                "Content-Type": "application/octet-stream",
                "Chunk-Range": f"0-{max(0, file_size - 1)}",
            },
            content=file_bytes,
        )
        if not upload_resp.get("ok_http"):
            return self._step_failure(
                "upload_bytes",
                upload_resp,
                job_id=upload_job_id,
            )

        complete_resp = self.request(
            "POST",
            "interop",
            f"/files/upload/{upload_job_id}/complete",
        )
        if not complete_resp.get("ok_http"):
            return self._step_failure(
                "complete_upload",
                complete_resp,
                job_id=upload_job_id,
            )

        wait_resp = self.wait_for_interop_job(upload_job_id)
        return {
            "ok": bool(wait_resp.get("ok")),
            "job_id": upload_job_id,
            "overwrite": overwrite,
            "overwrite_delete_response": overwrite_delete_resp,
            "create_response": create_resp,
            "upload_response": upload_resp,
            "complete_response": complete_resp,
            "wait_response": wait_resp,
        }

    def wait_for_interop_job(self, job_id: str) -> dict[str, Any]:
        return self._wait_for_job(
            service="interop",
            path=f"/status/jobs/{job_id}",
            fetch_error="Could not fetch interop job status",
        )

    def run_import_metadata_job(
        self,
        app_name: str,
        job_name: str,
        zip_file_name: str,
    ) -> dict[str, Any]:
        payload = {"jobType": "IMPORT_METADATA", "jobName": job_name, "parameters": {"importZipFileName": zip_file_name}}
        run_resp = self.request(
            "POST",
            "planning",
            f"/applications/{app_name}/jobs",
            json_body=payload,
        )
        if not run_resp.get("ok_http"):
            return self._step_failure("run_job", run_resp)

        job_id = self._extract_job_id(run_resp.get("data"))
        if not job_id:
            return self._step_failure(
                "run_job",
                run_resp,
                error="Could not derive Planning job id",
            )

        wait_resp = self.wait_for_planning_job(app_name, job_id)
        return {
            "ok": bool(wait_resp.get("ok")),
            "job_id": job_id,
            "run_response": run_resp,
            "wait_response": wait_resp,
        }

    def wait_for_planning_job(self, app_name: str, job_id: str) -> dict[str, Any]:
        return self._wait_for_job(
            service="planning",
            path=f"/applications/{app_name}/jobs/{job_id}",
            fetch_error="Could not fetch Planning job status",
        )

def _resolve_env_value_for_dimension(prefix: str, dimension_name: str) -> str | None:
    dim_norm = _normalize_key(dimension_name)
    normalized_prefix = prefix.upper()
    exact_match: str | None = None
    partial_match: str | None = None

    for key, value in os.environ.items():
        if not key.upper().startswith(normalized_prefix) or not str(value).strip():
            continue
        suffix = key[len(normalized_prefix):]
        suffix_norm = _normalize_key(suffix)
        if suffix_norm == dim_norm:
            exact_match = str(value).strip()
            break
        if suffix_norm.endswith(dim_norm):
            partial_match = str(value).strip()

    return exact_match or partial_match

def _resolve_import_job_name(dimension_name: str) -> str | None:
    return _resolve_env_value_for_dimension("INCIMPORT_", dimension_name)

def _resolve_import_csv_name(dimension_name: str) -> str | None:
    return _resolve_env_value_for_dimension("IMPJOB_FILE_", dimension_name)

def _is_data_storage_key(key: str) -> bool:
    return "datastorage" in _normalize_key(key)

def _build_dimension_header(template_header: list[str], dimension_name: str) -> tuple[list[str], int, int]:
    header = list(template_header)
    if not header:
        header = [dimension_name, "Parent", "Alias: Default"]
    header[0] = dimension_name

    parent_index = -1
    alias_index = -1
    for idx, cell in enumerate(header):
        lowered = cell.strip().lower()
        if lowered == "parent":
            parent_index = idx
        if lowered.startswith("alias"):
            alias_index = idx

    if parent_index < 0:
        header.append("Parent")
        parent_index = len(header) - 1
    if alias_index < 0:
        header.append("Alias: Default")
        alias_index = len(header) - 1
    if not any(_is_data_storage_key(cell) for cell in header):
        header.append("Data Storage")

    return header, parent_index, alias_index

def _load_existing_row_overrides(loadable_csv_path: Path) -> dict[str, dict[str, str]]:
    if not loadable_csv_path.exists():
        return {}

    with loadable_csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.reader(handle)
        header = next(reader, None)
        if not header:
            return {}
        normalized_header = [_normalize_key(cell) for cell in header]
        overrides: dict[str, dict[str, str]] = {}

        for row in reader:
            if not row:
                continue
            member_name = row[0].strip() if len(row) > 0 else ""
            if not member_name:
                continue
            row_values: dict[str, str] = {}
            for idx, raw_cell in enumerate(row):
                if idx >= len(normalized_header):
                    break
                key = normalized_header[idx]
                if not key:
                    continue
                value = raw_cell.strip()
                if value:
                    row_values[key] = value
            if row_values:
                overrides[_member_compare_key(member_name)] = row_values

    return overrides

def _lookup_property_value(
    property_values: dict[str, str],
    header_key: str,
) -> str | None:
    if not header_key:
        return None
    if header_key in property_values:
        return property_values[header_key]
    # Small tolerance for singular/plural header/property differences.
    if header_key.endswith("s"):
        singular = header_key[:-1]
        if singular in property_values:
            return property_values[singular]
    plural = f"{header_key}s"
    if plural in property_values:
        return property_values[plural]
    return None

def _write_loadable_csv(
    output_path: Path,
    header: list[str],
    parent_index: int,
    alias_index: int,
    row_specs: list[LoadRowSpec],
    existing_member_by_key: dict[str, ExistingMember],
    existing_row_overrides: dict[str, dict[str, str]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        width = len(header)
        header_keys = [_normalize_key(cell) for cell in header]
        data_storage_indexes = [
            idx for idx, key in enumerate(header_keys) if _is_data_storage_key(key)
        ]

        for row_spec in row_specs:
            row = [""] * width
            member_key = _member_compare_key(row_spec.member_name)

            explicit_values = existing_row_overrides.get(member_key, {})
            for idx, key in enumerate(header_keys):
                if not key:
                    continue
                explicit = explicit_values.get(key)
                if explicit:
                    row[idx] = explicit

            sibling_properties: dict[str, str] = {}
            sibling_name = row_spec.sibling_source_name
            if sibling_name:
                sibling = existing_member_by_key.get(_member_compare_key(sibling_name))
                if sibling:
                    sibling_properties = sibling.property_values

            if sibling_properties:
                for idx, key in enumerate(header_keys):
                    if idx in {0, parent_index, alias_index}:
                        continue
                    if not key or row[idx].strip():
                        continue
                    sibling_value = _lookup_property_value(sibling_properties, key)
                    if sibling_value:
                        row[idx] = sibling_value

            row[0] = row_spec.member_name
            row[parent_index] = row_spec.parent_name
            if not row[alias_index].strip():
                row[alias_index] = row_spec.member_name
            for index in data_storage_indexes:
                row[index] = "Store"
            writer.writerow(row)

def _zip_and_delete_csv(csv_path: Path) -> Path:
    zip_path = csv_path.with_suffix(".zip")
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_handle:
        zip_handle.write(csv_path, arcname=csv_path.name)
    csv_path.unlink(missing_ok=False)
    return zip_path

def _cleanup_old_validation_required_files(working_dir: Path, *, days_to_keep: int = 7) -> int:
    deleted = 0
    cutoff = datetime.now() - timedelta(days=days_to_keep)
    for path in working_dir.glob("Validation required_*.txt"):
        try:
            modified = datetime.fromtimestamp(path.stat().st_mtime)
        except OSError:
            continue
        if modified < cutoff:
            try:
                path.unlink()
                deleted += 1
            except OSError:
                pass
    return deleted

def _write_validation_required_note(
    working_dir: Path,
    validation_records: list[dict[str, Any]],
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"Validation required_{timestamp}.txt"
    note_path = working_dir / file_name

    lines: list[str] = [
        "Validation Required - New Member Parent Review",
        f"Generated At: {datetime.now().isoformat(timespec='seconds')}",
        "",
        "Review the member-parent mappings below, then approve in the same running session.",
        "",
    ]

    if not validation_records:
        lines.append("No validation records were available.")
    else:
        for idx, record in enumerate(validation_records, start=1):
            lines.append(f"{idx}. Dimension: {record.get('dimension_name')}")
            lines.append(f"   Zip File: {record.get('zip_file_name')}")
            pairs = record.get("child_parent_pairs") or []
            if not pairs:
                lines.append("   - No member-parent rows.")
            else:
                for pair in pairs:
                    lines.append(f"   - {pair.get('child')} -> {pair.get('parent')}")
            lines.append("")

    note_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
    return note_path

def _wait_for_human_validation_for_run(
    validation_records: list[dict[str, Any]],
    note_path: Path,
    *,
    pending_action_count: int,
    dry_run: bool,
) -> dict[str, Any]:
    dimension_count = len(validation_records)
    mapping_count = sum(len((record.get("child_parent_pairs") or [])) for record in validation_records)
    if dry_run:
        proceed_line = "[Validation] Do you want to proceed and complete this dry-run execution?"
    elif pending_action_count > 0:
        proceed_line = "[Validation] Do you want to proceed with next steps (upload + import jobs)?"
    else:
        proceed_line = "[Validation] Do you want to proceed and complete this execution? (No upload/import jobs are pending.)"
    prompt_lines = [
        "",
        "[Validation] Consolidated approval required for this run.",
        f"[Validation] Dimensions requiring validation: {dimension_count}",
        f"[Validation] Total new member-parent mappings: {mapping_count}",
        f"[Validation] Review file: {note_path}",
        proceed_line,
        "[Validation] Enter Yes/No and press Enter to continue: ",
    ]
    note_content = ""
    try:
        note_content = note_path.read_text(encoding="utf-8").strip()
    except OSError:
        note_content = ""

    if note_content:
        print("\n[Validation] -------- Validation Details --------")
        print(note_content)
        print("[Validation] -----------------------------------")

    next_prompt = "\n".join(prompt_lines)
    while True:
        try:
            raw = input(next_prompt)
        except EOFError:
            return {
                "approved": False,
                "reason": "no_input_stream",
                "raw_input": None,
            }

        response = (raw or "").strip().lower()
        if response in {"yes", "y"}:
            return {"approved": True, "raw_input": raw}
        if response in {"no", "n"}:
            return {
                "approved": False,
                "reason": "rejected_by_user",
                "raw_input": raw,
            }
        print("[Validation] Invalid response. Please type Yes or No.")
        next_prompt = "[Validation] Please enter Yes/No and press Enter: "

def _build_runtime_config(
    app_name_override: str | None,
    working_dir_override: str | None,
    similarity_threshold: float,
) -> RunConfig:
    base_url = _normalize_base_url(os.getenv("EPBCS_BASE_URL") or os.getenv("EPBCS_URL"))
    app_name = (app_name_override or os.getenv("EPBCS_APP_NAME") or "").strip()
    if not base_url:
        raise IncrementalMetadataSyncError("Missing EPBCS_BASE_URL (or EPBCS_URL) in environment.")
    if not app_name:
        raise IncrementalMetadataSyncError("Missing EPBCS_APP_NAME in environment and function input.")

    username = os.getenv("EPBCS_USERNAME")
    password = os.getenv("EPBCS_PASSWORD")
    password_file = os.getenv("EPBCS_PASSWORD_FILE")
    oauth_token = os.getenv("EPBCS_OAUTH_BEARER_TOKEN")
    if not oauth_token and not (username and (password or _read_password_file(password_file))):
        raise IncrementalMetadataSyncError(
            "Authentication is missing. Set EPBCS_OAUTH_BEARER_TOKEN or EPBCS_USERNAME "
            "with EPBCS_PASSWORD/EPBCS_PASSWORD_FILE."
        )

    if working_dir_override:
        work_dir = Path(working_dir_override).expanduser().resolve()
    else:
        work_dir = Path(os.getenv("EPBCS_WORK_DIR", "Files")).expanduser()
        if not work_dir.is_absolute():
            work_dir = (Path.cwd() / work_dir).resolve()

    return RunConfig(
        base_url=base_url,
        app_name=app_name,
        username=username,
        password=password,
        password_file=password_file,
        oauth_bearer_token=oauth_token,
        verify_ssl=_as_bool(os.getenv("EPBCS_VERIFY_SSL"), default=True),
        timeout_seconds=float(os.getenv("EPBCS_REQUEST_TIMEOUT_SEC", "30")),
        planning_api_version=os.getenv("EPBCS_API_VERSION", "v3"),
        interop_api_version=os.getenv("EPBCS_INTEROP_API_VERSION", "v2"),
        poll_interval_seconds=float(os.getenv("EPBCS_POLL_INTERVAL_SECONDS", "2")),
        poll_timeout_seconds=float(os.getenv("EPBCS_POLL_TIMEOUT_SECONDS", "600")),
        working_dir=work_dir,
        similarity_threshold=similarity_threshold,
    )

def _assert_environment_available(api: EpbcsRestClient) -> None:
    """Abort early when EPBCS environment is unreachable/down."""
    probe = api.request("GET", "planning", "/applications")
    http_status = probe.get("http_status")

    # 200: available and authenticated
    # 401/403: environment reachable, auth issue (not "down")
    if http_status in {200, 401, 403}:
        return

    detail = probe.get("error") or probe.get("data")
    raise IncrementalMetadataSyncError(
        f"Environemnt is not available or down. "
        f"error_code=ENV_UNAVAILABLE, http_status={http_status}, detail={detail}",
        error_code="ENV_UNAVAILABLE",
        exit_code=23,
    )

def _require_dimension_mapping(
    value: str | None,
    *,
    env_prefix: str,
    mapping_kind: str,
    dimension_name: str,
    example_value: str,
) -> str:
    if value:
        return value
    raise IncrementalMetadataSyncError(
        f"Missing {env_prefix}<Dimension> {mapping_kind} for '{dimension_name}'. "
        f"Add an environment entry like {env_prefix}{_safe_name(dimension_name)}=\"{example_value}\"."
    )

def _build_member_plan(
    new_members: list[str],
    existing_members: list[ExistingMember],
    existing_keys: set[str],
    *,
    dimension_name: str,
    similarity_threshold: float,
) -> tuple[list[tuple[str, str]], list[LoadRowSpec], list[dict[str, Any]]]:
    placeholder_name = f"TobeAdded_{dimension_name}"
    placeholder_exists = _member_compare_key(placeholder_name) in existing_keys
    placeholder_inserted_in_file = False

    child_parent_rows: list[tuple[str, str]] = []
    load_row_specs: list[LoadRowSpec] = []
    member_decisions: list[dict[str, Any]] = []

    def add_row(
        member: str,
        parent: str,
        *,
        method: str,
        sibling_source_name: str | None,
        details: dict[str, Any] | None = None,
    ) -> None:
        child_parent_rows.append((member, parent))
        load_row_specs.append(
            LoadRowSpec(
                member_name=member,
                parent_name=parent,
                sibling_source_name=sibling_source_name,
            )
        )
        decision: dict[str, Any] = {
            "member": member,
            "parent": parent,
            "method": method,
        }
        if details is not None:
            decision["details"] = details
        member_decisions.append(decision)

    for member_name in new_members:
        parent_name, reasoning = _best_parent_by_similarity(
            member_name,
            existing_members,
            dimension_name=dimension_name,
            threshold=similarity_threshold,
        )
        sibling_source_name = str(reasoning.get("closest_member") or "").strip() or None
        if parent_name:
            add_row(
                member_name,
                parent_name,
                method="closest_match_parent",
                sibling_source_name=sibling_source_name,
                details=reasoning,
            )
            continue

        if not placeholder_exists and not placeholder_inserted_in_file:
            add_row(
                placeholder_name,
                dimension_name,
                method="create_fallback_parent",
                sibling_source_name=sibling_source_name,
            )
            placeholder_inserted_in_file = True

        add_row(
            member_name,
            placeholder_name,
            method="fallback_parent",
            sibling_source_name=sibling_source_name,
            details=reasoning,
        )
    return child_parent_rows, load_row_specs, member_decisions

def _prepare_dimension_action(
    *,
    api: EpbcsRestClient,
    config: RunConfig,
    cube_name: str,
    entry: MappingEntry,
    col_values: set[str],
    template_header: list[str],
) -> tuple[dict[str, Any], dict[str, Any] | None, dict[str, Any] | None]:
    dimension = entry.dimension_name
    dimension_result: dict[str, Any] = {
        "column_number": entry.column_number,
        "dimension_name": dimension,
        "source_unique_count": len(col_values),
        "status": "skipped_no_values",
    }
    if not col_values:
        return dimension_result, None, None

    dim_resp = api.get_dimension(config.app_name, cube_name, dimension)
    if not dim_resp.get("ok"):
        raise IncrementalMetadataSyncError(
            f"Failed to fetch dimension '{dimension}' in cube '{cube_name}'. "
            f"HTTP={dim_resp.get('http_status')} data={dim_resp.get('data')}"
        )
    dim_payload = dim_resp.get("data")
    if not isinstance(dim_payload, dict):
        raise IncrementalMetadataSyncError(
            f"Unexpected dimension payload for '{dimension}': {type(dim_payload)}"
        )

    existing_members = _flatten_dimension_members(dim_payload)
    existing_names = {member.name for member in existing_members}
    existing_keys = {_member_compare_key(member.name) for member in existing_members}
    existing_member_by_key = {
        _member_compare_key(member.name): member for member in existing_members
    }
    new_members = sorted(
        value for value in col_values if _member_compare_key(value) not in existing_keys
    )
    dimension_result.update(
        {
            "existing_member_count": len(existing_names),
            "new_members": new_members,
            "new_member_count": len(new_members),
        }
    )
    if not new_members:
        dimension_result["status"] = "no_new_members"
        return dimension_result, None, None

    job_name = _require_dimension_mapping(
        _resolve_import_job_name(dimension),
        env_prefix="INCIMPORT_",
        mapping_kind="job mapping",
        dimension_name=dimension,
        example_value="Job Name",
    )
    child_parent_rows, load_row_specs, member_decisions = _build_member_plan(
        new_members,
        existing_members,
        existing_keys,
        dimension_name=dimension,
        similarity_threshold=config.similarity_threshold,
    )

    header, parent_idx, alias_idx = _build_dimension_header(template_header, dimension)
    csv_name = _require_dimension_mapping(
        _resolve_import_csv_name(dimension),
        env_prefix="IMPJOB_FILE_",
        mapping_kind="csv mapping",
        dimension_name=dimension,
        example_value="FileName.csv",
    )
    csv_path = config.working_dir / csv_name
    _write_loadable_csv(
        csv_path,
        header,
        parent_idx,
        alias_idx,
        load_row_specs,
        existing_member_by_key,
        _load_existing_row_overrides(csv_path),
    )

    zip_path = _zip_and_delete_csv(csv_path)
    validation_record = {
        "dimension_name": dimension,
        "zip_file_name": zip_path.name,
        "zip_file_path": str(zip_path),
        "child_parent_pairs": [
            {"child": child, "parent": parent} for child, parent in child_parent_rows
        ],
    }
    dimension_result.update(
        {
            "status": "prepared_zip",
            "import_job_name": job_name,
            "epm_zip_file": zip_path.name,
            "local_zip_path": str(zip_path),
            "member_decisions": member_decisions,
        }
    )
    return dimension_result, {
        "dimension": dimension,
        "job_name": job_name,
        "zip_path": zip_path,
        "dimension_result": dimension_result,
    }, validation_record

def _raise_validation_rejected(validation_feedback: dict[str, Any], note_path: Path) -> None:
    reason = validation_feedback.get("reason")
    if reason == "no_input_stream":
        raise IncrementalMetadataSyncError(
            "No console input stream was available for human validation. "
            "Run in an interactive terminal so the program can wait for Yes/No input. "
            f"Validation required note generated at: {note_path}",
            error_code="VALIDATION_INPUT_UNAVAILABLE",
            exit_code=24,
        )
    if reason == "rejected_by_user":
        raise IncrementalMetadataSyncError(
            "Human validation rejected by user. Program aborted before upload/job execution.",
            error_code="VALIDATION_ABORTED_BY_USER",
            exit_code=26,
        )
    raise IncrementalMetadataSyncError(
        "Invalid response received for human validation. Program aborted.",
        error_code="VALIDATION_REJECTED",
        exit_code=25,
    )

def _job_result_summary(result: dict[str, Any]) -> dict[str, Any]:
    wait_response = (result.get("wait_response", {}) or {})
    return {"job_id": result.get("job_id"), "final_status": wait_response.get("final_status")}

def pbcs_run_incmetadata_sync(
    source_csv_path: str,
    mapping_csv_path: str | None = None,
    app_name: str | None = None,
    working_dir: str | None = None,
    similarity_threshold: float = 0.72,
    dry_run: bool = False,
    prompt_for_missing_mapping: bool = True,
    require_human_validation: bool = True,
    human_validation_timeout_seconds: int = 300,
) -> dict[str, Any]:
    """Incremental metadata sync runner."""
    # Backward compatibility: manual validation is blocking, mandatory for every run, and no longer time-limited.
    _ = require_human_validation
    _ = human_validation_timeout_seconds

    load_dotenv(override=False)

    config = _build_runtime_config(
        app_name_override=app_name,
        working_dir_override=working_dir,
        similarity_threshold=similarity_threshold,
    )
    api = EpbcsRestClient(config)
    _assert_environment_available(api)

    source_path = Path(source_csv_path).expanduser().resolve()
    template_path = _resolve_default_template_path()
    cube_name = _resolve_cube_name()
    mapping_path = _resolve_mapping_path(mapping_csv_path, prompt_if_missing=prompt_for_missing_mapping)

    deleted_old_notes = _cleanup_old_validation_required_files(config.working_dir, days_to_keep=7)

    mapping_entries = _load_mapping_entries(mapping_path)
    unique_by_col, _ = _read_source_unique_values(source_path, mapping_entries)
    template_header = _read_template_header(template_path)
    validation_records: list[dict[str, Any]] = []

    summary: dict[str, Any] = {
        "ok": True,
        "app_name": config.app_name,
        "cube_name": cube_name,
        "source_csv_path": str(source_path),
        "mapping_csv_path": str(mapping_path),
        "template_csv_path": str(template_path),
        "dry_run": dry_run,
        "human_validation_enabled": True,
        "human_validation_mode": "blocking_yes_no",
        "old_validation_note_files_deleted": deleted_old_notes,
        "dimensions": [],
    }
    pending_upload_actions: list[dict[str, Any]] = []

    for entry in mapping_entries:
        dimension_result, pending_action, validation_record = _prepare_dimension_action(
            api=api,
            config=config,
            cube_name=cube_name,
            entry=entry,
            col_values=unique_by_col.get(entry.column_number, set()),
            template_header=template_header,
        )
        summary["dimensions"].append(dimension_result)
        if pending_action:
            pending_upload_actions.append(pending_action)
        if validation_record:
            validation_records.append(validation_record)

    note_path = _write_validation_required_note(config.working_dir, validation_records)
    summary["human_validation_note_path"] = str(note_path)
    validation_feedback = _wait_for_human_validation_for_run(
        validation_records=validation_records,
        note_path=note_path,
        pending_action_count=len(pending_upload_actions),
        dry_run=dry_run,
    )
    summary["human_validation"] = validation_feedback
    for action in pending_upload_actions:
        action["dimension_result"]["human_validation"] = validation_feedback
    if not validation_feedback.get("approved"):
        _raise_validation_rejected(validation_feedback, note_path)

    if dry_run:
        for action in pending_upload_actions:
            action["dimension_result"]["status"] = "dry_run_prepared_only"
        return summary

    for action in pending_upload_actions:
        dimension = action["dimension"]
        zip_path: Path = action["zip_path"]
        job_name = action["job_name"]
        dimension_result: dict[str, Any] = action["dimension_result"]

        upload_result = api.upload_file_to_interop(zip_path, zip_path.name)
        if not upload_result.get("ok"):
            raise IncrementalMetadataSyncError(
                f"Failed to upload {zip_path.name} for dimension {dimension}: {upload_result}"
            )
        # Upload is complete in EPBCS Interop; remove local zip to keep workspace clean.
        zip_path.unlink(missing_ok=True)
        dimension_result["local_zip_deleted"] = True

        run_result = api.run_import_metadata_job(
            app_name=config.app_name,
            job_name=job_name,
            zip_file_name=zip_path.name,
        )
        if not run_result.get("ok"):
            raise IncrementalMetadataSyncError(
                f"Import metadata job failed for dimension {dimension}: {run_result}"
            )

        dimension_result["status"] = "completed"
        dimension_result["upload_result"] = _job_result_summary(upload_result)
        dimension_result["job_result"] = _job_result_summary(run_result)

    return summary

def main() -> None:
    """CLI helper for local execution."""
    import argparse

    parser = argparse.ArgumentParser(description="Run EPBCS incremental metadata sync.")
    add = parser.add_argument
    add("--source", required=True, help="Path to source data CSV file")
    add("--mapping", default=None, help="Path to FileColtoDim.csv")
    add("--app", default=None, help="Override EPBCS app name")
    add("--workdir", default=None, help="Working directory for generated files")
    add("--similarity-threshold", type=float, default=0.72, help="Similarity threshold for parent derivation")
    add("--dry-run", action="store_true", help="Only create ZIP payloads locally, do not upload or run jobs")
    add("--no-prompt", action="store_true", help="Do not prompt when FileColtoDim.csv is missing")
    add(
        "--human-validation-timeout-seconds",
        type=int,
        default=300,
        help="Deprecated: ignored because validation is mandatory and blocks until Yes/No input",
    )

    args = parser.parse_args()
    try:
        result = pbcs_run_incmetadata_sync(
            source_csv_path=args.source,
            mapping_csv_path=args.mapping,
            app_name=args.app,
            working_dir=args.workdir,
            similarity_threshold=args.similarity_threshold,
            dry_run=args.dry_run,
            prompt_for_missing_mapping=not args.no_prompt,
            require_human_validation=True,
            human_validation_timeout_seconds=args.human_validation_timeout_seconds,
        )
        print(json.dumps(result, indent=2))
    except IncrementalMetadataSyncError as exc:
        print(json.dumps({"ok": False, "error_code": getattr(exc, "error_code", "INC_METADATA_SYNC_ERROR"), "message": str(exc)}, indent=2))
        raise SystemExit(getattr(exc, "exit_code", 1))
    except Exception as exc:  # pragma: no cover
        print(json.dumps({"ok": False, "error_code": "UNHANDLED_ERROR", "message": str(exc)}, indent=2))
        raise SystemExit(1)

if __name__ == "__main__":
    main()
