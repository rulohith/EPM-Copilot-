from __future__ import annotations

import csv
import json
import logging
import os
import re
import sys
import time
import zipfile
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
        results.append(
            ExistingMember(
                name=name,
                parent_name=node.get("parentName"),
                path=path,
                generation=node.get("generation"),
                level=node.get("level"),
                object_type=node.get("objectType"),
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
    new_norm = _normalize_key(new_member)

    for member in candidates:
        score = SequenceMatcher(None, new_norm, _normalize_key(member.name)).ratio()
        if score > best_score:
            best_score = score
            best_member = member

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
        }

    parent = best_member.parent_name or dimension_name
    return parent, {
        "reason": "matched",
        "closest_member": best_member.name,
        "score": round(best_score, 4),
    }


class EpbcsRestClient:
    """Minimal REST client for required EPBCS operations."""

    def __init__(self, config: RunConfig) -> None:
        self.config = config
        self.config.working_dir.mkdir(parents=True, exist_ok=True)

    def _auth(self) -> httpx.BasicAuth | None:
        if self.config.oauth_bearer_token:
            return None
        if not self.config.username:
            return None
        password = self.config.password or _read_password_file(self.config.password_file)
        if not password:
            return None
        return httpx.BasicAuth(self.config.username, password)

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

        content_type = response.headers.get("content-type", "").lower()
        if "application/json" in content_type:
            try:
                data = response.json()
            except ValueError:
                data = response.text
        else:
            data = response.text

        ok_http = 200 <= response.status_code < 300
        epm_status = None
        ok_epm = True
        if isinstance(data, dict) and "status" in data:
            try:
                epm_status = int(data["status"])
                ok_epm = epm_status in {-1, 0}
            except (TypeError, ValueError):
                ok_epm = True

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
    def _extract_job_id(payload: Any) -> str | None:
        if not isinstance(payload, dict):
            return None
        for key in ("jobId", "jobID", "jobIdentifier", "id"):
            value = payload.get(key)
            if value is not None:
                return str(value)
        links = payload.get("links")
        if isinstance(links, list):
            for link in links:
                if not isinstance(link, dict):
                    continue
                href = str(link.get("href") or "")
                if not href:
                    continue
                match = re.search(r"/jobs/([^/?#]+)", href)
                if match:
                    return match.group(1)
                match = re.search(r"/files/upload/([^/?#]+)", href)
                if match:
                    return match.group(1)
                match = re.search(r"/status/download/([^/?#]+)", href)
                if match:
                    return match.group(1)
        return None

    def get_dimension(self, app_name: str, cube_name: str, dimension_name: str) -> dict[str, Any]:
        return self.request(
            "GET",
            "planning",
            f"/applications/{app_name}/plantypes/{cube_name}/dimensions/{dimension_name}",
        )

    def upload_file_to_interop(self, local_zip_path: Path, remote_file_name: str) -> dict[str, Any]:
        file_size = local_zip_path.stat().st_size
        create_resp = self.request(
            "POST",
            "interop",
            "/files/upload",
            json_body={"fileName": remote_file_name, "fileSize": str(file_size)},
        )
        if not create_resp.get("ok_http"):
            return {"ok": False, "step": "create_upload", "response": create_resp}

        upload_job_id = self._extract_job_id(create_resp.get("data"))
        if not upload_job_id:
            return {
                "ok": False,
                "step": "create_upload",
                "error": "Could not derive upload job id",
                "response": create_resp,
            }

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
            return {
                "ok": False,
                "step": "upload_bytes",
                "job_id": upload_job_id,
                "response": upload_resp,
            }

        complete_resp = self.request(
            "POST",
            "interop",
            f"/files/upload/{upload_job_id}/complete",
        )
        if not complete_resp.get("ok_http"):
            return {
                "ok": False,
                "step": "complete_upload",
                "job_id": upload_job_id,
                "response": complete_resp,
            }

        wait_resp = self.wait_for_interop_job(upload_job_id)
        return {
            "ok": bool(wait_resp.get("ok")),
            "job_id": upload_job_id,
            "create_response": create_resp,
            "upload_response": upload_resp,
            "complete_response": complete_resp,
            "wait_response": wait_resp,
        }

    def wait_for_interop_job(self, job_id: str) -> dict[str, Any]:
        started = time.time()
        attempts = 0
        while True:
            attempts += 1
            status_resp = self.request("GET", "interop", f"/status/jobs/{job_id}")
            if not status_resp.get("ok_http"):
                return {
                    "ok": False,
                    "error": "Could not fetch interop job status",
                    "attempts": attempts,
                    "response": status_resp,
                }

            data = status_resp.get("data")
            parsed_status = None
            if isinstance(data, dict):
                raw = data.get("status")
                try:
                    parsed_status = int(raw)
                except (TypeError, ValueError):
                    parsed_status = None
            if parsed_status is not None and parsed_status != -1:
                return {
                    "ok": parsed_status == 0,
                    "final_status": parsed_status,
                    "attempts": attempts,
                    "elapsed_seconds": round(time.time() - started, 2),
                    "response": status_resp,
                }

            if time.time() - started > self.config.poll_timeout_seconds:
                return {
                    "ok": False,
                    "timed_out": True,
                    "attempts": attempts,
                    "elapsed_seconds": round(time.time() - started, 2),
                    "response": status_resp,
                }
            time.sleep(self.config.poll_interval_seconds)

    def run_import_metadata_job(
        self,
        app_name: str,
        job_name: str,
        zip_file_name: str,
    ) -> dict[str, Any]:
        payload = {
            "jobType": "IMPORT_METADATA",
            "jobName": job_name,
            "parameters": {
                "importZipFileName": zip_file_name,
            },
        }
        run_resp = self.request(
            "POST",
            "planning",
            f"/applications/{app_name}/jobs",
            json_body=payload,
        )
        if not run_resp.get("ok_http"):
            return {"ok": False, "step": "run_job", "response": run_resp}

        job_id = self._extract_job_id(run_resp.get("data"))
        if not job_id:
            return {
                "ok": False,
                "step": "run_job",
                "error": "Could not derive Planning job id",
                "response": run_resp,
            }

        wait_resp = self.wait_for_planning_job(app_name, job_id)
        return {
            "ok": bool(wait_resp.get("ok")),
            "job_id": job_id,
            "run_response": run_resp,
            "wait_response": wait_resp,
        }

    def wait_for_planning_job(self, app_name: str, job_id: str) -> dict[str, Any]:
        started = time.time()
        attempts = 0
        while True:
            attempts += 1
            status_resp = self.request(
                "GET",
                "planning",
                f"/applications/{app_name}/jobs/{job_id}",
            )
            if not status_resp.get("ok_http"):
                return {
                    "ok": False,
                    "error": "Could not fetch Planning job status",
                    "attempts": attempts,
                    "response": status_resp,
                }

            data = status_resp.get("data")
            parsed_status = None
            if isinstance(data, dict):
                raw = data.get("status")
                try:
                    parsed_status = int(raw)
                except (TypeError, ValueError):
                    parsed_status = None

            if parsed_status is not None and parsed_status != -1:
                return {
                    "ok": parsed_status == 0,
                    "final_status": parsed_status,
                    "attempts": attempts,
                    "elapsed_seconds": round(time.time() - started, 2),
                    "response": status_resp,
                }

            if time.time() - started > self.config.poll_timeout_seconds:
                return {
                    "ok": False,
                    "timed_out": True,
                    "attempts": attempts,
                    "elapsed_seconds": round(time.time() - started, 2),
                    "response": status_resp,
                }
            time.sleep(self.config.poll_interval_seconds)


def _resolve_import_job_name(dimension_name: str) -> str | None:
    dim_norm = _normalize_key(dimension_name)
    prefix = "INCIMPORT_"
    exact_match: str | None = None
    partial_match: str | None = None

    for key, value in os.environ.items():
        if not key.upper().startswith(prefix) or not str(value).strip():
            continue
        suffix = key[len(prefix):]
        suffix_norm = _normalize_key(suffix)
        if suffix_norm == dim_norm:
            exact_match = str(value).strip()
            break
        if suffix_norm.endswith(dim_norm):
            partial_match = str(value).strip()

    return exact_match or partial_match


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

    return header, parent_index, alias_index


def _write_loadable_csv(
    output_path: Path,
    header: list[str],
    parent_index: int,
    alias_index: int,
    child_parent_pairs: list[tuple[str, str]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        width = len(header)
        for child, parent in child_parent_pairs:
            row = [""] * width
            row[0] = child
            row[parent_index] = parent
            row[alias_index] = child
            writer.writerow(row)


def _zip_and_delete_csv(csv_path: Path) -> Path:
    zip_path = csv_path.with_suffix(".zip")
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_handle:
        zip_handle.write(csv_path, arcname=csv_path.name)
    csv_path.unlink(missing_ok=False)
    return zip_path


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
    """
    First-step health check.
    If EPBCS environment is down/unreachable, abort early with explicit error code.
    """
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


def pbcs_run_incmetadata_sync(
    source_csv_path: str,
    template_csv_path: str,
    cube_name: str,
    mapping_csv_path: str | None = None,
    app_name: str | None = None,
    working_dir: str | None = None,
    similarity_threshold: float = 0.72,
    dry_run: bool = False,
    prompt_for_missing_mapping: bool = True,
) -> dict[str, Any]:
    """
    Incremental metadata sync runner.

    Call format requested:
    pbcs_run_incmetadata_sync(source_csv_path=..., template_csv_path=..., cube_name=...)
    """
    load_dotenv(override=False)

    config = _build_runtime_config(
        app_name_override=app_name,
        working_dir_override=working_dir,
        similarity_threshold=similarity_threshold,
    )
    api = EpbcsRestClient(config)
    _assert_environment_available(api)

    source_path = Path(source_csv_path).expanduser().resolve()
    template_path = Path(template_csv_path).expanduser().resolve()
    mapping_path = _resolve_mapping_path(mapping_csv_path, prompt_if_missing=prompt_for_missing_mapping)

    mapping_entries = _load_mapping_entries(mapping_path)
    unique_by_col, source_header = _read_source_unique_values(source_path, mapping_entries)
    template_header = _read_template_header(template_path)

    summary: dict[str, Any] = {
        "ok": True,
        "app_name": config.app_name,
        "cube_name": cube_name,
        "source_csv_path": str(source_path),
        "mapping_csv_path": str(mapping_path),
        "template_csv_path": str(template_path),
        "dry_run": dry_run,
        "dimensions": [],
    }

    for entry in mapping_entries:
        dimension = entry.dimension_name
        col_values = unique_by_col.get(entry.column_number, set())

        dimension_result: dict[str, Any] = {
            "column_number": entry.column_number,
            "dimension_name": dimension,
            "source_unique_count": len(col_values),
            "status": "skipped_no_values",
        }

        if not col_values:
            summary["dimensions"].append(dimension_result)
            continue

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
        new_members = sorted(value for value in col_values if value not in existing_names)

        dimension_result["existing_member_count"] = len(existing_names)
        dimension_result["new_members"] = new_members
        dimension_result["new_member_count"] = len(new_members)

        if not new_members:
            dimension_result["status"] = "no_new_members"
            summary["dimensions"].append(dimension_result)
            continue

        job_name = _resolve_import_job_name(dimension)
        if not job_name:
            raise IncrementalMetadataSyncError(
                f"Missing INCIMPORT_<Dimension> job mapping for '{dimension}'. "
                f"Add an environment entry like INCIMPORT_{_safe_name(dimension)}=\"Job Name\"."
            )

        placeholder_name = f"TobeAdded_{dimension}"
        placeholder_exists = placeholder_name in existing_names
        placeholder_inserted_in_file = False

        child_parent_rows: list[tuple[str, str]] = []
        member_decisions: list[dict[str, Any]] = []
        for member_name in new_members:
            parent_name, reasoning = _best_parent_by_similarity(
                member_name,
                existing_members,
                dimension_name=dimension,
                threshold=config.similarity_threshold,
            )
            if parent_name:
                child_parent_rows.append((member_name, parent_name))
                member_decisions.append(
                    {
                        "member": member_name,
                        "parent": parent_name,
                        "method": "closest_match_parent",
                        "details": reasoning,
                    }
                )
                continue

            if not placeholder_exists and not placeholder_inserted_in_file:
                child_parent_rows.append((placeholder_name, dimension))
                placeholder_inserted_in_file = True
                member_decisions.append(
                    {
                        "member": placeholder_name,
                        "parent": dimension,
                        "method": "create_fallback_parent",
                    }
                )
            child_parent_rows.append((member_name, placeholder_name))
            member_decisions.append(
                {
                    "member": member_name,
                    "parent": placeholder_name,
                    "method": "fallback_parent",
                    "details": reasoning,
                }
            )

        header, parent_idx, alias_idx = _build_dimension_header(template_header, dimension)
        csv_name = f"EPMLoadable_{_safe_name(dimension)}.csv"
        csv_path = config.working_dir / csv_name
        _write_loadable_csv(csv_path, header, parent_idx, alias_idx, child_parent_rows)

        zip_path = _zip_and_delete_csv(csv_path)

        dimension_result.update(
            {
                "status": "prepared_zip",
                "import_job_name": job_name,
                "epm_zip_file": zip_path.name,
                "local_zip_path": str(zip_path),
                "member_decisions": member_decisions,
            }
        )

        if dry_run:
            dimension_result["status"] = "dry_run_prepared_only"
            summary["dimensions"].append(dimension_result)
            continue

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
        dimension_result["upload_result"] = {
            "job_id": upload_result.get("job_id"),
            "final_status": (
                upload_result.get("wait_response", {}) or {}
            ).get("final_status"),
        }
        dimension_result["job_result"] = {
            "job_id": run_result.get("job_id"),
            "final_status": (
                run_result.get("wait_response", {}) or {}
            ).get("final_status"),
        }
        summary["dimensions"].append(dimension_result)

    return summary


def main() -> None:
    """
    CLI helper for local execution.

    Example:
    python IncrementalMetadata_Sync.py --source .\\Files\\Test_DataFile_Feb.csv --template .\\Files\\Metadata_Template.csv --cube CSNPLAN
    """
    import argparse

    parser = argparse.ArgumentParser(description="Run EPBCS incremental metadata sync.")
    parser.add_argument("--source", required=True, help="Path to source data CSV file")
    parser.add_argument("--template", required=True, help="Path to Metadata_Template.csv")
    parser.add_argument("--cube", required=True, help="Cube/plan type name")
    parser.add_argument("--mapping", default=None, help="Path to FileColtoDim.csv")
    parser.add_argument("--app", default=None, help="Override EPBCS app name")
    parser.add_argument("--workdir", default=None, help="Working directory for generated files")
    parser.add_argument(
        "--similarity-threshold",
        type=float,
        default=0.72,
        help="Similarity threshold for parent derivation",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Only create ZIP payloads locally, do not upload or run jobs",
    )
    parser.add_argument(
        "--no-prompt",
        action="store_true",
        help="Do not prompt when FileColtoDim.csv is missing",
    )

    args = parser.parse_args()
    try:
        result = pbcs_run_incmetadata_sync(
            source_csv_path=args.source,
            template_csv_path=args.template,
            cube_name=args.cube,
            mapping_csv_path=args.mapping,
            app_name=args.app,
            working_dir=args.workdir,
            similarity_threshold=args.similarity_threshold,
            dry_run=args.dry_run,
            prompt_for_missing_mapping=not args.no_prompt,
        )
        print(json.dumps(result, indent=2))
    except IncrementalMetadataSyncError as exc:
        print(
            json.dumps(
                {
                    "ok": False,
                    "error_code": getattr(exc, "error_code", "INC_METADATA_SYNC_ERROR"),
                    "message": str(exc),
                },
                indent=2,
            )
        )
        raise SystemExit(getattr(exc, "exit_code", 1))
    except Exception as exc:  # pragma: no cover
        print(json.dumps({"ok": False, "error_code": "UNHANDLED_ERROR", "message": str(exc)}, indent=2))
        raise SystemExit(1)


if __name__ == "__main__":
    main()
