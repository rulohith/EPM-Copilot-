from __future__ import annotations

import base64
import csv
import hashlib
import json
import os
import re
import sys
import time
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import quote

import openpyxl
import requests
from dotenv import load_dotenv


ACCOUNT_ZIP_NAME = "Export_Account.zip"
ACCOUNT_CSV_NAME = "ankita.roy@oracle.com_ExportedMetadata_Measures.csv"

LOB_ZIP_NAME = "Export_LOB.zip"
LOB_CSV_NAME = "ankita.roy@oracle.com_ExportedMetadata_LOB.csv"

EXCEL_FILE_NAME = "AccountsToBeAdded.xlsx"
# Excel file to use for LOB updates (separate from account updates)
LOB_EXCEL_FILE_NAME = "LOBsToBeAdded.xlsx"

DEBUG_MERGE = True


def _normalize_base_url(base_url: str) -> str:
    """
    Normalize EPM base URL to the environment root.

    Examples:
    - https://host/epmcloud                    -> https://host
    - https://host/epmcloud/HyperionPlanning   -> https://host
    - https://host/HyperionPlanning            -> https://host
    """
    base = (base_url or "").strip().rstrip("/")
    if not base:
        return base

    for suffix in ("/epmcloud/HyperionPlanning", "/HyperionPlanning", "/epmcloud"):
        if base.endswith(suffix):
            base = base[: -len(suffix)]
    return base.rstrip("/")


# Account export removes the full unwanted set.
ACCOUNT_COLUMNS_TO_REMOVE = [
    "Alias: CY Table",
    "Valid For Consolidations",
    "Hierarchy Type",
    "Enable for Dynamic Children",
    "Number of Possible Dynamic Children",
    "Access Granted to Member Creator",
    "Allow Upper Level Entity Input",
    "Process Management Enabled",
    "UUID",
    "Data Id",
    "Old Name",
    "Old Unique Name",
    "Operation",
]

# LOB export only removes this one column.
LOB_COLUMNS_TO_REMOVE = [
    "Alias: CY Table",
]


def _debug(msg: str) -> None:
    if DEBUG_MERGE:
        print(msg)


# ---------------- AUTH ----------------
def _basic_auth_header(username: str, password: str) -> Dict[str, str]:
    token = base64.b64encode(f"{username}:{password}".encode()).decode()
    return {
        "Authorization": f"Basic {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }


def _load_config() -> Dict[str, str]:
    script_dir = Path(__file__).resolve().parent
    load_dotenv(script_dir / ".env", override=True)

    cfg = {
        "base_url": _normalize_base_url(os.getenv("EPM_BASE_URL", "")),
        "username": os.getenv("EPM_USERNAME", "").strip(),
        "password": os.getenv("EPM_PASSWORD", "").strip(),
        "api_version": os.getenv("EPM_API_VERSION", "v3").strip(),
        "application": os.getenv("EPM_APPLICATION", "").strip(),
        "account_export_job_name": os.getenv("EPM_JOB_NAME", "Export Account").strip(),
        "account_import_job_name": os.getenv("EPM_IMPORT_JOB_NAME", "Import Account").strip(),
        "lob_export_job_name": os.getenv("EPM_LOB_EXPORT_JOB_NAME", "Export LOB").strip(),
        "lob_import_job_name": os.getenv("EPM_LOB_IMPORT_JOB_NAME", "Import LOB").strip(),
    }

    missing = [k for k, v in cfg.items() if not v]
    if missing:
        raise ValueError(f"Missing required environment variables: {', '.join(missing)}")

    return cfg


# ---------------- REQ / JOB HELPERS ----------------
def _request_json(
    session: requests.Session,
    method: str,
    url: str,
    headers: Dict[str, str],
    *,
    params=None,
    payload=None,
    timeout: int = 60,
) -> Dict[str, Any]:
    resp = session.request(
        method=method,
        url=url,
        headers=headers,
        params=params,
        json=payload,
        timeout=timeout,
    )
    resp.raise_for_status()
    if not resp.text.strip():
        return {}
    return resp.json()


def _find_self_link(payload: Dict[str, Any]) -> Optional[str]:
    links = payload.get("links")
    if not isinstance(links, list):
        return None

    for link in links:
        if isinstance(link, dict) and link.get("rel") == "self" and link.get("href"):
            return str(link["href"])
    return None


def _extract_job_id(payload: Dict[str, Any]) -> Optional[str]:
    for key in ("jobId", "jobID", "id"):
        value = payload.get(key)
        if value is not None and str(value).strip():
            return str(value)
    return None


def _wait_for_job_completion(
    session: requests.Session,
    status_url: str,
    headers: Dict[str, str],
    timeout_seconds: int = 1800,
) -> Dict[str, Any]:
    deadline = time.time() + timeout_seconds

    while time.time() < deadline:
        data = _request_json(session, "GET", status_url, headers, timeout=60)
        status = data.get("status")

        if status in (-1, "-1", None):
            time.sleep(5)
            continue

        if status in (0, "0"):
            return data

        raise RuntimeError(
            "Job failed.\n"
            f"Status: {data.get('status')}\n"
            f"Job: {data.get('jobName')}\n"
            f"Details: {data.get('details') or data.get('descriptiveStatus') or data}"
        )

    raise TimeoutError("Timed out waiting for job completion.")


def _validate_job_definition(
    session: requests.Session,
    cfg: Dict[str, str],
    headers: Dict[str, str],
    job_type: str,
    job_name: str,
) -> None:
    jobdefs_url = (
        f"{cfg['base_url']}/HyperionPlanning/rest/{cfg['api_version']}/applications/"
        f"{cfg['application']}/jobdefinitions"
    )
    query = f'{{"jobType":"{job_type}"}}'
    data = _request_json(session, "GET", jobdefs_url, headers, params={"q": query}, timeout=60)

    items = data.get("items", [])
    if not isinstance(items, list):
        return

    exact = [item for item in items if isinstance(item, dict) and item.get("jobName") == job_name]
    if exact:
        return

    available = [item.get("jobName") for item in items if isinstance(item, dict) and item.get("jobName")]
    available_text = ", ".join(map(str, available[:20])) if available else "(none returned)"
    raise RuntimeError(
        f"Planning does not expose a {job_type} job named '{job_name}'.\n"
        f"Available {job_type} jobs include: {available_text}"
    )


# ---------------- REPOSITORY HELPERS ----------------
def _download_repository_file(
    session: requests.Session,
    cfg: Dict[str, str],
    headers: Dict[str, str],
    remote_file_name: str,
    local_dir: Path,
) -> Path:
    encoded_name = quote(remote_file_name, safe="")
    download_url = (
        f"{cfg['base_url']}/interop/rest/11.1.2.3.600/applicationsnapshots/"
        f"{encoded_name}/contents"
    )

    resp = session.get(
        download_url,
        headers={"Authorization": headers["Authorization"]},
        stream=True,
        timeout=120,
    )

    content_type = resp.headers.get("Content-Type", "").lower()
    if "application/json" in content_type:
        try:
            err = resp.json()
        except Exception:
            err = {"details": resp.text}
        raise RuntimeError(f"Failed to download repository file: {err}")

    resp.raise_for_status()

    local_dir.mkdir(parents=True, exist_ok=True)
    local_path = local_dir / remote_file_name
    with open(local_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=1024 * 1024):
            if chunk:
                f.write(chunk)

    return local_path


def _delete_repository_file(
    session: requests.Session,
    cfg: Dict[str, str],
    headers: Dict[str, str],
    file_name: str,
) -> None:
    delete_url = f"{cfg['base_url']}/interop/rest/v3/files/delete"
    payload = {"fileName": file_name}

    resp = session.post(
        delete_url,
        json=payload,
        headers={
            "Authorization": headers["Authorization"],
            "Accept": "application/json",
            "Content-Type": "application/json",
        },
        timeout=120,
    )

    if resp.status_code == 404:
        return

    resp.raise_for_status()


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _wait_until_repository_file_matches(
    session: requests.Session,
    cfg: Dict[str, str],
    headers: Dict[str, str],
    local_file: Path,
    remote_name: str,
    timeout_seconds: int = 300,
    poll_seconds: int = 5,
) -> None:
    expected_hash = _sha256_file(local_file)
    deadline = time.time() + timeout_seconds
    last_err: Optional[Exception] = None

    while time.time() < deadline:
        try:
            tmp_download = _download_repository_file(
                session=session,
                cfg=cfg,
                headers=headers,
                remote_file_name=remote_name,
                local_dir=Path(local_file.parent),
            )
            actual_hash = _sha256_file(tmp_download)
            if actual_hash == expected_hash:
                return
        except Exception as exc:
            last_err = exc
        time.sleep(poll_seconds)

    if last_err:
        raise TimeoutError(
            f"Repository file '{remote_name}' did not match the uploaded ZIP in time. Last error: {last_err}"
        )
    raise TimeoutError(f"Repository file '{remote_name}' did not match the uploaded ZIP in time.")


def _upload_repository_file(
    session: requests.Session,
    cfg: Dict[str, str],
    headers: Dict[str, str],
    local_file: Path,
    remote_name: str,
) -> None:
    _delete_repository_file(session, cfg, headers, remote_name)

    encoded_name = quote(remote_name, safe="")
    upload_url = (
        f"{cfg['base_url']}/interop/rest/11.1.2.3.600/applicationsnapshots/"
        f"{encoded_name}/contents"
    )

    with open(local_file, "rb") as f:
        resp = session.post(
            upload_url,
            headers={
                "Authorization": headers["Authorization"],
                "Accept": "application/json",
                "Content-Type": "application/octet-stream",
            },
            data=f,
            timeout=300,
        )

    resp.raise_for_status()

    if resp.text.strip():
        try:
            data = resp.json()
            self_link = _find_self_link(data)
            if self_link:
                _wait_for_job_completion(session, self_link, headers, timeout_seconds=1800)
        except Exception:
            pass

    _wait_until_repository_file_matches(session, cfg, headers, local_file, remote_name)


# ---------------- UTILS ----------------
def _clean(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _norm_key(value: Any) -> str:
    text = _clean(value)
    text = text.replace("\u00A0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.casefold()


def _alphanum_key(value: str):
    parts = re.split(r"(\d+)", _clean(value))
    return [int(p) if p.isdigit() else p.lower() for p in parts if p != ""]


def _next_member_value(last_value: str) -> str:
    last_value = _clean(last_value)
    if not last_value:
        return "1"

    match = re.match(r"^(.*?)(\d+)$", last_value)
    if match:
        prefix, digits = match.groups()
        return f"{prefix}{str(int(digits) + 1).zfill(len(digits))}"

    return f"{last_value}1"


# ---------------- MERGE LOGIC ----------------
def _merge_accounts_from_excel(
    csv_path: Path,
    excel_path: Path,
    columns_to_remove: list[str],
    member_prefix: Optional[str] = None,
    comparison_field_index: int = 2,
) -> bool:
    cols_to_remove_set = {c.strip() for c in columns_to_remove}
    csv_changed = False

    _debug(f"[MERGE] Reading CSV: {csv_path.name}")

    with open(csv_path, "r", newline="", encoding="utf-8-sig") as src:
        reader = csv.reader(src, skipinitialspace=True)
        header = next(reader)
        header = [_clean(col) for col in header]

        keep_indices = [i for i, col in enumerate(header) if col not in cols_to_remove_set]
        filtered_header = [header[i] for i in keep_indices]

        rows: list[list[str]] = []
        for row in reader:
            cleaned_row = []
            for i in keep_indices:
                cleaned_row.append(_clean(row[i]) if i < len(row) else "")
            rows.append(cleaned_row)

    if not excel_path.is_file():
        print(f"Excel file not found: {excel_path}")
        with open(csv_path, "w", newline="", encoding="utf-8") as dst:
            writer = csv.writer(dst)
            writer.writerow(filtered_header)
            writer.writerows(rows)
        return False

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    source_label = "LOB" if member_prefix else "ACCOUNT"
    _debug(f"[MERGE] Source workbook: {excel_path.name} | Target mode: {source_label}")
    if member_prefix:
        _debug(f"[MERGE] LOB member prefix will be applied as needed: {member_prefix}")

    def normalize_member_for_target(member_value: str) -> str:
        member_value = _clean(member_value)
        if member_prefix and member_value:
            if not member_value.upper().startswith(member_prefix.upper()):
                return f"{member_prefix}{member_value}"
        return member_value

    def rebuild_indexes():
        name_to_index: dict[str, int] = {}
        parent_to_rows: dict[str, list[list[str]]] = defaultdict(list)
        parent_member_to_index: dict[str, dict[str, int]] = defaultdict(dict)
        parent_member_field_to_index: dict[str, dict[tuple[str, str], int]] = defaultdict(dict)

        for i, row in enumerate(rows):
            if len(row) >= 1:
                name_to_index[_norm_key(row[0])] = i
            if len(row) >= 2:
                parent_key = _norm_key(row[1])
                parent_to_rows[parent_key].append(row)

                member_value = _clean(row[0]) if len(row) > 0 else ""
                member_key = _norm_key(member_value) if member_value else ""
                third_value = _clean(row[2]) if len(row) > 2 else ""
                third_key = _norm_key(third_value) if third_value else ""

                if member_key:
                    parent_member_to_index[parent_key][member_key] = i
                    if third_key:
                        parent_member_field_to_index[parent_key][(member_key, third_key)] = i

        return name_to_index, parent_to_rows, parent_member_to_index, parent_member_field_to_index

    name_to_index, parent_to_rows, parent_member_to_index, parent_member_field_to_index = rebuild_indexes()

    affected_parents: set[str] = set()
    inserted_count = 0
    updated_count = 0
    skipped_count = 0

    for idx, excel_row in enumerate(ws.iter_rows(values_only=True), start=1):
        if idx == 1:
            continue

        member_input = _clean(excel_row[0] if len(excel_row) > 0 else "")
        parent = _clean(excel_row[1] if len(excel_row) > 1 else "")
        third_field = _clean(
            excel_row[comparison_field_index] if len(excel_row) > comparison_field_index else ""
        )

        if not member_input and not parent and not third_field:
            skipped_count += 1
            _debug(f"[SKIP] Excel row {idx + 1}: empty row")
            continue

        parent_key = _norm_key(parent) if parent else ""
        if not parent_key or parent_key not in parent_to_rows or not parent_to_rows[parent_key]:
            skipped_count += 1
            _debug(f"[SKIP] Excel row {idx + 1}: parent not found in CSV -> parent={parent!r}")
            continue

        member = normalize_member_for_target(member_input)
        if not member:
            last_member = _clean(parent_to_rows[parent_key][-1][0]) if parent_to_rows[parent_key] else ""
            generated_member = _next_member_value(last_member)
            member = normalize_member_for_target(generated_member)
            _debug(
                f"[AUTO] Excel row {idx + 1}: blank member, generated member={member!r} from last_member={last_member!r}"
            )
        elif member != member_input:
            _debug(
                f"[NORM] Excel row {idx + 1}: member normalized from {member_input!r} to {member!r}"
            )

        member_key = _norm_key(member)
        third_key = _norm_key(third_field) if third_field else ""

        exact_index = None
        if third_key:
            exact_index = parent_member_field_to_index.get(parent_key, {}).get((member_key, third_key))

        if exact_index is not None:
            skipped_count += 1
            _debug(
                f"[SKIP] Excel row {idx + 1}: exact match already exists -> member={member!r}, parent={parent!r}, third_field={third_field!r}"
            )
            continue

        existing_index = parent_member_to_index.get(parent_key, {}).get(member_key)

        if existing_index is not None:
            existing_row = rows[existing_index]
            old_parent = _clean(existing_row[1]) if len(existing_row) > 1 else ""
            old_third_field = _clean(existing_row[2]) if len(existing_row) > 2 else ""

            if len(existing_row) < len(filtered_header):
                existing_row.extend([""] * (len(filtered_header) - len(existing_row)))

            if (
                _clean(existing_row[0]) != member
                or _clean(existing_row[1]) != parent
                or old_third_field != third_field
            ):
                existing_row[0] = member
                if len(existing_row) > 1:
                    existing_row[1] = parent
                if len(existing_row) > 2:
                    existing_row[2] = third_field

                if parent_key and parent_key in parent_to_rows and parent_to_rows[parent_key]:
                    sibling = parent_to_rows[parent_key][0]
                    for col_index in range(3, len(filtered_header)):
                        value = _clean(sibling[col_index]) if col_index < len(sibling) else ""
                        if col_index < len(existing_row):
                            existing_row[col_index] = value
                        else:
                            existing_row.append(value)

                csv_changed = True
                updated_count += 1
                affected_parents.add(_norm_key(old_parent))
                affected_parents.add(parent_key)
                _debug(
                    f"[UPDATE] Excel row {idx + 1}: updated existing row -> member={member!r}, parent={parent!r}, third_field={third_field!r}"
                )
                name_to_index, parent_to_rows, parent_member_to_index, parent_member_field_to_index = rebuild_indexes()
            else:
                skipped_count += 1
                _debug(
                    f"[SKIP] Excel row {idx + 1}: same member and parent already present with same comparison field -> member={member!r}, parent={parent!r}"
                )
            continue

        sibling = parent_to_rows[parent_key][0]
        new_row = ["" for _ in range(len(filtered_header))]

        if len(new_row) > 0:
            new_row[0] = member
        if len(new_row) > 1:
            new_row[1] = parent
        if len(new_row) > 2:
            new_row[2] = third_field

        for col_index in range(3, len(filtered_header)):
            new_row[col_index] = _clean(sibling[col_index]) if col_index < len(sibling) else ""

        rows.append(new_row)
        csv_changed = True
        inserted_count += 1
        affected_parents.add(parent_key)
        _debug(
            f"[INSERT] Excel row {idx + 1}: added new row -> member={member!r}, parent={parent!r}, third_field={third_field!r}"
        )

        name_to_index, parent_to_rows, parent_member_to_index, parent_member_field_to_index = rebuild_indexes()

    for parent_key in affected_parents:
        parent_indices = [
            i for i, row in enumerate(rows)
            if len(row) >= 2 and _norm_key(row[1]) == parent_key
        ]
        if len(parent_indices) <= 1:
            continue

        sorted_group = sorted(
            [rows[i] for i in parent_indices],
            key=lambda r: _alphanum_key(r[0] if len(r) > 0 else ""),
        )

        for i, sorted_row in zip(parent_indices, sorted_group):
            rows[i] = sorted_row

    with open(csv_path, "w", newline="", encoding="utf-8") as dst:
        writer = csv.writer(dst)
        writer.writerow(filtered_header)
        writer.writerows(rows)

    print(f"Inserted {inserted_count} rows from {excel_path.name}")
    print(f"Updated {updated_count} existing rows from {excel_path.name}")
    print(f"Skipped {skipped_count} Excel rows")
    print(f"CSV changed: {csv_changed}")

    wb.close()
    return csv_changed


# ---------------- JOB WORKFLOW ----------------
def _submit_export_job(
    cfg: Dict[str, str],
    export_job_name: str,
    export_zip_name: str,
    exported_csv_name: str,
    columns_to_remove: list[str],
    member_prefix: Optional[str] = None,
    comparison_field_index: int = 2,
    excel_file_name: str = EXCEL_FILE_NAME,
) -> Dict[str, Any]:
    headers = _basic_auth_header(cfg["username"], cfg["password"])
    session = requests.Session()

    _validate_job_definition(session, cfg, headers, job_type="EXPORT_METADATA", job_name=export_job_name)

    job_url = (
        f"{cfg['base_url']}/HyperionPlanning/rest/{cfg['api_version']}/applications/"
        f"{cfg['application']}/jobs"
    )

    payload = {
        "jobType": "EXPORT_METADATA",
        "jobName": export_job_name,
        "parameters": {"exportZipFileName": export_zip_name},
    }

    print(f"Submitting jobType=EXPORT_METADATA jobName={export_job_name!r}")
    print(f"Payload: {json.dumps(payload)}")

    resp = session.post(job_url, json=payload, headers=headers, timeout=120)
    resp.raise_for_status()
    data = resp.json()

    self_link = _find_self_link(data)
    if self_link:
        result = _wait_for_job_completion(session, self_link, headers)
    else:
        job_id = _extract_job_id(data)
        if job_id:
            status_url = (
                f"{cfg['base_url']}/HyperionPlanning/rest/{cfg['api_version']}/applications/"
                f"{cfg['application']}/jobs/{job_id}"
            )
            result = _wait_for_job_completion(session, status_url, headers)
        else:
            result = data

    download_dir = Path(__file__).resolve().parent
    local_zip = _download_repository_file(session, cfg, headers, export_zip_name, download_dir)
    print(f"Downloaded ZIP to: {local_zip}")

    try:
        # Extract the exported CSV from the ZIP
        with zipfile.ZipFile(local_zip, "r") as zf:
            members = [name for name in zf.namelist() if name.lower().endswith(".csv") and not name.endswith("/")]

            if not members:
                raise RuntimeError("No CSV file found inside the exported ZIP.")

            target_member = exported_csv_name if exported_csv_name in members else members[0]
            extracted_csv_path = Path(zf.extract(target_member, path=download_dir))

        # Ensure the extracted CSV has the expected filename
        if extracted_csv_path.name != exported_csv_name:
            renamed_csv_path = download_dir / exported_csv_name
            if renamed_csv_path.exists():
                try:
                    renamed_csv_path.unlink()
                except Exception as rm_err:
                    print(f"Warning: could not remove existing CSV before rename: {rm_err}", file=sys.stderr)
            extracted_csv_path.replace(renamed_csv_path)
            extracted_csv_path = renamed_csv_path

        # Determine which Excel workbook to use for the merge
        excel_path = download_dir / excel_file_name
        _debug(
            f"[MERGE] Processing exported CSV {exported_csv_name!r} using workbook {excel_path.name!r}"
        )
        csv_changed = _merge_accounts_from_excel(
            extracted_csv_path,
            excel_path,
            columns_to_remove,
            member_prefix=member_prefix,
            comparison_field_index=comparison_field_index,
        )

        # Re‑package the (potentially) updated CSV back into the ZIP
        with zipfile.ZipFile(local_zip, "w", zipfile.ZIP_DEFLATED) as zf_out:
            zf_out.write(extracted_csv_path, arcname=exported_csv_name)

        print(f"Updated CSV and repackaged ZIP: {local_zip}")

        if csv_changed:
            _upload_repository_file(session, cfg, headers, local_zip, export_zip_name)
            print(f"Uploaded updated ZIP back to repository as: {export_zip_name}")
        else:
            print("No CSV changes detected. Skipping upload. Import job will still run.")

    except Exception as e:
        print(f"Warning: failed to process export CSV: {e}", file=sys.stderr)

    return result


def _submit_import_job(
    cfg: Dict[str, str],
    import_job_name: str,
    import_zip_name: str,
) -> Dict[str, Any]:
    headers = _basic_auth_header(cfg["username"], cfg["password"])
    session = requests.Session()

    _validate_job_definition(session, cfg, headers, job_type="IMPORT_METADATA", job_name=import_job_name)

    job_url = (
        f"{cfg['base_url']}/HyperionPlanning/rest/{cfg['api_version']}/applications/"
        f"{cfg['application']}/jobs"
    )

    payload = {
        "jobType": "IMPORT_METADATA",
        "jobName": import_job_name,
        "parameters": {
            "importZipFileName": import_zip_name,
            "errorFile": "Error_Log.zip",
            "refreshDatabase": "true",
        },
    }

    print(f"Submitting jobType=IMPORT_METADATA jobName={import_job_name!r}")
    print(f"Payload: {json.dumps(payload)}")

    resp = session.post(job_url, json=payload, headers=headers, timeout=120)
    resp.raise_for_status()
    data = resp.json()

    self_link = _find_self_link(data)
    if self_link:
        result = _wait_for_job_completion(session, self_link, headers)
    else:
        job_id = _extract_job_id(data)
        if job_id:
            status_url = (
                f"{cfg['base_url']}/HyperionPlanning/rest/{cfg['api_version']}/applications/"
                f"{cfg['application']}/jobs/{job_id}"
            )
            result = _wait_for_job_completion(session, status_url, headers)
        else:
            result = data

    print("Import metadata job completed.")
    return result


def main() -> int:
    try:
        cfg = _load_config()

        export_account_result = _submit_export_job(
            cfg=cfg,
            export_job_name=cfg["account_export_job_name"],
            export_zip_name=ACCOUNT_ZIP_NAME,
            exported_csv_name=ACCOUNT_CSV_NAME,
            columns_to_remove=ACCOUNT_COLUMNS_TO_REMOVE,
            member_prefix=None,
            comparison_field_index=2,  # Excel 3rd column
        )
        print("Export Account job completed successfully.")
        print(json.dumps(export_account_result, indent=2))

        import_account_result = _submit_import_job(
            cfg=cfg,
            import_job_name=cfg["account_import_job_name"],
            import_zip_name=ACCOUNT_ZIP_NAME,
        )
        print("Import Account job completed successfully.")
        print(json.dumps(import_account_result, indent=2))

        # For LOB merges we want to match on the "Alias: Default" column in the exported CSV.
        # This column is the third column (index 2) in the CSV file. The Excel source only provides
        # Name, Parent, and Description, so we use the Description field as the value to compare
        # against the "Alias: Default" column. Therefore we set ``comparison_field_index`` to ``2``
        # (zero‑based) which corresponds to the third column in the CSV.
        export_lob_result = _submit_export_job(
            cfg=cfg,
            export_job_name=cfg["lob_export_job_name"],
            export_zip_name=LOB_ZIP_NAME,
            exported_csv_name=LOB_CSV_NAME,
            columns_to_remove=LOB_COLUMNS_TO_REMOVE,
            member_prefix="LOB_",
            comparison_field_index=2,  # Match on "Alias: Default" column
            excel_file_name=LOB_EXCEL_FILE_NAME,
        )
        print("Export LOB job completed successfully.")
        print(json.dumps(export_lob_result, indent=2))

        import_lob_result = _submit_import_job(
            cfg=cfg,
            import_job_name=cfg["lob_import_job_name"],
            import_zip_name=LOB_ZIP_NAME,
        )
        print("Import LOB job completed successfully.")
        print(json.dumps(import_lob_result, indent=2))

        return 0
    except Exception as exc:
        print(f"Error during export/import: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
