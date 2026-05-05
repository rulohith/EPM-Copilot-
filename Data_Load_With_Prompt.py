import argparse
import os
import base64
import json
import time
from pathlib import Path
from typing import Dict, Any, Optional, Tuple, List
from urllib.parse import quote

import requests
from dotenv import load_dotenv


SCRIPT_DIR = Path(__file__).resolve().parent


class FileSelectionError(RuntimeError):
    """Raised when an entity-specific file cannot be uniquely identified."""


def _normalize_base_url(base_url: str) -> str:
    """Normalize EPM base URL to the host root used by Interop/Planning REST APIs."""
    base = (base_url or "").strip().rstrip("/")
    if not base:
        return base

    for suffix in ("/epmcloud/HyperionPlanning", "/HyperionPlanning", "/epmcloud"):
        if base.endswith(suffix):
            base = base[: -len(suffix)]
    return base.rstrip("/")


def _load_config() -> Dict[str, str]:
    load_dotenv(SCRIPT_DIR / ".env", override=True)

    cfg = {
        "base_url": _normalize_base_url(os.getenv("EPM_BASE_URL", "")),
        "username": os.getenv("EPM_USERNAME", "").strip(),
        "password": os.getenv("EPM_PASSWORD", "").strip(),
        "api_version": os.getenv("EPM_API_VERSION", "v3").strip(),
        "application": os.getenv("EPM_APPLICATION", "").strip(),
        "local_folder": os.getenv("LOCAL_FOLDER", "").strip() or os.getcwd(),
        # Interop upload API version. Default aligns with the endpoint used by
        # metadata import/export flows.
        "upload_version": (
            os.getenv("EPM_UPLOAD_VERSION")
            or os.getenv("EPM_INTEROP_VERSION")
            or "11.1.2.3.600"
        ).strip(),
        # Optional: falls back to entity lookup when unset so we shouldn't
        # force every environment to define it.
        "filename": os.getenv("FILENAME", "").strip(),
        "groovy_rule_actuals": os.getenv("GROOVY_RULE_ACTUALS", "LoadData").strip() or "LoadData",
        "groovy_rule_forecast": os.getenv("GROOVY_RULE_FORECAST", "LoadData").strip() or "LoadData",
        "rtp_filename": (
            os.getenv("GROOVY_RTP_FILENAME")
            or os.getenv("GROOVY_RTP_NAME")
            or "RTP_FileName"
        ).strip(),
        "rtp_load_option": os.getenv("GROOVY_RTP_LOADOPTION", "DataLoadOption").strip()
        or "DataLoadOption",
        "job_poll_interval": int(os.getenv("JOB_POLL_INTERVAL", "5")),
        "job_timeout": int(os.getenv("JOB_TIMEOUT_SECONDS", "300")),
    }

    # "filename" is optional because the entity/file arguments can drive the
    # selection. Everything else is required for the REST calls.
    missing = [k for k, v in cfg.items() if not v and k != "filename"]
    if missing:
        raise ValueError(f"Missing required environment variables: {', '.join(missing)}")

    return cfg


def _basic_auth_header(username: str, password: str) -> Dict[str, str]:
    token = base64.b64encode(f"{username}:{password}".encode()).decode()
    return {
        "Authorization": f"Basic {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }


def _delete_inbox_file(
    session: requests.Session,
    cfg: Dict[str, Any],
    headers: Dict[str, str],
    remote_name: str,
) -> None:
    delete_url = f"{cfg['base_url']}/interop/rest/v3/files/delete"
    payload = {"fileName": remote_name}

    resp = session.post(
        delete_url,
        json=payload,
        headers={
            "Authorization": headers["Authorization"],
            "Accept": "application/json",
            "Content-Type": "application/json",
        },
        timeout=60,
    )

    if resp.status_code not in (200, 404):
        resp.raise_for_status()


def _upload_to_inbox(
    session: requests.Session,
    cfg: Dict[str, Any],
    headers: Dict[str, str],
    local_file_path: Path,
    remote_name: str,
) -> None:
    _delete_inbox_file(session, cfg, headers, remote_name)

    encoded_name = quote(remote_name, safe="")
    upload_url = (
        f"{cfg['base_url']}/interop/rest/{cfg['upload_version']}"
        f"/applicationsnapshots/{encoded_name}/contents"
    )

    with open(local_file_path, "rb") as file_body:
        raw_bytes = file_body.read()

    req_headers = {
        "Authorization": headers["Authorization"],
        "Content-Type": "application/octet-stream",
        "Content-Length": str(len(raw_bytes)),
    }

    params = {"isForce": "true"}

    resp = session.post(
        upload_url,
        headers=req_headers,
        params=params,
        data=raw_bytes,
        timeout=120,
    )

    if resp.status_code == 200:
        print(f"Successfully uploaded {remote_name} to inbox (force overwrite enabled)")
    else:
        print("Upload failed:", resp.status_code, resp.text)
        resp.raise_for_status()


def _run_groovy_script(
    session: requests.Session,
    cfg: Dict[str, Any],
    headers: Dict[str, str],
    file_name: str,
    load_option: str,
    dataset: str,
    skip_poll: bool = False,
) -> Dict[str, Any]:
    """Invoke the Groovy rule via the Planning jobs REST API."""

    jobs_url = (
        f"{cfg['base_url']}/HyperionPlanning/rest/{cfg['api_version']}"
        f"/applications/{cfg['application']}/jobs"
    )

    rule_name = (
        cfg["groovy_rule_forecast"] if dataset == "forecast" else cfg["groovy_rule_actuals"]
    )

    payload = {
        "jobType": "Rules",
        "jobName": rule_name,
        "parameters": {
            cfg["rtp_filename"]: file_name,
            cfg["rtp_load_option"]: load_option,
        },
    }

    resp = session.post(jobs_url, headers=headers, data=json.dumps(payload), timeout=60)
    if not resp.ok:
        print("Groovy job submission failed:", resp.status_code, resp.text)
        resp.raise_for_status()

    data = resp.json()

    # Some Planning responses return "details" as a string, so normalize to dict.
    details = data.get("details")
    if isinstance(details, str):
        try:
            details = json.loads(details)
        except json.JSONDecodeError:
            details = {"message": details}
    elif details is None:
        details = {}
    elif not isinstance(details, dict):
        details = {"details": details}

    job_id = details.get("jobId") or data.get("jobId")
    if not job_id:
        raise RuntimeError(
            "Groovy job response missing jobId. Payload: " + json.dumps(data, indent=2)
        )

    print(f"Triggered Groovy rule '{rule_name}' as job {job_id} for {file_name}.")

    if skip_poll:
        print("Skipping status polling per --no-poll flag. Monitor the job in Planning if needed.")
        return {"jobId": job_id, "status": "Submitted", "polled": False}

    result = _poll_job_until_complete(session, cfg, headers, job_id)

    return result


def _poll_job_until_complete(
    session: requests.Session,
    cfg: Dict[str, Any],
    headers: Dict[str, str],
    job_id: int,
) -> Dict[str, Any]:
    status_url = (
        f"{cfg['base_url']}/HyperionPlanning/rest/{cfg['api_version']}"
        f"/applications/{cfg['application']}/jobs/{job_id}"
    )

    deadline = time.time() + cfg["job_timeout"]
    while True:
        resp = session.get(status_url, headers=headers, timeout=30)
        if not resp.ok:
            print("Failed to poll job status:", resp.status_code, resp.text)
            resp.raise_for_status()

        payload = resp.json()
        status = payload.get("status")
        details = payload.get("details", "")
        print(f"Job {job_id} status: {status} {details}")

        normalized_status = _normalize_job_status(status)

        if normalized_status == "success":
            return payload
        if normalized_status == "failure":
            raise RuntimeError(f"Groovy job {job_id} failed: {details}")

        if time.time() > deadline:
            raise TimeoutError(f"Groovy job {job_id} did not complete within timeout")

        time.sleep(cfg["job_poll_interval"])




def _normalize_job_status(status: Any) -> str:
    """Map Planning job status responses to canonical success/failure labels.

    Planning REST APIs sometimes return descriptive strings ("Completed", "Error")
    and in other contexts numeric codes (0/1). This helper keeps the polling loop
    consistent so the caller can opt-out of polling entirely if desired.
    """

    if isinstance(status, str):
        normalized = status.strip().lower()
        if normalized in {"completed", "succeeded", "success"}:
            return "success"
        if normalized in {"error", "failed", "failure"}:
            return "failure"
        return "pending"

    # Planning job service uses 0 for success, 1 for failure while running jobs
    # such as rules and data loads.
    if isinstance(status, (int, float)):
        if status == 0:
            return "success"
        if status == 1:
            return "failure"

    return "pending"


def _find_entity_file(entity: str, dataset: str, search_root: Path) -> Tuple[Path, str]:
    """Locate the CSV for the requested entity and dataset (Actuals/Forecast)."""

    suffix = "Forecast" if dataset == "forecast" else "Actuals"
    expected_name = f"{entity} {suffix}.csv"
    matches: List[Path] = [
        path
        for path in search_root.rglob("*.csv")
        if path.name.lower() == expected_name.lower()
    ]

    if not matches:
        raise FileSelectionError("File not found in local system")
    if len(matches) > 1:
        raise FileSelectionError("Duplicate file with same name found in local system")

    return matches[0], matches[0].name


def main():
    parser = argparse.ArgumentParser(description="Upload Actuals/Forecast CSV and trigger Groovy load")
    parser.add_argument(
        "--entity",
        help="Entity identifier (expects '<ENTITY> Actuals.csv' or '<ENTITY> Forecast.csv' based on --dataset)",
    )
    parser.add_argument(
        "--file",
        help="Override full path to a CSV instead of resolving via entity",
    )
    parser.add_argument(
        "--dataset",
        choices=["actuals", "forecast"],
        default="actuals",
        help="Which dataset to load (drives file naming and Groovy rule).",
    )
    parser.add_argument(
        "--load-option",
        choices=["Add", "Overwrite"],
        default="Overwrite",
        help="How to load data: Add (merge) or Overwrite",
    )
    parser.add_argument(
        "--no-poll",
        action="store_true",
        help="Submit the Groovy job and exit without waiting for completion",
    )
    args = parser.parse_args()

    cfg = _load_config()
    headers = _basic_auth_header(cfg["username"], cfg["password"])


    with requests.Session() as session:
        session.headers.update({"Accept": "application/json", "Content-Type": "application/json"})

        local_folder = Path(cfg["local_folder"]).resolve()
        remote_name: str

        if args.file:
            test_file = Path(args.file).expanduser().resolve()
            if not test_file.exists():
                raise FileNotFoundError(f"Provided file not found: {test_file}")
            remote_name = test_file.name
        elif args.entity:
            test_file, remote_name = _find_entity_file(args.entity.strip(), args.dataset, local_folder)
        else:
            if not cfg["filename"]:
                raise FileSelectionError(
                    "No entity, file argument, or FILENAME env provided."
                )

            test_file = (local_folder / cfg["filename"]).resolve()
            if not test_file.exists():
                raise FileNotFoundError(f"{cfg['filename']} not found in {local_folder}")
            remote_name = test_file.name

        _upload_to_inbox(session, cfg, headers, test_file, remote_name)
        result = _run_groovy_script(
            session,
            cfg,
            headers,
            remote_name,
            load_option=args.load_option,
            dataset=args.dataset,
            skip_poll=args.no_poll,
        )
        print("Groovy job completed:", json.dumps(result, indent=2))


if __name__ == "__main__":
    main()