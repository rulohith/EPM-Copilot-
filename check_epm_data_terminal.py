from __future__ import annotations

from itertools import islice
import base64
import time
from pathlib import Path

from connections import connect_to_epm_environment
from generate_dimension_reports import (
    _base_url_candidates,
    discover_dimensions_with_links,
    export_data,
    fetch_dimension_members,
)


def main() -> None:
    print("=== EPM Metadata + Data Check (Terminal) ===")
    session, cfg = connect_to_epm_environment()
    cfg["plantype"] = cfg.get("plantype") or cfg["application"]

    print(f"Application: {cfg['application']}")
    print(f"Base URL: {cfg['base_url']}")
    print("\n[1] Discovering dimensions...")
    dimensions, dim_links = discover_dimensions_with_links(session, cfg)

    if not dimensions:
        print("No dimensions discovered.")
        return

    print(f"Discovered {len(dimensions)} dimensions:")
    for d in dimensions:
        print(f" - {d}")

    print("\n[2] Fetching sample members per dimension (up to 20 each)...")
    for d in dimensions:
        members = fetch_dimension_members(session, cfg, d, dim_links.get(d, ""))
        sample = list(islice(members, 20))
        print(f"\nDimension: {d}")
        print(f"Total members fetched: {len(members)}")
        print("Sample members:")
        if sample:
            for m in sample:
                print(f"   - {m}")
        else:
            print("   (no members returned)")

    print("\n[3] Attempting data export...")
    items = export_data(session, cfg, dimensions)
    print(f"Data rows returned: {len(items)}")

    if items:
        print("\nSample data rows (first 5):")
        for i, row in enumerate(items[:5], start=1):
            print(f"Row {i}: {row}")
    else:
        print("No data rows returned by dataexport endpoint for current user/environment.")

    print("\n[4] Attempting fallback using EPM Job: Export Account ...")
    token = base64.b64encode(f"{cfg['username']}:{cfg['password']}".encode()).decode()
    headers = {
        "Authorization": f"Basic {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }

    job_payload = {"jobType": "EXPORT_METADATA", "jobName": "Export Account"}
    job_id = None
    job_url_used = ""

    for root in _base_url_candidates(cfg["base_url"]):
        candidates = [
            f"{root}/HyperionPlanning/rest/{cfg['api_version']}/applications/{cfg['application']}/jobs",
            f"{root}/rest/{cfg['api_version']}/applications/{cfg['application']}/jobs",
        ]
        for u in candidates:
            try:
                r = session.post(u, headers=headers, json=job_payload, timeout=60, verify=False)
                print(f"POST {u} -> {r.status_code}")
                if r.status_code in (200, 201, 202):
                    data = r.json()
                    job_id = data.get("jobId") or data.get("jobID") or data.get("id")
                    job_url_used = u
                    break
                else:
                    print((r.text or "")[:200])
            except Exception as ex:
                print(f"POST {u} error: {ex}")
        if job_id:
            break

    if not job_id:
        print("Could not start Export Account job via REST endpoints.")
        return

    print(f"Started job id: {job_id}")

    # Poll status
    job_status_url = f"{job_url_used}/{job_id}"
    final_job_payload = {}
    for _ in range(15):
        try:
            rs = session.get(job_status_url, headers=headers, timeout=60, verify=False)
            if rs.status_code == 200:
                js = rs.json()
                final_job_payload = js
                status = str(js.get("status") or js.get("jobStatus") or "").lower()
                print(f"Job status: {status}")
                if status in ("success", "completed", "4"):
                    break
                if status in ("error", "failed", "3"):
                    print("Job failed")
                    print(js)
                    return
            else:
                print(f"Status check HTTP {rs.status_code}")
        except Exception as ex:
            print(f"Status check error: {ex}")
        time.sleep(2)

    if final_job_payload:
        print("Job payload:")
        print(final_job_payload)

    # Try download known metadata zip from Outbox
    out_file = "Export_Account.zip"
    out_path = Path("Export_Account.zip")
    downloaded = False
    # First try direct links from job payload
    links = final_job_payload.get("links", []) if isinstance(final_job_payload, dict) else []
    for lnk in links if isinstance(links, list) else []:
        if not isinstance(lnk, dict):
            continue
        href = str(lnk.get("href") or "")
        if not href:
            continue
        try:
            rd = session.get(href, headers=headers, timeout=60, verify=False)
            print(f"GET {href} -> {rd.status_code}")
            if rd.status_code == 200 and rd.content:
                out_path.write_bytes(rd.content)
                downloaded = True
                print(f"Downloaded via job link: {out_path.resolve()}")
                break
        except Exception as ex:
            print(f"Download error {href}: {ex}")

    for root in _base_url_candidates(cfg["base_url"]):
        if downloaded:
            break
        dl_candidates = [
            f"{root}/interop/rest/v1/files/{out_file}",
            f"{root}/interop/rest/v2/files/{out_file}",
            f"{root}/HyperionPlanning/rest/{cfg['api_version']}/applications/{cfg['application']}/files/{out_file}",
            f"{root}/rest/{cfg['api_version']}/applications/{cfg['application']}/files/{out_file}",
        ]
        for u in dl_candidates:
            try:
                rd = session.get(u, headers=headers, timeout=60, verify=False)
                print(f"GET {u} -> {rd.status_code}")
                if rd.status_code == 200 and rd.content:
                    out_path.write_bytes(rd.content)
                    downloaded = True
                    print(f"Downloaded: {out_path.resolve()}")
                    break
            except Exception as ex:
                print(f"Download error {u}: {ex}")
        if downloaded:
            break

    if not downloaded:
        print("Could not download Export_Account.zip from REST file endpoints.")


if __name__ == "__main__":
    main()
