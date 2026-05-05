#!/usr/bin/env python3
"""Utility to load a flat (CSV) extract into EPBCS via the MCP API.

The script performs the following steps:

1. Load environment variables (MCP server URL, EPBCS base URL, username and
   password) from a ``.env`` file or the process environment.
2. Retrieve the list of dimensions from the EPBCS instance so we can map the
   column headers in the flat file to EPBCS dimensions.
3. Parse the CSV file provided on the command line. Each column is assumed to
   correspond to a dimension name (or a member of a dimension) as returned by
   the EPBCS dimensions endpoint.
4. Build a payload that represents an EPBCS *load* operation. The exact payload
   depends on the MCP ``/invoke`` contract used by the existing
   ``MCP_EPBCS_Client`` – we follow the same pattern as ``get_application_details``
   and ``run_business_rule``.
5. Send the payload to the MCP server via a POST request.

The implementation is deliberately defensive – it validates that every column
header maps to a known dimension and raises a clear error if the mapping is
ambiguous.  The script can be used as a library or run directly from the command
line.
"""

from __future__ import annotations

import csv
import os
import sys
from pathlib import Path
from typing import Dict, List, Sequence

import httpx
from dotenv import load_dotenv

# Re‑use the client implementation from the existing code base for consistency.
# Importing the class directly avoids duplicating authentication logic.
try:
    from mcp_epbcs_client import MCP_EPBCS_Client
except Exception:  # pragma: no cover – import may fail if path is unusual.
    # Fallback: define a minimal client with the same public interface.
    class MCP_EPBCS_Client:  # type: ignore
        def __init__(self) -> None:
            self.mcp_url = os.getenv("MCP_SERVER_URL")
            self.epbcs_url = os.getenv("EPBCS_BASE_URL")
            self.username = os.getenv("EPBCS_USERNAME")
            self.password = os.getenv("EPBCS_PASSWORD")
            if not all([self.mcp_url, self.epbcs_url, self.username, self.password]):
                raise ValueError("Missing required environment variables for MCP client")
            credentials = f"{self.username}:{self.password}"
            encoded = httpx._basic_auth._basic_auth(username=self.username, password=self.password)  # noqa: SLF001
            self.headers = {
                "Authorization": f"Basic {encoded}",
                "Content-Type": "application/json",
            }

        def _call_mcp(self, method: str, endpoint: str, payload: Dict | None = None):
            url = f"{self.mcp_url}{endpoint}"
            response = httpx.request(method=method, url=url, headers=self.headers, json=payload)
            response.raise_for_status()
            return response.json()

        def invoke(self, payload: Dict) -> Dict:
            """Convenience wrapper for the generic ``/invoke`` endpoint used in the
            existing client implementation.
            """
            return self._call_mcp("POST", "/invoke", payload)


def load_dimensions(client: MCP_EPBCS_Client) -> Dict[str, str]:
    """Return a mapping of dimension name -> storage type (Dense/Sparse).

    The function mirrors the logic from ``test_epbcs.py`` which queries the
    ``/dimensions`` endpoint.  Only the dimension name is needed for mapping CSV
    columns.
    """
    url = f"{client.epbcs_url.rstrip('/')}/HyperionPlanning/rest/v3/applications/{{application}}/plantypes/{{plantype}}/dimensions"
    # The placeholders ``{{application}}`` and ``{{plantype}}`` are expected to be
    # supplied via environment variables.  If they are missing we raise a clear
    # error.
    application = os.getenv("EPM_APPLICATION")
    plantype = os.getenv("EPM_PLANTYPE")
    if not application or not plantype:
        raise ValueError("EPM_APPLICATION and EPM_PLANTYPE must be set in the environment")

    url = url.format(application=application, plantype=plantype)
    headers = {
        "Authorization": client.headers["Authorization"],
        "Accept": "application/json",
    }
    resp = httpx.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    dimensions = data.get("items", [])
    mapping: Dict[str, str] = {}
    for dim in dimensions:
        if isinstance(dim, dict):
            name = (
                dim.get("name")
                or dim.get("dimensionName")
                or dim.get("id")
                or ""
            )
            if name:
                mapping[name] = dim.get("storage", "Sparse")
    return mapping


def parse_flat_file(csv_path: Path) -> List[Dict[str, str]]:
    """Read a CSV file and return a list of rows as dictionaries.

    The first row is interpreted as the header.  All values are kept as strings –
    EPBCS expects members/values to be strings in the load payload.
    """
    if not csv_path.is_file():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")
    with csv_path.open(newline="", encoding="utf-8") as fp:
        reader = csv.DictReader(fp)
        rows = [dict(row) for row in reader]
    if not rows:
        raise ValueError("CSV file contains no data rows")
    return rows


def build_load_payload(
    rows: Sequence[Dict[str, str]], dimensions: Dict[str, str]
) -> Dict:
    """Construct a payload suitable for the MCP ``/invoke`` load endpoint.

    The exact contract depends on the EPBCS implementation; we follow a generic
    structure used by the existing client where ``method`` is ``POST`` and the
    target URL points to the EPBCS ``/load`` endpoint.  Each row is wrapped under
    a ``data`` key.
    """
    # Validate that every column maps to a known dimension.
    for row in rows:
        for col in row.keys():
            if col not in dimensions:
                raise ValueError(
                    f"Column '{col}' does not match any known EPBCS dimension. "
                    f"Available dimensions: {', '.join(dimensions)}"
                )

    payload = {
        "method": "POST",
        "url": f"{os.getenv('EPBCS_BASE_URL').rstrip('/')}/load",
        "body": {
            "rows": rows,
            "dimensions": list(dimensions.keys()),
        },
    }
    return payload


def main() -> int:
    # Load environment variables from a .env file located next to this script.
    script_dir = Path(__file__).resolve().parent
    load_dotenv(script_dir / ".env")

    if len(sys.argv) != 2:
        print("Usage: python load_flat_file.py <path-to-csv>", file=sys.stderr)
        return 1

    csv_path = Path(sys.argv[1])
    try:
        rows = parse_flat_file(csv_path)
    except Exception as exc:
        print(f"Failed to read CSV file: {exc}", file=sys.stderr)
        return 1

    try:
        client = MCP_EPBCS_Client()
        dimensions = load_dimensions(client)
        payload = build_load_payload(rows, dimensions)
        response = client.invoke(payload)
        print("Load request submitted successfully. Server response:")
        print(response)
        return 0
    except Exception as exc:
        print(f"Error during integration: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
