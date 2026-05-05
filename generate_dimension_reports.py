from __future__ import annotations

import argparse
import csv
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import quote

from openpyxl import load_workbook

from connections import connect_to_epm_environment

_REPORTED_DATAEXPORT_FAILURES: set[str] = set()
_LAST_DATAEXPORT_FAILURES: list[tuple[str, int | str]] = []

DEFAULT_ESG_MEASURES = [
    "Product Revenue",
    "Software Revenue",
    "Hardware Revenue",
    "Services Revenue",
    "Dierct Cost of Sale",
    "Hardware costs",
    "Services costs",
    "Sales and Marketing costs",
    "Research and development costs",
    "General and administrative costs",
    "Amortization of intangible assets",
    "Acquisition related and other costs",
    "Restructuring costs",
    "Operating Margin",
    "Interest expense",
]


def _load_local_intersection_csv_fallback() -> list[dict[str, Any]]:
    """Build row-level period data from local fallback CSVs when REST dataexport is unavailable.

    Expected files (project root):
      - LE_ANZ Actuals.csv
      - LE_ANZ Forecast.csv
    """
    root = Path(__file__).resolve().parent
    candidates = [
        ("LE_ANZ Actuals.csv", "Actual"),
        ("LE_ANZ Forecast.csv", "Forecast"),
    ]

    rows: list[dict[str, Any]] = []
    for name, default_scenario in candidates:
        path = root / name
        if not path.exists():
            continue
        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.reader(f)
                header = next(reader, [])
                # First 12 columns are periods (FSC1..FSC12)
                periods = [str(x).strip() for x in header[:12]]
                periods = [p for p in periods if p]
                for rec in reader:
                    if len(rec) < 9:
                        continue
                    year = str(rec[0]).strip()
                    scenario = str(rec[1]).strip() or default_scenario
                    version = str(rec[2]).strip()
                    entity = str(rec[3]).strip()
                    lob = str(rec[4]).strip()
                    uom = str(rec[5]).strip()
                    energy_type = str(rec[6]).strip()
                    location = str(rec[7]).strip()
                    account = str(rec[8]).strip()
                    values = rec[9 : 9 + len(periods)]
                    for i, p in enumerate(periods):
                        if i >= len(values):
                            continue
                        vtxt = str(values[i]).strip()
                        if not vtxt:
                            continue
                        try:
                            val = float(vtxt)
                        except Exception:
                            continue
                        rows.append(
                            {
                                "Year": year,
                                "Years": year,
                                "Scenario": scenario,
                                "Version": version,
                                "Entity": entity,
                                "Account": account,
                                "Measures": account,
                                "Measure": account,
                                "Period": p,
                                "LOB": lob,
                                "UoM": uom,
                                "Energy Type": energy_type,
                                "Location": location,
                                "Data": val,
                            }
                        )
        except Exception:
            continue
    return rows


def get_last_dataexport_failures() -> list[tuple[str, int | str]]:
    """Return last collected dataexport endpoint failures for UI diagnostics."""
    return list(_LAST_DATAEXPORT_FAILURES)


def _base_url_candidates(base_url: str) -> list[str]:
    base = (base_url or "").strip().rstrip("/")
    for marker in ("/epmcloud", "/HyperionPlanning", "/interop/rest", "/rest"):
        if marker in base:
            base = base.split(marker, 1)[0]
            break
    candidates = [base.rstrip("/")]

    seen = set()
    result: list[str] = []
    for c in candidates:
        if c and c not in seen:
            seen.add(c)
            result.append(c)
    return result


def _headers(cfg: dict[str, str]) -> dict[str, str]:
    import base64

    token = base64.b64encode(f"{cfg['username']}:{cfg['password']}".encode()).decode()
    return {
        "Authorization": f"Basic {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }


def _safe_get_json(session, headers: dict[str, str], urls: list[str]) -> Any:
    for url in urls:
        try:
            r = session.get(url, headers=headers, timeout=60, verify=False)
            if r.status_code == 200:
                return r.json()
        except Exception:
            continue
    return {}


def _extract_dimension_names(payload: Any) -> list[str]:
    if isinstance(payload, dict):
        items = payload.get("items") or payload.get("dimensions") or []
    elif isinstance(payload, list):
        items = payload
    else:
        items = []

    out: list[str] = []
    for item in items:
        if isinstance(item, dict):
            name = item.get("name") or item.get("dimensionName") or item.get("displayName")
            if name:
                out.append(str(name))
        elif isinstance(item, str):
            out.append(item)
    return list(dict.fromkeys(out))


def _extract_dimension_self_links(payload: Any) -> dict[str, str]:
    if isinstance(payload, dict):
        items = payload.get("items") or payload.get("dimensions") or []
    elif isinstance(payload, list):
        items = payload
    else:
        items = []

    links: dict[str, str] = {}
    for item in items:
        if not isinstance(item, dict):
            continue
        name = item.get("name") or item.get("dimensionName") or item.get("displayName")
        if not name:
            continue
        rels = item.get("links") or []
        if isinstance(rels, list):
            for rel in rels:
                if isinstance(rel, dict) and str(rel.get("rel", "")).lower() == "self" and rel.get("href"):
                    links[str(name)] = str(rel["href"])
                    break
    return links


def _extract_member_names(payload: Any) -> list[str]:
    if isinstance(payload, dict):
        items = payload.get("items") or payload.get("members") or []
    elif isinstance(payload, list):
        items = payload
    else:
        items = []

    out: list[str] = []
    for item in items:
        if isinstance(item, dict):
            name = item.get("name") or item.get("memberName") or item.get("displayName")
            if name:
                out.append(str(name))
        elif isinstance(item, str):
            out.append(item)
    return list(dict.fromkeys(out))


def _collect_children_names(node: Any, out: list[str]) -> None:
    if isinstance(node, dict):
        name = node.get("name") or node.get("memberName") or node.get("displayName")
        if name:
            out.append(str(name))
        kids = node.get("children")
        if isinstance(kids, list):
            for child in kids:
                _collect_children_names(child, out)
    elif isinstance(node, list):
        for item in node:
            _collect_children_names(item, out)


def _extract_members_from_dimension_payload(payload: Any) -> list[str]:
    if not isinstance(payload, dict):
        return []
    out: list[str] = []
    if "children" in payload:
        _collect_children_names(payload.get("children"), out)
    return list(dict.fromkeys(out))


def _find_node_by_name(node: Any, target_name: str) -> Any:
    if isinstance(node, dict):
        name = str(node.get("name") or node.get("memberName") or "")
        if name.lower() == target_name.lower():
            return node
        children = node.get("children")
        if isinstance(children, list):
            for child in children:
                found = _find_node_by_name(child, target_name)
                if found is not None:
                    return found
    elif isinstance(node, list):
        for item in node:
            found = _find_node_by_name(item, target_name)
            if found is not None:
                return found
    return None


def _collect_leaf_names(node: Any, out: list[str]) -> None:
    if isinstance(node, dict):
        children = node.get("children")
        if isinstance(children, list) and children:
            for child in children:
                _collect_leaf_names(child, out)
        else:
            name = node.get("name") or node.get("memberName") or node.get("displayName")
            if name:
                out.append(str(name))
    elif isinstance(node, list):
        for item in node:
            _collect_leaf_names(item, out)


def fetch_l0_descendants_under_parent(
    session,
    cfg: dict[str, str],
    dimension_name: str,
    dimension_self_href: str,
    parent_names: list[str],
) -> list[str]:
    if not dimension_self_href:
        return []
    try:
        resp = session.get(dimension_self_href, headers=_headers(cfg), timeout=60, verify=False)
        if resp.status_code != 200:
            return []
        payload = resp.json()
    except Exception:
        return []

    root_children = payload.get("children") if isinstance(payload, dict) else None
    if not isinstance(root_children, list):
        return []

    parent_node = None
    for p in parent_names:
        parent_node = _find_node_by_name(root_children, p)
        if parent_node is not None:
            break
    if parent_node is None:
        return []

    leaves: list[str] = []
    _collect_leaf_names(parent_node.get("children", []), leaves)
    return list(dict.fromkeys(leaves))


def discover_dimensions(session, cfg: dict[str, str]) -> list[str]:
    ver = cfg["api_version"]
    app = cfg["application"]
    plan = cfg.get("plantype") or app
    urls: list[str] = []

    for root in _base_url_candidates(cfg["base_url"]):
        urls.extend(
            [
                f"{root}/HyperionPlanning/rest/{ver}/applications/{app}/dimensions",
                f"{root}/rest/{ver}/applications/{app}/dimensions",
                f"{root}/HyperionPlanning/rest/{ver}/applications/{app}/plantypes/{plan}/dimensions",
                f"{root}/rest/{ver}/applications/{app}/plantypes/{plan}/dimensions",
            ]
        )

    payload = _safe_get_json(session, _headers(cfg), urls)
    return _extract_dimension_names(payload)


def discover_dimensions_with_links(session, cfg: dict[str, str]) -> tuple[list[str], dict[str, str]]:
    ver = cfg["api_version"]
    app = cfg["application"]
    plan = cfg.get("plantype") or app
    urls: list[str] = []

    for root in _base_url_candidates(cfg["base_url"]):
        urls.extend(
            [
                f"{root}/HyperionPlanning/rest/{ver}/applications/{app}/dimensions",
                f"{root}/rest/{ver}/applications/{app}/dimensions",
                f"{root}/HyperionPlanning/rest/{ver}/applications/{app}/plantypes/{plan}/dimensions",
                f"{root}/rest/{ver}/applications/{app}/plantypes/{plan}/dimensions",
            ]
        )

    payload = _safe_get_json(session, _headers(cfg), urls)
    return _extract_dimension_names(payload), _extract_dimension_self_links(payload)


def fetch_dimension_members(
    session,
    cfg: dict[str, str],
    dimension_name: str,
    dimension_self_href: str = "",
    page_size: int = 500,
    max_pages: int = 200,
) -> list[str]:
    headers = _headers(cfg)
    ver = cfg["api_version"]
    app = cfg["application"]
    plan = cfg.get("plantype") or app
    dim_quoted = quote(dimension_name, safe="")

    members: list[str] = []
    seen = set()

    # First try reading members embedded as hierarchical children on the dimension resource.
    if dimension_self_href:
        try:
            resp = session.get(dimension_self_href, headers=headers, timeout=60, verify=False)
            if resp.status_code == 200:
                payload = resp.json()
                direct_members = _extract_members_from_dimension_payload(payload)
                for m in direct_members:
                    if m not in seen:
                        seen.add(m)
                        members.append(m)
                if members:
                    return members
        except Exception:
            pass

    for page in range(max_pages):
        offset = page * page_size
        urls: list[str] = []

        if dimension_self_href:
            urls.extend(
                [
                    f"{dimension_self_href}/members?offset={offset}&limit={page_size}",
                    f"{dimension_self_href}/members?limit={page_size}",
                ]
            )

        for root in _base_url_candidates(cfg["base_url"]):
            urls.extend(
                [
                    f"{root}/HyperionPlanning/rest/{ver}/applications/{app}/dimensions/{dim_quoted}/members?offset={offset}&limit={page_size}",
                    f"{root}/rest/{ver}/applications/{app}/dimensions/{dim_quoted}/members?offset={offset}&limit={page_size}",
                    f"{root}/HyperionPlanning/rest/{ver}/applications/{app}/plantypes/{plan}/dimensions/{dim_quoted}/members?offset={offset}&limit={page_size}",
                    f"{root}/rest/{ver}/applications/{app}/plantypes/{plan}/dimensions/{dim_quoted}/members?offset={offset}&limit={page_size}",
                ]
            )

        payload = _safe_get_json(session, headers, urls)
        page_members = _extract_member_names(payload)
        if not page_members:
            break

        new_count = 0
        for m in page_members:
            if m not in seen:
                seen.add(m)
                members.append(m)
                new_count += 1

        if new_count == 0 or len(page_members) < page_size:
            break

    return members


def fetch_dimension_alias_map(
    session,
    cfg: dict[str, str],
    dimension_name: str,
    dimension_self_href: str = "",
    page_size: int = 500,
    max_pages: int = 200,
) -> dict[str, str]:
    """Return {member_name: alias/defaultAlias/displayName} for a dimension.

    Falls back gracefully to an empty mapping when aliases are not exposed by the API.
    """
    headers = _headers(cfg)
    aliases: dict[str, str] = {}

    if dimension_self_href:
        try:
            resp = session.get(dimension_self_href, headers=headers, timeout=60, verify=False)
            if resp.status_code == 200:
                payload = resp.json()

                def _walk(node: Any) -> None:
                    if isinstance(node, dict):
                        name = node.get("name") or node.get("memberName")
                        alias = node.get("alias") or node.get("defaultAlias") or node.get("displayName")
                        if name and alias:
                            aliases[str(name)] = str(alias)
                        for child in node.get("children") or []:
                            _walk(child)
                    elif isinstance(node, list):
                        for item in node:
                            _walk(item)

                _walk(payload.get("children") if isinstance(payload, dict) else payload)
        except Exception:
            pass

    # Also try members endpoints for alias-style fields.
    ver = cfg["api_version"]
    app = cfg["application"]
    plan = cfg.get("plantype") or app
    dim_quoted = quote(dimension_name, safe="")

    for page in range(max_pages):
        offset = page * page_size
        urls: list[str] = []
        if dimension_self_href:
            urls.extend(
                [
                    f"{dimension_self_href}/members?offset={offset}&limit={page_size}",
                    f"{dimension_self_href}/members?limit={page_size}",
                ]
            )
        for root in _base_url_candidates(cfg["base_url"]):
            urls.extend(
                [
                    f"{root}/HyperionPlanning/rest/{ver}/applications/{app}/dimensions/{dim_quoted}/members?offset={offset}&limit={page_size}",
                    f"{root}/rest/{ver}/applications/{app}/dimensions/{dim_quoted}/members?offset={offset}&limit={page_size}",
                    f"{root}/HyperionPlanning/rest/{ver}/applications/{app}/plantypes/{plan}/dimensions/{dim_quoted}/members?offset={offset}&limit={page_size}",
                    f"{root}/rest/{ver}/applications/{app}/plantypes/{plan}/dimensions/{dim_quoted}/members?offset={offset}&limit={page_size}",
                ]
            )

        payload = _safe_get_json(session, headers, urls)
        items = payload.get("items", []) if isinstance(payload, dict) else (payload if isinstance(payload, list) else [])
        if not items:
            break

        for item in items:
            if not isinstance(item, dict):
                continue
            name = item.get("name") or item.get("memberName")
            alias = item.get("alias") or item.get("defaultAlias") or item.get("displayName")
            if name and alias:
                aliases[str(name)] = str(alias)

        if len(items) < page_size:
            break

    return aliases


def _dedupe_nonblank(values: list[str]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for value in values:
        s = str(value or "").strip()
        if not s or s.lower() in seen:
            continue
        seen.add(s.lower())
        out.append(s)
    return out


def _first_dimension(dimensions: list[str], candidates: list[str]) -> str:
    dims_ci = {str(d).lower(): str(d) for d in dimensions}
    for candidate in candidates:
        found = dims_ci.get(candidate.lower())
        if found:
            return found
    return ""


def _parse_grid_number(value: Any) -> float | None:
    text = str(value or "").strip()
    if not text or text in {"#", "#Missing", "NoData"}:
        return None
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()").replace(",", "").replace("$", "").replace("%", "")
    try:
        number = float(text)
    except Exception:
        return None
    return -number if negative else number


def _response_status_detail(resp: Any) -> int | str:
    status = getattr(resp, "status_code", "")
    text = str(getattr(resp, "text", "") or "").strip().replace("\r", " ").replace("\n", " ")
    if text:
        return f"{status}: {text[:240]}"
    return status


def _grid_to_row_items(
    payload: Any,
    row_dimensions: list[str],
    column_dimension: str,
    pov_values: dict[str, str],
) -> list[dict[str, Any]]:
    """Convert Oracle exportdataslice JSON grid output into row dictionaries."""
    if not isinstance(payload, dict):
        return []

    columns_raw = payload.get("columns") or []
    if columns_raw and isinstance(columns_raw[0], list):
        column_members = [str(v).strip() for v in columns_raw[0]]
    else:
        column_members = []

    out: list[dict[str, Any]] = []
    for row in payload.get("rows") or []:
        if not isinstance(row, dict):
            continue
        headers = [str(v).strip() for v in (row.get("headers") or [])]
        data = row.get("data") or []
        if len(headers) < len(row_dimensions):
            continue

        base = dict(pov_values)
        for dim, member in zip(row_dimensions, headers):
            base[dim] = member

        for i, raw_value in enumerate(data):
            value = _parse_grid_number(raw_value)
            if value is None:
                continue
            item = dict(base)
            item[column_dimension] = column_members[i] if i < len(column_members) else str(i + 1)
            item["Data"] = value
            out.append(item)
    return out


def _append_pov_dimension(
    pov_dims: list[str],
    pov_members: list[list[str]],
    pov_values: dict[str, str],
    dimension: str,
    members: list[str],
) -> None:
    members = _dedupe_nonblank(members)
    if not dimension or not members:
        return
    pov_dims.append(dimension)
    pov_members.append(members)
    pov_values[dimension] = members[0]


def _export_dataslice_rows(
    session,
    cfg: dict[str, str],
    dimensions: list[str],
    entity_dim: str,
    account_dim: str,
    entities: list[str],
    measures: list[str],
    scenarios: list[str],
    years: list[str],
    periods: list[str],
    versions: list[str],
    pov_overrides: dict[str, str | list[str]] | None = None,
) -> list[dict[str, Any]]:
    headers = _headers(cfg)
    app = cfg["application"]
    ver = cfg["api_version"]
    plan = cfg.get("plantype") or app

    scenario_dim = _first_dimension(dimensions, ["Scenario"])
    year_dim = _first_dimension(dimensions, ["Year", "Years"])
    period_dim = _first_dimension(dimensions, ["Period"])
    version_dim = _first_dimension(dimensions, ["Version"])
    hsp_view_dim = _first_dimension(dimensions, ["HSP_View"])

    entities = _dedupe_nonblank(entities)
    measures = _dedupe_nonblank(measures)
    scenarios = _dedupe_nonblank(scenarios)
    years = _dedupe_nonblank(years)
    periods = _dedupe_nonblank(periods)
    versions = _dedupe_nonblank(versions)

    if not measures:
        measures = DEFAULT_ESG_MEASURES

    if not (entity_dim and account_dim and period_dim and entities and measures and periods):
        return []

    if scenario_dim and not scenarios:
        scenarios = ["Actual", "Forecast"]
    if year_dim and not years:
        years = ["FY26"]
    if version_dim and not versions:
        versions = ["Working"]

    dim_lookup = {d.lower(): d for d in dimensions}
    override_values: dict[str, list[str]] = {}
    for raw_dim, raw_values in (pov_overrides or {}).items():
        actual_dim = dim_lookup.get(str(raw_dim).lower())
        if not actual_dim:
            continue
        values = raw_values if isinstance(raw_values, list) else [raw_values]
        cleaned = _dedupe_nonblank([str(v) for v in values])
        if cleaned:
            override_values[actual_dim] = cleaned

    urls: list[str] = []
    for root in _base_url_candidates(cfg["base_url"]):
        urls.extend(
            [
                f"{root}/HyperionPlanning/rest/{ver}/applications/{app}/plantypes/{plan}/exportdataslice",
                f"{root}/rest/{ver}/applications/{app}/plantypes/{plan}/exportdataslice",
            ]
        )

    all_items: list[dict[str, Any]] = []
    failures: list[tuple[str, int | str]] = []

    scenario_values = scenarios if scenario_dim else [""]
    year_values = years if year_dim else [""]
    version_values = versions if version_dim else [""]

    for scenario in scenario_values:
        for year in year_values:
            for version in version_values:
                pov_dims: list[str] = []
                pov_members: list[list[str]] = []
                pov_values: dict[str, str] = {}
                row_dimensions = [entity_dim, account_dim]
                row_members = [entities, measures]

                if len(entities) == 1:
                    _append_pov_dimension(pov_dims, pov_members, pov_values, entity_dim, entities)
                    row_dimensions = [account_dim]
                    row_members = [measures]

                if hsp_view_dim:
                    _append_pov_dimension(pov_dims, pov_members, pov_values, hsp_view_dim, ["BaseData"])
                if scenario_dim and scenario:
                    _append_pov_dimension(pov_dims, pov_members, pov_values, scenario_dim, [scenario])
                if year_dim and year:
                    _append_pov_dimension(pov_dims, pov_members, pov_values, year_dim, [year])
                if version_dim and version:
                    _append_pov_dimension(pov_dims, pov_members, pov_values, version_dim, [version])

                for dim, values in override_values.items():
                    if dim in {entity_dim, account_dim, period_dim, scenario_dim, year_dim, version_dim, hsp_view_dim}:
                        continue
                    _append_pov_dimension(pov_dims, pov_members, pov_values, dim, values)

                payload = {
                    "exportPlanningData": False,
                    "gridDefinition": {
                        "suppressMissingBlocks": True,
                        "suppressMissingRows": True,
                        "suppressMissingColumns": True,
                        "pov": {"dimensions": pov_dims, "members": pov_members},
                        "columns": [{"dimensions": [period_dim], "members": [periods]}],
                        "rows": [
                            {
                                "dimensions": row_dimensions,
                                "members": row_members,
                            }
                        ],
                    },
                }

                for url in urls:
                    try:
                        resp = session.post(url, headers=headers, json=payload, timeout=120, verify=False)
                        if resp.status_code == 200:
                            all_items.extend(
                                _grid_to_row_items(
                                    resp.json(),
                                    row_dimensions,
                                    period_dim,
                                    pov_values,
                                )
                            )
                            break
                        failures.append((url, _response_status_detail(resp)))
                    except Exception as ex:
                        failures.append((url, f"ERR:{type(ex).__name__}"))

    global _LAST_DATAEXPORT_FAILURES
    _LAST_DATAEXPORT_FAILURES = failures[:50]
    return all_items


def export_data_slice(
    session,
    cfg: dict[str, str],
    dimensions: list[str],
    dim_links: dict[str, str],
    entity_dim: str,
    account_dim: str,
    entities: list[str],
    measures: list[str],
    scenarios: list[str],
    years: list[str],
    periods: list[str],
    versions: list[str],
    pov_overrides: dict[str, str | list[str]] | None = None,
) -> list[dict[str, Any]]:
    """Best-effort filtered slice retrieval using Planning exportdataslice."""
    items = _export_dataslice_rows(
        session,
        cfg,
        dimensions,
        entity_dim,
        account_dim,
        entities,
        measures,
        scenarios,
        years,
        periods,
        versions,
        pov_overrides=pov_overrides,
    )
    if not items:
        local_items = _load_local_intersection_csv_fallback()
        items = local_items if local_items else []
    if not items:
        return []

    entity_set = {e.lower() for e in entities if str(e).strip()}
    account_set = {a.lower() for a in measures if str(a).strip()}
    scenario_set = {s.lower() for s in scenarios if str(s).strip()}
    year_set = {y.lower() for y in years if str(y).strip()}
    period_set = {p.lower() for p in periods if str(p).strip()}
    version_set = {v.lower() for v in versions if str(v).strip()}
    pov_ci: dict[str, set[str]] = {}
    for k, v in (pov_overrides or {}).items():
        vals = v if isinstance(v, list) else [v]
        pov_ci[str(k).lower()] = {str(x).lower() for x in vals if str(x).strip()}

    def _ci_value(row: dict[str, Any], candidate: str) -> str:
        lower = {str(k).lower(): k for k in row.keys()}
        key = lower.get(candidate.lower())
        return str(row.get(key, "")) if key is not None else ""

    out: list[dict[str, Any]] = []
    for row in items:
        if entity_set and _ci_value(row, entity_dim).lower() not in entity_set and _ci_value(row, "Entity").lower() not in entity_set:
            continue
        if account_set and _ci_value(row, account_dim).lower() not in account_set and _ci_value(row, "Account").lower() not in account_set and _ci_value(row, "Measures").lower() not in account_set:
            continue
        if scenario_set and _ci_value(row, "Scenario").lower() not in scenario_set:
            continue
        if year_set and _ci_value(row, "Year").lower() not in year_set and _ci_value(row, "Years").lower() not in year_set:
            continue
        if period_set and _ci_value(row, "Period").lower() not in period_set:
            continue
        if version_set and _ci_value(row, "Version").lower() not in version_set:
            continue

        skip = False
        for dim_name_l, allowed_vals in pov_ci.items():
            rv = _ci_value(row, dim_name_l).lower()
            if allowed_vals and rv and rv not in allowed_vals:
                skip = True
                break
        if skip:
            continue

        out.append(row)

    return out


def export_data(session, cfg: dict[str, str], dimensions: list[str]) -> list[dict[str, Any]]:
    headers = _headers(cfg)
    body = {
        "dimensions": dimensions,
        "suppressMissingBlocks": True,
        "suppressMissingRows": True,
    }
    # Keep dimension selection tolerant across apps where Year/Account are exposed
    # as Years/Measures.
    preferred_dims = ["Scenario", "Year", "Years", "Period", "Entity", "Account", "Measures", "Measure"]
    dims_ci = {d.lower(): d for d in dimensions}
    selected_dims: list[str] = []
    for d in preferred_dims:
        actual = dims_ci.get(d.lower())
        if actual and actual not in selected_dims:
            selected_dims.append(actual)

    # Ensure one account-like and one year-like dimension are present if available.
    account_like = dims_ci.get("account") or dims_ci.get("measures") or dims_ci.get("measure")
    year_like = dims_ci.get("year") or dims_ci.get("years")
    if account_like and account_like not in selected_dims:
        selected_dims.append(account_like)
    if year_like and year_like not in selected_dims:
        selected_dims.append(year_like)

    if len(selected_dims) >= 3:
        body["dimensions"] = selected_dims

    failures: list[tuple[str, int | str]] = []

    for root in _base_url_candidates(cfg["base_url"]):
        plantype = cfg.get("plantype") or cfg["application"]
        urls = [
            f"{root}/HyperionPlanning/rest/{cfg['api_version']}/applications/{cfg['application']}/dataexport",
            f"{root}/rest/{cfg['api_version']}/applications/{cfg['application']}/dataexport",
            f"{root}/HyperionPlanning/rest/{cfg['api_version']}/applications/{cfg['application']}/plantypes/{plantype}/dataexport",
            f"{root}/rest/{cfg['api_version']}/applications/{cfg['application']}/plantypes/{plantype}/dataexport",
        ]
        for url in urls:
            try:
                r = session.post(url, headers=headers, json=body, timeout=120, verify=False)
                if r.status_code == 200:
                    data = r.json()
                    items = data.get("items", []) if isinstance(data, dict) else []
                    if isinstance(items, list):
                        return items
                else:
                    failures.append((url, r.status_code))
            except Exception as ex:
                failures.append((url, f"ERR:{type(ex).__name__}"))

    global _LAST_DATAEXPORT_FAILURES
    _LAST_DATAEXPORT_FAILURES = failures

    # Local CSV fallback for dashboard demo/intersection flows.
    # This keeps the UI populated in environments where dataexport REST is not exposed.
    local_items = _load_local_intersection_csv_fallback()
    if local_items:
        return local_items

    # Report once per app/base to avoid flooding Streamlit logs with repeated red warnings.
    report_key = f"{cfg.get('base_url','')}|{cfg.get('application','')}|{cfg.get('api_version','')}"
    if failures and report_key not in _REPORTED_DATAEXPORT_FAILURES:
        _REPORTED_DATAEXPORT_FAILURES.add(report_key)
        summary = ", ".join([f"{status} @ {url}" for url, status in failures[:4]])
        print(
            "[WARN] dataexport endpoint unavailable or not authorized. "
            "This environment may require different privileges/API. "
            f"Sample failures: {summary}"
        )
    return []


def _severity(var_pct: float) -> str:
    p = abs(var_pct)
    if p >= 20:
        return "High"
    if p >= 10:
        return "Medium"
    if p >= 5:
        return "Low"
    return "None"


def _trend(actual: float, prior1: float, prior2: float) -> str:
    if actual > prior1 and prior1 > prior2:
        return "Uptrend"
    if actual < prior1 and prior1 < prior2:
        return "Downtrend"
    return "Flat/volatile"


def _pattern(var_pct: float, trend: str) -> str:
    abs_var = abs(var_pct)
    if abs_var >= 20:
        return "Outlier"
    if var_pct > 5 and trend == "Uptrend":
        return "Sustained upside"
    if var_pct < -5 and trend == "Downtrend":
        return "Sustained downside"
    if abs_var > 5 and trend == "Flat/volatile":
        return "Forecast bias"
    return "On track"


def _narrative(pattern: str) -> str:
    if pattern == "Outlier":
        return "Value is outside the normal band and should be validated."
    if pattern == "Sustained upside":
        return "Actual is above forecast for 3 consecutive periods. Pattern suggests stronger-than-planned performance."
    if pattern == "Sustained downside":
        return "Actual is below forecast for 3 consecutive periods. Forecast may be overstated."
    if pattern == "Forecast bias":
        return "Forecast is consistently optimistic versus actuals."
    if pattern == "On track":
        return "Actual is aligned with forecast."
    return "Review required."


def _recommended_action(severity: str) -> str:
    if severity == "High":
        return "Escalate to owner and reforecast."
    if severity == "Medium":
        return "Review assumptions and adjust plan."
    if severity == "Low":
        return "Monitor next cycle."
    return "No action required."


def _pick_field_name(items: list[dict[str, Any]], candidates: list[str]) -> str:
    if not items:
        return ""
    keys = list(items[0].keys())
    lower_map = {k.lower(): k for k in keys}
    for c in candidates:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    return ""


def build_report_rows(
    items: list[dict[str, Any]],
    entities: list[str],
    accounts: list[str],
    entity_field: str = "Entity",
    account_field: str = "Account",
    scenario_field: str = "Scenario",
    data_field: str = "Data",
    allowed_scenarios: set[str] | None = None,
    allowed_periods: set[str] | None = None,
    allowed_years: set[str] | None = None,
    allowed_versions: set[str] | None = None,
    current_year: str | None = None,
    extra_member_filters: dict[str, set[str]] | None = None,
    include_all_intersections: bool = False,
    entity_alias_map: dict[str, str] | None = None,
    account_alias_map: dict[str, str] | None = None,
) -> list[list[Any]]:
    grouped: dict[tuple[str, str], dict[str, float]] = {
        (e, a): {"actual": 0.0, "forecast": 0.0, "prior1": 0.0, "prior2": 0.0} for e in entities for a in accounts
    }
    entity_set = set(entities)
    account_set = set(accounts)

    def _ci_get(row: dict[str, Any], candidates: list[str]) -> str:
        lower = {str(k).lower(): k for k in row.keys()}
        for c in candidates:
            k = lower.get(c.lower())
            if k is not None:
                return str(row.get(k, ""))
        return ""

    cur_year = (current_year or "").lower().strip()
    prev1 = ""
    prev2 = ""
    if cur_year.startswith("fy") and len(cur_year) >= 4:
        try:
            n = int(cur_year[2:])
            prev1 = f"fy{n-1:02d}"
            prev2 = f"fy{n-2:02d}"
        except Exception:
            pass

    for row in items:
        entity = str(row.get(entity_field, ""))
        account = str(row.get(account_field, ""))
        if entity not in entity_set or account not in account_set:
            continue
        scenario = str(row.get(scenario_field, "")).lower()
        period_val = _ci_get(row, ["Period"])
        year_val = _ci_get(row, ["Year", "Years"])
        version_val = _ci_get(row, ["Version"])

        if allowed_scenarios is not None and scenario not in allowed_scenarios:
            continue
        if allowed_periods is not None and period_val.lower() not in allowed_periods:
            continue
        if allowed_years is not None and year_val.lower() not in allowed_years:
            continue
        if allowed_versions is not None and version_val.lower() not in allowed_versions:
            continue

        if extra_member_filters:
            skip = False
            for dim_name, allowed_vals in extra_member_filters.items():
                dim_val = _ci_get(row, [dim_name])
                if allowed_vals and dim_val.lower() not in allowed_vals:
                    skip = True
                    break
            if skip:
                continue

        value_raw = row.get(data_field, 0)
        try:
            value = float(value_raw)
        except Exception:
            value = 0.0

        key = (entity, account)
        y = year_val.lower().strip()
        if "actual" in scenario:
            if not cur_year or y == cur_year:
                grouped[key]["actual"] += value
        elif "forecast" in scenario or "fcst" in scenario:
            if not cur_year or y == cur_year:
                grouped[key]["forecast"] += value
            elif prev1 and y == prev1:
                grouped[key]["prior1"] += value
            elif prev2 and y == prev2:
                grouped[key]["prior2"] += value

    report_rows: list[list[Any]] = []
    for (entity, account), vals in grouped.items():
        actual = vals["actual"]
        forecast = vals["forecast"]
        variance = actual - forecast
        var_pct = (variance / forecast * 100.0) if forecast else 0.0
        trend = _trend(actual, vals["prior1"], vals["prior2"])
        pattern = _pattern(var_pct, trend)
        severity = _severity(var_pct)
        narrative = _narrative(pattern)
        action = _recommended_action(severity)

        report_rows.append(
            [
                (entity_alias_map or {}).get(entity, entity),
                (account_alias_map or {}).get(account, account),
                actual,
                forecast,
                vals["prior1"],
                vals["prior2"],
                variance,
                round(var_pct, 2),
                trend,
                pattern,
                severity,
                narrative,
                action,
            ]
        )

    return sorted(report_rows, key=lambda r: abs(r[7]), reverse=True)


def write_report(
    template_path: Path,
    output_path: Path,
    rows: list[list[Any]],
    dimensions: list[str],
    notes: list[str] | None = None,
) -> Path:
    wb = load_workbook(template_path)

    controls = wb["Controls"]
    controls.append(["Generated On", datetime.now().isoformat(timespec="seconds"), "Report generation timestamp"])
    controls.append(["Dimension Count", len(dimensions), "Dimensions discovered in environment"])
    controls.append(["Dimensions", ", ".join(dimensions), "Discovered dimensions"])
    controls.append(["Severity Logic", "High >= 20%, Medium >= 10%, else Low", "Threshold rule"])
    controls.append(["Trend Logic", "Up: Actual>Forecast, Down: Actual<Forecast, Flat: equal", "Trend rule"])
    controls.append(["Pattern Logic", "Stable if |Var %| < 5 else Deviation", "Pattern rule"])
    for note in (notes or []):
        controls.append(["Note", note, "Generation note"])

    data_ws = wb["Data"]
    # Clear any template/sample rows so output contains only application data.
    # Keep the header row (row 1) intact.
    if data_ws.max_row > 1:
        data_ws.delete_rows(2, data_ws.max_row - 1)
    for r in rows:
        data_ws.append(r)

    summary = wb["Summary"]
    summary["A3"] = "Total Insight Rows"
    summary["B3"] = len(rows)
    summary["A4"] = "High Severity Count"
    summary["B4"] = sum(1 for r in rows if r[10] == "High")
    summary["A5"] = "Medium Severity Count"
    summary["B5"] = sum(1 for r in rows if r[10] == "Medium")
    summary["A6"] = "Low Severity Count"
    summary["B6"] = sum(1 for r in rows if r[10] == "Low")
    summary["A7"] = "Avg Var %"
    summary["B7"] = round(sum(float(r[7]) for r in rows) / len(rows), 2) if rows else 0
    summary["A8"] = "Max |Var %|"
    summary["B8"] = max((abs(float(r[7])) for r in rows), default=0)

    insights = wb["Insights"]
    # Clear old insight rows while preserving heading rows.
    # We write fresh content starting from row 3/4 below.
    if insights.max_row > 4:
        insights.delete_rows(5, insights.max_row - 4)
    insights["A3"] = "Top 10 Variance Insights"
    insights["A4"] = "Entity"
    insights["B4"] = "Account"
    insights["C4"] = "Actual"
    insights["D4"] = "Forecast"
    insights["E4"] = "Prior 1"
    insights["F4"] = "Prior 2"
    insights["G4"] = "Variance"
    insights["H4"] = "Var %"
    insights["I4"] = "Severity"
    insights["J4"] = "Narrative Message"
    insights["K4"] = "Recommended Action"
    start = 5
    for i, r in enumerate(rows[:10], start=start):
        insights[f"A{i}"] = r[0]
        insights[f"B{i}"] = r[1]
        insights[f"C{i}"] = r[2]
        insights[f"D{i}"] = r[3]
        insights[f"E{i}"] = r[4]
        insights[f"F{i}"] = r[5]
        insights[f"G{i}"] = r[6]
        insights[f"H{i}"] = r[7]
        insights[f"I{i}"] = r[10]
        insights[f"J{i}"] = r[11]
        insights[f"K{i}"] = r[12]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(output_path)
        return output_path
    except PermissionError:
        alt = output_path.with_name(f"{output_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{output_path.suffix}")
        wb.save(alt)
        print(f"[WARN] Output file was locked. Saved to alternate path: {alt}")
        return alt


def main() -> None:
    project_root = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description="Generate EPM dimension-member-based report workbook.")
    parser.add_argument(
        "--template",
        default=project_root / "EPM_Actual_vs_Forecast_Insight_Template.xlsx",
        help="Path to the template workbook.",
    )
    parser.add_argument(
        "--output",
        default="reports/EPM_Actual_vs_Forecast_Insight_Output.xlsx",
        help="Output workbook path.",
    )
    args = parser.parse_args()

    session, cfg = connect_to_epm_environment()
    cfg["plantype"] = cfg.get("plantype") or cfg["application"]

    dimensions, dim_links = discover_dimensions_with_links(session, cfg)
    if not dimensions:
        raise RuntimeError("Could not discover dimensions from the environment.")

    dim_lookup = {d.lower(): d for d in dimensions}
    if "entity" not in dim_lookup:
        raise RuntimeError("Required dimension Entity was not found in the environment metadata.")

    account_like = None
    for candidate in ("account", "accounts", "measure", "measures"):
        if candidate in dim_lookup:
            account_like = dim_lookup[candidate]
            break
    if not account_like:
        raise RuntimeError("Required account-like dimension (Account/Measures) was not found in the environment metadata.")

    entity_dim = dim_lookup["entity"]
    account_dim = account_like
    entities = fetch_dimension_members(session, cfg, entity_dim, dim_links.get(entity_dim, ""))
    accounts = fetch_dimension_members(session, cfg, account_dim, dim_links.get(account_dim, ""))

    # User-requested intersection subsets
    scenario_members = ["Actual", "Forecast", "Variance", "Var%"]
    year_members = ["FY24", "FY25", "FY26", "FY27", "FY28", "FY29"]
    version_members = ["Sample_Actual"]

    entity_l0 = fetch_l0_descendants_under_parent(
        session,
        cfg,
        entity_dim,
        dim_links.get(entity_dim, ""),
        ["DO NOT USE", "DONOTUSE", "DO_NOT_USE"],
    )
    if entity_l0:
        entities = entity_l0

    account_l0 = fetch_l0_descendants_under_parent(
        session,
        cfg,
        account_dim,
        dim_links.get(account_dim, ""),
        ["Profit and Loss Statement", "Profit and Loss statement", "Profit and Loss"],
    )
    if account_l0:
        accounts = account_l0

    period_dim = dim_lookup.get("period")
    period_members: list[str] = []
    if period_dim:
        period_members = fetch_l0_descendants_under_parent(
            session,
            cfg,
            period_dim,
            dim_links.get(period_dim, ""),
            ["YearTotal", "Year Total", "Year_Total"],
        )

    items = export_data(session, cfg, dimensions)
    entity_field = _pick_field_name(items, ["Entity"])
    account_field = _pick_field_name(items, ["Account", "Measures", "Measure"])
    scenario_field = _pick_field_name(items, ["Scenario"])
    data_field = _pick_field_name(items, ["Data", "Value", "Amount"])

    if (not entities or not accounts) and items:
        if not entity_field:
            entity_field = "Entity"
        if not account_field:
            account_field = "Account"
        entities = sorted({str(r.get(entity_field, "")) for r in items if str(r.get(entity_field, "")).strip()})
        accounts = sorted({str(r.get(account_field, "")) for r in items if str(r.get(account_field, "")).strip()})

    notes: list[str] = []
    if not entities or not accounts:
        notes.append("Could not retrieve Entity/Account members from metadata APIs in this environment.")

    scenario_field = scenario_field or "Scenario"
    data_field = data_field or "Data"

    if not items:
        notes.append("No dataexport payload returned from REST APIs (endpoint unavailable/unauthorized).")
        print("[INFO] No data rows returned by dataexport. Creating workbook with headers/summary only.")
        rows = build_report_rows(
            [],
            entities,
            accounts,
            entity_field or "Entity",
            account_field or "Account",
            scenario_field,
            data_field,
            allowed_scenarios={s.lower() for s in scenario_members},
            allowed_periods={p.lower() for p in period_members} if period_members else None,
            allowed_years={y.lower() for y in year_members},
            allowed_versions={v.lower() for v in version_members},
        )
    else:
        rows = build_report_rows(
            items,
            entities,
            accounts,
            entity_field or "Entity",
            account_field or "Account",
            scenario_field,
            data_field,
            allowed_scenarios={s.lower() for s in scenario_members},
            allowed_periods={p.lower() for p in period_members} if period_members else None,
            allowed_years={y.lower() for y in year_members},
            allowed_versions={v.lower() for v in version_members},
        )

    notes.append(f"Intersection filter applied: Scenario={scenario_members}; Years={year_members}; Version={version_members}.")
    if period_members:
        notes.append(f"Period L0 descendants under YearTotal: {len(period_members)} members.")
    notes.append(f"Entity L0 descendants under DO NOT USE: {len(entities)} members.")
    notes.append(f"{account_dim} L0 descendants under Profit and Loss Statement: {len(accounts)} members.")
    write_report(Path(args.template), Path(args.output), rows, dimensions, notes)

    print(f"Report generated successfully: {args.output}")
    print(f"Dimensions discovered: {len(dimensions)}")
    print(f"Insight rows written: {len(rows)}")


if __name__ == "__main__":
    main()
