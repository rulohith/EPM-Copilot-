from __future__ import annotations

import argparse
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
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import quote

import openpyxl
import requests
from dotenv import load_dotenv
import urllib3

# Add visualization libraries
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
from io import BytesIO

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


ACCOUNT_ZIP_NAME = "Export_Account.zip"
ACCOUNT_CSV_NAME = "ankita.roy@oracle.com_ExportedMetadata_Measures.csv"

LOB_ZIP_NAME = "Export_LOB.zip"
LOB_CSV_NAME = "ankita.roy@oracle.com_ExportedMetadata_LOB.csv"

EXCEL_FILE_NAME = "AccountsToBeAdded.xlsx"
LOB_EXCEL_FILE_NAME = "LOBsToBeAdded.xlsx"

DEBUG_MERGE = True

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

LOB_COLUMNS_TO_REMOVE = [
    "Alias: CY Table",
]


@dataclass
class MemberRequest:
    dimension: str  # ACCOUNT or LOB
    member: str
    parent: str
    description: str


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
        "base_url": os.getenv("EPM_BASE_URL", "").strip().rstrip("/"),
        "username": os.getenv("EPM_USERNAME", "").strip(),
        "password": os.getenv("EPM_PASSWORD", "").strip(),
        "api_version": os.getenv("EPM_API_VERSION", "v3").strip(),
        "application": os.getenv("EPM_APPLICATION", "").strip(),
        "plantype": os.getenv("EPM_PLANTYPE", "").strip(),
        "account_export_job_name": os.getenv("EPM_JOB_NAME", "Export Account").strip(),
        "account_import_job_name": os.getenv("EPM_IMPORT_JOB_NAME", "Import Account").strip(),
        "lob_export_job_name": os.getenv("EPM_LOB_EXPORT_JOB_NAME", "Export LOB").strip(),
        "lob_import_job_name": os.getenv("EPM_LOB_IMPORT_JOB_NAME", "Import LOB").strip(),
    }

    required = ["base_url", "username", "password", "api_version", "application"]
    missing = [k for k in required if not cfg.get(k)]
    if missing:
        raise ValueError(f"Missing required environment variables: {', '.join(missing)}")

    return cfg


def _base_url_candidates(base_url: str) -> list[str]:
    """Return likely REST roots for EPM cloud where customers may configure varying base URLs."""
    base = base_url.rstrip("/")
    candidates = [base]

    if base.endswith("/epmcloud"):
        candidates.append(base[: -len("/epmcloud")])

    if "/HyperionPlanning" in base:
        candidates.append(base.split("/HyperionPlanning", 1)[0])

    # Ordered unique
    unique: list[str] = []
    seen = set()
    for c in candidates:
        if c and c not in seen:
            unique.append(c)
            seen.add(c)
    return unique


def _connect_to_epm_instance(cfg: Dict[str, str]) -> requests.Session:
    """
    Connect to the EPM (EPBCS) instance and authenticate.
    Returns an authenticated requests Session object.
    """
    headers = _basic_auth_header(cfg["username"], cfg["password"])
    session = requests.Session()
    
    print(f"[CONNECTION] Connecting to EPM Instance...")
    print(f"[CONNECTION] Base URL: {cfg['base_url']}")
    print(f"[CONNECTION] Application: {cfg['application']}")
    print(f"[CONNECTION] API Version: {cfg['api_version']}")
    print(f"[CONNECTION] Username: {cfg['username']}")
    
    # Try multiple possible API paths and base roots
    possible_urls: list[str] = []
    for root in _base_url_candidates(cfg["base_url"]):
        possible_urls.extend([
            f"{root}/HyperionPlanning/rest/{cfg['api_version']}/applications/{cfg['application']}",
            f"{root}/HyperionPlanning/rest/{cfg['api_version']}/applications",
            f"{root}/rest/{cfg['api_version']}/applications/{cfg['application']}",
            f"{root}/rest/{cfg['api_version']}/applications",
            f"{root}/interop/rest",
        ])
    
    for test_url in possible_urls:
        print(f"\n[CONNECTION] Testing URL: {test_url}")
        try:
            resp = session.get(test_url, headers=headers, timeout=30, verify=False)
            print(f"[CONNECTION] Response Status: {resp.status_code}")
            
            if resp.status_code == 200:
                try:
                    app_info = resp.json()
                    print(f"[CONNECTION] SUCCESS - Connected to EPM instance")
                    print(f"[CONNECTION] Response: {json.dumps(app_info, indent=2)}")
                    return session
                except ValueError:
                    # Response is 200 but HTML usually means unauthenticated SSO/login page.
                    sample = (resp.text or "")[:200].lower()
                    if "<html" in sample or "login" in sample:
                        print("[CONNECTION] 200 HTML login response detected; trying next endpoint...")
                        continue
                    print(f"[CONNECTION] SUCCESS - Connected (200 OK)")
                    print(f"[CONNECTION] Response: {resp.text[:200]}")
                    return session
            elif resp.status_code == 401:
                print(f"[CONNECTION] Authentication failed (401)")
                raise RuntimeError("Invalid credentials")
            elif resp.status_code == 404:
                print(f"[CONNECTION] Endpoint not found (404)")
                continue
            else:
                print(f"[CONNECTION] Status {resp.status_code}")
                
        except Exception as e:
            print(f"[CONNECTION] Error: {str(e)[:100]}")
            continue
    
    # If we get here, none of the URLs worked
    raise RuntimeError(
        f"Could not connect to EPM instance. Tried multiple endpoints.\n"
        f"Verify the EPM_BASE_URL and EPM_APPLICATION settings in .env"
    )


def _get_cube_metadata(session: requests.Session, cfg: Dict[str, str]) -> Dict[str, Any]:
    """
    Retrieve metadata for the cube/application.
    Returns empty dict if metadata endpoint is not available.
    """
    headers = _basic_auth_header(cfg["username"], cfg["password"])
    
    print(f"[CUBE] Attempting to retrieve metadata for cube: {cfg['application']}")
    
    # Try multiple possible endpoints for dimensions with base-root fallbacks
    possible_urls: list[str] = []
    for root in _base_url_candidates(cfg["base_url"]):
        possible_urls.extend([
            f"{root}/HyperionPlanning/rest/{cfg['api_version']}/applications/{cfg['application']}/dimensions",
            f"{root}/rest/{cfg['api_version']}/applications/{cfg['application']}/dimensions",
        ])
        if cfg.get("plantype"):
            possible_urls.extend([
                f"{root}/HyperionPlanning/rest/{cfg['api_version']}/applications/{cfg['application']}/plantypes/{cfg['plantype']}/dimensions",
                f"{root}/rest/{cfg['api_version']}/applications/{cfg['application']}/plantypes/{cfg['plantype']}/dimensions",
            ])
    
    for test_url in possible_urls:
        try:
            resp = session.get(test_url, headers=headers, timeout=30, verify=False)
            if resp.status_code == 200:
                dimensions = resp.json()
                print(f"[CUBE] SUCCESS - Retrieved dimensions for cube: {cfg['application']}")
                print(f"[CUBE] Available Dimensions: {json.dumps(dimensions, indent=2)}")
                return dimensions
        except Exception as e:
            pass
    
    print(f"[CUBE] INFO - Cube metadata endpoint not available (expected in some EPM versions)")
    return {}


def initialize_epm_connection() -> tuple[requests.Session, Dict[str, str], Dict[str, Any]]:
    """
    Initialize the connection to EPM instance and retrieve cube metadata.
    Returns: (session, config, cube_metadata)
    """
    print("\n" + "="*60)
    print("INITIALIZING EPM CONNECTION AND CUBE METADATA")
    print("="*60 + "\n")
    
    # Load configuration from .env
    cfg = _load_config()
    
    # Connect to EPM instance
    session = _connect_to_epm_instance(cfg)
    
    # Get cube metadata
    cube_metadata = _get_cube_metadata(session, cfg)
    
    print("\n" + "="*60)
    print("EPM CONNECTION INITIALIZED SUCCESSFULLY")
    print("="*60 + "\n")
    
    return session, cfg, cube_metadata


# ---------------- DATA EXTRACTION ----------------
def _extract_dimension_names(cube_metadata: Any) -> list[str]:
    """
    Extract dimension names from cube metadata response in a tolerant way.
    """
    dimension_names: list[str] = []

    if not cube_metadata:
        return dimension_names

    # Common response shapes: {"items": [...]}, {"dimensions": [...]}, or direct list payload
    candidates: Any
    if isinstance(cube_metadata, dict):
        candidates = cube_metadata.get("items") or cube_metadata.get("dimensions") or []
    elif isinstance(cube_metadata, list):
        candidates = cube_metadata
    else:
        candidates = []

    for item in candidates:
        if isinstance(item, dict):
            dim_name = (
                item.get("name")
                or item.get("dimensionName")
                or item.get("displayName")
                or item.get("id")
            )
            if dim_name:
                dimension_names.append(str(dim_name))
        elif isinstance(item, str):
            dimension_names.append(item)

    # Preserve order while removing duplicates
    seen = set()
    ordered_unique = []
    for dim in dimension_names:
        if dim not in seen:
            ordered_unique.append(dim)
            seen.add(dim)

    return ordered_unique


def _extract_dimension_self_links(cube_metadata: Any) -> Dict[str, str]:
    """
    Extract mapping of dimension name -> self href URL from dimension payload.
    """
    result: Dict[str, str] = {}

    if isinstance(cube_metadata, dict):
        candidates = cube_metadata.get("items") or cube_metadata.get("dimensions") or []
    elif isinstance(cube_metadata, list):
        candidates = cube_metadata
    else:
        candidates = []

    for item in candidates:
        if not isinstance(item, dict):
            continue

        name = item.get("name") or item.get("dimensionName") or item.get("displayName") or item.get("id")
        if not name:
            continue

        links = item.get("links") or []
        self_href = ""
        if isinstance(links, list):
            for lnk in links:
                if isinstance(lnk, dict) and str(lnk.get("rel", "")).lower() == "self" and lnk.get("href"):
                    self_href = str(lnk["href"])
                    break

        if self_href:
            result[str(name)] = self_href

    return result


def _extract_member_names(member_payload: Any) -> list[str]:
    """
    Extract member names from a dimension members API payload.
    """
    member_names: list[str] = []

    if isinstance(member_payload, dict):
        candidates = member_payload.get("items") or member_payload.get("members") or []
    elif isinstance(member_payload, list):
        candidates = member_payload
    else:
        candidates = []

    for item in candidates:
        if isinstance(item, dict):
            member_name = (
                item.get("name")
                or item.get("memberName")
                or item.get("displayName")
                or item.get("id")
            )
            if member_name:
                member_names.append(str(member_name))
        elif isinstance(item, str):
            member_names.append(item)

    # Preserve order and uniqueness
    seen = set()
    unique_members = []
    for member in member_names:
        if member not in seen:
            unique_members.append(member)
            seen.add(member)

    return unique_members


def _safe_get_json(session: requests.Session, headers: Dict[str, str], urls: list[str]) -> Any:
    """
    Try multiple URLs and return the first successful JSON payload.
    """
    for url in urls:
        try:
            resp = session.get(url, headers=headers, timeout=60, verify=False)
            if resp.status_code == 200:
                try:
                    return resp.json()
                except ValueError:
                    continue
        except Exception:
            continue
    return {}


def _fetch_dimension_members_paginated(
    session: requests.Session,
    headers: Dict[str, str],
    base_url: str,
    api_version: str,
    application: str,
    dimension_name: str,
    plan_type: Optional[str] = None,
    dimension_self_href: str = "",
    page_size: int = 500,
    max_pages: int = 200,
) -> list[str]:
    """
    Fetch members for a dimension with pagination support.
    Some EPM APIs only return a partial list unless offset/limit are iterated.
    """
    members: list[str] = []
    seen = set()

    # IMPORTANT: quote with safe="" so '/' and special chars in dimension names are encoded
    dim_quoted = quote(dimension_name, safe="")

    for page in range(max_pages):
        offset = page * page_size
        member_urls: list[str] = []

        if dimension_self_href:
            member_urls.extend([
                f"{dimension_self_href}/members?offset={offset}&limit={page_size}",
                f"{dimension_self_href}/members?limit={page_size}",
            ])

        for root in _base_url_candidates(base_url):
            member_urls.extend([
                f"{root}/HyperionPlanning/rest/{api_version}/applications/{application}/dimensions/{dim_quoted}/members?offset={offset}&limit={page_size}",
                f"{root}/rest/{api_version}/applications/{application}/dimensions/{dim_quoted}/members?offset={offset}&limit={page_size}",
            ])
            if plan_type:
                member_urls.extend([
                    f"{root}/HyperionPlanning/rest/{api_version}/applications/{application}/plantypes/{plan_type}/dimensions/{dim_quoted}/members?offset={offset}&limit={page_size}",
                    f"{root}/rest/{api_version}/applications/{application}/plantypes/{plan_type}/dimensions/{dim_quoted}/members?offset={offset}&limit={page_size}",
                ])

        payload = _safe_get_json(session, headers, member_urls)
        page_members = _extract_member_names(payload)

        # Stop when no more rows are returned
        if not page_members:
            break

        new_count = 0
        for member in page_members:
            if member not in seen:
                seen.add(member)
                members.append(member)
                new_count += 1

        # Stop if API keeps returning the same page
        if new_count == 0:
            break

        # If page is smaller than requested, we reached the end
        if len(page_members) < page_size:
            break

    return members


def discover_and_print_dimensions_members(session: requests.Session, cfg: Dict[str, str]) -> Dict[str, list[str]]:
    """
    Discover dimensions and members from the application and print them.
    Returns a map: {dimension_name: [member1, member2, ...]}.
    """
    headers = _basic_auth_header(cfg["username"], cfg["password"])
    base = cfg["base_url"]
    ver = cfg["api_version"]
    app = cfg["application"]
    plan_type = cfg.get("plantype") or app

    print("\n" + "="*60)
    print("DISCOVERING DIMENSIONS AND MEMBERS FROM APPLICATION")
    print("="*60)

    dimension_urls: list[str] = []
    for root in _base_url_candidates(base):
        dimension_urls.extend([
            f"{root}/HyperionPlanning/rest/{ver}/applications/{app}/dimensions",
            f"{root}/rest/{ver}/applications/{app}/dimensions",
            f"{root}/HyperionPlanning/rest/{ver}/applications/{app}/plantypes/{plan_type}/dimensions",
            f"{root}/rest/{ver}/applications/{app}/plantypes/{plan_type}/dimensions",
            f"{root}/HyperionPlanning/rest/{ver}/applications/{app}/plantypes/Plan1/dimensions",
        ])

    dimension_payload = _safe_get_json(session, headers, dimension_urls)
    dimensions = _extract_dimension_names(dimension_payload)
    dimension_self_links = _extract_dimension_self_links(dimension_payload)

    if not dimensions:
        print("[METADATA] Could not retrieve dimensions from metadata APIs.")
        return {}

    print(f"[METADATA] Discovered Dimensions ({len(dimensions)}): {dimensions}")

    result: Dict[str, list[str]] = {}
    for dim in dimensions:
        members = _fetch_dimension_members_paginated(
            session=session,
            headers=headers,
            base_url=base,
            api_version=ver,
            application=app,
            dimension_name=dim,
            plan_type=plan_type,
            dimension_self_href=dimension_self_links.get(dim, ""),
        )
        result[dim] = members

        print(f"\n[METADATA] Dimension: {dim}")
        if members:
            print(f"[METADATA] Members fetched ({len(members)}):")
            # Print first N for readability while still returning full list
            preview_limit = 50
            for m in members[:preview_limit]:
                print(f"  - {m}")
            if len(members) > preview_limit:
                print(f"  ... and {len(members) - preview_limit} more")
        else:
            print("[METADATA] Members: <none returned or endpoint unavailable>")

    print("="*60 + "\n")
    return result


def extract_epm_data(session: requests.Session, cfg: Dict[str, str],
                    dimensions: Optional[list[str]] = None,
                    scenario_filter: Optional[str] = None) -> pd.DataFrame:
    """
    Extract data from EPM cube for reporting purposes.
    Returns only application data from EPM. No sample/fallback data is used.
    """
    headers = _basic_auth_header(cfg["username"], cfg["password"])
    
    print(f"[DATA EXTRACTION] Extracting data from cube: {cfg['application']}")
    
    # Try real data extraction first
    data_url = (
        f"{cfg['base_url']}/HyperionPlanning/rest/{cfg['api_version']}/applications/"
        f"{cfg['application']}/dataexport"
    )
    
    try:
        # Dimensions should come from the application metadata.
        # If not passed, try to retrieve dynamically from the application.
        if not dimensions:
            metadata = _get_cube_metadata(session, cfg)
            dimensions = _extract_dimension_names(metadata)

        if not dimensions:
            raise RuntimeError(
                "Could not determine application dimensions for data export. "
                "Please verify dimension metadata API access."
            )
        
        pov = {
            "dimensions": dimensions,
            "suppressMissingBlocks": True,
            "suppressMissingRows": True
        }
        
        resp = session.post(data_url, headers=headers, json=pov, timeout=120)
        resp.raise_for_status()
        
        data = resp.json()
        
        if 'items' in data:
            df = pd.DataFrame(data['items'])
        else:
            df = pd.DataFrame([data])

        if scenario_filter and 'Scenario' in df.columns:
            df = df[df['Scenario'].astype(str).str.lower() == scenario_filter.lower()]
        
        print(f"[DATA EXTRACTION] Extracted {len(df)} records from EPM")
        return df
        
    except requests.exceptions.RequestException as e:
        print(f"[DATA EXTRACTION] API extraction failed: {e}")
        return pd.DataFrame()


# ---------------- DATA PROCESSING ----------------
def process_comparison_data(df: pd.DataFrame, 
                          comparison_type: str = "actuals_vs_forecast") -> pd.DataFrame:
    """
    Process data for different types of comparisons.
    """
    print(f"[DATA PROCESSING] Processing data for {comparison_type} comparison")
    
    if df.empty:
        print("[DATA PROCESSING] Warning: Empty dataset provided")
        return df
    
    # Ensure we have the necessary columns
    required_cols = ['Account', 'Period', 'Year', 'Scenario', 'Data']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        print(f"[DATA PROCESSING] Warning: Missing columns: {missing_cols}")
        # Try to map common variations
        column_mapping = {
            'Account': ['Account', 'Accounts', 'ACCT'],
            'Period': ['Period', 'Time', 'Month'],
            'Year': ['Year', 'FY'],
            'Scenario': ['Scenario', 'Scenarios', 'SCE'],
            'Data': ['Data', 'Value', 'Amount']
        }
        
        for req_col in missing_cols:
            for alt_col in column_mapping.get(req_col, []):
                if alt_col in df.columns:
                    df = df.rename(columns={alt_col: req_col})
                    break
    
    # Filter and pivot data based on comparison type
    if comparison_type == "actuals_vs_forecast":
        # Filter for Actual and Forecast scenarios
        filtered_df = df[df['Scenario'].isin(['Actual', 'Forecast', 'ACTUAL', 'FORECAST'])]
        
        # Pivot to create comparison columns
        pivot_df = filtered_df.pivot_table(
            index=['Account', 'Period', 'Year'],
            columns='Scenario',
            values='Data',
            aggfunc='sum'
        ).reset_index()
        
        # Calculate variance
        if 'Actual' in pivot_df.columns and 'Forecast' in pivot_df.columns:
            pivot_df['Variance'] = pivot_df['Actual'] - pivot_df['Forecast']
            pivot_df['Variance_%'] = (pivot_df['Variance'] / pivot_df['Forecast'] * 100).round(2)
        
        return pivot_df
    
    elif comparison_type == "forecast_vs_forecast":
        # Compare different forecast versions
        forecast_scenarios = [col for col in df['Scenario'].unique() 
                            if 'forecast' in str(col).lower() or 'fcst' in str(col).lower()]
        
        if len(forecast_scenarios) >= 2:
            filtered_df = df[df['Scenario'].isin(forecast_scenarios)]
            
            pivot_df = filtered_df.pivot_table(
                index=['Account', 'Period', 'Year'],
                columns='Scenario',
                values='Data',
                aggfunc='sum'
            ).reset_index()
            
            # Calculate variance between forecasts
            cols = [col for col in pivot_df.columns if col not in ['Account', 'Period', 'Year']]
            if len(cols) >= 2:
                pivot_df['Variance'] = pivot_df[cols[0]] - pivot_df[cols[1]]
                pivot_df['Variance_%'] = (pivot_df['Variance'] / pivot_df[cols[1]] * 100).round(2)
            
            return pivot_df
    
    return df


# ---------------- VISUALIZATION ----------------
def create_comparison_charts(df: pd.DataFrame, 
                           comparison_type: str = "actuals_vs_forecast",
                           output_file: str = "comparison_report.png") -> str:
    """
    Create bar graphs and charts for data comparisons.
    Returns the path to the generated chart file.
    """
    print(f"[VISUALIZATION] Creating {comparison_type} charts")
    
    if df.empty:
        print("[VISUALIZATION] Warning: No data to visualize")
        return ""
    
    # Set up the plotting style
    plt.style.use('seaborn-v0_8')
    sns.set_palette("husl")
    
    # Create figure with subplots
    fig, axes = plt.subplots(2, 2, figsize=(15, 10))
    fig.suptitle(f'EPM Data Comparison Report - {comparison_type.replace("_", " ").title()}', 
                fontsize=16, fontweight='bold')
    
    try:
        if comparison_type == "actuals_vs_forecast":
            # Chart 1: Actual vs Forecast Bar Chart
            if 'Actual' in df.columns and 'Forecast' in df.columns:
                ax1 = axes[0, 0]
                df_sample = df.head(10)  # Sample first 10 records
                df_sample.set_index('Account')[['Actual', 'Forecast']].plot(
                    kind='bar', ax=ax1, width=0.8
                )
                ax1.set_title('Actual vs Forecast by Account')
                ax1.set_ylabel('Amount')
                ax1.tick_params(axis='x', rotation=45)
                ax1.legend(loc='upper right')
            
            # Chart 2: Variance Analysis
            if 'Variance' in df.columns:
                ax2 = axes[0, 1]
                df['Variance_Color'] = df['Variance'].apply(lambda x: 'red' if x < 0 else 'green')
                df.head(10).plot(kind='bar', x='Account', y='Variance', 
                               ax=ax2, color=df.head(10)['Variance_Color'])
                ax2.set_title('Variance (Actual - Forecast)')
                ax2.set_ylabel('Variance Amount')
                ax2.tick_params(axis='x', rotation=45)
                ax2.axhline(y=0, color='black', linestyle='-', alpha=0.3)
            
            # Chart 3: Variance Percentage
            if 'Variance_%' in df.columns:
                ax3 = axes[1, 0]
                df.head(10).plot(kind='bar', x='Account', y='Variance_%', ax=ax3, color='orange')
                ax3.set_title('Variance Percentage')
                ax3.set_ylabel('Variance %')
                ax3.tick_params(axis='x', rotation=45)
                ax3.axhline(y=0, color='black', linestyle='-', alpha=0.3)
            
            # Chart 4: Trend Analysis (by Period)
            ax4 = axes[1, 1]
            if 'Period' in df.columns:
                period_data = df.groupby('Period')[['Actual', 'Forecast']].sum().reset_index()
                period_data.set_index('Period').plot(kind='line', ax=ax4, marker='o')
                ax4.set_title('Actual vs Forecast Trend by Period')
                ax4.set_ylabel('Amount')
                ax4.tick_params(axis='x', rotation=45)
                ax4.legend(loc='upper left')
        
        elif comparison_type == "forecast_vs_forecast":
            # Similar charts for forecast vs forecast comparison
            forecast_cols = [col for col in df.columns if 'Forecast' in str(col) or 'FCST' in str(col)]
            
            if len(forecast_cols) >= 2:
                ax1 = axes[0, 0]
                df.head(10).set_index('Account')[forecast_cols].plot(
                    kind='bar', ax=ax1, width=0.8
                )
                ax1.set_title('Forecast vs Forecast Comparison')
                ax1.set_ylabel('Amount')
                ax1.tick_params(axis='x', rotation=45)
                ax1.legend(loc='upper right')
                
                # Add other charts as needed...
        
        # Adjust layout and save
        plt.tight_layout()
        plt.savefig(output_file, dpi=300, bbox_inches='tight')
        plt.close()
        
        print(f"[VISUALIZATION] Charts saved to: {output_file}")
        return output_file
        
    except Exception as e:
        print(f"[VISUALIZATION] Error creating charts: {e}")
        plt.close()
        return ""


def generate_comprehensive_report(session: requests.Session, cfg: Dict[str, str],
                                cube_metadata: Dict[str, Any],
                                output_dir: str = "reports") -> dict:
    """
    Generate a comprehensive report with multiple comparison charts.
    """
    print("\n" + "="*60)
    print("GENERATING COMPREHENSIVE EPM COMPARISON REPORT")
    print("="*60 + "\n")
    
    # Create output directory
    import os
    os.makedirs(output_dir, exist_ok=True)
    
    reports_generated = {}
    
    # Define comparison types
    comparison_types = [
        "actuals_vs_forecast",
        "forecast_vs_forecast"
    ]
    
    metadata_map = discover_and_print_dimensions_members(session, cfg)
    app_dimensions = list(metadata_map.keys()) or _extract_dimension_names(cube_metadata)
    if app_dimensions:
        print(f"[REPORT] Using application dimensions: {app_dimensions}")
    else:
        print("[REPORT] Warning: Could not read dimensions from initial metadata; will retry during extraction")

    for comp_type in comparison_types:
        print(f"[REPORT] Generating {comp_type} analysis...")
        
        try:
            # Extract data
            raw_data = extract_epm_data(session, cfg, dimensions=app_dimensions)
            
            if not raw_data.empty:
                # Process data
                processed_data = process_comparison_data(raw_data, comp_type)
                
                if not processed_data.empty:
                    # Create charts
                    chart_file = f"{output_dir}/{comp_type}_report.png"
                    chart_path = create_comparison_charts(processed_data, comp_type, chart_file)
                    
                    if chart_path:
                        reports_generated[comp_type] = {
                            'data': processed_data,
                            'chart_path': chart_path,
                            'records': len(processed_data)
                        }
                        print(f"[REPORT] {comp_type} report generated successfully")
                    else:
                        print(f"[REPORT] Failed to create charts for {comp_type}")
                else:
                    print(f"[REPORT] No processed data for {comp_type}")
            else:
                print(f"[REPORT] No raw data extracted for {comp_type}")
                
        except Exception as e:
            print(f"[REPORT] Error generating {comp_type} report: {e}")
    
    print(f"\n[REPORT] Generated {len(reports_generated)} comparison reports")
    return reports_generated


# ---------------- MAIN EXECUTION ----------------
if __name__ == "__main__":
    try:
        # Initialize connection
        session, cfg, cube_metadata = initialize_epm_connection()
        
        # Generate comprehensive reports
        reports = generate_comprehensive_report(session, cfg, cube_metadata)
        
        if reports:
            print("\n[SUCCESS] All reports generated successfully!")
            for report_type, details in reports.items():
                print(f"  - {report_type}: {details['records']} records, chart saved to {details['chart_path']}")
        else:
            print("\n[WARNING] No reports were generated. Check data extraction and processing.")
            
        sys.exit(0)
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)
