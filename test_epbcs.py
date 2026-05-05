"""Saved copy of test_epbcs.py after fixing authentication issues.

This script verifies EPBCS API authentication and fetches the list of
applications. It uses helper functions from ``epbcs_mcp_basic_auth.py`` for
configuration loading, header construction, request handling, and a quick
authentication check.
"""

from epbcs_mcp_basic_auth import load_config, build_basic_auth_header, request_json, authenticate_to_epbcs

def main() -> None:
    cfg = load_config()
    headers = build_basic_auth_header(cfg["username"], cfg["password"])

    try:
        authenticate_to_epbcs(cfg["base_url"], headers)
    except Exception as exc:
        print(f"Authentication failed: {exc}")
        return

    apps_url = f"{cfg['base_url']}/HyperionPlanning/rest/{cfg['api_version']}/applications"
    try:
        apps = request_json("GET", apps_url, headers=headers)
        print("Applications response:")
        print(apps)
    except Exception as exc:
        print(f"Failed to fetch applications: {exc}")

    epmcloud_url = f"{cfg['base_url']}/epmcloud/HyperionPlanning/rest/{cfg['api_version']}/applications"
    try:
        apps_ec = request_json("GET", epmcloud_url, headers=headers)
        print("\nEPMCloud Applications response:")
        print(apps_ec)
    except Exception as exc:
        print(f"Failed EPmCloud request: {exc}")


if __name__ == "__main__":
    main()