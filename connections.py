from __future__ import annotations

import base64
import os
from pathlib import Path
from typing import Dict

import requests
import urllib3
from dotenv import load_dotenv

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


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
        "plantype": os.getenv("EPM_PLANTYPE", "").strip(),
    }

    required = ["base_url", "username", "password", "api_version", "application"]
    missing = [k for k in required if not cfg.get(k)]
    if missing:
        raise ValueError(f"Missing required environment variables: {', '.join(missing)}")

    return cfg


def _normalize_base_url(base_url: str) -> str:
    """Normalize .env EPM URL to the tenant root used by Planning/Interop REST."""
    base = (base_url or "").strip().rstrip("/")
    for marker in ("/epmcloud", "/HyperionPlanning", "/interop/rest", "/rest"):
        if marker in base:
            base = base.split(marker, 1)[0]
            break
    return base.rstrip("/")


def _base_url_candidates(base_url: str) -> list[str]:
    base = _normalize_base_url(base_url)
    candidates = [base]

    unique: list[str] = []
    seen = set()
    for c in candidates:
        if c and c not in seen:
            unique.append(c)
            seen.add(c)
    return unique


def connect_to_epm_environment() -> tuple[requests.Session, Dict[str, str]]:
    """
    Establishes a connection to the EPM environment using .env values.
    Returns: (authenticated requests.Session, config dict)
    """
    cfg = _load_config()
    headers = _basic_auth_header(cfg["username"], cfg["password"])
    session = requests.Session()

    possible_urls: list[str] = []
    for root in _base_url_candidates(cfg["base_url"]):
        possible_urls.extend([
            f"{root}/HyperionPlanning/rest/{cfg['api_version']}/applications/{cfg['application']}",
            f"{root}/HyperionPlanning/rest/{cfg['api_version']}/applications",
            f"{root}/rest/{cfg['api_version']}/applications/{cfg['application']}",
            f"{root}/rest/{cfg['api_version']}/applications",
            f"{root}/interop/rest",
        ])

    def _try_connect(sess: requests.Session) -> tuple[requests.Session, bool]:
        for test_url in possible_urls:
            try:
                resp = sess.get(test_url, headers=headers, timeout=30, verify=False)
                if resp.status_code == 200:
                    # If response is HTML login page, keep trying.
                    sample = (resp.text or "")[:200].lower()
                    if "<html" in sample or "login" in sample:
                        continue
                    return sess, True
                if resp.status_code == 401:
                    raise RuntimeError("Authentication failed (401). Check credentials.")
            except Exception:
                continue
        return sess, False

    session, ok = _try_connect(session)
    if ok:
        return session, cfg

    # Fallback: ignore environment proxy vars that may block direct EPM calls.
    # This helps when shell/system proxy settings are stale or invalid.
    no_env_session = requests.Session()
    no_env_session.trust_env = False
    no_env_session, ok = _try_connect(no_env_session)
    if ok:
        return no_env_session, cfg

    raise RuntimeError(
        "Could not connect to EPM instance. Verify EPM_BASE_URL and EPM_APPLICATION in .env."
    )


if __name__ == "__main__":
    sess, config = connect_to_epm_environment()
    print("Connected to EPM environment successfully.")
    print(f"Base URL: {config['base_url']}")
    print(f"Application: {config['application']}")
