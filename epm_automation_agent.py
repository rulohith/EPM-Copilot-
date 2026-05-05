from __future__ import annotations

import subprocess
import sys
import re
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent


def _run_command(args: list[str]) -> int:
    print(f"\nRunning: {' '.join(args)}")
    completed = subprocess.run(args, cwd=str(PROJECT_DIR))
    return completed.returncode


def _run_command_capture(args: list[str]) -> tuple[int, str]:
    print(f"\nRunning: {' '.join(args)}")
    completed = subprocess.run(
        args,
        cwd=str(PROJECT_DIR),
        text=True,
        capture_output=True,
    )
    output = (completed.stdout or "") + (completed.stderr or "")
    if output.strip():
        print(output)
    return completed.returncode, output


def _contains_any(text: str, keywords: list[str]) -> bool:
    return any(k in text for k in keywords)


def _handle_update_metadata(user_text: str) -> None:
    txt = user_text.lower().strip()

    if _contains_any(
        txt,
        [
            "i have updated the metadata files",
            "updated the metadata files",
            "update the metadata files",
            "workbook mode",
            "run update metadata",
            "made a lot of changes",
            "made many changes",
            "i have made a lot of changes",
        ],
    ):
        print("\nGot it - running workbook metadata update flow.")
        cmd = [sys.executable, str(PROJECT_DIR / "update_metadata.py")]
        rc, _ = _run_command_capture(cmd)
    else:
        print("\nGot it - running metadata update with prompt.")
        print(f"Prompt: {user_text.strip()}")
        cmd = [
            sys.executable,
            str(PROJECT_DIR / "run_update_prompt.py"),
            "--prompt",
            user_text.strip(),
        ]
        rc, output = _run_command_capture(cmd)

        if _contains_any(output.lower(), ["already exists", "exact match already exists", "member already exists"]):
            print("Member already exists. Stopping metadata update.")
            return

    if rc == 0:
        print("Metadata update flow completed successfully.")
    else:
        print("Metadata update failed. Please review the logs above.")


def _handle_load_actuals(user_text: str) -> None:
    print("\nGreat - let's run data load.")
    dataset = "forecast" if "forecast" in user_text else "actuals"

    print("Provide either an Entity name (to auto-pick '<ENTITY> Actuals/Forecast.csv')")
    print("or a full CSV file path.")
    entity = input("Entity (optional): ").strip()
    csv_path = input("CSV file path (optional): ").strip()
    load_option = input("Load option [Add/Overwrite] (default Overwrite): ").strip() or "Overwrite"

    cmd = [
        sys.executable,
        str(PROJECT_DIR / "Data_Load_With_Prompt.py"),
        "--dataset",
        dataset,
        "--load-option",
        load_option,
    ]

    if csv_path:
        cmd.extend(["--file", csv_path])
    elif entity:
        cmd.extend(["--entity", entity])

    rc = _run_command(cmd)
    if rc == 0:
        print("Data load flow completed successfully.")
    else:
        print("Data load failed. Please review the logs above.")


def _extract_entity(user_text: str) -> str:
    # Supports phrases like "load actuals for LE_ANZ" / "validate forecast entity LE_ANZ"
    m = re.search(r"\b(?:for|entity)\s+([a-zA-Z0-9_\-]+)", user_text, flags=re.IGNORECASE)
    return m.group(1) if m else ""


def _handle_validate_data(user_text: str) -> None:
    dataset = "Forecast" if "forecast" in user_text else "Actuals"
    entity = _extract_entity(user_text) or "LE_ANZ"
    print(f"\nRunning validation for {dataset} (entity: {entity})...")

    cmd = [
        sys.executable,
        str(PROJECT_DIR / "Data_Validation_With_Prompt.py"),
        "--entity",
        entity,
        "--dataset",
        dataset,
    ]

    rc = _run_command(cmd)
    if rc == 0:
        print("Data validation completed successfully.")
    else:
        print("Data validation failed. Please review the logs above.")


def _handle_generate_insights(_: str) -> None:
    print("\nLaunching Streamlit insights dashboard...")
    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(PROJECT_DIR / "streamlit_epm_dashboard.py"),
    ]
    print("Press Ctrl+C in this terminal when you want to stop Streamlit.")
    _run_command(cmd)


def _route_command(user_text: str) -> bool:
    txt = user_text.lower().strip()
    if not txt:
        return True

    if txt in {"exit", "quit", "bye"}:
        print("Thanks! Your EPM automation agent is signing off.")
        return False

    if _contains_any(
        txt,
        [
            "update metadata",
            "metadata update",
            "update master",
            "metadata",
            "add specific member",
            "add member",
            "add under",
        ],
    ):
        _handle_update_metadata(txt)
        return True

    if _contains_any(txt, ["load actual", "load actuals", "data load", "load data", "load forecast"]):
        _handle_load_actuals(txt)
        return True

    if _contains_any(txt, ["validate actual", "validate actuals", "validate forecast", "validate data"]):
        _handle_validate_data(txt)
        return True

    if _contains_any(txt, ["generate insights", "insights", "dashboard", "streamlit"]):
        _handle_generate_insights(txt)
        return True

    print("\nI can help with:")
    print("   • Update metadata")
    print("   • Load actuals / forecast data")
    print("   • Validate actuals / forecast data")
    print("   • Generate insights (Streamlit dashboard)")
    return True


def main() -> None:
    print("Hi! I am your EPM automation agent - ready to automate metadata updates, data loads, and insight generation. How can I help you today?")
    while True:
        try:
            user_text = input("\nYou: ")
        except (EOFError, KeyboardInterrupt):
            print("\nSession ended.")
            break
        if not _route_command(user_text):
            break


if __name__ == "__main__":
    main()
