## Load Actuals SOP

1. When the user says **"Load Actuals"**, ask them to **Specify Entity**.
2. Look for a CSV named `{entity} Actuals.csv` in the same folder as `Data_Load_With_Prompt.py`.
3. If the file is missing (no matches), respond with **"File Not Found"**.
4. If the file exists, run `Data_Load_With_Prompt.py` with `--dataset actuals` and `--load-option Overwrite`. For example:

```bash
python Data_Load_With_Prompt.py --entity "LE_ANZ" --dataset actuals --load-option Overwrite
```

This script reads credentials from `.env`, uploads the CSV, and triggers the Groovy rule defined in the env vars. Use `--file <path>` to bypass entity lookup, or `--no-poll` to skip job polling if needed.

## Load Forecast SOP

1. When the user says **"Load Forecast"**, ask them to **Specify Entity**.
2. Look for a CSV named `{entity} Forecast.csv` in the same folder as `Data_Load_With_Prompt.py`.
3. If the file is missing (no matches), respond with **"File Not Found"**.
4. If the file exists, run `Data_Load_With_Prompt.py` with `--dataset forecast` and `--load-option Overwrite`. For example:

```bash
python Data_Load_With_Prompt.py --entity "LE_ANZ" --dataset forecast --load-option Overwrite
```

This script reads credentials from `.env`, uploads the CSV, and triggers the Groovy rule defined in the env vars. Use `--file <path>` to bypass entity lookup, or `--no-poll` to skip job polling if needed.

## Validate SOP (Actuals & Forecast)

1. Ask which dataset the user wants to validate (**"Actuals"** or **"Forecast"**). The script expects both the CSV (`{entity} Actuals.csv` or `{entity} Forecast.csv`) and the latest `Export Account*.zip` in the working folder.
2. Run the helper with the appropriate dataset flag:

```bash
# Actuals (default)
python Data_Validation_With_Prompt.py --entity "LE_ANZ"

# Forecast
python Data_Validation_With_Prompt.py --entity "LE_ANZ" --dataset Forecast
```

3. The script will:
   - Open the latest `Export Account*.zip`, find the `{something}Measures.csv`, and compare every 9th column value from the target CSV (excluding the header) against the `Measures` column.
   - Build `validation_details` according to the mismatch rules and generate a Tahoma-styled Word report (`Validation_Report.docx`).
   - Upload the report to Planning and call the Send Mail REST API. The console prints the job status link (e.g., `/status/jobs/<id>`); use that if you need to confirm delivery.
4. If email delivery fails, copy the `validation_details` from the console output and share it manually, then investigate the Interop mail job via the provided status link.