# run_incMetaJob Skill

## First Instruction (Mandatory)
- Read `Human_Validation_Instruction.md` before any other step in every run, including a new session.
- At the validation prompt, pause and wait for explicit human `Yes/No` input.
- Never auto-approve by entering `Yes` or `No`.
- If interactive input is unavailable, stop the run and keep the generated validation note.

## Goal
Run incremental metadata synchronization for EPBCS using:
- source CSV file name (looked up from EPBCS Inbox/Outbox, then downloaded to local `.\Files`)
- column-to-dimension mapping (`FileColtoDim.csv`)
- fixed metadata template CSV path (`.\Files\Metadata_Template.csv`)
- EPBCS REST APIs (no EPM Automate)

The script entry point is:

```python
pbcs_run_incmetadata_sync(
    source_csv_path=...
)
```

File: `IncrementalMetadata_Sync.py`

## Required Environment
Set these in `.env`:

- `EPBCS_BASE_URL`
- `EPBCS_APP_NAME`
- `EPBCS_USERNAME`
- `EPBCS_PASSWORD` or `EPBCS_PASSWORD_FILE`
- `EPBCS_VERIFY_SSL` (optional, default `true`)
- `EPBCS_REQUEST_TIMEOUT_SEC` (optional, default `30`)
- `EPBCS_POLL_INTERVAL_SECONDS` (optional, default `2`)
- `EPBCS_POLL_TIMEOUT_SECONDS` (optional, default `600`)
- `EPBCS_CUBE_NAME` (optional, default `OEP_FS`)

Per-dimension import job mapping is mandatory when new members are detected:

- `INCIMPORT_<DimensionName>="Import Metadata Job Name"`
- `IMPJOB_FILE_<DimensionName>="EPMLoadable_<DimensionName>.csv"`

Examples:
- `INCIMPORT_Future1="ImportMetadata_Future1"`
- `INCIMPORT_CostCentre="ImportMetadata_CostCentre"`
- `IMPJOB_FILE_Future1="EPMLoadable_Future1.csv"`
- `IMPJOB_FILE_CostCentre="EPMLoadable_CostCentre.csv"`

## Process Steps

### 1. Read Column-to-Dimension Mapping
- Use `FileColtoDim.csv` to map:
  - column number (1-based)
  - corresponding EPBCS dimension
- If mapping file is missing in default paths, script prompts for full file path.
- Only mapped columns are processed.

### 2. Detect New Members and Build Load Files
For each mapped column/dimension:
- Read unique values from source CSV column.
- Fetch current dimension hierarchy from EPBCS.
- Compare source unique values against existing members.
- If no new member: skip dimension.
- If new members exist:
  - Create `EPMLoadable_<Dimension>.csv` using `Metadata_Template.csv`.
  - Try closest-name match against existing members.
  - If closest match confidence passes threshold: assign parent as immediate parent of closest match.
  - Else:
    - Ensure fallback node `TobeAdded_<Dimension>` exists/created under dimension root.
    - Add new members under fallback node.
  - Zip to `EPMLoadable_<Dimension>.zip`.
  - Delete original `EPMLoadable_<Dimension>.csv`.
After all ZIPs are generated for the run:
- Pause once for consolidated human validation (single prompt per run, not per dimension).
- Script generates one consolidated validation note in working folder:
  - `Validation required_<YYYYMMDD_HHMMSS>.txt`
  - Includes all new-member -> parent mappings from all generated ZIPs in that run.
- Waits for explicit `Yes/No` input (blocking, no timeout).
- `Yes`: continue to upload/jobs for all dimensions in scope.
- `No`: stop run.
- Non-interactive run (no stdin): stop run with `VALIDATION_INPUT_UNAVAILABLE` and keep the consolidated validation note for review.

### 3. Upload and Run Import Metadata Job
For each dimension with new members:
- Resolve import job from `INCIMPORT_<Dimension>`.
- If missing: raise exception and stop run.
- Resolve load file name from `IMPJOB_FILE_<Dimension>`.
- If missing: raise exception and stop run.
- Upload generated zip to Interop via REST API.
  - If same file name already exists in Interop, script auto-deletes old file and overwrites with new upload.
  - First-time/new-dimension file delete responses like `Invalid file` are treated as missing-file and do not block upload.
- Execute import metadata job via Planning REST API.
- Poll job status until completion/failure/timeout.

### 4. Iterate Through All Mapped Columns
- Continue same logic for each mapped column.
- Do not scan unmapped columns.
- Return JSON summary with per-dimension outcome.
- At run start, script auto-deletes old `Validation required_*.txt` files older than 7 days.

## Run Examples

### Python call
```python
from IncrementalMetadata_Sync import pbcs_run_incmetadata_sync

result = pbcs_run_incmetadata_sync(
    source_csv_path="Test_DataFile_Feb.csv"
)
print(result)
```

### CLI call
```powershell
.\.venv\Scripts\python.exe .\IncrementalMetadata_Sync.py `
  --source Test_DataFile_Feb.csv
```

Human approval is mandatory for every run; there is no skip flag.

### Dry run (local file generation only)
```powershell
.\.venv\Scripts\python.exe .\IncrementalMetadata_Sync.py `
  --source Test_DataFile_Feb.csv `
  --dry-run
```
