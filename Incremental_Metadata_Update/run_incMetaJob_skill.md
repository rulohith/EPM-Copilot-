# run_incMetaJob Skill

## Goal
Run incremental metadata synchronization for EPBCS using:
- source data CSV
- column-to-dimension mapping (`FileColtoDim.csv`)
- metadata template CSV (`Metadata_Template.csv`)
- EPBCS REST APIs (no EPM Automate)

The script entry point is:

```python
pbcs_run_incmetadata_sync(
    source_csv_path=...,
    template_csv_path=...,
    cube_name=...
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

Per-dimension import job mapping is mandatory when new members are detected:

- `INCIMPORT_<DimensionName>="Import Metadata Job Name"`

Examples:
- `INCIMPORT_Future1="ImportMetadata_Future1"`
- `INCIMPORT_CostCentre="ImportMetadata_CostCentre"`

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

### 3. Upload and Run Import Metadata Job
For each dimension with new members:
- Resolve import job from `INCIMPORT_<Dimension>`.
- If missing: raise exception and stop run.
- Upload generated zip to Interop via REST API.
- Execute import metadata job via Planning REST API.
- Poll job status until completion/failure/timeout.

### 4. Iterate Through All Mapped Columns
- Continue same logic for each mapped column.
- Do not scan unmapped columns.
- Return JSON summary with per-dimension outcome.

## Run Examples

### Python call
```python
from IncrementalMetadata_Sync import pbcs_run_incmetadata_sync

result = pbcs_run_incmetadata_sync(
    source_csv_path=r".\Files\Test_DataFile_Feb.csv",
    template_csv_path=r".\Files\Metadata_Template.csv",
    cube_name="CSNPLAN"
)
print(result)
```

### CLI call
```powershell
.\.venv\Scripts\python.exe .\IncrementalMetadata_Sync.py `
  --source .\Files\Test_DataFile_Feb.csv `
  --template .\Files\Metadata_Template.csv `
  --cube CSNPLAN
```

### Dry run (local file generation only)
```powershell
.\.venv\Scripts\python.exe .\IncrementalMetadata_Sync.py `
  --source .\Files\Test_DataFile_Feb.csv `
  --template .\Files\Metadata_Template.csv `
  --cube CSNPLAN `
  --dry-run
```
