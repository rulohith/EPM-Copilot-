"""Compatibility shim for scripts that import ``update_account_csv``.

The original project placed the core metadata handling logic in ``update_metadata.py``.
``Update_Metadata_With_Prompt.py`` imports that module under the name ``update_account_csv``
and expects several *private* helpers (prefixed with ``_``) as well as constant values.

Python's ``import *`` only re‑exports names that do **not** start with an underscore,
so the required symbols were missing, leading to the runtime error:

``module 'update_account_csv' has no attribute '_load_config'``

This shim explicitly imports the original module as ``_um`` and then forwards the
necessary attributes so that the importing script sees the expected API.
"""

# Import the original implementation module under an alias.
import update_metadata as _um

# Re‑export the private helpers that ``Update_Metadata_With_Prompt.py`` relies on.
_load_config = _um._load_config
_submit_export_job = _um._submit_export_job
_submit_import_job = _um._submit_import_job
_norm_key = _um._norm_key

# Re‑export constant definitions used by the prompt script.
ACCOUNT_ZIP_NAME = _um.ACCOUNT_ZIP_NAME
LOB_ZIP_NAME = _um.LOB_ZIP_NAME
ACCOUNT_CSV_NAME = _um.ACCOUNT_CSV_NAME
LOB_CSV_NAME = _um.LOB_CSV_NAME
# ``update_metadata`` defines the workbook filenames as ``EXCEL_FILE_NAME`` (the
# account workbook) and ``LOB_EXCEL_FILE_NAME`` (the LOB workbook).  The prompt
# script accesses them via ``ACCOUNT_EXCEL_NAME`` and ``LOB_EXCEL_NAME``.  We expose
# both the original names and the expected aliases for compatibility.
EXCEL_FILE_NAME = _um.EXCEL_FILE_NAME
LOB_EXCEL_FILE_NAME = _um.LOB_EXCEL_FILE_NAME
ACCOUNT_EXCEL_NAME = getattr(_um, "ACCOUNT_EXCEL_NAME", _um.EXCEL_FILE_NAME)
LOB_EXCEL_NAME = getattr(_um, "LOB_EXCEL_NAME", _um.LOB_EXCEL_FILE_NAME)
ACCOUNT_COLUMNS_TO_REMOVE = _um.ACCOUNT_COLUMNS_TO_REMOVE
LOB_COLUMNS_TO_REMOVE = _um.LOB_COLUMNS_TO_REMOVE

# Any other symbols that might be accessed via ``epm`` can be added here in the
# future.  Keeping the shim explicit makes the dependency clear and avoids the
# pitfalls of ``import *`` with private names.
