import threading
import os
from pathlib import Path

# Thread-local storage for request-specific paths
_local = threading.local()

class DynamicPathProxy:
    """
    A path proxy that dynamically evaluates the actual path object on demand.
    This allows thread-safe overriding of output folders for multi-user isolation.
    """
    def __init__(self, get_path_fn):
        self._get_path_fn = get_path_fn

    def __truediv__(self, other):
        # Allow path concatenation like path / 'subdir'
        return self._get_path_fn() / other

    def __str__(self):
        return str(self._get_path_fn())

    def __repr__(self):
        return repr(self._get_path_fn())

    def __fspath__(self):
        # Support os.fspath() and open() functions expecting os.PathLike
        return os.fspath(self._get_path_fn())

    def __getattr__(self, name):
        # Delegate all other pathlib.Path methods (e.g. resolve, mkdir, exists)
        return getattr(self._get_path_fn(), name)

BASE_DIR = Path(__file__).parent
DATA_FOLDER = BASE_DIR / "data"
TEMPLATES_FOLDER = BASE_DIR / "template"
OUTPUT_FOLDER = BASE_DIR / "output"

# Default directories for standalone run or fallback
DOCX_OUTPUT_FOLDER_DEFAULT = OUTPUT_FOLDER / "docx"
PDF_OUTPUT_FOLDER_DEFAULT = OUTPUT_FOLDER / "pdf"

# Dynamic proxies pointing to thread-local paths if configured, otherwise fallback to defaults
DOCX_OUTPUT_FOLDER = DynamicPathProxy(lambda: getattr(_local, "docx_output_folder", DOCX_OUTPUT_FOLDER_DEFAULT))
PDF_OUTPUT_FOLDER = DynamicPathProxy(lambda: getattr(_local, "pdf_output_folder", PDF_OUTPUT_FOLDER_DEFAULT))

