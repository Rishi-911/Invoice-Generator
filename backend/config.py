import threading
import os
from pathlib import Path

_local = threading.local()

class DynamicPathProxy:
    """
    A path proxy that dynamically evaluates the actual path object on demand.
    This allows thread-safe overriding of output folders for multi-user isolation.
    """
    def __init__(self, get_path_fn):
        self._get_path_fn = get_path_fn

    def __truediv__(self, other):
        return self._get_path_fn() / other

    def __str__(self):
        return str(self._get_path_fn())

    def __repr__(self):
        return repr(self._get_path_fn())

    def __fspath__(self):
        return os.fspath(self._get_path_fn())

    def __getattr__(self, name):
        return getattr(self._get_path_fn(), name)

BASE_DIR = Path(__file__).parent
DATA_FOLDER = BASE_DIR / "data"
TEMPLATES_FOLDER = BASE_DIR / "template"
OUTPUT_FOLDER = BASE_DIR / "output"

DOCX_OUTPUT_FOLDER_DEFAULT = OUTPUT_FOLDER / "docx"
PDF_OUTPUT_FOLDER_DEFAULT = OUTPUT_FOLDER / "pdf"

DOCX_OUTPUT_FOLDER = DynamicPathProxy(lambda: getattr(_local, "docx_output_folder", DOCX_OUTPUT_FOLDER_DEFAULT))
PDF_OUTPUT_FOLDER = DynamicPathProxy(lambda: getattr(_local, "pdf_output_folder", PDF_OUTPUT_FOLDER_DEFAULT))

