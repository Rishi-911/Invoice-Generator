from word_generator import replace_placeholders
import config
from fastapi import FastAPI, UploadFile, File
from starlette.background import BackgroundTask
from fastapi.responses import FileResponse
import zipfile
import shutil
from fastapi.middleware.cors import CORSMiddleware
import os
import uuid
import logging
from pathlib import Path

# Setup structured logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("invoice_generator")

app = FastAPI()

def cleanup(req_dir: Path):
    """
    Cleans up request-specific temporary storage safely.
    
    Why cleanup only removes the request-specific directory:
    - Isolates deletion so that only files related to this specific request are removed.
    - Does not use broad operations, ensuring base/permanent files and other users' active files are never deleted.
    """
    try:
        if req_dir.exists() and req_dir.is_dir():
            # Safety check: Ensure we only delete folders inside the "temporary" directory to prevent accidental base folder deletion
            if "temporary" in req_dir.parts and req_dir != config.BASE_DIR:
                shutil.rmtree(req_dir)
                logger.info(f"Successfully cleaned up temporary request directory: {req_dir}")
            else:
                logger.warning(f"Cleanup safety check failed for path: {req_dir}")
        else:
            logger.info(f"Cleanup skipped: directory does not exist: {req_dir}")
    except Exception as e:
        logger.error(f"Cleanup error for directory {req_dir}: {e}")

# Keep CORS functionality completely unchanged
app.add_middleware(
       CORSMiddleware,
       allow_origins = ["http://localhost:5173","http://127.0.0.1:5173"],    
       allow_credentials = False,
       allow_methods = ["*"],
       allow_headers = ["*"]
)

@app.post("/api/generate")
def generate_invoices(
        template: UploadFile = File(...),
        data: UploadFile = File(...)
):      
        # Generate a unique request ID to isolate this request's context
        request_id = uuid.uuid4().hex

        # Why each request gets a unique directory:
        # - Prevents name collisions when multiple users run the application simultaneously.
        # - Isolates files (inputs, outputs, and the final ZIP) per user session.
        req_dir = config.BASE_DIR / "temporary" / request_id
        docx_out = req_dir / "output" / "docx"
        pdf_out = req_dir / "output" / "pdf"

        # Dynamically set the thread-local variables in config.py.
        # Since this FastAPI endpoint is synchronous, it runs in a thread pool executor.
        # Using thread-local storage isolates the output paths for concurrent request threads.
        config._local.docx_output_folder = docx_out
        config._local.pdf_output_folder = pdf_out

        # Ensure all required temporary request directories are created before saving uploads.
        # Parents=True ensures "temporary" and the request-specific subdirectories are created.
        req_dir.mkdir(parents=True, exist_ok=True)
        docx_out.mkdir(parents=True, exist_ok=True)
        pdf_out.mkdir(parents=True, exist_ok=True)

        # Why permanent template folders are not used for uploaded files:
        # - Prevents unauthorized overwriting of permanent template assets.
        # - Prevents deletion of permanent templates during background task cleanup.
        # - Keeps the filesystem secure and tidy.
        
        # Handle filenames safely by ignoring potentially malicious user-provided paths.
        # We extract only the extension and save the files with fixed safe names inside the unique request folder.
        template_ext = Path(template.filename).suffix if template.filename else ".docx"
        data_ext = Path(data.filename).suffix if data.filename else ".xlsx"

        template_path = req_dir / f"template{template_ext}"
        data_path = req_dir / f"data{data_ext}"

        # Write uploaded template file to the temporary request directory
        with open(template_path, "wb") as file:
                shutil.copyfileobj(
                        template.file,
                        file    
                )

        # Write uploaded data file to the temporary request directory
        with open(data_path, "wb") as file:
                shutil.copyfileobj(
                        data.file,
                        file    
                )

        # Run invoice generation using the unique local paths
        pdf_files, docx_files = replace_placeholders(template_path, data_path)

        # Why the ZIP must be unique:
        # - Allows concurrent users to download their respective zip archives without overwriting each other.
        # - Placing it inside the request-specific directory isolates the archive per request.
        zip_path = req_dir / "invoices.zip"

        with zipfile.ZipFile(
                zip_path,
                "w",
                zipfile.ZIP_DEFLATED
        ) as zip_file:
                for pdf_path in pdf_files:
                        zip_file.write(
                                pdf_path,
                                arcname = pdf_path.name
                        )

        # Return the ZIP file. The BackgroundTask will trigger cleanup of the request directory once delivery completes.
        return FileResponse(
                path = zip_path,
                media_type = "application/zip",
                filename = "invoices.zip",
                background = BackgroundTask(
                        cleanup,
                        req_dir
                )
        )
