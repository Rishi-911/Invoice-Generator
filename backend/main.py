from word_generator import replace_placeholders
import config
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from starlette.background import BackgroundTask
from fastapi.responses import FileResponse
import zipfile
import shutil
from fastapi.middleware.cors import CORSMiddleware
import os
import uuid
import logging
from pathlib import Path

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("invoice_generator")

app = FastAPI()

tasks = {}

def cleanup(req_dir: Path):
    try:
        if req_dir.exists() and req_dir.is_dir():
            if "temporary" in req_dir.parts and req_dir != config.BASE_DIR:
                shutil.rmtree(req_dir)
                logger.info(f"Successfully cleaned up temporary request directory: {req_dir}")
            else:
                logger.warning(f"Cleanup safety check failed for path: {req_dir}")
        else:
            logger.info(f"Cleanup skipped: directory does not exist: {req_dir}")
    except Exception as e:
        logger.error(f"Cleanup error for directory {req_dir}: {e}")

def cleanup_task(task_id: str, req_dir: Path):
    cleanup(req_dir)
    if task_id in tasks:
        del tasks[task_id]

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "https://invoice-generator-psi-ashy.vercel.app"
    ],
    allow_credentials=True,
    allow_methods=["GET", "POST"],
    allow_headers=["*"]
)

def generate_invoices_background(
    task_id: str,
    template_path: Path,
    data_path: Path,
    req_dir: Path,
    docx_out: Path,
    pdf_out: Path
):
    try:
        config._local.docx_output_folder = docx_out
        config._local.pdf_output_folder = pdf_out

        pdf_files, docx_files = replace_placeholders(template_path, data_path)

        zip_path = req_dir / "invoices.zip"

        with zipfile.ZipFile(
                zip_path,
                "w",
                zipfile.ZIP_DEFLATED
        ) as zip_file:
                for pdf_path in pdf_files:
                        zip_file.write(
                                pdf_path,
                                arcname=pdf_path.name
                        )

        tasks[task_id] = {
            "status": "completed",
            "zip_path": zip_path,
            "req_dir": req_dir,
            "error": None
        }
        logger.info(f"Background task {task_id} completed successfully.")
    except Exception as e:
        logger.error(f"Error in background task {task_id}: {e}", exc_info=True)
        tasks[task_id] = {
            "status": "failed",
            "zip_path": None,
            "req_dir": req_dir,
            "error": str(e)
        }
        cleanup(req_dir)

@app.post("/api/generate")
def generate_invoices(
        background_tasks: BackgroundTasks,
        template: UploadFile = File(...),
        data: UploadFile = File(...)
):      
        logger.info("Received request to generate invoices.")
        
        request_id = uuid.uuid4().hex

        req_dir = config.BASE_DIR / "temporary" / request_id
        docx_out = req_dir / "output" / "docx"
        pdf_out = req_dir / "output" / "pdf"

        req_dir.mkdir(parents=True, exist_ok=True)
        docx_out.mkdir(parents=True, exist_ok=True)
        pdf_out.mkdir(parents=True, exist_ok=True)

        template_ext = Path(template.filename).suffix if template.filename else ".docx"
        data_ext = Path(data.filename).suffix if data.filename else ".xlsx"

        template_path = req_dir / f"template{template_ext}"
        data_path = req_dir / f"data{data_ext}"

        with open(template_path, "wb") as file:
                shutil.copyfileobj(template.file, file)

        with open(data_path, "wb") as file:
                shutil.copyfileobj(data.file, file)

        logger.info("Successfully copied uploaded files to disk. Starting background worker.")

        tasks[request_id] = {
            "status": "processing",
            "zip_path": None,
            "req_dir": req_dir,
            "error": None
        }

        background_tasks.add_task(
            generate_invoices_background,
            request_id,
            template_path,
            data_path,
            req_dir,
            docx_out,
            pdf_out
        )

        return {
            "task_id": request_id,
            "status": "processing"
        }

@app.get("/api/status/{task_id}")
def get_task_status(task_id: str):
    if task_id not in tasks:
        raise HTTPException(status_code=404, detail="Task not found")
    return {
        "status": tasks[task_id]["status"],
        "error": tasks[task_id]["error"]
    }

@app.get("/api/download/{task_id}")
def download_invoices(task_id: str):
    if task_id not in tasks:
        raise HTTPException(status_code=404, detail="Task not found")
    
    task = tasks[task_id]
    if task["status"] != "completed":
        raise HTTPException(status_code=400, detail=f"Task is in status: {task['status']}")
    
    zip_path = task["zip_path"]
    req_dir = task["req_dir"]

    if not zip_path or not zip_path.exists():
        raise HTTPException(status_code=404, detail="Generated ZIP file not found")

    return FileResponse(
        path=zip_path,
        media_type="application/zip",
        filename="invoices.zip",
        background=BackgroundTask(
            cleanup_task,
            task_id,
            req_dir
        )
    )

