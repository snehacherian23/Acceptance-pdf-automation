"""
app/utils/job_tracker.py
---------------------------
A minimal in-memory store for tracking processing jobs and their logs.

Deliberately simple (a plain dict, not a database table) for now --
this is enough to support "upload, process, see results, download
ZIP" within a single running server process. If this needs to survive
server restarts or scale across multiple worker processes later, this
is the seam where we'd swap in a database-backed or Redis-backed
implementation without touching the route logic.

Also tracks live progress (current file, processed/total count, start
time) so the dashboard can poll /processing/status/<job_id> and show
real-time feedback while a background thread runs the pipeline.
"""

import uuid
import time
from datetime import datetime, timezone

_jobs = {}


def create_job():
    job_id = uuid.uuid4().hex[:12]
    _jobs[job_id] = {
        "id": job_id,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "status": "pending",       # pending -> processing -> completed
        "logs": [],
        "records": [],             # one entry per PDF -- feeds the Excel report
        "stats": {
            "total": 0,
            "success": 0,
            "failed": 0,
        },
        "zip_path": None,
        "zip_download_name": "documents.zip",
        "report_path": None,
        # --- live progress (polled by the dashboard while processing) ---
        "total_files": 0,
        "processed_count": 0,
        "current_file": None,
        "started_at": None,        # monotonic seconds, for duration math
        "finished_at": None,
        "duration_seconds": None,
    }
    return job_id


def get_job(job_id):
    return _jobs.get(job_id)


def log(job_id, message, level="info"):
    job = _jobs.get(job_id)
    if job is not None:
        job["logs"].append({
            "message": message,
            "level": level,  # info | success | warning | error
            "time": datetime.now(timezone.utc).isoformat(),
        })


def add_record(job_id, **fields):
    """
    Record one outcome row, later exported as processing_report.xlsx.
    Called for EVERY input item (Excel row, or signed document) --
    success or failure -- so the report always has one row per input.

    Generalized to accept arbitrary fields (**fields) rather than a
    fixed signature, since Workflow 1 (Excel rows: customer_name,
    account_number, property_id...) and Workflow 2 (signed documents:
    input_file, consolidated_count...) have genuinely different data
    shapes. report_service.generate_report is told which columns to
    render via its own `columns` argument -- this function just stores
    whatever it's given.
    """
    job = _jobs.get(job_id)
    if job is not None:
        job["records"].append({
            **fields,
            "processed_at": datetime.now(timezone.utc).isoformat(),
        })


def set_status(job_id, status):
    job = _jobs.get(job_id)
    if job is not None:
        job["status"] = status


def set_zip_path(job_id, zip_path, download_name=None):
    job = _jobs.get(job_id)
    if job is not None:
        job["zip_path"] = zip_path
        if download_name:
            job["zip_download_name"] = download_name


def set_report_path(job_id, report_path):
    job = _jobs.get(job_id)
    if job is not None:
        job["report_path"] = report_path


def increment_stat(job_id, key):
    job = _jobs.get(job_id)
    if job is not None and key in job["stats"]:
        job["stats"][key] += 1


# --- Live progress helpers -------------------------------------------------
def start_timer(job_id, total_files):
    job = _jobs.get(job_id)
    if job is not None:
        job["total_files"] = total_files
        job["started_at"] = time.monotonic()


def set_current_file(job_id, filename):
    job = _jobs.get(job_id)
    if job is not None:
        job["current_file"] = filename


def increment_processed(job_id):
    job = _jobs.get(job_id)
    if job is not None:
        job["processed_count"] += 1


def finish_timer(job_id):
    job = _jobs.get(job_id)
    if job is not None and job["started_at"] is not None:
        job["finished_at"] = time.monotonic()
        job["duration_seconds"] = round(job["finished_at"] - job["started_at"], 1)
        job["current_file"] = None
