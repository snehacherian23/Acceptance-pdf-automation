"""
app/routes/processing_routes.py
----------------------------------
The orchestration layer for BOTH business workflows. Each has its own
pipeline function and its own /run endpoint, but they share the same
job-tracking, reporting, and zipping infrastructure (job_tracker,
report_service, zip_service) -- and the same /status, /download, and
/download-report endpoints, since those are workflow-agnostic (they
just serve whatever job_tracker has stored for that job_id).

------------------------------------------------------------------------
WORKFLOW 1 -- Generate Acceptance Test Documents
------------------------------------------------------------------------
Input:  one or more Excel sheets.
  1. Read every row across all sheets (excel_service.read_all_rows),
     auto-detecting the real header row and mapping known columns to
     Acceptance Test fields.
  2. For each row: generate one Acceptance Test .docx (docx_template_service).
     Blank optional fields are left blank, never a failure.
  3. Generate processing_report.xlsx, zip the documents + report.
No Consolidated Sheets are used in this workflow.

------------------------------------------------------------------------
WORKFLOW 2 -- Prepare Final Signed Documents
------------------------------------------------------------------------
Input:  one or more signed Acceptance Test documents (.docx or .pdf),
        and one or more Consolidated Sheet files (.docx and/or .pdf).
  1. Convert every signed .docx to PDF (pdf_conversion_service, via
     LibreOffice headless -- batched in one call per file type for speed).
     Signed documents already in .pdf form are used directly, no
     conversion needed.
  2. Any Consolidated Sheet that's a .docx also gets converted to PDF
     the same way; native .pdf Consolidated Sheets are used as-is.
  3. For each signed document's PDF: append all Consolidated Sheet
     PDFs, in upload order (pdf_merge_service).
  4. Generate processing_report.xlsx, zip the final PDFs + report.

------------------------------------------------------------------------
Both pipelines run on a background thread so the dashboard can poll
/processing/status/<job_id> for live progress (current item, X/Y
completed) instead of the request hanging until the whole batch
finishes. See job_tracker for the in-memory progress store.
"""

import os
import threading
from flask import Blueprint, request, jsonify, send_file, current_app
from flask_login import login_required

from app.services import (
    excel_service, docx_template_service, zip_service, report_service,
    pdf_conversion_service, pdf_merge_service,
)
from app.utils import job_tracker

processing_bp = Blueprint("processing", __name__, url_prefix="/processing")

ALLOWED_EXCEL_EXT = {"xlsx", "xls"}
ALLOWED_SIGNED_EXT = {"docx", "pdf"}
ALLOWED_CONSOLIDATED_EXT = {"docx", "pdf"}


def _allowed(filename, allowed_exts):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed_exts


def _save_uploads(files, upload_dir):
    """Save a list of FileStorage objects to disk, return [(filename, path), ...]."""
    entries = []
    for f in files:
        path = os.path.join(upload_dir, f.filename)
        f.save(path)
        entries.append((f.filename, path))
    return entries


# =========================================================================
# WORKFLOW 1 -- Generate Acceptance Test Documents
# =========================================================================
def _run_workflow1_pipeline(job_id, job_output_dir, excel_entries):
    """Excel rows -> Acceptance Test .docx documents. Background thread."""
    job_tracker.set_status(job_id, "processing")
    job_tracker.log(job_id, f"Received {len(excel_entries)} Excel sheet(s).")

    try:
        excel_result = excel_service.read_all_rows(excel_entries)
    except excel_service.ExcelLookupError as exc:
        job_tracker.log(job_id, str(exc), level="error")
        job_tracker.set_status(job_id, "completed")
        job_tracker.finish_timer(job_id)
        return

    for msg in excel_result["logs"]:
        level = "warning" if "no column found" in msg else "info"
        job_tracker.log(job_id, msg, level=level)

    rows = excel_result["rows"]
    job_tracker.start_timer(job_id, total_files=len(rows))

    successful_output_paths = []

    for row in rows:
        row_label = f"{row['_source_file']} (row {row['_row_number']})"
        job_tracker.set_current_file(job_id, row_label)
        job_tracker.increment_stat(job_id, "total")

        # Blank optional fields (Property ID, Bandwidth, Installation Date,
        # Router Type, etc.) are NOT treated as failures -- they're just
        # left blank in the generated document. Log which fields were
        # blank for visibility, but always attempt generation.
        if row["_blank_fields"]:
            blank_labels = ", ".join(excel_service.FIELD_LABELS[f] for f in row["_blank_fields"])
            job_tracker.log(job_id, f"{row_label}: blank field(s) left empty in the document: {blank_labels}", level="info")

        output_name = docx_template_service.build_output_filename(row)
        output_path = os.path.join(job_output_dir, output_name)

        try:
            docx_template_service.generate_acceptance_document(row, output_path)
        except Exception as exc:
            reason = f"Document generation failed: {exc}"
            job_tracker.log(job_id, f"FAILED {row_label}: {reason}", level="error")
            job_tracker.increment_stat(job_id, "failed")
            job_tracker.add_record(
                job_id, source_file=row["_source_file"], row_number=row["_row_number"],
                customer_name=row.get("customer_name", ""), account_number=row.get("account_number", ""),
                status="Failed", error_message=reason,
            )
            job_tracker.increment_processed(job_id)
            continue

        job_tracker.log(job_id, f"{row_label} -> '{output_name}' generated successfully.", level="success")
        job_tracker.increment_stat(job_id, "success")
        job_tracker.add_record(
            job_id, source_file=row["_source_file"], row_number=row["_row_number"],
            customer_name=row.get("customer_name", ""), account_number=row.get("account_number", ""),
            property_id=row.get("property_id", ""), status="Success", output_document_name=output_name,
        )
        successful_output_paths.append(output_path)
        job_tracker.increment_processed(job_id)

    job = job_tracker.get_job(job_id)
    report_path = os.path.join(job_output_dir, "processing_report.xlsx")
    report_service.generate_report(job["records"], report_path, columns=report_service.WORKFLOW1_COLUMNS)
    job_tracker.set_report_path(job_id, report_path)
    job_tracker.log(job_id, "Generated processing_report.xlsx.", level="info")

    zip_entries = [(p, os.path.basename(p)) for p in successful_output_paths]
    zip_entries.append((report_path, "processing_report.xlsx"))

    zip_path = os.path.join(job_output_dir, "acceptance_documents.zip")
    zip_ok, zip_msg = zip_service.create_zip_archive(zip_entries, zip_path)
    if zip_ok:
        job_tracker.set_zip_path(job_id, zip_path, download_name="acceptance_documents.zip")
        job_tracker.log(job_id, zip_msg, level="success")
    else:
        job_tracker.log(job_id, zip_msg, level="error")

    job_tracker.finish_timer(job_id)
    job_tracker.set_status(job_id, "completed")


@processing_bp.route("/run", methods=["POST"])
@login_required
def run_processing():
    """
    Workflow 1. Accepts one or more Excel sheets, saves them to disk,
    and starts document generation on a background thread.

    Expected multipart form field:
      - excel_files  (1+ files, .xlsx/.xls)
    """
    excel_files = request.files.getlist("excel_files")

    if not excel_files or not all(_allowed(f.filename, ALLOWED_EXCEL_EXT) for f in excel_files):
        return jsonify({"error": "At least one valid Excel file (.xlsx/.xls) is required."}), 400

    job_id = job_tracker.create_job()
    job_upload_dir = os.path.join(current_app.config["UPLOAD_FOLDER"], job_id)
    job_output_dir = os.path.join(current_app.config["PROCESSED_FOLDER"], job_id)
    os.makedirs(job_upload_dir, exist_ok=True)
    os.makedirs(job_output_dir, exist_ok=True)

    excel_entries = _save_uploads(excel_files, job_upload_dir)

    thread = threading.Thread(
        target=_run_workflow1_pipeline,
        args=(job_id, job_output_dir, excel_entries),
        daemon=True,
    )
    thread.start()

    return jsonify({"job_id": job_id, "status": "processing"}), 202


# =========================================================================
# WORKFLOW 2 -- Prepare Final Signed Documents
# =========================================================================
def _run_workflow2_pipeline(job_id, job_upload_dir, job_output_dir, signed_entries, consolidated_paths):
    """Signed .docx + Consolidated Sheets -> final PDFs. Background thread."""
    job_tracker.set_status(job_id, "processing")
    job_tracker.log(job_id, f"Received {len(signed_entries)} signed document(s) and {len(consolidated_paths)} Consolidated Sheet(s).")
    job_tracker.start_timer(job_id, total_files=len(signed_entries))

    # --- Step 1: Convert Consolidated Sheets to PDF where needed (once, reused for every signed doc) ---
    consolidated_docx = [p for p in consolidated_paths if p.lower().endswith(".docx")]
    consolidated_pdf_native = [p for p in consolidated_paths if p.lower().endswith(".pdf")]

    conversion_dir = os.path.join(job_output_dir, "_converted")
    docx_to_pdf_map = {}
    if consolidated_docx:
        job_tracker.log(job_id, f"Converting {len(consolidated_docx)} Consolidated Sheet .docx file(s) to PDF...")
        try:
            docx_to_pdf_map = pdf_conversion_service.convert_docx_files_to_pdf(consolidated_docx, conversion_dir)
        except pdf_conversion_service.ConversionError as exc:
            job_tracker.log(job_id, f"FATAL: {exc}", level="error")
            job_tracker.set_status(job_id, "completed")
            job_tracker.finish_timer(job_id)
            return

        failed_conversions = [p for p in consolidated_docx if p not in docx_to_pdf_map]
        for p in failed_conversions:
            job_tracker.log(job_id, f"WARNING: Consolidated Sheet '{os.path.basename(p)}' failed to convert to PDF and will be skipped.", level="warning")

    # Rebuild the consolidated list in ORIGINAL upload order, substituting
    # converted PDFs for .docx entries -- order matters per the spec.
    ordered_consolidated_pdfs = []
    for p in consolidated_paths:
        if p.lower().endswith(".pdf"):
            ordered_consolidated_pdfs.append(p)
        elif p in docx_to_pdf_map:
            ordered_consolidated_pdfs.append(docx_to_pdf_map[p])
        # else: conversion failed, already warned above -- skip it

    # --- Step 2: Convert signed .docx documents to PDF; .pdf documents are used as-is ---
    signed_docx_paths = [path for _, path in signed_entries if path.lower().endswith(".docx")]
    signed_pdf_map = {}
    if signed_docx_paths:
        job_tracker.log(job_id, f"Converting {len(signed_docx_paths)} signed .docx document(s) to PDF...")
        try:
            signed_pdf_map = pdf_conversion_service.convert_docx_files_to_pdf(signed_docx_paths, conversion_dir)
        except pdf_conversion_service.ConversionError as exc:
            job_tracker.log(job_id, f"FATAL: {exc}", level="error")
            job_tracker.set_status(job_id, "completed")
            job_tracker.finish_timer(job_id)
            return

    # Signed .pdf documents need no conversion -- map them to themselves so
    # the lookup below works the same way regardless of original format.
    for _, path in signed_entries:
        if path.lower().endswith(".pdf"):
            signed_pdf_map[path] = path

    # --- Step 3: For each signed document, merge its PDF with the Consolidated Sheets ---
    successful_output_paths = []

    for filename, original_path in signed_entries:
        job_tracker.set_current_file(job_id, filename)
        job_tracker.increment_stat(job_id, "total")

        signed_pdf_path = signed_pdf_map.get(original_path)
        if not signed_pdf_path:
            reason = "Could not convert this document to PDF (it may be corrupted or password-protected)."
            job_tracker.log(job_id, f"FAILED '{filename}': {reason}", level="error")
            job_tracker.increment_stat(job_id, "failed")
            job_tracker.add_record(
                job_id, input_file=filename, consolidated_count=len(ordered_consolidated_pdfs),
                status="Failed", error_message=reason,
            )
            job_tracker.increment_processed(job_id)
            continue

        output_name = os.path.splitext(filename)[0] + ".pdf"
        output_path = os.path.join(job_output_dir, output_name)

        merge_result = pdf_merge_service.merge_pdfs(signed_pdf_path, ordered_consolidated_pdfs, output_path)
        if not merge_result.success:
            job_tracker.log(job_id, f"FAILED '{filename}': {merge_result.message}", level="error")
            job_tracker.increment_stat(job_id, "failed")
            job_tracker.add_record(
                job_id, input_file=filename, consolidated_count=len(ordered_consolidated_pdfs),
                status="Failed", error_message=merge_result.message,
            )
            job_tracker.increment_processed(job_id)
            continue

        job_tracker.log(job_id, f"'{filename}' -> '{output_name}' finalized successfully.", level="success")
        job_tracker.increment_stat(job_id, "success")
        job_tracker.add_record(
            job_id, input_file=filename, consolidated_count=len(ordered_consolidated_pdfs),
            status="Success", output_file=output_name,
        )
        successful_output_paths.append(output_path)
        job_tracker.increment_processed(job_id)

    # --- Step 4: Report + ZIP ---
    job = job_tracker.get_job(job_id)
    report_path = os.path.join(job_output_dir, "final_documents_report.xlsx")
    report_service.generate_report(job["records"], report_path, columns=report_service.WORKFLOW2_COLUMNS)
    job_tracker.set_report_path(job_id, report_path)
    job_tracker.log(job_id, "Generated final_documents_report.xlsx.", level="info")

    zip_entries = [(p, os.path.basename(p)) for p in successful_output_paths]
    zip_entries.append((report_path, "final_documents_report.xlsx"))

    zip_path = os.path.join(job_output_dir, "final_signed_documents.zip")
    zip_ok, zip_msg = zip_service.create_zip_archive(zip_entries, zip_path)
    if zip_ok:
        job_tracker.set_zip_path(job_id, zip_path, download_name="final_signed_documents.zip")
        job_tracker.log(job_id, zip_msg, level="success")
    else:
        job_tracker.log(job_id, zip_msg, level="error")

    job_tracker.finish_timer(job_id)
    job_tracker.set_status(job_id, "completed")


@processing_bp.route("/run-final", methods=["POST"])
@login_required
def run_final_processing():
    """
    Workflow 2. Accepts one or more signed Acceptance Test documents
    (.docx or .pdf) and one or more Consolidated Sheet files (.docx/.pdf),
    saves them to disk, and starts the finalization pipeline on a
    background thread.

    Expected multipart form fields:
      - signed_files        (1+ files, .docx or .pdf)
      - consolidated_files  (1+ files, .docx/.pdf)
    """
    signed_files = request.files.getlist("signed_files")
    consolidated_files = request.files.getlist("consolidated_files")

    if not signed_files or not all(_allowed(f.filename, ALLOWED_SIGNED_EXT) for f in signed_files):
        return jsonify({"error": "At least one valid signed document (.docx or .pdf) is required."}), 400

    if not consolidated_files or not all(_allowed(f.filename, ALLOWED_CONSOLIDATED_EXT) for f in consolidated_files):
        return jsonify({"error": "At least one valid Consolidated Sheet (.docx or .pdf) is required."}), 400

    job_id = job_tracker.create_job()
    job_upload_dir = os.path.join(current_app.config["UPLOAD_FOLDER"], job_id)
    job_output_dir = os.path.join(current_app.config["PROCESSED_FOLDER"], job_id)
    os.makedirs(job_upload_dir, exist_ok=True)
    os.makedirs(job_output_dir, exist_ok=True)

    signed_entries = _save_uploads(signed_files, job_upload_dir)
    # Saved in upload order -- preserved through to the final merge.
    consolidated_paths = [path for _, path in _save_uploads(consolidated_files, job_upload_dir)]

    thread = threading.Thread(
        target=_run_workflow2_pipeline,
        args=(job_id, job_upload_dir, job_output_dir, signed_entries, consolidated_paths),
        daemon=True,
    )
    thread.start()

    return jsonify({"job_id": job_id, "status": "processing"}), 202


# =========================================================================
# Shared endpoints (workflow-agnostic -- just serve whatever job_tracker has)
# =========================================================================
@processing_bp.route("/download/<job_id>")
@login_required
def download_zip(job_id):
    job = job_tracker.get_job(job_id)
    if job is None or not job["zip_path"] or not os.path.exists(job["zip_path"]):
        return jsonify({"error": "No processed ZIP available for this job."}), 404

    return send_file(job["zip_path"], as_attachment=True, download_name=job["zip_download_name"])


@processing_bp.route("/download-report/<job_id>")
@login_required
def download_report(job_id):
    """Serve just the Excel report on its own, without the full ZIP."""
    job = job_tracker.get_job(job_id)
    if job is None or not job["report_path"] or not os.path.exists(job["report_path"]):
        return jsonify({"error": "No processing report available for this job."}), 404

    return send_file(job["report_path"], as_attachment=True, download_name=os.path.basename(job["report_path"]))


@processing_bp.route("/status/<job_id>")
@login_required
def job_status(job_id):
    job = job_tracker.get_job(job_id)
    if job is None:
        return jsonify({"error": "Job not found."}), 404
    return jsonify(job)
