"""
app/services/pdf_merge_service.py
------------------------------------
Appends one or more Consolidated Sheet PDFs to the end of a base PDF
(a signed Acceptance Test document that's already been converted from
.docx). This is Workflow 2's second step.

Uses PyMuPDF's insert_pdf, which appends all pages of one PDF onto
another directly -- no rasterization, no quality loss, pages stay
fully text-searchable. Any Consolidated Sheet that arrives as .docx
must be converted to PDF first (see pdf_conversion_service);
by the time a path reaches this module, everything is already a PDF.
"""

import os
import fitz  # PyMuPDF


class MergeResult:
    def __init__(self, success, output_path=None, message=""):
        self.success = success
        self.output_path = output_path
        self.message = message


def merge_pdfs(base_pdf_path, additional_pdf_paths, output_path):
    """
    Append each PDF in additional_pdf_paths, in order, to the end of
    base_pdf_path, and save the combined result to output_path.

    Args:
        base_pdf_path: the signed Acceptance Test PDF (already converted from .docx).
        additional_pdf_paths: list of Consolidated Sheet PDF paths, in upload order.
        output_path: where to save the final merged PDF.

    Returns:
        MergeResult
    """
    if not os.path.exists(base_pdf_path):
        return MergeResult(False, message=f"Base PDF not found: {base_pdf_path}")

    for path in additional_pdf_paths:
        if not os.path.exists(path):
            return MergeResult(False, message=f"Consolidated Sheet PDF not found: {path}")

    try:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        combined = fitz.open(base_pdf_path)
        for path in additional_pdf_paths:
            addition = fitz.open(path)
            combined.insert_pdf(addition)
            addition.close()

        combined.save(output_path)
        combined.close()

        return MergeResult(True, output_path=output_path)

    except Exception as exc:
        return MergeResult(False, message=f"Failed to merge PDFs: {exc}")
