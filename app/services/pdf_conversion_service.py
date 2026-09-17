"""
app/services/pdf_conversion_service.py
------------------------------------------
Converts Word (.docx) documents to PDF, using LibreOffice in headless
mode as an external process. This is what powers Workflow 2's first
step: "convert each signed Acceptance Test document to PDF."

Why LibreOffice: it's free, works the same way on Windows/Mac/Linux,
and produces reliable, high-fidelity PDF output from .docx -- the
standard approach for server-side conversion without requiring
Microsoft Word to be installed.

SYSTEM REQUIREMENT: LibreOffice must be installed on the machine
running this app, and the `soffice` (or `libreoffice`) command must be
on PATH. Download: https://www.libreoffice.org/download/
"""

import os
import shutil
import subprocess
import tempfile
import uuid


class ConversionError(Exception):
    """Raised when LibreOffice isn't available, or a conversion fails outright."""
    pass


def _find_soffice():
    for candidate in ("soffice", "libreoffice"):
        path = shutil.which(candidate)
        if path:
            return path
    return None


def convert_docx_files_to_pdf(docx_paths, output_dir, timeout=300):
    """
    Batch-convert multiple .docx files to PDF in ONE LibreOffice
    invocation -- meaningfully faster than launching a new process per
    file, since each LibreOffice startup has real overhead. This
    matters for batches of hundreds of documents.

    Each job gets an isolated LibreOffice user profile (via
    -env:UserInstallation) so concurrent jobs (this app processes
    jobs on background threads) don't contend over the same profile
    lock, which can otherwise cause hangs or silent failures.

    Args:
        docx_paths: list of .docx file paths to convert.
        output_dir: directory to write the resulting PDFs into.
        timeout: seconds to allow the whole batch conversion to run.

    Returns:
        {docx_path: pdf_path} for every file that converted successfully.
        A path missing from this dict means that specific file failed
        to convert (e.g. corrupt docx) -- callers should treat that as
        a per-file failure, not raise for the whole batch.

    Raises:
        ConversionError if LibreOffice isn't installed at all, or the
        whole batch times out.
    """
    if not docx_paths:
        return {}

    soffice = _find_soffice()
    if not soffice:
        raise ConversionError(
            "LibreOffice (soffice) is not installed or not on PATH. "
            "Workflow 2 requires LibreOffice to convert Word documents to PDF. "
            "Install it from https://www.libreoffice.org/download/ and ensure "
            "the 'soffice' command is available in your system PATH."
        )

    os.makedirs(output_dir, exist_ok=True)

    profile_dir = os.path.join(tempfile.gettempdir(), f"lo_profile_{uuid.uuid4().hex[:8]}")
    profile_uri = f"file://{profile_dir}" if os.name != "nt" else f"file:///{profile_dir.replace(os.sep, '/')}"

    cmd = [
        soffice, "--headless", "--norestore",
        f"-env:UserInstallation={profile_uri}",
        "--convert-to", "pdf", "--outdir", output_dir,
    ] + list(docx_paths)

    try:
        subprocess.run(cmd, capture_output=True, timeout=timeout, check=False)
    except subprocess.TimeoutExpired as exc:
        raise ConversionError(f"LibreOffice conversion timed out after {timeout}s: {exc}")
    except FileNotFoundError as exc:
        raise ConversionError(f"Could not run LibreOffice: {exc}")
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)

    results = {}
    for docx_path in docx_paths:
        stem = os.path.splitext(os.path.basename(docx_path))[0]
        expected_pdf = os.path.join(output_dir, f"{stem}.pdf")
        if os.path.exists(expected_pdf):
            results[docx_path] = expected_pdf
    return results
