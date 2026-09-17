"""
app/services/zip_service.py
------------------------------
Packages all successfully processed PDFs into a single downloadable
ZIP file.

Streams files into the archive one at a time (rather than loading
every PDF into memory first) so this scales to 300+ files without
memory pressure, per the project's performance requirement.
"""

import os
import zipfile


def create_zip_archive(entries, output_zip_path):
    """
    Create a ZIP archive containing the given files.

    Args:
        entries: list of items, where each item is either:
                   - a plain file path (str) -> placed at the ZIP root
                     using just its filename, or
                   - a (source_path, arcname) tuple for explicit control
                     over where the file lands inside the archive
                     (e.g. ("...manual_review/foo.pdf", "manual_review/foo.pdf")).
        output_zip_path: full path where the .zip should be written.

    Returns:
        (success: bool, message: str)
    """
    try:
        os.makedirs(os.path.dirname(output_zip_path), exist_ok=True)

        written = 0
        with zipfile.ZipFile(output_zip_path, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for entry in entries:
                if isinstance(entry, tuple):
                    path, arcname = entry
                else:
                    path, arcname = entry, os.path.basename(entry)

                if os.path.exists(path):
                    zf.write(path, arcname=arcname)
                    written += 1

        return True, f"Created ZIP with {written} file(s)."

    except Exception as exc:
        return False, f"Failed to create ZIP archive: {exc}"
