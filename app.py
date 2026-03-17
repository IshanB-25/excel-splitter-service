import gc
import logging
import os
import re
import tempfile
import time
import zipfile
from datetime import datetime
from functools import wraps
from typing import Optional, Tuple

from flask import Flask, after_this_request, jsonify, request, send_file
from openpyxl import Workbook, load_workbook
from werkzeug.exceptions import RequestEntityTooLarge

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

MAX_FILE_SIZE_MB = int(os.environ.get("MAX_FILE_SIZE_MB", 50))
MAX_FILE_SIZE = MAX_FILE_SIZE_MB * 1024 * 1024
MAX_SHEETS = int(os.environ.get("MAX_SHEETS", 100))
ALLOWED_EXTENSIONS = {"xlsx", "xls", "xlsm", "xlsb"}
PORT = int(os.environ.get("PORT", 3070))

app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_SIZE


def timing_decorator(f):
    """Decorator to measure and log function execution time."""

    @wraps(f)
    def wrapper(*args, **kwargs):
        start = time.time()
        result = f(*args, **kwargs)
        duration = time.time() - start
        logger.info("%s took %.2f seconds", f.__name__, duration)
        return result

    return wrapper


def sanitize_filename(filename: str) -> str:
    """Sanitize filename for safe file system usage."""
    sanitized = re.sub(r'[<>:"/\\|?*]', "_", filename).strip(". ")
    if len(sanitized) > 100:
        name, ext = os.path.splitext(sanitized)
        sanitized = name[: 100 - len(ext)] + ext
    return sanitized or "unnamed"


def allowed_file(filename: str) -> bool:
    """Check if file has an allowed extension."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def _write_sheet_as_workbook(source_ws, sheet_name: str, output_path: str) -> None:
    """
    Write one worksheet into a new workbook.

    This intentionally copies only cell values/formulas for speed and low memory.
    """
    output_wb = Workbook(write_only=True)
    try:
        target_ws = output_wb.create_sheet(title=sheet_name)
        for row in source_ws.iter_rows(values_only=True):
            target_ws.append(row)
        output_wb.save(output_path)
    finally:
        try:
            output_wb.close()
        except Exception:
            pass


@timing_decorator
def split_excel_by_sheets_simple(
    uploaded_stream,
    original_filename: str,
    temp_dir: str,
) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str]]:
    """
    Split workbook into one workbook per visible sheet.

    Returns:
        Tuple of (response_path, mime_type, download_name, error_message)
    """
    base_name = os.path.splitext(original_filename)[0]
    source_wb = None
    try:
        uploaded_stream.seek(0)
        try:
            source_wb = load_workbook(
                uploaded_stream,
                read_only=True,
                data_only=False,
                keep_links=False,
            )
        except Exception as e:
            logger.warning("Excel validation failed: %s", e)
            return None, None, None, f"Invalid Excel file: {e}"

        if len(source_wb.sheetnames) > MAX_SHEETS:
            return (
                None,
                None,
                None,
                f"File contains too many sheets ({len(source_wb.sheetnames)}). Maximum allowed: {MAX_SHEETS}",
            )

        visible_sheets = []
        hidden_sheets = []
        for sheet_name in source_wb.sheetnames:
            sheet = source_wb[sheet_name]
            if sheet.sheet_state == "visible":
                visible_sheets.append(sheet_name)
            else:
                hidden_sheets.append(sheet_name)

        logger.info("Processing %s visible sheets from %s", len(visible_sheets), original_filename)
        if hidden_sheets:
            logger.info("Skipping %s hidden sheets", len(hidden_sheets))
        if not visible_sheets:
            return None, None, None, "No visible sheets found in workbook"

        generated_paths = []
        for sheet_name in visible_sheets:
            try:
                logger.info("Processing sheet: %s", sheet_name)
                source_ws = source_wb[sheet_name]
                output_filename = f"{base_name}_{sanitize_filename(sheet_name)}.xlsx"
                output_path = os.path.join(temp_dir, output_filename)
                _write_sheet_as_workbook(source_ws, sheet_name, output_path)
                generated_paths.append(output_path)
            except Exception as e:
                logger.error("Error processing sheet '%s': %s", sheet_name, e)
                continue
            finally:
                gc.collect()

        if not generated_paths:
            return None, None, None, "No sheets could be processed successfully"

        if len(generated_paths) == 1:
            one_file = generated_paths[0]
            return (
                one_file,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                os.path.basename(one_file),
                None,
            )

        zip_path = os.path.join(temp_dir, f"{base_name}_split.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zip_file:
            for generated_path in generated_paths:
                zip_file.write(generated_path, arcname=os.path.basename(generated_path))
                try:
                    os.remove(generated_path)
                except OSError:
                    pass

        return zip_path, "application/zip", os.path.basename(zip_path), None
    except Exception as e:
        logger.error("Error splitting Excel file: %s", e, exc_info=True)
        return None, None, None, f"Error processing Excel file: {e}"
    finally:
        if source_wb is not None:
            try:
                source_wb.close()
            except Exception:
                pass


@app.route("/", methods=["GET"])
def index():
    """Service information endpoint."""
    return jsonify(
        {
            "service": "Excel File Splitter",
            "version": "2.2.0",
            "description": "Split Excel files by visible sheets with low-memory processing",
            "endpoints": {
                "POST /split-excel": "Upload Excel file to split",
                "GET /health": "Health check endpoint",
                "GET /": "Service information",
            },
            "features": [
                "Low-memory sheet splitting",
                "Skips hidden sheets",
                "Preserves cell values and formulas",
                "Single-sheet direct response or multi-sheet ZIP",
            ],
            "configuration": {
                "max_file_size": f"{MAX_FILE_SIZE / (1024 * 1024):.1f} MB",
                "max_sheets": MAX_SHEETS,
                "allowed_extensions": list(ALLOWED_EXTENSIONS),
            },
            "timestamp": datetime.utcnow().isoformat(),
        }
    )


@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint for monitoring."""
    return (
        jsonify(
            {
                "status": "healthy",
                "timestamp": datetime.utcnow().isoformat(),
                "service": "excel-splitter",
            }
        ),
        200,
    )


@app.route("/split-excel", methods=["POST"])
def split_excel():
    """Split an uploaded workbook by visible sheets."""
    file = None
    temp_dir_ctx = None
    cleanup_registered = False
    try:
        if "file" not in request.files:
            logger.warning("No file in request")
            return jsonify({"error": "No file provided"}), 400

        file = request.files["file"]
        if file.filename == "":
            logger.warning("Empty filename")
            return jsonify({"error": "No file selected"}), 400

        if not allowed_file(file.filename):
            logger.warning("Invalid file extension: %s", file.filename)
            return (
                jsonify({"error": f'Invalid file type. Allowed types: {", ".join(ALLOWED_EXTENSIONS)}'}),
                400,
            )

        file.seek(0, os.SEEK_END)
        upload_size = file.tell()
        file.seek(0)
        logger.info("Received file: %s (%.2f MB)", file.filename, upload_size / (1024 * 1024))

        temp_dir_ctx = tempfile.TemporaryDirectory(prefix="excel-split-")
        response_path, mime_type, download_name, error = split_excel_by_sheets_simple(
            file.stream,
            file.filename,
            temp_dir_ctx.name,
        )

        if error:
            logger.error("Splitting failed: %s", error)
            status_code = 400 if error.startswith("Invalid Excel file") else 500
            return jsonify({"error": error}), status_code

        if not response_path:
            return jsonify({"error": "No sheets found in Excel file"}), 400

        tmp_ref = temp_dir_ctx

        @after_this_request
        def _cleanup_temp_dir(response):
            try:
                tmp_ref.cleanup()
            except Exception:
                logger.warning("Could not cleanup temporary directory", exc_info=True)
            return response

        cleanup_registered = True
        logger.info("Returning processed file: %s", download_name)
        return send_file(response_path, mimetype=mime_type, as_attachment=True, download_name=download_name)
    except RequestEntityTooLarge:
        logger.warning("File too large (max: %.1f MB)", MAX_FILE_SIZE / (1024 * 1024))
        return jsonify({"error": f"File too large. Maximum size: {MAX_FILE_SIZE / (1024 * 1024):.1f} MB"}), 413
    except Exception as e:
        logger.error("Unexpected error: %s", e, exc_info=True)
        return jsonify({"error": "An unexpected error occurred"}), 500
    finally:
        if file is not None:
            try:
                file.close()
            except Exception:
                pass
        if temp_dir_ctx is not None and not cleanup_registered:
            try:
                temp_dir_ctx.cleanup()
            except Exception:
                pass
        gc.collect()


@app.errorhandler(413)
def request_entity_too_large(_):
    """Handle file size limit exceeded."""
    return jsonify({"error": f"File too large. Maximum size: {MAX_FILE_SIZE / (1024 * 1024):.1f} MB"}), 413


@app.errorhandler(500)
def internal_server_error(e):
    """Handle internal server errors."""
    logger.error("Internal server error: %s", e, exc_info=True)
    return jsonify({"error": "Internal server error"}), 500


if __name__ == "__main__":
    logger.info("Starting Excel Splitter Service on port %s", PORT)
    logger.info("Configuration: MAX_FILE_SIZE=%sMB, MAX_SHEETS=%s", MAX_FILE_SIZE_MB, MAX_SHEETS)
    app.run(host="0.0.0.0", port=PORT, debug=False)