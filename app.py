import gc
import csv
import json
import logging
import os
import re
import tempfile
import time
import zipfile
from datetime import datetime
from functools import wraps
from typing import Optional, Tuple

from flask import Flask, Response, after_this_request, jsonify, request, send_file, stream_with_context
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
ZIP_COMPRESSION_LEVEL = max(0, min(9, int(os.environ.get("ZIP_COMPRESSION_LEVEL", 1))))
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

    This intentionally skips style formatting but preserves worksheet structure.
    """
    output_wb = Workbook(write_only=False)
    try:
        output_wb.remove(output_wb.active)
        target_ws = output_wb.create_sheet(title=sheet_name)

        # Preserve worksheet structure (without cell formatting styles).
        target_ws.freeze_panes = source_ws.freeze_panes
        if source_ws.auto_filter and source_ws.auto_filter.ref:
            target_ws.auto_filter.ref = source_ws.auto_filter.ref

        for col_letter, col_dim in source_ws.column_dimensions.items():
            target_dim = target_ws.column_dimensions[col_letter]
            target_dim.width = col_dim.width
            target_dim.hidden = col_dim.hidden
            target_dim.outlineLevel = col_dim.outlineLevel
            target_dim.collapsed = col_dim.collapsed

        for row_idx, row_dim in source_ws.row_dimensions.items():
            target_dim = target_ws.row_dimensions[row_idx]
            target_dim.height = row_dim.height
            target_dim.hidden = row_dim.hidden
            target_dim.outlineLevel = row_dim.outlineLevel
            target_dim.collapsed = row_dim.collapsed

        for merged_range in source_ws.merged_cells.ranges:
            target_ws.merge_cells(str(merged_range))

        # Copy only populated cells to avoid scanning huge empty grids.
        for source_cell in source_ws._cells.values():
            if source_cell.value is not None:
                target_ws.cell(
                    row=source_cell.row,
                    column=source_cell.column,
                    value=source_cell.value,
                )

        output_wb.save(output_path)
    finally:
        try:
            output_wb.close()
        except Exception:
            pass


def _write_sheet_as_txt(source_ws, sheet_name: str, output_path: str) -> None:
    """
    Write worksheet to tab-separated text while keeping grid structure.
    
    The top comment lines include structural metadata such as merged ranges.
    """
    max_row = source_ws.max_row or 0
    max_col = source_ws.max_column or 0
    merged_ranges = ""
    if hasattr(source_ws, "merged_cells"):
        merged_ranges = ",".join(str(r) for r in source_ws.merged_cells.ranges)

    with open(output_path, "w", encoding="utf-8", newline="") as txt_file:
        txt_file.write(f"# sheet_name\t{sheet_name}\n")
        txt_file.write(f"# max_row\t{max_row}\n")
        txt_file.write(f"# max_col\t{max_col}\n")
        txt_file.write(f"# merged_ranges\t{merged_ranges}\n")

        writer = csv.writer(
            txt_file,
            delimiter="\t",
            lineterminator="\n",
            quoting=csv.QUOTE_MINIMAL,
        )
        for row in source_ws.iter_rows(values_only=True):
            writer.writerow(["" if value is None else value for value in row])


def _row_to_content_string(row_values) -> str:
    """Convert row tuple to tab-separated text content."""
    return "\t".join("" if value is None else str(value) for value in row_values)


@timing_decorator
def split_excel_by_sheets_simple(
    uploaded_stream,
    original_filename: str,
    temp_dir: str,
    output_format: str = "xlsx",
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
        read_only_mode = output_format == "txt"
        try:
            source_wb = load_workbook(
                uploaded_stream,
                read_only=read_only_mode,
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
                output_filename = f"{base_name}_{sanitize_filename(sheet_name)}.{output_format}"
                output_path = os.path.join(temp_dir, output_filename)
                if output_format == "txt":
                    _write_sheet_as_txt(source_ws, sheet_name, output_path)
                else:
                    _write_sheet_as_workbook(source_ws, sheet_name, output_path)
                generated_paths.append(output_path)
            except Exception as e:
                logger.error("Error processing sheet '%s': %s", sheet_name, e)
                continue

        if not generated_paths:
            return None, None, None, "No sheets could be processed successfully"

        if len(generated_paths) == 1:
            one_file = generated_paths[0]
            mime_type = (
                "text/plain"
                if output_format == "txt"
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            return (
                one_file,
                mime_type,
                os.path.basename(one_file),
                None,
            )

        zip_path = os.path.join(temp_dir, f"{base_name}_split.zip")
        zip_compression = (
            zipfile.ZIP_STORED
            if output_format == "txt" or ZIP_COMPRESSION_LEVEL == 0
            else zipfile.ZIP_DEFLATED
        )
        zip_kwargs = {}
        if zip_compression == zipfile.ZIP_DEFLATED:
            zip_kwargs["compresslevel"] = ZIP_COMPRESSION_LEVEL

        with zipfile.ZipFile(zip_path, "w", zip_compression, **zip_kwargs) as zip_file:
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
            "description": "Split Excel files by visible sheets while preserving structure",
            "endpoints": {
                "POST /split-excel": "Upload Excel file to split",
                "POST /split-excel-ndjson": "Upload Excel and stream NDJSON rows",
                "GET /health": "Health check endpoint",
                "GET /": "Service information",
            },
            "features": [
                "Sheet structure preservation (no cell styling)",
                "Skips hidden sheets",
                "Preserves cell values and formulas",
                "Single-sheet direct response or multi-sheet ZIP",
            ],
            "configuration": {
                "max_file_size": f"{MAX_FILE_SIZE / (1024 * 1024):.1f} MB",
                "max_sheets": MAX_SHEETS,
                "allowed_extensions": list(ALLOWED_EXTENSIONS),
                "output_formats": ["xlsx", "txt", "ndjson"],
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

        output_format = (
            request.args.get("format")
            or request.form.get("format")
            or "xlsx"
        ).strip().lower()
        if output_format not in {"xlsx", "txt"}:
            return jsonify({"error": "Invalid format. Allowed formats: xlsx, txt"}), 400

        file.seek(0, os.SEEK_END)
        upload_size = file.tell()
        file.seek(0)
        logger.info(
            "Received file: %s (%.2f MB), output_format=%s",
            file.filename,
            upload_size / (1024 * 1024),
            output_format,
        )

        temp_dir_ctx = tempfile.TemporaryDirectory(prefix="excel-split-")
        response_path, mime_type, download_name, error = split_excel_by_sheets_simple(
            file.stream,
            file.filename,
            temp_dir_ctx.name,
            output_format,
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


@app.route("/split-excel-ndjson", methods=["POST"])
def split_excel_ndjson():
    """
    Stream workbook content as NDJSON with one row object per line.

    Output line schema:
    {"name": "<sheet_name>", "row": <row_number>, "content": "<tab-separated row text>"}
    """
    file = None
    upload_path = None
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file provided"}), 400

        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "No file selected"}), 400
        if not allowed_file(file.filename):
            return (
                jsonify({"error": f'Invalid file type. Allowed types: {", ".join(ALLOWED_EXTENSIONS)}'}),
                400,
            )

        suffix = os.path.splitext(file.filename)[1] or ".xlsx"
        with tempfile.NamedTemporaryFile(prefix="excel-upload-", suffix=suffix, delete=False) as tmp_file:
            upload_path = tmp_file.name
            total_bytes = 0
            while True:
                chunk = file.stream.read(1024 * 1024)
                if not chunk:
                    break
                total_bytes += len(chunk)
                if total_bytes > MAX_FILE_SIZE:
                    raise RequestEntityTooLarge()
                tmp_file.write(chunk)

        logger.info(
            "Received file for ndjson: %s (%.2f MB)",
            file.filename,
            total_bytes / (1024 * 1024),
        )

        base_name = os.path.splitext(file.filename)[0]

        @stream_with_context
        def generate():
            source_wb = None
            try:
                source_wb = load_workbook(
                    upload_path,
                    read_only=True,
                    data_only=False,
                    keep_links=False,
                )
                if len(source_wb.sheetnames) > MAX_SHEETS:
                    yield json.dumps(
                        {
                            "error": f"File contains too many sheets ({len(source_wb.sheetnames)}). Maximum allowed: {MAX_SHEETS}"
                        }
                    ) + "\n"
                    return

                visible_sheets = []
                for sheet_name in source_wb.sheetnames:
                    sheet = source_wb[sheet_name]
                    if sheet.sheet_state == "visible":
                        visible_sheets.append(sheet_name)

                if not visible_sheets:
                    yield json.dumps({"error": "No visible sheets found in workbook"}) + "\n"
                    return

                for sheet_name in visible_sheets:
                    ws = source_wb[sheet_name]
                    row_number = 0
                    for row in ws.iter_rows(values_only=True):
                        row_number += 1
                        yield json.dumps(
                            {
                                "name": sheet_name,
                                "row": row_number,
                                "content": _row_to_content_string(row),
                            },
                            ensure_ascii=False,
                        ) + "\n"
            except Exception as e:
                logger.error("Error streaming NDJSON: %s", e, exc_info=True)
                yield json.dumps({"error": f"Error processing Excel file: {e}"}) + "\n"
            finally:
                if source_wb is not None:
                    try:
                        source_wb.close()
                    except Exception:
                        pass
                if upload_path and os.path.exists(upload_path):
                    try:
                        os.remove(upload_path)
                    except OSError:
                        pass
                gc.collect()

        response = Response(generate(), mimetype="application/x-ndjson")
        response.headers["Content-Disposition"] = f'attachment; filename="{base_name}_split.ndjson"'
        return response
    except RequestEntityTooLarge:
        return jsonify({"error": f"File too large. Maximum size: {MAX_FILE_SIZE / (1024 * 1024):.1f} MB"}), 413
    except Exception as e:
        logger.error("Unexpected error in ndjson endpoint: %s", e, exc_info=True)
        return jsonify({"error": "An unexpected error occurred"}), 500
    finally:
        if file is not None:
            try:
                file.close()
            except Exception:
                pass


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