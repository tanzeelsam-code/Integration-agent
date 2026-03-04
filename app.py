#!/usr/bin/env python3
"""
AGENT ZEE — Flask Web UI
Multi-job document processing assistant with drag-and-drop interface.
"""

from urllib.parse import quote
import os
import uuid
import tempfile
from pathlib import Path

from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

from formatter.engine import reformat_document
from jobs import list_jobs, get_job
from input_adapter import (
    is_supported_filename,
    normalize_input_to_docx,
    supported_extensions_csv,
)
from gis.map_engine import generate_full_map, export_map_html
from gis.data_loader import (
    load_default_data,
    load_geojson_from_string,
    get_feature_summary,
    get_available_layers,
)
from gis.layers import ENERGY_LAYERS


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "integration_agent_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)
SUPPORTED_UPLOADS = supported_extensions_csv()


def _build_download_name(prefix: str, source_filename: str, extension: str = ".docx") -> str:
    stem = Path(secure_filename(source_filename or "document")).stem or "document"
    ext = extension if extension.startswith(".") else f".{extension}"
    return f"{prefix}_{stem}{ext}"


def _append_conversion_notes(summary: str, notes: list[str]) -> str:
    filtered = [n for n in notes if n]
    if not filtered:
        return summary
    note_block = "\n".join(f"- {n}" for n in filtered)
    return f"{summary}\nINPUT CONVERSION NOTES\n{'=' * 40}\n{note_block}\n"


# Shared GIS state (per-process; fine for single-server)
_custom_geojson_data = None


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/jobs")
def api_jobs():
    """Return JSON list of all available jobs."""
    return jsonify(list_jobs())


@app.route("/map")
def map_page():
    return render_template("map.html")


@app.route("/api/map", methods=["POST"])
def api_generate_map():
    """Generate an interactive map with the requested layers."""
    global _custom_geojson_data

    body = request.get_json(silent=True) or {}
    active_layers = body.get("layers", list(ENERGY_LAYERS.keys()))

    data = _custom_geojson_data if _custom_geojson_data else load_default_data()

    try:
        m = generate_full_map(data=data, active_layers=active_layers)
        map_html = export_map_html(m)
        summary = get_feature_summary(data)
        total = sum(summary.get(k, 0) for k in active_layers)

        return jsonify(
            {
                "map_html": map_html,
                "total_features": total,
                "layer_summary": summary,
            }
        )
    except Exception as e:
        return jsonify({"error": f"Map generation failed: {str(e)}"}), 500


@app.route("/api/map/upload", methods=["POST"])
def api_upload_geojson():
    """Upload a custom GeoJSON file to overlay on the map."""
    global _custom_geojson_data

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not file.filename.lower().endswith((".geojson", ".json")):
        return jsonify({"error": "Only .geojson or .json files are supported"}), 400

    try:
        raw = file.read().decode("utf-8")
        data = load_geojson_from_string(raw)

        # Merge uploaded features with defaults
        default = load_default_data()
        merged_features = default.get("features", []) + data.get("features", [])
        _custom_geojson_data = {
            "type": "FeatureCollection",
            "features": merged_features,
        }

        return jsonify(
            {
                "success": True,
                "features_count": len(data.get("features", [])),
                "total_features": len(merged_features),
                "layers": get_available_layers(_custom_geojson_data),
            }
        )
    except (ValueError, UnicodeDecodeError) as e:
        return jsonify({"error": f"Invalid GeoJSON: {str(e)}"}), 400
    except Exception as e:
        return jsonify({"error": f"Upload failed: {str(e)}"}), 500


@app.route("/api/map/export")
def api_export_map():
    """Export the current map as a standalone HTML file."""
    global _custom_geojson_data

    data = _custom_geojson_data if _custom_geojson_data else load_default_data()

    try:
        m = generate_full_map(data=data)
        output_path = os.path.join(UPLOAD_DIR, f"map_{uuid.uuid4().hex[:8]}.html")
        export_map_html(m, output_path=output_path)

        return send_file(
            output_path,
            mimetype="text/html",
            as_attachment=True,
            download_name="gb_energy_map.html",
        )
    except Exception as e:
        return jsonify({"error": f"Export failed: {str(e)}"}), 500


@app.route("/api/upload_chunk", methods=["POST"])
def upload_chunk():
    """Endpoint to handle chunked file uploads to bypass 32MB Cloud Run limits."""
    if "file" not in request.files:
        return jsonify({"error": "No file chunk provided"}), 400

    chunk = request.files["file"]
    job_id = request.form.get("job_id")
    chunk_index = int(request.form.get("chunk_index", 0))
    total_chunks = int(request.form.get("total_chunks", 1))
    filename = request.form.get("filename", "document.docx")
    
    if not job_id:
        return jsonify({"error": "Missing job_id parameter"}), 400

    safe_filename = secure_filename(filename)
    chunk_path = os.path.join(UPLOAD_DIR, f"{job_id}_chunked_{safe_filename}")

    # Append chunk data to the target file
    mode = "ab" if chunk_index > 0 else "wb"
    try:
        with open(chunk_path, mode) as f:
            f.write(chunk.read())
            
        return jsonify({
            "success": True, 
            "message": f"Chunk {chunk_index + 1}/{total_chunks} received.",
            "job_id": job_id
        })
    except Exception as e:
        return jsonify({"error": f"Failed to save chunk: {str(e)}"}), 500


@app.route("/upload", methods=["POST"])
def upload():
    """Legacy endpoint — Document Formatting job."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not is_supported_filename(file.filename):
        return jsonify({"error": f"Unsupported file type. Supported: {SUPPORTED_UPLOADS}"}), 400

    report_name = request.form.get("report_name", "Report")
    year = request.form.get("year", "2026")
    project_number = request.form.get("project_number", "PRJ-001")

    job_id = str(uuid.uuid4())[:8]
    input_ext = Path(file.filename).suffix.lower()
    input_path = os.path.join(UPLOAD_DIR, f"{job_id}_input{input_ext}")
    prepared_input_path = os.path.join(UPLOAD_DIR, f"{job_id}_input_prepared.docx")
    output_path = os.path.join(UPLOAD_DIR, f"{job_id}_output.docx")

    file.save(input_path)

    try:
        prepared = normalize_input_to_docx(input_path, prepared_input_path)
        summary = reformat_document(
            input_path=prepared["path"],
            output_path=output_path,
            report_name=report_name,
            year=year,
            project_number=project_number,
        )
        summary = _append_conversion_notes(summary, [prepared.get("note")])

        out_filename = _build_download_name("FORMATTED", file.filename, ".docx")

        return jsonify(
            {
                "success": True,
                "job_id": job_id,
                "summary": summary,
                "filename": out_filename,
                "download_url": f"/download/{job_id}/{quote(out_filename)}",
            }
        )

    except Exception as e:
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

    finally:
        for p in [input_path, prepared_input_path]:
            if p and os.path.exists(p):
                os.remove(p)


@app.route("/process/<job_type>", methods=["POST"])
def process_job(job_type):
    """Universal endpoint for all job types."""
    job_def = get_job(job_type)
    if not job_def:
        return jsonify({"error": f"Unknown job type: {job_type}"}), 400

    if job_type == "formatting":
        return upload()

    # File handling
    chunked_job_id1 = request.form.get("chunked_job_id1")
    chunked_filename1 = request.form.get("chunked_filename1")
    chunked_job_id2 = request.form.get("chunked_job_id2")
    chunked_filename2 = request.form.get("chunked_filename2")

    file1 = None
    file1_name = ""
    file2 = None
    file2_name = ""

    if job_def.get("multi_file"):
        if not chunked_job_id1 and "file" not in request.files:
            return jsonify({"error": "First file required for comparison"}), 400
        if not chunked_job_id2 and "file2" not in request.files:
            return jsonify({"error": "Second file required for comparison"}), 400
            
        if not chunked_job_id1:
            file1 = request.files["file"]
            file1_name = file1.filename
        else:
            file1_name = chunked_filename1
            
        if not chunked_job_id2:
            file2 = request.files["file2"]
            file2_name = file2.filename
        else:
            file2_name = chunked_filename2
            
        if not file1_name or not file2_name:
            return jsonify({"error": "Both files must be selected"}), 400
    else:
        if not chunked_job_id1 and "file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400
            
        if not chunked_job_id1:
            file1 = request.files["file"]
            file1_name = file1.filename
        else:
            file1_name = chunked_filename1
            
        if not file1_name:
            return jsonify({"error": "No file selected"}), 400

    allowed_exts = [e.strip() for e in job_def.get("accept", ".docx").split(",")]

    def _allowed(filename):
        lower = filename.lower()
        return any(lower.endswith(ext) for ext in allowed_exts)

    if not _allowed(file1_name):
        return jsonify({"error": f"Unsupported file type. Allowed: {', '.join(allowed_exts)}"}), 400
    if file2_name and not _allowed(file2_name):
        return jsonify({"error": f"Second file must be one of: {', '.join(allowed_exts)}"}), 400

    if not is_supported_filename(file1_name):
        return jsonify({"error": f"Unsupported file type. Supported: {SUPPORTED_UPLOADS}"}), 400
    if file2_name and not is_supported_filename(file2_name):
        return jsonify({"error": f"Unsupported second file type. Supported: {SUPPORTED_UPLOADS}"}), 400

    job_id = str(uuid.uuid4())[:8]
    ext1 = Path(file1_name).suffix.lower()
    input_path = os.path.join(UPLOAD_DIR, f"{job_id}_input{ext1}")
    prepared_input1 = os.path.join(UPLOAD_DIR, f"{job_id}_input_prepared.docx")

    if file2_name:
        ext2 = Path(file2_name).suffix.lower()
        input_path2 = os.path.join(UPLOAD_DIR, f"{job_id}_input2{ext2}")
        prepared_input2 = os.path.join(UPLOAD_DIR, f"{job_id}_input2_prepared.docx")
    else:
        input_path2 = None
        prepared_input2 = None

    output_ext = ".csv" if job_type == "gis" else ".docx"
    output_path = os.path.join(UPLOAD_DIR, f"{job_id}_output{output_ext}")

    if chunked_job_id1:
        # Move chunked file to designated input path
        chunk_source = os.path.join(UPLOAD_DIR, f"{chunked_job_id1}_chunked_{secure_filename(chunked_filename1)}")
        if os.path.exists(chunk_source):
            os.rename(chunk_source, input_path)
        else:
            return jsonify({"error": "Chunked upload file 1 not found"}), 400
    else:
        file1.save(input_path)
        
    if file2_name and input_path2:
        if chunked_job_id2:
            chunk_source2 = os.path.join(UPLOAD_DIR, f"{chunked_job_id2}_chunked_{secure_filename(chunked_filename2)}")
            if os.path.exists(chunk_source2):
                os.rename(chunk_source2, input_path2)
            else:
                return jsonify({"error": "Chunked upload file 2 not found"}), 400
        else:
            file2.save(input_path2)

    process_path1 = input_path
    process_path2 = input_path2
    conversion_notes = []

    try:
        prepared1 = normalize_input_to_docx(input_path, prepared_input1)
        process_path1 = prepared1["path"]
        conversion_notes.append(prepared1.get("note"))

        if file2 and input_path2 and prepared_input2:
            prepared2 = normalize_input_to_docx(input_path2, prepared_input2)
            process_path2 = prepared2["path"]
            conversion_notes.append(prepared2.get("note"))

        params = {}
        for field in job_def.get("fields", []):
            params[field["id"]] = request.form.get(field["id"], field.get("default", ""))

        processor = job_def["processor"]
        if job_def.get("multi_file") and process_path2:
            summary = processor([process_path1, process_path2], output_path, **params)
        else:
            summary = processor(process_path1, output_path, **params)

        summary = _append_conversion_notes(summary, conversion_notes)
        out_filename = _build_download_name(job_type.upper(), file1_name, output_ext)

        return jsonify(
            {
                "success": True,
                "job_id": job_id,
                "summary": summary,
                "filename": out_filename,
                "download_url": f"/download/{job_id}/{quote(out_filename)}",
            }
        )

    except Exception as e:
        import traceback

        traceback.print_exc()
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

    finally:
        for p in [input_path, input_path2, prepared_input1, prepared_input2]:
            if p and os.path.exists(p):
                os.remove(p)


@app.route("/download/<job_id>/<filename>")
def download(job_id, filename):
    if not job_id.isalnum():
        return jsonify({"error": "Invalid job ID"}), 400

    safe_filename = secure_filename(filename) or f"{job_id}.docx"
    file_ext = ".csv" if safe_filename.lower().endswith(".csv") else ".docx"
    if not safe_filename.lower().endswith((".docx", ".csv")):
        safe_filename = safe_filename + file_ext

    output_path = os.path.join(UPLOAD_DIR, f"{job_id}_output{file_ext}")
    if not os.path.exists(output_path):
        return jsonify({"error": "File not found or expired"}), 404

    from flask import make_response, after_this_request

    with open(output_path, "rb") as f:
        data = f.read()

    @after_this_request
    def cleanup(response):
        try:
            if os.path.exists(output_path):
                os.remove(output_path)
        except OSError:
            pass
        return response

    response = make_response(data)
    if file_ext == ".csv":
        response.headers["Content-Type"] = "text/csv"
    else:
        response.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    response.headers["Content-Disposition"] = f'attachment; filename="{safe_filename}"'
    return response


if __name__ == "__main__":
    print("\n  🌐 AGENT ZEE — Multi-Job AI Assistant")
    print("  Open http://localhost:5000 in your browser\n")
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
