#!/usr/bin/env python3
"""
AGENT ZEE — Flask Web UI
Multi-job document processing assistant with drag-and-drop interface.
"""

from urllib.parse import quote
import os
import uuid
import tempfile
from flask import Flask, render_template, request, send_file, jsonify

from formatter.engine import reformat_document
from jobs import JOB_REGISTRY, list_jobs, get_job

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "integration_agent_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/jobs")
def api_jobs():
    """Return JSON list of all available jobs."""
    return jsonify(list_jobs())


@app.route("/upload", methods=["POST"])
def upload():
    """Legacy endpoint — Document Formatting job."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not file.filename.lower().endswith(".docx"):
        return jsonify({"error": "Only .docx files are supported"}), 400

    report_name = request.form.get("report_name", "Report")
    year = request.form.get("year", "2026")
    project_number = request.form.get("project_number", "PRJ-001")

    job_id = str(uuid.uuid4())[:8]
    input_path = os.path.join(UPLOAD_DIR, f"{job_id}_input.docx")
    output_path = os.path.join(UPLOAD_DIR, f"{job_id}_output.docx")

    file.save(input_path)

    try:
        summary = reformat_document(
            input_path=input_path,
            output_path=output_path,
            report_name=report_name,
            year=year,
            project_number=project_number,
        )

        out_filename = f"FORMATTED_{file.filename}"

        return jsonify({
            "success": True,
            "job_id": job_id,
            "summary": summary,
            "filename": out_filename,
            "download_url": f"/download/{job_id}/{quote(out_filename)}",
        })

    except Exception as e:
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

    finally:
        if os.path.exists(input_path):
            os.remove(input_path)


@app.route("/process/<job_type>", methods=["POST"])
def process_job(job_type):
    """Universal endpoint for all job types."""
    job_def = get_job(job_type)
    if not job_def:
        return jsonify({"error": f"Unknown job type: {job_type}"}), 400

    # Handle formatting job via legacy path
    if job_type == "formatting":
        return upload()

    # File handling
    if job_def.get("multi_file"):
        if "file" not in request.files or "file2" not in request.files:
            return jsonify({"error": "Two files required for comparison"}), 400
        file1 = request.files["file"]
        file2 = request.files["file2"]
        if not file1.filename or not file2.filename:
            return jsonify({"error": "Both files must be selected"}), 400
    else:
        if "file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400
        file1 = request.files["file"]
        file2 = None
        if file1.filename == "":
            return jsonify({"error": "No file selected"}), 400

    # Validate file extension
    if not file1.filename.lower().endswith(".docx"):
        return jsonify({"error": "Only .docx files are supported"}), 400
    if file2 and not file2.filename.lower().endswith(".docx"):
        return jsonify({"error": "Second file must also be .docx"}), 400

    job_id = str(uuid.uuid4())[:8]
    input_path = os.path.join(UPLOAD_DIR, f"{job_id}_input.docx")
    input_path2 = os.path.join(UPLOAD_DIR, f"{job_id}_input2.docx") if file2 else None
    output_path = os.path.join(UPLOAD_DIR, f"{job_id}_output.docx")

    file1.save(input_path)
    if file2:
        file2.save(input_path2)

    try:
        # Collect form parameters
        params = {}
        for field in job_def.get("fields", []):
            val = request.form.get(field["id"], field.get("default", ""))
            params[field["id"]] = val

        processor = job_def["processor"]

        if job_def.get("multi_file") and input_path2:
            summary = processor([input_path, input_path2], output_path, **params)
        else:
            summary = processor(input_path, output_path, **params)

        prefix = job_type.upper()
        out_filename = f"{prefix}_{file1.filename}"

        return jsonify({
            "success": True,
            "job_id": job_id,
            "summary": summary,
            "filename": out_filename,
            "download_url": f"/download/{job_id}/{quote(out_filename)}",
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

    finally:
        for p in [input_path, input_path2]:
            if p and os.path.exists(p):
                os.remove(p)


@app.route("/download/<job_id>/<filename>")
def download(job_id, filename):
    if not job_id.isalnum():
        return jsonify({"error": "Invalid job ID"}), 400

    output_path = os.path.join(UPLOAD_DIR, f"{job_id}_output.docx")
    if not os.path.exists(output_path):
        return jsonify({"error": "File not found or expired"}), 404

    if not filename.lower().endswith(".docx"):
        filename = filename + ".docx"

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
    response.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    response.headers["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


if __name__ == "__main__":
    print("\n  🌐 AGENT ZEE — Multi-Job AI Assistant")
    print("  Open http://localhost:5000 in your browser\n")
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
