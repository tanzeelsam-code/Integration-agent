#!/usr/bin/env python3
"""
Formating AI Assistance — Flask Web UI
Drag-and-drop interface for reformatting .docx files.

Usage:
    python app.py
    Open http://localhost:5000 in your browser
"""

from urllib.parse import quote

import os
import uuid
import tempfile
from flask import Flask, render_template, request, send_file, jsonify

from formatter.engine import reformat_document

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "integration_agent_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not file.filename.lower().endswith(".docx"):
        return jsonify({"error": "Only .docx files are supported"}), 400

    # Get optional parameters
    report_name = request.form.get("report_name", "Report")
    year = request.form.get("year", "2026")
    project_number = request.form.get("project_number", "PRJ-001")

    # Save uploaded file
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
        # Clean up input file
        if os.path.exists(input_path):
            os.remove(input_path)


@app.route("/download/<job_id>/<filename>")
def download(job_id, filename):
    # Validate job_id to prevent path traversal
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
    print("\n  🌐 Formating AI Assistance Web UI")
    print("  Open http://localhost:5000 in your browser\n")
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
