#!/usr/bin/env python3
"""
Formating AI Assistance — Flask Web UI
Drag-and-drop interface for reformatting .docx files + GIS mapping.

Usage:
    python app.py
    Open http://localhost:5000 in your browser
"""

from urllib.parse import quote

import os
import uuid
import json
import tempfile
from flask import Flask, render_template, request, send_file, jsonify, make_response

from formatter.engine import reformat_document
from gis.map_engine import generate_full_map, export_map_html
from gis.data_loader import (
    load_default_data, load_geojson_from_string,
    get_feature_summary, get_available_layers,
)
from gis.layers import ENERGY_LAYERS

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "integration_agent_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Shared GIS state (per-process; fine for single-server)
_custom_geojson_data = None  # Holds uploaded GeoJSON, if any


@app.route("/")
def index():
    return render_template("index.html")


# GIS MAP ROUTES

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
        total = sum(
            summary.get(k, 0) for k in active_layers
        )

        return jsonify({
            "map_html": map_html,
            "total_features": total,
            "layer_summary": summary,
        })
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

        # Merge with default data
        default = load_default_data()
        merged_features = default.get("features", []) + data.get("features", [])
        _custom_geojson_data = {
            "type": "FeatureCollection",
            "features": merged_features,
        }

        return jsonify({
            "success": True,
            "features_count": len(data.get("features", [])),
            "total_features": len(merged_features),
            "layers": get_available_layers(_custom_geojson_data),
        })
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


# DOCUMENT FORMATTER ROUTES

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
    output_path = os.path.join(UPLOAD_DIR, f"{job_id}_output.docx")
    if not os.path.exists(output_path):
        return jsonify({"error": "File not found or expired"}), 404

    if not filename.lower().endswith(".docx"):
        filename = filename + ".docx"

    with open(output_path, "rb") as f:
        data = f.read()

    response = make_response(data)
    response.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    response.headers["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


if __name__ == "__main__":
    print("\n  🌐 Formating AI Assistance Web UI")
    print("  Open http://localhost:5000 in your browser\n")
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
