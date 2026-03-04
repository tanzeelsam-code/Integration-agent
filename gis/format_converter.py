"""
INTEGRATION Energy Plus — GIS Format Converter
Convert various GIS formats to GeoJSON for map display.

Supported formats:
  - GeoJSON (.geojson, .json)    — passthrough
  - KML / KMZ (.kml, .kmz)      — via fiona/geopandas
  - Shapefile (.zip containing .shp) — via geopandas
  - CSV with lat/lon columns (.csv) — pandas → GeoJSON
  - GPX tracks/waypoints (.gpx)  — via geopandas
"""

import json
import os
import io
import csv
import tempfile
import zipfile
from typing import Tuple

# Lazy imports to avoid hard failures if optional deps missing
_GEOPANDAS_AVAILABLE = None


def _check_geopandas():
    global _GEOPANDAS_AVAILABLE
    if _GEOPANDAS_AVAILABLE is None:
        try:
            import geopandas  # noqa: F401
            _GEOPANDAS_AVAILABLE = True
        except ImportError:
            _GEOPANDAS_AVAILABLE = False
    return _GEOPANDAS_AVAILABLE


# ─── Supported extensions ─────────────────────────────────────────────
SUPPORTED_GIS_EXTENSIONS = {
    ".geojson": "GeoJSON",
    ".json": "GeoJSON",
    ".kml": "KML",
    ".kmz": "KMZ",
    ".zip": "Shapefile (ZIP)",
    ".csv": "CSV (lat/lon)",
    ".gpx": "GPX",
}


def supported_gis_formats_csv() -> str:
    """Return comma-separated list of accepted extensions."""
    return ",".join(sorted(SUPPORTED_GIS_EXTENSIONS.keys()))


def is_supported_gis_file(filename: str) -> bool:
    """Check if a filename has a supported GIS extension."""
    ext = os.path.splitext(filename.lower())[1]
    return ext in SUPPORTED_GIS_EXTENSIONS


def convert_to_geojson(file_bytes: bytes, filename: str) -> Tuple[dict, str]:
    """
    Convert uploaded file bytes to a GeoJSON dict.

    Args:
        file_bytes: Raw file content.
        filename: Original filename (used to detect format).

    Returns:
        Tuple of (geojson_dict, format_name).

    Raises:
        ValueError: If the format is unsupported or conversion fails.
    """
    ext = os.path.splitext(filename.lower())[1]
    fmt = SUPPORTED_GIS_EXTENSIONS.get(ext)

    if not fmt:
        raise ValueError(
            f"Unsupported format '{ext}'. "
            f"Supported: {', '.join(sorted(SUPPORTED_GIS_EXTENSIONS.keys()))}"
        )

    if ext in (".geojson", ".json"):
        return _convert_geojson(file_bytes), fmt

    if ext == ".csv":
        return _convert_csv(file_bytes), fmt

    if ext == ".kml":
        return _convert_kml(file_bytes), fmt

    if ext == ".kmz":
        return _convert_kmz(file_bytes), fmt

    if ext == ".zip":
        return _convert_shapefile_zip(file_bytes), fmt

    if ext == ".gpx":
        return _convert_gpx(file_bytes), fmt

    raise ValueError(f"Conversion for '{ext}' not implemented")


# ─── Format-specific converters ──────────────────────────────────────

def _convert_geojson(raw: bytes) -> dict:
    """Parse GeoJSON bytes."""
    try:
        data = json.loads(raw.decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError) as e:
        raise ValueError(f"Invalid GeoJSON: {e}")

    if not isinstance(data, dict) or "type" not in data:
        raise ValueError("Not a valid GeoJSON object")
    return data


def _convert_csv(raw: bytes) -> dict:
    """
    Convert a CSV with latitude/longitude columns to GeoJSON points.
    Auto-detects common column names: lat/latitude/y, lon/lng/longitude/x.
    """
    text = raw.decode("utf-8")
    reader = csv.DictReader(io.StringIO(text))

    if not reader.fieldnames:
        raise ValueError("CSV has no headers")

    # Find lat/lon columns
    headers_lower = {h.lower().strip(): h for h in reader.fieldnames}
    lat_col = None
    lon_col = None

    for name in ["latitude", "lat", "y"]:
        if name in headers_lower:
            lat_col = headers_lower[name]
            break

    for name in ["longitude", "lon", "lng", "long", "x"]:
        if name in headers_lower:
            lon_col = headers_lower[name]
            break

    if not lat_col or not lon_col:
        raise ValueError(
            f"CSV must have lat/lon columns. Found headers: {list(reader.fieldnames)}. "
            f"Expected: latitude/lat/y and longitude/lon/lng/x"
        )

    features = []
    for i, row in enumerate(reader):
        try:
            lat = float(row[lat_col])
            lon = float(row[lon_col])
        except (ValueError, TypeError):
            continue  # Skip rows with invalid coordinates

        # All other columns become properties
        props = {k: v for k, v in row.items() if k not in (lat_col, lon_col)}
        props["_row"] = i + 1

        # Try to detect a name column
        name = props.get("name") or props.get("Name") or props.get("NAME") or f"Point {i+1}"

        features.append({
            "type": "Feature",
            "properties": {**props, "name": name, "layer": "uploaded"},
            "geometry": {"type": "Point", "coordinates": [lon, lat]},
        })

    if not features:
        raise ValueError("No valid coordinate rows found in CSV")

    return {"type": "FeatureCollection", "features": features}


def _convert_kml(raw: bytes) -> dict:
    """Convert KML bytes to GeoJSON using geopandas/fiona."""
    if not _check_geopandas():
        raise ValueError("KML support requires geopandas. Install with: pip install geopandas")

    import geopandas as gpd

    with tempfile.NamedTemporaryFile(suffix=".kml", delete=False) as tmp:
        tmp.write(raw)
        tmp_path = tmp.name

    try:
        gdf = gpd.read_file(tmp_path, driver="KML")
        return json.loads(gdf.to_json())
    except Exception as e:
        raise ValueError(f"Failed to parse KML: {e}")
    finally:
        os.unlink(tmp_path)


def _convert_kmz(raw: bytes) -> dict:
    """Extract KML from KMZ (ZIP) and convert to GeoJSON."""
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            kml_files = [n for n in zf.namelist() if n.lower().endswith(".kml")]
            if not kml_files:
                raise ValueError("KMZ archive contains no .kml files")
            kml_bytes = zf.read(kml_files[0])
    except zipfile.BadZipFile:
        raise ValueError("Invalid KMZ file (not a valid ZIP archive)")

    return _convert_kml(kml_bytes)


def _convert_shapefile_zip(raw: bytes) -> dict:
    """Convert a zipped Shapefile to GeoJSON using geopandas."""
    if not _check_geopandas():
        raise ValueError("Shapefile support requires geopandas. Install with: pip install geopandas")

    import geopandas as gpd

    # Validate ZIP
    try:
        zf = zipfile.ZipFile(io.BytesIO(raw))
    except zipfile.BadZipFile:
        raise ValueError("Invalid ZIP file")

    shp_files = [n for n in zf.namelist() if n.lower().endswith(".shp")]
    if not shp_files:
        raise ValueError("ZIP must contain a .shp file")

    zf.close()

    # Write to temp and read with geopandas
    with tempfile.NamedTemporaryFile(suffix=".zip", delete=False) as tmp:
        tmp.write(raw)
        tmp_path = tmp.name

    try:
        gdf = gpd.read_file(f"zip://{tmp_path}")
        return json.loads(gdf.to_json())
    except Exception as e:
        raise ValueError(f"Failed to parse Shapefile: {e}")
    finally:
        os.unlink(tmp_path)


def _convert_gpx(raw: bytes) -> dict:
    """Convert GPX file to GeoJSON using geopandas."""
    if not _check_geopandas():
        raise ValueError("GPX support requires geopandas. Install with: pip install geopandas")

    import geopandas as gpd

    with tempfile.NamedTemporaryFile(suffix=".gpx", delete=False) as tmp:
        tmp.write(raw)
        tmp_path = tmp.name

    try:
        # GPX has multiple layers — try waypoints first, then tracks
        features = []
        for layer_name in ["waypoints", "tracks", "track_points", "routes"]:
            try:
                gdf = gpd.read_file(tmp_path, layer=layer_name)
                if len(gdf) > 0:
                    fc = json.loads(gdf.to_json())
                    for f in fc.get("features", []):
                        f["properties"]["_gpx_layer"] = layer_name
                    features.extend(fc.get("features", []))
            except Exception:
                continue

        if not features:
            raise ValueError("No waypoints or tracks found in GPX file")

        return {"type": "FeatureCollection", "features": features}
    except ValueError:
        raise
    except Exception as e:
        raise ValueError(f"Failed to parse GPX: {e}")
    finally:
        os.unlink(tmp_path)
