"""
INTEGRATION Energy Plus — GIS Data Loader
Load and validate GeoJSON / JSON spatial data.
"""

import json
import os
from typing import Optional


_DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
_DEFAULT_FILE = os.path.join(_DATA_DIR, "gb_energy.geojson")


def load_geojson(path: str) -> dict:
    """
    Load a GeoJSON file and return it as a Python dict.

    Args:
        path: Absolute or relative path to a .geojson / .json file.

    Returns:
        Parsed GeoJSON dict (FeatureCollection or single Feature).

    Raises:
        FileNotFoundError: If the file does not exist.
        ValueError: If the file is not valid GeoJSON.
    """
    if not os.path.isfile(path):
        raise FileNotFoundError(f"GeoJSON file not found: {path}")

    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    _validate_geojson(data)
    return data


def load_default_data() -> dict:
    """Load the bundled Gilgit-Baltistan energy GeoJSON dataset."""
    return load_geojson(_DEFAULT_FILE)


def load_geojson_from_string(raw: str) -> dict:
    """Parse a GeoJSON string (e.g. from an upload) and validate it."""
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON: {e}")

    _validate_geojson(data)
    return data


def filter_features_by_layer(data: dict, layer_key: str) -> dict:
    """
    Return a new FeatureCollection containing only features
    whose 'properties.layer' matches *layer_key*.
    """
    if data.get("type") != "FeatureCollection":
        return data

    filtered = [
        f for f in data.get("features", [])
        if f.get("properties", {}).get("layer") == layer_key
    ]
    return {
        "type": "FeatureCollection",
        "features": filtered,
    }


def get_available_layers(data: dict) -> list[str]:
    """Return a sorted list of unique layer keys present in the data."""
    layers = set()
    for feat in data.get("features", []):
        lyr = feat.get("properties", {}).get("layer")
        if lyr:
            layers.add(lyr)
    return sorted(layers)


def get_feature_summary(data: dict) -> dict:
    """
    Return a summary dict: { layer_key: count } for each layer
    present in the FeatureCollection.
    """
    summary: dict[str, int] = {}
    for feat in data.get("features", []):
        lyr = feat.get("properties", {}).get("layer", "other")
        summary[lyr] = summary.get(lyr, 0) + 1
    return summary


# ─── Private helpers ──────────────────────────────────────────────────

def _validate_geojson(data: dict) -> None:
    """Basic structural validation of a GeoJSON object."""
    if not isinstance(data, dict):
        raise ValueError("GeoJSON root must be a JSON object")

    geo_type = data.get("type")
    valid_types = {
        "FeatureCollection", "Feature", "Point", "MultiPoint",
        "LineString", "MultiLineString", "Polygon", "MultiPolygon",
        "GeometryCollection",
    }
    if geo_type not in valid_types:
        raise ValueError(
            f"Invalid GeoJSON type '{geo_type}'. "
            f"Expected one of: {', '.join(sorted(valid_types))}"
        )

    if geo_type == "FeatureCollection":
        features = data.get("features")
        if not isinstance(features, list):
            raise ValueError("FeatureCollection must have a 'features' array")
