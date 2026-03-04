"""
INTEGRATION Energy Plus — GIS Layer Definitions
Predefined energy layer configurations with INTEGRATION brand styling.
"""

# ─── INTEGRATION brand colours ────────────────────────────────────────
GREEN_HAZE = "#009959"
GREEN_DARK = "#007a47"
GREEN_LIGHT = "#00b86b"
BLUE_HYDRO = "#2196F3"
YELLOW_SOLAR = "#FFC107"
RED_WIND = "#E91E63"
ORANGE_GRID = "#FF5722"
GREY_BOUNDARY = "#78909C"

# ─── Layer definitions ────────────────────────────────────────────────
ENERGY_LAYERS = {
    "hydropower": {
        "name": "Hydropower Plants",
        "icon": "tint",
        "color": BLUE_HYDRO,
        "prefix": "fa",
        "description": "Existing and planned hydropower stations",
        "default_on": True,
    },
    "solar": {
        "name": "Solar Potential Zones",
        "icon": "sun-o",
        "color": YELLOW_SOLAR,
        "prefix": "fa",
        "description": "High solar irradiance zones suitable for PV",
        "default_on": True,
    },
    "wind": {
        "name": "Wind Corridors",
        "icon": "flag",
        "color": RED_WIND,
        "prefix": "fa",
        "description": "Identified wind energy corridors",
        "default_on": True,
    },
    "grid": {
        "name": "Transmission Grid",
        "icon": "bolt",
        "color": ORANGE_GRID,
        "prefix": "fa",
        "description": "Existing and proposed transmission lines",
        "default_on": False,
    },
    "boundaries": {
        "name": "District Boundaries",
        "icon": "map",
        "color": GREY_BOUNDARY,
        "prefix": "fa",
        "description": "Administrative district boundaries of GB",
        "default_on": False,
    },
}

# ─── Map tile providers ──────────────────────────────────────────────
TILE_LAYERS = {
    "default": {
        "name": "OpenStreetMap",
        "tiles": "OpenStreetMap",
    },
    "satellite": {
        "name": "Satellite",
        "tiles": "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        "attr": "Esri",
    },
    "terrain": {
        "name": "Terrain",
        "tiles": "https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png",
        "attr": "OpenTopoMap",
    },
}


def get_marker_icon_options(layer_key: str) -> dict:
    """Return Folium Icon kwargs for a given energy layer."""
    layer = ENERGY_LAYERS.get(layer_key, {})
    return {
        "icon": layer.get("icon", "info-sign"),
        "prefix": layer.get("prefix", "glyphicon"),
        "icon_color": "white",
        "color": _folium_color(layer.get("color", GREEN_HAZE)),
    }


def _folium_color(hex_color: str) -> str:
    """Map hex colours to Folium's named colour palette (closest match)."""
    mapping = {
        BLUE_HYDRO: "blue",
        YELLOW_SOLAR: "orange",
        RED_WIND: "red",
        ORANGE_GRID: "darkred",
        GREY_BOUNDARY: "gray",
        GREEN_HAZE: "green",
        GREEN_DARK: "darkgreen",
        GREEN_LIGHT: "lightgreen",
    }
    return mapping.get(hex_color, "green")
