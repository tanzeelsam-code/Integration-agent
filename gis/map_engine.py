"""
INTEGRATION Energy Plus — GIS Map Engine
Generate interactive Folium/Leaflet maps for Gilgit-Baltistan
energy infrastructure visualization.
"""

import folium
from folium.plugins import MarkerCluster, Fullscreen, MiniMap
from typing import Optional

from .layers import (
    ENERGY_LAYERS, TILE_LAYERS,
    get_marker_icon_options,
    GREEN_HAZE, BLUE_HYDRO, YELLOW_SOLAR, RED_WIND, ORANGE_GRID, GREY_BOUNDARY,
)
from .data_loader import load_default_data, filter_features_by_layer, load_geojson_from_string


# ─── Gilgit-Baltistan centre coordinates ─────────────────────────────
GB_CENTER = [35.8, 74.5]
GB_ZOOM = 7


def create_base_map(
    center: list = None,
    zoom: int = None,
    tile_key: str = "default",
) -> folium.Map:
    """
    Create a base Folium map centred on Gilgit-Baltistan.

    Args:
        center: [lat, lon] centre point (default: GB centre).
        zoom: Initial zoom level (default: 7).
        tile_key: Tile provider key from TILE_LAYERS.

    Returns:
        A folium.Map instance.
    """
    center = center or GB_CENTER
    zoom = zoom or GB_ZOOM

    tile_cfg = TILE_LAYERS.get(tile_key, TILE_LAYERS["default"])

    if tile_key == "default":
        m = folium.Map(location=center, zoom_start=zoom, tiles="OpenStreetMap")
    else:
        m = folium.Map(
            location=center,
            zoom_start=zoom,
            tiles=tile_cfg["tiles"],
            attr=tile_cfg.get("attr", ""),
        )

    # Add all tile layers as switchable options
    for key, cfg in TILE_LAYERS.items():
        if key == tile_key:
            continue
        if key == "default":
            folium.TileLayer("OpenStreetMap", name=cfg["name"]).add_to(m)
        else:
            folium.TileLayer(
                tiles=cfg["tiles"],
                attr=cfg.get("attr", ""),
                name=cfg["name"],
            ).add_to(m)

    # Add helpful controls
    Fullscreen(position="topleft").add_to(m)
    MiniMap(toggle_display=True, position="bottomright").add_to(m)

    return m


def _build_popup_html(props: dict) -> str:
    """Build a rich HTML popup for a map feature."""
    name = props.get("name", "Unknown")
    layer = props.get("layer", "")
    status = props.get("status", "")
    capacity = props.get("capacity_mw")
    voltage = props.get("voltage_kv")
    desc = props.get("description", "")
    district = props.get("district", "")
    population = props.get("population")
    area = props.get("area_km2")

    layer_cfg = ENERGY_LAYERS.get(layer, {})
    color = layer_cfg.get("color", GREEN_HAZE)

    html = f"""
    <div style="font-family: Arial, sans-serif; min-width: 220px; max-width: 300px;">
        <h4 style="margin:0 0 8px; color:{color}; font-size:14px; border-bottom:2px solid {color}; padding-bottom:4px;">
            {name}
        </h4>
    """

    if status:
        badge_bg = "#e8f5e9" if status == "Operational" else "#fff3e0" if status in ("Planned", "Proposed") else "#e3f2fd"
        badge_color = "#2e7d32" if status == "Operational" else "#e65100" if status in ("Planned", "Proposed") else "#1565c0"
        html += f"""
        <span style="background:{badge_bg}; color:{badge_color}; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600;">
            {status}
        </span><br><br>
        """

    if capacity:
        html += f"<b>Capacity:</b> {capacity} MW<br>"
    if voltage:
        html += f"<b>Voltage:</b> {voltage} kV<br>"
    if district:
        html += f"<b>District:</b> {district}<br>"
    if population:
        html += f"<b>Population:</b> {population:,}<br>"
    if area:
        html += f"<b>Area:</b> {area:,} km²<br>"
    if desc:
        html += f"<p style='margin:8px 0 0; font-size:12px; color:#555;'>{desc}</p>"

    html += "</div>"
    return html


def add_point_layer(
    m: folium.Map,
    data: dict,
    layer_key: str,
    use_clusters: bool = True,
) -> folium.Map:
    """
    Add point features for a given layer to the map.

    Args:
        m: Folium Map instance.
        data: GeoJSON FeatureCollection.
        layer_key: Energy layer key (e.g. 'hydropower').
        use_clusters: Whether to cluster markers.

    Returns:
        The modified map.
    """
    filtered = filter_features_by_layer(data, layer_key)
    features = filtered.get("features", [])

    layer_cfg = ENERGY_LAYERS.get(layer_key, {})
    layer_name = layer_cfg.get("name", layer_key.title())

    # Create a feature group for this layer
    fg = folium.FeatureGroup(name=layer_name, show=layer_cfg.get("default_on", True))

    if use_clusters:
        cluster = MarkerCluster()
    else:
        cluster = fg

    for feat in features:
        geom = feat.get("geometry", {})
        props = feat.get("properties", {})

        if geom.get("type") == "Point":
            coords = geom.get("coordinates", [])
            if len(coords) >= 2:
                icon_opts = get_marker_icon_options(layer_key)
                marker = folium.Marker(
                    location=[coords[1], coords[0]],  # GeoJSON is [lon, lat]
                    popup=folium.Popup(_build_popup_html(props), max_width=320),
                    tooltip=props.get("name", ""),
                    icon=folium.Icon(**icon_opts),
                )
                marker.add_to(cluster)

    if use_clusters:
        cluster.add_to(fg)

    fg.add_to(m)
    return m


def add_polygon_layer(
    m: folium.Map,
    data: dict,
    layer_key: str,
) -> folium.Map:
    """Add polygon features (zones, boundaries) for a given layer."""
    filtered = filter_features_by_layer(data, layer_key)
    features = filtered.get("features", [])

    layer_cfg = ENERGY_LAYERS.get(layer_key, {})
    layer_name = layer_cfg.get("name", layer_key.title())
    color = layer_cfg.get("color", GREEN_HAZE)

    fg = folium.FeatureGroup(name=layer_name, show=layer_cfg.get("default_on", True))

    for feat in features:
        geom = feat.get("geometry", {})
        props = feat.get("properties", {})

        if geom.get("type") == "Polygon":
            # Folium expects [ [lat, lon], ... ] — GeoJSON is [ [lon, lat], ... ]
            coords = geom.get("coordinates", [[]])[0]
            latlon = [[c[1], c[0]] for c in coords]

            folium.Polygon(
                locations=latlon,
                color=color,
                weight=2,
                fill=True,
                fill_color=color,
                fill_opacity=0.15,
                popup=folium.Popup(_build_popup_html(props), max_width=320),
                tooltip=props.get("name", ""),
            ).add_to(fg)

    fg.add_to(m)
    return m


def add_line_layer(
    m: folium.Map,
    data: dict,
    layer_key: str,
) -> folium.Map:
    """Add line features (corridors, transmission lines) for a given layer."""
    filtered = filter_features_by_layer(data, layer_key)
    features = filtered.get("features", [])

    layer_cfg = ENERGY_LAYERS.get(layer_key, {})
    layer_name = layer_cfg.get("name", layer_key.title())
    color = layer_cfg.get("color", GREEN_HAZE)

    fg = folium.FeatureGroup(name=layer_name, show=layer_cfg.get("default_on", True))

    for feat in features:
        geom = feat.get("geometry", {})
        props = feat.get("properties", {})

        if geom.get("type") == "LineString":
            coords = geom.get("coordinates", [])
            latlon = [[c[1], c[0]] for c in coords]

            folium.PolyLine(
                locations=latlon,
                color=color,
                weight=3,
                opacity=0.8,
                popup=folium.Popup(_build_popup_html(props), max_width=320),
                tooltip=props.get("name", ""),
                dash_array="10" if props.get("status") in ("Planned", "Proposed") else None,
            ).add_to(fg)

    fg.add_to(m)
    return m


def generate_full_map(
    data: dict = None,
    active_layers: list = None,
    center: list = None,
    zoom: int = None,
) -> folium.Map:
    """
    Generate a complete interactive map with all energy layers.

    Args:
        data: GeoJSON FeatureCollection. Uses default GB data if None.
        active_layers: List of layer keys to include. Uses all if None.
        center: [lat, lon] map centre.
        zoom: Initial zoom level.

    Returns:
        A fully-configured folium.Map.
    """
    if data is None:
        data = load_default_data()

    if active_layers is None:
        active_layers = list(ENERGY_LAYERS.keys())

    m = create_base_map(center=center, zoom=zoom)

    # Determine geometry types per layer and add accordingly
    for layer_key in active_layers:
        if layer_key not in ENERGY_LAYERS:
            continue

        filtered = filter_features_by_layer(data, layer_key)
        features = filtered.get("features", [])

        geom_types = {f.get("geometry", {}).get("type") for f in features}

        if "Point" in geom_types:
            add_point_layer(m, data, layer_key)
        if "Polygon" in geom_types:
            add_polygon_layer(m, data, layer_key)
        if "LineString" in geom_types:
            add_line_layer(m, data, layer_key)

    # Add layer control toggle
    folium.LayerControl(collapsed=False).add_to(m)

    return m


def export_map_html(m: folium.Map, output_path: str = None) -> str:
    """
    Export a Folium map to a standalone HTML string or file.

    Args:
        m: Folium Map instance.
        output_path: If given, save to this file path.

    Returns:
        The HTML string of the map.
    """
    html = m._repr_html_()

    if output_path:
        m.save(output_path)
        with open(output_path, "r", encoding="utf-8") as f:
            html = f.read()

    return html


def get_map_iframe_html(m: folium.Map, width: str = "100%", height: str = "600px") -> str:
    """Return the map as an embeddable HTML fragment (no iframe)."""
    return m._repr_html_()
