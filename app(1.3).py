# -*- coding: utf-8 -*-

"""
Rasuwa 2026 Flood WASH Map
Streamlit + PyDeck

This version is designed for Streamlit Cloud.
External config.py and geopandas are NOT required.
"""

import streamlit as st
import pandas as pd
import pydeck as pdk


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Rasuwa 2026 Flood WASH Map",
    page_icon="🌊",
    layout="wide",
)


# ============================================================
# DEFAULT CONFIGURATION
# ============================================================

# Default map location
DEFAULT_LAT = 27.7172
DEFAULT_LON = 85.3240
DEFAULT_ZOOM = 8


# Holding Centre colours
HOLDING_CENTRE_COLORS = {
    "Active": [214, 39, 40, 220],
    "Open": [214, 39, 40, 220],
    "Closed": [128, 128, 128, 200],
    "Planned": [255, 165, 0, 220],
    "Temporary": [255, 221, 87, 220],
}

DEFAULT_HOLDING_CENTRE_COLOR = [
    128,
    128,
    128,
    200,
]


# ============================================================
# DISTRICT LAYER
# ============================================================

def build_district_layer(features):
    """
    Create District polygon layer.

    Priority:
        High   -> Red
        Medium -> Orange
        Low    -> Yellow
        Other  -> Gray
    """

    if features is None:
        features = []

    return pdk.Layer(
        "GeoJsonLayer",

        {
            "type": "FeatureCollection",
            "features": features,
        },

        pickable=True,
        auto_highlight=True,
        stroked=True,
        filled=True,

        get_fill_color="""
            properties.Priority == 'High'
                ? [214, 39, 40, 190]
                : properties.Priority == 'Medium'
                    ? [255, 165, 0, 180]
                    : properties.Priority == 'Low'
                        ? [255, 221, 87, 170]
                        : [200, 200, 200, 120]
        """,

        get_line_color=[
            60,
            60,
            60,
            200,
        ],

        line_width_min_pixels=1,

        highlight_color=[
            0,
            120,
            255,
            120,
        ],

        id="district-layer",
    )


# ============================================================
# HOLDING CENTRE LAYER
# ============================================================

def build_holding_centre_layer(df):
    """
    Create Holding Centre point layer.

    Only records with valid Latitude and Longitude
    are displayed.
    """

    records = []

    # --------------------------------------------------------
    # No data
    # --------------------------------------------------------

    if df is None:
        return pdk.Layer(
            "ScatterplotLayer",
            [],
            get_position=[
                "Longitude",
                "Latitude",
            ],
            get_fill_color=[
                128,
                128,
                128,
                200,
            ],
            get_radius=250,
            radius_min_pixels=5,
            radius_max_pixels=20,
            pickable=True,
            stroked=True,
            id="holding-centre-layer",
        )

    if df.empty:
        return pdk.Layer(
            "ScatterplotLayer",
            [],
            get_position=[
                "Longitude",
                "Latitude",
            ],
            get_fill_color=[
                128,
                128,
                128,
                200,
            ],
            get_radius=250,
            radius_min_pixels=5,
            radius_max_pixels=20,
            pickable=True,
            stroked=True,
            id="holding-centre-layer",
        )

    # --------------------------------------------------------
    # Check columns
    # --------------------------------------------------------

    required_columns = [
        "Latitude",
        "Longitude",
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing_columns:

        return pdk.Layer(
            "ScatterplotLayer",
            [],
            get_position=[
                "Longitude",
                "Latitude",
            ],
            get_fill_color=[
                128,
                128,
                128,
                200,
            ],
            get_radius=250,
            radius_min_pixels=5,
            radius_max_pixels=20,
            pickable=True,
            stroked=True,
            id="holding-centre-layer",
        )

    # --------------------------------------------------------
    # Prepare data
    # --------------------------------------------------------

    try:

        plot_df = df.copy()

        # Convert coordinates to numeric
        plot_df["Latitude"] = pd.to_numeric(
            plot_df["Latitude"],
            errors="coerce",
        )

        plot_df["Longitude"] = pd.to_numeric(
            plot_df["Longitude"],
            errors="coerce",
        )

        # Remove rows without coordinates
        plot_df = plot_df.dropna(
            subset=[
                "Latitude",
                "Longitude",
            ]
        )

        # ----------------------------------------------------
        # Status colour
        # ----------------------------------------------------

        if "Status" in plot_df.columns:

            plot_df["radius_color"] = (
                plot_df["Status"]
                .map(HOLDING_CENTRE_COLORS)
                .apply(
                    lambda color:
                    color
                    if isinstance(color, list)
                    else DEFAULT_HOLDING_CENTRE_COLOR
                )
            )

        else:

            plot_df["radius_color"] = [
                DEFAULT_HOLDING_CENTRE_COLOR
            ] * len(plot_df)

        # Convert to records
        records = plot_df.to_dict(
            orient="records"
        )

    except Exception:

        records = []

    # --------------------------------------------------------
    # Create layer
    # --------------------------------------------------------

    return pdk.Layer(
        "ScatterplotLayer",

        records,

        get_position=[
            "Longitude",
            "Latitude",
        ],

        get_fill_color="radius_color",

        get_radius=250,

        radius_min_pixels=5,
        radius_max_pixels=20,

        pickable=True,
        stroked=True,

        get_line_color=[
            255,
            255,
            255,
            220,
        ],

        line_width_min_pixels=1,

        id="holding-centre-layer",
    )


# ============================================================
# VIEW STATE
# ============================================================

def build_view_state(gdf=None):
    """
    Create initial map view.

    GeoPandas is not required.
    """

    # --------------------------------------------------------
    # No boundary data
    # --------------------------------------------------------

    if gdf is None:

        return pdk.ViewState(
            latitude=DEFAULT_LAT,
            longitude=DEFAULT_LON,
            zoom=DEFAULT_ZOOM,
            pitch=0,
        )

    # --------------------------------------------------------
    # Empty data
    # --------------------------------------------------------

    try:

        if gdf.empty:

            return pdk.ViewState(
                latitude=DEFAULT_LAT,
                longitude=DEFAULT_LON,
                zoom=DEFAULT_ZOOM,
                pitch=0,
            )

    except Exception:

        return pdk.ViewState(
            latitude=DEFAULT_LAT,
            longitude=DEFAULT_LON,
            zoom=DEFAULT_ZOOM,
            pitch=0,
        )

    # --------------------------------------------------------
    # Bounding box
    # --------------------------------------------------------

    try:

        minx, miny, maxx, maxy = gdf.total_bounds

        latitude = (
            float(miny) +
            float(maxy)
        ) / 2

        longitude = (
            float(minx) +
            float(maxx)
        ) / 2

        return pdk.ViewState(
            latitude=latitude,
            longitude=longitude,
            zoom=DEFAULT_ZOOM,
            pitch=0,
        )

    except Exception:

        return pdk.ViewState(
            latitude=DEFAULT_LAT,
            longitude=DEFAULT_LON,
            zoom=DEFAULT_ZOOM,
            pitch=0,
        )


# ============================================================
# DECK MAP
# ============================================================

def build_deck(
    features=None,
    gdf=None,
    extra_layers=None,
):
    """
    Build final PyDeck map.
    """

    # --------------------------------------------------------
    # District layer
    # --------------------------------------------------------

    layers = [
        build_district_layer(
            features
        )
    ]

    # --------------------------------------------------------
    # Additional layers
    # --------------------------------------------------------

    if extra_layers:

        for layer in extra_layers:

            if layer is not None:

                layers.append(layer)

    # --------------------------------------------------------
    # View state
    # --------------------------------------------------------

    view_state = build_view_state(
        gdf
    )

    # --------------------------------------------------------
    # Tooltip
    # --------------------------------------------------------

    tooltip = {
        "html": """
            <b>{District}</b>
        """,

        "style": {
            "backgroundColor": "#1F4E78",
            "color": "white",
        },
    }

    # --------------------------------------------------------
    # Deck
    # --------------------------------------------------------

    return pdk.Deck(

        layers=layers,

        initial_view_state=view_state,

        tooltip=tooltip,

        map_style="light",
    )


# ============================================================
# STREAMLIT APP
# ============================================================

st.title(
    "🌊 Rasuwa 2026 Flood WASH Map"
)

st.markdown(
    """
    **UNICEF WASH Emergency Response**
    
    Interactive map for visualising district priority
    and temporary Holding Centre information.
    """
)


# ============================================================
# INFORMATION
# ============================================================

with st.expander(
    "ℹ️ Map information",
    expanded=False,
):

    st.write(
        """
        The map displays:

        • District polygons according to Priority  
        • Holding Centre locations  
        • Holding Centre status  
        • Interactive map information
        """
    )


# ============================================================
# DEMO / EMPTY MAP
# ============================================================

# The application can run even when no GeoJSON
# or Holding Centre data is available.

features = []

gdf = None

extra_layers = []


# ============================================================
# BUILD MAP
# ============================================================

try:

    deck = build_deck(
        features=features,
        gdf=gdf,
        extra_layers=extra_layers,
    )

    st.pydeck_chart(
        deck,
        use_container_width=True,
    )

except Exception as e:

    st.error(
        "The map could not be loaded."
    )

    st.code(
        str(e)
    )


# ============================================================
# FOOTER
# ============================================================

st.caption(
    "UNICEF Nepal | WASH Emergency Response | Rasuwa 2026"
)