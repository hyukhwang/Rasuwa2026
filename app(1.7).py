# -*- coding: utf-8 -*-

"""
UNICEF Nepal
Flood WASH Emergency Decision Support Dashboard

Main Streamlit application.

Required repository structure:

rasuwa2026/
│
├── app(1.6).py
├── config.py
├── data_loader.py
├── map_utils.py
├── requirements.txt
│
└── data/
    ├── population.csv
    └── flood.geojson
"""

from pathlib import Path

import pandas as pd
import streamlit as st
import geopandas as gpd

import config
import data_loader
import map_utils


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="UNICEF Nepal - Flood WASH Dashboard",
    page_icon="🌊",
    layout="wide",
)


# ============================================================
# BASE DIRECTORY
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

POPULATION_PATH = DATA_DIR / "population.csv"
FLOOD_GEOJSON_PATH = DATA_DIR / "flood.geojson"


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def file_exists(path):
    """Check whether a file exists."""
    return Path(path).exists()


def format_number(value):
    """Format numeric values."""
    if value is None:
        return "No data"

    try:
        if pd.isna(value):
            return "No data"
    except Exception:
        pass

    try:
        return f"{int(float(value)):,}"
    except (ValueError, TypeError):
        return str(value)


def format_currency(value):
    """Format currency values."""
    if value is None:
        return "No data"

    try:
        if pd.isna(value):
            return "No data"
    except Exception:
        pass

    try:
        return f"USD {int(float(value)):,}"
    except (ValueError, TypeError):
        return str(value)


def format_text(value):
    """Format text values."""
    if value is None:
        return "No data"

    try:
        if pd.isna(value):
            return "No data"
    except Exception:
        pass

    value = str(value).strip()

    if value == "":
        return "No data"

    return value


# ============================================================
# REQUIRED FILE CHECK
# ============================================================

required_files = {
    "population.csv": POPULATION_PATH,
    "flood.geojson": FLOOD_GEOJSON_PATH,
}

missing_files = []

for filename, filepath in required_files.items():
    if not file_exists(filepath):
        missing_files.append(
            f"data/{filename}"
        )


if missing_files:

    st.error(
        "Required data files are missing."
    )

    st.markdown(
        "The following files could not be found:"
    )

    for filename in missing_files:
        st.code(filename)

    st.markdown(
        "Please check the GitHub repository structure."
    )

    st.stop()


# ============================================================
# LOAD POPULATION CSV
# ============================================================

@st.cache_data
def load_population_csv(filepath):
    """
    Load population CSV.
    """
    return pd.read_csv(filepath)


try:

    population_df = load_population_csv(
        POPULATION_PATH
    )

except Exception as e:

    st.error(
        "Failed to load population.csv."
    )

    st.code(
        str(e)
    )

    st.write(
        f"File path: {POPULATION_PATH}"
    )

    st.stop()


# ============================================================
# LOAD FLOOD GEOJSON
# ============================================================

@st.cache_data
def load_flood_geojson(filepath):
    """
    Load GeoJSON using pyogrio.
    """

    return gpd.read_file(
        filepath,
        engine="pyogrio",
    )


try:

    flood_gdf = load_flood_geojson(
        FLOOD_GEOJSON_PATH
    )

except Exception as e:

    st.error(
        "Failed to load flood.geojson."
    )

    st.code(
        str(e)
    )

    st.markdown(
        """
        Please check the following:

        1. flood.geojson exists in the data folder.
        2. pyogrio is included in requirements.txt.
        3. flood.geojson is a valid GeoJSON file.
        """
    )

    st.write(
        f"File path: {FLOOD_GEOJSON_PATH}"
    )

    st.stop()


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title(
    "⚙️ Data Sources"
)

st.sidebar.markdown(
    "### Repository data"
)

st.sidebar.caption(
    "Population data"
)

st.sidebar.code(
    "data/population.csv"
)

st.sidebar.caption(
    "Flood boundary"
)

st.sidebar.code(
    "data/flood.geojson"
)


# ============================================================
# EXCEL SOURCE
# ============================================================

default_excel_path = getattr(
    config,
    "DEFAULT_EXCEL_PATH",
    "",
)

excel_path = st.sidebar.text_input(
    "Excel file path",
    value=str(default_excel_path),
)


# ============================================================
# GEOJSON SOURCE
# ============================================================

geojson_path = st.sidebar.text_input(
    "Flood GeoJSON path",
    value=str(FLOOD_GEOJSON_PATH),
)


# ============================================================
# ADVANCED DATA SOURCES
# ============================================================

with st.sidebar.expander(
    "Advanced data sources",
    expanded=False,
):

    default_event_info = getattr(
        config,
        "DEFAULT_EVENT_INFO_PATH",
        "",
    )

    default_markers = getattr(
        config,
        "DEFAULT_EVENT_MARKERS_PATH",
        "",
    )

    default_corridor = getattr(
        config,
        "DEFAULT_FLOOD_CORRIDOR_PATH",
        "",
    )

    event_info_path = st.text_input(
        "Event information",
        value=str(default_event_info),
    )

    markers_path = st.text_input(
        "Event markers",
        value=str(default_markers),
    )

    corridor_path = st.text_input(
        "Flood corridor",
        value=str(default_corridor),
    )


# ============================================================
# REFRESH BUTTON
# ============================================================

if st.sidebar.button(
    "🔄 Refresh data"
):

    st.cache_data.clear()

    st.rerun()


# ============================================================
# DATA STATUS
# ============================================================

st.sidebar.markdown(
    "---"
)

st.sidebar.caption(
    f"Population records: {len(population_df):,}"
)

st.sidebar.caption(
    f"Flood features: {len(flood_gdf):,}"
)

st.sidebar.caption(
    "GeoJSON engine: pyogrio"
)


# ============================================================
# LOAD EVENT INFORMATION
# ============================================================

try:

    event_info, event_info_error = (
        data_loader.load_event_info(
            event_info_path
        )
    )

except Exception as e:

    event_info = {}
    event_info_error = str(e)


if event_info is None:
    event_info = {}


# ============================================================
# PAGE TITLE
# ============================================================

page_title = event_info.get(
    "Title",
    "UNICEF Nepal - Flood WASH Dashboard",
)

page_subtitle = event_info.get(
    "Subtitle",
    "",
)

st.title(
    page_title
)


if page_subtitle:

    st.markdown(
        f"### {page_subtitle}"
    )


# ============================================================
# EVENT METADATA
# ============================================================

metadata = []

corridor_note = event_info.get(
    "Corridor_note",
    "",
)

data_asof = event_info.get(
    "Data_asof",
    "",
)

if corridor_note:
    metadata.append(
        str(corridor_note)
    )

if data_asof:
    metadata.append(
        str(data_asof)
    )

if metadata:

    st.caption(
        " | ".join(metadata)
    )


st.caption(
    "Select a district on the map to view detailed flood and WASH information."
)


# ============================================================
# LEGEND
# ============================================================

severity_legend = getattr(
    config,
    "SEVERITY_LEGEND",
    [],
)


if severity_legend:

    with st.expander(
        "ℹ️ Map information and Impact Level",
        expanded=True,
    ):

        legend_columns = st.columns(
            len(severity_legend)
        )

        for column, item in zip(
            legend_columns,
            severity_legend,
        ):

            try:

                label, hex_color = item

            except (ValueError, TypeError):

                continue

            column.markdown(
                f"""
                <div>
                    <span
                        style="
                            display:inline-block;
                            width:14px;
                            height:14px;
                            background:{hex_color};
                            border-radius:3px;
                            margin-right:6px;
                        "
                    ></span>
                    {label}
                </div>
                """,
                unsafe_allow_html=True,
            )


# ============================================================
# LOAD DISTRICT DATA
# ============================================================

try:

    district_df, district_error = (
        data_loader.load_district_situation(
            excel_path
        )
    )

except Exception as e:

    district_df = None
    district_error = str(e)


if district_error:

    st.error(
        "Failed to load District_Situation data."
    )

    st.code(
        str(district_error)
    )

    st.stop()


if district_df is None:

    st.error(
        "District data could not be loaded."
    )

    st.stop()


# ============================================================
# PREPARE GEOJSON
# ============================================================

boundary_gdf = flood_gdf.copy()


# ============================================================
# JOIN KEY
# ============================================================

join_key = getattr(
    config,
    "JOIN_KEY",
    "District_ID",
)


# ============================================================
# CHECK GEOJSON JOIN KEY
# ============================================================

if join_key not in boundary_gdf.columns:

    st.error(
        f"GeoJSON does not contain the required join field: {join_key}"
    )

    st.markdown(
        "Available GeoJSON fields:"
    )

    st.code(
        "\n".join(
            str(column)
            for column in boundary_gdf.columns
        )
    )

    st.stop()


# ============================================================
# CHECK DISTRICT DATA JOIN KEY
# ============================================================

if join_key not in district_df.columns:

    st.error(
        f"District data does not contain the required join field: {join_key}"
    )

    st.markdown(
        "Available District data fields:"
    )

    st.code(
        "\n".join(
            str(column)
            for column in district_df.columns
        )
    )

    st.stop()


# ============================================================
# NORMALIZE JOIN KEY
# ============================================================

boundary_gdf[join_key] = (
    boundary_gdf[join_key]
    .astype(str)
    .str.strip()
)


district_df[join_key] = (
    district_df[join_key]
    .astype(str)
    .str.strip()
)


# ============================================================
# JOIN DISTRICT DATA
# ============================================================

try:

    merged = map_utils.join_district_data(
        boundary_gdf,
        district_df,
    )

except Exception as e:

    st.error(
        "Failed to join GeoJSON and District data."
    )

    st.code(
        str(e)
    )

    st.stop()


# ============================================================
# JOIN QUALITY CHECK
# ============================================================

geojson_ids = set(
    boundary_gdf[join_key]
    .dropna()
    .astype(str)
)

district_ids = set(
    district_df[join_key]
    .dropna()
    .astype(str)
)


missing_in_district = (
    geojson_ids - district_ids
)

missing_in_geojson = (
    district_ids - geojson_ids
)


if (
    missing_in_district
    or missing_in_geojson
):

    with st.expander(
        "🔍 Join quality check",
        expanded=False,
    ):

        if missing_in_district:

            st.warning(
                "IDs found in GeoJSON but not in District data."
            )

            st.code(
                "\n".join(
                    sorted(
                        missing_in_district
                    )
                )
            )

        if missing_in_geojson:

            st.warning(
                "IDs found in District data but not in GeoJSON."
            )

            st.code(
                "\n".join(
                    sorted(
                        missing_in_geojson
                    )
                )
            )


# ============================================================
# LOAD FLOOD CORRIDOR
# ============================================================

try:

    corridor_df, corridor_error = (
        data_loader.load_flood_corridor(
            corridor_path
        )
    )

except Exception as e:

    corridor_df = None
    corridor_error = str(e)


# ============================================================
# LOAD EVENT MARKERS
# ============================================================

try:

    markers_df, markers_error = (
        data_loader.load_event_markers(
            markers_path
        )
    )

except Exception as e:

    markers_df = None
    markers_error = str(e)


# ============================================================
# BUILD EXTRA MAP LAYERS
# ============================================================

extra_layers = []


if corridor_df is not None:

    try:

        corridor_layer = (
            map_utils.build_corridor_layer(
                corridor_df
            )
        )

        if corridor_layer is not None:

            extra_layers.append(
                corridor_layer
            )

    except Exception:
        pass


if markers_df is not None:

    try:

        point_layer, text_layer = (
            map_utils.build_marker_layers(
                markers_df
            )
        )

        if point_layer is not None:

            extra_layers.append(
                point_layer
            )

        if text_layer is not None:

            extra_layers.append(
                text_layer
            )

    except Exception:
        pass


# ============================================================
# MAP AND DETAIL COLUMNS
# ============================================================

map_column, detail_column = st.columns(
    [2, 1]
)


# ============================================================
# BUILD MAP
# ============================================================

with map_column:

    try:

        deck = map_utils.build_deck(
            merged,
            extra_layers=extra_layers,
        )

    except Exception as e:

        st.error(
            "Failed to build the map."
        )

        st.code(
            str(e)
        )

        st.stop()


    try:

        map_event = st.pydeck_chart(
            deck,
            on_select="rerun",
            selection_mode="single-object",
            key="district_map",
            use_container_width=True,
            height=620,
        )

    except Exception as e:

        st.error(
            "Failed to display the map."
        )

        st.code(
            str(e)
        )

        st.stop()


# ============================================================
# GET SELECTED DISTRICT
# ============================================================

selected_properties = None


try:

    selection = map_event.selection

    objects = (
        selection
        .get("objects", {})
        .get(
            "district-layer",
            [],
        )
    )

    if objects:

        selected_properties = objects[0]

except Exception:

    selected_properties = None


# ============================================================
# DETAIL PANEL
# ============================================================

with detail_column:

    st.subheader(
        "📋 District Detail"
    )

    if not selected_properties:

        st.info(
            "Click a district on the map to view detailed information."
        )

    else:

        district_name = format_text(
            selected_properties.get(
                "District",
                "Unknown District",
            )
        )

        severity = format_text(
            selected_properties.get(
                "Severity",
                "No Data",
            )
        )


        # ----------------------------------------------------
        # DISTRICT NAME
        # ----------------------------------------------------

        st.markdown(
            f"### {district_name}"
        )


        # ----------------------------------------------------
        # SEVERITY
        # ----------------------------------------------------

        severity_colors = dict(
            severity_legend
        ) if severity_legend else {}


        severity_color = severity_colors.get(
            severity,
            "#808080",
        )


        st.markdown(
            f"""
            <div
                style="
                    display:inline-block;
                    background:{severity_color};
                    color:white;
                    padding:4px 12px;
                    border-radius:12px;
                    font-weight:600;
                    margin-bottom:10px;
                "
            >
                {severity}
            </div>
            """,
            unsafe_allow_html=True,
        )


        # ----------------------------------------------------
        # KEY INDICATORS
        # ----------------------------------------------------

        col1, col2 = st.columns(2)


        col1.metric(
            "Affected population",
            format_number(
                selected_properties.get(
                    "Affected_population"
                )
            ),
        )


        col2.metric(
            "Water access population",
            format_number(
                selected_properties.get(
                    "Water_access_population"
                )
            ),
        )


        col3, col4 = st.columns(2)


        col3.metric(
            "Casualties",
            format_number(
                selected_properties.get(
                    "Casualties"
                )
            ),
        )


        col4.metric(
            "Estimated funding",
            format_currency(
                selected_properties.get(
                    "Estimated_funding_USD"
                )
            ),
        )


        # ----------------------------------------------------
        # WATER SOURCE
        # ----------------------------------------------------

        st.markdown(
            "---"
        )

        st.markdown(
            "**Current water source**"
        )

        st.write(
            format_text(
                selected_properties.get(
                    "Water_source_current"
                )
            )
        )


        # ----------------------------------------------------
        # KEY PROBLEMS
        # ----------------------------------------------------

        st.markdown(
            "**Key problems / needs**"
        )

        st.write(
            format_text(
                selected_properties.get(
                    "Key_problems"
                )
            )
        )


        # ----------------------------------------------------
        # SEVERITY REASON
        # ----------------------------------------------------

        st.markdown(
            "**Severity reason**"
        )

        severity_reason = (
            selected_properties.get(
                "Severity_reason"
            )
        )

        if severity_reason is None:

            severity_reason = (
                selected_properties.get(
                    "Priority_reason"
                )
            )

        st.write(
            format_text(
                severity_reason
            )
        )


        # ----------------------------------------------------
        # INTERVENTIONS
        # ----------------------------------------------------

        st.markdown(
            "---"
        )

        st.markdown(
            "**Immediate intervention**"
        )

        st.write(
            format_text(
                selected_properties.get(
                    "Immediate_intervention"
                )
            )
        )


        st.markdown(
            "**Medium-term intervention**"
        )

        st.write(
            format_text(
                selected_properties.get(
                    "Medium_term_intervention"
                )
            )
        )


        st.markdown(
            "**Long-term intervention**"
        )

        st.write(
            format_text(
                selected_properties.get(
                    "Long_term_intervention"
                )
            )
        )


        # ----------------------------------------------------
        # DATA INFORMATION
        # ----------------------------------------------------

        st.markdown(
            "---"
        )

        data_source = format_text(
            selected_properties.get(
                "Data_source"
            )
        )

        last_updated = format_text(
            selected_properties.get(
                "Last_updated"
            )
        )

        data_status = format_text(
            selected_properties.get(
                "Data_status"
            )
        )


        st.caption(
            f"Data source: {data_source}"
        )

        st.caption(
            f"Last updated: {last_updated}"
        )

        st.caption(
            f"Data status: {data_status}"
        )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    "---"
)

st.caption(
    "UNICEF Nepal | WASH Emergency Response"
)

st.caption(
    f"Population source: {POPULATION_PATH.name} | "
    f"Flood layer: {FLOOD_GEOJSON_PATH.name} | "
    "GeoJSON engine: pyogrio"
)