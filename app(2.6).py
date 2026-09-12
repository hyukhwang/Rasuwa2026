# -*- coding: utf-8 -*-
"""
UNICEF Nepal
Flood WASH Emergency Decision Support Dashboard

Streamlit + Folium version

Design goals
1. The map is the first thing the user sees when the app loads.
2. No API key is required for the base map (CartoDB Positron tiles).
3. All user-facing text is in English, no emoji anywhere.
4. Every flood-affected District (or Municipality) can carry its own
   situation data, sourced from the flood.geojson attributes and/or
   the WASH Excel template / CSV files, joined on District name
   (or District_ID when available).
5. The detail panel is organized around five sections requested by
   the response team:
       - Affected population & current situation
       - Problem / Needs
       - Required solution
       - Estimated financial requirement
       - Priority areas
   Schools and health facilities can be shown as additional,
   independently toggleable map layers.
6. Data files (Excel / CSV) are optional. Missing files never crash
   the app; the dashboard degrades gracefully and tells the user
   what is missing.

How to update the data later
------------------------------
Fill in data/WASH_data_template.xlsx (or the matching CSV files) and
click "Refresh data" in the sidebar. No code changes are needed.

    Sheet: District_Situation (one row per District)
        District, Affected_population, Water_access_population,
        Casualties, Estimated_funding_USD, Water_source_current,
        Key_problems, Severity_reason, Priority_reason,
        Host_community_needs, Immediate_intervention,
        Medium_term_intervention, Long_term_intervention,
        Data_source, Last_updated, Data_status, Severity

    Sheet: Water_Access / Interventions / Funding
        Any columns you need, as long as a District column is
        included so rows can be matched to the selected District.

    Sheet: Facilities (to show schools / health facilities on the
    map as toggleable layers)
        Facility_Name, Facility_Type (School / Health),
        Latitude, Longitude, District, Status

    File: data/population.csv
        District (or District_ID), plus any population breakdown
        columns you want to display.

The District name used everywhere must match the "District" value
stored in flood.geojson exactly (case and spacing are normalized
automatically, but the underlying name must be the same place).
"""

import os
import json
from pathlib import Path

import pandas as pd
import geopandas as gpd
import folium
import branca.colormap as bcm
import streamlit as st
from streamlit_folium import st_folium


# ============================================================
# 0. Streamlit configuration
# ============================================================

st.set_page_config(
    page_title="UNICEF Nepal - Flood WASH Dashboard",
    page_icon=None,
    layout="wide",
)


# ============================================================
# 1. Basic configuration
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

POPULATION_PATH = DATA_DIR / "population.csv"
FLOOD_GEOJSON_PATH = DATA_DIR / "flood.geojson"
MUNICIPALITY_GEOJSON_PATH = DATA_DIR / "municipality.geojson"

MUNICIPALITY_SITUATION_PATH = (
    DATA_DIR / "municipality_situation.csv"
)

DEFAULT_EXCEL_PATH = (
    DATA_DIR / "WASH_data_template.xlsx"
)

DEFAULT_DISTRICT_CSV_PATH = (
    DATA_DIR / "district_situation.csv"
)

DEFAULT_EVENT_INFO_PATH = (
    DATA_DIR / "event_info.csv"
)

DEFAULT_EVENT_MARKERS_PATH = (
    DATA_DIR / "event_markers.csv"
)

DEFAULT_FLOOD_CORRIDOR_PATH = (
    DATA_DIR / "flood_corridor_path.csv"
)


# ============================================================
# 2. Field names
# ============================================================

DISTRICT_ID_FIELD = "District_ID"
DISTRICT_NAME_FIELD = "District"

MUNICIPALITY_ID_FIELD = "Municipality_ID"
MUNICIPALITY_NAME_FIELD = "Municipality"


# ============================================================
# 3. Default map
# ============================================================

DEFAULT_LAT = 28.05
DEFAULT_LON = 84.85
DEFAULT_ZOOM = 8

# CartoDB Positron requires no API key or token.
BASE_MAP_TILES = "CartoDB positron"


# ============================================================
# 4. Severity
# ============================================================

SEVERITY_LEGEND = [
    ("Severe", "#7A0F0F"),
    ("Very High", "#D62728"),
    ("High", "#F4912D"),
    ("Moderate", "#FFDD57"),
    ("No Data", "#C8C5BD"),
]

SEVERITY_HEX = dict(SEVERITY_LEGEND)

DEFAULT_SEVERITY_HEX = (
    SEVERITY_HEX["No Data"]
)


# ============================================================
# 4-1. Selectable indicators for the map / table
# ============================================================

# Each entry: (label shown to the user, column name, kind)
# kind is "categorical" (uses the Severity-style legend) or
# "numeric" (uses a continuous color ramp).
INDICATOR_OPTIONS = [
    ("Severity", "Severity", "categorical"),
    (
        "Affected population",
        "Affected_population",
        "numeric",
    ),
    (
        "Water access population",
        "Water_access_population",
        "numeric",
    ),
    ("Casualties", "Casualties", "numeric"),
    (
        "Estimated funding (USD)",
        "Estimated_funding_USD",
        "numeric",
    ),
]

NUMERIC_COLOR_RAMP = [
    "#FFF3B0",  # low
    "#FFDD57",
    "#F4912D",
    "#D62728",
    "#7A0F0F",  # high
]


# ============================================================
# 5. Map marker colors
# ============================================================

MARKER_HEX = {
    "epicentre": "#7A0F0F",
    "annotation": "#282828",
}

DEFAULT_MARKER_HEX = "#505050"

CORRIDOR_HEX = "#1E5AC8"

FACILITY_HEX = {
    "school": "#1E5AC8",
    "health": "#2E9E5B",
}

DEFAULT_FACILITY_HEX = "#505050"


# ============================================================
# 6. Excel sheet names
# ============================================================

SHEET_DISTRICT_SITUATION = (
    "District_Situation"
)

SHEET_WATER_ACCESS = (
    "Water_Access"
)

SHEET_INTERVENTIONS = (
    "Interventions"
)

SHEET_FUNDING = (
    "Funding"
)

SHEET_FACILITIES = (
    "Facilities"
)


# ============================================================
# 7. Formatting helpers
# ============================================================

def format_number(value):
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
    if value is None:
        return "No data"

    try:
        if pd.isna(value):
            return "No data"
    except Exception:
        pass

    value = str(value).strip()

    if not value:
        return "No data"

    return value


# ============================================================
# 8. Column normalization
# ============================================================

def normalize_column_names(df):

    if df is None or df.empty:
        return df

    df = df.copy()

    rename_map = {}

    for col in df.columns:

        clean = str(col).strip()
        lower = clean.lower()

        if lower in [
            "district_id",
            "districtid",
            "district id",
        ]:
            rename_map[col] = DISTRICT_ID_FIELD

        elif lower in [
            "district",
            "district_name",
            "district name",
        ]:
            rename_map[col] = DISTRICT_NAME_FIELD

        elif lower in [
            "municipality_id",
            "municipalityid",
            "municipality id",
        ]:
            rename_map[col] = MUNICIPALITY_ID_FIELD

        elif lower in [
            "municipality",
            "municipality_name",
            "municipality name",
        ]:
            rename_map[col] = MUNICIPALITY_NAME_FIELD

        else:
            rename_map[col] = clean

    return df.rename(columns=rename_map)


# ============================================================
# 9. Identifier normalization
# ============================================================

def normalize_identifier_columns(df):

    if df is None or df.empty:
        return df

    df = normalize_column_names(df)

    identifier_columns = [
        DISTRICT_ID_FIELD,
        DISTRICT_NAME_FIELD,
        MUNICIPALITY_ID_FIELD,
        MUNICIPALITY_NAME_FIELD,
    ]

    for col in identifier_columns:

        if col not in df.columns:
            continue

        df[col] = (
            df[col]
            .astype("string")
            .str.strip()
        )

        df.loc[
            df[col]
            .str.lower()
            .isin(
                [
                    "nan",
                    "none",
                    "null",
                    "",
                    "<na>",
                ]
            ),
            col,
        ] = pd.NA

    return df


# ============================================================
# 10. File signature
# ============================================================

def _file_signature(path):

    try:
        return (
            str(path),
            os.path.getmtime(path),
        )

    except OSError:
        return (
            str(path),
            None,
        )


# ============================================================
# 11. CSV readers
# ============================================================

@st.cache_data(show_spinner=False)
def _read_csv(path, signature):

    return pd.read_csv(path)


def _safe_read_csv(path):

    if not path:
        return (
            pd.DataFrame(),
            "No file path was provided.",
        )

    if not os.path.exists(path):

        return (
            pd.DataFrame(),
            f"File not found: {path}",
        )

    try:

        df = _read_csv(
            str(path),
            _file_signature(path),
        )

        return (
            normalize_identifier_columns(df),
            None,
        )

    except Exception as e:

        return (
            pd.DataFrame(),
            f"Error reading CSV: {e}",
        )


@st.cache_data(show_spinner=False)
def _read_csv_raw(path, signature):

    return pd.read_csv(path)


def _safe_read_csv_raw(path):

    if not path:
        return (
            pd.DataFrame(),
            "No file path was provided.",
        )

    if not os.path.exists(path):

        return (
            pd.DataFrame(),
            f"File not found: {path}",
        )

    try:

        df = _read_csv_raw(
            str(path),
            _file_signature(path),
        )

        return (
            normalize_identifier_columns(df),
            None,
        )

    except Exception as e:

        return (
            pd.DataFrame(),
            f"Error reading CSV: {e}",
        )


# ============================================================
# 12. Excel readers
# ============================================================

@st.cache_data(show_spinner=False)
def _read_excel_sheet(
    path,
    sheet_name,
    signature,
):

    return pd.read_excel(
        path,
        sheet_name=sheet_name,
        engine="openpyxl",
    )


def _safe_read_excel_sheet(
    path,
    sheet_name,
):

    if not path:
        return (
            pd.DataFrame(),
            "No Excel file path was provided.",
        )

    if not os.path.exists(path):

        return (
            pd.DataFrame(),
            f"File not found: {path}",
        )

    try:

        df = _read_excel_sheet(
            str(path),
            sheet_name,
            _file_signature(path),
        )

        return (
            normalize_identifier_columns(df),
            None,
        )

    except ValueError as e:

        return (
            pd.DataFrame(),
            f"Sheet '{sheet_name}' not found: {e}",
        )

    except Exception as e:

        return (
            pd.DataFrame(),
            f"Error reading sheet '{sheet_name}': {e}",
        )


# ============================================================
# 13. Population
# ============================================================

@st.cache_data(show_spinner=False)
def load_population_csv(
    filepath,
    signature=None,
):

    df = pd.read_csv(filepath)

    return normalize_identifier_columns(df)


# ============================================================
# 14. GeoJSON reader
# ============================================================

@st.cache_data(show_spinner=False)
def _read_geo_file_cached(
    path,
    signature,
):

    try:

        return gpd.read_file(
            path,
            engine="pyogrio",
        )

    except Exception:

        return gpd.read_file(path)


def ensure_wgs84(
    gdf,
    label,
):

    notes = []

    if gdf.crs is None:

        notes.append(
            f"{label}: no CRS found, assumed EPSG:4326."
        )

        gdf = gdf.set_crs(
            epsg=4326,
            allow_override=True,
        )

    else:

        try:
            epsg = gdf.crs.to_epsg()
        except Exception:
            epsg = None

        if epsg != 4326:

            notes.append(
                f"{label}: converted CRS from "
                f"{gdf.crs} to EPSG:4326."
            )

            gdf = gdf.to_crs(
                epsg=4326
            )

    return gdf, notes


def clean_geometries(
    gdf,
    label,
):

    notes = []

    if gdf is None or gdf.empty:
        return gdf, notes

    before = len(gdf)

    gdf = gdf[
        gdf.geometry.notna()
    ].copy()

    if gdf.empty:

        notes.append(
            f"{label}: no valid geometries found."
        )

        return gdf, notes

    try:

        invalid_mask = (
            ~gdf.geometry.is_valid
        )

        if invalid_mask.any():

            try:

                gdf.loc[
                    invalid_mask,
                    "geometry",
                ] = (
                    gdf.loc[
                        invalid_mask,
                        "geometry",
                    ]
                    .buffer(0)
                )

            except Exception:
                pass

    except Exception:
        pass

    after = len(gdf)

    if after < before:

        notes.append(
            f"{label}: removed {before - after} "
            "empty geometries."
        )

    return gdf, notes


def read_geo_file(
    path,
    label,
):

    if not path:

        return (
            None,
            [
                f"No path was provided for {label}."
            ],
        )

    if not os.path.exists(path):

        return (
            None,
            [
                f"File not found for {label}: {path}"
            ],
        )

    try:

        gdf = _read_geo_file_cached(
            str(path),
            _file_signature(path),
        )

    except Exception as e:

        return (
            None,
            [
                f"Error reading {label}: {e}"
            ],
        )

    if gdf is None or gdf.empty:

        return (
            None,
            [
                f"{label} contains no valid data."
            ],
        )

    if "geometry" not in gdf.columns:

        return (
            None,
            [
                f"{label} has no geometry column."
            ],
        )

    gdf = normalize_identifier_columns(
        gdf
    )

    gdf, crs_notes = ensure_wgs84(
        gdf,
        label,
    )

    gdf, geometry_notes = clean_geometries(
        gdf,
        label,
    )

    return (
        gdf,
        crs_notes + geometry_notes,
    )


# ============================================================
# 15. Data loaders
# ============================================================

def load_district_situation(
    excel_path,
    csv_path=None,
):

    if csv_path and os.path.exists(csv_path):

        return _safe_read_csv(
            csv_path
        )

    if excel_path and os.path.exists(excel_path):

        return _safe_read_excel_sheet(
            excel_path,
            SHEET_DISTRICT_SITUATION,
        )

    return (
        pd.DataFrame(),
        "District Situation file not found.",
    )


def load_event_info(path=None):

    path = (
        path
        or DEFAULT_EVENT_INFO_PATH
    )

    df, err = _safe_read_csv_raw(path)

    if err:
        return {}, err

    if df.empty:
        return {}, None

    if (
        "Key" not in df.columns
        or "Value" not in df.columns
    ):
        return {}, None

    return (
        dict(
            zip(
                df["Key"].astype(str),
                df["Value"],
            )
        ),
        None,
    )


def load_event_markers(path=None):

    path = (
        path
        or DEFAULT_EVENT_MARKERS_PATH
    )

    return _safe_read_csv_raw(path)


def load_flood_corridor(path=None):

    path = (
        path
        or DEFAULT_FLOOD_CORRIDOR_PATH
    )

    df, err = _safe_read_csv_raw(path)

    if (
        err is None
        and not df.empty
        and "Sequence" in df.columns
    ):

        df["Sequence"] = pd.to_numeric(
            df["Sequence"],
            errors="coerce",
        )

        df = df.sort_values(
            "Sequence"
        )

    return df, err


def load_water_access(excel_path):

    return _safe_read_excel_sheet(
        excel_path,
        SHEET_WATER_ACCESS,
    )


def load_interventions(excel_path):

    return _safe_read_excel_sheet(
        excel_path,
        SHEET_INTERVENTIONS,
    )


def load_funding(excel_path):

    return _safe_read_excel_sheet(
        excel_path,
        SHEET_FUNDING,
    )


def load_facilities(excel_path):

    return _safe_read_excel_sheet(
        excel_path,
        SHEET_FACILITIES,
    )


def load_municipality_situation(path):

    return _safe_read_csv_raw(path)


# ============================================================
# 16. Join key
# ============================================================

def determine_join_key(
    flood_gdf,
    district_df,
):

    if (
        flood_gdf is None
        or flood_gdf.empty
    ):
        return None

    if (
        district_df is None
        or district_df.empty
    ):
        return None

    if (
        DISTRICT_ID_FIELD in flood_gdf.columns
        and DISTRICT_ID_FIELD in district_df.columns
    ):
        return DISTRICT_ID_FIELD

    if (
        DISTRICT_NAME_FIELD in flood_gdf.columns
        and DISTRICT_NAME_FIELD in district_df.columns
    ):
        return DISTRICT_NAME_FIELD

    return None


# ============================================================
# 17. District data join
# ============================================================

def prepare_flood_gdf(
    flood_gdf,
):

    gdf = flood_gdf.copy()

    # --------------------------------------------------------
    # Ensure District
    # --------------------------------------------------------

    if DISTRICT_NAME_FIELD in gdf.columns:

        gdf[DISTRICT_NAME_FIELD] = (
            gdf[DISTRICT_NAME_FIELD]
            .astype("string")
            .str.strip()
        )

    if DISTRICT_ID_FIELD in gdf.columns:

        gdf[DISTRICT_ID_FIELD] = (
            gdf[DISTRICT_ID_FIELD]
            .astype("string")
            .str.strip()
        )

    # --------------------------------------------------------
    # Severity
    # --------------------------------------------------------

    if "Severity" not in gdf.columns:

        if "Priority" in gdf.columns:

            gdf["Severity"] = (
                gdf["Priority"]
            )

        else:

            gdf["Severity"] = "No Data"

    gdf["Severity"] = (
        gdf["Severity"]
        .fillna("No Data")
        .astype(str)
        .str.strip()
    )

    severity_mapping = {
        "severe": "Severe",
        "very high": "Very High",
        "very_high": "Very High",
        "high": "High",
        "moderate": "Moderate",
        "medium": "Moderate",
        "no data": "No Data",
        "none": "No Data",
        "nan": "No Data",
        "": "No Data",
    }

    gdf["Severity"] = (
        gdf["Severity"]
        .str.lower()
        .map(severity_mapping)
        .fillna("No Data")
    )

    gdf["severity_hex"] = (
        gdf["Severity"]
        .map(SEVERITY_HEX)
        .fillna(DEFAULT_SEVERITY_HEX)
    )

    gdf["is_flood_affected"] = (
        gdf["Severity"] != "No Data"
    )

    # --------------------------------------------------------
    # Ensure common fields
    #
    # These columns line up with the five reporting sections:
    #   Affected population & current situation
    #   Problem / Needs
    #   Required solution
    #   Estimated financial requirement
    #   Priority areas
    # --------------------------------------------------------

    default_columns = [
        "Affected_population",
        "Water_access_population",
        "Estimated_funding_USD",
        "Casualties",
        "Water_source_current",
        "Key_problems",
        "Severity_reason",
        "Priority_reason",
        "Host_community_needs",
        "Immediate_intervention",
        "Medium_term_intervention",
        "Long_term_intervention",
        "Data_source",
        "Last_updated",
        "Data_status",
    ]

    for col in default_columns:

        if col not in gdf.columns:
            gdf[col] = None

    return gdf


def join_district_data(
    boundary_gdf,
    district_df,
    join_key,
):

    boundary_gdf = boundary_gdf.copy()

    if (
        district_df is None
        or district_df.empty
        or join_key is None
    ):

        return prepare_flood_gdf(
            boundary_gdf
        )

    district_df = district_df.copy()

    if join_key not in boundary_gdf.columns:

        return prepare_flood_gdf(
            boundary_gdf
        )

    if join_key not in district_df.columns:

        return prepare_flood_gdf(
            boundary_gdf
        )

    boundary_gdf[join_key] = (
        boundary_gdf[join_key]
        .astype("string")
        .str.strip()
    )

    district_df[join_key] = (
        district_df[join_key]
        .astype("string")
        .str.strip()
    )

    # Avoid duplicate rows in situation data
    district_df = district_df.drop_duplicates(
        subset=[join_key],
        keep="first",
    )

    merged = boundary_gdf.merge(
        district_df,
        on=join_key,
        how="left",
        suffixes=(
            "_boundary",
            "",
        ),
    )

    # --------------------------------------------------------
    # Recover boundary fields
    # --------------------------------------------------------

    for field in [
        DISTRICT_NAME_FIELD,
        DISTRICT_ID_FIELD,
    ]:

        boundary_field = (
            f"{field}_boundary"
        )

        if (
            field not in merged.columns
            and boundary_field in merged.columns
        ):

            merged[field] = merged[
                boundary_field
            ]

    # --------------------------------------------------------
    # If District name is missing after merge
    # --------------------------------------------------------

    if DISTRICT_NAME_FIELD not in merged.columns:

        boundary_name = (
            f"{DISTRICT_NAME_FIELD}_boundary"
        )

        if boundary_name in merged.columns:

            merged[DISTRICT_NAME_FIELD] = (
                merged[boundary_name]
            )

    return prepare_flood_gdf(
        merged
    )


# ============================================================
# 18. Generic selection filter
# ============================================================

def filter_data_by_selection(
    df,
    selected_properties,
    join_key=None,
):

    if (
        df is None
        or df.empty
        or selected_properties is None
    ):

        return pd.DataFrame()

    df = normalize_identifier_columns(
        df.copy()
    )

    # --------------------------------------------------------
    # Primary join key
    # --------------------------------------------------------

    if (
        join_key
        and join_key in df.columns
    ):

        selected_value = (
            selected_properties.get(
                join_key
            )
        )

        if selected_value is not None:

            selected_value = str(
                selected_value
            ).strip()

            result = df[
                df[join_key]
                .astype(str)
                .str.strip()
                == selected_value
            ]

            if not result.empty:
                return result

    # --------------------------------------------------------
    # District ID fallback
    # --------------------------------------------------------

    selected_id = (
        selected_properties.get(
            DISTRICT_ID_FIELD
        )
    )

    if (
        selected_id is not None
        and DISTRICT_ID_FIELD in df.columns
    ):

        result = df[
            df[DISTRICT_ID_FIELD]
            .astype(str)
            .str.strip()
            == str(selected_id).strip()
        ]

        if not result.empty:
            return result

    # --------------------------------------------------------
    # District name fallback
    # --------------------------------------------------------

    selected_name = (
        selected_properties.get(
            DISTRICT_NAME_FIELD
        )
    )

    if (
        selected_name is not None
        and DISTRICT_NAME_FIELD in df.columns
    ):

        result = df[
            df[DISTRICT_NAME_FIELD]
            .astype(str)
            .str.strip()
            == str(selected_name).strip()
        ]

        if not result.empty:
            return result

    return pd.DataFrame()


# ============================================================
# 19. Municipality filter
# ============================================================

def filter_municipality_rows(
    df,
    selected_properties,
):

    if (
        df is None
        or df.empty
        or selected_properties is None
    ):

        return pd.DataFrame()

    df = normalize_identifier_columns(
        df.copy()
    )

    # Municipality ID
    selected_id = (
        selected_properties.get(
            MUNICIPALITY_ID_FIELD
        )
    )

    if (
        selected_id is not None
        and MUNICIPALITY_ID_FIELD in df.columns
    ):

        result = df[
            df[MUNICIPALITY_ID_FIELD]
            .astype(str)
            .str.strip()
            == str(selected_id).strip()
        ]

        if not result.empty:
            return result

    # Municipality name
    selected_name = (
        selected_properties.get(
            MUNICIPALITY_NAME_FIELD
        )
    )

    if (
        selected_name is not None
        and MUNICIPALITY_NAME_FIELD in df.columns
    ):

        result = df[
            df[MUNICIPALITY_NAME_FIELD]
            .astype(str)
            .str.strip()
            == str(selected_name).strip()
        ]

        if not result.empty:
            return result

    return pd.DataFrame()


# ============================================================
# 19-1. Indicator display layer (choropleth coloring)
# ============================================================

def compute_display_layer(
    gdf,
    column,
    kind,
):
    """
    Given the merged district GeoDataFrame, compute a
    'display_hex' column used to color the map according to
    the indicator currently selected by the user, plus legend
    information to render under the map.

    Returns (gdf_with_display_hex, legend_info) where
    legend_info is one of:
        {"kind": "categorical", "items": [(label, color), ...]}
        {"kind": "numeric", "vmin": ..., "vmax": ...,
         "colors": [...], "no_data_color": ...}
        {"kind": "empty"}   # no usable data at all
    """

    gdf = gdf.copy()

    if (
        gdf is None
        or gdf.empty
        or column not in gdf.columns
    ):

        gdf["display_hex"] = DEFAULT_SEVERITY_HEX

        return gdf, {"kind": "empty"}

    if kind == "categorical":

        gdf["display_hex"] = (
            gdf[column]
            .map(SEVERITY_HEX)
            .fillna(DEFAULT_SEVERITY_HEX)
        )

        legend_info = {
            "kind": "categorical",
            "items": SEVERITY_LEGEND,
        }

        return gdf, legend_info

    # --------------------------------------------------------
    # Numeric indicator
    # --------------------------------------------------------

    values = pd.to_numeric(
        gdf[column],
        errors="coerce",
    )

    valid = values.dropna()

    if valid.empty:

        gdf["display_hex"] = DEFAULT_SEVERITY_HEX

        return gdf, {"kind": "empty"}

    vmin = float(valid.min())
    vmax = float(valid.max())

    if vmin == vmax:
        vmax = vmin + 1.0

    colormap = bcm.LinearColormap(
        colors=NUMERIC_COLOR_RAMP,
        vmin=vmin,
        vmax=vmax,
    )

    def _to_hex(v):

        if pd.isna(v):
            return DEFAULT_SEVERITY_HEX

        try:
            return colormap(float(v))
        except Exception:
            return DEFAULT_SEVERITY_HEX

    gdf["display_hex"] = values.map(_to_hex)

    legend_info = {
        "kind": "numeric",
        "vmin": vmin,
        "vmax": vmax,
        "colors": NUMERIC_COLOR_RAMP,
        "no_data_color": DEFAULT_SEVERITY_HEX,
    }

    return gdf, legend_info


# ============================================================
# 20. Folium styles
# ============================================================

def district_style_function(feature):

    props = (
        feature.get(
            "properties",
            {},
        )
        or {}
    )

    color = (
        props.get("display_hex")
        or props.get("severity_hex")
        or DEFAULT_SEVERITY_HEX
    )

    affected = (
        props.get(
            "is_flood_affected",
            False,
        )
        is True
    )

    return {
        "fillColor": color,
        "color": (
            "#1f1f1f"
            if affected
            else "#8a8a8a"
        ),
        "weight": (
            3
            if affected
            else 1.2
        ),
        "fillOpacity": 0.65,
    }


def district_highlight_function(
    feature,
):

    return {
        "weight": 4,
        "color": "#0078FF",
        "fillOpacity": 0.8,
    }


def municipality_style_function(
    feature,
):

    return {
        "fillColor": "#ffffff",
        "color": "#333333",
        "weight": 1,
        "fillOpacity": 0.03,
        "dashArray": "3, 3",
    }


def municipality_highlight_function(
    feature,
):

    return {
        "weight": 3,
        "color": "#FF7A00",
        "fillOpacity": 0.15,
    }


# ============================================================
# 21. Flood corridor
# ============================================================

def add_corridor_to_map(
    m,
    corridor_df,
):

    if (
        corridor_df is None
        or corridor_df.empty
    ):
        return

    required = {
        "Latitude",
        "Longitude",
    }

    if not required.issubset(
        corridor_df.columns
    ):
        return

    df = corridor_df.copy()

    df["Latitude"] = pd.to_numeric(
        df["Latitude"],
        errors="coerce",
    )

    df["Longitude"] = pd.to_numeric(
        df["Longitude"],
        errors="coerce",
    )

    coords = (
        df[
            [
                "Latitude",
                "Longitude",
            ]
        ]
        .dropna()
        .values
        .tolist()
    )

    if not coords:
        return

    folium.PolyLine(
        locations=coords,
        color=CORRIDOR_HEX,
        weight=5,
        opacity=0.85,
        tooltip="Flood corridor",
    ).add_to(m)


# ============================================================
# 22. Event markers
# ============================================================

def add_markers_to_map(
    m,
    markers_df,
):

    if (
        markers_df is None
        or markers_df.empty
    ):
        return

    required = {
        "Latitude",
        "Longitude",
        "Label",
    }

    if not required.issubset(
        markers_df.columns
    ):
        return

    df = markers_df.copy()

    df["Latitude"] = pd.to_numeric(
        df["Latitude"],
        errors="coerce",
    )

    df["Longitude"] = pd.to_numeric(
        df["Longitude"],
        errors="coerce",
    )

    df = df.dropna(
        subset=[
            "Latitude",
            "Longitude",
        ]
    )

    for _, row in df.iterrows():

        marker_type = str(
            row.get(
                "Type",
                "annotation",
            )
        ).lower().strip()

        color = MARKER_HEX.get(
            marker_type,
            DEFAULT_MARKER_HEX,
        )

        label = format_text(
            row.get(
                "Label",
                "",
            )
        )

        note = format_text(
            row.get(
                "Note",
                label,
            )
        )

        folium.CircleMarker(
            location=[
                row["Latitude"],
                row["Longitude"],
            ],
            radius=(
                7
                if marker_type == "epicentre"
                else 5
            ),
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.9,
            tooltip=label,
            popup=note,
        ).add_to(m)


# ============================================================
# 22-1. Facilities layer (schools / health facilities)
#
# These are shown as independently toggleable layers so that
# the response team can turn them on or off without affecting
# the District choropleth or the flood corridor.
# ============================================================

def add_facilities_to_map(
    m,
    facilities_df,
):

    if (
        facilities_df is None
        or facilities_df.empty
    ):
        return

    required = {
        "Latitude",
        "Longitude",
    }

    if not required.issubset(
        facilities_df.columns
    ):
        return

    df = facilities_df.copy()

    df["Latitude"] = pd.to_numeric(
        df["Latitude"],
        errors="coerce",
    )

    df["Longitude"] = pd.to_numeric(
        df["Longitude"],
        errors="coerce",
    )

    df = df.dropna(
        subset=[
            "Latitude",
            "Longitude",
        ]
    )

    if df.empty:
        return

    type_col = (
        "Facility_Type"
        if "Facility_Type" in df.columns
        else None
    )

    feature_groups = {}

    for _, row in df.iterrows():

        raw_type = (
            str(row.get(type_col, "Facility")).strip()
            if type_col
            else "Facility"
        )

        type_label = raw_type if raw_type else "Facility"

        color = FACILITY_HEX.get(
            type_label.lower(),
            DEFAULT_FACILITY_HEX,
        )

        if type_label not in feature_groups:

            feature_groups[type_label] = folium.FeatureGroup(
                name=f"Facilities: {type_label}",
                show=True,
            )

        name = format_text(
            row.get("Facility_Name", type_label)
        )

        popup_lines = [name]

        if DISTRICT_NAME_FIELD in df.columns:

            popup_lines.append(
                f"District: "
                f"{format_text(row.get(DISTRICT_NAME_FIELD))}"
            )

        if "Status" in df.columns:

            popup_lines.append(
                f"Status: {format_text(row.get('Status'))}"
            )

        folium.CircleMarker(
            location=[
                row["Latitude"],
                row["Longitude"],
            ],
            radius=6,
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.9,
            tooltip=f"{type_label}: {name}",
            popup="<br>".join(popup_lines),
        ).add_to(feature_groups[type_label])

    for group in feature_groups.values():
        group.add_to(m)


# ============================================================
# 23. Build Folium map
# ============================================================

def build_folium_map(
    district_gdf,
    municipality_gdf=None,
    corridor_df=None,
    markers_df=None,
    facilities_df=None,
    extra_tooltip=None,
):

    m = folium.Map(
        location=[
            DEFAULT_LAT,
            DEFAULT_LON,
        ],
        zoom_start=DEFAULT_ZOOM,
        tiles=BASE_MAP_TILES,
        control_scale=True,
    )

    # --------------------------------------------------------
    # District layer
    # --------------------------------------------------------

    if (
        district_gdf is not None
        and not district_gdf.empty
    ):

        district_geojson = json.loads(
            district_gdf.to_json()
        )

        tooltip_fields = []
        tooltip_aliases = []

        if DISTRICT_NAME_FIELD in district_gdf.columns:

            tooltip_fields.append(
                DISTRICT_NAME_FIELD
            )

            tooltip_aliases.append(
                "District:"
            )

        if DISTRICT_ID_FIELD in district_gdf.columns:

            tooltip_fields.append(
                DISTRICT_ID_FIELD
            )

            tooltip_aliases.append(
                "District ID:"
            )

        if "Severity" in district_gdf.columns:

            tooltip_fields.append(
                "Severity"
            )

            tooltip_aliases.append(
                "Severity:"
            )

        if (
            extra_tooltip
            and extra_tooltip[0] in district_gdf.columns
            and extra_tooltip[0] not in tooltip_fields
        ):

            tooltip_fields.append(
                extra_tooltip[0]
            )

            tooltip_aliases.append(
                f"{extra_tooltip[1]}:"
            )

        if tooltip_fields:

            district_layer = folium.GeoJson(
                district_geojson,
                name="Districts",
                style_function=(
                    district_style_function
                ),
                highlight_function=(
                    district_highlight_function
                ),
                tooltip=folium.GeoJsonTooltip(
                    fields=tooltip_fields,
                    aliases=tooltip_aliases,
                    sticky=True,
                ),
                zoom_on_click=False,
            )

        else:

            district_layer = folium.GeoJson(
                district_geojson,
                name="Districts",
                style_function=(
                    district_style_function
                ),
                highlight_function=(
                    district_highlight_function
                ),
                zoom_on_click=False,
            )

        district_layer.add_to(m)

    # --------------------------------------------------------
    # Municipality layer
    # --------------------------------------------------------

    if (
        municipality_gdf is not None
        and not municipality_gdf.empty
    ):

        muni_geojson = json.loads(
            municipality_gdf.to_json()
        )

        muni_fields = []
        muni_aliases = []

        if (
            MUNICIPALITY_NAME_FIELD
            in municipality_gdf.columns
        ):

            muni_fields.append(
                MUNICIPALITY_NAME_FIELD
            )

            muni_aliases.append(
                "Municipality:"
            )

        if (
            DISTRICT_NAME_FIELD
            in municipality_gdf.columns
        ):

            muni_fields.append(
                DISTRICT_NAME_FIELD
            )

            muni_aliases.append(
                "District:"
            )

        if (
            DISTRICT_ID_FIELD
            in municipality_gdf.columns
        ):

            muni_fields.append(
                DISTRICT_ID_FIELD
            )

            muni_aliases.append(
                "District ID:"
            )

        if (
            MUNICIPALITY_ID_FIELD
            in municipality_gdf.columns
        ):

            muni_fields.append(
                MUNICIPALITY_ID_FIELD
            )

            muni_aliases.append(
                "Municipality ID:"
            )

        kwargs = {
            "name": "Municipalities",
            "style_function": (
                municipality_style_function
            ),
            "highlight_function": (
                municipality_highlight_function
            ),
            "zoom_on_click": False,
        }

        if muni_fields:

            kwargs["tooltip"] = (
                folium.GeoJsonTooltip(
                    fields=muni_fields,
                    aliases=muni_aliases,
                    sticky=True,
                )
            )

        folium.GeoJson(
            muni_geojson,
            **kwargs,
        ).add_to(m)

    # --------------------------------------------------------
    # Corridor
    # --------------------------------------------------------

    add_corridor_to_map(
        m,
        corridor_df,
    )

    # --------------------------------------------------------
    # Event markers
    # --------------------------------------------------------

    add_markers_to_map(
        m,
        markers_df,
    )

    # --------------------------------------------------------
    # Facilities (schools / health facilities)
    # --------------------------------------------------------

    add_facilities_to_map(
        m,
        facilities_df,
    )

    # --------------------------------------------------------
    # Layer control
    # --------------------------------------------------------

    folium.LayerControl(
        collapsed=False
    ).add_to(m)

    # --------------------------------------------------------
    # Bounds
    # --------------------------------------------------------

    try:

        if (
            district_gdf is not None
            and not district_gdf.empty
        ):

            minx, miny, maxx, maxy = (
                district_gdf.total_bounds
            )

            if all(
                pd.notna(
                    [
                        minx,
                        miny,
                        maxx,
                        maxy,
                    ]
                )
            ):

                m.fit_bounds(
                    [
                        [
                            miny,
                            minx,
                        ],
                        [
                            maxy,
                            maxx,
                        ],
                    ]
                )

    except Exception:
        pass

    return m


# ============================================================
# 24. Render table helper
# ============================================================

def render_sheet_table(
    df,
    err,
    selected_properties,
    join_key,
    empty_message,
):

    if (
        err
        and (
            df is None
            or df.empty
        )
    ):

        st.info(
            f"{empty_message}\n\n"
            f"Details: {err}"
        )

        return

    if (
        df is None
        or df.empty
    ):

        st.info(
            empty_message
        )

        return

    sub = filter_data_by_selection(
        df,
        selected_properties,
        join_key,
    )

    if sub.empty:

        st.info(
            empty_message
        )

        return

    st.dataframe(
        sub.reset_index(
            drop=True
        ),
        use_container_width=True,
        hide_index=True,
    )

    numeric_cols = [
        c
        for c in sub.columns
        if pd.api.types.is_numeric_dtype(
            sub[c]
        )
    ]

    if numeric_cols:

        summary_cols = st.columns(
            min(
                len(numeric_cols),
                4,
            )
        )

        for i, cname in enumerate(
            numeric_cols
        ):

            summary_cols[
                i % len(summary_cols)
            ].metric(
                f"Total: {cname}",
                format_number(
                    sub[cname].sum()
                ),
            )


# ============================================================
# 25. Collect sources
# ============================================================

def collect_sources(
    selected_properties,
    join_key,
    water_access_df,
    interventions_df,
    funding_df,
    facilities_df,
):

    frames = {
        "Water Access": water_access_df,
        "Interventions": interventions_df,
        "Funding": funding_df,
        "Facilities": facilities_df,
    }

    rows = []

    for label, df in frames.items():

        if (
            df is None
            or df.empty
        ):
            continue

        sub = filter_data_by_selection(
            df,
            selected_properties,
            join_key,
        )

        if sub.empty:
            continue

        source_cols = [
            c
            for c in sub.columns
            if "source" in c.lower()
        ]

        for col in source_cols:

            values = (
                sub[col]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
            )

            for value in values:

                if (
                    value
                    and value.lower()
                    != "nan"
                ):

                    rows.append(
                        {
                            "Sheet": label,
                            "Column": col,
                            "Source": value,
                        }
                    )

    return rows


# ============================================================
# 26. Main application
# ============================================================

def main():

    # --------------------------------------------------------
    # Sidebar: data source configuration
    # --------------------------------------------------------

    st.sidebar.title("Data Sources")

    st.sidebar.markdown("Repository data")

    st.sidebar.code(
        "data/flood.geojson"
    )

    if POPULATION_PATH.exists():

        st.sidebar.code(
            "data/population.csv"
        )

    excel_path = st.sidebar.text_input(
        "Excel file path",
        value=str(
            DEFAULT_EXCEL_PATH
        ),
    )

    geojson_path = st.sidebar.text_input(
        "Flood GeoJSON path",
        value=str(
            FLOOD_GEOJSON_PATH
        ),
    )

    municipality_path = st.sidebar.text_input(
        "Municipality GeoJSON path",
        value=str(
            MUNICIPALITY_GEOJSON_PATH
        ),
    )

    with st.sidebar.expander(
        "Advanced data sources",
        expanded=False,
    ):

        municipality_situation_path = (
            st.text_input(
                "Municipality situation CSV",
                value=str(
                    MUNICIPALITY_SITUATION_PATH
                ),
            )
        )

        event_info_path = st.text_input(
            "Event information",
            value=str(
                DEFAULT_EVENT_INFO_PATH
            ),
        )

        markers_path = st.text_input(
            "Event markers",
            value=str(
                DEFAULT_EVENT_MARKERS_PATH
            ),
        )

        corridor_path = st.text_input(
            "Flood corridor",
            value=str(
                DEFAULT_FLOOD_CORRIDOR_PATH
            ),
        )

        district_csv_path = st.text_input(
            "District situation CSV",
            value=str(
                DEFAULT_DISTRICT_CSV_PATH
            ),
        )

    with st.sidebar.expander(
        "How to update the data",
        expanded=False,
    ):

        st.markdown(
            "Fill in the Excel template (or the matching CSV "
            "files) below, then click Refresh data. No code "
            "changes are required."
        )

        st.markdown(
            "- District_Situation sheet: one row per District. "
            "Columns include Affected_population, "
            "Water_access_population, Casualties, "
            "Estimated_funding_USD, Water_source_current, "
            "Key_problems, Severity_reason, Priority_reason, "
            "Host_community_needs, Immediate_intervention, "
            "Medium_term_intervention, Long_term_intervention, "
            "Data_source, Last_updated, Data_status, Severity."
        )

        st.markdown(
            "- Water_Access / Interventions / Funding sheets: "
            "any columns you need, as long as a District column "
            "is included."
        )

        st.markdown(
            "- Facilities sheet: Facility_Name, Facility_Type "
            "(School or Health), Latitude, Longitude, District, "
            "Status. These appear as toggleable layers on the "
            "map."
        )

        st.markdown(
            "The District name must match the District value "
            "stored in flood.geojson exactly."
        )

    if st.sidebar.button(
        "Refresh data"
    ):

        st.cache_data.clear()
        st.rerun()

    # --------------------------------------------------------
    # Population
    # --------------------------------------------------------

    population_df = pd.DataFrame()

    if POPULATION_PATH.exists():

        try:

            population_df = load_population_csv(
                str(POPULATION_PATH),
                _file_signature(
                    POPULATION_PATH
                ),
            )

        except Exception as e:

            st.sidebar.warning(
                "Could not read population.csv."
            )

            st.sidebar.caption(
                str(e)
            )

    st.sidebar.caption(
        f"Population records: "
        f"{len(population_df):,}"
    )

    # --------------------------------------------------------
    # Required flood GeoJSON
    # --------------------------------------------------------

    if not Path(
        geojson_path
    ).exists():

        st.error(
            "Required file `data/flood.geojson` was not found."
        )

        st.code(
            "repository/\n"
            "├── app.py\n"
            "└── data/\n"
            "    └── flood.geojson"
        )

        st.info(
            "Upload flood.geojson into the data folder of the "
            "GitHub repository."
        )

        st.stop()

    # --------------------------------------------------------
    # Load flood GeoJSON
    # --------------------------------------------------------

    flood_gdf, flood_notes = read_geo_file(
        geojson_path,
        "flood.geojson",
    )

    if flood_gdf is None:

        st.error(
            "Failed to load flood.geojson."
        )

        for note in flood_notes:
            st.write(
                f"- {note}"
            )

        st.stop()

    # --------------------------------------------------------
    # Municipality (loaded before District validation so that
    # it can be used as a fallback source for District info)
    # --------------------------------------------------------

    municipality_gdf = None
    muni_notes = []

    if (
        municipality_path
        and Path(
            municipality_path
        ).exists()
    ):

        municipality_gdf, muni_notes = (
            read_geo_file(
                municipality_path,
                "municipality.geojson",
            )
        )

        if municipality_gdf is not None:

            if (
                MUNICIPALITY_NAME_FIELD
                not in municipality_gdf.columns
            ):

                st.sidebar.warning(
                    "municipality.geojson has no "
                    "'Municipality' column."
                )

                municipality_gdf = None

        if municipality_gdf is not None:

            for col in [
                DISTRICT_NAME_FIELD,
                DISTRICT_ID_FIELD,
                MUNICIPALITY_NAME_FIELD,
                MUNICIPALITY_ID_FIELD,
            ]:

                if col in municipality_gdf.columns:

                    municipality_gdf[col] = (
                        municipality_gdf[col]
                        .astype("string")
                        .str.strip()
                    )

            st.sidebar.caption(
                f"Municipality features: "
                f"{len(municipality_gdf):,}"
            )

        if muni_notes:

            with st.sidebar.expander(
                "Municipality processing log",
                expanded=False,
            ):

                for note in muni_notes:
                    st.caption(note)

    else:

        st.sidebar.info(
            "municipality.geojson not found. "
            "Showing the District layer only."
        )

    # --------------------------------------------------------
    # Ensure District field exists in flood_gdf
    #
    # If flood.geojson has no District column at all, the app
    # does not stop. Instead:
    #   1) If municipality.geojson has District information, a
    #      spatial join is used to estimate District for each
    #      flood feature.
    #   2) Otherwise, temporary names ("Zone 1", "Zone 2", ...)
    #      are generated so the map and click interaction still
    #      work. Replacing flood.geojson with real administrative
    #      District boundaries is strongly recommended.
    # --------------------------------------------------------

    district_field_source = None

    if DISTRICT_NAME_FIELD not in flood_gdf.columns:

        joined_from_municipality = False

        if (
            municipality_gdf is not None
            and not municipality_gdf.empty
            and DISTRICT_NAME_FIELD in municipality_gdf.columns
        ):

            try:

                flood_gdf = flood_gdf.reset_index(drop=True)
                flood_gdf["_flood_row_id"] = flood_gdf.index

                points = flood_gdf.copy()
                points["geometry"] = (
                    points.geometry.representative_point()
                )

                muni_cols = [
                    c
                    for c in [
                        DISTRICT_NAME_FIELD,
                        DISTRICT_ID_FIELD,
                    ]
                    if c in municipality_gdf.columns
                ]

                joined = gpd.sjoin(
                    points[["_flood_row_id", "geometry"]],
                    municipality_gdf[muni_cols + ["geometry"]],
                    how="left",
                    predicate="intersects",
                )

                joined = joined.drop_duplicates(
                    subset=["_flood_row_id"],
                    keep="first",
                )

                district_lookup = joined.set_index(
                    "_flood_row_id"
                )[muni_cols]

                flood_gdf = flood_gdf.merge(
                    district_lookup,
                    left_on="_flood_row_id",
                    right_index=True,
                    how="left",
                )

                flood_gdf = flood_gdf.drop(
                    columns=["_flood_row_id"]
                )

                matched = (
                    flood_gdf[DISTRICT_NAME_FIELD]
                    .notna()
                    .sum()
                )

                if matched > 0:

                    joined_from_municipality = True
                    district_field_source = "municipality_spatial_join"

                    st.sidebar.info(
                        "flood.geojson has no District column. "
                        "District was estimated using a spatial "
                        f"join with municipality.geojson for "
                        f"{matched}/{len(flood_gdf)} features."
                    )

                    unmatched = (
                        len(flood_gdf) - matched
                    )

                    if unmatched > 0:

                        flood_gdf[DISTRICT_NAME_FIELD] = (
                            flood_gdf[DISTRICT_NAME_FIELD]
                            .fillna("Unknown")
                        )

                        st.sidebar.caption(
                            f"{unmatched} feature(s) could not "
                            "be matched and are labeled "
                            "'Unknown'."
                        )

            except Exception as e:

                st.sidebar.warning(
                    "Spatial join with municipality.geojson "
                    f"failed: {e}"
                )

        if not joined_from_municipality:

            flood_gdf = flood_gdf.reset_index(drop=True)

            flood_gdf[DISTRICT_NAME_FIELD] = [
                f"Zone {i + 1}"
                for i in range(len(flood_gdf))
            ]

            district_field_source = "auto_generated"

            st.sidebar.warning(
                "flood.geojson has no District column. "
                "Temporary names (Zone 1, Zone 2, ...) were "
                "generated so the app can run. These will not "
                "automatically match District_Situation, Water "
                "Access, or Funding data. Replacing flood.geojson "
                "with real administrative District boundaries is "
                "recommended."
            )

    # --------------------------------------------------------
    # Validate District field (should now always exist)
    # --------------------------------------------------------

    if (
        DISTRICT_NAME_FIELD
        not in flood_gdf.columns
    ):

        st.error(
            "Could not create a District column for "
            "flood.geojson."
        )

        st.write(
            "Current flood.geojson columns:"
        )

        st.code(
            "\n".join(
                str(c)
                for c in flood_gdf.columns
            )
        )

        st.stop()

    flood_gdf = prepare_flood_gdf(
        flood_gdf
    )

    if flood_notes:

        with st.sidebar.expander(
            "flood.geojson processing log",
            expanded=False,
        ):

            for note in flood_notes:
                st.caption(note)

    st.sidebar.caption(
        f"District features: "
        f"{len(flood_gdf):,}"
    )

    # --------------------------------------------------------
    # District situation
    # --------------------------------------------------------

    district_df, district_error = (
        load_district_situation(
            excel_path,
            district_csv_path,
        )
    )

    # --------------------------------------------------------
    # Join district data if available
    # --------------------------------------------------------

    JOIN_KEY = determine_join_key(
        flood_gdf,
        district_df,
    )

    join_status_message = None
    join_status_kind = "info"

    if JOIN_KEY:

        merged = join_district_data(
            flood_gdf,
            district_df,
            JOIN_KEY,
        )

        join_status_kind = "success"
        join_status_message = (
            f"District data joined on: {JOIN_KEY}"
        )

    else:

        # District Situation data is optional. The map still
        # works using flood.geojson attributes alone.

        merged = prepare_flood_gdf(
            flood_gdf
        )

        JOIN_KEY = (
            DISTRICT_ID_FIELD
            if DISTRICT_ID_FIELD
            in merged.columns
            else DISTRICT_NAME_FIELD
        )

        if (
            district_df is None
            or district_df.empty
        ):

            join_status_kind = "info"
            join_status_message = (
                "No District Situation data found. Using "
                "flood.geojson attributes only."
            )

        else:

            join_status_kind = "warning"
            join_status_message = (
                "No common join key between District Situation "
                "data and flood.geojson. Using flood.geojson "
                "attributes only."
            )

    # --------------------------------------------------------
    # Additional WASH data
    # --------------------------------------------------------

    water_access_df, water_access_err = (
        load_water_access(
            excel_path
        )
    )

    interventions_df, interventions_err = (
        load_interventions(
            excel_path
        )
    )

    funding_df, funding_err = (
        load_funding(
            excel_path
        )
    )

    facilities_df, facilities_err = (
        load_facilities(
            excel_path
        )
    )

    # --------------------------------------------------------
    # Corridor and markers
    # --------------------------------------------------------

    corridor_df, corridor_error = (
        load_flood_corridor(
            corridor_path
        )
    )

    markers_df, markers_error = (
        load_event_markers(
            markers_path
        )
    )

    # ==========================================================
    # TOP OF PAGE: THE MAP COMES FIRST
    # ==========================================================

    st.subheader("Flood-Affected Districts Map")

    indicator_labels = [
        label
        for label, _, _ in INDICATOR_OPTIONS
    ]

    selected_label = st.selectbox(
        "Indicator shown on the map and in the table below",
        indicator_labels,
        index=0,
        key="selected_indicator",
    )

    selected_column, selected_kind = next(
        (col, kind)
        for label, col, kind in INDICATOR_OPTIONS
        if label == selected_label
    )

    merged, legend_info = compute_display_layer(
        merged,
        selected_column,
        selected_kind,
    )

    try:

        fmap = build_folium_map(
            merged,
            municipality_gdf,
            corridor_df,
            markers_df,
            facilities_df,
            extra_tooltip=(
                selected_column,
                selected_label,
            ),
        )

    except Exception as e:

        st.error("Could not build the map.")
        st.exception(e)
        st.stop()

    try:

        map_state = st_folium(
            fmap,
            use_container_width=True,
            height=650,
            key="flood_map",
        )

    except Exception as e:

        st.error("Could not display the map.")
        st.exception(e)
        st.stop()

    # --------------------------------------------------------
    # Legend for the currently selected indicator
    # --------------------------------------------------------

    if legend_info.get("kind") == "categorical":

        legend_columns = st.columns(
            len(legend_info["items"])
        )

        for column, (
            label,
            hex_color,
        ) in zip(
            legend_columns,
            legend_info["items"],
        ):

            column.markdown(
                f"""
                <div>
                    <span style="
                        display:inline-block;
                        width:14px;
                        height:14px;
                        background:{hex_color};
                        border-radius:3px;
                        margin-right:6px;
                    "></span>
                    {label}
                </div>
                """,
                unsafe_allow_html=True,
            )

    elif legend_info.get("kind") == "numeric":

        gradient = ", ".join(
            legend_info["colors"]
        )

        st.markdown(
            f"""
            <div style="
                background: linear-gradient(
                    to right, {gradient}
                );
                height: 14px;
                border-radius: 3px;
            "></div>
            <div style="
                display:flex;
                justify-content:space-between;
                font-size: 0.85rem;
                margin-top: 2px;
            ">
                <span>{format_number(legend_info['vmin'])}</span>
                <span>{format_number(legend_info['vmax'])}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.caption(
            "Darker red means a higher value. Gray means no "
            "data has been entered yet for that District."
        )

    else:

        st.info(
            "No data available yet for the selected indicator. "
            "Fill in the District_Situation sheet to see it here."
        )

    st.caption(
        "Thick border = District with flood impact data. "
        "Dashed line = Municipality boundary. "
        "Use the layer control (top right of the map) to toggle "
        "Districts, Municipalities, and Facilities."
    )

    # --------------------------------------------------------
    # Map click detection
    # --------------------------------------------------------

    clicked_props = None
    selection_type = None

    if (
        map_state
        and map_state.get(
            "last_active_drawing"
        )
    ):

        active_drawing = (
            map_state[
                "last_active_drawing"
            ]
        )

        if isinstance(
            active_drawing,
            dict,
        ):

            clicked_props = (
                active_drawing.get(
                    "properties",
                    {},
                )
                or {}
            )

            if (
                MUNICIPALITY_NAME_FIELD
                in clicked_props
            ):

                selection_type = (
                    "municipality"
                )

            elif (
                DISTRICT_NAME_FIELD
                in clicked_props
            ):

                selection_type = (
                    "district"
                )

    # ==========================================================
    # PAGE HEADER AND EVENT INFORMATION (shown below the map)
    # ==========================================================

    event_info, event_info_error = (
        load_event_info(
            event_info_path
        )
    )

    page_title = event_info.get(
        "Title",
        "UNICEF Nepal - Flood WASH Dashboard",
    )

    page_subtitle = event_info.get(
        "Subtitle",
        "",
    )

    st.markdown("---")

    st.title(page_title)

    if page_subtitle:

        st.markdown(f"### {page_subtitle}")

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
        metadata.append(str(corridor_note))

    if data_asof:
        metadata.append(str(data_asof))

    if metadata:
        st.caption(" | ".join(metadata))

    st.caption(
        "Select a District or Municipality on the map above to "
        "see its detailed situation and response plan below."
    )

    if join_status_message:

        if join_status_kind == "success":
            st.sidebar.success(join_status_message)
        elif join_status_kind == "warning":
            st.sidebar.warning(join_status_message)
        else:
            st.sidebar.info(join_status_message)

    if district_field_source == "auto_generated":

        st.warning(
            "The District boundaries currently shown are "
            "temporary placeholders (Zone 1, Zone 2, ...) "
            "because flood.geojson has no real administrative "
            "District information. To connect District_Situation "
            "and other data automatically, add a real District "
            "name column to flood.geojson."
        )

    # --------------------------------------------------------
    # Join quality check (District Situation vs flood.geojson)
    # --------------------------------------------------------

    if (
        district_df is not None
        and not district_df.empty
        and JOIN_KEY
        and JOIN_KEY in flood_gdf.columns
        and JOIN_KEY in district_df.columns
    ):

        try:

            geo_values = set(
                flood_gdf[JOIN_KEY]
                .dropna()
                .astype(str)
                .str.strip()
            )

            district_values = set(
                district_df[JOIN_KEY]
                .dropna()
                .astype(str)
                .str.strip()
            )

            missing_in_district = (
                geo_values - district_values
            )

            missing_in_geojson = (
                district_values - geo_values
            )

            if (
                missing_in_district
                or missing_in_geojson
            ):

                with st.expander(
                    "Join quality check",
                    expanded=False,
                ):

                    if missing_in_district:

                        st.warning(
                            "In flood.geojson but missing from "
                            "District Situation data:"
                        )

                        st.code(
                            "\n".join(
                                sorted(missing_in_district)
                            )
                        )

                    if missing_in_geojson:

                        st.warning(
                            "In District Situation data but "
                            "missing from flood.geojson:"
                        )

                        st.code(
                            "\n".join(
                                sorted(missing_in_geojson)
                            )
                        )

        except Exception:
            pass

    # ==========================================================
    # DETAIL PANEL
    #
    # Organized around the five sections requested by the
    # response team:
    #   1. Affected population & current situation
    #   2. Problem / Needs
    #   3. Required solution
    #   4. Estimated financial requirement
    #   5. Priority areas
    # ==========================================================

    st.markdown("## District / Municipality Details")

    # ----------------------------------------------------------
    # Municipality
    # ----------------------------------------------------------

    if selection_type == "municipality":

        selected_properties = clicked_props

        muni_name = format_text(
            selected_properties.get(
                MUNICIPALITY_NAME_FIELD
            )
        )

        muni_district_name = (
            selected_properties.get(
                DISTRICT_NAME_FIELD
            )
        )

        muni_district_id = (
            selected_properties.get(
                DISTRICT_ID_FIELD
            )
        )

        if muni_district_name:
            parent_name = format_text(muni_district_name)
        else:
            parent_name = format_text(muni_district_id)

        st.markdown(f"### {muni_name}")

        st.caption(f"Parent District: {parent_name}")

        muni_situation_df, muni_situation_err = (
            load_municipality_situation(
                municipality_situation_path
            )
        )

        if (
            muni_situation_err
            or muni_situation_df.empty
        ):

            st.info(
                "municipality_situation.csv not found. "
                "Showing raw map attributes only."
            )

            display_props = {
                k: v
                for k, v in selected_properties.items()
                if k
                not in [
                    "severity_hex",
                    "is_flood_affected",
                    "display_hex",
                ]
            }

            st.json(display_props)

        else:

            sub = filter_municipality_rows(
                muni_situation_df,
                selected_properties,
            )

            if sub.empty:

                st.info(
                    "No detailed data found for this "
                    "Municipality."
                )

            else:

                st.dataframe(
                    sub.reset_index(drop=True),
                    use_container_width=True,
                    hide_index=True,
                )

    # ----------------------------------------------------------
    # District
    # ----------------------------------------------------------

    elif selection_type == "district":

        selected_properties = clicked_props

        district_name = format_text(
            selected_properties.get(
                DISTRICT_NAME_FIELD,
                "Unknown District",
            )
        )

        severity = format_text(
            selected_properties.get(
                "Severity",
                "No Data",
            )
        )

        st.markdown(f"### {district_name}")

        severity_color = SEVERITY_HEX.get(
            severity,
            DEFAULT_SEVERITY_HEX,
        )

        st.markdown(
            f"""
            <div style="
                display:inline-block;
                background:{severity_color};
                color:white;
                padding:4px 12px;
                border-radius:12px;
                font-weight:600;
                margin-bottom:10px;
            ">
                Severity: {severity}
            </div>
            """,
            unsafe_allow_html=True,
        )

        (
            tab_situation,
            tab_problem,
            tab_solution,
            tab_funding,
            tab_priority,
            tab_facilities,
            tab_sources,
            tab_population,
        ) = st.tabs(
            [
                "Affected population & current situation",
                "Problem / Needs",
                "Required solution",
                "Estimated financial requirement",
                "Priority areas",
                "Facilities",
                "Sources",
                "Population",
            ]
        )

        # --------------------------------------------------
        # 1. Affected population & current situation
        # --------------------------------------------------

        with tab_situation:

            st.caption(
                "Number of affected people and where/how they "
                "are currently accessing drinking water."
            )

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
                "Severity",
                severity,
            )

            st.markdown("---")

            st.markdown("**Current water source**")

            st.write(
                format_text(
                    selected_properties.get(
                        "Water_source_current"
                    )
                )
            )

            render_sheet_table(
                water_access_df,
                water_access_err,
                selected_properties,
                JOIN_KEY,
                "No Water Access data found for this "
                "District.",
            )

        # --------------------------------------------------
        # 2. Problem / Needs
        # --------------------------------------------------

        with tab_problem:

            st.caption(
                "Key issues related to water availability, "
                "quantity, quality, and safety."
            )

            st.markdown("**Key problems / needs**")

            st.write(
                format_text(
                    selected_properties.get(
                        "Key_problems"
                    )
                )
            )

            st.markdown("**Severity reason**")

            severity_reason = (
                selected_properties.get(
                    "Severity_reason"
                )
            )

            if not severity_reason:

                severity_reason = (
                    selected_properties.get(
                        "Priority_reason"
                    )
                )

            st.write(format_text(severity_reason))

        # --------------------------------------------------
        # 3. Required solution
        # --------------------------------------------------

        with tab_solution:

            st.caption(
                "Proposed immediate, medium-term, and "
                "long-term interventions."
            )

            st.markdown("**Immediate intervention**")

            st.write(
                format_text(
                    selected_properties.get(
                        "Immediate_intervention"
                    )
                )
            )

            st.markdown("**Medium-term intervention**")

            st.write(
                format_text(
                    selected_properties.get(
                        "Medium_term_intervention"
                    )
                )
            )

            st.markdown("**Long-term intervention**")

            st.write(
                format_text(
                    selected_properties.get(
                        "Long_term_intervention"
                    )
                )
            )

            render_sheet_table(
                interventions_df,
                interventions_err,
                selected_properties,
                JOIN_KEY,
                "No Interventions data found for this "
                "District.",
            )

        # --------------------------------------------------
        # 4. Estimated financial requirement
        # --------------------------------------------------

        with tab_funding:

            st.caption(
                "Approximate funding required for the "
                "proposed response."
            )

            st.metric(
                "Estimated funding required",
                format_currency(
                    selected_properties.get(
                        "Estimated_funding_USD"
                    )
                ),
            )

            render_sheet_table(
                funding_df,
                funding_err,
                selected_properties,
                JOIN_KEY,
                "No Funding data found for this District.",
            )

        # --------------------------------------------------
        # 5. Priority areas
        # --------------------------------------------------

        with tab_priority:

            st.caption(
                "Prioritization based on affected population, "
                "water requirement, water safety, host "
                "community needs, and related factors."
            )

            st.markdown("**Priority reason**")

            st.write(
                format_text(
                    selected_properties.get(
                        "Priority_reason"
                    )
                )
            )

            st.markdown("**Host community needs**")

            st.write(
                format_text(
                    selected_properties.get(
                        "Host_community_needs"
                    )
                )
            )

            st.info(
                "Schools and health facilities are available as "
                "separate, toggleable layers on the map above "
                "(see the layer control)."
            )

        # --------------------------------------------------
        # Facilities table for this District
        # --------------------------------------------------

        with tab_facilities:

            render_sheet_table(
                facilities_df,
                facilities_err,
                selected_properties,
                JOIN_KEY,
                "No Facilities data found for this District.",
            )

        # --------------------------------------------------
        # Sources
        # --------------------------------------------------

        with tab_sources:

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

            st.markdown("**District Situation source**")

            st.caption(f"Data source: {data_source}")
            st.caption(f"Last updated: {last_updated}")
            st.caption(f"Data status: {data_status}")

            project_sources = event_info.get("Sources")

            if project_sources:

                st.markdown("---")
                st.markdown("**Project-level sources**")
                st.write(format_text(project_sources))

            other_sources = collect_sources(
                selected_properties,
                JOIN_KEY,
                water_access_df,
                interventions_df,
                funding_df,
                facilities_df,
            )

            if other_sources:

                st.markdown("---")
                st.markdown("**Other sheet sources**")

                st.dataframe(
                    pd.DataFrame(other_sources),
                    use_container_width=True,
                    hide_index=True,
                )

        # --------------------------------------------------
        # Population
        # --------------------------------------------------

        with tab_population:

            st.markdown("**population.csv data**")

            if population_df.empty:

                st.info("population.csv not found.")

            else:

                pop_sub = filter_data_by_selection(
                    population_df,
                    selected_properties,
                    JOIN_KEY,
                )

                if (
                    pop_sub.empty
                    and DISTRICT_NAME_FIELD
                    in population_df.columns
                ):

                    pop_sub = population_df[
                        population_df[
                            DISTRICT_NAME_FIELD
                        ]
                        .astype(str)
                        .str.strip()
                        == district_name
                    ]

                if pop_sub.empty:

                    st.info(
                        "No population.csv data found for "
                        "this District."
                    )

                else:

                    st.dataframe(
                        pop_sub.reset_index(drop=True),
                        use_container_width=True,
                        hide_index=True,
                    )

                    numeric_cols = [
                        c
                        for c in pop_sub.columns
                        if pd.api.types.is_numeric_dtype(
                            pop_sub[c]
                        )
                    ]

                    if numeric_cols:

                        cols = st.columns(
                            min(len(numeric_cols), 4)
                        )

                        for i, cname in enumerate(
                            numeric_cols
                        ):

                            cols[
                                i % len(cols)
                            ].metric(
                                cname,
                                format_number(
                                    pop_sub[cname].sum()
                                ),
                            )

    # ----------------------------------------------------------
    # Nothing selected
    # ----------------------------------------------------------

    else:

        st.info(
            "Click a District or Municipality on the map above "
            "to see its detailed situation and response plan."
        )

    # ==========================================================
    # Full indicator table (all districts, current indicator)
    #
    # Shown regardless of whether a District is selected, so
    # the response team can compare all Districts at a glance.
    # ==========================================================

    st.markdown("---")

    st.markdown("## All Districts")

    st.caption(
        f"Sorted by the indicator currently shown on the map: "
        f"{selected_label}"
    )

    table_columns = [
        c
        for c in [
            DISTRICT_NAME_FIELD,
            DISTRICT_ID_FIELD,
            "Severity",
            "Affected_population",
            "Water_access_population",
            "Casualties",
            "Estimated_funding_USD",
            "Priority_reason",
        ]
        if c in merged.columns
    ]

    display_table = pd.DataFrame(
        merged.drop(columns="geometry")
    )[table_columns].copy()

    if selected_column in display_table.columns:

        if selected_kind == "numeric":

            display_table["_sort_key"] = pd.to_numeric(
                display_table[selected_column],
                errors="coerce",
            )

            display_table = display_table.sort_values(
                "_sort_key",
                ascending=False,
                na_position="last",
            ).drop(columns="_sort_key")

        else:

            severity_order = {
                label: i
                for i, (label, _) in enumerate(
                    SEVERITY_LEGEND
                )
            }

            display_table["_sort_key"] = (
                display_table[selected_column]
                .map(severity_order)
                .fillna(len(SEVERITY_LEGEND))
            )

            display_table = display_table.sort_values(
                "_sort_key"
            ).drop(columns="_sort_key")

    for col in [
        "Affected_population",
        "Water_access_population",
        "Casualties",
    ]:

        if col in display_table.columns:

            display_table[col] = display_table[
                col
            ].apply(format_number)

    if "Estimated_funding_USD" in display_table.columns:

        display_table["Estimated_funding_USD"] = (
            display_table["Estimated_funding_USD"].apply(
                format_currency
            )
        )

    st.dataframe(
        display_table.reset_index(drop=True),
        use_container_width=True,
        hide_index=True,
    )

    st.caption(
        "Values update automatically once "
        "WASH_data_template.xlsx (District_Situation sheet) or "
        "district_situation.csv is filled in and Refresh data "
        "is clicked in the sidebar."
    )

    # --------------------------------------------------------
    # Footer
    # --------------------------------------------------------

    st.markdown("---")

    st.caption(
        "UNICEF Nepal | WASH Emergency Response"
    )

    district_layer_name = (
        Path(geojson_path).name
        if geojson_path
        else "not loaded"
    )

    municipality_layer_name = (
        Path(municipality_path).name
        if (
            municipality_gdf is not None
            and municipality_path
        )
        else "not loaded"
    )

    st.caption(
        f"District layer: {district_layer_name} | "
        f"Municipality layer: {municipality_layer_name}"
    )

    st.caption(
        f"Base map tiles: {BASE_MAP_TILES} "
        "(no API key required)"
    )


# ============================================================
# 27. Run
# ============================================================

if __name__ == "__main__":
    main()
