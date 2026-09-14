# -*- coding: utf-8 -*-
"""
UNICEF Nepal - Flood WASH Emergency Decision Support Dashboard

Streamlit + GeoPandas + Folium

Purpose
-------
Help decision makers answer, for any selected District or
Municipality:
    1. What is the current situation?
    2. Who and how many people are affected?
    3. What WASH infrastructure has been damaged?
    4. What has UNICEF / partners already done?
    5. What are the remaining gaps?
    6. What needs to be done next?
    7. How much funding is required?
    8. Where should resources be prioritized?

Design principles
------------------
- The map is interactive: clicking a District or Municipality
  updates every section below it.
- No paid API or API key is required. The base map uses free
  CartoDB / OpenStreetMap tiles. If a Mapbox or other token is
  ever configured, it is read from Streamlit secrets only -
  never hard-coded.
- Missing data never crashes the app. Missing values show
  "N/A", missing cost figures show "Cost data not available",
  and missing files are reported in the sidebar "Data status"
  panel instead of stopping the app.
- District matching uses a three-tier strategy:
      1) District_ID (exact)
      2) District name (exact, case/whitespace normalized)
      3) District name (loosely normalized: lowercase, spaces
         and punctuation removed)
  so that small naming differences between flood.geojson and
  the Excel/CSV data do not break the join.
- No numbers are invented. If a figure is not present in the
  source data, the dashboard says so explicitly instead of
  guessing.

How to update the data later
-----------------------------
See the "How to update the data" panel in the sidebar for full
instructions. In short: edit the CSV files (or the matching
sheets in WASH_data_template.xlsx) in the data/ folder, push to
the GitHub repository, then click "Refresh data" in the sidebar
(or redeploy on Streamlit Cloud).
"""

import os
import re
import io
import json
from pathlib import Path
from datetime import datetime

import pandas as pd
import geopandas as gpd
import folium
import branca.colormap as bcm
import streamlit as st
from streamlit_folium import st_folium


# ============================================================
# 0. Streamlit page configuration
# ============================================================

st.set_page_config(
    page_title="UNICEF Nepal - Flood WASH Dashboard",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# 1. File paths
#
# Every path below can be overridden from the sidebar, so these
# are only the defaults used when the repository follows the
# recommended folder layout.
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

FLOOD_GEOJSON_PATH = DATA_DIR / "flood.geojson"
MUNICIPALITY_GEOJSON_PATH = DATA_DIR / "municipality.geojson"
RIVERS_GEOJSON_PATH = DATA_DIR / "rivers.geojson"

POPULATION_PATH = DATA_DIR / "population.csv"
MUNICIPALITY_SITUATION_PATH = DATA_DIR / "municipality_situation.csv"

DEFAULT_EXCEL_PATH = DATA_DIR / "WASH_data_template.xlsx"

DISTRICT_SITUATION_CSV_PATH = DATA_DIR / "district_situation.csv"
WATER_ACCESS_CSV_PATH = DATA_DIR / "water_access.csv"
INTERVENTIONS_CSV_PATH = DATA_DIR / "interventions.csv"
RESPONSE_CSV_PATH = DATA_DIR / "response.csv"
RESPONSE_PLAN_CSV_PATH = DATA_DIR / "response_plan.csv"
BUDGET_CSV_PATH = DATA_DIR / "budget.csv"
FACILITIES_CSV_PATH = DATA_DIR / "facilities.csv"

EVENT_INFO_PATH = DATA_DIR / "event_info.csv"
EVENT_MARKERS_PATH = DATA_DIR / "event_markers.csv"
FLOOD_CORRIDOR_PATH = DATA_DIR / "flood_corridor_path.csv"


# ============================================================
# 2. Field names
# ============================================================

DISTRICT_ID_FIELD = "District_ID"
DISTRICT_NAME_FIELD = "District"

MUNICIPALITY_ID_FIELD = "Municipality_ID"
MUNICIPALITY_NAME_FIELD = "Municipality"


# ============================================================
# 3. Map defaults (no API key needed for any of these tiles)
# ============================================================

DEFAULT_LAT = 28.05
DEFAULT_LON = 84.85
DEFAULT_ZOOM = 8

# Free, key-less basemap choices offered in the sidebar.
# Google Satellite is provided as an (tile_url, attribution)
# pair since folium has no built-in name for it.
BASEMAP_OPTIONS = {
    "CartoDB Positron (light)": "CartoDB positron",
    "OpenStreetMap": "OpenStreetMap",
    "CartoDB Dark Matter": "CartoDB dark_matter",
    "Google Satellite": (
        "https://{s}.google.com/vt/lyrs=s&x={x}&y={y}&z={z}",
        "Imagery &copy; Google",
    ),
}

# If a Mapbox / other token is later added to Streamlit secrets,
# it is read here - never hard-coded in the source file.
def get_secret(name):
    try:
        return st.secrets.get(name)
    except Exception:
        return None


# ============================================================
# 4. Severity legend and map indicators
# ============================================================

SEVERITY_LEGEND = [
    ("Severe", "#7A0F0F"),
    ("Very High", "#D62728"),
    ("High", "#F4912D"),
    ("Moderate", "#FFDD57"),
    ("No Data", "#C8C5BD"),
]

SEVERITY_HEX = dict(SEVERITY_LEGEND)
DEFAULT_SEVERITY_HEX = SEVERITY_HEX["No Data"]

# (label, column, kind) - kind is "categorical" or "numeric"
INDICATOR_OPTIONS = [
    ("Severity", "Severity", "categorical"),
    ("Affected population", "Affected_Population", "numeric"),
    ("Affected households", "Affected_HH", "numeric"),
    ("People currently reached", "People_Reached", "numeric"),
    ("Remaining population gap", "Population_Gap", "numeric"),
    (
        "Estimated budget requirement (USD)",
        "Estimated_Budget_Requirement_USD",
        "numeric",
    ),
]

NUMERIC_COLOR_RAMP = [
    "#FFF3B0",
    "#FFDD57",
    "#F4912D",
    "#D62728",
    "#7A0F0F",
]

MARKER_HEX = {"epicentre": "#7A0F0F", "annotation": "#282828"}
DEFAULT_MARKER_HEX = "#505050"
CORRIDOR_HEX = "#1E5AC8"
RIVER_HEX = "#2E6F95"

FACILITY_HEX = {"school": "#1E5AC8", "health": "#2E9E5B"}
DEFAULT_FACILITY_HEX = "#505050"


# ============================================================
# 5. Excel sheet names
#
# These are the sheet names used both by the legacy
# WASH_data_template.xlsx (matching CSV file takes priority
# when no workbook has been uploaded - see load_optional_table)
# and by the consolidated Data_Entry.xlsx workbook that can be
# uploaded from the sidebar "Load Excel Files" section (which,
# once uploaded, takes priority over everything else - see
# load_table_with_upload_priority).
# ============================================================

SHEET_INSTRUCTIONS = "Instructions"
SHEET_GENERAL_INFO = "General_Information"
SHEET_DISTRICT_SITUATION = "District_Situation"
SHEET_WATER_ACCESS = "Water_Access"
SHEET_INTERVENTIONS = "Interventions"
SHEET_RESPONSE = "Response"
SHEET_RESPONSE_PLAN = "Response_Plan"
SHEET_BUDGET = "Budget"
SHEET_FACILITIES = "Facilities"
SHEET_POPULATION = "Population"
SHEET_MUNICIPALITY_DATA = "Municipality_Data"
SHEET_EVENT_MARKERS = "Event_Markers"
SHEET_FLOOD_CORRIDOR = "Flood_Corridor"

DATA_ENTRY_DEFAULT_FILENAME = "Data_Entry.xlsx"

# Sheets that are actively used to drive the dashboard (used for
# upload validation/status reporting). Instructions is a
# reference-only sheet and is intentionally excluded.
DATA_ENTRY_SHEETS = [
    SHEET_GENERAL_INFO,
    SHEET_DISTRICT_SITUATION,
    SHEET_WATER_ACCESS,
    SHEET_INTERVENTIONS,
    SHEET_RESPONSE,
    SHEET_RESPONSE_PLAN,
    SHEET_BUDGET,
    SHEET_FACILITIES,
    SHEET_POPULATION,
    SHEET_MUNICIPALITY_DATA,
    SHEET_EVENT_MARKERS,
    SHEET_FLOOD_CORRIDOR,
]

# Minimum columns each sheet needs to be usable. Any other
# column is free-form and simply passed through to the relevant
# table/tab. This is intentionally a light check - the goal is
# a clear message, not to reject a whole sheet.
REQUIRED_COLUMNS_BY_SHEET = {
    SHEET_GENERAL_INFO: ["Key", "Value"],
    SHEET_DISTRICT_SITUATION: [DISTRICT_NAME_FIELD],
    SHEET_MUNICIPALITY_DATA: [MUNICIPALITY_NAME_FIELD],
    SHEET_FACILITIES: ["Latitude", "Longitude"],
    SHEET_EVENT_MARKERS: ["Latitude", "Longitude", "Label"],
    SHEET_FLOOD_CORRIDOR: ["Latitude", "Longitude"],
    SHEET_POPULATION: [DISTRICT_NAME_FIELD],
}


# ============================================================
# 6. Formatting helpers
#
#   - format_number / format_text  -> "N/A" when missing
#   - format_cost                  -> "Cost data not available"
#   - format_percent               -> "N/A" when missing
# ============================================================

def _is_missing(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    if isinstance(value, str) and not value.strip():
        return True
    return False


def format_number(value):
    if _is_missing(value):
        return "N/A"
    try:
        return f"{int(float(value)):,}"
    except (ValueError, TypeError):
        return str(value)


def format_percent(value):
    if _is_missing(value):
        return "N/A"
    try:
        return f"{float(value):.1f}%"
    except (ValueError, TypeError):
        return str(value)


def format_cost(value):
    if _is_missing(value):
        return "Cost data not available"
    try:
        return f"USD {int(float(value)):,}"
    except (ValueError, TypeError):
        return str(value)


def format_text(value):
    if _is_missing(value):
        return "N/A"
    return str(value).strip()


# ============================================================
# 7. Column normalization
# ============================================================

_COLUMN_ALIASES = {
    DISTRICT_ID_FIELD: [
        "district_id", "districtid", "district id",
    ],
    DISTRICT_NAME_FIELD: [
        "district", "district_name", "district name",
    ],
    MUNICIPALITY_ID_FIELD: [
        "municipality_id", "municipalityid", "municipality id",
    ],
    MUNICIPALITY_NAME_FIELD: [
        "municipality", "municipality_name", "municipality name",
    ],
}


def normalize_column_names(df):
    if df is None or df.empty:
        return df

    df = df.copy()
    rename_map = {}

    for col in df.columns:
        clean = str(col).strip()
        lower = clean.lower()
        target = clean

        for canonical, aliases in _COLUMN_ALIASES.items():
            if lower in aliases:
                target = canonical
                break

        rename_map[col] = target

    return df.rename(columns=rename_map)


def clean_str_series(series):
    """
    Returns a plain, object-dtype string Series where missing
    values are real Python None - never pandas' nullable pd.NA
    sentinel.

    This matters because GeoPandas' GeoDataFrame.to_json() (used
    every time the interactive map is built) serializes feature
    properties with Python's standard json module, which cannot
    encode pd.NA and raises TypeError: "Object of type NAType is
    not JSON serializable". Using pandas' nullable "string" dtype
    for District/Municipality identifier columns silently
    introduced pd.NA values, which broke map generation entirely
    whenever any District/Municipality identifier was missing.
    This helper is the fix: always use it instead of
    `.astype("string")` for identifier columns that end up in a
    GeoDataFrame passed to Folium.
    """
    cleaned = series.astype(str).str.strip()
    blank_mask = cleaned.str.lower().isin(
        ["nan", "none", "null", "", "<na>"]
    )
    return cleaned.mask(blank_mask, None)


def normalize_identifier_columns(df):
    if df is None or df.empty:
        return df

    df = normalize_column_names(df)

    id_columns = [
        DISTRICT_ID_FIELD,
        DISTRICT_NAME_FIELD,
        MUNICIPALITY_ID_FIELD,
        MUNICIPALITY_NAME_FIELD,
    ]

    for col in id_columns:
        if col not in df.columns:
            continue

        df[col] = clean_str_series(df[col])

    return df


def normalize_name_key(value):
    """
    Loosest-tier District/Municipality name match: lowercase,
    strip, remove anything that is not a letter or digit. Used
    only as a last-resort fallback when exact matches fail.
    """
    if _is_missing(value):
        return ""
    text = str(value).strip().lower()
    return re.sub(r"[^a-z0-9]", "", text)


# ============================================================
# 8. File / cache helpers
# ============================================================

def _file_signature(path):
    try:
        return (str(path), os.path.getmtime(path))
    except OSError:
        return (str(path), None)


@st.cache_data(show_spinner=False)
def _read_csv_cached(path, signature):
    return pd.read_csv(path)


def safe_read_csv(path):
    if not path:
        return pd.DataFrame(), "No file path was provided."

    if not os.path.exists(path):
        return pd.DataFrame(), f"File not found: {path}"

    try:
        df = _read_csv_cached(str(path), _file_signature(path))
        return normalize_identifier_columns(df), None
    except Exception as e:
        return pd.DataFrame(), f"Error reading CSV: {e}"


@st.cache_data(show_spinner=False)
def _read_excel_sheet_cached(path, sheet_name, signature):
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")


def safe_read_excel_sheet(path, sheet_name):
    if not path:
        return pd.DataFrame(), "No Excel file path was provided."

    if not os.path.exists(path):
        return pd.DataFrame(), f"File not found: {path}"

    try:
        df = _read_excel_sheet_cached(
            str(path), sheet_name, _file_signature(path)
        )
        return normalize_identifier_columns(df), None
    except ValueError as e:
        return pd.DataFrame(), f"Sheet '{sheet_name}' not found: {e}"
    except Exception as e:
        return pd.DataFrame(), f"Error reading sheet '{sheet_name}': {e}"


def load_optional_table(csv_path, excel_path, sheet_name):
    """
    Generic loader used for every optional WASH data table.
    Preference order: dedicated CSV file, then the matching
    sheet inside the Excel template. Never raises - always
    returns (DataFrame, error_message_or_None).
    """
    if csv_path and os.path.exists(csv_path):
        return safe_read_csv(csv_path)

    if excel_path and os.path.exists(excel_path):
        return safe_read_excel_sheet(excel_path, sheet_name)

    return (
        pd.DataFrame(),
        f"No data found for '{sheet_name}' "
        f"(checked CSV file and the Excel template).",
    )


def read_sheet_from_uploaded_workbook(sheet_name):
    """
    Reads one sheet from the Data_Entry.xlsx workbook uploaded via
    the sidebar "Load Excel Files" section (its raw bytes are kept
    in st.session_state). Returns (DataFrame, error_message_or_None).
    Returns (None, None) when no workbook has been uploaded, so the
    caller can fall back to the previous CSV/Excel-path behavior.
    """
    workbook_bytes = st.session_state.get("data_entry_bytes")

    if not workbook_bytes:
        return None, None

    try:
        df = pd.read_excel(
            io.BytesIO(workbook_bytes), sheet_name=sheet_name, engine="openpyxl"
        )
        return normalize_identifier_columns(df), None
    except ValueError:
        return (
            pd.DataFrame(),
            f"Sheet '{sheet_name}' was not found in the uploaded "
            f"{DATA_ENTRY_DEFAULT_FILENAME}.",
        )
    except Exception as e:
        return (
            pd.DataFrame(),
            f"Error reading sheet '{sheet_name}' from the uploaded "
            f"{DATA_ENTRY_DEFAULT_FILENAME}: {e}",
        )


def load_table_with_upload_priority(sheet_name, fallback_csv_path, fallback_excel_path):
    """
    Single Source of Truth loader: if a Data_Entry.xlsx workbook has
    been uploaded from the sidebar, its matching sheet is always
    used. Otherwise, behavior is unchanged from before - the
    dedicated CSV file, then the default Excel template.
    """
    df, err = read_sheet_from_uploaded_workbook(sheet_name)

    if df is not None:
        return df, err

    return load_optional_table(fallback_csv_path, fallback_excel_path, sheet_name)


def validate_uploaded_workbook(workbook_bytes):
    """
    Opens the uploaded workbook once and checks, for every sheet the
    dashboard actually uses, whether the sheet exists and whether its
    required columns are present. Never raises - returns a list of
    human-readable issue strings (empty list means no issues found)
    plus a summary dict used for the sidebar upload status display.
    """
    issues = []
    summary = {
        "sheets_found": 0,
        "sheets_expected": len(DATA_ENTRY_SHEETS),
        "total_records": 0,
        "districts": set(),
        "municipalities": set(),
    }

    try:
        workbook = pd.ExcelFile(io.BytesIO(workbook_bytes), engine="openpyxl")
    except Exception as e:
        return [f"Could not open the uploaded file as an Excel workbook: {e}"], summary

    available_sheets = set(workbook.sheet_names)

    for sheet_name in DATA_ENTRY_SHEETS:
        if sheet_name not in available_sheets:
            issues.append(
                f"Missing sheet: '{sheet_name}' was not found in the "
                f"uploaded {DATA_ENTRY_DEFAULT_FILENAME}."
            )
            continue

        try:
            df = workbook.parse(sheet_name)
        except Exception as e:
            issues.append(f"Could not read sheet '{sheet_name}': {e}")
            continue

        summary["sheets_found"] += 1
        summary["total_records"] += len(df)

        required_columns = REQUIRED_COLUMNS_BY_SHEET.get(sheet_name, [])
        normalized_columns = {str(c).strip() for c in df.columns}

        for required_column in required_columns:
            if required_column not in normalized_columns:
                issues.append(
                    f"Missing column: the required column "
                    f"'{required_column}' is missing from the "
                    f"'{sheet_name}' sheet."
                )

        if DISTRICT_NAME_FIELD in df.columns:
            summary["districts"].update(
                df[DISTRICT_NAME_FIELD].dropna().astype(str).str.strip().unique()
            )

        if MUNICIPALITY_NAME_FIELD in df.columns:
            summary["municipalities"].update(
                df[MUNICIPALITY_NAME_FIELD]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
            )

        if sheet_name == SHEET_DISTRICT_SITUATION and DISTRICT_NAME_FIELD in df.columns:
            duplicate_ids = df[DISTRICT_NAME_FIELD][
                df[DISTRICT_NAME_FIELD].duplicated(keep=False)
                & df[DISTRICT_NAME_FIELD].notna()
            ].unique()
            if len(duplicate_ids) > 0:
                issues.append(
                    "Duplicate District entries in 'District_Situation': "
                    + ", ".join(str(v) for v in duplicate_ids)
                )

    return issues, summary


def load_event_info(csv_path, excel_path):
    df, err = load_table_with_upload_priority(SHEET_GENERAL_INFO, csv_path, excel_path)
    if err or df.empty:
        return {}, err
    if "Key" not in df.columns or "Value" not in df.columns:
        return {}, None
    return dict(zip(df["Key"].astype(str), df["Value"])), None


# ============================================================
# 9. GeoJSON reading, CRS handling, geometry cleaning
# ============================================================

@st.cache_data(show_spinner=False)
def _read_geo_file_cached(path, signature):
    try:
        return gpd.read_file(path, engine="pyogrio")
    except Exception:
        return gpd.read_file(path)


def ensure_wgs84(gdf, label):
    notes = []

    if gdf.crs is None:
        notes.append(f"{label}: no CRS found, assumed EPSG:4326.")
        gdf = gdf.set_crs(epsg=4326, allow_override=True)
    else:
        try:
            epsg = gdf.crs.to_epsg()
        except Exception:
            epsg = None

        if epsg != 4326:
            notes.append(
                f"{label}: converted CRS from {gdf.crs} to EPSG:4326."
            )
            gdf = gdf.to_crs(epsg=4326)

    return gdf, notes


def clean_geometries(gdf, label):
    """
    Drops only the rows with missing/invalid geometry so a
    problem with one feature never takes down the rest of the
    dataset (or the app).
    """
    notes = []

    if gdf is None or gdf.empty:
        return gdf, notes

    before = len(gdf)
    gdf = gdf[gdf.geometry.notna()].copy()

    if gdf.empty:
        notes.append(f"{label}: no valid geometries found.")
        return gdf, notes

    try:
        invalid_mask = ~gdf.geometry.is_valid
        if invalid_mask.any():
            try:
                gdf.loc[invalid_mask, "geometry"] = (
                    gdf.loc[invalid_mask, "geometry"].buffer(0)
                )
            except Exception:
                pass
    except Exception:
        pass

    after = len(gdf)
    if after < before:
        notes.append(
            f"{label}: removed {before - after} invalid/empty geometries."
        )

    return gdf, notes


def read_geo_file(path, label):
    if not path:
        return None, [f"No path was provided for {label}."]

    if not os.path.exists(path):
        return None, [f"File not found for {label}: {path}"]

    try:
        gdf = _read_geo_file_cached(str(path), _file_signature(path))
    except Exception as e:
        return None, [f"Error reading {label}: {e}"]

    if gdf is None or gdf.empty:
        return None, [f"{label} contains no valid data."]

    if "geometry" not in gdf.columns:
        return None, [f"{label} has no geometry column."]

    gdf = normalize_identifier_columns(gdf)
    gdf, crs_notes = ensure_wgs84(gdf, label)
    gdf, geom_notes = clean_geometries(gdf, label)

    return gdf, crs_notes + geom_notes


# ============================================================
# 10. District data preparation
# ============================================================

# These columns line up with the reporting sections requested
# by the response team. If a column is missing from the source
# data it is created as empty, so downstream code never has to
# guess whether a column exists.
DISTRICT_DEFAULT_COLUMNS = [
    "Affected_Population",
    "Affected_HH",
    "Affected_Percentage",
    "Damaged_WASH_Systems",
    "Systems_Requiring_Repair",
    "Systems_Requiring_Rehabilitation",
    "People_Reached",
    "Population_Gap",
    "WASH_Supplies_Distributed",
    "Response_Status",
    "Estimated_Recovery_Requirement_USD",
    "Estimated_Budget_Requirement_USD",
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


def prepare_flood_gdf(flood_gdf):
    gdf = flood_gdf.copy()

    if DISTRICT_NAME_FIELD in gdf.columns:
        gdf[DISTRICT_NAME_FIELD] = clean_str_series(gdf[DISTRICT_NAME_FIELD])

    if DISTRICT_ID_FIELD in gdf.columns:
        gdf[DISTRICT_ID_FIELD] = clean_str_series(gdf[DISTRICT_ID_FIELD])

    # ---- Severity ----
    if "Severity" not in gdf.columns:
        gdf["Severity"] = gdf["Priority"] if "Priority" in gdf.columns else "No Data"

    gdf["Severity"] = gdf["Severity"].fillna("No Data").astype(str).str.strip()

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
        gdf["Severity"].str.lower().map(severity_mapping).fillna("No Data")
    )

    gdf["severity_hex"] = gdf["Severity"].map(SEVERITY_HEX).fillna(
        DEFAULT_SEVERITY_HEX
    )
    gdf["is_flood_affected"] = gdf["Severity"] != "No Data"

    # ---- Derived population gap, if not supplied directly ----
    for col in DISTRICT_DEFAULT_COLUMNS:
        if col not in gdf.columns:
            gdf[col] = None

    if gdf["Population_Gap"].isna().all():
        affected = pd.to_numeric(gdf["Affected_Population"], errors="coerce")
        reached = pd.to_numeric(gdf["People_Reached"], errors="coerce")
        computed_gap = affected - reached
        gdf["Population_Gap"] = gdf["Population_Gap"].where(
            gdf["Population_Gap"].notna(), computed_gap
        )

    return gdf


def _recover_boundary_fields(boundary_gdf, merged):
    for field in [DISTRICT_NAME_FIELD, DISTRICT_ID_FIELD]:
        boundary_field = f"{field}_boundary"
        if field not in merged.columns and boundary_field in merged.columns:
            merged[field] = merged[boundary_field]

    if DISTRICT_NAME_FIELD not in merged.columns:
        boundary_name = f"{DISTRICT_NAME_FIELD}_boundary"
        if boundary_name in merged.columns:
            merged[DISTRICT_NAME_FIELD] = merged[boundary_name]

    return merged


def robust_join_district_data(boundary_gdf, data_df):
    """
    Joins flood.geojson District boundaries with the District
    Situation table using a three-tier matching strategy:
        1) District_ID (exact)
        2) District name (exact, normalized)
        3) District name (loosely normalized fallback)

    Returns (merged_gdf, join_info) where join_info documents
    which tier succeeded and how many of the total District
    features were matched, so this can be shown to the user
    instead of silently guessing.
    """
    total = len(boundary_gdf) if boundary_gdf is not None else 0

    if boundary_gdf is None or boundary_gdf.empty:
        return boundary_gdf, {"tier": "none", "matched": 0, "total": total}

    if data_df is None or data_df.empty:
        return (
            prepare_flood_gdf(boundary_gdf),
            {"tier": "none", "matched": 0, "total": total},
        )

    data_df = data_df.copy()

    # ---- Tier 1: District_ID ----
    if (
        DISTRICT_ID_FIELD in boundary_gdf.columns
        and DISTRICT_ID_FIELD in data_df.columns
    ):
        b = boundary_gdf.copy()
        b[DISTRICT_ID_FIELD] = clean_str_series(b[DISTRICT_ID_FIELD])
        d = data_df.copy()
        d[DISTRICT_ID_FIELD] = clean_str_series(d[DISTRICT_ID_FIELD])
        d = d.drop_duplicates(subset=[DISTRICT_ID_FIELD], keep="first")

        merged = b.merge(d, on=DISTRICT_ID_FIELD, how="left", suffixes=("_boundary", ""))
        merged = _recover_boundary_fields(b, merged)

        matched = int(
            b[DISTRICT_ID_FIELD].dropna().isin(d[DISTRICT_ID_FIELD].dropna()).sum()
        )

        if matched > 0:
            return (
                prepare_flood_gdf(merged),
                {"tier": "District_ID", "matched": matched, "total": total},
            )

    # ---- Tier 2: District name, exact ----
    if (
        DISTRICT_NAME_FIELD in boundary_gdf.columns
        and DISTRICT_NAME_FIELD in data_df.columns
    ):
        b = boundary_gdf.copy()
        b[DISTRICT_NAME_FIELD] = clean_str_series(b[DISTRICT_NAME_FIELD])
        d = data_df.copy()
        d[DISTRICT_NAME_FIELD] = clean_str_series(d[DISTRICT_NAME_FIELD])
        d = d.drop_duplicates(subset=[DISTRICT_NAME_FIELD], keep="first")

        merged = b.merge(
            d, on=DISTRICT_NAME_FIELD, how="left", suffixes=("_boundary", "")
        )
        merged = _recover_boundary_fields(b, merged)

        matched = int(
            b[DISTRICT_NAME_FIELD]
            .dropna()
            .isin(d[DISTRICT_NAME_FIELD].dropna())
            .sum()
        )

        if matched > 0:
            return (
                prepare_flood_gdf(merged),
                {
                    "tier": "District name (exact)",
                    "matched": matched,
                    "total": total,
                },
            )

    # ---- Tier 3: District name, loosely normalized ----
    if (
        DISTRICT_NAME_FIELD in boundary_gdf.columns
        and DISTRICT_NAME_FIELD in data_df.columns
    ):
        b = boundary_gdf.copy()
        d = data_df.copy()

        b["_norm_key"] = b[DISTRICT_NAME_FIELD].map(normalize_name_key)
        d["_norm_key"] = d[DISTRICT_NAME_FIELD].map(normalize_name_key)

        d = d[d["_norm_key"] != ""].drop_duplicates(
            subset=["_norm_key"], keep="first"
        )

        d_for_merge = d.drop(columns=[DISTRICT_NAME_FIELD])

        merged = b.merge(d_for_merge, on="_norm_key", how="left", suffixes=("_boundary", ""))
        merged = merged.drop(columns=["_norm_key"])
        merged = _recover_boundary_fields(b, merged)

        matched = int(b["_norm_key"].isin(d["_norm_key"]).sum())

        if matched > 0:
            return (
                prepare_flood_gdf(merged),
                {
                    "tier": "District name (normalized match)",
                    "matched": matched,
                    "total": total,
                },
            )

    return (
        prepare_flood_gdf(boundary_gdf),
        {"tier": "none", "matched": 0, "total": total},
    )


# ============================================================
# 11. Generic row selection helpers
# ============================================================

def filter_rows_for_area(df, selected_properties, id_field, name_field):
    """
    Generic row filter used for every optional WASH table
    (Water Access, Interventions, Response, Response Plan,
    Budget, Facilities, population, etc.), applying the same
    three-tier matching strategy as the map join.
    """
    if df is None or df.empty or not selected_properties:
        return pd.DataFrame()

    df = normalize_identifier_columns(df.copy())

    # Tier 1: ID field
    selected_id = selected_properties.get(id_field)
    if selected_id is not None and id_field in df.columns:
        result = df[df[id_field].astype(str).str.strip() == str(selected_id).strip()]
        if not result.empty:
            return result

    # Tier 2: exact name
    selected_name = selected_properties.get(name_field)
    if selected_name is not None and name_field in df.columns:
        result = df[
            df[name_field].astype(str).str.strip() == str(selected_name).strip()
        ]
        if not result.empty:
            return result

    # Tier 3: normalized name
    if selected_name is not None and name_field in df.columns:
        target_key = normalize_name_key(selected_name)
        if target_key:
            keys = df[name_field].map(normalize_name_key)
            result = df[keys == target_key]
            if not result.empty:
                return result

    return pd.DataFrame()


def filter_district_rows(df, selected_properties):
    return filter_rows_for_area(
        df, selected_properties, DISTRICT_ID_FIELD, DISTRICT_NAME_FIELD
    )


def filter_municipality_rows(df, selected_properties):
    return filter_rows_for_area(
        df, selected_properties, MUNICIPALITY_ID_FIELD, MUNICIPALITY_NAME_FIELD
    )


# ============================================================
# 12. Map indicator coloring (choropleth)
# ============================================================

def compute_display_layer(gdf, column, kind):
    gdf = gdf.copy()

    if gdf is None or gdf.empty or column not in gdf.columns:
        gdf["display_hex"] = DEFAULT_SEVERITY_HEX
        return gdf, {"kind": "empty"}

    if kind == "categorical":
        gdf["display_hex"] = (
            gdf[column].map(SEVERITY_HEX).fillna(DEFAULT_SEVERITY_HEX)
        )
        return gdf, {"kind": "categorical", "items": SEVERITY_LEGEND}

    values = pd.to_numeric(gdf[column], errors="coerce")
    valid = values.dropna()

    if valid.empty:
        gdf["display_hex"] = DEFAULT_SEVERITY_HEX
        return gdf, {"kind": "empty"}

    vmin, vmax = float(valid.min()), float(valid.max())
    if vmin == vmax:
        vmax = vmin + 1.0

    colormap = bcm.LinearColormap(
        colors=NUMERIC_COLOR_RAMP, vmin=vmin, vmax=vmax
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
    }

    return gdf, legend_info


# ============================================================
# 13. Folium style functions
# ============================================================

def district_style_function(feature):
    props = feature.get("properties", {}) or {}
    color = props.get("display_hex") or props.get("severity_hex") or DEFAULT_SEVERITY_HEX
    affected = props.get("is_flood_affected", False) is True

    return {
        "fillColor": color,
        "color": "#1f1f1f" if affected else "#8a8a8a",
        "weight": 3 if affected else 1.2,
        "fillOpacity": 0.65,
    }


def district_highlight_function(feature):
    return {"weight": 4, "color": "#0078FF", "fillOpacity": 0.8}


def district_selected_style_function(feature):
    return {
        "fillColor": "#0078FF",
        "color": "#00264D",
        "weight": 5,
        "fillOpacity": 0.55,
    }


def municipality_style_function(feature):
    return {
        "fillColor": "#ffffff",
        "color": "#333333",
        "weight": 1,
        "fillOpacity": 0.03,
        "dashArray": "3, 3",
    }


def municipality_highlight_function(feature):
    return {"weight": 3, "color": "#FF7A00", "fillOpacity": 0.15}


def municipality_selected_style_function(feature):
    return {
        "fillColor": "#FF7A00",
        "color": "#8A3D00",
        "weight": 3,
        "fillOpacity": 0.35,
    }


def river_style_function(feature):
    return {"color": RIVER_HEX, "weight": 2, "opacity": 0.8}


# ============================================================
# 14. Optional map overlays: corridor, markers, facilities
# ============================================================

def add_corridor_to_map(m, corridor_df):
    if corridor_df is None or corridor_df.empty:
        return
    if not {"Latitude", "Longitude"}.issubset(corridor_df.columns):
        return

    df = corridor_df.copy()
    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    coords = df[["Latitude", "Longitude"]].dropna().values.tolist()

    if not coords:
        return

    folium.PolyLine(
        locations=coords,
        color=CORRIDOR_HEX,
        weight=5,
        opacity=0.85,
        tooltip="Flood corridor",
    ).add_to(m)


def add_markers_to_map(m, markers_df):
    if markers_df is None or markers_df.empty:
        return
    if not {"Latitude", "Longitude", "Label"}.issubset(markers_df.columns):
        return

    df = markers_df.copy()
    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    df = df.dropna(subset=["Latitude", "Longitude"])

    for _, row in df.iterrows():
        marker_type = str(row.get("Type", "annotation")).lower().strip()
        color = MARKER_HEX.get(marker_type, DEFAULT_MARKER_HEX)
        label = format_text(row.get("Label", ""))
        note = format_text(row.get("Note", label))

        folium.CircleMarker(
            location=[row["Latitude"], row["Longitude"]],
            radius=7 if marker_type == "epicentre" else 5,
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.9,
            tooltip=label,
            popup=note,
        ).add_to(m)


def add_facilities_to_map(m, facilities_df):
    """
    Schools and health facilities, each shown as an
    independently toggleable layer via the map layer control.
    """
    if facilities_df is None or facilities_df.empty:
        return
    if not {"Latitude", "Longitude"}.issubset(facilities_df.columns):
        return

    df = facilities_df.copy()
    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    df = df.dropna(subset=["Latitude", "Longitude"])

    if df.empty:
        return

    type_col = "Facility_Type" if "Facility_Type" in df.columns else None
    groups = {}

    for _, row in df.iterrows():
        type_label = (
            str(row.get(type_col, "Facility")).strip() if type_col else "Facility"
        ) or "Facility"

        color = FACILITY_HEX.get(type_label.lower(), DEFAULT_FACILITY_HEX)

        if type_label not in groups:
            groups[type_label] = folium.FeatureGroup(
                name=f"Facilities: {type_label}", show=True
            )

        name = format_text(row.get("Facility_Name", type_label))
        popup_lines = [name]

        if DISTRICT_NAME_FIELD in df.columns:
            popup_lines.append(
                f"District: {format_text(row.get(DISTRICT_NAME_FIELD))}"
            )
        if "Status" in df.columns:
            popup_lines.append(f"Status: {format_text(row.get('Status'))}")

        folium.CircleMarker(
            location=[row["Latitude"], row["Longitude"]],
            radius=6,
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.9,
            tooltip=f"{type_label}: {name}",
            popup="<br>".join(popup_lines),
        ).add_to(groups[type_label])

    for group in groups.values():
        group.add_to(m)


def sanitize_for_geojson(gdf):
    """
    Defense-in-depth: replace any remaining pandas NA-like values
    (pd.NA, pd.NaT, numpy.nan wrapped in a nullable dtype) in the
    non-geometry columns with plain Python None, immediately
    before the GeoDataFrame is converted to GeoJSON. Plain float
    NaN is left untouched since Python's json module can already
    encode it; this only targets values that would otherwise
    raise TypeError and silently break map rendering.
    """
    if gdf is None or gdf.empty:
        return gdf

    gdf = gdf.copy()
    geometry_col = gdf.geometry.name

    for col in gdf.columns:
        if col == geometry_col:
            continue

        dtype_str = str(gdf[col].dtype)

        if dtype_str in ("string", "Int64", "Float64", "boolean") or dtype_str.startswith(
            ("string", "Int", "Float", "boolean")
        ):
            gdf[col] = gdf[col].astype(object).where(gdf[col].notna(), None)

    return gdf


def add_rivers_to_map(m, rivers_gdf):
    if rivers_gdf is None or rivers_gdf.empty:
        return

    layer = folium.GeoJson(
        json.loads(sanitize_for_geojson(rivers_gdf).to_json()),
        name="Rivers",
        style_function=river_style_function,
    )
    layer.add_to(m)


# ============================================================
# 15. Build the Folium map
# ============================================================

def build_folium_map(
    district_gdf,
    municipality_gdf,
    rivers_gdf,
    corridor_df,
    markers_df,
    facilities_df,
    tiles,
    extra_tooltip,
    selected_district_key,
    selected_municipality_key,
    show_layers,
):
    district_gdf = sanitize_for_geojson(district_gdf)
    municipality_gdf = sanitize_for_geojson(municipality_gdf)

    if isinstance(tiles, tuple):
        tile_url, tile_attr = tiles
        m = folium.Map(
            location=[DEFAULT_LAT, DEFAULT_LON],
            zoom_start=DEFAULT_ZOOM,
            tiles=None,
            control_scale=True,
        )
        folium.TileLayer(
            tiles=tile_url,
            attr=tile_attr,
            name="Base map",
            subdomains=["mt0", "mt1", "mt2", "mt3"],
            control=False,
        ).add_to(m)
    else:
        m = folium.Map(
            location=[DEFAULT_LAT, DEFAULT_LON],
            zoom_start=DEFAULT_ZOOM,
            tiles=tiles,
            control_scale=True,
        )

    # ---- Rivers (drawn first, underneath everything else) ----
    if show_layers.get("rivers", True):
        add_rivers_to_map(m, rivers_gdf)

    # ---- District layer ----
    if show_layers.get("districts", True) and district_gdf is not None and not district_gdf.empty:

        tooltip_fields, tooltip_aliases = [], []

        if DISTRICT_NAME_FIELD in district_gdf.columns:
            tooltip_fields.append(DISTRICT_NAME_FIELD)
            tooltip_aliases.append("District:")

        if DISTRICT_ID_FIELD in district_gdf.columns:
            tooltip_fields.append(DISTRICT_ID_FIELD)
            tooltip_aliases.append("District ID:")

        if "Severity" in district_gdf.columns:
            tooltip_fields.append("Severity")
            tooltip_aliases.append("Severity:")

        if (
            extra_tooltip
            and extra_tooltip[0] in district_gdf.columns
            and extra_tooltip[0] not in tooltip_fields
        ):
            tooltip_fields.append(extra_tooltip[0])
            tooltip_aliases.append(f"{extra_tooltip[1]}:")

        # The selected District is drawn as its own highlighted
        # layer, on top of the base choropleth layer, so the
        # current selection is always visually obvious.
        base_features = []
        selected_features = []

        district_geojson = json.loads(district_gdf.to_json())

        for feature in district_geojson.get("features", []):
            props = feature.get("properties", {}) or {}
            key = props.get(DISTRICT_ID_FIELD) or props.get(DISTRICT_NAME_FIELD)

            if selected_district_key and str(key) == str(selected_district_key):
                selected_features.append(feature)
            else:
                base_features.append(feature)

        base_fc = {"type": "FeatureCollection", "features": base_features}

        tooltip = (
            folium.GeoJsonTooltip(fields=tooltip_fields, aliases=tooltip_aliases, sticky=True)
            if tooltip_fields
            else None
        )

        base_kwargs = dict(
            name="Districts",
            style_function=district_style_function,
            highlight_function=district_highlight_function,
            zoom_on_click=False,
        )

        if tooltip is not None:
            base_kwargs["tooltip"] = tooltip

        folium.GeoJson(base_fc, **base_kwargs).add_to(m)

        if selected_features:
            selected_fc = {
                "type": "FeatureCollection",
                "features": selected_features,
            }

            selected_kwargs = dict(
                name="Selected District",
                style_function=district_selected_style_function,
                zoom_on_click=False,
            )

            if tooltip is not None:
                selected_kwargs["tooltip"] = folium.GeoJsonTooltip(
                    fields=tooltip_fields, aliases=tooltip_aliases, sticky=True
                )

            folium.GeoJson(selected_fc, **selected_kwargs).add_to(m)

    # ---- Municipality layer ----
    if (
        show_layers.get("municipalities", True)
        and municipality_gdf is not None
        and not municipality_gdf.empty
    ):

        muni_fields, muni_aliases = [], []

        for field, alias in [
            (MUNICIPALITY_NAME_FIELD, "Municipality:"),
            (DISTRICT_NAME_FIELD, "District:"),
            (DISTRICT_ID_FIELD, "District ID:"),
            (MUNICIPALITY_ID_FIELD, "Municipality ID:"),
        ]:
            if field in municipality_gdf.columns:
                muni_fields.append(field)
                muni_aliases.append(alias)

        muni_geojson = json.loads(municipality_gdf.to_json())

        base_features, selected_features = [], []

        for feature in muni_geojson.get("features", []):
            props = feature.get("properties", {}) or {}
            key = props.get(MUNICIPALITY_ID_FIELD) or props.get(
                MUNICIPALITY_NAME_FIELD
            )

            if selected_municipality_key and str(key) == str(
                selected_municipality_key
            ):
                selected_features.append(feature)
            else:
                base_features.append(feature)

        base_fc = {"type": "FeatureCollection", "features": base_features}

        muni_kwargs = dict(
            name="Municipalities",
            style_function=municipality_style_function,
            highlight_function=municipality_highlight_function,
            zoom_on_click=False,
        )

        if muni_fields:
            muni_kwargs["tooltip"] = folium.GeoJsonTooltip(
                fields=muni_fields, aliases=muni_aliases, sticky=True
            )

        folium.GeoJson(base_fc, **muni_kwargs).add_to(m)

        if selected_features:
            selected_fc = {
                "type": "FeatureCollection",
                "features": selected_features,
            }

            sel_kwargs = dict(
                name="Selected Municipality",
                style_function=municipality_selected_style_function,
                zoom_on_click=False,
            )

            if muni_fields:
                sel_kwargs["tooltip"] = folium.GeoJsonTooltip(
                    fields=muni_fields, aliases=muni_aliases, sticky=True
                )

            folium.GeoJson(selected_fc, **sel_kwargs).add_to(m)

    if show_layers.get("corridor", True):
        add_corridor_to_map(m, corridor_df)

    if show_layers.get("markers", True):
        add_markers_to_map(m, markers_df)

    if show_layers.get("facilities", True):
        add_facilities_to_map(m, facilities_df)

    folium.LayerControl(collapsed=False).add_to(m)

    try:
        if district_gdf is not None and not district_gdf.empty:
            minx, miny, maxx, maxy = district_gdf.total_bounds
            if all(pd.notna([minx, miny, maxx, maxy])):
                m.fit_bounds([[miny, minx], [maxy, maxx]])
    except Exception:
        pass

    return m


# ============================================================
# 16. Generic table renderer for optional sheets
# ============================================================

def render_optional_table(df, err, selected_properties, filter_fn, empty_message):
    if df is None or df.empty:
        if err:
            st.info(f"{empty_message}\n\nDetails: {err}")
        else:
            st.info(empty_message)
        return

    sub = filter_fn(df, selected_properties)

    if sub.empty:
        st.info(empty_message)
        return

    st.dataframe(sub.reset_index(drop=True), use_container_width=True, hide_index=True)

    numeric_cols = [
        c for c in sub.columns if pd.api.types.is_numeric_dtype(sub[c])
    ]

    if numeric_cols:
        cols = st.columns(min(len(numeric_cols), 4))
        for i, cname in enumerate(numeric_cols):
            cols[i % len(cols)].metric(f"Total: {cname}", format_number(sub[cname].sum()))


# ============================================================
# 17. Target -> Reached -> Gap visualization
# ============================================================

def render_target_reached_gap(label, target, reached, gap=None):
    target_val = pd.to_numeric(pd.Series([target]), errors="coerce").iloc[0]
    reached_val = pd.to_numeric(pd.Series([reached]), errors="coerce").iloc[0]

    if gap is None or pd.isna(gap):
        if pd.notna(target_val) and pd.notna(reached_val):
            gap_val = target_val - reached_val
        else:
            gap_val = None
    else:
        gap_val = pd.to_numeric(pd.Series([gap]), errors="coerce").iloc[0]

    st.markdown(f"**{label}**")

    col1, col2, col3 = st.columns(3)
    col1.metric("Target", format_number(target_val))
    col2.metric("Reached", format_number(reached_val))
    col3.metric(
        "Remaining gap",
        format_number(gap_val) if gap_val is not None else "N/A",
    )

    if pd.notna(target_val) and target_val and pd.notna(reached_val):
        fraction = max(0.0, min(1.0, float(reached_val) / float(target_val)))
        st.progress(fraction)
    else:
        st.caption("Not enough data to draw a progress bar for this indicator.")


# ============================================================
# 18. Overall (all-District) KPI aggregation
# ============================================================

def sum_or_na(series):
    numeric = pd.to_numeric(series, errors="coerce")
    if numeric.dropna().empty:
        return None
    return numeric.sum()


def compute_overall_kpis(merged_gdf, municipality_gdf):
    kpis = {}

    kpis["affected_districts"] = int(
        (merged_gdf["is_flood_affected"] == True).sum()  # noqa: E712
    ) if "is_flood_affected" in merged_gdf.columns else None

    if municipality_gdf is not None and not municipality_gdf.empty:
        if MUNICIPALITY_NAME_FIELD in municipality_gdf.columns:
            kpis["affected_municipalities"] = int(
                municipality_gdf[MUNICIPALITY_NAME_FIELD].dropna().nunique()
            )
        else:
            kpis["affected_municipalities"] = None
    else:
        kpis["affected_municipalities"] = None

    for col, key in [
        ("Affected_Population", "total_affected_population"),
        ("Affected_HH", "total_affected_hh"),
        ("Damaged_WASH_Systems", "total_damaged_systems"),
        ("Systems_Requiring_Repair", "total_repair"),
        ("Systems_Requiring_Rehabilitation", "total_rehab"),
        ("People_Reached", "total_reached"),
        ("Population_Gap", "total_gap"),
    ]:
        kpis[key] = sum_or_na(merged_gdf[col]) if col in merged_gdf.columns else None

    return kpis


# ============================================================
# 19. Main application
# ============================================================

def main():

    # --------------------------------------------------------
    # Session state defaults
    # --------------------------------------------------------

    st.session_state.setdefault("selected_district_key", None)
    st.session_state.setdefault("selected_district_name", None)
    st.session_state.setdefault("selected_municipality_key", None)
    st.session_state.setdefault("selected_municipality_name", None)

    # ==========================================================
    # SIDEBAR: Load Excel Files
    #
    # This is the primary, single-source-of-truth data entry
    # point. Uploading Data_Entry.xlsx here overrides every other
    # data source below (the individual CSV/Excel path fields
    # under "Data Sources" / "Advanced data sources" remain as a
    # fallback for anyone who is not yet using the consolidated
    # workbook, or who has not uploaded one yet).
    # ==========================================================

    st.sidebar.subheader("Load Excel Files")

    uploaded_workbook = st.sidebar.file_uploader(
        f"Upload {DATA_ENTRY_DEFAULT_FILENAME}",
        type=["xlsx"],
        key="data_entry_uploader",
    )

    if uploaded_workbook is not None:
        st.session_state["data_entry_bytes"] = uploaded_workbook.getvalue()
        st.session_state["data_entry_filename"] = uploaded_workbook.name
        st.session_state["data_entry_uploaded_at"] = datetime.now().strftime(
            "%Y-%m-%d %H:%M"
        )

    if st.sidebar.button("Load / Refresh data"):
        st.cache_data.clear()
        st.rerun()

    if st.session_state.get("data_entry_bytes"):
        issues, summary = validate_uploaded_workbook(
            st.session_state["data_entry_bytes"]
        )

        if issues:
            st.sidebar.error("Data loading error")
            for issue in issues:
                st.sidebar.caption(issue)
        else:
            st.sidebar.success("Data loaded successfully")

        st.sidebar.caption(f"File: {st.session_state.get('data_entry_filename')}")
        st.sidebar.caption(
            f"Sheets loaded: {summary['sheets_found']}/{summary['sheets_expected']}"
        )
        st.sidebar.caption(f"Records loaded: {summary['total_records']:,}")
        st.sidebar.caption(f"Districts included: {len(summary['districts'])}")
        st.sidebar.caption(
            f"Municipalities included: {len(summary['municipalities'])}"
        )
        st.sidebar.caption(
            f"Last updated: {st.session_state.get('data_entry_uploaded_at')}"
        )

        if st.sidebar.button("Clear uploaded file"):
            for key in [
                "data_entry_bytes",
                "data_entry_filename",
                "data_entry_uploaded_at",
            ]:
                st.session_state.pop(key, None)
            st.rerun()
    else:
        st.sidebar.info(
            f"Upload the latest {DATA_ENTRY_DEFAULT_FILENAME} to update all "
            "Dashboard figures and visualizations."
        )

    # ==========================================================
    # SIDEBAR: data sources, filters, layers, help
    # ==========================================================

    with st.sidebar.expander("Data Sources", expanded=False):
        excel_path = st.text_input(
            "WASH Excel template path", value=str(DEFAULT_EXCEL_PATH)
        )
        geojson_path = st.text_input(
            "Flood District GeoJSON path", value=str(FLOOD_GEOJSON_PATH)
        )
        municipality_path = st.text_input(
            "Municipality GeoJSON path", value=str(MUNICIPALITY_GEOJSON_PATH)
        )
        rivers_path = st.text_input(
            "Rivers GeoJSON path (optional)", value=str(RIVERS_GEOJSON_PATH)
        )

    with st.sidebar.expander("Advanced data sources", expanded=False):
        district_csv_path = st.text_input(
            "District situation CSV", value=str(DISTRICT_SITUATION_CSV_PATH)
        )
        water_access_csv_path = st.text_input(
            "Water access CSV", value=str(WATER_ACCESS_CSV_PATH)
        )
        interventions_csv_path = st.text_input(
            "Interventions CSV", value=str(INTERVENTIONS_CSV_PATH)
        )
        response_csv_path = st.text_input(
            "Current response CSV", value=str(RESPONSE_CSV_PATH)
        )
        response_plan_csv_path = st.text_input(
            "Response plan CSV", value=str(RESPONSE_PLAN_CSV_PATH)
        )
        budget_csv_path = st.text_input(
            "Budget / funding CSV", value=str(BUDGET_CSV_PATH)
        )
        facilities_csv_path = st.text_input(
            "Facilities CSV", value=str(FACILITIES_CSV_PATH)
        )
        population_path = st.text_input(
            "Population CSV", value=str(POPULATION_PATH)
        )
        municipality_situation_path = st.text_input(
            "Municipality situation CSV",
            value=str(MUNICIPALITY_SITUATION_PATH),
        )
        event_info_path = st.text_input(
            "Event information CSV", value=str(EVENT_INFO_PATH)
        )
        markers_path = st.text_input(
            "Event markers CSV", value=str(EVENT_MARKERS_PATH)
        )
        corridor_path = st.text_input(
            "Flood corridor CSV", value=str(FLOOD_CORRIDOR_PATH)
        )

    st.sidebar.markdown("---")
    st.sidebar.subheader("Map")

    basemap_label = st.sidebar.selectbox(
        "Base map style", list(BASEMAP_OPTIONS.keys()), index=0
    )
    tiles = BASEMAP_OPTIONS[basemap_label]

    show_layers = {
        "districts": st.sidebar.checkbox("Show District layer", value=True),
        "municipalities": st.sidebar.checkbox(
            "Show Municipality layer", value=True
        ),
        "rivers": st.sidebar.checkbox("Show rivers", value=True),
        "facilities": st.sidebar.checkbox(
            "Show schools / health facilities", value=True
        ),
        "corridor": st.sidebar.checkbox("Show flood corridor", value=True),
        "markers": st.sidebar.checkbox("Show event markers", value=True),
    }

    if st.sidebar.button("Refresh data"):
        st.cache_data.clear()
        st.rerun()

    with st.sidebar.expander("How to update the data", expanded=False):
        st.markdown("**Step 1. Update Excel/CSV data**")
        st.markdown(
            "- `district_situation.csv` (or the `District_Situation` sheet "
            "in WASH_data_template.xlsx): one row per District with the "
            "current situation, response, gap, and cost figures.\n"
            "- `water_access.csv`: detailed water access records per "
            "District.\n"
            "- `interventions.csv`: planned or ongoing interventions.\n"
            "- `response.csv`: supplies distributed and people reached "
            "so far (Aqua tabs, hygiene kits, buckets, etc.).\n"
            "- `response_plan.csv`: planned activities, target "
            "population, and cost per activity.\n"
            "- `budget.csv`: funding requirement by response area.\n"
            "- `facilities.csv`: schools and health facilities with "
            "coordinates, to plot them on the map.\n"
            "- `population.csv`: baseline population figures per "
            "District."
        )

        st.markdown("**Step 2. Keep column names consistent**")
        st.dataframe(
            pd.DataFrame(
                {
                    "Column": [
                        "District_ID",
                        "District",
                        "Municipality_ID",
                        "Municipality",
                        "Affected_Population",
                        "Affected_HH",
                        "Affected_Percentage",
                        "Damaged_WASH_Systems",
                        "Systems_Requiring_Repair",
                        "Systems_Requiring_Rehabilitation",
                        "People_Reached",
                        "Population_Gap",
                        "Estimated_Recovery_Requirement_USD",
                        "Estimated_Budget_Requirement_USD",
                    ],
                    "Used for": [
                        "Primary District match key (preferred)",
                        "District name (fallback match key)",
                        "Primary Municipality match key",
                        "Municipality name",
                        "Affected population count",
                        "Affected households count",
                        "Affected population as % of District population",
                        "Number of damaged WASH systems",
                        "Systems needing repair",
                        "Systems needing full rehabilitation",
                        "People currently reached by the response",
                        "Remaining population gap (auto-computed if blank)",
                        "Estimated cost to recover damaged systems",
                        "Estimated total budget requirement",
                    ],
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("**Step 3. Update flood.geojson**")
        st.markdown(
            "flood.geojson currently has no real administrative District "
            "information, so temporary placeholder boundaries (Zone 1, "
            "Zone 2, ...) are shown instead. To connect District_Situation "
            "and other data automatically, add these properties to every "
            "feature in flood.geojson:"
        )
        st.code("District_ID\nDistrict")
        st.markdown(
            "Municipality boundaries are supported as a separate file "
            "(municipality.geojson) with `Municipality` and, ideally, "
            "`District` properties so each Municipality can be linked "
            "back to its parent District."
        )

        st.markdown("**Step 4. Refresh the Dashboard**")
        st.markdown(
            "Push the updated files to the GitHub repository's `data/` "
            "folder. On Streamlit Cloud the app will pick up the new "
            "files on the next run; locally, click **Refresh data** in "
            "the sidebar above."
        )

    # ==========================================================
    # LOAD DATA
    # ==========================================================

    data_status = []

    population_df, population_err = load_table_with_upload_priority(
        SHEET_POPULATION, population_path, excel_path
    )
    data_status.append(("population.csv", population_df, population_err))

    if not Path(geojson_path).exists():
        st.error("Required file `flood.geojson` was not found.")
        st.code(
            "repository/\n"
            "├── app.py\n"
            "└── data/\n"
            "    └── flood.geojson"
        )
        st.info("Upload flood.geojson into the data folder of the repository.")
        st.stop()

    flood_gdf, flood_notes = read_geo_file(geojson_path, "flood.geojson")

    if flood_gdf is None:
        st.error("Failed to load flood.geojson.")
        for note in flood_notes:
            st.write(f"- {note}")
        st.stop()

    municipality_gdf, muni_notes = (None, [])
    if municipality_path and Path(municipality_path).exists():
        municipality_gdf, muni_notes = read_geo_file(
            municipality_path, "municipality.geojson"
        )
        if municipality_gdf is not None and MUNICIPALITY_NAME_FIELD not in municipality_gdf.columns:
            st.sidebar.warning(
                "municipality.geojson has no 'Municipality' column; "
                "Municipality selection is disabled."
            )
            municipality_gdf = None

    rivers_gdf, river_notes = (None, [])
    if rivers_path and Path(rivers_path).exists():
        rivers_gdf, river_notes = read_geo_file(rivers_path, "rivers.geojson")

    # ---- Ensure District field exists ----
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
                flood_gdf["_row_id"] = flood_gdf.index

                points = flood_gdf.copy()
                points["geometry"] = points.geometry.representative_point()

                muni_cols = [
                    c
                    for c in [DISTRICT_NAME_FIELD, DISTRICT_ID_FIELD]
                    if c in municipality_gdf.columns
                ]

                joined = gpd.sjoin(
                    points[["_row_id", "geometry"]],
                    municipality_gdf[muni_cols + ["geometry"]],
                    how="left",
                    predicate="intersects",
                ).drop_duplicates(subset=["_row_id"], keep="first")

                lookup = joined.set_index("_row_id")[muni_cols]
                flood_gdf = flood_gdf.merge(
                    lookup, left_on="_row_id", right_index=True, how="left"
                ).drop(columns=["_row_id"])

                matched = flood_gdf[DISTRICT_NAME_FIELD].notna().sum()

                if matched > 0:
                    joined_from_municipality = True
                    district_field_source = "municipality_spatial_join"
                    flood_gdf[DISTRICT_NAME_FIELD] = flood_gdf[
                        DISTRICT_NAME_FIELD
                    ].fillna("Unknown")

                    st.sidebar.info(
                        "flood.geojson has no District column. District was "
                        f"estimated using a spatial join with "
                        f"municipality.geojson for {matched}/{len(flood_gdf)} "
                        "features."
                    )
            except Exception as e:
                st.sidebar.warning(f"Spatial join with municipality.geojson failed: {e}")

        if not joined_from_municipality:
            flood_gdf = flood_gdf.reset_index(drop=True)
            flood_gdf[DISTRICT_NAME_FIELD] = [
                f"Zone {i + 1}" for i in range(len(flood_gdf))
            ]
            district_field_source = "auto_generated"

    flood_gdf = prepare_flood_gdf(flood_gdf)

    data_status.append(("flood.geojson", flood_gdf, None))
    data_status.append(("municipality.geojson", municipality_gdf, None))
    data_status.append(("rivers.geojson", rivers_gdf, None))

    # ---- Load optional WASH tables ----
    district_df, district_err = load_table_with_upload_priority(
        SHEET_DISTRICT_SITUATION, district_csv_path, excel_path
    )
    water_access_df, water_access_err = load_table_with_upload_priority(
        SHEET_WATER_ACCESS, water_access_csv_path, excel_path
    )
    interventions_df, interventions_err = load_table_with_upload_priority(
        SHEET_INTERVENTIONS, interventions_csv_path, excel_path
    )
    response_df, response_err = load_table_with_upload_priority(
        SHEET_RESPONSE, response_csv_path, excel_path
    )
    response_plan_df, response_plan_err = load_table_with_upload_priority(
        SHEET_RESPONSE_PLAN, response_plan_csv_path, excel_path
    )
    budget_df, budget_err = load_table_with_upload_priority(
        SHEET_BUDGET, budget_csv_path, excel_path
    )
    facilities_df, facilities_err = load_table_with_upload_priority(
        SHEET_FACILITIES, facilities_csv_path, excel_path
    )
    municipality_situation_df, municipality_situation_err = (
        load_table_with_upload_priority(
            SHEET_MUNICIPALITY_DATA, municipality_situation_path, excel_path
        )
    )
    event_info, event_info_err = load_event_info(event_info_path, excel_path)
    markers_df, markers_err = load_table_with_upload_priority(
        SHEET_EVENT_MARKERS, markers_path, excel_path
    )
    corridor_df, corridor_err = load_table_with_upload_priority(
        SHEET_FLOOD_CORRIDOR, corridor_path, excel_path
    )

    for label, df, err in [
        ("District_Situation", district_df, district_err),
        ("Water_Access", water_access_df, water_access_err),
        ("Interventions", interventions_df, interventions_err),
        ("Response", response_df, response_err),
        ("Response_Plan", response_plan_df, response_plan_err),
        ("Budget", budget_df, budget_err),
        ("Facilities", facilities_df, facilities_err),
        ("Municipality_Situation", municipality_situation_df, municipality_situation_err),
    ]:
        data_status.append((label, df, err))

    # ---- Join District Situation data onto the boundaries ----
    merged, join_info = robust_join_district_data(flood_gdf, district_df)

    # ==========================================================
    # TOP OF PAGE: TITLE, PURPOSE, MAP
    # ==========================================================

    st.title("UNICEF Nepal - Flood WASH Dashboard")

    with st.expander("About this dashboard", expanded=False):
        st.markdown(
            "This dashboard supports flood WASH emergency response "
            "decisions in Nepal by bringing together, in one place:"
        )
        st.markdown(
            "- The current flood-affected WASH situation\n"
            "- Affected population and households\n"
            "- Affected Districts and Municipalities\n"
            "- Damaged WASH facilities and water systems\n"
            "- Current UNICEF and partner response\n"
            "- Remaining WASH needs and gaps\n"
            "- The response and estimated funding required to recover "
            "and rebuild"
        )

    # ---- Overall KPI cards ----

    st.subheader("Overall Emergency Situation")

    kpis = compute_overall_kpis(merged, municipality_gdf)

    row1 = st.columns(4)
    row1[0].metric(
        "Total affected population",
        format_number(kpis["total_affected_population"]),
    )
    row1[1].metric(
        "Total affected households", format_number(kpis["total_affected_hh"])
    )
    row1[2].metric(
        "Affected Districts",
        format_number(kpis["affected_districts"]),
    )
    row1[3].metric(
        "Affected Municipalities",
        format_number(kpis["affected_municipalities"]),
    )

    row2 = st.columns(4)
    row2[0].metric(
        "Damaged WASH systems", format_number(kpis["total_damaged_systems"])
    )
    row2[1].metric(
        "Systems requiring repair", format_number(kpis["total_repair"])
    )
    row2[2].metric(
        "Systems requiring rehabilitation", format_number(kpis["total_rehab"])
    )
    row2[3].metric(
        "Population currently reached", format_number(kpis["total_reached"])
    )

    # ---- District / Municipality selection filters ----

    st.markdown("---")
    st.subheader("Interactive Flood & WASH Map")

    district_names = sorted(
        [d for d in merged[DISTRICT_NAME_FIELD].dropna().unique()]
    ) if DISTRICT_NAME_FIELD in merged.columns else []

    filter_col1, filter_col2, filter_col3 = st.columns([2, 2, 2])

    with filter_col1:
        district_filter_choice = st.selectbox(
            "District",
            ["(select on map)"] + district_names,
            index=0,
        )

    municipality_names = []
    if (
        municipality_gdf is not None
        and not municipality_gdf.empty
        and MUNICIPALITY_NAME_FIELD in municipality_gdf.columns
    ):
        muni_subset = municipality_gdf
        if (
            district_filter_choice != "(select on map)"
            and DISTRICT_NAME_FIELD in municipality_gdf.columns
        ):
            muni_subset = municipality_gdf[
                municipality_gdf[DISTRICT_NAME_FIELD] == district_filter_choice
            ]
        municipality_names = sorted(
            muni_subset[MUNICIPALITY_NAME_FIELD].dropna().unique()
        )

    with filter_col2:
        municipality_filter_choice = st.selectbox(
            "Municipality",
            ["(select on map)"] + list(municipality_names),
            index=0,
        )

    with filter_col3:
        indicator_labels = [label for label, _, _ in INDICATOR_OPTIONS]
        selected_label = st.selectbox(
            "Indicator", indicator_labels, index=0
        )

    selected_column, selected_kind = next(
        (col, kind) for label, col, kind in INDICATOR_OPTIONS if label == selected_label
    )

    # Sidebar selections override the map click.
    if district_filter_choice != "(select on map)":
        match = merged[merged[DISTRICT_NAME_FIELD] == district_filter_choice]
        if not match.empty:
            key = match.iloc[0].get(DISTRICT_ID_FIELD) or match.iloc[0].get(
                DISTRICT_NAME_FIELD
            )
            st.session_state["selected_district_key"] = key
            st.session_state["selected_district_name"] = district_filter_choice
            st.session_state["selected_municipality_key"] = None
            st.session_state["selected_municipality_name"] = None

    if municipality_filter_choice != "(select on map)":
        st.session_state["selected_municipality_key"] = municipality_filter_choice
        st.session_state["selected_municipality_name"] = municipality_filter_choice
        st.session_state["selected_district_key"] = None
        st.session_state["selected_district_name"] = None

    merged, legend_info = compute_display_layer(merged, selected_column, selected_kind)

    try:
        fmap = build_folium_map(
            merged,
            municipality_gdf,
            rivers_gdf,
            corridor_df,
            markers_df,
            facilities_df,
            tiles=tiles,
            extra_tooltip=(selected_column, selected_label),
            selected_district_key=st.session_state["selected_district_key"],
            selected_municipality_key=st.session_state["selected_municipality_key"],
            show_layers=show_layers,
        )

        map_render_failed = False
    except Exception as e:
        st.error(
            "The map could not be built. The rest of the dashboard will "
            "still work below."
        )
        st.exception(e)
        fmap = None
        map_render_failed = True

    map_state = None

    if not map_render_failed:
        try:
            # Newer streamlit-folium versions support
            # use_container_width; older ones only accept a
            # fixed width. Try the modern signature first and
            # fall back automatically so a version mismatch
            # never takes down the whole map.
            try:
                map_state = st_folium(
                    fmap,
                    use_container_width=True,
                    height=600,
                    key="flood_map",
                )
            except TypeError:
                map_state = st_folium(
                    fmap,
                    width=1200,
                    height=600,
                    key="flood_map",
                )
        except Exception as e:
            st.error(
                "The map could not be displayed. The rest of the dashboard "
                "will still work below."
            )
            st.exception(e)

    # ---- Legend ----
    if legend_info.get("kind") == "categorical":
        legend_cols = st.columns(len(legend_info["items"]))
        for col, (label, hex_color) in zip(legend_cols, legend_info["items"]):
            col.markdown(
                f"""<div><span style="display:inline-block;width:14px;
                height:14px;background:{hex_color};border-radius:3px;
                margin-right:6px;"></span>{label}</div>""",
                unsafe_allow_html=True,
            )
    elif legend_info.get("kind") == "numeric":
        gradient = ", ".join(legend_info["colors"])
        st.markdown(
            f"""<div style="background:linear-gradient(to right,{gradient});
            height:14px;border-radius:3px;"></div>
            <div style="display:flex;justify-content:space-between;
            font-size:0.85rem;margin-top:2px;">
            <span>{format_number(legend_info['vmin'])}</span>
            <span>{format_number(legend_info['vmax'])}</span></div>""",
            unsafe_allow_html=True,
        )
        st.caption("Darker red means a higher value. Gray means no data yet.")
    else:
        st.info("No data available yet for the selected indicator.")

    # ---- Map click detection ----
    if map_state and map_state.get("last_active_drawing"):
        active = map_state["last_active_drawing"]
        if isinstance(active, dict):
            props = active.get("properties", {}) or {}

            if MUNICIPALITY_NAME_FIELD in props:
                key = props.get(MUNICIPALITY_ID_FIELD) or props.get(
                    MUNICIPALITY_NAME_FIELD
                )
                st.session_state["selected_municipality_key"] = key
                st.session_state["selected_municipality_name"] = props.get(
                    MUNICIPALITY_NAME_FIELD
                )
                st.session_state["selected_district_key"] = None
                st.session_state["selected_district_name"] = None

            elif DISTRICT_NAME_FIELD in props:
                key = props.get(DISTRICT_ID_FIELD) or props.get(
                    DISTRICT_NAME_FIELD
                )
                st.session_state["selected_district_key"] = key
                st.session_state["selected_district_name"] = props.get(
                    DISTRICT_NAME_FIELD
                )
                st.session_state["selected_municipality_key"] = None
                st.session_state["selected_municipality_name"] = None

    selected_district_name = st.session_state["selected_district_name"]
    selected_municipality_name = st.session_state["selected_municipality_name"]

    # Build the "selected_properties" dict used by all filter
    # functions below, from whichever District row matches the
    # current selection (works for both map clicks and sidebar
    # jump-to selections).
    selected_district_row = None
    if selected_district_name and DISTRICT_NAME_FIELD in merged.columns:
        match = merged[merged[DISTRICT_NAME_FIELD] == selected_district_name]
        if not match.empty:
            selected_district_row = match.iloc[0]

    # ==========================================================
    # SELECTED AREA DETAILS
    # ==========================================================

    st.markdown("---")

    if selected_municipality_name:

        st.subheader(f"Selected Municipality: {selected_municipality_name}")

        muni_props = {
            MUNICIPALITY_NAME_FIELD: selected_municipality_name,
            MUNICIPALITY_ID_FIELD: st.session_state["selected_municipality_key"],
        }

        muni_sub = filter_municipality_rows(
            municipality_situation_df, muni_props
        )

        if muni_sub.empty:
            st.info(
                "No detailed situation data found for this Municipality. "
                "Add a matching row to municipality_situation.csv to see "
                "it here."
            )
        else:
            row = muni_sub.iloc[0].to_dict()
            parent_district = format_text(row.get(DISTRICT_NAME_FIELD))
            st.caption(f"Parent District: {parent_district}")

            render_summary_grid(row)

            st.dataframe(
                muni_sub.reset_index(drop=True),
                use_container_width=True,
                hide_index=True,
            )

    elif selected_district_row is not None:

        selected_properties = selected_district_row.to_dict()
        district_name = format_text(
            selected_properties.get(DISTRICT_NAME_FIELD, "District information unavailable")
        )
        severity = format_text(selected_properties.get("Severity", "No Data"))

        st.subheader(f"Selected District: {district_name}")

        severity_color = SEVERITY_HEX.get(severity, DEFAULT_SEVERITY_HEX)
        st.markdown(
            f"""<div style="display:inline-block;background:{severity_color};
            color:white;padding:4px 12px;border-radius:12px;
            font-weight:600;margin-bottom:10px;">
            Severity: {severity}</div>""",
            unsafe_allow_html=True,
        )

        affected_muni_count = "N/A"
        if (
            municipality_gdf is not None
            and DISTRICT_NAME_FIELD in municipality_gdf.columns
            and MUNICIPALITY_NAME_FIELD in municipality_gdf.columns
        ):
            count = municipality_gdf[
                municipality_gdf[DISTRICT_NAME_FIELD] == district_name
            ][MUNICIPALITY_NAME_FIELD].nunique()
            affected_muni_count = count if count else "N/A"

        render_summary_grid(
            selected_properties, extra_field=("Number of affected municipalities", affected_muni_count)
        )

        (
            tab_situation,
            tab_damage,
            tab_response,
            tab_gap,
            tab_plan,
            tab_funding,
            tab_sources,
        ) = st.tabs(
            [
                "WASH Situation",
                "Problem / Needs",
                "Current Response",
                "Remaining Gap",
                "Response Plan",
                "Funding Requirement",
                "Sources",
            ]
        )

        # ---- WASH Situation ----
        with tab_situation:
            st.caption(
                "Number of affected people and where/how they are "
                "currently accessing drinking water."
            )

            c1, c2, c3 = st.columns(3)
            c1.metric(
                "Affected population",
                format_number(selected_properties.get("Affected_Population")),
            )
            c2.metric(
                "Affected households",
                format_number(selected_properties.get("Affected_HH")),
            )
            c3.metric(
                "Affected percentage",
                format_percent(selected_properties.get("Affected_Percentage")),
            )

            st.markdown("**Current water source**")
            st.write(format_text(selected_properties.get("Water_source_current")))

            render_optional_table(
                water_access_df,
                water_access_err,
                selected_properties,
                filter_district_rows,
                "No Water Access data found for this District.",
            )

        # ---- Problem / Needs ----
        with tab_damage:
            st.caption(
                "Key issues related to water availability, quantity, "
                "quality, and safety."
            )

            d1, d2, d3 = st.columns(3)
            d1.metric(
                "Damaged WASH systems",
                format_number(selected_properties.get("Damaged_WASH_Systems")),
            )
            d2.metric(
                "Systems requiring repair",
                format_number(selected_properties.get("Systems_Requiring_Repair")),
            )
            d3.metric(
                "Systems requiring rehabilitation",
                format_number(
                    selected_properties.get("Systems_Requiring_Rehabilitation")
                ),
            )

            st.markdown("**Key problems / needs**")
            st.write(format_text(selected_properties.get("Key_problems")))

            st.markdown("**Severity / priority reason**")
            reason = selected_properties.get("Severity_reason") or selected_properties.get(
                "Priority_reason"
            )
            st.write(format_text(reason))

            st.markdown("**Host community needs**")
            st.write(format_text(selected_properties.get("Host_community_needs")))

        # ---- Current Response ----
        with tab_response:
            st.caption(
                "What UNICEF and partners have already delivered in this "
                "District."
            )

            response_sub = filter_district_rows(response_df, selected_properties)

            if response_sub.empty:
                if response_err:
                    st.info(
                        "No Current Response data found for this District."
                        f"\n\nDetails: {response_err}"
                    )
                else:
                    st.info("No Current Response data found for this District.")
            else:
                row = response_sub.iloc[0].to_dict()

                response_fields = [
                    ("Aqua tabs distributed", "Aqua_Tabs_Distributed"),
                    ("Water treatment units", "Water_Treatment_Units"),
                    ("Hygiene kits", "Hygiene_Kits"),
                    ("Buckets", "Buckets"),
                    ("Mugs", "Mugs"),
                    ("Dignity kits", "Dignity_Kits"),
                    ("Water supply schemes under response", "Water_Supply_Schemes"),
                    ("People reached", "People_Reached"),
                    ("Households reached", "Households_Reached"),
                ]

                cols = st.columns(3)
                for i, (label, col_name) in enumerate(response_fields):
                    cols[i % 3].metric(label, format_number(row.get(col_name)))

                st.markdown("**Current response status**")
                st.write(
                    format_text(
                        row.get("Response_Status")
                        or selected_properties.get("Response_Status")
                    )
                )

                st.dataframe(
                    response_sub.reset_index(drop=True),
                    use_container_width=True,
                    hide_index=True,
                )

            render_optional_table(
                interventions_df,
                interventions_err,
                selected_properties,
                filter_district_rows,
                "No Interventions data found for this District.",
            )

        # ---- Remaining Gap ----
        with tab_gap:
            st.caption(
                "Target versus reached versus remaining gap for this "
                "District."
            )

            render_target_reached_gap(
                "Population coverage",
                selected_properties.get("Affected_Population"),
                selected_properties.get("People_Reached"),
                selected_properties.get("Population_Gap"),
            )

            st.markdown("**WASH supplies distributed**")
            st.write(format_text(selected_properties.get("WASH_Supplies_Distributed")))

        # ---- Response Plan ----
        with tab_plan:
            st.caption(
                "Proposed immediate, medium-term, and long-term "
                "interventions, and the response plan activities."
            )

            st.markdown("**Immediate intervention**")
            st.write(format_text(selected_properties.get("Immediate_intervention")))

            st.markdown("**Medium-term intervention**")
            st.write(format_text(selected_properties.get("Medium_term_intervention")))

            st.markdown("**Long-term intervention**")
            st.write(format_text(selected_properties.get("Long_term_intervention")))

            st.markdown("---")
            st.markdown("**Planned activities**")

            plan_sub = filter_district_rows(response_plan_df, selected_properties)

            if plan_sub.empty:
                if response_plan_err:
                    st.info(
                        "No Response Plan data found for this District."
                        f"\n\nDetails: {response_plan_err}"
                    )
                else:
                    st.info("No Response Plan data found for this District.")
            else:
                display_plan = plan_sub.copy()

                if "Estimated_Unit_Cost" in display_plan.columns:
                    display_plan["Estimated_Unit_Cost"] = display_plan[
                        "Estimated_Unit_Cost"
                    ].apply(format_cost)

                if "Estimated_Total_Cost" in display_plan.columns:
                    display_plan["Estimated_Total_Cost"] = display_plan[
                        "Estimated_Total_Cost"
                    ].apply(format_cost)

                st.dataframe(
                    display_plan.reset_index(drop=True),
                    use_container_width=True,
                    hide_index=True,
                )

            st.info(
                "Schools and health facilities are available as separate, "
                "toggleable layers on the map above (see the layer "
                "control)."
            )

        # ---- Funding Requirement ----
        with tab_funding:
            st.caption(
                "Approximate funding required for the proposed response "
                "in this District."
            )

            c1, c2 = st.columns(2)
            c1.metric(
                "Estimated recovery requirement",
                format_cost(selected_properties.get("Estimated_Recovery_Requirement_USD")),
            )
            c2.metric(
                "Estimated budget requirement",
                format_cost(selected_properties.get("Estimated_Budget_Requirement_USD")),
            )

            st.markdown("---")
            st.markdown("**Budget breakdown**")

            render_budget_table(budget_df, budget_err, selected_properties, filter_district_rows)

        # ---- Sources ----
        with tab_sources:
            st.markdown("**District Situation source**")
            st.caption(f"Data source: {format_text(selected_properties.get('Data_source'))}")
            st.caption(f"Last updated: {format_text(selected_properties.get('Last_updated'))}")
            st.caption(f"Data status: {format_text(selected_properties.get('Data_status'))}")

            project_sources = event_info.get("Sources")
            if project_sources:
                st.markdown("---")
                st.markdown("**Project-level sources**")
                st.write(format_text(project_sources))

            methodology = event_info.get("Methodology")
            if methodology:
                st.markdown("---")
                st.markdown("**Methodology**")
                st.write(format_text(methodology))

            st.markdown("---")
            st.markdown("**How this District was matched to the map**")
            st.caption(
                f"Match tier used across the dataset: {join_info['tier']} "
                f"({join_info['matched']}/{join_info['total']} District "
                "features matched)."
            )

    else:
        st.info(
            "Click a District or Municipality on the map above, or use "
            "the District / Municipality filters above, to see its "
            "detailed situation and response plan here."
        )

    # ==========================================================
    # GLOBAL FUNDING REQUIREMENT (all Districts)
    # ==========================================================

    st.markdown("---")
    st.subheader("Estimated Funding Requirement (All Districts)")

    render_overall_budget_table(budget_df, budget_err, merged)

    # ==========================================================
    # ALL DISTRICTS TABLE
    # ==========================================================

    st.markdown("---")
    st.subheader("All Districts")
    st.caption(f"Sorted by the indicator currently shown on the map: {selected_label}")

    render_all_districts_table(merged, selected_column, selected_kind, selected_label)

    # ==========================================================
    # DATA SOURCES / METHODOLOGY / DATA STATUS
    # ==========================================================

    st.markdown("---")
    st.subheader("Data Sources / Methodology")

    sources_text = event_info.get("Sources")
    methodology_text = event_info.get("Methodology")

    if sources_text:
        st.markdown(f"**Sources:** {format_text(sources_text)}")
    if methodology_text:
        st.markdown(f"**Methodology:** {format_text(methodology_text)}")
    if not sources_text and not methodology_text:
        st.info(
            "No Sources/Methodology information found. Add Key/Value rows "
            "'Sources' and 'Methodology' to event_info.csv to show them "
            "here."
        )

    with st.sidebar.expander("Data status", expanded=False):
        for label, df, err in data_status:
            if df is None:
                st.write(f"{label}: not loaded")
            elif hasattr(df, "empty") and df.empty:
                st.write(f"{label}: not found or empty ({err or 'no data'})")
            else:
                st.write(f"{label}: {len(df):,} record(s) loaded")

    with st.sidebar.expander("Data source / methodology", expanded=False):
        st.write(format_text(sources_text) if sources_text else "N/A")
        st.write(format_text(methodology_text) if methodology_text else "N/A")

    # --------------------------------------------------------
    # Footer
    # --------------------------------------------------------

    st.markdown("---")
    st.caption("UNICEF Nepal | WASH Emergency Response")
    st.caption(
        f"District layer: {Path(geojson_path).name if geojson_path else 'not loaded'} | "
        f"Municipality layer: "
        f"{Path(municipality_path).name if (municipality_gdf is not None and municipality_path) else 'not loaded'}"
    )
    st.caption(f"Base map tiles: {basemap_label} (no API key required)")


# ============================================================
# 20. Shared rendering helpers used inside main()
# ============================================================

SUMMARY_FIELDS = [
    ("Affected population", "Affected_Population", "number"),
    ("Affected households", "Affected_HH", "number"),
    ("Affected percentage", "Affected_Percentage", "percent"),
    ("Damaged WASH systems", "Damaged_WASH_Systems", "number"),
    ("Systems requiring repair", "Systems_Requiring_Repair", "number"),
    (
        "Systems requiring rehabilitation",
        "Systems_Requiring_Rehabilitation",
        "number",
    ),
    ("People currently reached", "People_Reached", "number"),
    ("Remaining population gap", "Population_Gap", "number"),
    ("WASH supplies distributed", "WASH_Supplies_Distributed", "text"),
    ("Current response status", "Response_Status", "text"),
    (
        "Estimated recovery requirement",
        "Estimated_Recovery_Requirement_USD",
        "cost",
    ),
    (
        "Estimated budget requirement",
        "Estimated_Budget_Requirement_USD",
        "cost",
    ),
]


def render_summary_grid(properties, extra_field=None):
    """
    Renders the standard set of area-level fields (used for both
    District and Municipality selections) as a grid of metrics.
    """
    fields = list(SUMMARY_FIELDS)

    cols = st.columns(3)
    slot = 0

    if extra_field is not None:
        label, value = extra_field
        cols[slot % 3].metric(label, format_number(value) if isinstance(value, (int, float)) else value)
        slot += 1

    for label, col_name, kind in fields:
        value = properties.get(col_name)

        if kind == "number":
            display_value = format_number(value)
        elif kind == "percent":
            display_value = format_percent(value)
        elif kind == "cost":
            display_value = format_cost(value)
        else:
            display_value = format_text(value)

        cols[slot % 3].metric(label, display_value)
        slot += 1


def render_budget_table(budget_df, budget_err, selected_properties, filter_fn):
    sub = filter_fn(budget_df, selected_properties)

    if sub.empty:
        if budget_err:
            st.info(f"No Budget data found for this District.\n\nDetails: {budget_err}")
        else:
            st.info("No Budget data found for this District.")
        return

    display = sub.copy()

    for col in ["Unit_Cost", "Estimated_Cost"]:
        if col in display.columns:
            display[col] = display[col].apply(format_cost)

    st.dataframe(display.reset_index(drop=True), use_container_width=True, hide_index=True)

    if "Estimated_Cost" in sub.columns:
        total = sum_or_na(sub["Estimated_Cost"])
        st.metric("Estimated total for this District", format_cost(total))


def render_overall_budget_table(budget_df, budget_err, merged_gdf):
    """
    Response Area | Target | Unit Cost | Estimated Cost | Status
    aggregated across all Districts, plus a prominent grand
    total. Never fabricates numbers - if cost data is missing,
    says so explicitly.
    """
    if budget_df is None or budget_df.empty:
        st.info(
            "Cost data not available. Add a Budget sheet/CSV with columns "
            "Response_Area, Target, Unit_Cost, Estimated_Cost, Status to "
            "populate this section."
        )
        return

    display = budget_df.copy()

    group_col = None
    for candidate in ["Response_Area", "Category", "Activity"]:
        if candidate in display.columns:
            group_col = candidate
            break

    if group_col and "Estimated_Cost" in display.columns:
        summary = (
            display.groupby(group_col, dropna=False)
            .agg(
                Target=("Target", "sum") if "Target" in display.columns else ("Estimated_Cost", "size"),
                Estimated_Cost=("Estimated_Cost", "sum"),
            )
            .reset_index()
        )

        if "Unit_Cost" in display.columns:
            unit_cost_avg = display.groupby(group_col, dropna=False)["Unit_Cost"].mean()
            summary["Unit_Cost"] = summary[group_col].map(unit_cost_avg)

        if "Status" in display.columns:
            status_mode = display.groupby(group_col, dropna=False)["Status"].agg(
                lambda s: s.dropna().iloc[0] if not s.dropna().empty else None
            )
            summary["Status"] = summary[group_col].map(status_mode)

        for col in ["Unit_Cost", "Estimated_Cost"]:
            if col in summary.columns:
                summary[col] = summary[col].apply(format_cost)

        st.dataframe(summary.reset_index(drop=True), use_container_width=True, hide_index=True)

        total = sum_or_na(display["Estimated_Cost"])
        st.metric("Estimated Total Funding Requirement", format_cost(total))
    else:
        st.dataframe(display.reset_index(drop=True), use_container_width=True, hide_index=True)

        if "Estimated_Cost" in display.columns:
            total = sum_or_na(display["Estimated_Cost"])
            st.metric("Estimated Total Funding Requirement", format_cost(total))
        else:
            st.info(
                "Cost data not available. Add an Estimated_Cost column to "
                "the Budget sheet/CSV to see the total here."
            )


def render_all_districts_table(merged, selected_column, selected_kind, selected_label):
    table_columns = [
        c
        for c in [
            DISTRICT_NAME_FIELD,
            DISTRICT_ID_FIELD,
            "Severity",
            "Affected_Population",
            "Affected_HH",
            "People_Reached",
            "Population_Gap",
            "Estimated_Budget_Requirement_USD",
        ]
        if c in merged.columns
    ]

    display_table = pd.DataFrame(merged.drop(columns="geometry"))[table_columns].copy()

    if selected_column in display_table.columns:
        if selected_kind == "numeric":
            display_table["_sort_key"] = pd.to_numeric(
                display_table[selected_column], errors="coerce"
            )
            display_table = display_table.sort_values(
                "_sort_key", ascending=False, na_position="last"
            ).drop(columns="_sort_key")
        else:
            severity_order = {label: i for i, (label, _) in enumerate(SEVERITY_LEGEND)}
            display_table["_sort_key"] = (
                display_table[selected_column].map(severity_order).fillna(len(SEVERITY_LEGEND))
            )
            display_table = display_table.sort_values("_sort_key").drop(columns="_sort_key")

    for col in ["Affected_Population", "Affected_HH", "People_Reached", "Population_Gap"]:
        if col in display_table.columns:
            display_table[col] = display_table[col].apply(format_number)

    if "Estimated_Budget_Requirement_USD" in display_table.columns:
        display_table["Estimated_Budget_Requirement_USD"] = display_table[
            "Estimated_Budget_Requirement_USD"
        ].apply(format_cost)

    st.dataframe(display_table.reset_index(drop=True), use_container_width=True, hide_index=True)


# ============================================================
# 21. Run
# ============================================================

if __name__ == "__main__":
    main()
