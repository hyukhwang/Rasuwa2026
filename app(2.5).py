# -*- coding: utf-8 -*-
"""
UNICEF Nepal
Flood WASH Emergency Decision Support Dashboard

Streamlit + Folium version

Main features
1. District_ID is optional in flood.geojson
2. District_ID is preferred when available
3. District name is used as fallback join key
4. District and Municipality boundaries are displayed
5. Designed for Streamlit Cloud deployment
6. Automatic CRS conversion to EPSG:4326
7. Automatic geometry cleaning
8. District click shows WASH information
9. Municipality click shows Municipality information
10. Optional data files do not stop the application
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
    page_icon="🌊",
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
    ("심각도 (Severity)", "Severity", "categorical"),
    (
        "영향 인구수 (Affected population)",
        "Affected_population",
        "numeric",
    ),
    (
        "식수 접근 인구수 (Water access population)",
        "Water_access_population",
        "numeric",
    ),
    ("사상자 수 (Casualties)", "Casualties", "numeric"),
    (
        "예상 소요 자금 (Estimated funding, USD)",
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
            "파일 경로가 지정되지 않았습니다.",
        )

    if not os.path.exists(path):

        return (
            pd.DataFrame(),
            f"파일을 찾을 수 없습니다: {path}",
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
            f"CSV를 읽는 중 오류: {e}",
        )


@st.cache_data(show_spinner=False)
def _read_csv_raw(path, signature):

    return pd.read_csv(path)


def _safe_read_csv_raw(path):

    if not path:
        return (
            pd.DataFrame(),
            "파일 경로가 지정되지 않았습니다.",
        )

    if not os.path.exists(path):

        return (
            pd.DataFrame(),
            f"파일을 찾을 수 없습니다: {path}",
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
            f"CSV를 읽는 중 오류: {e}",
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
            "Excel 파일 경로가 지정되지 않았습니다.",
        )

    if not os.path.exists(path):

        return (
            pd.DataFrame(),
            f"파일을 찾을 수 없습니다: {path}",
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
            f"'{sheet_name}' 시트를 찾을 수 없습니다: {e}",
        )

    except Exception as e:

        return (
            pd.DataFrame(),
            f"'{sheet_name}' 시트를 읽는 중 오류: {e}",
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
            f"{label}: CRS 정보가 없어 "
            "EPSG:4326으로 가정했습니다."
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
                f"{label}: CRS를 "
                f"{gdf.crs}에서 EPSG:4326으로 "
                "변환했습니다."
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
            f"{label}: 유효한 geometry가 없습니다."
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
            f"{label}: 비어있는 geometry "
            f"{before - after}건을 제거했습니다."
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
                f"{label} 경로가 지정되지 않았습니다."
            ],
        )

    if not os.path.exists(path):

        return (
            None,
            [
                f"{label} 파일을 찾을 수 없습니다: {path}"
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
                f"{label} 파일을 읽는 중 오류: {e}"
            ],
        )

    if gdf is None or gdf.empty:

        return (
            None,
            [
                f"{label}에 유효한 데이터가 없습니다."
            ],
        )

    if "geometry" not in gdf.columns:

        return (
            None,
            [
                f"{label}에 geometry 컬럼이 없습니다."
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
        "District Situation 파일이 없습니다.",
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
# 23. Build Folium map
# ============================================================

def build_folium_map(
    district_gdf,
    municipality_gdf=None,
    corridor_df=None,
    markers_df=None,
    extra_tooltip=None,
):

    m = folium.Map(
        location=[
            DEFAULT_LAT,
            DEFAULT_LON,
        ],
        zoom_start=DEFAULT_ZOOM,
        tiles="CartoDB positron",
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
            f"상세 오류: {err}"
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
                f"합계: {cname}",
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
    # Sidebar
    # --------------------------------------------------------

    st.sidebar.title(
        "⚙️ Data Sources"
    )

    st.sidebar.markdown(
        "### Repository data"
    )

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

    if st.sidebar.button(
        "🔄 Refresh data"
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
                "population.csv를 읽지 못했습니다."
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
            "필수 파일인 "
            "`data/flood.geojson`을 "
            "찾을 수 없습니다."
        )

        st.code(
            "repository/\n"
            "├── app.py\n"
            "└── data/\n"
            "    └── flood.geojson"
        )

        st.info(
            "GitHub repository의 data 폴더에 "
            "flood.geojson을 업로드해주세요."
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
            "flood.geojson을 불러오지 못했습니다."
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
                    "municipality.geojson에 "
                    "`Municipality` 컬럼이 없습니다."
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
                "ℹ️ municipality processing log",
                expanded=False,
            ):

                for note in muni_notes:
                    st.caption(note)

    else:

        st.sidebar.info(
            "municipality.geojson이 없어 "
            "District 레이어만 표시합니다."
        )

    # --------------------------------------------------------
    # Ensure District field exists in flood_gdf
    #
    # flood.geojson에 District 컬럼이 아예 없어도 앱이 멈추지
    # 않도록, 다음 순서로 District 정보를 자동 생성합니다.
    #
    #   1) municipality.geojson에 District 정보가 있으면
    #      공간 조인(spatial join)으로 District를 추정합니다.
    #   2) 그것도 불가능하면 "Zone 1", "Zone 2" ... 형태의
    #      임시 District 이름을 자동 생성합니다.
    #      (이 경우 District_Situation.csv/xlsx 등과는
    #      자동으로 연결되지 않으니, 실제 행정구역 데이터로
    #      교체하는 것을 권장한다는 안내를 표시합니다.)
    # --------------------------------------------------------

    district_field_source = None

    if DISTRICT_NAME_FIELD not in flood_gdf.columns:

        # ------------------------------------------------
        # 1) Try spatial join with municipality boundaries
        # ------------------------------------------------

        joined_from_municipality = False

        if (
            municipality_gdf is not None
            and not municipality_gdf.empty
            and DISTRICT_NAME_FIELD in municipality_gdf.columns
        ):

            try:

                flood_gdf = flood_gdf.reset_index(drop=True)
                flood_gdf["_flood_row_id"] = flood_gdf.index

                # Use representative point so that the join
                # also works for polygons that only partially
                # overlap a municipality boundary.
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
                        "flood.geojson에 `District` 컬럼이 없어 "
                        f"municipality.geojson과의 공간 조인으로 "
                        f"{matched}/{len(flood_gdf)}개 구역의 "
                        "District를 추정했습니다."
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
                            f"District를 찾지 못한 "
                            f"{unmatched}개 구역은 "
                            "'Unknown'으로 표시됩니다."
                        )

            except Exception as e:

                st.sidebar.warning(
                    "municipality.geojson과의 공간 조인 중 "
                    f"오류가 발생했습니다: {e}"
                )

        # ------------------------------------------------
        # 2) Fallback: auto-generate temporary District names
        # ------------------------------------------------

        if not joined_from_municipality:

            flood_gdf = flood_gdf.reset_index(drop=True)

            flood_gdf[DISTRICT_NAME_FIELD] = [
                f"Zone {i + 1}"
                for i in range(len(flood_gdf))
            ]

            district_field_source = "auto_generated"

            st.sidebar.warning(
                "flood.geojson에 `District` 컬럼이 없어 "
                "'Zone 1', 'Zone 2' ... 형태의 임시 District 이름을 "
                "자동으로 생성했습니다. District_Situation 데이터, "
                "Water Access, Funding 등 다른 표와는 자동으로 "
                "연결되지 않으므로, 가능하면 flood.geojson에 실제 "
                "행정구역명을 담은 `District` 컬럼을 추가하는 것을 "
                "권장합니다."
            )

    # --------------------------------------------------------
    # Validate District field (should now always exist)
    # --------------------------------------------------------

    if (
        DISTRICT_NAME_FIELD
        not in flood_gdf.columns
    ):

        # This should not normally happen anymore, but kept as
        # a final safety net in case of unexpected data issues.

        st.error(
            "flood.geojson에 "
            "`District` 컬럼을 생성하지 못했습니다."
        )

        st.write(
            "현재 flood.geojson 컬럼:"
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
            "ℹ️ flood.geojson processing log",
            expanded=False,
        ):

            for note in flood_notes:
                st.caption(note)

    st.sidebar.caption(
        f"District features: "
        f"{len(flood_gdf):,}"
    )

    # --------------------------------------------------------
    # Event information
    # --------------------------------------------------------

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

    st.title(
        page_title
    )

    if page_subtitle:

        st.markdown(
            f"### {page_subtitle}"
        )

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
        "지도에서 District 또는 Municipality를 "
        "클릭하면 상세 정보를 확인할 수 있습니다."
    )

    if district_field_source == "auto_generated":

        st.warning(
            "⚠️ 현재 District 경계는 flood.geojson에 실제 행정구역 "
            "정보가 없어 임시로 생성된 것입니다 (Zone 1, Zone 2 ...). "
            "District_Situation 등 다른 데이터와 연결하려면 "
            "flood.geojson에 실제 District 이름을 추가해 주세요."
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

    if JOIN_KEY:

        merged = join_district_data(
            flood_gdf,
            district_df,
            JOIN_KEY,
        )

        st.sidebar.success(
            f"District Join: {JOIN_KEY}"
        )

        # Join quality
        try:

            geo_values = set(
                flood_gdf[
                    JOIN_KEY
                ]
                .dropna()
                .astype(str)
                .str.strip()
            )

            district_values = set(
                district_df[
                    JOIN_KEY
                ]
                .dropna()
                .astype(str)
                .str.strip()
            )

            missing_in_district = (
                geo_values
                - district_values
            )

            missing_in_geojson = (
                district_values
                - geo_values
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
                            "flood.geojson에는 있지만 "
                            "District 데이터에는 없는 값:"
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
                            "District 데이터에는 있지만 "
                            "flood.geojson에는 없는 값:"
                        )

                        st.code(
                            "\n".join(
                                sorted(
                                    missing_in_geojson
                                )
                            )
                        )

        except Exception:
            pass

    else:

        # ----------------------------------------------------
        # IMPORTANT:
        # District CSV/Excel is optional.
        # The map will still work using flood.geojson.
        # ----------------------------------------------------

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

            st.sidebar.info(
                "District Situation 데이터가 없어 "
                "flood.geojson 속성만 사용합니다."
            )

        else:

            st.sidebar.warning(
                "District 데이터와 "
                "flood.geojson 사이에 "
                "공통 Join key가 없어 "
                "flood.geojson 속성만 사용합니다."
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

    # --------------------------------------------------------
    # Indicator selector
    #
    # 이 지표 선택에 따라 지도의 District 색상과, 지도 아래
    # 전체 수치표가 함께 바뀝니다.
    # --------------------------------------------------------

    st.markdown(
        "#### 🗂️ 지도에 표시할 지표 선택"
    )

    indicator_labels = [
        label
        for label, _, _ in INDICATOR_OPTIONS
    ]

    selected_label = st.selectbox(
        "지표를 선택하면 지도 색상과 아래 표가 함께 바뀝니다.",
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

    # --------------------------------------------------------
    # Map legend (dynamic, based on the selected indicator)
    # --------------------------------------------------------

    with st.expander(
        "ℹ️ Map information and legend",
        expanded=True,
    ):

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
                "색이 진한 빨강일수록 수치가 높은 District입니다. "
                "회색 = 아직 데이터가 입력되지 않은 District."
            )

        else:

            st.info(
                "선택한 지표에 대한 데이터가 아직 없습니다. "
                "Excel(District_Situation 등)에 값을 채우면 "
                "지도와 표에 자동으로 반영됩니다."
            )

        st.caption(
            "굵은 테두리 = 홍수 영향 데이터가 있는 District | "
            "점선 = Municipality 경계"
        )

    # --------------------------------------------------------
    # Map and detail columns
    # --------------------------------------------------------

    map_column, detail_column = (
        st.columns(
            [2, 1]
        )
    )

    # --------------------------------------------------------
    # Build map
    # --------------------------------------------------------

    with map_column:

        try:

            fmap = build_folium_map(
                merged,
                municipality_gdf,
                corridor_df,
                markers_df,
                extra_tooltip=(
                    selected_column,
                    selected_label,
                ),
            )

        except Exception as e:

            st.error(
                "지도를 생성하지 못했습니다."
            )

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

            st.error(
                "지도를 표시하지 못했습니다."
            )

            st.exception(e)

            st.stop()

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

    # --------------------------------------------------------
    # Detail panel
    # --------------------------------------------------------

    with detail_column:

        st.subheader(
            "📋 Detail"
        )

        # ====================================================
        # Municipality
        # ====================================================

        if (
            selection_type
            == "municipality"
        ):

            selected_properties = (
                clicked_props
            )

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

                parent_name = format_text(
                    muni_district_name
                )

            else:

                parent_name = format_text(
                    muni_district_id
                )

            st.markdown(
                f"### 🏘️ {muni_name}"
            )

            st.caption(
                f"Parent District: "
                f"{parent_name}"
            )

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
                    "municipality_situation.csv가 없어 "
                    "기본 속성만 표시합니다."
                )

                display_props = {
                    k: v
                    for k, v
                    in selected_properties.items()
                    if k not in [
                        "severity_hex",
                        "is_flood_affected",
                    ]
                }

                st.json(
                    display_props
                )

            else:

                sub = (
                    filter_municipality_rows(
                        muni_situation_df,
                        selected_properties,
                    )
                )

                if sub.empty:

                    st.info(
                        "이 Municipality에 대한 "
                        "상세 데이터가 없습니다."
                    )

                else:

                    st.dataframe(
                        sub.reset_index(
                            drop=True
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )

        # ====================================================
        # District
        # ====================================================

        elif (
            selection_type
            == "district"
        ):

            selected_properties = (
                clicked_props
            )

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

            st.markdown(
                f"### 🗺️ {district_name}"
            )

            severity_color = (
                SEVERITY_HEX.get(
                    severity,
                    DEFAULT_SEVERITY_HEX,
                )
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
                    {severity}
                </div>
                """,
                unsafe_allow_html=True,
            )

            (
                tab_situation,
                tab_water,
                tab_interventions,
                tab_funding,
                tab_facilities,
                tab_sources,
                tab_population,
            ) = st.tabs(
                [
                    "🗺️ Situation",
                    "💧 Water Access",
                    "🛠️ Interventions",
                    "💰 Funding",
                    "🏫 Facilities",
                    "📚 Sources",
                    "👥 Population",
                ]
            )

            # ----------------------------------------------
            # Situation
            # ----------------------------------------------

            with tab_situation:

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

                st.markdown("---")

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

                st.markdown(
                    "**Severity reason**"
                )

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

                st.write(
                    format_text(
                        severity_reason
                    )
                )

                st.markdown("---")

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

            # ----------------------------------------------
            # Water
            # ----------------------------------------------

            with tab_water:

                render_sheet_table(
                    water_access_df,
                    water_access_err,
                    selected_properties,
                    JOIN_KEY,
                    "이 District에 대한 "
                    "Water Access 데이터가 없습니다.",
                )

            # ----------------------------------------------
            # Interventions
            # ----------------------------------------------

            with tab_interventions:

                render_sheet_table(
                    interventions_df,
                    interventions_err,
                    selected_properties,
                    JOIN_KEY,
                    "이 District에 대한 "
                    "Interventions 데이터가 없습니다.",
                )

            # ----------------------------------------------
            # Funding
            # ----------------------------------------------

            with tab_funding:

                render_sheet_table(
                    funding_df,
                    funding_err,
                    selected_properties,
                    JOIN_KEY,
                    "이 District에 대한 "
                    "Funding 데이터가 없습니다.",
                )

            # ----------------------------------------------
            # Facilities
            # ----------------------------------------------

            with tab_facilities:

                render_sheet_table(
                    facilities_df,
                    facilities_err,
                    selected_properties,
                    JOIN_KEY,
                    "이 District에 대한 "
                    "Facilities 데이터가 없습니다.",
                )

            # ----------------------------------------------
            # Sources
            # ----------------------------------------------

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

                st.markdown(
                    "**District Situation source**"
                )

                st.caption(
                    f"Data source: "
                    f"{data_source}"
                )

                st.caption(
                    f"Last updated: "
                    f"{last_updated}"
                )

                st.caption(
                    f"Data status: "
                    f"{data_status}"
                )

                project_sources = (
                    event_info.get(
                        "Sources"
                    )
                )

                if project_sources:

                    st.markdown("---")

                    st.markdown(
                        "**Project-level sources**"
                    )

                    st.write(
                        format_text(
                            project_sources
                        )
                    )

                other_sources = (
                    collect_sources(
                        selected_properties,
                        JOIN_KEY,
                        water_access_df,
                        interventions_df,
                        funding_df,
                        facilities_df,
                    )
                )

                if other_sources:

                    st.markdown("---")

                    st.markdown(
                        "**Other sheet sources**"
                    )

                    st.dataframe(
                        pd.DataFrame(
                            other_sources
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )

            # ----------------------------------------------
            # Population
            # ----------------------------------------------

            with tab_population:

                st.markdown(
                    "**population.csv 데이터**"
                )

                if population_df.empty:

                    st.info(
                        "population.csv가 없습니다."
                    )

                else:

                    pop_sub = (
                        filter_data_by_selection(
                            population_df,
                            selected_properties,
                            JOIN_KEY,
                        )
                    )

                    if (
                        pop_sub.empty
                        and
                        DISTRICT_NAME_FIELD
                        in population_df.columns
                    ):

                        pop_sub = (
                            population_df[
                                population_df[
                                    DISTRICT_NAME_FIELD
                                ]
                                .astype(str)
                                .str.strip()
                                ==
                                district_name
                            ]
                        )

                    if pop_sub.empty:

                        st.info(
                            "population.csv에서 "
                            "이 District에 해당하는 "
                            "데이터를 찾을 수 없습니다."
                        )

                    else:

                        st.dataframe(
                            pop_sub.reset_index(
                                drop=True
                            ),
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
                                min(
                                    len(
                                        numeric_cols
                                    ),
                                    4,
                                )
                            )

                            for i, cname in enumerate(
                                numeric_cols
                            ):

                                cols[
                                    i % len(cols)
                                ].metric(
                                    cname,
                                    format_number(
                                        pop_sub[
                                            cname
                                        ].sum()
                                    ),
                                )

        # ====================================================
        # Nothing selected
        # ====================================================

        else:

            st.info(
                "지도에서 District 또는 "
                "Municipality를 클릭하면 "
                "상세 정보가 표시됩니다."
            )

    # --------------------------------------------------------
    # Full indicator table (all districts, current indicator)
    #
    # 지도에서 선택한 지표를 기준으로 전체 District를 정렬해서
    # 정확한 수치를 표로 보여줍니다. District를 클릭하지 않아도
    # 여기서 전체 현황을 한 번에 확인할 수 있습니다.
    # --------------------------------------------------------

    st.markdown("---")

    st.markdown(
        "### 📊 전체 District 수치표"
    )

    st.caption(
        f"현재 지도에 표시된 지표: **{selected_label}** "
        "(정렬 기준)"
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
        ]
        if c in merged.columns
    ]

    display_table = pd.DataFrame(
        merged.drop(columns="geometry")
    )[table_columns].copy()

    # Sort by the selected indicator so the most affected /
    # highest-value districts appear first.
    if selected_column in display_table.columns:

        if selected_kind == "numeric":

            display_table["_sort_key"] = pd.to_numeric(
                display_table[selected_column],
                errors="coerce",
            )

            display_table = (
                display_table.sort_values(
                    "_sort_key",
                    ascending=False,
                    na_position="last",
                ).drop(columns="_sort_key")
            )

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

            display_table = (
                display_table.sort_values(
                    "_sort_key"
                ).drop(columns="_sort_key")
            )

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
            display_table[
                "Estimated_funding_USD"
            ].apply(format_currency)
        )

    st.dataframe(
        display_table.reset_index(drop=True),
        use_container_width=True,
        hide_index=True,
    )

    st.caption(
        "표의 값은 Excel(WASH_data_template.xlsx)의 "
        "District_Situation 시트 또는 district_situation.csv를 "
        "채우면 자동으로 업데이트됩니다."
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
        f"District layer: "
        f"{district_layer_name} | "
        f"Municipality layer: "
        f"{municipality_layer_name}"
    )

    st.caption(
        "Map tiles: CartoDB Positron "
        "(no API key required)"
    )


# ============================================================
# 27. Run
# ============================================================

if __name__ == "__main__":
    main()
