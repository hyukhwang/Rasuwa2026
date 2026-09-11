# -*- coding: utf-8 -*-
"""
UNICEF Nepal
Flood WASH Emergency Decision Support Dashboard

Streamlit + Folium version

주요 특징
1. flood.geojson의 District_ID가 없어도 작동
2. District_ID가 있으면 District_ID를 우선 사용
3. District_ID가 없으면 District 이름으로 자동 Join
4. District 및 Municipality 실제 지도 표시
5. Streamlit Cloud 배포에 적합
6. CRS 자동 변환
7. Geometry 오류 자동 정리
8. District 클릭 시 WASH 상세정보 표시
9. Municipality 클릭 시 Municipality 상세정보 표시

권장 GitHub 구조

rasuwa2026/
│
├── app.py
│
├── requirements.txt
│
└── data/
    ├── population.csv
    ├── flood.geojson
    ├── municipality.geojson
    ├── municipality_situation.csv
    ├── WASH_data_template.xlsx
    ├── district_situation.csv
    ├── event_info.csv
    ├── event_markers.csv
    └── flood_corridor_path.csv

핵심 Join 방식

flood.geojson
    ↓
District_ID가 있으면
    ↓
District_ID 기준 Join

District_ID가 없으면
    ↓
District 기준 Join
"""

import os
import json
from pathlib import Path

import pandas as pd
import geopandas as gpd
import folium
import streamlit as st
from streamlit_folium import st_folium


# ============================================================
# 0. 기본 설정
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

POPULATION_PATH = DATA_DIR / "population.csv"
FLOOD_GEOJSON_PATH = DATA_DIR / "flood.geojson"
MUNICIPALITY_GEOJSON_PATH = DATA_DIR / "municipality.geojson"
MUNICIPALITY_SITUATION_PATH = DATA_DIR / "municipality_situation.csv"

DEFAULT_EXCEL_PATH = DATA_DIR / "WASH_data_template.xlsx"
DEFAULT_DISTRICT_CSV_PATH = DATA_DIR / "district_situation.csv"
DEFAULT_EVENT_INFO_PATH = DATA_DIR / "event_info.csv"
DEFAULT_EVENT_MARKERS_PATH = DATA_DIR / "event_markers.csv"
DEFAULT_FLOOD_CORRIDOR_PATH = DATA_DIR / "flood_corridor_path.csv"


# 가능한 Join 컬럼
DISTRICT_ID_FIELD = "District_ID"
DISTRICT_NAME_FIELD = "District"

MUNICIPALITY_ID_FIELD = "Municipality_ID"
MUNICIPALITY_NAME_FIELD = "Municipality"


# 홍수 영향 District
FLOOD_AFFECTED_DISTRICTS = [
    "Dhading",
    "Nuwakot",
    "Rasuwa",
    "Tanahu",
    "Gorkha",
    "Chitwan",
]


# 기본 지도 위치
DEFAULT_LAT = 28.05
DEFAULT_LON = 84.85
DEFAULT_ZOOM = 8


# Severity 색상
SEVERITY_LEGEND = [
    ("Severe", "#7A0F0F"),
    ("Very High", "#D62728"),
    ("High", "#F4912D"),
    ("Moderate", "#FFDD57"),
    ("No Data", "#C8C5BD"),
]

SEVERITY_HEX = dict(SEVERITY_LEGEND)
DEFAULT_SEVERITY_HEX = SEVERITY_HEX["No Data"]


MARKER_HEX = {
    "epicentre": "#7A0F0F",
    "annotation": "#282828",
}

DEFAULT_MARKER_HEX = "#505050"
CORRIDOR_HEX = "#1E5AC8"


# Excel sheets
SHEET_DISTRICT_SITUATION = "District_Situation"
SHEET_WATER_ACCESS = "Water_Access"
SHEET_INTERVENTIONS = "Interventions"
SHEET_FUNDING = "Funding"
SHEET_FACILITIES = "Facilities"


# Streamlit 설정
st.set_page_config(
    page_title="UNICEF Nepal - Flood WASH Dashboard",
    page_icon="🌊",
    layout="wide",
)


# ============================================================
# 1. 기본 Helper
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

    return value if value else "No data"


# ============================================================
# 2. Column 이름 정리
# ============================================================

def normalize_column_names(df):
    """
    컬럼명의 공백 및 대소문자 문제를 보정합니다.

    예:
    district_id → District_ID
    DISTRICT_ID → District_ID
    district → District
    """

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

        elif lower == "district":
            rename_map[col] = DISTRICT_NAME_FIELD

        elif lower in [
            "municipality_id",
            "municipalityid",
            "municipality id",
        ]:
            rename_map[col] = MUNICIPALITY_ID_FIELD

        elif lower == "municipality":
            rename_map[col] = MUNICIPALITY_NAME_FIELD

        else:
            rename_map[col] = clean

    df = df.rename(columns=rename_map)

    return df


# ============================================================
# 3. ID / Name 정리
# ============================================================

def normalize_identifier_columns(df):

    if df is None or df.empty:
        return df

    df = normalize_column_names(df)

    for col in [
        DISTRICT_ID_FIELD,
        DISTRICT_NAME_FIELD,
        MUNICIPALITY_ID_FIELD,
        MUNICIPALITY_NAME_FIELD,
    ]:

        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.strip()
            )

            df.loc[
                df[col].str.lower().isin(
                    ["nan", "none", "null", ""]
                ),
                col,
            ] = pd.NA

    return df


# ============================================================
# 4. 파일 Signature
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
# 5. CSV / Excel
# ============================================================

@st.cache_data(show_spinner=False)
def _read_csv(path, signature):

    df = pd.read_csv(path)

    return normalize_identifier_columns(df)


def _safe_read_csv(path):

    if not path or not os.path.exists(path):

        return (
            pd.DataFrame(),
            f"파일을 찾을 수 없습니다: {path}",
        )

    try:

        df = _read_csv(
            path,
            _file_signature(path),
        )

        return df, None

    except Exception as e:

        return (
            pd.DataFrame(),
            f"CSV를 읽는 중 오류: {e}",
        )


@st.cache_data(show_spinner=False)
def _read_csv_no_normalization(path, signature):

    return pd.read_csv(path)


def _safe_read_csv_no_id_filter(path):

    if not path or not os.path.exists(path):

        return (
            pd.DataFrame(),
            f"파일을 찾을 수 없습니다: {path}",
        )

    try:

        df = _read_csv_no_normalization(
            path,
            _file_signature(path),
        )

        return normalize_identifier_columns(df), None

    except Exception as e:

        return (
            pd.DataFrame(),
            f"CSV를 읽는 중 오류: {e}",
        )


@st.cache_data(show_spinner=False)
def _read_sheet(path, sheet_name, signature):

    return pd.read_excel(
        path,
        sheet_name=sheet_name,
        engine="openpyxl",
    )


def _safe_read_excel_sheet(path, sheet_name):

    if not path or not os.path.exists(path):

        return (
            pd.DataFrame(),
            f"파일을 찾을 수 없습니다: {path}",
        )

    try:

        df = _read_sheet(
            path,
            sheet_name,
            _file_signature(path),
        )

        df = normalize_identifier_columns(df)

        return df, None

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
# 6. 데이터 로더
# ============================================================

def load_district_situation(
    excel_path,
    csv_path=None,
):

    csv_path = (
        csv_path
        or DEFAULT_DISTRICT_CSV_PATH
    )

    if os.path.exists(csv_path):

        return _safe_read_csv(csv_path)

    return _safe_read_excel_sheet(
        excel_path,
        SHEET_DISTRICT_SITUATION,
    )


def load_event_info(path=None):

    path = (
        path
        or DEFAULT_EVENT_INFO_PATH
    )

    df, err = _safe_read_csv_no_id_filter(path)

    if (
        err
        or df.empty
        or "Key" not in df.columns
        or "Value" not in df.columns
    ):

        return {}, err

    return dict(
        zip(
            df["Key"],
            df["Value"],
        )
    ), None


def load_event_markers(path=None):

    path = (
        path
        or DEFAULT_EVENT_MARKERS_PATH
    )

    return _safe_read_csv_no_id_filter(path)


def load_flood_corridor(path=None):

    path = (
        path
        or DEFAULT_FLOOD_CORRIDOR_PATH
    )

    df, err = _safe_read_csv_no_id_filter(path)

    if (
        not err
        and "Sequence" in df.columns
    ):

        df = df.sort_values("Sequence")

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

    return _safe_read_csv_no_id_filter(path)


# ============================================================
# 7. Population
# ============================================================

@st.cache_data(show_spinner=False)
def load_population_csv(filepath):

    df = pd.read_csv(filepath)

    return normalize_identifier_columns(df)


# ============================================================
# 8. Join Key 자동 결정
# ============================================================

def determine_join_key(
    flood_gdf,
    district_df,
):
    """
    가장 중요한 함수입니다.

    1. flood.geojson과 District 데이터 모두 District_ID가 있으면
       District_ID 사용

    2. District_ID가 없으면 District 이름 사용

    3. 둘 다 불가능하면 None
    """

    flood_has_id = (
        DISTRICT_ID_FIELD in flood_gdf.columns
    )

    district_has_id = (
        DISTRICT_ID_FIELD in district_df.columns
    )

    flood_has_name = (
        DISTRICT_NAME_FIELD in flood_gdf.columns
    )

    district_has_name = (
        DISTRICT_NAME_FIELD in district_df.columns
    )

    if flood_has_id and district_has_id:

        return DISTRICT_ID_FIELD

    if flood_has_name and district_has_name:

        return DISTRICT_NAME_FIELD

    return None


# ============================================================
# 9. Geometry
# ============================================================

@st.cache_data(show_spinner=False)
def _read_geo_file_cached(
    path,
    signature,
):

    try:

        gdf = gpd.read_file(
            path,
            engine="pyogrio",
        )

    except Exception:

        gdf = gpd.read_file(path)

    return gdf


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

    elif gdf.crs.to_epsg() != 4326:

        notes.append(
            f"{label}: CRS를 "
            f"{gdf.crs}에서 EPSG:4326으로 변환했습니다."
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

    before = len(gdf)

    gdf = gdf[
        gdf.geometry.notna()
    ].copy()

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
                ].buffer(0)
            )

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

    if (
        not path
        or not os.path.exists(path)
    ):

        return (
            None,
            [
                f"{label} 파일을 찾을 수 없습니다: {path}"
            ],
        )

    try:

        gdf = _read_geo_file_cached(
            path,
            _file_signature(path),
        )

    except Exception as e:

        return (
            None,
            [
                f"{label} 파일을 읽는 중 오류: {e}"
            ],
        )

    if (
        gdf is None
        or gdf.empty
    ):

        return (
            None,
            [
                f"{label}에 유효한 geometry가 없습니다."
            ],
        )

    gdf = normalize_identifier_columns(
        gdf
    )

    gdf, crs_notes = ensure_wgs84(
        gdf,
        label,
    )

    gdf, geom_notes = clean_geometries(
        gdf,
        label,
    )

    return (
        gdf,
        crs_notes + geom_notes,
    )


# ============================================================
# 10. District Join
# ============================================================

def join_district_data(
    boundary_gdf,
    district_df,
    join_key,
):

    boundary_gdf = boundary_gdf.copy()
    district_df = district_df.copy()

    # Join 컬럼 정리
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

    merged = boundary_gdf.merge(
        district_df,
        on=join_key,
        how="left",
        suffixes=(
            "_boundary",
            "",
        ),
    )

    # District 이름이 없다면 boundary 값 사용
    if (
        DISTRICT_NAME_FIELD
        not in merged.columns
    ):

        if (
            f"{DISTRICT_NAME_FIELD}_boundary"
            in merged.columns
        ):

            merged[
                DISTRICT_NAME_FIELD
            ] = merged[
                f"{DISTRICT_NAME_FIELD}_boundary"
            ]

    # Severity
    if "Severity" not in merged.columns:

        if "Priority" in merged.columns:

            merged["Severity"] = (
                merged["Priority"]
            )

        else:

            merged["Severity"] = "No Data"

    merged["Severity"] = (
        merged["Severity"]
        .fillna("No Data")
        .astype(str)
        .str.strip()
    )

    merged.loc[
        ~merged["Severity"].isin(
            SEVERITY_HEX.keys()
        ),
        "Severity",
    ] = "No Data"

    merged["severity_hex"] = (
        merged["Severity"]
        .map(SEVERITY_HEX)
        .fillna(
            DEFAULT_SEVERITY_HEX
        )
    )

    merged["is_flood_affected"] = (
        merged["Severity"]
        != "No Data"
    )

    # 주요 숫자 컬럼
    for numeric_col in [
        "Affected_population",
        "Water_access_population",
        "Estimated_funding_USD",
    ]:

        if numeric_col not in merged.columns:

            merged[numeric_col] = None

    return merged


# ============================================================
# 11. 선택된 District 필터
# ============================================================

def filter_data_by_selection(
    df,
    selected_properties,
    join_key,
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

    # 우선 현재 Join key 사용
    selected_value = (
        selected_properties.get(
            join_key
        )
    )

    if (
        selected_value is not None
        and join_key in df.columns
    ):

        selected_value = str(
            selected_value
        ).strip()

        return df[
            df[join_key]
            .astype(str)
            .str.strip()
            == selected_value
        ]

    # fallback: District ID
    selected_id = (
        selected_properties.get(
            DISTRICT_ID_FIELD
        )
    )

    if (
        selected_id is not None
        and DISTRICT_ID_FIELD
        in df.columns
    ):

        return df[
            df[DISTRICT_ID_FIELD]
            .astype(str)
            .str.strip()
            == str(selected_id).strip()
        ]

    # fallback: District 이름
    selected_name = (
        selected_properties.get(
            DISTRICT_NAME_FIELD
        )
    )

    if (
        selected_name is not None
        and DISTRICT_NAME_FIELD
        in df.columns
    ):

        return df[
            df[DISTRICT_NAME_FIELD]
            .astype(str)
            .str.strip()
            == str(selected_name).strip()
        ]

    return pd.DataFrame()


# ============================================================
# 12. Folium Style
# ============================================================

def district_style_function(
    feature
):

    props = (
        feature.get(
            "properties",
            {}
        )
        or {}
    )

    color = (
        props.get(
            "severity_hex"
        )
        or DEFAULT_SEVERITY_HEX
    )

    affected = bool(
        props.get(
            "is_flood_affected"
        )
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
    feature
):

    return {
        "weight": 4,
        "color": "#0078FF",
        "fillOpacity": 0.8,
    }


def municipality_style_function(
    feature
):

    return {
        "fillColor": "#ffffff",
        "color": "#333333",
        "weight": 1,
        "fillOpacity": 0.03,
        "dashArray": "3, 3",
    }


def municipality_highlight_function(
    feature
):

    return {
        "weight": 3,
        "color": "#FF7A00",
        "fillOpacity": 0.15,
    }


# ============================================================
# 13. Corridor
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

    coords = (
        corridor_df[
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
# 14. Event Markers
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
        ).lower()

        color = MARKER_HEX.get(
            marker_type,
            DEFAULT_MARKER_HEX,
        )

        folium.CircleMarker(
            location=[
                row["Latitude"],
                row["Longitude"],
            ],
            radius=(
                7
                if marker_type
                == "epicentre"
                else 5
            ),
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.9,
            tooltip=str(
                row.get(
                    "Label",
                    "",
                )
            ),
            popup=str(
                row.get(
                    "Note",
                    row.get(
                        "Label",
                        "",
                    ),
                )
            ),
        ).add_to(m)


# ============================================================
# 15. Map Build
# ============================================================

def build_folium_map(
    district_gdf,
    municipality_gdf,
    corridor_df,
    markers_df,
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
    # District
    # --------------------------------------------------------

    district_geojson = json.loads(
        district_gdf.to_json()
    )

    district_tooltip_fields = []

    if (
        DISTRICT_NAME_FIELD
        in district_gdf.columns
    ):

        district_tooltip_fields.append(
            DISTRICT_NAME_FIELD
        )

    if (
        DISTRICT_ID_FIELD
        in district_gdf.columns
    ):

        district_tooltip_fields.append(
            DISTRICT_ID_FIELD
        )

    if "Severity" in district_gdf.columns:

        district_tooltip_fields.append(
            "Severity"
        )

    if district_tooltip_fields:

        aliases = []

        for field in district_tooltip_fields:

            if field == DISTRICT_NAME_FIELD:
                aliases.append(
                    "District:"
                )

            elif field == DISTRICT_ID_FIELD:
                aliases.append(
                    "District ID:"
                )

            elif field == "Severity":
                aliases.append(
                    "Severity:"
                )

            else:
                aliases.append(
                    f"{field}:"
                )

        district_layer = folium.GeoJson(
            district_geojson,
            name="Districts",
            style_function=district_style_function,
            highlight_function=district_highlight_function,
            tooltip=folium.GeoJsonTooltip(
                fields=district_tooltip_fields,
                aliases=aliases,
                sticky=True,
            ),
        )

    else:

        district_layer = folium.GeoJson(
            district_geojson,
            name="Districts",
            style_function=district_style_function,
            highlight_function=district_highlight_function,
        )

    district_layer.add_to(m)

    # --------------------------------------------------------
    # Municipality
    # --------------------------------------------------------

    if municipality_gdf is not None:

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

        if muni_fields:

            folium.GeoJson(
                muni_geojson,
                name="Municipalities",
                style_function=municipality_style_function,
                highlight_function=municipality_highlight_function,
                tooltip=folium.GeoJsonTooltip(
                    fields=muni_fields,
                    aliases=muni_aliases,
                    sticky=True,
                ),
            ).add_to(m)

        else:

            folium.GeoJson(
                muni_geojson,
                name="Municipalities",
                style_function=municipality_style_function,
                highlight_function=municipality_highlight_function,
            ).add_to(m)

    # --------------------------------------------------------
    # Corridor / Markers
    # --------------------------------------------------------

    add_corridor_to_map(
        m,
        corridor_df,
    )

    add_markers_to_map(
        m,
        markers_df,
    )

    # --------------------------------------------------------
    # Layer Control
    # --------------------------------------------------------

    folium.LayerControl(
        collapsed=False
    ).add_to(m)

    # --------------------------------------------------------
    # Map Bounds
    # --------------------------------------------------------

    try:

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
# 16. 필수 파일 확인
# ============================================================

if not FLOOD_GEOJSON_PATH.exists():

    st.error(
        "필수 파일인 data/flood.geojson을 찾을 수 없습니다."
    )

    st.code(
        "data/flood.geojson"
    )

    st.stop()


# ============================================================
# 17. Population
# ============================================================

population_df = pd.DataFrame()

if POPULATION_PATH.exists():

    try:

        population_df = load_population_csv(
            POPULATION_PATH
        )

    except Exception as e:

        st.warning(
            "population.csv를 읽지 못했습니다."
        )

        st.caption(
            str(e)
        )


# ============================================================
# 18. Sidebar
# ============================================================

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

    municipality_situation_path = st.text_input(
        "Municipality situation CSV",
        value=str(
            MUNICIPALITY_SITUATION_PATH
        ),
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


st.sidebar.markdown("---")

st.sidebar.caption(
    f"Population records: {len(population_df):,}"
)


# ============================================================
# 19. flood.geojson 로드
# ============================================================

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


# ============================================================
# 20. flood.geojson 기본 검증
# ============================================================

if (
    DISTRICT_NAME_FIELD
    not in flood_gdf.columns
):

    st.error(
        "flood.geojson에 'District' 컬럼이 없습니다."
    )

    st.markdown(
        "현재 flood.geojson 컬럼:"
    )

    st.code(
        "\n".join(
            str(c)
            for c in flood_gdf.columns
        )
    )

    st.info(
        "District 이름을 기준으로 지도와 WASH 데이터를 연결하려면 "
        "flood.geojson에 District 컬럼이 필요합니다."
    )

    st.stop()


# District_ID가 없어도 계속 진행
has_flood_id = (
    DISTRICT_ID_FIELD
    in flood_gdf.columns
)


if has_flood_id:

    flood_gdf[
        DISTRICT_ID_FIELD
    ] = (
        flood_gdf[
            DISTRICT_ID_FIELD
        ]
        .astype(str)
        .str.strip()
    )


flood_gdf[
    DISTRICT_NAME_FIELD
] = (
    flood_gdf[
        DISTRICT_NAME_FIELD
    ]
    .astype(str)
    .str.strip()
)


if flood_notes:

    with st.sidebar.expander(
        "ℹ️ flood.geojson processing log",
        expanded=False,
    ):

        for note in flood_notes:

            st.caption(note)


st.sidebar.caption(
    f"District features: {len(flood_gdf):,}"
)


# ============================================================
# 21. Municipality
# ============================================================

municipality_gdf = None

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
                "'Municipality' 컬럼이 없습니다."
            )

            municipality_gdf = None

    if municipality_gdf is not None:

        if (
            DISTRICT_NAME_FIELD
            in municipality_gdf.columns
        ):

            municipality_gdf[
                DISTRICT_NAME_FIELD
            ] = (
                municipality_gdf[
                    DISTRICT_NAME_FIELD
                ]
                .astype(str)
                .str.strip()
            )

        if (
            DISTRICT_ID_FIELD
            in municipality_gdf.columns
        ):

            municipality_gdf[
                DISTRICT_ID_FIELD
            ] = (
                municipality_gdf[
                    DISTRICT_ID_FIELD
                ]
                .astype(str)
                .str.strip()
            )

        if (
            MUNICIPALITY_NAME_FIELD
            in municipality_gdf.columns
        ):

            municipality_gdf[
                MUNICIPALITY_NAME_FIELD
            ] = (
                municipality_gdf[
                    MUNICIPALITY_NAME_FIELD
                ]
                .astype(str)
                .str.strip()
            )

    if muni_notes:

        with st.sidebar.expander(
            "ℹ️ municipality processing log",
            expanded=False,
        ):

            for note in muni_notes:

                st.caption(note)

    if municipality_gdf is not None:

        st.sidebar.caption(
            f"Municipality features: "
            f"{len(municipality_gdf):,}"
        )

else:

    st.sidebar.info(
        "municipality.geojson이 없어 "
        "District 레이어만 표시합니다."
    )


# ============================================================
# 22. Event Info
# ============================================================

try:

    event_info, event_info_error = (
        load_event_info(
            event_info_path
        )
    )

except Exception as e:

    event_info = {}
    event_info_error = str(e)


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
    "지도에서 District 또는 Municipality를 클릭하면 "
    "상세 정보를 확인할 수 있습니다."
)


# ============================================================
# 23. District Situation 로드
# ============================================================

try:

    district_df, district_error = (
        load_district_situation(
            excel_path,
            district_csv_path,
        )
    )

except Exception as e:

    district_df = pd.DataFrame()
    district_error = str(e)


if district_error:

    st.warning(
        "District Situation 데이터를 불러오지 못했습니다."
    )

    st.caption(
        str(district_error)
    )


if (
    district_df is None
    or district_df.empty
):

    st.error(
        "District 데이터가 없습니다."
    )

    st.info(
        "district_situation.csv 또는 "
        "WASH_data_template.xlsx의 "
        "District_Situation 시트를 확인해주세요."
    )

    st.stop()


district_df = normalize_identifier_columns(
    district_df
)


# ============================================================
# 24. Join Key 결정
# ============================================================

JOIN_KEY = determine_join_key(
    flood_gdf,
    district_df,
)


if JOIN_KEY is None:

    st.error(
        "flood.geojson과 District 데이터 사이에 "
        "공통 Join 컬럼이 없습니다."
    )

    st.markdown(
        """
        다음 중 하나가 필요합니다.

        - District_ID가 양쪽 파일에 존재
        - District가 양쪽 파일에 존재
        """
    )

    st.write(
        "flood.geojson columns:",
        list(
            flood_gdf.columns
        ),
    )

    st.write(
        "District data columns:",
        list(
            district_df.columns
        ),
    )

    st.stop()


# ============================================================
# 25. Join Key 표시
# ============================================================

if JOIN_KEY == DISTRICT_ID_FIELD:

    st.sidebar.success(
        "District Join: District_ID"
    )

else:

    st.sidebar.success(
        "District Join: District name"
    )


# ============================================================
# 26. District 데이터 Join
# ============================================================

try:

    merged = join_district_data(
        flood_gdf,
        district_df,
        JOIN_KEY,
    )

except Exception as e:

    st.error(
        "flood.geojson과 District 데이터를 "
        "Join하지 못했습니다."
    )

    st.code(
        str(e)
    )

    st.stop()


# ============================================================
# 27. Join Quality Check
# ============================================================

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


# ============================================================
# 28. 추가 데이터
# ============================================================

try:

    water_access_df, water_access_err = (
        load_water_access(
            excel_path
        )
    )

except Exception as e:

    water_access_df = pd.DataFrame()
    water_access_err = str(e)


try:

    interventions_df, interventions_err = (
        load_interventions(
            excel_path
        )
    )

except Exception as e:

    interventions_df = pd.DataFrame()
    interventions_err = str(e)


try:

    funding_df, funding_err = (
        load_funding(
            excel_path
        )
    )

except Exception as e:

    funding_df = pd.DataFrame()
    funding_err = str(e)


try:

    facilities_df, facilities_err = (
        load_facilities(
            excel_path
        )
    )

except Exception as e:

    facilities_df = pd.DataFrame()
    facilities_err = str(e)


# ============================================================
# 29. Flood Corridor / Markers
# ============================================================

try:

    corridor_df, corridor_error = (
        load_flood_corridor(
            corridor_path
        )
    )

except Exception as e:

    corridor_df = pd.DataFrame()
    corridor_error = str(e)


try:

    markers_df, markers_error = (
        load_event_markers(
            markers_path
        )
    )

except Exception as e:

    markers_df = pd.DataFrame()
    markers_error = str(e)


# ============================================================
# 30. Legend
# ============================================================

with st.expander(
    "ℹ️ Map information and Impact Level",
    expanded=True,
):

    legend_columns = st.columns(
        len(SEVERITY_LEGEND)
    )

    for column, (
        label,
        hex_color,
    ) in zip(
        legend_columns,
        SEVERITY_LEGEND,
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

    st.caption(
        "굵은 테두리 = 홍수 영향 데이터가 있는 District | "
        "점선 = Municipality 경계"
    )


# ============================================================
# 31. Map + Detail
# ============================================================

map_column, detail_column = (
    st.columns(
        [2, 1]
    )
)


with map_column:

    try:

        fmap = build_folium_map(
            merged,
            municipality_gdf,
            corridor_df,
            markers_df,
        )

    except Exception as e:

        st.error(
            "지도를 생성하지 못했습니다."
        )

        st.code(
            str(e)
        )

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

        st.code(
            str(e)
        )

        st.stop()


# ============================================================
# 32. Map Click
# ============================================================

clicked_props = None
selection_type = None


if (
    map_state
    and map_state.get(
        "last_active_drawing"
    )
):

    clicked_props = (
        map_state[
            "last_active_drawing"
        ].get(
            "properties",
            {},
        )
        or {}
    )

    # Municipality
    if (
        MUNICIPALITY_NAME_FIELD
        in clicked_props
        and DISTRICT_NAME_FIELD
        not in clicked_props
    ):

        selection_type = (
            "municipality"
        )

    # District
    elif (
        DISTRICT_NAME_FIELD
        in clicked_props
    ):

        selection_type = (
            "district"
        )


# ============================================================
# 33. Table Helper
# ============================================================

def render_sheet_table(
    df,
    err,
    selected_properties,
    empty_message,
):

    if err:

        st.info(
            f"이 데이터는 아직 준비되지 않았습니다. "
            f"({err})"
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
        JOIN_KEY,
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
# 34. Sources
# ============================================================

def collect_sources(
    selected_properties,
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
            JOIN_KEY,
        )

        source_cols = [
            c
            for c in sub.columns
            if "source" in c.lower()
        ]

        for c in source_cols:

            for v in (
                sub[c]
                .dropna()
                .unique()
            ):

                v = str(v).strip()

                if (
                    v
                    and v.lower()
                    != "nan"
                ):

                    rows.append(
                        {
                            "Sheet": label,
                            "Column": c,
                            "Source": v,
                        }
                    )

    return rows


# ============================================================
# 35. Detail Panel
# ============================================================

with detail_column:

    st.subheader(
        "📋 Detail"
    )

    # ========================================================
    # District
    # ========================================================

    if selection_type == "district":

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

        district_id = (
            selected_properties.get(
                DISTRICT_ID_FIELD
            )
        )

        st.markdown(
            f"### 🗺️ {district_name}"
        )

        severity_color = (
            SEVERITY_HEX.get(
                severity,
                "#808080",
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

        # ----------------------------------------------------
        # Situation
        # ----------------------------------------------------

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
                or
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

        # ----------------------------------------------------
        # Water
        # ----------------------------------------------------

        with tab_water:

            render_sheet_table(
                water_access_df,
                water_access_err,
                selected_properties,
                "이 District에 대한 Water Access 데이터가 없습니다.",
            )

        # ----------------------------------------------------
        # Interventions
        # ----------------------------------------------------

        with tab_interventions:

            render_sheet_table(
                interventions_df,
                interventions_err,
                selected_properties,
                "이 District에 대한 Interventions 데이터가 없습니다.",
            )

        # ----------------------------------------------------
        # Funding
        # ----------------------------------------------------

        with tab_funding:

            render_sheet_table(
                funding_df,
                funding_err,
                selected_properties,
                "이 District에 대한 Funding 데이터가 없습니다.",
            )

        # ----------------------------------------------------
        # Facilities
        # ----------------------------------------------------

        with tab_facilities:

            render_sheet_table(
                facilities_df,
                facilities_err,
                selected_properties,
                "이 District에 대한 Facilities 데이터가 없습니다.",
            )

        # ----------------------------------------------------
        # Sources
        # ----------------------------------------------------

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
                f"Data source: {data_source}"
            )

            st.caption(
                f"Last updated: {last_updated}"
            )

            st.caption(
                f"Data status: {data_status}"
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
                    selected_properties
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

        # ----------------------------------------------------
        # Population
        # ----------------------------------------------------

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

                # Join key로 못 찾으면 District 이름으로 한번 더 확인
                if (
                    pop_sub.empty
                    and DISTRICT_NAME_FIELD
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

    # ========================================================
    # Municipality
    # ========================================================

    elif selection_type == "municipality":

        muni_name = format_text(
            clicked_props.get(
                MUNICIPALITY_NAME_FIELD
            )
        )

        muni_district_id = (
            clicked_props.get(
                DISTRICT_ID_FIELD
            )
        )

        muni_district_name = (
            clicked_props.get(
                DISTRICT_NAME_FIELD
            )
        )

        parent_name = (
            format_text(
                muni_district_name
            )
            if muni_district_name
            else format_text(
                muni_district_id
            )
        )

        st.markdown(
            f"### 🏘️ {muni_name}"
        )

        st.caption(
            f"Parent District: {parent_name}"
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
                in clicked_props.items()
                if k not in (
                    "severity_hex",
                    "is_flood_affected",
                )
            }

            st.json(
                display_props
            )

        else:

            sub = filter_municipality_rows(
                muni_situation_df,
                clicked_props,
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

    # ========================================================
    # Nothing selected
    # ========================================================

    else:

        st.info(
            "지도에서 District 또는 Municipality를 "
            "클릭하면 상세 정보가 표시됩니다."
        )


# ============================================================
# 36. Footer
# ============================================================

st.markdown("---")

st.caption(
    "UNICEF Nepal | WASH Emergency Response"
)

st.caption(
    f"District layer: "
    f"{Path(geojson_path).name} | "
    f"Municipality layer: "
    f"{Path(municipality_path).name "
    if municipality_gdf is not None
    else 'not loaded'}"
)

st.caption(
    "Map tiles: CartoDB Positron "
    "(no API key required)"
)