# -*- coding: utf-8 -*-
"""
UNICEF Nepal
Flood WASH Emergency Decision Support Dashboard
(단일 파일 통합 버전 — 기존 app.py + config.py + data_loader.py + map_utils.py 를 합침)

필요한 폴더 구조:

rasuwa2026/
│
├── app.py                              (이 파일 하나만 있으면 됩니다)
│
└── data/
    ├── population.csv                  (필수 — 원본 인구 데이터)
    ├── flood.geojson                   (필수 — District 경계)
    ├── WASH_data_template.xlsx         (District_Situation / Water_Access /
    │                                     Interventions / Funding / Facilities 시트 포함)
    ├── district_situation.csv          (선택 — 있으면 Excel의
    │                                     District_Situation 시트보다 우선 사용)
    ├── event_info.csv                  (선택 — Key,Value 2컬럼. 타이틀/서브타이틀/
    │                                     Corridor_note/Data_asof/Sources 등)
    ├── event_markers.csv               (선택 — 진앙지 등 포인트 주석)
    └── flood_corridor_path.csv         (선택 — 하천 범람 경로선)

각 시트/CSV에는 최소한 'District_ID' 컬럼이 있어야 지도(GeoJSON)와 Join됩니다.
"""

import os
import json
import datetime
from pathlib import Path

import pandas as pd
import geopandas as gpd
import pydeck as pdk
import streamlit as st


# ============================================================
# 0. 전역 설정 (구 config.py)
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

POPULATION_PATH = DATA_DIR / "population.csv"
FLOOD_GEOJSON_PATH = DATA_DIR / "flood.geojson"

DEFAULT_EXCEL_PATH = DATA_DIR / "WASH_data_template.xlsx"
DEFAULT_DISTRICT_CSV_PATH = DATA_DIR / "district_situation.csv"
DEFAULT_EVENT_INFO_PATH = DATA_DIR / "event_info.csv"
DEFAULT_EVENT_MARKERS_PATH = DATA_DIR / "event_markers.csv"
DEFAULT_FLOOD_CORRIDOR_PATH = DATA_DIR / "flood_corridor_path.csv"

JOIN_KEY = "District_ID"

DEFAULT_LAT = 28.15
DEFAULT_LON = 84.80
DEFAULT_ZOOM = 6.6

SEVERITY_COLORS = {
    "Severe":    [122, 15, 15, 220],
    "Very High": [214, 39, 40, 200],
    "High":      [244, 145, 45, 190],
    "Moderate":  [255, 221, 87, 180],
    "No Data":   [200, 197, 189, 110],
}
DEFAULT_SEVERITY_COLOR = SEVERITY_COLORS["No Data"]

SEVERITY_LEGEND = [
    ("Severe", "#7A0F0F"),
    ("Very High", "#D62728"),
    ("High", "#F4912D"),
    ("Moderate", "#FFDD57"),
    ("No Data", "#C8C5BD"),
]

FACILITY_COLORS = {
    "School": [31, 119, 180, 220],
    "Health Facility": [44, 160, 44, 220],
}
DEFAULT_FACILITY_COLOR = [128, 128, 128, 220]

CORRIDOR_COLOR = [30, 90, 200, 220]
CORRIDOR_WIDTH_M = 900

MARKER_COLORS = {
    "epicentre": [122, 15, 15, 255],
    "annotation": [40, 40, 40, 255],
}
DEFAULT_MARKER_COLOR = [80, 80, 80, 255]

SHEET_DISTRICT_SITUATION = "District_Situation"
SHEET_WATER_ACCESS = "Water_Access"
SHEET_INTERVENTIONS = "Interventions"
SHEET_FUNDING = "Funding"
SHEET_FACILITIES = "Facilities"


st.set_page_config(
    page_title="UNICEF Nepal - Flood WASH Dashboard",
    page_icon="🌊",
    layout="wide",
)


# ============================================================
# 1. 포맷 헬퍼
# ============================================================

def file_exists(path):
    return Path(path).exists()


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
    return value if value != "" else "No data"


# ============================================================
# 2. 데이터 로딩 (구 data_loader.py)
# ============================================================

def _file_signature(path):
    """캐시 무효화를 위한 파일 signature (경로+수정시각)."""
    try:
        return (str(path), os.path.getmtime(path))
    except OSError:
        return (str(path), None)


@st.cache_data(show_spinner=False)
def _read_sheet(path, sheet_name, signature):
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")


def _normalize_id_columns(df):
    for id_col in ("District_ID", "Facility_ID"):
        if id_col in df.columns:
            df[id_col] = df[id_col].astype(str).str.strip()
    id_col = "District_ID" if "District_ID" in df.columns else (
        "Facility_ID" if "Facility_ID" in df.columns else None
    )
    if id_col:
        df = df[df[id_col].notna() & (df[id_col].astype(str).str.lower() != "nan")]
    return df


def _safe_read_excel_sheet(path, sheet_name):
    if not path or not os.path.exists(path):
        return pd.DataFrame(), f"파일을 찾을 수 없습니다: {path}"
    try:
        df = _read_sheet(path, sheet_name, _file_signature(path))
    except ValueError as e:
        return pd.DataFrame(), f"'{sheet_name}' 시트를 찾을 수 없습니다: {e}"
    except Exception as e:
        return pd.DataFrame(), f"'{sheet_name}' 시트를 읽는 중 오류: {e}"
    return _normalize_id_columns(df), None


@st.cache_data(show_spinner=False)
def _read_csv(path, signature):
    return pd.read_csv(path)


def _safe_read_csv(path):
    if not path or not os.path.exists(path):
        return pd.DataFrame(), f"파일을 찾을 수 없습니다: {path}"
    try:
        df = _read_csv(path, _file_signature(path))
    except Exception as e:
        return pd.DataFrame(), f"CSV를 읽는 중 오류: {e}"
    return _normalize_id_columns(df), None


def _safe_read_csv_no_id_filter(path):
    if not path or not os.path.exists(path):
        return pd.DataFrame(), f"파일을 찾을 수 없습니다: {path}"
    try:
        df = _read_csv(path, _file_signature(path))
    except Exception as e:
        return pd.DataFrame(), f"CSV를 읽는 중 오류: {e}"
    return df, None


def load_district_situation(excel_path, csv_path=None):
    """CSV(district_situation.csv)가 있으면 우선, 없으면 Excel 시트를 읽음."""
    csv_path = csv_path or DEFAULT_DISTRICT_CSV_PATH
    if os.path.exists(csv_path):
        return _safe_read_csv(csv_path)
    return _safe_read_excel_sheet(excel_path, SHEET_DISTRICT_SITUATION)


def load_event_info(path=None):
    path = path or DEFAULT_EVENT_INFO_PATH
    df, err = _safe_read_csv_no_id_filter(path)
    if err or df.empty or "Key" not in df.columns or "Value" not in df.columns:
        return {}, err
    return dict(zip(df["Key"], df["Value"])), None


def load_event_markers(path=None):
    path = path or DEFAULT_EVENT_MARKERS_PATH
    return _safe_read_csv_no_id_filter(path)


def load_flood_corridor(path=None):
    path = path or DEFAULT_FLOOD_CORRIDOR_PATH
    df, err = _safe_read_csv_no_id_filter(path)
    if not err and "Sequence" in df.columns:
        df = df.sort_values("Sequence")
    return df, err


def load_water_access(excel_path):
    return _safe_read_excel_sheet(excel_path, SHEET_WATER_ACCESS)


def load_interventions(excel_path):
    return _safe_read_excel_sheet(excel_path, SHEET_INTERVENTIONS)


def load_funding(excel_path):
    return _safe_read_excel_sheet(excel_path, SHEET_FUNDING)


def load_facilities(excel_path):
    return _safe_read_excel_sheet(excel_path, SHEET_FACILITIES)


@st.cache_data
def load_population_csv(filepath):
    df = pd.read_csv(filepath)
    if JOIN_KEY in df.columns:
        df[JOIN_KEY] = df[JOIN_KEY].astype(str).str.strip()
    return df


@st.cache_data
def load_flood_geojson(filepath):
    return gpd.read_file(filepath, engine="pyogrio")


# ============================================================
# 3. 지도 유틸 (구 map_utils.py)
# ============================================================

def join_district_data(boundary_gdf, district_df):
    """boundary(GeoDataFrame) + District_Situation(DataFrame)을 District_ID로 Join."""
    merged = boundary_gdf.merge(
        district_df,
        on=JOIN_KEY,
        how="left",
        suffixes=("_boundary", ""),
    )

    if "Severity" not in merged.columns:
        merged["Severity"] = merged["Priority"] if "Priority" in merged.columns else "No Data"
    merged["Severity"] = merged["Severity"].fillna("No Data")
    merged.loc[~merged["Severity"].isin(SEVERITY_COLORS.keys()), "Severity"] = "No Data"

    merged["fill_color"] = merged["Severity"].map(SEVERITY_COLORS)

    for numeric_col in ("Affected_population", "Water_access_population", "Estimated_funding_USD"):
        if numeric_col not in merged.columns:
            merged[numeric_col] = None

    return merged


def geodataframe_to_pydeck_features(merged_gdf):
    return json.loads(merged_gdf.to_json())


def build_district_layer(feature_collection):
    return pdk.Layer(
        "GeoJsonLayer",
        feature_collection,
        id="district-layer",
        pickable=True,
        auto_highlight=True,
        stroked=True,
        filled=True,
        get_fill_color="properties.fill_color",
        get_line_color=[70, 70, 70, 220],
        line_width_min_pixels=1,
        highlight_color=[0, 120, 255, 130],
    )


def build_corridor_layer(corridor_df):
    if corridor_df is None or corridor_df.empty:
        return None
    if not {"Latitude", "Longitude"}.issubset(corridor_df.columns):
        return None

    path = corridor_df[["Longitude", "Latitude"]].values.tolist()
    data = [{"path": path, "name": "Flood corridor"}]

    return pdk.Layer(
        "PathLayer",
        data,
        id="corridor-layer",
        get_path="path",
        get_color=CORRIDOR_COLOR,
        get_width=CORRIDOR_WIDTH_M,
        width_min_pixels=2,
        pickable=False,
    )


def build_marker_layers(markers_df):
    if markers_df is None or markers_df.empty:
        return None, None
    required = {"Latitude", "Longitude", "Label"}
    if not required.issubset(markers_df.columns):
        return None, None

    df = markers_df.copy()
    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    df = df.dropna(subset=["Latitude", "Longitude"])
    if df.empty:
        return None, None

    marker_type = df["Type"] if "Type" in df.columns else pd.Series(["annotation"] * len(df))
    df["marker_color"] = marker_type.map(MARKER_COLORS).apply(
        lambda c: c if isinstance(c, list) else DEFAULT_MARKER_COLOR
    )

    point_layer = pdk.Layer(
        "ScatterplotLayer",
        df,
        id="event-marker-points",
        get_position=["Longitude", "Latitude"],
        get_fill_color="marker_color",
        get_radius=600,
        radius_min_pixels=6,
        radius_max_pixels=14,
        pickable=True,
        stroked=True,
        get_line_color=[255, 255, 255, 230],
        line_width_min_pixels=1,
    )

    text_layer = pdk.Layer(
        "TextLayer",
        df,
        id="event-marker-labels",
        get_position=["Longitude", "Latitude"],
        get_text="Label",
        get_size=13,
        get_color=[40, 40, 40, 255],
        get_pixel_offset=[0, -16],
        get_alignment_baseline="'bottom'",
    )

    return point_layer, text_layer


def build_view_state(merged_gdf):
    try:
        if merged_gdf is None or merged_gdf.empty:
            raise ValueError
        minx, miny, maxx, maxy = merged_gdf.total_bounds
        lat = (miny + maxy) / 2
        lon = (minx + maxx) / 2
        return pdk.ViewState(latitude=lat, longitude=lon, zoom=DEFAULT_ZOOM, pitch=0)
    except Exception:
        return pdk.ViewState(
            latitude=DEFAULT_LAT,
            longitude=DEFAULT_LON,
            zoom=DEFAULT_ZOOM,
            pitch=0,
        )


def build_deck(merged_gdf, extra_layers=None):
    feature_collection = geodataframe_to_pydeck_features(merged_gdf)
    layers = [build_district_layer(feature_collection)]
    if extra_layers:
        layers.extend([l for l in extra_layers if l is not None])

    tooltip = {
        "html": (
            "<b>{District}</b><br/>"
            "Severity: {Severity}<br/>"
            "Affected population: {Affected_population}"
        ),
        "style": {"backgroundColor": "#1F4E78", "color": "white"},
    }

    return pdk.Deck(
        layers=layers,
        initial_view_state=build_view_state(merged_gdf),
        tooltip=tooltip,
        map_style="light",
    )


# ============================================================
# 4. 필수 파일 체크
# ============================================================

required_files = {
    "population.csv": POPULATION_PATH,
    "flood.geojson": FLOOD_GEOJSON_PATH,
}

missing_files = [f"data/{name}" for name, path in required_files.items() if not file_exists(path)]

if missing_files:
    st.error("Required data files are missing.")
    st.markdown("The following files could not be found:")
    for filename in missing_files:
        st.code(filename)
    st.markdown("Please check the GitHub repository structure.")
    st.stop()


# ============================================================
# 5. population.csv / flood.geojson 로드
# ============================================================

try:
    population_df = load_population_csv(POPULATION_PATH)
except Exception as e:
    st.error("Failed to load population.csv.")
    st.code(str(e))
    st.write(f"File path: {POPULATION_PATH}")
    st.stop()

try:
    flood_gdf = load_flood_geojson(FLOOD_GEOJSON_PATH)
except Exception as e:
    st.error("Failed to load flood.geojson.")
    st.code(str(e))
    st.markdown(
        """
        Please check the following:

        1. flood.geojson exists in the data folder.
        2. pyogrio is included in requirements.txt.
        3. flood.geojson is a valid GeoJSON file.
        """
    )
    st.write(f"File path: {FLOOD_GEOJSON_PATH}")
    st.stop()


# ============================================================
# 6. 사이드바 — 데이터 소스
# ============================================================

st.sidebar.title("⚙️ Data Sources")

st.sidebar.markdown("### Repository data")
st.sidebar.caption("Population data")
st.sidebar.code("data/population.csv")
st.sidebar.caption("Flood boundary")
st.sidebar.code("data/flood.geojson")

excel_path = st.sidebar.text_input(
    "Excel file path (District_Situation / Water_Access / Interventions / Funding / Facilities)",
    value=str(DEFAULT_EXCEL_PATH),
)

geojson_path = st.sidebar.text_input(
    "Flood GeoJSON path",
    value=str(FLOOD_GEOJSON_PATH),
)

with st.sidebar.expander("Advanced data sources", expanded=False):
    event_info_path = st.text_input("Event information", value=str(DEFAULT_EVENT_INFO_PATH))
    markers_path = st.text_input("Event markers", value=str(DEFAULT_EVENT_MARKERS_PATH))
    corridor_path = st.text_input("Flood corridor", value=str(DEFAULT_FLOOD_CORRIDOR_PATH))
    district_csv_path = st.text_input("District situation CSV (optional override)", value=str(DEFAULT_DISTRICT_CSV_PATH))

if st.sidebar.button("🔄 Refresh data"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption(f"Population records: {len(population_df):,}")
st.sidebar.caption(f"Flood features: {len(flood_gdf):,}")
st.sidebar.caption("GeoJSON engine: pyogrio")


# ============================================================
# 7. 이벤트 메타정보 / 타이틀
# ============================================================

try:
    event_info, event_info_error = load_event_info(event_info_path)
except Exception as e:
    event_info, event_info_error = {}, str(e)

if event_info is None:
    event_info = {}

page_title = event_info.get("Title", "UNICEF Nepal - Flood WASH Dashboard")
page_subtitle = event_info.get("Subtitle", "")

st.title(page_title)
if page_subtitle:
    st.markdown(f"### {page_subtitle}")

metadata = []
corridor_note = event_info.get("Corridor_note", "")
data_asof = event_info.get("Data_asof", "")
if corridor_note:
    metadata.append(str(corridor_note))
if data_asof:
    metadata.append(str(data_asof))
if metadata:
    st.caption(" | ".join(metadata))

st.caption("Select a district on the map to view detailed flood and WASH information.")


# ============================================================
# 8. 범례
# ============================================================

if SEVERITY_LEGEND:
    with st.expander("ℹ️ Map information and Impact Level", expanded=True):
        legend_columns = st.columns(len(SEVERITY_LEGEND))
        for column, item in zip(legend_columns, SEVERITY_LEGEND):
            try:
                label, hex_color = item
            except (ValueError, TypeError):
                continue
            column.markdown(
                f"""
                <div>
                    <span style="display:inline-block;width:14px;height:14px;
                        background:{hex_color};border-radius:3px;margin-right:6px;"></span>
                    {label}
                </div>
                """,
                unsafe_allow_html=True,
            )


# ============================================================
# 9. District_Situation + Excel 시트 전체 로드
# ============================================================

try:
    district_df, district_error = load_district_situation(excel_path, district_csv_path)
except Exception as e:
    district_df, district_error = None, str(e)

if district_error:
    st.error("Failed to load District_Situation data.")
    st.code(str(district_error))
    st.stop()

if district_df is None:
    st.error("District data could not be loaded.")
    st.stop()

# 나머지 4개 시트는 "선택" — 없어도 지도 자체는 동작하고, 해당 탭에서만 안내 표시
water_access_df, water_access_err = load_water_access(excel_path)
interventions_df, interventions_err = load_interventions(excel_path)
funding_df, funding_err = load_funding(excel_path)
facilities_df, facilities_err = load_facilities(excel_path)


# ============================================================
# 10. GeoJSON 준비 + Join Key 정합성 체크
# ============================================================

boundary_gdf = flood_gdf.copy()

if JOIN_KEY not in boundary_gdf.columns:
    st.error(f"GeoJSON does not contain the required join field: {JOIN_KEY}")
    st.markdown("Available GeoJSON fields:")
    st.code("\n".join(str(c) for c in boundary_gdf.columns))
    st.stop()

if JOIN_KEY not in district_df.columns:
    st.error(f"District data does not contain the required join field: {JOIN_KEY}")
    st.markdown("Available District data fields:")
    st.code("\n".join(str(c) for c in district_df.columns))
    st.stop()

boundary_gdf[JOIN_KEY] = boundary_gdf[JOIN_KEY].astype(str).str.strip()
district_df[JOIN_KEY] = district_df[JOIN_KEY].astype(str).str.strip()

try:
    merged = join_district_data(boundary_gdf, district_df)
except Exception as e:
    st.error("Failed to join GeoJSON and District data.")
    st.code(str(e))
    st.stop()

geojson_ids = set(boundary_gdf[JOIN_KEY].dropna().astype(str))
district_ids = set(district_df[JOIN_KEY].dropna().astype(str))
missing_in_district = geojson_ids - district_ids
missing_in_geojson = district_ids - geojson_ids

if missing_in_district or missing_in_geojson:
    with st.expander("🔍 Join quality check", expanded=False):
        if missing_in_district:
            st.warning("IDs found in GeoJSON but not in District data.")
            st.code("\n".join(sorted(missing_in_district)))
        if missing_in_geojson:
            st.warning("IDs found in District data but not in GeoJSON.")
            st.code("\n".join(sorted(missing_in_geojson)))


# ============================================================
# 11. Flood corridor / event markers → 추가 지도 레이어
# ============================================================

try:
    corridor_df, corridor_error = load_flood_corridor(corridor_path)
except Exception as e:
    corridor_df, corridor_error = None, str(e)

try:
    markers_df, markers_error = load_event_markers(markers_path)
except Exception as e:
    markers_df, markers_error = None, str(e)

extra_layers = []

if corridor_df is not None:
    try:
        corridor_layer = build_corridor_layer(corridor_df)
        if corridor_layer is not None:
            extra_layers.append(corridor_layer)
    except Exception:
        pass

if markers_df is not None:
    try:
        point_layer, text_layer = build_marker_layers(markers_df)
        if point_layer is not None:
            extra_layers.append(point_layer)
        if text_layer is not None:
            extra_layers.append(text_layer)
    except Exception:
        pass


# ============================================================
# 12. 지도 + 상세 패널 레이아웃
# ============================================================

map_column, detail_column = st.columns([2, 1])

with map_column:
    try:
        deck = build_deck(merged, extra_layers=extra_layers)
    except Exception as e:
        st.error("Failed to build the map.")
        st.code(str(e))
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
        st.error("Failed to display the map.")
        st.code(str(e))
        st.stop()

selected_properties = None
try:
    selection = map_event.selection
    objects = selection.get("objects", {}).get("district-layer", [])
    if objects:
        selected_properties = objects[0]
except Exception:
    selected_properties = None


# ============================================================
# 13. 상세 패널 헬퍼 — 시트별 표 렌더링
# ============================================================

def get_district_id(props):
    if not props:
        return None
    return props.get(JOIN_KEY) or props.get("District_ID")


def render_sheet_table(df, err, district_id, empty_message):
    """District_ID로 필터링한 표 + 숫자 컬럼 합계를 렌더링."""
    if err:
        st.info(f"이 데이터는 아직 준비되지 않았습니다. ({err})")
        return
    if df is None or df.empty:
        st.info(empty_message)
        return

    if JOIN_KEY in df.columns and district_id is not None:
        sub = df[df[JOIN_KEY].astype(str) == str(district_id)]
    else:
        sub = df

    if sub.empty:
        st.info(empty_message)
        return

    st.dataframe(sub.reset_index(drop=True), use_container_width=True, hide_index=True)

    numeric_cols = [c for c in sub.columns if pd.api.types.is_numeric_dtype(sub[c])]
    if numeric_cols:
        summary_cols = st.columns(min(len(numeric_cols), 4))
        for i, cname in enumerate(numeric_cols):
            summary_cols[i % len(summary_cols)].metric(
                f"합계: {cname}", format_number(sub[cname].sum())
            )


def collect_sources(district_id):
    """Water_Access/Interventions/Funding/Facilities 안에서 'source' 관련 컬럼을 모두 수집."""
    frames = {
        "Water Access": water_access_df,
        "Interventions": interventions_df,
        "Funding": funding_df,
        "Facilities": facilities_df,
    }
    rows = []
    for label, df in frames.items():
        if df is None or df.empty:
            continue
        if JOIN_KEY in df.columns and district_id is not None:
            sub = df[df[JOIN_KEY].astype(str) == str(district_id)]
        else:
            sub = df
        source_cols = [c for c in sub.columns if "source" in c.lower()]
        for c in source_cols:
            for v in sub[c].dropna().unique():
                v = str(v).strip()
                if v and v.lower() != "nan":
                    rows.append({"Sheet": label, "Column": c, "Source": v})
    return rows


# ============================================================
# 14. 상세 패널 — 탭 메뉴
# ============================================================

with detail_column:
    st.subheader("📋 District Detail")

    if not selected_properties:
        st.info("Click a district on the map to view detailed information.")
    else:
        district_id = get_district_id(selected_properties)
        district_name = format_text(selected_properties.get("District", "Unknown District"))
        severity = format_text(selected_properties.get("Severity", "No Data"))

        st.markdown(f"### {district_name}")

        severity_colors = dict(SEVERITY_LEGEND) if SEVERITY_LEGEND else {}
        severity_color = severity_colors.get(severity, "#808080")
        st.markdown(
            f"""
            <div style="display:inline-block;background:{severity_color};color:white;
                padding:4px 12px;border-radius:12px;font-weight:600;margin-bottom:10px;">
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

        # ---------------- District Situation ----------------
        with tab_situation:
            col1, col2 = st.columns(2)
            col1.metric("Affected population", format_number(selected_properties.get("Affected_population")))
            col2.metric("Water access population", format_number(selected_properties.get("Water_access_population")))

            col3, col4 = st.columns(2)
            col3.metric("Casualties", format_number(selected_properties.get("Casualties")))
            col4.metric("Estimated funding", format_currency(selected_properties.get("Estimated_funding_USD")))

            st.markdown("---")
            st.markdown("**Current water source**")
            st.write(format_text(selected_properties.get("Water_source_current")))

            st.markdown("**Key problems / needs**")
            st.write(format_text(selected_properties.get("Key_problems")))

            st.markdown("**Severity reason**")
            severity_reason = selected_properties.get("Severity_reason")
            if severity_reason is None:
                severity_reason = selected_properties.get("Priority_reason")
            st.write(format_text(severity_reason))

            st.markdown("---")
            st.markdown("**Immediate intervention**")
            st.write(format_text(selected_properties.get("Immediate_intervention")))
            st.markdown("**Medium-term intervention**")
            st.write(format_text(selected_properties.get("Medium_term_intervention")))
            st.markdown("**Long-term intervention**")
            st.write(format_text(selected_properties.get("Long_term_intervention")))

        # ---------------- Water Access ----------------
        with tab_water:
            render_sheet_table(
                water_access_df, water_access_err, district_id,
                "이 District에 대한 Water Access 데이터가 없습니다.",
            )

        # ---------------- Interventions ----------------
        with tab_interventions:
            render_sheet_table(
                interventions_df, interventions_err, district_id,
                "이 District에 대한 Interventions 데이터가 없습니다.",
            )

        # ---------------- Funding ----------------
        with tab_funding:
            render_sheet_table(
                funding_df, funding_err, district_id,
                "이 District에 대한 Funding 데이터가 없습니다.",
            )

        # ---------------- Facilities ----------------
        with tab_facilities:
            render_sheet_table(
                facilities_df, facilities_err, district_id,
                "이 District에 대한 Facilities 데이터가 없습니다.",
            )

        # ---------------- Sources ----------------
        with tab_sources:
            data_source = format_text(selected_properties.get("Data_source"))
            last_updated = format_text(selected_properties.get("Last_updated"))
            data_status = format_text(selected_properties.get("Data_status"))

            st.markdown("**District Situation source**")
            st.caption(f"Data source: {data_source}")
            st.caption(f"Last updated: {last_updated}")
            st.caption(f"Data status: {data_status}")

            project_sources = event_info.get("Sources")
            if project_sources:
                st.markdown("---")
                st.markdown("**Project-level sources**")
                st.write(format_text(project_sources))

            other_sources = collect_sources(district_id)
            if other_sources:
                st.markdown("---")
                st.markdown("**Other sheet sources (Water Access / Interventions / Funding / Facilities)**")
                st.dataframe(pd.DataFrame(other_sources), use_container_width=True, hide_index=True)

        # ---------------- Population (raw) ----------------
        with tab_population:
            st.markdown("**원본 population.csv 데이터**")
            if JOIN_KEY in population_df.columns and district_id is not None:
                pop_sub = population_df[population_df[JOIN_KEY].astype(str) == str(district_id)]
            elif "District" in population_df.columns:
                pop_sub = population_df[population_df["District"].astype(str) == district_name]
            else:
                pop_sub = pd.DataFrame()

            if pop_sub.empty:
                st.info("population.csv에서 이 District에 해당하는 데이터를 찾을 수 없습니다.")
            else:
                st.dataframe(pop_sub.reset_index(drop=True), use_container_width=True, hide_index=True)

                numeric_cols = [c for c in pop_sub.columns if pd.api.types.is_numeric_dtype(pop_sub[c])]
                if numeric_cols:
                    cols = st.columns(min(len(numeric_cols), 4))
                    for i, cname in enumerate(numeric_cols):
                        cols[i % len(cols)].metric(cname, format_number(pop_sub[cname].sum()))

                affected = selected_properties.get("Affected_population")
                if affected is not None and "Affected_population" in pop_sub.columns:
                    st.caption(
                        "District_Situation의 Affected_population: "
                        f"{format_number(affected)} vs population.csv 값과 비교해 보세요."
                    )


# ============================================================
# 15. 푸터
# ============================================================

st.markdown("---")
st.caption("UNICEF Nepal | WASH Emergency Response")
st.caption(
    f"Population source: {POPULATION_PATH.name} | "
    f"Flood layer: {FLOOD_GEOJSON_PATH.name} | "
    "GeoJSON engine: pyogrio"
)
