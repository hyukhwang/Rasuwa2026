# -*- coding: utf-8 -*-
"""
UNICEF Nepal
Flood WASH Emergency Decision Support Dashboard
(단일 파일 통합 버전 — folium 기반 실제 지도 + District/Municipality 레이어)

지도 엔진을 pydeck → folium(streamlit-folium) 으로 교체했습니다.
pydeck의 "light" 맵 스타일은 Mapbox/Carto API 토큰이 필요한데, Streamlit Cloud에는
기본적으로 그 토큰이 없어 타일이 로드되지 않고 빈 화면만 나옵니다.
folium + CartoDB Positron 타일은 API 키 없이 항상 렌더링됩니다.

필요한 폴더 구조 (GitHub repository 기준):

rasuwa2026/
│
├── app.py                              (이 파일 하나만 있으면 됩니다)
│
└── data/
    ├── population.csv                  [필수] 원본 인구 데이터 (District_ID 컬럼 포함)
    ├── flood.geojson                   [필수] District(구/군) 경계. District_ID, District 컬럼 필요
    ├── municipality.geojson            [선택, 강력 권장] Municipality(시/면) 경계.
    │                                     District_ID(부모 District 연결용), Municipality(이름) 컬럼 필요
    ├── municipality_situation.csv      [선택] Municipality 클릭 시 보여줄 상세 정보.
    │                                     Municipality_ID 또는 (District_ID + Municipality) 컬럼으로 매칭
    ├── WASH_data_template.xlsx         [필수] District_Situation 시트 포함
    │                                     (Water_Access / Interventions / Funding / Facilities 시트는 선택)
    ├── district_situation.csv          [선택] 있으면 Excel의 District_Situation 시트보다 우선
    ├── event_info.csv                  [선택] Key,Value 2컬럼 — 타이틀/서브타이틀/Sources 등
    ├── event_markers.csv               [선택] 진앙지 등 포인트 주석 (Latitude, Longitude, Label)
    └── flood_corridor_path.csv         [선택] 하천 범람 경로선 (Sequence, Latitude, Longitude)

flood.geojson 과 municipality.geojson의 좌표계(CRS)가 서로 다르거나 없는 경우,
이 앱이 자동으로 감지해서 EPSG:4326(WGS84)으로 변환하므로 두 레이어가 항상
같은 지도 위에 정확히 겹쳐서 표시됩니다.
"""

import os
import json
from pathlib import Path

import pandas as pd
import geopandas as gpd
import folium
from streamlit_folium import st_folium
import streamlit as st


# ============================================================
# 0. 전역 설정
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

JOIN_KEY = "District_ID"
MUNICIPALITY_NAME_FIELD = "Municipality"

# 홍수 영향 District (참고용 — 지도 초기 범위를 잡을 때 사용)
FLOOD_AFFECTED_DISTRICTS = ["Dhading", "Nuwakot", "Rasuwa", "Tanahu", "Gorkha", "Chitwan"]

DEFAULT_LAT = 28.05
DEFAULT_LON = 84.85
DEFAULT_ZOOM = 8

SEVERITY_LEGEND = [
    ("Severe", "#7A0F0F"),
    ("Very High", "#D62728"),
    ("High", "#F4912D"),
    ("Moderate", "#FFDD57"),
    ("No Data", "#C8C5BD"),
]
SEVERITY_HEX = dict(SEVERITY_LEGEND)
DEFAULT_SEVERITY_HEX = SEVERITY_HEX["No Data"]

MARKER_HEX = {"epicentre": "#7A0F0F", "annotation": "#282828"}
DEFAULT_MARKER_HEX = "#505050"
CORRIDOR_HEX = "#1E5AC8"

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
# 2. 표 데이터 로딩 (Excel / CSV)
# ============================================================

def _file_signature(path):
    try:
        return (str(path), os.path.getmtime(path))
    except OSError:
        return (str(path), None)


@st.cache_data(show_spinner=False)
def _read_sheet(path, sheet_name, signature):
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")


def _normalize_id_columns(df):
    for id_col in ("District_ID", "Facility_ID", "Municipality_ID"):
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


def load_municipality_situation(path):
    return _safe_read_csv_no_id_filter(path)


@st.cache_data
def load_population_csv(filepath):
    df = pd.read_csv(filepath)
    if JOIN_KEY in df.columns:
        df[JOIN_KEY] = df[JOIN_KEY].astype(str).str.strip()
    return df


def filter_municipality_rows(df, props):
    if df is None or df.empty:
        return pd.DataFrame()
    if "Municipality_ID" in df.columns and "Municipality_ID" in props:
        return df[df["Municipality_ID"].astype(str) == str(props["Municipality_ID"])]
    if MUNICIPALITY_NAME_FIELD in df.columns:
        sub = df[df[MUNICIPALITY_NAME_FIELD].astype(str) == str(props.get(MUNICIPALITY_NAME_FIELD))]
        if JOIN_KEY in df.columns and JOIN_KEY in props:
            sub = sub[sub[JOIN_KEY].astype(str) == str(props.get(JOIN_KEY))]
        return sub
    return pd.DataFrame()


# ============================================================
# 3. GeoJSON 로딩 — CRS 자동 정렬 + geometry 검증 (요구사항 6, 9)
# ============================================================

@st.cache_data(show_spinner=False)
def _read_geo_file_cached(path, signature):
    try:
        gdf = gpd.read_file(path, engine="pyogrio")
    except Exception:
        gdf = gpd.read_file(path)
    return gdf


def ensure_wgs84(gdf, label):
    notes = []
    if gdf.crs is None:
        notes.append(f"{label}: 좌표계(CRS) 정보가 없어 EPSG:4326(WGS84)으로 가정했습니다.")
        gdf = gdf.set_crs(epsg=4326, allow_override=True)
    elif gdf.crs.to_epsg() != 4326:
        notes.append(f"{label}: 좌표계를 {gdf.crs}에서 EPSG:4326(WGS84)으로 자동 변환했습니다.")
        gdf = gdf.to_crs(epsg=4326)
    return gdf, notes


def clean_geometries(gdf, label):
    notes = []
    before = len(gdf)
    gdf = gdf[gdf.geometry.notna()].copy()
    invalid_mask = ~gdf.geometry.is_valid
    if invalid_mask.any():
        gdf.loc[invalid_mask, "geometry"] = gdf.loc[invalid_mask, "geometry"].buffer(0)
    after = len(gdf)
    if after < before:
        notes.append(f"{label}: 비어있거나 유효하지 않은 geometry {before - after}건을 정리했습니다.")
    return gdf, notes


def read_geo_file(path, label):
    """GeoJSON/Shapefile 로드 + CRS 자동 변환 + geometry 정리. (gdf 또는 None, notes) 반환."""
    if not path or not os.path.exists(path):
        return None, [f"{label} 파일을 찾을 수 없습니다: {path}"]
    try:
        gdf = _read_geo_file_cached(path, _file_signature(path))
    except Exception as e:
        return None, [f"{label} 파일을 읽는 중 오류가 발생했습니다: {e}"]

    if gdf is None or gdf.empty:
        return None, [f"{label} 파일에 유효한 geometry가 없습니다: {path}"]

    gdf, crs_notes = ensure_wgs84(gdf, label)
    gdf, geom_notes = clean_geometries(gdf, label)
    return gdf, crs_notes + geom_notes


# ============================================================
# 4. District/Municipality Join
# ============================================================

def join_district_data(boundary_gdf, district_df):
    merged = boundary_gdf.merge(
        district_df,
        on=JOIN_KEY,
        how="left",
        suffixes=("_boundary", ""),
    )

    if "Severity" not in merged.columns:
        merged["Severity"] = merged["Priority"] if "Priority" in merged.columns else "No Data"
    merged["Severity"] = merged["Severity"].fillna("No Data")
    merged.loc[~merged["Severity"].isin(SEVERITY_HEX.keys()), "Severity"] = "No Data"

    merged["severity_hex"] = merged["Severity"].map(SEVERITY_HEX).fillna(DEFAULT_SEVERITY_HEX)
    merged["is_flood_affected"] = merged["Severity"] != "No Data"

    for numeric_col in ("Affected_population", "Water_access_population", "Estimated_funding_USD"):
        if numeric_col not in merged.columns:
            merged[numeric_col] = None

    return merged


# ============================================================
# 5. folium 지도 빌드 (요구사항 1, 2, 3, 4, 5)
# ============================================================

def district_style_function(feature):
    props = feature.get("properties", {}) or {}
    color = props.get("severity_hex") or DEFAULT_SEVERITY_HEX
    affected = bool(props.get("is_flood_affected"))
    return {
        "fillColor": color,
        "color": "#1f1f1f" if affected else "#8a8a8a",
        "weight": 3 if affected else 1.2,
        "fillOpacity": 0.65,
    }


def district_highlight_function(feature):
    return {"weight": 4, "color": "#0078FF", "fillOpacity": 0.8}


def municipality_style_function(feature):
    return {
        "fillColor": "#ffffff",
        "color": "#333333",
        "weight": 1,
        "fillOpacity": 0.03,   # 0이 아니어야 내부 클릭도 감지됩니다
        "dashArray": "3, 3",
    }


def municipality_highlight_function(feature):
    return {"weight": 3, "color": "#FF7A00", "fillOpacity": 0.15}


def add_corridor_to_map(m, corridor_df):
    if corridor_df is None or corridor_df.empty:
        return
    if not {"Latitude", "Longitude"}.issubset(corridor_df.columns):
        return
    coords = corridor_df[["Latitude", "Longitude"]].values.tolist()
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
    required = {"Latitude", "Longitude", "Label"}
    if not required.issubset(markers_df.columns):
        return
    df = markers_df.copy()
    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    df = df.dropna(subset=["Latitude", "Longitude"])
    for _, row in df.iterrows():
        marker_type = str(row.get("Type", "annotation")).lower()
        color = MARKER_HEX.get(marker_type, DEFAULT_MARKER_HEX)
        folium.CircleMarker(
            location=[row["Latitude"], row["Longitude"]],
            radius=7 if marker_type == "epicentre" else 5,
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.9,
            tooltip=str(row.get("Label", "")),
            popup=str(row.get("Note", row.get("Label", ""))),
        ).add_to(m)


def build_folium_map(district_gdf, municipality_gdf, corridor_df, markers_df):
    m = folium.Map(
        location=[DEFAULT_LAT, DEFAULT_LON],
        zoom_start=DEFAULT_ZOOM,
        tiles="CartoDB positron",
        control_scale=True,
    )

    # ---- District 레이어 (요구사항 2, 3) ----
    district_geojson = json.loads(district_gdf.to_json())
    folium.GeoJson(
        district_geojson,
        name="Districts (Flood Severity)",
        style_function=district_style_function,
        highlight_function=district_highlight_function,
        tooltip=folium.GeoJsonTooltip(
            fields=["District", "Severity"],
            aliases=["District:", "Severity:"],
            sticky=True,
        ),
    ).add_to(m)

    # ---- Municipality 레이어 (요구사항 4) ----
    if municipality_gdf is not None:
        muni_geojson = json.loads(municipality_gdf.to_json())
        muni_fields = [MUNICIPALITY_NAME_FIELD]
        muni_aliases = ["Municipality:"]
        if JOIN_KEY in municipality_gdf.columns:
            muni_fields.append(JOIN_KEY)
            muni_aliases.append("District ID:")

        folium.GeoJson(
            muni_geojson,
            name="Municipalities",
            style_function=municipality_style_function,
            highlight_function=municipality_highlight_function,
            tooltip=folium.GeoJsonTooltip(fields=muni_fields, aliases=muni_aliases, sticky=True),
        ).add_to(m)

    add_corridor_to_map(m, corridor_df)
    add_markers_to_map(m, markers_df)

    folium.LayerControl(collapsed=False).add_to(m)

    try:
        minx, miny, maxx, maxy = district_gdf.total_bounds
        if all(pd.notna([minx, miny, maxx, maxy])):
            m.fit_bounds([[miny, minx], [maxy, maxx]])
    except Exception:
        pass

    return m


# ============================================================
# 6. 필수 파일 체크 (요구사항 9)
# ============================================================

required_files = {
    "population.csv": POPULATION_PATH,
    "flood.geojson": FLOOD_GEOJSON_PATH,
}

missing_files = [f"data/{name}" for name, path in required_files.items() if not Path(path).exists()]

if missing_files:
    st.error("필수 데이터 파일을 찾을 수 없습니다. (Required data files are missing.)")
    st.markdown("다음 파일들이 없습니다:")
    for filename in missing_files:
        st.code(filename)
    st.markdown(
        "GitHub repository의 `data/` 폴더 구조를 확인해주세요. "
        "필요한 파일 목록은 이 파일 최상단 주석(docstring)에 정리되어 있습니다."
    )
    st.stop()


# ============================================================
# 7. population.csv 로드
# ============================================================

try:
    population_df = load_population_csv(POPULATION_PATH)
except Exception as e:
    st.error("population.csv 로드에 실패했습니다.")
    st.code(str(e))
    st.write(f"File path: {POPULATION_PATH}")
    st.stop()


# ============================================================
# 8. 사이드바 — 데이터 소스
# ============================================================

st.sidebar.title("⚙️ Data Sources")

st.sidebar.markdown("### Repository data")
st.sidebar.caption("Population data")
st.sidebar.code("data/population.csv")

excel_path = st.sidebar.text_input(
    "Excel file path (District_Situation / Water_Access / Interventions / Funding / Facilities)",
    value=str(DEFAULT_EXCEL_PATH),
)

geojson_path = st.sidebar.text_input(
    "Flood GeoJSON path (District boundary)",
    value=str(FLOOD_GEOJSON_PATH),
)

municipality_path = st.sidebar.text_input(
    "Municipality GeoJSON path (선택)",
    value=str(MUNICIPALITY_GEOJSON_PATH),
)

with st.sidebar.expander("Advanced data sources", expanded=False):
    municipality_situation_path = st.text_input(
        "Municipality situation CSV (선택)", value=str(MUNICIPALITY_SITUATION_PATH)
    )
    event_info_path = st.text_input("Event information", value=str(DEFAULT_EVENT_INFO_PATH))
    markers_path = st.text_input("Event markers", value=str(DEFAULT_EVENT_MARKERS_PATH))
    corridor_path = st.text_input("Flood corridor", value=str(DEFAULT_FLOOD_CORRIDOR_PATH))
    district_csv_path = st.text_input(
        "District situation CSV (optional override)", value=str(DEFAULT_DISTRICT_CSV_PATH)
    )

if st.sidebar.button("🔄 Refresh data"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption(f"Population records: {len(population_df):,}")


# ============================================================
# 9. flood.geojson (District 경계) 로드 — 필수
# ============================================================

flood_gdf, flood_notes = read_geo_file(geojson_path, "flood.geojson (District 경계)")

if flood_gdf is None:
    st.error("flood.geojson을 불러오지 못했습니다. (Failed to load flood.geojson.)")
    for note in flood_notes:
        st.write(f"- {note}")
    st.markdown(
        """
        다음을 확인해주세요:

        1. `data/flood.geojson` 파일이 실제로 존재하는지
        2. `pyogrio` 가 requirements.txt에 포함되어 있는지
        3. flood.geojson이 유효한 GeoJSON 형식이며, `District_ID` / `District` 컬럼을 포함하는지
        """
    )
    st.write(f"File path: {geojson_path}")
    st.stop()

if JOIN_KEY not in flood_gdf.columns:
    st.error(f"flood.geojson에 필수 join 컬럼 '{JOIN_KEY}'이(가) 없습니다.")
    st.markdown("flood.geojson에 포함된 컬럼:")
    st.code("\n".join(str(c) for c in flood_gdf.columns))
    st.stop()

if flood_notes:
    with st.sidebar.expander("ℹ️ flood.geojson 처리 로그", expanded=False):
        for note in flood_notes:
            st.caption(note)

st.sidebar.caption(f"Flood/District features: {len(flood_gdf):,}")
flood_gdf[JOIN_KEY] = flood_gdf[JOIN_KEY].astype(str).str.strip()


# ============================================================
# 10. municipality.geojson 로드 — 선택 (요구사항 4, 6, 9)
# ============================================================

municipality_gdf = None

if municipality_path and Path(municipality_path).exists():
    municipality_gdf, muni_notes = read_geo_file(municipality_path, "municipality.geojson (Municipality 경계)")

    if municipality_gdf is not None:
        if MUNICIPALITY_NAME_FIELD not in municipality_gdf.columns:
            st.sidebar.warning(
                f"municipality.geojson에 '{MUNICIPALITY_NAME_FIELD}' 컬럼이 없어 "
                "Municipality 레이어를 표시하지 않습니다."
            )
            municipality_gdf = None
        elif JOIN_KEY not in municipality_gdf.columns:
            st.sidebar.info(
                f"municipality.geojson에 '{JOIN_KEY}' 컬럼이 없어 "
                "District와의 연결 정보 없이 경계만 표시됩니다."
            )

    if municipality_gdf is not None and JOIN_KEY in municipality_gdf.columns:
        municipality_gdf[JOIN_KEY] = municipality_gdf[JOIN_KEY].astype(str).str.strip()

    if muni_notes:
        with st.sidebar.expander("ℹ️ municipality.geojson 처리 로그", expanded=False):
            for note in muni_notes:
                st.caption(note)

    if municipality_gdf is not None:
        st.sidebar.caption(f"Municipality features: {len(municipality_gdf):,}")
else:
    st.sidebar.info("municipality.geojson이 없어 District 레이어만 표시합니다.")


# ============================================================
# 11. 타이틀 / 메타정보
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

st.caption("지도에서 District 또는 Municipality를 클릭하면 상세 정보를 확인할 수 있습니다.")


# ============================================================
# 12. 범례
# ============================================================

if SEVERITY_LEGEND:
    with st.expander("ℹ️ Map information and Impact Level", expanded=True):
        legend_columns = st.columns(len(SEVERITY_LEGEND))
        for column, (label, hex_color) in zip(legend_columns, SEVERITY_LEGEND):
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
        st.caption("굵은 테두리 = 홍수 영향 데이터가 있는 District / 점선 = Municipality 경계")


# ============================================================
# 13. District_Situation + 관련 시트 로드
# ============================================================

try:
    district_df, district_error = load_district_situation(excel_path, district_csv_path)
except Exception as e:
    district_df, district_error = None, str(e)

if district_error:
    st.error("District_Situation 데이터를 불러오지 못했습니다.")
    st.code(str(district_error))
    st.stop()

if district_df is None or district_df.empty:
    st.error("District 데이터가 비어 있습니다.")
    st.stop()

if JOIN_KEY not in district_df.columns:
    st.error(f"District 데이터에 필수 join 컬럼 '{JOIN_KEY}'이(가) 없습니다.")
    st.markdown("District 데이터에 포함된 컬럼:")
    st.code("\n".join(str(c) for c in district_df.columns))
    st.stop()

district_df[JOIN_KEY] = district_df[JOIN_KEY].astype(str).str.strip()

water_access_df, water_access_err = load_water_access(excel_path)
interventions_df, interventions_err = load_interventions(excel_path)
funding_df, funding_err = load_funding(excel_path)
facilities_df, facilities_err = load_facilities(excel_path)


# ============================================================
# 14. Join + Join 품질 체크
# ============================================================

try:
    merged = join_district_data(flood_gdf, district_df)
except Exception as e:
    st.error("flood.geojson과 District 데이터를 Join하지 못했습니다.")
    st.code(str(e))
    st.stop()

geojson_ids = set(flood_gdf[JOIN_KEY].dropna().astype(str))
district_ids = set(district_df[JOIN_KEY].dropna().astype(str))
missing_in_district = geojson_ids - district_ids
missing_in_geojson = district_ids - geojson_ids

if missing_in_district or missing_in_geojson:
    with st.expander("🔍 Join quality check", expanded=False):
        if missing_in_district:
            st.warning("flood.geojson에는 있지만 District 데이터에는 없는 ID")
            st.code("\n".join(sorted(missing_in_district)))
        if missing_in_geojson:
            st.warning("District 데이터에는 있지만 flood.geojson에는 없는 ID")
            st.code("\n".join(sorted(missing_in_geojson)))


# ============================================================
# 15. Flood corridor / event markers
# ============================================================

try:
    corridor_df, corridor_error = load_flood_corridor(corridor_path)
except Exception as e:
    corridor_df, corridor_error = None, str(e)

try:
    markers_df, markers_error = load_event_markers(markers_path)
except Exception as e:
    markers_df, markers_error = None, str(e)


# ============================================================
# 16. 지도 + 상세 패널 레이아웃
# ============================================================

map_column, detail_column = st.columns([2, 1])

with map_column:
    try:
        fmap = build_folium_map(merged, municipality_gdf, corridor_df, markers_df)
    except Exception as e:
        st.error("지도를 생성하지 못했습니다.")
        st.code(str(e))
        st.stop()

    try:
        map_state = st_folium(
            fmap,
            use_container_width=True,
            height=620,
            key="flood_map",
        )
    except Exception as e:
        st.error("지도를 표시하지 못했습니다.")
        st.code(str(e))
        st.stop()

clicked_props = None
selection_type = None

if map_state and map_state.get("last_active_drawing"):
    clicked_props = map_state["last_active_drawing"].get("properties", {}) or {}
    if MUNICIPALITY_NAME_FIELD in clicked_props and "District" not in clicked_props:
        selection_type = "municipality"
    elif "District" in clicked_props:
        selection_type = "district"


# ============================================================
# 17. 상세 패널 헬퍼
# ============================================================

def get_district_id(props):
    if not props:
        return None
    return props.get(JOIN_KEY)


def render_sheet_table(df, err, district_id, empty_message):
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
            summary_cols[i % len(summary_cols)].metric(f"합계: {cname}", format_number(sub[cname].sum()))


def collect_sources(district_id):
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
# 18. 상세 패널 — District / Municipality 분기 (요구사항 5)
# ============================================================

with detail_column:
    st.subheader("📋 Detail")

    if selection_type == "district":
        selected_properties = clicked_props
        district_id = get_district_id(selected_properties)
        district_name = format_text(selected_properties.get("District", "Unknown District"))
        severity = format_text(selected_properties.get("Severity", "No Data"))

        st.markdown(f"### 🗺️ {district_name}")

        severity_colors = dict(SEVERITY_LEGEND)
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
            tab_situation, tab_water, tab_interventions,
            tab_funding, tab_facilities, tab_sources, tab_population,
        ) = st.tabs(
            ["🗺️ Situation", "💧 Water Access", "🛠️ Interventions",
             "💰 Funding", "🏫 Facilities", "📚 Sources", "👥 Population"]
        )

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
            severity_reason = selected_properties.get("Severity_reason") or selected_properties.get("Priority_reason")
            st.write(format_text(severity_reason))

            st.markdown("---")
            st.markdown("**Immediate intervention**")
            st.write(format_text(selected_properties.get("Immediate_intervention")))
            st.markdown("**Medium-term intervention**")
            st.write(format_text(selected_properties.get("Medium_term_intervention")))
            st.markdown("**Long-term intervention**")
            st.write(format_text(selected_properties.get("Long_term_intervention")))

        with tab_water:
            render_sheet_table(water_access_df, water_access_err, district_id, "이 District에 대한 Water Access 데이터가 없습니다.")

        with tab_interventions:
            render_sheet_table(interventions_df, interventions_err, district_id, "이 District에 대한 Interventions 데이터가 없습니다.")

        with tab_funding:
            render_sheet_table(funding_df, funding_err, district_id, "이 District에 대한 Funding 데이터가 없습니다.")

        with tab_facilities:
            render_sheet_table(facilities_df, facilities_err, district_id, "이 District에 대한 Facilities 데이터가 없습니다.")

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
                st.markdown("**Other sheet sources**")
                st.dataframe(pd.DataFrame(other_sources), use_container_width=True, hide_index=True)

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
                        f"District_Situation의 Affected_population: {format_number(affected)} "
                        "vs population.csv 값과 비교해 보세요."
                    )

    elif selection_type == "municipality":
        muni_name = format_text(clicked_props.get(MUNICIPALITY_NAME_FIELD))
        muni_district_id = clicked_props.get(JOIN_KEY)

        parent_row = district_df[district_df[JOIN_KEY].astype(str) == str(muni_district_id)] if muni_district_id else pd.DataFrame()
        parent_name = (
            format_text(parent_row.iloc[0].get("District"))
            if not parent_row.empty and "District" in parent_row.columns
            else format_text(muni_district_id)
        )

        st.markdown(f"### 🏘️ {muni_name}")
        st.caption(f"Parent District: {parent_name}")

        muni_situation_df, muni_situation_err = load_municipality_situation(municipality_situation_path)
        if muni_situation_err or muni_situation_df.empty:
            st.info("municipality_situation.csv가 없어 기본 속성만 표시합니다. (data/municipality_situation.csv 추가 시 상세 정보가 표시됩니다.)")
            display_props = {k: v for k, v in clicked_props.items() if k not in ("severity_hex", "is_flood_affected")}
            st.json(display_props)
        else:
            sub = filter_municipality_rows(muni_situation_df, clicked_props)
            if sub.empty:
                st.info("이 Municipality에 대한 상세 데이터가 없습니다.")
            else:
                st.dataframe(sub.reset_index(drop=True), use_container_width=True, hide_index=True)

    else:
        st.info("지도에서 District 또는 Municipality를 클릭하면 상세 정보가 표시됩니다.")


# ============================================================
# 19. 푸터
# ============================================================

st.markdown("---")
st.caption("UNICEF Nepal | WASH Emergency Response")
st.caption(
    f"District layer: {Path(geojson_path).name} | "
    f"Municipality layer: {Path(municipality_path).name if municipality_gdf is not None else 'not loaded'} | "
    "Map tiles: CartoDB Positron (no API key required)"
)
