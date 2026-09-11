# -*- coding: utf-8 -*-
"""
UNICEF Nepal — Flood WASH Emergency Decision Support Dashboard
Phase 1 MVP

GitHub / Streamlit Cloud 파일 구조
----------------------------------

├── app.py
├── config.py
├── data_loader.py
├── map_utils.py
├── requirements.txt
│
└── data/
    ├── population.csv
    └── flood.geojson

중요:
- app.py는 동일 Repository에 있는 config.py, data_loader.py, map_utils.py를 import합니다.
- population.csv와 flood.geojson은 data/ 폴더에서 별도로 읽습니다.
- 모든 데이터 경로는 절대경로가 아닌 Repository 기준 상대경로로 설정합니다.
- 따라서 Streamlit Cloud에서도 동일한 구조로 실행할 수 있습니다.

실행:
    streamlit run app.py
"""

from pathlib import Path

import pandas as pd
import streamlit as st
import geopandas as gpd

import config
import data_loader
import map_utils

from config import *
from data_loader import load_data


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="UNICEF Nepal — Flood WASH Dashboard",
    page_icon="🌊",
    layout="wide",
)


# ============================================================
# REPOSITORY / DATA FILE PATHS
# ============================================================
#
# GitHub Repository 구조를 기준으로 상대경로를 설정합니다.
#
# app.py
# └── data/
#       ├── population.csv
#       └── flood.geojson
#
# Streamlit Cloud에서도 이 구조가 그대로 유지되므로
# 별도의 로컬 컴퓨터 경로를 지정할 필요가 없습니다.
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

POPULATION_PATH = DATA_DIR / "population.csv"
FLOOD_GEOJSON_PATH = DATA_DIR / "flood.geojson"


# ============================================================
# REQUIRED FILE CHECK
# ============================================================

required_files = {
    "population.csv": POPULATION_PATH,
    "flood.geojson": FLOOD_GEOJSON_PATH,
}

missing_files = [
    name for name, path in required_files.items()
    if not path.exists()
]

if missing_files:
    st.error(
        "필수 데이터 파일을 찾을 수 없습니다.\n\n"
        + "\n".join([f"- {file}" for file in missing_files])
        + "\n\n"
        "GitHub Repository에 다음 구조가 존재하는지 확인해주세요:\n\n"
        "data/population.csv\n"
        "data/flood.geojson"
    )
    st.stop()


# ============================================================
# LOAD REPOSITORY DATA FILES
# ============================================================
#
# population.csv와 flood.geojson은 Excel과 별도의 독립적인
# 데이터 파일입니다.
#
# 즉, Dashboard 실행 시 다음과 같이 각각 읽습니다.
#
# population.csv  -> pandas DataFrame
# flood.geojson   -> GeoDataFrame
#
# ============================================================

@st.cache_data
def load_population_data(file_path):
    """Load population.csv from the repository data folder."""
    return pd.read_csv(file_path)


@st.cache_data
def load_flood_geojson(file_path):
    """Load flood.geojson from the repository data folder."""
    return gpd.read_file(file_path)


try:
    population_df = load_population_data(POPULATION_PATH)
except Exception as e:
    st.error(
        f"population.csv를 불러오지 못했습니다.\n\n"
        f"파일 위치: {POPULATION_PATH}\n\n"
        f"오류: {e}"
    )
    st.stop()


try:
    flood_gdf = load_flood_geojson(FLOOD_GEOJSON_PATH)
except Exception as e:
    st.error(
        f"flood.geojson을 불러오지 못했습니다.\n\n"
        f"파일 위치: {FLOOD_GEOJSON_PATH}\n\n"
        f"오류: {e}"
    )
    st.stop()


# ============================================================
# SIDEBAR — DATA SOURCE
# ============================================================

st.sidebar.title("⚙️ Data source")

# Repository 내부 population.csv
st.sidebar.caption(
    f"Population data: `{POPULATION_PATH.relative_to(BASE_DIR)}`"
)

# Repository 내부 flood.geojson
st.sidebar.caption(
    f"Flood GeoJSON: `{FLOOD_GEOJSON_PATH.relative_to(BASE_DIR)}`"
)

excel_path = st.sidebar.text_input(
    "Excel 파일 경로",
    value=str(config.DEFAULT_EXCEL_PATH),
    help=(
        "WASH_data_template.xlsx 의 실제 경로. "
        "팀 공유 경로로 바꿔도 됩니다."
    ),
)

# 기본 GeoJSON은 Repository의 flood.geojson 사용
geojson_path = st.sidebar.text_input(
    "District boundary / Flood GeoJSON 파일 경로",
    value=str(FLOOD_GEOJSON_PATH),
    help=(
        "GitHub Repository의 data/flood.geojson을 사용합니다."
    ),
)


# ============================================================
# ADVANCED DATA SOURCES
# ============================================================

with st.sidebar.expander(
    "고급: 이벤트 메타 / corridor / marker 경로",
    expanded=False,
):

    event_info_path = st.text_input(
        "이벤트 정보 CSV (제목/부제목)",
        value=str(config.DEFAULT_EVENT_INFO_PATH),
    )

    markers_path = st.text_input(
        "주석 marker CSV (epicentre 등)",
        value=str(config.DEFAULT_EVENT_MARKERS_PATH),
    )

    corridor_path = st.text_input(
        "Flood corridor 경로 CSV",
        value=str(config.DEFAULT_FLOOD_CORRIDOR_PATH),
    )


# ============================================================
# REFRESH
# ============================================================

if st.sidebar.button("🔄 데이터 새로고침"):
    st.cache_data.clear()
    st.rerun()


st.sidebar.markdown("---")

st.sidebar.caption(
    f"Population records: **{len(population_df):,}**"
)

st.sidebar.caption(
    f"Flood GeoJSON features: **{len(flood_gdf):,}**"
)

st.sidebar.caption(
    f"Workbook last modified: "
    f"**{data_loader.workbook_last_updated(excel_path)}**"
)


# ============================================================
# DATA FILE INFORMATION
# ============================================================

with st.sidebar.expander("📁 Repository data files", expanded=False):

    st.write("**Population CSV**")
    st.code(str(POPULATION_PATH.relative_to(BASE_DIR)))

    st.write("**Flood GeoJSON**")
    st.code(str(FLOOD_GEOJSON_PATH.relative_to(BASE_DIR)))

    st.write("**Population columns**")
    st.write(list(population_df.columns))

    st.write("**Flood GeoJSON columns**")
    st.write(list(flood_gdf.columns))


# ============================================================
# TITLE
# ============================================================

event_info, event_info_err = data_loader.load_event_info(
    event_info_path
)

st.markdown(
    f"""
    <h1 style='margin-bottom:0;'>
        {event_info.get(
            'Title',
            'UNICEF Nepal — Flood WASH Dashboard'
        )}
    </h1>
    """,
    unsafe_allow_html=True,
)


if event_info.get("Subtitle"):
    st.markdown(
        f"""
        <h3 style='color:#B22222;margin-top:2px;'>
            {event_info['Subtitle']}
        </h3>
        """,
        unsafe_allow_html=True,
    )


meta_bits = [
    v
    for v in [
        event_info.get("Corridor_note"),
        event_info.get("Data_asof"),
    ]
    if v
]


if meta_bits:
    st.caption("  ·  ".join(meta_bits))


st.caption(
    "District를 클릭하면 상세 flood/WASH 상황이 표시됩니다."
)


# ============================================================
# MAP LEGEND
# ============================================================

with st.expander(
    "ℹ️ 지도 정보 / 범례 (Impact level)",
    expanded=True,
):

    legend_cols = st.columns(
        len(config.SEVERITY_LEGEND)
    )

    for col, (label, hexcolor) in zip(
        legend_cols,
        config.SEVERITY_LEGEND,
    ):

        col.markdown(
            f"""
            <span style='display:inline-block;
            width:14px;height:14px;
            background:{hexcolor};
            border-radius:3px;
            margin-right:6px;'></span>
            {label}
            """,
            unsafe_allow_html=True,
        )

    st.write(
        "District 색상은 `district_situation.csv` "
        "(없으면 Excel의 `District_Situation` 시트)의 "
        "`Severity` 컬럼을 따릅니다. "
        "CSV 값을 바꾸고 새로고침하면 지도에 바로 반영됩니다."
    )


# ============================================================
# LOAD DISTRICT DATA
# ============================================================

district_df, district_err = (
    data_loader.load_district_situation(excel_path)
)


# ============================================================
# LOAD FLOOD GEOJSON
# ============================================================
#
# flood.geojson은 Repository의 data/ 폴더에서 직접 읽습니다.
#
# 이후 map_utils.py에서 필요한 경우 이 GeoDataFrame을
# flood layer 또는 district boundary로 활용할 수 있습니다.
# ============================================================

boundary_gdf = flood_gdf.copy()
boundary_err = None


# ============================================================
# ERROR CHECK
# ============================================================

if district_err:
    st.error(
        "District_Situation 시트를 불러오지 못했습니다: "
        f"{district_err}"
    )
    st.stop()


if boundary_err:
    st.error(
        "Flood GeoJSON을 불러오지 못했습니다: "
        f"{boundary_err}"
    )
    st.stop()


# ============================================================
# JOIN DISTRICT DATA
# ============================================================

merged = map_utils.join_district_data(
    boundary_gdf,
    district_df,
)


# ============================================================
# JOIN QUALITY CHECK
# ============================================================

if config.JOIN_KEY in boundary_gdf.columns:

    boundary_ids = set(
        boundary_gdf[config.JOIN_KEY]
    )

else:

    boundary_ids = set()


if config.JOIN_KEY in district_df.columns:

    excel_ids = set(
        district_df[config.JOIN_KEY]
    )

else:

    excel_ids = set()


missing_in_excel = (
    boundary_ids - excel_ids
)

missing_in_boundary = (
    excel_ids - boundary_ids
)


if missing_in_excel or missing_in_boundary:

    with st.expander(
        "🔍 Join 점검 (District_ID 불일치)",
        expanded=False,
    ):

        if missing_in_excel:

            st.write(
                "경계 파일에는 있지만 Excel에 데이터가 없는 "
                f"District_ID ({len(missing_in_excel)}개):"
            )

            st.code(
                ", ".join(
                    sorted(
                        map(str, missing_in_excel)
                    )
                )
            )

        if missing_in_boundary:

            st.write(
                "Excel에는 있지만 경계 파일에 없는 "
                f"District_ID ({len(missing_in_boundary)}개):"
            )

            st.code(
                ", ".join(
                    sorted(
                        map(str, missing_in_boundary)
                    )
                )
            )


# ============================================================
# MAP DATA
# ============================================================

corridor_df, corridor_err = (
    data_loader.load_flood_corridor(
        corridor_path
    )
)

markers_df, markers_err = (
    data_loader.load_event_markers(
        markers_path
    )
)


extra_layers = [
    map_utils.build_corridor_layer(
        corridor_df
    )
]


point_layer, text_layer = (
    map_utils.build_marker_layers(
        markers_df
    )
)


extra_layers += [
    point_layer,
    text_layer,
]


# ============================================================
# MAP + DETAIL PANEL
# ============================================================

map_col, detail_col = st.columns(
    [2, 1]
)


# ============================================================
# MAP
# ============================================================

with map_col:

    deck = map_utils.build_deck(
        merged,
        extra_layers=extra_layers,
    )

    event = st.pydeck_chart(
        deck,
        on_select="rerun",
        selection_mode="single-object",
        key="district_map",
        use_container_width=True,
        height=620,
    )


# ============================================================
# SELECTED DISTRICT DETAIL PANEL
# ============================================================

def _fmt_number(value):

    if value is None or pd.isna(value):
        return "No data"

    try:
        return f"{int(value):,}"

    except (ValueError, TypeError):
        return str(value)


def _fmt_currency(value):

    if value is None or pd.isna(value):
        return "No data"

    try:
        return f"USD {int(value):,}"

    except (ValueError, TypeError):
        return str(value)


def _fmt_text(value):

    if (
        value is None
        or pd.isna(value)
        or str(value).strip() == ""
    ):
        return "_(not recorded)_"

    return str(value)


selected_props = None


try:

    objects = event.selection[
        "objects"
    ].get(
        "district-layer",
        []
    )

    if objects:
        selected_props = objects[0]

except Exception:

    selected_props = None


# ============================================================
# DETAIL PANEL
# ============================================================

with detail_col:

    st.subheader("📋 District detail")

    if not selected_props:

        st.info(
            "지도에서 District를 클릭하면 "
            "상세 정보가 여기에 표시됩니다."
        )

    else:

        severity = selected_props.get(
            "Severity",
            "No Data",
        )

        severity_hex = dict(
            config.SEVERITY_LEGEND
        ).get(
            severity,
            "#BEBEBE",
        )

        st.markdown(
            f"### {selected_props.get('District', 'Unknown')}"
        )

        st.markdown(
            f"""
            <span style='background:{severity_hex};
            color:white;padding:2px 10px;
            border-radius:10px;
            font-weight:600;'>
            {str(severity).upper()}
            </span>
            """,
            unsafe_allow_html=True,
        )

        c1, c2 = st.columns(2)

        c1.metric(
            "Affected population",
            _fmt_number(
                selected_props.get(
                    "Affected_population"
                )
            ),
        )

        c2.metric(
            "Water access population",
            _fmt_number(
                selected_props.get(
                    "Water_access_population"
                )
            ),
        )

        c3, c4 = st.columns(2)

        c3.metric(
            "Casualties",
            _fmt_number(
                selected_props.get(
                    "Casualties"
                )
            ),
        )

        c4.metric(
            "Estimated funding requirement",
            _fmt_currency(
                selected_props.get(
                    "Estimated_funding_USD"
                )
            ),
        )

        st.markdown(
            "**Current water source:** "
            + _fmt_text(
                selected_props.get(
                    "Water_source_current"
                )
            )
        )

        st.markdown("---")

        st.markdown("**Problem / Needs**")

        st.write(
            _fmt_text(
                selected_props.get(
                    "Key_problems"
                )
            )
        )

        st.markdown("**Severity reason**")

        st.write(
            _fmt_text(
                selected_props.get(
                    "Severity_reason",
                    selected_props.get(
                        "Priority_reason"
                    ),
                )
            )
        )

        st.markdown("---")

        st.markdown(
            "**Immediate intervention**"
        )

        st.write(
            _fmt_text(
                selected_props.get(
                    "Immediate_intervention"
                )
            )
        )

        st.markdown(
            "**Medium-term intervention**"
        )

        st.write(
            _fmt_text(
                selected_props.get(
                    "Medium_term_intervention"
                )
            )
        )

        st.markdown(
            "**Long-term intervention**"
        )

        st.write(
            _fmt_text(
                selected_props.get(
                    "Long_term_intervention"
                )
            )
        )

        st.markdown("---")

        c3, c4 = st.columns(2)

        c3.caption(
            f"Data source: "
            f"{selected_props.get('Data_source', '-')}"
        )

        c4.caption(
            f"Last updated: "
            f"{selected_props.get('Last_updated', '-')}"
        )

        st.caption(
            f"Data status: "
            f"{selected_props.get('Data_status', '-')}"
        )


# ============================================================
# FOOTER
# ============================================================

st.markdown("---")

st.caption(
    "UNICEF Nepal | WASH Emergency Response | "
    f"Source: {excel_path} | "
    f"Population: {POPULATION_PATH.name} | "
    f"Flood layer: {FLOOD_GEOJSON_PATH.name} | "
    f"Workbook last modified: "
    f"{data_loader.workbook_last_updated(excel_path)}"
)