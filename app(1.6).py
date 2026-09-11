```python
# -*- coding: utf-8 -*-
"""
UNICEF Nepal — Flood WASH Emergency Decision Support Dashboard
Phase 1 MVP

GitHub Repository structure
---------------------------

├── app(1.5).py
├── config.py
├── data_loader.py
├── map_utils.py
├── requirements.txt
│
└── data/
    ├── population.csv
    └── flood.geojson

Important:
- config.py, data_loader.py, map_utils.py are imported from the same repository.
- population.csv and flood.geojson are stored separately under data/.
- All paths are relative to the repository.
- GeoJSON is loaded using pyogrio instead of Fiona.
- This structure is designed for Streamlit Cloud deployment.

Run:
    streamlit run "app(1.5).py"
"""

from pathlib import Path

import pandas as pd
import streamlit as st
import geopandas as gpd

import config
import data_loader
import map_utils


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="UNICEF Nepal — Flood WASH Dashboard",
    page_icon="🌊",
    layout="wide",
)


# ============================================================
# REPOSITORY PATHS
# ============================================================

# Repository root
BASE_DIR = Path(__file__).resolve().parent

# Data folder
DATA_DIR = BASE_DIR / "data"

# Individual data files
POPULATION_PATH = DATA_DIR / "population.csv"
FLOOD_GEOJSON_PATH = DATA_DIR / "flood.geojson"


# ============================================================
# CHECK REQUIRED FILES
# ============================================================

required_files = {
    "population.csv": POPULATION_PATH,
    "flood.geojson": FLOOD_GEOJSON_PATH,
}

missing_files = [
    name
    for name, path in required_files.items()
    if not path.exists()
]

if missing_files:

    st.error(
        "Required data files are missing from the repository.\n\n"
        + "\n".join(
            f"- data/{file}"
            for file in missing_files
        )
        + "\n\n"
        "Please make sure the GitHub repository has the following structure:\n\n"
        "data/population.csv\n"
        "data/flood.geojson"
    )

    st.stop()


# ============================================================
# LOAD POPULATION CSV
# ============================================================

@st.cache_data
def load_population_data(file_path):
    """
    Load population.csv from the repository data folder.
    """

    return pd.read_csv(file_path)


try:

    population_df = load_population_data(
        POPULATION_PATH
    )

except Exception as e:

    st.error(
        "Failed to load population.csv.\n\n"
        f"File: {POPULATION_PATH}\n\n"
        f"Error: {e}"
    )

    st.stop()


# ============================================================
# LOAD FLOOD GEOJSON
# ============================================================

@st.cache_data
def load_flood_geojson(file_path):
    """
    Load flood.geojson using pyogrio.

    Fiona is intentionally NOT used.
    This avoids the GDAL / fiona build problem
    encountered on Streamlit Cloud.
    """

    return gpd.read_file(
        file_path,
        engine="pyogrio",
    )


try:

    flood_gdf = load_flood_geojson(
        FLOOD_GEOJSON_PATH
    )

except Exception as e:

    st.error(
        "Failed to load flood.geojson.\n\n"
        f"File: {FLOOD_GEOJSON_PATH}\n\n"
        f"Error: {e}\n\n"
        "Please check that pyogrio is included in "
        "requirements.txt."
    )

    st.stop()


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("⚙️ Data source")


st.sidebar.caption(
    "Repository data files"
)


st.sidebar.code(
    "data/population.csv"
)


st.sidebar.code(
    "data/flood.geojson"
)


# ============================================================
# EXCEL DATA SOURCE
# ============================================================

excel_path = st.sidebar.text_input(
    "Excel 파일 경로",
    value=str(
        config.DEFAULT_EXCEL_PATH
    ),
    help=(
        "WASH_data_template.xlsx 의 실제 경로. "
        "필요한 경우 팀 공유 경로로 변경할 수 있습니다."
    ),
)


# ============================================================
# GEOJSON DATA SOURCE
# ============================================================

geojson_path = st.sidebar.text_input(
    "District boundary / Flood GeoJSON",
    value=str(
        FLOOD_GEOJSON_PATH
    ),
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
        value=str(
            config.DEFAULT_EVENT_INFO_PATH
        ),
    )

    markers_path = st.text_input(
        "주석 marker CSV (epicentre 등)",
        value=str(
            config.DEFAULT_EVENT_MARKERS_PATH
        ),
    )

    corridor_path = st.text_input(
        "Flood corridor 경로 CSV",
        value=str(
            config.DEFAULT_FLOOD_CORRIDOR_PATH
        ),
    )


# ============================================================
# REFRESH
# ============================================================

if st.sidebar.button("🔄 데이터 새로고침"):

    st.cache_data.clear()

    st.rerun()


# ============================================================
# DATA STATUS
# ============================================================

st.sidebar.markdown("---")

st.sidebar.caption(
    f"Population records: **{len(population_df):,}**"
)

st.sidebar.caption(
    f"Flood GeoJSON features: **{len(flood_gdf):,}**"
)

st.sidebar.caption(
    "GeoJSON engine: **pyogrio**"
)

st.sidebar.caption(
    f"Workbook last modified: "
    f"**{data_loader.workbook_last_updated(excel_path)}**"
)


# ============================================================
# REPOSITORY DATA INFORMATION
# ============================================================

with st.sidebar.expander(
    "📁 Repository data files",
    expanded=False,
):

    st.write("**Population CSV**")

    st.code(
        str(
            POPULATION_PATH.relative_to(
                BASE_DIR
            )
        )
    )

    st.write("**Flood GeoJSON**")

    st.code(
        str(
            FLOOD_GEOJSON_PATH.relative_to(
                BASE_DIR
            )
        )
    )

    st.write("**Population columns**")

    st.write(
        list(
            population_df.columns
        )
    )

    st.write("**Flood GeoJSON columns**")

    st.write(
        list(
            flood_gdf.columns
        )
    )


# ============================================================
# TITLE
# ============================================================

event_info, event_info_err = (
    data_loader.load_event_info(
        event_info_path
    )
)


st.markdown(
    f"""
    <h1 style='margin-bottom:0;'>
        {event_info.get(
            "Title",
            "UNICEF Nepal — Flood WASH Dashboard"
        )}
    </h1>
    """,
    unsafe_allow_html=True,
)


if event_info.get("Subtitle"):

    st.markdown(
        f"""
        <h3 style='color:#B22222;margin-top:2px;'>
            {event_info["Subtitle"]}
        </h3>
        """,
        unsafe_allow_html=True,
    )


meta_bits = [
    value
    for value in [
        event_info.get("Corridor_note"),
        event_info.get("Data_asof"),
    ]
    if value
]


if meta_bits:

    st.caption(
        "  ·  ".join(meta_bits)
    )


st.caption(
    "District를 클릭하면 상세 flood/WASH 상황이 표시됩니다."
)


# ============================================================
# LEGEND
# ============================================================

with st.expander(
    "ℹ️ 지도 정보 / 범례 (Impact level)",
    expanded=True,
):

    legend_cols = st.columns(
        len(config.SEVERITY_LEGEND)
    )

    for col, (
        label,
        hexcolor,
    ) in zip(
        legend_cols,
        config.SEVERITY_LEGEND,
    ):

        col.markdown(
            f"""
            <span style='
                display:inline-block;
                width:14px;
                height:14px;
                background:{hexcolor};
                border-radius:3px;
                margin-right:6px;
            '></span>
            {label}
            """,
            unsafe_allow_html=True,
        )

    st.write(
        "District 색상은 "
        "`district_situation.csv` "
        "(없으면 Excel의 "
        "`District_Situation` 시트)의 "
        "`Severity` 컬럼을 따릅니다."
    )


# ============================================================
# LOAD DISTRICT DATA
# ============================================================

district_df, district_err = (
    data_loader.load_district_situation(
        excel_path
    )
)


if district_err:

    st.error(
        "District_Situation 시트를 "
        "불러오지 못했습니다: "
        f"{district_err}"
    )

    st.stop()


# ============================================================
# USE FLOOD GEOJSON AS BOUNDARY DATA
# ============================================================

boundary_gdf = flood_gdf.copy()


# ============================================================
# CHECK JOIN KEY
# ============================================================

if config.JOIN_KEY not in boundary_gdf.columns:

    st.error(
        f"GeoJSON에 JOIN KEY "
        f"`{config.JOIN_KEY}` 컬럼이 없습니다.\n\n"
        "현재 GeoJSON columns:\n"
        f"{list(boundary_gdf.columns)}"
    )

    st.stop()


if config.JOIN_KEY not in district_df.columns:

    st.error(
        f"District data에 JOIN KEY "
        f"`{config.JOIN_KEY}` 컬럼이 없습니다.\n\n"
        "현재 District data columns:\n"
        f"{list(district_df.columns)}"
    )

    st.stop()


# ============================================================
# JOIN
# ============================================================

merged = map_utils.join_district_data(
    boundary_gdf,
    district_df,
)


# ============================================================
# JOIN QUALITY CHECK
# ============================================================

boundary_ids = set(
    boundary_gdf[
        config.JOIN_KEY
    ].astype(str)
)


district_ids = set(
    district_df[
        config.JOIN_KEY
    ].astype(str)
)


missing_in_excel = (
    boundary_ids - district_ids
)


missing_in_boundary = (
    district_ids - boundary_ids
)


if (
    missing_in_excel
    or missing_in_boundary
):

    with st.expander(
        "🔍 Join 점검 (District_ID 불일치)",
        expanded=False,
    ):

        if missing_in_excel:

            st.write(
                "GeoJSON에는 있지만 District data에 "
                f"없는 District_ID "
                f"({len(missing_in_excel)}개):"
            )

            st.code(
                ", ".join(
                    sorted(
                        missing_in_excel
                    )
                )
            )

        if missing_in_boundary:

            st.write(
                "District data에는 있지만 GeoJSON에 "
                f"없는 District_ID "
                f"({len(missing_in_boundary)}개):"
            )

            st.code(
                ", ".join(
                    sorted(
                        missing_in_boundary
                    )
                )
            )


# ============================================================
# LOAD CORRIDOR / MARKERS
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


extra_layers = []


if corridor_df is not None:

    extra_layers.append(
        map_utils.build_corridor_layer(
            corridor_df
        )
    )


if markers_df is not None:

    point_layer, text_layer = (
        map_utils.build_marker_layers(
            markers_df
        )
    )

    extra_layers.extend(
        [
            point_layer,
            text_layer,
        ]
    )


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
# FORMAT FUNCTIONS
# ============================================================

def _fmt_number(value):

    if value is None or pd.isna(value):

        return "No data"

    try:

        return f"{int(value):,}"

    except (
        ValueError,
        TypeError,
    ):

        return str(value)


def _fmt_currency(value):

    if value is None or pd.isna(value):

        return "No data"

    try:

        return f"USD {int(value):,}"

    except (
        ValueError,
        TypeError,
    ):

        return str(value)


def _fmt_text(value):

    if (
        value is None
        or pd.isna(value)
        or str(value).strip() == ""
    ):

        return "_(not recorded)_"

    return str(value)


# ============================================================
# SELECTED DISTRICT
# ============================================================

selected_props = None


try:

    objects = (
        event.selection
        .get("objects", {})
        .get(
            "district-layer",
            [],
        )
    )

    if objects:

        selected_props = objects[0]

except Exception:

    selected_props = None


# ============================================================
# DETAIL PANEL
# ============================================================

with detail_col:

    st.subheader(
        "📋 District detail"
    )

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
            f"### {selected_props.get(
                'District',
                'Unknown'
            )}"
        )

        st.markdown(
            f"""
            <span style='
                background:{severity_hex};
                color:white;
                padding:2px 10px;
                border-radius:10px;
                font-weight:600;
            '>
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

        st.markdown(
            "**Problem / Needs**"
        )

        st.write(
            _fmt_text(
                selected_props.get(
                    "Key_problems"
                )
            )
        )

        st.markdown(
            "**Severity reason**"
        )

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
            "Data source: "
            f"{selected_props.get(
                'Data_source',
                '-'
            )}"
        )

        c4.caption(
            "Last updated: "
            f"{selected_props.get(
                'Last_updated',
                '-'
            )}"
        )

        st.caption(
            "Data status: "
            f"{selected_props.get(
                'Data_status',
                '-'
            )}"
        )


# ============================================================
# FOOTER
# ============================================================

st.markdown("---")

st.caption(
    "UNICEF Nepal | WASH Emergency Response | "
    f"Population: {POPULATION_PATH.name} | "
    f"Flood layer: {FLOOD_GEOJSON_PATH.name} | "
    f"Workbook last modified: "
    f"{data_loader.workbook_last_updated(excel_path)}"
)
```

그리고 **가장 중요한 것은 `requirements.txt`도 반드시 같이 수정하는 것**입니다. 현재 오류의 직접적인 원인인 `fiona`를 제거해야 합니다.

```txt
streamlit
pandas
geopandas
pydeck
openpyxl
shapely
pyogrio
```

즉, 기존의

```txt
fiona==1.10.1
```

은 **삭제**하세요.

이번 오류의 핵심은 `GeoPandas → Fiona → GDAL` 과정에서 Streamlit Cloud가 `gdal-config`를 찾지 못한 것입니다. 위와 같이 `GeoPandas → Pyogrio → GDAL` 구조로 변경하면 Fiona을 별도로 빌드할 필요가 없어서 현재 발생한 `A GDAL API version must be specified` 오류를 피할 수 있습니다.
