# -*- coding: utf-8 -*-
"""
UNICEF Nepal — Flood WASH Emergency Decision Support Dashboard
Phase 1 MVP

- Excel(Source of Truth) 읽기
- District boundary(GeoJSON) 읽기
- District_ID 기준 Join
- Priority 색상으로 District polygon 표시
- District 클릭 -> 상세정보 패널 표시
- Last updated 표시

실행: streamlit run app.py
"""

import pandas as pd
import streamlit as st

import config
import data_loader
import map_utils

st.set_page_config(
    page_title="UNICEF Nepal — Flood WASH Dashboard",
    page_icon="🌊",
    layout="wide",
)

# ============================================================
# SIDEBAR — 데이터 소스 설정
# ============================================================
st.sidebar.title("⚙️ Data source")

excel_path = st.sidebar.text_input(
    "Excel 파일 경로",
    value=str(config.DEFAULT_EXCEL_PATH),
    help="WASH_data_template.xlsx 의 실제 경로. 팀 공유 경로로 바꿔도 됩니다.",
)
geojson_path = st.sidebar.text_input(
    "District boundary 파일 경로 (GeoJSON)",
    value=str(config.DEFAULT_GEOJSON_PATH),
    help="⚠️ 기본값은 테스트용 PLACEHOLDER 경계입니다. 실제 Nepal District "
         "경계 파일로 교체하세요 (README.md 의 출처 링크 참고).",
)

with st.sidebar.expander("고급: 이벤트 메타 / corridor / marker 경로", expanded=False):
    event_info_path = st.text_input("이벤트 정보 CSV (제목/부제목)", value=str(config.DEFAULT_EVENT_INFO_PATH))
    markers_path = st.text_input("주석 marker CSV (epicentre 등)", value=str(config.DEFAULT_EVENT_MARKERS_PATH))
    corridor_path = st.text_input("Flood corridor 경로 CSV", value=str(config.DEFAULT_FLOOD_CORRIDOR_PATH))

if st.sidebar.button("🔄 데이터 새로고침"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption(f"Workbook last modified: **{data_loader.workbook_last_updated(excel_path)}**")

if str(geojson_path) == str(config.DEFAULT_GEOJSON_PATH):
    st.sidebar.warning(
        "현재 데모용 PLACEHOLDER District 경계를 사용 중입니다. "
        "실제 서비스 전에 진짜 경계 파일로 교체하세요.",
        icon="⚠️",
    )

# ============================================================
# TITLE  (event_info.csv 기반 — 하드코딩 없음)
# ============================================================
event_info, event_info_err = data_loader.load_event_info(event_info_path)

st.markdown(
    f"<h1 style='margin-bottom:0;'>{event_info.get('Title', 'UNICEF Nepal — Flood WASH Dashboard')}</h1>",
    unsafe_allow_html=True,
)
if event_info.get("Subtitle"):
    st.markdown(
        f"<h3 style='color:#B22222;margin-top:2px;'>{event_info['Subtitle']}</h3>",
        unsafe_allow_html=True,
    )
meta_bits = [v for v in [event_info.get("Corridor_note"), event_info.get("Data_asof")] if v]
if meta_bits:
    st.caption("  ·  ".join(meta_bits))
st.caption("District를 클릭하면 상세 flood/WASH 상황이 표시됩니다.")

with st.expander("ℹ️ 지도 정보 / 범례 (Impact level)", expanded=True):
    legend_cols = st.columns(len(config.SEVERITY_LEGEND))
    for col, (label, hexcolor) in zip(legend_cols, config.SEVERITY_LEGEND):
        col.markdown(
            f"<span style='display:inline-block;width:14px;height:14px;"
            f"background:{hexcolor};border-radius:3px;margin-right:6px;'></span>{label}",
            unsafe_allow_html=True,
        )
    st.write(
        "District 색상은 `district_situation.csv`(없으면 Excel의 "
        "`District_Situation` 시트)의 `Severity` 컬럼을 따릅니다. "
        "CSV 값을 바꾸고 새로고침하면 지도에 바로 반영됩니다."
    )

# ============================================================
# LOAD DATA
# ============================================================
district_df, district_err = data_loader.load_district_situation(excel_path)
boundary_gdf, boundary_err = map_utils.load_district_boundary(geojson_path)

if district_err:
    st.error(f"District_Situation 시트를 불러오지 못했습니다: {district_err}")
    st.stop()

if boundary_err:
    st.error(f"District boundary를 불러오지 못했습니다: {boundary_err}")
    st.stop()

merged = map_utils.join_district_data(boundary_gdf, district_df)

# Join 품질 체크: 경계에는 있는데 Excel에 없는 District_ID (혹은 반대)
boundary_ids = set(boundary_gdf[config.JOIN_KEY])
excel_ids = set(district_df[config.JOIN_KEY]) if config.JOIN_KEY in district_df.columns else set()
missing_in_excel = boundary_ids - excel_ids
missing_in_boundary = excel_ids - boundary_ids

if missing_in_excel or missing_in_boundary:
    with st.expander("🔍 Join 점검 (District_ID 불일치)", expanded=False):
        if missing_in_excel:
            st.write(f"경계 파일에는 있지만 Excel에 데이터가 없는 District_ID ({len(missing_in_excel)}개):")
            st.code(", ".join(sorted(missing_in_excel)))
        if missing_in_boundary:
            st.write(f"Excel에는 있지만 경계 파일에 없는 District_ID ({len(missing_in_boundary)}개):")
            st.code(", ".join(sorted(missing_in_boundary)))

# ============================================================
# MAP
# ============================================================
corridor_df, corridor_err = data_loader.load_flood_corridor(corridor_path)
markers_df, markers_err = data_loader.load_event_markers(markers_path)

extra_layers = [map_utils.build_corridor_layer(corridor_df)]
point_layer, text_layer = map_utils.build_marker_layers(markers_df)
extra_layers += [point_layer, text_layer]

map_col, detail_col = st.columns([2, 1])

with map_col:
    deck = map_utils.build_deck(merged, extra_layers=extra_layers)
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
    if value is None or pd.isna(value) or str(value).strip() == "":
        return "_(not recorded)_"
    return str(value)


selected_props = None
try:
    objects = event.selection["objects"].get("district-layer", [])
    if objects:
        selected_props = objects[0]
except Exception:
    selected_props = None

with detail_col:
    st.subheader("📋 District detail")

    if not selected_props:
        st.info("지도에서 District를 클릭하면 상세 정보가 여기에 표시됩니다.")
    else:
        severity = selected_props.get("Severity", "No Data")
        severity_hex = dict(config.SEVERITY_LEGEND).get(severity, "#BEBEBE")

        st.markdown(f"### {selected_props.get('District', 'Unknown')}")
        st.markdown(
            f"<span style='background:{severity_hex};color:white;padding:2px 10px;"
            f"border-radius:10px;font-weight:600;'>{str(severity).upper()}</span>",
            unsafe_allow_html=True,
        )

        c1, c2 = st.columns(2)
        c1.metric("Affected population", _fmt_number(selected_props.get("Affected_population")))
        c2.metric("Water access population", _fmt_number(selected_props.get("Water_access_population")))
        c3, c4 = st.columns(2)
        c3.metric("Casualties", _fmt_number(selected_props.get("Casualties")))
        c4.metric("Estimated funding requirement", _fmt_currency(selected_props.get("Estimated_funding_USD")))

        st.markdown("**Current water source:** " + _fmt_text(selected_props.get("Water_source_current")))

        st.markdown("---")
        st.markdown("**Problem / Needs**")
        st.write(_fmt_text(selected_props.get("Key_problems")))

        st.markdown("**Severity reason**")
        st.write(_fmt_text(selected_props.get("Severity_reason", selected_props.get("Priority_reason"))))

        st.markdown("---")
        st.markdown("**Immediate intervention**")
        st.write(_fmt_text(selected_props.get("Immediate_intervention")))
        st.markdown("**Medium-term intervention**")
        st.write(_fmt_text(selected_props.get("Medium_term_intervention")))
        st.markdown("**Long-term intervention**")
        st.write(_fmt_text(selected_props.get("Long_term_intervention")))

        st.markdown("---")
        c3, c4 = st.columns(2)
        c3.caption(f"Data source: {selected_props.get('Data_source', '-')}")
        c4.caption(f"Last updated: {selected_props.get('Last_updated', '-')}")
        st.caption(f"Data status: {selected_props.get('Data_status', '-')}")

# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.caption(
    "UNICEF Nepal | WASH Emergency Response | "
    f"Source: {excel_path} | Workbook last modified: {data_loader.workbook_last_updated(excel_path)}"
)
