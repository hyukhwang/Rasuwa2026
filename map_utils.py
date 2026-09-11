# -*- coding: utf-8 -*-
"""
District boundary(GeoJSON/Shapefile) + Excel District_Situation 을
District_ID 기준으로 Join 하고, PyDeck 레이어 / ViewState 를 만드는 모듈.
"""

import os
import json

import geopandas as gpd
import pandas as pd
import pydeck as pdk

import config


def load_district_boundary(geojson_path=None):
    """District boundary 파일을 GeoDataFrame으로 로드.

    Shapefile(.shp)이나 GeoJSON(.geojson/.json) 모두 지원합니다.
    boundary 파일에는 최소한 District_ID 컬럼(혹은 그에 준하는 고유 ID)이
    있어야 Excel과 Join할 수 있습니다.
    """
    geojson_path = geojson_path or config.DEFAULT_GEOJSON_PATH
    if not os.path.exists(geojson_path):
        return None, f"경계 파일을 찾을 수 없습니다: {geojson_path}"
    try:
        gdf = gpd.read_file(geojson_path)
    except Exception as e:
        return None, f"경계 파일을 읽는 중 오류: {e}"

    if config.JOIN_KEY not in gdf.columns:
        return None, (
            f"경계 파일에 '{config.JOIN_KEY}' 컬럼이 없습니다. "
            f"현재 컬럼: {list(gdf.columns)}\n"
            "boundary 파일에 District_ID 컬럼을 추가하거나, "
            "이름 매핑 테이블로 먼저 채워주세요."
        )

    gdf[config.JOIN_KEY] = gdf[config.JOIN_KEY].astype(str).str.strip()
    if gdf.crs is not None and str(gdf.crs) != "EPSG:4326":
        gdf = gdf.to_crs(epsg=4326)
    return gdf, None


def join_district_data(boundary_gdf, district_df):
    """boundary(GeoDataFrame) + District_Situation(DataFrame)을 District_ID로 Join.

    Left join(경계 기준)이므로, Excel에 데이터가 아직 없는 District도
    'No Data' 색상으로 지도에는 항상 표시됩니다.
    """
    merged = boundary_gdf.merge(
        district_df,
        on=config.JOIN_KEY,
        how="left",
        suffixes=("_boundary", ""),
    )

    if "Severity" not in merged.columns:
        # 하위 호환: CSV가 없어 Excel의 District_Situation(Priority 컬럼)만
        # 읽은 경우에도 지도가 동작하도록 Priority를 Severity로 사용.
        merged["Severity"] = merged["Priority"] if "Priority" in merged.columns else "No Data"
    merged["Severity"] = merged["Severity"].fillna("No Data")
    merged.loc[~merged["Severity"].isin(config.SEVERITY_COLORS.keys()), "Severity"] = "No Data"

    merged["fill_color"] = merged["Severity"].map(config.SEVERITY_COLORS)

    for numeric_col in ("Affected_population", "Water_access_population", "Estimated_funding_USD"):
        if numeric_col not in merged.columns:
            merged[numeric_col] = None

    return merged


def geodataframe_to_pydeck_features(merged_gdf):
    """GeoDataFrame -> PyDeck GeoJsonLayer가 바로 먹을 수 있는 dict(FeatureCollection)."""
    geojson_str = merged_gdf.to_json()
    return json.loads(geojson_str)


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
    """flood_corridor_path.csv(Sequence, Latitude, Longitude)로 PathLayer 생성."""
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
        get_color=config.CORRIDOR_COLOR,
        get_width=config.CORRIDOR_WIDTH_M,
        width_min_pixels=2,
        pickable=False,
    )


def build_marker_layers(markers_df):
    """event_markers.csv(Marker_ID, Label, Type, Latitude, Longitude, Note)로
    포인트 + 라벨 레이어 생성. Type 값에 따라 색상이 달라짐(예: epicentre)."""
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
    df["marker_color"] = marker_type.map(config.MARKER_COLORS).apply(
        lambda c: c if isinstance(c, list) else config.DEFAULT_MARKER_COLOR
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
        return pdk.ViewState(latitude=lat, longitude=lon, zoom=config.DEFAULT_ZOOM, pitch=0)
    except Exception:
        return pdk.ViewState(
            latitude=config.DEFAULT_LAT,
            longitude=config.DEFAULT_LON,
            zoom=config.DEFAULT_ZOOM,
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
