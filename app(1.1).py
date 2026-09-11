# -*- coding: utf-8 -*-
"""
map_builder.py
pydeck layer / view state 생성을 담당합니다.
"""
import pydeck as pdk
import geopandas as gpd

import config


def build_district_layer(features: list) -> pdk.Layer:
    """
    Priority에 따라 색이 달라지는 District polygon layer.
    get_fill_color 에 JS 표현식을 사용해, 데이터가 바뀌어도 Python 재실행 없이 색이 매핑되게 함.
    """
    return pdk.Layer(
        "GeoJsonLayer",
        {"type": "FeatureCollection", "features": features},
        pickable=True,
        auto_highlight=True,
        stroked=True,
        filled=True,
        get_fill_color="""
            properties.Priority == 'High'   ? [214, 39, 40, 190] :
            properties.Priority == 'Medium' ? [255, 165, 0, 180] :
            properties.Priority == 'Low'    ? [255, 221, 87, 170] :
                                               [200, 200, 200, 120]
        """,
        get_line_color=[60, 60, 60, 200],
        line_width_min_pixels=1,
        highlight_color=[0, 120, 255, 120],
        id="district-layer",
    )


def build_holding_centre_layer(df) -> pdk.Layer:
    """
    Holding_Centres 시트 중 Latitude/Longitude가 채워진 행만 point layer로 표시.
    (좌표가 없는 행은 지도에 못 그리므로 app.py에서 별도로 목록/경고를 보여줌)
    """
    import pandas as pd  # local import to keep module import-light

    if df is None or df.empty:
        records = []
    else:
        plot_df = df.dropna(subset=["Latitude", "Longitude"]).copy()
        plot_df["radius_color"] = plot_df["Status"].map(config.HOLDING_CENTRE_COLORS).apply(
            lambda c: c if isinstance(c, list) else config.DEFAULT_HOLDING_CENTRE_COLOR
        )
        records = plot_df.to_dict("records")

    return pdk.Layer(
        "ScatterplotLayer",
        records,
        get_position=["Longitude", "Latitude"],
        get_fill_color="radius_color",
        get_radius=250,
        radius_min_pixels=5,
        radius_max_pixels=20,
        pickable=True,
        stroked=True,
        get_line_color=[255, 255, 255, 220],
        line_width_min_pixels=1,
        id="holding-centre-layer",
    )


def build_view_state(gdf: gpd.GeoDataFrame) -> pdk.ViewState:
    """boundary 전체 bounding box 중심으로 초기 view를 잡습니다."""
    if gdf.empty:
        return pdk.ViewState(
            latitude=config.DEFAULT_LAT, longitude=config.DEFAULT_LON, zoom=config.DEFAULT_ZOOM
        )
    minx, miny, maxx, maxy = gdf.total_bounds
    return pdk.ViewState(
        latitude=(miny + maxy) / 2,
        longitude=(minx + maxx) / 2,
        zoom=config.DEFAULT_ZOOM,
        pitch=0,
    )


def build_deck(features: list, gdf: gpd.GeoDataFrame, extra_layers: list = None) -> pdk.Deck:
    layers = [build_district_layer(features)]
    if extra_layers:
        layers.extend(extra_layers)
    view_state = build_view_state(gdf)
    # 레이어마다 필드 구조가 다르므로(District polygon vs Holding centre point),
    # 두 레이어 모두에 존재하는 필드만 tooltip에 사용합니다.
    # 상세 정보는 오른쪽 패널(클릭)과 하단 Holding Centre 목록에서 확인합니다.
    tooltip = {
        "html": "<b>{District}</b>",
        "style": {"backgroundColor": "#1F4E78", "color": "white"},
    }
    return pdk.Deck(
        layers=layers,
        initial_view_state=view_state,
        tooltip=tooltip,
        map_style="light",
    )