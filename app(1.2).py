# -*- coding: utf-8 -*-
"""
map_builder.py

PyDeck layer / view state 생성을 담당합니다.

주의:
- Streamlit Cloud에서 GeoPandas가 설치되지 않아도 import 단계에서
  앱이 중단되지 않도록 geopandas 의존성을 제거했습니다.
- build_view_state()는 GeoDataFrame 자체를 요구하지 않고,
  total_bounds를 사용할 수 있는 객체를 받도록 처리합니다.
"""

import pydeck as pdk

import config


def build_district_layer(features: list) -> pdk.Layer:
    """
    Priority에 따라 색상이 달라지는 District polygon layer를 생성합니다.

    Priority:
        High   -> Red
        Medium -> Orange
        Low    -> Yellow
        기타   -> Gray
    """

    # features가 None인 경우 빈 FeatureCollection으로 처리
    if features is None:
        features = []

    return pdk.Layer(
        "GeoJsonLayer",

        {
            "type": "FeatureCollection",
            "features": features,
        },

        pickable=True,
        auto_highlight=True,

        stroked=True,
        filled=True,

        # Priority에 따른 색상
        get_fill_color="""
            properties.Priority == 'High'
                ? [214, 39, 40, 190]
                : properties.Priority == 'Medium'
                    ? [255, 165, 0, 180]
                    : properties.Priority == 'Low'
                        ? [255, 221, 87, 170]
                        : [200, 200, 200, 120]
        """,

        get_line_color=[60, 60, 60, 200],
        line_width_min_pixels=1,

        highlight_color=[0, 120, 255, 120],

        id="district-layer",
    )


def build_holding_centre_layer(df) -> pdk.Layer:
    """
    Holding_Centres 시트에서 Latitude / Longitude가 존재하는
    행만 point layer로 지도에 표시합니다.

    좌표가 없거나 필요한 컬럼이 없는 경우에는
    빈 layer를 반환하여 앱 전체가 중단되지 않도록 합니다.
    """

    records = []

    # DataFrame이 없는 경우
    if df is None:
        return pdk.Layer(
            "ScatterplotLayer",
            [],
            get_position=["Longitude", "Latitude"],
            get_fill_color=[128, 128, 128, 150],
            get_radius=250,
            radius_min_pixels=5,
            radius_max_pixels=20,
            pickable=True,
            stroked=True,
            get_line_color=[255, 255, 255, 220],
            line_width_min_pixels=1,
            id="holding-centre-layer",
        )

    # DataFrame이 비어 있는 경우
    try:
        if df.empty:
            return pdk.Layer(
                "ScatterplotLayer",
                [],
                get_position=["Longitude", "Latitude"],
                get_fill_color=[128, 128, 128, 150],
                get_radius=250,
                radius_min_pixels=5,
                radius_max_pixels=20,
                pickable=True,
                stroked=True,
                get_line_color=[255, 255, 255, 220],
                line_width_min_pixels=1,
                id="holding-centre-layer",
            )
    except Exception:
        pass

    # 필요한 컬럼 확인
    required_columns = [
        "Latitude",
        "Longitude",
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    # 좌표 컬럼이 없으면 빈 layer 반환
    if missing_columns:
        return pdk.Layer(
            "ScatterplotLayer",
            [],
            get_position=["Longitude", "Latitude"],
            get_fill_color=[128, 128, 128, 150],
            get_radius=250,
            radius_min_pixels=5,
            radius_max_pixels=20,
            pickable=True,
            stroked=True,
            get_line_color=[255, 255, 255, 220],
            line_width_min_pixels=1,
            id="holding-centre-layer",
        )

    try:
        # 좌표가 있는 행만 선택
        plot_df = df.dropna(
            subset=["Latitude", "Longitude"]
        ).copy()

        # Status 컬럼이 존재하는 경우 색상 매핑
        if "Status" in plot_df.columns:

            plot_df["radius_color"] = (
                plot_df["Status"]
                .map(config.HOLDING_CENTRE_COLORS)
                .apply(
                    lambda color:
                    color
                    if isinstance(color, list)
                    else config.DEFAULT_HOLDING_CENTRE_COLOR
                )
            )

        else:
            # Status 컬럼이 없는 경우 기본 색상 사용
            plot_df["radius_color"] = [
                config.DEFAULT_HOLDING_CENTRE_COLOR
            ] * len(plot_df)

        # 숫자로 변환
        plot_df["Latitude"] = (
            plot_df["Latitude"]
            .astype(float)
        )

        plot_df["Longitude"] = (
            plot_df["Longitude"]
            .astype(float)
        )

        records = plot_df.to_dict(
            orient="records"
        )

    except Exception:
        # 데이터 처리 중 문제가 발생하더라도
        # 앱 전체가 종료되지 않도록 빈 layer 반환
        records = []

    return pdk.Layer(
        "ScatterplotLayer",
        records,

        get_position=[
            "Longitude",
            "Latitude",
        ],

        get_fill_color="radius_color",

        get_radius=250,

        radius_min_pixels=5,
        radius_max_pixels=20,

        pickable=True,
        stroked=True,

        get_line_color=[
            255,
            255,
            255,
            220,
        ],

        line_width_min_pixels=1,

        id="holding-centre-layer",
    )


def build_view_state(gdf) -> pdk.ViewState:
    """
    Boundary 전체 bounding box를 기준으로
    초기 지도 중심을 설정합니다.

    GeoPandas의 타입 annotation을 제거하여
    GeoPandas가 설치되지 않은 환경에서도
    모듈 import가 가능하도록 했습니다.
    """

    # 데이터가 없는 경우 기본 위치 사용
    if gdf is None:
        return pdk.ViewState(
            latitude=config.DEFAULT_LAT,
            longitude=config.DEFAULT_LON,
            zoom=config.DEFAULT_ZOOM,
            pitch=0,
        )

    try:
        # 빈 GeoDataFrame 처리
        if gdf.empty:
            return pdk.ViewState(
                latitude=config.DEFAULT_LAT,
                longitude=config.DEFAULT_LON,
                zoom=config.DEFAULT_ZOOM,
                pitch=0,
            )
    except Exception:
        return pdk.ViewState(
            latitude=config.DEFAULT_LAT,
            longitude=config.DEFAULT_LON,
            zoom=config.DEFAULT_ZOOM,
            pitch=0,
        )

    try:
        # 전체 bounding box
        minx, miny, maxx, maxy = gdf.total_bounds

        # 좌표값이 정상적인지 확인
        if any(
            value is None
            for value in [minx, miny, maxx, maxy]
        ):
            raise ValueError(
                "Invalid bounding box."
            )

        latitude = (
            float(miny) + float(maxy)
        ) / 2

        longitude = (
            float(minx) + float(maxx)
        ) / 2

        return pdk.ViewState(
            latitude=latitude,
            longitude=longitude,
            zoom=config.DEFAULT_ZOOM,
            pitch=0,
        )

    except Exception:
        # bounding box를 계산할 수 없는 경우
        # 기본 Nepal 위치 사용
        return pdk.ViewState(
            latitude=config.DEFAULT_LAT,
            longitude=config.DEFAULT_LON,
            zoom=config.DEFAULT_ZOOM,
            pitch=0,
        )


def build_deck(
    features: list,
    gdf,
    extra_layers: list = None,
) -> pdk.Deck:
    """
    District polygon layer와 추가 layer를 결합하여
    최종 PyDeck map을 생성합니다.
    """

    # 기본 District layer
    layers = [
        build_district_layer(features)
    ]

    # 추가 layer가 있는 경우 추가
    if extra_layers:

        for layer in extra_layers:

            if layer is not None:
                layers.append(layer)

    # 초기 View 설정
    view_state = build_view_state(gdf)

    # Tooltip
    tooltip = {
        "html": """
            <b>{District}</b>
        """,

        "style": {
            "backgroundColor": "#1F4E78",
            "color": "white",
        },
    }

    # 최종 Deck 생성
    return pdk.Deck(
        layers=layers,

        initial_view_state=view_state,

        tooltip=tooltip,

        map_style="light",
    )