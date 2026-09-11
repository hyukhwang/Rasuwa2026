# -*- coding: utf-8 -*-
"""
전역 설정.
⚠️ 여기에는 '설정값'만 둡니다 (경로, 색상, 컬럼명, 기본 지도 위치).
District 별 실제 데이터는 절대 이 파일에 넣지 않습니다 — 전부 Excel에서 읽습니다.
"""

from pathlib import Path

# ------------------------------------------------------------------
# 경로 (Streamlit Cloud에서도 상대경로로 동작하도록 프로젝트 루트 기준)
# ------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent

DEFAULT_EXCEL_PATH = BASE_DIR / "data" / "WASH_data_template.xlsx"

# District_Situation을 CSV로 꾸준히 업데이트하는 운영 방식을 위한 경로.
# 존재하면 이 CSV가 Excel의 District_Situation 시트보다 우선합니다.
# (Water_Access / Interventions / Funding / Facilities는 계속 Excel에서 읽습니다.)
DEFAULT_DISTRICT_CSV_PATH = BASE_DIR / "data" / "district_situation.csv"

# 실제 District 경계 파일이 준비되면 이 경로에 두고 파일명을 바꿔주세요.
# (지금은 테스트용 PLACEHOLDER 사각형 geometry 입니다 — README 참고)
DEFAULT_GEOJSON_PATH = BASE_DIR / "data" / "sample_districts_PLACEHOLDER.geojson"

# 지도 상단 타이틀/서브타이틀 등 이벤트 메타정보 (하드코딩 대신 CSV에서 읽음)
DEFAULT_EVENT_INFO_PATH = BASE_DIR / "data" / "event_info.csv"

# Epicentre, "not flooded" 같은 주석용 point marker
DEFAULT_EVENT_MARKERS_PATH = BASE_DIR / "data" / "event_markers.csv"

# Flood corridor(하천 경로) 선을 그리기 위한 순서 있는 좌표 목록
DEFAULT_FLOOD_CORRIDOR_PATH = BASE_DIR / "data" / "flood_corridor_path.csv"

# ------------------------------------------------------------------
# Join 기준 키
# ------------------------------------------------------------------
JOIN_KEY = "District_ID"

# ------------------------------------------------------------------
# 지도 기본 위치 (Kathmandu)
# ------------------------------------------------------------------
DEFAULT_LAT = 28.15
DEFAULT_LON = 84.80
DEFAULT_ZOOM = 6.6

# ------------------------------------------------------------------
# Severity(피해 심각도)별 색상 (RGBA) — 참고 이미지의 4단계 + No Data 체계
# District_Situation의 'Severity' 컬럼 값과 정확히 일치해야 합니다.
# ------------------------------------------------------------------
SEVERITY_COLORS = {
    "Severe":    [122, 15, 15, 220],    # 진한 적갈색 — epicentre 급
    "Very High": [214, 39, 40, 200],    # 빨강
    "High":      [244, 145, 45, 190],   # 주황
    "Moderate":  [255, 221, 87, 180],   # 노랑
    "No Data":   [200, 197, 189, 110],  # 회색 (배경톤)
}
DEFAULT_SEVERITY_COLOR = SEVERITY_COLORS["No Data"]

SEVERITY_LEGEND = [
    ("Severe", "#7A0F0F"),
    ("Very High", "#D62728"),
    ("High", "#F4912D"),
    ("Moderate", "#FFDD57"),
    ("No Data", "#C8C5BD"),
]

SEVERITY_ORDER = ["Severe", "Very High", "High", "Moderate", "No Data"]

# 예전 버전과의 호환용 별칭 (Priority 기준 코드를 참조하는 곳이 있다면 유지)
PRIORITY_COLORS = SEVERITY_COLORS
DEFAULT_PRIORITY_COLOR = DEFAULT_SEVERITY_COLOR
PRIORITY_LEGEND = SEVERITY_LEGEND

# ------------------------------------------------------------------
# Facility 색상 (Phase 3에서 사용)
# ------------------------------------------------------------------
FACILITY_COLORS = {
    "School": [31, 119, 180, 220],
    "Health Facility": [44, 160, 44, 220],
}
DEFAULT_FACILITY_COLOR = [128, 128, 128, 220]

# ------------------------------------------------------------------
# Sheet 이름 (Excel 시트명 — 여기서 바꾸면 코드 전체에 반영됨)
# ------------------------------------------------------------------
# ------------------------------------------------------------------
# 지도 상단 corridor / marker 스타일
# ------------------------------------------------------------------
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
