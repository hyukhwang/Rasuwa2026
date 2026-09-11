# -*- coding: utf-8 -*-
"""
Excel(Source of Truth) 읽기 전담 모듈.
Streamlit 캐시를 사용하되, Excel의 '수정 시각(mtime)'을 캐시 키에 포함시켜
파일이 업데이트되면 자동으로 새로 읽도록 합니다.
"""

import os
import pandas as pd
import streamlit as st

import config


def _file_signature(path):
    """캐시 무효화를 위한 파일 signature (경로+수정시각)."""
    try:
        return (str(path), os.path.getmtime(path))
    except OSError:
        return (str(path), None)


@st.cache_data(show_spinner=False)
def _read_sheet(path, sheet_name, signature):
    """signature가 바뀔 때만 실제로 다시 읽음 (Streamlit 캐시 트릭)."""
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")


def _safe_read(path, sheet_name):
    if not os.path.exists(path):
        return pd.DataFrame(), f"파일을 찾을 수 없습니다: {path}"
    try:
        df = _read_sheet(path, sheet_name, _file_signature(path))
    except ValueError as e:
        return pd.DataFrame(), f"'{sheet_name}' 시트를 찾을 수 없습니다: {e}"
    except Exception as e:
        return pd.DataFrame(), f"'{sheet_name}' 시트를 읽는 중 오류: {e}"

    # District_ID / Facility_ID 는 항상 문자열로 (엑셀이 숫자로 인식하는 것 방지)
    for id_col in ("District_ID", "Facility_ID"):
        if id_col in df.columns:
            df[id_col] = df[id_col].astype(str).str.strip()

    # 완전히 빈 행 제거 (District_ID/Facility_ID 기준)
    id_col = "District_ID" if "District_ID" in df.columns else (
        "Facility_ID" if "Facility_ID" in df.columns else None
    )
    if id_col:
        df = df[df[id_col].notna() & (df[id_col].str.lower() != "nan")]

    return df, None


@st.cache_data(show_spinner=False)
def _read_csv(path, signature):
    return pd.read_csv(path)


def _safe_read_csv(path):
    if not os.path.exists(path):
        return pd.DataFrame(), f"파일을 찾을 수 없습니다: {path}"
    try:
        df = _read_csv(path, _file_signature(path))
    except Exception as e:
        return pd.DataFrame(), f"CSV를 읽는 중 오류: {e}"

    for id_col in ("District_ID", "Facility_ID"):
        if id_col in df.columns:
            df[id_col] = df[id_col].astype(str).str.strip()

    id_col = "District_ID" if "District_ID" in df.columns else (
        "Facility_ID" if "Facility_ID" in df.columns else None
    )
    if id_col:
        df = df[df[id_col].notna() & (df[id_col].astype(str).str.lower() != "nan")]

    return df, None


def load_district_situation(path=None, csv_path=None):
    """District_Situation 데이터를 읽는다.

    운영 방식: 팀이 꾸준히 업데이트하는 CSV(district_situation.csv)가 있으면
    그것을 우선 사용하고, 없으면 Excel 워크북의 District_Situation 시트를 읽는다.
    (Water_Access / Interventions / Funding / Facilities는 계속 Excel 기준.)
    """
    csv_path = csv_path or config.DEFAULT_DISTRICT_CSV_PATH
    if os.path.exists(csv_path):
        return _safe_read_csv(csv_path)

    path = path or config.DEFAULT_EXCEL_PATH
    return _safe_read(path, config.SHEET_DISTRICT_SITUATION)


def load_event_info(path=None):
    """지도 상단 타이틀/서브타이틀 등 이벤트 메타정보 (Key,Value 2컬럼 CSV)."""
    path = path or config.DEFAULT_EVENT_INFO_PATH
    df, err = _safe_read_csv_no_id_filter(path)
    if err or df.empty:
        return {}, err
    return dict(zip(df["Key"], df["Value"])), None


def _safe_read_csv_no_id_filter(path):
    if not os.path.exists(path):
        return pd.DataFrame(), f"파일을 찾을 수 없습니다: {path}"
    try:
        df = _read_csv(path, _file_signature(path))
    except Exception as e:
        return pd.DataFrame(), f"CSV를 읽는 중 오류: {e}"
    return df, None


def load_event_markers(path=None):
    """Epicentre 등 주석용 point marker 목록."""
    path = path or config.DEFAULT_EVENT_MARKERS_PATH
    return _safe_read_csv_no_id_filter(path)


def load_flood_corridor(path=None):
    """Flood corridor(하천 경로) polyline 좌표 목록 (Sequence 순서대로)."""
    path = path or config.DEFAULT_FLOOD_CORRIDOR_PATH
    df, err = _safe_read_csv_no_id_filter(path)
    if not err and "Sequence" in df.columns:
        df = df.sort_values("Sequence")
    return df, err


def load_water_access(path=None):
    path = path or config.DEFAULT_EXCEL_PATH
    return _safe_read(path, config.SHEET_WATER_ACCESS)


def load_interventions(path=None):
    path = path or config.DEFAULT_EXCEL_PATH
    return _safe_read(path, config.SHEET_INTERVENTIONS)


def load_funding(path=None):
    path = path or config.DEFAULT_EXCEL_PATH
    return _safe_read(path, config.SHEET_FUNDING)


def load_facilities(path=None):
    path = path or config.DEFAULT_EXCEL_PATH
    return _safe_read(path, config.SHEET_FACILITIES)


def workbook_last_updated(path=None):
    """워크북 파일 자체의 최종 수정 시각 (파일 시스템 기준, 표시용)."""
    path = path or config.DEFAULT_EXCEL_PATH
    try:
        ts = os.path.getmtime(path)
        import datetime
        return datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
    except OSError:
        return "Unknown"
