[README.md](https://github.com/user-attachments/files/32087065/README.md)
# UNICEF Nepal — Flood WASH Situation Dashboard

Excel을 source of truth로 사용하는 Streamlit 기반 interactive WASH 대시보드 (Phase 1 MVP).

## 1. 폴더 구조

```
unicef_wash_dashboard/
├── app.py                        # Streamlit 진입점 (Phase 1 MVP)
├── config.py                     # 경로 / 컬럼명 / 색상 등 전역 설정 (여기만 고치면 대부분 대응 가능)
├── requirements.txt
├── README.md
├── data/
│   ├── WASH_data_template.xlsx   # ★ source of truth (지속 업데이트할 실제 파일)
│   └── nepal_districts.geojson   # ★ 사용자가 넣어야 하는 Nepal District boundary (아래 3번 참고)
└── src/
    ├── data_loader.py            # Excel 읽기 + 정제
    ├── geo_utils.py               # GeoJSON 읽기 + District_ID join
    └── map_builder.py            # pydeck layer 생성
```

코드에는 District 데이터가 전혀 하드코딩되어 있지 않습니다. `data/WASH_data_template.xlsx` 를
수정하고 앱에서 **Reload data** 버튼만 누르면 지도가 즉시 갱신됩니다.

## 2. Excel 구조 (제안 반영본)

요청하신 구조를 거의 그대로 사용하되, 아래 2가지를 추가/정리했습니다.

- **`Districts_Lookup` 시트 신규 추가**: District_ID ↔ District ↔ Province 매핑을 별도로 관리하는
  "마스터 목록" 시트입니다. 다른 모든 시트(District_Situation, Water_Access, Interventions, Funding)는
  District_ID만 갖고 있으면 되고, District 이름 철자가 헷갈릴 일이 없습니다. GeoJSON과의 join도
  이 시트를 기준으로 검증할 수 있습니다.
- **District_Situation의 숫자/텍스트 분리는 이미 잘 설계되어 있어 그대로 유지**했습니다
  (`Affected_population`처럼 숫자는 숫자 컬럼, `Key_problems`처럼 서술은 텍스트 컬럼).
- 중복 우려가 있었던 `Water_quantity/quality/safety`가 `District_Situation`과 `Water_Access` 양쪽에
  있는데, 이는 중복이 아니라 **역할이 다릅니다**:
  - `District_Situation`의 값 = District 전체를 대표하는 요약값 (지도 클릭 시 바로 보이는 값)
  - `Water_Access`의 값 = 개별 water source별 값 (하나의 District 안에 여러 row, 상세 분석/차트용)
  → 그대로 유지하는 것을 권장합니다.

시트별 컬럼은 원안 그대로이며, 실제 생성된 템플릿 파일(`data/WASH_data_template.xlsx`)에
예시 1행씩 채워두었으니 열어서 형식을 확인하시면 됩니다. (노란 헤더 = 필수 시트,
Priority/Status 등은 드롭다운 목록으로 입력 오류를 방지했습니다.)

## 3. District boundary(GeoJSON) 연결 방법

현재 실제 Nepal District boundary 파일이 없어 `data/nepal_districts.geojson`은
빈 상태입니다. 아래 중 하나를 받아 이 경로에 넣어주세요.

- UN OCHA HDX — "Nepal - Subnational Administrative Boundaries" (admin2 = District 레벨)
- Survey Department of Nepal (측량국) 공식 shapefile
- ICIMOD RDS / DoLIDAR 배포본

파일을 넣은 뒤 다음을 확인하세요.

1. GeoJSON의 District 식별 필드명이 `District_ID`가 아니라면 `config.py`의
   `GEOJSON_ID_FIELD` 값을 실제 필드명으로 바꿔주세요 (예: `"DIST_ID"`, `"gid"` 등).
2. 그 필드의 **값 자체**가 Excel의 `District_ID`와 정확히 일치해야 합니다 (대소문자/공백 포함).
   보통 GeoJSON 쪽은 숫자 코드(예: `27`)이고 Excel은 문자 코드(예: `NPL-RAS`)로 서로 다른 경우가
   많으므로, 이 경우 `Districts_Lookup` 시트에 GeoJSON 코드값을 별도 컬럼으로 추가해
   매핑표로 활용하시는 걸 권장합니다. (실제 파일을 보내주시면 이 매핑 코드를 바로 만들어 드립니다.)
3. Shapefile(.shp)을 쓰실 경우 `.shp .shx .dbf .prj` 4개 파일이 같은 폴더에 함께 있어야
   `geopandas.read_file()`이 정상 작동합니다.

`app.py`는 이 파일이 없거나 필드가 안 맞으면 화면에 구체적인 에러 메시지를 보여주도록
만들어 두었습니다.

## 4. 실행 방법

```bash
cd unicef_wash_dashboard
pip install -r requirements.txt
streamlit run app.py
```

브라우저가 자동으로 열리고, 좌측 사이드바에서 Excel/GeoJSON 경로를 바꾸거나
**Reload data** 버튼으로 최신 데이터를 즉시 반영할 수 있습니다.

- Streamlit 1.38 이상: 지도 위 District를 **직접 클릭**하면 오른쪽 패널에 상세정보가 뜹니다
  (`st.pydeck_chart(..., on_select="rerun")` 기능 사용).
- 그보다 낮은 버전: 클릭 선택 대신 오른쪽 드롭다운으로 District를 선택하도록 자동 전환됩니다.
  (`pip install --upgrade streamlit` 로 최신화하면 클릭 선택이 활성화됩니다.)

Streamlit Community Cloud에 올려 팀원들과 온라인 공유하실 때도 동일한 `requirements.txt`와
폴더 구조 그대로 배포하면 됩니다. `data/WASH_data_template.xlsx`는 깃 저장소에 커밋해두고
주기적으로 업데이트 → git push 하는 방식을 권장합니다 (팀원 여러 명이 동시에 로컬에서
직접 파일을 고치는 방식보다 충돌이 적습니다).

## 5. 단계별 개발 로드맵 (요청하신 Phase 그대로)

| Phase | 내용 | 상태 |
|---|---|---|
| **Phase 1** | Excel/GeoJSON 읽기, District_ID join, Priority 색상 polygon 지도, 클릭 시 Affected population / Priority / 기본 정보 표시, Last updated 표시 | ✅ 이번에 구현 완료 |
| **Phase 2** | Water access/availability/quantity/quality/safety, Problem/Needs, 3단계 intervention(immediate/medium/long), Funding requirement/gap을 상세 패널에 전부 표시 + `Water_Access`/`Funding` 시트 기반 차트(plotly) 추가 | 다음 단계 |
| **Phase 3** | Schools/Health facilities/Water points 등 `Facilities` 시트 기반 point layer 추가 (pydeck `ScatterplotLayer` + layer on/off 토글) | 이후 |
| **Phase 4** | 전체 통계 요약, Province 필터, District 비교, funding gap 랭킹, Excel 다운로드/내보내기 | 이후 |

`src/data_loader.py`에 이미 `load_facilities()` 함수를 미리 만들어 두었기 때문에,
Phase 3에서는 `map_builder.py`에 `ScatterplotLayer`를 추가하고 `app.py`의 layer 리스트에
붙이기만 하면 됩니다 (기존 Phase 1 코드는 그대로 유지).
