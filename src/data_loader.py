# src/data_loader.py
import pandas as pd
import numpy as np
from config import DATA_FILE, WEIGHTS

# 시트명: '전체 프로젝트', 헤더는 2행(엑셀 기준), 0-index = 1
SHEET_NAME = '전체 프로젝트'
HEADER_ROW = 1  # 엑셀의 2행

# 화면/유사도에서 사용할 핵심 컬럼 (실제 엑셀 헤더와 일치해야 함)
REQUIRED_COLS = [
    '현장코드', '프로젝트명', '착공연도', '발주처', '수주여부',
    '공사금액', '사업구도', '지역1', '지역2', '지역3',
    '건축행위', '건축물종류', '대지면적(㎡)', '건축면적(㎡)', '연면적(㎡)',
    '지하층', '지상층',
    '지하 구조', '지상 구조', '구조체', '기초',
    '역타공법1', '역타공법2', '흙막이 공법1', '흙막이 공법2',
    '외장마감1', '외장마감2', '공사개월',
]

# 결과창에 표시할 '층수'는 지하/지상을 합쳐 만든 가상 컬럼
def _make_layer(row):
    u = row.get('지하층', '-'); a = row.get('지상층', '-')
    u_s = '-' if pd.isna(u) or str(u).strip() in ('', 'nan', '-') else str(u).strip()
    a_s = '-' if pd.isna(a) or str(a).strip() in ('', 'nan', '-') else str(a).strip()
    if u_s == '-' and a_s == '-': return '-'
    return f"B{u_s}/{a_s}F"


def load_projects(path=DATA_FILE):
    """엑셀 로드 — 2행을 헤더로, 3행(단위 행) 제거, 데이터 정제."""
    df = pd.read_excel(path, sheet_name=SHEET_NAME, header=HEADER_ROW)

    # 컬럼명 공백 정리
    df.columns = [str(c).strip() for c in df.columns]

    # 첫 행이 단위 행("(m2)", "(본)" 등)일 가능성이 높음 → 제거
    # 판단 기준: 'No.' 컬럼이 NaN이거나 숫자가 아니면 헤더 보조행으로 간주
    if 'No.' in df.columns:
        first = df.iloc[0]
        try:
            float(first['No.'])
        except (TypeError, ValueError):
            df = df.iloc[1:].reset_index(drop=True)

    # 모든 셀이 NaN인 행 제거
    df = df.dropna(how='all').reset_index(drop=True)

    # 현장코드 / 프로젝트명이 비어있는 행 제거
    if '프로젝트명' in df.columns:
        df = df[df['프로젝트명'].notna() & (df['프로젝트명'].astype(str).str.strip() != '')]
        df = df.reset_index(drop=True)

    # 숫자형 컬럼 변환
    numeric_cols = ['착공연도', '공사금액', '대지면적(㎡)', '건축면적(㎡)',
                    '연면적(㎡)', '지하층', '지상층', '공사개월']
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # 현장코드 문자열화 (502005.0 → 502005)
    if '현장코드' in df.columns:
        df['현장코드'] = df['현장코드'].apply(
            lambda x: '' if pd.isna(x) else str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) else str(x).strip()
        )

    # 가상 컬럼 '층수' 생성 (결과창 표시용)
    df['층수'] = df.apply(_make_layer, axis=1)

    return df


def get_building_types(df):
    """건축물종류 유니크 리스트 (드롭다운용)."""
    if '건축물종류' not in df.columns:
        return []
    types = df['건축물종류'].dropna().astype(str).str.strip()
    types = [t for t in types.unique() if t and t != '-' and t.lower() != 'nan']
    return sorted(types)


# ============================================================
# 유사도 계산
# ============================================================
SIMILAR_TYPES = {
    '업무시설': ['업무시설', '복합시설', '교육연구시설'],
    '창고시설': ['창고시설', '공장', '데이터센터'],
    '판매시설': ['판매시설', '복합시설'],
    '숙박시설': ['숙박시설'],
    '교육연구시설': ['교육연구시설', '업무시설'],
    '공장': ['공장', '창고시설', 'GMP'],
    '의료시설': ['의료시설'],
    '데이터센터': ['데이터센터', '창고시설'],
}


def _type_score(target_type, row_type):
    if not isinstance(row_type, str) or not isinstance(target_type, str): return 0.0
    if row_type == target_type: return 1.0
    similars = SIMILAR_TYPES.get(target_type, [])
    if row_type in similars: return 0.7
    return 0.0


def _num_distance(t, v, scale):
    if pd.isna(v) or pd.isna(t) or scale == 0: return 0.0
    return max(0.0, 1.0 - abs(float(t) - float(v)) / scale)


def calculate_similarity(df, target):
    """유사도 점수 계산 후 내림차순 정렬된 DataFrame 반환."""
    out = df.copy()
    t_type  = target.get('건축물종류', '')
    t_area  = float(target.get('연면적(㎡)', 0) or 0)
    t_under = float(target.get('지하층', 0) or 0)
    t_above = float(target.get('지상층', 0) or 0)

    area_scale  = max(t_area, 50000)  # 면적 정규화 스케일
    under_scale = 10
    above_scale = 30

    scores = []
    for _, row in out.iterrows():
        s_type  = _type_score(t_type, row.get('건축물종류'))
        s_area  = _num_distance(t_area,  row.get('연면적(㎡)'), area_scale)
        s_under = _num_distance(t_under, row.get('지하층'),     under_scale)
        s_above = _num_distance(t_above, row.get('지상층'),     above_scale)
        score = (WEIGHTS['building_type'] * s_type
               + WEIGHTS['area']          * s_area
               + WEIGHTS['underground']   * s_under
               + WEIGHTS['aboveground']   * s_above)
        scores.append(score)

    out['similarity'] = scores
    out = out.sort_values('similarity', ascending=False).reset_index(drop=True)
    return out
