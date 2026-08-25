# src/config.py
import sys
from pathlib import Path
import pandas as pd

# ============================================================
# 기본 경로
# PyInstaller 단일 exe 실행 시: sys.executable 위치 기준
# 일반 Python 실행 시: 소스 파일 위치 기준
# ============================================================
if getattr(sys, 'frozen', False):
    # PyInstaller로 빌드된 실행 파일로 실행될 때
    BASE_DIR = Path(sys.executable).parent
else:
    # 일반 Python 스크립트로 실행될 때
    BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / 'data'
DATA_FILE = DATA_DIR / 'data.xlsx'
REPORTS_ORIGINAL_DIR = BASE_DIR / 'Reports'
PACKAGE_FILE = BASE_DIR / 'reports.pkg'
OUTPUT_DIR = BASE_DIR / 'output'
LOGS_DIR = BASE_DIR / 'logs'

OUTPUT_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)

# ============================================================
# 보안 설정
# ============================================================
MASTER_PASSWORD = "CJ_ProjectHub_2026!@#"  # 배포 전 반드시 변경

# ============================================================
# UI 텍스트
# ============================================================
APP_NAME = "CJ ProjectHub"
APP_VERSION = "v1.1"
APP_TITLE = f"{APP_NAME} {APP_VERSION}"
TAB1_TITLE = "프로젝트 검색"
TAB2_TITLE = "키워드 본문 검색"
TAB3_TITLE = "준공보고서 검색"

# ============================================================
# 색상 (CJ 브랜드)
# ============================================================
COLOR_PRIMARY = "#003D7A"   # CJ Navy
COLOR_ACCENT  = "#E60012"   # CJ Red
COLOR_GRAY    = "#666666"
COLOR_LIGHT   = "#F5F5F5"

# ============================================================
# 유사도 가중치
# ============================================================
WEIGHTS = {
    'building_type': 0.40,
    'area':          0.30,
    'underground':   0.15,
    'aboveground':   0.15,
}

# ============================================================
# 결과창 / PDF 표시 필드 (18개 항목)
# 키 = 엑셀 컬럼명, 값 = 화면 표시명
# ============================================================
DISPLAY_FIELDS = [
    ('현장코드',       '현장코드'),
    ('프로젝트명',     '프로젝트명'),
    ('착공연도',       '착공연도'),
    ('건축물종류',     '건축물종류'),
    ('건축행위',       '건축행위'),
    ('연면적(㎡)',     '연면적(㎡)'),
    ('층수',           '층수'),
    ('공사금액',       '공사금액(억)'),
    ('공사개월',       '공사개월'),
    ('외장마감1',      '외장마감1'),
    ('외장마감2',      '외장마감2'),
    ('지하 구조',      '지하 구조'),
    ('지상 구조',      '지상 구조'),
    ('구조체',         '구조체'),
    ('기초',           '기초'),
    ('역타공법1',      '역타공법1'),
    ('역타공법2',      '역타공법2'),
    ('흙막이 공법1',   '흙막이 공법1'),
    ('흙막이 공법2',   '흙막이 공법2'),
]

# 보고서 검색 탭에서 사용할 검색 필드 (드롭다운)
REPORT_SEARCH_FIELDS = [
    '프로젝트명',
    '착공연도',
    '사업구도',
    '지역1',          # 화면 표시는 '지역'
    '건축행위',
    '건축물종류',
    '지하 구조',
    '지상 구조',
    '구조체',
    '외장마감1',
    '발주처',
]

# 화면 표시명 매핑
FIELD_DISPLAY_NAME = {
    '지역1': '지역',
}

# ============================================================
# 공란 표시
# ============================================================
EMPTY_MARK = '-'

def fmt(value, field_name=None):
    """값이 비어있으면 '-' 반환, 아니면 적절히 포맷."""
    if value is None or pd.isna(value) or str(value).strip() == '':
        return EMPTY_MARK
    s = str(value).strip()
    if s.lower() in ('nan', 'none', 'null'):
        return EMPTY_MARK
    if field_name == '연면적(㎡)':
        try: return f"{float(value):,.0f}"
        except: return s
    if field_name == '공사금액':
        try: return f"{float(value)/1e8:,.0f}"
        except: return s
    if field_name == '공사개월':
        try: return f"{float(value):.0f}"
        except: return s
    if field_name == '착공연도':
        try: return f"{int(float(value))}"
        except: return s
    return s
