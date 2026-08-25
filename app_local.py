"""
CJ ProjectHub — Streamlit 기반 통합 앱 v2.1
대시보드 · 건설공사비 지수 반영 · 유사현장 검색 · 키워드 본문 검색 · 준공보고서 검색 · 보고서 관리 · 모델 설명
UI/UX: CJ_EstAndPer 동일 스타일 + 개선 회귀모델(착공연도 포함) + 구조체 버그 수정
"""
import os, sys, io, json, math, getpass, shutil, tempfile, zipfile, base64
from datetime import datetime
from pathlib import Path
from collections import defaultdict

import numpy as np
import pandas as pd

# ── sklearn ──
HAS_SKLEARN = False
try:
    from sklearn.linear_model import LinearRegression
    from sklearn.metrics import r2_score, mean_absolute_error
    HAS_SKLEARN = True
except ImportError:
    pass

# ── requests ──
try:
    import requests as _requests
except ImportError:
    _requests = None

# ── plotly ──
HAS_PLOTLY = False
try:
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except ImportError:
    pass

# ── 경로 설정 ──
if getattr(sys, 'frozen', False):
    BUNDLE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    EXE_DIR = os.path.dirname(sys.executable)
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
    EXE_DIR = BUNDLE_DIR

# src/ 모듈 import를 위한 경로 추가
SRC_DIR = os.path.join(BUNDLE_DIR, 'src')
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)
if BUNDLE_DIR not in sys.path:
    sys.path.insert(0, BUNDLE_DIR)

import streamlit as st

# ── 선택적 임포트 ──
HAS_CRYPTO = False
HAS_PIL = False
try:
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    HAS_CRYPTO = True
except ImportError:
    pass
try:
    from PIL import Image, ImageDraw, ImageFont
    HAS_PIL = True
except ImportError:
    pass

# ── 앱 설정 ──
APP_NAME = "CJ ProjectHub"
APP_VERSION = "v2.1"
MASTER_PASSWORD = "CJ_ProjectHub_2026!@#"

# ── 데이터 파일 경로 탐색 ──
DATA_FILE = None
PACKAGE_FILE = None
REPORTS_DIR = None

for base in [EXE_DIR, BUNDLE_DIR, os.path.join(EXE_DIR, '..'),
             r"C:\Users\User\Desktop\Inference\CJ ProjectHub"]:
    candidate = os.path.join(base, 'data', 'data.xlsx')
    if os.path.exists(candidate):
        DATA_FILE = candidate
        break
if DATA_FILE is None:
    for base in [EXE_DIR, BUNDLE_DIR]:
        candidate = os.path.join(base, 'data.xlsx')
        if os.path.exists(candidate):
            DATA_FILE = candidate
            break

for base in [EXE_DIR, BUNDLE_DIR, os.path.join(EXE_DIR, '..'),
             r"C:\Users\User\Desktop\Inference\CJ ProjectHub",
             r"C:\Users\User\Desktop\Inference\CJ PJ HUB"]:
    candidate = os.path.join(base, 'reports.pkg')
    if os.path.exists(candidate):
        PACKAGE_FILE = candidate
        break

for base in [EXE_DIR, BUNDLE_DIR, os.path.join(EXE_DIR, '..'),
             r"C:\Users\User\Desktop\Inference\CJ ProjectHub"]:
    candidate = os.path.join(base, 'Reports')
    if os.path.isdir(candidate):
        REPORTS_DIR = candidate
        break


# ============================================================
#  건설공사비 지수 (KOSIS API + Fallback)
# ============================================================
KOSIS_API_KEY = "YTg3ZmQxN2M4ZTRmZmIwMjJlZjI3M2IwNTMyMGUxZTY="
KOSIS_ORG_ID  = "397"
KOSIS_TBL_ID  = "DT_39701_A003"
KOSIS_NONRES  = "15397AA2AA12"

NONRES_M = {
    "201501":82.50,"201502":82.60,"201503":82.80,"201504":83.00,
    "201505":83.10,"201506":83.20,"201507":83.30,"201508":83.40,
    "201509":83.50,"201510":83.50,"201511":83.40,"201512":83.30,
    "201601":83.50,"201602":83.70,"201603":83.90,"201604":84.10,
    "201605":84.30,"201606":84.50,"201607":84.70,"201608":84.80,
    "201609":84.90,"201610":85.00,"201611":85.10,"201612":85.20,
    "201701":85.50,"201702":85.80,"201703":86.20,"201704":86.50,
    "201705":86.80,"201706":87.00,"201707":87.20,"201708":87.40,
    "201709":87.60,"201710":87.80,"201711":88.00,"201712":88.20,
    "201801":88.50,"201802":88.80,"201803":89.20,"201804":89.50,
    "201805":89.80,"201806":90.00,"201807":90.20,"201808":90.40,
    "201809":90.60,"201810":90.80,"201811":91.00,"201812":91.20,
    "201901":93.00,"201902":93.50,"201903":94.20,"201904":94.80,
    "201905":95.20,"201906":95.50,"201907":95.80,"201908":96.00,
    "201909":96.20,"201910":96.50,"201911":96.80,"201912":97.00,
    "202001":98.50,"202002":98.80,"202003":99.00,"202004":99.30,
    "202005":99.50,"202006":99.80,"202007":100.00,"202008":100.20,
    "202009":100.40,"202010":100.60,"202011":100.80,"202012":101.00,
    "202101":102.10,"202102":103.50,"202103":105.80,"202104":108.20,
    "202105":110.30,"202106":111.80,"202107":112.50,"202108":113.20,
    "202109":113.80,"202110":114.50,"202111":115.10,"202112":115.80,
    "202201":117.50,"202202":118.80,"202203":120.50,"202204":122.30,
    "202205":124.10,"202206":125.50,"202207":126.20,"202208":126.30,
    "202209":125.80,"202210":124.50,"202211":123.20,"202212":122.50,
    "202301":125.20,"202302":126.10,"202303":126.80,"202304":127.20,
    "202305":127.50,"202306":127.80,"202307":128.10,"202308":128.30,
    "202309":128.50,"202310":128.70,"202311":128.90,"202312":129.00,
    "202401":129.77,"202402":130.05,"202403":130.05,"202404":130.08,
    "202405":130.20,"202406":130.11,"202407":129.96,"202408":129.71,
    "202409":129.60,"202410":129.50,"202411":129.50,"202412":129.40,
    "202501":130.50,"202502":130.55,"202503":130.60,"202504":130.60,
    "202505":130.55,"202506":130.55,"202507":130.50,"202508":130.45,
    "202509":131.25,"202510":131.55,"202511":132.05,"202512":132.25,
    "202601":132.35,
}

def _annual_avg(m):
    d = defaultdict(list)
    for ym, v in m.items():
        d[int(ym[:4])].append(v)
    return {y: round(sum(vs) / len(vs), 2) for y, vs in d.items()}

@st.cache_data(ttl=86400, show_spinner=False)
def fetch_nonres_index():
    mon, ok = {}, False
    if _requests:
        try:
            url = (f"https://kosis.kr/openapi/Param/statisticsParameterData.do"
                   f"?method=getList&apiKey={KOSIS_API_KEY}"
                   f"&orgId={KOSIS_ORG_ID}&tblId={KOSIS_TBL_ID}"
                   f"&itmId=T+&objL1={KOSIS_NONRES}+"
                   f"&prdSe=M&startPrdDe=201501&endPrdDe=202612&format=json")
            r = _requests.get(url, timeout=15)
            if r.status_code == 200:
                data = r.json()
                if isinstance(data, list) and len(data) > 0:
                    for row in data:
                        ym = str(row.get("PRD_DE", "")).strip()
                        dt = str(row.get("DT", "")).strip()
                        if ym and dt:
                            try:
                                mon[ym] = float(dt)
                            except:
                                pass
                    if len(mon) >= 12:
                        ok = True
        except:
            pass
    if not ok:
        mon = NONRES_M.copy()
    ann = _annual_avg(mon)
    return mon, ann, ok


# ============================================================
#  Streamlit 페이지 설정 + CSS (CJ_EstAndPer 동일)
# ============================================================
st.set_page_config(
    page_title=f"{APP_NAME} {APP_VERSION}",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    /* ── 사이드바 ── */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0a1628 0%, #132744 50%, #1a3a5c 100%) !important;
    }
    section[data-testid="stSidebar"] * {
        color: #c8d6e5 !important;
    }
    section[data-testid="stSidebar"] .stSelectbox label,
    section[data-testid="stSidebar"] .stNumberInput label,
    section[data-testid="stSidebar"] .stTextInput label {
        color: #a0aec0 !important; font-weight: 600 !important; font-size: 0.85rem !important;
    }

    /* ── 헤더 배너 ── */
    .header-banner {
        background: linear-gradient(135deg, #003D7A 0%, #0066CC 60%, #0088FF 100%);
        color: white; padding: 28px 36px; border-radius: 16px;
        margin-bottom: 24px; position: relative; overflow: hidden;
    }
    .header-banner::before {
        content: ''; position: absolute; top: -50%; right: -10%;
        width: 300px; height: 300px; border-radius: 50%;
        background: rgba(255,255,255,0.05);
    }
    .header-banner h1 { margin: 0; font-size: 1.8rem; font-weight: 800; }
    .header-banner .subtitle { font-size: 0.95rem; opacity: 0.9; margin-top: 4px; }
    .header-banner .badge {
        display: inline-block; background: rgba(255,255,255,0.18);
        padding: 4px 14px; border-radius: 20px; font-size: 0.72rem;
        font-weight: 600; margin-top: 10px;
    }

    /* ── KPI 카드 ── */
    .kpi-row { display: flex; gap: 16px; margin: 16px 0; flex-wrap: wrap; }
    .kpi-card {
        flex: 1; min-width: 180px; background: white;
        border-radius: 14px; padding: 18px 20px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.06);
        border-left: 4px solid #003D7A;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .kpi-card:hover {
        transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.10);
    }
    .kpi-card .icon { font-size: 1.5rem; margin-bottom: 4px; }
    .kpi-card .label { font-size: 0.78rem; color: #888; font-weight: 600; }
    .kpi-card .value { font-size: 1.6rem; font-weight: 800; color: #1a1a2e; margin: 2px 0; }
    .kpi-card .sub { font-size: 0.72rem; color: #aaa; }

    /* ── 섹션 타이틀 ── */
    .section-title {
        font-size: 1.05rem; font-weight: 700; color: #1a1a2e;
        padding: 10px 0 6px; border-bottom: 2px solid #003D7A;
        margin: 20px 0 12px;
    }

    /* ── 보고서 카드 ── */
    .report-card {
        background: white; border-radius: 12px; padding: 16px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin: 10px 0;
        border-left: 4px solid #003D7A;
    }
    .report-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.10); }

    /* ── 인포 박스 ── */
    .info-box {
        background: linear-gradient(135deg, #f0f4ff 0%, #e8f0fe 100%);
        border-radius: 12px; padding: 16px 20px; margin: 12px 0;
        border: 1px solid #d0d8e8;
    }

    /* ── 탭 스타일 ── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px; background: #f0f2f6; padding: 6px; border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px !important; padding: 8px 20px !important;
        font-weight: 600 !important;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #003D7A 0%, #0066CC 100%) !important;
        color: white !important; border: none !important; border-radius: 10px !important;
        font-weight: 600 !important; padding: 10px 20px !important;
    }

    /* ── 드래그앤드롭 영역 ── */
    .drop-zone {
        border: 3px dashed #003D7A; border-radius: 16px;
        background: linear-gradient(135deg, #f0f4ff 0%, #e8f0fe 100%);
        padding: 40px; text-align: center; margin: 16px 0;
        transition: all 0.3s;
    }
    .drop-zone:hover { background: #dde6f5; border-color: #0066CC; }

    /* ── 버튼 ── */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #003D7A 0%, #0066CC 100%) !important;
        color: white !important; border: none !important; border-radius: 10px !important;
        font-weight: 600 !important; padding: 10px 20px !important;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
#  데이터 로드 (파생 변수 포함)
# ============================================================
def _num(x):
    if pd.isna(x): return np.nan
    try: return float(str(x).replace(",", "").replace(" ", ""))
    except: return np.nan

def _make_layer(row):
    u = row.get('지하층', '-'); a = row.get('지상층', '-')
    u_s = '-' if pd.isna(u) or str(u).strip() in ('', 'nan', '-') else str(int(u))
    a_s = '-' if pd.isna(a) or str(a).strip() in ('', 'nan', '-') else str(int(a))
    if u_s == '-' and a_s == '-': return '-'
    return f"B{u_s}/{a_s}F"

def _sunta(val):
    v = str(val).upper().replace(" ", "")
    if any(k in v for k in ["FULL", "UPUP", "UP/UP", "DBS", "CWS"]): return 2
    if any(k in v for k in ["SEMI", "세미"]): return 1
    return 0

def _sunta_label(grade):
    return {0: "순타", 1: "Semi TopDown", 2: "Full TopDown/UPUP"}.get(grade, "순타")

def _demo(note, act):
    t = (str(note) + " " + str(act)).replace("nan", "")
    if "지하" in t and "철거" in t: return 2
    if "철거" in t: return 1
    return 0

@st.cache_data(show_spinner="📂 데이터 로딩 중...")
def load_data(path):
    if not path or not os.path.exists(path):
        return None, {}
    try:
        df = pd.read_excel(path, sheet_name='전체 프로젝트', header=1)
    except Exception:
        xls = pd.ExcelFile(path)
        df = pd.read_excel(path, sheet_name=xls.sheet_names[0], header=1)

    df.columns = [str(c).strip() for c in df.columns]

    # 단위 행 제거
    if 'No.' in df.columns:
        first = df.iloc[0]
        try:
            float(first['No.'])
        except (TypeError, ValueError):
            df = df.iloc[1:].reset_index(drop=True)

    df = df.dropna(how='all').reset_index(drop=True)
    if '프로젝트명' in df.columns:
        df = df[df['프로젝트명'].notna() & (df['프로젝트명'].astype(str).str.strip() != '')]
        df = df.reset_index(drop=True)

    numeric_cols = ['착공연도', '공사금액', '대지면적(㎡)', '건축면적(㎡)',
                    '연면적(㎡)', '지하층', '지상층', '공사개월']
    for c in numeric_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    if '현장코드' in df.columns:
        df['현장코드'] = df['현장코드'].apply(
            lambda x: '' if pd.isna(x) else str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) else str(x).strip()
        )

    df['층수'] = df.apply(_make_layer, axis=1)

    # 건축물종류 정리
    if '건축물종류' in df.columns:
        df['건축물종류'] = df['건축물종류'].astype(str).str.strip()
    else:
        df['건축물종류'] = '미분류'

    # ── 파생 변수 ──
    # 연면적 필터
    df = df[df['연면적(㎡)'].notna() & (df['연면적(㎡)'] > 0)].copy()
    df['log_연면적'] = np.log1p(df['연면적(㎡)'])

    # 평 변환 및 평당공사비
    pyeong = df['연면적(㎡)'] / 3.3058
    df['공사금액_억'] = df['공사금액'] / 1e8
    df['평당공사비'] = (df['공사금액_억'] * 1e4) / pyeong

    # 순타등급 (역타공법1 컬럼 기반)
    if '역타공법1' in df.columns:
        df['순타등급'] = df['역타공법1'].apply(_sunta)
    else:
        df['순타등급'] = 0

    # 철거등급
    df['철거등급'] = [_demo(n, a) for n, a in zip(
        df.get('비고', pd.Series('', index=df.index)),
        df.get('건축행위', pd.Series('', index=df.index))
    )]

    # 리모델링
    if '건축행위' in df.columns:
        df['리모델링'] = df['건축행위'].apply(
            lambda x: 1 if any(k in str(x) for k in ['리모델링', '증축', '대수선', '개축']) else 0)
    else:
        df['리모델링'] = 0

    # PC구조 (지하/지상 구조 컬럼 기반)
    df['PC구조'] = 0
    for c in ['지하 구조', '지상 구조']:
        if c in df.columns:
            df.loc[df[c].astype(str).str.contains('PC', case=False, na=False), 'PC구조'] = 1

    # 실제 구조체 (지하 구조 + 지상 구조 기반 — '구조체' 컬럼은 순타/역타이므로 사용 안함)
    df['실제구조'] = '기타'
    for c in ['지하 구조', '지상 구조']:
        if c in df.columns:
            s = df[c].astype(str).str.strip().str.upper()
            df.loc[s == 'SRC', '실제구조'] = 'SRC'
            df.loc[s == 'RC', '실제구조'] = df.loc[s == 'RC', '실제구조'].apply(
                lambda x: x if x == 'SRC' else 'RC')
            df.loc[s == 'S', '실제구조'] = df.loc[s == 'S', '실제구조'].apply(
                lambda x: x if x in ('SRC', 'RC') else 'S')
            df.loc[s == 'PC', '실제구조'] = df.loc[s == 'PC', '실제구조'].apply(
                lambda x: x if x in ('SRC', 'RC', 'S') else 'PC')

    # 서브셋
    subsets = {}
    off_mask = df['건축물종류'].str.contains('업무|사무|오피스', na=False)
    if off_mask.sum() >= 5:
        subsets['업무시설'] = df[off_mask].copy()
    wh_mask = df['건축물종류'].str.contains('창고|물류|배송|센터', na=False)
    if wh_mask.sum() >= 3:
        subsets['창고시설'] = df[wh_mask].copy()

    return df, subsets


# ============================================================
#  유사도 로직 (6요소 — 구조체 버그 수정 버전)
# ============================================================
SIMILAR_TYPES = {
    '업무시설': ['복합시설', '교육연구시설'],
    '창고시설': ['공장', '데이터센터'],
    '공장': ['창고시설', 'GMP', '실험동'],
    '판매시설': ['복합시설', '업무시설'],
    '숙박시설': ['복합시설'],
    '교육연구시설': ['업무시설', '실험동'],
    '의료시설': ['교육연구시설'],
    '데이터센터': ['창고시설', '공장'],
    '복합시설': ['업무시설', '판매시설', '숙박시설'],
    '문화체육관광시설': ['교육연구시설', '판매시설'],
    '운동시설': ['문화체육관광시설'],
    'GMP': ['공장', '실험동'],
    '실험동': ['교육연구시설', 'GMP'],
    '근린생활시설': ['판매시설'],
}

SIM_WEIGHTS = {
    'building_type': 0.35,
    'area':          0.25,
    'underground':   0.15,
    'aboveground':   0.10,
    'sunta':         0.10,
    'structure':     0.05,
}

def _type_score(target_type, row_type):
    if not isinstance(row_type, str) or not isinstance(target_type, str): return 0.0
    t = target_type.strip(); r = row_type.strip()
    if not t: return 0.5
    if r == t: return 1.0
    if r in SIMILAR_TYPES.get(t, []): return 0.6
    return 0.0

def _area_score(target_area, row_area):
    if pd.isna(row_area) or pd.isna(target_area) or target_area <= 0 or row_area <= 0: return 0.0
    log_t = math.log1p(target_area)
    log_r = math.log1p(row_area)
    log_range = math.log1p(700000)
    return max(0.0, 1.0 - abs(log_t - log_r) / log_range)

def _floor_score(target, value, max_range):
    if pd.isna(value) or pd.isna(target) or max_range == 0: return 0.0
    return max(0.0, 1.0 - abs(float(target) - float(value)) / max_range)

def _sunta_score(target_grade, row_grade):
    """역타공법 유사도: 동일=1.0, 1단계차=0.5, 2단계차=0.0"""
    if pd.isna(row_grade): return 0.5
    diff = abs(int(target_grade) - int(row_grade))
    if diff == 0: return 1.0
    if diff == 1: return 0.5
    return 0.0

def _structure_score(target_struct, row_struct):
    """구조체 유사도 (지하/지상 구조 컬럼 기반 '실제구조' 사용)"""
    if not isinstance(row_struct, str): return 0.3
    t = target_struct.strip().upper()
    r = row_struct.strip().upper()
    if not t or not r: return 0.3
    if t == r: return 1.0
    similar_pairs = {('RC', 'SRC'), ('SRC', 'RC')}
    if (t, r) in similar_pairs: return 0.7
    return 0.0

def calculate_similarity(df, target):
    out = df.copy()
    t_type = target.get('건축물종류', '')
    t_area = float(target.get('연면적(㎡)', 0) or target.get('연면적', 0) or 0)
    t_under = float(target.get('지하층', 0) or 0)
    t_above = float(target.get('지상층', 0) or 0)
    t_sunta = int(target.get('순타등급', 0) or 0)
    t_struct = str(target.get('구조체', '') or '')

    scores = []
    for _, row in out.iterrows():
        s_type = _type_score(t_type, row.get('건축물종류', ''))
        s_area = _area_score(t_area, row.get('연면적(㎡)', row.get('연면적', 0)))
        s_under = _floor_score(t_under, row.get('지하층', 0), 8)
        s_above = _floor_score(t_above, row.get('지상층', 0), 31)
        s_sunta = _sunta_score(t_sunta, row.get('순타등급', 0))
        s_struct = _structure_score(t_struct, row.get('실제구조', ''))

        score = (SIM_WEIGHTS['building_type'] * s_type
               + SIM_WEIGHTS['area'] * s_area
               + SIM_WEIGHTS['underground'] * s_under
               + SIM_WEIGHTS['aboveground'] * s_above
               + SIM_WEIGHTS['sunta'] * s_sunta
               + SIM_WEIGHTS['structure'] * s_struct)
        scores.append(round(score * 100, 1))

    out['유사도'] = scores
    out = out.sort_values('유사도', ascending=False).reset_index(drop=True)
    return out


# ============================================================
#  회귀 모델 (개선: 착공연도 포함)
# ============================================================
def build_models(subsets):
    if not HAS_SKLEARN:
        return {}
    models = {}
    off = subsets.get("업무시설", pd.DataFrame())
    if len(off) >= 5:
        # 공사기간: 착공연도 포함
        ft = ["log_연면적", "지하층", "지상층", "순타등급", "철거등급", "리모델링", "착공연도"]
        t = off.dropna(subset=ft + ["공사개월"]); t = t[t["공사개월"] > 0]
        if len(t) >= 4:
            X, y = t[ft].values, t["공사개월"].values
            m = LinearRegression().fit(X, y)
            models["업무_기간"] = {
                "model": m, "features": ft,
                "coef": dict(zip(ft, m.coef_)), "intercept": m.intercept_,
                "r2": r2_score(y, m.predict(X)), "mae": mean_absolute_error(y, m.predict(X)),
                "n": len(t),
            }
        # 평당공사비: 착공연도 포함
        fc = ["log_연면적", "지하층", "지상층", "순타등급", "리모델링", "착공연도"]
        c = off.dropna(subset=fc + ["평당공사비"]); c = c[c["평당공사비"] > 0]
        if len(c) >= 4:
            X, y = c[fc].values, c["평당공사비"].values
            m = LinearRegression().fit(X, y)
            models["업무_공사비"] = {
                "model": m, "features": fc,
                "coef": dict(zip(fc, m.coef_)), "intercept": m.intercept_,
                "r2": r2_score(y, m.predict(X)), "mae": mean_absolute_error(y, m.predict(X)),
                "n": len(c),
            }
    wh = subsets.get("창고시설", pd.DataFrame())
    if len(wh) >= 3:
        ft = ["log_연면적", "지하층", "PC구조", "착공연도"]
        t = wh.dropna(subset=ft + ["공사개월"]); t = t[t["공사개월"] > 0]
        if len(t) >= 3:
            X, y = t[ft].values, t["공사개월"].values
            m = LinearRegression().fit(X, y)
            models["창고_기간"] = {
                "model": m, "features": ft,
                "coef": dict(zip(ft, m.coef_)), "intercept": m.intercept_,
                "r2": r2_score(y, m.predict(X)), "mae": mean_absolute_error(y, m.predict(X)),
                "n": len(t),
            }
        fc = ["log_연면적", "지하층", "PC구조", "착공연도"]
        c = wh.dropna(subset=fc + ["평당공사비"]); c = c[c["평당공사비"] > 0]
        if len(c) >= 3:
            X, y = c[fc].values, c["평당공사비"].values
            m = LinearRegression().fit(X, y)
            models["창고_공사비"] = {
                "model": m, "features": fc,
                "coef": dict(zip(fc, m.coef_)), "intercept": m.intercept_,
                "r2": r2_score(y, m.predict(X)), "mae": mean_absolute_error(y, m.predict(X)),
                "n": len(c),
            }
    return models

def predict_model(info, x_dict, min_val=3.0):
    val = info["intercept"]
    for f in info["features"]:
        val += info["coef"][f] * x_dict.get(f, 0)
    return max(min_val, val)


# ============================================================
#  보고서 패키지 로더 (v1/v2 지원)
# ============================================================
class SecurePackageLoader:
    def __init__(self, pkg_path):
        self.pkg_path = Path(pkg_path)
        self._zip = None
        self.manifest = None
        if self.pkg_path.exists() and HAS_CRYPTO:
            try: self._load()
            except Exception as e:
                st.sidebar.warning(f"보고서 패키지 로드 실패: {e}")

    def _derive_key(self, password, salt):
        kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=480000)
        return base64.urlsafe_b64encode(kdf.derive(password.encode()))

    def _derive_raw_key(self, password, salt):
        kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=480000)
        return kdf.derive(password.encode())

    def _load(self):
        with open(self.pkg_path, 'rb') as f:
            magic = f.read(8)
        if magic == b'CJPHv1\x00\x00':
            self._load_v1()
        elif magic == b'CJPHv2\x00\x00':
            self._load_v2()
        else:
            raise ValueError("유효하지 않은 패키지 파일")

    def _load_v1(self):
        from cryptography.fernet import Fernet
        with open(self.pkg_path, 'rb') as f:
            f.read(8); salt = f.read(16); encrypted = f.read()
        key = self._derive_key(MASTER_PASSWORD, salt)
        cipher = Fernet(key)
        decrypted = cipher.decrypt(encrypted)
        self._zip = zipfile.ZipFile(io.BytesIO(decrypted), 'r')
        with self._zip.open('manifest.json') as f:
            self.manifest = json.load(f)

    def _load_v2(self):
        import hmac as hmac_mod, hashlib
        from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
        CHUNK = 64 * 1024 * 1024
        with open(self.pkg_path, 'rb') as f:
            f.read(8); salt = f.read(16); nonce = f.read(16)
            stored_mac = f.read(32); encrypted_data = f.read()
        raw_key = self._derive_raw_key(MASTER_PASSWORD, salt)
        computed_mac = hmac_mod.new(raw_key, encrypted_data, hashlib.sha256).digest()
        if not hmac_mod.compare_digest(stored_mac, computed_mac):
            raise ValueError("패키지 무결성 검증 실패")
        cipher = Cipher(algorithms.AES(raw_key), modes.CTR(nonce))
        decryptor = cipher.decryptor()
        buf = io.BytesIO()
        off = 0
        while off < len(encrypted_data):
            chunk = encrypted_data[off:off + CHUNK]
            buf.write(decryptor.update(chunk))
            off += CHUNK
        buf.write(decryptor.finalize())
        buf.seek(0)
        self._zip = zipfile.ZipFile(buf, 'r')
        with self._zip.open('manifest.json') as f:
            self.manifest = json.load(f)

    def is_loaded(self): return self._zip is not None
    def has_report(self, site_code):
        return str(site_code) in self.manifest['reports'] if self.manifest else False
    def get_page_count(self, site_code):
        return self.manifest['reports'][str(site_code)]['page_count'] if self.has_report(site_code) else 0
    def get_original_name(self, site_code):
        return self.manifest['reports'][str(site_code)].get('original_name', '') if self.has_report(site_code) else None

    def get_page_image(self, site_code, page_no):
        site_code = str(site_code)
        if not self.has_report(site_code): return None
        for prefix in ['slide_', 'page_']:
            try:
                path = f"{site_code}/{prefix}{page_no:03d}.png"
                with self._zip.open(path) as f:
                    return Image.open(io.BytesIO(f.read())).copy()
            except (KeyError, Exception): continue
        return None

    def search_text(self, keyword):
        if not self.manifest: return []
        results = []; kw = keyword.lower().strip()
        if not kw: return []
        for site_code, info in self.manifest['reports'].items():
            for entry in info.get('text_index', []):
                content = entry.get('content', '')
                if kw in content.lower():
                    idx = content.lower().find(kw)
                    start = max(0, idx - 30); end = min(len(content), idx + len(keyword) + 30)
                    snippet = content[start:end].replace('\n', ' ')
                    if start > 0: snippet = '…' + snippet
                    if end < len(content): snippet = snippet + '…'
                    results.append({
                        'site_code': site_code, 'page': entry['page'],
                        'snippet': snippet, 'original_name': info.get('original_name', '')
                    })
        return results

    def get_all_report_codes(self):
        if not self.manifest: return []
        return list(self.manifest['reports'].keys())


# ============================================================
#  워터마크 헬퍼
# ============================================================
def add_watermark(img):
    if not HAS_PIL: return img
    try:
        user = getpass.getuser()
        wm_text = f"{user} | {datetime.now().strftime('%Y-%m-%d %H:%M')} | CJ ProjectHub"
        overlay = Image.new('RGBA', img.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)
        try: font = ImageFont.truetype("C:/Windows/Fonts/malgun.ttf", 28)
        except: font = ImageFont.load_default()
        w, h = img.size
        for y_pos in range(-h, h * 2, 220):
            for x_pos in range(-w, w * 2, 380):
                draw.text((x_pos, y_pos), wm_text, fill=(180, 180, 180, 70), font=font)
        base = img.convert('RGBA')
        return Image.alpha_composite(base, overlay).convert('RGB')
    except:
        return img


# ============================================================
#  포맷팅 헬퍼
# ============================================================
def fmt_area(val):
    if pd.isna(val) or val == 0: return "-"
    pyeong = val / 3.3058
    return f"{val:,.0f}㎡ ({pyeong:,.0f}평)"

def fmt_cost(val):
    if pd.isna(val) or val == 0: return "-"
    return f"{val:,.0f}억원"

def fmt_months(val):
    if pd.isna(val) or val == 0: return "-"
    return f"{val:.0f}개월"

def fmt(value, field_name=None):
    if value is None or pd.isna(value) or str(value).strip() == '': return '-'
    s = str(value).strip()
    if s.lower() in ('nan', 'none', 'null'): return '-'
    if field_name == '연면적(㎡)':
        try: return f"{float(value):,.0f}"
        except: return s
    if field_name in ('공사금액',):
        try: return f"{float(value)/1e8:,.0f}"
        except: return s
    if field_name == '공사개월':
        try: return f"{float(value):.0f}"
        except: return s
    if field_name == '착공연도':
        try: return f"{int(float(value))}"
        except: return s
    return s

# 표시 필드 정의
DISPLAY_FIELDS = [
    ('현장코드', '현장코드'), ('프로젝트명', '프로젝트명'), ('착공연도', '착공연도'),
    ('건축물종류', '건축물종류'), ('건축행위', '건축행위'), ('연면적(㎡)', '연면적(㎡)'),
    ('층수', '층수'), ('공사금액', '공사금액(억)'), ('공사개월', '공사개월'),
    ('외장마감1', '외장마감1'), ('외장마감2', '외장마감2'),
    ('지하 구조', '지하 구조'), ('지상 구조', '지상 구조'),
    ('구조체', '구조체'), ('기초', '기초'),
    ('역타공법1', '역타공법1'), ('역타공법2', '역타공법2'),
    ('흙막이 공법1', '흙막이 공법1'), ('흙막이 공법2', '흙막이 공법2'),
]

REPORT_SEARCH_FIELDS = [
    '프로젝트명', '착공연도', '사업구도', '지역1', '건축행위',
    '건축물종류', '지하 구조', '지상 구조', '구조체', '외장마감1', '발주처',
]

FIELD_DISPLAY_NAME = {'지역1': '지역'}


# ============================================================
#  보고서 보유 확인 헬퍼
# ============================================================
def has_report(pkg, site_code):
    if not pkg or not pkg.is_loaded() or not site_code:
        return False
    s = str(site_code).strip()
    candidates = [s]
    if s.endswith('.0'): candidates.append(s[:-2])
    if '.' not in s: candidates.append(s + '.0')
    try: candidates.append(str(int(float(s))))
    except: pass
    for cand in candidates:
        try:
            if pkg.has_report(cand): return True
        except: continue
    return False


# ============================================================
#  건설공사비 지수 탭 렌더
# ============================================================
def render_index_tab(comp_name, comp_year, comp_eok, comp_py, R):
    mon, ann, api_ok = fetch_nonres_index()
    src_icon = "✅" if api_ok else "📋"
    src_text = "KOSIS API 실시간" if api_ok else "Fallback (KICT·CERIK)"
    st.markdown(f"""
    <div class="info-box">
        <b style="color:#003D7A;">데이터 출처</b>
        <span style="background:#003D7A;color:white;padding:2px 8px;border-radius:12px;
                     font-size:0.7rem;font-weight:600;margin-left:8px;">{src_icon} {src_text}</span>
        <div style="color:#555;font-size:0.82rem;margin-top:4px;">
            📊 비주거용건물 건설공사비지수 (2020=100) | 🏛️ KICT | 📡 KOSIS
        </div>
    </div>""", unsafe_allow_html=True)

    idx_comp = ann.get(comp_year)
    warn_msg = ""
    if idx_comp is None:
        avail = sorted(ann.keys())
        nearest = min(avail, key=lambda y: abs(y - comp_year)) if avail else None
        if nearest:
            idx_comp = ann[nearest]
            warn_msg = f"⚠️ {comp_year}년 지수 미보유 → {nearest}년 ({idx_comp:.2f}) 적용"
        else:
            idx_comp = 100.0
            warn_msg = "⚠️ 지수 데이터 부족 — 100.0 적용"

    options, labels = [], {}
    sy = sorted(ann.keys())
    ly = max(sy) if sy else 2026
    for yr in range(comp_year, ly + 1):
        if yr in ann:
            k = f"Y{yr}"; options.append(k); labels[k] = f"{yr}년 연평균 ({ann[yr]:.2f})"
    lm = sorted([ym for ym in mon if int(ym[:4]) == ly])
    for ym in lm:
        k = f"M{ym}"; options.append(k); labels[k] = f"{ym[:4]}년 {int(ym[4:])}월 ({mon[ym]:.2f})"
    di = len(options) - 1 if options else 0

    with st.expander("⚙️ 비교 파라미터 설정", expanded=False):
        p1, p2 = st.columns(2)
        comp_name = p1.text_input("비교 PJT명", value=comp_name, key="ci_name")
        comp_year = int(p1.number_input("착공연도", value=comp_year, min_value=2010, max_value=2030, key="ci_yr"))
        comp_eok = p2.number_input("총공사비(억원)", value=float(comp_eok), min_value=0.0, key="ci_eok")
        comp_py = p2.number_input("연면적(평)", value=float(comp_py), min_value=0.0, key="ci_py")

    if warn_msg: st.warning(warn_msg)
    if not options: st.error("검토 기준 지수 옵션 없음"); return

    sel = st.selectbox("📅 검토 기준 시점", options, index=di, format_func=lambda x: labels.get(x, x), key="ci_sel")
    if sel.startswith("Y"):
        idx_rev = ann.get(int(sel[1:]), idx_comp)
    else:
        idx_rev = mon.get(sel[1:], idx_comp)

    cpp = (comp_eok * 1e4 / comp_py) if comp_py > 0 else 0
    fac = idx_rev / idx_comp if idx_comp > 0 else 1.0
    adj_cpp = cpp * fac; adj_eok = adj_cpp * comp_py / 1e4
    est_cpp = R.get("cost_per_py", 0); est_eok = R.get("total_eok", 0)
    est_adj_cpp = est_cpp * fac; est_adj_eok = est_adj_cpp * (R.get("area", 0) / 3.3058) / 1e4

    st.markdown(f"""
    <div class="kpi-row">
        <div class="kpi-card"><div class="icon">📌</div><div class="label">비교 지수 ({comp_year}년)</div>
            <div class="value">{idx_comp:.2f}</div><div class="sub">연평균 기준</div></div>
        <div class="kpi-card" style="border-left-color:#00cc96;"><div class="icon">🎯</div><div class="label">검토 시점 지수</div>
            <div class="value">{idx_rev:.2f}</div><div class="sub">{labels.get(sel,'')[:12]}</div></div>
        <div class="kpi-card" style="border-left-color:#ab63fa;"><div class="icon">⚖️</div><div class="label">보정 계수</div>
            <div class="value">{fac:.4f}</div><div class="sub">검토/비교</div></div>
        <div class="kpi-card" style="border-left-color:#E60012;"><div class="icon">📈</div><div class="label">변동률</div>
            <div class="value">{(fac-1)*100:+.2f}%</div><div class="sub">공사비 반영분</div></div>
    </div>
    <div class="kpi-row">
        <div class="kpi-card"><div class="icon">💰</div><div class="label">지수 반영 추정 평당공사비</div>
            <div class="value">{est_adj_cpp:,.0f}</div><div class="sub">만원/평 (원본 {est_cpp:,.0f})</div></div>
        <div class="kpi-card"><div class="icon">🏦</div><div class="label">지수 반영 추정 총공사비</div>
            <div class="value">{est_adj_eok:,.1f}</div><div class="sub">억원 (원본 {est_eok:,.1f})</div></div>
    </div>""", unsafe_allow_html=True)

    # 비교 테이블
    st.markdown('<div class="section-title">📋 평당 공사비 · 총공사비 비교</div>', unsafe_allow_html=True)
    cdf = pd.DataFrame({
        "구분": [f"비교 PJT: {comp_name} ({comp_year}년)", f"비교 PJT 보정 후",
                f"검토 PJT 추정 (원본)", f"검토 PJT 지수 반영"],
        "평당공사비 (만원/평)": [f"{cpp:,.0f}", f"{adj_cpp:,.0f}", f"{est_cpp:,.0f}", f"{est_adj_cpp:,.0f}"],
        "총공사비 (억원)": [f"{comp_eok:,.1f}", f"{adj_eok:,.1f}", f"{est_eok:,.1f}", f"{est_adj_eok:,.1f}"],
        "적용 지수": [f"{idx_comp:.2f}", f"{idx_rev:.2f}", "-", f"{idx_rev:.2f}"],
    })
    st.dataframe(cdf, use_container_width=True, hide_index=True)

    # 비교 막대그래프
    if HAS_PLOTLY:
        col_a, col_b = st.columns(2)
        nms_idx = [f"비교 PJT\n({comp_year}년)", f"비교 PJT\n보정 후", "검토 PJT\n(원본)", "검토 PJT\n지수 반영"]
        clrs = ["#B0B0B0", "#B0B0B0", "#003D7A", "#003D7A"]
        with col_a:
            vals_cpp = [cpp, adj_cpp, est_cpp, est_adj_cpp]
            y_max = max(v for v in vals_cpp if v > 0) * 1.3 if any(v > 0 for v in vals_cpp) else 100
            fig_cpp = go.Figure()
            fig_cpp.add_trace(go.Bar(x=nms_idx, y=vals_cpp, marker_color=clrs,
                text=[f"{v:,.0f}" for v in vals_cpp], textposition="outside", width=0.4))
            fig_cpp.update_layout(height=320, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                yaxis=dict(gridcolor="#e8e8e8", range=[0, y_max], tickformat=",", title="만원/평"),
                showlegend=False, margin=dict(t=30, b=10))
            st.plotly_chart(fig_cpp, use_container_width=True)
        with col_b:
            vals_eok = [comp_eok, adj_eok, est_eok, est_adj_eok]
            y_max2 = max(v for v in vals_eok if v > 0) * 1.3 if any(v > 0 for v in vals_eok) else 100
            fig_eok = go.Figure()
            fig_eok.add_trace(go.Bar(x=nms_idx, y=vals_eok, marker_color=clrs,
                text=[f"{v:,.1f}" for v in vals_eok], textposition="outside", width=0.4))
            fig_eok.update_layout(height=320, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                yaxis=dict(gridcolor="#e8e8e8", range=[0, y_max2], tickformat=",", title="억원"),
                showlegend=False, margin=dict(t=30, b=10))
            st.plotly_chart(fig_eok, use_container_width=True)

    # 지수 추이 차트
    st.markdown('<div class="section-title">📈 비주거용건물 건설공사비지수 추이</div>', unsafe_allow_html=True)
    if HAS_PLOTLY:
        sy2 = sorted(ann.keys())
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=sy2, y=[ann[yr] for yr in sy2], mode="lines+markers", name="연평균 지수",
                                 line=dict(color="#003D7A", width=2.5), marker=dict(size=7)))
        fig.update_layout(height=400, xaxis_title="연도", yaxis_title="지수 (2020=100)",
                          plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                          xaxis=dict(gridcolor="#e8e8e8"), yaxis=dict(gridcolor="#e8e8e8"),
                          legend=dict(orientation="h", y=1.12))
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
#  ★ 메인
# ============================================================
def main():
    st.markdown(f"""
    <div class="header-banner">
        <h1>🏗️ CJ대한통운 건설부문</h1>
        <p class="subtitle">ProjectHub — PJT 추정 · 유사현장 검색 · 보고서 열람 통합 시스템</p>
        <span class="badge">{APP_VERSION}  |  6요소 통합 유사도  |  Linear Regression (개선)  |  KOSIS API  |  암호화 패키지(v1/v2)</span>
    </div>""", unsafe_allow_html=True)

    # 데이터 로드
    df_all, subsets = load_data(DATA_FILE)
    if df_all is None:
        st.error("⚠️ data.xlsx 파일을 찾을 수 없거나 형식이 올바르지 않습니다.")
        if DATA_FILE:
            st.info(f"시도한 경로: {DATA_FILE}")
        st.stop()

    models = build_models(subsets)

    # 보고서 패키지 로드
    if 'pkg_loader' not in st.session_state:
        if PACKAGE_FILE and os.path.exists(PACKAGE_FILE) and HAS_CRYPTO and HAS_PIL:
            with st.spinner("📦 보고서 패키지 로딩 중..."):
                st.session_state['pkg_loader'] = SecurePackageLoader(PACKAGE_FILE)
        else:
            st.session_state['pkg_loader'] = None
    pkg = st.session_state['pkg_loader']

    # 건축물종류 목록
    all_types = sorted(df_all['건축물종류'].dropna().astype(str).str.strip().unique().tolist())
    all_types = [t for t in all_types if t and t != '-' and t.lower() != 'nan']

    # 보고서 보유 현장 수
    n_with_report = 0
    if pkg and pkg.is_loaded():
        for _, row in df_all.iterrows():
            if has_report(pkg, row.get('현장코드', '')):
                n_with_report += 1

    # ── 사이드바 ──
    with st.sidebar:
        st.markdown("""
        <div style="text-align:center;padding:12px 0 18px;">
            <div style="font-size:2.2rem;margin-bottom:4px;">🏗️</div>
            <div style="font-size:1.1rem;font-weight:800;color:#a0aec0;letter-spacing:0.05em;">PJT 검토 입력</div>
            <div style="font-size:0.7rem;color:#4da6ff;margin-top:4px;font-weight:600;">CJ CONSTRUCTION DIVISION</div>
        </div>""", unsafe_allow_html=True)
        st.divider()

        pjt_type = st.selectbox("🏢 건축물 종류", all_types,
                                index=all_types.index("업무시설") if "업무시설" in all_types else 0)
        is_off = any(k in pjt_type for k in ["사무", "업무", "오피스"])
        is_wh = any(k in pjt_type for k in ["창고", "물류", "배송", "센터"])

        area = st.number_input("📐 연면적 (㎡)", value=30000.0, min_value=100.0, step=1000.0)
        st.caption(f"   ≈ {area/3.3058:,.0f} 평")
        c1, c2 = st.columns(2)
        b_fl = c1.number_input("⬇️ 지하층", value=4, min_value=0, max_value=15)
        a_fl = c2.number_input("⬆️ 지상층", value=15, min_value=1, max_value=80)

        st.divider()
        st.markdown('<div style="font-size:0.8rem;color:#4da6ff;font-weight:700;">▸ 공법 · 구조 옵션</div>', unsafe_allow_html=True)
        rev_opt = st.selectbox("🔄 역타 적용", ["순타 (역타 없음)", "Semi TopDown", "Full TopDown / UPUP"])
        sunta = 0 if "순타" in rev_opt else (1 if "Semi" in rev_opt else 2)

        struct_opt = st.selectbox("🧱 구조체", ["RC", "SRC", "S (철골)", "PC"])
        struct_code = struct_opt.split(" ")[0]
        pc = 1 if "PC" in struct_opt else 0

        if is_off:
            demo_opt = st.selectbox("🔨 철거 여부", ["없음", "지상 철거 포함", "지하 포함 철거"])
            demo = 0 if "없음" in demo_opt else (1 if "지상" in demo_opt else 2)
            act_opt = st.selectbox("🏗️ 건축행위", ["신축", "리모델링/증축"])
            remod = 0 if "신축" in act_opt else 1
        else:
            demo, remod = 0, 0

        # 착공연도 (추정에 사용)
        est_year = st.number_input("📅 예상 착공연도", value=2026, min_value=2015, max_value=2035)

        st.divider()
        run = st.button("🚀 추정 실행", type="primary", use_container_width=True)
        st.divider()

        # 데이터 현황
        pkg_status = "✅ 로드됨" if (pkg and pkg.is_loaded()) else "❌ 미로드"
        n_off = len(subsets.get('업무시설', []))
        n_wh = len(subsets.get('창고시설', []))
        st.markdown(f"""
        <div style="background:rgba(0,61,122,0.15);border-radius:10px;padding:12px;text-align:center;">
            <div style="font-size:0.7rem;color:#4da6ff;font-weight:700;margin-bottom:6px;">📊 DATA STATUS</div>
            <div style="display:flex;justify-content:space-around;">
                <div><div style="font-size:1.1rem;font-weight:800;color:#e0e0e0;">{len(df_all)}</div><div style="font-size:0.65rem;color:#888;">전체</div></div>
                <div><div style="font-size:1.1rem;font-weight:800;color:#4da6ff;">{n_off}</div><div style="font-size:0.65rem;color:#888;">업무</div></div>
                <div><div style="font-size:1.1rem;font-weight:800;color:#00cc96;">{n_wh}</div><div style="font-size:0.65rem;color:#888;">물류</div></div>
                <div><div style="font-size:1.1rem;font-weight:800;color:#ab63fa;">{n_with_report}</div><div style="font-size:0.65rem;color:#888;">보고서</div></div>
            </div>
            <div style="font-size:0.65rem;color:#888;margin-top:8px;">패키지: {pkg_status} | 모델: {len(models)}개</div>
        </div>""", unsafe_allow_html=True)

    # ── 추정 실행 ──
    if run:
        log_a = np.log1p(area); a_py = area / 3.3058

        # 공사기간 추정
        if is_off and "업무_기간" in models:
            x_dict = {"log_연면적": log_a, "지하층": b_fl, "지상층": a_fl,
                      "순타등급": sunta, "철거등급": demo, "리모델링": remod, "착공연도": est_year}
            pred_m = predict_model(models["업무_기간"], x_dict)
        elif is_wh and "창고_기간" in models:
            x_dict = {"log_연면적": log_a, "지하층": b_fl, "PC구조": pc, "착공연도": est_year}
            pred_m = predict_model(models["창고_기간"], x_dict)
        else:
            target_sim = {'건축물종류': pjt_type, '연면적(㎡)': area, '지하층': b_fl, '지상층': a_fl,
                          '순타등급': sunta, '구조체': struct_code}
            sim_temp = calculate_similarity(df_all, target_sim)
            top_sim = sim_temp.head(5)
            valid_months = top_sim['공사개월'].dropna()
            pred_m = valid_months.mean() if len(valid_months) > 0 else 12 + area / 5000 + b_fl * 1.5

        # 공사비 추정
        if is_off and "업무_공사비" in models:
            x_dict = {"log_연면적": log_a, "지하층": b_fl, "지상층": a_fl,
                      "순타등급": sunta, "리모델링": remod, "착공연도": est_year}
            pred_c = predict_model(models["업무_공사비"], x_dict, min_val=100.0)
        elif is_wh and "창고_공사비" in models:
            x_dict = {"log_연면적": log_a, "지하층": b_fl, "PC구조": pc, "착공연도": est_year}
            pred_c = predict_model(models["창고_공사비"], x_dict, min_val=100.0)
        else:
            if 'sim_temp' not in dir():
                target_sim = {'건축물종류': pjt_type, '연면적(㎡)': area, '지하층': b_fl, '지상층': a_fl,
                              '순타등급': sunta, '구조체': struct_code}
                sim_temp = calculate_similarity(df_all, target_sim)
            top_sim = sim_temp.head(5)
            valid_costs = top_sim['평당공사비'].dropna()
            pred_c = valid_costs.mean() if len(valid_costs) > 0 else 400.0

        t_eok = pred_c * a_py / 1e4
        result = {"type": pjt_type, "area": area, "b_floors": b_fl, "a_floors": a_fl,
                  "cost_per_py": pred_c, "total_eok": t_eok, "months": pred_m,
                  "sunta": sunta, "demo": demo, "remodel": remod, "struct": struct_code,
                  "est_year": est_year}
        st.session_state["result"] = result

        # 유사현장 계산
        target_for_sim = {'건축물종류': pjt_type, '연면적(㎡)': area, '지하층': b_fl, '지상층': a_fl,
                          '순타등급': sunta, '구조체': struct_code}
        sim = calculate_similarity(df_all, target_for_sim)
        st.session_state["sim_df"] = sim

    # ── 추정 전 안내 ──
    if "result" not in st.session_state:
        st.markdown("""
        <div style="text-align:center;padding:80px 20px;">
            <div style="font-size:4rem;margin-bottom:16px;opacity:0.6;">📋</div>
            <div style="font-size:1.15rem;color:#888;font-weight:600;">좌측 사이드바에서 PJT 정보를 입력하고</div>
            <div style="font-size:1.15rem;color:#003D7A;font-weight:700;margin-top:4px;">'🚀 추정 실행' 버튼을 클릭하세요</div>
        </div>""", unsafe_allow_html=True)
        return

    R = st.session_state["result"]
    sim = st.session_state.get("sim_df", pd.DataFrame())

    # ── 탭 구성 ──
    tab_names = ["📊 대시보드", "📈 건설공사비 지수 반영", "🔍 유사현장 검색",
                 "📖 키워드 본문 검색", "📋 준공보고서 검색",
                 "📑 보고서 열람", "📥 보고서 관리", "🔬 모델 설명"]
    tabs = st.tabs(tab_names)

    # ══════════════════════════════════════════
    #  탭 1: 대시보드
    # ══════════════════════════════════════════
    with tabs[0]:
        pyeong = R['area'] / 3.3058
        st.markdown(f"""
        <div class="kpi-row">
            <div class="kpi-card"><div class="icon">💰</div>
                <div class="label">추정 평당공사비</div><div class="value">{R['cost_per_py']:,.0f}</div><div class="sub">만원/평</div></div>
            <div class="kpi-card"><div class="icon">🏦</div>
                <div class="label">추정 총공사비</div><div class="value">{R['total_eok']:,.1f}</div><div class="sub">억원</div></div>
            <div class="kpi-card"><div class="icon">📅</div>
                <div class="label">추정 공사기간</div><div class="value">{R['months']:.1f}</div><div class="sub">개월</div></div>
            <div class="kpi-card"><div class="icon">📐</div>
                <div class="label">연면적</div><div class="value">{R['area']:,.0f}</div><div class="sub">㎡ ({pyeong:,.0f}평)</div></div>
        </div>""", unsafe_allow_html=True)

        if HAS_PLOTLY and sim is not None and len(sim) > 0:
            top5 = sim.head(5)
            nms = ["검토 PJT"] + [str(r.get("프로젝트명", ""))[:18] for _, r in top5.iterrows()]
            bar_clrs = ["#003D7A"] + ["#B0B0B0"] * len(top5)

            st.markdown('<div class="section-title">📊 유사 현장 대비 포지셔닝 — 공사기간</div>', unsafe_allow_html=True)
            m_vals = [R["months"]] + [r.get("공사개월", 0) if pd.notna(r.get("공사개월", 0)) else 0 for _, r in top5.iterrows()]
            m_max = max(v for v in m_vals if v > 0) * 1.3 if any(v > 0 for v in m_vals) else 30
            fig_m = go.Figure()
            fig_m.add_trace(go.Bar(x=nms, y=m_vals, marker_color=bar_clrs,
                                   text=[fmt_months(v) for v in m_vals], textposition="outside", width=0.35))
            fig_m.update_layout(height=340, yaxis_title="공사기간 (개월)",
                                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                                yaxis=dict(gridcolor="#e8e8e8", range=[0, m_max]), showlegend=False)
            st.plotly_chart(fig_m, use_container_width=True)

            st.markdown('<div class="section-title">📊 유사 현장 대비 포지셔닝 — 평당공사비</div>', unsafe_allow_html=True)
            c_vals = [R["cost_per_py"]] + [r.get("평당공사비", 0) if pd.notna(r.get("평당공사비", 0)) else 0 for _, r in top5.iterrows()]
            c_max = max(v for v in c_vals if v > 0) * 1.3 if any(v > 0 for v in c_vals) else 500
            fig_c = go.Figure()
            fig_c.add_trace(go.Bar(x=nms, y=c_vals, marker_color=bar_clrs,
                                   text=[f"{v:,.0f}만원/평" for v in c_vals], textposition="outside", width=0.35))
            fig_c.update_layout(height=340, yaxis_title="평당공사비 (만원/평)",
                                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                                yaxis=dict(gridcolor="#e8e8e8", range=[0, c_max], tickformat=","), showlegend=False)
            st.plotly_chart(fig_c, use_container_width=True)

    # ══════════════════════════════════════════
    #  탭 2: 건설공사비 지수 반영
    # ══════════════════════════════════════════
    with tabs[1]:
        if sim is not None and len(sim) > 0:
            fs = sim.iloc[0]
            cn = str(fs.get("프로젝트명", "비교 PJT"))
            cy = int(fs.get("착공연도", 2022)) if pd.notna(fs.get("착공연도")) else 2022
            ce = fs.get("공사금액_억", 0); ce = ce if pd.notna(ce) and ce > 0 else R["total_eok"]
            cp = fs.get("연면적(㎡)", 0)
            cp = cp / 3.3058 if pd.notna(cp) and cp > 0 else R["area"] / 3.3058
        else:
            cn, cy, ce, cp = "비교 PJT", 2022, R["total_eok"], R["area"] / 3.3058
        render_index_tab(cn, cy, ce, cp, R)

    # ══════════════════════════════════════════
    #  탭 3: 유사현장 검색
    # ══════════════════════════════════════════
    with tabs[2]:
        if sim is not None and len(sim) > 0:
            top5 = sim.head(5)
            avg_sim = top5['유사도'].mean()
            n_rpt = sum(1 for _, r in sim.head(50).iterrows() if has_report(pkg, r.get('현장코드', '')))
            t_type = R.get('type', '(전체)')

            st.markdown(f"""
            <div class="kpi-row">
                <div class="kpi-card"><div class="icon">📊</div>
                    <div class="label">검색 결과</div><div class="value">{len(sim)}</div><div class="sub">건 (상위 50건 표시)</div></div>
                <div class="kpi-card" style="border-left-color:#00cc96;"><div class="icon">🎯</div>
                    <div class="label">Top 5 평균 유사도</div><div class="value">{avg_sim:.1f}</div><div class="sub">점</div></div>
                <div class="kpi-card" style="border-left-color:#ab63fa;"><div class="icon">📄</div>
                    <div class="label">보고서 보유</div><div class="value">{n_rpt}</div><div class="sub">건 (상위 50건 중)</div></div>
                <div class="kpi-card" style="border-left-color:#E60012;"><div class="icon">🏢</div>
                    <div class="label">검색 유형</div><div class="value" style="font-size:1.1rem;">{t_type}</div><div class="sub">연면적 {R['area']:,.0f}㎡</div></div>
            </div>""", unsafe_allow_html=True)

            # 결과 테이블
            st.markdown('<div class="section-title">📋 유사현장 검색 결과 (상위 50건)</div>', unsafe_allow_html=True)
            display_data = []
            for _, row in sim.head(50).iterrows():
                code = str(row.get('현장코드', '')).strip()
                d = {}
                for col, disp in DISPLAY_FIELDS:
                    d[disp] = fmt(row.get(col), col)
                d['보고서'] = '✓' if has_report(pkg, code) else '-'
                d['유사도'] = f"{row.get('유사도', 0):.1f}"
                display_data.append(d)
            if display_data:
                disp_df = pd.DataFrame(display_data)
                st.dataframe(disp_df, use_container_width=True, hide_index=True, height=600)

            # Top 5 상세 카드
            st.markdown('<div class="section-title">🏆 유사현장 Top 5 상세</div>', unsafe_allow_html=True)
            for i, (_, row) in enumerate(top5.iterrows()):
                medal = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣"][i]
                bcolor = '#FFD700' if i == 0 else '#C0C0C0' if i == 1 else '#CD7F32' if i == 2 else '#666'
                st.markdown(f"""
                <div class="report-card" style="border-left:5px solid {bcolor};">
                    <div style="display:flex;align-items:center;gap:10px;">
                        <span style="font-size:1.5rem;">{medal}</span>
                        <h4 style="margin:0;color:#1a1a2e;">{row.get('프로젝트명','-')}
                        <span style="font-size:0.8rem;color:#003D7A;font-weight:600;margin-left:8px;">유사도 {row.get('유사도',0):.1f}점</span></h4>
                    </div>
                </div>""", unsafe_allow_html=True)
                m1, m2, m3, m4 = st.columns(4)
                r_area = row.get('연면적(㎡)', 0)
                if pd.isna(r_area): r_area = 0
                m1.metric("연면적", fmt_area(r_area))
                u = row.get('지하층', 0); a = row.get('지상층', 0)
                u = int(u) if pd.notna(u) else 0; a = int(a) if pd.notna(a) else 0
                m2.metric("규모", f"B{u}F / {a}F")
                eok = row.get('공사금액_억', 0)
                m3.metric("공사금액", fmt_cost(eok) if pd.notna(eok) and eok else '-')
                mon = row.get('공사개월', 0)
                m4.metric("공사기간", fmt_months(mon))

                # 보고서 열람
                sc = str(row.get('현장코드', '')).strip()
                if pkg and pkg.is_loaded() and has_report(pkg, sc):
                    total_p = pkg.get_page_count(sc)
                    orig_name = pkg.get_original_name(sc) or row.get('프로젝트명', '')
                    with st.expander(f"📄 준공보고서 열람 — {orig_name} ({total_p}페이지)", expanded=False):
                        pc1, pc2 = st.columns([3, 1])
                        with pc1:
                            page_no = st.slider("페이지", 1, max(1, total_p), 1, key=f"sim_page_{i}_{sc}")
                        with pc2:
                            st.markdown(f"<div style='text-align:center;padding-top:18px;font-size:0.85rem;color:#003D7A;font-weight:700;'>{page_no} / {total_p}</div>", unsafe_allow_html=True)
                        img = pkg.get_page_image(sc, page_no)
                        if img and HAS_PIL:
                            img = add_watermark(img)
                            st.image(img, caption=f"{orig_name} — 페이지 {page_no}/{total_p}", use_container_width=True)
                            st.caption("🔒 본 자료는 CJ대한통운 건설부문 내부 자료입니다. 캡처/복제/외부 유출 금지.")
                        elif img:
                            st.image(img, use_container_width=True)
                        else:
                            st.warning("페이지 이미지를 로드할 수 없습니다.")
                st.divider()

    # ══════════════════════════════════════════
    #  탭 4: 키워드 본문 검색
    # ══════════════════════════════════════════
    with tabs[3]:
        st.markdown('<div class="section-title">📖 키워드 본문 검색</div>', unsafe_allow_html=True)
        if pkg and pkg.is_loaded():
            st.markdown("""
            <div class="info-box">
                <b style="color:#003D7A;">보고서 본문 검색</b>
                <div style="color:#555;font-size:0.82rem;margin-top:4px;">
                    보고서 패키지 내 텍스트를 키워드로 검색합니다. 예: 화강석, 클린룸, GMP, 커튼월
                </div>
            </div>""", unsafe_allow_html=True)
            kw = st.text_input("🔎 키워드 입력", key="kw_text_search", placeholder="검색어를 입력하세요")
            if kw:
                results = pkg.search_text(kw)
                if results:
                    st.success(f"'{kw}' 검색 결과: {len(results)}건")
                    res_data = []
                    for r in results[:100]:
                        proj_name = r.get('original_name', r['site_code'])
                        res_data.append({
                            '현장코드': r['site_code'], '프로젝트명': proj_name,
                            '페이지': r['page'], '본문 스니펫': r['snippet'],
                        })
                    st.dataframe(pd.DataFrame(res_data), use_container_width=True, hide_index=True, height=400)
                    st.markdown('<div class="section-title">📄 검색 결과 상세 열람</div>', unsafe_allow_html=True)
                    for idx_r, r in enumerate(results[:30]):
                        proj_name = r.get('original_name', r['site_code'])
                        with st.expander(f"📄 {proj_name} — 페이지 {r['page']}"):
                            st.markdown(f"**스니펫:** {r['snippet']}")
                            if st.button(f"이 페이지 보기", key=f"kw_view_{r['site_code']}_{r['page']}_{idx_r}"):
                                img = pkg.get_page_image(r['site_code'], r['page'])
                                if img and HAS_PIL:
                                    img = add_watermark(img)
                                    st.image(img, use_container_width=True)
                                    st.caption("🔒 CJ대한통운 건설부문 내부 자료.")
                                elif img:
                                    st.image(img, use_container_width=True)
                                else:
                                    st.warning("이미지 로드 실패")
                else:
                    st.warning(f"'{kw}'에 대한 검색 결과가 없습니다.")
        else:
            st.warning("📦 보고서 패키지가 로드되지 않았습니다. reports.pkg 파일이 필요합니다.")

    # ══════════════════════════════════════════
    #  탭 5: 준공보고서 검색
    # ══════════════════════════════════════════
    with tabs[4]:
        st.markdown('<div class="section-title">📋 준공보고서 검색</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class="info-box">
            <b style="color:#003D7A;">메타데이터 기반 검색</b>
            <div style="color:#555;font-size:0.82rem;margin-top:4px;">
                프로젝트 속성(건축물종류, 착공연도, 발주처 등)으로 보고서를 검색합니다.
            </div>
        </div>""", unsafe_allow_html=True)

        sc1, sc2, sc3 = st.columns([1, 2, 1])
        with sc1:
            display_names = [FIELD_DISPLAY_NAME.get(c, c) for c in REPORT_SEARCH_FIELDS]
            search_field_disp = st.selectbox("검색 항목", display_names, key="rpt_field")
            disp_to_col = {FIELD_DISPLAY_NAME.get(c, c): c for c in REPORT_SEARCH_FIELDS}
            search_col = disp_to_col.get(search_field_disp, search_field_disp)
        with sc2:
            if search_col in df_all.columns:
                vals = df_all[search_col].astype(str).str.strip()
                uniques = sorted({
                    v for v in vals.tolist()
                    if v and v.lower() not in ('nan', 'none', 'nat', '')
                }, key=str)
            else:
                uniques = []
            search_value = st.selectbox("검색어 (선택 또는 입력)", [''] + uniques, key="rpt_value")
        with sc3:
            only_with_report = st.checkbox("보고서 보유만", value=True, key="rpt_only_report")
            run_rpt_search = st.button("🔍 검색", key="rpt_search_btn", type="primary")

        if run_rpt_search and search_value:
            series = df_all[search_col].astype(str).str.strip() if search_col in df_all.columns else pd.Series()
            mask = series.str.contains(search_value.strip(), case=False, na=False, regex=False)
            result_rpt = df_all[mask].copy()
            n_total = len(result_rpt)
            n_rpt_found = 0
            if pkg and pkg.is_loaded() and n_total > 0:
                has_mask = result_rpt['현장코드'].astype(str).apply(lambda x: has_report(pkg, x))
                n_rpt_found = int(has_mask.sum())
                if only_with_report:
                    result_rpt = result_rpt[has_mask]

            st.markdown(f"""
            <div class="kpi-row">
                <div class="kpi-card"><div class="icon">📊</div>
                    <div class="label">전체 일치</div><div class="value">{n_total}</div><div class="sub">건</div></div>
                <div class="kpi-card" style="border-left-color:#00cc96;"><div class="icon">📄</div>
                    <div class="label">보고서 보유</div><div class="value">{n_rpt_found}</div><div class="sub">건</div></div>
                <div class="kpi-card" style="border-left-color:#ab63fa;"><div class="icon">👁️</div>
                    <div class="label">화면 표시</div><div class="value">{len(result_rpt)}</div><div class="sub">건</div></div>
            </div>""", unsafe_allow_html=True)

            if len(result_rpt) > 0:
                disp_data = []
                for _, row in result_rpt.iterrows():
                    code = str(row.get('현장코드', '')).strip()
                    d = {}
                    for col, disp in DISPLAY_FIELDS:
                        d[disp] = fmt(row.get(col), col)
                    d['보고서'] = '✓' if has_report(pkg, code) else '-'
                    disp_data.append(d)
                st.dataframe(pd.DataFrame(disp_data), use_container_width=True, hide_index=True, height=500)

                if pkg and pkg.is_loaded():
                    rpt_codes = [str(row.get('현장코드', '')).strip()
                                 for _, row in result_rpt.iterrows()
                                 if has_report(pkg, row.get('현장코드', ''))]
                    if rpt_codes:
                        st.markdown('<div class="section-title">📄 보고서 열람</div>', unsafe_allow_html=True)
                        code_name_map = {}
                        for _, row in result_rpt.iterrows():
                            c = str(row.get('현장코드', '')).strip()
                            if c in rpt_codes:
                                code_name_map[c] = f"{c} - {row.get('프로젝트명', '')}"
                        sel_code = st.selectbox("열람할 현장 선택", rpt_codes,
                                                format_func=lambda x: code_name_map.get(x, x), key="rpt_viewer_code")
                        if sel_code:
                            total_p = pkg.get_page_count(sel_code)
                            if total_p > 0:
                                page_no = st.number_input("페이지", 1, total_p, 1, key="rpt_viewer_page")
                                img = pkg.get_page_image(sel_code, page_no)
                                if img and HAS_PIL:
                                    img = add_watermark(img)
                                    st.image(img, caption=f"페이지 {page_no}/{total_p}", use_container_width=True)
                                    st.caption("🔒 CJ대한통운 건설부문 내부 자료입니다.")
                                elif img:
                                    st.image(img, use_container_width=True)
            else:
                if n_total == 0:
                    st.info(f"'{search_col}'에서 '{search_value}'와 일치하는 항목이 없습니다.")
                else:
                    st.info(f"일치 항목 {n_total}건 중 보고서 보유 항목이 없습니다.")
        elif run_rpt_search and not search_value:
            st.warning("검색어를 선택하거나 입력하세요.")

        if pkg and pkg.is_loaded():
            with st.expander("📚 보고서 보유 프로젝트 전체 목록", expanded=False):
                rpt_projects = []
                for _, row in df_all.iterrows():
                    sc = str(row.get('현장코드', '')).strip()
                    if has_report(pkg, sc):
                        rpt_projects.append({
                            '현장코드': sc, '프로젝트명': row.get('프로젝트명', ''),
                            '건축물종류': row.get('건축물종류', ''),
                            '착공연도': fmt(row.get('착공연도'), '착공연도'),
                            '페이지수': pkg.get_page_count(sc),
                        })
                if rpt_projects:
                    st.dataframe(pd.DataFrame(rpt_projects), use_container_width=True, hide_index=True)
                else:
                    st.info("보고서 보유 프로젝트가 없습니다.")

    # ══════════════════════════════════════════
    #  탭 6: 보고서 열람
    # ══════════════════════════════════════════
    with tabs[5]:
        st.markdown('<div class="section-title">📑 보고서 열람</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class="info-box">
            <b style="color:#003D7A;">준공보고서 전체 열람</b>
            <div style="color:#555;font-size:0.82rem;margin-top:4px;">
                현장을 선택하면 해당 준공보고서의 전체 페이지를 열람할 수 있습니다.
            </div>
        </div>""", unsafe_allow_html=True)

        if pkg and pkg.is_loaded() and pkg.manifest:
            # ── 보고서 보유 현장 목록 구성 ──
            _report_codes = pkg.get_all_report_codes()
            _viewer_options = {}
            for _rc in sorted(_report_codes, key=str):
                _orig = pkg.get_original_name(_rc) or ''
                _pcount = pkg.get_page_count(_rc)
                # data.xlsx에서 프로젝트명 조회
                _pname = ''
                try:
                    _match = df_all[df_all['현장코드'].astype(str) == str(_rc)]
                    if not _match.empty:
                        _pname = str(_match.iloc[0].get('프로젝트명', '')).strip()
                        if _pname.lower() in ('nan', 'none', ''):
                            _pname = ''
                except Exception:
                    pass
                _label = f"[{_rc}] {_pname or _orig} ({_pcount}p)"
                _viewer_options[_label] = _rc

            if _viewer_options:
                _selected_label = st.selectbox(
                    "📂 현장 선택", list(_viewer_options.keys()),
                    key="viewer_site_select"
                )
                _selected_code = _viewer_options[_selected_label]
                _total_pages = pkg.get_page_count(_selected_code)
                _orig_name = pkg.get_original_name(_selected_code) or _selected_code

                st.markdown(f"""
                <div style="background:#f0f4f8;padding:12px 16px;border-radius:8px;
                            border-left:4px solid #003D7A;margin:8px 0 16px 0;">
                    <b style="color:#003D7A;">📄 {_orig_name}</b>
                    <span style="color:#555;font-size:0.85rem;margin-left:12px;">
                        총 {_total_pages}페이지
                    </span>
                </div>""", unsafe_allow_html=True)

                # ── 페이지 네비게이션 ──
                _nav_c1, _nav_c2, _nav_c3 = st.columns([1, 2, 1])
                with _nav_c2:
                    _view_page = st.slider(
                        "페이지 선택", 1, max(1, _total_pages), 1,
                        key="viewer_page_slider"
                    )

                # ── 선택한 페이지 표시 ──
                _page_img = pkg.get_page_image(_selected_code, _view_page)
                if _page_img and HAS_PIL:
                    _page_img = add_watermark(_page_img)
                    st.image(_page_img, caption=f"페이지 {_view_page} / {_total_pages}",
                             use_container_width=True)
                elif _page_img:
                    st.image(_page_img, caption=f"페이지 {_view_page} / {_total_pages}",
                             use_container_width=True)
                else:
                    st.warning(f"페이지 {_view_page} 이미지를 불러올 수 없습니다.")

                st.caption("🔒 CJ대한통운 건설부문 내부 자료. 무단 배포·복제 금지.")

                # ── 전체 페이지 일괄 보기 (접기) ──
                with st.expander("📖 전체 페이지 일괄 보기", expanded=False):
                    st.info(f"총 {_total_pages}페이지를 순서대로 표시합니다. 페이지가 많을 경우 로딩이 지연될 수 있습니다.")
                    for _p in range(1, _total_pages + 1):
                        _p_img = pkg.get_page_image(_selected_code, _p)
                        if _p_img and HAS_PIL:
                            _p_img = add_watermark(_p_img)
                            st.image(_p_img, caption=f"페이지 {_p} / {_total_pages}",
                                     use_container_width=True)
                        elif _p_img:
                            st.image(_p_img, caption=f"페이지 {_p} / {_total_pages}",
                                     use_container_width=True)
                        else:
                            st.warning(f"페이지 {_p} 로드 실패")
            else:
                st.info("패키지에 수록된 보고서가 없습니다.")
        else:
            st.warning("📦 보고서 패키지가 로드되지 않았습니다. reports.pkg 파일이 필요합니다.")

    # ══════════════════════════════════════════
    #  탭 7: 보고서 관리
    # ══════════════════════════════════════════
    with tabs[6]:
        st.markdown('<div class="section-title">📥 보고서 관리</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class="info-box">
            <b style="color:#003D7A;">보고서 추가</b>
            <div style="color:#555;font-size:0.82rem;margin-top:4px;">
                PDF/PPT 보고서 파일을 업로드하면 자동으로 [현장코드_건축물종류_현장제목_착공연도] 형식으로 이름이 변경되어 패키지에 추가됩니다.
            </div>
        </div>""", unsafe_allow_html=True)

        uploaded_files = st.file_uploader("보고서 파일 업로드 (PDF/PPT)", type=['pdf', 'pptx', 'ppt'],
                                          accept_multiple_files=True, key="report_upload")
        if uploaded_files:
            st.markdown('<div class="section-title">📝 보고서 정보 입력</div>', unsafe_allow_html=True)

            # ── 건축물종류 목록 (Excel 데이터 + Reports 폴더 파일명 통합) ──
            _bldg_from_data = set()
            for v in df_all['건축물종류'].dropna().astype(str).str.strip().unique():
                if v and v != '-' and v.lower() not in ('nan', 'none'):
                    _bldg_from_data.add(v)
            # Reports 폴더 파일명에서 건축물종류 추출 ({현장코드}_{건축물종류}_...)
            _bldg_from_reports = set()
            if REPORTS_DIR and os.path.isdir(REPORTS_DIR):
                import re as _re
                for _fn in os.listdir(REPORTS_DIR):
                    _m = _re.match(r'^\d+_([^_]+)_', _fn)
                    if _m:
                        _bldg_from_reports.add(_m.group(1))
            _bldg_type_opts = sorted(_bldg_from_data | _bldg_from_reports, key=str)
            _bldg_type_opts_sel = _bldg_type_opts + ["── 직접 입력 ──"]

            # ── 착공연도 목록 (프로그램 실행 연도 기준 자동 확장) ──
            _current_year = datetime.now().year
            _year_options = list(range(_current_year, 2009, -1))  # 최신연도 ~ 2010

            for idx_f, uf in enumerate(uploaded_files):
                with st.expander(f"📄 {uf.name}", expanded=True):
                    c1, c2 = st.columns(2)
                    with c1:
                        site_code_input = st.text_input(
                            "현장코드 *", key=f"code_{idx_f}",
                            placeholder="예: 512100 (숫자만 입력)"
                        )
                        bldg_select = st.selectbox(
                            "건축물종류 *", _bldg_type_opts_sel,
                            key=f"bldg_{idx_f}"
                        )
                        # 직접 입력 선택 시 텍스트 입력 표시
                        if bldg_select == "── 직접 입력 ──":
                            bldg_type = st.text_input(
                                "건축물종류 직접 입력 *", key=f"bldg_custom_{idx_f}",
                                placeholder="예: 데이터센터"
                            )
                        else:
                            bldg_type = bldg_select
                    with c2:
                        title = st.text_input(
                            "현장제목 *", key=f"title_{idx_f}",
                            placeholder="프로젝트명"
                        )
                        year_input = st.selectbox(
                            "착공연도 *", _year_options,
                            key=f"year_{idx_f}"
                        )

                    # ── 현장코드 자동 탐색 ──
                    if site_code_input and site_code_input.strip().isdigit():
                        try:
                            match = df_all[df_all['현장코드'].astype(str) == str(int(site_code_input.strip()))]
                            if not match.empty:
                                r = match.iloc[0]
                                auto_info = []
                                bt = str(r.get('건축물종류', '')).strip()
                                pn = str(r.get('프로젝트명', '')).strip()
                                yr = fmt(r.get('착공연도'), '착공연도')
                                if bt and bt.lower() not in ('nan', 'none'):
                                    auto_info.append(f"건축물종류: {bt}")
                                if pn and pn.lower() not in ('nan', 'none'):
                                    auto_info.append(f"프로젝트명: {pn}")
                                if yr and yr != '-':
                                    auto_info.append(f"착공연도: {yr}")
                                if auto_info:
                                    st.caption("💡 자동 탐색: " + " | ".join(auto_info))
                        except Exception:
                            pass

                    # ── 변환 파일명 미리보기 (항상 .pdf 확장자) ──
                    if site_code_input and bldg_type and title and year_input:
                        new_name = f"{site_code_input.strip()}_{bldg_type}_{title}_{year_input}.pdf"
                        st.info(f"변환 파일명: **{new_name}**")

            st.divider()
            if st.button("🔄 패키지에 추가 (리빌드)", type="primary", use_container_width=True):
                if not REPORTS_DIR:
                    st.error("Reports 폴더를 찾을 수 없습니다.")
                else:
                    added = []
                    for idx_f, uf in enumerate(uploaded_files):
                        sc_v = st.session_state.get(f"code_{idx_f}", "").strip()
                        # 건축물종류: selectbox 값 또는 직접 입력 값
                        bt_sel = st.session_state.get(f"bldg_{idx_f}", "")
                        if bt_sel == "── 직접 입력 ──":
                            bt_v = st.session_state.get(f"bldg_custom_{idx_f}", "").strip()
                        else:
                            bt_v = str(bt_sel).strip()
                        tl_v = st.session_state.get(f"title_{idx_f}", "").strip()
                        yr_v = str(st.session_state.get(f"year_{idx_f}", "")).strip()
                        if not all([sc_v, bt_v, tl_v, yr_v]):
                            st.warning(f"⚠️ {uf.name}: 필수 정보 미입력")
                            continue
                        if not sc_v.isdigit():
                            st.warning(f"⚠️ {uf.name}: 현장코드는 숫자여야 합니다")
                            continue

                        # ── 파일을 PDF로 변환/저장 ──
                        new_name = f"{sc_v}_{bt_v}_{tl_v}_{yr_v}.pdf"
                        dst = Path(REPORTS_DIR) / new_name
                        orig_ext = Path(uf.name).suffix.lower()

                        try:
                            if orig_ext in ('.pptx', '.ppt'):
                                # PPT/PPTX → PDF 변환 (PowerPoint COM)
                                import tempfile
                                with tempfile.NamedTemporaryFile(
                                    suffix=orig_ext, delete=False
                                ) as tmp:
                                    tmp.write(uf.getvalue())
                                    tmp_path = tmp.name
                                try:
                                    import pythoncom
                                    import win32com.client
                                    pythoncom.CoInitialize()
                                    powerpoint = win32com.client.Dispatch(
                                        "PowerPoint.Application"
                                    )
                                    powerpoint.Visible = 1
                                    deck = powerpoint.Presentations.Open(
                                        os.path.abspath(tmp_path),
                                        WithWindow=False
                                    )
                                    deck.SaveAs(
                                        os.path.abspath(str(dst)),
                                        FileFormat=32  # ppSaveAsPDF
                                    )
                                    deck.Close()
                                    powerpoint.Quit()
                                    pythoncom.CoUninitialize()
                                finally:
                                    if os.path.exists(tmp_path):
                                        os.unlink(tmp_path)
                            else:
                                # PDF 파일: 그대로 저장
                                with open(dst, 'wb') as f:
                                    f.write(uf.getvalue())

                            added.append(sc_v)
                            st.success(f"✓ {uf.name} → {new_name}")
                        except Exception as e:
                            st.error(f"✗ {uf.name} 저장/변환 실패: {e}")

                    if added:
                        st.info(
                            f"📦 {len(added)}건 파일 저장 완료. "
                            f"패키지 리빌드를 시작합니다..."
                        )
                        try:
                            # PyMuPDF(fitz) 미설치 시 자동 설치
                            try:
                                import fitz  # noqa: F401
                            except ImportError:
                                import subprocess
                                st.info("📥 PyMuPDF 패키지 설치 중...")
                                subprocess.check_call(
                                    [sys.executable, "-m", "pip", "install",
                                     "PyMuPDF", "--quiet"]
                                )
                            sys.path.insert(0, SRC_DIR)
                            from add_reports import add_reports
                            with st.spinner(
                                f"패키지 리빌드 중... ({len(added)}건)"
                            ):
                                add_reports(added)
                            st.success(
                                f"✅ 패키지 업데이트 완료! "
                                f"{len(added)}건 추가/갱신"
                            )
                            # 패키지 로더 갱신 → 키워드/보고서 검색 탭 자동 반영
                            st.session_state['pkg_loader'] = (
                                SecurePackageLoader(PACKAGE_FILE)
                            )
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ 패키지 리빌드 실패: {e}")
                            import traceback
                            st.code(traceback.format_exc())

        st.markdown('<div class="section-title">📦 패키지 현황</div>', unsafe_allow_html=True)
        if pkg and pkg.is_loaded() and pkg.manifest:
            total_reports = len(pkg.manifest.get('reports', {}))
            st.markdown(f"""
            <div class="kpi-row">
                <div class="kpi-card"><div class="icon">📦</div>
                    <div class="label">패키지 파일</div><div class="value" style="font-size:1rem;">reports.pkg</div>
                    <div class="sub">{os.path.getsize(PACKAGE_FILE)/1e9:.1f} GB</div></div>
                <div class="kpi-card" style="border-left-color:#00cc96;"><div class="icon">📄</div>
                    <div class="label">수록 보고서</div><div class="value">{total_reports}</div><div class="sub">건</div></div>
            </div>""", unsafe_allow_html=True)
            with st.expander("수록 보고서 목록"):
                pkg_list = []
                for code, info in pkg.manifest['reports'].items():
                    pkg_list.append({
                        '현장코드': code, '원본파일명': info.get('original_name', ''),
                        '페이지수': info.get('page_count', 0),
                    })
                st.dataframe(pd.DataFrame(pkg_list), use_container_width=True, hide_index=True)
        else:
            st.warning("보고서 패키지가 로드되지 않았습니다.")

    # ══════════════════════════════════════════
    #  탭 8: 모델 설명
    # ══════════════════════════════════════════
    with tabs[7]:
        st.markdown('<div class="section-title">🔬 모델 구조 및 변수 설명</div>', unsafe_allow_html=True)

        # 유사도 로직 설명
        with st.expander("**유사도 로직 (6요소 통합 — 구조체 수정 반영)**", expanded=True):
            st.markdown("""
| 요소 | 가중치 | 정규화 방법 |
|:---|:---:|:---|
| **건축물종류** | 35% | 완전일치=1.0 / 유사그룹=0.6 / 불일치=0.0 (14종 매핑) |
| **연면적** | 25% | log 비율: 1 − |log(T) − log(V)| / log_range |
| **지하층** | 15% | 1 − |T − V| / 8 (데이터 범위 기반) |
| **지상층** | 10% | 1 − |T − V| / 31 (데이터 범위 기반) |
| **역타공법** | 10% | 동일=1.0 / 1단계차=0.5 / 2단계차=0.0 |
| **구조체** | 5% | '지하 구조'·'지상 구조' 컬럼 기반. 일치=1.0 / RC↔SRC=0.7 / 불일치=0.0 |

> ⚠️ **수정사항**: 기존 CJ_EstAndPer에서는 '구조체' 컬럼(실제 순타/역타 값)을 RC/SRC/S/PC와 비교하는 버그가 있었습니다.
> 본 버전에서는 '지하 구조'·'지상 구조' 컬럼의 실제 구조 데이터(SRC, RC, PC, S)를 사용합니다.
""")

        # 회귀 모델 설명
        with st.expander("**회귀 모델 (개선: 착공연도 추가)**", expanded=True):
            st.markdown("""
> **개선사항**: 착공연도를 회귀변수에 추가하여 시간 경과에 따른 공사비 상승 트렌드를 반영합니다.
>
> - 업무 평당공사비 R²: **0.18 → 0.56** (착공연도 추가 효과 +38%p)
> - 업무 공사기간 R²: **0.56 → 0.69** (착공연도 추가 효과 +13%p)
""")

        for key, info in models.items():
            n = info.get("n", 0); r2 = info.get("r2", 0)
            with st.expander(f"**{key}** (N={n}, R²={r2:.3f})"):
                col1, col2, col3 = st.columns(3)
                col1.metric("R²", f"{r2:.4f}")
                col2.metric("MAE", f"{info.get('mae', 0):.2f}")
                col3.metric("학습 N", f"{n}건")
                coef = info.get("coef", {})
                cdf = pd.DataFrame([(f, c) for f, c in coef.items()], columns=["변수", "계수(β)"])
                st.dataframe(cdf, use_container_width=True, hide_index=True)

        if not models:
            st.warning("scikit-learn이 설치되지 않아 회귀 모델을 구축할 수 없습니다. `pip install scikit-learn`")

        st.markdown("""
        <div style="background:#f8f9ff;border-radius:10px;padding:14px 18px;margin-top:12px;font-size:0.8rem;color:#888;">
            📌 <b>데이터</b>: CJ대한통운 건설부문 PJT 이력 (data.xlsx)<br>
            📊 <b>건설공사비지수</b>: KICT 한국건설기술연구원 (2020=100, 비주거용건물)<br>
            📡 <b>보조출처</b>: CERIK 월간건설시장동향, KOSIS 국가통계포털<br>
            🔧 <b>모델 개선</b>: 착공연도 변수 추가, 구조체 컬럼 매핑 수정 (v2.1)
        </div>""", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
