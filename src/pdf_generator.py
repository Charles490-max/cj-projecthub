# src/pdf_generator.py
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth
from datetime import datetime

from config import COLOR_PRIMARY, COLOR_ACCENT, DISPLAY_FIELDS, EMPTY_MARK, fmt

# ============================================================
# 한글 폰트 등록
# ============================================================
try:
    pdfmetrics.registerFont(TTFont('Malgun', 'C:/Windows/Fonts/malgun.ttf'))
    pdfmetrics.registerFont(TTFont('MalgunBold', 'C:/Windows/Fonts/malgunbd.ttf'))
    FONT, FONT_B = 'Malgun', 'MalgunBold'
except Exception:
    FONT = FONT_B = 'Helvetica'

# 카드에 표시할 필드 (현장코드/프로젝트명은 헤더에 별도 표시)
CARD_FIELDS = [(c, d) for c, d in DISPLAY_FIELDS if c not in ('현장코드', '프로젝트명')]


# ============================================================
# 유틸
# ============================================================
def _hex(c):
    c = c.lstrip('#')
    return tuple(int(c[i:i+2], 16) / 255 for i in (0, 2, 4))


def _truncate(text, font, size, max_width_mm):
    """문자열이 max_width_mm 를 초과하면 말줄임표(…) 처리."""
    if text is None:
        return EMPTY_MARK
    s = str(text)
    max_w = max_width_mm * mm
    if stringWidth(s, font, size) <= max_w:
        return s
    ell = '…'
    while s and stringWidth(s + ell, font, size) > max_w:
        s = s[:-1]
    return (s + ell) if s else ell


def _default_has_report(_code):
    """has_report_func 미전달 시 사용되는 기본값 (모두 미보유 처리)."""
    return False


# ============================================================
# 메인 함수
# ============================================================
def generate_reference_pdf(top_df, target, output_path,
                           has_report_func=None, user=None, **_kwargs):
    """A4 1페이지 PDF: 검색조건 + TOP3 카드(18개 항목).

    Parameters
    ----------
    top_df : pd.DataFrame
        검색 결과 상위 N건 (TOP3만 사용됨). 비어 있어도 안전 처리.
    target : dict
        검색 조건 (건축물종류, 연면적(㎡), 지하층, 지상층 등).
    output_path : str | Path
        출력 PDF 경로.
    has_report_func : callable | None
        site_code -> bool 함수. 미전달 시 모두 False 로 처리.
    user : str | None
        하단 푸터에 표시할 발급자 ID (선택).
    """
    if has_report_func is None:
        has_report_func = _default_has_report
    target = target or {}

    c = canvas.Canvas(str(output_path), pagesize=A4)
    W, H = A4

    # ---------- 헤더 ----------
    c.setFillColorRGB(*_hex(COLOR_PRIMARY))
    c.rect(0, H - 25 * mm, W, 25 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont(FONT_B, 18)
    c.drawString(15 * mm, H - 15 * mm, "CJ ProjectHub")
    c.setFont(FONT, 10)
    c.drawString(15 * mm, H - 21 * mm, "준공 프로젝트 레퍼런스 보고서")
    c.setFont(FONT, 9)
    c.drawRightString(W - 15 * mm, H - 15 * mm,
                      datetime.now().strftime('%Y-%m-%d %H:%M'))

    # ---------- 검색 조건 ----------
    y = H - 32 * mm
    c.setFillColorRGB(0.95, 0.95, 0.95)
    c.rect(15 * mm, y - 10 * mm, W - 30 * mm, 10 * mm, fill=1, stroke=0)
    c.setFillColorRGB(0, 0, 0)
    c.setFont(FONT_B, 10)
    c.drawString(18 * mm, y - 6.5 * mm, "■ 검색 조건")
    c.setFont(FONT, 9)

    # target 값 안전 추출 (None/문자열 모두 대응)
    def _safe_num(v, default=0):
        try:
            return float(v)
        except (TypeError, ValueError):
            return default

    bld = target.get('건축물종류', '-') or '-'
    area = _safe_num(target.get('연면적(㎡)', 0), 0)
    under = _safe_num(target.get('지하층', 0), 0)
    above = _safe_num(target.get('지상층', 0), 0)
    cond = (f"건축물종류: {bld}    "
            f"연면적: {area:,.0f}㎡    "
            f"지하층: {under:.0f}    "
            f"지상층: {above:.0f}")
    c.drawString(60 * mm, y - 6.5 * mm, cond)

    # ---------- TOP3 카드 ----------
    card_top = y - 14 * mm
    card_bottom_limit = 15 * mm
    card_h = (card_top - card_bottom_limit) / 3

    # 상위 3건만 사용
    if top_df is not None and not top_df.empty:
        top3 = top_df.head(3)
        for idx, (_, row) in enumerate(top3.iterrows()):
            top = card_top - idx * card_h
            bot = top - card_h + 2 * mm
            _draw_card(c, row, idx + 1, top, bot, W, has_report_func)

    # ---------- 푸터 ----------
    c.setFont(FONT, 7)
    c.setFillColorRGB(0.5, 0.5, 0.5)
    footer_text = "본 자료는 CJ대한통운 건설부문 내부용입니다. 무단 배포·복제 금지."
    if user:
        footer_text = f"발급: {user}  |  " + footer_text
    c.drawCentredString(W / 2, 8 * mm, footer_text)

    c.save()
    return output_path


# ============================================================
# 카드 그리기
# ============================================================
def _draw_card(c, row, rank, top, bot, W, has_report_func):
    # ---------- 외곽선 ----------
    c.setStrokeColorRGB(*_hex(COLOR_PRIMARY))
    c.setLineWidth(0.8)
    c.rect(15 * mm, bot, W - 30 * mm, top - bot, fill=0, stroke=1)

    # ---------- 랭크 배지 ----------
    c.setFillColorRGB(*_hex(COLOR_ACCENT))
    c.rect(15 * mm, top - 7 * mm, 14 * mm, 7 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont(FONT_B, 11)
    c.drawCentredString(22 * mm, top - 5 * mm, f"#{rank}")

    # ---------- 프로젝트명 ----------
    code = fmt(row.get('현장코드'), '현장코드')
    pname = fmt(row.get('프로젝트명'), '프로젝트명')

    # 유사도: '유사도'(%) 우선, 없으면 'similarity'(0~1 가정 → ×100)
    if '유사도' in row.index:
        try:
            sim = float(row.get('유사도', 0))
        except Exception:
            sim = 0.0
    else:
        try:
            sim_raw = float(row.get('similarity', 0))
            sim = sim_raw * 100.0 if sim_raw <= 1.0 else sim_raw
        except Exception:
            sim = 0.0

    title_text = f"[{code}] {pname}"
    title_text = _truncate(title_text, FONT_B, 11, 120)
    c.setFillColorRGB(0, 0, 0)
    c.setFont(FONT_B, 11)
    c.drawString(32 * mm, top - 5 * mm, title_text)

    # ---------- 유사도 / 보고서 보유 ----------
    c.setFont(FONT, 9)
    c.setFillColorRGB(*_hex(COLOR_PRIMARY))
    c.drawRightString(W - 18 * mm, top - 5 * mm, f"유사도 {sim:.1f}%")

    has = ('✓ 보고서 보유'
           if has_report_func(str(row.get('현장코드', '')))
           else f'{EMPTY_MARK} 보고서 없음')
    c.setFont(FONT, 8)
    c.setFillColorRGB(0.3, 0.3, 0.3)
    c.drawRightString(W - 18 * mm, top - 9.5 * mm, has)

    # ---------- 항목 그리드 (2열 × 9행) ----------
    grid_top = top - 13 * mm
    inner_left = 18 * mm
    inner_right = W - 18 * mm
    inner_w = inner_right - inner_left
    n_cols = 2
    col_w = inner_w / n_cols
    row_h = 4.3 * mm

    LABEL_W_MM = 26
    GAP_MM = 2

    LABEL_FONT, LABEL_SIZE = FONT_B, 8
    VALUE_FONT, VALUE_SIZE = FONT, 8.5

    for i, (col_name, disp) in enumerate(CARD_FIELDS):
        col = i % n_cols
        rw = i // n_cols
        cx_label = inner_left + col * col_w
        cy = grid_top - rw * row_h

        # 라벨
        label_text = f"{disp} :"
        label_text = _truncate(label_text, LABEL_FONT, LABEL_SIZE, LABEL_W_MM - 1)
        c.setFont(LABEL_FONT, LABEL_SIZE)
        c.setFillColorRGB(*_hex(COLOR_PRIMARY))
        c.drawString(cx_label, cy, label_text)

        # 값
        cx_value = cx_label + LABEL_W_MM * mm + GAP_MM * mm
        value_max_mm = (col_w - LABEL_W_MM - GAP_MM - 2)
        v = fmt(row.get(col_name), col_name)
        v = _truncate(v, VALUE_FONT, VALUE_SIZE, value_max_mm)
        c.setFont(VALUE_FONT, VALUE_SIZE)
        c.setFillColorRGB(0, 0, 0)
        c.drawString(cx_value, cy, v)
