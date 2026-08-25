"""
변환된 PDF 1건이 정상인지 빠르게 검증
"""
import sys
from pathlib import Path
import fitz  # PyMuPDF

# 테스트할 PDF 파일 경로 (변환 완료된 것 1개)
TEST_PDF = Path(r"C:\Users\User\Desktop\CJ ProjectHub\Reports\변환된파일명.pdf")

print(f"검증 대상: {TEST_PDF.name}\n")

# 1. PDF 시그니처 확인
with open(TEST_PDF, 'rb') as f:
    sig = f.read(5)
print(f"[1] PDF 시그니처: {sig}")
print(f"    정상 여부: {'✓' if sig == b'%PDF-' else '✗ DRM 보호된 PDF'}")

# 2. PyMuPDF로 페이지 읽기 시도
try:
    doc = fitz.open(str(TEST_PDF))
    print(f"\n[2] 페이지 수: {doc.page_count}")
    
    page = doc[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
    
    test_output = TEST_PDF.parent.parent / 'output' / 'pdf_test_page1.png'
    test_output.parent.mkdir(exist_ok=True)
    pix.save(str(test_output))
    
    print(f"    ✓ 페이지 이미지 추출 성공")
    print(f"    → {test_output}")
    print(f"\n결과: 이 PDF는 CJ ProjectHub에서 정상 사용 가능합니다.")
    doc.close()
except Exception as e:
    print(f"\n[2] ✗ PDF 읽기 실패: {e}")
    print(f"    → DRM이 PDF에 상속되어 있을 가능성이 높습니다.")
