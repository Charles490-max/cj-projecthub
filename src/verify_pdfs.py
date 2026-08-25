"""
Reports 폴더의 모든 PDF 파일 일괄 검증
- PDF 시그니처 확인
- PyMuPDF로 페이지 읽기 테스트
- 첫 페이지 이미지 추출 가능 여부
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import fitz  # PyMuPDF
from config import REPORTS_ORIGINAL_DIR

print("=" * 65)
print("  PDF 일괄 검증")
print("=" * 65)

# Reports 폴더의 모든 PDF 수집
pdf_files = sorted(REPORTS_ORIGINAL_DIR.glob('*.pdf'))
print(f"\n  검증 대상: {len(pdf_files)}건\n")

ok_count = 0
fail_count = 0
failed_list = []

for idx, pdf_path in enumerate(pdf_files, 1):
    # 현장코드 추출
    site_code = pdf_path.stem.split('_')[0]
    
    # 1. PDF 시그니처 확인
    with open(pdf_path, 'rb') as f:
        sig = f.read(5)
    
    if sig != b'%PDF-':
        print(f"[{idx:>3}/{len(pdf_files)}] ✗ {pdf_path.name}")
        print(f"           DRM 보호된 PDF (시그니처: {sig})")
        fail_count += 1
        failed_list.append((pdf_path.name, "DRM 보호"))
        continue
    
    # 2. PyMuPDF 읽기 테스트
    try:
        doc = fitz.open(str(pdf_path))
        page_count = doc.page_count
        
        # 첫 페이지 픽셀 추출 시도
        if page_count > 0:
            page = doc[0]
            pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
            _ = pix.tobytes()
        
        doc.close()
        
        size_mb = pdf_path.stat().st_size / 1024 / 1024
        print(f"[{idx:>3}/{len(pdf_files)}] ✓ [{site_code}] "
              f"{page_count:>3}p / {size_mb:>5.1f}MB - {pdf_path.name[:50]}")
        ok_count += 1
    
    except Exception as e:
        print(f"[{idx:>3}/{len(pdf_files)}] ✗ {pdf_path.name}")
        print(f"           읽기 실패: {e}")
        fail_count += 1
        failed_list.append((pdf_path.name, str(e)))

# 결과 요약
print("\n" + "=" * 65)
print(f"  검증 결과")
print("=" * 65)
print(f"  ✓ 정상: {ok_count}건")
print(f"  ✗ 실패: {fail_count}건")

if failed_list:
    print(f"\n  실패 파일 목록:")
    for name, reason in failed_list:
        print(f"    - {name}")
        print(f"      사유: {reason}")
    print(f"\n  ※ 실패 파일은 재변환이 필요합니다.")
else:
    print(f"\n  → 모든 PDF가 정상입니다. 다음 단계로 진행하세요.")
