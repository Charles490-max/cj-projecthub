"""
PDF 기반 보고서가 정상 표시되는지 확인
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from config import PACKAGE_FILE
from package_loader import SecurePackageLoader

loader = SecurePackageLoader(PACKAGE_FILE)

# PDF 원본인 보고서들의 첫 페이지 점검
pdf_codes = ['502116', '502156', '502158', '502160', '502183', 
             '502224', '502261', '502262', '502316', '502333',
             '502413', '502417', '502425', '502435', '502445',
             '502507', '512011', '512012', '512045', '512083',
             '512099', '552200']

print("=" * 60)
print("  PDF 기반 보고서 점검")
print("=" * 60)

ok_count = 0
for code in pdf_codes:
    if not loader.has_report(code):
        continue
    
    page_count = loader.get_page_count(code)
    name = loader.get_original_name(code)
    
    # 첫 페이지 PNG 시그니처 확인
    for prefix in ['page_', 'slide_']:
        try:
            path = f"{code}/{prefix}001.png"
            with loader._zip.open(path) as f:
                data = f.read()
            
            is_png = data[:8] == b'\x89PNG\r\n\x1a\n'
            status = '✓ 정상' if is_png else '✗ 손상'
            
            print(f"\n  [{code}] {name}")
            print(f"    페이지: {page_count} | 첫 페이지: {status} ({len(data):,} bytes)")
            
            if is_png:
                ok_count += 1
            break
        except KeyError:
            continue

print(f"\n{'=' * 60}")
print(f"  PDF 출처 보고서 정상: {ok_count}건")
print("=" * 60)
