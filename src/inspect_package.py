"""
패키지 내부 이미지 파일 상태 점검
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from config import PACKAGE_FILE
from package_loader import SecurePackageLoader

loader = SecurePackageLoader(PACKAGE_FILE)

print("=" * 60)
print("  패키지 내부 이미지 점검")
print("=" * 60)

# 무신사 P1의 모든 페이지 파일 정보 확인
site_code = '502389'
print(f"\n[현장코드 {site_code}]")

for page_no in range(1, 6):  # 첫 5페이지만 점검
    for prefix in ['slide_', 'page_']:
        try:
            path = f"{site_code}/{prefix}{page_no:03d}.png"
            with loader._zip.open(path) as f:
                data = f.read()
            
            # 파일 크기와 시그니처 확인
            size = len(data)
            # PNG 시그니처: 89 50 4E 47 0D 0A 1A 0A
            is_png = data[:8] == b'\x89PNG\r\n\x1a\n' if size >= 8 else False
            
            print(f"  {path}")
            print(f"    크기: {size:,} bytes")
            print(f"    PNG 시그니처: {'✓' if is_png else '✗'}")
            print(f"    첫 16바이트: {data[:16].hex() if size >= 16 else data.hex()}")
            break
        except KeyError:
            continue

# 다른 보고서도 같은 문제인지 확인
print(f"\n[전체 보고서 첫 페이지 점검]")
broken = []
ok = []
for code in list(loader.manifest['reports'].keys())[:20]:
    for prefix in ['slide_', 'page_']:
        try:
            path = f"{code}/{prefix}001.png"
            with loader._zip.open(path) as f:
                data = f.read()
            is_png = data[:8] == b'\x89PNG\r\n\x1a\n' if len(data) >= 8 else False
            if is_png:
                ok.append(code)
            else:
                broken.append((code, len(data)))
            break
        except KeyError:
            continue

print(f"  정상 PNG: {len(ok)}건")
print(f"  손상/비정상: {len(broken)}건")
if broken:
    print(f"\n  손상 파일 (앞 10건):")
    for code, size in broken[:10]:
        print(f"    {code}: {size} bytes")
