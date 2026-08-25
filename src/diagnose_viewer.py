"""
뷰어 진단 스크립트
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from config import PACKAGE_FILE, REPORTS_ORIGINAL_DIR
from package_loader import SecurePackageLoader

print("=" * 60)
print("  CJ ProjectHub - 뷰어 진단")
print("=" * 60)

# 1. reports.pkg 존재 확인
print(f"\n[1] reports.pkg 파일")
print(f"    경로: {PACKAGE_FILE}")
print(f"    존재: {'✓' if PACKAGE_FILE.exists() else '✗ (빌드 필요)'}")

if not PACKAGE_FILE.exists():
    print("\n→ build_package.py를 먼저 실행하세요.")
    print("  cd src")
    print("  python build_package.py")
    sys.exit()

print(f"    크기: {PACKAGE_FILE.stat().st_size/1024/1024:.1f} MB")

# 2. 패키지 로드
print(f"\n[2] 패키지 로드 시도")
try:
    loader = SecurePackageLoader(PACKAGE_FILE)
    print(f"    ✓ 로드 성공")
    print(f"    보고서 수: {len(loader.manifest['reports'])}건")
except Exception as e:
    print(f"    ✗ 로드 실패: {e}")
    sys.exit()

# 3. 무신사 P1 (502389) 확인
print(f"\n[3] 무신사 P1 (현장코드 502389) 확인")
site_code = '502389'
if loader.has_report(site_code):
    page_count = loader.get_page_count(site_code)
    original_name = loader.get_original_name(site_code)
    print(f"    ✓ 보유 중")
    print(f"    원본: {original_name}")
    print(f"    페이지 수: {page_count}")
    
    # 첫 페이지 이미지 추출 시도
    img = loader.get_page_image(site_code, 1)
    if img is None:
        print(f"    ✗ 1페이지 이미지 로드 실패")
    else:
        print(f"    ✓ 1페이지 이미지 크기: {img.size}")
        # 테스트용 저장
        test_path = Path(__file__).parent.parent / 'output' / 'test_page1.png'
        img.save(test_path)
        print(f"    → 테스트 저장: {test_path}")
else:
    print(f"    ✗ 패키지에 없음")
    print(f"    빌드된 보고서 목록 (앞 10건):")
    for code in list(loader.manifest['reports'].keys())[:10]:
        print(f"      - {code}: {loader.manifest['reports'][code]['original_name']}")

# 4. 원본 PPT 확인
print(f"\n[4] 원본 Reports 폴더 점검")
musinsa_files = list(REPORTS_ORIGINAL_DIR.glob('502389_*'))
if musinsa_files:
    for f in musinsa_files:
        print(f"    ✓ {f.name} ({f.stat().st_size/1024/1024:.1f} MB)")
else:
    print(f"    ✗ 502389로 시작하는 파일 없음")

print("\n" + "=" * 60)
