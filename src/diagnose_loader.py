# src/diagnose_loader.py
"""
CJ ProjectHub - 패키지 로더 진단 스크립트
package_loader.py 와 reports.pkg 의 상태, 메서드 인터페이스를 점검한다.
"""
import sys
import traceback
from pathlib import Path

print("=" * 70)
print("CJ ProjectHub - 패키지 로더 진단")
print("=" * 70)

# ----------------------------------------------------------
# [1] config 확인
# ----------------------------------------------------------
try:
    from config import PACKAGE_FILE, MASTER_PASSWORD
    print(f"\n[1] config 로드: OK")
    print(f"    PACKAGE_FILE = {PACKAGE_FILE}")
    print(f"    파일 존재 여부: {PACKAGE_FILE.exists()}")
    if PACKAGE_FILE.exists():
        size_mb = PACKAGE_FILE.stat().st_size / 1024 / 1024
        print(f"    파일 크기: {size_mb:.1f} MB")
    print(f"    MASTER_PASSWORD 길이: {len(MASTER_PASSWORD)} 자")
except Exception as e:
    print(f"\n[1] config 로드 실패: {e}")
    traceback.print_exc()
    sys.exit(1)

# ----------------------------------------------------------
# [2] package_loader 모듈 import
# ----------------------------------------------------------
try:
    import package_loader
    print(f"\n[2] package_loader 모듈 import: OK")
    print(f"    파일 위치: {package_loader.__file__}")
    members = [m for m in dir(package_loader) if not m.startswith('_')]
    print(f"    공개 멤버: {members}")
except Exception as e:
    print(f"\n[2] package_loader import 실패: {e}")
    traceback.print_exc()
    sys.exit(1)

# ----------------------------------------------------------
# [3] SecurePackageLoader 클래스 확인
# ----------------------------------------------------------
try:
    cls = getattr(package_loader, 'SecurePackageLoader', None)
    if cls is None:
        # 다른 이름의 클래스가 있는지 탐색
        candidates = [name for name in dir(package_loader)
                      if not name.startswith('_')
                      and isinstance(getattr(package_loader, name), type)]
        print(f"\n[3] SecurePackageLoader 클래스 없음!")
        print(f"    package_loader.py 안의 클래스 후보: {candidates}")
        sys.exit(1)
    print(f"\n[3] SecurePackageLoader 클래스: OK")
    methods = [m for m in dir(cls)
               if not m.startswith('_') and callable(getattr(cls, m))]
    attrs = [m for m in dir(cls)
             if not m.startswith('_') and not callable(getattr(cls, m))]
    print(f"    메서드 목록: {methods}")
    print(f"    속성 목록: {attrs}")
except Exception as e:
    print(f"\n[3] 클래스 확인 실패: {e}")
    traceback.print_exc()
    sys.exit(1)

# ----------------------------------------------------------
# [4] 인스턴스화 (실제 패키지 로드)
# ----------------------------------------------------------
loader = None
try:
    loader = cls(PACKAGE_FILE)
    print(f"\n[4] 패키지 로드: OK")
    inst_attrs = [a for a in dir(loader)
                  if not a.startswith('_') and not callable(getattr(loader, a))]
    print(f"    인스턴스 속성: {inst_attrs}")
except Exception as e:
    print(f"\n[4] 패키지 로드 실패: {type(e).__name__}: {e}")
    traceback.print_exc()
    sys.exit(1)

# ----------------------------------------------------------
# [5] 보고서 목록 / 개수 확인
# ----------------------------------------------------------
keys = []
try:
    if hasattr(loader, 'list_reports'):
        result = loader.list_reports()
        keys = list(result) if result is not None else []
        print(f"\n[5] list_reports() 호출: OK")
        print(f"    보고서 수: {len(keys)}")
        if keys:
            print(f"    첫 5건: {keys[:5]}")
    elif hasattr(loader, 'reports'):
        r = loader.reports
        print(f"\n[5] loader.reports 속성: 타입={type(r).__name__}")
        if isinstance(r, dict):
            keys = list(r.keys())
            print(f"    개수: {len(keys)}")
            print(f"    첫 5건: {keys[:5]}")
            # 첫 항목 구조
            if keys:
                first_key = keys[0]
                first_val = r[first_key]
                print(f"    샘플 [{first_key}] 타입={type(first_val).__name__}")
                if isinstance(first_val, dict):
                    print(f"    샘플 필드: {list(first_val.keys())}")
        elif isinstance(r, list):
            keys = r
            print(f"    개수: {len(keys)}")
            print(f"    첫 5건: {keys[:5]}")
    elif hasattr(loader, 'manifest'):
        m = loader.manifest
        print(f"\n[5] loader.manifest 속성: 타입={type(m).__name__}")
        if isinstance(m, dict):
            print(f"    manifest 키: {list(m.keys())[:10]}")
            if 'reports' in m:
                keys = list(m['reports'].keys()) if isinstance(m['reports'], dict) else list(m['reports'])
                print(f"    보고서 수: {len(keys)}, 첫 5건: {keys[:5]}")
    else:
        print(f"\n[5] 보고서 목록 조회 가능한 속성/메서드를 찾을 수 없음")
except Exception as e:
    print(f"\n[5] 보고서 목록 조회 실패: {type(e).__name__}: {e}")
    traceback.print_exc()

# ----------------------------------------------------------
# [6] 메서드 인터페이스 점검 (secure_viewer 가 호출하는 메서드 위주)
# ----------------------------------------------------------
print(f"\n[6] secure_viewer.py 가 사용하는 메서드 존재 여부:")
target_methods = [
    'has_report',
    'get_page_count',
    'get_page_image',
    'get_project_name',
    'search_keyword',
    'get_pages',
    'get_image',
    'get_report',
]
for method in target_methods:
    exists = hasattr(loader, method)
    mark = '✓' if exists else '✗'
    print(f"    {mark} {method}")

# ----------------------------------------------------------
# [7] 첫 번째 보고서로 실제 호출 테스트
# ----------------------------------------------------------
if keys:
    site = str(keys[0])
    print(f"\n[7] 테스트 대상 site_code = {site}")

    # has_report
    if hasattr(loader, 'has_report'):
        try:
            r = loader.has_report(site)
            print(f"    has_report({site}) = {r}")
        except Exception as e:
            print(f"    has_report 실패: {e}")

    # get_page_count
    page_count = None
    if hasattr(loader, 'get_page_count'):
        try:
            page_count = loader.get_page_count(site)
            print(f"    get_page_count({site}) = {page_count}")
        except Exception as e:
            print(f"    get_page_count 실패: {e}")

    # get_page_image
    if hasattr(loader, 'get_page_image') and (page_count is None or page_count > 0):
        try:
            img = loader.get_page_image(site, 1)
            if img:
                size = len(img) if hasattr(img, '__len__') else '?'
                head = img[:8] if isinstance(img, (bytes, bytearray)) else 'N/A'
                print(f"    get_page_image({site}, 1) = {size} bytes, head={head}")
            else:
                print(f"    get_page_image({site}, 1) = 빈 데이터")
        except Exception as e:
            print(f"    get_page_image 실패: {e}")

    # get_project_name
    if hasattr(loader, 'get_project_name'):
        try:
            name = loader.get_project_name(site)
            print(f"    get_project_name({site}) = {name}")
        except Exception as e:
            print(f"    get_project_name 실패: {e}")
else:
    print(f"\n[7] 테스트할 보고서가 없음 (keys 비어있음)")

print("\n" + "=" * 70)
print("진단 완료 — 위 결과를 그대로 복사해 알려주세요")
print("=" * 70)
