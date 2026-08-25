# src/verify_reports.py
"""
Reports 폴더 <-> reports.pkg <-> data.xlsx 3자 정합성 검증
"""
import sys
import io
import re

# Windows 콘솔 인코딩 문제 해결: stdout을 UTF-8로
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from pathlib import Path
from config import REPORTS_ORIGINAL_DIR, PACKAGE_FILE, DATA_FILE
from data_loader import load_projects
from package_loader import SecurePackageLoader

print("=" * 75)
print(" 준공보고서 정합성 검증 (Reports 폴더 / Package / Excel)")
print("=" * 75)

# ──────────────────────────────────────────────────────────────
# [1] Reports 폴더 실제 파일
# ──────────────────────────────────────────────────────────────
print(f"\n[1] Reports 폴더: {REPORTS_ORIGINAL_DIR}")
if not REPORTS_ORIGINAL_DIR.exists():
    print("  [X] 폴더가 존재하지 않습니다.")
    raise SystemExit(1)

all_files = []
for ext in ('*.pdf', '*.PDF', '*.pptx', '*.PPTX', '*.ppt', '*.PPT'):
    all_files.extend(REPORTS_ORIGINAL_DIR.glob(ext))
for ext in ('*.pdf', '*.PDF', '*.pptx', '*.PPTX'):
    all_files.extend(REPORTS_ORIGINAL_DIR.rglob(ext))
all_files = list(set(all_files))
print(f"  총 파일 수: {len(all_files)}")

code_pattern = re.compile(r'(\d{6})')
file_codes = {}
no_code_files = []
for f in all_files:
    m = code_pattern.search(f.name)
    if m:
        code = m.group(1)
        file_codes.setdefault(code, []).append(f.name)
    else:
        no_code_files.append(f.name)

print(f"  6자리 코드 추출 성공: {len(file_codes)}개 코드 ({sum(len(v) for v in file_codes.values())}개 파일)")

if no_code_files:
    print(f"\n  [!] 코드 추출 실패 파일 {len(no_code_files)}개 (전체 목록):")
    for n in no_code_files:
        print(f"    - {n}")

dup = {c: fs for c, fs in file_codes.items() if len(fs) > 1}
if dup:
    print(f"\n  [!] 같은 코드에 파일 여러 개 ({len(dup)}개 코드):")
    for c, fs in dup.items():
        print(f"    {c}: {fs}")

# ──────────────────────────────────────────────────────────────
# [2] reports.pkg 패키지 내부
# ──────────────────────────────────────────────────────────────
print(f"\n[2] Package: {PACKAGE_FILE}")
loader = SecurePackageLoader(PACKAGE_FILE)
if hasattr(loader, 'reports'):
    pkg_codes = set(str(k) for k in loader.reports.keys())
else:
    pkg_codes = set()
print(f"  패키지 내부 보고서 키 개수: {len(pkg_codes)}")
print(f"  샘플 키(10개): {sorted(pkg_codes)[:10]}")

# ──────────────────────────────────────────────────────────────
# [3] data.xlsx 의 현장코드
# ──────────────────────────────────────────────────────────────
print(f"\n[3] Excel: {DATA_FILE}")
df = load_projects(DATA_FILE)
excel_codes = set(df['현장코드'].astype(str).str.strip().tolist())
excel_codes.discard('nan')
excel_codes.discard('')
print(f"  Excel 행 수: {len(df)}")
print(f"  고유 현장코드: {len(excel_codes)}")
print(f"  샘플 코드(10개): {sorted(excel_codes)[:10]}")

# ──────────────────────────────────────────────────────────────
# [4] 3자 교집합/차집합 분석
# ──────────────────────────────────────────────────────────────
folder_codes = set(file_codes.keys())

print("\n" + "=" * 75)
print(" [4] 정합성 분석")
print("=" * 75)
print(f"  Reports 폴더 코드          : {len(folder_codes):3d}개")
print(f"  Package 내부 코드          : {len(pkg_codes):3d}개")
print(f"  Excel 현장코드             : {len(excel_codes):3d}개")
print()
print(f"  폴더 AND Package          : {len(folder_codes & pkg_codes):3d}개")
print(f"  폴더 AND Excel            : {len(folder_codes & excel_codes):3d}개")
print(f"  Package AND Excel         : {len(pkg_codes & excel_codes):3d}개")
print(f"  폴더 AND Package AND Excel: {len(folder_codes & pkg_codes & excel_codes):3d}개  <- 시스템이 '보고서 보유'로 인식")

# ──────────────────────────────────────────────────────────────
# [5] 차집합(누락 위치 추적)
# ──────────────────────────────────────────────────────────────
print("\n" + "=" * 75)
print(" [5] 누락 원인 추적")
print("=" * 75)

in_folder_not_pkg = folder_codes - pkg_codes
print(f"\n  [A] 폴더에 있지만 Package에 없음: {len(in_folder_not_pkg)}개")
for c in sorted(in_folder_not_pkg):
    fname = file_codes[c][0]
    print(f"    {c} - {fname}")

in_pkg_not_folder = pkg_codes - folder_codes
print(f"\n  [B] Package에 있지만 폴더에 없음: {len(in_pkg_not_folder)}개")
for c in sorted(in_pkg_not_folder)[:20]:
    print(f"    {c}")

in_pkg_not_excel = pkg_codes - excel_codes
print(f"\n  [C] Package에 있지만 Excel에 없음: {len(in_pkg_not_excel)}개")
for c in sorted(in_pkg_not_excel)[:20]:
    print(f"    {c}")

in_folder_not_excel = folder_codes - excel_codes
print(f"\n  [D] 폴더에 있지만 Excel에 없음: {len(in_folder_not_excel)}개")
for c in sorted(in_folder_not_excel):
    fname = file_codes[c][0]
    print(f"    {c} - {fname}")

# ──────────────────────────────────────────────────────────────
# [6] has_report 동작 실측
# ──────────────────────────────────────────────────────────────
print("\n" + "=" * 75)
print(" [6] loader.has_report() 실측")
print("=" * 75)
hits = 0
miss_in_pkg = []  # Excel에 있고 Package에도 있는데 has_report=False인 경우
for code in excel_codes:
    if loader.has_report(code):
        hits += 1
    elif code in pkg_codes:
        miss_in_pkg.append(code)
print(f"  Excel 코드 {len(excel_codes)}개 중 has_report()=True : {hits}개")
if miss_in_pkg:
    print(f"  [!] Package에 키는 있는데 has_report=False인 코드 {len(miss_in_pkg)}개:")
    for c in miss_in_pkg[:10]:
        print(f"    {c}")

print("\n" + "=" * 75)
print(" 검증 완료")
print("=" * 75)
