"""
PPT → PDF 일괄 변환 스크립트
※ Charles님 PC에서 PowerPoint 인증된 상태로 1회 실행
※ DRM 보호된 PPT를 본인 권한으로 PDF 변환 (정상 동작)
"""
import os
import sys
import time
import shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from config import REPORTS_ORIGINAL_DIR

# 변환된 PDF가 저장될 폴더
PDF_OUTPUT_DIR = REPORTS_ORIGINAL_DIR.parent / 'Reports_PDF'
# 원본 PPT를 백업할 폴더 (변환 후 자동 이동)
PPT_BACKUP_DIR = REPORTS_ORIGINAL_DIR.parent / 'Reports_PPT_원본'

PDF_OUTPUT_DIR.mkdir(exist_ok=True)
PPT_BACKUP_DIR.mkdir(exist_ok=True)

# PowerPoint PDF 저장 포맷 상수
ppSaveAsPDF = 32


def convert_pptx_to_pdf(ppt_path: Path, pdf_path: Path):
    """단일 PPT 파일을 PDF로 변환"""
    import win32com.client
    import pythoncom
    
    pythoncom.CoInitialize()
    powerpoint = None
    deck = None
    
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1
        powerpoint.DisplayAlerts = 0
        
        # PPT 열기 (DRM 인증된 사용자 권한으로)
        deck = powerpoint.Presentations.Open(
            str(ppt_path.absolute()),
            ReadOnly=False,
            WithWindow=False
        )
        
        # PDF로 저장
        deck.SaveAs(str(pdf_path.absolute()), ppSaveAsPDF)
        time.sleep(0.5)
        
        return True
    
    except Exception as e:
        print(f"      ✗ 변환 실패: {e}")
        return False
    
    finally:
        try:
            if deck:
                deck.Close()
        except Exception:
            pass
        try:
            if powerpoint:
                powerpoint.Quit()
        except Exception:
            pass
        # 좀비 프로세스 정리
        os.system('taskkill /F /IM POWERPNT.EXE >nul 2>&1')
        pythoncom.CoUninitialize()
        time.sleep(0.5)


def main():
    print("=" * 65)
    print("  PPT → PDF 일괄 변환")
    print("=" * 65)
    print(f"\n  입력: {REPORTS_ORIGINAL_DIR}")
    print(f"  출력 (PDF): {PDF_OUTPUT_DIR}")
    print(f"  백업 (원본 PPT): {PPT_BACKUP_DIR}")
    print()
    
    # PPT 파일 수집
    ppt_files = []
    for f in REPORTS_ORIGINAL_DIR.glob('*'):
        if f.suffix.lower() in ['.pptx', '.ppt']:
            ppt_files.append(f)
    
    print(f"  처리 대상 PPT: {len(ppt_files)}건\n")
    
    if not ppt_files:
        print("  변환할 PPT가 없습니다.")
        return
    
    # 사용자 확인
    response = input(f"  진행하시겠습니까? (y/n): ").strip().lower()
    if response != 'y':
        print("  취소됨")
        return
    
    print()
    success_count = 0
    fail_count = 0
    skip_count = 0
    failed_files = []
    
    for idx, ppt_path in enumerate(ppt_files, 1):
        # PDF 출력 경로 (확장자만 변경)
        pdf_name = ppt_path.stem + '.pdf'
        pdf_path = PDF_OUTPUT_DIR / pdf_name
        
        print(f"[{idx}/{len(ppt_files)}] {ppt_path.name}")
        
        # 이미 변환된 경우 스킵
        if pdf_path.exists() and pdf_path.stat().st_size > 10000:
            print(f"      ⊘ 이미 변환됨 (스킵)")
            skip_count += 1
            continue
        
        # 변환 실행
        if convert_pptx_to_pdf(ppt_path, pdf_path):
            # 검증: PDF 파일 크기 + PDF 시그니처
            if pdf_path.exists() and pdf_path.stat().st_size > 10000:
                with open(pdf_path, 'rb') as f:
                    sig = f.read(5)
                if sig == b'%PDF-':
                    size_mb = pdf_path.stat().st_size / 1024 / 1024
                    print(f"      ✓ 변환 완료 ({size_mb:.1f} MB)")
                    success_count += 1
                    
                    # 원본 PPT를 백업 폴더로 이동
                    backup_path = PPT_BACKUP_DIR / ppt_path.name
                    try:
                        shutil.move(str(ppt_path), str(backup_path))
                        print(f"      → 원본 PPT 백업 완료")
                    except Exception as e:
                        print(f"      ⚠ 원본 이동 실패: {e}")
                else:
                    print(f"      ✗ PDF 시그니처 불일치 (DRM 차단 가능성)")
                    fail_count += 1
                    failed_files.append(ppt_path.name)
                    pdf_path.unlink(missing_ok=True)
            else:
                print(f"      ✗ PDF 미생성 또는 크기 비정상")
                fail_count += 1
                failed_files.append(ppt_path.name)
        else:
            fail_count += 1
            failed_files.append(ppt_path.name)
    
    # 결과 요약
    print("\n" + "=" * 65)
    print(f"  변환 결과")
    print("=" * 65)
    print(f"  ✓ 성공: {success_count}건")
    print(f"  ⊘ 스킵: {skip_count}건 (기변환)")
    print(f"  ✗ 실패: {fail_count}건")
    
    if failed_files:
        print(f"\n  실패 파일 목록:")
        for name in failed_files:
            print(f"    - {name}")
        print(f"\n  ※ 실패한 파일은 PowerPoint에서 수동으로 PDF 변환 필요")
    
    print(f"\n  변환된 PDF: {PDF_OUTPUT_DIR}")
    print(f"  원본 백업: {PPT_BACKUP_DIR}")


if __name__ == '__main__':
    main()
