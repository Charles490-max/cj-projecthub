# src/secure_viewer.py
"""
CJ ProjectHub - 보안 준공보고서 뷰어
- 패키지(reports.pkg) 내 페이지 이미지를 메모리에서만 디코딩하여 표시
- 사용자에게 원본 파일을 노출하지 않음 (파일 시스템에 저장하지 않음)
- 캡처/복사/저장 단축키 차단, 워터마크 표시, 열람 로그 기록
"""

import io
import getpass
import platform
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox

from PIL import Image, ImageTk, ImageDraw, ImageFont

from config import (
    APP_NAME, APP_VERSION,
    COLOR_PRIMARY, COLOR_ACCENT, COLOR_GRAY, COLOR_LIGHT,
    LOGS_DIR,
)


# ==========================================================
# 보안 뷰어 (Toplevel 윈도우)
# ==========================================================
class SecureReportViewer(tk.Toplevel):
    """패키지에서 페이지 이미지를 받아 화면에만 표시하는 보안 뷰어."""

    # 줌 단계
    ZOOM_LEVELS = [0.5, 0.75, 1.0, 1.25, 1.5, 2.0]

    def __init__(self, master, loader, site_code, project_info=None, start_page=1):
        super().__init__(master)
        self.loader = loader
        self.site_code = str(site_code)
        self.project_info = project_info or {}
        self.user = getpass.getuser()
        self.host = platform.node()

        # 상태
        self.total_pages = self.loader.get_page_count(self.site_code)
        if self.total_pages <= 0:
            messagebox.showwarning("알림", "표시할 페이지가 없습니다.", parent=master)
            self.destroy()
            return

        self.current_page = max(1, min(start_page, self.total_pages))
        self.zoom_idx = 2  # 기본 100%
        self._tk_img = None     # ImageTk 참조 보존
        self._raw_image = None  # 원본 PIL Image (현재 페이지)

        # 창 설정
        self.title(f"{APP_NAME} - 보고서 뷰어 [{self.site_code}]")
        self.geometry("1200x850")
        self.minsize(800, 600)
        self.configure(bg=COLOR_LIGHT)

        # 닫기 동작 통제
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # UI / 보안 / 로그
        self._build_ui()
        self._bind_security()
        self._log_open()

        # 첫 페이지 렌더
        self.after(50, self._render_current_page)

        # 모달처럼 동작
        self.transient(master)
        self.lift()
        self.focus_force()

    # ------------------------------------------------------
    # 프로젝트명 안전하게 가져오기
    # ------------------------------------------------------
    def _get_display_title(self):
        """헤더 표시용 제목.
        1) 호출자가 넘긴 project_info['프로젝트명'] 우선
        2) loader.get_original_name(site_code) 사용 (원본 파일명)
        3) 모두 실패 시 기본값
        """
        name = self.project_info.get('프로젝트명', '') if self.project_info else ''
        if not name:
            try:
                if hasattr(self.loader, 'get_original_name'):
                    name = self.loader.get_original_name(self.site_code) or ''
            except Exception:
                name = ''
        return name or '준공보고서'

    # ------------------------------------------------------
    # UI
    # ------------------------------------------------------
    def _build_ui(self):
        # 헤더
        header = tk.Frame(self, bg=COLOR_PRIMARY, height=55)
        header.pack(fill='x')
        header.pack_propagate(False)

        title = self._get_display_title()
        tk.Label(header, text=f"📄 {title}",
                 font=('맑은 고딕', 13, 'bold'),
                 fg='white', bg=COLOR_PRIMARY).pack(side='left', padx=15, pady=10)
        tk.Label(header, text=f"현장코드 {self.site_code}",
                 font=('맑은 고딕', 10),
                 fg='white', bg=COLOR_PRIMARY).pack(side='left', padx=10)
        tk.Label(header,
                 text=f"열람자: {self.user}@{self.host}  |  {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                 font=('맑은 고딕', 9),
                 fg='#cfd8e3', bg=COLOR_PRIMARY).pack(side='right', padx=15)

        # 보안 알림 바
        notice = tk.Frame(self, bg='#FFF3CD', height=24)
        notice.pack(fill='x')
        notice.pack_propagate(False)
        tk.Label(notice,
                 text="🔒 본 자료는 CJ대한통운 건설부문 내부 자료입니다. 캡처/복제/외부 유출이 금지됩니다.",
                 font=('맑은 고딕', 9), fg='#856404', bg='#FFF3CD').pack(side='left', padx=12)

        # 툴바
        toolbar = tk.Frame(self, bg='white', height=40)
        toolbar.pack(fill='x')
        toolbar.pack_propagate(False)

        btn_style = {'font': ('맑은 고딕', 10), 'bg': 'white',
                     'relief': 'flat', 'cursor': 'hand2', 'padx': 8}

        tk.Button(toolbar, text='◀ 이전', command=self.prev_page, **btn_style).pack(side='left', padx=4, pady=5)
        self.page_label = tk.Label(toolbar, text='', font=('맑은 고딕', 10, 'bold'),
                                    bg='white', width=14)
        self.page_label.pack(side='left', padx=4)
        tk.Button(toolbar, text='다음 ▶', command=self.next_page, **btn_style).pack(side='left', padx=4)

        ttk.Separator(toolbar, orient='vertical').pack(side='left', fill='y', padx=8, pady=8)

        tk.Button(toolbar, text='🔍-', command=self.zoom_out, **btn_style).pack(side='left', padx=4)
        self.zoom_label = tk.Label(toolbar, text='100%', font=('맑은 고딕', 10),
                                    bg='white', width=6)
        self.zoom_label.pack(side='left', padx=2)
        tk.Button(toolbar, text='🔍+', command=self.zoom_in, **btn_style).pack(side='left', padx=4)
        tk.Button(toolbar, text='맞춤', command=self.fit_window, **btn_style).pack(side='left', padx=4)

        ttk.Separator(toolbar, orient='vertical').pack(side='left', fill='y', padx=8, pady=8)

        # 페이지 이동 입력
        tk.Label(toolbar, text='페이지 이동:', font=('맑은 고딕', 9), bg='white').pack(side='left', padx=4)
        self.goto_entry = tk.Entry(toolbar, width=5, font=('맑은 고딕', 10))
        self.goto_entry.pack(side='left')
        self.goto_entry.bind('<Return>', lambda e: self._goto_from_entry())
        tk.Button(toolbar, text='이동', command=self._goto_from_entry, **btn_style).pack(side='left', padx=4)

        tk.Button(toolbar, text='✕ 닫기', command=self._on_close,
                  font=('맑은 고딕', 10, 'bold'), bg=COLOR_ACCENT, fg='white',
                  relief='flat', cursor='hand2', padx=12).pack(side='right', padx=8, pady=5)

        # 본문 (스크롤 가능한 캔버스)
        body = tk.Frame(self, bg=COLOR_LIGHT)
        body.pack(fill='both', expand=True)

        self.canvas = tk.Canvas(body, bg='#3a3a3a', highlightthickness=0)
        vsb = ttk.Scrollbar(body, orient='vertical', command=self.canvas.yview)
        hsb = ttk.Scrollbar(body, orient='horizontal', command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.canvas.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        # 마우스 휠 스크롤
        self.canvas.bind('<MouseWheel>', self._on_mousewheel)
        self.canvas.bind('<Button-4>', lambda e: self.canvas.yview_scroll(-1, 'units'))
        self.canvas.bind('<Button-5>', lambda e: self.canvas.yview_scroll(1, 'units'))

        # 키 바인딩 (탐색)
        self.bind('<Left>',  lambda e: self.prev_page())
        self.bind('<Right>', lambda e: self.next_page())
        self.bind('<Prior>', lambda e: self.prev_page())  # PageUp
        self.bind('<Next>',  lambda e: self.next_page())  # PageDown
        self.bind('<Home>',  lambda e: self.goto_page(1))
        self.bind('<End>',   lambda e: self.goto_page(self.total_pages))
        self.bind('<Escape>', lambda e: self._on_close())

        # 푸터
        footer = tk.Frame(self, bg='white', height=22)
        footer.pack(fill='x', side='bottom')
        footer.pack_propagate(False)
        tk.Label(footer,
                 text=f"{APP_NAME} {APP_VERSION}  |  열람 기록이 저장됩니다.",
                 font=('맑은 고딕', 8), fg=COLOR_GRAY, bg='white').pack(side='left', padx=10)

    # ------------------------------------------------------
    # 보안 키 바인딩 (캡처/복사/저장 차단)
    # ------------------------------------------------------
    def _bind_security(self):
        """보안 키 바인딩.
        - PrintScreen(<Snapshot>)은 Windows가 OS 레벨에서 처리하므로
          Tkinter 이벤트로 전달되지 않아 바인딩 불가 → 제외.
        - 환경별로 미지원되는 키는 try/except로 무시.
        """
        blocked_keys = [
            '<Control-c>', '<Control-C>',
            '<Control-v>', '<Control-V>',
            '<Control-x>', '<Control-X>',
            '<Control-s>', '<Control-S>',
            '<Control-p>', '<Control-P>',
            '<Control-Insert>',
            '<Shift-Insert>',
        ]
        for key in blocked_keys:
            try:
                self.bind_all(key, self._block_action)
            except Exception:
                pass  # 일부 키는 환경에 따라 실패할 수 있음 — 무시

    def _block_action(self, event=None):
        try:
            self.bell()
        except Exception:
            pass
        self._log('BLOCKED', f'key={getattr(event, "keysym", "?")}')
        return 'break'

    # ------------------------------------------------------
    # 페이지 렌더링
    # ------------------------------------------------------
    def _render_current_page(self):
        """현재 페이지 이미지를 캔버스에 렌더 (워터마크 포함).
        loader.get_page_image() 가 bytes / PIL.Image / fitz.Pixmap /
        파일경로 / numpy.ndarray 중 무엇을 반환하더라도 안전하게 처리한다.
        """
        try:
            page_data = self.loader.get_page_image(self.site_code, self.current_page)
        except Exception as e:
            messagebox.showerror("오류", f"페이지를 불러올 수 없습니다:\n{e}", parent=self)
            return

        if page_data is None:
            messagebox.showwarning("알림", "페이지 데이터가 비어 있습니다.", parent=self)
            return

        img = self._normalize_to_pil(page_data)
        if img is None:
            messagebox.showerror("오류",
                f"이미지 변환 실패: 알 수 없는 타입 ({type(page_data).__name__})",
                parent=self)
            return

        # 워터마크 적용
        img = self._apply_watermark(img)
        self._raw_image = img
        self._draw_to_canvas()

        # 라벨 업데이트
        self.page_label.config(text=f"{self.current_page} / {self.total_pages}")
        self.zoom_label.config(text=f"{int(self.ZOOM_LEVELS[self.zoom_idx]*100)}%")
        self.title(f"{APP_NAME} - 보고서 뷰어 [{self.site_code}]  ({self.current_page}/{self.total_pages})")

    def _normalize_to_pil(self, data):
        """다양한 반환 타입을 PIL.Image (RGB) 로 정규화."""
        # 1) 이미 PIL Image 인 경우
        if isinstance(data, Image.Image):
            try:
                return data.convert('RGB')
            except Exception:
                return data

        # 2) bytes / bytearray / memoryview → BytesIO 로 디코딩
        if isinstance(data, (bytes, bytearray, memoryview)):
            try:
                return Image.open(io.BytesIO(bytes(data))).convert('RGB')
            except Exception as e:
                messagebox.showerror("오류", f"이미지 디코딩 실패:\n{e}", parent=self)
                return None

        # 3) fitz.Pixmap (PyMuPDF)
        try:
            import fitz  # 선택적 import
            if isinstance(data, fitz.Pixmap):
                pix = data
                if pix.alpha:  # RGBA → RGB
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                return Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
        except ImportError:
            pass
        except Exception as e:
            messagebox.showerror("오류", f"Pixmap 변환 실패:\n{e}", parent=self)
            return None

        # 4) 파일 경로 문자열 (Path 포함)
        if isinstance(data, (str, Path)):
            try:
                return Image.open(str(data)).convert('RGB')
            except Exception as e:
                messagebox.showerror("오류", f"이미지 파일 열기 실패:\n{e}", parent=self)
                return None

        # 5) numpy.ndarray (혹시 모를 경우)
        try:
            import numpy as np
            if isinstance(data, np.ndarray):
                return Image.fromarray(data).convert('RGB')
        except ImportError:
            pass
        except Exception:
            return None

        return None  # 어느 타입에도 해당 안 됨

    def _draw_to_canvas(self):
        if self._raw_image is None:
            return
        zoom = self.ZOOM_LEVELS[self.zoom_idx]
        w, h = self._raw_image.size
        new_w, new_h = max(1, int(w * zoom)), max(1, int(h * zoom))
        try:
            resized = self._raw_image.resize((new_w, new_h), Image.LANCZOS)
        except Exception:
            resized = self._raw_image

        self._tk_img = ImageTk.PhotoImage(resized)
        self.canvas.delete('all')
        # 캔버스 중앙 정렬
        cw = max(self.canvas.winfo_width(), new_w)
        ch = max(self.canvas.winfo_height(), new_h)
        x = max((cw - new_w) // 2, 0)
        y = max((ch - new_h) // 2, 0)
        self.canvas.create_image(x, y, anchor='nw', image=self._tk_img)
        self.canvas.configure(scrollregion=(0, 0, cw, ch))

    def _apply_watermark(self, img):
        """대각선 반복 워터마크 (열람자 ID + 시각)."""
        try:
            wm_text = f"{self.user}  |  {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  CJ ProjectHub"
            overlay = Image.new('RGBA', img.size, (0, 0, 0, 0))
            draw = ImageDraw.Draw(overlay)
            try:
                font = ImageFont.truetype("C:/Windows/Fonts/malgun.ttf", 28)
            except Exception:
                font = ImageFont.load_default()
            w, h = img.size
            step_x, step_y = 380, 220
            for y in range(-h, h * 2, step_y):
                for x in range(-w, w * 2, step_x):
                    draw.text((x, y), wm_text, fill=(180, 180, 180, 70), font=font)
            base = img.convert('RGBA')
            combined = Image.alpha_composite(base, overlay).convert('RGB')
            return combined
        except Exception:
            return img  # 워터마크 실패 시 원본 반환

    # ------------------------------------------------------
    # 페이지 이동 / 줌
    # ------------------------------------------------------
    def prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self._render_current_page()
            self._log('PAGE', f'p={self.current_page}')

    def next_page(self):
        if self.current_page < self.total_pages:
            self.current_page += 1
            self._render_current_page()
            self._log('PAGE', f'p={self.current_page}')

    def goto_page(self, page):
        page = max(1, min(int(page), self.total_pages))
        if page != self.current_page:
            self.current_page = page
            self._render_current_page()
            self._log('PAGE', f'p={self.current_page}')

    def _goto_from_entry(self):
        try:
            v = int(self.goto_entry.get())
            self.goto_page(v)
        except ValueError:
            messagebox.showinfo("알림", "페이지 번호를 숫자로 입력하세요.", parent=self)
        finally:
            self.goto_entry.delete(0, 'end')

    def zoom_in(self):
        if self.zoom_idx < len(self.ZOOM_LEVELS) - 1:
            self.zoom_idx += 1
            self._draw_to_canvas()
            self.zoom_label.config(text=f"{int(self.ZOOM_LEVELS[self.zoom_idx]*100)}%")

    def zoom_out(self):
        if self.zoom_idx > 0:
            self.zoom_idx -= 1
            self._draw_to_canvas()
            self.zoom_label.config(text=f"{int(self.ZOOM_LEVELS[self.zoom_idx]*100)}%")

    def fit_window(self):
        """창 크기에 맞도록 자동 스케일."""
        if self._raw_image is None:
            return
        cw = self.canvas.winfo_width() or 1000
        ch = self.canvas.winfo_height() or 700
        iw, ih = self._raw_image.size
        if iw <= 0 or ih <= 0:
            return
        ratio = min(cw / iw, ch / ih)
        # 가장 가까운 ZOOM_LEVELS 선택
        best = min(range(len(self.ZOOM_LEVELS)),
                   key=lambda i: abs(self.ZOOM_LEVELS[i] - ratio))
        self.zoom_idx = best
        self._draw_to_canvas()
        self.zoom_label.config(text=f"{int(self.ZOOM_LEVELS[self.zoom_idx]*100)}%")

    def _on_mousewheel(self, event):
        # Windows 휠 단위(120)
        delta = -1 if event.delta > 0 else 1
        self.canvas.yview_scroll(delta, 'units')

    # ------------------------------------------------------
    # 로그
    # ------------------------------------------------------
    def _log_open(self):
        self._log('OPEN',
                  f"site={self.site_code} pages={self.total_pages} "
                  f"name={self.project_info.get('프로젝트명', '') if self.project_info else ''}")

    def _log(self, action, detail=''):
        try:
            log_path = Path(LOGS_DIR) / f"viewer_{datetime.now().strftime('%Y%m')}.log"
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\t"
                        f"{self.user}@{self.host}\t{action}\t{detail}\n")
        except Exception:
            pass  # 로그 실패는 사용자 흐름을 막지 않음

    # ------------------------------------------------------
    # 종료
    # ------------------------------------------------------
    def _on_close(self):
        self._log('CLOSE', f'last_page={self.current_page}')
        try:
            # 전역 바인딩 해제
            for key in ('<Control-c>', '<Control-C>', '<Control-v>', '<Control-V>',
                        '<Control-x>', '<Control-X>', '<Control-s>', '<Control-S>',
                        '<Control-p>', '<Control-P>',
                        '<Control-Insert>', '<Shift-Insert>'):
                try:
                    self.unbind_all(key)
                except Exception:
                    pass
        finally:
            self.destroy()
