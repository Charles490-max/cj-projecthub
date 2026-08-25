# src/main.py
"""
CJ ProjectHub - 메인 애플리케이션
프로젝트 검색 / 키워드 본문 검색 / 준공보고서 검색
"""
import os
import sys
import subprocess
import socket
import getpass
import traceback
import threading
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

import pandas as pd

from config import (
    APP_NAME, APP_VERSION, APP_TITLE,
    TAB1_TITLE, TAB2_TITLE, TAB3_TITLE,
    COLOR_PRIMARY, COLOR_ACCENT, COLOR_GRAY, COLOR_LIGHT,
    DATA_FILE, PACKAGE_FILE, OUTPUT_DIR, REPORTS_ORIGINAL_DIR,
    DISPLAY_FIELDS, REPORT_SEARCH_FIELDS, FIELD_DISPLAY_NAME,
    EMPTY_MARK, fmt,
)
from data_loader import load_projects, calculate_similarity, get_building_types
from package_loader import SecurePackageLoader
from secure_viewer import SecureReportViewer
from pdf_generator import generate_reference_pdf

# tkinterdnd2 선택적 임포트 (드래그앤드롭용)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False


class CJProjectHubApp:
    """메인 애플리케이션 클래스"""

    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1500x850")
        self.root.configure(bg=COLOR_LIGHT)

        self.user = getpass.getuser()
        try:
            self.host = socket.gethostname()
        except Exception:
            self.host = 'unknown'

        self.df = None
        self.loader = None
        self._load_data()

        self.last_search_result = None
        self.last_search_query = None

        self._build_header()
        self._build_tabs()
        self._build_footer()

    # ──────────────────────────────────────────────────────────
    # 초기 로드
    # ──────────────────────────────────────────────────────────
    def _load_data(self):
        try:
            self.df = load_projects(DATA_FILE)
            print(f"[INFO] 데이터 로드: {len(self.df)}건")
        except Exception as e:
            messagebox.showerror("데이터 로드 실패", f"data.xlsx 로드 실패:\n{e}")
            self.df = pd.DataFrame()

        try:
            self.loader = SecurePackageLoader(PACKAGE_FILE)
            if not self.loader.is_loaded():
                self.loader = None
                messagebox.showwarning("알림", "보고서 패키지 로드 실패. 본문 검색/뷰어 사용 불가.")
        except Exception as e:
            self.loader = None
            messagebox.showwarning("알림", f"보고서 패키지 로드 실패:\n{e}")

    # ──────────────────────────────────────────────────────────
    # UI 헤더 / 탭 / 푸터
    # ──────────────────────────────────────────────────────────
    def _build_header(self):
        header = tk.Frame(self.root, bg=COLOR_PRIMARY, height=50)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text=APP_TITLE, bg=COLOR_PRIMARY, fg='white',
                 font=('맑은 고딕', 14, 'bold')).pack(side='left', padx=20, pady=10)
        tk.Label(header, text=f"User: {self.user}@{self.host}",
                 bg=COLOR_PRIMARY, fg='white',
                 font=('맑은 고딕', 9)).pack(side='right', padx=20)

    def _build_tabs(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        self.tab1 = tk.Frame(self.notebook, bg='white')
        self.tab2 = tk.Frame(self.notebook, bg='white')
        self.tab3 = tk.Frame(self.notebook, bg='white')
        self.tab4 = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.tab1, text=TAB1_TITLE)
        self.notebook.add(self.tab2, text=TAB2_TITLE)
        self.notebook.add(self.tab3, text=TAB3_TITLE)
        self.notebook.add(self.tab4, text="보고서 관리")
        self._build_tab1()
        self._build_tab2()
        self._build_tab3()
        self._build_tab4()

    def _build_footer(self):
        footer = tk.Frame(self.root, bg=COLOR_GRAY, height=22)
        footer.pack(fill='x', side='bottom')
        footer.pack_propagate(False)
        tk.Label(footer,
                 text=f"{APP_NAME} {APP_VERSION}  |  본 화면의 모든 접근은 로그로 기록됩니다.",
                 bg=COLOR_GRAY, fg='white',
                 font=('맑은 고딕', 8)).pack(side='left', padx=10)

    # ──────────────────────────────────────────────────────────
    # Tab 1: 프로젝트 검색 (유사도)
    # ──────────────────────────────────────────────────────────
    def _build_tab1(self):
        top = tk.LabelFrame(self.tab1, text="프로젝트 검색 (유사도)",
                            font=('맑은 고딕', 11, 'bold'), padx=10, pady=10)
        top.pack(fill='x', padx=5, pady=5)

        tk.Label(top, text="건축물종류:").grid(row=0, column=0, sticky='w', padx=5, pady=3)
        self.tab1_building = ttk.Combobox(top, width=18, state='readonly')
        self.tab1_building.grid(row=0, column=1, padx=5)
        try:
            types = get_building_types(self.df)
        except Exception as e:
            print(f"[WARN] get_building_types 실패: {e}")
            types = []
        types = sorted({str(t).strip() for t in types
                        if t and str(t).strip().lower() not in ('nan', 'none', '')})
        self.tab1_building['values'] = ['(전체)'] + types
        self.tab1_building.current(0)

        tk.Label(top, text="연면적(㎡):").grid(row=0, column=2, sticky='w', padx=5)
        self.tab1_area = tk.Entry(top, width=12)
        self.tab1_area.grid(row=0, column=3, padx=5)

        tk.Label(top, text="지하층:").grid(row=0, column=4, sticky='w', padx=5)
        self.tab1_under = tk.Entry(top, width=6)
        self.tab1_under.grid(row=0, column=5, padx=5)

        tk.Label(top, text="지상층:").grid(row=0, column=6, sticky='w', padx=5)
        self.tab1_above = tk.Entry(top, width=6)
        self.tab1_above.grid(row=0, column=7, padx=5)

        tk.Button(top, text="🔍 검색", command=self.search_similar,
                  bg=COLOR_PRIMARY, fg='white',
                  font=('맑은 고딕', 10, 'bold'), width=10
                  ).grid(row=0, column=8, padx=10)

        tk.Button(top, text="📄 1페이지 PDF 출력", command=self.export_pdf,
                  bg=COLOR_ACCENT, fg='white',
                  font=('맑은 고딕', 10, 'bold'), width=18
                  ).grid(row=0, column=9, padx=5)

        tk.Label(top, text="(빈 칸은 0으로 간주됨. 건축물종류 가중치가 가장 큼)",
                 fg='gray', font=('맑은 고딕', 9)
                 ).grid(row=1, column=0, columnspan=10, sticky='w', padx=5, pady=(3, 0))

        self.tab1_status_label = tk.Label(
            self.tab1,
            text="검색 조건을 입력하고 검색 버튼을 누르세요.",
            fg=COLOR_PRIMARY, font=('맑은 고딕', 10, 'bold'), anchor='w'
        )
        self.tab1_status_label.pack(fill='x', padx=10, pady=(5, 2))

        self.tree1 = self._make_tree(self.tab1, with_similarity=True)
        self.tree1.bind('<Double-1>', self._on_tree1_double_click)

    def _parse_float(self, s, default=0.0):
        s = (s or '').strip()
        if not s:
            return default
        try:
            return float(s.replace(',', ''))
        except Exception:
            return default

    def _parse_int(self, s, default=0):
        s = (s or '').strip()
        if not s:
            return default
        try:
            return int(float(s))
        except Exception:
            return default

    def search_similar(self):
        """프로젝트 검색 — calculate_similarity(df, target_dict) 시그니처에 맞춰 호출."""
        if self.df is None or self.df.empty:
            messagebox.showwarning("알림", "데이터가 없습니다.")
            return

        bld_raw = self.tab1_building.get().strip()
        bld = '' if bld_raw in ('', '(전체)') else bld_raw
        area = self._parse_float(self.tab1_area.get(), 0.0)
        under = self._parse_int(self.tab1_under.get(), 0)
        above = self._parse_int(self.tab1_above.get(), 0)

        target = {
            '건축물종류': bld,
            '연면적(㎡)': area,
            '지하층': under,
            '지상층': above,
        }
        print(f"[INFO] Tab1 검색 — target={target}")

        try:
            result = calculate_similarity(self.df, target)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("오류", f"유사도 계산 실패:\n{type(e).__name__}: {e}")
            return

        if not isinstance(result, pd.DataFrame) or result.empty:
            messagebox.showinfo("결과", "검색 결과가 없습니다.")
            self.tab1_status_label.config(text="검색 결과: 0건")
            for i in self.tree1.get_children():
                self.tree1.delete(i)
            return

        # similarity → 유사도(%) 환산
        if 'similarity' in result.columns and '유사도' not in result.columns:
            sim_series = result['similarity'].astype(float)
            max_val = sim_series.max() if not sim_series.empty else 0
            if max_val <= 1.0:
                result['유사도'] = sim_series * 100.0
            else:
                result['유사도'] = sim_series
        elif '유사도' not in result.columns:
            result['유사도'] = 0.0

        result = result.sort_values('유사도', ascending=False).reset_index(drop=True)

        self.last_search_result = result
        self.last_search_query = target

        for i in self.tree1.get_children():
            self.tree1.delete(i)

        top_n = result.head(50)
        n_with_report = 0
        for _, row in top_n.iterrows():
            code = str(row.get('현장코드', '')).strip()
            try:
                sim = f"{float(row.get('유사도', 0)):.1f}%"
            except Exception:
                sim = '-'
            values = [fmt(row.get(c), c) for c, _ in DISPLAY_FIELDS]
            has_flag = self._has_report(code)
            if has_flag:
                n_with_report += 1
            values.append('✓' if has_flag else EMPTY_MARK)
            values.append(sim)
            self.tree1.insert('', 'end', values=values, tags=(code,))

        self.tab1_status_label.config(
            text=(f"검색 결과 — 전체 {len(result)}건, "
                  f"화면 표시 상위 {len(top_n)}건, 보고서 보유 {n_with_report}건  "
                  f"(조건: 건축물종류='{bld or '(전체)'}', "
                  f"연면적={area:,.0f}㎡, 지하={under}, 지상={above})")
        )
        print(f"[INFO] Tab1 결과: 전체 {len(result)}건, 표시 {len(top_n)}건, "
              f"보고서 {n_with_report}건")

    # ──────────────────────────────────────────────────────────
    # Tab 2: 키워드 본문 검색
    # ──────────────────────────────────────────────────────────
    def _build_tab2(self):
        top = tk.LabelFrame(self.tab2, text="키워드 본문 검색",
                            font=('맑은 고딕', 11, 'bold'), padx=10, pady=10)
        top.pack(fill='x', padx=5, pady=5)

        tk.Label(top, text="검색어:").grid(row=0, column=0, padx=5, pady=3)
        self.kw_entry = tk.Entry(top, width=40)
        self.kw_entry.grid(row=0, column=1, padx=5)
        self.kw_entry.bind('<Return>', lambda e: self.search_keyword())

        tk.Button(top, text="🔍 검색", command=self.search_keyword,
                  bg=COLOR_PRIMARY, fg='white',
                  font=('맑은 고딕', 10, 'bold'), width=10
                  ).grid(row=0, column=2, padx=10)

        cols = ('현장코드', '프로젝트명', '페이지', '본문 스니펫')
        self.tree2 = ttk.Treeview(self.tab2, columns=cols, show='headings', height=25)
        for c, w in zip(cols, (90, 280, 70, 800)):
            self.tree2.heading(c, text=c)
            self.tree2.column(c, width=w, anchor='w' if c != '페이지' else 'center')
        vsb = ttk.Scrollbar(self.tab2, orient='vertical', command=self.tree2.yview)
        self.tree2.configure(yscrollcommand=vsb.set)
        self.tree2.pack(side='left', fill='both', expand=True, padx=(5, 0), pady=5)
        vsb.pack(side='left', fill='y', pady=5)
        self.tree2.bind('<Double-1>', self._on_tree2_double_click)

    def search_keyword(self):
        if self.loader is None:
            messagebox.showwarning("알림", "보고서 패키지가 로드되지 않았습니다.")
            return
        kw = self.kw_entry.get().strip()
        if not kw:
            return
        try:
            if hasattr(self.loader, 'search_text'):
                raw_results = self.loader.search_text(kw)
            elif hasattr(self.loader, 'search_keyword'):
                raw_results = self.loader.search_keyword(kw, limit=200)
            else:
                messagebox.showerror("오류", "패키지 로더에 본문 검색 메서드가 없습니다.")
                return
        except Exception as e:
            messagebox.showerror("오류", f"검색 실패:\n{e}")
            return

        results = self._normalize_search_results(raw_results, kw)
        for i in self.tree2.get_children():
            self.tree2.delete(i)
        if not results:
            messagebox.showinfo("결과", "검색 결과가 없습니다.")
            return
        for r in results[:200]:
            code = str(r.get('site_code', ''))
            page = r.get('page', 1)
            snippet = r.get('snippet', '')
            row = self.df[self.df['현장코드'].astype(str) == code]
            pname = row.iloc[0].get('프로젝트명', '') if not row.empty else ''
            self.tree2.insert('', 'end', values=(code, pname, page, snippet),
                              tags=(code, str(page)))

    def _normalize_search_results(self, raw, keyword):
        if not raw:
            return []
        results = []
        if isinstance(raw, dict):
            for code, pages in raw.items():
                if isinstance(pages, (list, tuple, set)):
                    for p in pages:
                        if isinstance(p, dict):
                            results.append({'site_code': code,
                                            'page': p.get('page', 1),
                                            'snippet': p.get('snippet', '')})
                        else:
                            results.append({'site_code': code, 'page': p, 'snippet': ''})
                else:
                    results.append({'site_code': code, 'page': pages, 'snippet': ''})
            return results
        if isinstance(raw, (list, tuple)):
            for item in raw:
                if isinstance(item, dict):
                    results.append({
                        'site_code': item.get('site_code') or item.get('code') or '',
                        'page': item.get('page', 1),
                        'snippet': item.get('snippet') or item.get('text') or ''
                    })
                elif isinstance(item, (list, tuple)):
                    if len(item) >= 3:
                        results.append({'site_code': str(item[0]),
                                        'page': item[1], 'snippet': str(item[2])})
                    elif len(item) == 2:
                        results.append({'site_code': str(item[0]),
                                        'page': item[1], 'snippet': ''})
                    elif len(item) == 1:
                        results.append({'site_code': str(item[0]),
                                        'page': 1, 'snippet': ''})
                elif isinstance(item, str):
                    results.append({'site_code': item, 'page': 1, 'snippet': ''})
            return results
        return []

    # ──────────────────────────────────────────────────────────
    # Tab 3: 준공보고서 검색
    # ──────────────────────────────────────────────────────────
    def _build_tab3(self):
        top = tk.LabelFrame(self.tab3, text="준공보고서 검색",
                            font=('맑은 고딕', 11, 'bold'), padx=10, pady=10)
        top.pack(fill='x', padx=5, pady=5)

        tk.Label(top, text="검색 항목:").grid(row=0, column=0, sticky='w', padx=5, pady=3)
        display_to_col = {FIELD_DISPLAY_NAME.get(c, c): c for c in REPORT_SEARCH_FIELDS}
        self._tab3_field_map = display_to_col
        self.tab3_field = ttk.Combobox(top, values=list(display_to_col.keys()),
                                       width=18, state='readonly')
        self.tab3_field.grid(row=0, column=1, padx=5)
        self.tab3_field.current(0)
        self.tab3_field.bind('<<ComboboxSelected>>', self._on_tab3_field_change)

        tk.Label(top, text="검색어:").grid(row=0, column=2, sticky='w', padx=5)
        self.tab3_value = ttk.Combobox(top, width=30)
        self.tab3_value.grid(row=0, column=3, padx=5)
        self.tab3_value.bind('<Return>', lambda e: self.search_report())

        tk.Button(top, text="🔍 검색", command=self.search_report,
                  bg=COLOR_PRIMARY, fg='white',
                  font=('맑은 고딕', 10, 'bold'), width=10
                  ).grid(row=0, column=4, padx=10)

        tk.Button(top, text="↻ 초기화", command=self._reset_tab3, width=8
                  ).grid(row=0, column=5, padx=5)

        tk.Label(top, text="(부분 일치 검색)", fg='gray',
                 font=('맑은 고딕', 9)).grid(row=0, column=6, padx=10)

        self.tab3_only_with_report = tk.BooleanVar(value=True)
        tk.Checkbutton(top, text="보고서 보유만 표시",
                       variable=self.tab3_only_with_report,
                       font=('맑은 고딕', 9)
                       ).grid(row=0, column=7, padx=10)

        self.tab3_status_label = tk.Label(
            self.tab3,
            text="검색 항목을 고르고 검색어를 선택 또는 입력 후 검색을 누르세요.",
            fg=COLOR_PRIMARY, font=('맑은 고딕', 10, 'bold'), anchor='w'
        )
        self.tab3_status_label.pack(fill='x', padx=10, pady=(5, 2))

        self.tree3 = self._make_tree(self.tab3, with_similarity=False)
        self.tree3.bind('<Double-1>', self._on_tree3_double_click)

        self._on_tab3_field_change()

    def _on_tab3_field_change(self, event=None):
        display_name = (self.tab3_field.get() or '').strip()
        col = self._tab3_field_map.get(display_name)
        if not col or col not in self.df.columns:
            self.tab3_value['values'] = []
            self.tab3_value.set('')
            return
        vals = (self.df[col].astype(str).str.strip()
                .replace({'nan': '', 'None': '', 'NaT': ''}))
        uniques = sorted({v for v in vals.tolist() if v})
        self.tab3_value['values'] = uniques
        self.tab3_value.set('')

    def _reset_tab3(self):
        self.tab3_field.current(0)
        self._on_tab3_field_change()
        for i in self.tree3.get_children():
            self.tree3.delete(i)
        self.tab3_status_label.config(
            text="검색 항목을 고르고 검색어를 선택 또는 입력 후 검색을 누르세요."
        )

    def search_report(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("알림", "데이터가 없습니다.")
            return

        display_name = (self.tab3_field.get() or '').strip()
        keyword = (self.tab3_value.get() or '').strip()

        col = self._tab3_field_map.get(display_name)
        if not col:
            if display_name in self.df.columns:
                col = display_name
            else:
                for real, disp in FIELD_DISPLAY_NAME.items():
                    if disp == display_name:
                        col = real
                        break

        if not col:
            messagebox.showwarning("알림",
                f"검색 항목 '{display_name}'에 해당하는 컬럼을 찾을 수 없습니다.")
            return
        if not keyword:
            messagebox.showwarning("알림", "검색어를 선택하거나 입력하세요.")
            return
        if col not in self.df.columns:
            messagebox.showerror("오류",
                f"컬럼 '{col}' 이(가) 데이터에 없습니다.\n"
                f"사용 가능 컬럼 예: {list(self.df.columns)[:10]}")
            return

        series = (self.df[col].astype(str).str.strip()
                  .replace({'nan': '', 'None': '', 'NaT': ''}))
        kw = keyword.strip()

        mask = series.str.contains(kw, case=False, na=False, regex=False)
        if not mask.any():
            mask = (series == kw)

        result = self.df[mask].copy()
        n_total = len(result)

        n_with_report = 0
        if self.loader is not None and n_total > 0:
            has_mask = result['현장코드'].astype(str).apply(self._has_report)
            n_with_report = int(has_mask.sum())
            if self.tab3_only_with_report.get():
                result = result[has_mask]

        for i in self.tree3.get_children():
            self.tree3.delete(i)

        self.tab3_status_label.config(
            text=(f"검색 결과 — 컬럼 '{col}', 키워드 '{kw}'  |  "
                  f"일치 {n_total}건, 보고서 보유 {n_with_report}건, "
                  f"화면 표시 {len(result)}건")
        )

        if result.empty:
            if n_total == 0:
                sample_vals = [v for v in series.tolist() if v][:5]
                msg = (f"'{col}' 컬럼에서 '{kw}'와 일치하는 항목이 없습니다.\n\n"
                       f"• 컬럼 샘플값: {sample_vals}\n"
                       f"• 드롭다운에서 선택한 정확한 값을 사용해 보세요.")
            else:
                msg = (f"일치 항목 {n_total}건 중 보고서 보유 항목이 0건입니다.\n\n"
                       f"'보고서 보유만 표시' 체크를 해제하면 전체 일치 항목을 확인할 수 있습니다.")
            messagebox.showinfo("결과", msg)
            return

        for _, row in result.iterrows():
            code = str(row.get('현장코드', '')).strip()
            values = [fmt(row.get(c), c) for c, _ in DISPLAY_FIELDS]
            has = '✓' if self._has_report(code) else EMPTY_MARK
            values.append(has)
            self.tree3.insert('', 'end', values=values, tags=(code,))

    # ──────────────────────────────────────────────────────────
    # Tab 4: 보고서 관리 (드래그앤드롭 / 파일선택으로 추가)
    # ──────────────────────────────────────────────────────────
    def _build_tab4(self):
        # 상단 안내
        top = tk.LabelFrame(self.tab4, text="보고서 추가",
                            font=('맑은 고딕', 11, 'bold'), padx=10, pady=10)
        top.pack(fill='x', padx=5, pady=5)

        tk.Label(top, text=(
            "PDF/PPT 보고서 파일을 아래 영역에 드래그하거나 '파일 선택' 버튼으로 추가합니다.\n"
            "파일은 자동으로 [현장코드_건축물종류_현장제목_연도] 형식으로 이름이 변경되어 패키지에 추가됩니다."
        ), font=('맑은 고딕', 9), fg=COLOR_GRAY, justify='left',
            anchor='w').pack(fill='x', padx=5, pady=2)

        btn_frame = tk.Frame(top)
        btn_frame.pack(fill='x', pady=5)
        tk.Button(btn_frame, text="📁 파일 선택", command=self._select_files_for_add,
                  bg=COLOR_PRIMARY, fg='white',
                  font=('맑은 고딕', 10, 'bold'), width=14
                  ).pack(side='left', padx=5)
        tk.Button(btn_frame, text="🔄 패키지 리빌드", command=self._rebuild_package,
                  bg=COLOR_ACCENT, fg='white',
                  font=('맑은 고딕', 10, 'bold'), width=16
                  ).pack(side='left', padx=5)
        tk.Label(btn_frame,
                 text="(리빌드: 대기열 파일을 일괄 처리하여 reports.pkg 갱신)",
                 fg='gray', font=('맑은 고딕', 8)).pack(side='left', padx=10)

        # 드래그앤드롭 영역
        self.drop_frame = tk.Frame(self.tab4, bg='#E8F0FE',
                                   highlightbackground=COLOR_PRIMARY,
                                   highlightthickness=2, height=120)
        self.drop_frame.pack(fill='x', padx=10, pady=5)
        self.drop_frame.pack_propagate(False)
        self.drop_label = tk.Label(
            self.drop_frame,
            text="📥  여기에 보고서 파일(PDF/PPT)을 드래그하세요",
            bg='#E8F0FE', fg=COLOR_PRIMARY,
            font=('맑은 고딕', 12, 'bold')
        )
        self.drop_label.pack(expand=True)

        # 드래그앤드롭 바인딩 (tkinterdnd2 사용 가능시)
        if HAS_DND and isinstance(self.root, TkinterDnD.Tk):
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self._on_drop_files)
            self.drop_frame.dnd_bind('<<DragEnter>>', self._on_drag_enter)
            self.drop_frame.dnd_bind('<<DragLeave>>', self._on_drag_leave)
        else:
            self.drop_label.config(
                text="📥  드래그앤드롭을 사용하려면 tkinterdnd2를 설치하세요\n"
                     "     (pip install tkinterdnd2)\n"
                     "     또는 위 '파일 선택' 버튼을 사용하세요",
                font=('맑은 고딕', 10)
            )

        # 대기열 Treeview
        queue_label = tk.Label(self.tab4, text="추가 대기열",
                               font=('맑은 고딕', 10, 'bold'),
                               fg=COLOR_PRIMARY, anchor='w')
        queue_label.pack(fill='x', padx=10, pady=(10, 2))

        q_cols = ('원본파일명', '현장코드', '건축물종류', '현장제목', '연도',
                  '변환파일명', '상태')
        q_frame = tk.Frame(self.tab4, bg='white')
        q_frame.pack(fill='both', expand=True, padx=5, pady=5)

        self.queue_tree = ttk.Treeview(q_frame, columns=q_cols,
                                        show='headings', height=15)
        q_widths = {'원본파일명': 300, '현장코드': 80, '건축물종류': 100,
                    '현장제목': 220, '연도': 60, '변환파일명': 320, '상태': 80}
        for c in q_cols:
            self.queue_tree.heading(c, text=c)
            self.queue_tree.column(c, width=q_widths.get(c, 100), anchor='w')

        vsb = ttk.Scrollbar(q_frame, orient='vertical',
                             command=self.queue_tree.yview)
        self.queue_tree.configure(yscrollcommand=vsb.set)
        self.queue_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        q_frame.rowconfigure(0, weight=1)
        q_frame.columnconfigure(0, weight=1)

        # 더블클릭으로 편집
        self.queue_tree.bind('<Double-1>', self._on_queue_double_click)

        # 상태 라벨
        self.tab4_status = tk.Label(
            self.tab4, text="파일을 추가하면 대기열에 표시됩니다.",
            fg=COLOR_PRIMARY, font=('맑은 고딕', 9), anchor='w'
        )
        self.tab4_status.pack(fill='x', padx=10, pady=(0, 5))

        # 내부 큐 데이터
        self._add_queue = []  # list of dicts

    def _on_drag_enter(self, event):
        self.drop_frame.config(bg='#C5D9F1')
        self.drop_label.config(bg='#C5D9F1', text="📥  놓으면 추가됩니다!")

    def _on_drag_leave(self, event):
        self.drop_frame.config(bg='#E8F0FE')
        self.drop_label.config(bg='#E8F0FE',
                               text="📥  여기에 보고서 파일(PDF/PPT)을 드래그하세요")

    def _on_drop_files(self, event):
        self.drop_frame.config(bg='#E8F0FE')
        self.drop_label.config(bg='#E8F0FE',
                               text="📥  여기에 보고서 파일(PDF/PPT)을 드래그하세요")
        # 파싱: tkinterdnd2는 경로를 중괄호나 공백으로 구분
        raw = event.data
        files = []
        if '{' in raw:
            # 중괄호로 묶인 경로들
            import re
            files = re.findall(r'\{(.+?)\}', raw)
            # 나머지 중괄호 밖의 경로
            rest = re.sub(r'\{.+?\}', '', raw).strip()
            if rest:
                files.extend(rest.split())
        else:
            files = raw.split()

        valid = [f for f in files
                 if Path(f).suffix.lower() in ('.pdf', '.pptx', '.ppt')]
        if not valid:
            messagebox.showinfo("알림", "PDF 또는 PPT 파일만 추가할 수 있습니다.")
            return
        self._process_dropped_files(valid)

    def _select_files_for_add(self):
        files = filedialog.askopenfilenames(
            title="보고서 파일 선택",
            filetypes=[
                ("보고서 파일", "*.pdf *.pptx *.ppt"),
                ("PDF", "*.pdf"),
                ("PowerPoint", "*.pptx *.ppt"),
            ]
        )
        if files:
            self._process_dropped_files(list(files))

    def _process_dropped_files(self, file_paths):
        """드롭/선택된 파일을 분석하여 대기열에 추가"""
        for fpath in file_paths:
            fp = Path(fpath)
            if not fp.exists():
                continue

            # 현장코드 입력 대화상자
            info = self._ask_report_info(fp)
            if info is None:
                continue  # 사용자 취소

            site_code = info['site_code']
            bldg_type = info['bldg_type']
            title = info['title']
            year = info['year']
            ext = fp.suffix.lower()
            if ext == '.ppt':
                ext = '.pptx'

            new_name = f"{site_code}_{bldg_type}_{title}_{year}{ext}"

            entry = {
                'original_path': str(fp),
                'original_name': fp.name,
                'site_code': site_code,
                'bldg_type': bldg_type,
                'title': title,
                'year': year,
                'new_name': new_name,
                'status': '대기',
            }
            self._add_queue.append(entry)
            self.queue_tree.insert('', 'end', values=(
                fp.name, site_code, bldg_type, title, year, new_name, '대기'
            ))

        n = len(self._add_queue)
        self.tab4_status.config(
            text=f"대기열: {n}건  |  '패키지 리빌드' 버튼을 눌러 일괄 처리하세요."
        )

    def _ask_report_info(self, filepath: Path):
        """보고서 정보 입력 대화상자"""
        dialog = tk.Toplevel(self.root)
        dialog.title("보고서 정보 입력")
        dialog.geometry("500x320")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        result = {'cancelled': True}

        tk.Label(dialog, text=f"파일: {filepath.name}",
                 font=('맑은 고딕', 9, 'bold'), fg=COLOR_PRIMARY
                 ).grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky='w')

        tk.Label(dialog, text="현장코드 *:").grid(row=1, column=0, padx=10, pady=5, sticky='w')
        code_entry = tk.Entry(dialog, width=20)
        code_entry.grid(row=1, column=1, padx=10, pady=5, sticky='w')

        tk.Label(dialog, text="건축물종류 *:").grid(row=2, column=0, padx=10, pady=5, sticky='w')
        bldg_combo = ttk.Combobox(dialog, width=18, state='normal')
        # 기존 데이터에서 건축물종류 목록 가져오기
        try:
            types = sorted({str(t).strip() for t in get_building_types(self.df)
                           if t and str(t).strip().lower() not in ('nan', 'none', '')})
        except Exception:
            types = []
        bldg_combo['values'] = types
        bldg_combo.grid(row=2, column=1, padx=10, pady=5, sticky='w')

        tk.Label(dialog, text="현장제목 *:").grid(row=3, column=0, padx=10, pady=5, sticky='w')
        title_entry = tk.Entry(dialog, width=40)
        title_entry.grid(row=3, column=1, padx=10, pady=5, sticky='w')

        tk.Label(dialog, text="연도 *:").grid(row=4, column=0, padx=10, pady=5, sticky='w')
        year_entry = tk.Entry(dialog, width=10)
        year_entry.grid(row=4, column=1, padx=10, pady=5, sticky='w')

        # 자동 채우기 시도 (data.xlsx에서)
        tk.Label(dialog, text="현장코드 입력 후 Tab을 누르면 자동 채우기",
                 fg='gray', font=('맑은 고딕', 8)
                 ).grid(row=5, column=0, columnspan=2, padx=10, pady=2, sticky='w')

        def auto_fill(event=None):
            code = code_entry.get().strip()
            if not code:
                return
            try:
                code_num = int(float(code))
                row = self.df[self.df['현장코드'].astype(float).astype(int) == code_num]
                if not row.empty:
                    r = row.iloc[0]
                    bt = str(r.get('건축물종류', '')).strip()
                    pn = str(r.get('프로젝트명', '')).strip()
                    yr = str(r.get('착공연도', '')).strip()
                    if bt and bt.lower() not in ('nan', 'none'):
                        bldg_combo.set(bt)
                    if pn and pn.lower() not in ('nan', 'none'):
                        title_entry.delete(0, tk.END)
                        title_entry.insert(0, pn)
                    if yr and yr.lower() not in ('nan', 'none'):
                        try:
                            year_entry.delete(0, tk.END)
                            year_entry.insert(0, str(int(float(yr))))
                        except Exception:
                            pass
            except Exception:
                pass

        code_entry.bind('<FocusOut>', auto_fill)
        code_entry.bind('<Tab>', auto_fill)

        def on_ok():
            code = code_entry.get().strip()
            bldg = bldg_combo.get().strip()
            title = title_entry.get().strip()
            year = year_entry.get().strip()

            if not code or not code.isdigit():
                messagebox.showwarning("알림", "현장코드를 숫자로 입력하세요.", parent=dialog)
                return
            if not bldg:
                messagebox.showwarning("알림", "건축물종류를 입력하세요.", parent=dialog)
                return
            if not title:
                messagebox.showwarning("알림", "현장제목을 입력하세요.", parent=dialog)
                return
            if not year:
                messagebox.showwarning("알림", "연도를 입력하세요.", parent=dialog)
                return

            result['cancelled'] = False
            result['site_code'] = code
            result['bldg_type'] = bldg
            result['title'] = title
            result['year'] = year
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=15)
        tk.Button(btn_frame, text="확인", command=on_ok,
                  bg=COLOR_PRIMARY, fg='white',
                  font=('맑은 고딕', 10, 'bold'), width=10
                  ).pack(side='left', padx=10)
        tk.Button(btn_frame, text="취소", command=on_cancel,
                  width=10).pack(side='left', padx=10)

        dialog.wait_window()

        if result['cancelled']:
            return None
        return result

    def _rebuild_package(self):
        """대기열의 파일들을 Reports/에 복사하고 패키지 리빌드"""
        if not self._add_queue:
            messagebox.showinfo("알림", "대기열에 파일이 없습니다.\n"
                                "먼저 보고서 파일을 추가하세요.")
            return

        # 파일 복사
        copied_codes = []
        for entry in self._add_queue:
            if entry['status'] != '대기':
                continue
            src = Path(entry['original_path'])
            dst = REPORTS_ORIGINAL_DIR / entry['new_name']
            try:
                import shutil
                shutil.copy2(str(src), str(dst))
                entry['status'] = '복사완료'
                copied_codes.append(entry['site_code'])
                print(f"[복사] {src.name} → {dst.name}")
            except Exception as e:
                entry['status'] = f'오류: {e}'
                print(f"[오류] {src.name}: {e}")

        if not copied_codes:
            messagebox.showwarning("알림", "복사할 파일이 없습니다.")
            return

        # reports_filelist.xlsx 업데이트
        self._update_filelist(copied_codes)

        # 대기열 트리뷰 갱신
        for item in self.queue_tree.get_children():
            self.queue_tree.delete(item)
        for entry in self._add_queue:
            self.queue_tree.insert('', 'end', values=(
                entry['original_name'], entry['site_code'],
                entry['bldg_type'], entry['title'], entry['year'],
                entry['new_name'], entry['status']
            ))

        # 패키지 리빌드 (백그라운드 스레드)
        self.tab4_status.config(
            text=f"패키지 리빌드 중... ({len(copied_codes)}건 처리)")

        def do_rebuild():
            try:
                from add_reports import add_reports
                add_reports(copied_codes)
                self.root.after(0, self._on_rebuild_done, True, "")
            except Exception as e:
                traceback.print_exc()
                self.root.after(0, self._on_rebuild_done, False, str(e))

        t = threading.Thread(target=do_rebuild, daemon=True)
        t.start()

    def _on_rebuild_done(self, success, error_msg):
        """패키지 리빌드 완료 콜백"""
        if success:
            # 패키지 리로드
            try:
                self.loader = SecurePackageLoader(PACKAGE_FILE)
                if not self.loader.is_loaded():
                    self.loader = None
            except Exception:
                self.loader = None

            # 대기열 상태 갱신
            for entry in self._add_queue:
                if entry['status'] == '복사완료':
                    entry['status'] = '✓ 완료'
            for item in self.queue_tree.get_children():
                self.queue_tree.delete(item)
            for entry in self._add_queue:
                self.queue_tree.insert('', 'end', values=(
                    entry['original_name'], entry['site_code'],
                    entry['bldg_type'], entry['title'], entry['year'],
                    entry['new_name'], entry['status']
                ))

            done_count = sum(1 for e in self._add_queue if e['status'] == '✓ 완료')
            total = len(self.loader.manifest['reports']) if self.loader and self.loader.manifest else '?'
            self.tab4_status.config(
                text=f"✓ 패키지 업데이트 완료! {done_count}건 추가/갱신, 총 {total}건")
            self._toast("패키지 업데이트 완료!")
            # 완료된 항목 큐에서 제거
            self._add_queue = [e for e in self._add_queue if e['status'] != '✓ 완료']
        else:
            self.tab4_status.config(
                text=f"✗ 패키지 리빌드 실패: {error_msg}")
            messagebox.showerror("오류", f"패키지 리빌드 실패:\n{error_msg}")

    def _update_filelist(self, codes):
        """reports_filelist.xlsx에 신규 항목 추가"""
        try:
            filelist_path = DATA_FILE.parent / 'reports_filelist.xlsx'
            if not filelist_path.exists():
                return

            fl = pd.read_excel(filelist_path)
            # 합계 행 제거
            fl = fl[fl['No.'] != '합계'].copy()
            fl['코드번호'] = pd.to_numeric(fl['코드번호'], errors='coerce')
            fl = fl.dropna(subset=['코드번호'])
            fl['코드번호'] = fl['코드번호'].astype(int)

            for entry in self._add_queue:
                code = int(entry['site_code'])
                if code in fl['코드번호'].values:
                    # 기존 항목 업데이트
                    mask = fl['코드번호'] == code
                    fl.loc[mask, '파일명'] = entry['new_name']
                    fl.loc[mask, '파일형식'] = ('PDF' if entry['new_name'].endswith('.pdf')
                                              else 'PPTX')
                else:
                    # 신규 항목 추가
                    new_row = {
                        '코드번호': code,
                        '시설유형': entry['bldg_type'],
                        '프로젝트명': entry['title'],
                        '준공연도': entry['year'],
                        '파일형식': ('PDF' if entry['new_name'].endswith('.pdf')
                                    else 'PPTX'),
                        '파일명': entry['new_name'],
                    }
                    fl = pd.concat([fl, pd.DataFrame([new_row])],
                                   ignore_index=True)

            fl = fl.sort_values('코드번호').reset_index(drop=True)
            fl['No.'] = range(1, len(fl) + 1)

            total = len(fl)
            summary = pd.DataFrame([{'No.': '합계', '코드번호': f'총 {total}건'}])
            fl_out = pd.concat([fl, summary], ignore_index=True)
            fl_out.to_excel(filelist_path, index=False)
            print(f"[INFO] reports_filelist.xlsx 업데이트: 총 {total}건")

        except Exception as e:
            print(f"[WARN] reports_filelist.xlsx 업데이트 실패: {e}")

    def _on_queue_double_click(self, event):
        """대기열 항목 더블클릭 시 편집"""
        item = self.queue_tree.focus()
        if not item:
            return
        values = self.queue_tree.item(item, 'values')
        if not values:
            return

        # 해당 큐 항목 찾기
        orig_name = values[0]
        idx = None
        for i, entry in enumerate(self._add_queue):
            if entry['original_name'] == orig_name and entry['status'] == '대기':
                idx = i
                break
        if idx is None:
            return

        entry = self._add_queue[idx]

        # 편집 대화상자
        dialog = tk.Toplevel(self.root)
        dialog.title("보고서 정보 수정")
        dialog.geometry("450x250")
        dialog.transient(self.root)
        dialog.grab_set()

        fields = [
            ('현장코드', entry['site_code']),
            ('건축물종류', entry['bldg_type']),
            ('현장제목', entry['title']),
            ('연도', entry['year']),
        ]
        entries = {}
        for r, (label, val) in enumerate(fields):
            tk.Label(dialog, text=f"{label}:").grid(row=r, column=0, padx=10, pady=5, sticky='w')
            e = tk.Entry(dialog, width=35)
            e.insert(0, val)
            e.grid(row=r, column=1, padx=10, pady=5, sticky='w')
            entries[label] = e

        def on_save():
            entry['site_code'] = entries['현장코드'].get().strip()
            entry['bldg_type'] = entries['건축물종류'].get().strip()
            entry['title'] = entries['현장제목'].get().strip()
            entry['year'] = entries['연도'].get().strip()
            ext = Path(entry['original_path']).suffix.lower()
            if ext == '.ppt':
                ext = '.pptx'
            entry['new_name'] = (f"{entry['site_code']}_{entry['bldg_type']}_"
                                 f"{entry['title']}_{entry['year']}{ext}")
            # 트리뷰 갱신
            self.queue_tree.item(item, values=(
                entry['original_name'], entry['site_code'],
                entry['bldg_type'], entry['title'], entry['year'],
                entry['new_name'], entry['status']
            ))
            dialog.destroy()

        tk.Button(dialog, text="저장", command=on_save,
                  bg=COLOR_PRIMARY, fg='white', width=10
                  ).grid(row=len(fields), column=0, columnspan=2, pady=15)

        dialog.wait_window()

    # ──────────────────────────────────────────────────────────
    # 공통 헬퍼
    # ──────────────────────────────────────────────────────────
    def _make_tree(self, parent, with_similarity=False):
        cols = [disp for disp, _ in DISPLAY_FIELDS] + ['보고서']
        if with_similarity:
            cols.append('유사도')

        frame = tk.Frame(parent, bg='white')
        frame.pack(fill='both', expand=True, padx=5, pady=5)

        tree = ttk.Treeview(frame, columns=cols, show='headings', height=25)
        widths = {
            '현장코드': 70, '프로젝트명': 200, '착공연도': 70,
            '건축물종류': 90, '건축행위': 80, '연면적(㎡)': 90,
            '층수': 70, '공사금액': 80, '공사개월': 70,
            '외장마감1': 90, '외장마감2': 90,
            '지하 구조': 90, '지상 구조': 90,
            '구조체': 80, '기초': 80,
            '역타공법1': 80, '역타공법2': 80,
            '흙막이 공법1': 90, '흙막이 공법2': 90,
            '보고서': 60, '유사도': 70,
        }
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=widths.get(c, 80),
                        anchor='center' if c in ('보고서', '유사도', '착공연도') else 'w')

        vsb = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def _has_report(self, site_code):
        if not site_code or self.loader is None:
            return False
        s = str(site_code).strip()
        candidates = [s]
        if s.endswith('.0'):
            candidates.append(s[:-2])
        if '.' not in s:
            candidates.append(s + '.0')
        try:
            candidates.append(str(int(float(s))))
        except Exception:
            pass
        for cand in candidates:
            try:
                if self.loader.has_report(cand):
                    return True
            except Exception:
                continue
        return False

    def _get_project_name(self, site_code):
        row = self.df[self.df['현장코드'].astype(str) == str(site_code)]
        if not row.empty:
            return str(row.iloc[0].get('프로젝트명', ''))
        return ''

    def _open_viewer(self, site_code, start_page=1):
        if self.loader is None:
            messagebox.showwarning("알림", "보고서 패키지가 로드되지 않았습니다.")
            return
        if not self._has_report(site_code):
            messagebox.showinfo("알림", "해당 프로젝트의 준공보고서가 없습니다.")
            return
        row = self.df[self.df['현장코드'].astype(str) == str(site_code)]
        info = row.iloc[0].to_dict() if not row.empty else {}
        try:
            SecureReportViewer(self.root, self.loader, str(site_code),
                               project_info=info, start_page=start_page)
        except Exception as e:
            messagebox.showerror("오류", f"보고서 뷰어 실행 실패:\n{e}")

    def _on_tree1_double_click(self, event):
        item = self.tree1.focus()
        if not item:
            return
        code = self.tree1.item(item, 'tags')[0]
        self._open_viewer(code)

    def _on_tree2_double_click(self, event):
        item = self.tree2.focus()
        if not item:
            return
        tags = self.tree2.item(item, 'tags')
        code = tags[0] if tags else ''
        page = int(tags[1]) if len(tags) > 1 else 1
        self._open_viewer(code, start_page=page)

    def _on_tree3_double_click(self, event):
        item = self.tree3.focus()
        if not item:
            return
        code = self.tree3.item(item, 'tags')[0]
        self._open_viewer(code)

    # ──────────────────────────────────────────────────────────
    # PDF 출력  ★★★ 핵심 수정 영역 ★★★
    # ──────────────────────────────────────────────────────────
    def export_pdf(self):
        if self.last_search_result is None or self.last_search_result.empty:
            messagebox.showwarning("알림", "먼저 프로젝트 검색을 수행하세요.")
            return
        try:
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            out = OUTPUT_DIR / f"reference_{ts}.pdf"
            # 시그니처: (top_df, target, output_path, has_report_func, user)
            generate_reference_pdf(
                self.last_search_result.head(3),   # TOP3 카드
                self.last_search_query or {},
                out,
                has_report_func=self._has_report,
                user=self.user,
            )
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("오류", f"PDF 생성 실패:\n{type(e).__name__}: {e}")
            return

        opened = self._open_file_with_default_app(out)
        msg = f"PDF 생성 완료\n{out.name}" + ("  (기본 뷰어로 열림)" if opened else "")
        self._toast(msg)

    def _open_file_with_default_app(self, filepath):
        try:
            if sys.platform.startswith('win'):
                os.startfile(str(filepath))
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(filepath)])
            else:
                subprocess.Popen(['xdg-open', str(filepath)])
            return True
        except Exception as e:
            print(f"[WARN] 파일 자동 실행 실패: {e}")
            return False

    def _toast(self, message, duration_ms=2500):
        try:
            toast = tk.Toplevel(self.root)
            toast.overrideredirect(True)
            toast.configure(bg=COLOR_PRIMARY)
            w, h = 380, 80
            self.root.update_idletasks()
            x = self.root.winfo_rootx() + self.root.winfo_width() - w - 30
            y = self.root.winfo_rooty() + self.root.winfo_height() - h - 60
            toast.geometry(f"{w}x{h}+{x}+{y}")
            tk.Label(toast, text=message, bg=COLOR_PRIMARY, fg='white',
                     font=('맑은 고딕', 10, 'bold'), justify='left'
                     ).pack(expand=True, fill='both', padx=12, pady=10)
            toast.after(duration_ms, toast.destroy)
        except Exception:
            messagebox.showinfo("알림", message)


# ──────────────────────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────────────────────
def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
        print("[INFO] tkinterdnd2 활성화 - 드래그앤드롭 사용 가능")
    else:
        root = tk.Tk()
        print("[INFO] tkinterdnd2 미설치 - 파일선택 대화상자만 사용 가능")
        print("       드래그앤드롭을 사용하려면: pip install tkinterdnd2")
    app = CJProjectHubApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
