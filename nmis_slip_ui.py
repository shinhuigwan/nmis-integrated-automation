"""
NMIS (통합) 전표 & 26년 월보고 자동화 시스템 (Modern CustomTkinter UI Edition)
- Dark Mode / Neon Accent Palette / Card-based Responsive Layout
- Playwright CDP 연동, 실시간 세입세출표, 회원현황, 4번째 시트, 5번째 직원가입 시트 및 발송대장 5건 자동 등록
"""

import os
import sys
import json
import calendar
import datetime
from datetime import date
from pathlib import Path
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Callable

import customtkinter as ctk
from playwright.sync_api import sync_playwright

# ── 모듈 임포트 ─────────────────────────────────────────────────────────────
CDP_URL = "http://127.0.0.1:9222"

from nmis_slip_automation import (
    MemberVerificationResult,
    Transaction,
    fetch_sheet4_data_only,
    fill_all_monthly_reports_sequentially,
    fill_member_status_from_nmis,
    fill_monthly_member_and_revenue_report,
    fill_monthly_report_from_nmis,
    fill_staff_join_excel_from_data,
    find_nmis_page,
    read_excel,
    register_ship_documents_on_nmis,
    verify_member_info_from_nmis,
)

def find_chrome() -> Path | None:
    candidates = (
        Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
        Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
        Path(os.environ.get("LOCALAPPDATA", "")) / "Google" / "Chrome" / "Application" / "chrome.exe",
    )
    return next((p for p in candidates if p.is_file()), None)

def open_attachable_chrome() -> bool:
    chrome = find_chrome()
    if not chrome:
        return False
    profile = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "nmis-cdp-profile"
    try:
        subprocess.Popen([str(chrome), "--remote-debugging-port=9222",
                          f"--user-data-dir={profile}", "http://nmis.foodservice.or.kr/"],
                         close_fds=True)
        return True
    except Exception:
        return False

# ── CustomTkinter 글로벌 설정 ───────────────────────────────────────────────
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ── 설정 파일 및 글로벌 변수 ────────────────────────────────────────────────
SETTINGS_FILE = Path(__file__).parent / "settings.json"

NMIS_USER_ID = "e20240056"
NMIS_PASSWORD = "a4848665"

KEYWORD_RULES: dict[str, str] = {
    "회비": "member_fee",
    "월회비": "member_fee",
    "입회비": "join_fee",
    "가입금": "join_fee",
    "급여": "salary",
    "상여": "salary",
}

CONFIRM_SELECTORS = [
    "button:has-text('확인')",
    "a:has-text('확인')",
    "input[value='확인']",
    ".btn-primary:has-text('확인')",
]

MACRO_STEPS: list[dict] = []

def load_settings() -> None:
    global KEYWORD_RULES, CONFIRM_SELECTORS, MACRO_STEPS, NMIS_USER_ID, NMIS_PASSWORD
    if not SETTINGS_FILE.is_file():
        return
    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        if "keyword_rules" in data:
            KEYWORD_RULES = data["keyword_rules"]
        if "confirm_selectors" in data:
            CONFIRM_SELECTORS = data["confirm_selectors"]
        if "macro_steps" in data:
            MACRO_STEPS = data["macro_steps"]
        if "nmis_user_id" in data:
            NMIS_USER_ID = data["nmis_user_id"]
        if "nmis_password" in data:
            NMIS_PASSWORD = data["nmis_password"]
    except Exception as e:
        print(f"설정 로드 실패: {e}")

def save_settings() -> None:
    data = {
        "keyword_rules": KEYWORD_RULES,
        "confirm_selectors": CONFIRM_SELECTORS,
        "macro_steps": MACRO_STEPS,
        "nmis_user_id": NMIS_USER_ID,
        "nmis_password": NMIS_PASSWORD,
    }
    try:
        SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"설정 저장 실패: {e}")

def cdp_is_ready() -> bool:
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.connect_over_cdp(CDP_URL)
            page = find_nmis_page(browser)
            return page is not None
    except Exception:
        return False


# ── 메인 UI 클래스 (CustomTkinter 기반) ──────────────────────────────────────

class ModernSlipUI(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        load_settings()

        self.title("통합 자동화 시스템 — Modern Dark Edition")
        self.geometry("1100x760")
        self.minsize(980, 680)

        # 변수 데이터
        self.file_var = tk.StringVar()
        self.browser_status_var = tk.StringVar(value="🔴 Chrome 연결 확인 전")
        self.run_status_var = tk.StringVar(value="대기")
        self.running = False

        self.all_transactions: dict[str, Transaction] = {}
        self.tx_type: dict[str, str] = {}
        self.tx_settings: dict[str, dict] = {}

        # 메인 그리드 구성 (사이드바 0, 메인 1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main_area()

        # 거래내역 초기화
        if len(sys.argv) > 1:
            candidate = Path(sys.argv[1]).expanduser()
            if candidate.is_file():
                self.file_var.set(str(candidate.resolve()))
                self.after(100, self.load_excel)

        self.after(300, self.check_browser_connection)

    # ── 사이드바 ────────────────────────────────────────────────────────────

    def _build_sidebar(self) -> None:
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#141126")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(6, weight=1)

        # 로고
        logo_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_frame.grid(row=0, column=0, padx=20, pady=(24, 20), sticky="ew")

        ctk.CTkLabel(
            logo_frame,
            text="⚡ 통합 자동화",
            font=ctk.CTkFont(family="맑은 고딕", size=20, weight="bold"),
            text_color="#A855F7"
        ).pack(anchor="w")
        ctk.CTkLabel(
            logo_frame,
            text="전표 & 26년 월보고 시스템",
            font=ctk.CTkFont(family="맑은 고딕", size=10),
            text_color="#94A3B8"
        ).pack(anchor="w")

        # 탭 선택 버튼
        self.btn_tab_monthly = ctk.CTkButton(
            self.sidebar,
            text="📊 26년 월보고 자동 연동",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            height=42,
            corner_radius=10,
            command=lambda: self._select_tab("monthly")
        )
        self.btn_tab_monthly.grid(row=1, column=0, padx=16, pady=6, sticky="ew")

        self.btn_tab_slip = ctk.CTkButton(
            self.sidebar,
            text="📋 전표 자동등록",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
            fg_color="transparent",
            hover_color="#26214A",
            text_color="#CBD5E1",
            height=42,
            corner_radius=10,
            command=lambda: self._select_tab("slip")
        )
        self.btn_tab_slip.grid(row=2, column=0, padx=16, pady=6, sticky="ew")

        self.btn_tab_member = ctk.CTkButton(
            self.sidebar,
            text="👥 회원 정보 검수",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
            fg_color="transparent",
            hover_color="#26214A",
            text_color="#CBD5E1",
            height=42,
            corner_radius=10,
            command=lambda: self._select_tab("member")
        )
        self.btn_tab_member.grid(row=3, column=0, padx=16, pady=6, sticky="ew")

        # Chrome Status Box in Sidebar
        status_box = ctk.CTkFrame(self.sidebar, fg_color="#1E1B3A", corner_radius=12)
        status_box.grid(row=5, column=0, padx=16, pady=16, sticky="ew")

        ctk.CTkLabel(
            status_box,
            textvariable=self.browser_status_var,
            font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold"),
            text_color="#10B981"
        ).pack(anchor="w", padx=12, pady=(12, 8))

        ctk.CTkButton(
            status_box,
            text="🚀 Chrome 연결 열기",
            font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold"),
            fg_color="#10B981",
            hover_color="#059669",
            text_color="#FFFFFF",
            height=32,
            corner_radius=8,
            command=self.start_attachable_chrome
        ).pack(fill="x", padx=12, pady=(0, 10))

        ctk.CTkButton(
            status_box,
            text="⚙ 계정 및 시스템 설정",
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            fg_color="#374151",
            hover_color="#4B5563",
            text_color="#E2E8F0",
            height=30,
            corner_radius=8,
            command=self.open_settings
        ).pack(fill="x", padx=12, pady=(0, 12))

    # ── 메인 영역 ────────────────────────────────────────────────────────────

    def _build_main_area(self) -> None:
        self.main_container = ctk.CTkFrame(self, fg_color="#0D0B18", corner_radius=0)
        self.main_container.grid(row=0, column=1, sticky="nsew")
        self.main_container.grid_rowconfigure(1, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)

        # 1. 헤더
        header = ctk.CTkFrame(self.main_container, fg_color="transparent")
        header.grid(row=0, column=0, padx=24, pady=(20, 10), sticky="ew")

        self.main_title_label = ctk.CTkLabel(
            header,
            text="📊 26년 월보고 자동 연동 시스템",
            font=ctk.CTkFont(family="맑은 고딕", size=20, weight="bold"),
            text_color="#FFFFFF"
        )
        self.main_title_label.pack(anchor="w")

        # 2. 탭 컨텐츠 프레임
        self.tab_monthly_frame = ctk.CTkScrollableFrame(self.main_container, fg_color="transparent")
        self.tab_slip_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.tab_member_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")

        self._build_monthly_tab(self.tab_monthly_frame)
        self._build_slip_tab(self.tab_slip_frame)
        self._build_member_tab(self.tab_member_frame)

        # 초기 탭 표시 (월보고 자동 연동)
        self._select_tab("monthly")

        # 3. 하단 실시간 콘솔 드로어
        self._build_console_drawer()

    def _select_tab(self, tab_name: str) -> None:
        self.tab_monthly_frame.grid_forget()
        self.tab_slip_frame.grid_forget()
        self.tab_member_frame.grid_forget()

        self.btn_tab_monthly.configure(fg_color="transparent", text_color="#CBD5E1")
        self.btn_tab_slip.configure(fg_color="transparent", text_color="#CBD5E1")
        self.btn_tab_member.configure(fg_color="transparent", text_color="#CBD5E1")

        if tab_name == "monthly":
            self.tab_monthly_frame.grid(row=1, column=0, padx=24, pady=10, sticky="nsew")
            self.btn_tab_monthly.configure(fg_color="#8B5CF6", text_color="#FFFFFF")
            self.main_title_label.configure(text="📊 26년 월보고 자동 연동 시스템")
        elif tab_name == "slip":
            self.tab_slip_frame.grid(row=1, column=0, padx=24, pady=10, sticky="nsew")
            self.btn_tab_slip.configure(fg_color="#8B5CF6", text_color="#FFFFFF")
            self.main_title_label.configure(text="📋 전표 자동등록 시스템")
        elif tab_name == "member":
            self.tab_member_frame.grid(row=1, column=0, padx=24, pady=10, sticky="nsew")
            self.btn_tab_member.configure(fg_color="#8B5CF6", text_color="#FFFFFF")
            self.main_title_label.configure(text="👥 회원 정보(대표자/신고번호) 검수 시스템")

    # ── [탭 2] 월보고 자동 연동 뷰 ──────────────────────────────────────────

    def _build_monthly_tab(self, parent: ctk.CTkScrollableFrame) -> None:
        # Card 1: 기본 설정 및 엑셀 지정
        card1 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=16)
        card1.pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(
            card1,
            text="📁 [1] 월보고 기본 설정 및 엑셀 파일 지정",
            font=ctk.CTkFont(family="맑은 고딕", size=15, weight="bold"),
            text_color="#A855F7"
        ).pack(anchor="w", padx=20, pady=(16, 12))

        # 파일 선택 줄
        file_row = ctk.CTkFrame(card1, fg_color="transparent")
        file_row.pack(fill="x", padx=20, pady=(0, 12))

        ctk.CTkLabel(file_row, text="월보고 엑셀 파일:", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#E2E8F0").pack(side="left", padx=(0, 8))

        desktop_report_file = Path.home() / "Desktop" / "26년 월보고.xls"
        default_path = str(desktop_report_file) if desktop_report_file.is_file() else ""
        self.monthly_excel_var = tk.StringVar(value=default_path)

        excel_entry = ctk.CTkEntry(
            file_row,
            textvariable=self.monthly_excel_var,
            font=ctk.CTkFont(family="맑은 고딕", size=12),
            fg_color="#120F24",
            border_color="#3B326B",
            corner_radius=8
        )
        excel_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        def browse_monthly_file():
            p = filedialog.askopenfilename(
                title="월보고 엑셀 파일 선택",
                filetypes=(("Excel 파일", "*.xls;*.xlsx"), ("모든 파일", "*.*")),
            )
            if p:
                self.monthly_excel_var.set(p)

        def analyze_monthly_file():
            p = Path(self.monthly_excel_var.get()).expanduser()
            if not p.is_file():
                messagebox.showerror("파일 오류", "월보고 엑셀 파일 경로를 확인해주세요.")
                return
            messagebox.showinfo(
                "분석 완료",
                f"월보고 파일 분석 완료!\n• 파일명: {p.name}\n• 감지된 시트 목록: 세입세출표, 회원현황, 재산변동사항, 월별 회원현황 및 세입실적, 직원회원가입실적, 자율지도추진실적, 개인별추진실적"
            )

        ctk.CTkButton(file_row, text="파일 선택", width=90, fg_color="#374151", hover_color="#4B5563", font=ctk.CTkFont(family="맑은 고딕", size=12), corner_radius=8, command=browse_monthly_file).pack(side="left", padx=(0, 6))
        ctk.CTkButton(file_row, text="파일 분석", width=90, fg_color="#374151", hover_color="#4B5563", font=ctk.CTkFont(family="맑은 고딕", size=12), corner_radius=8, command=analyze_monthly_file).pack(side="left")

        # 파라미터 줄
        param_row = ctk.CTkFrame(card1, fg_color="transparent")
        param_row.pack(fill="x", padx=20, pady=(0, 16))

        now_today = datetime.datetime.now().strftime("%Y.%m.%d")
        now_ym = datetime.datetime.now().strftime("%Y-%m")
        now_m = datetime.datetime.now().month
        default_month_label = f"{12 if now_m == 1 else now_m - 1}월말"

        ctk.CTkLabel(param_row, text="📅 작성 기준년월:", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#CBD5E1").pack(side="left", padx=(0, 6))
        self.member_ym_var = tk.StringVar(value=now_ym)
        ctk.CTkEntry(param_row, textvariable=self.member_ym_var, width=100, font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), fg_color="#120F24", border_color="#3B326B", corner_radius=8).pack(side="left", padx=(0, 24))

        ctk.CTkLabel(param_row, text="📅 발송일자:", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#CBD5E1").pack(side="left", padx=(0, 6))
        self.ship_send_date_var = tk.StringVar(value=now_today)
        ctk.CTkEntry(param_row, textvariable=self.ship_send_date_var, width=110, font=ctk.CTkFont(family="맑은 고딕", size=12), fg_color="#120F24", border_color="#3B326B", corner_radius=8).pack(side="left", padx=(0, 24))

        ctk.CTkLabel(param_row, text="🏷️ 보고월:", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#CBD5E1").pack(side="left", padx=(0, 6))
        self.ship_report_month_var = tk.StringVar(value=default_month_label)
        ctk.CTkEntry(param_row, textvariable=self.ship_report_month_var, width=80, font=ctk.CTkFont(family="맑은 고딕", size=12), fg_color="#120F24", border_color="#3B326B", corner_radius=8).pack(side="left")

        # Card 2: 원클릭 메인 자동 실행 패널
        card2 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=16)
        card2.pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(
            card2,
            text="🚀 [2] 원클릭 메인 자동 실행 패널",
            font=ctk.CTkFont(family="맑은 고딕", size=15, weight="bold"),
            text_color="#10B981"
        ).pack(anchor="w", padx=20, pady=(16, 12))

        hero_grid = ctk.CTkFrame(card2, fg_color="transparent")
        hero_grid.pack(fill="x", padx=20, pady=(0, 20))
        hero_grid.grid_columnconfigure(0, weight=1)
        hero_grid.grid_columnconfigure(1, weight=1)

        # Hero Action 1: Neon Green Button
        h1_box = ctk.CTkFrame(hero_grid, fg_color="#120F24", corner_radius=12, border_color="#272248", border_width=1)
        h1_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        ctk.CTkButton(
            h1_box,
            text="🚀 월보고 전체 4단계 자동 작성",
            font=ctk.CTkFont(family="맑은 고딕", size=15, weight="bold"),
            fg_color="#10B981",
            hover_color="#059669",
            text_color="#FFFFFF",
            height=48,
            corner_radius=10,
            command=self.start_run_monthly_fill
        ).pack(fill="x", padx=12, pady=(12, 8))

        ctk.CTkLabel(
            h1_box,
            text="• 세입세출표 ➔ 회원현황 ➔ 월별 실적 ➔ 직원가입실적 4개 시트 연속 라이브 기입",
            font=ctk.CTkFont(family="맑은 고딕", size=10),
            text_color="#94A3B8"
        ).pack(anchor="w", padx=14, pady=(0, 12))

        # Hero Action 2: Neon Purple Button
        h2_box = ctk.CTkFrame(hero_grid, fg_color="#120F24", corner_radius=12, border_color="#272248", border_width=1)
        h2_box.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        ctk.CTkButton(
            h2_box,
            text="📨 통합 발송대장 5건 자동 등록",
            font=ctk.CTkFont(family="맑은 고딕", size=15, weight="bold"),
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            text_color="#FFFFFF",
            height=48,
            corner_radius=10,
            command=lambda: self.start_register_ship_documents(only_first_doc=False)
        ).pack(fill="x", padx=12, pady=(12, 8))

        ctk.CTkLabel(
            h2_box,
            text="• 관리자 > 문서대장관리 > 발송대장관리 5개 문서를 자동 추가 및 기입",
            font=ctk.CTkFont(family="맑은 고딕", size=10),
            text_color="#94A3B8"
        ).pack(anchor="w", padx=14, pady=(0, 12))

        # Card 3: 단계별 수동 제어 모음
        card3 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=16)
        card3.pack(fill="x", pady=(0, 14))

        card3_hdr = ctk.CTkFrame(card3, fg_color="transparent")
        card3_hdr.pack(fill="x", padx=20, pady=14)

        ctk.CTkLabel(
            card3_hdr,
            text="⚙️ [3] 단계별 개별 시트 수동 제어 모음",
            font=ctk.CTkFont(family="맑은 고딕", size=14, weight="bold"),
            text_color="#CBD5E1"
        ).pack(side="left")

        self.step_toggle_var = tk.BooleanVar(value=False)

        def toggle_step_section():
            if self.step_toggle_var.get():
                step_content.pack(fill="x", padx=20, pady=(0, 16))
                toggle_btn.configure(text="▲ 수동 메뉴 닫기")
            else:
                step_content.pack_forget()
                toggle_btn.configure(text="▶ 수동 메뉴 열기")

        toggle_btn = ctk.CTkButton(
            card3_hdr,
            text="▶ 수동 메뉴 열기",
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            fg_color="#374151",
            hover_color="#4B5563",
            width=110,
            height=30,
            corner_radius=8,
            command=lambda: [
                self.step_toggle_var.set(not self.step_toggle_var.get()),
                toggle_step_section()
            ]
        )
        toggle_btn.pack(side="right")

        step_content = ctk.CTkFrame(card3, fg_color="transparent")
        step_content.grid_columnconfigure(0, weight=1)
        step_content.grid_columnconfigure(1, weight=1)

        # 1단계
        s1 = ctk.CTkFrame(step_content, fg_color="#120F24", corner_radius=10)
        s1.grid(row=0, column=0, sticky="ew", padx=(0, 4), pady=4)
        ctk.CTkLabel(s1, text="[1단계] 세입세출표", font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold"), text_color="#A855F7").pack(anchor="w", padx=10, pady=(8, 4))
        b1_f = ctk.CTkFrame(s1, fg_color="transparent")
        b1_f.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkButton(b1_f, text="📊 1단계 작성", font=ctk.CTkFont(family="맑은 고딕", size=11), fg_color="#8B5CF6", width=100, command=self.start_run_step1_only).pack(side="left", padx=(0, 4))
        ctk.CTkButton(b1_f, text="🔍 데이터 추출", font=ctk.CTkFont(family="맑은 고딕", size=11), fg_color="#374151", width=100, command=self.start_extract_settlement).pack(side="left")

        # 2단계
        s2 = ctk.CTkFrame(step_content, fg_color="#120F24", corner_radius=10)
        s2.grid(row=0, column=1, sticky="ew", padx=(4, 0), pady=4)
        ctk.CTkLabel(s2, text="[2단계] 회원현황", font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold"), text_color="#A855F7").pack(anchor="w", padx=10, pady=(8, 4))
        ctk.CTkButton(s2, text="👥 2단계 작성", font=ctk.CTkFont(family="맑은 고딕", size=11), fg_color="#8B5CF6", width=120, command=self.start_run_member_fill).pack(anchor="w", padx=10, pady=(0, 8))

        # 3단계
        s3 = ctk.CTkFrame(step_content, fg_color="#120F24", corner_radius=10)
        s3.grid(row=1, column=0, sticky="ew", padx=(0, 4), pady=4)
        ctk.CTkLabel(s3, text="[3단계] 월별 회원현황 및 세입실적", font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold"), text_color="#A855F7").pack(anchor="w", padx=10, pady=(8, 4))
        b3_f = ctk.CTkFrame(s3, fg_color="transparent")
        b3_f.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkButton(b3_f, text="📈 3단계 작성", font=ctk.CTkFont(family="맑은 고딕", size=11), fg_color="#8B5CF6", width=100, command=self.start_run_sheet4_fill).pack(side="left", padx=(0, 4))
        ctk.CTkButton(b3_f, text="🔍 미리보기", font=ctk.CTkFont(family="맑은 고딕", size=11), fg_color="#374151", width=100, command=self.start_preview_sheet4).pack(side="left")

        # 4단계
        s4 = ctk.CTkFrame(step_content, fg_color="#120F24", corner_radius=10)
        s4.grid(row=1, column=1, sticky="ew", padx=(4, 0), pady=4)
        ctk.CTkLabel(s4, text="[4단계] 직원회원가입실적", font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold"), text_color="#A855F7").pack(anchor="w", padx=10, pady=(8, 4))
        ctk.CTkButton(s4, text="📋 4단계 작성", font=ctk.CTkFont(family="맑은 고딕", size=11), fg_color="#8B5CF6", width=120, command=self.start_run_staff_fill).pack(anchor="w", padx=10, pady=(0, 8))

    # ── [탭 1] 전표 자동등록 뷰 ──────────────────────────────────────────

    def _build_slip_tab(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(2, weight=1)

        # 1. 파일선택 및 날짜 설정 카드
        card1 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=14)
        card1.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        # Row 1: File
        f_row = ctk.CTkFrame(card1, fg_color="transparent")
        f_row.pack(fill="x", padx=16, pady=(12, 6))

        ctk.CTkLabel(f_row, text="1. 거래내역 파일:", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#E2E8F0").pack(side="left", padx=(0, 8))
        ctk.CTkEntry(f_row, textvariable=self.file_var, font=ctk.CTkFont(family="맑은 고딕", size=12), fg_color="#120F24", border_color="#3B326B", corner_radius=8).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(f_row, text="파일 선택", width=90, fg_color="#374151", hover_color="#4B5563", font=ctk.CTkFont(family="맑은 고딕", size=12), corner_radius=8, command=self.browse_file).pack(side="left", padx=(0, 4))
        ctk.CTkButton(f_row, text="불러오기", width=90, fg_color="#8B5CF6", hover_color="#7C3AED", font=ctk.CTkFont(family="맑은 고딕", size=12), corner_radius=8, command=self.load_excel).pack(side="left")

        # Row 2: Dates
        d_row = ctk.CTkFrame(card1, fg_color="transparent")
        d_row.pack(fill="x", padx=16, pady=(0, 12))

        today = date.today()
        first_day = today.replace(day=1).strftime("%Y-%m-%d")
        last_day = today.replace(day=calendar.monthrange(today.year, today.month)[1]).strftime("%Y-%m-%d")
        self.from_date_var = tk.StringVar(value=first_day)
        self.to_date_var = tk.StringVar(value=last_day)

        ctk.CTkLabel(d_row, text="📅 전표일자 범위:", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#CBD5E1").pack(side="left", padx=(0, 8))
        ctk.CTkEntry(d_row, textvariable=self.from_date_var, width=110, font=ctk.CTkFont(family="맑은 고딕", size=12), fg_color="#120F24", border_color="#3B326B", corner_radius=8).pack(side="left")
        ctk.CTkLabel(d_row, text="~", text_color="#94A3B8").pack(side="left", padx=4)
        ctk.CTkEntry(d_row, textvariable=self.to_date_var, width=110, font=ctk.CTkFont(family="맑은 고딕", size=12), fg_color="#120F24", border_color="#3B326B", corner_radius=8).pack(side="left")

        # 2. 거래내역 표 카드
        card2 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=14)
        card2.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        card2.grid_rowconfigure(0, weight=1)
        card2.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#120F24", foreground="#E2E8F0", fieldbackground="#120F24", rowheight=28, font=("맑은 고딕", 9))
        style.configure("Treeview.Heading", background="#1E1B3A", foreground="#A855F7", font=("맑은 고딕", 9, "bold"))
        style.map("Treeview", background=[("selected", "#3B326B")], foreground=[("selected", "#FFFFFF")])

        cols = ("date", "type", "dir", "amount", "content", "acct", "status")
        self.tree = ttk.Treeview(card2, columns=cols, show="headings", height=8)
        col_cfg = {
            "date": ("거래일시", 130, "center"),
            "type": ("전표유형", 75, "center"),
            "dir": ("구분", 55, "center"),
            "amount": ("금액", 110, "e"),
            "content": ("내용", 220, "w"),
            "acct": ("계정코드", 160, "w"),
            "status": ("상태", 75, "center"),
        }
        for col, (heading, width, anchor) in col_cfg.items():
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, anchor=anchor)

        sy = ttk.Scrollbar(card2, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sy.set)
        sy.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Delete>", self.delete_selected)
        self.tree.bind("<Button-3>", self.on_right_click)

        self.tree.tag_configure("cms", background="#1E1B3A")
        self.tree.tag_configure("salary", background="#162E25")
        self.tree.tag_configure("done", background="#163A2A")
        self.tree.tag_configure("error", background="#3E1A1A")
        self.tree.tag_configure("warn", background="#3E381A")

        self.ctx_menu = tk.Menu(self, tearoff=0, bg="#1E1B3A", fg="#E2E8F0")
        self.ctx_menu.add_command(label="설정 편집", command=self._ctx_edit)
        self.ctx_menu.add_command(label="이 위치부터 동일구분 연속 등록", command=self.start_batch_registration)
        self.ctx_menu.add_command(label="이 위치부터 전체 거래 연속 등록", command=self.start_all_batch_registration)
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="선택 거래 삭제", command=self.delete_selected)

        # 3. 실행 제어 패널
        card3 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=14)
        card3.grid(row=3, column=0, sticky="ew")

        c3_row = ctk.CTkFrame(card3, fg_color="transparent")
        c3_row.pack(fill="x", padx=16, pady=10)

        ctk.CTkLabel(c3_row, textvariable=self.run_status_var, font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#10B981").pack(side="left")

        ctk.CTkButton(c3_row, text="선택 삭제", width=90, fg_color="#EF4444", hover_color="#DC2626", font=ctk.CTkFont(family="맑은 고딕", size=12), corner_radius=8, command=self.delete_selected).pack(side="right", padx=(6, 0))

        self.all_batch_run_button = ctk.CTkButton(
            c3_row, text="🚀 전체 거래 연속 등록", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), fg_color="#10B981", hover_color="#059669", corner_radius=8, command=self.start_all_batch_registration
        )
        self.all_batch_run_button.pack(side="right", padx=(0, 6))

        self.batch_run_button = ctk.CTkButton(
            c3_row, text="⚡ 동일구분 연속 등록", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), fg_color="#8B5CF6", hover_color="#7C3AED", corner_radius=8, command=self.start_batch_registration
        )
        self.batch_run_button.pack(side="right", padx=(0, 6))

        self.run_button = ctk.CTkButton(
            c3_row, text="선택 1건 등록", font=ctk.CTkFont(family="맑은 고딕", size=12), fg_color="#374151", hover_color="#4B5563", corner_radius=8, command=self.start_registration
        )
        self.run_button.pack(side="right", padx=(0, 6))

    # ── 하단 콘솔 드로어 ──────────────────────────────────────────────────

    def _build_console_drawer(self) -> None:
        self.log_drawer = ctk.CTkFrame(self.main_container, fg_color="#120F24", border_color="#2E2756", border_width=1, corner_radius=12)
        self.log_drawer.grid(row=2, column=0, padx=24, pady=(0, 16), sticky="ew")

        hdr = ctk.CTkFrame(self.log_drawer, fg_color="transparent")
        hdr.pack(fill="x", padx=12, pady=6)

        ctk.CTkLabel(hdr, text="💻 실시간 작업 콘솔 로그", font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold"), text_color="#10B981").pack(side="left")

        ctk.CTkButton(hdr, text="지우기", width=60, height=22, font=ctk.CTkFont(family="맑은 고딕", size=10), fg_color="#374151", hover_color="#4B5563", command=self._clear_log).pack(side="right")

        self.log_textbox = ctk.CTkTextbox(
            self.log_drawer,
            height=90,
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color="#0D0B18",
            text_color="#10B981",
            corner_radius=8
        )
        self.log_textbox.pack(fill="both", expand=True, padx=12, pady=(0, 10))

    def _clear_log(self) -> None:
        self.log_textbox.delete("1.0", "end")

    def log(self, msg: str) -> None:
        def append():
            t_str = datetime.datetime.now().strftime("[%H:%M:%S] ")
            self.log_textbox.insert("end", t_str + msg + "\n")
            self.log_textbox.see("end")
        self.after(0, append)

    def thread_log(self, msg: str) -> None:
        self.log(msg)

    def toggle_log(self) -> None:
        pass

    # ── 이벤트 및 로직 구현 ──────────────────────────────────────────────────

    def check_browser_connection(self) -> None:
        def worker():
            ok = cdp_is_ready()
            st = "🟢 Chrome 연결됨 (Port 9222)" if ok else "🔴 Chrome 미연결"
            self.after(0, lambda: self.browser_status_var.set(st))
        threading.Thread(target=worker, daemon=True).start()

    def start_attachable_chrome(self) -> None:
        if cdp_is_ready():
            self.browser_status_var.set("🟢 Chrome 연결됨 (Port 9222)")
            self.log("이미 Chrome이 실행 중이며 9222 포트로 연결되어 있습니다.")
            return

        chrome = find_chrome()
        if not chrome:
            messagebox.showerror("Chrome 오류", "Chrome 실행 파일(chrome.exe)을 찾지 못했습니다.")
            return

        profile = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "nmis-cdp-profile"
        try:
            subprocess.Popen([
                str(chrome),
                "--remote-debugging-port=9222",
                f"--user-data-dir={profile}",
                "http://nmis.foodservice.or.kr/"
            ], close_fds=True)
            self.browser_status_var.set("🟡 Chrome 시작 중...")
            self.log("🌐 연결용 Chrome 브라우저 오픈 실행 중... 자동 로그인을 진행합니다.")
        except Exception as e:
            self.log(f"❌ Chrome 실행 실패: {e}")
            messagebox.showerror("Chrome 오류", f"Chrome 실행 실패: {e}")
            return

        def auto_login_worker():
            import time
            time.sleep(3)  # Chrome 로딩 대기

            try:
                with sync_playwright() as pw:
                    browser = pw.chromium.connect_over_cdp(CDP_URL)
                    page = find_nmis_page(browser)
                    if not page:
                        self.log("연결된 Chrome에서 NMIS 페이지를 찾지 못했습니다.")
                        self.after(0, lambda: self.browser_status_var.set("🟢 Chrome 연결됨 (Port 9222)"))
                        return

                    time.sleep(1)

                    # 아이디/비밀번호 자동 입력
                    user_id_loc = page.locator("#userId")
                    if user_id_loc.count() > 0 and user_id_loc.is_visible():
                        user_id_loc.fill(NMIS_USER_ID)
                        self.log(f"아이디 입력 완료: {NMIS_USER_ID}")

                        pw_loc = page.locator("#userPass")
                        if pw_loc.count() > 0 and pw_loc.is_visible():
                            pw_loc.fill(NMIS_PASSWORD)
                            self.log("비밀번호 입력 완료")
                            pw_loc.press("Enter")
                            self.log("로그인 요청 전송 완료")
                            time.sleep(2)

                    # 비밀번호 갱신 팝업 닫기 처리
                    cancel_btn = page.locator("button[ng-click*='fnCancel']")
                    deadline = time.monotonic() + 4.0
                    while time.monotonic() < deadline:
                        if cancel_btn.count() > 0 and cancel_btn.first.is_visible():
                            cancel_btn.first.click(force=True)
                            self.log("비밀번호 갱신 팝업 취소 처리 완료")
                            break
                        time.sleep(0.3)

                    self.log("🎉 Chrome 연결 및 NMIS 로그인 완결!")
                    self.after(0, lambda: self.browser_status_var.set("🟢 Chrome 연결됨 (Port 9222)"))
            except Exception as e:
                self.log(f"ℹ️ Chrome 연결 상태: {e}")
                self.after(0, lambda: self.browser_status_var.set("🟢 Chrome 연결됨 (Port 9222)"))

        threading.Thread(target=auto_login_worker, daemon=True).start()

    def browse_file(self) -> None:
        p = filedialog.askopenfilename(
            title="거래내역 엑셀 파일 선택",
            filetypes=(("Excel 파일", "*.xls;*.xlsx"), ("모든 파일", "*.*")),
        )
        if p:
            self.file_var.set(p)
            self.load_excel()

    def load_excel(self) -> None:
        path_str = self.file_var.get().strip()
        if not path_str:
            return
        p = Path(path_str)
        if not p.is_file():
            messagebox.showerror("오류", f"파일이 존재하지 않습니다:\n{p}")
            return
        try:
            txs = parse_excel(p)
        except Exception as e:
            messagebox.showerror("엑셀 파싱 오류", f"엑셀을 읽는 데 실패했습니다:\n{e}")
            return

        self.all_transactions = {tx.raw_id: tx for tx in txs}
        self.tx_type.clear()
        self.tx_settings.clear()

        for tx in txs:
            t_type = "기타"
            for kw, mapped_type in KEYWORD_RULES.items():
                if kw in tx.content:
                    t_type = mapped_type
                    break
            self.tx_type[tx.raw_id] = t_type

        self._refresh_tree()
        self.log(f"엑셀 파일 불러오기 완료: {p.name} (총 {len(txs)}건)")

    def _refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for raw_id, tx in self.all_transactions.items():
            t_type = self.tx_type.get(raw_id, "기타")
            t_type_korean = {
                "member_fee": "회비",
                "join_fee": "가입금",
                "salary": "급여/상여",
                "기타": "기타",
            }.get(t_type, t_type)

            acct_disp = "-"
            if t_type in ("member_fee", "join_fee"):
                acct_disp = "회비/가입금 자동"
            elif t_type == "salary":
                acct_disp = "인건비 자동"

            tag = "cms" if t_type in ("member_fee", "join_fee") else ("salary" if t_type == "salary" else "")

            self.tree.insert(
                "",
                "end",
                iid=raw_id,
                values=(
                    tx.date_str,
                    t_type_korean,
                    tx.direction,
                    f"{tx.amount:,}",
                    tx.content,
                    acct_disp,
                    "대기",
                ),
                tags=(tag,) if tag else (),
            )

    def on_double_click(self, event: tk.Event) -> None:
        item_id = self.tree.focus()
        if not item_id or item_id not in self.all_transactions:
            return
        self._open_edit_dialog(item_id)

    def on_right_click(self, event: tk.Event) -> None:
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.tree.focus(item)
            self.ctx_menu.tk_popup(event.x_root, event.y_root)

    def _ctx_edit(self) -> None:
        item_id = self.tree.focus()
        if item_id and item_id in self.all_transactions:
            self._open_edit_dialog(item_id)

    def _open_edit_dialog(self, item_id: str) -> None:
        tx = self.all_transactions[item_id]
        curr_type = self.tx_type.get(item_id, "기타")

        dlg = ctk.CTkToplevel(self)
        dlg.title(f"거래 설정 — {tx.content}")
        dlg.geometry("400x300")
        dlg.grab_set()

        ctk.CTkLabel(dlg, text=f"📌 {tx.date_str} | {tx.direction} | {tx.amount:,}원", font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold")).pack(pady=12)
        ctk.CTkLabel(dlg, text=f"내용: {tx.content}", font=ctk.CTkFont(family="맑은 고딕", size=11)).pack(pady=(0, 12))

        type_var = tk.StringVar(value=curr_type)

        f_type = ctk.CTkFrame(dlg, fg_color="transparent")
        f_type.pack(fill="x", padx=20, pady=8)

        ctk.CTkLabel(f_type, text="전표 유형:", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold")).pack(side="left", padx=(0, 10))

        combo = ctk.CTkComboBox(f_type, values=["member_fee", "join_fee", "salary", "기타"], variable=type_var)
        combo.pack(side="left", fill="x", expand=True)

        def save_and_close():
            self.tx_type[item_id] = type_var.get()
            self._refresh_tree()
            dlg.destroy()

        ctk.CTkButton(dlg, text="저장 및 적용", fg_color="#10B981", hover_color="#059669", command=save_and_close).pack(pady=20)

    def delete_selected(self, event: tk.Event = None) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        for iid in selected:
            if iid in self.all_transactions:
                del self.all_transactions[iid]
            if iid in self.tx_type:
                del self.tx_type[iid]
            self.tree.delete(iid)
        self.log(f"선택한 {len(selected)}건 거래 삭제 완료.")

    # ── 전표 실행 핸들러 ────────────────────────────────────────────────────

    def start_registration(self) -> None:
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("선택 없음", "등록할 거래를 하나 이상 선택하세요.")
            return
        target_txs = [self.all_transactions[iid] for iid in selected if iid in self.all_transactions]
        self._run_batch_worker(target_txs, mode="selected")

    def start_batch_registration(self) -> None:
        selected = self.tree.selection()
        all_children = list(self.tree.get_children())
        if not all_children:
            return
        start_idx = all_children.index(selected[0]) if selected else 0
        first_tx = self.all_transactions.get(all_children[start_idx])
        if not first_tx:
            return
        target_dir = first_tx.direction

        target_txs = []
        for iid in all_children[start_idx:]:
            tx = self.all_transactions.get(iid)
            if tx and tx.direction == target_dir:
                target_txs.append(tx)
            else:
                break
        self._run_batch_worker(target_txs, mode=f"same_dir_{target_dir}")

    def start_all_batch_registration(self) -> None:
        selected = self.tree.selection()
        all_children = list(self.tree.get_children())
        if not all_children:
            return
        start_idx = all_children.index(selected[0]) if selected else 0
        target_txs = [self.all_transactions[iid] for iid in all_children[start_idx:] if iid in self.all_transactions]
        self._run_batch_worker(target_txs, mode="all_continuous")

    def _run_batch_worker(self, target_txs: list[Transaction], mode: str) -> None:
        if not target_txs:
            return
        if not cdp_is_ready():
            messagebox.showwarning("Chrome 미연결", "먼저 [Chrome 연결 열기] 버튼을 눌러 Chrome을 연결하세요.")
            return

        def worker():
            self.running = True
            self.run_status_var.set("⚡ 전표 등록 진행 중...")
            try:
                with sync_playwright() as pw:
                    browser = pw.chromium.connect_over_cdp(CDP_URL)
                    page = find_nmis_page(browser)
                    if not page:
                        self.log("페이지를 찾지 못했습니다.")
                        return

                    for tx in target_txs:
                        t_type = self.tx_type.get(tx.raw_id, "기타")
                        res = register_transactions(page, [tx], {tx.raw_id: t_type}, self.log)
                        if res.get("success", False):
                            self.after(0, lambda i=tx.raw_id: self.tree.item(i, values=(
                                tx.date_str, t_type, tx.direction, f"{tx.amount:,}", tx.content, "자동", "✅ 완료"
                            ), tags=("done",)))
            except Exception as e:
                self.log(f"오류 발생: {e}")
            finally:
                self.running = False
                self.run_status_var.set("대기")

        threading.Thread(target=worker, daemon=True).start()

    # ── 월보고 실행 핸들러들 ─────────────────────────────────────────────────

    def start_run_monthly_fill(self) -> None:
        """1~4단계 연속 자동 작성"""
        if not cdp_is_ready():
            messagebox.showwarning("Chrome 미연결", "먼저 [Chrome 연결 열기] 버튼을 누르세요.")
            return
        excel_path = Path(self.monthly_excel_var.get()).expanduser()
        if not excel_path.is_file():
            messagebox.showerror("파일 오류", "월보고 엑셀 파일 경로를 확인해주세요.")
            return

        target_ym = self.member_ym_var.get().strip()

        def worker():
            self.log(f"🚀 [월보고 4단계 연속 자동 작성] 기준년월({target_ym}) 라이브 기입 시작...")
            try:
                with sync_playwright() as pw:
                    browser = pw.chromium.connect_over_cdp(CDP_URL)
                    page = find_nmis_page(browser)
                    if not page:
                        self.log("통합 웹페이지를 찾을 수 없습니다.")
                        return

                    fill_all_monthly_reports_sequentially(page, excel_path=excel_path, target_year_month=target_ym, log_cb=self.log)

                    msg = f"🎉 [월보고 4단계 전체 완성!] {target_ym} 기준 세입세출/회원/실적/직원가입 시트 기입 완결!"
                    self.log(msg)
                    self.after(0, lambda: messagebox.showinfo("월보고 자동작성 완료", msg))
            except Exception as e:
                self.log(f"❌ 월보고 자동 입력 오류: {e}")
                self.after(0, lambda: messagebox.showerror("오류", f"월보고 작성 중 오류: {e}"))

        threading.Thread(target=worker, daemon=True).start()

    def start_run_step1_only(self) -> None:
        if not cdp_is_ready():
            messagebox.showwarning("Chrome 미연결", "먼저 [Chrome 연결 열기] 버튼을 누르세요.")
            return
        excel_path = Path(self.monthly_excel_var.get()).expanduser()
        target_ym = self.member_ym_var.get().strip()
        def worker():
            with sync_playwright() as pw:
                page = find_nmis_page(pw.chromium.connect_over_cdp(CDP_URL))
                fill_monthly_report_from_nmis(page, excel_path=excel_path, target_year_month=target_ym, log_cb=self.log)
        threading.Thread(target=worker, daemon=True).start()

    def start_extract_settlement(self) -> None:
        if not cdp_is_ready():
            messagebox.showwarning("Chrome 미연결", "먼저 [Chrome 연결 열기] 버튼을 누르세요.")
            return
        target_ym = self.member_ym_var.get().strip()
        def worker():
            with sync_playwright() as pw:
                page = find_nmis_page(pw.chromium.connect_over_cdp(CDP_URL))
                fetch_sheet4_data_only(page, target_year_month=target_ym, log_cb=self.log)
        threading.Thread(target=worker, daemon=True).start()

    def start_run_member_fill(self) -> None:
        if not cdp_is_ready():
            messagebox.showwarning("Chrome 미연결", "먼저 [Chrome 연결 열기] 버튼을 누르세요.")
            return
        excel_path = Path(self.monthly_excel_var.get()).expanduser()
        target_ym = self.member_ym_var.get().strip()
        def worker():
            with sync_playwright() as pw:
                page = find_nmis_page(pw.chromium.connect_over_cdp(CDP_URL))
                fill_member_status_from_nmis(page, excel_path=excel_path, target_year_month=target_ym, log_cb=self.log)
        threading.Thread(target=worker, daemon=True).start()

    def start_run_sheet4_fill(self) -> None:
        if not cdp_is_ready():
            messagebox.showwarning("Chrome 미연결", "먼저 [Chrome 연결 열기] 버튼을 누르세요.")
            return
        excel_path = Path(self.monthly_excel_var.get()).expanduser()
        target_ym = self.member_ym_var.get().strip()
        def worker():
            with sync_playwright() as pw:
                page = find_nmis_page(pw.chromium.connect_over_cdp(CDP_URL))
                fill_monthly_member_and_revenue_report(page, excel_path=excel_path, target_year_month=target_ym, log_cb=self.log)
        threading.Thread(target=worker, daemon=True).start()

    def start_preview_sheet4(self) -> None:
        if not cdp_is_ready():
            messagebox.showwarning("Chrome 미연결", "먼저 [Chrome 연결 열기] 버튼을 누르세요.")
            return
        target_ym = self.member_ym_var.get().strip()
        def worker():
            with sync_playwright() as pw:
                page = find_nmis_page(pw.chromium.connect_over_cdp(CDP_URL))
                fetch_sheet4_data_only(page, target_year_month=target_ym, log_cb=self.log)
        threading.Thread(target=worker, daemon=True).start()

    def start_run_staff_fill(self) -> None:
        if not cdp_is_ready():
            messagebox.showwarning("Chrome 미연결", "먼저 [Chrome 연결 열기] 버튼을 누르세요.")
            return
        excel_path = Path(self.monthly_excel_var.get()).expanduser()
        target_ym = self.member_ym_var.get().strip()
        def worker():
            with sync_playwright() as pw:
                page = find_nmis_page(pw.chromium.connect_over_cdp(CDP_URL))
                data = fetch_sheet4_data_only(page, target_year_month=target_ym, log_cb=self.log)
                fill_staff_join_excel_from_data(excel_path=excel_path, data=data, log_cb=self.log)
        threading.Thread(target=worker, daemon=True).start()

    def start_register_ship_documents(self, only_first_doc: bool = False) -> None:
        if not cdp_is_ready():
            messagebox.showwarning("Chrome 미연결", "먼저 [Chrome 연결 열기] 버튼을 누르세요.")
            return
        send_date = self.ship_send_date_var.get().strip()
        month_label = self.ship_report_month_var.get().strip()
        def worker():
            with sync_playwright() as pw:
                page = find_nmis_page(pw.chromium.connect_over_cdp(CDP_URL))
                register_ship_documents_on_nmis(page, send_date=send_date, report_month_label=month_label, only_first_doc=only_first_doc, log_cb=self.log)
        threading.Thread(target=worker, daemon=True).start()

    def open_settings(self) -> None:
        dlg = ctk.CTkToplevel(self)
        dlg.title("계정 및 시스템 설정")
        dlg.geometry("450x380")
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="🔑 통합 로그인 계정 설정", font=ctk.CTkFont(family="맑은 고딕", size=14, weight="bold"), text_color="#A855F7").pack(pady=16)

        f_id = ctk.CTkFrame(dlg, fg_color="transparent")
        f_id.pack(fill="x", padx=30, pady=8)
        ctk.CTkLabel(f_id, text="아이디 (ID):", font=ctk.CTkFont(family="맑은 고딕", size=12)).pack(side="left", padx=(0, 10))
        e_id = ctk.CTkEntry(f_id, width=200)
        e_id.pack(side="left")
        e_id.insert(0, NMIS_USER_ID)

        f_pw = ctk.CTkFrame(dlg, fg_color="transparent")
        f_pw.pack(fill="x", padx=30, pady=8)
        ctk.CTkLabel(f_pw, text="비밀번호 (PW):", font=ctk.CTkFont(family="맑은 고딕", size=12)).pack(side="left", padx=(0, 10))
        e_pw = ctk.CTkEntry(f_pw, width=200, show="●")
        e_pw.pack(side="left")
        e_pw.insert(0, NMIS_PASSWORD)

        def save_and_close():
            global NMIS_USER_ID, NMIS_PASSWORD
            NMIS_USER_ID = e_id.get().strip()
            NMIS_PASSWORD = e_pw.get().strip()
            save_settings()
            messagebox.showinfo("저장 완료", "계정 설정이 성공적으로 저장되었습니다.")
            dlg.destroy()

        ctk.CTkButton(dlg, text="저장하기", fg_color="#10B981", hover_color="#059669", command=save_and_close).pack(pady=24)

    # ── [탭 3] 회원 정보 검수 뷰 ──────────────────────────────────────────

    def _build_member_tab(self, parent: ctk.CTkFrame) -> None:
        parent.grid_rowconfigure(2, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        self.member_results_list: list[MemberVerificationResult] = []
        self.member_stop_event = threading.Event()
        self.is_member_verifying = False

        # Card 1: 검수 기본 설정
        card1 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=16)
        card1.grid(row=0, column=0, sticky="ew", pady=(0, 12))

        ctk.CTkLabel(
            card1,
            text="👥 [1] 회원 데이터 검수 기본 설정 및 엑셀 파일 지정",
            font=ctk.CTkFont(family="맑은 고딕", size=15, weight="bold"),
            text_color="#A855F7"
        ).pack(anchor="w", padx=20, pady=(16, 12))

        # 엑셀 파일 선택 Row
        file_row = ctk.CTkFrame(card1, fg_color="transparent")
        file_row.pack(fill="x", padx=20, pady=(0, 10))

        ctk.CTkLabel(file_row, text="검수 대상 엑셀:", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#E2E8F0").pack(side="left", padx=(0, 8))

        default_dl_file = Path.home() / "Downloads" / "일반음식점현황(6.30.기준일).xlsx"
        default_member_path = str(default_dl_file) if default_dl_file.is_file() else ""
        self.member_excel_var = tk.StringVar(value=default_member_path)

        excel_entry = ctk.CTkEntry(
            file_row,
            textvariable=self.member_excel_var,
            font=ctk.CTkFont(family="맑은 고딕", size=12),
            fg_color="#120F24",
            border_color="#3B326B",
            corner_radius=8
        )
        excel_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        def browse_member_excel():
            p = filedialog.askopenfilename(
                title="일반음식점현황 엑셀 파일 선택",
                filetypes=(("Excel 파일", "*.xlsx;*.xls"), ("모든 파일", "*.*")),
            )
            if p:
                self.member_excel_var.set(p)

        ctk.CTkButton(
            file_row,
            text="📁 파일 선택",
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            fg_color="#374151",
            hover_color="#4B5563",
            width=90,
            height=32,
            command=browse_member_excel
        ).pack(side="left")

        # 검수 옵션 Row
        opt_row = ctk.CTkFrame(card1, fg_color="transparent")
        opt_row.pack(fill="x", padx=20, pady=(0, 14))

        self.check_license_var = tk.BooleanVar(value=False)
        chk_lic = ctk.CTkCheckBox(
            opt_row,
            text="인허가번호(신고번호) 일치 여부 추가 검수 (체크 시 엑셀 D열 vs NMIS 신고번호 동시 비교)",
            variable=self.check_license_var,
            font=ctk.CTkFont(family="맑은 고딕", size=12),
            text_color="#CBD5E1",
            fg_color="#8B5CF6",
            hover_color="#7C3AED"
        )
        chk_lic.pack(side="left")

        # 실행 버튼 그룹 Row
        btn_row = ctk.CTkFrame(card1, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(0, 16))

        self.btn_start_member = ctk.CTkButton(
            btn_row,
            text="🚀 회원 정보 검수 시작",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            height=38,
            corner_radius=10,
            command=self.start_member_verification
        )
        self.btn_start_member.pack(side="left", padx=(0, 10))

        self.btn_stop_member = ctk.CTkButton(
            btn_row,
            text="⏹ 중지",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
            fg_color="#EF4444",
            hover_color="#DC2626",
            height=38,
            corner_radius=10,
            state="disabled",
            command=self.stop_member_verification
        )
        self.btn_stop_member.pack(side="left", padx=(0, 10))

        self.btn_export_member = ctk.CTkButton(
            btn_row,
            text="📥 검수 결과 엑셀 저장",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
            fg_color="#10B981",
            hover_color="#059669",
            height=38,
            corner_radius=10,
            command=self.export_member_results
        )
        self.btn_export_member.pack(side="left")

        # Card 2: 검수 진행 및 실시간 통계 카운트
        card2 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=16)
        card2.grid(row=1, column=0, sticky="ew", pady=(0, 12))

        stat_row = ctk.CTkFrame(card2, fg_color="transparent")
        stat_row.pack(fill="x", padx=20, pady=12)

        self.lbl_member_status = ctk.CTkLabel(
            stat_row,
            text="준비 완료 (검수 시작 버튼을 눌러주세요)",
            font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"),
            text_color="#94A3B8"
        )
        self.lbl_member_status.pack(side="left", padx=(0, 20))

        # 카운터 뱃지들
        self.lbl_stat_total = ctk.CTkLabel(stat_row, text="전체: 0건", font=ctk.CTkFont(family="맑은 고딕", size=12), text_color="#E2E8F0")
        self.lbl_stat_total.pack(side="left", padx=8)

        self.lbl_stat_match = ctk.CTkLabel(stat_row, text="✅ 일치: 0건", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#10B981")
        self.lbl_stat_match.pack(side="left", padx=8)

        self.lbl_stat_mismatch = ctk.CTkLabel(stat_row, text="❌ 불일치: 0건", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#EF4444")
        self.lbl_stat_mismatch.pack(side="left", padx=8)

        self.lbl_stat_notfound = ctk.CTkLabel(stat_row, text="🔍 미검색: 0건", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#F59E0B")
        self.lbl_stat_notfound.pack(side="left", padx=8)

        self.progress_member = ctk.CTkProgressBar(card2, fg_color="#120F24", progress_color="#8B5CF6", height=8)
        self.progress_member.pack(fill="x", padx=20, pady=(0, 12))
        self.progress_member.set(0.0)

        # Card 3: 실시간 결과 리스트업 테이블
        card3 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=16)
        card3.grid(row=2, column=0, sticky="nsew")
        card3.grid_rowconfigure(1, weight=1)
        card3.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            card3,
            text="📋 실시간 검수 결과 및 불일치/미검색 리스트",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
            text_color="#A855F7"
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(14, 8))

        table_frame = ctk.CTkFrame(card3, fg_color="#120F24")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 14))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        cols = ("seq", "store_name", "excel_owner", "web_owner", "excel_license", "web_license", "status", "reason")
        self.member_tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")

        self.member_tree.heading("seq", text="순번")
        self.member_tree.heading("store_name", text="업소명 (F열)")
        self.member_tree.heading("excel_owner", text="엑셀 대표자 (G열)")
        self.member_tree.heading("web_owner", text="웹 대표자")
        self.member_tree.heading("excel_license", text="엑셀 인허가 (D열)")
        self.member_tree.heading("web_license", text="웹 신고번호")
        self.member_tree.heading("status", text="상태")
        self.member_tree.heading("reason", text="검수 사유")

        self.member_tree.column("seq", width=50, anchor="center")
        self.member_tree.column("store_name", width=140, anchor="w")
        self.member_tree.column("excel_owner", width=100, anchor="center")
        self.member_tree.column("web_owner", width=100, anchor="center")
        self.member_tree.column("excel_license", width=110, anchor="center")
        self.member_tree.column("web_license", width=110, anchor="center")
        self.member_tree.column("status", width=75, anchor="center")
        self.member_tree.column("reason", width=220, anchor="w")

        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.member_tree.yview)
        self.member_tree.configure(yscrollcommand=scroll.set)

        self.member_tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

    # ── 회원 정보 검수 실행 핸들러 ──────────────────────────────────────────

    def start_member_verification(self) -> None:
        if self.is_member_verifying:
            return

        excel_p = self.member_excel_var.get().strip()
        if not excel_p or not os.path.isfile(excel_p):
            messagebox.showerror("파일 오류", "올바른 엑셀 파일 경로를 선택해 주세요.")
            return

        self.is_member_verifying = True
        self.member_stop_event.clear()
        self.member_results_list.clear()

        # 트리 초기화
        for item in self.member_tree.get_children():
            self.member_tree.delete(item)

        self.btn_start_member.configure(state="disabled")
        self.btn_stop_member.configure(state="normal")
        self.lbl_member_status.configure(text="🔍 NMIS 크롬 브라우저 연동 및 검수 시작...", text_color="#A855F7")
        self.progress_member.set(0.0)

        def status_cb(data: dict):
            t = data.get("type")
            if t == "log":
                self.log(data.get("message", ""))
            elif t == "item_processed":
                item: MemberVerificationResult = data["item"]
                self.member_results_list.append(item)

                current = data["current"]
                total = data["total"]
                pct = current / total if total > 0 else 0.0

                def update_ui():
                    self.progress_member.set(pct)
                    self.lbl_stat_total.configure(text=f"전체: {total}건")
                    self.lbl_stat_match.configure(text=f"✅ 일치: {data['match_count']}건")
                    self.lbl_stat_mismatch.configure(text=f"❌ 불일치: {data['mismatch_count']}건")
                    self.lbl_stat_notfound.configure(text=f"🔍 미검색: {data['not_found_count']}건")
                    self.lbl_member_status.configure(text=f"검수 진행 중... ({current}/{total}) [{item.store_name}]")

                    # 트리뷰 행 추가
                    row_id = self.member_tree.insert(
                        "",
                        "end",
                        values=(
                            item.seq,
                            item.store_name,
                            item.excel_owner,
                            item.web_owner,
                            item.excel_license,
                            item.web_license,
                            item.status,
                            item.reason,
                        ),
                    )
                    self.member_tree.see(row_id)

                self.after(0, update_ui)

        def worker():
            try:
                with sync_playwright() as pw:
                    browser = pw.chromium.connect_over_cdp(CDP_URL)
                    res = verify_member_info_from_nmis(
                        target=browser,
                        excel_path=excel_p,
                        check_license=self.check_license_var.get(),
                        status_callback=status_cb,
                        stop_event=self.member_stop_event,
                    )
                    def on_complete():
                        self.is_member_verifying = False
                        self.btn_start_member.configure(state="normal")
                        self.btn_stop_member.configure(state="disabled")
                        self.lbl_member_status.configure(
                            text=f"🎉 검수 완결! (일치: {res['match_count']}건, 불일치: {res['mismatch_count']}건, 미검색: {res['not_found_count']}건)",
                            text_color="#10B981"
                        )
                        messagebox.showinfo(
                            "검수 완료",
                            f"회원 데이터 검수가 완료되었습니다!\n\n• 전체 검수: {res['total']}건\n• ✅ 일치: {res['match_count']}건\n• ❌ 불일치: {res['mismatch_count']}건\n• 🔍 미검색: {res['not_found_count']}건"
                        )
                    self.after(0, on_complete)
            except Exception as e:
                err_msg = str(e)
                def on_error():
                    self.is_member_verifying = False
                    self.btn_start_member.configure(state="normal")
                    self.btn_stop_member.configure(state="disabled")
                    self.lbl_member_status.configure(text=f"❌ 오류 발생: {err_msg[:40]}", text_color="#EF4444")
                    messagebox.showerror("검수 오류", f"회원 검수 중 오류가 발생했습니다:\n{err_msg}")
                self.after(0, on_error)

        threading.Thread(target=worker, daemon=True).start()

    def stop_member_verification(self) -> None:
        if self.is_member_verifying:
            self.member_stop_event.set()
            self.lbl_member_status.configure(text="⏹ 사용자 요청으로 중지 중...", text_color="#EF4444")

    def export_member_results(self) -> None:
        if not self.member_results_list:
            messagebox.showwarning("저장 경고", "저장할 검수 결과 데이터가 없습니다. 먼저 검수를 진행해 주세요.")
            return

        out_path = filedialog.asksaveasfilename(
            title="검수 결과 엑셀 저장",
            defaultextension=".xlsx",
            filetypes=(("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")),
            initialfile=f"NMIS_회원검수결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not out_path:
            return

        try:
            import pandas as pd
            data = []
            for r in self.member_results_list:
                data.append({
                    "순번": r.seq,
                    "업소명(F열)": r.store_name,
                    "엑셀대표자(G열)": r.excel_owner,
                    "웹대표자": r.web_owner,
                    "엑셀인허가번호(D열)": r.excel_license,
                    "웹신고번호": r.web_license,
                    "검수상태": r.status,
                    "검수세부사유": r.reason
                })
            df = pd.DataFrame(data)
            df.to_excel(out_path, index=False)
            messagebox.showinfo("저장 완료", f"검수 결과가 성공적으로 저장되었습니다!\n\n경로: {out_path}")
        except Exception as e:
            messagebox.showerror("저장 실패", f"엑셀 저장 중 오류 발생:\n{e}")



def main() -> None:
    app = ModernSlipUI()
    app.mainloop()


if __name__ == "__main__":
    main()
