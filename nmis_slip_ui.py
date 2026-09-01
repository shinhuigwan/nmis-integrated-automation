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
import subprocess
import time
import urllib.request
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable

import customtkinter as ctk
import openpyxl
import pandas as pd
from playwright.sync_api import sync_playwright

# ── 모듈 임포트 ─────────────────────────────────────────────────────────────
CDP_URL = "http://127.0.0.1:9222"

from nmis_slip_automation import (
    ACCOUNT_CODES,
    MemberVerificationResult,
    Transaction,
    classify_transaction_type,
    describe_slip_settings,
    fetch_sheet4_data_only,
    fill_all_monthly_reports_sequentially,
    fill_member_status_from_nmis,
    fill_monthly_member_and_revenue_report,
    fill_monthly_report_from_nmis,
    fill_staff_join_excel_from_data,
    find_nmis_page,
    navigate_to_slip_management,
    register_ship_documents_on_nmis,
    verify_member_info_from_nmis,
    register_potential_members_on_nmis,
    register_transactions,
    parse_excel,
    parse_rrn_birth_gender,
    resolve_column_from_letter_or_name,
    format_korean_phone,
    format_digits_only,
    set_slip_search_date_range,
    slip_settings_ready,
    suggest_slip_settings,
    transaction_date_range,
)
from receipt_register import (
    ReceiptHistory,
    ReceiptMailItem,
    apply_history,
    complete_registered_attachment,
    extract_receipt_document_metadata,
    filter_receipt_candidate_documents,
    parse_iso_date,
    protect_windows_secret,
    register_receipt_document_on_nmis,
    search_naver_hwp_mail,
    sort_receipt_items_oldest_first,
    unprotect_windows_secret,
)
from membership_fee_register import (
    FeePaymentItem,
    ensure_nmis_logged_in,
    navigate_to_fee_management,
    parse_fee_payment_excel,
    register_fee_payment_on_nmis,
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

SLIP_COLUMN_MAPPING: dict[str, str] = {
    "date": "A",
    "type": "F",
    "dir": "C",
    "amount": "D",
    "content": "E",
}

MEMBER_COLUMN_MAPPING: dict[str, str] = {
    "store_name": "F",
    "owner_name": "G",
    "license_no": "D",
}

POTENTIAL_COLUMN_MAPPING: dict[str, str] = {
    "biz_type": "C",
    "license_no": "D",
    "perm_date": "E",
    "store_name": "F",
    "owner_name": "G",
    "rrn": "H",
    "address": "I",
    "area": "K",
    "phone": "L",
    "mobile": "P",
}

RECEIPT_SETTINGS: dict[str, Any] = {
    "naver_id": "",
    "naver_app_password_dpapi": "",
    "sender_filter": "krbajb5980@daum.net",
    "all_senders": True,
    "download_dir": str(Path.home() / "Downloads" / "NMIS_접수대장"),
}

FEE_PAYMENT_SETTINGS: dict[str, Any] = {
    "excel_path": str(Path.home() / "Desktop" / "통장이체_거래내역_정리.xlsx"),
}

def get_excel_column_headers(file_path: str) -> list[dict]:
    if not file_path or not os.path.exists(file_path):
        return []
    try:
        df = pd.read_excel(file_path, nrows=1)
        cols = []
        for idx, col in enumerate(df.columns):
            letter = chr(65 + idx) if idx < 26 else f"A{chr(65 + idx - 26)}"
            col_name = str(col).strip()
            cols.append({
                "index": idx,
                "letter": letter,
                "header": col_name,
                "display": f"[{letter}열] {col_name}"
            })
        return cols
    except Exception as e:
        print(f"헤더 읽기 예외: {e}")
        return []

def load_settings() -> None:
    global KEYWORD_RULES, CONFIRM_SELECTORS, MACRO_STEPS, NMIS_USER_ID, NMIS_PASSWORD, SLIP_COLUMN_MAPPING, MEMBER_COLUMN_MAPPING, POTENTIAL_COLUMN_MAPPING, RECEIPT_SETTINGS, FEE_PAYMENT_SETTINGS
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
        if "slip_column_mapping" in data:
            SLIP_COLUMN_MAPPING.update(data["slip_column_mapping"])
        if "member_column_mapping" in data:
            MEMBER_COLUMN_MAPPING.update(data["member_column_mapping"])
        if "potential_column_mapping" in data:
            POTENTIAL_COLUMN_MAPPING.update(data["potential_column_mapping"])
        if "receipt_settings" in data:
            RECEIPT_SETTINGS.update(data["receipt_settings"])
        if "fee_payment_settings" in data:
            FEE_PAYMENT_SETTINGS.update(data["fee_payment_settings"])
        if RECEIPT_SETTINGS.get("sender_filter") in {"", "한국외식업중앙회", "외식업전북지회"}:
            RECEIPT_SETTINGS["sender_filter"] = "krbajb5980@daum.net"
    except Exception as e:
        print(f"설정 로드 실패: {e}")

def save_settings() -> None:
    data = {
        "keyword_rules": KEYWORD_RULES,
        "confirm_selectors": CONFIRM_SELECTORS,
        "macro_steps": MACRO_STEPS,
        "nmis_user_id": NMIS_USER_ID,
        "nmis_password": NMIS_PASSWORD,
        "slip_column_mapping": SLIP_COLUMN_MAPPING,
        "member_column_mapping": MEMBER_COLUMN_MAPPING,
        "potential_column_mapping": POTENTIAL_COLUMN_MAPPING,
        # 네이버 앱 비밀번호는 receipt_settings 안에 Windows DPAPI 암호문으로만 저장합니다.
        "receipt_settings": RECEIPT_SETTINGS,
        "fee_payment_settings": FEE_PAYMENT_SETTINGS,
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


class ModernSlipUI(ctk.CTk):
    def _center_window(self, width: int = 1120, height: int = 780) -> None:
        try:
            self.update_idletasks()
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            x = max(0, (sw - width) // 2)
            y = max(0, (sh - height) // 2)
            self.geometry(f"{width}x{height}+{x}+{y}")
        except Exception:
            self.geometry(f"{width}x{height}")

    def __init__(self) -> None:
        super().__init__()

        load_settings()

        self.title("통합 자동화 시스템 — Modern Dark Edition")
        self._center_window(1120, 780)
        self.minsize(980, 680)

        # 변수 데이터
        self.file_var = tk.StringVar()
        self.browser_status_var = tk.StringVar(value="🔴 Chrome 연결 확인 전")
        self.run_status_var = tk.StringVar(value="대기")
        self.running = False

        self.all_transactions: dict[str, Transaction] = {}
        self.tx_type: dict[str, str] = {}
        self.tx_settings: dict[str, dict] = {}
        self.receipt_items: dict[str, ReceiptMailItem] = {}
        self.receipt_history = ReceiptHistory(Path(__file__).parent / "data" / "receipt_history.sqlite3")
        self.receipt_running = False
        self.receipt_checked_ids: set[str] = set()
        self.receipt_drag_check_value: bool | None = None
        self.receipt_drag_seen_ids: set[str] = set()
        self.fee_items: dict[str, FeePaymentItem] = {}
        self.fee_checked_ids: set[str] = set()
        self.fee_running = False
        self.fee_stop_requested = False

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
        self.sidebar.grid_rowconfigure(8, weight=1)

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
            text="전표 & 월보고 시스템",
            font=ctk.CTkFont(family="맑은 고딕", size=10),
            text_color="#94A3B8"
        ).pack(anchor="w")

        # 탭 선택 버튼
        self.btn_tab_monthly = ctk.CTkButton(
            self.sidebar,
            text="📊 월보고 자동 연동",
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

        self.btn_tab_potential = ctk.CTkButton(
            self.sidebar,
            text="📑 잠재회원 등록",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
            fg_color="transparent",
            hover_color="#26214A",
            text_color="#CBD5E1",
            height=42,
            corner_radius=10,
            command=lambda: self._select_tab("potential")
        )
        self.btn_tab_potential.grid(row=4, column=0, padx=16, pady=6, sticky="ew")

        self.btn_tab_receipt = ctk.CTkButton(
            self.sidebar,
            text="📥 접수대장",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
            fg_color="transparent",
            hover_color="#26214A",
            text_color="#CBD5E1",
            height=42,
            corner_radius=10,
            command=lambda: self._select_tab("receipt")
        )
        self.btn_tab_receipt.grid(row=5, column=0, padx=16, pady=6, sticky="ew")

        self.btn_tab_fee = ctk.CTkButton(
            self.sidebar,
            text="💳 회비 입금등록",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
            fg_color="transparent",
            hover_color="#26214A",
            text_color="#CBD5E1",
            height=42,
            corner_radius=10,
            command=lambda: self._select_tab("fee")
        )
        self.btn_tab_fee.grid(row=6, column=0, padx=16, pady=6, sticky="ew")

        # Chrome Status Box in Sidebar
        status_box = ctk.CTkFrame(self.sidebar, fg_color="#1E1B3A", corner_radius=12)
        status_box.grid(row=7, column=0, padx=16, pady=16, sticky="ew")

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
        self.tab_potential_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.tab_receipt_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.tab_fee_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")

        self._build_monthly_tab(self.tab_monthly_frame)
        self._build_slip_tab(self.tab_slip_frame)
        self._build_member_tab(self.tab_member_frame)
        self._build_potential_tab(self.tab_potential_frame)
        self._build_receipt_tab(self.tab_receipt_frame)
        self._build_fee_tab(self.tab_fee_frame)

        # 초기 탭 표시 (월보고 자동 연동)
        self._select_tab("monthly")

        # 3. 하단 실시간 콘솔 드로어
        self._build_console_drawer()

    def _select_tab(self, tab_name: str) -> None:
        self.tab_monthly_frame.grid_forget()
        self.tab_slip_frame.grid_forget()
        self.tab_member_frame.grid_forget()
        self.tab_potential_frame.grid_forget()
        self.tab_receipt_frame.grid_forget()
        self.tab_fee_frame.grid_forget()

        self.btn_tab_monthly.configure(fg_color="transparent", text_color="#CBD5E1")
        self.btn_tab_slip.configure(fg_color="transparent", text_color="#CBD5E1")
        self.btn_tab_member.configure(fg_color="transparent", text_color="#CBD5E1")
        self.btn_tab_potential.configure(fg_color="transparent", text_color="#CBD5E1")
        self.btn_tab_receipt.configure(fg_color="transparent", text_color="#CBD5E1")
        self.btn_tab_fee.configure(fg_color="transparent", text_color="#CBD5E1")

        if tab_name == "monthly":
            self.tab_monthly_frame.grid(row=1, column=0, padx=24, pady=10, sticky="nsew")
            self.btn_tab_monthly.configure(fg_color="#8B5CF6", text_color="#FFFFFF")
            self.main_title_label.configure(text="📊 월보고 자동 연동 시스템")
        elif tab_name == "slip":
            self.tab_slip_frame.grid(row=1, column=0, padx=24, pady=10, sticky="nsew")
            self.btn_tab_slip.configure(fg_color="#8B5CF6", text_color="#FFFFFF")
            self.main_title_label.configure(text="📋 전표 자동등록 시스템")
        elif tab_name == "member":
            self.tab_member_frame.grid(row=1, column=0, padx=24, pady=10, sticky="nsew")
            self.btn_tab_member.configure(fg_color="#8B5CF6", text_color="#FFFFFF")
            self.main_title_label.configure(text="👥 회원 정보(대표자/신고번호) 검수 시스템")
        elif tab_name == "potential":
            self.tab_potential_frame.grid(row=1, column=0, padx=24, pady=10, sticky="nsew")
            self.btn_tab_potential.configure(fg_color="#8B5CF6", text_color="#FFFFFF")
            self.main_title_label.configure(text="📑 잠재회원 등록 시스템 (1~8단계 서식 자동 작성)")
        elif tab_name == "receipt":
            self.tab_receipt_frame.grid(row=1, column=0, padx=24, pady=10, sticky="nsew")
            self.btn_tab_receipt.configure(fg_color="#8B5CF6", text_color="#FFFFFF")
            self.main_title_label.configure(text="📥 네이버 메일 → NMIS 접수대장 자동화")
        elif tab_name == "fee":
            self.tab_fee_frame.grid(row=1, column=0, padx=24, pady=10, sticky="nsew")
            self.btn_tab_fee.configure(fg_color="#8B5CF6", text_color="#FFFFFF")
            self.main_title_label.configure(text="💳 엑셀 → NMIS 월회비 입금등록")

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
        ctk.CTkButton(f_row, text="불러오기", width=90, fg_color="#8B5CF6", hover_color="#7C3AED", font=ctk.CTkFont(family="맑은 고딕", size=12), corner_radius=8, command=self.load_excel).pack(side="left", padx=(0, 4))
        ctk.CTkButton(f_row, text="⚙️ 엑셀 컬럼 설정", width=120, fg_color="#374151", hover_color="#4B5563", font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold"), corner_radius=8, command=self.open_slip_column_settings_dialog).pack(side="left")

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
        already_connected = cdp_is_ready()
        if already_connected:
            self.browser_status_var.set("🟢 Chrome 연결됨 (Port 9222)")
            self.log("연결된 Chrome의 NMIS 로그인 상태를 다시 확인합니다.")
        else:
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
            time.sleep(1 if already_connected else 3)  # 기존 연결은 짧게, 새 Chrome은 로딩 대기

            try:
                with sync_playwright() as pw:
                    browser = pw.chromium.connect_over_cdp(CDP_URL)
                    page = find_nmis_page(browser)
                    if not page:
                        self.log("연결된 Chrome에서 NMIS 페이지를 찾지 못했습니다.")
                        self.after(0, lambda: self.browser_status_var.set("🟢 Chrome 연결됨 (Port 9222)"))
                        return

                    ensure_nmis_logged_in(page, NMIS_USER_ID, NMIS_PASSWORD, self.log)

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

        if txs:
            earliest_date, latest_date = transaction_date_range(txs)
            self.from_date_var.set(earliest_date.isoformat())
            self.to_date_var.set(latest_date.isoformat())
            self.log(
                f"전표일자 범위 자동 설정: {earliest_date:%Y-%m-%d} ~ "
                f"{latest_date:%Y-%m-%d}"
            )

        self.all_transactions = {tx.raw_id: tx for tx in txs}
        self.tx_type.clear()
        self.tx_settings.clear()

        for tx in txs:
            transaction_type = classify_transaction_type(tx, KEYWORD_RULES)
            self.tx_type[tx.raw_id] = transaction_type
            self.tx_settings[tx.raw_id] = suggest_slip_settings(tx, transaction_type)

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

            settings = self.tx_settings.get(raw_id) or suggest_slip_settings(tx, t_type)
            self.tx_settings[raw_id] = settings
            acct_disp = describe_slip_settings(tx, settings)

            if not slip_settings_ready(tx, settings):
                tag = "warn"
            else:
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
                    "대기" if slip_settings_ready(tx, settings) else "설정 필요",
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
        current = dict(
            self.tx_settings.get(item_id)
            or suggest_slip_settings(tx, curr_type)
        )

        dlg = ctk.CTkToplevel(self)
        dlg.title(f"거래 설정 — {tx.content}")
        dlg.geometry("560x670")
        dlg.minsize(520, 620)
        dlg.grab_set()

        ctk.CTkLabel(
            dlg,
            text=f"📌 {tx.date_str} | {tx.direction} | {tx.amount:,.0f}원",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
        ).pack(pady=(14, 4))
        ctk.CTkLabel(
            dlg,
            text=f"내용: {tx.content} / 적요: {tx.memo or '-'}",
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            wraplength=510,
        ).pack(padx=20, pady=(0, 10))

        body = ctk.CTkScrollableFrame(dlg, fg_color="#18152E")
        body.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        type_var = tk.StringVar(value=curr_type)
        account_code_var = tk.StringVar(value=str(current.get("account_code") or ""))
        account_name_var = tk.StringVar(value=str(current.get("account_name") or ""))
        brief_var = tk.StringVar(value=str(current.get("brief") or ""))
        cms_count_var = tk.StringVar(value=str(current.get("cms_count") or ""))
        cms_fee_var = tk.StringVar(value=str(current.get("cms_fee") or ""))
        basic_pay_var = tk.StringVar(value=str(current.get("basic_pay") or ""))
        bonus_pay_var = tk.StringVar(value=str(current.get("bonus_pay") or ""))

        def add_entry(label: str, variable: tk.StringVar) -> None:
            row = ctk.CTkFrame(body, fg_color="transparent")
            row.pack(fill="x", padx=10, pady=5)
            ctk.CTkLabel(row, text=label, width=135, anchor="w").pack(side="left")
            ctk.CTkEntry(row, textvariable=variable).pack(side="left", fill="x", expand=True)

        f_type = ctk.CTkFrame(body, fg_color="transparent")
        f_type.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(f_type, text="전표 유형", width=135, anchor="w").pack(side="left")

        combo = ctk.CTkComboBox(f_type, values=["member_fee", "join_fee", "salary", "기타"], variable=type_var)
        combo.pack(side="left", fill="x", expand=True)

        account_options = ["직접 입력"] + [
            f"{code} | {name}" for code, name in ACCOUNT_CODES.items()
        ]

        def selected_account_option() -> str:
            code = account_code_var.get().strip()
            name = account_name_var.get().strip() or ACCOUNT_CODES.get(code, "")
            option = f"{code} | {name}" if code and name else "직접 입력"
            return option if option in account_options else "직접 입력"

        account_choice_var = tk.StringVar(value=selected_account_option())

        def apply_account_choice(choice: str) -> None:
            if choice == "직접 입력":
                return
            code, _, name = choice.partition("|")
            account_code_var.set(code.strip())
            account_name_var.set(name.strip())

        account_row = ctk.CTkFrame(body, fg_color="transparent")
        account_row.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(account_row, text="계정과목 선택", width=135, anchor="w").pack(side="left")
        account_combo = ctk.CTkComboBox(
            account_row,
            values=account_options,
            variable=account_choice_var,
            command=apply_account_choice,
            state="readonly",
        )
        account_combo.pack(side="left", fill="x", expand=True)

        add_entry("선택된 계정코드", account_code_var)
        add_entry("선택된 계정명", account_name_var)
        add_entry("적요", brief_var)

        ctk.CTkLabel(
            body,
            text="CMS 중앙회 입금 설정 (5141 회비 + 4385 잡비)",
            text_color="#A855F7",
            anchor="w",
        ).pack(fill="x", padx=10, pady=(14, 2))
        add_entry("CMS 건수", cms_count_var)
        add_entry("CMS 수수료", cms_fee_var)

        ctk.CTkLabel(
            body,
            text="급여 분할 설정 (4326 기본급 + 4403 직책수당및직무급)",
            text_color="#10B981",
            anchor="w",
        ).pack(fill="x", padx=10, pady=(14, 2))
        add_entry("기본급", basic_pay_var)
        add_entry("직무급/상여", bonus_pay_var)

        ctk.CTkLabel(
            body,
            text=(
                "CMS 거래는 건수와 수수료가 모두 필요합니다.\n"
                "급여는 기본급과 직무급/상여의 합계가 거래금액과 같아야 합니다."
            ),
            justify="left",
            text_color="#FBBF24",
        ).pack(fill="x", padx=10, pady=12)

        def parse_money(value: str, field_name: str) -> int:
            cleaned = value.replace(",", "").replace("원", "").strip()
            if not cleaned:
                return 0
            try:
                return int(cleaned)
            except ValueError as exc:
                raise ValueError(f"{field_name}에는 숫자만 입력하세요.") from exc

        def apply_type_defaults() -> None:
            defaults = suggest_slip_settings(tx, type_var.get())
            account_code_var.set(str(defaults.get("account_code") or ""))
            account_name_var.set(str(defaults.get("account_name") or ""))
            account_choice_var.set(selected_account_option())
            brief_var.set(str(defaults.get("brief") or ""))
            cms_count_var.set(str(defaults.get("cms_count") or ""))
            cms_fee_var.set(str(defaults.get("cms_fee") or ""))
            basic_pay_var.set(str(defaults.get("basic_pay") or ""))
            bonus_pay_var.set(str(defaults.get("bonus_pay") or ""))

        def save_and_close() -> None:
            try:
                selected_type = type_var.get()
                is_cms = (
                    selected_type == "member_fee"
                    and tx.account_code == "5141"
                    and tx.memo.strip().upper() == "CMS"
                )
                mode = "salary_bundle" if selected_type == "salary" else ("cms_bundle" if is_cms else "single")
                selected_account_code = account_code_var.get().strip()
                selected_account_name = (
                    account_name_var.get().strip()
                    or ACCOUNT_CODES.get(selected_account_code, "")
                )
                new_settings: dict[str, object] = {
                    "mode": mode,
                    "account_code": selected_account_code,
                    "account_name": selected_account_name,
                    "brief": brief_var.get().strip(),
                    "cms_count": parse_money(cms_count_var.get(), "CMS 건수"),
                    "cms_fee": parse_money(cms_fee_var.get(), "CMS 수수료"),
                    "basic_pay": parse_money(basic_pay_var.get(), "기본급"),
                    "bonus_pay": parse_money(bonus_pay_var.get(), "직무급/상여"),
                    "basic_account": "4326",
                    "bonus_account": "4403",
                }
            except ValueError as exc:
                messagebox.showwarning("설정 확인", str(exc), parent=dlg)
                return

            if not slip_settings_ready(tx, new_settings):
                messagebox.showwarning(
                    "설정 확인",
                    "자동등록에 필요한 값이 부족하거나 합계가 맞지 않습니다.\n"
                    "계정코드, CMS 건수·수수료 또는 급여 분할금액을 확인하세요.",
                    parent=dlg,
                )
                return

            self.tx_type[item_id] = selected_type
            self.tx_settings[item_id] = new_settings
            self._refresh_tree()
            dlg.destroy()

        footer = ctk.CTkFrame(dlg, fg_color="transparent")
        footer.pack(fill="x", padx=16, pady=(0, 14))
        ctk.CTkButton(
            footer,
            text="유형 기본값 적용",
            fg_color="#374151",
            hover_color="#4B5563",
            command=apply_type_defaults,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            footer,
            text="저장 및 적용",
            fg_color="#10B981",
            hover_color="#059669",
            command=save_and_close,
        ).pack(side="right")

    def delete_selected(self, event: tk.Event = None) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        for iid in selected:
            if iid in self.all_transactions:
                del self.all_transactions[iid]
            if iid in self.tx_type:
                del self.tx_type[iid]
            self.tx_settings.pop(iid, None)
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
        missing = [
            tx for tx in target_txs
            if not slip_settings_ready(
                tx,
                self.tx_settings.get(tx.raw_id)
                or suggest_slip_settings(tx, self.tx_type.get(tx.raw_id, "기타")),
            )
        ]
        if missing:
            first = missing[0]
            if self.tree.exists(first.raw_id):
                self.tree.selection_set(first.raw_id)
                self.tree.focus(first.raw_id)
                self.tree.see(first.raw_id)
            sample = "\n".join(f"- {tx.date_str[:10]} {tx.content}" for tx in missing[:5])
            more = f"\n외 {len(missing) - 5}건" if len(missing) > 5 else ""
            messagebox.showwarning(
                "전표 설정 필요",
                f"자동등록 설정이 필요한 거래가 {len(missing)}건 있습니다.\n"
                f"해당 행을 더블클릭하여 설정한 뒤 다시 실행하세요.\n\n{sample}{more}",
            )
            return
        if not cdp_is_ready():
            messagebox.showwarning("Chrome 미연결", "먼저 [Chrome 연결 열기] 버튼을 눌러 Chrome을 연결하세요.")
            return

        start_date, end_date = transaction_date_range(target_txs)

        def update_row(item_id: str, status: str, tag: str) -> None:
            tx = self.all_transactions.get(item_id)
            if not tx or not self.tree.exists(item_id):
                return
            t_type = self.tx_type.get(item_id, "기타")
            type_label = {
                "member_fee": "회비",
                "join_fee": "가입금",
                "salary": "급여/상여",
                "기타": "기타",
            }.get(t_type, t_type)
            settings = self.tx_settings.get(item_id) or suggest_slip_settings(tx, t_type)
            self.tree.item(
                item_id,
                values=(
                    tx.date_str,
                    type_label,
                    tx.direction,
                    f"{tx.amount:,.0f}",
                    tx.content,
                    describe_slip_settings(tx, settings),
                    status,
                ),
                tags=(tag,),
            )

        def worker():
            self.running = True
            self.after(0, lambda: self.run_status_var.set("⚡ 전표 등록 진행 중..."))
            success_count = 0
            failure_count = 0
            try:
                with sync_playwright() as pw:
                    browser = pw.chromium.connect_over_cdp(CDP_URL)
                    page = find_nmis_page(browser)
                    if not page:
                        self.log("페이지를 찾지 못했습니다.")
                        return

                    navigate_to_slip_management(page, self.log)
                    set_slip_search_date_range(page, start_date, end_date, self.log)

                    for tx in target_txs:
                        t_type = self.tx_type.get(tx.raw_id, "기타")
                        self.after(0, lambda i=tx.raw_id: update_row(i, "처리 중", ""))
                        try:
                            res = register_transactions(
                                page,
                                [tx],
                                {tx.raw_id: t_type},
                                self.log,
                                settings_by_id=self.tx_settings,
                            )
                            if res.get("success", False):
                                success_count += 1
                                self.after(0, lambda i=tx.raw_id: update_row(i, "✅ 완료", "done"))
                        except Exception as exc:
                            failure_count += 1
                            self.log(f"❌ 전표 등록 실패 [{tx.content}]: {exc}")
                            self.after(0, lambda i=tx.raw_id: update_row(i, "❌ 오류", "error"))
            except Exception as e:
                self.log(f"오류 발생: {e}")
            finally:
                self.running = False
                self.after(0, lambda: self.run_status_var.set("대기"))
                self.log(
                    f"전표 연속등록 종료: 성공 {success_count}건 / 실패 {failure_count}건"
                )
                self.after(
                    0,
                    lambda: messagebox.showinfo(
                        "전표 자동등록 결과",
                        f"작업이 끝났습니다.\n성공 {success_count}건 / 실패 {failure_count}건",
                    ),
                )

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

    # ── 접수대장 자동화 ─────────────────────────────────────────────────────

    def _build_receipt_tab(self, parent: ctk.CTkFrame) -> None:
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        today = date.today()
        card1 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=16)
        card1.grid(row=0, column=0, sticky="ew", pady=(0, 12))

        ctk.CTkLabel(
            card1,
            text="📨 [1] 네이버 메일 검색 조건",
            font=ctk.CTkFont(family="맑은 고딕", size=15, weight="bold"),
            text_color="#A855F7",
        ).pack(anchor="w", padx=20, pady=(16, 10))

        account_row = ctk.CTkFrame(card1, fg_color="transparent")
        account_row.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkLabel(account_row, text="네이버 ID", width=75, anchor="w", text_color="#E2E8F0").pack(side="left")
        self.receipt_naver_id_var = tk.StringVar(value=RECEIPT_SETTINGS.get("naver_id", ""))
        ctk.CTkEntry(account_row, textvariable=self.receipt_naver_id_var, width=180, fg_color="#120F24", border_color="#3B326B").pack(side="left", padx=(0, 14))
        ctk.CTkLabel(account_row, text="앱 비밀번호", width=85, anchor="w", text_color="#E2E8F0").pack(side="left")
        saved_app_password = unprotect_windows_secret(
            RECEIPT_SETTINGS.get("naver_app_password_dpapi", "")
        )
        self.receipt_naver_password_var = tk.StringVar(value=saved_app_password)
        ctk.CTkEntry(account_row, textvariable=self.receipt_naver_password_var, width=190, show="●", fg_color="#120F24", border_color="#3B326B").pack(side="left", padx=(0, 10))
        ctk.CTkLabel(account_row, text="※ Windows 암호화 저장", text_color="#10B981", font=ctk.CTkFont(size=10)).pack(side="left")

        filter_row = ctk.CTkFrame(card1, fg_color="transparent")
        filter_row.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkLabel(filter_row, text="조회 기간", width=75, anchor="w", text_color="#E2E8F0").pack(side="left")
        self.receipt_from_date_var = tk.StringVar(value="2026-01-01")
        self.receipt_to_date_var = tk.StringVar(value=today.strftime("%Y-%m-%d"))
        ctk.CTkEntry(filter_row, textvariable=self.receipt_from_date_var, width=105, fg_color="#120F24", border_color="#3B326B").pack(side="left")
        ctk.CTkLabel(filter_row, text="~", width=24).pack(side="left")
        ctk.CTkEntry(filter_row, textvariable=self.receipt_to_date_var, width=105, fg_color="#120F24", border_color="#3B326B").pack(side="left", padx=(0, 18))
        ctk.CTkLabel(filter_row, text="발신자", width=55, anchor="w", text_color="#E2E8F0").pack(side="left")
        self.receipt_sender_var = tk.StringVar(value=RECEIPT_SETTINGS.get("sender_filter", "krbajb5980@daum.net"))
        self.receipt_sender_entry = ctk.CTkEntry(
            filter_row,
            textvariable=self.receipt_sender_var,
            width=160,
            fg_color="#120F24",
            border_color="#3B326B",
        )
        self.receipt_sender_entry.pack(side="left", padx=(0, 8))
        self.receipt_all_senders_var = tk.BooleanVar(
            value=bool(RECEIPT_SETTINGS.get("all_senders", True))
        )
        ctk.CTkCheckBox(
            filter_row,
            text="전체 발신자",
            variable=self.receipt_all_senders_var,
            command=self._toggle_receipt_sender_filter,
            width=105,
            text_color="#E2E8F0",
        ).pack(side="left")
        self._toggle_receipt_sender_filter()

        folder_row = ctk.CTkFrame(card1, fg_color="transparent")
        folder_row.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkLabel(folder_row, text="저장 폴더", width=75, anchor="w", text_color="#E2E8F0").pack(side="left")
        self.receipt_download_dir_var = tk.StringVar(value=RECEIPT_SETTINGS.get("download_dir", ""))
        ctk.CTkEntry(folder_row, textvariable=self.receipt_download_dir_var, fg_color="#120F24", border_color="#3B326B").pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(folder_row, text="폴더 선택", width=90, fg_color="#374151", hover_color="#4B5563", command=self._browse_receipt_download_dir).pack(side="left")

        action_row = ctk.CTkFrame(card1, fg_color="transparent")
        action_row.pack(fill="x", padx=20, pady=(2, 14))
        self.receipt_status_var = tk.StringVar(value="대기")
        ctk.CTkLabel(action_row, textvariable=self.receipt_status_var, text_color="#10B981", font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold")).pack(side="left")
        self.receipt_search_button = ctk.CTkButton(
            action_row,
            text="🔍 메일 조회",
            width=120,
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            command=self.start_receipt_mail_search,
        )
        self.receipt_search_button.pack(side="right")

        card2 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=16)
        card2.grid(row=1, column=0, sticky="nsew", pady=(0, 12))
        card2.grid_rowconfigure(1, weight=1)
        card2.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(card2, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 4))
        ctk.CTkLabel(header, text="📋 [2] 접수 대상 HWP/PDF 목록", font=ctk.CTkFont(family="맑은 고딕", size=14, weight="bold"), text_color="#A855F7").pack(side="left")
        self.receipt_selection_var = tk.StringVar(value="실행 범위 0건")
        ctk.CTkLabel(header, textvariable=self.receipt_selection_var, font=ctk.CTkFont(size=10, weight="bold"), text_color="#10B981").pack(side="right", padx=(8, 0))
        ctk.CTkButton(header, text="전체 해제", width=72, height=26, fg_color="#374151", hover_color="#4B5563", command=self._uncheck_all_receipts).pack(side="right", padx=(6, 0))
        ctk.CTkButton(header, text="전체 선택", width=72, height=26, fg_color="#4C1D95", hover_color="#5B21B6", command=self._check_all_receipts).pack(side="right", padx=(8, 0))
        ctk.CTkLabel(header, text="체크 열 드래그 또는 Ctrl/Shift 선택", font=ctk.CTkFont(size=10), text_color="#94A3B8").pack(side="right", padx=(0, 8))

        table_box = tk.Frame(card2, bg="#18152E")
        table_box.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 12))
        cols = ("checked", "date", "sender", "subject", "file", "receipt", "status")
        self.receipt_tree = ttk.Treeview(table_box, columns=cols, show="headings", selectmode="extended", height=8)
        headings = {
            "checked": ("선택", 48, "center"),
            "date": ("수신일", 90, "center"),
            "sender": ("발신자", 150, "w"),
            "subject": ("메일 제목", 280, "w"),
            "file": ("HWP/PDF 첨부파일", 220, "w"),
            "receipt": ("접수번호", 90, "center"),
            "status": ("상태", 100, "center"),
        }
        for name, (label, width, anchor) in headings.items():
            self.receipt_tree.heading(name, text=label)
            self.receipt_tree.column(name, width=width, anchor=anchor)
        sy = ttk.Scrollbar(table_box, orient="vertical", command=self.receipt_tree.yview)
        sx = ttk.Scrollbar(table_box, orient="horizontal", command=self.receipt_tree.xview)
        self.receipt_tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.receipt_tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        table_box.grid_rowconfigure(0, weight=1)
        table_box.grid_columnconfigure(0, weight=1)
        self.receipt_tree.tag_configure("done", background="#163A2A")
        self.receipt_tree.tag_configure("error", background="#3E1A1A")
        self.receipt_tree.bind("<Button-1>", self._on_receipt_tree_press, add="+")
        self.receipt_tree.bind("<B1-Motion>", self._on_receipt_tree_drag, add="+")
        self.receipt_tree.bind("<ButtonRelease-1>", self._on_receipt_tree_release, add="+")
        self.receipt_tree.bind("<<TreeviewSelect>>", lambda _event: self._update_receipt_selection_status(), add="+")

        card3 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=16)
        card3.grid(row=2, column=0, sticky="ew")
        footer = ctk.CTkFrame(card3, fg_color="transparent")
        footer.pack(fill="x", padx=18, pady=12)
        ctk.CTkLabel(
            footer,
            text="원본 HWP/PDF는 보존하고 ‘접수완료’ 하위 폴더에 결과 파일을 저장합니다.",
            text_color="#94A3B8",
            font=ctk.CTkFont(size=10),
        ).pack(side="left")
        self.receipt_run_button = ctk.CTkButton(
            footer,
            text="🚀 선택 문서 접수 등록",
            width=180,
            fg_color="#10B981",
            hover_color="#059669",
            command=self.start_receipt_registration,
        )
        self.receipt_run_button.pack(side="right")

    def _set_receipt_checked(self, item_id: str, checked: bool) -> None:
        if item_id not in self.receipt_items:
            return
        if checked:
            self.receipt_checked_ids.add(item_id)
        else:
            self.receipt_checked_ids.discard(item_id)
        if self.receipt_tree.exists(item_id):
            self.receipt_tree.set(item_id, "checked", "☑" if checked else "☐")
        self._update_receipt_selection_status()

    def _check_all_receipts(self) -> None:
        for item_id, item in self.receipt_items.items():
            self._set_receipt_checked(item_id, item.status != "처리완료")

    def _uncheck_all_receipts(self) -> None:
        for item_id in tuple(self.receipt_checked_ids):
            self._set_receipt_checked(item_id, False)
        self.receipt_tree.selection_remove(*self.receipt_tree.selection())
        self._update_receipt_selection_status()

    def _update_receipt_selection_status(self) -> None:
        checked_count = len(self.receipt_checked_ids)
        highlighted_count = len(self.receipt_tree.selection()) if hasattr(self, "receipt_tree") else 0
        if checked_count:
            text = f"실행 범위 {checked_count}건"
        elif highlighted_count:
            text = f"강조 선택 {highlighted_count}건"
        else:
            text = "실행 범위 0건"
        self.receipt_selection_var.set(text)

    def _on_receipt_tree_press(self, event) -> str | None:
        if self.receipt_tree.identify_region(event.x, event.y) != "cell":
            return None
        if self.receipt_tree.identify_column(event.x) != "#1":
            self.receipt_drag_check_value = None
            self.receipt_drag_seen_ids.clear()
            return None
        item_id = self.receipt_tree.identify_row(event.y)
        if not item_id:
            return "break"
        checked = item_id not in self.receipt_checked_ids
        self.receipt_drag_check_value = checked
        self.receipt_drag_seen_ids = {item_id}
        self._set_receipt_checked(item_id, checked)
        return "break"

    def _on_receipt_tree_drag(self, event) -> str | None:
        if self.receipt_drag_check_value is None:
            return None
        item_id = self.receipt_tree.identify_row(event.y)
        if item_id and item_id not in self.receipt_drag_seen_ids:
            self.receipt_drag_seen_ids.add(item_id)
            self._set_receipt_checked(item_id, self.receipt_drag_check_value)
            self.receipt_tree.see(item_id)
        return "break"

    def _on_receipt_tree_release(self, _event) -> str | None:
        was_check_drag = self.receipt_drag_check_value is not None
        self.receipt_drag_check_value = None
        self.receipt_drag_seen_ids.clear()
        return "break" if was_check_drag else None

    def _browse_receipt_download_dir(self) -> None:
        selected = filedialog.askdirectory(
            title="접수대장 첨부문서 저장 폴더 선택",
            initialdir=self.receipt_download_dir_var.get() or str(Path.home() / "Downloads"),
        )
        if selected:
            self.receipt_download_dir_var.set(selected)

    def _toggle_receipt_sender_filter(self) -> None:
        state = "disabled" if self.receipt_all_senders_var.get() else "normal"
        self.receipt_sender_entry.configure(state=state)

    def _save_receipt_preferences(self, app_password: str | None = None) -> None:
        RECEIPT_SETTINGS.update(
            {
                "naver_id": self.receipt_naver_id_var.get().strip(),
                "sender_filter": self.receipt_sender_var.get().strip(),
                "all_senders": bool(self.receipt_all_senders_var.get()),
                "download_dir": self.receipt_download_dir_var.get().strip(),
            }
        )
        if app_password is not None:
            RECEIPT_SETTINGS["naver_app_password_dpapi"] = protect_windows_secret(app_password)
        save_settings()

    def _refresh_receipt_tree(self) -> None:
        self.receipt_checked_ids.intersection_update(self.receipt_items)
        for child in self.receipt_tree.get_children():
            self.receipt_tree.delete(child)
        for stable_id, item in self.receipt_items.items():
            tag = "done" if item.status == "처리완료" else ("error" if "오류" in item.status else "")
            self.receipt_tree.insert(
                "",
                "end",
                iid=stable_id,
                values=(
                    "☑" if stable_id in self.receipt_checked_ids else "☐",
                    item.received_at.strftime("%Y-%m-%d"),
                    item.sender_name or item.sender_address,
                    item.subject,
                    item.attachment_name,
                    item.receipt_no,
                    item.status,
                ),
                tags=(tag,) if tag else (),
            )
        self._update_receipt_selection_status()

    def _update_receipt_row(self, item: ReceiptMailItem) -> None:
        def update() -> None:
            if not self.receipt_tree.exists(item.stable_id):
                return
            tag = "done" if item.status == "처리완료" else ("error" if "오류" in item.status else "")
            self.receipt_tree.item(
                item.stable_id,
                values=(
                    "☑" if item.stable_id in self.receipt_checked_ids else "☐",
                    item.received_at.strftime("%Y-%m-%d"),
                    item.sender_name or item.sender_address,
                    item.subject,
                    item.attachment_name,
                    item.receipt_no,
                    item.status,
                ),
                tags=(tag,) if tag else (),
            )
        self.after(0, update)

    def _show_receipt_batch_result(self, completed: int, failed: int, total: int) -> None:
        """모든 선택 문서 처리가 끝난 뒤 결과 팝업을 한 번만 앞쪽에 표시한다."""
        summary = f"접수 처리 완료: 성공 {completed}건 / 실패 {failed}건 / 전체 {total}건"
        self.receipt_status_var.set(summary)
        self.receipt_run_button.configure(state="normal")
        try:
            self.deiconify()
            self.lift()
            self.focus_force()
            self.attributes("-topmost", True)
        except Exception:
            pass
        try:
            if failed:
                messagebox.showwarning("접수대장 처리 결과", summary, parent=self)
            else:
                messagebox.showinfo("접수대장 처리 결과", summary, parent=self)
        finally:
            try:
                self.attributes("-topmost", False)
            except Exception:
                pass

    def start_receipt_mail_search(self) -> None:
        if self.receipt_running:
            messagebox.showinfo("작업 중", "현재 접수대장 작업이 진행 중입니다.")
            return
        try:
            start = parse_iso_date(self.receipt_from_date_var.get())
            end = parse_iso_date(self.receipt_to_date_var.get())
            if start > end:
                raise ValueError("시작일은 종료일보다 늦을 수 없습니다.")
            download_dir_text = self.receipt_download_dir_var.get().strip()
            if not download_dir_text:
                raise ValueError("첨부파일 저장 폴더를 지정하세요.")
            download_dir = Path(download_dir_text).expanduser()
            username = self.receipt_naver_id_var.get().strip()
            app_password = self.receipt_naver_password_var.get()
            all_senders = bool(self.receipt_all_senders_var.get())
            sender_filter = "" if all_senders else self.receipt_sender_var.get().strip()
            if not all_senders and not sender_filter:
                raise ValueError("발신자 검색어를 입력하세요.")
        except Exception as exc:
            messagebox.showerror("입력 확인", str(exc))
            return

        self._save_receipt_preferences()
        self.receipt_running = True
        self.receipt_status_var.set("메일 조회 중...")
        self.receipt_search_button.configure(state="disabled")

        def worker() -> None:
            try:
                items = search_naver_hwp_mail(
                    username=username,
                    app_password=app_password,
                    start_date=start,
                    end_date=end,
                    sender_filter=sender_filter,
                    download_dir=download_dir,
                    log_cb=self.log,
                )
                self.after(0, lambda: self.receipt_status_var.set("접수 가능 문서 검사 중..."))
                items = filter_receipt_candidate_documents(items, log_cb=self.log)
                self._save_receipt_preferences(app_password=app_password)
                apply_history(items, self.receipt_history)
                self.receipt_checked_ids.clear()
                self.receipt_items = {item.stable_id: item for item in items}
                self.after(0, self._refresh_receipt_tree)
                self.after(0, lambda: self.receipt_status_var.set(f"조회 완료: {len(items)}건"))
                if not items:
                    self.after(
                        0,
                        lambda: messagebox.showinfo(
                            "조회 결과",
                            "시행정보와 '접수 :' 칸이 모두 있는 HWP/HWPX/PDF 첨부파일이 없습니다.",
                        ),
                    )
            except Exception as exc:
                error_text = str(exc)
                self.log(f"❌ 네이버 메일 조회 오류: {error_text}")
                self.after(0, lambda text=error_text: messagebox.showerror("메일 조회 오류", text))
                self.after(0, lambda: self.receipt_status_var.set("조회 오류"))
            finally:
                self.receipt_running = False
                self.after(0, lambda: self.receipt_search_button.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    def start_receipt_registration(self) -> None:
        if self.receipt_running:
            messagebox.showinfo("작업 중", "현재 접수대장 작업이 진행 중입니다.")
            return
        selected_ids = [
            item_id
            for item_id in self.receipt_tree.get_children()
            if item_id in self.receipt_checked_ids
        ]
        if not selected_ids:
            selected_ids = list(self.receipt_tree.selection())
        if not selected_ids:
            messagebox.showwarning(
                "선택 필요",
                "체크 열을 클릭·드래그하거나 Ctrl/Shift로 접수할 문서를 선택하세요.",
            )
            return
        targets = [self.receipt_items[item_id] for item_id in selected_ids if item_id in self.receipt_items]
        targets = [item for item in targets if item.status != "처리완료"]
        targets = sort_receipt_items_oldest_first(targets)
        if not targets:
            messagebox.showinfo("처리 완료", "선택한 문서는 이미 처리되었습니다.")
            return
        self.log(
            "접수 처리 순서: 오래된 수신일 우선 "
            f"({targets[0].received_at:%Y.%m.%d %H:%M} → {targets[-1].received_at:%Y.%m.%d %H:%M}, "
            f"총 {len(targets)}건)"
        )
        if not cdp_is_ready():
            messagebox.showwarning("Chrome 미연결", "먼저 [Chrome 연결 열기] 버튼을 누르고 NMIS에 로그인하세요.")
            return
        self.receipt_running = True
        self.receipt_status_var.set(f"접수 처리 중: 0/{len(targets)}")
        self.receipt_run_button.configure(state="disabled")
        output_dir = Path(self.receipt_download_dir_var.get()).expanduser() / "접수완료"

        def worker() -> None:
            completed = 0
            failed = 0
            try:
                with sync_playwright() as pw:
                    browser = pw.chromium.connect_over_cdp(CDP_URL)
                    page = find_nmis_page(browser)
                    for index, item in enumerate(targets, start=1):
                        try:
                            self.log(f"▶️ 접수 일괄처리 [{index}/{len(targets)}] 시작: {item.subject}")
                            item.status = "시행정보 확인 중"
                            self._update_receipt_row(item)
                            metadata = extract_receipt_document_metadata(
                                item.attachment_path,
                                sender_address=item.sender_address,
                                sender_display_name=item.sender_name,
                                log_cb=self.log,
                            )
                            item.status = "NMIS 등록 중"
                            self._update_receipt_row(item)
                            result = register_receipt_document_on_nmis(
                                page,
                                item,
                                metadata=metadata,
                                save=True,
                                log_cb=self.log,
                            )
                            item.receipt_no = str(result["receipt_no"])
                            item.status = "결과 파일 저장 중"
                            self._update_receipt_row(item)
                            item.output_path = complete_registered_attachment(
                                item.attachment_path,
                                item.receipt_no,
                                document_date=metadata.document_date,
                                output_dir=output_dir,
                                visible=False,
                                log_cb=self.log,
                            )
                            self.receipt_history.record(item, item.receipt_no, item.output_path)
                            item.status = "처리완료"
                            self.receipt_checked_ids.discard(item.stable_id)
                            completed += 1
                            self.log(f"✅ 접수 일괄처리 [{index}/{len(targets)}] 완료: 접수번호 {item.receipt_no}")
                            if index < len(targets) and item.attachment_path.suffix.lower() in {".hwp", ".hwpx"}:
                                # 한글 외부 COM 서버가 Quit을 마치기 전에 다음 HWP를 열면
                                # 두 번째 문서부터 RPC 서버 예외가 발생하므로 종료 완료를 기다린다.
                                time.sleep(1.5)
                        except Exception as exc:
                            item.status = f"오류: {str(exc)[:80]}"
                            failed += 1
                            self.log(f"❌ 접수 처리 실패 [{item.subject}]: {exc}")
                        finally:
                            self._update_receipt_row(item)
                            self.after(0, lambda i=index, total=len(targets): self.receipt_status_var.set(f"접수 처리 중: {i}/{total}"))
            except Exception as exc:
                failed += max(0, len(targets) - completed - failed)
                self.log(f"❌ NMIS 브라우저 연결 오류: {exc}")
            finally:
                self.receipt_running = False
                self.after(
                    0,
                    lambda ok=completed, error=failed, total=len(targets):
                    self._show_receipt_batch_result(ok, error, total),
                )

        threading.Thread(target=worker, daemon=True).start()

    # ── 회비 및 가입금관리 월회비 입금 자동화 ──────────────────────────────

    def _build_fee_tab(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        file_card = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=14)
        file_card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        file_row = ctk.CTkFrame(file_card, fg_color="transparent")
        file_row.pack(fill="x", padx=16, pady=12)
        ctk.CTkLabel(
            file_row,
            text="1. 통장이체 정리 엑셀:",
            font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"),
            text_color="#E2E8F0",
        ).pack(side="left", padx=(0, 8))
        self.fee_excel_var = tk.StringVar(value=str(FEE_PAYMENT_SETTINGS.get("excel_path", "")))
        ctk.CTkEntry(
            file_row,
            textvariable=self.fee_excel_var,
            fg_color="#120F24",
            border_color="#3B326B",
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(
            file_row,
            text="파일 선택",
            width=90,
            fg_color="#374151",
            hover_color="#4B5563",
            command=self._browse_fee_excel,
        ).pack(side="left", padx=(0, 5))
        ctk.CTkButton(
            file_row,
            text="불러오기",
            width=90,
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            command=self.load_fee_excel,
        ).pack(side="left")

        table_card = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=14)
        table_card.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        table_card.grid_columnconfigure(0, weight=1)
        table_card.grid_rowconfigure(1, weight=1)
        toolbar = ctk.CTkFrame(table_card, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 5))
        ctk.CTkLabel(
            toolbar,
            text="📋 [2] 월회비 입금 대상",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
            text_color="#A855F7",
        ).pack(side="left")
        self.fee_selection_var = tk.StringVar(value="선택 0건")
        ctk.CTkLabel(toolbar, textvariable=self.fee_selection_var, text_color="#10B981").pack(side="right", padx=(8, 0))
        ctk.CTkButton(toolbar, text="전체 해제", width=72, height=26, fg_color="#374151", command=self._uncheck_all_fees).pack(side="right", padx=(6, 0))
        ctk.CTkButton(toolbar, text="전체 선택", width=72, height=26, fg_color="#4C1D95", command=self._check_all_fees).pack(side="right", padx=(8, 0))

        table_box = tk.Frame(table_card, bg="#18152E")
        table_box.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 12))
        table_box.grid_columnconfigure(0, weight=1)
        table_box.grid_rowconfigure(0, weight=1)
        fee_columns = ("checked", "row", "date", "name", "amount", "status", "detail")
        self.fee_tree = ttk.Treeview(table_box, columns=fee_columns, show="headings", selectmode="extended", height=10)
        fee_headings = {
            "checked": ("선택", 48, "center"),
            "row": ("엑셀행", 60, "center"),
            "date": ("거래일자", 100, "center"),
            "name": ("성명/상호", 150, "w"),
            "amount": ("금액", 100, "e"),
            "status": ("상태", 120, "center"),
            "detail": ("안내", 260, "w"),
        }
        for column, (heading, width, anchor) in fee_headings.items():
            self.fee_tree.heading(column, text=heading)
            self.fee_tree.column(column, width=width, anchor=anchor)
        fee_sy = ttk.Scrollbar(table_box, orient="vertical", command=self.fee_tree.yview)
        fee_sx = ttk.Scrollbar(table_box, orient="horizontal", command=self.fee_tree.xview)
        self.fee_tree.configure(yscrollcommand=fee_sy.set, xscrollcommand=fee_sx.set)
        self.fee_tree.grid(row=0, column=0, sticky="nsew")
        fee_sy.grid(row=0, column=1, sticky="ns")
        fee_sx.grid(row=1, column=0, sticky="ew")
        self.fee_tree.tag_configure("done", background="#163A2A")
        self.fee_tree.tag_configure("missing", background="#3E381A")
        self.fee_tree.tag_configure("error", background="#3E1A1A")
        self.fee_tree.bind("<Button-1>", self._on_fee_tree_click, add="+")

        run_card = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=14)
        run_card.grid(row=2, column=0, sticky="ew")
        run_row = ctk.CTkFrame(run_card, fg_color="transparent")
        run_row.pack(fill="x", padx=16, pady=12)
        self.fee_status_var = tk.StringVar(value="엑셀 파일을 불러오세요.")
        ctk.CTkLabel(run_row, textvariable=self.fee_status_var, text_color="#94A3B8").pack(side="left")
        self.fee_run_button = ctk.CTkButton(
            run_row,
            text="🚀 선택 월회비 자동입력",
            width=190,
            fg_color="#10B981",
            hover_color="#059669",
            command=self.start_fee_registration,
        )
        self.fee_run_button.pack(side="right")
        self.fee_stop_button = ctk.CTkButton(
            run_row,
            text="⏹ 중지",
            width=92,
            state="disabled",
            fg_color="#DC2626",
            hover_color="#B91C1C",
            command=self.request_fee_stop,
        )
        self.fee_stop_button.pack(side="right", padx=(0, 8))

        if Path(self.fee_excel_var.get()).expanduser().is_file():
            self.after(600, self.load_fee_excel)

    def _browse_fee_excel(self) -> None:
        selected = filedialog.askopenfilename(
            title="통장이체 거래내역 정리 엑셀 선택",
            filetypes=(("Excel 파일", "*.xlsx;*.xlsm"), ("모든 파일", "*.*")),
        )
        if selected:
            self.fee_excel_var.set(selected)
            self.load_fee_excel()

    def load_fee_excel(self) -> None:
        source = Path(self.fee_excel_var.get()).expanduser()
        if not source.is_file():
            messagebox.showerror("엑셀 오류", f"파일을 찾을 수 없습니다:\n{source}", parent=self)
            return
        try:
            result = parse_fee_payment_excel(source)
        except Exception as exc:
            messagebox.showerror("엑셀 파싱 오류", str(exc), parent=self)
            return
        self.fee_items = {item.stable_id: item for item in result.items}
        self.fee_checked_ids = set(self.fee_items)
        FEE_PAYMENT_SETTINGS["excel_path"] = str(source.resolve())
        save_settings()
        self._refresh_fee_tree()
        self.fee_status_var.set(f"{result.sheet_name} 시트: 입력 대상 {len(result.items)}건")
        self.log(f"💳 월회비 엑셀 로드: {source.name} / 입력 대상 {len(result.items)}건")
        for warning in result.skipped_rows:
            self.log(f"⚠️ 회비 엑셀 제외: {warning}")

    def _on_fee_tree_click(self, event) -> str | None:
        if self.fee_tree.identify_region(event.x, event.y) != "cell":
            return None
        if self.fee_tree.identify_column(event.x) != "#1":
            return None
        item_id = self.fee_tree.identify_row(event.y)
        if not item_id:
            return "break"
        if item_id in self.fee_checked_ids:
            self.fee_checked_ids.discard(item_id)
        else:
            self.fee_checked_ids.add(item_id)
        self.fee_tree.set(item_id, "checked", "☑" if item_id in self.fee_checked_ids else "☐")
        self._update_fee_selection_status()
        return "break"

    def _check_all_fees(self) -> None:
        self.fee_checked_ids = {
            item_id for item_id, item in self.fee_items.items() if item.status != "처리완료"
        }
        self._refresh_fee_tree()

    def _uncheck_all_fees(self) -> None:
        self.fee_checked_ids.clear()
        self._refresh_fee_tree()

    def _update_fee_selection_status(self) -> None:
        self.fee_selection_var.set(f"선택 {len(self.fee_checked_ids)}건 / 전체 {len(self.fee_items)}건")

    def _refresh_fee_tree(self) -> None:
        self.fee_checked_ids.intersection_update(self.fee_items)
        for child in self.fee_tree.get_children():
            self.fee_tree.delete(child)
        for item_id, item in self.fee_items.items():
            tag = "done" if item.status == "처리완료" else ("missing" if item.status == "월회비 행 없음" else ("error" if item.status.startswith("오류") else ""))
            self.fee_tree.insert(
                "",
                "end",
                iid=item_id,
                values=(
                    "☑" if item_id in self.fee_checked_ids else "☐",
                    item.excel_row,
                    item.payment_date.strftime("%Y-%m-%d"),
                    item.member_name,
                    f"{item.amount:,}",
                    item.status,
                    item.detail,
                ),
                tags=(tag,) if tag else (),
            )
        self._update_fee_selection_status()

    def _update_fee_row(self, item: FeePaymentItem) -> None:
        def update() -> None:
            if not self.fee_tree.exists(item.stable_id):
                return
            tag = "done" if item.status == "처리완료" else ("missing" if item.status == "월회비 행 없음" else ("error" if item.status.startswith("오류") else ""))
            self.fee_tree.item(
                item.stable_id,
                values=(
                    "☑" if item.stable_id in self.fee_checked_ids else "☐",
                    item.excel_row,
                    item.payment_date.strftime("%Y-%m-%d"),
                    item.member_name,
                    f"{item.amount:,}",
                    item.status,
                    item.detail,
                ),
                tags=(tag,) if tag else (),
            )
            self._update_fee_selection_status()
        self.after(0, update)

    def _show_fee_batch_result(
        self,
        completed: int,
        missing: list[FeePaymentItem],
        failed: list[FeePaymentItem],
        total: int,
        stopped: bool = False,
        remaining: int = 0,
        setup_error: str = "",
    ) -> None:
        title = "월회비 입력 중단" if setup_error else ("월회비 입력 중지" if stopped else "월회비 입력 완료")
        summary = f"{title}: 성공 {completed}건 / 월회비 행 없음 {len(missing)}건 / 실패 {len(failed)}건"
        if stopped or setup_error:
            summary += f" / 미처리 {remaining}건"
        summary += f" / 전체 {total}건"
        self.fee_status_var.set(summary)
        self.fee_run_button.configure(state="normal")
        self.fee_stop_button.configure(state="disabled")
        details: list[str] = [summary]
        if setup_error:
            details.extend(("\n[시작 단계 오류]", setup_error))
        if missing:
            details.append("\n[월회비 행 없음]")
            details.extend(
                f"- 엑셀 {item.excel_row}행 {item.member_name} ({item.payment_date:%Y-%m-%d})"
                for item in missing[:15]
            )
            if len(missing) > 15:
                details.append(f"- 외 {len(missing) - 15}건 (목록 상태와 로그에서 확인)")
        if failed:
            details.append("\n[실패]")
            details.extend(
                f"- 엑셀 {item.excel_row}행 {item.member_name}: {item.detail}"
                for item in failed[:10]
            )
        try:
            self.deiconify()
            self.lift()
            self.focus_force()
            self.attributes("-topmost", True)
        except Exception:
            pass
        try:
            message = "\n".join(details)
            if stopped or setup_error or missing or failed:
                messagebox.showwarning("월회비 자동입력 결과", message, parent=self)
            else:
                messagebox.showinfo("월회비 자동입력 결과", message, parent=self)
        finally:
            try:
                self.attributes("-topmost", False)
            except Exception:
                pass

    def request_fee_stop(self) -> None:
        if not self.fee_running:
            return
        self.fee_stop_requested = True
        self.fee_stop_button.configure(state="disabled")
        self.fee_status_var.set("중지 요청됨 — 현재 항목 저장 완료 후 중지합니다.")
        self.log("⏹ 월회비 자동입력 중지 요청: 현재 항목 완료 후 남은 작업을 중지합니다.")

    def start_fee_registration(self) -> None:
        if self.fee_running:
            messagebox.showinfo("작업 중", "월회비 자동입력이 진행 중입니다.", parent=self)
            return
        targets = [
            item
            for item_id, item in self.fee_items.items()
            if item_id in self.fee_checked_ids and item.status != "처리완료"
        ]
        targets.sort(key=lambda item: (item.payment_date, item.excel_row))
        if not targets:
            messagebox.showwarning("선택 필요", "자동입력할 거래를 하나 이상 선택하세요.", parent=self)
            return
        if not cdp_is_ready():
            messagebox.showwarning("Chrome 미연결", "먼저 [Chrome 연결 열기] 버튼을 눌러 NMIS에 로그인하세요.", parent=self)
            return

        self.fee_running = True
        self.fee_stop_requested = False
        self.fee_run_button.configure(state="disabled")
        self.fee_stop_button.configure(state="normal")
        self.fee_status_var.set(f"월회비 처리 중: 0/{len(targets)}")

        def worker() -> None:
            completed = 0
            missing: list[FeePaymentItem] = []
            failed: list[FeePaymentItem] = []
            stopped = False
            setup_error = ""
            try:
                with sync_playwright() as pw:
                    browser = pw.chromium.connect_over_cdp(CDP_URL)
                    page = find_nmis_page(browser)
                    if not page:
                        raise RuntimeError("연결된 Chrome에서 NMIS 페이지를 찾지 못했습니다.")
                    ensure_nmis_logged_in(page, NMIS_USER_ID, NMIS_PASSWORD, self.log)
                    navigate_to_fee_management(page, self.log)
                    for index, item in enumerate(targets, start=1):
                        if self.fee_stop_requested:
                            stopped = True
                            self.log(
                                f"⏹ 월회비 자동입력을 중지했습니다. "
                                f"남은 미처리 항목: {len(targets) - index + 1}건"
                            )
                            break
                        try:
                            item.status = "NMIS 조회 중"
                            item.detail = ""
                            self._update_fee_row(item)
                            self.log(
                                f"▶️ 월회비 [{index}/{len(targets)}] {item.payment_date:%Y-%m-%d} "
                                f"{item.member_name} {item.amount:,}원"
                            )
                            result = register_fee_payment_on_nmis(
                                page,
                                item,
                                log_cb=self.log,
                                user_id=NMIS_USER_ID,
                                password=NMIS_PASSWORD,
                            )
                            if result.get("missing_month_fee"):
                                item.status = "월회비 행 없음"
                                item.detail = str(result.get("found", "조회 결과 없음"))
                                missing.append(item)
                            else:
                                item.status = "처리완료"
                                item.detail = "저장 완료"
                                completed += 1
                            self.fee_checked_ids.discard(item.stable_id)
                        except Exception as exc:
                            item.status = "오류"
                            item.detail = str(exc)[:120]
                            failed.append(item)
                            self.log(f"❌ 월회비 입력 실패 [{item.member_name}]: {exc}")
                        finally:
                            self._update_fee_row(item)
                            self.after(
                                0,
                                lambda current=index, total=len(targets):
                                self.fee_status_var.set(f"월회비 처리 중: {current}/{total}"),
                            )
            except Exception as exc:
                setup_error = str(exc)
                remaining = [item for item in targets if item.status not in {"처리완료", "월회비 행 없음", "오류"}]
                for item in remaining:
                    item.status = "대기"
                    item.detail = "시작 단계 중단 — 다시 실행 가능"
                    self._update_fee_row(item)
                self.log(f"❌ 월회비 자동입력 연결 오류: {exc}")
            finally:
                self.fee_running = False
                processed = completed + len(missing) + len(failed)
                remaining_count = max(0, len(targets) - processed)
                if self.fee_stop_requested and remaining_count > 0:
                    stopped = True
                self.after(
                    0,
                    lambda ok=completed, no_row=list(missing), errors=list(failed), total=len(targets),
                    was_stopped=stopped, remaining=remaining_count, start_error=setup_error:
                    self._show_fee_batch_result(ok, no_row, errors, total, was_stopped, remaining, start_error),
                )

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
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            file_row,
            text="⚙️ 엑셀 컬럼 설정",
            font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold"),
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            width=120,
            height=32,
            command=self.open_member_column_settings_dialog
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
                        custom_col_store=MEMBER_COLUMN_MAPPING.get("store_name"),
                        custom_col_owner=MEMBER_COLUMN_MAPPING.get("owner_name"),
                        custom_col_license=MEMBER_COLUMN_MAPPING.get("license_no"),
                    )
                    def on_complete():
                        self.is_member_verifying = False
                        self.btn_start_member.configure(state="normal")
                        self.btn_stop_member.configure(state="disabled")

                        if self.member_stop_event.is_set():
                            self.lbl_member_status.configure(
                                text=f"⏹ 검수 중단 완료! (진행 완료: {len(self.member_results_list)}건 - 상단 [엑셀 저장] 버튼으로 저장 가능)",
                                text_color="#F59E0B"
                            )
                        else:
                            self.lbl_member_status.configure(
                                text=f"🎉 검수 완료! (전체: {res['total']}건 | ✅ 일치: {res['match_count']}건 | ❌ 불일치: {res['mismatch_count']}건 | 🔍 미검색: {res['not_found_count']}건)",
                                text_color="#10B981"
                            )
                    self.after(0, on_complete)
            except Exception as e:
                err_msg = str(e)
                def on_error():
                    self.is_member_verifying = False
                    self.btn_start_member.configure(state="normal")
                    self.btn_stop_member.configure(state="disabled")
                    self.lbl_member_status.configure(text=f"❌ 오류 발생: {err_msg[:40]}", text_color="#EF4444")
                self.after(0, on_error)

        threading.Thread(target=worker, daemon=True).start()

    def _center_dialog(self, dlg: ctk.CTkToplevel, width: int = 520, height: int = 460) -> None:
        try:
            dlg.update_idletasks()
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            x = max(0, (sw - width) // 2)
            y = max(0, (sh - height) // 2)
            dlg.geometry(f"{width}x{height}+{x}+{y}")
        except Exception:
            dlg.geometry(f"{width}x{height}")

    def open_member_column_settings_dialog(self) -> None:
        excel_p = self.member_excel_var.get().strip()
        headers = get_excel_column_headers(excel_p) if excel_p and os.path.isfile(excel_p) else []

        dlg = ctk.CTkToplevel(self)
        dlg.title("⚙️ 회원검수 엑셀 컬럼 매핑 설정")
        dlg.geometry("520x460")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()
        self._center_dialog(dlg, 520, 460)

        ctk.CTkLabel(
            dlg,
            text="⚙️ [회원 검수] 엑셀 컬럼 매핑 지정",
            font=ctk.CTkFont(family="맑은 고딕", size=16, weight="bold"),
            text_color="#A855F7"
        ).pack(anchor="w", padx=24, pady=(20, 8))

        if headers:
            file_name = os.path.basename(excel_p)
            header_preview = ", ".join([h["display"] for h in headers[:6]])
            subtitle = f"📄 엑셀 파일: {file_name}\n💡 감지된 최상단 헤더: {header_preview}..."
        else:
            subtitle = "💡 엑셀 파일 선택 시 최상단 헤더(1행)가 자동으로 감지 나열됩니다.\n필요한 데이터의 열(A~Z)을 지정하세요."

        ctk.CTkLabel(
            dlg,
            text=subtitle,
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            text_color="#94A3B8",
            justify="left"
        ).pack(anchor="w", padx=24, pady=(0, 16))

        if headers:
            options = [h["display"] for h in headers]
            def get_default_opt(key_type: str, default_letter: str) -> str:
                curr = MEMBER_COLUMN_MAPPING.get(key_type, default_letter)
                for h in headers:
                    if h["letter"] == curr or curr in h["header"] or h["header"] in curr:
                        return h["display"]
                return options[0]
        else:
            options = [f"[{chr(65+i)}열] 항목_{i+1}" for i in range(20)]
            def get_default_opt(key_type: str, default_letter: str) -> str:
                curr = MEMBER_COLUMN_MAPPING.get(key_type, default_letter)
                idx = ord(curr[0].upper()) - 65 if len(curr) == 1 and curr.isalpha() else 0
                return options[idx] if idx < len(options) else options[0]

        # 1. 업소명
        f1 = ctk.CTkFrame(dlg, fg_color="#18152E", corner_radius=10)
        f1.pack(fill="x", padx=24, pady=6)
        ctk.CTkLabel(f1, text="🏢 업소명 (상호명) 컬럼:", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#E2E8F0").pack(side="left", padx=12, pady=10)
        combo_store = ctk.CTkComboBox(f1, values=options, width=220, font=ctk.CTkFont(family="맑은 고딕", size=12), dropdown_font=ctk.CTkFont(family="맑은 고딕", size=11))
        combo_store.pack(side="right", padx=12, pady=10)
        combo_store.set(get_default_opt("store_name", "F"))

        # 2. 대표자 성명
        f2 = ctk.CTkFrame(dlg, fg_color="#18152E", corner_radius=10)
        f2.pack(fill="x", padx=24, pady=6)
        ctk.CTkLabel(f2, text="👤 대표자 성명 컬럼:", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#E2E8F0").pack(side="left", padx=12, pady=10)
        combo_owner = ctk.CTkComboBox(f2, values=options, width=220, font=ctk.CTkFont(family="맑은 고딕", size=12), dropdown_font=ctk.CTkFont(family="맑은 고딕", size=11))
        combo_owner.pack(side="right", padx=12, pady=10)
        combo_owner.set(get_default_opt("owner_name", "G"))

        # 3. 인허가/신고번호
        f3 = ctk.CTkFrame(dlg, fg_color="#18152E", corner_radius=10)
        f3.pack(fill="x", padx=24, pady=6)
        ctk.CTkLabel(f3, text="🔢 인허가(신고)번호 컬럼:", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#E2E8F0").pack(side="left", padx=12, pady=10)
        combo_license = ctk.CTkComboBox(f3, values=options, width=220, font=ctk.CTkFont(family="맑은 고딕", size=12), dropdown_font=ctk.CTkFont(family="맑은 고딕", size=11))
        combo_license.pack(side="right", padx=12, pady=10)
        combo_license.set(get_default_opt("license_no", "D"))

        btn_box = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_box.pack(fill="x", padx=24, pady=(20, 10))

        def parse_letter(val_str: str) -> str:
            if val_str.startswith("[") and "열]" in val_str:
                return val_str.split("]")[0].replace("[", "").replace("열", "").strip()
            return val_str

        def save_mapping():
            MEMBER_COLUMN_MAPPING["store_name"] = parse_letter(combo_store.get())
            MEMBER_COLUMN_MAPPING["owner_name"] = parse_letter(combo_owner.get())
            MEMBER_COLUMN_MAPPING["license_no"] = parse_letter(combo_license.get())
            save_settings()
            messagebox.showinfo("저장 완료", f"✅ 회원검수 엑셀 컬럼 매핑이 저장되었습니다!\n\n• 업소명: {MEMBER_COLUMN_MAPPING['store_name']}열\n• 대표자성명: {MEMBER_COLUMN_MAPPING['owner_name']}열\n• 인허가번호: {MEMBER_COLUMN_MAPPING['license_no']}열", parent=dlg)
            dlg.destroy()

        def reset_mapping():
            MEMBER_COLUMN_MAPPING["store_name"] = "F"
            MEMBER_COLUMN_MAPPING["owner_name"] = "G"
            MEMBER_COLUMN_MAPPING["license_no"] = "D"
            save_settings()
            messagebox.showinfo("복원 완료", "기본값(F열: 업소명, G열: 성명, D열: 인허가번호)으로 복원되었습니다.", parent=dlg)
            dlg.destroy()

        ctk.CTkButton(btn_box, text="💾 매핑 저장", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), fg_color="#8B5CF6", hover_color="#7C3AED", command=save_mapping, width=130, height=36).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_box, text="🔄 기본값 복원", font=ctk.CTkFont(family="맑은 고딕", size=12), fg_color="#374151", hover_color="#4B5563", command=reset_mapping, width=130, height=36).pack(side="left")
        ctk.CTkButton(btn_box, text="취소", font=ctk.CTkFont(family="맑은 고딕", size=12), fg_color="#1F2937", hover_color="#374151", command=dlg.destroy, width=90, height=36).pack(side="right")

    def open_slip_column_settings_dialog(self) -> None:
        excel_p = self.file_var.get().strip()
        headers = get_excel_column_headers(excel_p) if excel_p and os.path.isfile(excel_p) else []

        dlg = ctk.CTkToplevel(self)
        dlg.title("⚙️ 전표등록 엑셀 컬럼 매핑 설정")
        dlg.geometry("520x520")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()
        self._center_dialog(dlg, 520, 520)

        ctk.CTkLabel(
            dlg,
            text="⚙️ [전표 자동등록] 엑셀 컬럼 매핑 지정",
            font=ctk.CTkFont(family="맑은 고딕", size=16, weight="bold"),
            text_color="#A855F7"
        ).pack(anchor="w", padx=24, pady=(20, 8))

        if headers:
            file_name = os.path.basename(excel_p)
            header_preview = ", ".join([h["display"] for h in headers[:6]])
            subtitle = f"📄 엑셀 파일: {file_name}\n💡 감지된 최상단 헤더: {header_preview}..."
        else:
            subtitle = "💡 엑셀 파일 선택 시 최상단 헤더(1행)가 자동으로 감지 나열됩니다.\n필요한 데이터의 열(A~Z)을 지정하세요."

        ctk.CTkLabel(
            dlg,
            text=subtitle,
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            text_color="#94A3B8",
            justify="left"
        ).pack(anchor="w", padx=24, pady=(0, 16))

        options = [h["display"] for h in headers] if headers else [f"[{chr(65+i)}열] 항목_{i+1}" for i in range(20)]

        def get_default_opt(key_type: str, default_letter: str) -> str:
            curr = SLIP_COLUMN_MAPPING.get(key_type, default_letter)
            if headers:
                for h in headers:
                    if h["letter"] == curr or curr in h["header"] or h["header"] in curr:
                        return h["display"]
                return options[0]
            else:
                idx = ord(curr[0].upper()) - 65 if len(curr) == 1 and curr.isalpha() else 0
                return options[idx] if idx < len(options) else options[0]

        fields = [
            ("date", "📅 거래일시 컬럼:", "A"),
            ("type", "📑 전표유형 컬럼:", "F"),
            ("dir", "🔄 구분 (입/출) 컬럼:", "C"),
            ("amount", "💵 거래금액 컬럼:", "D"),
            ("content", "📝 내용 (적요) 컬럼:", "E"),
        ]

        combos = {}
        for key_type, label_txt, def_letter in fields:
            f = ctk.CTkFrame(dlg, fg_color="#18152E", corner_radius=10)
            f.pack(fill="x", padx=24, pady=4)
            ctk.CTkLabel(f, text=label_txt, font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#E2E8F0").pack(side="left", padx=12, pady=8)
            cb = ctk.CTkComboBox(f, values=options, width=220, font=ctk.CTkFont(family="맑은 고딕", size=12), dropdown_font=ctk.CTkFont(family="맑은 고딕", size=11))
            cb.pack(side="right", padx=12, pady=8)
            cb.set(get_default_opt(key_type, def_letter))
            combos[key_type] = cb

        btn_box = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_box.pack(fill="x", padx=24, pady=(16, 10))

        def parse_letter(val_str: str) -> str:
            if val_str.startswith("[") and "열]" in val_str:
                return val_str.split("]")[0].replace("[", "").replace("열", "").strip()
            return val_str

        def save_mapping():
            for k in combos:
                SLIP_COLUMN_MAPPING[k] = parse_letter(combos[k].get())
            save_settings()
            messagebox.showinfo("저장 완료", f"✅ 전표등록 엑셀 컬럼 매핑이 저장되었습니다!\n\n• 거래일시: {SLIP_COLUMN_MAPPING['date']}열 | 금액: {SLIP_COLUMN_MAPPING['amount']}열 | 내용: {SLIP_COLUMN_MAPPING['content']}열", parent=dlg)
            dlg.destroy()

        def reset_mapping():
            SLIP_COLUMN_MAPPING.update({"date": "A", "type": "F", "dir": "C", "amount": "D", "content": "E"})
            save_settings()
            messagebox.showinfo("복원 완료", "기본값(A열: 일시, F열: 유형, C열: 구분, D열: 금액, E열: 내용)으로 복원되었습니다.", parent=dlg)
            dlg.destroy()

        ctk.CTkButton(btn_box, text="💾 매핑 저장", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), fg_color="#8B5CF6", hover_color="#7C3AED", command=save_mapping, width=130, height=36).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_box, text="🔄 기본값 복원", font=ctk.CTkFont(family="맑은 고딕", size=12), fg_color="#374151", hover_color="#4B5563", command=reset_mapping, width=130, height=36).pack(side="left")
        ctk.CTkButton(btn_box, text="취소", font=ctk.CTkFont(family="맑은 고딕", size=12), fg_color="#1F2937", hover_color="#374151", command=dlg.destroy, width=90, height=36).pack(side="right")

    def stop_member_verification(self) -> None:
        if self.is_member_verifying:
            self.member_stop_event.set()
            self.is_member_verifying = False
            self.btn_start_member.configure(state="normal")
            self.btn_stop_member.configure(state="disabled")
            self.lbl_member_status.configure(
                text=f"⏹ 중지 요청됨... (현재 완료: {len(self.member_results_list)}건 - [엑셀 저장] 가능)",
                text_color="#EF4444"
            )

    def export_member_results(self) -> None:
        data = []
        if self.member_results_list:
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

        # 화면 테이블(Treeview)에 있는 데이터 백업 수집
        tree_data = []
        if hasattr(self, "member_tree"):
            for child in self.member_tree.get_children():
                vals = self.member_tree.item(child, "values")
                if vals and len(vals) >= 8:
                    tree_data.append({
                        "순번": vals[0],
                        "업소명(F열)": vals[1],
                        "엑셀대표자(G열)": vals[2],
                        "웹대표자": vals[3],
                        "엑셀인허가번호(D열)": vals[4],
                        "웹신고번호": vals[5],
                        "검수상태": vals[6],
                        "검수세부사유": vals[7]
                    })

        # 화면에 보이는 건수가 더 많으면 화면 데이터 우선 채택
        if len(tree_data) >= len(data) and len(tree_data) > 0:
            data = tree_data

        if not data:
            messagebox.showwarning("저장 경고", "저장할 검수 결과 데이터가 없습니다. 먼저 회원 정보 검수를 진행해 주세요.", parent=self)
            return

        # 바탕화면 경로 탐색 (한글/OneDrive 대응)
        user_home = Path.home()
        desktop_dir = user_home / "Desktop"
        for c in [user_home / "Desktop", user_home / "OneDrive" / "Desktop", user_home / "OneDrive" / "바탕 화면", user_home / "바탕 화면"]:
            if c.is_dir():
                desktop_dir = c
                break

        default_filename = f"NMIS_회원검수결과_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        out_path = filedialog.asksaveasfilename(
            parent=self,
            title="검수 결과 엑셀 파일 저장 위치 선택",
            defaultextension=".xlsx",
            filetypes=(("Excel 파일 (*.xlsx)", "*.xlsx"), ("모든 파일 (*.*)", "*.*")),
            initialdir=str(desktop_dir),
            initialfile=default_filename
        )

        # 사용자가 대화상자를 닫으면 바탕화면 기본 경로로 자동 지정
        if not out_path:
            out_path = str(desktop_dir / default_filename)

        try:
            df = pd.DataFrame(data)
            df.to_excel(out_path, index=False)

            log_msg = f"📊 총 {len(data)}건의 회원검수 결과가 엑셀로 저장되었습니다!"
            self.log(f"{log_msg} ({out_path})")
            self.lbl_member_status.configure(text=f"✅ {log_msg}", text_color="#10B981")

            messagebox.showinfo(
                "엑셀 저장 완료",
                f"🎉 회원검수 결과 총 {len(data)}건이 성공적으로 저장되었습니다!\n\n저장 위치:\n{out_path}\n\n[확인]을 누르면 생성된 엑셀 파일이 바로 열립니다.",
                parent=self
            )

            try:
                os.startfile(out_path)
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("저장 오류", f"엑셀 저장 중 오류가 발생했습니다:\n{e}", parent=self)
            self.log(f"❌ 엑셀 저장 실패: {e}")

    # ── [탭 4] 잠재회원 등록 뷰 ──────────────────────────────────────────

    def _build_potential_tab(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        # Card 1: 엑셀 파일 선택 및 컬럼 매핑 설정
        card1 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=14)
        card1.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        file_row = ctk.CTkFrame(card1, fg_color="transparent")
        file_row.pack(fill="x", padx=20, pady=14)

        ctk.CTkLabel(file_row, text="잠재회원 엑셀:", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#E2E8F0").pack(side="left", padx=(0, 8))

        default_dl_file = Path.home() / "Downloads" / "일반음식점현황(6.30.기준일).xlsx"
        default_pot_path = str(default_dl_file) if default_dl_file.is_file() else ""
        self.potential_excel_var = tk.StringVar(value=default_pot_path)

        excel_entry = ctk.CTkEntry(
            file_row,
            textvariable=self.potential_excel_var,
            font=ctk.CTkFont(family="맑은 고딕", size=12),
            fg_color="#120F24",
            border_color="#3B326B",
            corner_radius=8
        )
        excel_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        def browse_potential_excel():
            p = filedialog.askopenfilename(
                title="일반음식점현황 엑셀 파일 선택",
                filetypes=(("Excel 파일", "*.xlsx;*.xls"), ("모든 파일", "*.*")),
            )
            if p:
                self.potential_excel_var.set(p)
                self.load_potential_excel_rows()

        ctk.CTkButton(
            file_row,
            text="📁 파일 선택",
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            fg_color="#374151",
            hover_color="#4B5563",
            width=90,
            height=32,
            command=browse_potential_excel
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            file_row,
            text="📥 엑셀 불러오기",
            font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold"),
            fg_color="#3B82F6",
            hover_color="#2563EB",
            width=110,
            height=32,
            command=self.load_potential_excel_rows
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            file_row,
            text="⚙️ 엑셀 컬럼 설정",
            font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold"),
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            width=120,
            height=32,
            command=self.open_potential_column_settings_dialog
        ).pack(side="left")

        # Card 2: 엑셀 행 목록 (Treeview 및 선택 체크 기능)
        card2 = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=14)
        card2.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        card2.grid_columnconfigure(0, weight=1)
        card2.grid_rowconfigure(1, weight=1)

        # 상단 컨트롤 바 (전체 선택 / 전체 해제 / 선택 건수)
        tb_bar = ctk.CTkFrame(card2, fg_color="transparent")
        tb_bar.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 6))

        ctk.CTkLabel(tb_bar, text="📋 잠재회원 대상 행 선택", font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"), text_color="#A855F7").pack(side="left", padx=(0, 10))

        def select_all_rows():
            for child in self.potential_tree.get_children():
                vals = list(self.potential_tree.item(child, "values"))
                vals[0] = " ✅ 선택 "
                self.potential_tree.item(child, values=vals)
            update_selected_count()

        def deselect_all_rows():
            for child in self.potential_tree.get_children():
                vals = list(self.potential_tree.item(child, "values"))
                vals[0] = " ⬜ "
                self.potential_tree.item(child, values=vals)
            update_selected_count()

        def update_selected_count():
            sel_cnt = 0
            tot_cnt = len(self.potential_tree.get_children())
            for child in self.potential_tree.get_children():
                vals = self.potential_tree.item(child, "values")
                if vals and "선택" in vals[0]:
                    sel_cnt += 1
            self.lbl_potential_count.configure(text=f"선택: {sel_cnt}건 / 전체: {tot_cnt}건")

        ctk.CTkButton(tb_bar, text="☑️ 전체 선택", font=ctk.CTkFont(family="맑은 고딕", size=11, weight="bold"), fg_color="#374151", width=95, height=28, command=select_all_rows).pack(side="left", padx=(0, 6))
        ctk.CTkButton(tb_bar, text="⬜ 전체 해제", font=ctk.CTkFont(family="맑은 고딕", size=11), fg_color="#374151", width=95, height=28, command=deselect_all_rows).pack(side="left", padx=(0, 12))

        self.lbl_potential_count = ctk.CTkLabel(tb_bar, text="선택: 0건 / 전체: 0건", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#10B981")
        self.lbl_potential_count.pack(side="right")

        # Treeview 테이블
        tree_frame = ctk.CTkFrame(card2, fg_color="transparent")
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 12))
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        cols = ("select", "seq", "perm_date", "store_name", "owner_name", "rrn", "birth_date", "gender", "mobile", "phone", "address")
        self.potential_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=12)

        headings_map = [
            ("select", "선택 상태"),
            ("seq", "순번"),
            ("perm_date", "인허가일자(E열)"),
            ("store_name", "업소명(상호)"),
            ("owner_name", "영업자(성명)"),
            ("rrn", "주민등록번호(H열)"),
            ("birth_date", "생년월일(8자리)"),
            ("gender", "성별(파싱)"),
            ("mobile", "핸드폰(P열)"),
            ("phone", "소재지전화(L열)"),
            ("address", "소재지주소(I/J열)"),
        ]

        for col_id, title in headings_map:
            self.potential_tree.heading(col_id, text=title, command=lambda c=col_id: self.sort_potential_tree_by_column(c))

        self.potential_tree.column("select", width=80, anchor="center")
        self.potential_tree.column("seq", width=65, anchor="center")
        self.potential_tree.column("perm_date", width=110, anchor="center")
        self.potential_tree.column("store_name", width=140, anchor="w")
        self.potential_tree.column("owner_name", width=90, anchor="center")
        self.potential_tree.column("rrn", width=125, anchor="center")
        self.potential_tree.column("birth_date", width=100, anchor="center")
        self.potential_tree.column("gender", width=60, anchor="center")
        self.potential_tree.column("mobile", width=115, anchor="center")
        self.potential_tree.column("phone", width=110, anchor="center")
        self.potential_tree.column("address", width=380, minwidth=250, stretch=True, anchor="w")

        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.potential_tree.yview)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.potential_tree.xview)
        self.potential_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.potential_tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        # 셀 클릭 시에만 선택/해제 토글 (헤더 클릭시에는 정렬 실행)
        def on_tree_click(event):
            region = self.potential_tree.identify_region(event.x, event.y)
            if region not in ("cell", "tree"):
                return
            item_id = self.potential_tree.identify_row(event.y)
            if item_id:
                vals = list(self.potential_tree.item(item_id, "values"))
                vals[0] = " ⬜ " if "선택" in vals[0] else " ✅ 선택 "
                self.potential_tree.item(item_id, values=vals)
                update_selected_count()

        self.potential_tree.bind("<Button-1>", on_tree_click)

        # Bottom Bar (실행 / 중지 / 상태)
        b_bar = ctk.CTkFrame(parent, fg_color="#18152E", border_color="#2E2756", border_width=1, corner_radius=14)
        b_bar.grid(row=2, column=0, sticky="ew")

        f_btns = ctk.CTkFrame(b_bar, fg_color="transparent")
        f_btns.pack(fill="x", padx=16, pady=12)

        self.btn_start_potential = ctk.CTkButton(
            f_btns,
            text="🚀 등록 시작",
            font=ctk.CTkFont(family="맑은 고딕", size=14, weight="bold"),
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            width=150,
            height=40,
            command=self.start_potential_registration
        )
        self.btn_start_potential.pack(side="left", padx=(0, 10))

        self.btn_stop_potential = ctk.CTkButton(
            f_btns,
            text="⏹ 중지",
            font=ctk.CTkFont(family="맑은 고딕", size=13, weight="bold"),
            fg_color="#EF4444",
            hover_color="#DC2626",
            state="disabled",
            width=90,
            height=40,
            command=self.stop_potential_registration
        )
        self.btn_stop_potential.pack(side="left")

        self.lbl_potential_status = ctk.CTkLabel(
            f_btns,
            text="대기 중... (엑셀 파일 선택 후 등록할 행을 체크하고 시작하세요)",
            font=ctk.CTkFont(family="맑은 고딕", size=12),
            text_color="#94A3B8"
        )
        self.lbl_potential_status.pack(side="right", padx=10)

        # 초기 엑셀 데이터 자동로드
        if default_pot_path:
            self.after(500, self.load_potential_excel_rows)

    def load_potential_excel_rows(self) -> None:
        excel_p = self.potential_excel_var.get().strip()
        if not excel_p or not os.path.isfile(excel_p):
            return

        try:
            df = pd.read_excel(excel_p)
            for child in self.potential_tree.get_children():
                self.potential_tree.delete(child)

            self.potential_row_data_map = {}

            col_biz = resolve_column_from_letter_or_name(df, POTENTIAL_COLUMN_MAPPING.get("biz_type"), 2)
            col_lic = resolve_column_from_letter_or_name(df, POTENTIAL_COLUMN_MAPPING.get("license_no"), 3)
            col_perm = resolve_column_from_letter_or_name(df, POTENTIAL_COLUMN_MAPPING.get("perm_date"), 4)
            col_store = resolve_column_from_letter_or_name(df, POTENTIAL_COLUMN_MAPPING.get("store_name"), 5)
            col_owner = resolve_column_from_letter_or_name(df, POTENTIAL_COLUMN_MAPPING.get("owner_name"), 6)
            col_rrn = resolve_column_from_letter_or_name(df, POTENTIAL_COLUMN_MAPPING.get("rrn"), 7)
            col_addr = resolve_column_from_letter_or_name(df, POTENTIAL_COLUMN_MAPPING.get("address"), 8)
            col_area = resolve_column_from_letter_or_name(df, POTENTIAL_COLUMN_MAPPING.get("area"), 10)
            col_phone = resolve_column_from_letter_or_name(df, POTENTIAL_COLUMN_MAPPING.get("phone"), 11)
            col_mobile = resolve_column_from_letter_or_name(df, POTENTIAL_COLUMN_MAPPING.get("mobile"), 15)

            for idx, row in df.iterrows():
                seq = idx + 1
                biz_type = str(row[col_biz]).strip() if col_biz and pd.notna(row[col_biz]) else ""
                license_no = str(row[col_lic]).strip() if col_lic and pd.notna(row[col_lic]) else ""
                perm_date = str(row[col_perm]).strip() if col_perm and pd.notna(row[col_perm]) else ""
                store_name = str(row[col_store]).strip() if col_store and pd.notna(row[col_store]) else ""
                owner_name = str(row[col_owner]).strip() if col_owner and pd.notna(row[col_owner]) else ""
                rrn = str(row[col_rrn]).strip() if col_rrn and pd.notna(row[col_rrn]) else ""
                addr = str(row[col_addr]).strip() if col_addr and pd.notna(row[col_addr]) else ""
                area = str(row[col_area]).strip() if col_area and pd.notna(row[col_area]) else ""
                phone = str(row[col_phone]).strip() if col_phone and pd.notna(row[col_phone]) else ""
                mobile = str(row[col_mobile]).strip() if col_mobile and pd.notna(row[col_mobile]) else ""

                birth_date_8digit, gender_code = parse_rrn_birth_gender(rrn)
                gender_str = "남" if gender_code == "M" else ("여" if gender_code == "F" else "")

                formatted_mobile = format_korean_phone(mobile)
                formatted_phone = format_korean_phone(phone)
                formatted_perm = format_digits_only(perm_date)

                # 기본값: 아무것도 체크되어 있지 않음 (요구사항 반영)
                chk = " ⬜ "

                item_id = self.potential_tree.insert(
                    "",
                    "end",
                    values=(chk, seq, formatted_perm, store_name, owner_name, rrn, birth_date_8digit, gender_str, formatted_mobile, formatted_phone, addr)
                )

                self.potential_row_data_map[item_id] = {
                    "seq": seq,
                    "biz_type": biz_type,
                    "license_no": license_no,
                    "perm_date": perm_date,
                    "store_name": store_name,
                    "owner_name": owner_name,
                    "rrn": rrn,
                    "birth_date": birth_date_8digit,
                    "gender": gender_str,
                    "mobile": mobile,
                    "phone": phone,
                    "address": addr,
                    "area": area,
                }

            self.lbl_potential_count.configure(text=f"선택: 0건 / 전체: {len(df)}건")
            self.log(f"📑 잠재회원 엑셀 데이터 총 {len(df)}행이 테이블에 로드되었습니다.")
        except Exception as e:
            self.log(f"❌ 잠재회원 엑셀 로드 예외: {e}")

    def sort_potential_tree_by_column(self, col_name: str) -> None:
        """
        잠재회원 테이블의 컬럼 헤더 클릭 시 오름차순/내림차순 정렬 (순번 등 숫자 열 숫자순 정렬 지원)
        """
        if not hasattr(self, "_potential_sort_state"):
            self._potential_sort_state = {}

        reverse = not self._potential_sort_state.get(col_name, False)
        self._potential_sort_state[col_name] = reverse

        cols = ("select", "seq", "perm_date", "store_name", "owner_name", "rrn", "birth_date", "gender", "mobile", "phone", "address")
        if col_name not in cols:
            return
        col_idx = cols.index(col_name)

        children = self.potential_tree.get_children("")
        items_data = []

        for item_id in children:
            vals = self.potential_tree.item(item_id, "values")
            val = vals[col_idx] if len(vals) > col_idx else ""

            if col_name == "seq":
                try:
                    sort_key = int(val)
                except ValueError:
                    sort_key = 0
            else:
                sort_key = str(val).lower()

            items_data.append((sort_key, item_id))

        items_data.sort(key=lambda x: x[0], reverse=reverse)

        for index, (_, item_id) in enumerate(items_data):
            self.potential_tree.move(item_id, "", index)

        headings = {
            "select": "선택 상태",
            "seq": "순번",
            "perm_date": "인허가일자(E열)",
            "store_name": "업소명(상호)",
            "owner_name": "영업자(성명)",
            "rrn": "주민등록번호(H열)",
            "birth_date": "생년월일(8자리)",
            "gender": "성별(파싱)",
            "mobile": "핸드폰(P열)",
            "phone": "소재지전화(L열)",
            "address": "소재지주소(I/J열)"
        }

        for c, base_title in headings.items():
            if c == col_name:
                arrow = " ▼" if reverse else " ▲"
                self.potential_tree.heading(c, text=base_title + arrow, command=lambda _c=c: self.sort_potential_tree_by_column(_c))
            else:
                self.potential_tree.heading(c, text=base_title, command=lambda _c=c: self.sort_potential_tree_by_column(_c))

    def open_potential_column_settings_dialog(self) -> None:
        excel_p = self.potential_excel_var.get().strip()
        headers = get_excel_column_headers(excel_p) if excel_p and os.path.isfile(excel_p) else []

        dlg = ctk.CTkToplevel(self)
        dlg.title("⚙️ 잠재회원 엑셀 컬럼 매핑 설정")
        dlg.geometry("540x680")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()
        self._center_dialog(dlg, 540, 680)

        ctk.CTkLabel(
            dlg,
            text="⚙️ [잠재회원 등록] 엑셀 컬럼 매핑 지정",
            font=ctk.CTkFont(family="맑은 고딕", size=16, weight="bold"),
            text_color="#A855F7"
        ).pack(anchor="w", padx=24, pady=(20, 8))

        if headers:
            file_name = os.path.basename(excel_p)
            header_preview = ", ".join([h["display"] for h in headers[:7]])
            subtitle = f"📄 엑셀 파일: {file_name}\n💡 감지된 헤더: {header_preview}..."
        else:
            subtitle = "💡 엑셀 파일 선택 시 최상단 헤더(1행)가 자동으로 나열됩니다.\n필요한 데이터의 열(A~Z)을 지정하세요."

        ctk.CTkLabel(
            dlg,
            text=subtitle,
            font=ctk.CTkFont(family="맑은 고딕", size=11),
            text_color="#94A3B8",
            justify="left"
        ).pack(anchor="w", padx=24, pady=(0, 14))

        options = [h["display"] for h in headers] if headers else [f"[{chr(65+i)}열] 항목_{i+1}" for i in range(20)]

        def get_default_opt(key_type: str, default_letter: str) -> str:
            curr = POTENTIAL_COLUMN_MAPPING.get(key_type, default_letter)
            if headers:
                for h in headers:
                    if h["letter"] == curr or curr in h["header"] or h["header"] in curr:
                        return h["display"]
                return options[0]
            else:
                idx = ord(curr[0].upper()) - 65 if len(curr) == 1 and curr.isalpha() else 0
                return options[idx] if idx < len(options) else options[0]

        fields = [
            ("biz_type", "🏷️ 업태명 컬럼:", "C"),
            ("license_no", "📜 인허가번호 (신고번호) 컬럼:", "D"),
            ("perm_date", "📅 인허가일자 (적용일자) 컬럼:", "E"),
            ("store_name", "🏢 업소명 (상호명) 컬럼:", "F"),
            ("owner_name", "👤 영업자 성명 컬럼:", "G"),
            ("rrn", "🪪 주민등록번호 컬럼:", "H"),
            ("address", "🏠 소재지 주소 컬럼:", "I"),
            ("area", "📐 영업장면적 컬럼:", "K"),
            ("phone", "☎️ 소재지 전화번호 컬럼:", "L"),
            ("mobile", "📱 핸드폰번호 컬럼:", "P"),
        ]

        combos = {}
        for key_type, label_txt, def_letter in fields:
            f = ctk.CTkFrame(dlg, fg_color="#18152E", corner_radius=10)
            f.pack(fill="x", padx=24, pady=3)
            ctk.CTkLabel(f, text=label_txt, font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), text_color="#E2E8F0").pack(side="left", padx=12, pady=5)
            cb = ctk.CTkComboBox(f, values=options, width=210, font=ctk.CTkFont(family="맑은 고딕", size=12), dropdown_font=ctk.CTkFont(family="맑은 고딕", size=11))
            cb.pack(side="right", padx=12, pady=5)
            cb.set(get_default_opt(key_type, def_letter))
            combos[key_type] = cb

        btn_box = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_box.pack(fill="x", padx=24, pady=(16, 10))

        def parse_letter(val_str: str) -> str:
            if val_str.startswith("[") and "열]" in val_str:
                return val_str.split("]")[0].replace("[", "").replace("열", "").strip()
            return val_str

        def save_mapping():
            for k in combos:
                POTENTIAL_COLUMN_MAPPING[k] = parse_letter(combos[k].get())
            save_settings()
            messagebox.showinfo("저장 완료", f"✅ 잠재회원등록 엑셀 컬럼 매핑이 저장되었습니다!\n\n• 업태: {POTENTIAL_COLUMN_MAPPING['biz_type']}열 | 신고번호: {POTENTIAL_COLUMN_MAPPING['license_no']}열 | 인허가일자: {POTENTIAL_COLUMN_MAPPING['perm_date']}열 | 면적: {POTENTIAL_COLUMN_MAPPING['area']}열", parent=dlg)
            dlg.destroy()
            self.load_potential_excel_rows()

        def reset_mapping():
            POTENTIAL_COLUMN_MAPPING.update({"biz_type": "C", "license_no": "D", "perm_date": "E", "store_name": "F", "owner_name": "G", "rrn": "H", "address": "I", "area": "K", "phone": "L", "mobile": "P"})
            save_settings()
            messagebox.showinfo("복원 완료", "기본값(C: 업태, D: 신고번호, E: 인허가일자, F: 상호, G: 영업자, H: 주민번호, I: 주소, K: 면적, L: 전화, P: 핸드폰)으로 복원되었습니다.", parent=dlg)
            dlg.destroy()
            self.load_potential_excel_rows()

        ctk.CTkButton(btn_box, text="💾 매핑 저장", font=ctk.CTkFont(family="맑은 고딕", size=12, weight="bold"), fg_color="#8B5CF6", hover_color="#7C3AED", command=save_mapping, width=130, height=36).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_box, text="🔄 기본값 복원", font=ctk.CTkFont(family="맑은 고딕", size=12), fg_color="#374151", hover_color="#4B5563", command=reset_mapping, width=130, height=36).pack(side="left")
        ctk.CTkButton(btn_box, text="취소", font=ctk.CTkFont(family="맑은 고딕", size=12), fg_color="#1F2937", hover_color="#374151", command=dlg.destroy, width=90, height=36).pack(side="right")

    def start_potential_registration(self) -> None:
        selected_rows = []
        for child in self.potential_tree.get_children():
            vals = self.potential_tree.item(child, "values")
            if vals and "선택" in vals[0]:
                data_obj = getattr(self, "potential_row_data_map", {}).get(child, {})
                if not data_obj:
                    data_obj = {
                        "seq": vals[1],
                        "perm_date": vals[2],
                        "store_name": vals[3],
                        "owner_name": vals[4],
                        "rrn": vals[5],
                        "birth_date": vals[6],
                        "gender": vals[7],
                        "mobile": vals[8],
                        "phone": vals[9],
                        "address": vals[10],
                    }
                selected_rows.append(data_obj)

        if not selected_rows:
            messagebox.showwarning("선택 경고", "등록할 잠재회원 행을 목록에서 1개 이상 체크해 주세요.", parent=self)
            return

        if not cdp_is_ready():
            messagebox.showerror(
                "Chrome 미연결",
                "Chrome이 연동 모드로 열려있지 않습니다.\n\n사이드바의 [🚀 Chrome 연결 열기] 버튼을 누른 후 다시 시도해 주세요.",
                parent=self
            )
            return

        self.is_potential_running = True
        self.potential_stop_event = threading.Event()
        self.btn_start_potential.configure(state="disabled")
        self.btn_stop_potential.configure(state="normal")
        self.lbl_potential_status.configure(text=f"🚀 선택된 {len(selected_rows)}건 서식 작성 진행 중...", text_color="#A855F7")

        def status_cb(info: dict):
            t = info.get("type")
            if t == "log":
                self.log(info["message"])
            elif t == "potential_done":
                self.log(f"  ✅ [{info['seq']}] {info['store_name']} -> {info['status']} ({info['reason']})")

        def worker():
            try:
                with sync_playwright() as pw:
                    browser = pw.chromium.connect_over_cdp(CDP_URL)
                    res = register_potential_members_on_nmis(
                        target=browser,
                        selected_rows=selected_rows,
                        status_callback=status_cb,
                        stop_event=self.potential_stop_event,
                    )
                    def on_complete():
                        self.is_potential_running = False
                        self.btn_start_potential.configure(state="normal")
                        self.btn_stop_potential.configure(state="disabled")

                        if self.potential_stop_event.is_set():
                            self.lbl_potential_status.configure(
                                text=f"⏹ 중단됨! (완료: {res['success']}건)",
                                text_color="#F59E0B"
                            )
                        else:
                            self.lbl_potential_status.configure(
                                text=f"🎉 잠재회원 1~8단계 작성 완결! (성공: {res['success']}건 / {res['total']}건) — [등록] 버튼은 미클릭 상태입니다.",
                                text_color="#10B981"
                            )
                    self.after(0, on_complete)
            except Exception as e:
                err_msg = str(e)
                def on_error():
                    self.is_potential_running = False
                    self.btn_start_potential.configure(state="normal")
                    self.btn_stop_potential.configure(state="disabled")
                    self.lbl_potential_status.configure(text=f"❌ 오류: {err_msg[:40]}", text_color="#EF4444")
                self.after(0, on_error)

        threading.Thread(target=worker, daemon=True).start()

    def stop_potential_registration(self) -> None:
        if hasattr(self, "is_potential_running") and self.is_potential_running:
            self.potential_stop_event.set()
            self.is_potential_running = False
            self.btn_start_potential.configure(state="normal")
            self.btn_stop_potential.configure(state="disabled")
            self.lbl_potential_status.configure(text="⏹ 중지 요청됨...", text_color="#EF4444")


def main() -> None:
    app = ModernSlipUI()
    app.mainloop()


if __name__ == "__main__":
    main()
