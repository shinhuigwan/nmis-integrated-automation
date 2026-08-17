import tempfile
import unittest
import imaplib
import sys
import types
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from unittest.mock import patch

from receipt_register import (
    ReceiptHistory,
    ReceiptMailItem,
    copy_registered_pdf,
    decode_mail_header,
    ensure_hancom_security_module_registered,
    extract_document_number,
    extract_pdf_receipt_metadata,
    filter_receipt_candidate_documents,
    format_hwp_receipt_entry,
    has_receipt_slot,
    parse_hwp_receipt_metadata,
    parse_iso_date,
    protect_windows_secret,
    sanitize_filename,
    search_naver_hwp_mail,
    sort_receipt_items_oldest_first,
    unprotect_windows_secret,
)
from receipt_register import _find_text_hwp, _move_to_separate_receipt_value_cell


class ReceiptRegisterTests(unittest.TestCase):
    def test_selected_receipts_are_processed_oldest_first(self):
        def item(uid: str, received_at: datetime) -> ReceiptMailItem:
            return ReceiptMailItem(
                uid=uid,
                message_id=f"<mail-{uid}@example.com>",
                received_at=received_at,
                sender_name="발신자",
                sender_address="sender@example.com",
                subject=f"문서 {uid}",
                attachment_name=f"문서_{uid}.hwp",
                attachment_path=Path(f"문서_{uid}.hwp"),
                attachment_sha256=uid * 8,
            )

        selected = [
            item("30", datetime(2026, 1, 3, 9, 0)),
            item("10", datetime(2026, 1, 1, 9, 0)),
            item("20", datetime(2026, 1, 2, 9, 0)),
        ]
        ordered = sort_receipt_items_oldest_first(selected)
        self.assertEqual([entry.uid for entry in ordered], ["10", "20", "30"])

    def test_hancom_security_module_is_registered_and_loaded(self):
        registry_values = {}

        class FakeKey:
            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        fake_winreg = types.SimpleNamespace(
            HKEY_CURRENT_USER=1,
            KEY_READ=1,
            KEY_WRITE=2,
            REG_SZ=1,
            CreateKeyEx=lambda *_args: FakeKey(),
            QueryValueEx=lambda _key, name: (
                (registry_values[name], 1)
                if name in registry_values
                else (_ for _ in ()).throw(FileNotFoundError())
            ),
            SetValueEx=lambda _key, name, _reserved, _type, value: registry_values.__setitem__(name, value),
        )
        fake_hwp = types.SimpleNamespace(RegisterModule=lambda kind, name: kind == "FilePathCheckDLL" and name == "FilePathCheckerModuleExample")
        with patch.dict(sys.modules, {"winreg": fake_winreg}):
            path = ensure_hancom_security_module_registered(fake_hwp)
        self.assertTrue(path.name.endswith(".dll"))
        self.assertEqual(registry_values["FilePathCheckerModuleExample"], str(path))

    def test_hancom_security_module_reuses_first_success_for_next_hwp(self):
        registry_values = {}

        class FakeKey:
            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        fake_winreg = types.SimpleNamespace(
            HKEY_CURRENT_USER=1,
            KEY_READ=2,
            KEY_WRITE=4,
            REG_SZ=1,
            CreateKeyEx=lambda *_args: FakeKey(),
            QueryValueEx=lambda _key, name: (
                (registry_values[name], 1)
                if name in registry_values
                else (_ for _ in ()).throw(FileNotFoundError())
            ),
            SetValueEx=lambda _key, name, _reserved, _type, value: registry_values.__setitem__(name, value),
        )
        load_results = iter((True, False))
        fake_hwp = types.SimpleNamespace(RegisterModule=lambda *_args: next(load_results))
        with (
            patch.dict(sys.modules, {"winreg": fake_winreg}),
            patch("receipt_register._HANCOM_SECURITY_MODULE_LOADED", False),
        ):
            ensure_hancom_security_module_registered(fake_hwp)
            ensure_hancom_security_module_registered(fake_hwp)

    def test_two_cell_receipt_layout_moves_to_same_row_value_cell(self):
        class FakeAction:
            def __init__(self, owner):
                self.owner = owner

            def Run(self, action):
                if action == "TableRightCell":
                    self.owner.cell = "M5"
                    return True
                return True

        fake_hwp = types.SimpleNamespace(cell="J5")
        fake_hwp.HAction = FakeAction(fake_hwp)
        fake_hwp.KeyIndicator = lambda: (True, f"({fake_hwp.cell}): 문자 입력")
        self.assertTrue(_move_to_separate_receipt_value_cell(fake_hwp))

    def test_one_cell_receipt_layout_rejects_next_row_cell(self):
        class FakeAction:
            def __init__(self, owner):
                self.owner = owner

            def Run(self, action):
                if action == "TableRightCell":
                    self.owner.cell = "A6"
                    return True
                return True

        fake_hwp = types.SimpleNamespace(cell="F5")
        fake_hwp.HAction = FakeAction(fake_hwp)
        fake_hwp.KeyIndicator = lambda: (True, f"({fake_hwp.cell}): 문자 입력")
        self.assertFalse(_move_to_separate_receipt_value_cell(fake_hwp))

    def test_hwp_text_search_starts_at_document_begin_for_last_page(self):
        class Params:
            HSet = object()

        class FakeAction:
            def __init__(self):
                self.at_document_begin = False
                self.executed = False

            def Run(self, action):
                if action == "MoveDocBegin":
                    self.at_document_begin = True
                    return True
                return True

            def GetDefault(self, *_args):
                return None

            def Execute(self, action, _params):
                self.executed = action == "RepeatFind"
                return self.at_document_begin

        action = FakeAction()
        params = Params()
        fake_hwp = types.SimpleNamespace(
            HAction=action,
            HParameterSet=types.SimpleNamespace(HFindReplace=params),
            FindDir=lambda value: value,
        )
        self.assertTrue(_find_text_hwp(fake_hwp, "접수 :"))
        self.assertTrue(action.at_document_begin)
        self.assertTrue(action.executed)
        self.assertEqual(params.Direction, "Forward")

    def test_parse_iso_date_accepts_dot_and_dash(self):
        self.assertEqual(parse_iso_date("2026-01-02").isoformat(), "2026-01-02")
        self.assertEqual(parse_iso_date("2026.01.02").isoformat(), "2026-01-02")

    def test_decode_korean_mail_header(self):
        encoded = "=?UTF-8?B?7ZWc6rWt7Jm47Iud7JeF7KSR7JWZ7ZqM?="
        self.assertEqual(decode_mail_header(encoded), "한국외식업중앙회")

    def test_extract_document_number(self):
        self.assertEqual(extract_document_number("시행번호: 총무 120-33 업무보고"), "총무 120-33")
        self.assertEqual(extract_document_number("[정경 650] 음식문화개선"), "정경 650")

    def test_parse_hwp_enforcement_line_for_receipt_ledger(self):
        text = """
        수신처 :
        시행 : 경영800-1 (2026. 1. 2.)        접수 :
        전북특별자치도지회
        Email. krbajb5980@daum.net
        """
        metadata = parse_hwp_receipt_metadata(
            text,
            "문성운부장모친상알림시행문_20260102.hwp",
            "krbajb5980@daum.net",
        )
        self.assertEqual(metadata.classification_no, "경영800-1")
        self.assertEqual(metadata.document_date.isoformat(), "2026-01-02")
        self.assertEqual(metadata.document_title, "문성운부장모친상알림시행문")
        self.assertEqual(metadata.sender_name, "전북특별자치도지회")
        self.assertEqual(metadata.receiver_name, "진안군지부")
        self.assertEqual(metadata.in_out_label, "내부")

    def test_parse_hwp_strips_ho_suffix_and_requires_enforcement_line(self):
        metadata = parse_hwp_receipt_metadata(
            "시행 : 총무 120-108호(2026.04.13) 접수 :",
            "공문_20260413.hwp",
            "krbajb5980@daum.net",
        )
        self.assertEqual(metadata.classification_no, "총무120-108")
        self.assertEqual(metadata.document_date.isoformat(), "2026-04-13")
        with self.assertRaisesRegex(ValueError, "시행"):
            parse_hwp_receipt_metadata("일반 안내문", "안내.hwp")

    def test_external_sender_name_is_used_for_all_sender_mode(self):
        metadata = parse_hwp_receipt_metadata(
            "시행 : 행정지원과-1234 (2026. 8. 15.)",
            "군청공문_20260815.pdf",
            sender_address="office@example.go.kr",
            sender_display_name="진안군청",
        )
        self.assertEqual(metadata.sender_name, "진안군청")
        self.assertEqual(metadata.in_out_label, "외부")

    def test_receipt_entry_contains_number_and_document_date(self):
        self.assertEqual(
            format_hwp_receipt_entry("15", parse_iso_date("2026-04-13")),
            "15 (2026.04.13)",
        )

    def test_sanitize_filename_removes_windows_reserved_characters(self):
        self.assertEqual(sanitize_filename('보고서:<1>?*.hwp'), "보고서__1___.hwp")

    def test_history_round_trip(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "a.hwp"
            output = Path(tmp) / "done.hwp"
            source.write_bytes(b"source")
            output.write_bytes(b"output")
            item = ReceiptMailItem(
                uid="7",
                message_id="<mail-7@example.com>",
                received_at=datetime(2026, 1, 2, 10, 0),
                sender_name="한국외식업중앙회",
                sender_address="sender@example.com",
                subject="테스트 문서",
                attachment_name="a.hwp",
                attachment_path=source,
                attachment_sha256="abc",
            )
            history = ReceiptHistory(Path(tmp) / "history.sqlite3")
            history.record(item, "2026-001", output)
            found = history.find(item.stable_id)
            self.assertIsNotNone(found)
            self.assertEqual(found["receipt_no"], "2026-001")

    def test_receipt_date_uses_mail_received_date(self):
        item = ReceiptMailItem(
            uid="received-date",
            message_id="<received-date@example.com>",
            received_at=datetime(2026, 1, 5, 9, 30),
            sender_name="전북특별자치도지회",
            sender_address="krbajb5980@daum.net",
            subject="수신일 테스트",
            attachment_name="문서.hwp",
            attachment_path=Path("문서.hwp"),
            attachment_sha256="received-date",
        )
        self.assertEqual(item.received_date_raw, "20260105")

    def test_search_mail_filters_sender_and_saves_hwp_attachment(self):
        message = EmailMessage()
        message["From"] = "한국외식업중앙회 <sender@example.com>"
        message["To"] = "user@naver.com"
        message["Subject"] = "총무 120 업무보고"
        message["Date"] = "Fri, 02 Jan 2026 10:00:00 +0900"
        message["Message-ID"] = "<mail-1@example.com>"
        message.set_content("첨부 문서를 접수해 주세요.")
        message.add_attachment(
            b"fake-hwp",
            maintype="application",
            subtype="x-hwp",
            filename="업무보고.hwp",
        )
        raw = message.as_bytes()

        class FakeImap:
            def login(self, *_args):
                return "OK", []

            def select(self, *_args, **_kwargs):
                return "OK", []

            def uid(self, command, *_args):
                if command == "search":
                    return "OK", [b"1"]
                return "OK", [(b"1 (BODY[])", raw), b")"]

            def logout(self):
                return "BYE", []

        with tempfile.TemporaryDirectory() as tmp:
            with patch("receipt_register.imaplib.IMAP4_SSL", return_value=FakeImap()):
                items = search_naver_hwp_mail(
                    username="user",
                    app_password="app-password",
                    start_date=parse_iso_date("2026-01-01"),
                    end_date=parse_iso_date("2026-01-31"),
                    sender_filter="한국외식업중앙회",
                    download_dir=tmp,
                )
            self.assertEqual(len(items), 1)
            self.assertEqual(items[0].attachment_name, "업무보고.hwp")
            self.assertEqual(items[0].attachment_path.name, "업무보고_20260102.hwp")
            self.assertEqual(items[0].attachment_path.read_bytes(), b"fake-hwp")

    def test_search_mail_all_senders_includes_pdf_and_hwp(self):
        message = EmailMessage()
        message["From"] = "진안군청 <office@example.go.kr>"
        message["To"] = "user@naver.com"
        message["Subject"] = "군청 공문"
        message["Date"] = "Fri, 02 Jan 2026 10:00:00 +0900"
        message["Message-ID"] = "<mail-pdf@example.com>"
        message.set_content("PDF 공문을 확인하세요.")
        message.add_attachment(
            b"fake-pdf",
            maintype="application",
            subtype="pdf",
            filename="군청공문.pdf",
        )
        raw = message.as_bytes()

        class FakeImap:
            def login(self, *_args):
                return "OK", []

            def select(self, *_args, **_kwargs):
                return "OK", []

            def uid(self, command, *_args):
                if command == "search":
                    return "OK", [b"1"]
                return "OK", [(b"1 (BODY[])", raw), b")"]

            def logout(self):
                return "BYE", []

        with tempfile.TemporaryDirectory() as tmp:
            with patch("receipt_register.imaplib.IMAP4_SSL", return_value=FakeImap()):
                items = search_naver_hwp_mail(
                    username="user",
                    app_password="app-password",
                    start_date=parse_iso_date("2026-01-01"),
                    end_date=parse_iso_date("2026-01-31"),
                    sender_filter="",
                    download_dir=tmp,
                )
            self.assertEqual(len(items), 1)
            self.assertEqual(items[0].sender_name, "진안군청")
            self.assertEqual(items[0].attachment_path.name, "군청공문_20260102.pdf")

    def test_extract_pdf_metadata_from_text_pages(self):
        class FakePage:
            def extract_text(self):
                return "진안군청\n시행 : 행정지원과-42호 (2026. 8. 15.)\n접수 :"

        fake_pypdf = types.SimpleNamespace(
            PdfReader=lambda _path: types.SimpleNamespace(pages=[FakePage()])
        )
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "군청공문_20260815.pdf"
            source.write_bytes(b"fake-pdf")
            with patch.dict(sys.modules, {"pypdf": fake_pypdf}):
                metadata = extract_pdf_receipt_metadata(
                    source,
                    sender_address="office@example.go.kr",
                    sender_display_name="진안군청",
                )
            self.assertEqual(metadata.classification_no, "행정지원과-42")
            self.assertEqual(metadata.document_date.isoformat(), "2026-08-15")
            self.assertEqual(metadata.sender_name, "진안군청")
            self.assertEqual(metadata.in_out_label, "외부")

    def test_receipt_slot_detection_and_mail_prefilter(self):
        self.assertTrue(has_receipt_slot("시행 : 총무1 (2026.1.1.)  접 수 :"))
        self.assertFalse(has_receipt_slot("시행 : 총무1 (2026.1.1.)"))
        with tempfile.TemporaryDirectory() as tmp:
            valid_path = Path(tmp) / "valid.hwp"
            invalid_path = Path(tmp) / "invalid.hwp"
            valid_path.write_bytes(b"valid")
            invalid_path.write_bytes(b"invalid")

            def item(path: Path, uid: str) -> ReceiptMailItem:
                return ReceiptMailItem(
                    uid=uid,
                    message_id=f"<mail-{uid}@example.com>",
                    received_at=datetime(2026, 1, 2, 10, 0),
                    sender_name="전북특별자치도지회",
                    sender_address="krbajb5980@daum.net",
                    subject=path.stem,
                    attachment_name=path.name,
                    attachment_path=path,
                    attachment_sha256=uid,
                )

            texts = {
                valid_path: "시행 : 총무120-1 (2026.1.2.)\n접수 :",
                invalid_path: "시행 : 총무120-2 (2026.1.2.)",
            }
            with patch(
                "receipt_register.extract_attachment_text_fast",
                side_effect=lambda path: texts[Path(path)],
            ):
                filtered = filter_receipt_candidate_documents(
                    [item(valid_path, "1"), item(invalid_path, "2")]
                )
            self.assertEqual([entry.attachment_name for entry in filtered], ["valid.hwp"])
            self.assertEqual(filtered[0].status, "접수 가능")
            self.assertIsNotNone(filtered[0].metadata)

    def test_registered_pdf_is_copied_without_modification(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "공문.pdf"
            output_dir = Path(tmp) / "접수완료"
            source.write_bytes(b"original-pdf")
            output = copy_registered_pdf(
                source,
                "15",
                parse_iso_date("2026-08-15"),
                output_dir=output_dir,
            )
            self.assertEqual(output.read_bytes(), b"original-pdf")
            self.assertEqual(source.read_bytes(), b"original-pdf")

    def test_duplicate_attachment_names_keep_only_latest_mail(self):
        def build_message(date_header: str, message_id: str, payload: bytes) -> bytes:
            message = EmailMessage()
            message["From"] = "외식업전북지회 <krbajb5980@daum.net>"
            message["To"] = "user@naver.com"
            message["Subject"] = "중복 공문"
            message["Date"] = date_header
            message["Message-ID"] = message_id
            message.set_content("첨부 공문")
            message.add_attachment(
                payload,
                maintype="application",
                subtype="x-hwp",
                filename="같은이름.hwp",
            )
            return message.as_bytes()

        raws = {
            "1": build_message("Thu, 01 Jan 2026 10:00:00 +0900", "<old@example.com>", b"old"),
            "2": build_message("Sat, 03 Jan 2026 10:00:00 +0900", "<new@example.com>", b"new"),
        }

        class FakeImap:
            def login(self, *_args):
                return "OK", []

            def select(self, *_args, **_kwargs):
                return "OK", []

            def uid(self, command, *args):
                if command == "search":
                    return "OK", [b"1 2"]
                uid = str(args[0])
                return "OK", [(b"BODY", raws[uid]), b")"]

            def logout(self):
                return "BYE", []

        with tempfile.TemporaryDirectory() as tmp:
            with patch("receipt_register.imaplib.IMAP4_SSL", return_value=FakeImap()):
                items = search_naver_hwp_mail(
                    username="user",
                    app_password="app-password",
                    start_date=parse_iso_date("2026-01-01"),
                    end_date=parse_iso_date("2026-01-31"),
                    sender_filter="krbajb5980@daum.net",
                    download_dir=tmp,
                )
            self.assertEqual(len(items), 1)
            self.assertEqual(items[0].attachment_path.name, "같은이름_20260103.hwp")
            self.assertEqual(items[0].attachment_path.read_bytes(), b"new")

    def test_imap_auth_error_explains_app_password_requirement(self):
        class RejectingImap:
            def login(self, *_args):
                raise imaplib.IMAP4.error(b"AUTHENTICATIONFAILED")

            def logout(self):
                return "BYE", []

        with tempfile.TemporaryDirectory() as tmp:
            with patch("receipt_register.imaplib.IMAP4_SSL", return_value=RejectingImap()):
                with self.assertRaisesRegex(RuntimeError, "일반 로그인 비밀번호가 아닌"):
                    search_naver_hwp_mail(
                        username="user",
                        app_password="wrong-password",
                        start_date=parse_iso_date("2026-01-01"),
                        end_date=parse_iso_date("2026-01-31"),
                        sender_filter="한국외식업중앙회",
                        download_dir=tmp,
                    )

    def test_app_password_is_stored_as_dpapi_ciphertext(self):
        fake_win32crypt = types.SimpleNamespace(
            CryptProtectData=lambda data, *_args: ("", b"encrypted:" + data[::-1]),
            CryptUnprotectData=lambda data, *_args: ("", data.removeprefix(b"encrypted:")[::-1]),
        )
        with patch.dict(sys.modules, {"win32crypt": fake_win32crypt}):
            protected = protect_windows_secret("secret-value")
            self.assertTrue(protected.startswith("dpapi:"))
            self.assertNotIn("secret-value", protected)
            self.assertEqual(unprotect_windows_secret(protected), "secret-value")

    def test_app_password_supports_new_pywin32_bytes_result(self):
        fake_win32crypt = types.SimpleNamespace(
            CryptProtectData=lambda data, *_args: b"encrypted:" + data[::-1],
            CryptUnprotectData=lambda data, *_args: data.removeprefix(b"encrypted:")[::-1],
        )
        with patch.dict(sys.modules, {"win32crypt": fake_win32crypt}):
            protected = protect_windows_secret("new-api-secret")
            self.assertTrue(protected.startswith("dpapi:"))
            self.assertNotIn("new-api-secret", protected)
            self.assertEqual(unprotect_windows_secret(protected), "new-api-secret")


if __name__ == "__main__":
    unittest.main()
