from __future__ import annotations

import base64
import hashlib
import imaplib
import re
import shutil
import sqlite3
import struct
import zlib
import zipfile
from contextlib import closing
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from email import policy
from email.header import decode_header, make_header
from email.parser import BytesParser
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, Callable, Iterable
from xml.etree import ElementTree

if TYPE_CHECKING:
    from playwright.sync_api import Page
else:
    Page = Any


NAVER_IMAP_HOST = "imap.naver.com"
NAVER_IMAP_PORT = 993
HWP_SUFFIXES = {".hwp", ".hwpx"}
PDF_SUFFIXES = {".pdf"}
RECEIPT_ATTACHMENT_SUFFIXES = HWP_SUFFIXES | PDF_SUFFIXES
DPAPI_PREFIX = "dpapi:"
HANCOM_SECURITY_MODULE_NAME = "FilePathCheckerModuleExample"
HANCOM_SECURITY_REGISTRY_KEY = r"Software\HNC\HwpAutomation\Modules"
_HANCOM_SECURITY_MODULE_LOADED = False


def ensure_hancom_security_module_registered(hwp: Any | None = None) -> Path:
    """공식 한컴 파일경로 보안 모듈을 현재 사용자 계정에 등록하고 로드한다."""
    global _HANCOM_SECURITY_MODULE_LOADED
    module_path = (
        Path(__file__).resolve().parent
        / "hancom_security"
        / "FilePathCheckerModuleExample.dll"
    )
    if not module_path.is_file():
        raise FileNotFoundError(f"한컴 자동화 보안 모듈을 찾을 수 없습니다: {module_path}")
    try:
        import winreg

        with winreg.CreateKeyEx(
            winreg.HKEY_CURRENT_USER,
            HANCOM_SECURITY_REGISTRY_KEY,
            0,
            winreg.KEY_READ | winreg.KEY_WRITE,
        ) as key:
            try:
                registered_path, _ = winreg.QueryValueEx(key, HANCOM_SECURITY_MODULE_NAME)
            except FileNotFoundError:
                registered_path = ""
            if str(registered_path) != str(module_path):
                winreg.SetValueEx(
                    key,
                    HANCOM_SECURITY_MODULE_NAME,
                    0,
                    winreg.REG_SZ,
                    str(module_path),
                )
    except OSError as exc:
        raise RuntimeError(f"한컴 자동화 보안 모듈을 레지스트리에 등록하지 못했습니다: {exc}") from exc
    if hwp is not None:
        loaded_now = bool(
            hwp.RegisterModule("FilePathCheckDLL", HANCOM_SECURITY_MODULE_NAME)
        )
        if loaded_now:
            _HANCOM_SECURITY_MODULE_LOADED = True
        elif not _HANCOM_SECURITY_MODULE_LOADED:
            raise RuntimeError("한컴 자동화 보안 모듈을 불러오지 못했습니다.")
        # 같은 자동화 프로세스에서 두 번째 HWP를 열 때 RegisterModule이 False를
        # 반환해도 첫 문서에서 이미 로드된 모듈은 계속 유효하다.
    return module_path


def protect_windows_secret(secret: str) -> str:
    """Windows DPAPI로 현재 사용자 계정에 귀속된 암호문을 만든다."""
    if not secret:
        return ""
    try:
        import win32crypt
    except ImportError as exc:
        raise RuntimeError("앱 비밀번호 암호화에 필요한 pywin32를 찾지 못했습니다.") from exc
    protected_result = win32crypt.CryptProtectData(
        secret.encode("utf-8"),
        "NMIS Naver Mail App Password",
        None,
        None,
        None,
        0,
    )
    encrypted = protected_result[1] if isinstance(protected_result, tuple) else protected_result
    if not isinstance(encrypted, (bytes, bytearray)):
        raise RuntimeError("Windows 암호화 결과 형식을 해석하지 못했습니다.")
    return DPAPI_PREFIX + base64.b64encode(encrypted).decode("ascii")


def unprotect_windows_secret(protected: str) -> str:
    """현재 Windows 사용자 계정으로 DPAPI 암호문을 복호화한다."""
    if not protected or not protected.startswith(DPAPI_PREFIX):
        return ""
    try:
        import win32crypt

        encrypted = base64.b64decode(protected[len(DPAPI_PREFIX) :], validate=True)
        unprotected_result = win32crypt.CryptUnprotectData(encrypted, None, None, None, 0)
        decrypted = (
            unprotected_result[1]
            if isinstance(unprotected_result, tuple)
            else unprotected_result
        )
        if not isinstance(decrypted, (bytes, bytearray)):
            return ""
        return decrypted.decode("utf-8")
    except Exception:
        return ""


@dataclass
class ReceiptMailItem:
    uid: str
    message_id: str
    received_at: datetime
    sender_name: str
    sender_address: str
    subject: str
    attachment_name: str
    attachment_path: Path
    attachment_sha256: str
    body_preview: str = ""
    status: str = "대기"
    receipt_no: str = ""
    output_path: Path | None = None
    metadata: ReceiptDocumentMetadata | None = None

    @property
    def stable_id(self) -> str:
        raw = f"{self.message_id}|{self.attachment_sha256}|{self.attachment_name}"
        return hashlib.sha256(raw.encode("utf-8", errors="replace")).hexdigest()

    @property
    def sender_display(self) -> str:
        if self.sender_name and self.sender_address:
            return f"{self.sender_name} <{self.sender_address}>"
        return self.sender_name or self.sender_address

    @property
    def received_date_raw(self) -> str:
        return self.received_at.strftime("%Y%m%d")


def sort_receipt_items_oldest_first(items: Iterable[ReceiptMailItem]) -> list[ReceiptMailItem]:
    """체크 순서와 무관하게 수신일시가 가장 이른 문서부터 정렬한다."""
    return sorted(
        items,
        key=lambda item: (
            item.received_at.timestamp(),
            int(item.uid) if str(item.uid).isdigit() else float("inf"),
            str(item.uid),
            item.attachment_name,
        ),
    )


@dataclass(frozen=True)
class ReceiptDocumentMetadata:
    classification_no: str
    document_date: date
    document_title: str
    sender_name: str
    receiver_name: str
    in_out_label: str

    @property
    def document_date_raw(self) -> str:
        return self.document_date.strftime("%Y%m%d")


class ReceiptHistory:
    def __init__(self, db_path: str | Path) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        with closing(self._connect()) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS receipt_history (
                    stable_id TEXT PRIMARY KEY,
                    message_id TEXT NOT NULL,
                    attachment_sha256 TEXT NOT NULL,
                    receipt_no TEXT NOT NULL,
                    source_path TEXT NOT NULL,
                    output_path TEXT NOT NULL,
                    processed_at TEXT NOT NULL
                )
                """
            )
            conn.commit()

    def _connect(self) -> sqlite3.Connection:
        return sqlite3.connect(self.db_path)

    def find(self, stable_id: str) -> dict[str, str] | None:
        with closing(self._connect()) as conn:
            row = conn.execute(
                """
                SELECT receipt_no, source_path, output_path, processed_at
                FROM receipt_history WHERE stable_id = ?
                """,
                (stable_id,),
            ).fetchone()
        if not row:
            return None
        return {
            "receipt_no": row[0],
            "source_path": row[1],
            "output_path": row[2],
            "processed_at": row[3],
        }

    def record(self, item: ReceiptMailItem, receipt_no: str, output_path: Path) -> None:
        with closing(self._connect()) as conn:
            conn.execute(
                """
                INSERT OR REPLACE INTO receipt_history (
                    stable_id, message_id, attachment_sha256, receipt_no,
                    source_path, output_path, processed_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    item.stable_id,
                    item.message_id,
                    item.attachment_sha256,
                    receipt_no,
                    str(item.attachment_path),
                    str(output_path),
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )
            conn.commit()


def parse_iso_date(value: str) -> date:
    text = value.strip().replace(".", "-")
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError("날짜는 YYYY-MM-DD 형식으로 입력하세요.") from exc


def decode_mail_header(value: str | None) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value))).strip()
    except Exception:
        return value.strip()


def sanitize_filename(value: str, fallback: str = "attachment.hwp") -> str:
    name = Path(value or fallback).name
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return name or fallback


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _unique_attachment_path(
    download_dir: Path,
    received_at: datetime,
    uid: str,
    filename: str,
    data: bytes,
) -> Path:
    download_dir.mkdir(parents=True, exist_ok=True)
    safe_name = sanitize_filename(filename)
    safe_path = Path(safe_name)
    base_name = f"{safe_path.stem}_{received_at:%Y%m%d}{safe_path.suffix}"
    candidate = download_dir / base_name
    digest = sha256_bytes(data)
    if candidate.exists() and hashlib.sha256(candidate.read_bytes()).hexdigest() == digest:
        return candidate
    counter = 2
    while candidate.exists():
        candidate = download_dir / f"{Path(base_name).stem}_{counter}{Path(base_name).suffix}"
        counter += 1
    candidate.write_bytes(data)
    return candidate


def _plain_body_preview(message, limit: int = 300) -> str:
    try:
        body = message.get_body(preferencelist=("plain",))
        text = body.get_content() if body else ""
    except Exception:
        text = ""
    return " ".join(str(text).split())[:limit]


def _imap_date(value: date) -> str:
    return value.strftime("%d-%b-%Y")


def search_naver_hwp_mail(
    username: str,
    app_password: str,
    start_date: date,
    end_date: date,
    sender_filter: str,
    download_dir: str | Path,
    mailbox: str = "INBOX",
    log_cb: Callable[[str], None] | None = None,
) -> list[ReceiptMailItem]:
    """네이버 메일을 읽기 전용으로 검색하고 접수 가능한 HWP/HWPX/PDF를 내려받는다."""
    if start_date > end_date:
        raise ValueError("시작일은 종료일보다 늦을 수 없습니다.")
    username = username.strip()
    if not username:
        raise ValueError("네이버 메일 아이디를 입력하세요.")
    if "@" not in username:
        username = f"{username}@naver.com"
    if not app_password:
        raise ValueError("네이버 애플리케이션 비밀번호를 입력하세요.")

    log = log_cb or (lambda _msg: None)
    destination = Path(download_dir).expanduser().resolve()
    results: list[ReceiptMailItem] = []
    sender_match_count = 0
    attachment_count = 0
    attachment_extensions: dict[str, int] = {}
    seen_attachment_names: set[tuple[str, str]] = set()
    duplicate_attachment_count = 0
    client: imaplib.IMAP4_SSL | None = None
    try:
        log("네이버 메일에 안전하게 연결하는 중...")
        client = imaplib.IMAP4_SSL(NAVER_IMAP_HOST, NAVER_IMAP_PORT, timeout=30)
        client.login(username, app_password)
        status, _ = client.select(mailbox, readonly=True)
        if status != "OK":
            raise RuntimeError(f"메일함을 열 수 없습니다: {mailbox}")

        criteria = (
            "SINCE",
            _imap_date(start_date),
            "BEFORE",
            _imap_date(end_date + timedelta(days=1)),
        )
        status, data = client.uid("search", None, *criteria)
        if status != "OK":
            raise RuntimeError("기간 조건으로 메일을 검색하지 못했습니다.")

        uids = (data[0] or b"").split()
        sender_scope = f"발신자 '{sender_filter.strip()}'" if sender_filter.strip() else "전체 발신자"
        log(
            f"기간 내 메일 {len(uids)}건에서 {sender_scope}의 "
            "HWP/HWPX/PDF 첨부파일을 확인합니다."
        )
        sender_needle = sender_filter.strip().casefold()
        for uid_bytes in reversed(uids):
            uid = uid_bytes.decode("ascii", errors="ignore")
            status, fetched = client.uid("fetch", uid, "(BODY.PEEK[])")
            if status != "OK":
                log(f"메일 UID {uid} 읽기 실패 - 건너뜀")
                continue
            raw = next(
                (part[1] for part in fetched if isinstance(part, tuple) and len(part) > 1),
                None,
            )
            if not raw:
                continue

            message = BytesParser(policy=policy.default).parsebytes(raw)
            sender_header = decode_mail_header(message.get("From"))
            sender_name, sender_address = parseaddr(sender_header)
            sender_name = decode_mail_header(sender_name)
            sender_haystack = f"{sender_header} {sender_name} {sender_address}".casefold()
            if sender_needle and sender_needle not in sender_haystack:
                continue
            sender_match_count += 1

            try:
                received_at = parsedate_to_datetime(message.get("Date"))
                if received_at is None:
                    raise ValueError
                if received_at.tzinfo:
                    received_at = received_at.astimezone().replace(tzinfo=None)
            except Exception:
                received_at = datetime.combine(start_date, datetime.min.time())
            if not start_date <= received_at.date() <= end_date:
                continue

            subject = decode_mail_header(message.get("Subject")) or "(제목 없음)"
            message_id = (message.get("Message-ID") or f"naver-uid-{uid}").strip()
            body_preview = _plain_body_preview(message)
            for part in message.iter_attachments():
                filename = decode_mail_header(part.get_filename())
                suffix = Path(filename).suffix.lower()
                attachment_count += 1
                extension_label = suffix or "(확장자 없음)"
                attachment_extensions[extension_label] = attachment_extensions.get(extension_label, 0) + 1
                if suffix not in RECEIPT_ATTACHMENT_SUFFIXES:
                    continue
                payload = part.get_payload(decode=True) or b""
                if not payload:
                    continue
                duplicate_key = (
                    sender_address.strip().casefold(),
                    sanitize_filename(filename).casefold(),
                )
                if duplicate_key in seen_attachment_names:
                    duplicate_attachment_count += 1
                    continue
                seen_attachment_names.add(duplicate_key)
                saved_path = _unique_attachment_path(
                    destination, received_at, uid, filename, payload
                )
                results.append(
                    ReceiptMailItem(
                        uid=uid,
                        message_id=message_id,
                        received_at=received_at,
                        sender_name=sender_name,
                        sender_address=sender_address,
                        subject=subject,
                        attachment_name=filename,
                        attachment_path=saved_path,
                        attachment_sha256=sha256_bytes(payload),
                        body_preview=body_preview,
                    )
                )
        extension_summary = ", ".join(
            f"{extension} {count}개" for extension, count in sorted(attachment_extensions.items())
        ) or "첨부파일 없음"
        log(
            f"발신자 일치 메일 {sender_match_count}건 / 첨부파일 {attachment_count}개 "
            f"({extension_summary})"
        )
        log(f"조건에 맞는 HWP/HWPX/PDF 첨부파일 {len(results)}건을 찾았습니다.")
        if duplicate_attachment_count:
            log(
                "♻️ 같은 발신자가 같은 이름으로 재발송한 첨부파일 "
                f"{duplicate_attachment_count}건은 제외하고 최신 파일만 남겼습니다."
            )
        if sender_needle and sender_match_count == 0:
            log("⚠️ 발신자 검색어가 실제 보낸사람 이름 또는 이메일 주소와 일치하지 않습니다.")
        return results
    except imaplib.IMAP4.error as exc:
        raw_reason = " ".join(str(part) for part in getattr(exc, "args", ()) if part)
        server_reason = raw_reason.replace("b'", "").rstrip("'") or "인증 거절"
        raise RuntimeError(
            "네이버가 IMAP 로그인을 거절했습니다.\n"
            "1) 네이버 메일 > 환경설정 > POP3/IMAP 설정에서 IMAP/SMTP를 '사용함'으로 변경하고,\n"
            "2) 네이버ID 보안설정에서 2단계 인증을 켠 뒤 '애플리케이션 비밀번호'를 새로 만든 다음,\n"
            "3) 이 프로그램에는 일반 로그인 비밀번호가 아닌 그 애플리케이션 비밀번호를 입력하세요.\n"
            f"네이버 서버 응답: {server_reason}"
        ) from exc
    except (OSError, TimeoutError) as exc:
        raise RuntimeError(
            f"네이버 IMAP 서버({NAVER_IMAP_HOST}:{NAVER_IMAP_PORT})에 연결하지 못했습니다. "
            f"인터넷 또는 방화벽 상태를 확인하세요. 상세 오류: {exc}"
        ) from exc
    finally:
        if client is not None:
            try:
                client.logout()
            except Exception:
                pass


def extract_document_number(subject: str, body: str = "") -> str:
    text = f"{subject} {body}"
    patterns = (
        r"(?:문서번호|시행번호)\s*[:：]?\s*([가-힣A-Za-z]+\s*[-–]?\s*\d+(?:-\d+)*)",
        r"\b((?:총무|정경|회원|사업|위생)\s*[-–]?\s*\d+(?:-\d+)*)\b",
    )
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return re.sub(r"\s+", " ", match.group(1)).strip()
    return ""


def parse_receipt_document_metadata(
    text: str,
    attachment_name: str,
    sender_address: str = "",
    sender_display_name: str = "",
) -> ReceiptDocumentMetadata:
    """공문 본문의 '시행 : 분류번호 (YYYY. M. D.)'를 접수대장 값으로 변환한다."""
    normalized = str(text or "").replace("\xa0", " ")
    match = re.search(
        r"시\s*행\s*[:：]?\s*([^()\r\n]+?)\s*"
        r"\(\s*(\d{4})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(\d{1,2})\s*[.]?\s*\)",
        normalized,
        re.IGNORECASE,
    )
    if not match:
        raise ValueError(
            "첨부문서에서 '시행 : 분류번호 (YYYY. M. D.)' 형식을 찾지 못했습니다."
        )
    classification_no = re.sub(r"\s+", "", match.group(1)).strip()
    classification_no = re.sub(r"호$", "", classification_no).strip()
    if not classification_no:
        raise ValueError("첨부문서 시행란의 분류번호가 비어 있습니다.")
    try:
        document_date = date(int(match.group(2)), int(match.group(3)), int(match.group(4)))
    except ValueError as exc:
        raise ValueError("첨부문서 시행란의 문서일자가 올바르지 않습니다.") from exc

    document_title = Path(attachment_name).stem.strip()
    document_title = re.sub(r"_\d{8}(?:_\d+)?$", "", document_title).strip()
    if not document_title:
        raise ValueError("첨부파일에서 문서명을 만들 수 없습니다.")

    sender_address_lower = sender_address.strip().lower()
    is_jeonbuk_branch = (
        "전북특별자치도지회" in normalized
        or "전북특별자치도지회" in normalized.replace(" ", "")
        or sender_address_lower == "krbajb5980@daum.net"
    )
    fallback_sender = sender_display_name.strip() or sender_address.strip() or "발신자 미상"
    sender_name = "전북특별자치도지회" if is_jeonbuk_branch else fallback_sender
    in_out_label = "내부" if sender_name == "전북특별자치도지회" else "외부"

    return ReceiptDocumentMetadata(
        classification_no=classification_no,
        document_date=document_date,
        document_title=document_title,
        sender_name=sender_name,
        receiver_name="진안군지부",
        in_out_label=in_out_label,
    )


def parse_hwp_receipt_metadata(
    text: str,
    attachment_name: str,
    sender_address: str = "",
    sender_display_name: str = "",
) -> ReceiptDocumentMetadata:
    """이전 호출부와의 호환성을 유지하는 공문 시행정보 파서."""
    return parse_receipt_document_metadata(
        text,
        attachment_name,
        sender_address=sender_address,
        sender_display_name=sender_display_name,
    )


def has_receipt_slot(text: str) -> bool:
    """문서에 비어 있거나 작성된 '접수 :' 칸이 있는지 확인한다."""
    return bool(re.search(r"접\s*수\s*[:：]", str(text or "")))


def require_receipt_slot(text: str) -> None:
    if not has_receipt_slot(text):
        raise ValueError("첨부문서에 '접수 :' 칸이 없어 접수대장 등록 대상이 아닙니다.")


def _extract_binary_hwp_text(source: Path) -> str:
    """한글을 실행하지 않고 HWP OLE BodyText의 문단 문자열을 빠르게 읽는다."""
    try:
        import olefile
    except ImportError as exc:
        raise RuntimeError("HWP 빠른 검사 모듈(olefile)이 설치되지 않았습니다.") from exc

    chunks: list[str] = []
    try:
        with olefile.OleFileIO(str(source)) as ole:
            header = ole.openstream("FileHeader").read()
            if len(header) < 40:
                raise ValueError("HWP 파일 헤더가 올바르지 않습니다.")
            properties = struct.unpack_from("<I", header, 36)[0]
            compressed = bool(properties & 0x01)
            section_paths = sorted(
                (
                    path
                    for path in ole.listdir()
                    if len(path) == 2
                    and path[0] == "BodyText"
                    and path[1].startswith("Section")
                ),
                key=lambda path: int(re.sub(r"\D", "", path[1]) or 0),
            )
            for path in section_paths:
                data = ole.openstream(path).read()
                if compressed:
                    data = zlib.decompress(data, -15)
                offset = 0
                while offset + 4 <= len(data):
                    record_header = struct.unpack_from("<I", data, offset)[0]
                    offset += 4
                    tag_id = record_header & 0x3FF
                    size = (record_header >> 20) & 0xFFF
                    if size == 0xFFF:
                        if offset + 4 > len(data):
                            break
                        size = struct.unpack_from("<I", data, offset)[0]
                        offset += 4
                    payload = data[offset : offset + size]
                    offset += size
                    if tag_id == 67:
                        chunks.append(payload.decode("utf-16le", errors="ignore"))
    except Exception as exc:
        raise RuntimeError(f"HWP 본문 빠른 검사 실패: {source.name} ({exc})") from exc
    return "\n".join(chunks)


def _extract_hwpx_text(source: Path) -> str:
    chunks: list[str] = []
    try:
        with zipfile.ZipFile(source) as archive:
            section_names = sorted(
                name
                for name in archive.namelist()
                if re.fullmatch(r"Contents/section\d+\.xml", name, re.IGNORECASE)
            )
            for name in section_names:
                root = ElementTree.fromstring(archive.read(name))
                chunks.extend(text for text in root.itertext() if text)
    except Exception as exc:
        raise RuntimeError(f"HWPX 본문 빠른 검사 실패: {source.name} ({exc})") from exc
    return "\n".join(chunks)


def _extract_pdf_text(source: Path) -> str:
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise RuntimeError("PDF 읽기 모듈(pypdf)이 설치되지 않았습니다.") from exc
    try:
        reader = PdfReader(str(source))
        return "\n".join(str(page.extract_text() or "") for page in reader.pages)
    except Exception as exc:
        raise RuntimeError(f"PDF 본문 빠른 검사 실패: {source.name} ({exc})") from exc


def extract_attachment_text_fast(source_path: str | Path) -> str:
    source = Path(source_path).expanduser().resolve()
    suffix = source.suffix.lower()
    if suffix == ".hwp":
        return _extract_binary_hwp_text(source)
    if suffix == ".hwpx":
        return _extract_hwpx_text(source)
    if suffix in PDF_SUFFIXES:
        return _extract_pdf_text(source)
    raise ValueError("HWP, HWPX 또는 PDF 첨부파일만 검사할 수 있습니다.")


def filter_receipt_candidate_documents(
    items: Iterable[ReceiptMailItem],
    log_cb: Callable[[str], None] | None = None,
) -> list[ReceiptMailItem]:
    """메일 조회 직후 시행정보와 접수칸이 모두 있는 문서만 남긴다."""
    log = log_cb or (lambda _msg: None)
    source_items = list(items)
    candidates: list[ReceiptMailItem] = []
    excluded_no_receipt = 0
    excluded_invalid_enforcement = 0
    excluded_unreadable = 0
    log(f"접수 가능 여부를 빠르게 검사합니다: 첨부문서 {len(source_items)}건")
    for item in source_items:
        try:
            text = extract_attachment_text_fast(item.attachment_path)
            require_receipt_slot(text)
            item.metadata = parse_receipt_document_metadata(
                text,
                item.attachment_path.name,
                sender_address=item.sender_address,
                sender_display_name=item.sender_name,
            )
            item.status = "접수 가능"
            candidates.append(item)
        except ValueError as exc:
            if "접수" in str(exc):
                excluded_no_receipt += 1
            else:
                excluded_invalid_enforcement += 1
        except Exception:
            excluded_unreadable += 1
    log(
        f"접수 가능 문서 {len(candidates)}건 / 제외: 접수칸 없음 {excluded_no_receipt}건, "
        f"시행정보 형식 불일치 {excluded_invalid_enforcement}건, 본문 읽기 실패 {excluded_unreadable}건"
    )
    return candidates


def format_hwp_receipt_entry(receipt_no: str, document_date: date) -> str:
    receipt_no = str(receipt_no).strip()
    if not receipt_no:
        raise ValueError("접수번호가 비어 있습니다.")
    return f"{receipt_no} ({document_date:%Y.%m.%d})"


def extract_hwp_receipt_metadata(
    source_path: str | Path,
    sender_address: str = "",
    sender_display_name: str = "",
    log_cb: Callable[[str], None] | None = None,
) -> ReceiptDocumentMetadata:
    """한글 Automation으로 HWP의 시행번호와 문서일자를 읽는다."""
    source = Path(source_path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"HWP 파일을 찾을 수 없습니다: {source}")
    log = log_cb or (lambda _msg: None)
    hwp = None
    pythoncom = None
    try:
        import pythoncom as _pythoncom
        import win32com.client

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
        # EnsureDispatch는 직전 문서에서 Quit 중인 한글 인스턴스를 재사용할 수 있어
        # 두 번째 문서부터 RPC 서버 예외가 발생한다. 문서별 독립 인스턴스를 사용한다.
        hwp = win32com.client.DispatchEx("HWPFrame.HwpObject")
        ensure_hancom_security_module_registered(hwp)
        try:
            hwp.XHwpWindows.Item(0).Visible = False
        except Exception:
            pass
        if not hwp.Open(str(source), "", "forceopen:true"):
            raise RuntimeError("한글에서 HWP 파일을 열지 못했습니다.")
        text = str(hwp.GetTextFile("TEXT", "") or "")
        require_receipt_slot(text)
        metadata = parse_hwp_receipt_metadata(
            text,
            source.name,
            sender_address=sender_address,
            sender_display_name=sender_display_name,
        )
        log(
            "HWP 시행정보 확인: "
            f"분류번호={metadata.classification_no}, "
            f"문서일자={metadata.document_date:%Y.%m.%d}, "
            f"내외구분={metadata.in_out_label}"
        )
        return metadata
    finally:
        if hwp is not None:
            try:
                hwp.Clear(1)
            except Exception:
                pass
            try:
                hwp.Quit()
            except Exception:
                pass
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def extract_pdf_receipt_metadata(
    source_path: str | Path,
    sender_address: str = "",
    sender_display_name: str = "",
    log_cb: Callable[[str], None] | None = None,
) -> ReceiptDocumentMetadata:
    """텍스트 기반 PDF에서 시행번호와 문서일자를 읽는다."""
    source = Path(source_path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {source}")
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise RuntimeError(
            "PDF 읽기 모듈(pypdf)이 설치되지 않았습니다. "
            "프로그램 폴더에서 'python -m pip install -r requirements.txt'를 실행하세요."
        ) from exc

    try:
        reader = PdfReader(str(source))
        text = "\n".join(str(page.extract_text() or "") for page in reader.pages)
    except Exception as exc:
        raise RuntimeError(f"PDF 본문을 읽지 못했습니다: {source.name} ({exc})") from exc
    if not text.strip():
        raise ValueError(
            "PDF에서 글자를 읽지 못했습니다. 스캔 이미지 PDF는 OCR 처리 후 다시 시도하세요."
        )
    require_receipt_slot(text)

    metadata = parse_receipt_document_metadata(
        text,
        source.name,
        sender_address=sender_address,
        sender_display_name=sender_display_name,
    )
    log = log_cb or (lambda _msg: None)
    log(
        "PDF 시행정보 확인: "
        f"분류번호={metadata.classification_no}, "
        f"문서일자={metadata.document_date:%Y.%m.%d}, "
        f"내외구분={metadata.in_out_label}"
    )
    return metadata


def extract_receipt_document_metadata(
    source_path: str | Path,
    sender_address: str = "",
    sender_display_name: str = "",
    log_cb: Callable[[str], None] | None = None,
) -> ReceiptDocumentMetadata:
    """첨부파일 형식에 맞는 시행정보 추출기를 선택한다."""
    suffix = Path(source_path).suffix.lower()
    kwargs = {
        "sender_address": sender_address,
        "sender_display_name": sender_display_name,
        "log_cb": log_cb,
    }
    if suffix in HWP_SUFFIXES:
        return extract_hwp_receipt_metadata(source_path, **kwargs)
    if suffix in PDF_SUFFIXES:
        return extract_pdf_receipt_metadata(source_path, **kwargs)
    raise ValueError("HWP, HWPX 또는 PDF 첨부파일만 처리할 수 있습니다.")


def _find_receipt_route(page: Page) -> str:
    route = page.evaluate(
        """() => {
            const injector = window.angular && angular.element(document.body).injector();
            if (!injector) return {error: 'NMIS Angular 화면을 찾지 못했습니다.'};
            const state = injector.get('$state');
            const names = state.get().map(x => x && x.name).filter(Boolean);
            const preferred = [
                'admin/document/receive/list',
                'admin/document/receipt/list',
                'admin/document/reception/list',
                'admin/document/accept/list'
            ];
            const exact = preferred.find(x => names.includes(x));
            if (exact) return {route: exact};
            const states = state.get();
            const candidateState = states.find(x => {
                const name = (x && x.name) || '';
                const serialized = JSON.stringify(x || {});
                return /admin[\\/.]document/i.test(name) &&
                    (/(receive|receipt|reception|accept|arrival)/i.test(name) || serialized.includes('접수'));
            });
            return candidateState ? {route: candidateState.name} : {error: '접수대장 화면 경로를 찾지 못했습니다.', names};
        }"""
    )
    if route.get("error"):
        raise RuntimeError(route["error"])
    return str(route["route"])


def _read_existing_receipt(page: Page, match_data: dict) -> dict | None:
    return page.evaluate(
        """(matchData) => {
            const el = document.querySelector('epro-grid') || document.querySelector('div[ui-grid]') || document.querySelector('.ui-grid');
            if (!el || !window.angular) return null;
            const scopes = [angular.element(el).scope(), angular.element(el).isolateScope()].filter(Boolean);
            const unwrap = row => row && (row.entity || row);
            const firstValue = (obj, keys) => {
                for (const key of keys) {
                    const value = obj && obj[key];
                    if (value !== undefined && value !== null && String(value).trim()) return String(value).trim();
                }
                return '';
            };
            for (const sc of scopes) {
                for (const key of Object.keys(sc)) {
                    if (!/^gridData/i.test(key) || !Array.isArray(sc[key])) continue;
                    for (const wrapped of sc[key]) {
                        const row = unwrap(wrapped);
                        const title = firstValue(row, ['contents', 'subject', 'title', 'documentName', 'documentTitle']);
                        const sender = firstValue(row, ['senderName', 'sendName', 'sender', 'senderOrgName', 'organizationName', 'dispatchName']);
                        const rowDate = firstValue(row, ['receiptDate', 'receiveDate', 'receivedDate', 'acceptDate', 'regDate']).replace(/[^0-9]/g, '');
                        if (title === matchData.title && (!matchData.sender || sender.includes(matchData.sender)) && (!matchData.date || rowDate === matchData.date)) {
                            const explicitKeys = ['receiptNo', 'receiptNumber', 'registrationNo', 'acceptNo', 'receiveNo', 'receivedNo', 'documentReceiptNo'];
                            const dynamicKeys = Object.keys(row).filter(x =>
                                /(receipt|receive|received|accept|arrival).*(no|num|number|seq)|(no|num|number|seq).*(receipt|receive|received|accept|arrival)/i.test(x)
                            );
                            const optionCandidates = Object.keys(sc)
                                .filter(x => /grid.*option/i.test(x) && sc[x] && Array.isArray(sc[x].columnDefs))
                                .flatMap(x => sc[x].columnDefs)
                                .filter(col => /접수.*번호|번호.*접수/.test(String((col && (col.displayName || col.name || col.headerName)) || '')))
                                .map(col => col && (col.field || col.name))
                                .filter(Boolean);
                            const receiptNo = firstValue(row, [...explicitKeys, ...dynamicKeys, ...optionCandidates]);
                            return {receiptNo, title, sender, rowDate, keys: Object.keys(row)};
                        }
                    }
                }
            }
            return null;
        }""",
        match_data,
    )


def register_receipt_document_on_nmis(
    page: Page,
    item: ReceiptMailItem,
    metadata: ReceiptDocumentMetadata | None = None,
    save: bool = True,
    log_cb: Callable[[str], None] | None = None,
) -> dict:
    """접수대장에 한 행을 등록하고 NMIS가 발급한 접수번호를 반환한다."""
    log = log_cb or (lambda _msg: None)
    if metadata is None:
        raise ValueError("NMIS 등록 전에 첨부문서 시행정보를 먼저 읽어야 합니다.")
    route = _find_receipt_route(page)
    log(f"NMIS 접수대장 화면으로 이동합니다: {route}")
    page.evaluate(
        """route => {
            const injector = angular.element(document.body).injector();
            injector.get('$state').go(route);
        }""",
        route,
    )
    page.wait_for_timeout(1600)

    receipt_date_raw = item.received_date_raw
    match_data = {
        "title": metadata.document_title,
        "sender": metadata.sender_name,
        "date": receipt_date_raw,
    }
    existing = _read_existing_receipt(page, match_data)
    if existing:
        if existing.get("receiptNo"):
            log(f"이미 등록된 문서를 찾았습니다. 접수번호: {existing['receiptNo']}")
            return {"success": True, "duplicate": True, "receipt_no": existing["receiptNo"], "route": route}
        raise RuntimeError("같은 제목·발신자·접수일의 문서가 이미 있으나 접수번호를 읽지 못했습니다.")

    payload = {
        "receiptDate": receipt_date_raw,
        "inOutLabel": metadata.in_out_label,
        "classificationNo": metadata.classification_no,
        "documentDate": metadata.document_date_raw,
        "title": metadata.document_title,
        "sender": metadata.sender_name,
        "receiver": metadata.receiver_name,
        "senderAddress": item.sender_address,
        "attachmentName": item.attachment_name,
    }
    inserted = page.evaluate(
        """data => {
            const el = document.querySelector('epro-grid') || document.querySelector('div[ui-grid]') || document.querySelector('.ui-grid');
            if (!el || !window.angular) return {error: '접수대장 그리드를 찾지 못했습니다.'};
            const scopes = [angular.element(el).scope(), angular.element(el).isolateScope()].filter(Boolean);
            const sc = scopes.find(x => typeof x.fnInsertRow === 'function') || scopes[0];
            if (!sc) return {error: '접수대장 그리드 스코프를 찾지 못했습니다.'};
            const before = {};
            for (const key of Object.keys(sc)) if (/^gridData/i.test(key) && Array.isArray(sc[key])) before[key] = sc[key].length;
            if (typeof sc.fnInsertRow !== 'function') return {error: '접수대장 행추가 기능을 찾지 못했습니다.'};
            sc.fnInsertRow();
            const arrays = Object.keys(sc).filter(key => /^gridData/i.test(key) && Array.isArray(sc[key]));
            const dataKey = arrays.find(key => sc[key].length > (before[key] || 0)) || arrays.find(key => sc[key].length) || arrays[0];
            if (!dataKey || !sc[dataKey].length) return {error: '추가된 접수대장 행을 찾지 못했습니다.'};
            const wrapped = sc[dataKey][sc[dataKey].length - 1];
            const row = wrapped.entity || wrapped;
            const keys = new Set([...Object.keys(wrapped || {}), ...Object.keys(row || {})]);
            const setAlias = (aliases, value, fallback) => {
                const key = aliases.find(x => keys.has(x)) || fallback;
                wrapped[key] = value;
                row[key] = value;
                return key;
            };
            const existingRows = sc[dataKey].slice(0, -1).map(x => x && (x.entity || x)).filter(Boolean);
            const firstExistingValue = key => {
                for (const existingRow of existingRows) {
                    const value = existingRow[key];
                    if (value !== undefined && value !== null && String(value).trim()) return value;
                }
                return '';
            };
            wrapped.selected = sc.getProperty ? sc.getProperty('TRUE') : 'TRUE';
            row.selected = wrapped.selected;
            const inOutAliases = ['inOutType', 'inoutType', 'inOutDiv', 'inoutDiv', 'insideOutside', 'documentInOutType', 'internalExternalType'];
            const inOutKey = inOutAliases.find(x => keys.has(x)) || 'inOutType';
            const inOutValue = data.inOutLabel === '내부' ? (firstExistingValue(inOutKey) || 'I') : 'O';
            const mapped = {
                receiptDate: setAlias(['receiptDate', 'receiveDate', 'receivedDate', 'acceptDate', 'regDate'], data.receiptDate, 'receiptDate'),
                inOut: setAlias(inOutAliases, inOutValue, inOutKey),
                classificationNo: setAlias(['classificationNo', 'classNo', 'documentClassNo', 'documentNo', 'senderDocumentNo'], data.classificationNo, 'documentNo'),
                documentDate: setAlias(['documentDate', 'enforcementDate', 'dispatchDate', 'sendDate'], data.documentDate, 'documentDate'),
                title: setAlias(['contents', 'subject', 'title', 'documentName', 'documentTitle'], data.title, 'contents'),
                sender: setAlias(['senderName', 'sendName', 'sender', 'senderOrgName', 'organizationName', 'dispatchName'], data.sender, 'senderName'),
                receiver: setAlias(['receiverName', 'receiveName', 'receiver', 'recipientName', 'recipient'], data.receiver, 'receiverName')
            };
            const inOutNameKey = ['inOutTypeName', 'inoutTypeName', 'insideOutsideName'].find(x => keys.has(x));
            if (inOutNameKey) { wrapped[inOutNameKey] = data.inOutLabel; row[inOutNameKey] = data.inOutLabel; mapped.inOutName = inOutNameKey; }
            const fileKey = ['fileName', 'attachmentName', 'documentFileName'].find(x => keys.has(x));
            if (fileKey) { wrapped[fileKey] = data.attachmentName; row[fileKey] = data.attachmentName; mapped.file = fileKey; }
            if (typeof sc.$apply === 'function' && !sc.$$phase) sc.$apply();
            return {success: true, dataKey, mapped, keys: Array.from(keys)};
        }""",
        payload,
    )
    if inserted.get("error"):
        raise RuntimeError(inserted["error"])
    page.wait_for_timeout(250)
    try:
        last_row = page.locator(".ui-grid-row:visible").last
        in_out_select = last_row.locator("select:has(option:has-text('내부')):visible")
        if in_out_select.count() > 0:
            in_out_select.first.select_option(label=metadata.in_out_label)
    except Exception:
        pass
    log(
        "접수대장 행 입력 완료: "
        f"접수일자={item.received_at:%Y.%m.%d}, "
        f"내외구분={metadata.in_out_label}, "
        f"분류번호={metadata.classification_no}, "
        f"문서일자={metadata.document_date:%Y.%m.%d}, "
        f"문서명={metadata.document_title}, "
        f"발신자={metadata.sender_name}, 수신자={metadata.receiver_name}"
    )
    if not save:
        return {"success": True, "saved": False, "route": route, "diagnostics": inserted}

    save_button = page.locator(
        "button[ng-click*='fnSave']:visible, "
        "button:has(span.button_icon[lang-code='save']):visible, "
        "button:has(span[lang-code='save']):visible"
    )
    if save_button.count() > 0:
        save_button.first.click(force=True)
    else:
        invoked = page.evaluate(
            """() => {
                const el = document.querySelector('epro-grid') || document.querySelector('div[ui-grid]') || document.querySelector('.ui-grid');
                const sc = el && angular.element(el).scope();
                for (const name of ['fnSave', 'fnSaveData', 'fnSaveList']) {
                    if (sc && typeof sc[name] === 'function') { sc[name](); return name; }
                }
                return '';
            }"""
        )
        if not invoked:
            raise RuntimeError("접수대장 저장 버튼을 찾지 못했습니다. 행은 화면에만 입력된 상태입니다.")
    page.wait_for_timeout(500)

    for _ in range(3):
        confirm = page.locator(
            "button:has-text('확인'):visible, button:has-text('예'):visible, "
            "button.btn-success:visible, input[value='확인']:visible"
        )
        if confirm.count() == 0:
            break
        confirm.first.click(force=True)
        page.wait_for_timeout(650)

    deadline = datetime.now().timestamp() + 8
    saved: dict | None = None
    while datetime.now().timestamp() < deadline:
        saved = _read_existing_receipt(page, match_data)
        if saved and saved.get("receiptNo"):
            break
        page.wait_for_timeout(400)
    if not saved or not saved.get("receiptNo"):
        known_keys = ", ".join(inserted.get("keys", []))
        raise RuntimeError(
            "접수대장은 저장했지만 생성된 접수번호를 화면 모델에서 읽지 못했습니다. "
            f"첫 실행 시 개발자에게 아래 필드 목록을 전달해 주세요: {known_keys}"
        )
    receipt_no = str(saved["receiptNo"])
    log(f"NMIS 접수 완료 - 접수번호 {receipt_no}")
    return {"success": True, "saved": True, "duplicate": False, "receipt_no": receipt_no, "route": route}


def _all_replace_hwp(hwp, find_text: str, replace_text: str) -> bool:
    try:
        action = hwp.HAction
        params = hwp.HParameterSet.HFindReplace
        action.GetDefault("AllReplace", params.HSet)
        params.FindString = find_text
        params.ReplaceString = replace_text
        params.IgnoreMessage = 1
        params.Direction = hwp.FindDir("AllDoc")
        params.FindType = 1
        return bool(action.Execute("AllReplace", params.HSet))
    except Exception:
        return False


def _find_text_hwp(hwp, find_text: str) -> bool:
    try:
        try:
            # 저장 당시 커서가 뒤쪽 페이지에 있어도 항상 문서 첫 위치부터 찾는다.
            moved_to_begin = bool(hwp.HAction.Run("MoveDocBegin"))
            if not moved_to_begin:
                hwp.HAction.Run("MoveTopLevelBegin")
        except Exception:
            hwp.HAction.Run("MoveTopLevelBegin")
        action = hwp.HAction
        params = hwp.HParameterSet.HFindReplace
        action.GetDefault("RepeatFind", params.HSet)
        params.FindString = find_text
        params.IgnoreMessage = 1
        params.Direction = hwp.FindDir("Forward")
        params.FindType = 1
        params.MatchCase = 0
        params.WholeWordOnly = 0
        params.UseWildCards = 0
        return bool(action.Execute("RepeatFind", params.HSet))
    except Exception:
        return False


def _insert_text_hwp(hwp, text: str) -> bool:
    """한글 Automation의 InsertText 액션으로 현재 위치에 문자열을 입력한다."""
    try:
        params = hwp.HParameterSet.HInsertText
        hwp.HAction.GetDefault("InsertText", params.HSet)
        params.Text = text
        return bool(hwp.HAction.Execute("InsertText", params.HSet))
    except Exception:
        return False


def _current_hwp_table_cell(hwp) -> tuple[str, int] | None:
    """현재 한글 표 셀의 열 문자와 행 번호를 반환한다."""
    try:
        indicator = hwp.KeyIndicator()
        description = str(indicator[-1]) if indicator else ""
        found = re.search(r"\(([A-Z]+)(\d+)\)", description, re.IGNORECASE)
        if found:
            return found.group(1).upper(), int(found.group(2))
    except Exception:
        pass
    return None


def _move_to_separate_receipt_value_cell(hwp) -> bool:
    """접수 라벨 오른쪽이 같은 행의 별도 입력 셀이면 그 셀로 이동한다."""
    label_cell = _current_hwp_table_cell(hwp)
    try:
        hwp.HAction.Run("Cancel")
        moved = bool(hwp.HAction.Run("TableRightCell"))
    except Exception:
        return False
    value_cell = _current_hwp_table_cell(hwp)
    return bool(
        moved
        and label_cell
        and value_cell
        and label_cell[1] == value_cell[1]
        and label_cell[0] != value_cell[0]
    )


def write_receipt_number_to_hwp(
    source_path: str | Path,
    receipt_no: str,
    document_date: date | None = None,
    output_dir: str | Path | None = None,
    visible: bool = True,
    log_cb: Callable[[str], None] | None = None,
) -> Path:
    """원본을 보존하고 복사본의 접수 필드/표 칸에 접수번호를 입력한다."""
    source = Path(source_path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"HWP 파일을 찾을 수 없습니다: {source}")
    if source.suffix.lower() not in HWP_SUFFIXES:
        raise ValueError("HWP 또는 HWPX 파일만 처리할 수 있습니다.")
    receipt_no = str(receipt_no).strip()
    if not receipt_no:
        raise ValueError("접수번호가 비어 있습니다.")
    receipt_text = format_hwp_receipt_entry(receipt_no, document_date) if document_date else receipt_no

    log = log_cb or (lambda _msg: None)
    destination_dir = Path(output_dir).expanduser().resolve() if output_dir else source.parent / "접수완료"
    destination_dir.mkdir(parents=True, exist_ok=True)
    safe_receipt = re.sub(r"[^0-9A-Za-z가-힣._-]", "_", receipt_no)
    destination = destination_dir / f"접수완료_{safe_receipt}_{source.name}"
    if destination.exists():
        try:
            source_digest = hashlib.sha256(source.read_bytes()).hexdigest()
            destination_digest = hashlib.sha256(destination.read_bytes()).hexdigest()
            if source_digest == destination_digest:
                # 이전 실패로 남은, 접수번호가 입력되지 않은 동일 복사본만 정리한다.
                destination.unlink()
        except OSError:
            # 사용자가 파일을 열어 둔 경우에는 기존 파일을 건드리지 않는다.
            pass
    counter = 2
    while destination.exists():
        destination = destination_dir / f"접수완료_{safe_receipt}_{source.stem}_{counter}{source.suffix}"
        counter += 1
    shutil.copy2(source, destination)

    hwp = None
    pythoncom = None
    completed = False
    try:
        import pythoncom as _pythoncom
        import win32com.client

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
        # 직전 문서에서 Quit 중인 인스턴스에 다시 붙지 않도록 독립 COM 인스턴스를 요청한다.
        hwp = win32com.client.DispatchEx("HWPFrame.HwpObject")
        ensure_hancom_security_module_registered(hwp)
        try:
            hwp.XHwpWindows.Item(0).Visible = bool(visible)
        except Exception:
            pass
        if not hwp.Open(str(destination), "", "forceopen:true"):
            raise RuntimeError("한글에서 HWP 파일을 열지 못했습니다.")

        changed = False
        try:
            field_list = str(hwp.GetFieldList(0, 2) or "")
        except Exception:
            field_list = ""
        field_names = [x for x in re.split(r"[\x02\r\n]+", field_list) if x]
        for field_name in ("접수번호", "접수", "receipt_no", "receipt"):
            if field_name in field_names:
                hwp.PutFieldText(field_name, receipt_text)
                changed = True
                break

        if not changed:
            for placeholder in ("{{접수번호}}", "[[접수번호]]", "<접수번호>"):
                if _all_replace_hwp(hwp, placeholder, receipt_text):
                    changed = True
                    break

        if not changed:
            receipt_labels = tuple(
                f"접{letter_gap}수{colon_gap}:"
                for letter_gap in ("", " ", "  ", "   ", "\t")
                for colon_gap in ("", " ", "  ", "\t")
            )
            for receipt_label in receipt_labels:
                if not _find_text_hwp(hwp, receipt_label):
                    continue
                try:
                    if _move_to_separate_receipt_value_cell(hwp):
                        # 2칸형: 라벨은 왼쪽 셀에 두고 오른쪽 입력 셀에 값만 쓴다.
                        changed = _insert_text_hwp(hwp, receipt_text)
                        if changed:
                            log("HWP 접수칸 형식 확인: 2칸형 - 오른쪽 입력칸에 접수정보를 입력합니다.")
                    if not changed and _find_text_hwp(hwp, receipt_label):
                        # 1칸형: 다음 표 셀이 다른 행이면 라벨과 값을 같은 칸에 둔다.
                        hwp.HAction.Run("Delete")
                        changed = _insert_text_hwp(hwp, f"{receipt_label} {receipt_text}")
                        if changed:
                            log("HWP 접수칸 형식 확인: 1칸형 - 접수 라벨 뒤에 접수정보를 입력합니다.")
                    break
                except Exception:
                    changed = False

        if not changed and _find_text_hwp(hwp, "접수"):
            try:
                if _move_to_separate_receipt_value_cell(hwp):
                    changed = _insert_text_hwp(hwp, receipt_text)
                if not changed and _find_text_hwp(hwp, "접수"):
                    hwp.HAction.Run("Delete")
                    changed = _insert_text_hwp(hwp, f"접수 : {receipt_text}")
            except Exception:
                changed = False

        if not changed:
            raise RuntimeError(
                "HWP에서 '접수번호' 누름틀, 접수번호 자리표시자 또는 '접수' 표 칸을 찾지 못했습니다."
            )
        if not hwp.Save():
            raise RuntimeError("접수번호를 입력했지만 HWP 저장에 실패했습니다.")
        completed = True
        log(f"HWP 접수번호 입력 완료: {destination.name}")
        return destination
    except Exception:
        try:
            destination.unlink(missing_ok=True)
        except Exception:
            pass
        raise
    finally:
        if hwp is not None:
            try:
                hwp.Clear(1)
            except Exception:
                pass
            try:
                hwp.Quit()
            except Exception:
                pass
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        if not completed:
            try:
                destination.unlink(missing_ok=True)
            except Exception:
                pass


def copy_registered_pdf(
    source_path: str | Path,
    receipt_no: str,
    document_date: date,
    output_dir: str | Path | None = None,
    log_cb: Callable[[str], None] | None = None,
) -> Path:
    """등록 완료된 PDF 원본을 변경하지 않고 접수완료 폴더에 복사한다."""
    source = Path(source_path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {source}")
    if source.suffix.lower() not in PDF_SUFFIXES:
        raise ValueError("PDF 파일만 복사할 수 있습니다.")

    destination_dir = (
        Path(output_dir).expanduser().resolve()
        if output_dir
        else source.parent / "접수완료"
    )
    destination_dir.mkdir(parents=True, exist_ok=True)
    destination = destination_dir / source.name
    source_digest = hashlib.sha256(source.read_bytes()).hexdigest()
    if destination.exists():
        if hashlib.sha256(destination.read_bytes()).hexdigest() != source_digest:
            counter = 2
            while destination.exists():
                destination = destination_dir / f"{source.stem}_{counter}{source.suffix}"
                counter += 1
    if not destination.exists():
        shutil.copy2(source, destination)

    log = log_cb or (lambda _msg: None)
    log(
        f"PDF 접수 완료 복사: {destination.name} - "
        f"접수번호 {format_hwp_receipt_entry(receipt_no, document_date)}"
    )
    return destination


def complete_registered_attachment(
    source_path: str | Path,
    receipt_no: str,
    document_date: date,
    output_dir: str | Path | None = None,
    visible: bool = True,
    log_cb: Callable[[str], None] | None = None,
) -> Path:
    """HWP에는 접수번호를 쓰고 PDF는 원본 그대로 접수완료 폴더에 보관한다."""
    suffix = Path(source_path).suffix.lower()
    if suffix in HWP_SUFFIXES:
        return write_receipt_number_to_hwp(
            source_path,
            receipt_no,
            document_date=document_date,
            output_dir=output_dir,
            visible=visible,
            log_cb=log_cb,
        )
    if suffix in PDF_SUFFIXES:
        return copy_registered_pdf(
            source_path,
            receipt_no,
            document_date=document_date,
            output_dir=output_dir,
            log_cb=log_cb,
        )
    raise ValueError("HWP, HWPX 또는 PDF 첨부파일만 처리할 수 있습니다.")


def apply_history(items: Iterable[ReceiptMailItem], history: ReceiptHistory) -> None:
    for item in items:
        found = history.find(item.stable_id)
        if not found:
            continue
        item.status = "처리완료"
        item.receipt_no = found["receipt_no"]
        output_path = Path(found["output_path"])
        item.output_path = output_path if output_path.exists() else None
