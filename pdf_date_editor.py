"""Static NMIS PDF DATE label reader and editor.

UbiReport PDFs do not expose AcroForm fields. This module locates the printed
``DATE:`` value, reports the original timestamp/font, covers only that value,
and writes a replacement into a copy of the source PDF.
"""

from __future__ import annotations

import io
import os
import random
import re
import tempfile
from dataclasses import dataclass
from datetime import datetime, time, timedelta
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


PRINTED_DATE_RE = re.compile(
    r"^\d{4}[-./]\d{1,2}[-./]\d{1,2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$"
)
INITIAL_ENTRY_RE = re.compile(
    r"(?:최초\s*입력일|INITIAL\s*ENTRY)\s*[:：]?\s*"
    r"(\d{4}[-./]\d{1,2}[-./]\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?)",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class PdfDateLocation:
    page_number: int
    value: str
    font_name: str
    x: float
    baseline_y: float
    font_size: float


@dataclass(frozen=True)
class PdfDateInfo:
    source_path: Path
    page_count: int
    locations: tuple[PdfDateLocation, ...]
    initial_entries: tuple[str, ...]

    @property
    def initial_value(self) -> str:
        return self.initial_entries[0] if self.initial_entries else ""

    @property
    def current_date_value(self) -> str:
        return self.locations[0].value if self.locations else ""

    @property
    def font_names(self) -> tuple[str, ...]:
        return tuple(dict.fromkeys(location.font_name for location in self.locations))


@dataclass(frozen=True)
class PdfDateEditResult:
    output_path: Path
    changed_pages: tuple[int, ...]
    old_values: tuple[str, ...]
    new_value: str
    source_fonts: tuple[str, ...]
    font_matched: bool


def normalize_date_value(value: str) -> str:
    """Validate user input and return the canonical printed DATE value."""
    text_value = " ".join(str(value or "").strip().split())
    formats = (
        ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"),
        ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M"),
        ("%Y.%m.%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"),
        ("%Y.%m.%d %H:%M", "%Y-%m-%d %H:%M"),
        ("%Y-%m-%d", "%Y-%m-%d"),
        ("%Y.%m.%d", "%Y-%m-%d"),
    )
    for input_format, output_format in formats:
        try:
            return datetime.strptime(text_value, input_format).strftime(output_format)
        except ValueError:
            continue
    raise ValueError(
        "일자는 YYYY-MM-DD HH:MM 형식으로 입력하세요. "
        "예: 2026-09-01 14:30"
    )


def _parse_timestamp(value: str) -> datetime:
    normalized = normalize_date_value(value)
    for date_format in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            parsed = datetime.strptime(normalized, date_format)
            if date_format == "%Y-%m-%d":
                return parsed.replace(hour=9)
            return parsed
        except ValueError:
            continue
    raise ValueError(f"최초 입력일을 해석하지 못했습니다: {value}")


def generate_random_date_value(
    initial_value: str,
    rng: random.Random | None = None,
) -> str:
    """Return a random minute after the initial timestamp.

    Up to 17:00, the result stays on the same date and does not exceed 17:00.
    For an original timestamp at or after 17:00, the result is 1-10 minutes
    after that timestamp as requested by the user.
    """
    generator = rng or random.SystemRandom()
    initial = _parse_timestamp(initial_value)
    work_end = datetime.combine(initial.date(), time(hour=17))

    if initial < work_end:
        remaining_minutes = int((work_end - initial).total_seconds() // 60)
        if remaining_minutes <= 0:
            generated = work_end
        else:
            generated = initial + timedelta(minutes=generator.randint(1, remaining_minutes))
            if generated > work_end:
                generated = work_end
    else:
        generated = initial + timedelta(minutes=generator.randint(1, 10))

    return generated.strftime("%Y-%m-%d %H:%M")


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_DATE수정{input_path.suffix}")


def _clean_font_name(font_dict) -> str:
    if not font_dict:
        return "Unknown"
    base_font = str(font_dict.get("/BaseFont") or "Unknown").lstrip("/")
    if "+" in base_font:
        base_font = base_font.split("+", 1)[1]
    return base_font


def _find_date_box(page, page_number: int) -> PdfDateLocation | None:
    fragments: list[tuple[str, float, float, float, str]] = []

    def collect(text_value, _cm, tm, font_dict, font_size) -> None:
        cleaned = " ".join(str(text_value or "").strip().split())
        if cleaned:
            fragments.append(
                (
                    cleaned,
                    float(tm[4]),
                    float(tm[5]),
                    float(font_size or 9.0),
                    _clean_font_name(font_dict),
                )
            )

    page.extract_text(visitor_text=collect)
    for text_value, x, y, font_size, font_name in fragments:
        if text_value.upper().startswith("DATE:"):
            inline_value = text_value[5:].strip()
            if PRINTED_DATE_RE.match(inline_value):
                value_x = x + (font_size * 2.5) + 18.0
                return PdfDateLocation(
                    page_number, inline_value, font_name, value_x, y, font_size
                )

    labels = [fragment for fragment in fragments if fragment[0].upper() == "DATE:"]
    for _label_text, label_x, label_y, _label_size, _label_font in labels:
        candidates = sorted(
            (
                fragment
                for fragment in fragments
                if fragment[1] > label_x
                and abs(fragment[2] - label_y) <= 2.5
                and PRINTED_DATE_RE.match(fragment[0])
            ),
            key=lambda fragment: fragment[1],
        )
        if candidates:
            old_value, x, y, font_size, font_name = candidates[0]
            return PdfDateLocation(
                page_number, old_value, font_name, x, y, font_size
            )
    return None


def read_pdf_date_info(input_path: Path | str) -> PdfDateInfo:
    """Read the document's first-entry time and DATE values without editing."""
    source = Path(input_path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"PDF 파일을 찾지 못했습니다: {source}")
    if source.suffix.lower() != ".pdf":
        raise ValueError("PDF 파일만 처리할 수 있습니다.")

    reader = PdfReader(str(source))
    if reader.is_encrypted:
        raise ValueError("암호화된 PDF는 일자를 확인할 수 없습니다.")
    locations: list[PdfDateLocation] = []
    initial_entries: list[str] = []
    for page_number, page in enumerate(reader.pages, start=1):
        location = _find_date_box(page, page_number)
        if location is not None:
            locations.append(location)
        page_text = " ".join((page.extract_text() or "").split())
        for match in INITIAL_ENTRY_RE.finditer(page_text):
            normalized = normalize_date_value(match.group(1))
            if normalized not in initial_entries:
                initial_entries.append(normalized)
    if not locations:
        raise ValueError("PDF에서 우측 상단 'DATE:' 값을 찾지 못했습니다.")
    return PdfDateInfo(
        source,
        len(reader.pages),
        tuple(locations),
        tuple(initial_entries),
    )


_REGISTERED_FONTS: dict[str, str] = {}


def _register_matching_font(source_font: str) -> tuple[str, bool]:
    normalized = re.sub(r"[^a-z]", "", source_font.lower())
    standard_fonts = {
        "helvetica": "Helvetica",
        "helveticabold": "Helvetica-Bold",
        "courier": "Courier",
        "timesroman": "Times-Roman",
    }
    if normalized in standard_fonts:
        return standard_fonts[normalized], True

    windows_fonts = (
        ("gulimche", "NMIS-GulimChe", Path(r"C:\Windows\Fonts\gulim.ttc"), 1),
        ("gulim", "NMIS-Gulim", Path(r"C:\Windows\Fonts\gulim.ttc"), 0),
        ("dotumche", "NMIS-DotumChe", Path(r"C:\Windows\Fonts\gulim.ttc"), 3),
        ("dotum", "NMIS-Dotum", Path(r"C:\Windows\Fonts\gulim.ttc"), 2),
        ("batangche", "NMIS-BatangChe", Path(r"C:\Windows\Fonts\batang.ttc"), 1),
        ("batang", "NMIS-Batang", Path(r"C:\Windows\Fonts\batang.ttc"), 0),
    )
    for token, registered_name, font_path, subfont_index in windows_fonts:
        if token == normalized and font_path.is_file():
            if registered_name not in _REGISTERED_FONTS:
                pdfmetrics.registerFont(
                    TTFont(
                        registered_name,
                        str(font_path),
                        subfontIndex=subfont_index,
                    )
                )
                _REGISTERED_FONTS[registered_name] = str(font_path)
            return registered_name, True
    return "Helvetica", False


def _create_overlay_page(page, location: PdfDateLocation, value: str):
    page_width = float(page.mediabox.width)
    page_height = float(page.mediabox.height)
    right = page_width - 18.0
    rect_x = location.x - 1.5
    rect_y = location.baseline_y - 2.2
    rect_width = max(1.0, right - rect_x)
    rect_height = max(11.5, location.font_size + 3.0)
    font_name, matched = _register_matching_font(location.font_name)

    buffer = io.BytesIO()
    overlay_canvas = canvas.Canvas(
        buffer,
        pagesize=(page_width, page_height),
        pageCompression=1,
    )
    overlay_canvas.setFillColorRGB(1, 1, 1)
    overlay_canvas.rect(rect_x, rect_y, rect_width, rect_height, stroke=0, fill=1)
    overlay_canvas.setFillColorRGB(0, 0, 0)
    overlay_canvas.setFont(font_name, location.font_size)
    overlay_canvas.drawString(location.x, location.baseline_y, value)
    overlay_canvas.save()
    buffer.seek(0)
    return PdfReader(buffer).pages[0], matched


def replace_pdf_date_label(
    input_path: Path | str,
    output_path: Path | str,
    new_value: str,
) -> PdfDateEditResult:
    """Replace printed values following ``DATE:`` and preserve the source file."""
    source = Path(input_path).expanduser().resolve()
    destination = Path(output_path).expanduser().resolve()
    if source == destination:
        raise ValueError("원본을 보호하기 위해 출력 파일은 다른 이름으로 지정하세요.")

    info = read_pdf_date_info(source)
    printed_value = normalize_date_value(new_value)
    writer = PdfWriter(clone_from=str(source))
    locations_by_page = {location.page_number: location for location in info.locations}
    changed_pages: list[int] = []
    old_values: list[str] = []
    matched_flags: list[bool] = []

    for page_number, page in enumerate(writer.pages, start=1):
        location = locations_by_page.get(page_number)
        if location is None:
            continue
        overlay_page, matched = _create_overlay_page(page, location, printed_value)
        page.merge_page(overlay_page, over=True, expand=False)
        changed_pages.append(page_number)
        old_values.append(location.value)
        matched_flags.append(matched)

    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{destination.stem}-",
            suffix=".pdf",
            dir=destination.parent,
            delete=False,
        ) as temp_stream:
            temp_path = Path(temp_stream.name)
            writer.write(temp_stream)

        reopened = PdfReader(str(temp_path))
        if len(reopened.pages) != info.page_count:
            raise RuntimeError("저장된 PDF의 페이지 수 검증에 실패했습니다.")
        os.replace(temp_path, destination)
        temp_path = None
    finally:
        if temp_path is not None and temp_path.exists():
            temp_path.unlink()

    return PdfDateEditResult(
        output_path=destination,
        changed_pages=tuple(changed_pages),
        old_values=tuple(old_values),
        new_value=printed_value,
        source_fonts=info.font_names,
        font_matched=all(matched_flags),
    )
