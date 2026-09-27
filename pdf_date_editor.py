"""Static NMIS PDF DATE label editor.

UbiReport PDFs do not expose AcroForm fields.  This module locates the printed
``DATE:`` value, covers only that value, and writes a replacement into a copy.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    ArrayObject,
    DecodedStreamObject,
    DictionaryObject,
    NameObject,
)


PRINTED_DATE_RE = re.compile(
    r"^\d{4}[-./]\d{1,2}[-./]\d{1,2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$"
)


@dataclass(frozen=True)
class PdfDateEditResult:
    output_path: Path
    changed_pages: tuple[int, ...]
    old_values: tuple[str, ...]
    new_value: str


def normalize_date_value(value: str) -> str:
    """Validate user input and return the canonical printed DATE value."""
    text = " ".join(str(value or "").strip().split())
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
            return datetime.strptime(text, input_format).strftime(output_format)
        except ValueError:
            continue
    raise ValueError(
        "일자는 YYYY-MM-DD HH:MM 형식으로 입력하세요. "
        "예: 2026-09-01 14:30"
    )


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_DATE수정{input_path.suffix}")


def _find_date_box(page) -> tuple[float, float, float, str] | None:
    fragments: list[tuple[str, float, float, float]] = []

    def collect(text, _cm, tm, _font_dict, font_size) -> None:
        cleaned = " ".join(str(text or "").strip().split())
        if cleaned:
            fragments.append((cleaned, float(tm[4]), float(tm[5]), float(font_size or 9.0)))

    page.extract_text(visitor_text=collect)
    for text, x, y, font_size in fragments:
        if text.upper().startswith("DATE:"):
            inline_value = text[5:].strip()
            if PRINTED_DATE_RE.match(inline_value):
                # Some PDF producers combine the label and value into one text
                # fragment even when they were positioned separately.
                value_x = x + (font_size * 2.5) + 18.0
                return value_x, y, font_size, inline_value

    labels = [fragment for fragment in fragments if fragment[0].upper() == "DATE:"]
    for _label_text, label_x, label_y, _label_size in labels:
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
            old_value, x, y, font_size = candidates[0]
            return x, y, font_size, old_value
    return None


def _escape_pdf_text(value: str) -> str:
    return value.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _append_date_overlay(
    writer: PdfWriter,
    page,
    x: float,
    baseline_y: float,
    original_font_size: float,
    value: str,
) -> None:
    page_width = float(page.mediabox.width)
    right = page_width - 18.0
    available_width = max(1.0, right - x)
    font_size = min(9.0, max(6.5, original_font_size))
    estimated_width = len(value) * font_size * 0.53
    if estimated_width > available_width:
        font_size = max(6.5, font_size * available_width / estimated_width)

    resources = page.get("/Resources")
    if resources is None:
        resources_object = DictionaryObject()
        page[NameObject("/Resources")] = resources_object
    else:
        resources_object = resources.get_object()
    fonts = resources_object.get("/Font")
    if fonts is None:
        fonts_object = DictionaryObject()
        resources_object[NameObject("/Font")] = fonts_object
    else:
        fonts_object = fonts.get_object()

    font_key = NameObject("/NMISDateFont")
    if font_key not in fonts_object:
        font = DictionaryObject(
            {
                NameObject("/Type"): NameObject("/Font"),
                NameObject("/Subtype"): NameObject("/Type1"),
                NameObject("/BaseFont"): NameObject("/Helvetica"),
                NameObject("/Encoding"): NameObject("/WinAnsiEncoding"),
            }
        )
        fonts_object[font_key] = writer._add_object(font)

    rect_x = x - 1.5
    rect_y = baseline_y - 2.2
    rect_width = max(1.0, right - rect_x)
    rect_height = max(11.5, original_font_size + 3.0)
    commands = (
        "q\n"
        "1 1 1 rg\n"
        f"{rect_x:.2f} {rect_y:.2f} {rect_width:.2f} {rect_height:.2f} re f\n"
        "0 0 0 rg\n"
        "BT\n"
        f"/NMISDateFont {font_size:.2f} Tf\n"
        f"1 0 0 1 {x:.2f} {baseline_y:.2f} Tm\n"
        f"({_escape_pdf_text(value)}) Tj\n"
        "ET\n"
        "Q\n"
    )
    overlay_stream = DecodedStreamObject()
    overlay_stream.set_data(commands.encode("ascii"))
    overlay_ref = writer._add_object(overlay_stream)

    contents = page.get("/Contents")
    if contents is None:
        page[NameObject("/Contents")] = overlay_ref
    elif isinstance(contents, ArrayObject):
        contents.append(overlay_ref)
    else:
        page[NameObject("/Contents")] = ArrayObject([contents, overlay_ref])


def replace_pdf_date_label(
    input_path: Path | str,
    output_path: Path | str,
    new_value: str,
) -> PdfDateEditResult:
    """Replace printed values following ``DATE:`` and preserve the source file."""
    source = Path(input_path).expanduser().resolve()
    destination = Path(output_path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"PDF 파일을 찾지 못했습니다: {source}")
    if source.suffix.lower() != ".pdf":
        raise ValueError("PDF 파일만 처리할 수 있습니다.")
    if source == destination:
        raise ValueError("원본을 보호하기 위해 출력 파일은 다른 이름으로 지정하세요.")

    printed_value = normalize_date_value(new_value)
    reader = PdfReader(str(source))
    if reader.is_encrypted:
        raise ValueError("암호화된 PDF는 일자를 수정할 수 없습니다.")

    date_boxes = [_find_date_box(page) for page in reader.pages]

    if not any(date_boxes):
        raise ValueError("PDF에서 'DATE:' 뒤의 날짜 값을 찾지 못했습니다.")

    writer = PdfWriter(clone_from=str(source))
    changed_pages: list[int] = []
    old_values: list[str] = []

    for page_index, page in enumerate(writer.pages):
        box = date_boxes[page_index] if page_index < len(date_boxes) else None
        if box is not None:
            x, baseline_y, font_size, old_value = box
            _append_date_overlay(
                writer,
                page,
                x=x,
                baseline_y=baseline_y,
                original_font_size=font_size,
                value=printed_value,
            )
            changed_pages.append(page_index + 1)
            old_values.append(old_value)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as stream:
        writer.write(stream)

    # Reopen before returning so truncated or otherwise invalid output is caught.
    reopened = PdfReader(str(destination))
    if len(reopened.pages) != len(reader.pages):
        raise RuntimeError("저장된 PDF의 페이지 수 검증에 실패했습니다.")

    return PdfDateEditResult(
        output_path=destination,
        changed_pages=tuple(changed_pages),
        old_values=tuple(old_values),
        new_value=printed_value,
    )
