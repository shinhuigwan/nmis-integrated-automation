import random
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject
from reportlab.pdfgen import canvas

from pdf_date_editor import (
    default_output_path,
    generate_random_date_value,
    normalize_date_value,
    read_pdf_date_info,
    replace_pdf_date_label,
)


class PdfDateEditorTests(unittest.TestCase):
    @staticmethod
    def _create_source(path: Path, include_date: bool = True) -> None:
        writer = PdfWriter()
        page = writer.add_blank_page(width=595.28, height=841.89)
        font = DictionaryObject(
            {
                NameObject("/Type"): NameObject("/Font"),
                NameObject("/Subtype"): NameObject("/Type1"),
                NameObject("/BaseFont"): NameObject("/Helvetica"),
                NameObject("/Encoding"): NameObject("/WinAnsiEncoding"),
            }
        )
        font_ref = writer._add_object(font)
        page[NameObject("/Resources")] = DictionaryObject(
            {
                NameObject("/Font"): DictionaryObject(
                    {NameObject("/F1"): font_ref}
                )
            }
        )
        text = (
            "BT /F1 9 Tf 1 0 0 1 443.5 765 Tm (DATE:) Tj "
            "1 0 0 1 484 765 Tm (2026-09-27 23:49) Tj "
            "1 0 0 1 109 435 Tm (INITIAL ENTRY: 2026-09-03 14:21) Tj ET"
            if include_date
            else "BT /F1 9 Tf 1 0 0 1 50 700 Tm (No date label) Tj ET"
        )
        stream = DecodedStreamObject()
        stream.set_data(text.encode("ascii"))
        page[NameObject("/Contents")] = writer._add_object(stream)
        with path.open("wb") as output:
            writer.write(output)

    def test_normalize_date_value(self):
        self.assertEqual(
            normalize_date_value("2026.09.01 14:30"),
            "2026-09-01 14:30",
        )
        self.assertEqual(normalize_date_value("2026-09-01"), "2026-09-01")
        with self.assertRaises(ValueError):
            normalize_date_value("2026-99-99")

    def test_default_output_path_preserves_source(self):
        source = Path("account_invoice_out_bo.pdf")
        self.assertEqual(
            default_output_path(source).name,
            "account_invoice_out_bo_DATE수정.pdf",
        )

    def test_replace_static_date_value(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.pdf"
            output = Path(temp_dir) / "output.pdf"
            self._create_source(source)

            result = replace_pdf_date_label(
                source,
                output,
                "2026-09-01 14:30",
            )

            self.assertEqual(result.changed_pages, (1,))
            self.assertEqual(result.old_values, ("2026-09-27 23:49",))
            self.assertEqual(result.new_value, "2026-09-01 14:30")
            self.assertEqual(result.source_fonts, ("Helvetica",))
            self.assertTrue(result.font_matched)
            self.assertTrue(output.is_file())
            self.assertEqual(len(PdfReader(str(output)).pages), 1)

    def test_read_original_date_and_font(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.pdf"
            self._create_source(source)

            info = read_pdf_date_info(source)

            self.assertEqual(info.initial_value, "2026-09-03 14:21")
            self.assertEqual(info.initial_entries, ("2026-09-03 14:21",))
            self.assertEqual(info.font_names, ("Helvetica",))
            self.assertEqual(info.locations[0].font_size, 9.0)

    def test_random_time_before_17_is_within_next_10_minutes(self):
        generated = generate_random_date_value(
            "2026-09-01 15:09",
            rng=random.Random(7),
        )
        initial_dt = datetime.strptime("2026-09-01 15:09", "%Y-%m-%d %H:%M")
        generated_dt = datetime.strptime(generated, "%Y-%m-%d %H:%M")
        self.assertGreater(generated_dt, initial_dt)
        self.assertLessEqual((generated_dt - initial_dt).total_seconds(), 600)

    def test_random_times_keep_date_and_seconds(self):
        for initial in ("2026-09-01 09:00", "2026-09-01 16:59", "2026-09-01 23:55", "2026-09-01 23:59", "2026-09-01 23:59:31"):
            for seed in range(20):
                generated = datetime.fromisoformat(generate_random_date_value(initial, random.Random(seed)))
                original = datetime.fromisoformat(initial)
                self.assertEqual(generated.date(), original.date())
                self.assertGreaterEqual((generated - original).total_seconds(), 0)
                self.assertLessEqual((generated - original).total_seconds(), 600)

    @staticmethod
    def _create_multi_vouchers(path, missing_top=False):
        output = canvas.Canvas(str(path), pagesize=(595.28, 841.89))
        for page_index in range(2):
            for voucher_index, header in enumerate((765, 347)):
                output.setFont("Helvetica", 9)
                output.drawString(443.5, header, "DATE:")
                output.drawString(484, header, "2026-10-14 11:33")
                if not (missing_top and page_index == 0 and voucher_index == 0):
                    initial = f"2026-04-{14 + page_index:02d} 13:0{voucher_index}"
                    output.drawString(108, header - 327, f"INITIAL ENTRY: {initial}")
            output.showPage()
        output.save()

    def test_each_voucher_on_every_page_uses_its_own_initial(self):
        with tempfile.TemporaryDirectory() as directory:
            source, output = Path(directory) / "source.pdf", Path(directory) / "output.pdf"
            self._create_multi_vouchers(source)
            original = source.read_bytes()
            info = read_pdf_date_info(source)
            self.assertEqual([item.page_number for item in info.locations], [1, 1, 2, 2])
            self.assertEqual([item.initial_value for item in info.locations], [
                "2026-04-14 13:00", "2026-04-14 13:01", "2026-04-15 13:00", "2026-04-15 13:01",
            ])
            values = tuple(generate_random_date_value(item.initial_value, random.Random(i)) for i, item in enumerate(info.locations))
            result = replace_pdf_date_label(source, output, values)
            self.assertEqual(result.changed_pages, (1, 2))
            self.assertEqual(result.new_values, values)
            self.assertEqual(tuple(item.value for item in read_pdf_date_info(output).locations), values)
            self.assertEqual(source.read_bytes(), original)
            replace_pdf_date_label(output, Path(directory) / "common.pdf", "2026-04-16 10:00")
            self.assertEqual([item.value for item in read_pdf_date_info(Path(directory) / "common.pdf").locations], ["2026-04-16 10:00"] * 4)

    def test_missing_footer_does_not_use_another_vouchers_initial(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.pdf"
            self._create_multi_vouchers(source, missing_top=True)
            info = read_pdf_date_info(source)
            self.assertEqual(info.locations[0].initial_value, "")
            self.assertEqual(info.locations[1].initial_value, "2026-04-14 13:01")
            output = Path(directory) / "output.pdf"
            with self.assertRaisesRegex(ValueError, "일치"):
                replace_pdf_date_label(source, output, ("2026-04-14 13:05",))
            self.assertFalse(output.exists())

    def test_random_time_after_17_is_within_next_10_minutes(self):
        generated = generate_random_date_value(
            "2026-09-01 17:21",
            rng=random.Random(7),
        )
        initial_dt = datetime.strptime("2026-09-01 17:21", "%Y-%m-%d %H:%M")
        generated_dt = datetime.strptime(generated, "%Y-%m-%d %H:%M")
        delta_minutes = int((generated_dt - initial_dt).total_seconds() // 60)
        self.assertGreaterEqual(delta_minutes, 1)
        self.assertLessEqual(delta_minutes, 10)

    def test_missing_date_label_is_rejected(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.pdf"
            output = Path(temp_dir) / "output.pdf"
            self._create_source(source, include_date=False)

            with self.assertRaisesRegex(ValueError, "DATE"):
                replace_pdf_date_label(source, output, "2026-09-01 14:30")


if __name__ == "__main__":
    unittest.main()
