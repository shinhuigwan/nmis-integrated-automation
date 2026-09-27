import tempfile
import unittest
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject

from pdf_date_editor import (
    default_output_path,
    normalize_date_value,
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
            "1 0 0 1 484 765 Tm (2026-09-27 23:49) Tj ET"
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
            self.assertTrue(output.is_file())
            self.assertEqual(len(PdfReader(str(output)).pages), 1)

    def test_missing_date_label_is_rejected(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.pdf"
            output = Path(temp_dir) / "output.pdf"
            self._create_source(source, include_date=False)

            with self.assertRaisesRegex(ValueError, "DATE"):
                replace_pdf_date_label(source, output, "2026-09-01 14:30")


if __name__ == "__main__":
    unittest.main()
