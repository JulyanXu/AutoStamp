import os
import sys
import pytest
from unittest.mock import patch, MagicMock
from core.converter import convert_to_pdf


def test_convert_pdf_passthrough(tmp_path):
    """A .pdf input should be returned as-is, no conversion needed."""
    pdf_file = tmp_path / "test.pdf"
    pdf_file.write_bytes(b"%PDF-1.4 fake content")
    result = convert_to_pdf(str(pdf_file), str(tmp_path / "out"))
    assert result == str(pdf_file)


def test_convert_unsupported_format(tmp_path):
    txt_file = tmp_path / "test.txt"
    txt_file.write_text("hello")
    with pytest.raises(ValueError, match="不支持的文件格式"):
        convert_to_pdf(str(txt_file), str(tmp_path / "out"))


def test_convert_docx_missing_docx2pdf(tmp_path):
    """If docx2pdf is not importable, RuntimeError with helpful message."""
    docx_file = tmp_path / "test.docx"
    docx_file.write_bytes(b"fake docx")

    mock_modules = dict(sys.modules)
    mock_modules["docx2pdf"] = None
    with patch.dict("sys.modules", {"docx2pdf": None}):
        with pytest.raises((RuntimeError, ImportError)):
            convert_to_pdf(str(docx_file), str(tmp_path / "out"))


def test_convert_docx_conversion_failure(tmp_path):
    """When docx2pdf.convert raises, RuntimeError with friendly message."""
    docx_file = tmp_path / "test.docx"
    docx_file.write_bytes(b"fake docx")

    mock_docx2pdf = MagicMock()
    mock_docx2pdf.convert.side_effect = Exception("Word COM error")
    with patch.dict("sys.modules", {"docx2pdf": mock_docx2pdf}):
        with pytest.raises(RuntimeError, match="转换失败"):
            convert_to_pdf(str(docx_file), str(tmp_path / "out"))
