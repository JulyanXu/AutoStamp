import os
import platform
import pytest
from unittest.mock import patch, MagicMock
from core.converter import find_libreoffice, convert_to_pdf


def test_find_libreoffice_returns_string():
    # May return None if LibreOffice not installed, but should not raise
    result = find_libreoffice()
    assert result is None or isinstance(result, str)


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


@patch("core.converter.find_libreoffice", return_value=None)
def test_convert_docx_no_libreoffice(mock_find, tmp_path):
    docx_file = tmp_path / "test.docx"
    docx_file.write_bytes(b"fake docx")
    with pytest.raises(RuntimeError, match="LibreOffice"):
        convert_to_pdf(str(docx_file), str(tmp_path / "out"))
