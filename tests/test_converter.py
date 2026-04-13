import builtins
import sys
import subprocess
import types
import pytest
import core.converter as converter
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


def test_windows_prerequisites_do_not_require_libreoffice(monkeypatch):
    monkeypatch.setattr(converter.platform, "system", lambda: "Windows")
    monkeypatch.setattr(converter, "_module_available", lambda module_name: True, raising=False)
    monkeypatch.setattr(converter, "find_libreoffice", lambda: None, raising=False)

    assert converter.conversion_prerequisite_error(["a.docx", "b.xlsx"]) is None


def test_convert_docx_on_windows_uses_docx2pdf(tmp_path, monkeypatch):
    """Windows Word conversion should use local MS Office through docx2pdf."""
    docx_file = tmp_path / "test.docx"
    docx_file.write_bytes(b"fake docx")
    output_dir = tmp_path / "out"
    expected_pdf = output_dir / "test.pdf"
    calls = []

    def fake_convert(source, destination):
        calls.append((source, destination))
        expected_pdf.write_bytes(b"%PDF-1.4 fake content")

    monkeypatch.setattr(converter.platform, "system", lambda: "Windows")
    monkeypatch.setitem(sys.modules, "docx2pdf", types.SimpleNamespace(convert=fake_convert))

    result = convert_to_pdf(str(docx_file), str(output_dir))

    assert result == str(expected_pdf)
    assert calls == [(str(docx_file), str(expected_pdf))]


def test_convert_docx_on_windows_handles_windowed_stdio(tmp_path, monkeypatch):
    """PyInstaller --windowed sets stdio to None; docx2pdf still writes progress."""
    docx_file = tmp_path / "test.docx"
    docx_file.write_bytes(b"fake docx")
    output_dir = tmp_path / "out"
    expected_pdf = output_dir / "test.pdf"

    def fake_convert(source, destination):
        sys.stdout.write("starting")
        sys.stderr.write("progress")
        expected_pdf.write_bytes(b"%PDF-1.4 fake content")

    monkeypatch.setattr(converter.platform, "system", lambda: "Windows")
    monkeypatch.setitem(sys.modules, "docx2pdf", types.SimpleNamespace(convert=fake_convert))
    monkeypatch.setattr(sys, "stdout", None)
    monkeypatch.setattr(sys, "stderr", None)

    result = convert_to_pdf(str(docx_file), str(output_dir))

    assert result == str(expected_pdf)


def test_convert_xlsx_on_windows_uses_excel_com(tmp_path, monkeypatch):
    """Windows Excel conversion should use local MS Excel through COM."""
    xlsx_file = tmp_path / "test.xlsx"
    xlsx_file.write_bytes(b"fake xlsx")
    output_dir = tmp_path / "out"
    expected_pdf = output_dir / "test.pdf"
    opened = []
    exports = []
    closed = []
    quit_called = []

    class FakeWorkbook:
        def ExportAsFixedFormat(self, output_type, destination):
            exports.append((output_type, destination))
            expected_pdf.write_bytes(b"%PDF-1.4 fake content")

        def Close(self, save_changes=False):
            closed.append(save_changes)

    class FakeWorkbooks:
        def Open(self, source):
            opened.append(source)
            return FakeWorkbook()

    class FakeExcel:
        def __init__(self):
            self.Workbooks = FakeWorkbooks()
            self.Visible = True
            self.DisplayAlerts = True

        def Quit(self):
            quit_called.append(True)

    fake_client = types.SimpleNamespace(DispatchEx=lambda name: FakeExcel())

    monkeypatch.setattr(converter.platform, "system", lambda: "Windows")
    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=fake_client))
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    result = convert_to_pdf(str(xlsx_file), str(output_dir))

    assert result == str(expected_pdf)
    assert opened == [str(xlsx_file)]
    assert exports == [(0, str(expected_pdf))]
    assert closed == [False]
    assert quit_called == [True]


def test_convert_docx_on_windows_missing_docx2pdf(tmp_path, monkeypatch):
    docx_file = tmp_path / "test.docx"
    docx_file.write_bytes(b"fake docx")

    monkeypatch.setattr(converter.platform, "system", lambda: "Windows")
    monkeypatch.delitem(sys.modules, "docx2pdf", raising=False)
    original_import = builtins.__import__

    def fake_import(name, *args, **kwargs):
        if name == "docx2pdf":
            raise ImportError("blocked")
        return original_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)

    with pytest.raises(RuntimeError, match="docx2pdf"):
        convert_to_pdf(str(docx_file), str(tmp_path / "out"))


def test_convert_docx_runs_libreoffice_headless(tmp_path, monkeypatch):
    """Non-Windows Word files still convert through LibreOffice in headless mode."""
    docx_file = tmp_path / "test.docx"
    docx_file.write_bytes(b"fake docx")
    output_dir = tmp_path / "out"
    expected_pdf = output_dir / "test.pdf"
    calls = []

    def fake_run(cmd, capture_output, text, timeout):
        calls.append((cmd, capture_output, text, timeout))
        expected_pdf.write_bytes(b"%PDF-1.4 fake content")
        return subprocess.CompletedProcess(cmd, 0, stdout="ok", stderr="")

    monkeypatch.setattr(converter.platform, "system", lambda: "Darwin")
    monkeypatch.setattr(converter, "find_libreoffice", lambda: "/opt/libreoffice/soffice", raising=False)
    monkeypatch.setattr(converter, "subprocess", types.SimpleNamespace(run=fake_run), raising=False)

    result = convert_to_pdf(str(docx_file), str(output_dir))

    assert result == str(expected_pdf)
    assert calls
    cmd, capture_output, text, timeout = calls[0]
    assert cmd[0] == "/opt/libreoffice/soffice"
    assert "--headless" in cmd
    assert "--convert-to" in cmd
    assert "pdf" in cmd
    assert str(output_dir) in cmd
    assert str(docx_file) in cmd
    assert capture_output is True
    assert text is True
    assert timeout == 120
