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
    monkeypatch.setattr(converter, "_module_available", lambda module_name: module_name == "win32com.client", raising=False)
    monkeypatch.setattr(converter, "find_libreoffice", lambda: None, raising=False)

    assert converter.conversion_prerequisite_error(["a.docx", "b.xlsx"]) is None


def test_convert_docx_on_windows_uses_word_com(tmp_path, monkeypatch):
    """Windows Word conversion should call the Word COM automation backend."""
    docx_file = tmp_path / "test.docx"
    docx_file.write_bytes(b"fake docx")
    output_dir = tmp_path / "out"
    expected_pdf = output_dir / "test.pdf"
    calls = {"dispatch": [], "open": [], "save": [], "close": [], "quit": []}

    class FakeDocument:
        def SaveAs2(self, destination, FileFormat):
            calls["save"].append((destination, FileFormat))
            expected_pdf.write_bytes(b"%PDF-1.4 fake content")

        def Close(self, save_changes=False):
            calls["close"].append(save_changes)

    class FakeDocuments:
        def Open(self, source):
            calls["open"].append(source)
            return FakeDocument()

    class FakeWord:
        def __init__(self):
            self.Documents = FakeDocuments()

        def Quit(self):
            calls["quit"].append(True)

    def fake_dispatch(prog_id):
        calls["dispatch"].append(prog_id)
        return FakeWord()

    monkeypatch.setattr(converter.platform, "system", lambda: "Windows")
    fake_client = types.SimpleNamespace(DispatchEx=fake_dispatch)
    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=fake_client))
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    result = convert_to_pdf(str(docx_file), str(output_dir))

    assert result == str(expected_pdf)
    assert calls["dispatch"] == ["Word.Application"]
    assert calls["open"] == [str(docx_file)]
    assert calls["save"] == [(str(expected_pdf), 17)]
    assert calls["close"] == [False]
    assert calls["quit"] == [True]


def test_convert_docx_on_windows_falls_back_to_wps_writer(tmp_path, monkeypatch):
    """If Microsoft Word cannot save the file, try WPS Writer."""
    docx_file = tmp_path / "test.docx"
    docx_file.write_bytes(b"fake docx")
    output_dir = tmp_path / "out"
    expected_pdf = output_dir / "test.pdf"
    dispatched = []

    class BrokenDocument:
        def SaveAs2(self, destination, FileFormat):
            raise RuntimeError("Open.SaveAs")

        def Close(self, save_changes=False):
            pass

    class BrokenDocuments:
        def Open(self, source):
            return BrokenDocument()

    class BrokenWord:
        def __init__(self):
            self.Documents = BrokenDocuments()

        def Quit(self):
            pass

    class WpsDocument:
        def SaveAs(self, destination, FileFormat):
            expected_pdf.write_bytes(b"%PDF-1.4 fake content")

        def Close(self, save_changes=False):
            pass

    class WpsDocuments:
        def Open(self, source):
            return WpsDocument()

    class WpsWriter:
        def __init__(self):
            self.Documents = WpsDocuments()

        def Quit(self):
            pass

    def fake_dispatch(prog_id):
        dispatched.append(prog_id)
        if prog_id == "Word.Application":
            return BrokenWord()
        if prog_id == "KWPS.Application":
            return WpsWriter()
        raise RuntimeError("not installed")

    monkeypatch.setattr(converter.platform, "system", lambda: "Windows")
    fake_client = types.SimpleNamespace(DispatchEx=fake_dispatch)
    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=fake_client))
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    result = convert_to_pdf(str(docx_file), str(output_dir))

    assert result == str(expected_pdf)
    assert dispatched == ["Word.Application", "KWPS.Application"]


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


def test_convert_xlsx_on_windows_falls_back_to_wps_spreadsheets(tmp_path, monkeypatch):
    """If Microsoft Excel is unavailable, try WPS Spreadsheets."""
    xlsx_file = tmp_path / "test.xlsx"
    xlsx_file.write_bytes(b"fake xlsx")
    output_dir = tmp_path / "out"
    expected_pdf = output_dir / "test.pdf"
    dispatched = []

    class FakeWorkbook:
        def ExportAsFixedFormat(self, output_type, destination):
            expected_pdf.write_bytes(b"%PDF-1.4 fake content")

        def Close(self, save_changes=False):
            pass

    class FakeWorkbooks:
        def Open(self, source):
            return FakeWorkbook()

    class FakeWpsSpreadsheet:
        def __init__(self):
            self.Workbooks = FakeWorkbooks()

        def Quit(self):
            pass

    def fake_dispatch(prog_id):
        dispatched.append(prog_id)
        if prog_id == "Excel.Application":
            raise RuntimeError("Excel missing")
        if prog_id == "KET.Application":
            return FakeWpsSpreadsheet()
        raise RuntimeError("not installed")

    monkeypatch.setattr(converter.platform, "system", lambda: "Windows")
    fake_client = types.SimpleNamespace(DispatchEx=fake_dispatch)
    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=fake_client))
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    result = convert_to_pdf(str(xlsx_file), str(output_dir))

    assert result == str(expected_pdf)
    assert dispatched == ["Excel.Application", "KET.Application"]


def test_convert_docx_on_windows_missing_win32com(tmp_path, monkeypatch):
    docx_file = tmp_path / "test.docx"
    docx_file.write_bytes(b"fake docx")

    monkeypatch.setattr(converter.platform, "system", lambda: "Windows")
    monkeypatch.delitem(sys.modules, "win32com", raising=False)
    monkeypatch.delitem(sys.modules, "win32com.client", raising=False)
    monkeypatch.setattr(converter, "_import_win32_client", lambda: (_ for _ in ()).throw(ImportError("blocked")), raising=False)

    with pytest.raises(RuntimeError, match="win32com"):
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
