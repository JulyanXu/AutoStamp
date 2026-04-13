import os
import platform
import shutil
import subprocess


SUPPORTED_EXTENSIONS = {".docx", ".xlsx", ".pdf"}
WD_FORMAT_PDF = 17
XL_TYPE_PDF = 0
WORD_BACKENDS = [
    ("Microsoft Word", "Word.Application"),
    ("WPS Writer", "KWPS.Application"),
    ("WPS Writer", "WPS.Application"),
]
SPREADSHEET_BACKENDS = [
    ("Microsoft Excel", "Excel.Application"),
    ("WPS Spreadsheets", "KET.Application"),
    ("WPS Spreadsheets", "ET.Application"),
]


def find_libreoffice() -> str | None:
    system = platform.system()
    if system == "Darwin":
        candidate = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(candidate):
            return candidate

    return shutil.which("soffice")


def conversion_prerequisite_error(file_paths: list[str]) -> str | None:
    office_exts = {
        os.path.splitext(path)[1].lower()
        for path in file_paths
        if os.path.splitext(path)[1].lower() in {".docx", ".xlsx"}
    }
    if not office_exts:
        return None

    if platform.system() == "Windows":
        if not _module_available("win32com.client"):
            return "缺少 pywin32/win32com 依赖，无法调用 Microsoft Office 或 WPS 转换文件。请重新安装或重新打包程序。"
        return None

    if not find_libreoffice():
        return (
            "文件列表中包含 Word/Excel 文件，但未找到 LibreOffice。\n"
            "请确保 LibreOffice 已安装。"
        )
    return None


def _module_available(module_name: str) -> bool:
    try:
        __import__(module_name)
    except ImportError:
        return False
    return True


def convert_to_pdf(input_path: str, temp_dir: str) -> str:
    ext = os.path.splitext(input_path)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"不支持的文件格式: {ext}")
    if ext == ".pdf":
        return input_path

    os.makedirs(temp_dir, exist_ok=True)
    if platform.system() == "Windows":
        if ext == ".docx":
            return _convert_docx_with_windows_office(input_path, temp_dir)
        if ext == ".xlsx":
            return _convert_xlsx_with_windows_office(input_path, temp_dir)

    return _convert_with_libreoffice(input_path, temp_dir)


def _output_path(input_path: str, temp_dir: str) -> str:
    base = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(temp_dir, base + ".pdf")


def _import_win32_client():
    import win32com.client

    return win32com.client


def _dispatch(client, prog_id: str):
    dispatch = getattr(client, "DispatchEx", None) or client.Dispatch
    return dispatch(prog_id)


def _set_quiet(app) -> None:
    for attr, value in (("Visible", False), ("DisplayAlerts", False)):
        try:
            setattr(app, attr, value)
        except Exception:
            pass


def _safe_close(obj, method_name: str, *args) -> None:
    if obj is None:
        return
    try:
        getattr(obj, method_name)(*args)
    except Exception:
        pass


def _format_backend_errors(errors: list[tuple[str, str, Exception]]) -> str:
    return "; ".join(
        f"{name} ({prog_id}): {error}" for name, prog_id, error in errors
    )


def _convert_docx_with_windows_office(input_path: str, temp_dir: str) -> str:
    output_path = _output_path(input_path, temp_dir)
    try:
        client = _import_win32_client()
    except ImportError as e:
        raise RuntimeError("缺少 pywin32/win32com 依赖，无法调用 Microsoft Word 或 WPS Writer 转换文件。请重新安装或重新打包程序。") from e

    errors = []
    for name, prog_id in WORD_BACKENDS:
        try:
            _convert_docx_with_writer_backend(client, prog_id, input_path, output_path)
            if os.path.exists(output_path):
                return output_path
            raise RuntimeError(f"转换后未找到输出文件: {output_path}")
        except Exception as e:
            errors.append((name, prog_id, e))

    raise RuntimeError(
        f"转换失败: {os.path.basename(input_path)}\n"
        "请确认电脑已安装 Microsoft Word 或 WPS Office，并且文件未被占用。\n"
        f"详情: {_format_backend_errors(errors)}"
    )


def _convert_docx_with_writer_backend(client, prog_id: str, input_path: str, output_path: str) -> None:
    app = None
    doc = None
    try:
        app = _dispatch(client, prog_id)
        _set_quiet(app)
        doc = app.Documents.Open(input_path)
        try:
            doc.SaveAs2(output_path, FileFormat=WD_FORMAT_PDF)
        except AttributeError:
            doc.SaveAs(output_path, FileFormat=WD_FORMAT_PDF)
    finally:
        _safe_close(doc, "Close", False)
        _safe_close(app, "Quit")


def _convert_xlsx_with_windows_office(input_path: str, temp_dir: str) -> str:
    output_path = _output_path(input_path, temp_dir)
    try:
        client = _import_win32_client()
    except ImportError as e:
        raise RuntimeError("缺少 pywin32/win32com 依赖，无法调用 Microsoft Excel 或 WPS 表格转换文件。请重新安装或重新打包程序。") from e

    errors = []
    for name, prog_id in SPREADSHEET_BACKENDS:
        try:
            _convert_xlsx_with_spreadsheet_backend(client, prog_id, input_path, output_path)
            if os.path.exists(output_path):
                return output_path
            raise RuntimeError(f"转换后未找到输出文件: {output_path}")
        except Exception as e:
            errors.append((name, prog_id, e))

    raise RuntimeError(
        f"转换失败: {os.path.basename(input_path)}\n"
        "请确认电脑已安装 Microsoft Excel 或 WPS Office，并且文件未被占用。\n"
        f"详情: {_format_backend_errors(errors)}"
    )


def _convert_xlsx_with_spreadsheet_backend(client, prog_id: str, input_path: str, output_path: str) -> None:
    app = None
    workbook = None
    try:
        app = _dispatch(client, prog_id)
        _set_quiet(app)
        workbook = app.Workbooks.Open(input_path)
        workbook.ExportAsFixedFormat(XL_TYPE_PDF, output_path)
    finally:
        _safe_close(workbook, "Close", False)
        _safe_close(app, "Quit")


def _convert_with_libreoffice(input_path: str, temp_dir: str) -> str:
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError(
            "未找到 LibreOffice，无法转换文件。"
            "请确保 LibreOffice 已安装。"
        )

    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        temp_dir,
        input_path,
    ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"转换超时: {os.path.basename(input_path)}") from e
    except OSError as e:
        raise RuntimeError(f"无法启动 LibreOffice: {soffice}\n详情: {e}") from e

    if result.returncode != 0:
        details = result.stderr.strip() or result.stdout.strip() or "LibreOffice 未返回错误详情"
        raise RuntimeError(f"转换失败: {os.path.basename(input_path)}\n{details}")

    output_path = _output_path(input_path, temp_dir)
    if not os.path.exists(output_path):
        raise RuntimeError(f"转换后未找到输出文件: {output_path}")
    return output_path
