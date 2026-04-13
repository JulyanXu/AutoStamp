import os
import platform
import shutil
import subprocess


SUPPORTED_EXTENSIONS = {".docx", ".xlsx", ".pdf"}
XL_TYPE_PDF = 0


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
        if ".docx" in office_exts and not _module_available("docx2pdf"):
            return "缺少 docx2pdf 依赖，无法转换 Word 文件。请重新安装或重新打包程序。"
        if ".xlsx" in office_exts and not _module_available("win32com.client"):
            return "缺少 pywin32/win32com 依赖，无法调用 Excel 转换文件。请重新安装或重新打包程序。"
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
            return _convert_docx_with_docx2pdf(input_path, temp_dir)
        if ext == ".xlsx":
            return _convert_xlsx_with_excel(input_path, temp_dir)

    return _convert_with_libreoffice(input_path, temp_dir)


def _output_path(input_path: str, temp_dir: str) -> str:
    base = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(temp_dir, base + ".pdf")


def _convert_docx_with_docx2pdf(input_path: str, temp_dir: str) -> str:
    output_path = _output_path(input_path, temp_dir)
    try:
        from docx2pdf import convert
    except ImportError as e:
        raise RuntimeError("缺少 docx2pdf 依赖，无法转换 Word 文件。请重新安装或重新打包程序。") from e

    try:
        convert(input_path, output_path)
    except Exception as e:
        raise RuntimeError(
            f"转换失败: {os.path.basename(input_path)}\n"
            f"请确认电脑已安装 Microsoft Word。\n详情: {e}"
        ) from e

    if not os.path.exists(output_path):
        raise RuntimeError(f"转换后未找到输出文件: {output_path}")
    return output_path


def _convert_xlsx_with_excel(input_path: str, temp_dir: str) -> str:
    output_path = _output_path(input_path, temp_dir)
    try:
        import win32com.client
    except ImportError as e:
        raise RuntimeError("缺少 pywin32/win32com 依赖，无法调用 Excel 转换文件。请重新安装或重新打包程序。") from e

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(input_path)
        workbook.ExportAsFixedFormat(XL_TYPE_PDF, output_path)
    except Exception as e:
        raise RuntimeError(
            f"转换失败: {os.path.basename(input_path)}\n"
            f"请确认电脑已安装 Microsoft Excel。\n详情: {e}"
        ) from e
    finally:
        if workbook is not None:
            workbook.Close(False)
        if excel is not None:
            excel.Quit()

    if not os.path.exists(output_path):
        raise RuntimeError(f"转换后未找到输出文件: {output_path}")
    return output_path


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
