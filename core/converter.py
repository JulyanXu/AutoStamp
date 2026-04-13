import os
import platform
import tempfile


SUPPORTED_EXTENSIONS = {".docx", ".xlsx", ".pdf"}


def convert_to_pdf(input_path: str, temp_dir: str) -> str:
    ext = os.path.splitext(input_path)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"不支持的文件格式: {ext}")
    if ext == ".pdf":
        return input_path

    os.makedirs(temp_dir, exist_ok=True)

    try:
        from docx2pdf import convert
    except ImportError:
        raise RuntimeError("缺少 docx2pdf 依赖，请运行: pip install docx2pdf")

    if platform.system() == "Windows":
        # On Windows, docx2pdf uses Microsoft Word via COM — no extra install needed
        base = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(temp_dir, base + ".pdf")
        try:
            convert(input_path, output_path)
        except Exception as e:
            raise RuntimeError(
                f"转换失败: {os.path.basename(input_path)}\n"
                f"请确认电脑已安装 Microsoft Word/Excel。\n详情: {e}"
            )
    else:
        # macOS / Linux: docx2pdf uses LibreOffice
        base = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(temp_dir, base + ".pdf")
        try:
            convert(input_path, output_path)
        except Exception as e:
            raise RuntimeError(
                f"转换失败: {os.path.basename(input_path)}\n"
                f"macOS 请安装 LibreOffice: brew install --cask libreoffice\n详情: {e}"
            )

    if not os.path.exists(output_path):
        raise RuntimeError(f"转换后未找到输出文件: {output_path}")
    return output_path
