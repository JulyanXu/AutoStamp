import os
import platform
import subprocess
import shutil


SUPPORTED_EXTENSIONS = {".docx", ".xlsx", ".pdf"}


def find_libreoffice() -> str | None:
    system = platform.system()
    if system == "Darwin":
        path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(path):
            return path
    elif system == "Windows":
        # Look for portable version next to the executable
        app_dir = os.path.dirname(os.path.abspath(__file__))
        portable = os.path.join(app_dir, "..", "libreoffice", "App", "libreoffice", "program", "soffice.exe")
        portable = os.path.normpath(portable)
        if os.path.exists(portable):
            return portable
        # Fallback: standard install paths
        for prog_dir in [os.environ.get("PROGRAMFILES", ""), os.environ.get("PROGRAMFILES(X86)", "")]:
            if prog_dir:
                path = os.path.join(prog_dir, "LibreOffice", "program", "soffice.exe")
                if os.path.exists(path):
                    return path
    # Fallback: check PATH
    soffice = shutil.which("soffice")
    return soffice


def convert_to_pdf(input_path: str, temp_dir: str) -> str:
    ext = os.path.splitext(input_path)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"不支持的文件格式: {ext}")
    if ext == ".pdf":
        return input_path

    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError("未找到 LibreOffice，无法转换文件。请确保 LibreOffice 已安装或放置在程序目录下。")

    os.makedirs(temp_dir, exist_ok=True)
    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--convert-to", "pdf",
        "--outdir", temp_dir,
        input_path,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if result.returncode != 0:
        raise RuntimeError(f"转换失败: {input_path}\n{result.stderr}")

    base = os.path.splitext(os.path.basename(input_path))[0]
    output_path = os.path.join(temp_dir, base + ".pdf")
    if not os.path.exists(output_path):
        raise RuntimeError(f"转换后未找到输出文件: {output_path}")
    return output_path
