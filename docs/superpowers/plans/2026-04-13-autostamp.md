# AutoStamp Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a desktop GUI application that batch-stamps transparent PNG images onto Word/Excel/PDF files and outputs stamped PDFs.

**Architecture:** Three-layer design — `core/` for file conversion, PDF stamping, and batch processing; `ui/` for PySide6 GUI with preview, drag-to-position, and settings; `main.py` as the entry point. Config persisted to JSON.

**Tech Stack:** Python 3.11+, PySide6, PyMuPDF (fitz), Pillow, LibreOffice (headless conversion), PyInstaller

---

## File Map

| File | Responsibility |
|------|---------------|
| `main.py` | Entry point, create QApplication, load config, show MainWindow |
| `core/converter.py` | Detect LibreOffice path per OS, convert .docx/.xlsx → PDF via subprocess |
| `core/stamper.py` | Open PDF with PyMuPDF, insert stamp image at coordinates with scale/opacity |
| `core/batch.py` | QThread worker: iterate files, convert → stamp → save, emit progress signals |
| `core/config.py` | Load/save config.json (stamp path, scale, opacity, position, page mode, output dir) |
| `ui/main_window.py` | Main window: three-panel layout, menu bar, bottom action bar with progress |
| `ui/file_list_widget.py` | Left panel: file list with drag-drop, add/remove buttons, dedup |
| `ui/preview_widget.py` | Center panel: render PDF page as QPixmap, overlay draggable stamp, page nav |
| `ui/settings_panel.py` | Right panel: stamp picker, scale/opacity sliders, page mode radios, output dir |
| `tests/test_converter.py` | Tests for converter module |
| `tests/test_stamper.py` | Tests for stamper module |
| `tests/test_config.py` | Tests for config module |
| `tests/test_batch.py` | Tests for batch worker |
| `requirements.txt` | Project dependencies |

---

### Task 1: Project Scaffolding and Dependencies

**Files:**
- Create: `requirements.txt`
- Create: `main.py`
- Create: `core/__init__.py`
- Create: `ui/__init__.py`
- Create: `tests/__init__.py`
- Create: `resources/stamps/.gitkeep`

- [ ] **Step 1: Initialize git repo**

```bash
cd /Users/julyan/Documents/AutoStamp
git init
```

- [ ] **Step 2: Create requirements.txt**

```
PySide6>=6.6.0
PyMuPDF>=1.24.0
Pillow>=10.0.0
pyinstaller>=6.0.0
```

- [ ] **Step 3: Create directory structure and __init__.py files**

Create empty `core/__init__.py`, `ui/__init__.py`, `tests/__init__.py`, and `resources/stamps/.gitkeep`.

- [ ] **Step 4: Create minimal main.py**

```python
import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QLabel


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoStamp - 批量盖章工具")
        self.setMinimumSize(1000, 700)
        self.setCentralWidget(QLabel("AutoStamp 启动成功"))


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
```

- [ ] **Step 5: Install dependencies and verify**

```bash
pip install -r requirements.txt
python main.py
```

Expected: A window appears with title "AutoStamp - 批量盖章工具" and the label text. Close it manually.

- [ ] **Step 6: Create .gitignore and commit**

`.gitignore`:
```
__pycache__/
*.pyc
.venv/
venv/
dist/
build/
*.spec
config.json
```

```bash
git add .
git commit -m "feat: project scaffolding with PySide6 skeleton"
```

---

### Task 2: Config Module

**Files:**
- Create: `core/config.py`
- Create: `tests/test_config.py`

- [ ] **Step 1: Write failing tests for config**

```python
# tests/test_config.py
import json
import os
import pytest
from core.config import AppConfig


def test_default_config():
    config = AppConfig()
    assert config.stamp_path == ""
    assert config.stamp_scale == 100
    assert config.stamp_opacity == 100
    assert config.stamp_x == 50.0
    assert config.stamp_y == 50.0
    assert config.page_mode == "first"
    assert config.custom_pages == ""
    assert config.output_dir == ""


def test_save_and_load(tmp_path):
    path = tmp_path / "config.json"
    config = AppConfig()
    config.stamp_scale = 75
    config.stamp_opacity = 50
    config.stamp_x = 100.0
    config.stamp_y = 200.0
    config.page_mode = "custom"
    config.custom_pages = "1,3-5"
    config.save(str(path))

    loaded = AppConfig.load(str(path))
    assert loaded.stamp_scale == 75
    assert loaded.stamp_opacity == 50
    assert loaded.stamp_x == 100.0
    assert loaded.stamp_y == 200.0
    assert loaded.page_mode == "custom"
    assert loaded.custom_pages == "1,3-5"


def test_load_missing_file(tmp_path):
    path = tmp_path / "nonexistent.json"
    config = AppConfig.load(str(path))
    assert config.stamp_scale == 100


def test_parse_pages_first():
    config = AppConfig()
    config.page_mode = "first"
    assert config.get_page_indices(total_pages=5) == [0]


def test_parse_pages_last():
    config = AppConfig()
    config.page_mode = "last"
    assert config.get_page_indices(total_pages=5) == [4]


def test_parse_pages_all():
    config = AppConfig()
    config.page_mode = "all"
    assert config.get_page_indices(total_pages=5) == [0, 1, 2, 3, 4]


def test_parse_pages_custom():
    config = AppConfig()
    config.page_mode = "custom"
    config.custom_pages = "1,3-5"
    assert config.get_page_indices(total_pages=10) == [0, 2, 3, 4]


def test_parse_pages_custom_out_of_range():
    config = AppConfig()
    config.page_mode = "custom"
    config.custom_pages = "1,3,99"
    assert config.get_page_indices(total_pages=5) == [0, 2]
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_config.py -v
```

Expected: ModuleNotFoundError — `core.config` does not exist yet.

- [ ] **Step 3: Implement config module**

```python
# core/config.py
import json
import os
from dataclasses import dataclass, field, asdict


@dataclass
class AppConfig:
    stamp_path: str = ""
    stamp_scale: int = 100
    stamp_opacity: int = 100
    stamp_x: float = 50.0
    stamp_y: float = 50.0
    page_mode: str = "first"  # "first", "last", "all", "custom"
    custom_pages: str = ""
    output_dir: str = ""

    def save(self, path: str) -> None:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(asdict(self), f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, path: str) -> "AppConfig":
        if not os.path.exists(path):
            return cls()
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

    def get_page_indices(self, total_pages: int) -> list[int]:
        if self.page_mode == "first":
            return [0]
        if self.page_mode == "last":
            return [total_pages - 1]
        if self.page_mode == "all":
            return list(range(total_pages))
        # custom: parse "1,3-5,7"
        indices = []
        for part in self.custom_pages.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                start, end = part.split("-", 1)
                for i in range(int(start), int(end) + 1):
                    idx = i - 1  # 1-based to 0-based
                    if 0 <= idx < total_pages:
                        indices.append(idx)
            else:
                idx = int(part) - 1
                if 0 <= idx < total_pages:
                    indices.append(idx)
        return sorted(set(indices))
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_config.py -v
```

Expected: All 8 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add core/config.py tests/test_config.py
git commit -m "feat: config module with save/load and page index parsing"
```

---

### Task 3: Converter Module

**Files:**
- Create: `core/converter.py`
- Create: `tests/test_converter.py`

- [ ] **Step 1: Write failing tests for converter**

```python
# tests/test_converter.py
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
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_converter.py -v
```

Expected: ModuleNotFoundError.

- [ ] **Step 3: Implement converter module**

```python
# core/converter.py
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
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_converter.py -v
```

Expected: All 4 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add core/converter.py tests/test_converter.py
git commit -m "feat: converter module — LibreOffice detection and docx/xlsx to PDF"
```

---

### Task 4: Stamper Module

**Files:**
- Create: `core/stamper.py`
- Create: `tests/test_stamper.py`

- [ ] **Step 1: Write failing tests for stamper**

```python
# tests/test_stamper.py
import os
import fitz  # PyMuPDF
from PIL import Image
import pytest
from core.stamper import stamp_pdf


@pytest.fixture
def sample_pdf(tmp_path):
    """Create a simple 1-page PDF for testing."""
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)  # A4
    page.insert_text((100, 100), "Test Document", fontsize=24)
    path = str(tmp_path / "sample.pdf")
    doc.save(path)
    doc.close()
    return path


@pytest.fixture
def stamp_image(tmp_path):
    """Create a simple red circle PNG with transparency."""
    img = Image.new("RGBA", (100, 100), (0, 0, 0, 0))
    for x in range(100):
        for y in range(100):
            if (x - 50) ** 2 + (y - 50) ** 2 < 40 ** 2:
                img.putpixel((x, y), (255, 0, 0, 200))
    path = str(tmp_path / "stamp.png")
    img.save(path)
    return path


def test_stamp_pdf_creates_output(sample_pdf, stamp_image, tmp_path):
    output = str(tmp_path / "output.pdf")
    stamp_pdf(
        pdf_path=sample_pdf,
        stamp_path=stamp_image,
        output_path=output,
        x=100.0,
        y=100.0,
        scale=100,
        opacity=100,
        page_indices=[0],
    )
    assert os.path.exists(output)
    doc = fitz.open(output)
    assert len(doc) == 1
    doc.close()


def test_stamp_pdf_multiple_pages(tmp_path, stamp_image):
    doc = fitz.open()
    for i in range(3):
        page = doc.new_page(width=595, height=842)
        page.insert_text((100, 100), f"Page {i+1}", fontsize=24)
    pdf_path = str(tmp_path / "multi.pdf")
    doc.save(pdf_path)
    doc.close()

    output = str(tmp_path / "output.pdf")
    stamp_pdf(
        pdf_path=pdf_path,
        stamp_path=stamp_image,
        output_path=output,
        x=50.0,
        y=50.0,
        scale=100,
        opacity=80,
        page_indices=[0, 2],  # stamp pages 1 and 3
    )
    assert os.path.exists(output)
    doc = fitz.open(output)
    assert len(doc) == 3
    doc.close()


def test_stamp_pdf_with_scale(sample_pdf, stamp_image, tmp_path):
    output = str(tmp_path / "output.pdf")
    stamp_pdf(
        pdf_path=sample_pdf,
        stamp_path=stamp_image,
        output_path=output,
        x=0.0,
        y=0.0,
        scale=50,
        opacity=100,
        page_indices=[0],
    )
    assert os.path.exists(output)
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_stamper.py -v
```

Expected: ModuleNotFoundError.

- [ ] **Step 3: Implement stamper module**

```python
# core/stamper.py
import fitz  # PyMuPDF
from PIL import Image
import io


def stamp_pdf(
    pdf_path: str,
    stamp_path: str,
    output_path: str,
    x: float,
    y: float,
    scale: int,
    opacity: int,
    page_indices: list[int],
) -> None:
    """
    Stamp a PNG image onto specified pages of a PDF.

    Args:
        pdf_path: Path to source PDF.
        stamp_path: Path to stamp PNG image (transparent background).
        output_path: Path to save the stamped PDF.
        x: X position in PDF points (from left edge).
        y: Y position in PDF points (from top edge).
        scale: Scale percentage (100 = original size).
        opacity: Opacity percentage (100 = fully opaque).
        page_indices: 0-based list of page indices to stamp.
    """
    # Load and scale the stamp image
    stamp_img = Image.open(stamp_path).convert("RGBA")
    if scale != 100:
        new_w = int(stamp_img.width * scale / 100)
        new_h = int(stamp_img.height * scale / 100)
        stamp_img = stamp_img.resize((new_w, new_h), Image.LANCZOS)

    # Apply opacity
    if opacity < 100:
        r, g, b, a = stamp_img.split()
        a = a.point(lambda p: int(p * opacity / 100))
        stamp_img = Image.merge("RGBA", (r, g, b, a))

    # Convert to PNG bytes for PyMuPDF
    buf = io.BytesIO()
    stamp_img.save(buf, format="PNG")
    stamp_bytes = buf.getvalue()

    stamp_w = stamp_img.width * 72 / 96  # pixels to PDF points (assuming 96 DPI)
    stamp_h = stamp_img.height * 72 / 96

    doc = fitz.open(pdf_path)
    for page_idx in page_indices:
        if page_idx < 0 or page_idx >= len(doc):
            continue
        page = doc[page_idx]
        rect = fitz.Rect(x, y, x + stamp_w, y + stamp_h)
        page.insert_image(rect, stream=stamp_bytes)

    doc.save(output_path)
    doc.close()
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_stamper.py -v
```

Expected: All 3 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add core/stamper.py tests/test_stamper.py
git commit -m "feat: stamper module — insert PNG stamp onto PDF pages"
```

---

### Task 5: Batch Worker

**Files:**
- Create: `core/batch.py`
- Create: `tests/test_batch.py`

- [ ] **Step 1: Write failing tests for batch worker**

```python
# tests/test_batch.py
import os
import fitz
from PIL import Image
import pytest
from core.batch import BatchWorker
from core.config import AppConfig


@pytest.fixture
def stamp_image(tmp_path):
    img = Image.new("RGBA", (50, 50), (255, 0, 0, 200))
    path = str(tmp_path / "stamp.png")
    img.save(path)
    return path


@pytest.fixture
def sample_pdf(tmp_path):
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.insert_text((100, 100), "Test", fontsize=24)
    path = str(tmp_path / "test.pdf")
    doc.save(path)
    doc.close()
    return path


def test_batch_process_single_pdf(sample_pdf, stamp_image, tmp_path):
    output_dir = str(tmp_path / "output")
    config = AppConfig(
        stamp_path=stamp_image,
        stamp_scale=100,
        stamp_opacity=100,
        stamp_x=10.0,
        stamp_y=10.0,
        page_mode="first",
        output_dir=output_dir,
    )
    worker = BatchWorker(file_list=[sample_pdf], config=config)
    results = worker.process_all()
    assert results["success"] == 1
    assert results["failed"] == 0
    assert os.path.exists(os.path.join(output_dir, "test_已盖章.pdf"))


def test_batch_skips_bad_file(stamp_image, tmp_path):
    bad_file = str(tmp_path / "bad.pdf")
    with open(bad_file, "w") as f:
        f.write("not a pdf")
    output_dir = str(tmp_path / "output")
    config = AppConfig(
        stamp_path=stamp_image,
        stamp_scale=100,
        stamp_opacity=100,
        stamp_x=10.0,
        stamp_y=10.0,
        page_mode="first",
        output_dir=output_dir,
    )
    worker = BatchWorker(file_list=[bad_file], config=config)
    results = worker.process_all()
    assert results["failed"] == 1
    assert len(results["errors"]) == 1
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_batch.py -v
```

Expected: ModuleNotFoundError.

- [ ] **Step 3: Implement batch worker**

```python
# core/batch.py
import os
import tempfile
import shutil
from PySide6.QtCore import QThread, Signal
from core.config import AppConfig
from core.converter import convert_to_pdf
from core.stamper import stamp_pdf


class BatchWorker(QThread):
    progress = Signal(int, int, str)  # current, total, filename
    finished_all = Signal(dict)  # results summary

    def __init__(self, file_list: list[str], config: AppConfig, parent=None):
        super().__init__(parent)
        self.file_list = file_list
        self.config = config
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self):
        results = self.process_all()
        self.finished_all.emit(results)

    def process_all(self) -> dict:
        results = {"success": 0, "failed": 0, "errors": []}
        total = len(self.file_list)
        temp_dir = tempfile.mkdtemp(prefix="autostamp_")

        try:
            for i, file_path in enumerate(self.file_list):
                if self._cancelled:
                    break
                filename = os.path.basename(file_path)
                self.progress.emit(i + 1, total, filename)
                try:
                    self._process_one(file_path, temp_dir)
                    results["success"] += 1
                except Exception as e:
                    results["failed"] += 1
                    results["errors"].append((filename, str(e)))
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

        return results

    def _process_one(self, file_path: str, temp_dir: str) -> None:
        # Step 1: Convert to PDF if needed
        pdf_path = convert_to_pdf(file_path, temp_dir)

        # Step 2: Determine output path
        output_dir = self.config.output_dir
        if not output_dir:
            output_dir = os.path.join(os.path.dirname(file_path), "output")
        os.makedirs(output_dir, exist_ok=True)

        base = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(output_dir, f"{base}_已盖章.pdf")

        # Step 3: Get page indices
        import fitz
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        doc.close()
        page_indices = self.config.get_page_indices(total_pages)

        # Step 4: Stamp
        stamp_pdf(
            pdf_path=pdf_path,
            stamp_path=self.config.stamp_path,
            output_path=output_path,
            x=self.config.stamp_x,
            y=self.config.stamp_y,
            scale=self.config.stamp_scale,
            opacity=self.config.stamp_opacity,
            page_indices=page_indices,
        )
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_batch.py -v
```

Expected: All 2 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add core/batch.py tests/test_batch.py
git commit -m "feat: batch worker — threaded file processing with progress signals"
```

---

### Task 6: Settings Panel (Right Panel)

**Files:**
- Create: `ui/settings_panel.py`

- [ ] **Step 1: Implement settings panel**

```python
# ui/settings_panel.py
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QSlider, QRadioButton, QButtonGroup, QLineEdit, QFileDialog,
    QGroupBox,
)
from PySide6.QtCore import Qt, Signal
from core.config import AppConfig


class SettingsPanel(QWidget):
    stamp_changed = Signal(str)  # stamp image path
    settings_changed = Signal()  # any setting changed

    def __init__(self, config: AppConfig, parent=None):
        super().__init__(parent)
        self.config = config
        self.setFixedWidth(240)
        self._build_ui()
        self._load_from_config()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignTop)

        # Stamp image selector
        stamp_group = QGroupBox("印章图片")
        stamp_layout = QVBoxLayout(stamp_group)
        self.stamp_label = QLabel("未选择")
        self.stamp_label.setWordWrap(True)
        stamp_btn = QPushButton("选择印章图片...")
        stamp_btn.clicked.connect(self._select_stamp)
        stamp_layout.addWidget(self.stamp_label)
        stamp_layout.addWidget(stamp_btn)
        layout.addWidget(stamp_group)

        # Scale slider
        scale_group = QGroupBox("大小")
        scale_layout = QVBoxLayout(scale_group)
        self.scale_slider = QSlider(Qt.Horizontal)
        self.scale_slider.setRange(10, 200)
        self.scale_slider.setValue(100)
        self.scale_value_label = QLabel("100%")
        self.scale_slider.valueChanged.connect(self._on_scale_changed)
        scale_layout.addWidget(self.scale_slider)
        scale_layout.addWidget(self.scale_value_label)
        layout.addWidget(scale_group)

        # Opacity slider
        opacity_group = QGroupBox("透明度")
        opacity_layout = QVBoxLayout(opacity_group)
        self.opacity_slider = QSlider(Qt.Horizontal)
        self.opacity_slider.setRange(10, 100)
        self.opacity_slider.setValue(100)
        self.opacity_value_label = QLabel("100%")
        self.opacity_slider.valueChanged.connect(self._on_opacity_changed)
        opacity_layout.addWidget(self.opacity_slider)
        opacity_layout.addWidget(self.opacity_value_label)
        layout.addWidget(opacity_group)

        # Page mode
        page_group = QGroupBox("盖章页码")
        page_layout = QVBoxLayout(page_group)
        self.page_button_group = QButtonGroup(self)
        self.radio_first = QRadioButton("仅第一页")
        self.radio_last = QRadioButton("仅最后一页")
        self.radio_all = QRadioButton("所有页")
        self.radio_custom = QRadioButton("自定义:")
        self.custom_pages_input = QLineEdit()
        self.custom_pages_input.setPlaceholderText("如: 1,3-5")
        self.custom_pages_input.setEnabled(False)

        self.page_button_group.addButton(self.radio_first)
        self.page_button_group.addButton(self.radio_last)
        self.page_button_group.addButton(self.radio_all)
        self.page_button_group.addButton(self.radio_custom)
        self.radio_first.setChecked(True)

        self.radio_custom.toggled.connect(self.custom_pages_input.setEnabled)
        self.page_button_group.buttonToggled.connect(lambda: self._on_page_mode_changed())
        self.custom_pages_input.textChanged.connect(self._on_custom_pages_changed)

        page_layout.addWidget(self.radio_first)
        page_layout.addWidget(self.radio_last)
        page_layout.addWidget(self.radio_all)
        page_layout.addWidget(self.radio_custom)
        page_layout.addWidget(self.custom_pages_input)
        layout.addWidget(page_group)

        # Output directory
        output_group = QGroupBox("输出目录")
        output_layout = QVBoxLayout(output_group)
        self.output_label = QLabel("默认: 源文件目录/output/")
        self.output_label.setWordWrap(True)
        output_btn = QPushButton("选择输出目录...")
        output_btn.clicked.connect(self._select_output_dir)
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(output_btn)
        layout.addWidget(output_group)

        layout.addStretch()

    def _load_from_config(self):
        if self.config.stamp_path:
            self.stamp_label.setText(self.config.stamp_path)
        self.scale_slider.setValue(self.config.stamp_scale)
        self.opacity_slider.setValue(self.config.stamp_opacity)
        mode_map = {
            "first": self.radio_first,
            "last": self.radio_last,
            "all": self.radio_all,
            "custom": self.radio_custom,
        }
        radio = mode_map.get(self.config.page_mode, self.radio_first)
        radio.setChecked(True)
        self.custom_pages_input.setText(self.config.custom_pages)
        if self.config.output_dir:
            self.output_label.setText(self.config.output_dir)

    def _select_stamp(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择印章图片", "", "图片文件 (*.png *.jpg *.jpeg *.bmp)"
        )
        if path:
            self.config.stamp_path = path
            self.stamp_label.setText(path)
            self.stamp_changed.emit(path)
            self.settings_changed.emit()

    def _on_scale_changed(self, value):
        self.scale_value_label.setText(f"{value}%")
        self.config.stamp_scale = value
        self.settings_changed.emit()

    def _on_opacity_changed(self, value):
        self.opacity_value_label.setText(f"{value}%")
        self.config.stamp_opacity = value
        self.settings_changed.emit()

    def _on_page_mode_changed(self):
        if self.radio_first.isChecked():
            self.config.page_mode = "first"
        elif self.radio_last.isChecked():
            self.config.page_mode = "last"
        elif self.radio_all.isChecked():
            self.config.page_mode = "all"
        elif self.radio_custom.isChecked():
            self.config.page_mode = "custom"
        self.settings_changed.emit()

    def _on_custom_pages_changed(self, text):
        self.config.custom_pages = text
        self.settings_changed.emit()

    def _select_output_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if path:
            self.config.output_dir = path
            self.output_label.setText(path)
            self.settings_changed.emit()
```

- [ ] **Step 2: Verify it imports without errors**

```bash
python -c "from ui.settings_panel import SettingsPanel; print('OK')"
```

Expected: `OK`

- [ ] **Step 3: Commit**

```bash
git add ui/settings_panel.py
git commit -m "feat: settings panel — stamp picker, sliders, page mode, output dir"
```

---

### Task 7: File List Widget (Left Panel)

**Files:**
- Create: `ui/file_list_widget.py`

- [ ] **Step 1: Implement file list widget**

```python
# ui/file_list_widget.py
import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QListWidget, QListWidgetItem,
    QFileDialog, QMessageBox,
)
from PySide6.QtCore import Signal, Qt


SUPPORTED_EXTENSIONS = {".docx", ".xlsx", ".pdf"}


class FileListWidget(QWidget):
    file_selected = Signal(str)  # emitted when user clicks a file in the list

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedWidth(200)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.ExtendedSelection)
        self.list_widget.currentItemChanged.connect(self._on_item_changed)
        self.setAcceptDrops(True)

        add_btn = QPushButton("添加文件...")
        add_btn.clicked.connect(self._add_files)

        remove_btn = QPushButton("删除选中")
        remove_btn.clicked.connect(self._remove_selected)

        clear_btn = QPushButton("清空列表")
        clear_btn.clicked.connect(self._clear_list)

        layout.addWidget(self.list_widget)
        layout.addWidget(add_btn)
        layout.addWidget(remove_btn)
        layout.addWidget(clear_btn)

    def _add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择文件",
            "",
            "支持的文件 (*.docx *.xlsx *.pdf);;所有文件 (*)",
        )
        self._insert_files(files)

    def _insert_files(self, paths: list[str]):
        existing = set(self._all_paths())
        added = 0
        skipped_format = []
        for path in paths:
            ext = os.path.splitext(path)[1].lower()
            if ext not in SUPPORTED_EXTENSIONS:
                skipped_format.append(os.path.basename(path))
                continue
            if path in existing:
                continue
            item = QListWidgetItem(os.path.basename(path))
            item.setData(Qt.UserRole, path)
            item.setToolTip(path)
            self.list_widget.addItem(item)
            existing.add(path)
            added += 1
        if skipped_format:
            QMessageBox.warning(
                self,
                "格式不支持",
                f"以下文件格式不支持，已跳过:\n" + "\n".join(skipped_format),
            )

    def _remove_selected(self):
        for item in self.list_widget.selectedItems():
            self.list_widget.takeItem(self.list_widget.row(item))

    def _clear_list(self):
        self.list_widget.clear()

    def _on_item_changed(self, current, _previous):
        if current:
            self.file_selected.emit(current.data(Qt.UserRole))

    def _all_paths(self) -> list[str]:
        paths = []
        for i in range(self.list_widget.count()):
            paths.append(self.list_widget.item(i).data(Qt.UserRole))
        return paths

    def get_file_list(self) -> list[str]:
        return self._all_paths()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
        self._insert_files(paths)
```

- [ ] **Step 2: Verify it imports without errors**

```bash
python -c "from ui.file_list_widget import FileListWidget; print('OK')"
```

Expected: `OK`

- [ ] **Step 3: Commit**

```bash
git add ui/file_list_widget.py
git commit -m "feat: file list widget — add, remove, drag-drop, dedup"
```

---

### Task 8: Preview Widget (Center Panel)

**Files:**
- Create: `ui/preview_widget.py`

- [ ] **Step 1: Implement preview widget**

```python
# ui/preview_widget.py
import os
import fitz  # PyMuPDF
from PIL import Image
from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel
from PySide6.QtGui import QPixmap, QImage, QPainter, QMouseEvent
from PySide6.QtCore import Qt, Signal, QPoint, QRect


class PreviewWidget(QWidget):
    stamp_position_changed = Signal(float, float)  # PDF-point coords

    def __init__(self, parent=None):
        super().__init__(parent)
        self._pdf_doc = None
        self._current_page = 0
        self._total_pages = 0
        self._page_pixmap = None  # QPixmap of the rendered page
        self._stamp_pixmap = None  # QPixmap of the stamp image
        self._stamp_scale = 100
        self._stamp_opacity = 100
        self._stamp_pos = QPoint(50, 50)  # position in widget coords
        self._dragging = False
        self._drag_offset = QPoint(0, 0)
        self._page_rect = QRect()  # where the page is drawn in the widget
        self._pdf_page_width = 0.0
        self._pdf_page_height = 0.0
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        self._canvas = _Canvas(self)
        layout.addWidget(self._canvas, 1)

        nav_layout = QHBoxLayout()
        self._prev_btn = QPushButton("<")
        self._prev_btn.setFixedWidth(40)
        self._prev_btn.clicked.connect(self._prev_page)
        self._next_btn = QPushButton(">")
        self._next_btn.setFixedWidth(40)
        self._next_btn.clicked.connect(self._next_page)
        self._page_label = QLabel("无文件")
        self._page_label.setAlignment(Qt.AlignCenter)
        nav_layout.addWidget(self._prev_btn)
        nav_layout.addWidget(self._page_label, 1)
        nav_layout.addWidget(self._next_btn)
        layout.addLayout(nav_layout)

    def load_pdf(self, pdf_path: str):
        if self._pdf_doc:
            self._pdf_doc.close()
        self._pdf_doc = fitz.open(pdf_path)
        self._total_pages = len(self._pdf_doc)
        self._current_page = 0
        self._render_page()

    def set_stamp_image(self, stamp_path: str):
        if not stamp_path or not os.path.exists(stamp_path):
            self._stamp_pixmap = None
            self._canvas.update()
            return
        self._stamp_pixmap = QPixmap(stamp_path)
        self._canvas.update()

    def set_stamp_scale(self, scale: int):
        self._stamp_scale = scale
        self._canvas.update()

    def set_stamp_opacity(self, opacity: int):
        self._stamp_opacity = opacity
        self._canvas.update()

    def set_stamp_position_from_config(self, x: float, y: float):
        """Set stamp position from PDF-point coordinates."""
        if self._page_rect.width() > 0 and self._pdf_page_width > 0:
            ratio = self._page_rect.width() / self._pdf_page_width
            self._stamp_pos = QPoint(
                int(self._page_rect.x() + x * ratio),
                int(self._page_rect.y() + y * ratio),
            )
        else:
            self._stamp_pos = QPoint(int(x), int(y))
        self._canvas.update()

    def _render_page(self):
        if not self._pdf_doc or self._current_page >= self._total_pages:
            return
        page = self._pdf_doc[self._current_page]
        self._pdf_page_width = page.rect.width
        self._pdf_page_height = page.rect.height
        zoom = 2.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        self._page_pixmap = QPixmap.fromImage(img)
        self._page_label.setText(f"{self._current_page + 1} / {self._total_pages}")
        self._canvas.update()

    def _prev_page(self):
        if self._current_page > 0:
            self._current_page -= 1
            self._render_page()

    def _next_page(self):
        if self._current_page < self._total_pages - 1:
            self._current_page += 1
            self._render_page()

    def _get_scaled_stamp_pixmap(self) -> QPixmap | None:
        if not self._stamp_pixmap:
            return None
        if self._stamp_scale == 100 and self._page_rect.width() > 0:
            # Scale stamp relative to page display size
            ratio = self._page_rect.width() / self._pdf_page_width if self._pdf_page_width else 1
            w = int(self._stamp_pixmap.width() * 72 / 96 * ratio)
            h = int(self._stamp_pixmap.height() * 72 / 96 * ratio)
            return self._stamp_pixmap.scaled(w, h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        elif self._page_rect.width() > 0 and self._pdf_page_width > 0:
            ratio = self._page_rect.width() / self._pdf_page_width
            w = int(self._stamp_pixmap.width() * 72 / 96 * ratio * self._stamp_scale / 100)
            h = int(self._stamp_pixmap.height() * 72 / 96 * ratio * self._stamp_scale / 100)
            return self._stamp_pixmap.scaled(w, h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        return self._stamp_pixmap

    def _to_pdf_coords(self, widget_pos: QPoint) -> tuple[float, float]:
        if self._page_rect.width() <= 0 or self._pdf_page_width <= 0:
            return float(widget_pos.x()), float(widget_pos.y())
        ratio = self._pdf_page_width / self._page_rect.width()
        pdf_x = (widget_pos.x() - self._page_rect.x()) * ratio
        pdf_y = (widget_pos.y() - self._page_rect.y()) * ratio
        return pdf_x, pdf_y

    def close_doc(self):
        if self._pdf_doc:
            self._pdf_doc.close()
            self._pdf_doc = None


class _Canvas(QWidget):
    """Inner widget that handles painting and mouse events for the preview."""

    def __init__(self, preview: PreviewWidget):
        super().__init__(preview)
        self._preview = preview
        self.setMouseTracking(True)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), Qt.lightGray)

        p = self._preview
        if p._page_pixmap:
            # Scale page to fit widget while keeping aspect ratio
            scaled = p._page_pixmap.scaled(
                self.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation
            )
            x = (self.width() - scaled.width()) // 2
            y = (self.height() - scaled.height()) // 2
            p._page_rect = QRect(x, y, scaled.width(), scaled.height())
            painter.drawPixmap(x, y, scaled)

            # Draw stamp overlay
            stamp = p._get_scaled_stamp_pixmap()
            if stamp:
                painter.setOpacity(p._stamp_opacity / 100.0)
                painter.drawPixmap(p._stamp_pos, stamp)
                painter.setOpacity(1.0)

        painter.end()

    def mousePressEvent(self, event: QMouseEvent):
        p = self._preview
        stamp = p._get_scaled_stamp_pixmap()
        if not stamp or event.button() != Qt.LeftButton:
            return
        stamp_rect = QRect(p._stamp_pos, stamp.size())
        if stamp_rect.contains(event.pos()):
            p._dragging = True
            p._drag_offset = event.pos() - p._stamp_pos

    def mouseMoveEvent(self, event: QMouseEvent):
        p = self._preview
        if p._dragging:
            p._stamp_pos = event.pos() - p._drag_offset
            self.update()

    def mouseReleaseEvent(self, event: QMouseEvent):
        p = self._preview
        if p._dragging:
            p._dragging = False
            pdf_x, pdf_y = p._to_pdf_coords(p._stamp_pos)
            p.stamp_position_changed.emit(pdf_x, pdf_y)
```

- [ ] **Step 2: Verify it imports without errors**

```bash
python -c "from ui.preview_widget import PreviewWidget; print('OK')"
```

Expected: `OK`

- [ ] **Step 3: Commit**

```bash
git add ui/preview_widget.py
git commit -m "feat: preview widget — PDF page rendering with draggable stamp overlay"
```

---

### Task 9: Main Window — Assemble Everything

**Files:**
- Create: `ui/main_window.py`
- Modify: `main.py`

- [ ] **Step 1: Implement main window**

```python
# ui/main_window.py
import os
import tempfile
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QPushButton,
    QProgressBar, QMessageBox, QLabel,
)
from PySide6.QtCore import Qt
from core.config import AppConfig
from core.converter import convert_to_pdf, find_libreoffice
from core.batch import BatchWorker
from ui.file_list_widget import FileListWidget
from ui.preview_widget import PreviewWidget
from ui.settings_panel import SettingsPanel


CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "config.json")
CONFIG_PATH = os.path.normpath(CONFIG_PATH)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoStamp - 批量盖章工具")
        self.setMinimumSize(1100, 750)
        self.config = AppConfig.load(CONFIG_PATH)
        self._batch_worker = None
        self._temp_dir = tempfile.mkdtemp(prefix="autostamp_preview_")
        self._build_ui()
        self._connect_signals()
        self._load_initial_state()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        # Three-panel layout
        panels = QHBoxLayout()

        # Left: file list
        self.file_list = FileListWidget()
        panels.addWidget(self.file_list)

        # Center: preview
        self.preview = PreviewWidget()
        panels.addWidget(self.preview, 1)

        # Right: settings
        self.settings_panel = SettingsPanel(self.config)
        panels.addWidget(self.settings_panel)

        main_layout.addLayout(panels, 1)

        # Bottom: action bar
        bottom = QHBoxLayout()
        self.start_btn = QPushButton("开始盖章")
        self.start_btn.setFixedHeight(40)
        self.start_btn.setStyleSheet("font-size: 16px; font-weight: bold;")
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setFixedHeight(40)
        self.cancel_btn.setVisible(False)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_label = QLabel("")
        bottom.addWidget(self.start_btn)
        bottom.addWidget(self.cancel_btn)
        bottom.addWidget(self.progress_bar, 1)
        bottom.addWidget(self.progress_label)
        main_layout.addLayout(bottom)

    def _connect_signals(self):
        self.file_list.file_selected.connect(self._on_file_selected)
        self.settings_panel.stamp_changed.connect(self._on_stamp_changed)
        self.settings_panel.settings_changed.connect(self._on_settings_changed)
        self.preview.stamp_position_changed.connect(self._on_stamp_position_changed)
        self.start_btn.clicked.connect(self._start_batch)
        self.cancel_btn.clicked.connect(self._cancel_batch)

    def _load_initial_state(self):
        if self.config.stamp_path and os.path.exists(self.config.stamp_path):
            self.preview.set_stamp_image(self.config.stamp_path)
        self.preview.set_stamp_scale(self.config.stamp_scale)
        self.preview.set_stamp_opacity(self.config.stamp_opacity)

    def _on_file_selected(self, file_path: str):
        try:
            pdf_path = convert_to_pdf(file_path, self._temp_dir)
            self.preview.load_pdf(pdf_path)
            self.preview.set_stamp_position_from_config(
                self.config.stamp_x, self.config.stamp_y
            )
        except Exception as e:
            QMessageBox.warning(self, "预览失败", f"无法预览文件:\n{e}")

    def _on_stamp_changed(self, path: str):
        self.preview.set_stamp_image(path)

    def _on_settings_changed(self):
        self.preview.set_stamp_scale(self.config.stamp_scale)
        self.preview.set_stamp_opacity(self.config.stamp_opacity)
        self.config.save(CONFIG_PATH)

    def _on_stamp_position_changed(self, x: float, y: float):
        self.config.stamp_x = x
        self.config.stamp_y = y
        self.config.save(CONFIG_PATH)

    def _start_batch(self):
        file_list = self.file_list.get_file_list()
        if not file_list:
            QMessageBox.warning(self, "提示", "请先添加文件。")
            return
        if not self.config.stamp_path or not os.path.exists(self.config.stamp_path):
            QMessageBox.warning(self, "提示", "请先选择印章图片。")
            return

        lo = find_libreoffice()
        has_office_files = any(
            os.path.splitext(f)[1].lower() in (".docx", ".xlsx") for f in file_list
        )
        if has_office_files and not lo:
            QMessageBox.critical(
                self, "错误",
                "文件列表中包含 Word/Excel 文件，但未找到 LibreOffice。\n"
                "请确保 LibreOffice 已安装或放置在程序目录下的 libreoffice/ 文件夹中。",
            )
            return

        self.start_btn.setVisible(False)
        self.cancel_btn.setVisible(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(file_list))
        self.progress_bar.setValue(0)

        self._batch_worker = BatchWorker(file_list, self.config)
        self._batch_worker.progress.connect(self._on_batch_progress)
        self._batch_worker.finished_all.connect(self._on_batch_finished)
        self._batch_worker.start()

    def _cancel_batch(self):
        if self._batch_worker:
            self._batch_worker.cancel()

    def _on_batch_progress(self, current: int, total: int, filename: str):
        self.progress_bar.setValue(current)
        self.progress_label.setText(f"{current}/{total} - {filename}")

    def _on_batch_finished(self, results: dict):
        self.start_btn.setVisible(True)
        self.cancel_btn.setVisible(False)
        self.progress_bar.setVisible(False)
        self.progress_label.setText("")

        msg = f"处理完成！\n成功: {results['success']} 个文件\n失败: {results['failed']} 个文件"
        if results["errors"]:
            msg += "\n\n失败文件:\n"
            for fname, err in results["errors"]:
                msg += f"  - {fname}: {err}\n"
        QMessageBox.information(self, "完成", msg)

    def closeEvent(self, event):
        self.config.save(CONFIG_PATH)
        self.preview.close_doc()
        import shutil
        shutil.rmtree(self._temp_dir, ignore_errors=True)
        event.accept()
```

- [ ] **Step 2: Update main.py to use MainWindow**

```python
# main.py
import sys
from PySide6.QtWidgets import QApplication
from ui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Run the application manually to verify**

```bash
python main.py
```

Expected: The full three-panel window opens. Left panel has file list with add/remove buttons. Center panel shows gray preview area. Right panel has stamp settings. Bottom has "开始盖章" button. Close manually.

- [ ] **Step 4: Commit**

```bash
git add ui/main_window.py main.py
git commit -m "feat: main window — assemble file list, preview, settings, and batch controls"
```

---

### Task 10: Integration Test — Full Workflow

**Files:**
- Create: `tests/test_integration.py`

- [ ] **Step 1: Write integration test**

```python
# tests/test_integration.py
import os
import fitz
from PIL import Image
import pytest
from core.config import AppConfig
from core.batch import BatchWorker


@pytest.fixture
def stamp_image(tmp_path):
    img = Image.new("RGBA", (80, 80), (0, 0, 0, 0))
    for x in range(80):
        for y in range(80):
            if (x - 40) ** 2 + (y - 40) ** 2 < 35 ** 2:
                img.putpixel((x, y), (255, 0, 0, 180))
    path = str(tmp_path / "stamp.png")
    img.save(path)
    return path


@pytest.fixture
def multi_page_pdf(tmp_path):
    doc = fitz.open()
    for i in range(5):
        page = doc.new_page(width=595, height=842)
        page.insert_text((100, 100), f"Page {i + 1}", fontsize=24)
    path = str(tmp_path / "multi.pdf")
    doc.save(path)
    doc.close()
    return path


def test_full_workflow_first_page(multi_page_pdf, stamp_image, tmp_path):
    output_dir = str(tmp_path / "output")
    config = AppConfig(
        stamp_path=stamp_image,
        stamp_scale=80,
        stamp_opacity=70,
        stamp_x=200.0,
        stamp_y=300.0,
        page_mode="first",
        output_dir=output_dir,
    )
    worker = BatchWorker(file_list=[multi_page_pdf], config=config)
    results = worker.process_all()
    assert results["success"] == 1
    assert results["failed"] == 0
    output_file = os.path.join(output_dir, "multi_已盖章.pdf")
    assert os.path.exists(output_file)
    doc = fitz.open(output_file)
    assert len(doc) == 5
    doc.close()


def test_full_workflow_custom_pages(multi_page_pdf, stamp_image, tmp_path):
    output_dir = str(tmp_path / "output")
    config = AppConfig(
        stamp_path=stamp_image,
        stamp_scale=100,
        stamp_opacity=100,
        stamp_x=50.0,
        stamp_y=50.0,
        page_mode="custom",
        custom_pages="1,3,5",
        output_dir=output_dir,
    )
    worker = BatchWorker(file_list=[multi_page_pdf], config=config)
    results = worker.process_all()
    assert results["success"] == 1
    output_file = os.path.join(output_dir, "multi_已盖章.pdf")
    assert os.path.exists(output_file)


def test_full_workflow_all_pages(multi_page_pdf, stamp_image, tmp_path):
    output_dir = str(tmp_path / "output")
    config = AppConfig(
        stamp_path=stamp_image,
        stamp_scale=150,
        stamp_opacity=50,
        stamp_x=0.0,
        stamp_y=0.0,
        page_mode="all",
        output_dir=output_dir,
    )
    worker = BatchWorker(file_list=[multi_page_pdf], config=config)
    results = worker.process_all()
    assert results["success"] == 1


def test_multiple_files(stamp_image, tmp_path):
    output_dir = str(tmp_path / "output")
    pdfs = []
    for i in range(3):
        doc = fitz.open()
        page = doc.new_page(width=595, height=842)
        page.insert_text((100, 100), f"Doc {i + 1}", fontsize=24)
        path = str(tmp_path / f"doc{i + 1}.pdf")
        doc.save(path)
        doc.close()
        pdfs.append(path)

    config = AppConfig(
        stamp_path=stamp_image,
        stamp_scale=100,
        stamp_opacity=100,
        stamp_x=100.0,
        stamp_y=100.0,
        page_mode="first",
        output_dir=output_dir,
    )
    worker = BatchWorker(file_list=pdfs, config=config)
    results = worker.process_all()
    assert results["success"] == 3
    assert results["failed"] == 0
    for i in range(3):
        assert os.path.exists(os.path.join(output_dir, f"doc{i + 1}_已盖章.pdf"))
```

- [ ] **Step 2: Run all tests**

```bash
python -m pytest tests/ -v
```

Expected: All tests PASS (config: 8, converter: 4, stamper: 3, batch: 2, integration: 4 = 21 tests).

- [ ] **Step 3: Commit**

```bash
git add tests/test_integration.py
git commit -m "test: integration tests for full stamp workflow"
```

---

### Task 11: PyInstaller Spec and Build Script

**Files:**
- Create: `build.py`

- [ ] **Step 1: Create build script**

```python
# build.py
"""
Build script for AutoStamp.
Run on Windows to create distributable package.
On macOS, creates a local build for testing.

Usage: python build.py
"""
import os
import platform
import subprocess
import sys


def main():
    system = platform.system()
    app_name = "AutoStamp"

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", app_name,
        "--windowed",
        "--noconfirm",
        "--add-data", f"resources{os.pathsep}resources",
        "main.py",
    ]

    print(f"Building {app_name} for {system}...")
    subprocess.run(cmd, check=True)

    dist_dir = os.path.join("dist", app_name)
    print(f"\nBuild complete: {dist_dir}")

    if system == "Windows":
        print("\nNext steps:")
        print(f"1. Copy LibreOffice Portable to: {dist_dir}/libreoffice/")
        print(f"2. Zip the entire '{dist_dir}' folder for distribution")
        print(f"3. Users extract the zip and run {app_name}.exe")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Verify build script syntax**

```bash
python -c "import ast; ast.parse(open('build.py').read()); print('OK')"
```

Expected: `OK`

- [ ] **Step 3: Commit**

```bash
git add build.py
git commit -m "feat: PyInstaller build script for Windows distribution"
```

---

## Self-Review Checklist

- **Spec coverage**: All spec requirements mapped to tasks — config (T2), converter (T3), stamper (T4), batch (T5), settings panel (T6), file list (T7), preview with drag (T8), main window assembly (T9), integration tests (T10), build/packaging (T11).
- **Placeholder scan**: No TBD/TODO. All code blocks are complete.
- **Type consistency**: `AppConfig` fields, method names (`get_page_indices`, `convert_to_pdf`, `stamp_pdf`, `process_all`), and signal signatures are consistent across all tasks.
