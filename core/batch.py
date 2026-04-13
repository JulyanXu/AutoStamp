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
