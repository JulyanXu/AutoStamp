import os
import tempfile
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QPushButton,
    QProgressBar, QMessageBox, QLabel,
)
from PySide6.QtCore import Qt
from core.config import AppConfig
from core.converter import convert_to_pdf, conversion_prerequisite_error
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

        prerequisite_error = conversion_prerequisite_error(file_list)
        if prerequisite_error:
            QMessageBox.critical(self, "错误", prerequisite_error)
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
