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
