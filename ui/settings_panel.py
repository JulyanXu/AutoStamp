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
