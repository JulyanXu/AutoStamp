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
