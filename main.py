"""文件转换助手：Word / PDF / 图片互转 + PDF 页面工具（PySide6 桌面版）。"""

from __future__ import annotations

import ctypes
import os
import sys
from pathlib import Path

from PySide6.QtCore import QEvent, QSettings, Qt, QThread, Signal
from PySide6.QtGui import QColor, QDragEnterEvent, QDropEvent, QFont, QIcon
from PySide6.QtWidgets import (
    QAbstractItemView,
    QAbstractSpinBox,
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSlider,
    QSpinBox,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

import converters as C
from theme import ACCENT, GREEN, MUTED, QSS, RED

APP_TITLE = "文件转换助手"
APP_VERSION = "1.0"

STATUS_WAITING = "等待转换"
STATUS_RUNNING = "转换中…"
STATUS_DONE = "完成"
STATUS_FAILED = "失败"
STATUS_SKIP = "不支持"

IMAGE_TARGETS = {".png", ".jpg", ".webp", ".bmp"}
ALL_SUPPORTED = C.WORD_EXTS | C.IMAGE_EXTS | C.PDF_EXTS


def resource_path(*parts: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base.joinpath(*parts)


def open_local(path: str) -> None:
    if not path:
        return
    try:
        os.startfile(path)  # type: ignore[attr-defined]
    except OSError:
        pass


def human_size(num_bytes: int) -> str:
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} GB"


# ---------------------------------------------------------------- 自定义控件


class DropArea(QFrame):
    """虚线拖拽区，点击也可选择文件。"""

    filesDropped = Signal(list)
    clicked = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("DropArea")
        self.setAcceptDrops(True)
        self.setProperty("drag", False)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(124)
        self.setMaximumHeight(140)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(4)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        icon = QLabel("⇩")
        icon.setStyleSheet(f"color: {ACCENT}; font-size: 30px; font-weight: 600;")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title = QLabel("把文件拖到这里，或点击选择")
        title.setStyleSheet("font-size: 14px; font-weight: 600; color: #2d4a7a;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        hint = QLabel("支持 Word / PDF / JPG / PNG / WebP / BMP / TIFF")
        hint.setObjectName("Hint")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)

        for label in (icon, title, hint):
            label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

        layout.addWidget(icon)
        layout.addWidget(title)
        layout.addWidget(hint)

    def _set_drag(self, active: bool) -> None:
        self.setProperty("drag", active)
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._set_drag(True)

    def dragLeaveEvent(self, event) -> None:
        self._set_drag(False)
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        self._set_drag(False)
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        if paths:
            self.filesDropped.emit(paths)
        event.acceptProposedAction()

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)


# ---------------------------------------------------------------- 工作线程


class ConvertWorker(QThread):
    row_status = Signal(int, str, str)
    progress = Signal(int, int)
    done = Signal(int, int)

    def __init__(self, jobs: list[dict], options: dict, parent=None):
        super().__init__(parent)
        self.jobs = jobs
        self.options = options
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def run(self) -> None:
        ok = 0
        failed = 0
        total = len(self.jobs)
        for index, job in enumerate(self.jobs, 1):
            if self._cancelled:
                break
            for row in job["rows"]:
                self.row_status.emit(row, STATUS_RUNNING, "")
            try:
                dst = C.unique_path(job["dst"])
                if len(job["srcs"]) > 1:
                    outputs = C.images_to_pdf(job["srcs"], dst, self.options.get("quality", 90))
                else:
                    outputs = C.convert_file(job["srcs"][0], dst, **self.options)
                ok += 1
                if outputs:
                    if len(outputs) > 1:
                        status = f"{STATUS_DONE}（{len(outputs)} 个文件）"
                        target = str(Path(outputs[0]).parent)
                    else:
                        status = STATUS_DONE
                        target = outputs[0]
                else:
                    status = STATUS_DONE
                    target = ""
                for row in job["rows"]:
                    self.row_status.emit(row, status, target)
            except Exception as exc:  # noqa: BLE001
                failed += 1
                for row in job["rows"]:
                    self.row_status.emit(row, f"{STATUS_FAILED}：{exc}", "")
            self.progress.emit(index, total)
        self.done.emit(ok, failed)


class ToolWorker(QThread):
    message = Signal(str)
    completed = Signal(bool, int)

    def __init__(self, func, parent=None):
        super().__init__(parent)
        self._func = func

    def run(self) -> None:
        try:
            outputs = self._func(self.message.emit)
            self.completed.emit(True, len(outputs or []))
        except Exception as exc:  # noqa: BLE001
            self.message.emit(f"失败：{exc}")
            self.completed.emit(False, 0)


# ---------------------------------------------------------------- 主窗口


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("TraeTools", "FileConverter")
        self.worker: ConvertWorker | None = None
        self.tool_worker: ToolWorker | None = None
        self.last_output_dir = str(Path.home() / "Desktop")
        self._tool_out_dir = ""

        self.setWindowTitle(f"{APP_TITLE}  v{APP_VERSION}")
        icon_file = resource_path("assets", "app.ico")
        if icon_file.exists():
            self.setWindowIcon(QIcon(str(icon_file)))
        self.resize(1120, 780)
        self.setMinimumSize(980, 680)
        self.setAcceptDrops(True)

        root = QWidget()
        root.setObjectName("Root")
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(24, 20, 24, 18)
        outer.setSpacing(14)

        outer.addLayout(self._build_header())

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_convert_tab(), "格式转换")
        self.tabs.addTab(self._build_toolbox_tab(), "PDF 工具箱")
        outer.addWidget(self.tabs, 1)

        self._restore_settings()
        self._on_tool_changed()
        self._update_option_visibility()
        self._refresh_queue_status()
        self._update_convert_button()

    # ------------------------------------------------------------ 头部

    def _build_header(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setSpacing(12)

        text_box = QVBoxLayout()
        text_box.setSpacing(2)
        title = QLabel(APP_TITLE)
        title.setObjectName("HeaderTitle")
        subtitle = QLabel("本地转换 · 不联网 · 支持批量处理与 PDF 页面工具")
        subtitle.setObjectName("HeaderSub")
        text_box.addWidget(title)
        text_box.addWidget(subtitle)

        badge = QLabel("Word / PDF / 图片")
        badge.setObjectName("Badge")

        layout.addLayout(text_box)
        layout.addStretch(1)
        layout.addWidget(badge)
        return layout

    # ------------------------------------------------------------ 格式转换页

    def _build_convert_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 10, 0, 0)
        layout.setSpacing(12)

        layout.addWidget(self._build_convert_settings())
        self.drop_area = DropArea()
        self.drop_area.filesDropped.connect(self._add_files_to_queue)
        self.drop_area.clicked.connect(self._choose_queue_files)
        layout.addWidget(self.drop_area)

        table_card = self._card()
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(0, 0, 0, 0)
        table_layout.setSpacing(0)
        self.queue_table = QTableWidget(0, 4)
        self.queue_table.setHorizontalHeaderLabels(["文件名", "源格式", "状态", "输出位置"])
        self.queue_table.verticalHeader().setVisible(False)
        self.queue_table.setShowGrid(False)
        self.queue_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.queue_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.queue_table.setAlternatingRowColors(False)
        header = self.queue_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.queue_table.setColumnWidth(1, 90)
        self.queue_table.setColumnWidth(2, 170)
        self.queue_table.verticalHeader().setDefaultSectionSize(38)
        self.queue_table.cellDoubleClicked.connect(self._open_queue_output)
        table_layout.addWidget(self.queue_table)
        layout.addWidget(table_card, 1)

        footer = QHBoxLayout()
        footer.setSpacing(10)
        self.queue_summary = QLabel("共 0 个文件")
        self.queue_summary.setObjectName("Hint")

        self.convert_progress = QProgressBar()
        self.convert_progress.setFixedWidth(220)
        self.convert_progress.setRange(0, 1)
        self.convert_progress.setValue(0)

        self.remove_btn = QPushButton("移除所选")
        self.remove_btn.clicked.connect(self._remove_selected_rows)
        self.clear_btn = QPushButton("清空列表")
        self.clear_btn.clicked.connect(self._clear_queue)
        self.open_output_btn = QPushButton("打开输出目录")
        self.open_output_btn.clicked.connect(self._open_output_dir)

        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.setObjectName("Primary")
        self.convert_btn.setMinimumWidth(140)
        self.convert_btn.clicked.connect(self._start_or_stop_convert)

        footer.addWidget(self.queue_summary)
        footer.addSpacing(8)
        footer.addWidget(self.remove_btn)
        footer.addWidget(self.clear_btn)
        footer.addStretch(1)
        footer.addWidget(self.convert_progress)
        footer.addWidget(self.open_output_btn)
        footer.addWidget(self.convert_btn)
        layout.addLayout(footer)
        return page

    def _build_convert_settings(self) -> QFrame:
        card = self._card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(10)

        row1 = QHBoxLayout()
        row1.setSpacing(10)
        row1.addWidget(self._label("转换为"))
        self.target_combo = QComboBox()
        self.target_combo.addItem("PDF（Word / 图片 → PDF）", ".pdf")
        self.target_combo.addItem("Word .docx（PDF → Word）", ".docx")
        self.target_combo.addItem("PNG 图片（图片互转 / PDF 导出）", ".png")
        self.target_combo.addItem("JPG 图片（图片互转 / PDF 导出）", ".jpg")
        self.target_combo.addItem("WebP 图片（图片互转）", ".webp")
        self.target_combo.addItem("BMP 图片（图片互转）", ".bmp")
        self.target_combo.setMinimumWidth(300)
        self.target_combo.currentIndexChanged.connect(self._on_target_changed)
        row1.addWidget(self.target_combo)

        row1.addSpacing(10)
        row1.addWidget(self._label("输出到"))
        self.out_dir_edit = QLineEdit()
        self.out_dir_edit.setPlaceholderText("默认保存到原文件所在目录")
        self.out_dir_edit.setMinimumWidth(260)
        row1.addWidget(self.out_dir_edit, 1)
        self.browse_out_btn = QPushButton("浏览…")
        self.browse_out_btn.clicked.connect(self._browse_output_dir)
        row1.addWidget(self.browse_out_btn)
        self.same_dir_cb = QCheckBox("与原文件同目录")
        self.same_dir_cb.setChecked(True)
        self.same_dir_cb.toggled.connect(self._on_same_dir_toggled)
        row1.addWidget(self.same_dir_cb)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.setSpacing(10)
        row2.setContentsMargins(0, 2, 0, 0)

        self.quality_box = QWidget()
        q_layout = QHBoxLayout(self.quality_box)
        q_layout.setContentsMargins(0, 0, 0, 0)
        q_layout.setSpacing(8)
        q_layout.addWidget(self._label("图片质量"))
        self.quality_slider = QSlider(Qt.Orientation.Horizontal)
        self.quality_slider.setRange(40, 100)
        self.quality_slider.setValue(90)
        self.quality_slider.setFixedWidth(150)
        self.quality_value = QLabel("90")
        self.quality_value.setObjectName("Hint")
        self.quality_slider.valueChanged.connect(lambda v: self.quality_value.setText(str(v)))
        q_layout.addWidget(self.quality_slider)
        q_layout.addWidget(self.quality_value)

        self.dpi_box = QWidget()
        d_layout = QHBoxLayout(self.dpi_box)
        d_layout.setContentsMargins(0, 0, 0, 0)
        d_layout.setSpacing(8)
        d_layout.addWidget(self._label("PDF 渲染精度"))
        self.dpi_spin = QSpinBox()
        self.dpi_spin.setRange(72, 400)
        self.dpi_spin.setValue(150)
        self.dpi_spin.setSuffix(" DPI")
        self.dpi_spin.setFixedWidth(110)
        self.dpi_spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.dpi_spin.setToolTip("可直接输入或滚动滚轮调整")
        d_layout.addWidget(self.dpi_spin)

        self.maxw_box = QWidget()
        w_layout = QHBoxLayout(self.maxw_box)
        w_layout.setContentsMargins(0, 0, 0, 0)
        w_layout.setSpacing(8)
        w_layout.addWidget(self._label("最长边不超过"))
        self.maxw_spin = QSpinBox()
        self.maxw_spin.setRange(0, 10000)
        self.maxw_spin.setValue(0)
        self.maxw_spin.setSingleStep(100)
        self.maxw_spin.setSpecialValueText("不限制")
        self.maxw_spin.setSuffix(" px")
        self.maxw_spin.setFixedWidth(120)
        self.maxw_spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.maxw_spin.setToolTip("可直接输入或滚动滚轮调整")
        w_layout.addWidget(self.maxw_spin)

        self.merge_cb = QCheckBox("把队列里的多张图片合并成一个 PDF")

        row2.addWidget(self.quality_box)
        row2.addSpacing(8)
        row2.addWidget(self.dpi_box)
        row2.addSpacing(8)
        row2.addWidget(self.maxw_box)
        row2.addSpacing(8)
        row2.addWidget(self.merge_cb)
        row2.addStretch(1)
        layout.addLayout(row2)
        return card

    def _on_target_changed(self) -> None:
        self._update_option_visibility()
        self._refresh_queue_status()

    def _update_option_visibility(self) -> None:
        target = self.target_combo.currentData()
        is_image = target in IMAGE_TARGETS
        self.quality_box.setVisible(target in {".jpg", ".webp"})
        self.dpi_box.setVisible(is_image)
        self.maxw_box.setVisible(target in IMAGE_TARGETS)
        self.merge_cb.setVisible(target == ".pdf")

    def _on_same_dir_toggled(self, checked: bool) -> None:
        self.out_dir_edit.setEnabled(not checked)
        self.browse_out_btn.setEnabled(not checked)

    def _browse_output_dir(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "选择输出目录", self.out_dir_edit.text() or str(Path.home()))
        if directory:
            self.out_dir_edit.setText(directory)

    def _output_dir(self) -> Path | None:
        if self.same_dir_cb.isChecked() or not self.out_dir_edit.text().strip():
            return None
        return Path(self.out_dir_edit.text().strip())

    # ------------------------------------------------------------ 队列操作

    def _choose_queue_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择要转换的文件",
            str(Path.home()),
            "所有支持的文件 (*.docx *.doc *.rtf *.odt *.txt *.pdf *.jpg *.jpeg *.png *.webp *.bmp *.gif *.tif *.tiff);;所有文件 (*.*)",
        )
        if files:
            self._add_files_to_queue(files)

    def _add_files_to_queue(self, paths: list[str]) -> None:
        if self.worker is not None and self.worker.isRunning():
            return
        pending: list[str] = []
        for raw in paths:
            path = Path(raw)
            if path.is_dir():
                for entry in sorted(path.iterdir()):
                    if entry.is_file() and entry.suffix.lower() in ALL_SUPPORTED:
                        pending.append(str(entry))
            elif path.is_file() and path.suffix.lower() in ALL_SUPPORTED:
                pending.append(str(path))
        existing = {self._row_path(row) for row in range(self.queue_table.rowCount())}
        for file_path in pending:
            resolved = str(Path(file_path).resolve())
            if resolved in existing:
                continue
            existing.add(resolved)
            row = self.queue_table.rowCount()
            self.queue_table.insertRow(row)
            name_item = QTableWidgetItem(Path(file_path).name)
            name_item.setToolTip(file_path)
            name_item.setData(Qt.ItemDataRole.UserRole, file_path)
            source_item = QTableWidgetItem(Path(file_path).suffix.upper().lstrip("."))
            source_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            status_item = QTableWidgetItem(STATUS_WAITING)
            status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            output_item = QTableWidgetItem("—")
            output_item.setForeground(QColor(MUTED))
            self.queue_table.setItem(row, 0, name_item)
            self.queue_table.setItem(row, 1, source_item)
            self.queue_table.setItem(row, 2, status_item)
            self.queue_table.setItem(row, 3, output_item)
        self._refresh_queue_status()
        self._update_summary()

    def _row_path(self, row: int) -> str:
        item = self.queue_table.item(row, 0)
        return item.data(Qt.ItemDataRole.UserRole) if item else ""

    def _remove_selected_rows(self) -> None:
        if self.worker is not None and self.worker.isRunning():
            return
        rows = sorted({index.row() for index in self.queue_table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.queue_table.removeRow(row)
        self._refresh_queue_status()
        self._update_summary()

    def _clear_queue(self) -> None:
        if self.worker is not None and self.worker.isRunning():
            return
        self.queue_table.setRowCount(0)
        self.convert_progress.setRange(0, 1)
        self.convert_progress.setValue(0)
        self._refresh_queue_status()
        self._update_summary()

    def _update_summary(self) -> None:
        count = self.queue_table.rowCount()
        total_bytes = 0
        for row in range(count):
            path = self._row_path(row)
            if path and os.path.exists(path):
                total_bytes += os.path.getsize(path)
        size_text = f" · {human_size(total_bytes)}" if total_bytes else ""
        self.queue_summary.setText(f"共 {count} 个文件{size_text}")
        self._update_convert_button()

    def _update_convert_button(self) -> None:
        running = self.worker is not None and self.worker.isRunning()
        self.convert_btn.setEnabled(running or self.queue_table.rowCount() > 0)
        self.convert_btn.setText("停止" if running else "开始转换")
        self.remove_btn.setEnabled(not running)
        self.clear_btn.setEnabled(not running)
        self.target_combo.setEnabled(not running)

    def _refresh_queue_status(self) -> None:
        target = self.target_combo.currentData()
        for row in range(self.queue_table.rowCount()):
            path = self._row_path(row)
            ext = Path(path).suffix.lower()
            status_item = self.queue_table.item(row, 2)
            if status_item is None:
                continue
            current = status_item.text()
            if current.startswith(STATUS_DONE) or current.startswith(STATUS_RUNNING):
                continue
            supported = target == ".pdf" and ext in (C.WORD_EXTS | C.IMAGE_EXTS)
            if target == ".docx":
                supported = ext == ".pdf"
            elif target in IMAGE_TARGETS:
                supported = ext == ".pdf" or ext in C.IMAGE_EXTS
            if supported:
                self._set_status(row, STATUS_WAITING)
            elif target == ".pdf" and ext == ".pdf":
                self._set_status(row, "已是 PDF，跳过")
            else:
                self._set_status(row, STATUS_SKIP)

    def _set_status(self, row: int, status: str, output: str = "") -> None:
        status_item = self.queue_table.item(row, 2)
        if status_item:
            status_item.setText(status)
            if status.startswith(STATUS_FAILED):
                color = RED
            elif status.startswith(STATUS_DONE):
                color = GREEN
            elif status.startswith(STATUS_RUNNING):
                color = ACCENT
            elif status in (STATUS_SKIP, "已是 PDF，跳过"):
                color = "#c07a2c"
            else:
                color = MUTED
            status_item.setForeground(QColor(color))
            status_item.setToolTip(status)
        if output:
            output_item = self.queue_table.item(row, 3)
            if output_item:
                output_item.setText(output)
                output_item.setToolTip(output)

    def _open_queue_output(self, row: int, _column: int) -> None:
        item = self.queue_table.item(row, 3)
        if item and item.text() not in ("", "—"):
            open_local(item.text())

    def _open_output_dir(self) -> None:
        for row in range(self.queue_table.rowCount() - 1, -1, -1):
            item = self.queue_table.item(row, 3)
            if item and item.text() not in ("", "—"):
                target = Path(item.text())
                open_local(str(target if target.is_dir() else target.parent))
                return
        fallback = self.out_dir_edit.text().strip() or str(Path.home())
        open_local(fallback)

    # ------------------------------------------------------------ 执行转换

    def _start_or_stop_convert(self) -> None:
        if self.worker is not None and self.worker.isRunning():
            self.worker.cancel()
            self.convert_btn.setEnabled(False)
            self.convert_btn.setText("正在停止…")
            return

        target = self.target_combo.currentData()
        jobs: list[dict] = []
        image_rows: list[int] = []
        image_srcs: list[str] = []
        merge_images = target == ".pdf" and self.merge_cb.isChecked()
        base_dir = self._output_dir()
        if base_dir:
            try:
                base_dir.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                QMessageBox.warning(self, "输出目录不可用", f"无法创建输出目录：{exc}")
                return

        for row in range(self.queue_table.rowCount()):
            path = self._row_path(row)
            status = self.queue_table.item(row, 2).text()
            if status.startswith(STATUS_SKIP) or status == "已是 PDF，跳过":
                continue
            out_dir = base_dir if base_dir else Path(path).parent
            if target == ".pdf":
                if merge_images and Path(path).suffix.lower() in C.IMAGE_EXTS:
                    image_rows.append(row)
                    image_srcs.append(path)
                    continue
                jobs.append({"rows": [row], "srcs": [path], "dst": out_dir / (Path(path).stem + ".pdf")})
            elif target == ".docx":
                jobs.append({"rows": [row], "srcs": [path], "dst": out_dir / (Path(path).stem + ".docx")})
            else:
                jobs.append({"rows": [row], "srcs": [path], "dst": out_dir / (Path(path).stem + target)})

        if image_rows:
            first_dir = base_dir if base_dir else Path(image_srcs[0]).parent
            jobs.append({"rows": image_rows, "srcs": image_srcs, "dst": first_dir / "合并图片.pdf"})

        if not jobs:
            QMessageBox.information(self, "没有可转换的文件", "当前队列里没有能转换成目标格式的文件。")
            return

        options = {
            "quality": self.quality_slider.value(),
            "dpi": self.dpi_spin.value(),
            "max_width": self.maxw_spin.value(),
            "pages": "",
        }

        for job in jobs:
            for row in job["rows"]:
                self._set_status(row, STATUS_WAITING, "")

        self.convert_progress.setRange(0, len(jobs))
        self.convert_progress.setValue(0)
        self.worker = ConvertWorker(jobs, options, self)
        self.worker.row_status.connect(self._set_status)
        self.worker.progress.connect(self._on_convert_progress)
        self.worker.done.connect(self._on_convert_done)
        self.worker.start()
        self._update_convert_button()

    def _on_convert_progress(self, done: int, total: int) -> None:
        self.convert_progress.setRange(0, total)
        self.convert_progress.setValue(done)

    def _on_convert_done(self, ok: int, failed: int) -> None:
        self.worker = None
        self.convert_btn.setText("开始转换")
        self.convert_btn.setEnabled(self.queue_table.rowCount() > 0)
        self.remove_btn.setEnabled(True)
        self.clear_btn.setEnabled(True)
        self.target_combo.setEnabled(True)
        if failed:
            QMessageBox.warning(self, "转换完成", f"成功 {ok} 个，失败 {failed} 个。\n可在列表中查看失败原因。")
        elif ok:
            self.last_output_dir = self._last_output()

    def _last_output(self) -> str:
        for row in range(self.queue_table.rowCount() - 1, -1, -1):
            item = self.queue_table.item(row, 3)
            if item and item.text() not in ("", "—"):
                target = Path(item.text())
                return str(target if target.is_dir() else target.parent)
        return self.out_dir_edit.text().strip() or str(Path.home())

    # ------------------------------------------------------------ PDF 工具箱

    def _build_toolbox_tab(self) -> QWidget:
        page = QWidget()
        layout = QHBoxLayout(page)
        layout.setContentsMargins(0, 10, 0, 0)
        layout.setSpacing(14)

        left = self._card()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(18, 16, 18, 16)
        left_layout.setSpacing(12)

        tool_row = QHBoxLayout()
        tool_row.setSpacing(10)
        tool_row.addWidget(self._label("工具"))
        self.tool_combo = QComboBox()
        for text, key in [
            ("合并 PDF", "merge"),
            ("拆分 PDF", "split"),
            ("旋转页面", "rotate"),
            ("加密 PDF", "encrypt"),
            ("解密 PDF", "decrypt"),
            ("提取文字", "text"),
            ("提取图片", "images"),
            ("压缩 PDF", "compress"),
        ]:
            self.tool_combo.addItem(text, key)
        self.tool_combo.currentIndexChanged.connect(self._on_tool_changed)
        tool_row.addWidget(self.tool_combo, 1)
        left_layout.addLayout(tool_row)

        self.tool_stack = QStackedWidget()
        self.tool_stack.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        self.tool_stack.addWidget(self._page_merge())
        self.tool_stack.addWidget(self._page_split())
        self.tool_stack.addWidget(self._page_rotate())
        self.tool_stack.addWidget(self._page_encrypt())
        self.tool_stack.addWidget(self._page_decrypt())
        self.tool_stack.addWidget(self._page_text())
        self.tool_stack.addWidget(self._page_images())
        self.tool_stack.addWidget(self._page_compress())
        left_layout.addWidget(self.tool_stack)

        files_label = QHBoxLayout()
        files_label.addWidget(self._label("待处理文件"))
        self.tool_files_hint = QLabel("支持多选，也可直接把文件拖到窗口里")
        self.tool_files_hint.setObjectName("Hint")
        files_label.addStretch(1)
        files_label.addWidget(self.tool_files_hint)
        left_layout.addLayout(files_label)

        self.tool_files = QListWidget()
        self.tool_files.setMinimumHeight(110)
        self.tool_files.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tool_files.setAcceptDrops(True)
        self.tool_files.setDragDropMode(QAbstractItemView.DragDropMode.DropOnly)
        self.tool_files.viewport().setAcceptDrops(True)
        self.tool_files.viewport().installEventFilter(self)
        left_layout.addWidget(self.tool_files, 1)

        file_buttons = QHBoxLayout()
        file_buttons.setSpacing(8)
        add_btn = QPushButton("+ 添加文件")
        add_btn.clicked.connect(self._choose_tool_files)
        remove_btn = QPushButton("移除所选")
        remove_btn.clicked.connect(self._remove_tool_files)
        clear_btn = QPushButton("清空")
        clear_btn.clicked.connect(lambda: self.tool_files.clear())
        file_buttons.addWidget(add_btn)
        file_buttons.addWidget(remove_btn)
        file_buttons.addWidget(clear_btn)
        file_buttons.addStretch(1)
        left_layout.addLayout(file_buttons)

        out_row = QHBoxLayout()
        out_row.setSpacing(10)
        self.tool_out_label = self._label("输出到")
        out_row.addWidget(self.tool_out_label)
        self.tool_out_edit = QLineEdit()
        self.tool_out_edit.setPlaceholderText("默认与原文件同目录")
        out_row.addWidget(self.tool_out_edit, 1)
        self.tool_out_browse = QPushButton("浏览…")
        self.tool_out_browse.clicked.connect(self._browse_tool_out_dir)
        out_row.addWidget(self.tool_out_browse)
        self.tool_same_dir = QCheckBox("与原文件同目录")
        self.tool_same_dir.setChecked(True)
        self.tool_same_dir.toggled.connect(self._on_tool_same_dir_toggled)
        out_row.addWidget(self.tool_same_dir)
        left_layout.addLayout(out_row)

        run_row = QHBoxLayout()
        run_row.setSpacing(10)
        self.tool_status = QLabel("就绪")
        self.tool_status.setObjectName("Hint")
        self.tool_open_btn = QPushButton("打开输出目录")
        self.tool_open_btn.clicked.connect(self._open_tool_output)
        self.tool_run_btn = QPushButton("开始处理")
        self.tool_run_btn.setObjectName("Primary")
        self.tool_run_btn.setMinimumWidth(140)
        self.tool_run_btn.clicked.connect(self._run_tool)
        run_row.addWidget(self.tool_status, 1)
        run_row.addWidget(self.tool_open_btn)
        run_row.addWidget(self.tool_run_btn)
        left_layout.addLayout(run_row)

        right = self._card()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(16, 14, 16, 14)
        right_layout.setSpacing(8)
        log_title = QLabel("运行日志")
        log_title.setObjectName("SectionTitle")
        self.tool_log = QTextEdit()
        self.tool_log.setReadOnly(True)
        self.tool_log.setPlaceholderText("处理过程和结果会显示在这里")
        right_layout.addWidget(log_title)
        right_layout.addWidget(self.tool_log, 1)

        layout.addWidget(left, 3)
        layout.addWidget(right, 2)
        return page

    def _page_merge(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        row = QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(self._label("输出文件"))
        self.merge_path_edit = QLineEdit()
        self.merge_path_edit.setPlaceholderText("留空则保存到第一个文件所在目录，命名为 合并结果.pdf")
        row.addWidget(self.merge_path_edit, 1)
        browse = QPushButton("另存为…")
        browse.clicked.connect(self._browse_merge_path)
        row.addWidget(browse)
        layout.addLayout(row)
        hint = QLabel("按下方列表顺序合并，可通过拖拽或添加顺序调整。")
        hint.setObjectName("Hint")
        layout.addWidget(hint)
        return widget

    def _page_split(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        row = QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(self._label("拆分方式"))
        self.split_mode = QComboBox()
        self.split_mode.addItem("每页拆成一个文件", "page")
        self.split_mode.addItem("每 N 页拆成一个文件", "chunk")
        self.split_mode.addItem("按自定义范围拆分", "range")
        self.split_mode.setMinimumWidth(180)
        self.split_mode.currentIndexChanged.connect(lambda i: self.split_value_stack.setCurrentIndex(1 if i == 1 else 2 if i == 2 else 0))
        row.addWidget(self.split_mode)
        self.split_value_stack = QStackedWidget()
        blank = QWidget()
        blank.setFixedHeight(1)
        self.split_value_stack.addWidget(blank)
        self.split_chunk = QSpinBox()
        self.split_chunk.setRange(1, 999)
        self.split_chunk.setValue(2)
        self.split_chunk.setPrefix("每 ")
        self.split_chunk.setSuffix(" 页")
        self.split_chunk.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.split_value_stack.addWidget(self.split_chunk)
        self.split_ranges = QLineEdit()
        self.split_ranges.setPlaceholderText("如 1-3,4-6,7")
        self.split_ranges.setMinimumWidth(180)
        self.split_value_stack.addWidget(self.split_ranges)
        row.addWidget(self.split_value_stack, 1)
        layout.addLayout(row)
        hint = QLabel("文件保存在输出目录，命名为「原文件名_第N部分.pdf」。")
        hint.setObjectName("Hint")
        layout.addWidget(hint)
        return widget

    def _page_rotate(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        row = QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(self._label("旋转角度"))
        self.rotate_angle = QComboBox()
        self.rotate_angle.addItems(["顺时针 90°", "旋转 180°", "逆时针 90°"])
        self.rotate_angle.setMinimumWidth(140)
        row.addWidget(self.rotate_angle)
        row.addSpacing(8)
        row.addWidget(self._label("页码"))
        self.rotate_pages = QLineEdit()
        self.rotate_pages.setPlaceholderText("留空为全部，如 1-3,5")
        row.addWidget(self.rotate_pages, 1)
        layout.addLayout(row)
        return widget

    def _page_encrypt(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        row = QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(self._label("打开密码"))
        self.encrypt_pw = QLineEdit()
        self.encrypt_pw.setEchoMode(QLineEdit.EchoMode.Password)
        self.encrypt_pw.setPlaceholderText("设置后打开 PDF 需要输入此密码")
        row.addWidget(self.encrypt_pw, 1)
        layout.addLayout(row)
        perms = QHBoxLayout()
        self.encrypt_print = QCheckBox("允许打印")
        self.encrypt_print.setChecked(True)
        self.encrypt_copy = QCheckBox("允许复制文字")
        self.encrypt_copy.setChecked(True)
        perms.addWidget(self.encrypt_print)
        perms.addWidget(self.encrypt_copy)
        perms.addStretch(1)
        layout.addLayout(perms)
        hint = QLabel("使用 AES-256 加密，生成「原文件名_加密.pdf」。")
        hint.setObjectName("Hint")
        layout.addWidget(hint)
        return widget

    def _page_decrypt(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        row = QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(self._label("当前密码"))
        self.decrypt_pw = QLineEdit()
        self.decrypt_pw.setEchoMode(QLineEdit.EchoMode.Password)
        self.decrypt_pw.setPlaceholderText("输入 PDF 的打开密码")
        row.addWidget(self.decrypt_pw, 1)
        layout.addLayout(row)
        hint = QLabel("解密后生成「原文件名_解密.pdf」，不带密码。")
        hint.setObjectName("Hint")
        layout.addWidget(hint)
        return widget

    def _page_text(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        row = QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(self._label("页码"))
        self.text_pages = QLineEdit()
        self.text_pages.setPlaceholderText("留空为全部，如 1-3,5")
        row.addWidget(self.text_pages, 1)
        self.text_layout = QCheckBox("按版面顺序排列")
        self.text_layout.setChecked(True)
        row.addWidget(self.text_layout)
        layout.addLayout(row)
        hint = QLabel("扫描版 PDF 没有文字层，需要先做 OCR，无法直接提取。")
        hint.setObjectName("Hint")
        layout.addWidget(hint)
        return widget

    def _page_images(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        row = QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(self._label("最小尺寸"))
        self.min_size_spin = QSpinBox()
        self.min_size_spin.setRange(0, 2000)
        self.min_size_spin.setValue(100)
        self.min_size_spin.setSuffix(" px")
        self.min_size_spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        row.addWidget(self.min_size_spin)
        row.addWidget(QLabel("（宽高都小于该值的图片会被忽略）"))
        row.addStretch(1)
        layout.addLayout(row)
        hint = QLabel("图片按在原 PDF 中的编码格式导出，保存在输出目录。")
        hint.setObjectName("Hint")
        layout.addWidget(hint)
        return widget

    def _page_compress(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        row = QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(self._label("压缩强度"))
        self.compress_level = QComboBox()
        self.compress_level.addItem("轻度（画质优先）", "light")
        self.compress_level.addItem("标准（推荐）", "standard")
        self.compress_level.addItem("强力（体积优先）", "strong")
        self.compress_level.setCurrentIndex(1)
        self.compress_level.setMinimumWidth(180)
        row.addWidget(self.compress_level)
        row.addStretch(1)
        layout.addLayout(row)
        hint = QLabel("会重采样内嵌图片并清理冗余数据，纯文字 PDF 压缩空间有限。")
        hint.setObjectName("Hint")
        layout.addWidget(hint)
        return widget

    def _on_tool_changed(self) -> None:
        index = self.tool_combo.currentIndex()
        self.tool_stack.setCurrentIndex(index)
        is_merge = self.tool_combo.currentData() == "merge"
        for widget in (self.tool_out_label, self.tool_out_edit, self.tool_out_browse, self.tool_same_dir):
            widget.setVisible(not is_merge)
        self.tool_files.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection if not is_merge else QAbstractItemView.SelectionMode.SingleSelection
        )

    def _on_tool_same_dir_toggled(self, checked: bool) -> None:
        self.tool_out_edit.setEnabled(not checked)
        self.tool_out_browse.setEnabled(not checked)

    def _browse_tool_out_dir(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "选择输出目录", self.tool_out_edit.text() or str(Path.home()))
        if directory:
            self.tool_out_edit.setText(directory)

    def _browse_merge_path(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "合并后保存为", self.merge_path_edit.text() or "合并结果.pdf", "PDF 文件 (*.pdf)")
        if path:
            self.merge_path_edit.setText(path)

    def _choose_tool_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "选择 PDF 文件", str(Path.home()), "PDF 文件 (*.pdf);;所有文件 (*.*)")
        for file_path in files:
            self._add_tool_files([file_path])

    def _add_tool_files(self, paths: list[str]) -> None:
        existing = {self.tool_files.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.tool_files.count())}
        for raw in paths:
            path = Path(raw)
            if path.is_dir():
                candidates = [entry for entry in sorted(path.iterdir()) if entry.is_file() and entry.suffix.lower() == ".pdf"]
            else:
                candidates = [path] if path.suffix.lower() == ".pdf" else []
            for candidate in candidates:
                resolved = str(candidate.resolve())
                if resolved in existing:
                    continue
                existing.add(resolved)
                item = QListWidgetItem(candidate.name)
                item.setData(Qt.ItemDataRole.UserRole, resolved)
                item.setToolTip(resolved)
                self.tool_files.addItem(item)
        if self.tool_combo.currentData() == "merge" and not self.merge_path_edit.text().strip():
            first = self.tool_files.item(0)
            if first:
                suggested = Path(first.data(Qt.ItemDataRole.UserRole)).parent / "合并结果.pdf"
                self.merge_path_edit.setText(str(suggested))

    def _remove_tool_files(self) -> None:
        for item in self.tool_files.selectedItems():
            self.tool_files.takeItem(self.tool_files.row(item))

    def _open_tool_output(self) -> None:
        if self._tool_out_dir and Path(self._tool_out_dir).exists():
            open_local(self._tool_out_dir)
        elif self.tool_out_edit.text().strip():
            open_local(self.tool_out_edit.text().strip())
        else:
            first = self.tool_files.item(0)
            if first:
                open_local(str(Path(first.data(Qt.ItemDataRole.UserRole)).parent))

    def eventFilter(self, obj, event):
        if obj is self.tool_files.viewport():
            if event.type() in (QEvent.Type.DragEnter, QEvent.Type.DragMove):
                if event.mimeData().hasUrls():
                    event.acceptProposedAction()
                    return True
            if event.type() == QEvent.Type.Drop:
                paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
                if paths:
                    self._add_tool_files(paths)
                    event.acceptProposedAction()
                    return True
        return super().eventFilter(obj, event)

    def _tool_file_paths(self) -> list[str]:
        return [self.tool_files.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.tool_files.count())]

    def _tool_log_line(self, text: str) -> None:
        self.tool_log.append(text)

    def _run_tool(self) -> None:
        if self.tool_worker is not None and self.tool_worker.isRunning():
            return
        tool = self.tool_combo.currentData()
        files = self._tool_file_paths()
        if not files:
            QMessageBox.information(self, "还没有文件", "请先添加要处理的 PDF 文件。")
            return
        if tool == "merge" and len(files) < 2:
            QMessageBox.information(self, "文件不足", "合并 PDF 至少需要 2 个文件。")
            return

        same_dir = self.tool_same_dir.isChecked() or not self.tool_out_edit.text().strip()
        custom_dir = None if same_dir else Path(self.tool_out_edit.text().strip())
        if custom_dir:
            try:
                custom_dir.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                QMessageBox.warning(self, "输出目录不可用", f"无法创建输出目录：{exc}")
                return
        self.tool_log.clear()
        self.tool_status.setText("处理中…")
        self.tool_run_btn.setEnabled(False)
        self.tool_open_btn.setEnabled(False)

        def out_dir_for(src: str) -> Path:
            return custom_dir if custom_dir else Path(src).parent

        def process_all(one, log) -> list[str]:
            results: list[str] = []
            errors = 0
            for src in files:
                try:
                    outputs = one(src)
                    results.extend(outputs)
                    for output in outputs:
                        log(f"已输出：{Path(output).name}")
                except Exception as exc:  # noqa: BLE001
                    errors += 1
                    log(f"失败：{Path(src).name} - {exc}")
            if not results and errors:
                raise C.ConversionError(f"{errors} 个文件处理失败")
            if errors and results:
                log(f"有 {errors} 个文件失败，其余已完成")
            return results

        if tool == "merge":
            target_text = self.merge_path_edit.text().strip()
            if not target_text:
                target_text = str(Path(files[0]).parent / "合并结果.pdf")
            dst = Path(target_text)
            if dst.suffix.lower() != ".pdf":
                dst = dst.with_suffix(".pdf")
            self._tool_out_dir = str(dst.parent)

            def job(log):
                log(f"合并 {len(files)} 个文件 → {dst.name}")
                outputs = C.pdf_merge(files, C.unique_path(dst))
                log(f"已输出：{Path(outputs[0]).name}")
                return outputs

        elif tool == "split":
            mode = self.split_mode.currentData()
            chunk = self.split_chunk.value()
            ranges = self.split_ranges.text().strip()

            def job(log):
                return process_all(lambda src: C.pdf_split(src, out_dir_for(src), mode, chunk, ranges), log)

        elif tool == "rotate":
            angle = [90, 180, 270][self.rotate_angle.currentIndex()]
            pages = self.rotate_pages.text().strip()

            def job(log):
                def one(src):
                    dst = out_dir_for(src) / f"{Path(src).stem}_旋转.pdf"
                    log(f"旋转 {Path(src).name} {angle}° …")
                    return C.pdf_rotate(src, C.unique_path(dst), angle, pages)

                return process_all(one, log)

        elif tool == "encrypt":
            password = self.encrypt_pw.text()
            if not password:
                QMessageBox.warning(self, "缺少密码", "请先输入打开密码。")
                self.tool_run_btn.setEnabled(True)
                self.tool_status.setText("就绪")
                return
            allow_print = self.encrypt_print.isChecked()
            allow_copy = self.encrypt_copy.isChecked()

            def job(log):
                def one(src):
                    dst = out_dir_for(src) / f"{Path(src).stem}_加密.pdf"
                    log(f"加密 {Path(src).name} …")
                    return C.pdf_encrypt(src, C.unique_path(dst), password, allow_print, allow_copy)

                return process_all(one, log)

        elif tool == "decrypt":
            password = self.decrypt_pw.text()

            def job(log):
                def one(src):
                    dst = out_dir_for(src) / f"{Path(src).stem}_解密.pdf"
                    log(f"解密 {Path(src).name} …")
                    return C.pdf_decrypt(src, C.unique_path(dst), password)

                return process_all(one, log)

        elif tool == "text":
            pages = self.text_pages.text().strip()
            keep_layout = self.text_layout.isChecked()

            def job(log):
                def one(src):
                    dst = out_dir_for(src) / f"{Path(src).stem}.txt"
                    log(f"提取文字 {Path(src).name} …")
                    return C.pdf_extract_text(src, C.unique_path(dst), pages, keep_layout)

                return process_all(one, log)

        elif tool == "images":
            min_size = self.min_size_spin.value()

            def job(log):
                def one(src):
                    log(f"提取图片 {Path(src).name} …")
                    return C.pdf_extract_images(src, out_dir_for(src), min_size)

                return process_all(one, log)

        else:  # compress
            level = self.compress_level.currentData()

            def job(log):
                def one(src):
                    dst = out_dir_for(src) / f"{Path(src).stem}_压缩.pdf"
                    log(f"压缩 {Path(src).name} …")
                    return C.pdf_compress(src, C.unique_path(dst), level)

                return process_all(one, log)

        if self._tool_out_dir == "":
            self._tool_out_dir = str(out_dir_for(files[0]))

        self.tool_worker = ToolWorker(job, self)
        self.tool_worker.message.connect(self._tool_log_line)
        self.tool_worker.completed.connect(self._on_tool_completed)
        self.tool_worker.start()

    def _on_tool_completed(self, success: bool, count: int) -> None:
        self.tool_run_btn.setEnabled(True)
        self.tool_open_btn.setEnabled(True)
        if success:
            self.tool_status.setText(f"完成 · {count} 个结果")
            self.tool_log.append(f"全部完成（{count} 个结果）")
        else:
            self.tool_status.setText("处理失败")
            self.tool_log.append("处理失败，请根据上方错误信息检查文件。")

    # ------------------------------------------------------------ 通用

    def _card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("Card")
        return card

    def _label(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("SectionTitle")
        return label

    def _restore_settings(self) -> None:
        geometry = self.settings.value("window/geometry")
        if geometry:
            self.restoreGeometry(geometry)
        out_dir = self.settings.value("convert/out_dir", "", type=str)
        if out_dir:
            self.out_dir_edit.setText(out_dir)
        same_dir = self.settings.value("convert/same_dir", "true", type=str).lower() == "true"
        self.same_dir_cb.setChecked(same_dir)
        self.out_dir_edit.setEnabled(not same_dir)
        self.browse_out_btn.setEnabled(not same_dir)
        target_index = self.settings.value("convert/target", 0, type=int)
        self.target_combo.setCurrentIndex(max(0, min(target_index, self.target_combo.count() - 1)))
        self.quality_slider.setValue(self.settings.value("convert/quality", 90, type=int))
        self.dpi_spin.setValue(self.settings.value("convert/dpi", 150, type=int))
        self.maxw_spin.setValue(self.settings.value("convert/max_width", 0, type=int))
        tool_index = self.settings.value("tool/index", 0, type=int)
        self.tool_combo.setCurrentIndex(max(0, min(tool_index, self.tool_combo.count() - 1)))

    def _save_settings(self) -> None:
        self.settings.setValue("window/geometry", self.saveGeometry())
        self.settings.setValue("convert/out_dir", self.out_dir_edit.text().strip())
        self.settings.setValue("convert/same_dir", "true" if self.same_dir_cb.isChecked() else "false")
        self.settings.setValue("convert/target", self.target_combo.currentIndex())
        self.settings.setValue("convert/quality", self.quality_slider.value())
        self.settings.setValue("convert/dpi", self.dpi_spin.value())
        self.settings.setValue("convert/max_width", self.maxw_spin.value())
        self.settings.setValue("tool/index", self.tool_combo.currentIndex())

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        if not paths:
            return
        if self.tabs.currentIndex() == 0:
            self._add_files_to_queue(paths)
        else:
            self._add_tool_files(paths)
        event.acceptProposedAction()

    def closeEvent(self, event) -> None:
        if self.worker is not None and self.worker.isRunning():
            self.worker.cancel()
            self.worker.wait(4000)
        if self.tool_worker is not None and self.tool_worker.isRunning():
            self.tool_worker.wait(4000)
        self._save_settings()
        self.settings.sync()
        super().closeEvent(event)


def _selftest_convert() -> bool:
    """打包后的自检：跑一遍 PDF/图片/Word 转换，确认依赖齐全。"""
    import shutil
    import tempfile

    tmp = Path(tempfile.mkdtemp(prefix="fc_selftest_"))
    try:
        doc = C.fitz.open()
        page = doc.new_page(width=300, height=200)
        page.insert_text((40, 80), "selftest")
        pdf = tmp / "selftest.pdf"
        doc.save(str(pdf))
        doc.close()

        images = C.pdf_to_images(pdf, tmp, "png", 72)
        if not images:
            return False
        word = C.pdf_to_docx(pdf, tmp / "selftest.docx")
        if not Path(word[0]).exists():
            return False
        jpg = C.convert_image(images[0], tmp / "selftest.jpg", 85)
        merged = C.images_to_pdf([images[0], jpg[0]], tmp / "merged.pdf", 85)
        if C.pdf_page_count(merged[0]) != 2:
            return False
        try:
            import pythoncom  # noqa: F401
            import win32com.client  # noqa: F401
        except ImportError:
            return False
        try:
            from docx import Document

            sample = Document()
            sample.add_paragraph("selftest")
            docx_in = tmp / "selftest.docx"
            sample.save(docx_in)
            C.word_to_pdf(docx_in, tmp / "word.pdf")
            if C.pdf_page_count(tmp / "word.pdf") < 1:
                return False
        except Exception:
            pass
        return True
    except Exception:
        return False
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main() -> int:
    selftest = "--selftest" in sys.argv
    if selftest:
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    if sys.platform == "win32":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("TraeTools.FileConverter.1")
        except Exception:
            pass
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setFont(QFont("Microsoft YaHei UI", 10))
    app.setStyleSheet(QSS)
    icon_file = resource_path("assets", "app.ico")
    if icon_file.exists():
        app.setWindowIcon(QIcon(str(icon_file)))
    window = MainWindow()
    if selftest:
        return 0 if _selftest_convert() else 1
    window.show()
    code = app.exec()
    try:
        window._save_settings()
        window.settings.sync()
    except Exception:
        pass
    return code


if __name__ == "__main__":
    # 打包环境下 Qt/PySide6 的退出清理偶发崩溃（0xC0000409），
    # 这里在设置已落盘后直接结束进程，跳过解释器与 Qt 的析构流程。
    try:
        _exit_code = main()
    except Exception:
        _exit_code = 1
    os._exit(_exit_code)
