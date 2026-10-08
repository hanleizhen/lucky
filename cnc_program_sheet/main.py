from __future__ import annotations

import os
import sys
import traceback
import xml.etree.ElementTree as ElementTree
from pathlib import Path
from typing import Any, Callable

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from PySide6.QtCore import QThread, QTimer, Qt, Signal
from PySide6.QtGui import QColor, QFont, QIcon, QImage, QPalette, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStatusBar,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .app_paths import (
    copy_user_image,
    default_user_template,
    load_embedded_update_repository,
    load_settings,
    log_exception,
    resource_path,
    save_settings,
)
from .excel_service import (
    LAYOUT,
    TemplateError,
    automatic_cells_for_rows,
    cleared_program_cells,
    export_workbook,
    output_filename,
    verify_template,
)
from .models import ImagePlacement, ProgramRecord, UNRECOGNIZED
from .nc_parser import parse_file_records
from .update_service import (
    UpdateError,
    UpdateInfo,
    check_latest,
    download_verified_installer,
    open_release_page,
    start_installer_after_exit,
)
from .version import APP_NAME, __version__


RESULT_COLUMNS: list[tuple[str, str]] = [
    ("程序名称", "program_name"),
    ("区分", "tool_type"),
    ("刀具号", "tool_number"),
    ("D", "diameter"),
    ("R", "radius"),
    ("加工DATA", "machining_data"),
    ("Depth", "depth"),
    ("解析状态", "status"),
]
FIELD_NAMES = {field for _, field in RESULT_COLUMNS if field != "status"}


def is_supported_nc_file(path: Path) -> bool:
    """Accept conventional .NC files and controller programs with no suffix.

    Copper-electrode programs commonly use names such as ``O0788`` without an
    extension.  Files with another explicit suffix remain excluded so an
    accidental spreadsheet/text document cannot be imported as NC code.
    """

    return path.is_file() and (path.suffix.lower() == ".nc" or not path.suffix)


class Worker(QThread):
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, task: Callable[[], Any], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._task = task

    def run(self) -> None:
        try:
            self.succeeded.emit(self._task())
        except Exception as exc:
            self.failed.emit(str(exc))


class ExcelPreview(QTableWidget):
    """Editable in-app preview of the real sheet values and major formatting."""

    cell_edited = Signal(str, str)
    insert_blank_row_requested = Signal(int)
    remove_blank_row_requested = Signal(int)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._loading = False
        self._theme_colors: list[str] = []
        self._image_marker_cells: set[str] = set()
        # Excel cells with “no fill” are white. Do not inherit a dark Windows
        # palette that would make the template preview black.
        palette = self.palette()
        palette.setColor(QPalette.ColorRole.Base, QColor("#FFFFFF"))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#FFFFFF"))
        palette.setColor(QPalette.ColorRole.Text, QColor("#000000"))
        self.setPalette(palette)
        self.setAlternatingRowColors(False)
        self.setWordWrap(True)
        self.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked | QTableWidget.EditTrigger.EditKeyPressed)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_program_row_menu)
        self.itemChanged.connect(self._on_item_changed)

    @staticmethod
    def _is_program_table_row(row: int) -> bool:
        return LAYOUT.first_program_row - 1 <= row <= LAYOUT.last_program_row - 1

    def _show_program_row_menu(self, position: Any) -> None:
        index = self.indexAt(position)
        if not index.isValid() or not self._is_program_table_row(index.row()):
            return
        self.setCurrentCell(index.row(), index.column())
        menu = QMenu(self)
        insert = menu.addAction("在下方插入空行")
        insert.setEnabled(index.row() < LAYOUT.last_program_row - 1)
        insert.triggered.connect(lambda: self.insert_blank_row_requested.emit(index.row() + 1))
        remove = menu.addAction("取消空行（也可按 Delete）")
        remove.triggered.connect(lambda: self.remove_blank_row_requested.emit(index.row()))
        menu.exec(self.viewport().mapToGlobal(position))

    def keyPressEvent(self, event: Any) -> None:
        # Delete is intentionally reserved for removing an operator-inserted
        # blank program row. Individual values remain editable by double click.
        if event.key() == Qt.Key.Key_Delete and self._is_program_table_row(self.currentRow()):
            self.remove_blank_row_requested.emit(self.currentRow())
            event.accept()
            return
        super().keyPressEvent(event)

    def shift_program_row_values(self, start_table_row: int, direction: int) -> None:
        """Move visible program-row values while retaining every cell's style.

        This is a logical row move inside the template's fixed NC area, not an
        Excel row insertion. It therefore cannot disturb template merges,
        printing layout, row heights, or the blue input-cell formatting.
        """

        first = LAYOUT.first_program_row - 1
        last = LAYOUT.last_program_row - 1
        if direction not in (-1, 1) or not first <= start_table_row <= last:
            return
        if direction == 1:
            source_rows = range(last - 1, start_table_row - 1, -1)
            clear_row = start_table_row
        else:
            source_rows = range(start_table_row + 1, last + 1)
            clear_row = last
        self._loading = True
        try:
            # Column A contains the fixed serial numbers, so only move the
            # user/program data area B onward.
            for column in range(1, self.columnCount()):
                for source_row in source_rows:
                    source = self.item(source_row, column)
                    destination = self.item(source_row + direction, column)
                    if source is not None and destination is not None:
                        destination.setText(source.text())
                item = self.item(clear_row, column)
                if item is not None:
                    item.setText("")
        finally:
            self._loading = False

    @staticmethod
    def _theme_palette(theme: bytes | None) -> list[str]:
        if not theme:
            return []
        try:
            root = ElementTree.fromstring(theme)
            namespace = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
            scheme = root.find(f".//{namespace}clrScheme")
            if scheme is None:
                return []
            values: list[str] = []
            for child in list(scheme):
                color = child.find(f"{namespace}srgbClr")
                if color is None:
                    color = child.find(f"{namespace}sysClr")
                if color is not None:
                    values.append(color.attrib.get("val") or color.attrib.get("lastClr") or "000000")
            return values
        except ElementTree.ParseError:
            return []

    def _rgb(self, openpyxl_color: Any) -> QColor | None:
        if openpyxl_color is not None and getattr(openpyxl_color, "type", "") == "rgb":
            value = getattr(openpyxl_color, "rgb", None)
            if value and len(value) >= 6:
                return QColor("#" + value[-6:])
        if openpyxl_color is not None and getattr(openpyxl_color, "type", "") == "theme":
            index = int(getattr(openpyxl_color, "theme", -1))
            if 0 <= index < len(self._theme_colors):
                color = QColor("#" + self._theme_colors[index][-6:])
                tint = float(getattr(openpyxl_color, "tint", 0.0) or 0.0)
                if tint:
                    channels = [color.red(), color.green(), color.blue()]
                    if tint > 0:
                        channels = [round(channel + (255 - channel) * tint) for channel in channels]
                    else:
                        channels = [round(channel * (1 + tint)) for channel in channels]
                    return QColor(*channels)
                return color
        return None

    @staticmethod
    def _editable(cell: Any) -> bool:
        if cell.data_type == "f":
            return False
        # Static template captions are protected in the preview. Blue/form fields,
        # all program rows, and blank header information cells remain editable.
        if cell.row <= 6 and cell.value not in (None, ""):
            return False
        if cell.column == 1 and cell.row >= 5:
            return False
        return True

    def load_template(self, template_path: Path) -> None:
        workbook = load_workbook(template_path, read_only=False, data_only=False, keep_links=True)
        try:
            sheet = workbook[LAYOUT.sheet_name]
            self._loading = True
            self._theme_colors = self._theme_palette(workbook.loaded_theme)
            self._image_marker_cells.clear()
            self.clear()
            self.setRowCount(sheet.max_row)
            self.setColumnCount(sheet.max_column)
            self.setHorizontalHeaderLabels([self._column_name(index) for index in range(1, sheet.max_column + 1)])
            self.setVerticalHeaderLabels([str(index) for index in range(1, sheet.max_row + 1)])
            for column in range(1, sheet.max_column + 1):
                width = sheet.column_dimensions[self._column_name(column)].width
                self.setColumnWidth(column - 1, max(44, int((width or 10) * 7 + 10)))
            for row in range(1, sheet.max_row + 1):
                height = sheet.row_dimensions[row].height
                self.setRowHeight(row - 1, max(23, int((height or 15) * 1.35)))
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell, MergedCell):
                        continue
                    item = QTableWidgetItem("" if cell.value is None else str(cell.value))
                    item.setToolTip(cell.coordinate)
                    # A no-fill cell carries ``00000000`` in openpyxl. That
                    # is transparent/default, not a black Excel fill.
                    fill = self._rgb(cell.fill.fgColor) if cell.fill.fill_type else None
                    item.setBackground(fill or QColor("#FFFFFF"))
                    font_color = self._rgb(cell.font.color)
                    font = QFont(cell.font.name or "Arial", int(cell.font.sz or 10))
                    font.setBold(bool(cell.font.bold))
                    font.setItalic(bool(cell.font.italic))
                    item.setFont(font)
                    item.setForeground(font_color or QColor("#000000"))
                    alignment = Qt.AlignmentFlag.AlignVCenter
                    if cell.alignment.horizontal == "center":
                        alignment |= Qt.AlignmentFlag.AlignHCenter
                    elif cell.alignment.horizontal == "right":
                        alignment |= Qt.AlignmentFlag.AlignRight
                    else:
                        alignment |= Qt.AlignmentFlag.AlignLeft
                    item.setTextAlignment(int(alignment))
                    if not self._editable(cell):
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    item.setData(Qt.ItemDataRole.UserRole, cell.coordinate)
                    self.setItem(cell.row - 1, cell.column - 1, item)
            for merged in sheet.merged_cells.ranges:
                self.setSpan(merged.min_row - 1, merged.min_col - 1, merged.max_row - merged.min_row + 1, merged.max_col - merged.min_col + 1)
        finally:
            self._loading = False
            workbook.close()

    @staticmethod
    def _column_name(number: int) -> str:
        name = ""
        while number:
            number, remainder = divmod(number - 1, 26)
            name = chr(65 + remainder) + name
        return name

    def set_value(self, cell: str, value: str) -> None:
        column_letters = "".join(character for character in cell if character.isalpha())
        row_number = int("".join(character for character in cell if character.isdigit()))
        column = 0
        for character in column_letters:
            column = column * 26 + ord(character) - 64
        item = self.item(row_number - 1, column - 1)
        if item is None:
            item = QTableWidgetItem()
            self.setItem(row_number - 1, column - 1, item)
            item.setData(Qt.ItemDataRole.UserRole, cell)
        self._loading = True
        item.setText(value)
        self._loading = False

    def current_coordinate(self) -> str | None:
        """Return the selected Excel coordinate, including merged-cell roots."""

        row, column = self.currentRow(), self.currentColumn()
        if row < 0 or column < 0:
            return None
        item = self.item(row, column)
        coordinate = item.data(Qt.ItemDataRole.UserRole) if item is not None else None
        return str(coordinate) if coordinate else None

    def refresh_image_markers(self, placements: list[ImagePlacement]) -> None:
        """Show small image markers in the preview; export retains full size."""

        for coordinate in self._image_marker_cells:
            row_digits = "".join(character for character in coordinate if character.isdigit())
            column_letters = "".join(character for character in coordinate if character.isalpha())
            if not row_digits or not column_letters:
                continue
            column = 0
            for character in column_letters:
                column = column * 26 + ord(character) - 64
            item = self.item(int(row_digits) - 1, column - 1)
            if item is not None:
                item.setIcon(QIcon())
                item.setToolTip(coordinate)
        self._image_marker_cells.clear()

        grouped: dict[str, list[ImagePlacement]] = {}
        for placement in placements:
            grouped.setdefault(placement.anchor, []).append(placement)
        for coordinate, anchored_images in grouped.items():
            row_digits = "".join(character for character in coordinate if character.isdigit())
            column_letters = "".join(character for character in coordinate if character.isalpha())
            if not row_digits or not column_letters:
                continue
            column = 0
            for character in column_letters:
                column = column * 26 + ord(character) - 64
            item = self.item(int(row_digits) - 1, column - 1)
            if item is None:
                continue
            thumbnail = QPixmap(str(anchored_images[0].source_path))
            if thumbnail.isNull():
                continue
            thumbnail = thumbnail.scaled(72, 52, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            item.setIcon(QIcon(thumbnail))
            names = "、".join(image.source_path.name for image in anchored_images)
            item.setToolTip(f"{coordinate}\n已插入图片：{names}\n导出 Excel 将按设定尺寸显示")
            self._image_marker_cells.add(coordinate)

    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        coordinate = item.data(Qt.ItemDataRole.UserRole)
        if not self._loading and coordinate:
            self.cell_edited.emit(str(coordinate), item.text())


class SettingsDialog(QDialog):
    def __init__(self, current_repository: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("更新设置")
        self.repository = QLineEdit(current_repository)
        self.repository.setPlaceholderText("owner/repository")
        form = QFormLayout(self)
        form.addRow("GitHub 更新仓库：", self.repository)
        note = QLabel("发布版会自动写入该仓库。仅在私有部署或测试时需要手动填写。")
        note.setWordWrap(True)
        form.addRow(note)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        form.addRow(buttons)


class ImageInsertDialog(QDialog):
    """Choose the exported Excel size for an image anchored to one cell."""

    def __init__(self, image_path: Path, anchor: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("插入图片")
        image = QImage(str(image_path))
        if image.isNull():
            raise ValueError("无法读取图片文件")
        self._ratio = image.width() / image.height()
        self._resizing = False
        self.width_input = QSpinBox()
        self.height_input = QSpinBox()
        for control in (self.width_input, self.height_input):
            control.setRange(1, 10000)
        self.width_input.setValue(image.width())
        self.height_input.setValue(image.height())
        self.keep_ratio = QCheckBox("保持宽高比例")
        self.keep_ratio.setChecked(True)
        self.width_input.valueChanged.connect(self._width_changed)
        self.height_input.valueChanged.connect(self._height_changed)

        layout = QFormLayout(self)
        layout.addRow("图片：", QLabel(image_path.name))
        layout.addRow("插入单元格：", QLabel(anchor))
        layout.addRow("宽度（像素）：", self.width_input)
        layout.addRow("高度（像素）：", self.height_input)
        layout.addRow(self.keep_ratio)
        hint = QLabel("预览显示缩略图；导出的 Excel 会使用这里设置的实际图片尺寸。")
        hint.setWordWrap(True)
        layout.addRow(hint)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def _width_changed(self, width: int) -> None:
        if self._resizing or not self.keep_ratio.isChecked():
            return
        self._resizing = True
        self.height_input.setValue(max(1, round(width / self._ratio)))
        self._resizing = False

    def _height_changed(self, height: int) -> None:
        if self._resizing or not self.keep_ratio.isChecked():
            return
        self._resizing = True
        self.width_input.setValue(max(1, round(height * self._ratio)))
        self._resizing = False

    @property
    def image_size(self) -> tuple[int, int]:
        return self.width_input.value(), self.height_input.value()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.settings = load_settings()
        # ``None`` is an intentional visual spacer inserted by the operator.
        # Keeping it in the row list means programs below the spacer retain
        # their exact Excel row when the workbook is exported.
        self.records: list[ProgramRecord | None] = []
        self.manual_cells: dict[str, str] = {}
        self.image_placements: list[ImagePlacement] = []
        self.template_path: Path | None = None
        self._program_area_reset = False
        self._result_loading = False
        self._update_worker: Worker | None = None
        self.setWindowTitle(f"{APP_NAME}  v{__version__}")
        self.setMinimumSize(1180, 720)
        self.setAcceptDrops(True)
        self._build_ui()
        self._load_initial_template()
        QTimer.singleShot(1200, self._automatic_update_check)
        self._background_update_timer = QTimer(self)
        self._background_update_timer.setInterval(6 * 60 * 60 * 1000)
        self._background_update_timer.timeout.connect(self._automatic_update_check)
        self._background_update_timer.start()

    def _build_ui(self) -> None:
        root = QWidget(self)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(12, 12, 12, 12)
        toolbar = QHBoxLayout()
        self.save_button = QPushButton("完成并保存")
        self.save_button.setObjectName("primaryButton")
        self.save_button.clicked.connect(self.save_output)
        toolbar.addWidget(self.save_button)
        reparse = QPushButton("重新解析")
        reparse.clicked.connect(self.reparse)
        toolbar.addWidget(reparse)
        clear = QPushButton("清空")
        clear.clicked.connect(self.clear_all)
        toolbar.addWidget(clear)
        template = QPushButton("打开模板")
        template.clicked.connect(self.open_template)
        toolbar.addWidget(template)
        insert_image = QPushButton("插入图片")
        insert_image.clicked.connect(self.insert_image)
        toolbar.addWidget(insert_image)
        remove_image = QPushButton("删除图片")
        remove_image.clicked.connect(self.remove_images_at_selected_cell)
        toolbar.addWidget(remove_image)
        add_files = QPushButton("选择 NC 文件")
        add_files.clicked.connect(self.choose_nc_files)
        toolbar.addWidget(add_files)
        toolbar.addStretch(1)
        updates = QPushButton("检查更新")
        updates.clicked.connect(lambda: self.check_updates(silent=False))
        toolbar.addWidget(updates)
        update_settings = QPushButton("更新设置")
        update_settings.clicked.connect(self.edit_update_settings)
        toolbar.addWidget(update_settings)
        layout.addLayout(toolbar)

        self.drop_label = QLabel("将 NC 程序拖到这里\n支持一次导入多个 .NC / .nc 文件，顺序将保持不变")
        self.drop_label.setObjectName("dropArea")
        self.drop_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.drop_label.setMinimumHeight(95)
        layout.addWidget(self.drop_label)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.addWidget(QLabel("已导入 NC 文件"))
        self.file_list = QListWidget()
        self.file_list.setMinimumWidth(265)
        left_layout.addWidget(self.file_list, 2)
        left_layout.addWidget(QLabel("解析结果（可直接修改）"))
        self.result_table = QTableWidget(0, len(RESULT_COLUMNS))
        self.result_table.setHorizontalHeaderLabels([title for title, _ in RESULT_COLUMNS])
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked | QTableWidget.EditTrigger.EditKeyPressed)
        self.result_table.itemChanged.connect(self._on_result_changed)
        self.result_table.setMinimumWidth(610)
        self.result_table.horizontalHeader().setStretchLastSection(True)
        left_layout.addWidget(self.result_table, 3)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        self.template_label = QLabel("Excel 程序单预览")
        right_layout.addWidget(self.template_label)
        self.preview = ExcelPreview()
        self.preview.cell_edited.connect(self._on_preview_edited)
        self.preview.insert_blank_row_requested.connect(self.insert_blank_program_row)
        self.preview.remove_blank_row_requested.connect(self.remove_blank_program_row)
        right_layout.addWidget(self.preview)
        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setSizes([560, 820])
        layout.addWidget(splitter, 1)
        self.setCentralWidget(root)
        self.setStatusBar(QStatusBar())
        self.setStyleSheet(
            "QWidget { font-family: 'Microsoft YaHei UI', Arial; font-size: 12px; }"
            "#dropArea { border: 2px dashed #4b87b9; border-radius: 8px; background: #edf6fd; color: #245b85; font-size: 15px; }"
            "#primaryButton { background: #0f6cbd; color: white; font-weight: 600; padding: 6px 15px; border-radius: 4px; }"
            "#primaryButton:hover { background: #005a9e; }"
            "QTableWidget { gridline-color: #d4dbe2; }"
        )

    def _load_initial_template(self) -> None:
        candidate = Path(str(self.settings.get("template_path", "")))
        if not candidate.is_file():
            candidate = default_user_template()
        self.load_template(candidate)

    def load_template(self, path: Path) -> None:
        try:
            verify_template(path)
            self.template_path = path
            self.preview.load_template(path)
            self.image_placements.clear()
            self.preview.refresh_image_markers(self.image_placements)
            self._program_area_reset = False
            # The template's DATE value is a generated field, not a permanent
            # part of the source workbook. Show today's date immediately.
            self.manual_cells.update(automatic_cells_for_rows([]))
            self.preview.set_value(LAYOUT.date_cell, self.manual_cells[LAYOUT.date_cell])
            self.template_label.setText(f"Excel 程序单预览：{path.name}")
            self.settings["template_path"] = str(path)
            save_settings(self.settings)
            self.statusBar().showMessage("模板已加载；原始模板只读使用，不会被覆盖。", 5000)
        except Exception as exc:
            self.show_error("无法打开模板", str(exc))

    def open_template(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(self, "选择 CNC 程序单模板", str(self.template_path or Path.home()), "Excel 文件 (*.xlsx)")
        if filename:
            self.load_template(Path(filename))

    def insert_image(self) -> None:
        """Anchor a selected image to the currently selected Excel cell."""

        anchor = self.preview.current_coordinate()
        if not anchor:
            self.show_error("请选择插入位置", "请先在右侧 Excel 预览中点击一个单元格，再点击“插入图片”。")
            return
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "选择要插入的图片",
            str(Path.home()),
            "图片文件 (*.png *.jpg *.jpeg *.bmp *.gif)",
        )
        if not filename:
            return
        try:
            image_path = Path(filename)
            dialog = ImageInsertDialog(image_path, anchor, self)
            if dialog.exec() != QDialog.DialogCode.Accepted:
                return
            width, height = dialog.image_size
            cached_image = copy_user_image(image_path)
            self.image_placements.append(
                ImagePlacement(source_path=cached_image, anchor=anchor, width=width, height=height)
            )
            self.preview.refresh_image_markers(self.image_placements)
            self.statusBar().showMessage(f"已将图片插入到 {anchor}；导出 Excel 时会保留该图片。", 4500)
        except Exception as exc:
            self.show_error("无法插入图片", str(exc))

    def remove_images_at_selected_cell(self) -> None:
        anchor = self.preview.current_coordinate()
        if not anchor:
            self.show_error("请选择图片位置", "请先在右侧 Excel 预览中点击图片所在的单元格。")
            return
        original_count = len(self.image_placements)
        self.image_placements = [placement for placement in self.image_placements if placement.anchor != anchor]
        if len(self.image_placements) == original_count:
            self.statusBar().showMessage(f"{anchor} 没有已插入的图片。", 3500)
            return
        self.preview.refresh_image_markers(self.image_placements)
        self.statusBar().showMessage(f"已移除 {anchor} 的图片。", 3500)

    def choose_nc_files(self) -> None:
        filenames, _ = QFileDialog.getOpenFileNames(
            self,
            "选择 NC 程序",
            str(Path.home()),
            "NC 程序 (*.NC *.nc);;铜工无扩展名程序 (*)",
        )
        self.add_nc_files([Path(filename) for filename in filenames])

    def dragEnterEvent(self, event: Any) -> None:  # Qt event type differs across bindings
        urls = event.mimeData().urls() if event.mimeData().hasUrls() else []
        if any(is_supported_nc_file(Path(url.toLocalFile())) for url in urls):
            event.acceptProposedAction()

    def dropEvent(self, event: Any) -> None:
        paths = [Path(url.toLocalFile()) for url in event.mimeData().urls()]
        self.add_nc_files(paths)
        event.acceptProposedAction()

    @staticmethod
    def _program_row_index(table_row: int) -> int | None:
        index = table_row - (LAYOUT.first_program_row - 1)
        return index if 0 <= index < LAYOUT.capacity else None

    def _has_imported_records(self) -> bool:
        return any(record is not None for record in self.records)

    def _has_tracked_data(self, index: int) -> bool:
        """Whether moving this program row would discard a user-entered value."""

        excel_row = LAYOUT.first_program_row + index
        for column in range(2, self.preview.columnCount() + 1):
            coordinate = f"{ExcelPreview._column_name(column)}{excel_row}"
            if self.manual_cells.get(coordinate, ""):
                return True
        return False

    def _has_image_at_program_index(self, index: int) -> bool:
        row = LAYOUT.first_program_row + index
        return any(
            "".join(character for character in placement.anchor if character.isdigit()) == str(row)
            for placement in self.image_placements
        )

    def _shift_program_images(self, start_index: int, direction: int) -> None:
        """Keep image anchors aligned when a logical program row moves."""

        shifted: list[ImagePlacement] = []
        for placement in self.image_placements:
            letters = "".join(character for character in placement.anchor if character.isalpha())
            digits = "".join(character for character in placement.anchor if character.isdigit())
            if not digits:
                shifted.append(placement)
                continue
            index = int(digits) - LAYOUT.first_program_row
            if 0 <= index < LAYOUT.capacity and index >= start_index:
                destination = index + direction
                if 0 <= destination < LAYOUT.capacity:
                    shifted.append(
                        ImagePlacement(
                            source_path=placement.source_path,
                            anchor=f"{letters}{LAYOUT.first_program_row + destination}",
                            width=placement.width,
                            height=placement.height,
                        )
                    )
            else:
                shifted.append(placement)
        self.image_placements = shifted
        self.preview.refresh_image_markers(self.image_placements)

    def _shift_tracked_program_cells(self, start_index: int, direction: int) -> None:
        """Shift exported manual values in lockstep with the preview values."""

        moved: dict[str, str] = {}
        remove: list[str] = []
        first = LAYOUT.first_program_row
        last = LAYOUT.last_program_row
        for coordinate, value in self.manual_cells.items():
            letters = "".join(character for character in coordinate if character.isalpha())
            digits = "".join(character for character in coordinate if character.isdigit())
            if not digits:
                continue
            row = int(digits)
            index = row - first
            if first <= row <= last and index >= start_index:
                destination = index + direction
                if 0 <= destination < LAYOUT.capacity:
                    moved[f"{letters}{first + destination}"] = value
                remove.append(coordinate)
        for coordinate in remove:
            self.manual_cells.pop(coordinate, None)
        self.manual_cells.update(moved)

    def insert_blank_program_row(self, table_row: int) -> None:
        """Insert an empty logical row below the selected preview row."""

        index = self._program_row_index(table_row)
        if index is None:
            return
        if not self._has_imported_records():
            self.statusBar().showMessage("请先导入 NC 程序，再插入程序单空行。", 3500)
            return
        if self._has_tracked_data(LAYOUT.capacity - 1) or self._has_image_at_program_index(LAYOUT.capacity - 1) or (
            len(self.records) == LAYOUT.capacity and self.records[-1] is not None
        ):
            self.show_error("无法插入空行", "模板最后一行已有数据。请先腾出最后一个程序行，避免覆盖数据。")
            return
        # A row below the final imported program can still be selected in the
        # preview. Preserve its requested position by representing preceding
        # gaps explicitly.
        while len(self.records) < index:
            self.records.append(None)
        if len(self.records) == LAYOUT.capacity:
            self.records.pop()
        self.preview.shift_program_row_values(table_row, 1)
        self._shift_tracked_program_cells(index, 1)
        self._shift_program_images(index, 1)
        self.records.insert(index, None)
        self._apply_records_to_view()
        self.statusBar().showMessage("已在所选行下方插入空行；选中该空行后按 Delete 可取消。", 4500)

    def remove_blank_program_row(self, table_row: int) -> None:
        """Remove an operator-inserted empty row and close the gap."""

        index = self._program_row_index(table_row)
        if index is None or index >= len(self.records) or self.records[index] is not None:
            self.statusBar().showMessage("请选择通过“插入空行”创建的空白程序行，再按 Delete。", 4000)
            return
        if self._has_tracked_data(index) or self._has_image_at_program_index(index):
            self.statusBar().showMessage("该行已有手动填写内容或图片，无法作为空行取消。请先清空该行内容。", 4000)
            return
        self.preview.shift_program_row_values(table_row, -1)
        self._shift_tracked_program_cells(index, -1)
        self._shift_program_images(index, -1)
        self.records.pop(index)
        self._apply_records_to_view()
        self.statusBar().showMessage("已取消空行，后续程序数据已上移。", 3500)

    def add_nc_files(self, paths: list[Path]) -> None:
        paths = [path for path in paths if is_supported_nc_file(path)]
        if not paths:
            self.show_error("未导入文件", "请拖入或选择至少一个 .NC / .nc 文件，或无扩展名的铜工 NC 程序。")
            return
        existing = {record.source_path.resolve() for record in self.records if record is not None}
        overflow_records: list[ProgramRecord] = []
        for path in paths:  # Do not sort: Windows drop order is the program-sheet order.
            if path.resolve() in existing:
                continue
            try:
                parsed_records = parse_file_records(path)
                free_rows = LAYOUT.capacity - len(self.records)
                if free_rows <= 0:
                    overflow_records.extend(parsed_records)
                    continue
                self.records.extend(parsed_records[:free_rows])
                overflow_records.extend(parsed_records[free_rows:])
                existing.add(path.resolve())
            except Exception as exc:
                self.show_error("NC 文件解析失败", f"{path.name}\n\n{exc}")
        if overflow_records:
            self.show_capacity_overflow(overflow_records)
        self._apply_records_to_view()

    def show_capacity_overflow(self, overflow_records: list[ProgramRecord]) -> None:
        """Name every omitted program row and give a safe continuation path."""

        labels = [f"{record.program_name}（{record.tool_number}）" for record in overflow_records]
        preview_labels = labels[:20]
        displayed = "\n".join(f"• {label}" for label in preview_labels)
        if len(labels) > len(preview_labels):
            displayed += f"\n• ……其余 {len(labels) - len(preview_labels)} 个请点“显示详细信息”查看"
        box = QMessageBox(self)
        box.setWindowTitle("模板容量已满")
        box.setIcon(QMessageBox.Icon.Warning)
        box.setText(f"模板最多填写 {LAYOUT.capacity} 个程序行。以下 {len(labels)} 个程序行本次未导入：")
        box.setInformativeText(
            f"{displayed}\n\n"
            "请先点击“完成并保存”保存当前已整理的程序单，然后点击“清空”清除当前导入记录，"
            "再重新导入上述未导入的程序。\n\n"
            "“清空”只清除软件当前列表，不会删除任何原始 NC 文件。"
        )
        box.setDetailedText("未导入的程序行：\n" + "\n".join(labels))
        box.exec()

    def reparse(self) -> None:
        if not self._has_imported_records():
            self.statusBar().showMessage("当前没有可重新解析的 NC 文件。", 3500)
            return
        refreshed_by_path: dict[Path, list[ProgramRecord]] = {}
        for record in self.records:
            if record is None:
                continue
            source = record.source_path.resolve()
            if source not in refreshed_by_path:
                try:
                    refreshed_by_path[source] = parse_file_records(record.source_path)
                except Exception as exc:
                    refreshed_by_path[source] = [
                        ProgramRecord(
                            record.source_path,
                            record.program_name,
                            source_tool_index=record.source_tool_index,
                            warnings=[f"重新解析失败：{exc}"],
                        )
                    ]
        refreshed: list[ProgramRecord | None] = []
        for record in self.records:
            if record is None:
                refreshed.append(None)
                continue
            parsed_records = refreshed_by_path[record.source_path.resolve()]
            if record.source_tool_index < len(parsed_records):
                refreshed.append(parsed_records[record.source_tool_index])
            else:
                refreshed.append(
                    ProgramRecord(
                        record.source_path,
                        record.program_name,
                        source_tool_index=record.source_tool_index,
                        warnings=["重新解析后未找到对应刀具段"],
                    )
                )
        self.records = refreshed
        self._apply_records_to_view()
        self.statusBar().showMessage("已按原始 NC 文件重新解析。", 3500)

    def clear_all(self) -> None:
        if self._has_imported_records() and QMessageBox.question(self, "清空", "清空已导入文件及当前预览中的手动修改？") != QMessageBox.StandardButton.Yes:
            return
        self.records.clear()
        self.manual_cells.clear()
        self.image_placements.clear()
        self._program_area_reset = False
        self.file_list.clear()
        self._set_result_table()
        if self.template_path:
            self.preview.load_template(self.template_path)
            self.preview.refresh_image_markers(self.image_placements)
            self.manual_cells.update(automatic_cells_for_rows([]))
            self.preview.set_value(LAYOUT.date_cell, self.manual_cells[LAYOUT.date_cell])
        self.statusBar().showMessage("已恢复到未填写的模板预览。", 3500)

    def _apply_records_to_view(self) -> None:
        self.file_list.clear()
        listed_sources: set[Path] = set()
        for record in self.records:
            if record is not None and record.source_path.resolve() not in listed_sources:
                self.file_list.addItem(record.program_name)
                listed_sources.add(record.source_path.resolve())
        # The supplied workbook contains filled reference rows. On the first
        # import, blank only the NC-generated fields across the program area so
        # old sample entries cannot leak into the exported sheet.
        if not self._program_area_reset:
            cleared = cleared_program_cells()
            self.manual_cells.update(cleared)
            for cell, value in cleared.items():
                self.preview.set_value(cell, value)
            self._program_area_reset = True
        # Re-parsing deliberately resets automatic fields. Other cells edited in
        # the preview stay intact until the user chooses “清空”. Clear previous
        # automatic values first so a manually inserted blank row remains blank.
        automatic = automatic_cells_for_rows(self.records)
        automatic_area = cleared_program_cells()
        self.manual_cells.update(automatic_area)
        self.manual_cells.update(automatic)
        for cell, value in {**automatic_area, **automatic}.items():
            self.preview.set_value(cell, value)
        self._set_result_table()

    def _set_result_table(self) -> None:
        self._result_loading = True
        self.result_table.setRowCount(len(self.records))
        for row, record in enumerate(self.records):
            for column, (_, field) in enumerate(RESULT_COLUMNS):
                if record is None:
                    value = "空行（预览中按 Delete 取消）" if field == "status" else ""
                else:
                    value = record.status if field == "status" else str(getattr(record, field))
                item = QTableWidgetItem(value)
                if field == "status" or record is None:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    if record is not None and record.has_unrecognized_values:
                        item.setForeground(QColor("#a15c00"))
                self.result_table.setItem(row, column, item)
        self.result_table.resizeColumnsToContents()
        self._result_loading = False

    def _on_result_changed(self, item: QTableWidgetItem) -> None:
        if self._result_loading or item.row() >= len(self.records):
            return
        field = RESULT_COLUMNS[item.column()][1]
        if field not in FIELD_NAMES:
            return
        record = self.records[item.row()]
        if record is None:
            return
        setattr(record, field, item.text())
        target = LAYOUT.target_cell(field, item.row())
        self.manual_cells[target] = item.text()
        self.preview.set_value(target, item.text())

    def _on_preview_edited(self, cell: str, value: str) -> None:
        self.manual_cells[cell] = value
        for index, record in enumerate(self.records):
            if record is None:
                continue
            for field in FIELD_NAMES:
                if LAYOUT.target_cell(field, index) == cell:
                    setattr(record, field, value)
                    self._set_result_table()
                    return

    def _repository(self) -> str:
        return str(self.settings.get("update_repository") or load_embedded_update_repository()).strip()

    def edit_update_settings(self) -> None:
        dialog = SettingsDialog(self._repository(), self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.settings["update_repository"] = dialog.repository.text().strip()
            save_settings(self.settings)
            self.statusBar().showMessage("更新设置已保存。", 3500)

    def _automatic_update_check(self) -> None:
        if self._repository():
            self.check_updates(silent=True)

    def check_updates(self, silent: bool) -> None:
        if self._update_worker and self._update_worker.isRunning():
            return
        repository = self._repository()
        if not repository:
            if not silent:
                self.show_error("尚未配置更新源", "请先点击“更新设置”，填写 GitHub 仓库 owner/repository。")
            return
        self.statusBar().showMessage("正在检查 GitHub Release 更新…")
        worker = Worker(lambda: check_latest(repository), self)
        self._update_worker = worker
        worker.succeeded.connect(lambda update: self._on_update_check_success(update, silent))
        worker.failed.connect(lambda error: self._on_update_check_failed(error, silent))
        worker.finished.connect(worker.deleteLater)
        worker.start()

    def _on_update_check_success(self, update: UpdateInfo | None, silent: bool) -> None:
        self._update_worker = None
        if update is None:
            if not silent:
                self.statusBar().showMessage("已是最新版本。", 3500)
            return
        box = QMessageBox(self)
        box.setWindowTitle("发现新版本")
        box.setIcon(QMessageBox.Icon.Information)
        box.setText(f"发现 v{update.version}，当前为 v{__version__}。")
        box.setInformativeText("更新只替换程序安装文件；用户模板、配置和个人数据均保存在 AppData，不会被覆盖。")
        install = box.addButton("下载并安装", QMessageBox.ButtonRole.AcceptRole)
        release = box.addButton("查看 Release", QMessageBox.ButtonRole.ActionRole)
        box.addButton("稍后", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() is release:
            open_release_page(update)
        elif box.clickedButton() is install:
            self._download_update(update)

    def _on_update_check_failed(self, error: str, silent: bool) -> None:
        self._update_worker = None
        if not silent:
            self.show_error("检查更新失败", error)

    def _download_update(self, update: UpdateInfo) -> None:
        self.statusBar().showMessage("正在下载并校验更新安装包…")
        worker = Worker(lambda: download_verified_installer(update), self)
        self._update_worker = worker
        worker.succeeded.connect(self._on_update_downloaded)
        worker.failed.connect(lambda error: self._on_update_check_failed(error, False))
        worker.finished.connect(worker.deleteLater)
        worker.start()

    def _on_update_downloaded(self, installer: Path) -> None:
        self._update_worker = None
        try:
            start_installer_after_exit(installer)
            QMessageBox.information(self, "开始升级", "已下载并校验更新。程序关闭后将自动安装新版。")
            QApplication.instance().quit()
        except Exception as exc:
            self.show_error("无法启动更新安装程序", str(exc))

    def _validate_before_save(self) -> bool:
        if not self._has_imported_records():
            self.show_error("无法保存", "请先拖入或选择至少一个 NC 程序。")
            return False
        try:
            automatic_cells_for_rows(self.records)
        except TemplateError as exc:
            self.show_error("无法保存", str(exc))
            return False
        uncertain = [record.program_name for record in self.records if record is not None and record.has_unrecognized_values]
        if uncertain:
            text = "以下程序仍有“未识别”项：\n" + "、".join(uncertain)
            text += "\n\n可在左侧解析表或右侧 Excel 预览中手动修改，仍要保存吗？"
            return QMessageBox.question(self, "需要确认", text) == QMessageBox.StandardButton.Yes
        return True

    def save_output(self) -> None:
        if not self.template_path or not self._validate_before_save():
            return
        try:
            # Ensure table edits are applied even when a preview cell was never clicked.
            self.manual_cells.update(automatic_cells_for_rows(self.records))
            desktop = Path(os.path.join(os.path.expanduser("~"), "Desktop"))
            # QStandardPaths is more reliable with redirected Windows desktops.
            from PySide6.QtCore import QStandardPaths

            desktop_text = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DesktopLocation)
            if desktop_text:
                desktop = Path(desktop_text)
            target = output_filename([record for record in self.records if record is not None], desktop)
            export_workbook(self.template_path, target, self.manual_cells, images=self.image_placements)
            self._show_save_success(target)
        except Exception as exc:
            self.show_error("保存程序单失败", str(exc))

    def _show_save_success(self, target: Path) -> None:
        box = QMessageBox(self)
        box.setWindowTitle("程序单保存成功")
        box.setIcon(QMessageBox.Icon.Information)
        box.setText("程序单保存成功")
        box.setInformativeText(str(target))
        open_file = box.addButton("打开文件", QMessageBox.ButtonRole.AcceptRole)
        open_folder = box.addButton("打开文件夹", QMessageBox.ButtonRole.ActionRole)
        box.addButton("关闭", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() is open_file:
            os.startfile(target)  # type: ignore[attr-defined]
        elif box.clickedButton() is open_folder:
            os.startfile(target.parent)  # type: ignore[attr-defined]

    def show_error(self, title: str, details: str) -> None:
        log_path = log_exception(f"{title}\n{details}")
        QMessageBox.critical(self, title, f"{details}\n\n详细日志：{log_path}")


def run() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName("CNCProgramSheet")
    window = MainWindow()
    window.show()
    return app.exec()


def _exception_hook(exc_type: type[BaseException], value: BaseException, trace: Any) -> None:
    details = "".join(traceback.format_exception(exc_type, value, trace))
    log_path = log_exception(details)
    if QApplication.instance() is not None:
        QMessageBox.critical(None, "程序发生错误", f"程序没有正常完成操作。\n\n{value}\n\n详细日志：{log_path}")
    else:
        sys.stderr.write(f"程序发生错误：{value}\n详细日志：{log_path}\n")


def self_check() -> int:
    """Non-GUI packaging check used by CI and release validation."""

    marker_text = os.environ.get("CNC_PROGRAM_SHEET_SELF_CHECK_MARKER")
    marker = Path(marker_text) if marker_text else None
    template = resource_path("assets", "CNC程序单.xlsx")
    if marker:
        marker.write_text(f"started\ntemplate={template}\nexists={template.exists()}\n", encoding="utf-8")
    try:
        verify_template(template)
        expected = {"program_name": "B", "tool_type": "C", "tool_number": "I", "diameter": "D", "radius": "E", "machining_data": "F", "depth": "J"}
        if dict(LAYOUT.columns) != expected or LAYOUT.capacity != 19:
            raise TemplateError("内置模板映射校验失败。")
    except Exception as exc:
        if marker:
            marker.write_text(marker.read_text(encoding="utf-8") + f"error={exc}\n", encoding="utf-8")
        raise
    if marker:
        marker.write_text(marker.read_text(encoding="utf-8") + "ok\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.excepthook = _exception_hook
    raise SystemExit(self_check() if "--self-check" in sys.argv else run())
