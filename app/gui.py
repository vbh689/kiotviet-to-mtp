import sys
import os
import io
import contextlib
from dataclasses import dataclass
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGridLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from PyQt6.QtCore import Qt, pyqtSignal, QThread

from .kv_config import (
    CUSTOMER_HEADER_ALIASES,
    PRODUCT_HEADER_ALIASES,
    PROVIDER_HEADER_ALIASES,
)
from .kv_excel import read_excel_headers, resolve_alias_columns
from .kv_mapping import ColumnMappings, MAPPING_METADATA, SOURCE_TYPE_LABELS
from .kv_runner import convert_kiotviet_files, detect_source_type, get_default_outdir
from .kv_utils import clean_text, excel_column_letter


ALIASES_BY_SOURCE_TYPE = {
    "product": PRODUCT_HEADER_ALIASES,
    "customer": CUSTOMER_HEADER_ALIASES,
    "provider": PROVIDER_HEADER_ALIASES,
}


@dataclass(frozen=True)
class SourceFileInfo:
    path: Path
    source_type: str
    headers: list[str]


def format_source_column(index: int, header: str) -> str:
    label = clean_text(header) or "(không có tiêu đề)"
    return f"{excel_column_letter(index)} - {label}"


# Run conversion directly while redirecting console output into the GUI log.
def run_gui_conversion(
    files: list[str],
    column_mappings: ColumnMappings,
    output_dir: str | None = None,
    merge_dvt: bool = False,
) -> str:
    outdir = Path(output_dir) if output_dir else get_default_outdir()

    f = io.StringIO()
    try:
        with contextlib.redirect_stdout(f), contextlib.redirect_stderr(f):
            convert_kiotviet_files([Path(file) for file in files], outdir, column_mappings, merge_dvt=merge_dvt)
    except Exception as e:
        import traceback
        f.write(traceback.format_exc())

    return f.getvalue()


class ColumnMappingDialog(QDialog):
    def __init__(self, grouped_files: dict[str, list[SourceFileInfo]], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Chọn cột dữ liệu KiotViet")
        self.resize(860, 520)
        self.grouped_files = grouped_files
        self.combos: dict[str, dict[str, QComboBox]] = {}

        layout = QVBoxLayout(self)
        intro = QLabel(
            "Kiểm tra hoặc đổi cột nguồn cho lần chuyển đổi này. "
            "Các lựa chọn này không được lưu cho lần chạy sau."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        tabs = QTabWidget()
        layout.addWidget(tabs)

        for source_type, infos in grouped_files.items():
            tabs.addTab(
                self._build_source_tab(source_type, infos),
                SOURCE_TYPE_LABELS.get(source_type, source_type),
            )

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _build_source_tab(self, source_type: str, infos: list[SourceFileInfo]) -> QWidget:
        tab = QWidget()
        layout = QGridLayout(tab)
        layout.setColumnStretch(1, 1)
        layout.setColumnStretch(2, 2)

        first = infos[0]
        default_columns = resolve_alias_columns(first.headers, ALIASES_BY_SOURCE_TYPE[source_type])
        self.combos[source_type] = {}

        file_names = ", ".join(info.path.name for info in infos)
        file_label = QLabel(f"Áp dụng cho: {file_names}")
        file_label.setWordWrap(True)
        layout.addWidget(file_label, 0, 0, 1, 3)

        layout.addWidget(QLabel("Trường KiotViet"), 1, 0)
        layout.addWidget(QLabel("Cột nguồn"), 1, 1)
        layout.addWidget(QLabel("Dùng cho MTP"), 1, 2)

        for row_idx, item in enumerate(MAPPING_METADATA[source_type], start=2):
            layout.addWidget(QLabel(item.label), row_idx, 0)

            combo = QComboBox()
            combo.addItem("-- Chọn cột --", None)
            for col_idx, header in enumerate(first.headers, start=1):
                combo.addItem(format_source_column(col_idx, header), col_idx)

            default_idx = default_columns.get(item.field)
            if default_idx is not None:
                found = combo.findData(default_idx)
                if found >= 0:
                    combo.setCurrentIndex(found)

            self.combos[source_type][item.field] = combo
            layout.addWidget(combo, row_idx, 1)

            targets = "; ".join(
                f"{target.template} / {target.column} / {target.label}"
                for target in item.targets
            )
            target_label = QLabel(targets)
            target_label.setWordWrap(True)
            layout.addWidget(target_label, row_idx, 2)

        return tab

    def column_mappings(self) -> ColumnMappings:
        mappings: ColumnMappings = {}
        for source_type, fields in self.combos.items():
            mappings[source_type] = {}
            for field, combo in fields.items():
                mappings[source_type][field] = combo.currentData()
        return mappings

    def accept(self):
        mappings = self.column_mappings()
        missing: list[str] = []
        invalid: list[str] = []

        for source_type, fields in mappings.items():
            label = SOURCE_TYPE_LABELS.get(source_type, source_type)
            for item in MAPPING_METADATA[source_type]:
                col_idx = fields.get(item.field)
                if col_idx is None:
                    if not getattr(item, "is_optional", False):
                        missing.append(f"{label}: {item.label}")
                    continue
                for info in self.grouped_files[source_type]:
                    if col_idx > len(info.headers):
                        invalid.append(
                            f"{info.path.name}: {item.label} chọn cột {excel_column_letter(col_idx)} "
                            f"nhưng file chỉ có {len(info.headers)} cột"
                        )

        if missing:
            QMessageBox.warning(
                self,
                "Thiếu mapping",
                "Vui lòng chọn cột cho:\n" + "\n".join(missing),
            )
            return

        if invalid:
            QMessageBox.warning(
                self,
                "Mapping không hợp lệ",
                "Một số lựa chọn không tồn tại trong file cùng loại:\n" + "\n".join(invalid),
            )
            return

        super().accept()

class ConversionWorker(QThread):
    result_ready = pyqtSignal(str)

    def __init__(self, files, column_mappings: ColumnMappings, merge_dvt: bool = False):
        super().__init__()
        self.files = files
        self.column_mappings = column_mappings
        self.merge_dvt = merge_dvt

    def run(self):
        output = run_gui_conversion(self.files, self.column_mappings, merge_dvt=self.merge_dvt)
        self.result_ready.emit(output)


class DragDropArea(QLabel):
    files_dropped = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setText("Kéo thả hoặc click vào đây để chọn file Excel KiotViet (.xlsx, .xls)\n\n(DanhSachSanPham..., DanhSachKhachHang..., DanhSachNhaCungCap...)")
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #999;
                border-radius: 10px;
                background-color: #f8f9fa;
                font-size: 16px;
                color: #555;
            }
        """)
        self.setAcceptDrops(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            files, _ = QFileDialog.getOpenFileNames(
                self,
                "Chọn file Excel KiotViet",
                "",
                "Excel Files (*.xlsx *.xls)"
            )
            if files:
                self.files_dropped.emit(files)
        else:
            super().mousePressEvent(event)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
            self.setStyleSheet("""
                QLabel {
                    border: 2px dashed #4CAF50;
                    border-radius: 10px;
                    background-color: #e8f5e9;
                    font-size: 16px;
                    color: #4CAF50;
                }
            """)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #999;
                border-radius: 10px;
                background-color: #f8f9fa;
                font-size: 16px;
                color: #555;
            }
        """)

    def dropEvent(self, event):
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #999;
                border-radius: 10px;
                background-color: #f8f9fa;
                font-size: 16px;
                color: #555;
            }
        """)
        
        files = []
        for url in event.mimeData().urls():
            # Support macOS and Windows file drops
            if url.isLocalFile():
                files.append(url.toLocalFile())
        
        if files:
            self.files_dropped.emit(files)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("KiotViet to MTP - Công cụ chuyển đổi")
        self.resize(700, 500)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        self.drop_area = DragDropArea()
        self.drop_area.files_dropped.connect(self.process_files)
        # Make the drop area take about 1/3 of the vertical space
        self.drop_area.setMinimumHeight(150)
        
        layout.addWidget(self.drop_area, stretch=1)

        self.merge_dvt_checkbox = QCheckBox("Gộp ĐVT phụ (Multi-ĐVT)")
        self.merge_dvt_checkbox.setToolTip("Khi bật, các dòng sản phẩm có cùng tên nhưng khác ĐVT sẽ được gộp thành ĐVT phụ trên cùng một dòng sản phẩm.")
        layout.addWidget(self.merge_dvt_checkbox)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setStyleSheet("""
            QTextEdit {
                background-color: #2b2b2b;
                color: #f1f1f1;
                font-family: monospace;
                padding: 10px;
                border-radius: 5px;
            }
        """)
        
        layout.addWidget(self.log_output, stretch=2)
        
        self.worker = None

    def process_files(self, files):
        # Filter for .xlsx and .xls
        excel_files = [f for f in files if f.lower().endswith(('.xlsx', '.xls'))]
        if not excel_files:
            self.log_output.append(">>> Không tìm thấy file .xlsx hay .xls nào trong danh sách được thả.\n")
            return
            
        self.log_output.append(f">>> Đang xử lý {len(excel_files)} file...")
        for f in excel_files:
            self.log_output.append(f" - {os.path.basename(f)}")
        self.log_output.append("")

        try:
            grouped = self.prepare_mapping_context(excel_files)
        except ValueError as exc:
            self.log_output.append(f">>> Lỗi: {exc}\n")
            QMessageBox.warning(self, "Không thể đọc file", str(exc))
            return

        dialog = ColumnMappingDialog(grouped, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            self.log_output.append(">>> Đã hủy trước khi chuyển đổi.\n")
            return
        
        # Disable drop area while processing
        self.drop_area.setAcceptDrops(False)
        self.worker = ConversionWorker(excel_files, dialog.column_mappings(), merge_dvt=self.merge_dvt_checkbox.isChecked())
        self.worker.result_ready.connect(self.on_processing_finished)
        self.worker.start()

    def prepare_mapping_context(self, files: list[str]) -> dict[str, list[SourceFileInfo]]:
        grouped: dict[str, list[SourceFileInfo]] = {}
        for file in files:
            path = Path(file)
            headers = read_excel_headers(path)
            source_type = detect_source_type(path, headers)
            grouped.setdefault(source_type, []).append(
                SourceFileInfo(path=path, source_type=source_type, headers=headers)
            )
        return grouped
        
    def on_processing_finished(self, output):
        self.log_output.append(output)
        self.log_output.append(">>> Xử lý hoàn tất.\n")
        # Re-enable drop area
        self.drop_area.setAcceptDrops(True)


def run_gui():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()
