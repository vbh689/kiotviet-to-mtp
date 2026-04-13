import sys
import os
import io
import contextlib
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, QTextEdit, QHBoxLayout, QPushButton
)
from PyQt6.QtCore import Qt, pyqtSignal, QThread

# Run conversion by patching sys.argv and redirecting stdout
def run_cli_conversion(files: list[str], output_dir: str = None):
    from .kv_runner import main
    # Save original args and stdout
    original_argv = sys.argv[:]
    
    # We build the argv as if it was called from CLI
    new_argv = ["kiotviet_to_mtp", "--kiotviet"] + files
    if output_dir:
        new_argv += ["--outdir", output_dir]
    
    sys.argv = new_argv
    
    f = io.StringIO()
    try:
        with contextlib.redirect_stdout(f), contextlib.redirect_stderr(f):
            # Run the actual CLI main function
            main()
    except SystemExit as e:
        # argparse can call sys.exit, we just catch it harmlessly
        pass
    except Exception as e:
        import traceback
        f.write(traceback.format_exc())
    finally:
        sys.argv = original_argv
        
    return f.getvalue()

class ConversionWorker(QThread):
    result_ready = pyqtSignal(str)

    def __init__(self, files):
        super().__init__()
        self.files = files

    def run(self):
        output = run_cli_conversion(self.files)
        self.result_ready.emit(output)


class DragDropArea(QLabel):
    files_dropped = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setText("Kéo thả file Excel KiotViet (.xlsx) vào đây\n\n(DanhSachSanPham..., DanhSachKhachHang..., DanhSachNhaCungCap...)")
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
        # Filter for .xlsx
        xlsx_files = [f for f in files if f.lower().endswith('.xlsx')]
        if not xlsx_files:
            self.log_output.append(">>> Không tìm thấy file .xlsx nào trong danh sách được thả.\n")
            return
            
        self.log_output.append(f">>> Đang xử lý {len(xlsx_files)} file...")
        for f in xlsx_files:
            self.log_output.append(f" - {os.path.basename(f)}")
        self.log_output.append("")
        
        # Disable drop area while processing
        self.drop_area.setAcceptDrops(False)
        self.worker = ConversionWorker(xlsx_files)
        self.worker.result_ready.connect(self.on_processing_finished)
        self.worker.start()
        
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
