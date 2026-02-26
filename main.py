import sys
import os
import filetype
import fitz  # PyMuPDF
from PIL import Image

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QComboBox, QProgressBar, QMessageBox
)
from PySide6.QtCore import QThread, Signal, Qt

class ConvertWorker(QThread):
    """Background thread to handle file conversion without freezing the UI."""
    progress = Signal(int)
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, input_file, output_dir, target_format):
        super().__init__()
        self.input_file = input_file
        self.output_dir = output_dir
        self.target_format = target_format.lower()

    def run(self):
        try:
            kind = filetype.guess(self.input_file)
            mime_type = kind.mime if kind else ""

            base_name = os.path.splitext(os.path.basename(self.input_file))[0]

            # --- PDF TO IMAGE CONVERSION ---
            if "pdf" in mime_type:
                doc = fitz.open(self.input_file)
                total_pages = len(doc)
                
                for i in range(total_pages):
                    page = doc.load_page(i)
                    # dpi=300 for high quality images
                    pix = page.get_pixmap(dpi=300)
                    out_path = os.path.join(self.output_dir, f"{base_name}_page_{i+1}.{self.target_format}")
                    pix.save(out_path)
                    
                    # Update progress bar
                    self.progress.emit(int(((i + 1) / total_pages) * 100))
                    
                doc.close()
                self.finished.emit("PDF successfully converted to images!")

            # --- IMAGE TO IMAGE CONVERSION ---
            elif "image" in mime_type:
                with Image.open(self.input_file) as img:
                    # Convert to RGB if saving as JPG to prevent transparency errors
                    if self.target_format in ['jpg', 'jpeg'] and img.mode in ('RGBA', 'LA', 'P'):
                        img = img.convert('RGB')
                    
                    out_path = os.path.join(self.output_dir, f"{base_name}_converted.{self.target_format}")
                    img.save(out_path)
                    self.progress.emit(100)
                    self.finished.emit("Image successfully converted!")
            else:
                self.error.emit("Unsupported file type for conversion.")

        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FormatForge - File Converter")
        self.resize(450, 300)

        self.input_file = None
        self.output_dir = None

        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # File Selection
        self.lbl_file = QLabel("No file selected.")
        self.lbl_file.setAlignment(Qt.AlignCenter)
        self.lbl_file.setStyleSheet("padding: 10px; border: 1px dashed #aaa;")
        layout.addWidget(self.lbl_file)

        btn_select_file = QPushButton("Select File")
        btn_select_file.clicked.connect(self.select_file)
        layout.addWidget(btn_select_file)

        # Auto-Detected Type
        self.lbl_detected = QLabel("Detected Type: N/A")
        layout.addWidget(self.lbl_detected)

        # Format Selection
        layout.addWidget(QLabel("Select Target Format:"))
        self.combo_format = QComboBox()
        self.combo_format.setEnabled(False)
        layout.addWidget(self.combo_format)

        # Output Directory
        self.lbl_dir = QLabel("Output Folder: Not selected (Defaults to input folder)")
        layout.addWidget(self.lbl_dir)
        
        btn_select_dir = QPushButton("Select Output Folder")
        btn_select_dir.clicked.connect(self.select_output_dir)
        layout.addWidget(btn_select_dir)

        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # Convert Button
        self.btn_convert = QPushButton("Convert")
        self.btn_convert.setEnabled(False)
        self.btn_convert.setStyleSheet("background-color: #2b5797; color: white; padding: 10px;")
        self.btn_convert.clicked.connect(self.start_conversion)
        layout.addWidget(self.btn_convert)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File to Convert")
        if file_path:
            self.input_file = file_path
            self.lbl_file.setText(f"Selected: {os.path.basename(file_path)}")
            
            # Default output directory to the same folder as input
            if not self.output_dir:
                self.output_dir = os.path.dirname(file_path)
                self.lbl_dir.setText(f"Output Folder: {self.output_dir}")

            self.detect_file_type()

    def select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if dir_path:
            self.output_dir = dir_path
            self.lbl_dir.setText(f"Output Folder: {self.output_dir}")

    def detect_file_type(self):
        kind = filetype.guess(self.input_file)
        self.combo_format.clear()

        if kind is None:
            self.lbl_detected.setText("Detected Type: Unknown (Fallback to extension)")
            ext = os.path.splitext(self.input_file)[1].lower()
            if ext == '.pdf':
                mime_type = 'application/pdf'
            else:
                mime_type = 'unknown'
        else:
            mime_type = kind.mime
            self.lbl_detected.setText(f"Detected Type: {mime_type}")

        # Populate combo box based on detected type
        if "pdf" in mime_type:
            self.combo_format.addItems(["png", "jpg", "jpeg", "bmp"])
            self.combo_format.setEnabled(True)
            self.btn_convert.setEnabled(True)
        elif "image" in mime_type:
            self.combo_format.addItems(["png", "jpg", "jpeg", "webp", "bmp"])
            self.combo_format.setEnabled(True)
            self.btn_convert.setEnabled(True)
        else:
            self.combo_format.setEnabled(False)
            self.btn_convert.setEnabled(False)
            QMessageBox.warning(self, "Unsupported", "This file type is not currently supported for conversion.")

    def start_conversion(self):
        target_fmt = self.combo_format.currentText()
        
        self.btn_convert.setEnabled(False)
        self.progress_bar.setValue(0)

        # Setup and start background thread
        self.worker = ConvertWorker(self.input_file, self.output_dir, target_fmt)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.conversion_finished)
        self.worker.error.connect(self.conversion_error)
        self.worker.start()

    def update_progress(self, val):
        self.progress_bar.setValue(val)

    def conversion_finished(self, msg):
        self.btn_convert.setEnabled(True)
        QMessageBox.information(self, "Success", msg)

    def conversion_error(self, err_msg):
        self.btn_convert.setEnabled(True)
        QMessageBox.critical(self, "Error", f"An error occurred:\n{err_msg}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Set a clean, modern style
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())