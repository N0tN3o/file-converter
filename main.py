import sys
import os
import io
import zipfile
import filetype
import fitz  # PyMuPDF
import docx  # python-docx
from PIL import Image

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QComboBox, QProgressBar, QMessageBox
)
from PySide6.QtCore import QThread, Signal, Qt

# Mapping of what source formats can be converted into
CONVERSIONS = {
    "PDF": ["png", "jpg", "jpeg", "txt"],
    "Image": ["png", "jpg", "jpeg", "webp", "bmp", "gif"],
    "DOCX": ["txt", "md"],
    "TXT": ["docx", "md"],
    "MD": ["txt", "docx"]
}

class ConvertWorker(QThread):
    """Background thread to handle file conversions."""
    progress = Signal(int)
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, input_file, output_dir, source_format, target_format):
        super().__init__()
        self.input_file = input_file
        self.output_dir = output_dir
        self.source_format = source_format
        self.target_format = target_format.lower()

    def run(self):
        try:
            base_name = os.path.splitext(os.path.basename(self.input_file))[0]

            # ==========================================
            # 1. PDF CONVERSIONS
            # ==========================================
            if self.source_format == "PDF":
                doc = fitz.open(self.input_file)
                total_pages = len(doc)

                # PDF to Text
                if self.target_format == "txt":
                    text_content = ""
                    for i in range(total_pages):
                        text_content += doc.load_page(i).get_text()
                        self.progress.emit(int(((i + 1) / total_pages) * 100))
                    
                    out_path = os.path.join(self.output_dir, f"{base_name}.txt")
                    with open(out_path, 'w', encoding='utf-8') as f:
                        f.write(text_content)
                    self.finished.emit("PDF successfully extracted to Text!")

                # PDF to Image(s)
                elif self.target_format in ["png", "jpg", "jpeg"]:
                    if total_pages > 1:
                        # Output as ZIP
                        zip_path = os.path.join(self.output_dir, f"{base_name}_images.zip")
                        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                            for i in range(total_pages):
                                page = doc.load_page(i)
                                pix = page.get_pixmap(dpi=300)
                                # fitz tobytes supports png, jpeg, etc.
                                img_data = pix.tobytes(self.target_format)
                                zipf.writestr(f"{base_name}_page_{i+1}.{self.target_format}", img_data)
                                self.progress.emit(int(((i + 1) / total_pages) * 100))
                        self.finished.emit("Multiple PDF pages successfully zipped as images!")
                    else:
                        # Output single image
                        page = doc.load_page(0)
                        pix = page.get_pixmap(dpi=300)
                        out_path = os.path.join(self.output_dir, f"{base_name}.{self.target_format}")
                        pix.save(out_path)
                        self.progress.emit(100)
                        self.finished.emit("Single PDF page successfully converted to image!")
                doc.close()

            # ==========================================
            # 2. IMAGE CONVERSIONS
            # ==========================================
            elif self.source_format == "Image":
                with Image.open(self.input_file) as img:
                    if self.target_format in ['jpg', 'jpeg'] and img.mode in ('RGBA', 'LA', 'P'):
                        img = img.convert('RGB')
                    out_path = os.path.join(self.output_dir, f"{base_name}_converted.{self.target_format}")
                    img.save(out_path)
                    self.progress.emit(100)
                    self.finished.emit(f"Image converted to {self.target_format.upper()}!")

            # ==========================================
            # 3. DOCX CONVERSIONS
            # ==========================================
            elif self.source_format == "DOCX":
                doc = docx.Document(self.input_file)
                text_content = "\n".join([p.text for p in doc.paragraphs])
                
                out_path = os.path.join(self.output_dir, f"{base_name}.{self.target_format}")
                with open(out_path, 'w', encoding='utf-8') as f:
                    f.write(text_content)
                self.progress.emit(100)
                self.finished.emit(f"DOCX converted to {self.target_format.upper()}!")

            # ==========================================
            # 4. TEXT / MARKDOWN CONVERSIONS
            # ==========================================
            elif self.source_format in ["TXT", "MD"]:
                with open(self.input_file, 'r', encoding='utf-8') as f:
                    text_content = f.read()

                if self.target_format == "docx":
                    doc = docx.Document()
                    doc.add_paragraph(text_content)
                    out_path = os.path.join(self.output_dir, f"{base_name}.docx")
                    doc.save(out_path)
                else:
                    # Simple text-to-text / text-to-md save
                    out_path = os.path.join(self.output_dir, f"{base_name}.{self.target_format}")
                    with open(out_path, 'w', encoding='utf-8') as f:
                        f.write(text_content)
                
                self.progress.emit(100)
                self.finished.emit(f"File converted to {self.target_format.upper()}!")

        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FormatForge - Advanced File Converter")
        self.resize(500, 350)
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

        # Source Format Override
        layout.addWidget(QLabel("Source Format (Auto-detected, but you can change it):"))
        self.combo_source = QComboBox()
        self.combo_source.addItems(["Auto-Detecting..."] + list(CONVERSIONS.keys()))
        self.combo_source.setEnabled(False)
        self.combo_source.currentTextChanged.connect(self.update_target_formats)
        layout.addWidget(self.combo_source)

        # Target Format Selection
        layout.addWidget(QLabel("Select Target Format:"))
        self.combo_target = QComboBox()
        self.combo_target.setEnabled(False)
        layout.addWidget(self.combo_target)

        # Output Directory
        self.lbl_dir = QLabel("Output Folder: Not selected")
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
        self.combo_source.setEnabled(True)
        kind = filetype.guess(self.input_file)
        ext = os.path.splitext(self.input_file)[1].lower()

        # Combine magic numbers (filetype) and extension fallback
        detected_type = None
        if kind and "pdf" in kind.mime:
            detected_type = "PDF"
        elif kind and "image" in kind.mime:
            detected_type = "Image"
        elif ext == ".pdf":
            detected_type = "PDF"
        elif ext in [".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"]:
            detected_type = "Image"
        elif ext == ".docx":
            detected_type = "DOCX"
        elif ext == ".md":
            detected_type = "MD"
        elif ext == ".txt":
            detected_type = "TXT"

        if detected_type:
            # Set the dropdown to the detected type
            index = self.combo_source.findText(detected_type)
            if index >= 0:
                self.combo_source.setCurrentIndex(index)
        else:
            QMessageBox.warning(self, "Unknown File", "Could not reliably detect file type. Please manually select the Source Format.")
            self.combo_source.setCurrentIndex(1) # Default to first valid item (PDF)

    def update_target_formats(self, source_format):
        self.combo_target.clear()
        if source_format in CONVERSIONS:
            self.combo_target.addItems(CONVERSIONS[source_format])
            self.combo_target.setEnabled(True)
            self.btn_convert.setEnabled(True)
        else:
            self.combo_target.setEnabled(False)
            self.btn_convert.setEnabled(False)

    def start_conversion(self):
        source_fmt = self.combo_source.currentText()
        target_fmt = self.combo_target.currentText()
        
        self.btn_convert.setEnabled(False)
        self.progress_bar.setValue(0)

        self.worker = ConvertWorker(self.input_file, self.output_dir, source_fmt, target_fmt)
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
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())