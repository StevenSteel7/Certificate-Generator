import sys
import os
import csv
import re
import fitz  # PyMuPDF

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QPushButton, QLabel, QLineEdit, QSlider, QFileDialog, QMessageBox, QTabWidget
)
from PyQt6.QtGui import QPixmap, QImage, QFont
from PyQt6.QtCore import Qt, QSize

class PdfCertificateGenerator(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("PDF Certificate Batch Generator (PyQt6)")
        self.setGeometry(100, 100, 1300, 850)

        # --- Data Attributes ---
        self.doc_template = None
        self.page_width = 0
        self.page_height = 0
        self.template_path = ""
        self.csv_path = ""
        self.output_folder = ""
        self.certificate_data = []

        # Create font objects to access their binary data later
        self.font_name_bold = fitz.Font("Helvetica-Bold")
        self.font_achievement_reg = fitz.Font("Helvetica")

        # --- Main Layout ---
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # --- Controls Frame (Left) ---
        controls_widget = QWidget()
        controls_widget.setFixedWidth(400)
        controls_layout = QVBoxLayout(controls_widget)
        main_layout.addWidget(controls_widget)

        # --- File I/O Section ---
        io_group = QWidget()
        io_layout = QFormLayout(io_group)
        io_layout.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapAllRows)
        
        self.template_btn = QPushButton("Browse...")
        self.template_label = QLabel("Not selected")
        self.template_btn.clicked.connect(self.select_template_pdf)
        io_layout.addRow("1. Select Template PDF:", self.template_btn)
        io_layout.addRow(self.template_label)
        
        self.csv_btn = QPushButton("Browse...")
        self.csv_label = QLabel("Not selected")
        self.csv_btn.clicked.connect(self.select_csv_file)
        io_layout.addRow("2. Select Data CSV:", self.csv_btn)
        io_layout.addRow(self.csv_label)
        
        self.output_btn = QPushButton("Browse...")
        self.output_label = QLabel("Not selected")
        self.output_btn.clicked.connect(self.select_output_folder)
        io_layout.addRow("3. Select Output Folder:", self.output_btn)
        io_layout.addRow(self.output_label)
        
        controls_layout.addWidget(io_group)

        # --- Tabbed Controls for Text Positioning ---
        self.tabs = QTabWidget()
        
        self.name_widget = QWidget()
        self.name_text, self.name_x, self.name_y, self.name_size, self.name_rot = self.create_positioning_controls()
        self.name_widget.setLayout(self.create_form_layout_for_controls("Name", self.name_text, self.name_x, self.name_y, self.name_size, self.name_rot))
        self.tabs.addTab(self.name_widget, "Name Settings")
        
        self.ach_widget = QWidget()
        self.ach_text, self.ach_x, self.ach_y, self.ach_size, self.ach_rot = self.create_positioning_controls()
        self.ach_widget.setLayout(self.create_form_layout_for_controls("Achievement", self.ach_text, self.ach_x, self.ach_y, self.ach_size, self.ach_rot))
        self.tabs.addTab(self.ach_widget, "Achievement Settings")
        
        controls_layout.addWidget(self.tabs)
        controls_layout.addStretch()

        # --- Generate Button ---
        self.generate_button = QPushButton("Generate & Save All Certificates")
        self.generate_button.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.generate_button.setStyleSheet("background-color: #4CAF50; color: white; padding: 10px;")
        self.generate_button.clicked.connect(self.generate_all_certificates)
        controls_layout.addWidget(self.generate_button)

        # --- Image Display Label (Right) ---
        self.image_label = QLabel("Please select a Template PDF to begin")
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.setStyleSheet("background-color: #333333; color: white;")
        main_layout.addWidget(self.image_label, 1)

    def create_positioning_controls(self):
        text_entry = QLineEdit()
        text_entry.textChanged.connect(self.update_display)
        x_slider, _ = self.create_slider(0, 1000)
        y_slider, _ = self.create_slider(0, 1000)
        size_slider, _ = self.create_slider(8, 150, 24)
        rot_slider, _ = self.create_slider(-45, 45, 0)
        return text_entry, x_slider, y_slider, size_slider, rot_slider

    def create_form_layout_for_controls(self, group_name, text_widget, x_slider, y_slider, size_slider, rot_slider):
        layout = QFormLayout()
        layout.addRow(f"{group_name} Text:", text_widget)
        layout.addRow("X Position:", x_slider)
        layout.addRow("Y Position:", y_slider)
        layout.addRow("Font Size:", size_slider)
        layout.addRow("Rotation:", rot_slider)
        return layout

    def create_slider(self, min_val, max_val, initial_val=0):
        slider = QSlider(Qt.Orientation.Horizontal)
        slider.setRange(min_val, max_val)
        slider.setValue(initial_val)
        slider.valueChanged.connect(self.update_display)
        return slider, QLabel(f"Value: {initial_val}")

    def select_template_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Template PDF", "", "PDF Files (*.pdf)")
        if not path: return
        
        self.template_path = path
        self.template_label.setText(os.path.basename(path))
        self.template_label.setToolTip(path)
        try:
            if self.doc_template: self.doc_template.close()
            self.doc_template = fitz.open(self.template_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open PDF file: {e}"); return
        
        if not self.doc_template.page_count:
            QMessageBox.critical(self, "Error", "This PDF has no pages."); return

        page = self.doc_template[0]
        self.page_width = int(page.rect.width)
        self.page_height = int(page.rect.height)

        for slider in [self.name_x, self.ach_x]:
            slider.setRange(0, self.page_width); slider.setValue(self.page_width // 2)
        for slider in [self.name_y, self.ach_y]:
            slider.setRange(0, self.page_height); slider.setValue(self.page_height // 3)
        self.ach_y.setValue(int(self.page_height * 0.6))
        self.update_display()

    def select_csv_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Data CSV", "", "CSV Files (*.csv)")
        if not path: return
        self.csv_path = path
        self.csv_label.setText(os.path.basename(path))
        self.csv_label.setToolTip(path)
        self.parse_csv()

    def select_output_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not path: return
        self.output_folder = path
        self.output_label.setText(os.path.basename(path))
        self.output_label.setToolTip(path)

    def parse_csv(self):
        self.certificate_data = []
        try:
            with open(self.csv_path, mode='r', encoding='utf-8') as infile:
                reader = csv.reader(infile)
                for _ in range(3): next(reader)
                
                for row in reader:
                    if not row or (not row[0].strip() and not row[1].strip()):
                        continue
                    
                    name = row[0].strip()
                    achievement = row[1].strip()

                    if name:
                        self.certificate_data.append([name, achievement])
                    elif achievement and self.certificate_data:
                        self.certificate_data[-1][1] += f"\n{achievement}"

        except Exception as e:
            QMessageBox.critical(self, "CSV Error", f"Failed to read or parse CSV file:\n{e}"); return
            
        if not self.certificate_data:
            QMessageBox.warning(self, "CSV Warning", "No valid data found in CSV."); return
        
        first_name, first_ach = self.certificate_data[0]
        self.name_text.setText(first_name)
        self.ach_text.setText(first_ach)

    def update_display(self, _=None):
        if self.doc_template is None: return

        temp_doc = fitz.open()
        temp_page = temp_doc.new_page(width=self.page_width, height=self.page_height)
        temp_page.show_pdf_page(temp_page.rect, self.doc_template, 0)
        
        temp_page.insert_font(fontname="F0", fontbuffer=self.font_name_bold.buffer)
        temp_page.insert_font(fontname="F1", fontbuffer=self.font_achievement_reg.buffer)

        rect_name = fitz.Rect(self.name_x.value(), self.name_y.value(), self.page_width, self.page_height)
        temp_page.insert_textbox(rect_name, self.name_text.text(), 
                                fontsize=self.name_size.value(),
                                fontname="F0",
                                color=(1, 1, 1),  # Set color to white
                                align=fitz.TEXT_ALIGN_CENTER,
                                rotate=self.name_rot.value())

        rect_ach = fitz.Rect(self.ach_x.value(), self.ach_y.value(), self.page_width, self.page_height)
        temp_page.insert_textbox(rect_ach, self.ach_text.text(),
                                fontsize=self.ach_size.value(),
                                fontname="F1",
                                color=(1, 1, 1),  # Set color to white
                                align=fitz.TEXT_ALIGN_CENTER,
                                rotate=self.ach_rot.value())
        
        temp_pix = temp_page.get_pixmap(alpha=False)
        temp_doc.close()

        q_image = QImage(temp_pix.samples, temp_pix.width, temp_pix.height, temp_pix.stride, QImage.Format.Format_RGB888)
        pixmap = QPixmap.fromImage(q_image)

        scaled_pixmap = pixmap.scaled(self.image_label.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.image_label.setPixmap(scaled_pixmap)

    def generate_all_certificates(self):
        if not all([self.template_path, self.csv_path, self.output_folder]):
            QMessageBox.warning(self, "Warning", "Please select a template, a CSV file, and an output folder."); return
        if not self.certificate_data:
            QMessageBox.warning(self, "Warning", "No data loaded from the CSV file."); return

        name_rect = fitz.Rect(self.name_x.value(), self.name_y.value(), self.page_width, self.page_height)
        name_size = self.name_size.value()
        name_rot = self.name_rot.value()
        
        ach_rect = fitz.Rect(self.ach_x.value(), self.ach_y.value(), self.page_width, self.page_height)
        ach_size = self.ach_size.value()
        ach_rot = self.ach_rot.value()

        count = 0
        try:
            for name, achievement in self.certificate_data:
                doc = fitz.open(self.template_path)
                page = doc[0]

                page.insert_font(fontname="F0-name", fontbuffer=self.font_name_bold.buffer)
                page.insert_font(fontname="F1-ach", fontbuffer=self.font_achievement_reg.buffer)
                
                page.insert_textbox(name_rect, name, fontsize=name_size,
                                    fontname="F0-name",
                                    color=(1, 1, 1),  # Set color to white
                                    align=fitz.TEXT_ALIGN_CENTER, rotate=name_rot)
                
                page.insert_textbox(ach_rect, achievement, fontsize=ach_size,
                                    fontname="F1-ach",
                                    color=(1, 1, 1),  # Set color to white
                                    align=fitz.TEXT_ALIGN_CENTER, rotate=ach_rot)

                safe_name = re.sub(r'[\\/*?:"<>|]', "_", name)
                output_path = os.path.join(self.output_folder, f"Certificate - {safe_name}.pdf")
                doc.save(output_path, garbage=4, deflate=True)
                doc.close()
                count += 1
            
            QMessageBox.information(self, "Success", f"Successfully generated and saved {count} certificates to:\n{self.output_folder}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred during PDF generation:\n{e}")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_display()

    def closeEvent(self, event):
        self.font_name_bold.clean_font_data()
        self.font_achievement_reg.clean_font_data()
        if self.doc_template:
            self.doc_template.close()
        super().closeEvent(event)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PdfCertificateGenerator()
    window.show()
    sys.exit(app.exec())