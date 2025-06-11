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
        # self.original_page_pixmap is no longer needed
        self.page_width = 0
        self.page_height = 0
        self.template_path = ""
        self.csv_path = ""
        self.output_folder = ""
        self.certificate_data = [] # Will hold list of (name, achievement) tuples

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
        
        # Create Name Tab
        self.name_widget = QWidget()
        self.name_text, self.name_x, self.name_y, self.name_size, self.name_rot = self.create_positioning_controls()
        self.name_widget.setLayout(self.create_form_layout_for_controls("Name", self.name_text, self.name_x, self.name_y, self.name_size, self.name_rot))
        self.tabs.addTab(self.name_widget, "Name Settings")
        
        # Create Achievement Tab
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
        """Creates a set of controls (text, x, y, size, rot) for one text field."""
        text_entry = QLineEdit()
        text_entry.textChanged.connect(self.update_display)
        
        x_slider, _ = self.create_slider(0, 1000)
        y_slider, _ = self.create_slider(0, 1000)
        size_slider, _ = self.create_slider(8, 150, 24)
        rot_slider, _ = self.create_slider(-45, 45, 0)
        
        return text_entry, x_slider, y_slider, size_slider, rot_slider

    def create_form_layout_for_controls(self, group_name, text_widget, x_slider, y_slider, size_slider, rot_slider):
        """Helper to arrange a set of controls into a QFormLayout."""
        layout = QFormLayout()
        layout.addRow(f"{group_name} Text:", text_widget)
        layout.addRow("X Position:", x_slider)
        layout.addRow("Y Position:", y_slider)
        layout.addRow("Font Size:", size_slider)
        layout.addRow("Rotation:", rot_slider)
        return layout

    def create_slider(self, min_val, max_val, initial_val=0):
        """Helper to create a QSlider."""
        slider = QSlider(Qt.Orientation.Horizontal)
        slider.setRange(min_val, max_val)
        slider.setValue(initial_val)
        slider.valueChanged.connect(self.update_display)
        # The label is now part of the update_display logic
        return slider, QLabel(f"Value: {initial_val}")

    # --- File Selection and Parsing ---

    def select_template_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Template PDF", "", "PDF Files (*.pdf)")
        if not path:
            return
            
        self.template_path = path
        self.template_label.setText(os.path.basename(path))
        self.template_label.setToolTip(path)
        try:
            # Close previous document if one is open
            if self.doc_template:
                self.doc_template.close()
            self.doc_template = fitz.open(self.template_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open PDF file: {e}")
            return
        
        if not self.doc_template.page_count:
            QMessageBox.critical(self, "Error", "This PDF has no pages.")
            return

        page = self.doc_template[0]
        self.page_width = int(page.rect.width)
        self.page_height = int(page.rect.height)

        # Update all sliders
        for slider in [self.name_x, self.ach_x]:
            slider.setRange(0, self.page_width)
            slider.setValue(self.page_width // 2)
        for slider in [self.name_y, self.ach_y]:
            slider.setRange(0, self.page_height)
            slider.setValue(self.page_height // 3)
        self.ach_y.setValue(int(self.page_height * 0.6))
        
        # No longer need to store the original pixmap. update_display will handle all rendering.
        self.update_display()

    def select_csv_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Data CSV", "", "CSV Files (*.csv)")
        if not path:
            return
        
        self.csv_path = path
        self.csv_label.setText(os.path.basename(path))
        self.csv_label.setToolTip(path)
        self.parse_csv()

    def select_output_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not path:
            return
        self.output_folder = path
        self.output_label.setText(os.path.basename(path))
        self.output_label.setToolTip(path)

    def parse_csv(self):
        """Reads the CSV, skipping headers and populating the data list."""
        self.certificate_data = []
        try:
            with open(self.csv_path, mode='r', encoding='utf-8') as infile:
                reader = csv.reader(infile)
                # Skip the first 3 header lines based on the provided format
                next(reader)  # Skips "FY 2024-25,"
                next(reader)  # Skips "CHAMPIONS RECOGNITION,"
                next(reader)  # Skips "name ,achievement"
                
                for row in reader:
                    # Improved row checking
                    if not row or len(row) < 2 or not row[0].strip():
                         continue
                    
                    name = row[0].strip()
                    # Handle multiline achievement by checking if row[1] is empty
                    achievement = row[1].strip() if row[1].strip() else row[0].strip()

                    self.certificate_data.append((name, achievement))

        except Exception as e:
            QMessageBox.critical(self, "CSV Error", f"Failed to read or parse CSV file:\n{e}")
            return
            
        if not self.certificate_data:
            QMessageBox.warning(self, "CSV Warning", "No valid data (name, achievement) found in the CSV.")
            return
        
        # Load first entry into the text fields for live preview
        first_name, first_ach = self.certificate_data[0]
        self.name_text.setText(first_name)
        self.ach_text.setText(first_ach)
        # update_display will be triggered by the setText signals

    # --- Core Display and Generation Logic ---
    
    def update_display(self, _=None):
        """Redraws the text on the page image based on all control values."""
        # --- FIX: THIS IS THE CORRECTED LOGIC ---
        if self.doc_template is None:
            return

        # 1. Create a temporary, in-memory document and page
        temp_doc = fitz.open()
        temp_page = temp_doc.new_page(width=self.page_width, height=self.page_height)
        
        # 2. Use the original template page as a background for our temp page
        # This stamps the content of the first page of the template onto our temp page
        temp_page.show_pdf_page(temp_page.rect, self.doc_template, 0)
        
        # 3. Now, insert text onto the temporary PAGE object, not a pixmap
        # Get values for Name
        rect_name = fitz.Rect(self.name_x.value(), self.name_y.value(), self.page_width, self.page_height)
        temp_page.insert_textbox(rect_name, self.name_text.text(), 
                                fontsize=self.name_size.value(), fontname="helv-bold", 
                                color=(0, 0, 0), align=fitz.TEXT_ALIGN_CENTER,
                                rotate=self.name_rot.value())

        # Get values for Achievement
        rect_ach = fitz.Rect(self.ach_x.value(), self.ach_y.value(), self.page_width, self.page_height)
        temp_page.insert_textbox(rect_ach, self.ach_text.text(),
                                fontsize=self.ach_size.value(), fontname="helv", 
                                color=(0, 0, 0), align=fitz.TEXT_ALIGN_CENTER,
                                rotate=self.ach_rot.value())
        
        # 4. Render the MODIFIED temporary page to a pixmap
        temp_pix = temp_page.get_pixmap(alpha=False)
        temp_doc.close() # Clean up the temporary document

        # 5. Convert fitz.Pixmap to QImage for display (this part is the same)
        q_image = QImage(temp_pix.samples, temp_pix.width, temp_pix.height, temp_pix.stride, QImage.Format.Format_RGB888)
        pixmap = QPixmap.fromImage(q_image)

        scaled_pixmap = pixmap.scaled(self.image_label.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.image_label.setPixmap(scaled_pixmap)


    def generate_all_certificates(self):
        """Saves a new PDF for each entry in the CSV data."""
        if not all([self.template_path, self.csv_path, self.output_folder]):
            QMessageBox.warning(self, "Warning", "Please select a template, a CSV file, and an output folder before generating.")
            return
        if not self.certificate_data:
            QMessageBox.warning(self, "Warning", "No data loaded from the CSV file.")
            return

        name_pos = (self.name_x.value(), self.name_y.value())
        name_size = self.name_size.value()
        name_rot = self.name_rot.value()
        
        ach_pos = (self.ach_x.value(), self.ach_y.value())
        ach_size = self.ach_size.value()
        ach_rot = self.ach_rot.value()

        count = 0
        try:
            for name, achievement in self.certificate_data:
                doc = fitz.open(self.template_path)
                page = doc[0]

                # Insert Name
                page.insert_textbox(fitz.Rect(name_pos[0], name_pos[1], self.page_width, self.page_height),
                                    name, fontsize=name_size, fontname="helv-bold",
                                    color=(0, 0, 0), align=fitz.TEXT_ALIGN_CENTER, rotate=name_rot)

                # Insert Achievement
                page.insert_textbox(fitz.Rect(ach_pos[0], ach_pos[1], self.page_width, self.page_height),
                                    achievement, fontsize=ach_size, fontname="helv",
                                    color=(0, 0, 0), align=fitz.TEXT_ALIGN_CENTER, rotate=ach_rot)

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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PdfCertificateGenerator()
    window.show()
    sys.exit(app.exec())