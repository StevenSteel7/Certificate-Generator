import sys
import os
import csv
import re
import math # Needed for rotation calculation
import fitz  # PyMuPDF

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QPushButton, QLabel, QLineEdit, QSlider, QFileDialog, QMessageBox, QTabWidget,
    QCheckBox, QSpinBox
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

        # Create font objects
        self.font_name_bold = fitz.Font("Helvetica-Bold")

        # --- Load Achievement Font (Custom with Fallback) ---
        try:
            # IMPORTANT: The font file must be in the same directory as the script.
            # Change the filename below if yours is different (e.g., "Brixton Medium.ttf")
            font_filename = "Brixton_Medium.otf"
            script_dir = os.path.dirname(os.path.abspath(__file__))
            font_path = os.path.join(script_dir, font_filename)

            # Try to load the custom font from the file
            self.font_achievement_reg = fitz.Font(filename=font_path)
            print(f"Successfully loaded custom font: {font_filename}")

        except Exception as e:
            # If loading fails, print a warning and fall back to the default Helvetica font
            print(f"WARNING: Could not load custom font '{font_filename}'. Falling back to Helvetica.")
            print(f"Error details: {e}")
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
        name_controls = self.create_positioning_controls()
        self.name_text, self.name_x, self.name_y, self.name_size, self.name_rot = name_controls
        self.name_widget.setLayout(self.create_form_layout_for_controls("Name", name_controls))
        self.tabs.addTab(self.name_widget, "Name Settings")
        
        self.ach_widget = QWidget()
        ach_controls = self.create_positioning_controls(add_wh_sliders=True)
        self.ach_text, self.ach_x, self.ach_y, self.ach_size, self.ach_rot, self.ach_w, self.ach_h = ach_controls
        ach_layout = self.create_form_layout_for_controls("Achievement", ach_controls, has_wh_sliders=True)
        
        # Add underline spacing control to achievement tab
        self.underline_spacing = QSpinBox()
        self.underline_spacing.setRange(0, 20)
        self.underline_spacing.setValue(0)  # Default spacing
        self.underline_spacing.setSuffix(" px")
        self.underline_spacing.setToolTip("Distance between text and underline (in pixels)")
        self.underline_spacing.valueChanged.connect(self.update_display)
        ach_layout.addRow("Underline Spacing:", self.underline_spacing)
        
        self.ach_widget.setLayout(ach_layout)
        self.tabs.addTab(self.ach_widget, "Achievement Settings")
        
        controls_layout.addWidget(self.tabs)
        
        self.autoresize_checkbox = QCheckBox("Auto-resize text to fit boxes")
        self.autoresize_checkbox.setChecked(True)
        self.autoresize_checkbox.setToolTip("If checked, font size will be reduced automatically to fit the text box during generation.")
        controls_layout.addWidget(self.autoresize_checkbox)
        
        controls_layout.addStretch()

        self.generate_button = QPushButton("Generate & Save All Certificates")
        self.generate_button.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        self.generate_button.setStyleSheet("background-color: #4CAF50; color: white; padding: 10px;")
        self.generate_button.clicked.connect(self.generate_all_certificates)
        controls_layout.addWidget(self.generate_button)

        self.image_label = QLabel("Please select a Template PDF to begin")
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.setStyleSheet("background-color: #333333; color: white;")
        main_layout.addWidget(self.image_label, 1)

    def create_positioning_controls(self, add_wh_sliders=False):
        text_entry = QLineEdit()
        text_entry.textChanged.connect(self.update_display)
        x_slider, _ = self.create_slider(0, 1000)
        y_slider, _ = self.create_slider(0, 1000)
        size_slider, _ = self.create_slider(8, 150, 36)
        rot_slider, _ = self.create_slider(-45, 45, 0)
        
        if add_wh_sliders:
            w_slider, _ = self.create_slider(50, 1000)
            h_slider, _ = self.create_slider(50, 1000)
            return text_entry, x_slider, y_slider, size_slider, rot_slider, w_slider, h_slider
            
        return text_entry, x_slider, y_slider, size_slider, rot_slider

    def create_form_layout_for_controls(self, group_name, controls, has_wh_sliders=False):
        layout = QFormLayout()
        text, x, y, size, rot = controls[:5]
        
        layout.addRow(f"{group_name} Text:", text)
        layout.addRow("X Position:", x)
        layout.addRow("Y Position:", y)
        if has_wh_sliders:
            w, h = controls[5:]
            layout.addRow("Box Width:", w)
            layout.addRow("Box Height:", h)
        layout.addRow("Font Size:", size)
        layout.addRow("Rotation:", rot)
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

        for slider in [self.name_x, self.ach_x, self.ach_w]:
            slider.setRange(0, self.page_width)
        for slider in [self.name_y, self.ach_y, self.ach_h]:
            slider.setRange(0, self.page_height)
            
        self.name_x.setValue(self.page_width // 2)
        self.name_y.setValue(self.page_height // 3)
        self.ach_x.setValue(int(self.page_width * 0.1))
        self.ach_y.setValue(int(self.page_height * 0.6))
        self.ach_w.setValue(int(self.page_width * 0.8))
        self.ach_h.setValue(int(self.page_height * 0.2))
        
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
            with open(self.csv_path, mode='r', encoding='utf-8-sig') as infile:
                reader = csv.reader(infile)
                for _ in range(3):
                    try:
                        next(reader)
                    except StopIteration:
                        break
                
                for row in reader:
                    if not row or (len(row) < 2) or (not row[0].strip() and not row[1].strip()):
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
        self.insert_text_with_autoresize(temp_page, rect_name, self.name_text.text(), 
                                         self.font_name_bold.buffer, "F0",
                                         self.name_size.value(), self.name_rot.value(),
                                         underline=False)

        ach_x = self.ach_x.value()
        ach_y = self.ach_y.value()
        ach_w = self.ach_w.value()
        ach_h = self.ach_h.value()
        rect_ach = fitz.Rect(ach_x, ach_y, ach_x + ach_w, ach_y + ach_h)

        temp_page.draw_rect(rect_ach, color=(1, 0, 0), width=1.5)
        self.insert_text_with_autoresize(temp_page, rect_ach, self.ach_text.text(), 
                                         self.font_achievement_reg.buffer, "F1",
                                         self.ach_size.value(), self.ach_rot.value(),
                                         underline=True, underline_spacing=self.underline_spacing.value())
        
        temp_pix = temp_page.get_pixmap(alpha=False)
        temp_doc.close()

        q_image = QImage(temp_pix.samples, temp_pix.width, temp_pix.height, temp_pix.stride, QImage.Format.Format_RGB888)
        pixmap = QPixmap.fromImage(q_image)

        scaled_pixmap = pixmap.scaled(self.image_label.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.image_label.setPixmap(scaled_pixmap)

    def insert_text_with_autoresize(self, page, rect, text, font_buffer, font_alias, initial_fontsize, rotate, underline=False, underline_spacing=3, min_fontsize=8):
        final_size = initial_fontsize
        temp_font = fitz.Font(fontname=font_alias, fontbuffer=font_buffer)

        if self.autoresize_checkbox.isChecked():
            line_height_factor = temp_font.ascender - temp_font.descender
            while final_size >= min_fontsize:
                if '\n' not in text:
                    text_width = temp_font.text_length(text, fontsize=final_size)
                    if text_width <= rect.width:
                        break
                else:
                    lines = text.split('\n')
                    max_width = 0
                    for line in lines:
                        line_width = temp_font.text_length(line, fontsize=final_size)
                        if line_width > max_width:
                            max_width = line_width
                    total_height = len(lines) * line_height_factor * final_size
                    if max_width <= rect.width and total_height <= rect.height:
                        break
                final_size -= 1

            if final_size < min_fontsize:
                final_size = min_fontsize

        page.insert_textbox(
            rect, text,
            fontname=font_alias,
            fontsize=final_size,
            color=(1, 1, 1),
            align=fitz.TEXT_ALIGN_CENTER,
            rotate=rotate
        )

        if underline and text.strip():
            found_rects = page.search_for(text, clip=rect, quads=False)
            if found_rects:
                actual_text_rect = found_rects[-1]
                self.add_underline_to_text(page, actual_text_rect, text, temp_font, final_size, rotate, underline_spacing)

    def add_underline_to_text(self, page, text_actual_rect, text, font, fontsize, rotate, spacing):
        """
        Draws an underline using the ACTUAL rendered position of the text.
        This version is compatible with older PyMuPDF versions that lack the 'Rect.center' property.
        """
        non_empty_lines = [line for line in text.split('\n') if line.strip()]
        if not non_empty_lines:
            return
        last_line_width = font.text_length(non_empty_lines[-1], fontsize=fontsize)

        underline_y = text_actual_rect.y1 + spacing

        center_x = text_actual_rect.x0 + text_actual_rect.width / 2
        p1_x = center_x - last_line_width / 2
        p2_x = center_x + last_line_width / 2

        p1 = fitz.Point(p1_x, underline_y)
        p2 = fitz.Point(p2_x, underline_y)

        if rotate != 0:
            pivot = fitz.Point(text_actual_rect.x0 + text_actual_rect.width / 2, 
                               text_actual_rect.y0 + text_actual_rect.height / 2)
            mat = fitz.Matrix(1, 1).prerotate(rotate)
            p1 = (p1 - pivot) * mat + pivot
            p2 = (p2 - pivot) * mat + pivot

        page.draw_line(p1, p2, color=(1, 1, 1), width=max(0.7, fontsize * 0.05))

    def generate_all_certificates(self):
        if not all([self.template_path, self.csv_path, self.output_folder]):
            QMessageBox.warning(self, "Warning", "Please select a template, a CSV file, and an output folder."); return
        if not self.certificate_data:
            QMessageBox.warning(self, "Warning", "No data loaded from the CSV file."); return

        name_rect = fitz.Rect(self.name_x.value(), self.name_y.value(), self.page_width, self.page_height)
        name_size = self.name_size.value()
        name_rot = self.name_rot.value()
        
        ach_x = self.ach_x.value()
        ach_y = self.ach_y.value()
        ach_w = self.ach_w.value()
        ach_h = self.ach_h.value()
        ach_rect = fitz.Rect(ach_x, ach_y, ach_x + ach_w, ach_y + ach_h)
        ach_size = self.ach_size.value()
        ach_rot = self.ach_rot.value()
        
        underline_spacing = self.underline_spacing.value()

        count = 0
        try:
            for i, (name, achievement) in enumerate(self.certificate_data):
                doc = fitz.open(self.template_path)
                page = doc[0]

                page.insert_font(fontname=f"F0-{i}", fontbuffer=self.font_name_bold.buffer)
                page.insert_font(fontname=f"F1-{i}", fontbuffer=self.font_achievement_reg.buffer)

                self.insert_text_with_autoresize(page, name_rect, name, self.font_name_bold.buffer, f"F0-{i}", name_size, name_rot, underline=False)
                
                self.insert_text_with_autoresize(page, ach_rect, achievement, self.font_achievement_reg.buffer, f"F1-{i}", ach_size, ach_rot, underline=True, underline_spacing=underline_spacing)

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
        if self.doc_template:
            self.doc_template.close()
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PdfCertificateGenerator()
    window.show()
    sys.exit(app.exec())