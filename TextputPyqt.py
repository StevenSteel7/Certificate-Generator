import sys
import fitz  # PyMuPDF
from PIL import Image

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QPushButton, QLabel, QLineEdit, QSlider, QFileDialog, QMessageBox
)
from PyQt6.QtGui import QPixmap, QImage, QFont
from PyQt6.QtCore import Qt, QSize

class PdfTextEditor(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Live PDF Text Positioner (PyQt6)")
        self.setGeometry(100, 100, 1200, 800)

        # --- Data Attributes ---
        self.doc = None
        self.original_page_pixmap = None  # This will hold the fitz.Pixmap
        self.page_width = 0
        self.page_height = 0
        self.input_pdf_path = ""

        # --- Main Layout ---
        # Central widget to hold everything
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        # Main horizontal layout: Controls on left, Image on right
        main_layout = QHBoxLayout(central_widget)

        # --- Controls Frame (Left) ---
        controls_widget = QWidget()
        controls_widget.setFixedWidth(300)
        controls_layout = QVBoxLayout(controls_widget)
        main_layout.addWidget(controls_widget)

        # --- Widgets in Controls Frame ---
        title_label = QLabel("Controls")
        title_label.setFont(QFont("Arial", 20, QFont.Weight.Bold))
        controls_layout.addWidget(title_label)

        self.load_button = QPushButton("Load PDF")
        self.load_button.clicked.connect(self.load_pdf)
        controls_layout.addWidget(self.load_button)

        # Use QFormLayout for neatly aligned Label-Widget pairs
        form_layout = QFormLayout()
        
        self.text_entry = QLineEdit("Steven Moses")
        self.text_entry.setPlaceholderText("Enter your text here")
        self.text_entry.textChanged.connect(self.update_display)
        form_layout.addRow("Text to Add:", self.text_entry)

        # Sliders
        self.x_slider, self.x_label = self.create_slider(0, 1000)
        self.y_slider, self.y_label = self.create_slider(0, 1000)
        self.size_slider, self.size_label = self.create_slider(8, 100)
        self.rot_slider, self.rot_label = self.create_slider(-180, 180)
        
        form_layout.addRow(self.x_label, self.x_slider)
        form_layout.addRow(self.y_label, self.y_slider)
        form_layout.addRow(self.size_label, self.size_slider)
        form_layout.addRow(self.rot_label, self.rot_slider)
        
        # Set initial values
        self.size_slider.setValue(36)
        
        controls_layout.addLayout(form_layout)
        controls_layout.addStretch() # Pushes the save button to the bottom

        self.save_button = QPushButton("Save to New PDF")
        self.save_button.clicked.connect(self.save_pdf)
        controls_layout.addWidget(self.save_button)

        # --- Image Display Label (Right) ---
        self.image_label = QLabel("Load a PDF to begin")
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.setStyleSheet("background-color: #333333; color: white;")
        main_layout.addWidget(self.image_label, 1) # The '1' makes it take up remaining space


    def create_slider(self, min_val, max_val):
        """Helper to create a QSlider and its corresponding QLabel."""
        label = QLabel(f"Value: {min_val}")
        slider = QSlider(Qt.Orientation.Horizontal)
        slider.setRange(min_val, max_val)
        slider.valueChanged.connect(self.update_display)
        return slider, label

    def load_pdf(self):
        """Opens a file dialog to load a PDF and displays the first page."""
        path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf)")
        if not path:
            return
            
        self.input_pdf_path = path
        try:
            self.doc = fitz.open(self.input_pdf_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open PDF file: {e}")
            return
        
        if not self.doc.page_count:
            QMessageBox.critical(self, "Error", "This PDF has no pages.")
            return

        page = self.doc[0]
        self.page_width = int(page.rect.width)
        self.page_height = int(page.rect.height)

        # Update slider ranges to match page dimensions
        self.x_slider.setRange(0, self.page_width)
        self.y_slider.setRange(0, self.page_height)
        self.x_slider.setValue(self.page_width // 2)
        self.y_slider.setValue(self.page_height // 2)
        
        # Store the original page render from PyMuPDF
        self.original_page_pixmap = page.get_pixmap()
        self.update_display()

    def update_display(self, _=None): # The _ is to catch the signal's value argument
        """The core function: redraws the text on the page image based on slider values."""
        if self.original_page_pixmap is None:
            return

        # Get current values
        x = self.x_slider.value()
        y = self.y_slider.value()
        size = self.size_slider.value()
        rotation = self.rot_slider.value()
        text = self.text_entry.text()

        # Update info labels
        self.x_label.setText(f"X Position: {x}")
        self.y_label.setText(f"Y Position: {y}")
        self.size_label.setText(f"Font Size: {size}")
        self.rot_label.setText(f"Rotation: {rotation}")

        # --- Create the composite image ---
        # 1. Start with the clean background page image
        img = Image.frombytes("RGB", [self.original_page_pixmap.width, self.original_page_pixmap.height], self.original_page_pixmap.samples)

        # 2. Create a temporary, transparent pixmap for the text using fitz
        temp_doc = fitz.open()
        temp_page = temp_doc.new_page(width=self.page_width, height=self.page_height)
        
        temp_page.insert_text(
            (x, y), text, fontsize=size, fontname="helv",
            color=(1, 1, 1), # Bright red for visibility
            rotate=rotation,
        )

        # Render the text layer with a transparent background
        text_pix = temp_page.get_pixmap(alpha=True)
        text_img = Image.frombytes("RGBA", [text_pix.width, text_pix.height], text_pix.samples)

        # 3. Paste the text layer onto the background image using PIL
        img.paste(text_img, (0, 0), text_img)
        
        # --- Display the image in the GUI ---
        # Convert PIL Image to QPixmap for display in a QLabel
        q_image = QImage(img.tobytes("raw", "RGB"), img.width, img.height, QImage.Format.Format_RGB888)
        pixmap = QPixmap.fromImage(q_image)

        # Scale pixmap to fit the label while maintaining aspect ratio
        scaled_pixmap = pixmap.scaled(self.image_label.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.image_label.setPixmap(scaled_pixmap)

    def save_pdf(self):
        """Saves the final text to a new PDF file."""
        if not self.doc:
            QMessageBox.warning(self, "Warning", "Please load a PDF first.")
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF As...", "", "PDF Files (*.pdf)")
        if not save_path:
            return

        # Re-open the original document to ensure it's clean
        final_doc = fitz.open(self.input_pdf_path)
        page = final_doc[0]

        # Get final values from the GUI
        x = self.x_slider.value()
        y = self.y_slider.value()
        size = self.size_slider.value()
        rotation = self.rot_slider.value()
        text = self.text_entry.text()

        # Insert the text using the final values
        # FIX: Changed "helv-bold" to "helv" to use a standard PDF base font
        # that does not require an external font file for embedding.
        page.insert_text(
            (x, y),
            text,
            fontsize=size,
            fontname="helv", 
            color=(1, 1, 1), # Save it in black
            rotate=rotation
        )
        
        try:
            final_doc.save(save_path, garbage=4, deflate=True)
            QMessageBox.information(self, "Success", f"PDF saved successfully to:\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save PDF: {e}")
        finally:
            final_doc.close()
    
    def resizeEvent(self, event):
        """This method is called automatically when the window is resized."""
        super().resizeEvent(event)
        # We call update_display to rescale the pixmap to the new label size
        self.update_display()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PdfTextEditor()
    window.show()
    sys.exit(app.exec())