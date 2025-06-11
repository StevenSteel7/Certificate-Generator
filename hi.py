import customtkinter as ctk
import fitz  # PyMuPDF
from tkinter import filedialog, messagebox
from PIL import Image

class PdfTextEditor(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Live PDF Text Positioner")
        self.geometry("1200x800")
        ctk.set_appearance_mode("System")

        # --- Data Attributes ---
        self.doc = None
        self.original_page_pix = None
        self.page_width = 0
        self.page_height = 0
        self.input_pdf_path = ""

        # --- Layout ---
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Controls Frame (Left) ---
        self.controls_frame = ctk.CTkFrame(self, width=300, corner_radius=0)
        self.controls_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.controls_frame.grid_rowconfigure(8, weight=1)

        # --- Widgets in Controls Frame ---
        self.title_label = ctk.CTkLabel(self.controls_frame, text="Controls", font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.load_button = ctk.CTkButton(self.controls_frame, text="Load PDF", command=self.load_pdf)
        self.load_button.grid(row=1, column=0, padx=20, pady=10)

        self.text_entry_label = ctk.CTkLabel(self.controls_frame, text="Text to Add:", anchor="w")
        self.text_entry_label.grid(row=2, column=0, padx=20, pady=(10, 0))
        self.text_entry = ctk.CTkEntry(self.controls_frame, placeholder_text="Enter your text here")
        self.text_entry.grid(row=3, column=0, padx=20, pady=(0, 10))
        self.text_entry.insert(0, "Steven Moses")
        self.text_entry.bind("<KeyRelease>", self.update_display)

        self.x_slider, self.x_label = self.create_slider("X Position", 4, 0, 1000)
        self.y_slider, self.y_label = self.create_slider("Y Position", 5, 0, 1000)
        self.size_slider, self.size_label = self.create_slider("Font Size", 6, 8, 100)
        self.rot_slider, self.rot_label = self.create_slider("Rotation", 7, -180, 180)
        
        # Set initial values for sliders
        self.size_slider.set(36)

        self.save_button = ctk.CTkButton(self.controls_frame, text="Save to New PDF", command=self.save_pdf)
        self.save_button.grid(row=9, column=0, padx=20, pady=20)

        # --- Image Display Frame (Right) ---
        self.image_frame = ctk.CTkFrame(self, fg_color="gray20")
        self.image_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.image_label = ctk.CTkLabel(self.image_frame, text="Load a PDF to begin")
        self.image_label.pack(expand=True)
        
        # Bind the window resize event
        self.bind("<Configure>", self.update_display)


    def create_slider(self, label_text, row, from_, to):
        """Helper to create a label and a slider."""
        label = ctk.CTkLabel(self.controls_frame, text=f"{label_text}: 0", anchor="w")
        label.grid(row=row, column=0, padx=20, pady=(10, 0), sticky="w")
        slider = ctk.CTkSlider(self.controls_frame, from_=from_, to=to, command=self.update_display)
        slider.grid(row=row, column=0, padx=20, pady=(0, 10), sticky="we")
        return slider, label

    def load_pdf(self):
        """Opens a file dialog to load a PDF and displays the first page."""
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        
        self.input_pdf_path = path
        self.doc = fitz.open(self.input_pdf_path)
        
        if not self.doc.page_count:
            messagebox.showerror("Error", "This PDF has no pages.")
            return

        page = self.doc[0]
        self.page_width = int(page.rect.width)
        self.page_height = int(page.rect.height)

        # Update slider ranges to match page dimensions
        self.x_slider.configure(to=self.page_width)
        self.y_slider.configure(to=self.page_height)
        self.x_slider.set(self.page_width / 2)
        self.y_slider.set(self.page_height / 2)
        
        self.original_page_pix = page.get_pixmap()
        self.update_display()

    def update_display(self, event=None):
        """The core function: redraws the text on the page image based on slider values."""
        if self.original_page_pix is None:
            return

        # Get current values
        x = int(self.x_slider.get())
        y = int(self.y_slider.get())
        size = int(self.size_slider.get())
        rotation = int(self.rot_slider.get())
        text = self.text_entry.get()

        # Update info labels
        self.x_label.configure(text=f"X Position: {x}")
        self.y_label.configure(text=f"Y Position: {y}")
        self.size_label.configure(text=f"Font Size: {size}")
        self.rot_label.configure(text=f"Rotation: {rotation}")

        # --- Create the composite image ---
        # 1. Start with the clean background page image
        img = Image.frombytes("RGB", [self.original_page_pix.width, self.original_page_pix.height], self.original_page_pix.samples)

        # 2. Create a temporary, transparent pixmap for the text
        temp_doc = fitz.open()
        temp_page = temp_doc.new_page(width=self.page_width, height=self.page_height)
        
        # Use a rectangle for better alignment and rotation control
        rect = fitz.Rect(0, 0, self.page_width, self.page_height)
        
        temp_page.insert_text(
            (x, y),
            text,
            fontsize=size,
            fontname="helv",
            color=(0.9, 0.1, 0.1), # Bright red for visibility
            rotate=rotation,
        )

        # Render the text layer with a transparent background
        text_pix = temp_page.get_pixmap(alpha=True)
        text_img = Image.frombytes("RGBA", [text_pix.width, text_pix.height], text_pix.samples)

        # 3. Paste the text layer onto the background image
        img.paste(text_img, (0, 0), text_img)
        
        # --- Display the image in the GUI ---
        # Resize image to fit the display frame while maintaining aspect ratio
        frame_w = self.image_frame.winfo_width()
        frame_h = self.image_frame.winfo_height()
        
        img.thumbnail((frame_w - 20, frame_h - 20), Image.LANCZOS)
        
        ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
        
        self.image_label.configure(image=ctk_img, text="")
        self.image_label.image = ctk_img # Keep a reference

    def save_pdf(self):
        """Saves the final text to a new PDF file."""
        if not self.doc:
            messagebox.showwarning("Warning", "Please load a PDF first.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save PDF As..."
        )
        if not save_path:
            return

        # Re-open the original document to ensure it's clean
        final_doc = fitz.open(self.input_pdf_path)
        page = final_doc[0]

        # Get final values from the GUI
        x = int(self.x_slider.get())
        y = int(self.y_slider.get())
        size = int(self.size_slider.get())
        rotation = int(self.rot_slider.get())
        text = self.text_entry.get()

        # Insert the text using the final values
        page.insert_text(
            (x, y),
            text,
            fontsize=size,
            fontname="helv-bold",
            color=(0, 0, 0), # Save it in black
            rotate=rotation
        )
        
        try:
            final_doc.save(save_path, garbage=4, deflate=True)
            messagebox.showinfo("Success", f"PDF saved successfully to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save PDF: {e}")
        finally:
            final_doc.close()


if __name__ == "__main__":
    app = PdfTextEditor()
    app.mainloop()