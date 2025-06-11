import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import pandas as pd
import fitz  # PyMuPDF
from PIL import Image
import os
import re

class CertificateGenerator(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Bulk Certificate Generator")
        self.geometry("1300x850")
        ctk.set_appearance_mode("System")

        # --- Data ---
        self.template_path = ""
        self.csv_path = ""
        self.output_dir = ""
        self.names_list = []
        self.font_color = (0, 0, 0) # Default black

        # --- Layout ---
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Controls Frame (Left) ---
        self.controls_frame = ctk.CTkFrame(self, width=350, corner_radius=10)
        self.controls_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.controls_frame.grid_rowconfigure(12, weight=1)

        # --- Image Display Frame (Right) ---
        self.image_frame = ctk.CTkFrame(self, fg_color="gray20")
        self.image_frame.grid(row=0, column=1, padx=(0, 20), pady=20, sticky="nsew")
        self.image_label = ctk.CTkLabel(self.image_frame, text="Load a PDF Template to begin", text_color="gray70")
        self.image_label.pack(expand=True)
        
        self.setup_controls()
        
        # Bind window resize to update preview
        self.bind("<Configure>", self.update_preview)

    def setup_controls(self):
        """Creates all the widgets in the left-hand controls panel."""
        row = 0
        title_label = ctk.CTkLabel(self.controls_frame, text="Certificate Generator", font=ctk.CTkFont(size=24, weight="bold"))
        title_label.grid(row=row, column=0, columnspan=2, padx=20, pady=(20, 10))
        row += 1

        # --- File/Folder Selection ---
        self.template_button = ctk.CTkButton(self.controls_frame, text="1. Select PDF Template", command=self.select_template)
        self.template_button.grid(row=row, column=0, padx=20, pady=10, sticky="ew")
        self.template_label = ctk.CTkLabel(self.controls_frame, text="No file selected", anchor="w", text_color="gray")
        self.template_label.grid(row=row+1, column=0, padx=20, pady=(0,10), sticky="ew")
        row += 2
        
        self.csv_button = ctk.CTkButton(self.controls_frame, text="2. Select CSV File", command=self.select_csv)
        self.csv_button.grid(row=row, column=0, padx=20, pady=10, sticky="ew")
        self.csv_label = ctk.CTkLabel(self.controls_frame, text="No file selected", anchor="w", text_color="gray")
        self.csv_label.grid(row=row+1, column=0, padx=20, pady=(0,10), sticky="ew")
        row += 2

        self.column_label = ctk.CTkLabel(self.controls_frame, text="Column with Names:", anchor="w")
        self.column_label.grid(row=row, column=0, padx=20, pady=(10,0), sticky="ew")
        self.column_var = ctk.StringVar(value="Select a column")
        self.column_menu = ctk.CTkOptionMenu(self.controls_frame, variable=self.column_var, values=["Load a CSV first"], state="disabled")
        self.column_menu.grid(row=row+1, column=0, padx=20, pady=(0,10), sticky="ew")
        row += 2

        self.output_button = ctk.CTkButton(self.controls_frame, text="3. Select Output Folder", command=self.select_output_dir)
        self.output_button.grid(row=row, column=0, padx=20, pady=10, sticky="ew")
        self.output_label = ctk.CTkLabel(self.controls_frame, text="No folder selected", anchor="w", text_color="gray")
        self.output_label.grid(row=row+1, column=0, padx=20, pady=(0,20), sticky="ew")
        row += 2
        
        # --- Text Styling Controls ---
        style_label = ctk.CTkLabel(self.controls_frame, text="Text Styling & Position", font=ctk.CTkFont(size=16, weight="bold"))
        style_label.grid(row=row, column=0, padx=20, pady=(10,0), sticky="ew")
        row += 1

        self.x_slider, self.x_label = self.create_slider("X Position", row, 0, 1000)
        row += 1
        self.y_slider, self.y_label = self.create_slider("Y Position", row, 0, 1000)
        row += 1
        self.size_slider, self.size_label = self.create_slider("Font Size", row, 8, 150)
        row += 1
        
        self.color_button = ctk.CTkButton(self.controls_frame, text="Text Color", command=self.choose_color)
        self.color_button.grid(row=row, column=0, padx=20, pady=10, sticky="ew")
        row += 1
        
        # --- Generation ---
        self.generate_button = ctk.CTkButton(self.controls_frame, text="Generate Certificates", height=40, font=ctk.CTkFont(size=18), command=self.start_generation, state="disabled")
        self.generate_button.grid(row=row, column=0, padx=20, pady=20, sticky="ew")
        row += 1

        self.progressbar = ctk.CTkProgressBar(self.controls_frame)
        self.progressbar.set(0)
        self.progressbar.grid(row=row, column=0, padx=20, pady=(0, 5), sticky="ew")
        row += 1
        
        self.status_label = ctk.CTkLabel(self.controls_frame, text="", anchor="center")
        self.status_label.grid(row=row, column=0, padx=20, pady=(0, 20), sticky="ew")

    def create_slider(self, label_text, row, from_, to):
        label = ctk.CTkLabel(self.controls_frame, text=f"{label_text}: 0", anchor="w")
        label.grid(row=row, column=0, padx=20, pady=(10, 0), sticky="w")
        slider = ctk.CTkSlider(self.controls_frame, from_=from_, to=to, command=self.update_preview)
        slider.grid(row=row, column=0, padx=20, pady=(0, 10), sticky="we")
        return slider, label

    def select_template(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not path: return
        self.template_path = path
        self.template_label.configure(text=os.path.basename(path), text_color="white")
        try:
            doc = fitz.open(self.template_path)
            page = doc[0]
            self.x_slider.configure(to=page.rect.width)
            self.y_slider.configure(to=page.rect.height)
            self.x_slider.set(page.rect.width / 2)
            self.y_slider.set(page.rect.height / 2)
            self.size_slider.set(48)
            doc.close()
            self.update_preview()
            self.check_if_ready()
        except Exception as e:
            messagebox.showerror("Error", f"Could not open PDF: {e}")

    def select_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not path: return
        self.csv_path = path
        self.csv_label.configure(text=os.path.basename(path), text_color="white")
        try:
            df = pd.read_csv(self.csv_path)
            columns = df.columns.tolist()
            self.column_menu.configure(values=columns, state="normal")
            self.column_var.set(columns[0])
            self.check_if_ready()
        except Exception as e:
            messagebox.showerror("Error", f"Could not read CSV: {e}")

    def select_output_dir(self):
        path = filedialog.askdirectory()
        if not path: return
        self.output_dir = path
        self.output_label.configure(text=os.path.basename(path), text_color="white")
        self.check_if_ready()

    def choose_color(self):
        color_code = colorchooser.askcolor(title="Choose text color")
        if color_code and color_code[0]:
            # Convert (r, g, b) from 0-255 to 0-1 for PyMuPDF
            self.font_color = tuple(c/255.0 for c in color_code[0])
            self.update_preview()

    def update_preview(self, event=None):
        if not self.template_path: return

        x = self.x_slider.get()
        y = self.y_slider.get()
        size = self.size_slider.get()
        
        self.x_label.configure(text=f"X Position: {int(x)}")
        self.y_label.configure(text=f"Y Position: {int(y)}")
        self.size_label.configure(text=f"Font Size: {int(size)}")

        # Create the preview image
        try:
            doc = fitz.open(self.template_path)
            page = doc[0]
            
            # Draw text for preview
            page.insert_text((x, y), "Sample Name", fontsize=size, fontname="helv-bold", color=self.font_color, align=fitz.TEXT_ALIGN_CENTER)
            
            pix = page.get_pixmap(dpi=150)
            doc.close()

            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Resize to fit the frame
            frame_w = self.image_frame.winfo_width()
            frame_h = self.image_frame.winfo_height()
            if frame_w > 1 and frame_h > 1:
                img.thumbnail((frame_w - 20, frame_h - 20), Image.LANCZOS)

            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
            self.image_label.configure(image=ctk_img, text="")
            self.image_label.image = ctk_img
        except Exception as e:
            print(f"Error updating preview: {e}")

    def check_if_ready(self):
        if self.template_path and self.csv_path and self.output_dir and self.column_var.get() != "Select a column":
            self.generate_button.configure(state="normal")
        else:
            self.generate_button.configure(state="disabled")

    def start_generation(self):
        # --- Validation ---
        if not all([self.template_path, self.csv_path, self.output_dir]):
            messagebox.showerror("Error", "Please select template, CSV, and output folder.")
            return
        
        name_column = self.column_var.get()
        if name_column == "Select a column" or name_column == "Load a CSV first":
            messagebox.showerror("Error", "Please select the column containing the names.")
            return

        try:
            df = pd.read_csv(self.csv_path)
            if name_column not in df.columns:
                messagebox.showerror("Error", f"Column '{name_column}' not found in CSV.")
                return
            self.names_list = df[name_column].dropna().tolist()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process CSV: {e}")
            return
            
        self.generate_button.configure(state="disabled")
        self.status_label.configure(text="Starting...")
        self.progressbar.set(0)
        
        total = len(self.names_list)
        for i, name in enumerate(self.names_list):
            try:
                # Sanitize name for filename
                safe_name = re.sub(r'[\\/*?:"<>|]', "", str(name))
                safe_name = safe_name.replace(" ", "_")
                output_filename = os.path.join(self.output_dir, f"certificate_{safe_name}.pdf")

                doc = fitz.open(self.template_path)
                page = doc[0]

                x = self.x_slider.get()
                y = self.y_slider.get()
                size = self.size_slider.get()

                page.insert_text((x, y), str(name), fontsize=size, fontname="helv-bold", color=self.font_color, align=fitz.TEXT_ALIGN_CENTER)

                doc.save(output_filename, garbage=4, deflate=True)
                doc.close()

                # Update GUI
                progress = (i + 1) / total
                self.progressbar.set(progress)
                self.status_label.configure(text=f"Generated: {name} ({i+1}/{total})")
                self.update_idletasks() # Force GUI to refresh

            except Exception as e:
                print(f"Error processing '{name}': {e}")
                self.status_label.configure(text=f"Error on: {name}")

        self.status_label.configure(text=f"Done! {total} certificates generated.")
        messagebox.showinfo("Success", f"Process complete!\n{total} certificates have been generated in the output folder.")
        self.generate_button.configure(state="normal")

if __name__ == "__main__":
    app = CertificateGenerator()
    app.mainloop()