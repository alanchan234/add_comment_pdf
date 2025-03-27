import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import shutil
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO

class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Processor")
        self.root.geometry("600x400")

        # Variables
        self.source_folder = ""
        self.dest_folder = ""
        self.df = None

        # UI Elements
        self.create_widgets()

    def create_widgets(self):
        # Source Folder Selection
        tk.Label(self.root, text="Source PDF Folder:").pack(pady=5)
        self.source_btn = tk.Button(self.root, text="Select Folder", command=self.select_source_folder)
        self.source_btn.pack()
        self.source_label = tk.Label(self.root, text="No folder selected")
        self.source_label.pack()

        # Submit Button
        self.submit_btn = tk.Button(self.root, text="Generate Excel", command=self.generate_excel, state="disabled")
        self.submit_btn.pack(pady=10)

        # Destination Folder Selection
        tk.Label(self.root, text="Destination Folder for Processed PDFs:").pack(pady=5)
        self.dest_btn = tk.Button(self.root, text="Select Destination Folder", command=self.select_dest_folder, state="disabled")
        self.dest_btn.pack()
        self.dest_label = tk.Label(self.root, text="No destination folder selected")
        self.dest_label.pack()

        # Progress Bar
        self.progress = ttk.Progressbar(self.root, length=300, mode="determinate")
        self.progress.pack(pady=10)
        self.progress_label = tk.Label(self.root, text="")
        self.progress_label.pack()

        # Upload Excel Button
        self.upload_btn = tk.Button(self.root, text="Upload Excel and Process", command=self.upload_and_process, state="disabled")
        self.upload_btn.pack(pady=10)

    def select_source_folder(self):
        self.source_folder = filedialog.askdirectory()
        if self.source_folder:
            self.source_label.config(text=self.source_folder)
            self.submit_btn.config(state="normal")

    def generate_excel(self):
        # Debug: Print all files in the folder
        all_files = os.listdir(self.source_folder)
        print(f"Files in folder {self.source_folder}: {all_files}")

        pdf_files = [f for f in os.listdir(self.source_folder) if f.lower().endswith('.pdf')]
        if not pdf_files:
            tk.messagebox.showerror("Error", "No PDF files found in the selected folder")
            return

        self.df = pd.DataFrame({
            'File Name': pdf_files,
            'Invoice Num': [os.path.splitext(f)[0] for f in pdf_files],
            'Voucher Num': [''] * len(pdf_files)
        })

        excel_path = os.path.join(self.source_folder, 'pdf_list.xlsx')
        self.df.to_excel(excel_path, index=False)
        tk.messagebox.showinfo("Success", f"Excel file generated at: {excel_path}")
        # Enable the destination folder selection after generating the Excel file
        self.dest_btn.config(state="normal")

    def select_dest_folder(self):
        self.dest_folder = filedialog.askdirectory()
        if self.dest_folder:
            self.dest_label.config(text=self.dest_folder)
            # Enable the upload button only after the destination folder is selected
            self.upload_btn.config(state="normal")
        else:
            self.upload_btn.config(state="disabled")

    def upload_and_process(self):
        excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not excel_file:
            return

        self.df = pd.read_excel(excel_file)
        total_files = len(self.df)
        self.progress["maximum"] = total_files

        for index, row in self.df.iterrows():
            original_file = os.path.join(self.source_folder, row['File Name'])
            invoice_num = row['Invoice Num']
            voucher_num = str(row['Voucher Num'])  # Convert to string to avoid type issues

            if os.path.exists(original_file):
                # Read original PDF
                reader = PdfReader(original_file)
                writer = PdfWriter()

                # Get the first page to determine its size
                first_page = reader.pages[0]
                media_box = first_page.mediabox  # Get the page dimensions
                page_width = float(media_box.width)
                page_height = float(media_box.height)
                print(f"Processing {row['File Name']}: Page size = {page_width}x{page_height} points")

                # Create a temporary PDF with the voucher number annotation
                buffer = BytesIO()
                c = canvas.Canvas(buffer, pagesize=(page_width, page_height))  # Match the page size
                c.setFont("Helvetica-Bold", 12)
                # Add the voucher number in the top-left corner
                # Position: 50 points from left, 50 points from top
                x_position = 50
                y_position = page_height - 50  # Top of the page
                if voucher_num and voucher_num.strip():  # Check if voucher_num is not empty
                    # Draw the text
                    text = f"Voucher Num: {voucher_num}"
                    c.setFillColorRGB(0, 0, 0)  # Black text
                    c.drawString(x_position, y_position, text)
                    
                    # Calculate the text width to size the rectangle
                    text_width = c.stringWidth(text, "Helvetica-Bold", 12)
                    # Draw a black border around the text
                    c.setStrokeColorRGB(0, 0, 0)  # Black border
                    c.rect(x_position - 5, y_position - 5, text_width + 10, 20, fill=0)  # Border only, no fill
                    print(f"Added annotation '{text}' at ({x_position}, {y_position})")
                else:
                    print(f"No voucher number provided for {row['File Name']}")
                c.showPage()
                c.save()

                # Merge the annotation with the original PDF
                buffer.seek(0)
                annotation_pdf = PdfReader(buffer)
                annotation_page = annotation_pdf.pages[0]

                # Merge the annotation onto the first page
                first_page.merge_page(annotation_page)

                # Add the modified first page and remaining pages to the writer
                writer.add_page(first_page)
                for page_num in range(1, len(reader.pages)):
                    writer.add_page(reader.pages[page_num])

                # Save with new name directly to the destination folder
                new_filename = f"{invoice_num}.pdf"
                dest_path = os.path.join(self.dest_folder, new_filename)
                with open(dest_path, 'wb') as output_file:
                    writer.write(output_file)

                # Update progress
                self.progress["value"] = index + 1
                self.progress_label.config(text=f"Processing: {index + 1}/{total_files}")
                self.root.update()

        tk.messagebox.showinfo("Processing Complete", f"PDF files have been renamed, annotated, and moved to: {self.dest_folder}")
        # Reset UI
        self.reset_ui()

    def reset_ui(self):
        self.source_folder = ""
        self.dest_folder = ""
        self.source_label.config(text="No folder selected")
        self.dest_label.config(text="No destination folder selected")
        self.submit_btn.config(state="disabled")
        self.upload_btn.config(state="disabled")
        self.dest_btn.config(state="disabled")
        self.progress["value"] = 0
        self.progress_label.config(text="")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()