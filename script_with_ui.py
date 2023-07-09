import os
import pdfplumber
import threading
from pathlib import Path
from pdf2docx import Converter
from docx import Document
import tkinter as tk
from tkinter import filedialog, StringVar
from tqdm import tqdm

stop_thread = False

def convert_pdfs_to_docx_and_text(input_dir, output_docx_dir, output_text_dir, status_var, root):
    global stop_thread
    input_path = Path(input_dir)
    output_docx_path = Path(output_docx_dir)
    output_text_path = Path(output_text_dir)

    pdf_files = list(input_path.glob("*.pdf"))
    total_files = len(pdf_files)

    for index, file in enumerate(pdf_files):
        if stop_thread:
            break
        pdf_file_path = str(file)
        docx_file_path = str(output_docx_path / (file.stem + ".docx"))
        text_file_path = str(output_text_path / (file.stem + ".docx"))

        # Convert PDF to Word document with formatting
        cv = Converter(pdf_file_path)
        cv.convert(docx_file_path, start=0, end=None)
        cv.close()

        # Extract text from PDF and save as a Word document
        with pdfplumber.open(pdf_file_path) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
            doc = Document()
            doc.add_paragraph(text)
            doc.save(text_file_path)

        status_var.set(f"Processed file {index + 1} of {total_files}: {file.name}")
        root.update_idletasks()
        
    status_var.set("Finished processing files.")

def browse_directory(title):
    directory = filedialog.askdirectory(title=title)
    return directory

def browse_input_dir():
    input_dir.set(browse_directory("Select input directory"))

def start_conversion():
    global stop_thread
    stop_thread = False
    if input_dir.get():
        status.set("Processing files...")
        output_docx_path = os.path.join(input_dir.get(), 'docx_output')
        output_text_path = os.path.join(input_dir.get(), 'text_output')
        
        if not os.path.exists(output_docx_path):
            os.makedirs(output_docx_path)
        if not os.path.exists(output_text_path):
            os.makedirs(output_text_path)
        
        conversion_thread = threading.Thread(target=lambda: convert_pdfs_to_docx_and_text(input_dir.get(), output_docx_path, output_text_path, status, root))
        conversion_thread.start()
    else:
        status.set("Invalid directory. Please select the input directory.")

def stop_conversion():
    global stop_thread
    stop_thread = True
    status.set("Stopped processing files.")

def main():
    global root
    global input_dir
    global status
    root = tk.Tk()
    root.title("PDF Converter")

    input_dir = StringVar()
    status = StringVar()

    tk.Label(root, text="Input directory:").grid(row=0, column=0, sticky="e")
    tk.Entry(root, textvariable=input_dir, width=40).grid(row=0, column=1)
    tk.Button(root, text="Browse", command=browse_input_dir).grid(row=0, column=2)

    tk.Button(root, text="Start Conversion", command=start_conversion).grid(row=1, column=1, pady=10)
    tk.Button(root, text="Stop Conversion", command=stop_conversion).grid(row=1, column=2)

    tk.Label(root, text="Status:").grid(row=2, column=0, sticky="e")
    tk.Label(root, textvariable=status).grid(row=2, column=1, sticky="w")

    tk.Label(root, text="Made by Rares Ficiu v1").grid(row=3, column=1, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
