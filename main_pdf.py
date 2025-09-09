import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from pdf2docx import Converter

def convert_pdf_to_docx(pdf_path, docx_path):
    try:
        converter = Converter(pdf_path)
        converter.convert(docx_path)
        converter.close()
        return True
    except Exception as e:
        print(f"Conversion error: {e}")
        return False

def run_conversion_thread(pdf_file, docx_file):
    # Start the progress bar animation and disable button
    progress_bar.start()
    convert_button.config(state=tk.DISABLED)

    success = convert_pdf_to_docx(pdf_file, docx_file)

    # Stop progress bar and re-enable button
    progress_bar.stop()
    convert_button.config(state=tk.NORMAL)

    if success and os.path.isfile(docx_file):
        messagebox.showinfo("Success", f"PDF successfully converted to:\n{docx_file}")
    else:
        messagebox.showerror("Error", "Conversion failed.")

def browse_and_convert():
    pdf_file = filedialog.askopenfilename(
        title="Select PDF File",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not pdf_file:
        return

    docx_file = os.path.splitext(pdf_file)[0] + ".docx"

    # Run conversion in background thread to keep GUI responsive
    threading.Thread(target=run_conversion_thread, args=(pdf_file, docx_file), daemon=True).start()

# GUI setup
root = tk.Tk()
root.title("PDF to DOCX Converter")
root.geometry("300x180")

label = tk.Label(root, text="Click the button to convert PDF to DOCX")
label.pack(pady=10)

convert_button = tk.Button(root, text="Select PDF and Convert", command=browse_and_convert)
convert_button.pack(pady=10)

progress_bar = ttk.Progressbar(root, mode='indeterminate')
progress_bar.pack(pady=10, fill=tk.X, padx=20)

root.mainloop()
