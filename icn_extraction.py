import os
import zipfile
import re
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import queue
import sys

# ==============================================================================
# 1. CORE EXTRACTION LOGIC (This is the engine from our previous script)
#    No changes are needed here.
# ==============================================================================

def extract_images_with_tagged_icn(docx_path, output_dir):
    with zipfile.ZipFile(docx_path, 'r') as docx:
        media_files = sorted([f for f in docx.namelist() if f.startswith('word/media/')])
        if not media_files:
            print(f"⏭️ Skipping {os.path.basename(docx_path)}: No images found.")
            return False
        try:
            xml_content = docx.read("word/document.xml")
        except KeyError:
            print(f"❌ {os.path.basename(docx_path)} is missing document.xml.")
            return False

        plain_text = ""
        try:
            root = ET.fromstring(xml_content)
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            for t in root.findall('.//w:t', ns):
                if t.text:
                    plain_text += t.text
        except ET.ParseError:
            print(f"⚠️ Warning for {os.path.basename(docx_path)}: Could not parse XML, falling back to simple text search.")
            plain_text = xml_content.decode('utf-8', errors='ignore')

        icn_matches = re.findall(r'ICN-\s*([\w\-.]+)', plain_text)
        icn_labels = [f"ICN-{match}" for match in icn_matches]

        if len(icn_labels) != len(media_files):
            print(f"⚠️ Warning for {os.path.basename(docx_path)}: Found {len(media_files)} images but {len(icn_labels)} ICN tags. Using default names to avoid errors.")
            icn_labels = []

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        for i, media_file in enumerate(media_files):
            if i < len(icn_labels):
                label = icn_labels[i]
            else:
                label = f"image_{i + 1}"
            
            ext = os.path.splitext(media_file)[1]
            safe_label = re.sub(r'[<>:"/\\|?*]', '_', label)
            out_path = os.path.join(output_dir, f"{safe_label}{ext}")
            
            image_data = docx.read(media_file)
            with open(out_path, "wb") as out_file:
                out_file.write(image_data)
            
            print(f"✅ Saved: {os.path.basename(out_path)}")
        
        return True

def batch_process_folder(folder_path, output_root):
    files = [f for f in os.listdir(folder_path) if f.lower().endswith('.docx') and not f.startswith('~')]
    if not files:
        print(f"No .docx files found in '{folder_path}'")
        return

    for file in files:
        docx_path = os.path.join(folder_path, file)
        base_name = os.path.splitext(file)[0]
        output_dir = os.path.join(output_root, base_name)
        print(f"\n📂 Processing: {file}")
        
        created = extract_images_with_tagged_icn(docx_path, output_dir)

        if not created and os.path.exists(output_dir) and not os.listdir(output_dir):
            try:
                os.rmdir(output_dir)
            except OSError as e:
                print(f"Could not remove empty directory {output_dir}: {e}")

    print("\n🎉 Done: All files processed.")

# ==============================================================================
# 2. TKINTER GUI APPLICATION
# ==============================================================================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("DOCX Image Extractor")
        self.root.geometry("700x550")

        self.input_folder_path = tk.StringVar()
        self.output_folder_path = tk.StringVar()
        
        # --- UI Frames ---
        top_frame = tk.Frame(root, padx=10, pady=10)
        top_frame.pack(fill=tk.X)
        log_frame = tk.Frame(root, padx=10, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        # --- Input Folder Selection ---
        tk.Label(top_frame, text="1. Select Folder with DOCX Files:", anchor="w").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,2))
        self.input_entry = tk.Entry(top_frame, textvariable=self.input_folder_path, state="readonly", width=80)
        self.input_entry.grid(row=1, column=0, sticky="ew", padx=(0,5))
        self.input_btn = tk.Button(top_frame, text="Browse...", command=self.browse_input)
        self.input_btn.grid(row=1, column=1, sticky="ew")

        # --- Output Folder Selection ---
        tk.Label(top_frame, text="2. Select Parent Output Folder:", anchor="w").grid(row=2, column=0, columnspan=2, sticky="w", pady=(10,2))
        self.output_entry = tk.Entry(top_frame, textvariable=self.output_folder_path, state="readonly", width=80)
        self.output_entry.grid(row=3, column=0, sticky="ew", padx=(0,5))
        self.output_btn = tk.Button(top_frame, text="Browse...", command=self.browse_output)
        self.output_btn.grid(row=3, column=1, sticky="ew")
        
        top_frame.grid_columnconfigure(0, weight=1)

        # --- Action Button ---
        self.run_btn = tk.Button(top_frame, text="Start Extraction", command=self.start_extraction_thread, font=("Helvetica", 10, "bold"), bg="#4CAF50", fg="white")
        self.run_btn.grid(row=4, column=0, columnspan=2, pady=(15,0), sticky="ew")

        # --- Log Window ---
        tk.Label(log_frame, text="Status Log:", anchor="w").pack(fill=tk.X)
        self.log_widget = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state="disabled")
        self.log_widget.pack(fill=tk.BOTH, expand=True)
        
        self.log_queue = queue.Queue()
        self.root.after(100, self.process_log_queue)

    def browse_input(self):
        path = filedialog.askdirectory(title="Select Folder with DOCX Files")
        if path:
            self.input_folder_path.set(path)

    def browse_output(self):
        path = filedialog.askdirectory(title="Select Parent Folder to Save Output")
        if path:
            self.output_folder_path.set(path)

    def set_ui_state(self, state):
        """Enable or disable buttons."""
        if state == "disabled":
            self.run_btn.config(state=tk.DISABLED, text="Processing...")
            self.input_btn.config(state=tk.DISABLED)
            self.output_btn.config(state=tk.DISABLED)
        else: # 'normal'
            self.run_btn.config(state=tk.NORMAL, text="Start Extraction")
            self.input_btn.config(state=tk.NORMAL)
            self.output_btn.config(state=tk.NORMAL)

    def start_extraction_thread(self):
        input_folder = self.input_folder_path.get()
        parent_output_folder = self.output_folder_path.get()

        if not input_folder or not parent_output_folder:
            messagebox.showerror("Error", "Please select both an input and an output folder.")
            return

        self.set_ui_state("disabled")
        self.log_widget.config(state="normal")
        self.log_widget.delete('1.0', tk.END)
        self.log_widget.config(state="disabled")
        
        # Run the heavy work in a separate thread to keep the GUI responsive
        thread = threading.Thread(target=self.run_extraction, args=(input_folder, parent_output_folder), daemon=True)
        thread.start()

    def run_extraction(self, input_folder, parent_output_folder):
        """This function runs in the background thread."""
        
        # --- New output folder logic ---
        input_folder_name = os.path.basename(input_folder)
        final_output_root = os.path.join(parent_output_folder, input_folder_name)
        
        # Redirect print statements to our log queue
        sys.stdout = QueueWriter(self.log_queue)
        
        try:
            batch_process_folder(input_folder, final_output_root)
        except Exception as e:
            print(f"\n❌ AN UNEXPECTED ERROR OCCURRED:\n{e}")
        finally:
            # When done, restore stdout and send a "DONE" signal
            sys.stdout = sys.__stdout__
            self.log_queue.put("DONE")

    def process_log_queue(self):
        """Checks the queue for new log messages and updates the widget."""
        try:
            while True:
                line = self.log_queue.get_nowait()
                if line == "DONE":
                    self.set_ui_state("normal")
                    return
                else:
                    self.log_widget.config(state="normal")
                    self.log_widget.insert(tk.END, line)
                    self.log_widget.see(tk.END) # Auto-scroll
                    self.log_widget.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self.process_log_queue)

# Helper class to redirect stdout to the queue
class QueueWriter:
    def __init__(self, q):
        self.queue = q
    def write(self, text):
        self.queue.put(text)
    def flush(self):
        pass

# ==============================================================================
# 3. MAIN EXECUTION BLOCK
# ==============================================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()