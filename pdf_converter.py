import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdf2docx import Converter
import threading

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to DOCX Converter")
        self.root.geometry("500x650")  
        
        # Variables
        self.input_path = None
        self.output_dir = None
        self.conversion_type = tk.StringVar(value="file")
        self.output_location = tk.StringVar(value="same")
        self.progress_var = tk.DoubleVar(value=0)
        self.status_var = tk.StringVar(value="Ready")
        
        # Simple grid layout - 12 rows x 2 columns (added rows for progress)
        # ROW 0: Title
        title_label = tk.Label(root, text="PDF to DOCX Converter", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # ROW 1-3: Conversion options
        options_frame = tk.LabelFrame(root, text="Conversion Options")
        options_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        
        tk.Radiobutton(options_frame, text="Convert One File", variable=self.conversion_type, value="file").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        tk.Radiobutton(options_frame, text="Convert One Folder", variable=self.conversion_type, value="folder").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        tk.Radiobutton(options_frame, text="Convert Nested Folders", variable=self.conversion_type, value="nested").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        
        # ROW 4-5: Output location
        output_frame = tk.LabelFrame(root, text="Output Location")
        output_frame.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        
        tk.Radiobutton(output_frame, text="Save in the Same Folder", variable=self.output_location, value="same").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        tk.Radiobutton(output_frame, text="Save in a Different Folder", variable=self.output_location, value="different").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        
        # ROW 6: Selection buttons
        self.select_input_btn = tk.Button(root, text="Select Input", command=self.select_input)
        self.select_input_btn.grid(row=6, column=0, padx=5, pady=10, sticky="w")
        
        self.select_output_btn = tk.Button(root, text="Select Output Folder", command=self.select_output, state=tk.DISABLED)
        self.select_output_btn.grid(row=6, column=1, padx=5, pady=10, sticky="e")
        
        # ROW 7: Selected paths
        paths_frame = tk.Frame(root)
        paths_frame.grid(row=7, column=0, columnspan=2, padx=10, sticky="ew")
        
        tk.Label(paths_frame, text="Input:").grid(row=0, column=0, sticky="w")
        self.input_label = tk.Label(paths_frame, text="No input selected", wraplength=400)
        self.input_label.grid(row=0, column=1, sticky="w")
        
        tk.Label(paths_frame, text="Output:").grid(row=1, column=0, sticky="w")
        self.output_label = tk.Label(paths_frame, text="Same as input", wraplength=400)
        self.output_label.grid(row=1, column=1, sticky="w")
        
        # ROW 8-9: Progress section (new)
        progress_frame = tk.LabelFrame(root, text="Progress")
        progress_frame.grid(row=8, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100, 
            mode='determinate',
            length=450  # Set a good width
        )
        self.progress_bar.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        # Status label
        self.status_label = tk.Label(progress_frame, textvariable=self.status_var)
        self.status_label.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="w")
        
        # ROW 10: Main buttons - START and EXIT
        buttons_frame = tk.Frame(root)
        buttons_frame.grid(row=10, column=0, columnspan=2, pady=20)
        
        self.exit_btn = tk.Button(buttons_frame, text="EXIT", command=root.quit, bg="red", fg="white", width=15, height=2)
        self.exit_btn.pack(side=tk.LEFT, padx=10)
        
        self.start_btn = tk.Button(buttons_frame, text="START", command=self.start_conversion, bg="green", fg="white", width=15, height=2, state=tk.DISABLED)
        self.start_btn.pack(side=tk.LEFT, padx=10)
        
        # ROW 11: Footer
        footer_label = tk.Label(root, text="Frankie Sanzki, 2025", bg="lightgray", width=60)
        footer_label.grid(row=11, column=0, columnspan=2, sticky="ew", pady=(20, 0))
        
        # Configure grid to expand properly
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=1)
        
        # Configure progress frame column to expand
        progress_frame.grid_columnconfigure(0, weight=1)
        
        # Trace for output location changes
        self.output_location.trace_add("write", self.update_output_button_state)
    
    def update_output_button_state(self, *args):
        if self.output_location.get() == "different":
            self.select_output_btn.config(state=tk.NORMAL)
            self.output_label.config(text="No output folder selected")
        else:
            self.select_output_btn.config(state=tk.DISABLED)
            self.output_label.config(text="Same as input")
            self.output_dir = None
    
    def select_input(self):
        conversion_type = self.conversion_type.get()
        
        if conversion_type == "file":
            input_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if input_path:
                self.input_path = input_path
                self.input_label.config(text=input_path)
                self.start_btn.config(state=tk.NORMAL)
        else:  # folder or nested
            input_path = filedialog.askdirectory()
            if input_path:
                self.input_path = input_path
                self.input_label.config(text=input_path)
                self.start_btn.config(state=tk.NORMAL)
    
    def select_output(self):
        output_dir = filedialog.askdirectory()
        if output_dir:
            self.output_dir = output_dir
            self.output_label.config(text=output_dir)
    
    def start_conversion(self):
        if not self.input_path:
            messagebox.showerror("Error", "Please select an input file or folder first.")
            return
        
        if self.output_location.get() == "different" and not self.output_dir:
            messagebox.showerror("Error", "Please select an output folder.")
            return
        
        # Reset progress bar
        self.progress_var.set(0)
        self.status_var.set("Starting conversion...")
        
        # Start conversion in a separate thread
        self.start_btn.config(state=tk.DISABLED)
        self.select_input_btn.config(state=tk.DISABLED)
        self.select_output_btn.config(state=tk.DISABLED)
        
        conversion_thread = threading.Thread(
            target=self.perform_conversion,
            args=(self.conversion_type.get(), self.input_path, self.output_dir)
        )
        conversion_thread.daemon = True
        conversion_thread.start()
    
    def perform_conversion(self, conversion_type, input_path, output_dir):
        try:
            if conversion_type == "file":
                # For a single file, we update progress to 50% when starting conversion
                # and 100% when finished
                self.update_status("Converting file...", 10)
                self.convert_single_file(input_path, output_dir)
                self.update_status("Conversion completed successfully!", 100)
            
            elif conversion_type == "folder" or conversion_type == "nested":
                pdf_files = self.collect_pdf_files(input_path, nested=(conversion_type == "nested"))
                
                if not pdf_files:
                    raise ValueError("No PDF files found in the selected folder")
                
                total_files = len(pdf_files)
                self.update_status(f"Found {total_files} PDF files. Starting conversion...", 0)
                
                # Convert each PDF file and update progress
                for i, pdf_path in enumerate(pdf_files):
                    file_name = os.path.basename(pdf_path)
                    self.update_status(f"Converting {i+1}/{total_files}: {file_name}", (i / total_files) * 100)
                    
                    try:
                        if output_dir:
                            # Keep original folder structure for nested option
                            if conversion_type == "nested":
                                rel_path = os.path.relpath(pdf_path, input_path)
                                output_path = os.path.join(output_dir, rel_path.replace('.pdf', '.docx'))
                                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                            else:
                                output_filename = os.path.basename(pdf_path).replace('.pdf', '.docx')
                                output_path = os.path.join(output_dir, output_filename)
                        else:
                            output_path = pdf_path.replace('.pdf', '.docx')
                        
                        cv = Converter(pdf_path)
                        cv.convert(output_path)
                        cv.close()
                        
                    except Exception as e:
                        print(f"Error converting {pdf_path}: {str(e)}")
                
                # Final update
                self.update_status("All files converted successfully!", 100)
            
            messagebox.showinfo("Success", "Conversion completed successfully!")
        
        except Exception as e:
            self.update_status(f"Error: {str(e)}", 0)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        
        finally:
            self.start_btn.config(state=tk.NORMAL)
            self.select_input_btn.config(state=tk.NORMAL)
            if self.output_location.get() == "different":
                self.select_output_btn.config(state=tk.NORMAL)
    
    def update_status(self, message, progress):
        """Update status message and progress bar safely from a thread"""
        self.status_var.set(message)
        self.progress_var.set(progress)
        self.root.update_idletasks()  # Force update of the UI
    
    def collect_pdf_files(self, folder_path, nested=False):
        """Collect all PDF files in the given folder (and subfolders if nested=True)"""
        pdf_files = []
        
        if nested:
            for root, _, files in os.walk(folder_path):
                for file in files:
                    if file.lower().endswith('.pdf'):
                        pdf_files.append(os.path.join(root, file))
        else:
            for file in os.listdir(folder_path):
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(folder_path, file))
        
        return pdf_files
    
    def convert_single_file(self, pdf_path, output_dir=None):
        if not pdf_path.lower().endswith('.pdf'):
            raise ValueError("Selected file is not a PDF")
        
        # Determine output path
        if output_dir:
            output_filename = os.path.basename(pdf_path).replace('.pdf', '.docx')
            output_path = os.path.join(output_dir, output_filename)
        else:
            output_path = pdf_path.replace('.pdf', '.docx')
        
        # Convert PDF to DOCX
        cv = Converter(pdf_path)
        self.update_status("Converting file...", 50)
        cv.convert(output_path)
        cv.close()

# Create the main application file
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()