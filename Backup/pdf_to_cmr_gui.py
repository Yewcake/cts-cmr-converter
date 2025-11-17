#!/usr/bin/env python3
"""
PDF to CMR Converter - GUI Application
Simple graphical interface for converting packing list PDFs to CMR Excel
"""

import os
import sys
import threading
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

# Import from the main script
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pdf_to_cmr import PackingListExtractor, CMRExcelPopulator


class PDFtoCMRApp:
    """GUI Application for PDF to CMR conversion"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("CTS Packing List to CMR Converter")
        self.root.geometry("700x550")
        self.root.resizable(False, False)
        
        # Variables
        self.pdf_path = StringVar()
        self.template_path = StringVar(value="CTS_NL_CMR_Template.xlsx")
        self.output_dir = StringVar(value="./cmr_output")
        self.status_text = StringVar(value="Ready")
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the user interface"""
        
        # Title
        title_frame = Frame(self.root, bg="#0066cc", height=60)
        title_frame.pack(fill=X)
        title_frame.pack_propagate(False)
        
        title_label = Label(
            title_frame,
            text="CTS Packing List to CMR Converter",
            font=("Arial", 16, "bold"),
            bg="#0066cc",
            fg="white"
        )
        title_label.pack(pady=15)
        
        # Main content frame
        content_frame = Frame(self.root, padx=20, pady=20)
        content_frame.pack(fill=BOTH, expand=True)
        
        # PDF Selection
        pdf_frame = LabelFrame(content_frame, text="1. Select Packing List PDF", font=("Arial", 10, "bold"))
        pdf_frame.pack(fill=X, pady=(0, 15))
        
        pdf_inner = Frame(pdf_frame, padx=10, pady=10)
        pdf_inner.pack(fill=X)
        
        Entry(pdf_inner, textvariable=self.pdf_path, width=60, state="readonly").pack(side=LEFT, padx=(0, 10))
        Button(pdf_inner, text="Browse...", command=self.browse_pdf, width=12).pack(side=LEFT)
        
        # Template Selection
        template_frame = LabelFrame(content_frame, text="2. CMR Template (Optional)", font=("Arial", 10, "bold"))
        template_frame.pack(fill=X, pady=(0, 15))
        
        template_inner = Frame(template_frame, padx=10, pady=10)
        template_inner.pack(fill=X)
        
        Entry(template_inner, textvariable=self.template_path, width=60).pack(side=LEFT, padx=(0, 10))
        Button(template_inner, text="Browse...", command=self.browse_template, width=12).pack(side=LEFT)
        
        # Output Directory
        output_frame = LabelFrame(content_frame, text="3. Output Directory", font=("Arial", 10, "bold"))
        output_frame.pack(fill=X, pady=(0, 15))
        
        output_inner = Frame(output_frame, padx=10, pady=10)
        output_inner.pack(fill=X)
        
        Entry(output_inner, textvariable=self.output_dir, width=60).pack(side=LEFT, padx=(0, 10))
        Button(output_inner, text="Browse...", command=self.browse_output, width=12).pack(side=LEFT)
        
        # Process Button
        button_frame = Frame(content_frame)
        button_frame.pack(pady=20)
        
        self.process_btn = Button(
            button_frame,
            text="Convert to CMR",
            command=self.process_pdf,
            font=("Arial", 12, "bold"),
            bg="#0066cc",
            fg="white",
            width=20,
            height=2,
            cursor="hand2"
        )
        self.process_btn.pack()
        
        # Batch Process Button
        self.batch_btn = Button(
            button_frame,
            text="Batch Process Folder",
            command=self.batch_process,
            font=("Arial", 10),
            width=20,
            cursor="hand2"
        )
        self.batch_btn.pack(pady=(10, 0))
        
        # Status Bar
        status_frame = Frame(self.root, bg="#f0f0f0", height=40)
        status_frame.pack(fill=X, side=BOTTOM)
        status_frame.pack_propagate(False)
        
        self.status_label = Label(
            status_frame,
            textvariable=self.status_text,
            bg="#f0f0f0",
            anchor=W,
            padx=10
        )
        self.status_label.pack(fill=BOTH, expand=True)
        
        # Progress bar
        self.progress = ttk.Progressbar(content_frame, mode='indeterminate')
    
    def browse_pdf(self):
        """Browse for PDF file"""
        filename = filedialog.askopenfilename(
            title="Select Packing List PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if filename:
            self.pdf_path.set(filename)
    
    def browse_template(self):
        """Browse for template file"""
        filename = filedialog.askopenfilename(
            title="Select CMR Template",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if filename:
            self.template_path.set(filename)
    
    def browse_output(self):
        """Browse for output directory"""
        dirname = filedialog.askdirectory(title="Select Output Directory")
        if dirname:
            self.output_dir.set(dirname)
    
    def process_pdf(self):
        """Process single PDF file"""
        
        if not self.pdf_path.get():
            messagebox.showerror("Error", "Please select a PDF file")
            return
        
        if not os.path.exists(self.pdf_path.get()):
            messagebox.showerror("Error", "Selected PDF file does not exist")
            return
        
        # Disable buttons
        self.process_btn.config(state=DISABLED)
        self.batch_btn.config(state=DISABLED)
        
        # Show progress
        self.progress.pack(fill=X, pady=(0, 15))
        self.progress.start(10)
        
        # Run in thread to avoid freezing UI
        thread = threading.Thread(target=self._process_pdf_thread)
        thread.daemon = True
        thread.start()
    
    def _process_pdf_thread(self):
        """Thread function for processing PDF"""
        try:
            self.status_text.set("Processing...")
            
            # Create output directory
            os.makedirs(self.output_dir.get(), exist_ok=True)
            
            # Extract data
            self.status_text.set("Extracting data from PDF...")
            extractor = PackingListExtractor(self.pdf_path.get())
            data = extractor.extract()
            
            # Generate output filename
            packing_list_no = data.get('our_ref') or data.get('packing_list_number')
            if not packing_list_no:
                packing_list_no = os.path.splitext(os.path.basename(self.pdf_path.get()))[0]
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_filename = f"CMR_{packing_list_no}_{timestamp}.xlsx"
            output_path = os.path.join(self.output_dir.get(), output_filename)
            
            # Populate template
            self.status_text.set("Creating CMR document...")
            populator = CMRExcelPopulator(self.template_path.get())
            populator.populate(data, output_path)
            
            # Success
            self.root.after(0, lambda path=output_path: self._process_complete(path))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self._process_error(msg))
    
    def _process_complete(self, output_path):
        """Handle successful processing"""
        self.progress.stop()
        self.progress.pack_forget()
        self.process_btn.config(state=NORMAL)
        self.batch_btn.config(state=NORMAL)
        self.status_text.set(f"Success! Saved to: {os.path.basename(output_path)}")
        
        result = messagebox.askyesno(
            "Success",
            f"CMR document created successfully!\n\n{os.path.basename(output_path)}\n\nWould you like to open it?"
        )
        
        if result:
            if sys.platform == "win32":
                os.startfile(output_path)
            elif sys.platform == "darwin":
                os.system(f"open '{output_path}'")
            else:
                os.system(f"xdg-open '{output_path}'")
    
    def _process_error(self, error_msg):
        """Handle processing error"""
        self.progress.stop()
        self.progress.pack_forget()
        self.process_btn.config(state=NORMAL)
        self.batch_btn.config(state=NORMAL)
        self.status_text.set("Error occurred")
        
        messagebox.showerror("Error", f"Failed to process PDF:\n\n{error_msg}")
    
    def batch_process(self):
        """Batch process multiple PDFs"""
        dirname = filedialog.askdirectory(title="Select Folder with Packing List PDFs")
        
        if not dirname:
            return
        
        # Count PDF files
        pdf_files = [f for f in os.listdir(dirname) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            messagebox.showinfo("No PDFs", "No PDF files found in selected directory")
            return
        
        result = messagebox.askyesno(
            "Batch Process",
            f"Found {len(pdf_files)} PDF file(s).\n\nProcess all files?"
        )
        
        if not result:
            return
        
        # Disable buttons
        self.process_btn.config(state=DISABLED)
        self.batch_btn.config(state=DISABLED)
        self.progress.pack(fill=X, pady=(0, 15))
        self.progress.start(10)
        
        # Run batch process in thread
        thread = threading.Thread(target=self._batch_process_thread, args=(dirname,))
        thread.daemon = True
        thread.start()
    
    def _batch_process_thread(self, input_dir):
        """Thread function for batch processing"""
        try:
            os.makedirs(self.output_dir.get(), exist_ok=True)
            
            pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')]
            
            successful = 0
            failed = 0
            
            for i, pdf_file in enumerate(pdf_files):
                try:
                    pdf_path = os.path.join(input_dir, pdf_file)
                    self.status_text.set(f"Processing {i+1}/{len(pdf_files)}: {pdf_file}")
                    
                    # Extract and process
                    extractor = PackingListExtractor(pdf_path)
                    data = extractor.extract()
                    
                    packing_list_no = data.get('our_ref') or data.get('packing_list_number')
                    if not packing_list_no:
                        packing_list_no = os.path.splitext(pdf_file)[0]
                    
                    timestamp = datetime.now().strftime('%Y%m%d')
                    output_filename = f"CMR_{packing_list_no}_{timestamp}.xlsx"
                    output_path = os.path.join(self.output_dir.get(), output_filename)
                    
                    populator = CMRExcelPopulator(self.template_path.get())
                    populator.populate(data, output_path)
                    
                    successful += 1
                    
                except Exception as e:
                    failed += 1
                    print(f"Error processing {pdf_file}: {e}")
            
            self.root.after(0, lambda s=successful, f=failed: self._batch_complete(s, f))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self._process_error(msg))
    
    def _batch_complete(self, successful, failed):
        """Handle batch processing completion"""
        self.progress.stop()
        self.progress.pack_forget()
        self.process_btn.config(state=NORMAL)
        self.batch_btn.config(state=NORMAL)
        self.status_text.set(f"Batch complete: {successful} successful, {failed} failed")
        
        messagebox.showinfo(
            "Batch Processing Complete",
            f"Processing complete!\n\nSuccessful: {successful}\nFailed: {failed}\n\nOutput: {self.output_dir.get()}"
        )


def main():
    """Main application entry point"""
    root = Tk()
    app = PDFtoCMRApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
