#!/usr/bin/env python3
"""
PDF to CMR Converter - Modern GUI Application
Professional interface for converting packing list PDFs to CMR Excel
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

# Import updater
try:
    from updater import check_for_updates, get_current_version
    UPDATER_AVAILABLE = True
except ImportError:
    UPDATER_AVAILABLE = False
    def get_current_version():
        return "1.0.0"


class ModernButton(Canvas):
    """Modern styled button with hover effects"""
    
    def __init__(self, parent, text, command, bg_color="#2563eb", hover_color="#1d4ed8", 
                 text_color="white", width=200, height=45, font=("Segoe UI", 11, "bold")):
        super().__init__(parent, width=width, height=height, highlightthickness=0, bg=parent['bg'])
        
        self.command = command
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.text_color = text_color
        self.text = text
        self.font = font
        
        # Draw button
        self.rect = self.create_rectangle(0, 0, width, height, fill=bg_color, outline="", tags="button")
        self.text_item = self.create_text(width/2, height/2, text=text, fill=text_color, 
                                          font=font, tags="button")
        
        # Bind events
        self.tag_bind("button", "<Button-1>", lambda e: self.command())
        self.tag_bind("button", "<Enter>", self.on_enter)
        self.tag_bind("button", "<Leave>", self.on_leave)
        
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        
    def on_enter(self, event=None):
        self.itemconfig(self.rect, fill=self.hover_color)
        self.config(cursor="hand2")
        
    def on_leave(self, event=None):
        self.itemconfig(self.rect, fill=self.bg_color)
        self.config(cursor="")
    
    def set_state(self, state):
        """Enable or disable button"""
        if state == "disabled":
            self.itemconfig(self.rect, fill="#9ca3af")
            self.unbind("<Button-1>")
            self.tag_unbind("button", "<Button-1>")
            self.config(cursor="")
        else:
            self.itemconfig(self.rect, fill=self.bg_color)
            self.tag_bind("button", "<Button-1>", lambda e: self.command())


class PDFtoCMRApp:
    """Modern GUI Application for PDF to CMR conversion"""
    
    # Color scheme - Modern blue theme
    COLORS = {
        'primary': '#2563eb',        # Blue
        'primary_hover': '#1d4ed8',  # Darker blue
        'secondary': '#64748b',      # Slate
        'success': '#10b981',        # Green
        'danger': '#ef4444',         # Red
        'background': '#f8fafc',     # Very light gray
        'surface': '#ffffff',        # White
        'text_primary': '#1e293b',   # Dark slate
        'text_secondary': '#64748b', # Medium slate
        'border': '#e2e8f0',         # Light border
    }
    
    def __init__(self, root):
        self.root = root
        version = get_current_version()
        self.root.title(f"CTS CMR Converter v{version}")
        self.root.geometry("750x600")
        self.root.resizable(False, False)
        self.root.configure(bg=self.COLORS['background'])
        
        # Variables
        self.pdf_path = StringVar()
        self.output_dir = StringVar(value="./cmr_output")
        self.status_text = StringVar(value="Ready to convert packing lists")
        
        # Template is hardcoded
        self.template_path = "CTS_NL_CMR_Template.xlsx"
        
        # Configure modern ttk style
        self.setup_styles()
        self.setup_ui()
        
        # Check for updates after 2 seconds
        if UPDATER_AVAILABLE:
            self.root.after(2000, self.check_for_updates_async)
    
    def setup_styles(self):
        """Setup modern ttk styles"""
        style = ttk.Style()
        
        # Configure progressbar
        style.theme_use('clam')
        style.configure("Modern.Horizontal.TProgressbar",
                       troughcolor=self.COLORS['border'],
                       background=self.COLORS['primary'],
                       borderwidth=0,
                       thickness=6)
    
    def setup_ui(self):
        """Setup the modern user interface"""
        
        # Header with gradient effect
        header_frame = Frame(self.root, bg=self.COLORS['primary'], height=100)
        header_frame.pack(fill=X)
        header_frame.pack_propagate(False)
        
        # Logo/Title area
        title_container = Frame(header_frame, bg=self.COLORS['primary'])
        title_container.pack(expand=True, pady=20)
        
        # Icon (you can replace with actual logo later)
        icon_label = Label(
            title_container,
            text="📄",
            font=("Segoe UI", 28),
            bg=self.COLORS['primary'],
            fg="white"
        )
        icon_label.pack(side=LEFT, padx=(0, 15))
        
        # Title
        title_label = Label(
            title_container,
            text="CTS CMR Converter",
            font=("Segoe UI", 24, "bold"),
            bg=self.COLORS['primary'],
            fg="white"
        )
        title_label.pack(side=LEFT)
        
        # Subtitle
        subtitle = Label(
            title_container,
            text="Packing List to CMR Document",
            font=("Segoe UI", 10),
            bg=self.COLORS['primary'],
            fg="#bfdbfe"
        )
        subtitle.pack(side=LEFT, padx=(10, 0), pady=(5, 0))
        
        # Main content area with padding
        content_frame = Frame(self.root, bg=self.COLORS['background'])
        content_frame.pack(fill=BOTH, expand=True, padx=30, pady=30)
        
        # PDF Selection Card
        self.create_card(
            content_frame,
            "📁 Select Packing List PDF",
            "Choose the PDF file you want to convert",
            self.pdf_path,
            self.browse_pdf,
            "Browse PDF..."
        ).pack(fill=X, pady=(0, 20))
        
        # Output Directory Card
        self.create_card(
            content_frame,
            "💾 Output Folder",
            "Where to save the CMR document",
            self.output_dir,
            self.browse_output,
            "Browse Folder..."
        ).pack(fill=X, pady=(0, 30))
        
        # Action buttons
        button_frame = Frame(content_frame, bg=self.COLORS['background'])
        button_frame.pack(pady=20)
        
        # Convert button (primary)
        self.convert_btn = ModernButton(
            button_frame,
            text="Convert to CMR",
            command=self.process_pdf,
            bg_color=self.COLORS['primary'],
            hover_color=self.COLORS['primary_hover'],
            width=250,
            height=50,
            font=("Segoe UI", 12, "bold")
        )
        self.convert_btn.pack(side=LEFT, padx=5)
        
        # Batch button (secondary)
        self.batch_btn = ModernButton(
            button_frame,
            text="Batch Convert Folder",
            command=self.batch_process,
            bg_color=self.COLORS['secondary'],
            hover_color="#475569",
            width=200,
            height=50,
            font=("Segoe UI", 11)
        )
        self.batch_btn.pack(side=LEFT, padx=5)
        
        # Progress bar (hidden by default)
        self.progress_frame = Frame(content_frame, bg=self.COLORS['background'])
        
        progress_label = Label(
            self.progress_frame,
            text="Processing...",
            font=("Segoe UI", 10),
            bg=self.COLORS['background'],
            fg=self.COLORS['text_secondary']
        )
        progress_label.pack(pady=(0, 8))
        
        self.progress = ttk.Progressbar(
            self.progress_frame,
            mode='indeterminate',
            style="Modern.Horizontal.TProgressbar",
            length=400
        )
        self.progress.pack()
        
        # Status bar at bottom
        status_frame = Frame(self.root, bg=self.COLORS['surface'], height=50)
        status_frame.pack(fill=X, side=BOTTOM)
        status_frame.pack_propagate(False)
        
        # Add subtle shadow effect
        separator = Frame(status_frame, bg=self.COLORS['border'], height=1)
        separator.pack(fill=X)
        
        status_container = Frame(status_frame, bg=self.COLORS['surface'])
        status_container.pack(fill=BOTH, expand=True, padx=20, pady=10)
        
        # Status icon
        self.status_icon = Label(
            status_container,
            text="✓",
            font=("Segoe UI", 12),
            bg=self.COLORS['surface'],
            fg=self.COLORS['success']
        )
        self.status_icon.pack(side=LEFT, padx=(0, 10))
        
        # Status text
        self.status_label = Label(
            status_container,
            textvariable=self.status_text,
            font=("Segoe UI", 10),
            bg=self.COLORS['surface'],
            fg=self.COLORS['text_secondary'],
            anchor=W
        )
        self.status_label.pack(side=LEFT, fill=X, expand=True)
    
    def create_card(self, parent, title, subtitle, variable, browse_command, button_text):
        """Create a modern card UI element"""
        # Card container with shadow effect
        card_outer = Frame(parent, bg=self.COLORS['border'])
        card = Frame(card_outer, bg=self.COLORS['surface'])
        card.pack(padx=1, pady=1, fill=BOTH, expand=True)
        
        # Card header
        header = Frame(card, bg=self.COLORS['surface'])
        header.pack(fill=X, padx=20, pady=(15, 5))
        
        title_label = Label(
            header,
            text=title,
            font=("Segoe UI", 12, "bold"),
            bg=self.COLORS['surface'],
            fg=self.COLORS['text_primary'],
            anchor=W
        )
        title_label.pack(side=TOP, anchor=W)
        
        subtitle_label = Label(
            header,
            text=subtitle,
            font=("Segoe UI", 9),
            bg=self.COLORS['surface'],
            fg=self.COLORS['text_secondary'],
            anchor=W
        )
        subtitle_label.pack(side=TOP, anchor=W, pady=(2, 0))
        
        # Input area
        input_frame = Frame(card, bg=self.COLORS['surface'])
        input_frame.pack(fill=X, padx=20, pady=(10, 15))
        
        # Entry with modern styling
        entry_container = Frame(input_frame, bg=self.COLORS['border'], height=40)
        entry_container.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        entry_container.pack_propagate(False)
        
        entry = Entry(
            entry_container,
            textvariable=variable,
            font=("Segoe UI", 10),
            relief=FLAT,
            bg=self.COLORS['surface'],
            fg=self.COLORS['text_primary'],
            state="readonly" if browse_command == self.browse_pdf else "normal"
        )
        entry.pack(fill=BOTH, expand=True, padx=1, pady=1)
        
        # Modern browse button
        browse_btn = ModernButton(
            input_frame,
            text=button_text,
            command=browse_command,
            bg_color=self.COLORS['secondary'],
            hover_color="#475569",
            width=150,
            height=40,
            font=("Segoe UI", 10)
        )
        browse_btn.pack(side=LEFT)
        
        return card_outer
    
    def browse_pdf(self):
        """Browse for PDF file"""
        filename = filedialog.askopenfilename(
            title="Select Packing List PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if filename:
            self.pdf_path.set(filename)
            self.update_status("PDF selected: " + os.path.basename(filename), "info")
    
    def browse_output(self):
        """Browse for output directory"""
        dirname = filedialog.askdirectory(title="Select Output Directory")
        if dirname:
            self.output_dir.set(dirname)
            self.update_status("Output folder set", "info")
    
    def update_status(self, message, status_type="success"):
        """Update status bar with icon and color"""
        self.status_text.set(message)
        
        if status_type == "success":
            self.status_icon.config(text="✓", fg=self.COLORS['success'])
        elif status_type == "error":
            self.status_icon.config(text="✗", fg=self.COLORS['danger'])
        elif status_type == "info":
            self.status_icon.config(text="ℹ", fg=self.COLORS['primary'])
        elif status_type == "processing":
            self.status_icon.config(text="⟳", fg=self.COLORS['primary'])
    
    def process_pdf(self):
        """Process single PDF file"""
        
        if not self.pdf_path.get():
            messagebox.showerror("No File Selected", "Please select a PDF file to convert")
            return
        
        if not os.path.exists(self.pdf_path.get()):
            messagebox.showerror("File Not Found", "The selected PDF file does not exist")
            return
        
        # Disable buttons
        self.convert_btn.set_state("disabled")
        self.batch_btn.set_state("disabled")
        
        # Show progress
        self.progress_frame.pack(fill=X, pady=20)
        self.progress.start(10)
        self.update_status("Processing PDF...", "processing")
        
        # Run in thread
        thread = threading.Thread(target=self._process_pdf_thread)
        thread.daemon = True
        thread.start()
    
    def _process_pdf_thread(self):
        """Thread function for processing PDF"""
        try:
            # Create output directory
            os.makedirs(self.output_dir.get(), exist_ok=True)
            
            # Extract data
            self.root.after(0, lambda: self.update_status("Extracting data from PDF...", "processing"))
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
            self.root.after(0, lambda: self.update_status("Creating CMR document...", "processing"))
            populator = CMRExcelPopulator(self.template_path)
            populator.populate(data, output_path)
            
            # Success
            self.root.after(0, lambda path=output_path: self._process_complete(path))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self._process_error(msg))
    
    def _process_complete(self, output_path):
        """Handle successful processing"""
        self.progress.stop()
        self.progress_frame.pack_forget()
        self.convert_btn.set_state("normal")
        self.batch_btn.set_state("normal")
        self.update_status(f"Success! Created: {os.path.basename(output_path)}", "success")
        
        result = messagebox.askyesno(
            "Conversion Complete",
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
        self.progress_frame.pack_forget()
        self.convert_btn.set_state("normal")
        self.batch_btn.set_state("normal")
        self.update_status("Conversion failed", "error")
        
        messagebox.showerror("Conversion Failed", f"An error occurred:\n\n{error_msg}")
    
    def batch_process(self):
        """Batch process multiple PDFs"""
        dirname = filedialog.askdirectory(title="Select Folder with Packing List PDFs")
        
        if not dirname:
            return
        
        # Count PDF files
        pdf_files = [f for f in os.listdir(dirname) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            messagebox.showinfo("No PDFs Found", "No PDF files found in the selected folder")
            return
        
        result = messagebox.askyesno(
            "Batch Convert",
            f"Found {len(pdf_files)} PDF file(s).\n\nConvert all files?"
        )
        
        if not result:
            return
        
        # Disable buttons
        self.convert_btn.set_state("disabled")
        self.batch_btn.set_state("disabled")
        self.progress_frame.pack(fill=X, pady=20)
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
                    self.root.after(0, lambda i=i, total=len(pdf_files), name=pdf_file: 
                                   self.update_status(f"Processing {i+1}/{total}: {name}", "processing"))
                    
                    # Extract and process
                    extractor = PackingListExtractor(pdf_path)
                    data = extractor.extract()
                    
                    packing_list_no = data.get('our_ref') or data.get('packing_list_number')
                    if not packing_list_no:
                        packing_list_no = os.path.splitext(pdf_file)[0]
                    
                    timestamp = datetime.now().strftime('%Y%m%d')
                    output_filename = f"CMR_{packing_list_no}_{timestamp}.xlsx"
                    output_path = os.path.join(self.output_dir.get(), output_filename)
                    
                    populator = CMRExcelPopulator(self.template_path)
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
        self.progress_frame.pack_forget()
        self.convert_btn.set_state("normal")
        self.batch_btn.set_state("normal")
        self.update_status(f"Batch complete: {successful} successful, {failed} failed", 
                          "success" if failed == 0 else "error")
        
        messagebox.showinfo(
            "Batch Processing Complete",
            f"Processing complete!\n\nSuccessful: {successful}\nFailed: {failed}\n\nOutput: {self.output_dir.get()}"
        )
    
    def check_for_updates_async(self):
        """Check for updates in background thread"""
        thread = threading.Thread(target=self._check_updates_thread)
        thread.daemon = True
        thread.start()
    
    def _check_updates_thread(self):
        """Background thread for update checking"""
        try:
            update_info = check_for_updates()
            
            if update_info.get('available'):
                self.root.after(0, lambda: self._show_update_dialog(update_info))
        
        except Exception as e:
            print(f"Update check error: {e}")
    
    def _show_update_dialog(self, update_info):
        """Show update available dialog"""
        source = update_info.get('source', 'unknown')
        version = update_info['version']
        notes = update_info.get('release_notes', 'No release notes available')
        
        message = (
            f"A new version is available!\n\n"
            f"Current version: {get_current_version()}\n"
            f"New version: {version}\n\n"
            f"Release notes:\n{notes[:200]}{'...' if len(notes) > 200 else ''}\n\n"
            f"Would you like to download the update?"
        )
        
        result = messagebox.askyesno("Update Available", message)
        
        if result:
            if source == 'network_share':
                # For network share, open file location
                import subprocess
                path = update_info['download_path']
                folder = os.path.dirname(path)
                subprocess.Popen(f'explorer /select,"{path}"')
            else:
                # For GitHub/web, open browser
                import webbrowser
                url = update_info.get('download_url')
                if url:
                    webbrowser.open(url)


def main():
    """Main application entry point"""
    root = Tk()
    
    # Set window icon if available
    try:
        if os.path.exists('icon.ico'):
            root.iconbitmap('icon.ico')
    except:
        pass
    
    app = PDFtoCMRApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
