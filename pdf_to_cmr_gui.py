#!/usr/bin/env python3
"""
PDF to CMR Converter - Modern GUI Application
Professional interface with manual browse + smart search
"""

import os
import sys
import threading
import glob
from pathlib import Path
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


class PDFSearcher:
    """Smart PDF searcher for project and packing list numbers"""
    
    def __init__(self, base_path="P:\\"):
        self.base_path = base_path
    
    def find_by_project_and_pl(self, project_num, pl_num):
        """Search by both project number and packing list number"""
        results = []
        year_folders = self._get_year_folders()
        
        for year_folder in year_folders:
            # Look for folders STARTING with project number
            project_folders = glob.glob(os.path.join(year_folder, f"{project_num} *"))
            
            for project_folder in project_folders:
                transport_folder = os.path.join(project_folder, "Transport")
                
                if os.path.exists(transport_folder):
                    pdfs = self._find_packing_list_pdfs(transport_folder, pl_num)
                    results.extend(pdfs)
        
        if len(results) == 0:
            return False, f"No packing list found for Project {project_num} / PL {pl_num}\n\nTry:\n• Check numbers\n• Try PL number only\n• Browse manually"
        elif len(results) == 1:
            return True, results[0]['path']
        else:
            return True, results
    
    def find_by_project_only(self, project_num):
        """Search by project number only"""
        results = []
        year_folders = self._get_year_folders()
        
        for year_folder in year_folders:
            project_folders = glob.glob(os.path.join(year_folder, f"{project_num} *"))
            
            for project_folder in project_folders:
                transport_folder = os.path.join(project_folder, "Transport")
                
                if os.path.exists(transport_folder):
                    pdfs = self._find_all_packing_lists(transport_folder)
                    results.extend(pdfs)
        
        if len(results) == 0:
            return False, f"No packing lists found in Project {project_num}\n\nTry:\n• Browse manually\n• Check if Transport folder exists"
        elif len(results) == 1:
            return True, results[0]['path']
        else:
            return True, results
    
    def find_by_pl_only(self, pl_num):
        """Search by packing list number only"""
        results = []
        year_folders = self._get_year_folders()
        
        for year_folder in year_folders:
            # Get all project folders
            project_folders = [d for d in glob.glob(os.path.join(year_folder, "*")) 
                             if os.path.isdir(d)]
            
            for project_folder in project_folders:
                transport_folder = os.path.join(project_folder, "Transport")
                
                if os.path.exists(transport_folder):
                    pdfs = self._find_packing_list_pdfs(transport_folder, pl_num)
                    results.extend(pdfs)
        
        if len(results) == 0:
            return False, f"No packing list {pl_num} found\n\nTry:\n• Check number\n• Browse manually"
        elif len(results) == 1:
            return True, results[0]['path']
        else:
            return True, results
    
    def _get_year_folders(self):
        """Get all year folders (2024, 2025, etc.)"""
        if not os.path.exists(self.base_path):
            return []
        
        year_folders = []
        for item in os.listdir(self.base_path):
            full_path = os.path.join(self.base_path, item)
            if os.path.isdir(full_path) and item.isdigit() and len(item) == 4:
                year_folders.append(full_path)
        
        return sorted(year_folders, reverse=True)
    
    def _find_packing_list_pdfs(self, folder, pl_num):
        """Find PDFs containing packing list number - handles PL16008, PL 16008, Cl16008"""
        results = []
        
        try:
            for file in os.listdir(folder):
                if file.lower().endswith('.pdf'):
                    # Check for various patterns:
                    # PL16008, PL 16008, Cl16008, pl16008, etc.
                    file_upper = file.upper()
                    
                    # Look for the 5-digit number with common prefixes
                    patterns = [
                        f"PL{pl_num}",      # PL16008
                        f"PL {pl_num}",     # PL 16008
                        f"CL{pl_num}",      # Cl16008
                        f"CL {pl_num}",     # Cl 16008
                        pl_num              # Just the number
                    ]
                    
                    if any(pattern in file_upper for pattern in patterns):
                        full_path = os.path.join(folder, file)
                        results.append({
                            'path': full_path,
                            'filename': file,
                            'modified': datetime.fromtimestamp(os.path.getmtime(full_path))
                        })
        except Exception as e:
            print(f"Error searching folder {folder}: {e}")
        
        return results
    
    def _find_all_packing_lists(self, folder):
        """Find all PDFs that look like packing lists"""
        results = []
        
        try:
            for file in os.listdir(folder):
                if file.lower().endswith('.pdf'):
                    file_upper = file.upper()
                    
                    # Look for files with PL or CL followed by 5 digits
                    # Also check for common packing list keywords
                    keywords = ['PL', 'CL', 'PACKING', 'PACKINGLIST']
                    
                    # Check if it has 5 consecutive digits (likely a PL number)
                    has_pl_pattern = False
                    for i in range(len(file) - 4):
                        if file[i:i+5].isdigit():
                            has_pl_pattern = True
                            break
                    
                    if any(kw in file_upper for kw in keywords) or has_pl_pattern:
                        full_path = os.path.join(folder, file)
                        results.append({
                            'path': full_path,
                            'filename': file,
                            'modified': datetime.fromtimestamp(os.path.getmtime(full_path))
                        })
        except Exception as e:
            print(f"Error searching folder {folder}: {e}")
        
        return results


class FileSelectionDialog:
    """Dialog for selecting from multiple PDF matches"""
    
    def __init__(self, parent, files):
        self.result = None
        self.dialog = Toplevel(parent)
        self.dialog.title("Select Packing List")
        self.dialog.geometry("700x450")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 700) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 450) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        # Title
        title_label = Label(self.dialog, text="Multiple packing lists found - select one:",
                          font=("Segoe UI", 12, "bold"), fg="#1e40af")
        title_label.pack(pady=15, padx=20)
        
        # List
        list_frame = Frame(self.dialog, bg="white")
        list_frame.pack(fill=BOTH, expand=True, padx=20, pady=(0, 15))
        
        scrollbar = Scrollbar(list_frame)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        self.listbox = Listbox(list_frame, font=("Segoe UI", 10), 
                              yscrollcommand=scrollbar.set, selectmode=SINGLE)
        self.listbox.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)
        
        # Add files
        self.files = files
        for file in files:
            display = f"{file['filename']}"
            self.listbox.insert(END, display)
            self.listbox.insert(END, f"  Modified: {file['modified'].strftime('%Y-%m-%d %H:%M')}")
            self.listbox.insert(END, "")  # Empty line
        
        self.listbox.selection_set(0)
        
        # Buttons
        button_frame = Frame(self.dialog, bg="white")
        button_frame.pack(pady=(0, 15))
        
        select_btn = Button(button_frame, text="Convert Selected", command=self.on_select,
                           bg="#2563eb", fg="white", font=("Segoe UI", 10, "bold"),
                           padx=20, pady=10, relief=FLAT, cursor="hand2")
        select_btn.pack(side=LEFT, padx=5)
        
        cancel_btn = Button(button_frame, text="Cancel", command=self.on_cancel,
                           bg="#6b7280", fg="white", font=("Segoe UI", 10),
                           padx=20, pady=10, relief=FLAT, cursor="hand2")
        cancel_btn.pack(side=LEFT, padx=5)
        
        self.listbox.bind("<Double-Button-1>", lambda e: self.on_select())
    
    def on_select(self):
        selection = self.listbox.curselection()
        if selection:
            # Selection index / 3 because we have 3 lines per file
            file_index = selection[0] // 3
            if file_index < len(self.files):
                self.result = self.files[file_index]['path']
                self.dialog.destroy()
    
    def on_cancel(self):
        self.dialog.destroy()
    
    def show(self):
        self.dialog.wait_window()
        return self.result


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
        self.enabled = True
        
        # Draw button
        self.rect = self.create_rectangle(0, 0, width, height, fill=bg_color, outline="", tags="button")
        self.text_item = self.create_text(width/2, height/2, text=text, fill=text_color, 
                                          font=font, tags="button")
        
        # Bind events
        self.tag_bind("button", "<Button-1>", lambda e: self.command() if self.enabled else None)
        self.tag_bind("button", "<Enter>", self.on_enter)
        self.tag_bind("button", "<Leave>", self.on_leave)
    
    def on_enter(self, event):
        if self.enabled:
            self.itemconfig(self.rect, fill=self.hover_color)
            self.config(cursor="hand2")
    
    def on_leave(self, event):
        if self.enabled:
            self.itemconfig(self.rect, fill=self.bg_color)
            self.config(cursor="")
    
    def set_state(self, state):
        """Enable or disable button"""
        self.enabled = (state != "disabled")
        if not self.enabled:
            self.itemconfig(self.rect, fill="#9ca3af")
            self.config(cursor="")
        else:
            self.itemconfig(self.rect, fill=self.bg_color)


class PDFtoCMRApp:
    """Modern GUI with manual browse + smart search"""
    
    COLORS = {
        'primary': '#2563eb',
        'primary_hover': '#1d4ed8',
        'secondary': '#64748b',
        'success': '#10b981',
        'danger': '#ef4444',
        'background': '#f8fafc',
        'surface': '#ffffff',
        'text_primary': '#1e293b',
        'text_secondary': '#64748b',
        'border': '#e2e8f0',
    }
    
    def __init__(self, root):
        self.root = root
        version = get_current_version()
        self.root.title(f"CTS CMR Converter v{version}")
        self.root.geometry("750x700")
        self.root.resizable(True, True)
        self.root.configure(bg=self.COLORS['background'])
        
        # Variables
        self.selected_pdf = None
        self.template_path = "CTS_NL_CMR_Template.xlsx"
        self.searcher = PDFSearcher()
        
        # Build UI
        self.create_widgets()
        self.center_window()
        
        # Check for updates
        if UPDATER_AVAILABLE:
            self.root.after(2000, self.check_updates)
    
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_widgets(self):
        # Main container
        main_frame = Frame(self.root, bg=self.COLORS['background'])
        main_frame.pack(fill=BOTH, expand=True, padx=30, pady=30)
        
        # Header
        header_frame = Frame(main_frame, bg=self.COLORS['background'])
        header_frame.pack(fill=X, pady=(0, 25))
        
        title = Label(header_frame, text="CTS CMR Converter", 
                     font=("Segoe UI", 24, "bold"), fg="#1e40af", bg=self.COLORS['background'])
        title.pack()
        
        subtitle = Label(header_frame, text="Convert Packing Lists to CMR Documents",
                        font=("Segoe UI", 11), fg=self.COLORS['text_secondary'], bg=self.COLORS['background'])
        subtitle.pack()
        
        # Main browse card (original design)
        browse_card = self.create_card(main_frame)
        browse_card.pack(fill=X, pady=(0, 15))
        
        browse_content = Frame(browse_card, bg=self.COLORS['surface'])
        browse_content.pack(fill=BOTH, padx=25, pady=25)
        
        # Browse section
        browse_title = Label(browse_content, text="📁 Select Packing List PDF",
                            font=("Segoe UI", 14, "bold"), fg=self.COLORS['text_primary'],
                            bg=self.COLORS['surface'])
        browse_title.pack(anchor=W, pady=(0, 15))
        
        # File selection
        file_frame = Frame(browse_content, bg=self.COLORS['surface'])
        file_frame.pack(fill=X, pady=(0, 20))
        
        self.file_label = Label(file_frame, text="No file selected", 
                               font=("Segoe UI", 10), fg=self.COLORS['text_secondary'],
                               bg="#f1f5f9", anchor=W, padx=15, pady=12, relief=SOLID, bd=1)
        self.file_label.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))
        
        browse_file_btn = Button(file_frame, text="Browse...", command=self.browse_file,
                                bg=self.COLORS['secondary'], fg="white", 
                                font=("Segoe UI", 10, "bold"), relief=FLAT, 
                                padx=20, pady=10, cursor="hand2")
        browse_file_btn.pack(side=LEFT)
        
        # Convert button
        convert_frame = Frame(browse_content, bg=self.COLORS['surface'])
        convert_frame.pack(pady=(10, 0))
        
        self.convert_btn = ModernButton(convert_frame, "Convert to CMR", 
                                        self.convert_pdf, width=250, height=50)
        self.convert_btn.pack()
        self.convert_btn.set_state("disabled")
        
        # Smart search section (collapsible)
        search_section = Frame(main_frame, bg=self.COLORS['background'])
        search_section.pack(fill=X, pady=(15, 0))
        
        # Search toggle button
        toggle_frame = Frame(search_section, bg=self.COLORS['background'])
        toggle_frame.pack(fill=X, pady=(0, 10))
        
        self.search_visible = False
        self.toggle_btn = Button(toggle_frame, text="▼ Show Smart Search (Project/PL Number)",
                                command=self.toggle_search,
                                bg="#e5e7eb", fg=self.COLORS['text_primary'],
                                font=("Segoe UI", 10), relief=FLAT, anchor=W,
                                padx=15, pady=10, cursor="hand2")
        self.toggle_btn.pack(fill=X)
        
        # Search card (hidden by default)
        self.search_card = self.create_card(search_section)
        
        search_content = Frame(self.search_card, bg=self.COLORS['surface'])
        search_content.pack(fill=BOTH, padx=25, pady=20)
        
        # Base folder
        folder_frame = Frame(search_content, bg=self.COLORS['surface'])
        folder_frame.pack(fill=X, pady=(0, 10))
        
        Label(folder_frame, text="Base Folder:", font=("Segoe UI", 9), 
              fg=self.COLORS['text_secondary'], bg=self.COLORS['surface']).pack(anchor=W, pady=(0, 3))
        
        folder_input = Frame(folder_frame, bg=self.COLORS['surface'])
        folder_input.pack(fill=X)
        
        self.folder_var = StringVar(value="P:\\")
        folder_entry = Entry(folder_input, textvariable=self.folder_var,
                            font=("Segoe UI", 9), relief=SOLID, bd=1)
        folder_entry.pack(side=LEFT, fill=X, expand=True, ipady=5, padx=(0, 8))
        
        Button(folder_input, text="Browse", command=self.browse_base_folder,
              bg="#e5e7eb", fg=self.COLORS['text_primary'], font=("Segoe UI", 8),
              relief=FLAT, padx=12, pady=5, cursor="hand2").pack(side=LEFT)
        
        # Project and PL number side by side
        fields_frame = Frame(search_content, bg=self.COLORS['surface'])
        fields_frame.pack(fill=X, pady=(0, 10))
        
        # Left column - Project
        left_col = Frame(fields_frame, bg=self.COLORS['surface'])
        left_col.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))
        
        Label(left_col, text="Project (4 digits):",
              font=("Segoe UI", 9), fg=self.COLORS['text_secondary'],
              bg=self.COLORS['surface']).pack(anchor=W, pady=(0, 3))
        
        self.project_var = StringVar()
        Entry(left_col, textvariable=self.project_var,
             font=("Segoe UI", 10), relief=SOLID, bd=1).pack(fill=X, ipady=6)
        
        # Right column - PL
        right_col = Frame(fields_frame, bg=self.COLORS['surface'])
        right_col.pack(side=LEFT, fill=BOTH, expand=True)
        
        Label(right_col, text="PL Number (5 digits):",
              font=("Segoe UI", 9), fg=self.COLORS['text_secondary'],
              bg=self.COLORS['surface']).pack(anchor=W, pady=(0, 3))
        
        self.pl_var = StringVar()
        Entry(right_col, textvariable=self.pl_var,
             font=("Segoe UI", 10), relief=SOLID, bd=1).pack(fill=X, ipady=6)
        
        # Search button
        search_btn_frame = Frame(search_content, bg=self.COLORS['surface'])
        search_btn_frame.pack(pady=(5, 0))
        
        self.search_btn = ModernButton(search_btn_frame, "🔍 Search & Convert",
                                       self.smart_search, width=220, height=45,
                                       font=("Segoe UI", 10, "bold"))
        self.search_btn.pack()
        
        # Status bar
        status_frame = Frame(main_frame, bg="#f1f5f9", relief=FLAT)
        status_frame.pack(fill=X, pady=(20, 0))
        
        status_content = Frame(status_frame, bg="#f1f5f9")
        status_content.pack(fill=X, padx=20, pady=12)
        
        Label(status_content, text="Status:", font=("Segoe UI", 9, "bold"),
              fg=self.COLORS['text_secondary'], bg="#f1f5f9").pack(side=LEFT, padx=(0, 8))
        
        self.status_label = Label(status_content, text="Ready", font=("Segoe UI", 9),
                                 fg=self.COLORS['success'], bg="#f1f5f9")
        self.status_label.pack(side=LEFT)
    
    def create_card(self, parent):
        """Create a card-style frame"""
        card = Frame(parent, bg=self.COLORS['surface'], relief=FLAT, bd=0)
        # Shadow effect
        shadow = Frame(parent, bg="#e2e8f0", height=2)
        shadow.place(in_=card, relx=0, rely=1, relwidth=1)
        return card
    
    def toggle_search(self):
        """Toggle smart search visibility"""
        self.search_visible = not self.search_visible
        
        if self.search_visible:
            self.search_card.pack(fill=X)
            self.toggle_btn.config(text="▲ Hide Smart Search")
        else:
            self.search_card.pack_forget()
            self.toggle_btn.config(text="▼ Show Smart Search (Project/PL Number)")
    
    def browse_file(self):
        """Browse for PDF file"""
        file_path = filedialog.askopenfilename(
            title="Select Packing List PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.selected_pdf = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=filename, fg=self.COLORS['text_primary'])
            self.convert_btn.set_state("normal")
    
    def browse_base_folder(self):
        """Browse for base projects folder"""
        folder = filedialog.askdirectory(initialdir=self.folder_var.get(),
                                        title="Select Base Projects Folder")
        if folder:
            self.folder_var.set(folder)
            self.searcher.base_path = folder
    
    def smart_search(self):
        """Smart search based on user input"""
        project_num = self.project_var.get().strip()
        pl_num = self.pl_var.get().strip()
        
        if not project_num and not pl_num:
            messagebox.showwarning("Input Required",
                                  "Please enter a project number, packing list number, or both.")
            return
        
        self.searcher.base_path = self.folder_var.get()
        self.set_status("Searching...", "#0369a1")
        self.root.update()
        
        # Search
        try:
            if project_num and pl_num:
                success, result = self.searcher.find_by_project_and_pl(project_num, pl_num)
            elif project_num:
                success, result = self.searcher.find_by_project_only(project_num)
            else:
                success, result = self.searcher.find_by_pl_only(pl_num)
            
            if not success:
                self.set_status("Not found", "#dc2626")
                messagebox.showerror("Not Found", result)
                return
            
            # Handle result
            if isinstance(result, list):
                dialog = FileSelectionDialog(self.root, result)
                selected_path = dialog.show()
                
                if selected_path:
                    self.selected_pdf = selected_path
                    self.do_conversion()
                else:
                    self.set_status("Cancelled", self.COLORS['text_secondary'])
            else:
                self.selected_pdf = result
                self.do_conversion()
        except Exception as e:
            self.set_status("Error", "#dc2626")
            messagebox.showerror("Search Error", f"Search failed:\n\n{str(e)}")
    
    def convert_pdf(self):
        """Convert selected PDF"""
        if not self.selected_pdf:
            messagebox.showwarning("No File", "Please select a PDF file first.")
            return
        
        self.do_conversion()
    
    def do_conversion(self):
        """Perform the conversion"""
        self.set_status("Converting...", "#0369a1")
        self.convert_btn.set_state("disabled")
        self.search_btn.set_state("disabled")
        
        # Run in thread
        thread = threading.Thread(target=self._conversion_thread)
        thread.daemon = True
        thread.start()
    
    def _conversion_thread(self):
        """Conversion logic (runs in thread)"""
        try:
            extractor = PackingListExtractor(self.selected_pdf)
            data = extractor.extract()
            
            base_name = os.path.splitext(os.path.basename(self.selected_pdf))[0]
            output_dir = "cmr_output"
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, 
                                      f"CMR_{base_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            
            populator = CMRExcelPopulator(self.template_path)
            populator.populate(data, output_path)
            
            self.root.after(0, lambda: self.on_success(output_path))
            
        except Exception as e:
            self.root.after(0, lambda: self.on_error(str(e)))
    
    def on_success(self, output_path):
        """Handle success"""
        self.set_status("✓ Success!", self.COLORS['success'])
        self.convert_btn.set_state("normal")
        self.search_btn.set_state("normal")
        
        result = messagebox.askyesno("Success", 
                                    f"CMR created successfully!\n\n{output_path}\n\nOpen folder?")
        if result:
            os.startfile(os.path.dirname(output_path))
    
    def on_error(self, error_msg):
        """Handle error"""
        self.set_status("✗ Error", self.COLORS['danger'])
        self.convert_btn.set_state("normal")
        self.search_btn.set_state("normal")
        messagebox.showerror("Error", f"Conversion failed:\n\n{error_msg}")
    
    def set_status(self, text, color):
        """Update status"""
        self.status_label.config(text=text, fg=color)
    
    def check_updates(self):
        """Check for updates"""
        try:
            update_info = check_for_updates()
            if update_info and update_info.get('available'):
                result = messagebox.askyesno(
                    "Update Available",
                    f"New version available!\n\n"
                    f"Current: {get_current_version()}\n"
                    f"New: {update_info['version']}\n\n"
                    f"Download now?",
                    icon='info'
                )
                if result:
                    import webbrowser
                    webbrowser.open(update_info.get('download_url', 
                                   'https://github.com/Yewcake/cts-cmr-converter/releases'))
        except Exception as e:
            print(f"Update check failed: {e}")


def main():
    root = Tk()
    app = PDFtoCMRApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
