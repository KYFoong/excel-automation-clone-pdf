import customtkinter as ctk
from tkinter import ttk
from customtkinter import filedialog
from CTkMessagebox import CTkMessagebox
import os
import shutil
import threading
import pythoncom  # Required for COM in threads
from win32com import client

# Configuration
ctk.set_appearance_mode('System')
ctk.set_default_color_theme('blue')

class ExcelManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Setup
        self.title('Excel File Manager')
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # State Variables
        self.current_files = []
        self.is_processing = False

        # Icons (Unicode)
        self.ICON_CHECKED = '\u2611'
        self.ICON_UNCHECKED = '\u2610'

        self.setup_ui()

    def setup_ui(self):
        # --- Section 1: Source ---
        self.frame_source = ctk.CTkFrame(self)
        self.frame_source.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        self.frame_source.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.frame_source, text="SOURCE DIRECTORY", font=("Arial", 12, "bold"), text_color="gray").grid(row=0, column=0, sticky="w", padx=15, pady=(10, 0))

        self.entry_source = ctk.CTkEntry(self.frame_source, placeholder_text="Select a folder to scan...", height=35)
        self.entry_source.configure(state='disabled')
        self.entry_source.grid(row=1, column=0, columnspan=2, sticky="ew", padx=15, pady=(5, 15))

        ctk.CTkButton(self.frame_source, text='Browse Folder', command=self.select_source_folder, width=120).grid(row=1, column=2, padx=15, pady=(5, 15))

        # --- Section 2: File Table ---
        self.frame_table = ctk.CTkFrame(self)
        self.frame_table.grid(row=1, column=0, sticky="nsew", padx=20, pady=5)
        self.frame_table.grid_columnconfigure(0, weight=1)
        self.frame_table.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self.frame_table, text="DETECTED FILES", font=("Arial", 12, "bold"), text_color="gray").grid(row=0, column=0, sticky="w", padx=15, pady=10)

        # Treeview Container
        tree_container = ctk.CTkFrame(self.frame_table, fg_color="transparent")
        tree_container.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=(0, 10))
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(0, weight=1)

        columns = ('select', 'name', 'type', 'size')
        self.tree = ttk.Treeview(tree_container, columns=columns, show='headings', selectmode='browse')
        
        # Scrollbar
        scrollbar = ctk.CTkScrollbar(tree_container, orientation="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Tree Headers
        self.tree.heading("select", text="Select", command=self.toggle_all_selection)
        self.tree.heading("name", text="File Name", anchor="w")
        self.tree.heading("type", text="Type")
        self.tree.heading("size", text="Size")

        self.tree.column("select", width=50, anchor="center", stretch=False)
        self.tree.column("name", width=400, anchor="w")
        self.tree.column("type", width=100, anchor="center")
        self.tree.column("size", width=100, anchor="center")

        self.tree.bind('<Button-1>', self.on_click)

        # --- Section 3: Actions ---
        self.frame_dest = ctk.CTkFrame(self)
        self.frame_dest.grid(row=2, column=0, sticky="ew", padx=20, pady=(10, 20))
        self.frame_dest.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.frame_dest, text="ACTIONS & EXPORT", font=("Arial", 12, "bold"), text_color="gray").grid(row=0, column=0, sticky="w", padx=15, pady=(10, 0))

        self.entry_dest = ctk.CTkEntry(self.frame_dest, placeholder_text="Select destination folder...", height=35)
        self.entry_dest.configure(state='disabled')
        self.entry_dest.grid(row=1, column=0, columnspan=2, sticky="ew", padx=15, pady=(5, 15))

        ctk.CTkButton(self.frame_dest, text='Browse Folder', command=self.select_dest_folder, width=120).grid(row=1, column=2, padx=15, pady=(5, 15))

        # Action Buttons
        action_grid = ctk.CTkFrame(self.frame_dest, fg_color="transparent")
        action_grid.grid(row=2, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 15))
        action_grid.grid_columnconfigure((0, 1), weight=1)

        # Clone Button
        card_clone = ctk.CTkFrame(action_grid, border_width=1, border_color='#404040') 
        card_clone.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=0)
        card_clone.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(card_clone, text='Copy selected files to destination.', text_color='gray', font=('Arial', 11)).grid(row=0, column=0, padx=10, pady=5)
        self.btn_clone = ctk.CTkButton(card_clone, text='Clone Files', command=self.action_clone, fg_color="#2CC985", hover_color="#229965", text_color="black")
        self.btn_clone.grid(row=1, column=0, padx=5, pady=(0, 15), sticky="ew")

        # PDF Button
        card_pdf = ctk.CTkFrame(action_grid, border_width=1, border_color='#404040') 
        card_pdf.grid(row=0, column=1, sticky="nsew", padx=(0, 5), pady=0)
        card_pdf.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(card_pdf, text='Convert selected files to PDF.', text_color='gray', font=('Arial', 11)).grid(row=0, column=0, padx=10, pady=5)
        self.btn_pdf = ctk.CTkButton(card_pdf, text='Convert to PDF', command=self.start_pdf_thread, fg_color="#3B8ED0", hover_color="#2A699C")
        self.btn_pdf.grid(row=1, column=0, padx=5, pady=(0, 15), sticky="ew")

        # Progress Bar (Hidden by default)
        self.progress_bar = ctk.CTkProgressBar(self.frame_dest, mode='determinate')
        self.progress_bar.set(0)

    # --- Functions ---

    def select_source_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.update_entry(self.entry_source, folder)
            self.populate_table(folder)

    def select_dest_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.update_entry(self.entry_dest, folder)

    def update_entry(self, entry_widget, pathText):
        entry_widget.configure(state='normal')
        entry_widget.delete(0, 'end')
        entry_widget.insert(0, pathText)
        entry_widget.configure(state='disabled')

    def populate_table(self, folder_path):
        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.current_files.clear()

        try:
            files = os.listdir(folder_path)
        except OSError as e:
            CTkMessagebox(title="Error", message=f"Cannot access folder: {e}", icon="cancel")
            return

        for file in files:
            ext = os.path.splitext(file)[1].lower()
            # Filter excel files
            if ext in (".xlsx", ".xls", ".csv"):
                full_path = os.path.join(folder_path, file)
                try:
                    size_mb = os.path.getsize(full_path) / 1024  # KB
                    size_str = f"{size_mb:.2f} KB"
                except OSError:
                    size_str = "Unknown"

                self.tree.insert('', 'end', values=(self.ICON_CHECKED, file, ext, size_str))

    def on_click(self, event):
        region = self.tree.identify('region', event.x, event.y)
        if region == 'cell':
            column = self.tree.identify_column(event.x)
            row_id = self.tree.identify_row(event.y)
            
            # Column #1 is the select box
            if column == '#1' and row_id:
                current_values = self.tree.item(row_id, 'values')
                new_status = self.ICON_UNCHECKED if current_values[0] == self.ICON_CHECKED else self.ICON_CHECKED
                
                # Update only the first value, keep others
                new_values = (new_status, *current_values[1:])
                self.tree.item(row_id, values=new_values)

    def toggle_all_selection(self):
        # Check the state of the first item to decide
        children = self.tree.get_children()
        if not children: return

        first_val = self.tree.item(children[0], 'values')[0]
        target_val = self.ICON_UNCHECKED if first_val == self.ICON_CHECKED else self.ICON_CHECKED

        for item in children:
            vals = self.tree.item(item, 'values')
            self.tree.item(item, values=(target_val, *vals[1:]))

    def get_selected_files(self):
        selected = []
        for item in self.tree.get_children():
            vals = self.tree.item(item, 'values')
            if vals[0] == self.ICON_CHECKED:
                selected.append(vals[1])
        return selected

    def validate_paths(self):
        src = self.entry_source.get().strip()
        dst = self.entry_dest.get().strip()
        files = self.get_selected_files()

        if not src or not dst:
            CTkMessagebox(title="Missing Info", message="Please select Source and Destination folders.", icon="warning")
            return None
        if not files:
            CTkMessagebox(title="No Files", message="Please select at least one file.", icon="warning")
            return None
        return src, dst, files

    # --- Actions ---

    def action_clone(self):
        data = self.validate_paths()
        if not data: return
        src, dst, files = data

        success_count = 0
        for file in files:
            try:
                shutil.copy2(os.path.join(src, file), dst)
                success_count += 1
            except Exception as e:
                print(f"Error copying {file}: {e}")
        
        CTkMessagebox(title="Cloning Complete", message=f"Copied {success_count} files successfully.", icon="check")

    def start_pdf_thread(self):
        data = self.validate_paths()
        if not data: return
        
        if self.is_processing:
            return

        self.is_processing = True
        self.btn_pdf.configure(state="disabled", text="Processing...")
        self.btn_clone.configure(state="disabled")
        self.progress_bar.grid(row=3, column=0, columnspan=3, sticky="ew", padx=15, pady=(0, 15))
        self.progress_bar.set(0)

        # Run heavy task in separate thread
        threading.Thread(target=self.convert_pdf_worker, args=data, daemon=True).start()

    def convert_pdf_worker(self, src_path, dst_path, file_list):
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        excel = None
        count = 0
        errors = []
        total_files = len(file_list)

        try:
            excel = client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            for index, file in enumerate(file_list):
                try:
                    input_file = os.path.abspath(os.path.join(src_path, file))
                    pdf_name = os.path.splitext(file)[0] + '.pdf'
                    output_file = os.path.abspath(os.path.join(dst_path, pdf_name))

                    wb = excel.Workbooks.Open(input_file)
                    try:
                        # 0 = xlTypePDF
                        wb.ExportAsFixedFormat(0, output_file)
                    finally:
                        wb.Close(SaveChanges=False)
                    
                    count += 1
                except Exception as e:
                    errors.append(f"{file}: {e}")
                
                # Update progress bar safely on main thread
                progress = (index + 1) / total_files
                self.after(0, lambda p=progress: self.progress_bar.set(p))

        except Exception as e:
            errors.append(f"Excel Launch Error: {e}")
        finally:
            if excel:
                excel.Quit()
            # Uninitialize COM
            pythoncom.CoUninitialize()
            
            # Reset UI on main thread
            self.after(0, lambda: self.conversion_finished(count, errors))

    def conversion_finished(self, count, errors):
        self.is_processing = False
        self.btn_pdf.configure(state="normal", text="Convert to PDF")
        self.btn_clone.configure(state="normal")
        self.progress_bar.grid_forget()

        if errors:
            msg = f"Converted: {count}\nFailed: {len(errors)}\n\nFirst Error: {errors[0]}"
            CTkMessagebox(title="Done with Errors", message=msg, icon="warning")
        else:
            CTkMessagebox(title="Success", message=f"Successfully converted {count} files.", icon="check")