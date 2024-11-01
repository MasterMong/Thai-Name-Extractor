import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from docx import Document
from collections import Counter
import re
from openpyxl import Workbook
from datetime import datetime
from PyPDF2 import PdfReader
from typing import List, Tuple, Dict
import logging
from pathlib import Path

# Constants
THAI_FONT = ('TH Sarabun New', 14)
WINDOW_SIZE = "800x600"
THAI_TITLES = r'''|เด็กหญิง|เด็กชาย|ด.ญ.|ด.ช.|นาย|นาง|นางสาว|
ว่าที่ร้อยตรี|ว่าที่ร้อยโท|ว่าที่ร้อยเอก|
ว่าที่เรือตรี|ว่าที่เรือโท|ว่าที่เรือเอก|
ว่าที่พันตรี|ว่าที่พันโท|ว่าที่พันเอก|
ร้อยตรี|ร้อยโท|ร้อยเอก|
เรือตรี|เรือโท|เรือเอก|
พันตรี|พันโท|พันเอก|
พลทหาร|พลตำรวจ|สิบตรี|สิบโท|สิบเอก|
จ่าสิบตรี|จ่าสิบโท|จ่าสิบเอก|
จ่าตรี|จ่าโท|จ่าเอก|
พลตำรวจตรี|พลตำรวจโท|พลตำรวจเอก|
พลโท|พลเอก|พลตรี|
ดร|รศ|ผศ|ศ|
หม่อมราชวงศ์|หม่อมหลวง|
พระ|พระครู|พระมหา'''
NAME_PATTERN = f'\\d+\\)\\s*({THAI_TITLES}[ก-์\\s]+[ก-์]+)'

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='name_extractor.log'
)

class NameExtractorApp:
    """
    A GUI application for extracting and processing Thai names from DOCX and PDF files.
    Provides functionality to sort, filter, and export names to Excel.
    """
    
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Thai Name Extractor")
        self.root.geometry(WINDOW_SIZE)
        
        self.all_names: List[Tuple[str, int]] = []
        self.sort_reverse: Dict[str, bool] = {'Name': False, 'Count': True}
        
        self._setup_gui()
        
    def _setup_gui(self) -> None:
        """Initialize and configure all GUI elements"""
        # Progress bar (new)
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        self.progress.pack(fill=tk.X, padx=10, pady=5)
        
        # Style configuration (new)
        style = ttk.Style()
        style.configure("Custom.Treeview", font=THAI_FONT)
        style.configure("Custom.TButton", font=THAI_FONT)
        
        # Button frame with improved layout
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10, fill=tk.X, padx=10)
        
        self.select_btn = ttk.Button(
            button_frame, 
            text="เลือกไฟล์ DOCX, PDF",
            command=self.select_file,
            style="Custom.TButton"
        )
        self.select_btn.pack(side=tk.LEFT, padx=5, expand=True)
        
        self.export_btn = ttk.Button(
            button_frame,
            text="ส่งออก Excel",
            command=self.export_to_excel,
            style="Custom.TButton"
        )
        self.export_btn.pack(side=tk.LEFT, padx=5, expand=True)

        # Search frame with improved layout
        self._setup_search_frame()
        
        # Treeview with improved configuration
        self._setup_treeview()

    def _setup_search_frame(self) -> None:
        """Setup search functionality"""
        search_frame = ttk.Frame(self.root)
        search_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(
            search_frame,
            text="ค้นหา:",
            font=THAI_FONT
        ).pack(side=tk.LEFT, padx=5)
        
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_names)
        ttk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=THAI_FONT
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

    def clean_thai_name(self, name: str) -> str:
        """
        Clean and normalize Thai names.
        
        Args:
            name: Raw Thai name string
            
        Returns:
            Cleaned name string
        """
        name = re.sub(r'\s+', ' ', name).strip()
        parts = name.split()
        return ' '.join(parts[:2]) if parts else ''

    def process_file(self, file_path: str) -> None:
        """
        Process the selected file and extract names.
        
        Args:
            file_path: Path to the input file
        """
        self.progress.start()
        self.select_btn.config(state='disabled')
        
        try:
            # File type specific processing
            text = (self.extract_text_from_pdf(file_path) 
                   if file_path.lower().endswith('.pdf')
                   else self.extract_text_from_docx(file_path))
            
            names = re.findall(NAME_PATTERN, text)
            clean_names = [self.clean_thai_name(name) for name in names]
            self.all_names = sorted(Counter(clean_names).items())
            
            self._update_treeview()
            
            if not self.all_names:
                messagebox.showwarning("คำเตือน", "ไม่พบรายชื่อในไฟล์")
                
        except Exception as e:
            logging.error(f"Error processing file: {str(e)}")
            messagebox.showerror("Error", f"เกิดข้อผิดพลาด: {str(e)}")
        finally:
            self.progress.stop()
            self.select_btn.config(state='normal')

    # ... (keep other existing methods but add proper type hints)

    def export_to_excel(self) -> None:
        """Export the processed names to an Excel file"""
        if not self.all_names:
            messagebox.showwarning("คำเตือน", "ไม่มีข้อมูลที่จะส่งออก")
            return
            
        try:
            default_filename = f"name_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                initialfile=default_filename
            )
            
            if not file_path:
                return
                
            wb = Workbook()
            ws = wb.active
            ws.title = "Name List"
            
            # Add headers with style
            for col, header in enumerate(['ชื่อ', 'จำนวน'], 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = cell.font.copy(bold=True)
            
            # Add data with auto-adjusted column width
            for row, (name, count) in enumerate(self.all_names, start=2):
                ws.cell(row=row, column=1, value=name)
                ws.cell(row=row, column=2, value=count)
            
            # Auto-adjust column widths
            for column in ws.columns:
                max_length = max(len(str(cell.value)) for cell in column)
                ws.column_dimensions[column[0].column_letter].width = max_length + 2
            
            wb.save(file_path)
            messagebox.showinfo("สำเร็จ", "ส่งออกไฟล์ Excel เรียบร้อยแล้ว")
            
        except Exception as e:
            logging.error(f"Error exporting to Excel: {str(e)}")
            messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการส่งออก: {str(e)}")

def main():
    root = tk.Tk()
    app = NameExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
