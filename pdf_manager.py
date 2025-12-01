#!/usr/bin/env python3
"""
PDF Manager - A GUI application for PDF manipulation
Features:
- Slice PDF files to selected page ranges
- Merge multiple PDF files
- Convert PPTX files to PDF and merge them
- Select custom output paths
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from pptx import Presentation
from PIL import Image
import io
import tempfile


class PDFManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Manager")
        self.root.geometry("1600x1200")
        
        # Variables
        self.output_path = tk.StringVar(value=os.path.expanduser("~"))
        
        # Create notebook (tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_slice_tab()
        self.create_merge_tab()
        self.create_pptx_tab()
        
    def create_slice_tab(self):
        """Create the PDF slicing tab"""
        slice_frame = ttk.Frame(self.notebook)
        self.notebook.add(slice_frame, text="Slice PDF")
        
        # Input file
        ttk.Label(slice_frame, text="Select PDF File:").pack(pady=(20, 5))
        
        input_frame = ttk.Frame(slice_frame)
        input_frame.pack(fill='x', padx=20, pady=5)
        
        self.slice_input_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.slice_input_var, width=50).pack(side='left', fill='x', expand=True)
        ttk.Button(input_frame, text="Browse", command=self.browse_slice_input).pack(side='left', padx=(5, 0))
        
        # Page ranges section
        ttk.Label(slice_frame, text="Page Ranges (one per line, e.g., '1-5' or '10-15'):").pack(pady=(20, 5))
        
        # Text widget for multiple ranges
        ranges_frame = ttk.Frame(slice_frame)
        ranges_frame.pack(fill='both', expand=True, padx=20, pady=5)
        
        ranges_scrollbar = ttk.Scrollbar(ranges_frame)
        ranges_scrollbar.pack(side='right', fill='y')
        
        self.slice_ranges_text = tk.Text(ranges_frame, height=6, width=50, yscrollcommand=ranges_scrollbar.set)
        self.slice_ranges_text.pack(side='left', fill='both', expand=True)
        ranges_scrollbar.config(command=self.slice_ranges_text.yview)
        
        # Buttons for range management
        range_buttons_frame = ttk.Frame(slice_frame)
        range_buttons_frame.pack(pady=5)
        
        ttk.Button(range_buttons_frame, text="Add Current Range", command=self.add_current_range).pack(side='left', padx=5)
        ttk.Button(range_buttons_frame, text="Clear Ranges", command=self.clear_ranges).pack(side='left', padx=5)
        
        # Single page range input (for quick add)
        quick_range_frame = ttk.Frame(slice_frame)
        quick_range_frame.pack(padx=20, pady=5)
        
        ttk.Label(quick_range_frame, text="Quick Add - From:").pack(side='left')
        self.slice_start_var = tk.StringVar(value="1")
        ttk.Entry(quick_range_frame, textvariable=self.slice_start_var, width=10).pack(side='left', padx=5)
        
        ttk.Label(quick_range_frame, text="To:").pack(side='left', padx=(10, 0))
        self.slice_end_var = tk.StringVar(value="1")
        ttk.Entry(quick_range_frame, textvariable=self.slice_end_var, width=10).pack(side='left', padx=5)
        
        # Output path
        ttk.Label(slice_frame, text="Output Location:").pack(pady=(10, 5))
        
        output_frame = ttk.Frame(slice_frame)
        output_frame.pack(fill='x', padx=20, pady=5)
        
        self.slice_output_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.slice_output_var, width=50).pack(side='left', fill='x', expand=True)
        ttk.Button(output_frame, text="Browse", command=self.browse_slice_output).pack(side='left', padx=(5, 0))
        
        # Slice button (larger and centered)
        slice_button = ttk.Button(slice_frame, text="✂️ SLICE PDF", command=self.slice_pdf)
        slice_button.pack(pady=20, ipadx=20, ipady=10)
        
        # Status
        self.slice_status = ttk.Label(slice_frame, text="", foreground="blue")
        self.slice_status.pack(pady=5)
        
    def create_merge_tab(self):
        """Create the PDF merging tab"""
        merge_frame = ttk.Frame(self.notebook)
        self.notebook.add(merge_frame, text="Merge PDFs")
        
        # File list
        ttk.Label(merge_frame, text="Select PDF Files to Merge:").pack(pady=(20, 5))
        
        list_frame = ttk.Frame(merge_frame)
        list_frame.pack(fill='both', expand=True, padx=20, pady=5)
        
        # Listbox with scrollbar
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')
        
        self.merge_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=10)
        self.merge_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.merge_listbox.yview)
        
        # Buttons
        button_frame = ttk.Frame(merge_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Add Files", command=self.add_merge_files).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Remove Selected", command=self.remove_merge_file).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Clear All", command=self.clear_merge_files).pack(side='left', padx=5)
        
        # Output path
        ttk.Label(merge_frame, text="Output File:").pack(pady=(10, 5))
        
        output_frame = ttk.Frame(merge_frame)
        output_frame.pack(fill='x', padx=20, pady=5)
        
        self.merge_output_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.merge_output_var, width=50).pack(side='left', fill='x', expand=True)
        ttk.Button(output_frame, text="Browse", command=self.browse_merge_output).pack(side='left', padx=(5, 0))
        
        # Merge button (larger and centered)
        merge_button = ttk.Button(merge_frame, text="📄 MERGE PDFs", command=self.merge_pdfs)
        merge_button.pack(pady=20, ipadx=20, ipady=10)
        
        # Status
        self.merge_status = ttk.Label(merge_frame, text="", foreground="blue")
        self.merge_status.pack(pady=5)
        
    def create_pptx_tab(self):
        """Create the PPTX to PDF conversion tab"""
        pptx_frame = ttk.Frame(self.notebook)
        self.notebook.add(pptx_frame, text="PPTX to PDF")
        
        # File list
        ttk.Label(pptx_frame, text="Select PPTX Files:").pack(pady=(20, 5))
        
        list_frame = ttk.Frame(pptx_frame)
        list_frame.pack(fill='both', expand=True, padx=20, pady=5)
        
        # Listbox with scrollbar
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')
        
        self.pptx_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=10)
        self.pptx_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.pptx_listbox.yview)
        
        # Buttons
        button_frame = ttk.Frame(pptx_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Add Files", command=self.add_pptx_files).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Remove Selected", command=self.remove_pptx_file).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Clear All", command=self.clear_pptx_files).pack(side='left', padx=5)
        
        # Output path
        ttk.Label(pptx_frame, text="Output PDF File:").pack(pady=(10, 5))
        
        output_frame = ttk.Frame(pptx_frame)
        output_frame.pack(fill='x', padx=20, pady=5)
        
        self.pptx_output_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.pptx_output_var, width=50).pack(side='left', fill='x', expand=True)
        ttk.Button(output_frame, text="Browse", command=self.browse_pptx_output).pack(side='left', padx=(5, 0))
        
        # Convert button (larger and centered)
        convert_button = ttk.Button(pptx_frame, text="🔄 CONVERT TO PDF", command=self.convert_pptx_to_pdf)
        convert_button.pack(pady=20, ipadx=20, ipady=10)
        
        # Status
        self.pptx_status = ttk.Label(pptx_frame, text="", foreground="blue")
        self.pptx_status.pack(pady=5)
        
    # Slice PDF methods
    def add_current_range(self):
        """Add the current page range to the ranges text widget"""
        try:
            start = int(self.slice_start_var.get())
            end = int(self.slice_end_var.get())
            
            if start < 1 or end < start:
                messagebox.showerror("Error", "Invalid page range")
                return
            
            # Add range to text widget
            current_text = self.slice_ranges_text.get("1.0", tk.END).strip()
            if current_text:
                self.slice_ranges_text.insert(tk.END, f"\n{start}-{end}")
            else:
                self.slice_ranges_text.insert(tk.END, f"{start}-{end}")
            
            # Increment for next range (optional convenience)
            self.slice_start_var.set(str(end + 1))
            self.slice_end_var.set(str(end + 1))
            
        except ValueError:
            messagebox.showerror("Error", "Please enter valid page numbers")
    
    def clear_ranges(self):
        """Clear all page ranges"""
        self.slice_ranges_text.delete("1.0", tk.END)
        self.slice_status.config(text="Ranges cleared", foreground="blue")
    
    def browse_slice_input(self):
        filename = filedialog.askopenfilename(
            title="Select PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=os.path.expanduser("~")
        )
        if filename:
            self.slice_input_var.set(filename)
            # Update page count
            try:
                reader = PdfReader(filename)
                num_pages = len(reader.pages)
                self.slice_end_var.set(str(num_pages))
                self.slice_status.config(text=f"PDF loaded: {num_pages} pages", foreground="green")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read PDF: {str(e)}")
                
    def browse_slice_output(self):
        # Get directory from input file if available, otherwise use home
        initial_dir = os.path.dirname(self.slice_input_var.get()) if self.slice_input_var.get() else os.path.expanduser("~")
        filename = filedialog.asksaveasfilename(
            title="Save sliced PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=initial_dir
        )
        if filename:
            self.slice_output_var.set(filename)
            
    def slice_pdf(self):
        input_file = self.slice_input_var.get()
        output_file = self.slice_output_var.get()
        
        if not input_file or not os.path.exists(input_file):
            messagebox.showerror("Error", "Please select a valid input PDF file")
            return
            
        if not output_file:
            messagebox.showerror("Error", "Please select an output file location")
            return
        
        # Get all ranges from the text widget
        ranges_text = self.slice_ranges_text.get("1.0", tk.END).strip()
        
        if not ranges_text:
            messagebox.showerror("Error", "Please add at least one page range")
            return
            
        try:
            reader = PdfReader(input_file)
            total_pages = len(reader.pages)
            writer = PdfWriter()
            
            # Parse and process each range
            ranges = []
            for line in ranges_text.split('\n'):
                line = line.strip()
                if not line:
                    continue
                
                # Parse range (e.g., "1-5" or "10-15")
                if '-' in line:
                    parts = line.split('-')
                    if len(parts) != 2:
                        messagebox.showerror("Error", f"Invalid range format: {line}\nUse format: start-end (e.g., 1-5)")
                        return
                    
                    start_page = int(parts[0].strip()) - 1  # Convert to 0-indexed
                    end_page = int(parts[1].strip()) - 1
                    
                    if start_page < 0 or end_page < start_page:
                        messagebox.showerror("Error", f"Invalid page range: {line}")
                        return
                    
                    if end_page >= total_pages:
                        messagebox.showerror("Error", f"Range {line} exceeds total pages ({total_pages})")
                        return
                    
                    ranges.append((start_page, end_page))
                else:
                    # Single page
                    page_num = int(line.strip()) - 1
                    if page_num < 0 or page_num >= total_pages:
                        messagebox.showerror("Error", f"Page {line} is out of range (1-{total_pages})")
                        return
                    ranges.append((page_num, page_num))
            
            if not ranges:
                messagebox.showerror("Error", "No valid page ranges found")
                return
            
            # Add pages from all ranges to the writer
            total_pages_added = 0
            for start_page, end_page in ranges:
                for page_num in range(start_page, end_page + 1):
                    writer.add_page(reader.pages[page_num])
                    total_pages_added += 1
            
            # Write the output file
            with open(output_file, 'wb') as output:
                writer.write(output)
            
            # Create summary message
            ranges_summary = ", ".join([f"{s+1}-{e+1}" if s != e else f"{s+1}" for s, e in ranges])
            self.slice_status.config(
                text=f"Success! Extracted {total_pages_added} pages from ranges: {ranges_summary}", 
                foreground="green"
            )
            messagebox.showinfo(
                "Success", 
                f"PDF sliced successfully!\n\nExtracted pages: {ranges_summary}\nTotal pages: {total_pages_added}\nSaved to: {output_file}"
            )
            
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid page number format: {str(e)}")
        except Exception as e:
            self.slice_status.config(text="Error occurred", foreground="red")
            messagebox.showerror("Error", f"Failed to slice PDF: {str(e)}")
            
    # Merge PDF methods
    def add_merge_files(self):
        # Get directory from last file in list if available
        initial_dir = os.path.expanduser("~")
        if self.merge_listbox.size() > 0:
            last_file = self.merge_listbox.get(self.merge_listbox.size() - 1)
            if os.path.exists(last_file):
                initial_dir = os.path.dirname(last_file)
        
        filenames = filedialog.askopenfilenames(
            title="Select PDF files to merge",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=initial_dir
        )
        for filename in filenames:
            self.merge_listbox.insert(tk.END, filename)
            
    def remove_merge_file(self):
        selection = self.merge_listbox.curselection()
        if selection:
            self.merge_listbox.delete(selection[0])
            
    def clear_merge_files(self):
        self.merge_listbox.delete(0, tk.END)
        
    def browse_merge_output(self):
        # Get directory from first file in list if available
        initial_dir = os.path.expanduser("~")
        if self.merge_listbox.size() > 0:
            first_file = self.merge_listbox.get(0)
            if os.path.exists(first_file):
                initial_dir = os.path.dirname(first_file)
        
        filename = filedialog.asksaveasfilename(
            title="Save merged PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=initial_dir
        )
        if filename:
            self.merge_output_var.set(filename)
            
    def merge_pdfs(self):
        output_file = self.merge_output_var.get()
        
        if self.merge_listbox.size() == 0:
            messagebox.showerror("Error", "Please add PDF files to merge")
            return
            
        if not output_file:
            messagebox.showerror("Error", "Please select an output file location")
            return
            
        try:
            merger = PdfMerger()
            
            for i in range(self.merge_listbox.size()):
                pdf_file = self.merge_listbox.get(i)
                if not os.path.exists(pdf_file):
                    messagebox.showwarning("Warning", f"File not found: {pdf_file}")
                    continue
                merger.append(pdf_file)
                
            merger.write(output_file)
            merger.close()
            
            self.merge_status.config(text=f"Success! Merged {self.merge_listbox.size()} files", foreground="green")
            messagebox.showinfo("Success", f"PDFs merged successfully!\nSaved to: {output_file}")
            
        except Exception as e:
            self.merge_status.config(text="Error occurred", foreground="red")
            messagebox.showerror("Error", f"Failed to merge PDFs: {str(e)}")
            
    # PPTX to PDF methods
    def add_pptx_files(self):
        # Get directory from last file in list if available
        initial_dir = os.path.expanduser("~")
        if self.pptx_listbox.size() > 0:
            last_file = self.pptx_listbox.get(self.pptx_listbox.size() - 1)
            if os.path.exists(last_file):
                initial_dir = os.path.dirname(last_file)
        
        filenames = filedialog.askopenfilenames(
            title="Select PPTX files",
            filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")],
            initialdir=initial_dir
        )
        for filename in filenames:
            self.pptx_listbox.insert(tk.END, filename)
            
    def remove_pptx_file(self):
        selection = self.pptx_listbox.curselection()
        if selection:
            self.pptx_listbox.delete(selection[0])
            
    def clear_pptx_files(self):
        self.pptx_listbox.delete(0, tk.END)
        
    def browse_pptx_output(self):
        # Get directory from first file in list if available
        initial_dir = os.path.expanduser("~")
        if self.pptx_listbox.size() > 0:
            first_file = self.pptx_listbox.get(0)
            if os.path.exists(first_file):
                initial_dir = os.path.dirname(first_file)
        
        filename = filedialog.asksaveasfilename(
            title="Save converted PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=initial_dir
        )
        if filename:
            self.pptx_output_var.set(filename)
            
    def convert_pptx_to_pdf(self):
        output_file = self.pptx_output_var.get()
        
        if self.pptx_listbox.size() == 0:
            messagebox.showerror("Error", "Please add PPTX files to convert")
            return
            
        if not output_file:
            messagebox.showerror("Error", "Please select an output file location")
            return
            
        try:
            import sys
            import platform
            
            # Check if on Windows for COM support
            if platform.system() == 'Windows':
                self.convert_pptx_windows(output_file)
            else:
                self.convert_pptx_alternative(output_file)
                
        except Exception as e:
            self.pptx_status.config(text="Error occurred", foreground="red")
            messagebox.showerror("Error", f"Failed to convert PPTX: {str(e)}")
            
    def convert_pptx_windows(self, output_file):
        """Convert PPTX to PDF on Windows using COM"""
        try:
            import comtypes.client
            
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            powerpoint.Visible = 1
            
            temp_pdfs = []
            
            for i in range(self.pptx_listbox.size()):
                pptx_file = self.pptx_listbox.get(i)
                if not os.path.exists(pptx_file):
                    messagebox.showwarning("Warning", f"File not found: {pptx_file}")
                    continue
                    
                temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf').name
                
                deck = powerpoint.Presentations.Open(os.path.abspath(pptx_file))
                deck.SaveAs(os.path.abspath(temp_pdf), 32)  # 32 = PDF format
                deck.Close()
                
                temp_pdfs.append(temp_pdf)
                
            powerpoint.Quit()
            
            # Merge all PDFs
            if len(temp_pdfs) > 0:
                merger = PdfMerger()
                for pdf in temp_pdfs:
                    merger.append(pdf)
                merger.write(output_file)
                merger.close()
                
                # Clean up temp files
                for pdf in temp_pdfs:
                    os.unlink(pdf)
                    
                self.pptx_status.config(text=f"Success! Converted {len(temp_pdfs)} files", foreground="green")
                messagebox.showinfo("Success", f"PPTX files converted successfully!\nSaved to: {output_file}")
            else:
                messagebox.showerror("Error", "No valid PPTX files to convert")
                
        except Exception as e:
            raise Exception(f"Windows conversion failed: {str(e)}")
            
    def convert_pptx_alternative(self, output_file):
        """Alternative PPTX to PDF conversion (uses LibreOffice if available)"""
        import subprocess
        
        # Check if LibreOffice is available
        libreoffice_cmds = ['libreoffice', 'soffice']
        libreoffice_cmd = None
        
        for cmd in libreoffice_cmds:
            try:
                subprocess.run([cmd, '--version'], capture_output=True, check=True)
                libreoffice_cmd = cmd
                break
            except (subprocess.CalledProcessError, FileNotFoundError):
                continue
                
        if not libreoffice_cmd:
            messagebox.showerror(
                "Error",
                "PPTX to PDF conversion requires LibreOffice on Linux/Mac.\n\n"
                "Please install LibreOffice:\n"
                "- Ubuntu/Debian: sudo apt-get install libreoffice\n"
                "- Mac: brew install libreoffice\n"
                "- Or download from: https://www.libreoffice.org/"
            )
            return
            
        try:
            temp_pdfs = []
            temp_dir = tempfile.mkdtemp()
            
            for i in range(self.pptx_listbox.size()):
                pptx_file = self.pptx_listbox.get(i)
                if not os.path.exists(pptx_file):
                    messagebox.showwarning("Warning", f"File not found: {pptx_file}")
                    continue
                    
                # Convert using LibreOffice
                result = subprocess.run(
                    [libreoffice_cmd, '--headless', '--convert-to', 'pdf', 
                     '--outdir', temp_dir, pptx_file],
                    capture_output=True,
                    text=True
                )
                
                if result.returncode != 0:
                    raise Exception(f"LibreOffice conversion failed: {result.stderr}")
                    
                # Find the generated PDF
                pptx_basename = os.path.splitext(os.path.basename(pptx_file))[0]
                temp_pdf = os.path.join(temp_dir, f"{pptx_basename}.pdf")
                
                if os.path.exists(temp_pdf):
                    temp_pdfs.append(temp_pdf)
                    
            # Merge all PDFs
            if len(temp_pdfs) > 0:
                merger = PdfMerger()
                for pdf in temp_pdfs:
                    merger.append(pdf)
                merger.write(output_file)
                merger.close()
                
                # Clean up
                import shutil
                shutil.rmtree(temp_dir)
                
                self.pptx_status.config(text=f"Success! Converted {len(temp_pdfs)} files", foreground="green")
                messagebox.showinfo("Success", f"PPTX files converted successfully!\nSaved to: {output_file}")
            else:
                messagebox.showerror("Error", "No valid PPTX files to convert")
                
        except Exception as e:
            raise Exception(f"Conversion failed: {str(e)}")


def main():
    root = tk.Tk()
    app = PDFManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
