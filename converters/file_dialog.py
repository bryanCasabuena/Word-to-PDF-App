import os
import tkinter as tk
from tkinter import filedialog, messagebox

def select_word_file(entry_widget):
    file_path = filedialog.askopenfilename(filetypes=[('Microsoft Word', '*.docx')])
    if file_path:
        print(f"Selected Word file: {file_path}")
        
        # Verify if the file exists
        if not os.path.exists(file_path):
            print(f"Error: The file at {file_path} doesn't exist.")
            messagebox.showerror("File Not Found", f"The file at {file_path} doesn't exist.")
            return
        
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

#Open dialog to select save PDF location
def save_pdf_location(entry_widget):
    file_path = filedialog.asksaveasfilename(defaultextension='.pdf', filetypes=[('PDF files', '*.pdf')])
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)