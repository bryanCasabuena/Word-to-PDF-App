import os
import win32com.client as win32
import pythoncom
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to convert Word to PDF
def word_to_pdf(word_file, pdf_file):
    """Convert Word document to PDF."""
    if not os.path.exists(word_file):
        messagebox.showerror("Error", "Word file not found. Check the file path.")
        return

    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Launch Word application
        word = win32.Dispatch('Word.Application')
        
        # Open the Word document
        doc = word.Documents.Open(word_file)
        
        # Save the document as PDF
        doc.SaveAs(pdf_file, FileFormat=17)
        
        # Close document and quit Word
        doc.Close()
        word.Quit()
        
        messagebox.showinfo("Success", f'PDF saved successfully as {pdf_file}')
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        print(e)
    finally:
        word.Quit()

# Function to select a Word file
def select_word_file():
    """Open file dialog to select Word file."""
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if file_path:
        word_entry.delete(0, tk.END)
        word_entry.insert(0, file_path)

# Function to select PDF save location
def select_pdf_location():
    """Open file dialog to select PDF save location."""
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, file_path)

# Function to perform the conversion
def convert():
    """Convert the selected Word file to PDF."""
    word_file = word_entry.get()
    pdf_file = pdf_entry.get()

    print(f'Trying to open word file: {word_file} to pdf: {pdf_entry}')

    if not os.path.exists(word_file):
        messagebox.showerror('Error')

    word_to_pdf(word_file, pdf_file)

# Create the main window
root = tk.Tk()
root.title("Word to PDF Converter")
root.geometry("400x300")
root.configure(bg='black')

# Word file selection
tk.Label(root, text="Select Word File:").pack(pady=5)
word_entry = tk.Entry(root, width=50)
word_entry.pack(padx=10, pady=5)
tk.Button(root, text="Browse...", command=select_word_file).pack(pady=5)

# PDF save location
tk.Label(root, text="Save as PDF:").pack(pady=5)
pdf_entry = tk.Entry(root, width=50)
pdf_entry.pack(padx=10, pady=5)
tk.Button(root, text="Save As...", command=select_pdf_location).pack(pady=5)

# Convert button
tk.Button(root, text="Convert to PDF", command=convert, bg="blue", fg="white").pack(pady=20)

# Run the application
root.mainloop()
