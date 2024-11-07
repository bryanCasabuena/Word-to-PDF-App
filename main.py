import os
import win32com.client as win32
import pythoncom
import tkinter as tk
from tkinter import filedialog, messagebox

def word_to_pdf(word_file, pdf_file):
    # Check if the file exists
    if not os.path.exists(word_file):
        print("Error: Word file not found.")
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
        
        messagebox.showinfo('Success', f'PDF saved successfully AS {pdf_file}')
    except Exception as e:
        messagebox.showerror('Error', f'Error occured: {e}')
    finally:
        word.Quit()

#Open dialog to select the word file
def select_word_file():
    file_path = filedialog.askopenfilename(filetypes=[('Microsoft Word', '*.docx')])
    if file_path:
        print(f"Selected Word file: {file_path}")
        
        # Verify if the file exists
        if not os.path.exists(file_path):
            print(f"Error: The file at {file_path} doesn't exist.")
            messagebox.showerror("File Not Found", f"The file at {file_path} doesn't exist.")
            return
        
        word_entry.delete(0, tk.END)
        word_entry.insert(0, file_path)

#Open dialog to select save PDF location
def save_pdf_location():
    file_path = filedialog.asksaveasfilename(defaultextension='.pdf', filetypes=[('PDF files', '*.pdf')])
    if file_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, file_path)

#Convert file to pdf
def convert():
    word_file = word_entry.get()
    pdf_file = pdf_entry.get()
    word_to_pdf(word_file, pdf_file)
    

#CREATE THE MAIN WINDOW
root = tk.Tk()
root.title('Word to PDF Converter')
root.geometry('400x300')

#WORD FILE SELECTION
tk.Label(root, text='Select word file:').pack(pady=5)
word_entry = tk.Entry(root, width=50)
word_entry.pack(padx=10,pady=5)
tk.Button(root, text='Browse...', command=select_word_file).pack(pady=5)

#PDF SAVE LOCATION
tk.Label(root, text='Save as PDF:').pack(pady=5)
pdf_entry = tk.Entry(root, width=50)
pdf_entry.pack(padx=10, pady=5)
tk.Button(root, text='Browse...',command=save_pdf_location).pack(pady=5)

#CONVERT BUTTON
tk.Button(root, text='Convert', bg='blue', fg='white', command=convert).pack(pady=20)

root.mainloop()