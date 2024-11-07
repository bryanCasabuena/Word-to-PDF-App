import tkinter as tk 
from tkinter import messagebox
from functions.word_to_pdf import word_to_pdf
from converters.file_dialog import select_word_file, save_pdf_location

#Convert file to pdf
def convert():
    word_file = word_entry.get()
    pdf_file = pdf_entry.get()

    if not word_file or not pdf_file:
        messagebox.showerror('Error', 'Please select both input and output files')
        return
    
    try:
        word_to_pdf(word_file, pdf_file)
    except Exception as e:
        messagebox.showerror('Error', f'Error occured: {e}')

#CREATE THE MAIN WINDOW
root = tk.Tk()
root.title('Word to PDF Converter')
root.geometry('400x300')
root.configure(bg='black')

#WORD FILE SELECTION
tk.Label(root, text='Select word file:').pack(pady=5)
word_entry = tk.Entry(root, width=50)
word_entry.pack(padx=10,pady=5)
tk.Button(root, text='Browse...', command=lambda: select_word_file(word_entry)).pack(pady=5)

#PDF SAVE LOCATION
tk.Label(root, text='Save as PDF:').pack(pady=5)
pdf_entry = tk.Entry(root, width=50)
pdf_entry.pack(padx=10, pady=5)
tk.Button(root, text='Browse...',command=lambda: save_pdf_location(pdf_entry)).pack(pady=5)

#CONVERT BUTTON
tk.Button(root, text='Convert', bg='blue', fg='white', command=convert).pack(pady=20)

root.mainloop()