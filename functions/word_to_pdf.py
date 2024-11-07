import os
import win32com.client as win32
import pythoncom
import tkinter as tk
from tkinter import messagebox

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