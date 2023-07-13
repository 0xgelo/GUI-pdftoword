from tkinter import * 
from tkinter import filedialog
from tkinter.filedialog import askopenfile
from tkinter.filedialog import asksaveasfile
from tkinter import ttk 
from pdf2docx import Converter

PINK = "#e2979c"
RED = "#e7305b"
GREEN = "#9bdeac"
YELLOW = "#f7f5dd"
FONT_NAME = "Courier"
YELLOW = "#f7f5dd" 

window = Tk() 
window.title("PDF to WORD CONVERTER")
window.config(padx=100, pady=50, bg='white')

pdf_to_word = PhotoImage(file="new2.png")
canvas = Canvas(width=400, height=224,  bg= 'white', highlightthickness=0)
canvas.create_image(200, 112, image=pdf_to_word)
canvas.grid(row=1, column=1)

title_label = Label(text="PDF to WORD Converter", font=(FONT_NAME, 35, 'bold'), fg='black', bg= 'white')
title_label.grid(row=0, column=1)

upload_button = PhotoImage(file= 'button100.png')
b1 = Button(text='Upload File',command = lambda:upload_file(), image=upload_button, borderwidth=0)
b1.grid(row=2,column=1) 


path_name =StringVar()
path_label = Label(textvariable=path_name,fg=RED, bg='white')
path_label.grid(row=3,column=1) 
path_name.set("Upload a file to see the path")

convert_progress = Label(text="", font=(FONT_NAME, 20, 'bold'), fg ='black', bg= 'white')
convert_progress.grid(row=4,column=1)

def upload_file():
    file = filedialog.askopenfilename()
    if(file):
        path_name.set(file)
        convert_to_word(file)
        
def convert_to_word(file):
    doc_name = file.split('.')[0]
    pdf_file = file
    docx_file = f'{doc_name}.docx'
    
    try:
    # Converting PDF to Docx
        cv_obj = Converter(pdf_file)
        cv_obj.convert(docx_file)
        cv_obj.close()
    except:
        convert_progress.config(text='Failed to convert PDF', fg=RED)
    
    else:
        convert_progress.config(text='Successfully converted to document', fg=GREEN)
        

window.mainloop()