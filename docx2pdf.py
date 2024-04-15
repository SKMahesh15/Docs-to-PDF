import os
from pypdf import PdfWriter
import glob
import subprocess
import aspose.words as aw
from tkinter import *


def word_to_pdf_merger():

    input_dir = str(txt_input_word.get())
    output_dir = str(txt_input_pdf.get())

    input_dir = os.path.join(input_dir)
    output_dir = os.path.join(output_dir)
    output_dir = str(output_dir)
    
    filepaths_docx = glob.glob(os.path.join(input_dir, "*"))
    filepaths_pdf = glob.glob(os.path.join(output_dir, "*"))

    for filepath in filepaths_docx:
        if filepath.endswith('.docx'):
            doc = aw.Document(filepath)
            file_name = str(os.path.basename(filepath)).split('.')[0] 
            saving_files = output_dir + '\\' + file_name + ".pdf"
            doc.save(saving_files)
            
    merger = PdfWriter()
    for filepath in filepaths_pdf:
        if filepath.endswith(".pdf"):
            merger.append(filepath)

    merged_pdf_path = os.path.join(output_dir, "merged-pdf.pdf")
    merger.write(merged_pdf_path)
    merger.close()


# main window
window = Tk()

text_doc = StringVar()
label = Label(window, text="Enter Word Documents Folder: ", font=('bold', 13), pady=20)
label.grid(row=0, column=0, sticky=W)
txt_input_word = Entry(window, textvariable=text_doc, width=60)
txt_input_word.grid(row=0, column=1)

text_pdf = StringVar()
label = Label(window, text="Enter PDF Documents Folder: ", font=('bold', 13))
label.grid(row=1, column=0, sticky=W)
txt_input_pdf = Entry(window, textvariable=text_pdf, width=60)
txt_input_pdf.grid(row=1, column=1)


bttn_to_cnvrt = Button(window, text="Convert and merge", command=word_to_pdf_merger)
bttn_to_cnvrt.grid(row=3, column=1)

window.title("Docs to PDF converter and merger")
window.geometry("700x450")

window.mainloop()
