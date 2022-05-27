import os,sys

import xlwings as xw # pip install xlwings
from docxtpl import DocxTemplate # pip install docxtpl

import win32com.client as win32 # pip install pywin32

import tkinter as tk # pip install tk
from tkinter import filedialog, Text

#make system path current directory
os.chdir(sys.path[0])

#functions for file execution
def main():
    wb = xw.Book(leasedata)
    sheet = wb.sheets["Sheet1"]
    doc = DocxTemplate(leasetemplate)
    context = sheet.range("A2").options(dict, expand="table", numbers=int).value
    print(context)
    print(os.path.basename(leasetemplate))
    docname = os.path.basename(leasetemplate)
    slicedname = docname[:-13]

    for key in context:
        if(context[key] == None):
            context[key] = ""

    output_name = slicedname  + context["street_address"] + ' - filled.docx'
    doc.render(context)
    doc.save(output_name)

    opentemplate.config(bg='#F0F0F0')

    #docdirectory = os.path.join(os.getcwd(), output_name)
    #convert_to_pdf(docdirectory)

#converts word document to pdf
#def convert_to_pdf(doc):
#    word = win32.DispatchEx("Word.Application")
#    new_name = doc.replace(".docx", r".pdf")
#    worddoc = word.Documents.Open(doc)
#    worddoc.SaveAs(new_name, FileFormat=17)
#    worddoc.Close()
#    return None

#functions for button filesaving
def addData():
    global leasedata 
    leasedata = filedialog.askopenfilename(initialdir="/", title = "Select File", filetypes=[("Excel files",".xlsx .xls")])
    if(leasedata != None):
        opendata.config(bg='green')
    return None

def addTemplate():
    global leasetemplate 
    leasetemplate = filedialog.askopenfilename(initialdir="/", title = "Select File", filetypes=[('Word files','.docx')])

    if(leasetemplate != None):
        opentemplate.config(bg='green')
    return None


#gui root window
root = tk.Tk()
root.title('Lease Assistant')
root.resizable(False,False)
root.geometry('300x125')

#data button
opendata = tk.Button(root, text="Open Data", padx = 5, pady = 5, command = addData)
opendata.pack()

#template button
opentemplate = tk.Button(root, text="Open Template", padx = 5, pady = 5, command = addTemplate)
opentemplate.pack()

#insertion button
executeinsert = tk.Button(root, text="Insert Data to Template", padx = 5, pady = 5, command = main)
executeinsert.pack()

root.mainloop()




