from tkinter import *
from tkinter import filedialog, messagebox, font
from win32com import client
import os
from os.path import exists
# from PIL import Image, ImageTk
import win32com

main = Tk()
main.geometry("710x300")
main.title('PDF Converter by Concepts')
# main.iconbitmap('icon.ico')
main.minsize(710, 300)
main.maxsize(710, 300)
# title_label = Label(text = 'PDF Converter').pack()

# photo = Image.open("far.png")
# photopng = ImageTk.PhotoImage(photo)
# photo_label = Label(image = photopng).pack()


def exl2pdf(file_location):
    app = client.DispatchEx("Excel.Application")
    app.Interactive = False
    app.Visible = False
    newpath = os.path.abspath(file_location)
    print(newpath)
    workbook = app.Workbooks.open(newpath)
    output = os.path.splitext(newpath)[0]
    workbook.ActiveSheet.ExportAsFixedFormat(0, output)
    workbook.Close()


def wrd2pdf(file_location):
    wrd = client.Dispatch("Word.Application")
    wrd.Visible = False
    newpath = os.path.abspath(file_location)
    print(newpath)
    doc = wrd.Documents.Open(newpath)

    output = os.path.splitext(newpath)[0]
    doc.SaveAs(output, FileFormat=17)
    doc.Close()

def convertwrd():
    filename = filedialog.askopenfilename(initialdir=".")
    infile = os.path.splitext(filename)
    firstname = infile[0]

    pdf = firstname + '.pdf'
    # word = firstname + '.docx'
    # print(word)
    if exists(pdf):
        messagebox.showinfo('Concepts.pk Says', 'File already Converter or File with same name Exists in selected Location')
    else:

        wrd2pdf(filename)
        messagebox.showinfo('Concepts.pk Says', 'PDF File Converted. check your selected folder')

def convertXL():
    filename = filedialog.askopenfilename(initialdir=".")
    infile = os.path.splitext(filename)
    firstname= infile[0]

    pdf = firstname + '.pdf'
    if exists(pdf):
        messagebox.showinfo('Concepts.pk Says', 'File already Converter or File with same name Exists in selected Location')
    else:
        # print('converting file')
        exl2pdf(filename)
        messagebox.showinfo('Concepts.pk Says', 'File Converted. check your selected folder')


Label(text ="PDF Converter", font = 'Perpetua 20 bold', fg = 'green', bg ='light grey' , width = '50', relief= SUNKEN).pack(side = 'top')

canvas = Canvas()
canvas.create_rectangle(1.5, 2, 280, 150)
canvas.place(x=80,y=40)

Bw=Button(main, text="Convert to PDF", command=convertwrd, bg='purple', fg='white', font='Times 12 bold')
Bw.place(x=220, y=150)

canvas1 = Canvas()
canvas1.create_rectangle(1.5, 2, 280, 150)
canvas1.place(x=380,y=40)

Be = Button(main, text ="Convert to PDF", command=convertXL, bg='maroon', fg='white', font='Times 12 bold')
Be.place(x=520, y=150)

Label(text="Note: Make Sure you have Saved the file & close opened file before clicking convert button", font='Arial 12 underline ', fg='red', width='70').place(x=50, y=200)

Label(text="Covert Following Documents to PDF", underline=True, font='Times 11 bold underline', fg='blue', width='28').place(x=90, y=45)

Label(text="MS Word Documents & Text Files", font='Arial 12 ', fg='purple', width='26').place(x=95, y=75)

Label(text="Web Pages(HTML, JS, PHP) Files", font='Arial 12  ', fg='purple', width='27').place(x=92, y=95)

Label(text="All Type of File Containing Text", font='Arial 12', fg='purple', width='25').place(x=92, y=117)

Label(text="Covert Following Data Files to PDF", font='Times 12  bold underline', fg='blue', width='28').place(x=392,y=45)

Label(text="MS Excel Files, CVS Files", font='Arial 12  ', fg='maroon', width='26').place(x=383, y=75)

Label(text="Web Pages(PHP, Java Scripts)", font='Arial 12  ', fg='maroon', width='26').place(x=392, y=95)

Label(text="All Type of Spread Sheets", font='Arial 12  ', fg='maroon', width='25').place(x=392, y=117)

Label(text="Concepts Coding--www.concepts.pk--Farhan--03009665776", font='Times 15 bold', fg='Teal', bg='light grey' , width='70', relief=SUNKEN).pack(side='bottom')

Label(text="For feedback or suggestions email us : ENGR.FARHAN.92@gmail.com", font='Arial 10 italic', fg='green', bg='light grey', width = '70', relief=SUNKEN).pack(side='bottom')

main.mainloop()
