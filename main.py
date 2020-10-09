import glob
import win32com.client
import os
import img2pdf
from tkinter import *
from pathlib import Path
from tkinter.filedialog import askdirectory
from docx2pdf import convert


def run():
    window = Tk()
    window['bg'] = 'Dark Blue'

    frame1 = Frame(window)
    frame1.place(x=50, y=30, width=650, height=450)

    def pdf_to_docx(path):
        word = win32com.client.Dispatch("Word.Application")
        word.visible = 0

        pdfs_path = path  # folder where the .pdf files are stored
        for i, doc in enumerate(glob.iglob(pdfs_path + "*.pdf")):
            print(doc)
            filename = doc.split('\\')[-1]
            in_file = os.path.abspath(doc)
            print(in_file)
            wb = word.Documents.Open(in_file)
            out_file = os.path.abspath(pdfs_path + filename[0:-4] + ".docx".format(i))
            print("outfile\n", out_file)
            wb.SaveAs2(out_file, FileFormat=16)  # file format for docx
            print("success...")
            wb.Close()
        word.Quit()

    def docx_to_pdf(arq):
        convert(f"{arq}")

    def img_to_pdf(arq):
        with open(f"{str(arq)[0:-4]}.pdf", "wb") as f:
            f.write(img2pdf.convert(str(arq)))

    def Convert():
        path = Path(f"{local_lbl['text']}")
        inpt = clicked.get()
        outpt = clicked2.get()
        arquivos = clicked3.get()
        if local_lbl['text'] == "":
            error_lbl['text'] = "Please Select the directory"
        elif inpt == outpt:
            error_lbl['text'] = "Error, input file cant't be the same of the output file"
        else:
            error_lbl['text'] = ""
            print(inpt, outpt)

        if arquivos == "Archive":
            print(arquivos)

        elif arquivos == "Directory":
            for arq in path.glob('*'):
                if arq.suffix == inpt:
                    if inpt == ".pdf" and outpt == ".docx":
                        pdf_to_docx(path)
                    elif inpt == ".docx" and outpt == ".pdf":
                        docx_to_pdf(arq)
                    elif inpt == ".png" and outpt == ".pdf":
                        img_to_pdf(arq)
                    elif inpt == ".jpg" and outpt == ".pdf":
                        img_to_pdf(arq)

        else:
            print(arquivos)

    def filelocal():
        localfile = askdirectory()
        local_lbl['text'] = localfile

    options = [
        ".docx",
        ".pdf",
        ".png",
        ".jpg"
    ]

    qnt = [
        "Archive",
        "Directory",
        "Dir/Subdir"
    ]

    clicked = StringVar()
    clicked.set(options[0])
    clicked2 = StringVar()
    clicked2.set(options[1])
    clicked3 = StringVar()
    clicked3.set(qnt[0])
    r = IntVar()
    r.set("2")

    def clickedoverwrite(value):
        if value == 1:
            print("Overwrite")
            print(value)
        else:
            print("Don't Overwrite")
            print(value)

    # OptionMenu's
    archives_opt = OptionMenu(frame1, clicked3, *qnt)
    inputfile_opt = OptionMenu(frame1, clicked, *options)
    outputfile_opt = OptionMenu(frame1, clicked2, *options)

    # Labels
    logo_lbl = Label(frame1, text="|- - - - - - - - FILE CONVERTER- - - - - - - -|")
    from_lbl = Label(frame1, text="From: ")
    to_lbl = Label(frame1, text="to: ")
    overwrite_lbl = Label(frame1, text="Overwrite files? ")
    local_lbl = Label(frame1, text="", borderwidth=2, relief="ridge")
    error_lbl = Label(frame1, text="", fg="red")

    # Buttons
    selectdirectory_bt = Button(frame1, text="Select", bg='white', borderwidth=3, command=filelocal)
    converts_bt = Button(frame1, text='Convert', bg='white', borderwidth=5, command=Convert)
    overwriteTrue_bt = Radiobutton(frame1, text="Yes", value=1, variable=r, command=lambda: clickedoverwrite(r.get()))
    overwriteFalse_bt = Radiobutton(frame1, text="No", value=0, variable=r, command=lambda: clickedoverwrite(r.get()))

    # Position
    logo_lbl.place(x=205, y=50, height=100)
    from_lbl.place(x=170, y=200, width=50, height=40)
    to_lbl.place(x=320, y=200, width=40, height=40)
    overwrite_lbl.place(x=180, y=260)
    local_lbl.place(x=185, y=143, width=300, height=40)
    error_lbl.place(x=145, y=350, height=40, width=350)

    archives_opt.place(x=85, y=150)
    inputfile_opt.place(x=220, y=200, width=80, height=40)
    outputfile_opt.place(x=360, y=200, width=80, height=40)

    selectdirectory_bt.place(x=490, y=150, width=60, height=25)
    converts_bt.place(x=240, y=310, width=150, height=40)
    overwriteTrue_bt.place(x=340, y=260)
    overwriteFalse_bt.place(x=390, y=260)

    window.geometry("750x510+370+140")
    window.title("Conversor")
    window.mainloop()


if __name__ == '__main__':
    run()
