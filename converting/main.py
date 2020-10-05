from tkinter import *
from pathlib import Path
from tkinter.filedialog import askdirectory
import glob
import win32com.client
import os
from docx2pdf import convert


def run():
    window = Tk()
    window['bg'] = 'Dark Blue'

    frame1 = Frame(window)
    frame1.place(x=150, y=120, width=900, height=600) 
    
    logo = "|- - - - - - - - FILE CONVERSOR - - - - - - - - |"
    logolabel = Label(frame1, text=logo)
    logolabel.place(x=315, y=100, height=100)


    def pdf_to_docx():
        word = win32com.client.Dispatch("Word.Application")
        word.visible = 0

        pdfs_path = path # folder where the .pdf files are stored
        for i, doc in enumerate(glob.iglob(pdfs_path+"*.pdf")):
            print(doc)
            filename = doc.split('\\')[-1]
            in_file = os.path.abspath(doc)
            print(in_file)
            wb = word.Documents.Open(in_file)
            out_file = os.path.abspath(reqs_path +filename[0:-4]+ ".docx".format(i))
            print("outfile\n",out_file)
            wb.SaveAs2(out_file, FileFormat=16) # file format for docx
            print("success...")
            wb.Close()
        word.Quit()


    def docx_to_pdf(filename, path):
        convert(f"{filename}")
        convert(f"{filename}", f"""{filename - filemane.suffix() + ".pdf"}""")
        convert(f"{path}")


    def png_to_pdf():
        print("png to pdf")


    def jpg_to_pdf():
        print("jpg to pdf")


    def Convert():
        path = Path(f"{local['text']}")
        inpt = clicked.get()
        outpt = clicked2.get()
        if local['text'] == "":
            error['text'] = "Please Select the directory"
        elif inpt == outpt:
            error['text'] = "Error, input file cant't be the same of the output file"
        else:
            error['text'] = ""
            print(inpt, outpt)

        arquivos = clicked3.get()
        print(arquivos)

        if arquivos == "Current Directory" or arquivos == "Subdirectory":    
            for filename in path.glob('*'):
                if arquivos == "Current Directory":
                    if filename.suffix == inpt:
                        if inpt == ".pdf" and outpt == ".docx":
                            pdf_to_docx()
                        elif inpt == ".docx" and outpt == ".pdf":
                            docx_to_pdf(filename, path)
                        elif inpt == ".png" and outpt == ".pdf":
                            png_to_pdf()
                        elif inpt == ".jpg" and outpt == ".pdf":
                            jpg_to_pdf()
                else:
                    if filename.suffix == inpt:
                        print(filename)
    

    def filelocal():
        localfile = askdirectory()
        local['text'] = localfile

    options = [
            ".docx",
            ".pdf",
            ".png",
            ".jpg"
    ]
    
    qnt = [
            "Archive",
            "Current Directory",
            "Subdirectory"
    ]

    clicked = StringVar()
    clicked.set(options[0])
    clicked2 = StringVar()
    clicked2.set(options[1])
    clicked3 = StringVar()
    clicked3.set(qnt[0])

    # OptionMenu's
    quantidade = OptionMenu(frame1, clicked3, *qnt)
    quantidade.place(x=20, y=20)
    inputfile = OptionMenu(frame1, clicked, *options)
    inputfile.place(x=350, y=350, width=80, height=40)
    outputfile = OptionMenu(frame1, clicked2, *options)
    outputfile.place(x=490, y=350, width=80, height=40)

    # Labels
    from_ = Label(frame1, text="From: ")
    from_.place(x=300, y=350, width=50, height=40)
    to = Label(frame1, text="to: ")
    to.place(x=450, y=350, width=40, height=40)
    local = Label(frame1, text="", borderwidth=2, relief="ridge")
    local.place(x=315, y=243, width=300, height=40)
    directory = Label(frame1, text="Directory: ")
    directory.place(x=240, y=250)
    error = Label(frame1, text="")
    error.place(x=280, y=500, height=40, width=350)

    #Buttons
    selectdirectory = Button(frame1, text="Select", bg='white', borderwidth=3, command=filelocal)
    selectdirectory.place(x=620, y=250, width=60, height=25)
    converts = Button(frame1, text='Convert', bg='white', borderwidth=5, command=Convert)
    converts.place(x=370, y=450, width=150, height=40)

    window.geometry("1200x800+370+100")
    window.title("Conversor")
    window.mainloop()

if __name__ == '__main__':
    run()
