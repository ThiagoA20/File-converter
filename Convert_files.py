#import glob
#import win32com.client
import os
import img2pdf
from tkinter import Tk
from pathlib import Path
from tkinter.filedialog import askdirectory, askopenfile
from docx2pdf import convert


def docx_to_pdf(arq):
    convert(f"{arq}")

def img_to_pdf(arq):
    with open(f"{str(arq)[0:-4]}.pdf", "wb") as f:
        f.write(img2pdf.convert(str(arq)))

def Convert(filelocal, values): 
    inpt = "." + values[1].lower()
    outpt = values[2]

    if values[0] == "File":
        path = Path(f"{filelocal.name}")
        if path.suffix == inpt:
            try:
                if inpt == ".pdf" and outpt == "Docx":
                    print(".pdf --> .docx")
                elif inpt == ".docx" and outpt == "Pdf":
                    docx_to_pdf(path)
                elif inpt == ".png" or inpt == ".jpg" and outpt == "Pdf":
                    img_to_pdf(path)
            except:
                print("Erro!!!!!")
        else:
            print("Extensão de arquivo errada!")

    elif values[0] == "Directory":
        path = Path(f"{filelocal}")
        print(f"Directory = {path}")
        """
        elif values[0] == "Directory":
            for arq in path.glob('*'):
                if arq.suffix == inpt:
                    if inpt == "pdf" and outpt == "docx":
                        pdf_to_docx(path)
                    elif inpt == "docx" and outpt == "pdf":
                        print(arq)
                    elif inpt == "png" and outpt == "pdf":
                        print(arq)
                    elif inpt == "jpg" and outpt == "pdf":
                        print(arq)
        """

    else:
        path = Path(f"{filelocal}")
        print(f"Subdirectory = {path}")


def filelocal(arquivos):
    Tk().withdraw()
    if arquivos == "File":
        localfile = askopenfile()
        return localfile
    elif arquivos == "Directory" or value == "Subdirectory":
        localfile = askdirectory()
        return localfile
    else:
        print("Erro!")
