import glob
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
    path = Path(f"{filelocal.name}")
    values = values
    arquivos = values[0]
    inpt = values[1]
    outpt = values[2]

    if arquivos == "File":
        if inpt == "Pdf" and outpt == "Docx":
            print(f"Arquivo = {path}, input = {inpt}, output = {outpt}")
        elif inpt == "Docx" and outpt == "Pdf":
            print(f"Arquivo = {path}, input = {inpt}, output = {outpt}")
        elif inpt == "Png" and outpt == "Pdf":
            print(f"Arquivo = {path}, input = {inpt}, output = {outpt}")
        elif inpt == "Jpg" and outpt == "Pdf":
            print(f"Arquivo = {path}, input = {inpt}, output = {outpt}")

    elif arquivos == "Directory":
        print("Directory")

    else:
        print("Subdirectory")

    """
    elif arquivos == "Directory":
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


def filelocal():
    Tk().withdraw()
    localfile = askopenfile()
    return localfile
