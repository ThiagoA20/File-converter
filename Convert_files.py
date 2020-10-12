import glob
#import win32com.client
import os
import img2pdf
import shutil
from tkinter import Tk
from pathlib import Path
from tkinter.filedialog import askdirectory, askopenfile, asksaveasfile
from docx2pdf import convert
from PIL import Image

inpt = None
outpt = None

def pdf_to_docx(arq):
    pass


def docx_to_pdf(arq):
    pass


def img_to_pdf(arq):
    save_file = str(askdirectory(initialdir="/home"))
    for i in arq:
        fullname = str(i)[0:-4] + '.pdf'
        with open(fullname,"wb") as f:
            f.write(img2pdf.convert(str(i)))
        shutil.move(fullname, save_file)


def type_convert(arq):
    global inpt
    global outpt
    list_files = []
    for fil in arq:
        if str(fil).endswith(str(inpt)):
            list_files.append(fil)
    if inpt == ".pdf" and outpt == "docx":
        pdf_to_docx(list_files)
    elif inpt == ".docx" and outpt == "pdf":
        docx_to_pdf(list_files)
    elif inpt == ".png" or inpt == ".jpg"  and outpt == "pdf":
        img_to_pdf(list_files)


def Convert(filelocal, values):
    global inpt
    global outpt
    inpt = "." + values[1].lower()
    outpt = values[2].lower()

    if filelocal == None:
        raise

    if values[0] == "File":
        path = [Path(f"{filelocal.name}")]
        type_convert(path)

    elif values[0] == "Directory":
        path = Path(f"{filelocal}")
        fil = []
        for arq in path.glob('*'):
            fil.append(arq) 
        type_convert(fil)

    else:
        path = Path(f"{filelocal}")
        def getFiles(dirName):
            listOfFile = os.listdir(dirName)
            completeFileList = list()
            for file in listOfFile:
                completePath = os.path.join(dirName, file)
                if os.path.isdir(completePath):
                    completeFileList = completeFileList + getFiles(completePath)
                else:
                    completeFileList.append(completePath)
            return completeFileList
        fil = getFiles(path)
        type_convert(fil)


def filelocal(arquivos):
    Tk().withdraw()
    try:
        if arquivos == "File":
            localfile = askopenfile(initialdir="/home")
            return localfile
        elif arquivos == "Directory" or arquivos == "Subdirectory":
            localfile = askdirectory(initialdir="/home")
            return localfile
    except:
        print("Error while catching files path!")
