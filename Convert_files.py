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

"file types"
inpt = None
outpt = None

"these functions before type_convert function will receive an list with posixpath of all files that will be converted, convert they, rename and move to a specifyed local by the user"
def pdf_to_docx(arq):
    pass

def pdf_to_png(arq):
    pass

def pdf_to_jpg(arq):
    pass

def pdf_to_Json(arq):
    pass

def pdf_to_Csv(arq):
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

def json_to_pdf(arq):
    pass

def csv_to_pdf(arq):
    pass

def json_to_csv(arq):
    pass

def csv_to_json(arq):
    pass

"a type_convert function will identify the types of files with which the user will work, filter and put in a list only those files that end with the extension in the variable inpt, then it will pass this list to the respective function defined by the variables inpt and outpt"
def type_convert(arq):
    global inpt
    global outpt
    list_files = []
    for fil in arq:
        if str(fil).endswith(str(inpt)):
            list_files.append(fil)
    if inpt == ".pdf" and outpt == "docx":
        pdf_to_docx(list_files)
    elif inpt == ".pdf" and outpt == "png":
        pdf_to_png(list_files)
    elif inpt == ".pdf" and outpt == "jpg":
        pdf_to_jpg(list_files)
    elif inpt == ".pdf" and outpt == "json":
        pdf_to_json(list_files)
    elif inpt == ".pdf" and outpt == "csv":
        pdf_to_csv(list_files)
    elif inpt == ".docx" and outpt == "pdf":
        docx_to_pdf(list_files)
    elif inpt == ".png" or inpt == ".jpg"  and outpt == "pdf":
        img_to_pdf(list_files)
    elif inpt == ".json" and outpt == "pdf":
        json_to_pdf(list_files)
    elif inpt == ".csv" and outpt == "pdf":
        csv_to_pdf(list_files)
    elif inpt == ".json" and outpt == "csv":
        json_to_csv(list_files)
    elif inpt == ".csv" and outpt == "json":
        csv_to_json(list_files)
    else:
        print("invalid conversion!")

"receives from the convert function of main.py the location where the files are located and the parameters passed by the user to check how many files he wants to convert and generates the variables inpt and outpt to define which and to what type of file to convert"
def Convert(filelocal, values):
    global inpt
    global outpt
    inpt = "." + values[1].lower()
    outpt = values[2].lower()

    "If user don't specify the file local, this will raise an error to main.py, that will show a mensage asking for the file local"
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
        "This function will generate a list of all files in the current directory and subdirectories that will be passed to the type_convert function"
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
