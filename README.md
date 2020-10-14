# File-converter

converter pdf, docx, png, jpg, csv files

Program with a graphical interface made in kivy, you select if you want to convert just one file or all files in a folder with subdirectories or
no, then select the respective directory where the files are and finally which and to what format you want to convert.

The file widgets.kv will create the user interface and send parameters that user provides to main.py, main.py manages the screens exange and send the parameters to respective packages. Also Convert_files.py will identify the parameters and apply the correct function to convert the files
