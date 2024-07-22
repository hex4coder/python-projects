# program untuk generate report qr code siswa dan guru
import os
from docx import Document
from docx.shared import Inches

folder_guru = "./qr-guru"
folder_siswa = "./qr-siswa"


# generate tables for list data



# parsing files in dir
def read_files(directory):
    # baca file dalam folder tersebut
    files = os.listdir(directory)
    list_data = []

    if len(files) < 1:
        # no files
        return None
    

    for file in files:
        abs_file = directory + "/" + file
        list_data.append(abs_file)


    return list_data

def read_file_guru():
    qr_guru = read_files(folder_guru)
    for qr in qr_guru:
        print(qr)




read_file_guru()