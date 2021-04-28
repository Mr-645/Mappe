from datetime import datetime
from os import scandir

basepath = "C:/Users/drift/OneDrive/PROJECTS/Mappe/testing_ground/SOP_docs"

def convert_date(timestamp):
    d = datetime.utcfromtimestamp(timestamp)
    formated_date = d.strftime('%d %b %Y')
    return formated_date

def get_files():
    dir_entries = scandir(basepath)
    for entry in dir_entries:
        if entry.is_file():
            if entry.name.endswith(".docx"):
                info = entry.stat()
                print(f'{entry.name}\t Last Modified: {convert_date(info.st_mtime)}')

get_files()