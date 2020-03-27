# Helper function for first attempt to deal with .doc file format.
# More work needed to be done.

import subprocess
import os
import shutil

def convert_to_docx_and_give_path(DOCUMENT):
    if DOCUMENT.endswith('.docx'):
        return "./INPUTS/" + DOCUMENT
    elif DOCUMENT.endswith('.doc'):
        FILE_NAME = DOCUMENT.split("/")[-1].split(".")[0]
        
        shutil.copyfile("./inputs/" + DOCUMENT, "./tmp/" + DOCUMENT)
        
        CONVERT_CMD = '"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\wordconv.exe" -oice -nme "{}" "{}"'.format("./tmp/" + DOCUMENT,"./tmp/" + DOCUMENT + "x")
        subprocess.call(CONVERT_CMD, shell=True)
        
        return "./tmp/" + DOCUMENT + "x"
    else:
        print("Issue with {}".format(DOCUMENT))