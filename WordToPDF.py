import win32com.client
from pathlib import Path 
import os

docs_path = Path.cwd()
docs_files = os.walk(docs_path)

for dirpath, dirnames, filenames in docs_files:
    for filename in filenames:
        if filename.endswith("docx"):
            pdfFilename = filename[0:-4] + "pdf"

            wdFormatPDF = 17 # Word's numeric code for PDFs.
            wordObj = win32com.client.Dispatch('Word.Application')

            word_path = Path(dirpath) / Path(str(filename))
            pdf_path = Path(dirpath) / Path(pdfFilename)

            #For some reason the code wont work after closing certain documents in the directory
            try:
                docObj = wordObj.Documents.Open(str(word_path))

                docObj.SaveAs(str(pdf_path), FileFormat=wdFormatPDF)
                docObj.Close()
                wordObj.Quit()
            except:
                continue
