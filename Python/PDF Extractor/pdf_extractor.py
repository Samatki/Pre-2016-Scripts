import glob, os
from PyPDF2 import PdfFileWriter, PdfFileReader

directory = os.getcwd()
os.chdir(directory) 
for file in glob.glob("*.pdf"):
    fileop = file
    inputpdf = PdfFileReader(fileop,"rb")
  #  if inputpdf.isEncrypted:
   #     inputpdf.decrypt('')
    k = inputpdf.numPages
    if not k < 2 :
        for i in range(inputpdf.numPages):
            l = i+1
            output = PdfFileWriter()
            output.addPage(inputpdf.getPage(i))
            with open(file[:-4] + " page %s.pdf" % l, "wb") as outputStream:
                output.write(outputStream)