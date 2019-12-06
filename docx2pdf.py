#! python3
# doc/docx 2 pdf - duplicate all docx & doc files from a folder to pdf in same folder
import sys, os, comtypes.client, time
# set word format, details here https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat
wdFormatPDF = 17
# set path to work
in_Path='c:\\temp\\doc2pdf'
os.chdir(in_Path)
# get all the DOCX filenames.
docxFiles = []
for filename in os.listdir(in_Path):
    if filename.endswith('.docx'):
        docxFiles.append(filename)
docxFiles.sort(key=str.lower)
# get all the DOC filenames.
docFiles = []
for filename in os.listdir(in_Path):
    if filename.endswith('.doc'):
        docFiles.append(filename)
docFiles.sort(key=str.lower)
# create word process/COM object
word = comtypes.client.CreateObject('Word.Application')
# make word invisible
word.Visible = False
# wait for the COM Server to prepare well.
time.sleep(2)
#Function convert to PDF, requires input/output files name only.
def doc2pdf(in_File,out_File):
    #mount full in/out path+file
    in_File_Full = in_Path + '\\' + in_File
    out_File_Full = in_Path + '\\' + out_File
    #convert file to pdf 
    # open file in word
    doc=word.Documents.Open(in_File_Full) 
    # save file as pdf
    doc.SaveAs(out_File_Full, FileFormat=wdFormatPDF)
    # close word file 
    doc.Close() 
# loop through all the DOCX files.
for in_File in docxFiles:
    out_File=in_File+'.pdf'
    print(in_File)
    print(out_File)
    doc2pdf(in_File,out_File)
# loop through all the DOC files.
for in_File in docFiles:
    out_File=in_File+'.pdf'
    print('Converting doc file: ' + in_File)
    doc2pdf(in_File,out_File)
# close Word Application 
word.Quit()
# print end
Print('All doc/docx files converted, bye')