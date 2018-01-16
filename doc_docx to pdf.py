import sys
import os
import comtypes.client
import glob

wdFormatPDF = 17

#以下のディレクトリを変更してください．下は参考例
path_docx = "C:\\Users\\User\\Desktop\\*.docx"
path_doc = "C:\\Users\\User\\Desktop\\*.doc"

files_docx = []
files_doc = []
files_docx = glob.glob(path_docx)
files_doc = glob.glob(path_doc)

for afile in files_docx:
#一時ファイル~$が検出される場合があるのでそれを回避
    if "~$" in afile:
        pass
    else:
        print(afile)
        in_file = afile
        out_file = afile.replace(".docx",".pdf")
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        
for afile in files_doc:
    if "~$" in afile:
        pass
    else:
        print(afile)
        in_file = afile
        out_file = afile.replace(".doc",".pdf")
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()

#参考URL:https://stackoverflow.com/questions/6011115/doc-to-pdf-using-python
