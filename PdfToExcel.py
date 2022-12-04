from pathlib import Path
import win32com.client as win32com
#create an empty folder in desktop.
#The name of the folder is not important.
#If you get 000000000046x0x1x9 exception, 
#delete the folder and create new one with diffrent name.

gen_py_path = r'C:\Users\orish\Desktop\ABC' #<======== Change this
Path(gen_py_path).mkdir(parents=True, exist_ok=True)
win32com.__gen_path__ = gen_py_path

import subprocess
import time
from openpyxl import Workbook



macro_file_path = r"C:\Users\orish\Desktop\5.xlsx" #<======== Change this

pdf_path = r"C:\Users\orish\Desktop\5.pdf" #<======== Change this

docx_path = r"C:\Users\orish\Desktop\5.docx" #<======== Change this

powershell_script_text = (f'$filePath = "{pdf_path}"\n'\
                           f"    $wd = New-Object -ComObject Word.Application\n"\
                           f"$txt = $wd.Documents.Open($filePath,$false,$false,$false)\n"\
                           f'$wd.Documents[1].SaveAs("{docx_path}")\n'\
                           f"$wd.Documents[1].Close()\n"\
                           f'$wd.Quit()')

    
macro_script =(f'Sub WordToExcelWithFormatting()\n'\
               f"Dim Document, Word As Object\n"\
               f"Dim File As Variant\n"\
               f"Dim PG, Range\n"\
               f"Application.ScreenUpdating = False\n"\
               f'File = "{docx_path}"\n'\
               f"If File = False Then Exit Sub\n"\
               f'Set Word = CreateObject("Word.Application")\n'\
               f"Set Document = Word.Documents.Open(Filename:=File, ReadOnly:=True)\n"\
               f"Document.Activate\n"\
               f"PG = Document.Paragraphs.Count\n"\
               f"Set Range = Document.Range(Start:=Document.Paragraphs(1).Range.Start, _\n"\
               f"End:=Document.Paragraphs(PG).Range.End)\n"\
               f"Range.Select\n"\
               f"On Error Resume Next\n"\
               f"Word.Selection.Copy\n"\
               f'ActiveSheet.Range("B2").Select\n'\
               f"ActiveSheet.Paste\n"\
               f"Document.Close\n"\
               f"Word.Quit (wdDoNotSaveChanges)\n"\
               f"Application.ScreenUpdating = True\n"\
               f"End Sub")  

print("creating Word File")    
subprocess.run(['powershell.exe', '-Command', powershell_script_text], stdout=subprocess.PIPE)
time.sleep(2)
print("creating Excel File") 
wb = Workbook()
ws = wb.active
wb.save(macro_file_path)
time.sleep(3)

xlapp = win32com.Dispatch('Excel.Application')
xlapp.Visible = True 
xlapp.DisplayAlerts = False
xlwb = xlapp.Workbooks.Open(macro_file_path, False, False, None)

xlwb.VBProject.VBComponents.Add(1).CodeModule.AddFromString(macro_script.strip())
xlwb.Application.Run('Module1.WordToExcelWithFormatting')
xlwb.Save()
xlapp.Quit()
print("DONE") 
