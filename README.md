Hy guy's..this is my new project on VBA Automation.
Here i am showing you that how to copy data from one source of file and paste it into another file..
here i upload two files with  1) MONTHLY ANALYSIS MAIN 
                             2) JUL 20201 SALES
in which i have created the VBA Code into  MONTHLY ANALYSIS MAIN  this file which has one sheet with the name "DATA"...so what the code is do delete the data from that sheet and paste the data.
from  JUL 20201 SALES this worbook and in data sheet...

one more important thing that u hace to put both files and in same folder..and change the path in Concole sheet which is in this workbook MONTHLY ANALYSIS MAIN. 

You  just access the code by using ALT+F11 by open the VBA Editor...And go through the code...

if you want only code so here it is......

Sub import_data()

Application.ScreenUpdating = False

Dim wsCons As Worksheet, wsData As Worksheet
Set wsCons = ThisWorkbook.Sheets("console")
Set wsData = ThisWorkbook.Sheets("Data")

Dim Folder As String, importfile As String
Folder = wsCons.Range("G1").Value
importfile = wsCons.Range("G2").Value

Dim fullfilename As String
fullfilename = Folder & "\" & importfile

wsData.Columns("A:E").ClearContents

Dim wb As Workbook
Set wb = Workbooks.Open(fullfilename)

Dim rngTOCopy As Range
Set rngTOCopy = wb.Sheets("Sheet1").Range("A1").CurrentRegion

rngTOCopy.Copy wsData.Range("A1")

wb.Close savechanges:=False

Dim lrowdata As Long
lrowdata = wsData.Range("A1").CurrentRegion.Rows.Count

wsData.Range("F2:H2").AutoFill wsData.Range("F2:H" & lrowdata)

Application.ScreenUpdating = True

End Sub

