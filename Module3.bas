Attribute VB_Name = "Module2"
Sub CopyValuesFromHeat()

Dim myFile As String, path As String
Dim wsToCopy As Worksheet
Dim wsToPaste As Worksheet
Dim lCopyLastRow As Long

Dim rangeOne As Range

Dim erow As Long, col As Long

path = ThisWorkbook.path & "\Files\Heats\"
myFile = Dir(path & "*.xls")

Set wsToPaste = Workbooks("MacroChemistry.xlsm").Sheets("Master")

Application.ScreenUpdating = False

Do While myFile <> ""
Set wsToCopy = Workbooks.Open(path & myFile).Sheets("L3_SAP_Analysis - Wytopy")

Set rangeOne = wsToCopy.Range("F3,G3,H3,J3,K3,L3,N3,M3,S3,W3,O3,AA3,P3,R3,Q3,D3,Z3")

Windows("MacroChemistry.xlsm").Activate

erow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
col = 1
For Each cel In rangeOne
cel.Copy
Cells(erow, col).PasteSpecial xlPasteValues
col = col + 1
Next

Windows(myFile).Close savechanges:=False
myFile = Dir()
Loop

Range("A:Q").EntireColumn.AutoFit

Application.ScreenUpdating = True

MsgBox "Task Complete!"

End Sub
