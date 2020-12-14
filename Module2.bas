Attribute VB_Name = "Module1"
Sub CopyValues()

Dim myFile As String, path As String
Dim wsToCopy As Worksheet
Dim wsToPaste As Worksheet
Dim lCopyLastRow As Long

Dim erow As Long, col As Long

Dim MyRange As Range
Dim RowSelect As Range
Dim i As Integer

path = ThisWorkbook.path & "\Files\"
myFile = Dir(path & "*.xlsm")

Set wsToPaste = Workbooks("MakroChemia.xlsm").Sheets("Master")

Application.ScreenUpdating = False

Do While myFile <> ""
Set wsToCopy = Workbooks.Open(path & myFile).Sheets("Sheet1")

Set MyRange = wsToCopy.Range("A1").CurrentRegion
Set RowSelect = MyRange.Rows(22)
For i = 22 To MyRange.Rows.Count Step 2
Set RowSelect = Union(RowSelect, MyRange.Rows(i))
Next i

Windows("MakroChemia.xlsm").Activate

erow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
col = 1
For Each cel In RowSelect
cel.Copy
Cells(erow, col).PasteSpecial xlPasteValues
col = col + 1
If col = 18 Then
col = 1
erow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
End If
Next

Windows(myFile).Close savechanges:=False
myFile = Dir()
Loop

Range("A:R").EntireColumn.AutoFit

Application.ScreenUpdating = True

MsgBox "Task Complete!"

End Sub
