Attribute VB_Name = "Module1"
Sub CopyValues()

Dim myFile As String, path As String
Dim wsToCopy As Worksheet
Dim wsToPaste As Worksheet
Dim lCopyLastRow As Long

Dim rangeOne As Range

Dim erow As Long, col As Long

path = ThisWorkbook.path & "\Files\"
myFile = Dir(path & "*.xlsm")

Set wsToPaste = Workbooks("MakroForSliverDB.xlsm").Sheets("Master")

Application.ScreenUpdating = False

Do While myFile <> ""
Set wsToCopy = Workbooks.Open(path & myFile).Sheets("Informacje g³ówne- General Inf")

Set rangeOne = wsToCopy.Range("C7,C8,C6")

Windows("MakroForSliverDB.xlsm").Activate

erow = Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
col = 1
For Each cel In rangeOne
cel.Copy
Cells(erow, col).PasteSpecial xlPasteValues
col = col + 1
Next

Set wsToCopy = Workbooks.Open(path & myFile).Sheets("Sk³ad+param- Chem. comp+ param")

Set rangeOne = wsToCopy.Range("F2,C3,C14,A27,A27,C1")

Windows("MakroForSliverDB.xlsm").Activate

For Each cel In rangeOne
cel.Copy
If col = 6 Then
    If Left(cel.Value, 1) = "K" Then
        Cells(erow, col) = "DG"
    Else
        Cells(erow, col) = "Kr"
    End If
End If

If col = 6 Then
    col = col + 2
    If Mid(cel.Value, 8, 1) = "1" Then
        Cells(erow, col) = "Strand 1"
    Else
        Cells(erow, col) = "Strand 2"
    End If
    col = col - 1
End If

Cells(erow, col).PasteSpecial xlPasteValues

If col = 7 Then
    col = col + 2
Else
    col = col + 1
End If
Next

Set wsToCopy = Workbooks.Open(path & myFile).Sheets("Informacje g³ówne- General Inf")

Set rangeOne = wsToCopy.Range("C16")

Windows("MakroForSliverDB.xlsm").Activate

For Each cel In rangeOne
cel.Copy
Cells(erow, col).PasteSpecial xlPasteValues
col = col + 1
Next

Windows(myFile).Close savechanges:=False
myFile = Dir()
Loop

Range("A:K").EntireColumn.AutoFit

Application.ScreenUpdating = True

MsgBox "Task Complete!"

End Sub
