Attribute VB_Name = "Module22"
Sub step_13(ByVal FileName As String, ByVal Row_ As Long)

Dim j As Long
Dim x As Long

x = 1

Worksheets("Step 13").Cells.ClearContents

For j = 2 To Worksheets(FileName).range("B2").End(xlDown)

If Worksheets(FileName).range("B" + CStr(j)).Value = "" Then
Exit For
End If

If Worksheets(FileName).range("E" + CStr(j)).Value = Worksheets(FileName).range("E" + CStr(Row_)).Value Then
Worksheets(FileName).range("A" & j).EntireRow.Copy ThisWorkbook.Sheets("Step 13").range("A" & CStr(x))
x = x + 1
End If


Next j
 
 
 
 'MsgBox Rows.Count
Sheets("Step 13").Select
range("A1:AX" & Rows.count).Sort key1:=range("AW1"), Order1:=xlAscending, Header:=xlNo '
Sheets("Step 13").Select
'MsgBox "Copied!"
End Sub
