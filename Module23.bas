Attribute VB_Name = "Module23"
Sub step_14(ByVal FileName As String, ByVal Row_ As Long)

Dim j As Long
Dim x As Long

x = 1

Worksheets("Step 14").Cells.ClearContents

For j = 2 To Worksheets(FileName).range("B2").End(xlDown)

If Worksheets(FileName).range("B" + CStr(j)).Value = "" Then
Exit For
End If

If Worksheets(FileName).range("E" + CStr(j)).Value = Worksheets(FileName).range("E" + CStr(Row_)).Value And Worksheets(FileName).range("AJ" + CStr(j)).Value = Worksheets(FileName).range("AJ" + CStr(Row_)).Value Then
Worksheets(FileName).range("A" & j).EntireRow.Copy ThisWorkbook.Sheets("Step 14").range("A" & CStr(x))
x = x + 1
End If


Next j
 
 
 
 'MsgBox Rows.Count
Sheets("Step 14").Select
range("A1:AX" & Rows.count).Sort key1:=range("AK1"), Order1:=xlAscending, Header:=xlNo '
Sheets("Step 14").Select
'MsgBox "Copied!"
End Sub

