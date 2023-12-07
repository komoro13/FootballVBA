Attribute VB_Name = "Module14"
Sub step_5(ByVal FileName As String, ByVal Row_ As Long)

Dim j As Long
Dim x As Long

x = 1

Worksheets("Step 5").Cells.ClearContents

For j = 2 To Worksheets(FileName).range("B2").End(xlDown)

If Worksheets(FileName).range("B" + CStr(j)).Value = "" Then
Exit For
End If

If Worksheets(FileName).range("AD" + CStr(j)).Value = Worksheets(FileName).range("AD" + CStr(Row_)).Value And Worksheets(FileName).range("AE" + CStr(j)).Value = Worksheets(FileName).range("AE" + CStr(Row_)).Value And Worksheets(FileName).range("AF" + CStr(j)).Value = Worksheets(FileName).range("AF" + CStr(Row_)).Value Then
Worksheets(FileName).range("A" & j).EntireRow.Copy ThisWorkbook.Sheets("Step 5").range("A" & CStr(x))
x = x + 1
End If


Next j
 
 
 
 'MsgBox Rows.Count
Sheets("Step 5").Select
range("A1:AX" & Rows.count).Sort key1:=range("AW1"), Order1:=xlAscending, Header:=xlNo
Sheets("Step 5").Select
'MsgBox "Copied!"
End Sub

