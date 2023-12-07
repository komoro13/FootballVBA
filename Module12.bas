Attribute VB_Name = "Module12"
Sub step_3(ByVal FileName As String, ByVal Row_ As Long)

Dim j As Long
Dim x As Long

x = 1

Worksheets("Step 3").Cells.ClearContents

For j = 2 To Worksheets(FileName).range("B2").End(xlDown)

If Worksheets(FileName).range("B" + CStr(j)).Value = "" Then
Exit For
End If

If Worksheets(FileName).range("AJ" + CStr(j)).Value = Worksheets(FileName).range("AJ" + CStr(Row_)).Value And Worksheets(FileName).range("AK" + CStr(j)).Value = Worksheets(FileName).range("AK" + CStr(Row_)).Value Then
Worksheets(FileName).range("A" & j).EntireRow.Copy ThisWorkbook.Sheets("Step 3").range("A" & CStr(x))
x = x + 1
End If


Next j
 
 
 
 'MsgBox Rows.count
 Sheets("Step 3").Select
range("A1:AX" & Rows.count).Sort key1:=range("AU1"), Order1:=xlDescending, Header:=xlNo
Sheets("Step 3").Select
'MsgBox "Copied!"
End Sub

