Attribute VB_Name = "Module4"
Sub addRanks(ByVal url As String, ByVal start As Long, ByVal fin As Long, ByVal Sheet_ As String)
Dim curRnd As Integer
Dim url1 As String
'Dim ws1 As Worksheet
'Set ws1 = ThisWorkbook.Sheets("Generic")
 

 
Dim url_ As String
url_ = url


Dim j As Long

ThisWorkbook.Sheets("Generic").Select

range("B3:Z50").NumberFormat = "@"



For j = start To fin


If ThisWorkbook.Sheets(Sheet_).range("AG" & CStr(j)).Value <> curRnd Then

curRnd = ThisWorkbook.Sheets(Sheet_).range("AG" & CStr(j)).Value
ActiveSheet.Cells.ClearContents

url1 = url & curRnd

useQueryTable url1, 4, ActiveSheet, 2
If InStr(range("C2").Value, "Hom") > 0 Then
ActiveSheet.Cells.ClearContents
useQueryTable url1, 5, ActiveSheet, 2
End If

End If

For i = 3 To 50
Worksheets("Generic").range("D" & CStr(i)).Value = Split(Worksheets("Generic").range("D" & CStr(i)).Value, "(")
If Worksheets("Generic").range("D" & CStr(i)).Value = ThisWorkbook.Sheets(Sheet_).range("E" & CStr(j)).Value Then
If ActiveSheet.range("B" & CStr(i)).Value = "" Then
ThisWorkbook.Sheets(Sheet_).range("AJ" & CStr(j)).Value = ActiveSheet.range("B" & CStr(i - 1)).Value + 1
End If
ThisWorkbook.Sheets(Sheet_).range("AJ" & CStr(j)).Value = ActiveSheet.range("B" & CStr(i)).Value
End If

If Worksheets("Generic").range("D" & CStr(i)).Value = ThisWorkbook.Sheets(Sheet_).range("G" & CStr(j)).Value Then
ThisWorkbook.Sheets(Sheet_).range("AK" & CStr(j)).Value = ActiveSheet.range("B" & CStr(i)).Value
End If


If Worksheets("Generic").range("D" & CStr(i)).Value = ThisWorkbook.Sheets(Sheet_).range("E" & CStr(j)).Value Then
ThisWorkbook.Sheets(Sheet_).range("AP" & CStr(j)).Value = ActiveSheet.range("E" & CStr(i)).Value
End If

If Worksheets("Generic").range("D" & CStr(i)).Value = ThisWorkbook.Sheets(Sheet_).range("G" & CStr(j)).Value Then
ThisWorkbook.Sheets(Sheet_).range("AQ" & CStr(j)).Value = ActiveSheet.range("E" & CStr(i)).Value
End If



Next i


Next j


ThisWorkbook.Sheets(Sheet_).Select

End Sub
