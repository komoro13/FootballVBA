Attribute VB_Name = "Module7"
Function SgetLastJ(ByVal start As Long, ByVal Sheet_ As String) As Long
Dim j As Long

j = start
While ThisWorkbook.Sheets(Sheet_).range("D" & j).Value <> "" Or ThisWorkbook.Sheets(Sheet_).range("B" & j).Value <> ""
j = j + 1
Wend
SgetLastJ = j
End Function

