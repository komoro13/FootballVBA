Attribute VB_Name = "Module6"
Function getLastJ(ByVal cell_ As String, ByVal start As Long, ByVal Sheet_ As String) As Long
Dim j As Long

j = start
While ThisWorkbook.Sheets(Sheet_).range(cell_ & j).Value <> ""
j = j + 1
Wend
getLastJ = j
End Function
