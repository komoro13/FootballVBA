Attribute VB_Name = "Module31"
Function predictions_(ByVal Sheetname As String)
Dim J As Long
Dim x As Integer
x = SgetLastJ(2, "Predictions")

Worksheets("Predictions").Cells.ClearContents

For J = 2 To SgetLastJ(2, Sheetname) - 1
If Worksheets(Sheetname).Range("J" & J).Value = "" Then
Worksheets(Sheetname).Range("A" & CStr(J) & ":N" & CStr(J)).Copy Worksheets("Predictions").Range("A" & x)
x = x + 1
End If
Next J


End Function
