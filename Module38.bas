Attribute VB_Name = "Module38"
Function deleteCurrentSeason(ByVal Sheetname As String)
Dim J As Long
Dim lmatch

lmatch = SgetLastJ(2, Sheetname)

J = 2
While True
If CInt(Worksheets(Sheetname).Range("AG" & J).Value) > CInt(Worksheets(Sheetname).Range("AG" & CStr(J + 1)).Value) Then
Exit Function
End If
Worksheets(Sheetname).Range("B" & J).EntireRow.Delete
End Function
Sub update()

Dim J As Long
Dim leagues As Long

leagues = SgetLastJ(2, "Config")

For J = 2 To leagues
deleteCurrentSeason (Worksheets("Config").Range("D" & J).Value)
run Worksheets("Config").Range("D" & J).Value, Worksheets("Config").Range("D" & J).Value, Worksheets("Config").Range("B" & J).Value, link_2, Worksheets("Config").Range("C" & J).Value, 2
End Sub
