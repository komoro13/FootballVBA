Attribute VB_Name = "Module37"
Sub getFirstHalf(ByVal rec As Long)
Dim J As Long
Dim recs As Long

recs = SgetLastJ(2, "Novibet")

For J = 2 To recs


If Worksheets("Predictions").Range("E" & rec).Value = Worksheets("Novibet").Range("C" & J).Value And Worksheets("Predictions").Range("G" & rec).Value = Worksheets("Novibet").Range("D" & J).Value Then

If InStr(Worksheets("Novibet").Range("F" & J).Value, "+") > 0 Or Worksheets("Novibet").Range("F" & J).Value = "Interrupted" Or Worksheets("Novibet").Range("F" & J).Value = "Pen" Then
Exit Sub
End If


If TimeValue(Worksheets("Novibet").Range("F" & J).Value) <= TimeValue("45:00:00") Then
Worksheets("Predictions").Range("L" & rec).Value = Worksheets("Novibet").Range("E" & J).Value
Exit Sub
End If
Next J

End Sub
