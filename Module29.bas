Attribute VB_Name = "Module29"
Function checkscore(ByVal hscore As String, ByVal score As String) As Integer
Dim HHome As Integer
Dim HAway As Integer
Dim Home As Integer
Dim Away As Integer


HHome = CInt(Split(hscore, "-")(0))
HAway = CInt(Split(hscore, "-")(1))
Home = CInt(Split(score, "-")(0))
Away = CInt(Split(score, "-")(1))

checkscore = (Home + Away) - (HHome + HAway)
End Function
Sub predict(ByVal filename As String, ByVal row_ As Long, ByVal rec As Long)
Dim J As Long
Dim x As Long
Dim lastj As Long
Dim recs As Integer

Dim assoi As Integer
Dim xinaria As Integer
Dim dipla As Integer

Dim under As Integer
Dim over As Integer

Dim gg As Integer
Dim ng As Integer

Dim NextGoal As Integer

x = 1

Worksheets("Halftime").Cells.ClearContents
step_1 filename, row_
Worksheets("Step 1").Select
lastj = SgetLastJ(1, ActiveSheet.Name) - 1


For J = 1 To lastj
If ActiveSheet.Range("L" + CStr(J)).Value = Worksheets(filename).Range("L" + CStr(row_)).Value And ActiveSheet.Range("B" + CStr(J)).Value <> Worksheets(filename).Range("B" + CStr(row_)).Value Then
ActiveSheet.Range("A" & J).EntireRow.Copy Worksheets("Halftime").Range("A" & CStr(x))
x = x + 1
End If
Next J
If (x = 1) Then
MsgBox "No records"
Exit Sub
End If
Worksheets("Halftime").Select
lastj = SgetLastJ(1, ActiveSheet.Name) - 1


recs = CInt(lastj)

For J = 1 To lastj
If (ActiveSheet.Range("I" + CStr(J)).Value = "1") Then
    assoi = assoi + 1
ElseIf (ActiveSheet.Range("I" + CStr(J)).Value = "X") Then
    xinaria = xinaria + 1
ElseIf (ActiveSheet.Range("I" + CStr(J)).Value = "2") Then
    dipla = dipla + 1
End If
If (ActiveSheet.Range("M" + CStr(J)).Value = "Under") Then
    under = under + 1
ElseIf (ActiveSheet.Range("M" + CStr(J)).Value = "Over") Then
    over = over + 1
End If
If (ActiveSheet.Range("N" + CStr(J)).Value = "NG") Then
    ng = ng + 1
ElseIf (ActiveSheet.Range("N" + CStr(J)).Value = "G") Then
    gg = gg + 1
End If
If checkscore(ActiveSheet.Range("L" + CStr(J)).Value, ActiveSheet.Range("J" + CStr(J)).Value) > 0 Then
NextGoal = NextGoal + 1
End If
    
Next J
    
If ((assoi * 100) / recs > 80) Then
    Worksheets("Predictions").Range("I" & rec).Value = "1"
    Worksheets("Predictions").Range("U" & rec).Value = CDbl(recs / assoi)
ElseIf ((xinaria * 100) / recs > 80) Then
    MsgBox "X with " & (xinaria * 100) / recs
    Worksheets("Predictions").Range("I" & rec).Value = "X"
    Worksheets("Predictions").Range("U" & rec).Value = (recs / xinaria)
ElseIf ((dipla * 100) / recs > 80) Then
    MsgBox "2 with " & (dipla * 100) / recs
    Worksheets("Predictions").Range("1" & rec).Value = "2"
    Worksheets("Predictions").Range("U" & rec).Value = (recs / dipla)
End If
If ((under * 100) / recs > 80) Then
    MsgBox "Under with " & (under * 100) / recs
    Worksheets("Predictions").Range("M" & rec).Value = "Under"
    Worksheets("Predictions").Range("V" & rec).Value = (recs / under)
ElseIf ((over * 100) / recs > 80) Then
    MsgBox "Over with " & (over * 100) / recs
    Worksheets ("Predictions"), Range("M" & rec).Value = "Over"
    Worksheets("Predictions").Range("V" & rec).Value = (recs / over)
End If
If ((ng * 100) / recs > 80) Then
    MsgBox "NG with " & (under * 100) / recs
    Worksheets("Predictions").Range("N" & rec).Value = "NG"
    Worksheets("Predictions").Range("W" & rec).Value = (recs / ng)
ElseIf ((gg * 100) / recs > 80) Then
    MsgBox "GG with " & (over * 100) / recs
    Worksheets("Predictions").Range("N" & rec).Value = "GG"
    Worksheets("Predictions").Range("W" & rec).Value = (recs / gg)
End If

If (assoi = 0) And Not ((xinaria * 100) / recs > 80) And Not ((dipla * 100) / recs > 80) Then
    Worksheets("Predictions").Range("É" & rec).Value = "X/2"
End If
If (xinaria = 0) And Not ((assoi * 100) / recs > 80) And Not ((dipla * 100) / recs > 80) Then
    Worksheets("Predictions").Range("É" & rec).Value = "1/2"
End If
If (dipla = 0) And Not ((assoi * 100) / recs > 80) And Not ((xinaria * 100) / recs > 80) Then
    Worksheets("Predictions").Range("I" & rec).Value = "1/X"
End If
If ((NextGoal * 100) / recs > 80) Then
Worksheets("Predictions").Range("P" & rec).Value = "Yes"
Worksheets("Predictions").Range("X" & rec).Value = (recs / NextGoal)
Else
Worksheets("Predictions").Range("P" & rec).Value = "No"
End If








End Sub
