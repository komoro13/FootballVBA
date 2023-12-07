Attribute VB_Name = "Module34"
Sub mainLoop()

Dim J As Long
Dim rec As Long
Dim rec_ As Long
Dim leagues As Integer

leagues = SgetLastJ(2, "Config")
rec = SgetLastJ(1, "Halftime") - 1
rec_ = SgetLastJ(2, "Predictions") - 1

Dim updateTime1 As String
Dim updateTime1 As String
Dim updateTime1 As String
Dim updateTime1 As String

update Time = "19:00:00"

While True

If format(Now, "HH:MM") = "19:51" Then
update
End If


For J = 2 To leagues
Novibet
predictions_ (Worksheets("Config").Range("C" & J).Value)
Next J

Novibet

For J = 2 To rec_
getHalfTime rec_
Next J

For J = 1 To rec
predict Worksheets("Halftime").Range("C" & J).Value, J
Next J
Novibet

For J = 2 To rec_
checkValue (rec_)
Next J

Wend

End Sub
