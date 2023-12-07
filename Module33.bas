Attribute VB_Name = "Module33"
Private ch As Selenium.ChromeDriver
 Sub Novibet()
 
Dim data() As String

Set ch = New Selenium.ChromeDriver
Dim J As Long
Dim cf As String
Dim crlm As String


ch.start baseUrl:="https://www.novibet.gr"
ch.Get "/en/live-betting"
ch.Window.Maximize
DoEvents
Dim divs As Selenium.WebElements
Dim div As Selenium.WebElement
Set divs = ch.FindElementsByTag("div")
Dim teamB As Boolean
Dim score As Boolean
Dim scoreB As Boolean
Dim i As Long
teamB = False
scoreB = False

Worksheets("Novibet").Cells.ClearContents



DoEvents
Worksheets("Novibet").Range("Z2").Value = ch.FindElementByTag("app-in-play-events").Text
DoEvents
ch.Close

Worksheets("Novibet").Range("A1").Value = "Country"
Worksheets("Novibet").Range("B1").Value = "League"
Worksheets("Novibet").Range("C1").Value = "Team A"
Worksheets("Novibet").Range("D1").Value = "Team B"
Worksheets("Novibet").Range("E1").Value = "Score"
Worksheets("Novibet").Range("F1").Value = "Time"
Worksheets("Novibet").Range("G1").Value = "1"
Worksheets("Novibet").Range("H1").Value = "X"
Worksheets("Novibet").Range("I1").Value = "2"
Worksheets("Novibet").Range("J1").Value = "U"
Worksheets("Novibet").Range("K1").Value = "O"
Worksheets("Novibet").Range("L1").Value = "Ut"
Worksheets("Novibet").Range("M1").Value = "Ot"
Worksheets("Novibet").Range("N1").Value = "NG"
Worksheets("Novibet").Range("O1").Value = "GG"
Worksheets("Novibet").Range("P1").Value = "Next goal"


data = Split(Worksheets("Novibet").Range("Z2").Value, Chr(10))


J = 1

Dim mnav As Integer
mnav = 0
Dim s As Long
s = UBound(data, 1)

ReDim Preserve data(s + 20)


For i = 0 To s
mnav = 0

If (data(i) = "") Then
Exit For
End If

If (InStr(data(i), " - ") > 0) Then
cf = Split(data(i), " - ")(0)
crlm = Split(data(i), " - ")(1)

ElseIf InStr(data(i), ":") > 0 Then
Worksheets("Novibet").Range("F" & J).Value = data(i)

ElseIf data(i) = "Match interrupted" Or data(i) = "Match Interrupted" Then
Worksheets("Novibet").Range("F" & J).Value = "Interrupted"

ElseIf data(i) = "Pen" Then
Worksheets("Novibet").Range("F" & J).Value = "Pen"

ElseIf InStr(data(i), "+") = 1 Then
Worksheets("Novibet").Range("F" & J).Value = Worksheets("Novibet").Range("F" & J).Value & data(i)

ElseIf data(i) = "Full Time Result" Then

If data(i + 1) = "1" And data(i + 2) = "X" Then
Worksheets("Novibet").Range("G" & J).Value = "Locked"
Worksheets("Novibet").Range("H" & J).Value = "Locked"
Worksheets("Novibet").Range("I" & J).Value = "Locked"
mnav = mnav + 3


ElseIf data(i + 1) <> "Markets are not available" Then

Worksheets("Novibet").Range("G" & J).Value = data(i + 2)
Worksheets("Novibet").Range("H" & J).Value = data(i + 4)
Worksheets("Novibet").Range("I" & J).Value = data(i + 6)


ElseIf data(i + 1) = "Markets are not available" Then
Worksheets("Novibet").Range("G" & J).Value = "No bet"
Worksheets("Novibet").Range("H" & J).Value = "No bet"
Worksheets("Novibet").Range("I" & J).Value = "No bet"
mnav = mnav + 5
End If


If InStr(data(i + 8 - mnav), "O ") = 1 And InStr(data(i + 9 - mnav), "U ") = 1 Then
Worksheets("Novibet").Range("K" & J).Value = "Locked"
Worksheets("Novibet").Range("J" & J).Value = "Locked"
mnav = mnav + 2

ElseIf data(i + 8 - mnav) <> "Markets are not available" Then
Worksheets("Novibet").Range("M" & J).Value = Split(data(i + 8 - mnav), " ")(1)
Worksheets("Novibet").Range("L" & J).Value = Split(data(i + 10 - mnav), " ")(1)
Worksheets("Novibet").Range("K" & J).Value = data(i + 9 - mnav)
Worksheets("Novibet").Range("J" & J).Value = data(i + 11 - mnav)

ElseIf data(i + 8 - mnav) = "Markets are not available" Then
Worksheets("Novibet").Range("K" & J).Value = "No bet"
Worksheets("Novibet").Range("J" & J).Value = "No bet"
mnav = mnav + 3
End If


If data(i + 13 - mnav) = "GG" And data(i + 14 - mnav) = "NG" Then
Worksheets("Novibet").Range("N" & J).Value = "Locked"
Worksheets("Novibet").Range("O" & J).Value = "Locked"
mnav = mnav + 2

ElseIf data(i + 13 - mnav) <> "Markets are not available" Then

Worksheets("Novibet").Range("N" & J).Value = data(i + 14 - mnav)
 Worksheets("Novibet").Range("O" & J).Value = data(i + 16 - mnav)

ElseIf data(i + 13 - mnav) = "Markets are not available" Then  '----------------------OK-----------------------------------
Worksheets("Novibet").Range("O" & J).Value = "No bet"
Worksheets("Novibet").Range("N" & J).Value = "No bet"
mnav = mnav + 3
End If

i = i + 16 - mnav


ElseIf (IsNumeric(data(i)) Or data(i) = "Markets are not available") And (InStr(data(i - 1), ":") > 0 Or InStr(data(i - 1), "+") = 1 Or data(i - 1) = "Match Interrupted" Or data(i - 1) = "Pen") Or data(i - 1) = "Match Interrupted" Then

If data(i) = "1" And data(i + 1) = "X" Then
Worksheets("Novibet").Range("G" & J).Value = "Locked"
Worksheets("Novibet").Range("H" & J).Value = "Locked"
Worksheets("Novibet").Range("I" & J).Value = "Locked"
mnav = mnav + 3

ElseIf data(i) <> "Markets are not available" Then

Worksheets("Novibet").Range("G" & J).Value = data(i + 1)
Worksheets("Novibet").Range("H" & J).Value = data(i + 3)
Worksheets("Novibet").Range("I" & J).Value = data(i + 5)

ElseIf data(i) = "Markets are not available" Then
Worksheets("Novibet").Range("G" & J).Value = "No bet"
Worksheets("Novibet").Range("H" & J).Value = "No bet"
Worksheets("Novibet").Range("I" & J).Value = "No bet"
mnav = mnav + 5
End If

If InStr(data(i + 6 - mnav), "O ") = "1" And InStr(data(i + 7 - mnav), "U ") = 1 Then
Worksheets("Novibet").Range("K" & J).Value = "Locked"
Worksheets("Novibet").Range("J" & J).Value = "Locked"
mnav = mnav + 2

ElseIf data(i + 6 - mnav) <> "Markets are not available" Then

Worksheets("Novibet").Range("M" & J).Value = Split(data(i + 6 - mnav), " ")(1)
Worksheets("Novibet").Range("K" & J).Value = Split(data(i + 8 - mnav), " ")(1)
Worksheets("Novibet").Range("L" & J).Value = data(i + 7 - mnav)
Worksheets("Novibet").Range("J" & J).Value = data(i + 9 - mnav)

ElseIf data(i + 6 - mnav) = "Markets are not available" Then
Worksheets("Novibet").Range("K" & J).Value = "No bet"
Worksheets("Novibet").Range("J" & J).Value = "No bet"
mnav = mnav + 3
End If


If InStr(data(i + 10 - mnav), "GG") = 1 And InStr(data(i + 11 - mnav), "NG") = 1 Then
Worksheets("Novibet").Range("O" & J).Value = "Locked"
Worksheets("Novibet").Range("N" & J).Value = "Locked"
mnav = mnav + 2

ElseIf data(i + 10 - mnav) <> "Markets are not available" Then

Worksheets("Novibet").Range("O" & J).Value = data(i + 11 - mnav)
Worksheets("Novibet").Range("N" & J).Value = data(i + 13 - mnav)

ElseIf data(i + 10 - mnav) = "Markets are not available" Then '------------------------------------------OK--------------------------------
Worksheets("Novibet").Range("O" & J).Value = "No bet"
Worksheets("Novibet").Range("N" & J).Value = "No bet"
mnav = mnav + 3
End If
i = i + 13 - mnav

ElseIf IsNumeric(data(i)) = True And InStr(data(i), ".") = 0 Then
If scoreB = True Then
scoreB = False
Else
Worksheets("Novibet").Range("E" & J).Value = data(i) & "-" & data(i + 1)
scoreB = True
End If

ElseIf teamB = False Then

J = J + 1
Worksheets("Novibet").Range("C" & J).Value = data(i)
teamB = True
Else
Worksheets("Novibet").Range("D" & J).Value = data(i)
Worksheets("Novibet").Range("A" & J).Value = cf
Worksheets("Novibet").Range("B" & J).Value = crlm
teamB = False
End If


Next i
End Sub
