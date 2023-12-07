Attribute VB_Name = "Module2"
Sub format(ByVal url As String, ByVal start As Long, ByVal Sheet_ As String)
Attribute format.VB_ProcData.VB_Invoke_Func = "q\n14"


Dim j As Long
Dim i As Long




Dim BRowData As String
Dim ABRowData As String
Dim splArr(2) As String

Dim tempE As String
Dim tempD As String
Dim tempF As String

Dim curRnd As Integer

Dim lastj As Long

ActiveSheet.Hyperlinks.Delete



lastj = SgetLastJ(2, Sheet_)

range("C2:AM" & CStr(lastj - 1)).NumberFormat = "@"

range("A2:BB" & CStr(lastj - 1)).Font.Bold = True

With Worksheets(Sheet_).Columns("E")
 .ColumnWidth = 25
End With

With Worksheets(Sheet_).Columns("G")
 .ColumnWidth = 25
End With

With Worksheets(Sheet_).Columns("J")
 .ColumnWidth = 15
 End With

With Worksheets(Sheet_).Columns("F")
 .ColumnWidth = 15
 End With



For j = start To lastj


If range("B" & CStr(j)).Value = "" And range("D" & CStr(j)).Value = "" Then
Exit For
End If

If Right(range("B" & CStr(j)).Value, 5) = "Round" Then
j = j + 1

range("AG" & j).Value = Split(range("B" & CStr(j - 1)).Value, ".")(0)
range("B" & CStr(j - 1)).EntireRow.Delete
End If
Next j

For j = start To lastj - 1
If range("B" & CStr(j)).Value = "" And range("D" & CStr(j)).Value = "" Then
Exit For
End If
 range("E" & CStr(j)).Value = range("D" & CStr(j)).Value


tempG = range("G" & CStr(j)).Value
range("G" & CStr(j)).Value = tempE

tempF = range("F" & CStr(j)).Value
range("F" & CStr(j)).Value = "-"

range("J" & CStr(j)).Value = tempG
range("G" & CStr(j)).Value = tempF
range("D" & CStr(j)).Value = ""
  
Next j








lastj = getLastJ("E", 2, Sheet_)

'Exit Sub






For j = start To lastj - 1

If range("E" & CStr(j)).Value = "" Then
Exit For
End If


If range("B" & CStr(j)).Value <> "" Then
BRowData = range("B" & CStr(j)).Value
End If


If range("B" & CStr(j)).Value = "" Then
range("B" & CStr(j)).Value = BRowData
End If


If range("AG" & CStr(j)).Value <> "" Then
AGRowData = range("AG" & CStr(j)).Value
End If


If range("AG" & CStr(j)).Value = "" Then
range("AG" & CStr(j)).Value = AGRowData
End If


Next j

For j = start To lastj - 1

If range("E" & CStr(j)).Value = "" Then
Exit For
End If

tempF = range("J" & CStr(j)).Value
If InStr(tempF, "-") > 0 Then
range("J" & CStr(j)).Value = ""
range("L" & CStr(j)).Value = ""
Else



If InStr(tempF, "dec") > 0 Then
range("J" & CStr(j)).Value = Replace(Split(tempF, " ")(0), ":", "-")
range("L" & CStr(j)).Value = "?"

ElseIf tempF = "resch." Then
range("J" & CStr(j)).Value = ""
range("L" & CStr(j)).Value = ""

ElseIf InStr(tempF, "(") = 0 Then
'-----
range("J" & CStr(j)).Value = Replace(tempF, ":", "-")
'-----
range("L" & CStr(j)).Value = ""

Else
range("J" & CStr(j)).Value = Trim(Replace(Split(tempF, "(")(0), ":", "-"))
range("L" & CStr(j)).Value = Trim(Replace(Split(Split(tempF, "(")(1), ")")(0), ":", "-"))


End If
End If

Next j

With ActiveSheet.Rows("1:" & CStr(lastj - 1))
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .RowHeight = 18.8
End With

For j = 1 To lastj
range("C" & CStr(j)).Value = ""
Next j





range("A1").Value = ""
Columns("A").Interior.ColorIndex = 1
Columns("A").ColumnWidth = 1.33

range("B1").Value = "DATE"
Columns("B").Font.ColorIndex = 2
Columns("B").Interior.ColorIndex = 1
Columns("B").ColumnWidth = 9.78
Columns("B").NumberFormat = "dd/mm/yyyy"

range("C1").Value = ""
Columns("C").Interior.ColorIndex = 1
Columns("C").Font.ColorIndex = 2
Columns("C").ColumnWidth = 0.94

range("D1").Value = ""
Columns("D").Interior.ColorIndex = 1
Columns("D").Font.ColorIndex = 2
Columns("D").ColumnWidth = 5.22

range("E1").Value = "TEAM A"
Columns("E").Interior.ColorIndex = 1
Columns("E").Font.ColorIndex = 2
Columns("E").ColumnWidth = 25.67

range("F1").Value = ""
Columns("F").Interior.ColorIndex = 1
Columns("F").Font.ColorIndex = 2
Columns("A").ColumnWidth = 0.94


range("G1").Value = "TEAM B"
Columns("G").Interior.ColorIndex = 1
Columns("G").Font.ColorIndex = 2
Columns("G").ColumnWidth = 25.67

range("H1").Value = ""
Columns("H").Interior.ColorIndex = 1
Columns("H").Font.ColorIndex = 2
Columns("H").ColumnWidth = 1.56


range("I1").Value = "FIN"
'Columns("I").Interior.ColorIndex = 1
'Columns("I").Font.ColorIndex = 2
Columns("I").ColumnWidth = 3.56

range("J1").Value = "SCORE"
Columns("J").Interior.ColorIndex = 1
Columns("J").Font.ColorIndex = 2
Columns("J").ColumnWidth = 6.78


range("K1").Value = ""
Columns("K").Interior.ColorIndex = 20
Columns("K").Font.ColorIndex = 2
Columns("K").ColumnWidth = 1.11

range("L1").Value = "H/T" & vbNewLine & "SCORE"
Columns("L").Interior.ColorIndex = 1
Columns("L").Font.ColorIndex = 2
Columns("L").ColumnWidth = 6.22

range("M1").Value = "Over" & vbNewLine & "Under"
'Columns("M").Interior.ColorIndex = 13
'Columns("M").Font.ColorIndex = 2
Columns("M").ColumnWidth = 5.56


range("N1").Value = "GG" & vbNewLine & "NG"
'Columns("N").Interior.ColorIndex = 13
'Columns("N").Font.ColorIndex = 2
Columns("N").ColumnWidth = 11

range("O1").Value = "Nikes A"
Columns("O").ColumnWidth = 11
Columns("O").Interior.Color = vbRed

range("P1").Value = "Isop A"
Columns("P").ColumnWidth = 11
Columns("P").Interior.Color = vbRed

range("Q1").Value = "Httes A"
Columns("Q").ColumnWidth = 11
Columns("Q").Interior.Color = vbRed

range("R1").Value = "Nikes A" & vbNewLine & "Home"
Columns("R").ColumnWidth = 11
Columns("R").Interior.Color = vbRed

range("S1").Value = "Isop A" & vbNewLine & "Home"
Columns("S").ColumnWidth = 11
Columns("S").Interior.Color = vbRed

range("T1").Value = "Httes A" & vbNewLine & "Home"
Columns("T").ColumnWidth = 11
Columns("T").Interior.Color = vbRed

range("U1").Value = "Nikes A" & vbNewLine & "Away"
Columns("U").ColumnWidth = 11
Columns("U").Interior.Color = vbRed

range("V1").Value = "Isop A" & vbNewLine & "Away"
Columns("V").ColumnWidth = 11
Columns("V").Interior.Color = vbRed

range("W1").Value = "Httes A" & vbNewLine & "Away"
Columns("W").ColumnWidth = 11
Columns("W").Interior.Color = vbRed


range("X1").Value = "Nikes B"
Columns("X").ColumnWidth = 11
Columns("X").Interior.Color = vbGreen

range("Y1").Value = "Isop B"
Columns("Y").ColumnWidth = 11
Columns("Y").Interior.Color = vbGreen

range("Z1").Value = "Httes B"
Columns("Z").ColumnWidth = 11
Columns("Z").Interior.Color = vbGreen

range("AA1").Value = "Nikes B" & vbNewLine & "Home"
Columns("AA").ColumnWidth = 11
Columns("AA").Interior.Color = vbGreen

range("AB1").Value = "Isop B" & vbNewLine & "Home"
Columns("AB").ColumnWidth = 11
Columns("AB").Interior.Color = vbGreen

range("AC1").Value = "Httes B" & vbNewLine & "Home"
Columns("AC").ColumnWidth = 11
Columns("AC").Interior.Color = vbGreen

range("AD1").Value = "Nikes B" & vbNewLine & "Away"
Columns("AD").ColumnWidth = 11
Columns("AD").Interior.Color = vbGreen

range("AE1").Value = "Isop B" & vbNewLine & "Away"
Columns("AE").ColumnWidth = 11
Columns("AE").Interior.Color = vbGreen

range("AF1").Value = "Httes B" & vbNewLine & "Away"
Columns("AF").ColumnWidth = 11
Columns("AF").Interior.Color = vbGreen

range("AG1").Value = "ROUND"
Columns("AG").Interior.ColorIndex = 10
Columns("AG").Font.ColorIndex = 1
Columns("AG").ColumnWidth = 11

range("AH1").Value = "POINTS A"
Columns("AH").Font.ColorIndex = 2
Columns("AH").Interior.ColorIndex = 1
Columns("AH").ColumnWidth = 11

range("AI1").Value = "POINTS B"
Columns("AI").Interior.ColorIndex = 1
Columns("AI").Font.ColorIndex = 2
Columns("AI").ColumnWidth = 11

range("AJ1").Value = "RANK A"
Columns("AJ").Interior.ColorIndex = 1
Columns("AJ").Font.ColorIndex = 2
Columns("AJ").ColumnWidth = 11

range("AK1").Value = "RANK B"
Columns("AK").Interior.ColorIndex = 1
Columns("AK").Font.ColorIndex = 2
Columns("AK").ColumnWidth = 11

range("AL1").Value = "GOALS A" & vbNewLine & "(+)"
Columns("AL").Interior.ColorIndex = 13
Columns("AL").Font.ColorIndex = 2
Columns("AL").ColumnWidth = 11

range("AM1").Value = "GOALS B" & vbNewLine & "(+)"
Columns("AM").Interior.ColorIndex = 13
Columns("AM").Font.ColorIndex = 2
Columns("AM").ColumnWidth = 11

range("AN1").Value = "GOALS A" & vbNewLine & "(-)"
Columns("AN").Interior.ColorIndex = 13
Columns("AN").Font.ColorIndex = 2
Columns("AN").ColumnWidth = 11

range("AO1").Value = "GOALS B" & vbNewLine & "(-)"
Columns("AO").Interior.ColorIndex = 13
Columns("AO").Font.ColorIndex = 2
Columns("AO").ColumnWidth = 11

range("AP1").Value = "Played A" & vbNewLine & "(-)"
Columns("AP").Interior.ColorIndex = 33
Columns("AP").Font.ColorIndex = 1
Columns("AP").ColumnWidth = 11

range("AQ1").Value = "Played B" & vbNewLine & "(-)"
Columns("AQ").Interior.ColorIndex = 33
Columns("AQ").Font.ColorIndex = 1
Columns("AQ").ColumnWidth = 11

range("AR1").Value = "Points A" & vbNewLine & "Home"
Columns("AR").Font.ColorIndex = 2
Columns("AR").Interior.ColorIndex = 1
Columns("AR").ColumnWidth = 11

range("AS1").Value = "Points A" & vbNewLine & "Away"
Columns("AS").Font.ColorIndex = 2
Columns("AS").Interior.ColorIndex = 1
Columns("AS").ColumnWidth = 11

range("AT1").Value = "Points B" & vbNewLine & "Home"
Columns("AT").Font.ColorIndex = 2
Columns("AT").Interior.ColorIndex = 1
Columns("AT").ColumnWidth = 11

range("AU1").Value = "Points B" & vbNewLine & "Away"
Columns("AU").Font.ColorIndex = 2
Columns("AU").Interior.ColorIndex = 1
Columns("AU").ColumnWidth = 11

range("AV1").Value = "Sub Points"
Columns("AV").Font.ColorIndex = 1
Columns("AV").Interior.ColorIndex = 7
Columns("AV").ColumnWidth = 11

range("AW1").Value = "Sub Ranks"
Columns("AW").Font.ColorIndex = 1
Columns("AW").Interior.ColorIndex = 7
Columns("AW").ColumnWidth = 11

range("AX1").Value = "F1_1"
range("AY1").Value = "F1_X"
range("AZ1").Value = "F1_2"
range("BA1").Value = "F1_U"
range("BB1").Value = "F1_O"
range("BC1").Value = "F1_NG"
range("BD1").Value = "F1_GG"

range("BE1").Value = "F2_1"
range("BF1").Value = "F2_X"
range("BG1").Value = "F2_2"
range("BH1").Value = "F2_U"
range("BI1").Value = "F2_O"
range("BJ1").Value = "F2_NG"
range("BK1").Value = "F2_GG"

range("BL1").Value = "F3_1"
range("BM1").Value = "F3_X"
range("BN1").Value = "F3_2"
range("BO1").Value = "F3_U"
range("BP1").Value = "F3_O"
range("BQ1").Value = "F3_NG"
range("BR1").Value = "F3_GG"

range("BS1").Value = "F4_1"
range("BT1").Value = "F4_X"
range("BU1").Value = "F4_2"
range("BV1").Value = "F4_U"
range("BW1").Value = "F4_O"
range("BX1").Value = "F4_NG"
range("BY1").Value = "F4_GG"

range("BZ1").Value = "F5_1"
range("CA1").Value = "F5_X"
range("CB1").Value = "F5_2"
range("CC1").Value = "F5_U"
range("CD1").Value = "F5_O"
range("CE1").Value = "F5_NG"
range("CF1").Value = "F5_GG"

range("CG1").Value = "F6_1"
range("CH1").Value = "F6_X"
range("CI1").Value = "F6_2"
range("CJ1").Value = "F6_U"
range("CK1").Value = "F6_O"
range("CL1").Value = "F6_NG"
range("CM1").Value = "F6_GG"

range("CN1").Value = "F7_1"
range("CO1").Value = "F7_X"
range("CP1").Value = "F7_2"
range("CQ1").Value = "F7_U"
range("CR1").Value = "F7_O"
range("CS1").Value = "F7_NG"
range("CT1").Value = "F7_GG"

range("CU1").Value = "F8_1"
range("CV1").Value = "F8_X"
range("CW1").Value = "F8_2"
range("CX1").Value = "F8_U"
range("CY1").Value = "F8_O"
range("DZ1").Value = "F8_NG"
range("DA1").Value = "F8_GG"

range("DB1").Value = "F9_1"
range("DC1").Value = "F9_X"
range("DD1").Value = "F9_2"
range("DE1").Value = "F9_U"
range("DF1").Value = "F9_O"
range("DG1").Value = "F9_NG"
range("DH1").Value = "F9_GG"

range("DB1").Value = "F10_1"
range("DC1").Value = "F10_X"
range("DD1").Value = "F10_2"
range("DE1").Value = "F10_U"
range("DF1").Value = "F10_O"
range("DG1").Value = "F10_NG"
range("DH1").Value = "F10_GG"

range("DI1").Value = "F11_1"
range("DJ1").Value = "F11_X"
range("DK1").Value = "F11_2"
range("DL1").Value = "F11_U"
range("DM1").Value = "F11_O"
range("DN1").Value = "F11_NG"
range("DO1").Value = "F11_GG"

range("DP1").Value = "F12_1"
range("DQ1").Value = "F12_X"
range("DR1").Value = "F12_2"
range("DS1").Value = "F12_U"
range("DT1").Value = "F12_O"
range("DU1").Value = "F12_NG"
range("DV1").Value = "F12_GG"

range("DW1").Value = "F13_1"
range("DX1").Value = "F13_X"
range("DY1").Value = "F13_2"
range("DZ1").Value = "F13_U"
range("EA1").Value = "F13_O"
range("EB1").Value = "F13_NG"
range("EC1").Value = "F13_GG"

range("ED1").Value = "F14_1"
range("EE1").Value = "F14_X"
range("EF1").Value = "F14_2"
range("EG1").Value = "F14_U"
range("EH1").Value = "F14_O"
range("EI1").Value = "F14_NG"
range("EJ1").Value = "F14_GG"

Rows(1).Interior.ColorIndex = 15
Rows(1).Font.ColorIndex = 1
Rows(1).Font.Underline = False
Rows(1).RowHeight = 51.8


For j = start To lastj - 1
If range("J" & CStr(j)).Value = "" Or range("J" & CStr(j)).Value = "?" Then
GoTo nextIter
End If
If CInt(Left(range("J" & CStr(j)).Value, 1)) > CInt(Right(range("J" & CStr(j)).Value, 1)) Then
range("I" & CStr(j)).Value = "1"
range("I" & CStr(j)).Interior.ColorIndex = 4
range("I" & CStr(j)).Font.ColorIndex = 1
ElseIf CInt(Left(range("J" & CStr(j)).Value, 1)) < CInt(Right(range("J" & CStr(j)).Value, 1)) Then
range("I" & CStr(j)).Value = "2"
range("I" & CStr(j)).Interior.ColorIndex = 3
range("I" & CStr(j)).Font.ColorIndex = 2
Else
range("I" & CStr(j)).Value = "X"
range("I" & CStr(j)).Interior.ColorIndex = 2
range("I" & CStr(j)).Font.ColorIndex = 1
End If


If CInt(Left(range("J" & CStr(j)).Value, 1)) + CInt(Right(range("J" & CStr(j)).Value, 1)) > 2.5 Then
range("M" & CStr(j)).Value = "Over"
range("M" & CStr(j)).Interior.ColorIndex = 4
range("M" & CStr(j)).Font.ColorIndex = 1

Else
range("M" & CStr(j)).Value = "Under"
range("M" & CStr(j)).Interior.ColorIndex = 3
range("M" & CStr(j)).Font.ColorIndex = 2


End If

If CInt(Left(range("J" & CStr(j)).Value, 1)) > 0 And CInt(Right(range("J" & CStr(j)).Value, 1)) > 0 Then
range("N" & CStr(j)).Value = "G"
range("N" & CStr(j)).Interior.ColorIndex = 37
range("N" & CStr(j)).Font.ColorIndex = 1
Else
range("N" & CStr(j)).Value = "NG"
range("N" & CStr(j)).Interior.ColorIndex = 44
range("N" & CStr(j)).Font.ColorIndex = 1
End If

nextIter:
Next j


 addRanks url, start, lastj, Sheet_

DoEvents
 NikesHttesIsp start, lastj - 1
DoEvents

For j = start To lastj - 1
range("AV" & CStr(j)).Value = CInt(range("AU" & CStr(j)).Value) - CInt(range("AR" & CStr(j)).Value)
Next j

For j = start To lastj - 1
If range("AK" & CStr(j)).Value = "" Or range("AJ" & CStr(j)).Value = "" Then
GoTo niter
End If

range("AW" & CStr(j)).Value = CInt(range("AK" & CStr(j)).Value) - CInt(range("AJ" & CStr(j)).Value)
niter:
Next j

End Sub
