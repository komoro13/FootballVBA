Attribute VB_Name = "Module27"

Function countFilterStats(ByVal Sheetname As String) As data_
Dim lastj As Long
lastj = SgetLastJ(1, Sheetname)
Dim j


Dim dat0 As New data_
dat0.assoi = 0
dat0.xinaria = 0
dat0.dipla = 0
dat0.under = 0
dat0.over = 0
dat0.ng = 0
dat0.gg = 0


For j = 1 To lastj
If Worksheets(Sheetname).range("I" & CInt(j)).Value = "1" Then
dat0.assoi = dat0.assoi + 1
ElseIf Worksheets(Sheetname).range("I" & CInt(j)).Value = "X" Then
dat0.xinaria = dat0.xinaria + 1
ElseIf Worksheets(Sheetname).range("I" & CInt(j)).Value = "2" Then
dat0.dipla = dat0.dipla + 1
End If
If Worksheets(Sheetname).range("M" & CInt(j)).Value = "Under" Then
dat0.under = dat0.under + 1
ElseIf Worksheets(Sheetname).range("M" & CInt(j)).Value = "Over" Then
dat0.over = dat0.over + 1
End If
If Worksheets(Sheetname).range("N" & CInt(j)).Value = "NG" Then
dat0.ng = dat0.ng + 1
ElseIf Worksheets(Sheetname).range("N" & CInt(j)).Value = "G" Then
dat0.gg = dat0.gg + 1
End If
Next j
Set countFilterStats = dat0
End Function
