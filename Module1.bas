Attribute VB_Name = "Module1"
Sub run(ByVal start_ As String, ByVal end_ As String, ByVal link As String, ByVal link_2 As String, ByVal Sheet_ As String, ByVal process As Integer)
Attribute run.VB_ProcData.VB_Invoke_Func = "r\n14"
Dim year As String
Dim start As Long
Dim link_2_ As String
link_2_ = link_2
If link_2 = "" Then
link_2_ = link
End If
MsgBox "starting"
ThisWorkbook.Sheets(Sheet_).Select

Worksheets(Sheet_).Range("A1:AW" & Rows.count).Sort key1:=Range("B2"), Order1:=xlAscending, Header:=xlYes


If InStr(start_, "-") > 0 Then


year = Split(start_, "-")(0)
'loadingForm.Show
While year <= end_
'loadingForm.Bar.Width = (200 * year) / CInt(Split(start_, "-")(0))
'DoEvents
start = getLastJ("E", 2, Sheet_)
url = "https://www.worldfootball.net/all_matches/" & link & "-" & year & "-" & year + 1
rndurl = "https://www.worldfootball.net/schedule/" & link_2_ & "-" & year & "-" & year + 1 & "-spieltag/"

If process = 0 Or process = 2 Then

useQueryTable url, "2", Worksheets(Sheet_), start

End If

If process = 1 Or process = 2 Then
format_ rndurl, start, Sheet_
End If

year = year + 1
Wend

Else

year = start_
'loadingForm.Bar.Width = (200 * year) / end_
DoEvents
While year <= end_
start = getLastJ("E", 2, Sheet_)
url = "https://www.worldfootball.net/all_matches/" & link & "-" & year
rndurl = "https://www.worldfootball.net/schedule/" & link_2_ & "-" & year & "-spieltag/"

If process = 0 Or process = 2 Then
useQueryTable url, "2", Worksheets(Sheet_), start
End If

If process = 1 Or process = 2 Then
format_ rndurl, start, Sheet_
End If

year = year + 1
Wend
End If


Worksheets(Sheet_).Range("A1:AW" & Rows.count).Sort key1:=Range("B2"), Order1:=xlDescending, Header:=xlYes
'loadingForm.Hide

Worksheets("Config").Range("B" & (SgetLastJ(1, "Config"))).Value = link
Worksheets("Config").Range("C" & (SgetLastJ(1, "Config"))).Value = Sheet_
Worksheets("Config").Range("D" & (SgetLastJ(1, "Config"))).Value = end_
Worksheets("Config").Range("H" & (SgetLastJ(1, "Config"))).Value = Now
Worksheets("Config").Range("I" & (SgetLastJ(1, "Config"))).Value = link_2

MsgBox "Finished"
End Sub
