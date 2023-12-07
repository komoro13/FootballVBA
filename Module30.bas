Attribute VB_Name = "Module30"
Sub datesort()
range("A1:AW" & Rows.count).Sort key1:=range("B2"), Order1:=xlDescending, Header:=xlYes
End Sub
