Attribute VB_Name = "Module3"
Option Explicit
Public Sub useQueryTable(ByVal url As String, ByVal tb As String, Sheetname As Worksheet, ByVal start As Long)


Dim table As QueryTable
Set table = Sheetname.QueryTables.Add("URL;" & url, Sheetname.range("B" & start))
DoEvents
With table
    .WebSelectionType = xlSpecifiedTables
    .WebTables = tb
    .PreserveFormatting = True

    .WebFormatting = xlWebFormattingAll
    .Refresh BackgroundQuery:=False
End With

End Sub
