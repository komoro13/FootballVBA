Attribute VB_Name = "Module71"
Sub DeleteDuplicates()
Dim j As Long
Dim x As Long
x = 0

For j = 1 To range("B2").End(xlDown)

If range("B" + CStr(j)).Value = "" Then
Exit For
End If

If range("B" & CStr(j)).Value = range("B" & CStr(j + 1)).Value And range("E" & CStr(j)).Value = range("E" & CStr(j + 1)).Value And range("G" & CStr(j)).Value = range("G" & CStr(j + 1)).Value And range("J" & CStr(j)).Value = range("J" & CStr(j + 1)).Value Then
Rows(j).EntireRow.Delete
x = x + 1
End If
Next j
MsgBox "Done!" & vbNewLine & "Deleted " & x & " rows"
End Sub
