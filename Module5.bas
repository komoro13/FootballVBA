Attribute VB_Name = "Module5"
Public Type team
    Name As String
    
    Isp As Integer
    Nikes As Integer
    Httes As Integer
    

    
    IspH As Integer
    NikesH As Integer
    HttesH As Integer
 
    
    IspA As Integer
    NikesA As Integer
    HttesA As Integer
    
    goal_ef As Integer
    goal_ev As Integer
    
    points As Integer
    pointsH As Integer
    pointsA As Integer
    
End Type


Function NikesHttesIsp(ByVal JStart As Long, ByVal JLast As Long)

Dim z As Integer
Dim x As Integer
Dim j As Long
Dim i As Long






Dim Teams(100) As team




TeamsCount = 0

'================================================================VARIABLES===============================================================================
x = 0
While ThisWorkbook.Sheets("Generic").range("D" & x + 3).Value <> ""
Teams(x).Name = ThisWorkbook.Sheets("Generic").range("D" & x + 3).Value

x = x + 1
Wend











For j = JStart To JLast
For i = JStart To JLast

If i >= j Then
Exit For
End If
'======================================================================== SIDE 1 POINTS COUNT=================================================================
If range("J" & CStr(i)).Value = "" Or range("J" & CStr(i)).Value = "?" Then
GoTo ni1
End If
If range("E" & CStr(j)).Value = range("E" & CStr(i)).Value Then


    For z = 0 To 50
    If Teams(z).Name = range("E" & CStr(j)).Value Then
     Teams(z).goal_ev = Teams(z).goal_ev + CInt(Split(range("J" & CStr(i)).Value, "-")(0))
      Teams(z).goal_ef = Teams(z).goal_ef + CInt(Split(range("J" & CStr(i)).Value, "-")(1))
        If range("I" & CStr(i)).Value = "1" Then
            Teams(z).Nikes = Teams(z).Nikes + 1
            Teams(z).NikesH = Teams(z).NikesH + 1
            Teams(z).points = Teams(z).points + 3
            Teams(z).pointsH = Teams(z).pointsH + 3
            Exit For
        End If
        If range("I" & CStr(i)).Value = "X" Then
            Teams(z).Isp = Teams(z).Isp + 1
             Teams(z).IspH = Teams(z).IspH + 1
             Teams(z).points = Teams(z).points + 1
             Teams(z).pointsH = Teams(z).pointsH + 1
            Exit For
        End If
            If range("I" & CStr(i)).Value = "2" Then
            Teams(z).Httes = Teams(z).Httes + 1
            Teams(z).HttesH = Teams(z).HttesH + 1
            Exit For
        End If
        
     
    End If

Next z
End If
'======================================================================== SIDE 2 POINTS COUNT=================================================================



If range("E" & CStr(j)).Value = range("G" & CStr(i)).Value Then

    For z = 0 To 50
     If Teams(z).Name = range("E" & CStr(j)).Value Then
       Teams(z).goal_ef = Teams(z).goal_ef + CInt(Split(range("J" & CStr(i)).Value, "-")(0))
        Teams(z).goal_ev = Teams(z).goal_ev + CInt(Split(range("J" & CStr(i)).Value, "-")(1))
        If range("I" & CStr(i)).Value = "2" Then
            Teams(z).Nikes = Teams(z).Nikes + 1
            Teams(z).NikesA = Teams(z).NikesA + 1
            Teams(z).points = Teams(z).points + 3
            Teams(z).pointsA = Teams(z).pointsA + 3
            Exit For
        End If
        If range("I" & CStr(i)).Value = "X" Then
            Teams(z).Isp = Teams(z).Isp + 1
            Teams(z).IspA = Teams(z).IspA + 1
            Teams(z).points = Teams(z).points + 1
            Teams(z).pointsA = Teams(z).pointsA + 1
            Exit For
        End If
        If range("I" & CStr(i)).Value = "1" Then
        Teams(z).Httes = Teams(z).Httes + 1
        Teams(z).HttesA = Teams(z).HttesA + 1
        Exit For
        End If
      
    End If

Next z
End If
ni1:
Next i




'========================================================================SIDE 1 DISPLAY TEAM 1===========================================================
For z = 0 To 50
If Teams(z).Name = range("E" & CStr(j)).Value Then

range("O" & CStr(j)).Value = Teams(z).Nikes
range("P" & CStr(j)).Value = Teams(z).Isp
range("Q" & CStr(j)).Value = Teams(z).Httes
range("R" & CStr(j)).Value = Teams(z).NikesH
range("S" & CStr(j)).Value = Teams(z).IspH
range("T" & CStr(j)).Value = Teams(z).HttesH
range("U" & CStr(j)).Value = Teams(z).NikesA
range("V" & CStr(j)).Value = Teams(z).IspA
range("W" & CStr(j)).Value = Teams(z).HttesA


range("AL" & CStr(j)).Value = Teams(z).goal_ev
range("AN" & CStr(j)).Value = Teams(z).goal_ef
range("AH" & CStr(j)).Value = Teams(z).points
range("AR" & CStr(j)).Value = Teams(z).pointsH
range("AS" & CStr(j)).Value = Teams(z).pointsA

Teams(z).points = 0
Teams(z).pointsA = 0
Teams(z).pointsH = 0

Teams(z).goal_ef = 0
Teams(z).goal_ev = 0

Teams(z).Httes = 0
Teams(z).Nikes = 0
Teams(z).Isp = 0

Teams(z).HttesH = 0
Teams(z).NikesH = 0
Teams(z).IspH = 0

Teams(z).HttesA = 0
Teams(z).NikesA = 0
Teams(z).IspA = 0

Exit For
End If

Next z


If CInt(range("AG" & CStr(j)).Value) > CInt(range("AG" & CStr(j + 1)).Value) Then

GoTo c
Exit For
End If

Next j
'=============================================================================END OF SIDE 1==================================================================
'=============================================================================SIDE 2=========================================================================
c:

For j = JStart To JLast




'MsgBox DateDiff("D", "25/7/1981", Range("B" & CStr(J)).Value)


For i = JStart To JLast

If i >= j Then
'MsgBox "Interupting loop"
Exit For
End If
If range("J" & CStr(i)).Value = "" Or range("J" & CStr(i)).Value = "?" Then
GoTo ni2
End If
If range("G" & CStr(j)).Value = range("E" & CStr(i)).Value Then


    For z = 0 To 50
    If Teams(z).Name = range("G" & CStr(j)).Value Then
        Teams(z).goal_ev = Teams(z).goal_ev + CInt(Split(range("J" & CStr(i)).Value, "-")(0))
      Teams(z).goal_ef = Teams(z).goal_ef + CInt(Split(range("J" & CStr(i)).Value, "-")(1))
   
        If range("I" & CStr(i)).Value = "1" Then
            Teams(z).points = Teams(z).points + 3
            Teams(z).Nikes = Teams(z).Nikes + 1
            Teams(z).NikesH = Teams(z).NikesH + 1
            Teams(z).pointsH = Teams(z).pointsH + 3
            Exit For
        End If
        If range("I" & CStr(i)).Value = "X" Then
            Teams(z).points = Teams(z).points + 1
            Teams(z).Isp = Teams(z).Isp + 1
            Teams(z).IspH = Teams(z).IspH + 1
            Teams(z).pointsH = Teams(z).pointsH + 1
            Exit For
        End If
        If range("I" & CStr(i)).Value = "2" Then
            Teams(z).Httes = Teams(z).Httes + 1
            Teams(z).HttesH = Teams(z).HttesH + 1
            Exit For
        End If
        
    End If
    

'MsgBox Teams(z).Name & " " & Teams(z).Points

'MsgBox "Found " & Teams(z).Name & " " & Teams(z).Points
Next z
End If

'========================================================================SIDE 2 POINTS COUNT=================================================================

'MsgBox Range("E" & CStr(i)).Value & " " & Range("I" & CStr(i)).Value & " " & Range("B" & CStr(i)).Value & " i = " & i & " j = " & j & " " & Range("AB" & CStr(j)).Value

If range("G" & CStr(j)).Value = range("G" & CStr(i)).Value Then

    For z = 0 To 50
    If Teams(z).Name = range("G" & CStr(j)).Value Then
      Teams(z).goal_ef = Teams(z).goal_ef + CInt(Split(range("J" & CStr(i)).Value, "-")(0))
      Teams(z).goal_ev = Teams(z).goal_ev + CInt(Split(range("J" & CStr(i)).Value, "-")(1))
        If range("I" & CStr(i)).Value = "2" Then
            Teams(z).points = Teams(z).points + 3
            Teams(z).Nikes = Teams(z).Nikes + 1
            Teams(z).NikesA = Teams(z).NikesA + 1
            Teams(z).pointsA = Teams(z).pointsA + 3
            Exit For
        End If
        If range("I" & CStr(i)).Value = "X" Then
            Teams(z).points = Teams(z).points + 1
            Teams(z).Isp = Teams(z).Isp + 1
            Teams(z).IspA = Teams(z).IspA + 1
            Teams(z).pointsA = Teams(z).pointsA + 1
            Exit For
        End If
        If range("I" & CStr(i)).Value = "1" Then
            Teams(z).Httes = Teams(z).Httes + 1
            Teams(z).HttesA = Teams(z).HttesA + 1
            Exit For
        End If
     
    End If
    
    
'MsgBox Teams(z).Name & " " & Teams(z).Points

'MsgBox "Found " & Teams(z).Name & " " & Teams(z).Points
Next z
End If
ni2:
'MsgBox CInt(Right(Range("J" & CStr(i)), 1)) & " " & i
Next i





'========================================================================SIDE 2 DISPLAY TEAM 2===========================================================




For z = 0 To 50
If Teams(z).Name = range("G" & CStr(j)).Value Then


range("X" & CStr(j)).Value = Teams(z).Nikes
range("Y" & CStr(j)).Value = Teams(z).Isp
range("Z" & CStr(j)).Value = Teams(z).Httes
range("AA" & CStr(j)).Value = Teams(z).NikesH
range("AB" & CStr(j)).Value = Teams(z).IspH
range("AC" & CStr(j)).Value = Teams(z).HttesH
range("AD" & CStr(j)).Value = Teams(z).NikesA
range("AE" & CStr(j)).Value = Teams(z).IspA
range("AF" & CStr(j)).Value = Teams(z).HttesA

range("AI" & CStr(j)).Value = Teams(z).points
range("AM" & CStr(j)).Value = Teams(z).goal_ev
range("AO" & CStr(j)).Value = Teams(z).goal_ef
range("AT" & CStr(j)).Value = Teams(z).pointsH
range("AU" & CStr(j)).Value = Teams(z).pointsA

Teams(z).points = 0
Teams(z).pointsA = 0
Teams(z).pointsH = 0

Teams(z).goal_ef = 0
Teams(z).goal_ev = 0

Teams(z).Httes = 0
Teams(z).Nikes = 0
Teams(z).Isp = 0

Teams(z).HttesH = 0
Teams(z).NikesH = 0
Teams(z).IspH = 0


Teams(z).HttesA = 0
Teams(z).NikesA = 0
Teams(z).IspA = 0

Exit For
End If

Next z

If CInt(range("AG" & CStr(j)).Value) > CInt(range("AG" & CStr(j + 1)).Value) Then
Exit For
End If



Next j

ex:



End Function
