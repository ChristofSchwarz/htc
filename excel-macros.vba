Option Explicit
Const macroVersion = "v1.02"

Sub ShowMacros()
    Application.Goto Reference:="MainMacro"
End Sub

Function fnColIdx2Name(colIdx)
    ' convert colIdx to colName (2=>B, 3=>C ...)
    fnColIdx2Name = Split(Cells(1, colIdx).Address, "$")(1)
End Function

Sub MainMacro()
    Dim currCol
    currCol = Split(ActiveCell.Address, "$")(1)
    If ActiveSheet.Name <> "Qlik" Then
        MsgBox "Du bist nicht auf dem Qlik sheet", , macroVersion
        Exit Sub
    ElseIf Range(currCol & "1").Text = "Level" And Sheets(1).Name = "KER" Then
        KER_AutoLevel (currCol)
    ElseIf Range(currCol & "1").Text = "Formula" Then
        Summen_Aufloesung (currCol)
    Else
        MsgBox "Kein Makro für diese Spalte bekannt."
    End If

End Sub


Function KER_AutoLevel(scanCol)
    If MsgBox("Wirklich KER Auto-Level Makro starten? Es geht durch Spalte " & scanCol & " des Qlik Sheets" _
      , vbYesNo, macroVersion) = vbNo Then Exit Function
      Dim rw, level, lastRow
      'MsgBox ActiveCell.Font.Italic
      lastRow = Sheets("KER").Range("B65536").End(xlUp).Row
      For rw = 2 To lastRow
          If Sheets("KER").Range(scanCol & rw).Font.Italic Then
              level = 3
          ElseIf Sheets("KER").Range(scanCol & rw).Font.Bold Then
              level = 1
          Else
              level = 2
          End If
          Sheets("Qlik").Range(scanCol & rw).value = level
      Next rw

End Function



Function Summen_Aufloesung(scanCol)
    If MsgBox("Wirklich die ""=SUMME()"" auflösung starten? Es geht durch Spalte " & scanCol & " des Qlik Sheets" _
      , vbYesNo, macroVersion) = vbNo Then Exit Function
      
    Dim rw, level, lastRow, valBefore, needle, changedRows
    Dim changes, skip, valAfter, check, i
    Dim sumContent, sumFrom, sumTo, newSum
    
    lastRow = Sheets("Qlik").Range(scanCol & "65536").End(xlUp).Row
    changedRows = ""
    
    For rw = 2 To lastRow
        valBefore = Range(scanCol & rw).value
        valAfter = Replace(Range(scanCol & rw).value, "SUM(", "SUMME(")
        needle = 1
        
        skip = False
        ' Finde heraus, ob die Formel zu kompliziert ist
        check = Replace(Replace(Replace(valAfter, "SUMME(", ""), scanCol, ""), ":", "")
        check = Replace(Replace(Replace(Replace(check, "(", ""), ")", ""), "+", ""), "-", "")
        For i = 0 To 9
            check = Replace(check, i, "")
        Next i
        
        If check = "=" Then
            changes = 0
            While InStr(needle, valAfter, "SUMME(") > 0 And Not skip
                needle = InStr(needle, valAfter, "SUMME(") + 1
                sumContent = Split(Mid(valAfter, needle + 5, Len(valAfter)), ")")(0)
                check = Replace(Replace(sumContent, scanCol, ""), ":", "")
                For i = 0 To 9
                    check = Replace(check, i, "")
                Next i
                If Len(check) > 0 Then
                   skip = True
                Else
                    sumFrom = Replace(Split(sumContent, ":")(0), scanCol, "")
                    sumTo = Replace(Split(sumContent, ":")(1), scanCol, "")
                    newSum = "("
                    For i = sumFrom To sumTo
                        newSum = newSum & "+" & scanCol & i
                    Next
                    newSum = Replace(newSum & ")", "(+", "(")
                    changes = changes + 1
                    'MsgBox newSum, , sumContent
                End If
                valAfter = Replace(valAfter, "SUMME(" & sumContent & ")", newSum)
            Wend
            If changes > 0 And Not skip And valAfter <> valBefore Then
                If Len(Replace(valAfter, "-", "")) = Len(valAfter) Then
                    ' Keine Subtraktionen ... Klammern entfernen
                    valAfter = Replace(Replace(valAfter, ")", ""), "(", "")
                End If
                changedRows = changedRows & "," & rw
                If Not Range(scanCol & rw).Comment Is Nothing Then Range(scanCol & rw).Comment.Delete
                Range(scanCol & rw).AddComment (valBefore)
                Range(scanCol & rw).value = "'" & valAfter
                Range(scanCol & rw).Activate
                
            End If
        End If
    Next rw
    If Len(changedRows) = 0 Then
        MsgBox "Keine Änderungen machbar bis Zeile " & rw
    Else
        MsgBox "Änderungen in Zeilen " & changedRows & " gemacht"
    End If
    
End Function
