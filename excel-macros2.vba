Option Explicit



Sub Test()
    'Bist du auf dem richtigen Blatt?
    If Range("A1").Interior.ColorIndex <> 34 Then
        MsgBox "Du bist offenbar nicht auf einem Report Blatt. Die Zelle A1 ist nicht hellblau.", _
            vbCritical, Range("A1").Interior.ColorIndex
        End
    End If
    ' Es kann losgehn
    
    Dim StartJahr, StartMonat, EndeJahr, EndeMonat, Datenarten
    Dim j, m, dArt, fromMonth, toMonth, colName, col, i
    Const HeaderZeile = 2
    StartJahr = Range("StartJahr")
    StartMonat = Range("StartMonat")
    EndeJahr = Range("EndeJahr")
    EndeMonat = Range("EndeMonat")
    Datenarten = Split(Range("Datenarten"), ";")
    col = Range("AQ1").Column
    
    For j = StartJahr To EndeJahr
        If j = StartJahr Then fromMonth = StartMonat Else fromMonth = 1
        If j = EndeJahr Then toMonth = EndeMonat Else toMonth = 12
        For m = fromMonth To toMonth
            For dArt = 0 To UBound(Datenarten)
                For i = 1 To 2
                    colName = j & Right("0" & m, 2)
                    If i = 1 Then colName = colName & "#" Else colName = colName & "%"
                    colName = colName & Datenarten(dArt)
                    Cells(HeaderZeile, col).Value = colName
                    col = col + 1
                Next i
            Next dArt
        Next m
    Next j
    
    'MsgBox ActiveCell.Formula2R1C1, , "Formula2R1C1"
    'Call FillGetICVal("E:F", 6, LastFilledRow("B"))
End Sub

Function FillGetICVal(paramScanSpalten, ErsteZeile, LetzteZeile)
    'Const AusgabeSpalte = "AQ", SpalteScanDatenAbsolut = "E"
    Dim rw, formel, formelTeil, ScanSpalten
    Const AusgSpalte1 = "AQ", AusgSpalte2 = "AR"
    ScanSpalten = Split(paramScanSpalten, ":")
    
    For rw = ErsteZeile To LetzteZeile
        formel = Range(ScanSpalten(0) & rw).Formula2R1C1
        If formel Like "*GetICval(*" Then
            formel = Split(formel, ",")
            formelTeil = Split(formel(0), "(")
            formelTeil(1) = "BETRIEBNR"
            formelTeil = Join(formelTeil, "(")
            formel(0) = formelTeil
            formel(3) = "LEFT(R4C,4)"
            formel(4) = "MID(R4C,5,2)"
            formel(5) = "LEFT(R4C,4)"
            formel(6) = "MID(R4C,5,2)"
            formel(7) = "MID(R4C,8,32))" ' extra Klammer zu, letztes Argument
            formel = Join(formel, ",")
            Range(AusgSpalte1 & rw).Formula2R1C1 = formel
        End If
        formel = Range(ScanSpalten(1) & rw).Formula2R1C1
        Range(AusgSpalte2 & rw).Formula2R1C1 = formel
    Next rw
    
    FillGetICVal = 0 ' return value
End Function

Function LastFilledRow(Spalte)
    LastFilledRow = Range(Spalte & Rows.Count).End(xlUp).Row
End Function

Function AssignLevel(TitelSpalte, LevelSpalte, StartZeile, EndZeile)
    'Const TitelSpalte = "B", LevelSpalte = "A", StartZeile = 4
    Dim rw

    For rw = StartZeile To EndZeile
        If Trim(Range(TitelSpalte & rw).Value) = "" Then
            Range(LevelSpalte & rw) = ""
        ElseIf Range(TitelSpalte & rw).Interior.ColorIndex = 55 Then
            Range(LevelSpalte & rw) = 1
        Else
            ' Finde den Level an Zeilen-Gruppierung
            Range(LevelSpalte & rw) = Rows(rw).OutlineLevel + 1
        End If
    Next rw
    AssignLevel = 0 ' return value
End Function


