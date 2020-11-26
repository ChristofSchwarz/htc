Option Explicit

Function ColNumToText(ColNum)
   ' Gibt zu einer Spalten-Nummer den Spalten-TextNamen zurück, z.B. 256 -> "IV"
    ColNumToText = Split(ActiveSheet.Cells(1, ColNum).Address, "$")(1)
End Function

Function ColTextToNum(ColName)
   ' Gibt zu einem Spalten-TextNamen die laufende Spalten-Nummer zurück, z.B. "IV" -> 256
    ColTextToNum = ActiveSheet.Range(ColName & "1").Column
End Function

Sub CopyWorkbook()

    Const sourceSheetName = "KER nach Abteilungen", BetriebNr = 543
    Const firstRow = 4, TitelSpalte = "B", LevelSpalte = "A"
    Const startQlikColumn = "AQ"
    Const ScanSpalte1 = "E"
    Const ScanSpalte2 = "F"
    
    Dim sourceExcel As Workbook, targetExcel As Workbook
    Dim sourceSheet As Worksheet, targetSheet As Worksheet, indexSheet As Worksheet
    Dim col, ColName, lastRow, cLabel
    Dim StartJahr, StartMonat, EndeJahr, EndeMonat, Datenarten
    Dim i, j, m, dArt, fromMonth, toMonth ', fromRange, toRange
    Dim prevCalcSetting
    
    prevCalcSetting = Application.Calculation
    ' Stoppe automatisches Neuberechnen der Zellen,bis wir fertig sind
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    Set sourceExcel = Workbooks.Open("C:\Users\christof.schwarz\Downloads\KER.xls", False)
    Set sourceSheet = sourceExcel.Sheets(sourceSheetName)
    Set targetExcel = Workbooks.Add
    Set indexSheet = targetExcel.Sheets(1)
    indexSheet.Name = "Qlik"
    indexSheet.Range("A1").Value = "BetriebNr"
    indexSheet.Range("A2").Value = BetriebNr
    ActiveWorkbook.Names.Add Name:="BETRIEBNR", RefersToR1C1:="=Qlik!R2C1"
    indexSheet.Range("B1").Value = "SheetList"
    Set targetSheet = targetExcel.Sheets.Add(Null, indexSheet)
    targetSheet.Name = sourceSheetName

    sourceSheet.Activate
    lastRow = lastFilledRow(sourceSheet, TitelSpalte)
    sourceSheet.Rows(firstRow & ":" & lastRow).Select
    Application.CutCopyMode = False
    Selection.Copy
    targetSheet.Activate
    targetSheet.Range("A" & firstRow).Select
    targetSheet.Paste
    Range("A1").Value = "Level"
    Call AssignLevel(TitelSpalte, LevelSpalte, firstRow, lastRow)
    
    Range("B1").Value = "Label"
    
        
    StartJahr = 2017
    StartMonat = 11
    EndeJahr = 2020
    EndeMonat = 10
    Datenarten = Split("IST;FORECAST;PLAN 1", ";")
    col = ColTextToNum(startQlikColumn)
    
    For j = StartJahr To EndeJahr
        If j = StartJahr Then fromMonth = StartMonat Else fromMonth = 1
        If j = EndeJahr Then toMonth = EndeMonat Else toMonth = 12
        For m = fromMonth To toMonth
            For dArt = 0 To UBound(Datenarten)
                For i = 1 To 2
                    cLabel = j & Right("0" & m, 2)
                    If i = 1 Then cLabel = cLabel & "#" Else cLabel = cLabel & "%"
                    cLabel = cLabel & Datenarten(dArt)
                    Cells(1, col).Value = cLabel
                    col = col + 1
                Next i
                'FillGetICVal("E:F", 6, LastFilledRow("B"))
                If ColNumToText(col - 2) = startQlikColumn Then
                    ' Beim ersten Loop formen wir die Formeln um gemäß vorlage in Spalten E:F und kopieren sie
                    Call FillGetICVal(ScanSpalte1, ScanSpalte2, ColNumToText(col - 2), ColNumToText(col - 1), firstRow, lastRow)
                    Range(startQlikColumn & firstRow & ":" & ColNumToText(ColTextToNum(startQlikColumn) + 1) & lastRow).Select
                    Selection.Copy
                Else
                    ' Bei den folgenden Malen kopieren wir uns die spalten mit den neuen Formeln
                    'fromRange = startQlikColumn & firstRow & ":" & ColNumToText(ColTextToNum(startQlikColumn) + 1) & lastRow
                    'toRange = ColNumToText(col - 2) & firstRow
                    'MsgBox fromRange & " --> " & toRange
                    'Range(fromRange).Select
                    'Selection.Copy
                    Range(ColNumToText(col - 2) & firstRow).Activate
                    ActiveSheet.Paste
                End If
            Next dArt
        Next m
    Next j
    
    'Autofit alle Spaltenbreiten
    Columns("A:" & ColNumToText(col)).EntireColumn.AutoFit
    
    'Lösche die ehemaligen spalten zwischen Label und AQ, damit nicht unnötige Formeln berechnet werden
    For col = ColTextToNum(TitelSpalte) + 1 To ColTextToNum(startQlikColumn) - 1
        ColName = ColNumToText(col)
        'Range(colName & "1").UnMerge
        Range(ColName & "1").Value = "Ignore" & col
        If Range(ColName & (firstRow + 1)).Text <> "" Then
            ' Lösche spalteninhalt
            Range(ColName & "2:" & ColName & lastRow).UnMerge
            Range(ColName & "2:" & ColName & lastRow).Clear
        End If
        Columns(ColNumToText(col) & ":" & ColNumToText(col)).ColumnWidth = 0.75
    Next col


    Application.Calculate
    Application.Calculation = prevCalcSetting
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'targetExcel.SaveAs Filename:="C:\Users\christof.schwarz\Downloads\Book1.xlsx", FileFormat:=xlOpenXMLWorkbook
    'targetExcel.Close
    sourceExcel.Close
    Application.DisplayAlerts = True
End Sub


Sub Test()
    'Bist du auf dem richtigen Blatt?
    If Range("A1").Interior.ColorIndex <> 34 Then
        MsgBox "Du bist offenbar nicht auf einem Report Blatt. Die Zelle A1 ist nicht hellblau.", _
            vbCritical, Range("A1").Interior.ColorIndex
        End
    End If
    ' Es kann losgehn
End Sub

Function FillGetICVal(ScanSpalte1, ScanSpalte2, AusgSpalte1, AusgSpalte2, ErsteZeile, LetzteZeile)
    'Const AusgabeSpalte = "AQ", SpalteScanDatenAbsolut = "E"
    Dim rw, formel, formelTeil, ScanSpalten
   ' Const AusgSpalte1 = "AQ", AusgSpalte2 = "AR"
    
    
    For rw = ErsteZeile To LetzteZeile
        formel = Range(ScanSpalte1 & rw).Formula2R1C1
        If formel Like "*GetICval(*" Then
            formel = Split(formel, ",")
            formelTeil = Split(formel(0), "(")
            formelTeil(1) = "BETRIEBNR"
            formelTeil = Join(formelTeil, "(")
            formel(0) = formelTeil
            formel(3) = "LEFT(R1C,4)"
            formel(4) = "MID(R1C,5,2)"
            formel(5) = "LEFT(R1C,4)"
            formel(6) = "MID(R1C,5,2)"
            formel(7) = "MID(R1C,8,32))" ' extra Klammer zu, letztes Argument
            formel = Join(formel, ",")
            Range(AusgSpalte1 & rw).Formula2R1C1 = formel
        End If
        formel = Range(ScanSpalte2 & rw).Formula2R1C1
        Range(AusgSpalte2 & rw).Formula2R1C1 = formel
    Next rw
    
    FillGetICVal = 0 ' return value
End Function

Function lastFilledRow(Sheet, Spalte)
    lastFilledRow = Sheet.Range(Spalte & Rows.Count).End(xlUp).Row
End Function

Function AssignLevel(TitelSpalte, LevelSpalte, StartZeile, EndZeile)
    'Const TitelSpalte = "B", LevelSpalte = "A", StartZeile = 4
    Dim rw

    For rw = StartZeile To EndZeile
        If Trim(Range(TitelSpalte & rw).Formula) = "" Then
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


