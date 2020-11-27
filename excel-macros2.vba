Option Explicit
Const StartQlikColumn = "AQ"
Const LevelSpalte = "A"


Function ColNumToText(ColNum)
   ' Gibt zu einer Spalten-Nummer den Spalten-TextNamen zurück, z.B. 256 -> "IV"
    ColNumToText = Split(ActiveSheet.Cells(1, ColNum).Address, "$")(1)
End Function


Function ColTextToNum(colName)
   ' Gibt zu einem Spalten-TextNamen die laufende Spalten-Nummer zurück, z.B. "IV" -> 256
    ColTextToNum = ActiveSheet.range(colName & "1").Column
End Function


Function lastFilledRow(Sheet, col)
    ' returns the last row where there is a value filled in a given Spalte
    lastFilledRow = Sheet.range(col & Rows.Count).End(xlUp).row
End Function


Function IsValueInRange(value, range)
    ' returns True or Fales if the exact match of a given value is found in a given Range
    Dim find
    On Error Resume Next
    find = range.find(value, , xlValues, xlWhole, , , True)
    On Error GoTo 0
    IsValueInRange = VarType(find) = vbString
End Function


Sub ReadSheets()
    Dim sourceExcel As Workbook, sh As Worksheet, file, cell
    Dim knownSheets As range, find, settingsSheet As Worksheet
    Dim table, colName, newRow
    
    Set settingsSheet = ActiveSheet
    Set table = settingsSheet.ListObjects("SheetTable")
    Set knownSheets = table.ListColumns("Arbeitsblätter").DataBodyRange
    colName = Split(knownSheets.Address, "$")(1)
    If Not range("ROOTFOLDER") Like "*\" Then range("ROOTFOLDER") = range("ROOTFOLDER") & "\"
    If Not range("OUTPUTFOLDER") Like "*\" Then range("OUTPUTFOLDER") = range("OUTPUTFOLDER") & "\"
    file = range("ROOTFOLDER") & range("INPUTFILE")
    On Error GoTo Err
    Set sourceExcel = Workbooks.Open(file, False, True)
    On Error GoTo 0
    ActiveWindow.Visible = False
    For Each sh In sourceExcel.Sheets
        If Not IsValueInRange(sh.Name, knownSheets) Then
            table.ListRows.Add
            newRow = Replace(Split(table.ListRows(table.ListRows.Count).range.Address, "$")(2), ":", "")
            settingsSheet.range(colName & newRow).value = sh.Name
        End If
        'MsgBox sh.Name, ,
    Next sh
    MsgBox sourceExcel.Sheets.Count & " Sheets in '" & file & "' gefunden.", vbInformation
    
    Application.DisplayAlerts = False
    sourceExcel.Close
    Application.DisplayAlerts = True
    Exit Sub
Err:
    MsgBox "File '" & file & "' nicht gefunden.", vbCritical
End Sub



Function getTableFieldAddress(table, row, fieldName)
' returns the absolute address of a cell inside a data table, defined by row and field name
    Dim tableFirstCol
    tableFirstCol = table.ListColumns(1).range.Column
    On Error GoTo Err
    getTableFieldAddress = table.DataBodyRange(row, table.ListColumns(fieldName).range.Column - tableFirstCol + 1).Address
    On Error GoTo 0
    Exit Function
Err:
    MsgBox "Kein feld namens '" & fieldName & "' in Tabellenobjekt '" & table.Name & "'.", vbCritical
    End
End Function


Function getTableField(table, row, fieldName)
' returns the field name of a given table's row
    Dim tableFirstCol
    tableFirstCol = table.ListColumns(1).range.Column
    On Error GoTo Err
    getTableField = table.DataBodyRange(row, table.ListColumns(fieldName).range.Column - tableFirstCol + 1)
    On Error GoTo 0
    Exit Function
Err:
    MsgBox "Kein feld namens '" & fieldName & "' in Tabellenobjekt '" & table.Name & "'.", vbCritical
    End
End Function


Function CreateQlikSheet(targetExcel, betriebNr, betriebName)
    Dim indexSheet As Worksheet
    
    Set indexSheet = targetExcel.Sheets(1)
    indexSheet.Name = "Qlik"
    indexSheet.range("A1").value = "BetriebNr"
    indexSheet.range("A2").value = betriebNr
    targetExcel.Names.Add Name:="BETRIEBNR", RefersToR1C1:="=Qlik!R2C1"
    indexSheet.range("B1").value = "BetriebName"
    indexSheet.range("B2").value = betriebName
    indexSheet.range("C1").value = "SheetList"
    
End Function


Sub CopyWorkbook()

    Dim betriebNr, betriebName, sourceSheetName, firstRow, titelSpalte
    Dim scanSpalte1, scanSpalte2
    Dim saveAsFileName, saveFolder
    Dim sheetTable, betriebeTable, sheetTableRow, betriebeTableRow
    Dim sourceExcel As Workbook, targetExcel As Workbook
    Dim sourceSheet As Worksheet, settingsSheet As Worksheet
    'Dim col, colName, lastRow, cLabel
    Dim startJahr, startMonat, endeJahr, endeMonat, datenarten
    'Dim i, j, m, dArt, fromMonth, toMonth ', fromRange, toRange
    Dim prevCalcSetting
    
    If ActiveSheet.Name <> "BetriebSettings" Then
        MsgBox "Du bist nicht auf dem BetriebSettings Sheet.", vbCritical
        Exit Sub
    End If
    
    Set settingsSheet = ActiveSheet
    saveFolder = Sheets("SheetSettings").range("ROOTFOLDER")
    If Not saveFolder Like "*\" Then saveFolder = saveFolder & "\"
    saveFolder = saveFolder & Sheets("SheetSettings").range("OUTPUTFOLDER")
    If Not saveFolder Like "*\" Then saveFolder = saveFolder & "\"
     
    Set betriebeTable = ActiveSheet.ListObjects("BetriebeTable")
    Set sheetTable = Sheets("SheetSettings").ListObjects("SheetTable")
    
    For betriebeTableRow = 1 To betriebeTable.ListRows.Count
    
        betriebName = getTableField(betriebeTable, betriebeTableRow, "BetriebName")
        betriebNr = getTableField(betriebeTable, betriebeTableRow, "BetriebNr")
        saveAsFileName = saveFolder & getTableField(betriebeTable, betriebeTableRow, "OutputFile")
        startJahr = getTableField(betriebeTable, betriebeTableRow, "StartJahr")
        startMonat = getTableField(betriebeTable, betriebeTableRow, "StartMonat")
        endeJahr = getTableField(betriebeTable, betriebeTableRow, "EndeJahr")
        endeMonat = getTableField(betriebeTable, betriebeTableRow, "EndeMonat")
        datenarten = getTableField(betriebeTable, betriebeTableRow, "Datenarten")
        'MsgBox betriebName & "|" & betriebNr & "|" & saveAsFileName & "|" & startJahr & "|" & startMonat & "|" & endeJahr & "|" & endeMonat & "|" & datenarten
        
        For sheetTableRow = 1 To sheetTable.ListRows.Count
            If getTableField(sheetTable, sheetTableRow, "Aktiv") = 1 Then
            
                sourceSheetName = getTableField(sheetTable, sheetTableRow, "Arbeitsblätter")
                firstRow = getTableField(sheetTable, sheetTableRow, "ErsteZeile")
                titelSpalte = getTableField(sheetTable, sheetTableRow, "TitelSpalte")
                scanSpalte1 = getTableField(sheetTable, sheetTableRow, "ScanSpalte1")
                scanSpalte2 = getTableField(sheetTable, sheetTableRow, "ScanSpalte2")
                
                If sourceSheetName = "" Or firstRow = "" Or titelSpalte = "" Or scanSpalte1 = "" Or scanSpalte2 = "" Then
                    MsgBox "Zeile " & sheetTableRow & " (" & sourceSheetName & ") in SheetSettings ist nicht vollständig ausgefüllt", vbCritical
                    End
                End If
                
                If sourceExcel Is Nothing Then
                    prevCalcSetting = Application.Calculation
                    ' Stoppe automatisches Neuberechnen der Zellen,bis wir fertig sind
                    Application.Calculation = xlCalculationManual
                    Application.DisplayAlerts = False
                    Set sourceExcel = Workbooks.Open(range("ROOTFOLDER") & range("INPUTFILE"), False)
                End If
                
                Set sourceSheet = sourceExcel.Sheets(sourceSheetName)
                
                If targetExcel Is Nothing Then
                    Set targetExcel = Workbooks.Add
                    Call CreateQlikSheet(targetExcel, betriebNr, betriebName)
                End If
                Call CreateSheet(sourceSheet, targetExcel, titelSpalte, scanSpalte1, scanSpalte2, firstRow _
                    , startJahr, startMonat, endeJahr, endeMonat, datenarten)
                
            End If
        Next
        
        Application.Calculate
        Application.Calculation = prevCalcSetting
    
        targetExcel.SaveAs Filename:=saveAsFileName, FileFormat:=xlOpenXMLWorkbook
        targetExcel.Close
        Set targetExcel = Nothing
        ' Schreibe das Update Datum/Uhrzeit in die spalte "Updated"
        settingsSheet.range(getTableFieldAddress(betriebeTable, betriebeTableRow, "Updated")).value = Format(Now(), "DD.MM.YYYY hh:mm:ss")
    Next
        
    sourceExcel.Close
    Application.Calculation = prevCalcSetting
    Application.DisplayAlerts = True
End Sub


Function CreateSheet(sourceSheet, targetExcel, titelSpalte, scanSpalte1, scanSpalte2, firstRow, startJahr, startMonat, endeJahr, endeMonat, paramDatenarten)
    
    Dim datenarten, dArt, j, m, i, cLabel, col, colName, lastRow
    Dim fromMonth, toMonth
    Dim targetSheet As Worksheet
    Dim z As Workbook
    
    ' Trage SheetNamen im "Qlik" Sheet am Ende der Spalte C ein
    lastRow = lastFilledRow(targetExcel.Sheets("Qlik"), "C")
    targetExcel.Sheets("Qlik").range("C" & (lastRow + 1)).value = sourceSheet.Name
    
    ' Create sheet at the end of the given sheets
    Set targetSheet = targetExcel.Sheets.Add(, targetExcel.Sheets(targetExcel.Sheets.Count))
    targetSheet.Name = sourceSheet.Name
    sourceSheet.Activate
    lastRow = lastFilledRow(sourceSheet, titelSpalte)
    sourceSheet.Rows(firstRow & ":" & lastRow).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    targetSheet.Activate
    targetSheet.range("A" & firstRow).Select
    targetSheet.Paste
    range("A1").value = "Level"
    Call AssignLevel(titelSpalte, LevelSpalte, firstRow, lastRow)
    range("B1").value = "Label"
    
    datenarten = Split(paramDatenarten, ";")
    col = ColTextToNum(StartQlikColumn)
    
    For j = startJahr To endeJahr
        If j = startJahr Then fromMonth = startMonat Else fromMonth = 1
        If j = endeJahr Then toMonth = endeMonat Else toMonth = 12
        For m = fromMonth To toMonth
            For dArt = 0 To UBound(datenarten)
                For i = 1 To 2
                    cLabel = j & Right("0" & m, 2)
                    If i = 1 Then cLabel = cLabel & "#" Else cLabel = cLabel & "%"
                    cLabel = cLabel & datenarten(dArt)
                    Cells(1, col).value = cLabel
                    col = col + 1
                Next i
                'FillGetICVal("E:F", 6, LastFilledRow("B"))
                If ColNumToText(col - 2) = StartQlikColumn Then
                    ' Beim ersten Loop formen wir die Formeln um gemäß vorlage in Spalten E:F und kopieren sie
                    Call FillGetICVal(scanSpalte1, scanSpalte2, ColNumToText(col - 2), ColNumToText(col - 1), firstRow, lastRow)
                    range(StartQlikColumn & firstRow & ":" & ColNumToText(ColTextToNum(StartQlikColumn) + 1) & lastRow).Select
                    Selection.Copy
                Else
                    ' Bei den folgenden Malen kopieren wir uns die spalten mit den neuen Formeln
                    range(ColNumToText(col - 2) & firstRow).Activate
                    ActiveSheet.Paste
                End If
            Next dArt
        Next m
    Next j
    
    'Autofit alle Spaltenbreiten
    Columns("A:" & ColNumToText(col)).EntireColumn.AutoFit
    
    'Lösche die ehemaligen spalten zwischen Label und AQ, damit nicht unnötige Formeln berechnet werden
    For col = ColTextToNum(titelSpalte) + 1 To ColTextToNum(StartQlikColumn) - 1
        colName = ColNumToText(col)
        'Range(colName & "1").UnMerge
        range(colName & "1").value = "Ignore" & col
        If range(colName & (firstRow + 1)).Text <> "" Then
            ' Lösche spalteninhalt
            range(colName & "2:" & colName & lastRow).UnMerge
            range(colName & "2:" & colName & lastRow).Clear
        End If
        Columns(ColNumToText(col) & ":" & ColNumToText(col)).ColumnWidth = 0.75
    Next col

End Function


Function FillGetICVal(scanSpalte1, scanSpalte2, AusgSpalte1, AusgSpalte2, ErsteZeile, LetzteZeile)
    'Const AusgabeSpalte = "AQ", SpalteScanDatenAbsolut = "E"
    Dim rw, formel, formelTeil, ScanSpalten
   ' Const AusgSpalte1 = "AQ", AusgSpalte2 = "AR"
    
    
    For rw = ErsteZeile To LetzteZeile
        formel = range(scanSpalte1 & rw).Formula2R1C1
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
            range(AusgSpalte1 & rw).Formula2R1C1 = formel
        End If
        formel = range(scanSpalte2 & rw).Formula2R1C1
        range(AusgSpalte2 & rw).Formula2R1C1 = formel
    Next rw
    
    FillGetICVal = 0 ' return value
End Function



Function AssignLevel(titelSpalte, LevelSpalte, StartZeile, EndZeile)
    'Const TitelSpalte = "B", LevelSpalte = "A", StartZeile = 4
    Dim rw

    For rw = StartZeile To EndZeile
        If Trim(range(titelSpalte & rw).Formula) = "" Then
            range(LevelSpalte & rw) = ""
        ElseIf range(titelSpalte & rw).Interior.ColorIndex = 55 Then
            range(LevelSpalte & rw) = 1
        Else
            ' Finde den Level an Zeilen-Gruppierung
            range(LevelSpalte & rw) = Rows(rw).OutlineLevel + 1
        End If
    Next rw
    AssignLevel = 0 ' return value
End Function


