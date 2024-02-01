Sub addDaysRows()
    Dim daysCount As Variant
    Dim generalCarriers() As Variant
    Dim Carriers() As Variant
    Dim fullWs As Worksheet
    Dim shortWs As Worksheet
    
    Set fullWs = Sheets("Полная")
    Set shortWs = Sheets("Сводная")
    
    
    Application.Calculation = xlCalculationManual
    
    daysCount = InputBox("Введите количество дней, которое нужно добавить")
    daysCount = Replace(daysCount, ".", ",")
    
    If Not IsNumeric(daysCount) Then
        MsgBox ("Введено не число")
        Exit Sub
    End If
    
    badDaysCount = True
    Select Case daysCount
        Case Is = Empty
            MsgBox ("Вы ничего не ввели")
        Case Is > 31
            MsgBox ("Введено слишком большое число. Максимальное количество дней для добавления: 31")
        Case Is < 1
            MsgBox ("Количество дней должно быть больше 0")
        Case Is <> CInt(daysCount)
            MsgBox ("Введено не целое число")
        Case Else
                badDaysCount = False
    End Select
    
    If badDaysCount = True Then Exit Sub
    
    generalCarriers = Range("generalCarriers[Перевозчик]").Value
    Carriers = Range("Carrirers[Перевозчик]").Value
    
    'сводная
    shortWs.Select
    With shortWs
        lastRowShort = .Cells.SpecialCells(xlLastCell).Row
        lastColumnShort = .Cells.SpecialCells(xlLastCell).Column
        Set findDaysTableShort = .Range(.Cells(1, 1), .Cells(lastRowShort, lastColumnShort)).Find("Таблица по дням")
        For j = 1 To daysCount
            lastRowShort = .Cells.SpecialCells(xlLastCell).Row
            For i = LBound(generalCarriers) To UBound(generalCarriers)
                .Cells(lastRowShort + i, 1) = .Cells(lastRowShort - UBound(generalCarriers) + 1, 1).Value + 1
                .Cells(lastRowShort + i, 2) = generalCarriers(i, 1)
                .Cells(lastRowShort + i, 3) = .Cells(lastRowShort - UBound(generalCarriers) + 1, 3).Formula
                .Cells(lastRowShort + i, 4).FormulaR1C1 = .Cells(lastRowShort, 4).FormulaR1C1
                .Cells(lastRowShort + i, 5).FormulaR1C1 = .Cells(lastRowShort, 5).FormulaR1C1
                .Cells(lastRowShort + i, 6).FormulaR1C1 = .Cells(lastRowShort, 6).FormulaR1C1
                '.Cells(lastRowShort + i, 7) = CLng(.Cells(lastRowShort + i, 1)) & .Cells(lastRowShort + i, 2)
            Next i
        Next j
        
        firstRowForFormat = (lastRowShort + 1 - (UBound(generalCarriers) * daysCount))
        lastRowForFormat = (lastRowShort - (UBound(generalCarriers) * daysCount)) + UBound(generalCarriers)
        firstRowAfter = (lastRowShort + 1 - (UBound(generalCarriers) * daysCount)) + UBound(generalCarriers)
        lastRowAfter = lastRowShort + UBound(generalCarriers)
        
        .Rows(firstRowForFormat & ":" & lastRowForFormat).Select 'формат по образцу
        Selection.Copy
        .Rows(firstRowAfter & ":" & lastRowAfter).Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With
    
    'полная
    With fullWs
        For j = 1 To daysCount
        lastRowFull = .Cells(1, 1).CurrentRegion.Rows.Count
        lastColumnFull = .Cells(1, 1).CurrentRegion.Columns.Count
            For i = LBound(Carriers) To UBound(Carriers)
            .Cells(lastRowFull + i, 1) = .Cells(lastRowFull, 1).Value + 1
            .Cells(lastRowFull + i, 2).FormulaR1C1 = .Cells(lastRowFull, 2).FormulaR1C1
            .Cells(lastRowFull + i, 3).FormulaR1C1 = .Cells(lastRowFull, 3).FormulaR1C1
            .Cells(lastRowFull + i, 4) = Carriers(i, 1)
            .Cells(lastRowFull + i, 5) = .Cells(lastRowFull - UBound(Carriers) + i, 5).Formula
            .Cells(lastRowFull + i, 8).FormulaR1C1 = .Cells(lastRowFull, 8).FormulaR1C1
            Next i
        Next j
        
        firstRowForFormat = (lastRowFull + 1 - (UBound(Carriers) * daysCount))
        lastRowForFormat = (lastRowFull - (UBound(Carriers) * daysCount)) + UBound(Carriers)
        firstRowAfter = (lastRowFull + 1 - (UBound(Carriers) * daysCount)) + UBound(Carriers)
        lastRowAfter = lastRowFull + UBound(Carriers)
    End With
    
    fullWs.Select
    
    Rows(firstRowForFormat & ":" & lastRowForFormat).Select 'формат по образцу
    Selection.Copy
    Rows(firstRowAfter & ":" & lastRowAfter).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    ' rangeWithFormat = range(cells(firstRowForFormat, 1), cells(lastRowForFormat, lastColumnFull))
    ' rangeToFormat = range(Cells(firstRowAfter, 1), Cells(lastRowAfter, lastColumnFull))

    ' rangeToFormat.NumberFormat = rangeWithFormat.NumberFormat
    ' rangeToFormat.Font.Name = rangeWithFormat.Font.Name
    ' rangeToFormat.Font.Size = rangeWithFormat.Font.Size
    ' rangeToFormat.Font.Color = rangeWithFormat.Font.Color
    ' rangeToFormat.Interior.Color = rangeWithFormat.Interior.Color

    ' For Each cond In rangeWithFormat.FormatConditions
    '     Set condRule = rangeToFormat.FormatConditions.Add(Type:=cond.Type, Operator:=cond.Operator, Formula1:=cond.Formula1, Formula2:=cond.Formula2)
    '     condRule.Interior.Color = cond.Interior.Color
    '     condRule.Font.Color = cond.Font.Color
    ' Next cond

    Cells(lastRowFull + UBound(Carriers), 1).Select
    Application.Calculation = xlCalculationAutomatic

End Sub


