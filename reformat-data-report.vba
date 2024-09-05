Sub datareformat()
'
' datareformat Macro
'
' Keyboard Shortcut: Ctrl+h
'
    Range("A2").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Columns("F:F").ColumnWidth = 13.29
    Columns("C:C").ColumnWidth = 14.43
    Columns("C:C").Select
    Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 3), TrailingMinusNumbers:=True
    Columns("B:B").EntireColumn.AutoFit
    Range("A1").Select
    Dim thisLastRow
        thisLastRow = Last_Row()
        
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "A2:A4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "B2:B" & thisLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "C2:C" & thisLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range( _
        "F2:F" & thisLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:M" & thisLastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove ' RMK=K
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Cr"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<0,RC[-1],"""")"
    Range("J2").Select
    Selection.Copy
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range("J2:J" & thisLastRow).Select
    ActiveSheet.Paste
    Columns("J:J").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("J:J").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Db"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]>=0,RC[-1],"""")"
    Range("J2").Select
    Selection.Copy
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range("J2:J" & thisLastRow).Select
    ActiveSheet.Paste
    Columns("J:J").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("I:I").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Balance"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]&RC[-2]"
    Range("L2").Select
    Selection.Copy
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range("L2:L" & thisLastRow).Select
    ActiveSheet.Paste
    Columns("L:L").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("A1").Select
    Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(12), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    ActiveWindow.SmallScroll Down:=-3
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    Selection.AutoFilter
    
    With ActiveWindow
        .Width = 860
        .Height = 420
    End With
    
End Sub


' source: https://www.wallstreetmojo.com/vba-last-row/
Function Last_Row()

  Dim LR As Long
  'For understanding LR = Last Row

  LR = Cells(Rows.Count, 1).End(xlUp).Row

  'MsgBox LR
  Last_Row = LR
  

End Function
