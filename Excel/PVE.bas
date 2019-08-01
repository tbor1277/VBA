Sub copycolumns()
'
' copycolumns Macro
'
' Keyboard Shortcut: Ctrl+u
'
    Columns("D:D").EntireColumn.Select
    Selection.Copy
    Windows("Macro.xlsm").Activate
    Range("A1").Select
    ActiveSheet.Paste

End Sub

Sub extract()
'
' extract Macro
'
' Keyboard Shortcut: Ctrl+i
'
Dim lastCell As Long
lastCell = Cells(1, Columns.Count).End(xlToLeft).Column

    Range("A1", "A" & Columns.Count).Select
    Selection.Copy
    Cells(1, lastCell + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Cells(1, lastCell + 1).Select
    ActiveCell.FormulaR1C1 = "=extract_value(RC[-1])"
    Selection.AutoFill Destination:=Cells(1, lastCell + 1)
    'Range("B2", "B" & lastCell).Select
    'Selection.Copy
    
    
        
End Sub

Sub secondcopy()
'
' extract Macro
'
' Keyboard Shortcut: Ctrl+o
'
    Columns("R:R").EntireColumn.Select
    Selection.Copy
    Windows("Macro.xlsm").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Range("B2", "B" & Columns.Count).Select
    Selection.Copy
        
End Sub
