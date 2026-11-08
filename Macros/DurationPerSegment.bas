Attribute VB_Name = "DurationPerSegment"
Sub DurationPerSegment()
'
' DurationPerSegment Macro
'

'
    Sheets.Add After:=ActiveSheet
    Sheets("KeyStroke Data").Select
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Hoja1").Select
    Columns("H:H").Select
    Selection.Copy
    Sheets("KeyStroke Data").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("I2").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("Hoja2").Select
    ActiveCell.FormulaR1C1 = "Segment ID"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A49"), Type:=xlFillSeries
    Range("A2:A49").Select
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Duration"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIFS('KeyStroke Data'!C[7],'KeyStroke Data'!C[5],Hoja2!RC[-1])"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B49")
    Range("B2:B49").Select
End Sub
