Attribute VB_Name = "DurationInSeconds"

Sub DurationInSeconds()
'
' DurationInSeconds Macro
'

'
    Sheets("KeyStroke Data").Select
    Columns("H:H").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Replace What:="2024-03-10", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, FormulaVersion:= _
        xlReplaceFormula2
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=":", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), DecimalSeparator _
        :=".", ThousandsSeparator:=",", TrailingMinusNumbers:=True
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "hours"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "min"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "seconds"
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "hours in min"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "total min"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "min in seconds"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "total seconds"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "duration"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*60"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B4201")
    Range("B2:B4201").Select
    Range("D2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-1]"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D4201")
    Range("D2:D4201").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*60"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E4201")
    Range("E2:E4201").Select
    Range("G2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]+RC[-1]"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G4201")
    Range("G2:G4201").Select
    Range("H2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[1]C[-1]-RC[-1]"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H4201")
    Range("H2:H4201").Select
End Sub
