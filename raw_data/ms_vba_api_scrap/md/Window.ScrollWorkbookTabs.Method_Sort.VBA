Sub Macro1()
'
' Macro1 Macro
'

'
    Columns("J:J").Select
    ActiveWorkbook.Worksheets("INSP").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("INSP").Sort.SortFields.Add Key:=Range("J1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("INSP").Sort
        .SetRange Range("A2:N3525")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("H2").Select
    ActiveWindow.SmallScroll Down:=-12
    Columns("F:F").Select
    Selection.NumberFormat = "m/d/yyyy"
    Range("F1").Select
    ActiveWindow.SmallScroll Down:=-42
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("Summary").Select
    Range("E4:S4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("N27").Select
    Sheets("INSP").Select
    Sheets("INSP").Move Before:=Sheets(2)
    Sheets("MTSS").Select
    Sheets("MTSS").Move Before:=Sheets(4)
    Sheets("FOREX").Select
    Sheets("FOREX").Move Before:=Sheets(5)
    Sheets("SECURITY").Select
    Sheets("SECURITY").Move Before:=Sheets(4)
    Sheets("MTSS").Select
    Sheets("MTSS").Move Before:=Sheets(4)
    Sheets("HIGH RISK").Select
    Sheets("HIGH RISK").Move Before:=Sheets(7)
    Sheets("Summary").Select
    ActiveWorkbook.Save
End Sub
