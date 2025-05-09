Attribute VB_Name = "Module19"
Sub eMASSter_Formatting_2()
Attribute eMASSter_Formatting_2.VB_Description = "Formatting and filtering for C2SOC generated eMASSter ACAS scan reports."
Attribute eMASSter_Formatting_2.VB_ProcData.VB_Invoke_Func = "h\n14"
'
' eMASSter_Formatting_2 Macro
' Formatting and filtering for C2SOC generated eMASSter ACAS scan reports.
'
' Keyboard Shortcut: Ctrl+h
'
    Rows("2:2").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.RowHeight = 32
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("Z2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True
    Range("A2").Select
    Sheets("Nessus Summary").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.RowHeight = 32
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.SmallScroll Down:=-1
    Columns("A:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
    ActiveWindow.SmallScroll Down:=0
    Columns("E:E").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="Yes", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("G6").Select
    ActiveWindow.SmallScroll Down:=-1
    Sheets("Nessus Details").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.RowHeight = 32
    ActiveWindow.ScrollRow = 3657
    ActiveWindow.ScrollRow = 3652
    ActiveWindow.ScrollRow = 3632
    ActiveWindow.ScrollRow = 3622
    ActiveWindow.ScrollRow = 3597
    ActiveWindow.ScrollRow = 3563
    ActiveWindow.ScrollRow = 3366
    ActiveWindow.ScrollRow = 3262
    ActiveWindow.ScrollRow = 2788
    ActiveWindow.ScrollRow = 2557
    ActiveWindow.ScrollRow = 2088
    ActiveWindow.ScrollRow = 1901
    ActiveWindow.ScrollRow = 1467
    ActiveWindow.ScrollRow = 1279
    ActiveWindow.ScrollRow = 2
    ActiveWindow.SmallScroll Down:=-3
    Range("A:A,C:C,G:G,H:H,J:J").Select
    Range("J1").Activate
    ActiveWindow.LargeScroll ToRight:=1
    Range("A:A,C:C,G:G,H:H,J:J,L:L").Select
    Range("L1").Activate
    ActiveWindow.LargeScroll ToRight:=1
    Range("A:A,C:C,G:G,H:H,J:J,L:L,U:U").Select
    Range("U1").Activate
    ActiveWindow.LargeScroll ToRight:=1
    Range("A:A,C:C,G:G,H:H,J:J,L:L,U:U,AB:AH,AJ:AO").Select
    Range("AJ1").Activate
    ActiveWindow.LargeScroll ToRight:=1
    Range("A:A,C:C,G:G,H:H,J:J,L:L,U:U,AB:AH,AJ:AO,AQ:AQ").Select
    Range("AQ1").Activate
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.LargeScroll ToRight:=-2
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SmallScroll Down:=-2
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("X2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Hidden = True
    Range("D:D,F:F,G:G").Select
    Range("G1").Activate
    ActiveWindow.LargeScroll ToRight:=1
    Range("D:D,F:F,G:G,K:M").Select
    Range("K1").Activate
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    Range("D:D,F:F,G:G,K:M,V:V").Select
    Range("V1").Activate
    Selection.ColumnWidth = 32
    Columns("P:U").Select
    Range("U1").Activate
    ActiveWindow.SmallScroll ToRight:=-4
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    Selection.ColumnWidth = 15
    Range("G3").Select
    ActiveWorkbook.Worksheets("Nessus Details").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Nessus Details").Sort.SortFields.Add2 Key:=Range( _
        "H2:H3669"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "Critical,High,Moderate,Low", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Nessus Details").Sort.SortFields.Add2 Key:=Range( _
        "I2:I3669"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Nessus Details").Sort
        .SetRange Range("A1:V3669")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=-2
    ActiveSheet.Range("$A$1:$V$3669").AutoFilter Field:=8, Criteria1:="None"
    ActiveWindow.SmallScroll Down:=-7
    Rows("179:179").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$V$178").AutoFilter Field:=8
    ActiveWindow.SmallScroll Down:=-3
End Sub
